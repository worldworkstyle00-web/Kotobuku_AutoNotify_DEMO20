/************************************************
 * demo20_lib.gs
 * - DEMO20 Mailer v4 shared utilities
 ************************************************/

function __normHeader_(v) {
  return String(v ?? '')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\n/g, ' ')
    .trim();
}

function __getSheetOrThrow_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function __loadSheetByHeaderFlex_(ss, sheetName, requiredHeaders, scanRows) {
  const sheet = __getSheetOrThrow_(ss, sheetName);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) throw new Error(`sheet empty: ${sheetName}`);

  const scanN = Math.min(scanRows ?? 5, lastRow);
  const top = sheet.getRange(1, 1, scanN, lastCol).getValues();

  const req = (requiredHeaders || []).map(__normHeader_);

  let foundRow0 = -1;
  let idx0 = null;

  for (let r = 0; r < top.length; r++) {
    const row = top[r].map(__normHeader_);
    const m = {};
    for (let c = 0; c < row.length; c++) {
      const h = row[c];
      if (h) m[h] = c; // 0-based
    }
    const ok = req.every(h => m[h] !== undefined);
    if (ok) {
      foundRow0 = r;
      idx0 = m;
      break;
    }
  }

  if (foundRow0 < 0) {
    const sample = (top[0] || []).map(__normHeader_).filter(Boolean).slice(0, 20).join(' | ');
    throw new Error(`【${sheetName}】header not found in first ${scanN} rows. row0 sample: ${sample}`);
  }

  const bodyRows = lastRow - (foundRow0 + 1);
  const body = bodyRows > 0
    ? sheet.getRange(foundRow0 + 2, 1, bodyRows, lastCol).getValues()
    : [];

  const headerRow = top[foundRow0];
  const valuesAll = [headerRow, ...body];

  return { sheet, headerRowIndex0: foundRow0, idx0, valuesAll };
}

function __mustCol0_(idx0, headerName) {
  const h = __normHeader_(headerName);
  const c0 = idx0[h];
  if (c0 === undefined) throw new Error(`missing header: "${headerName}"`);
  return c0;
}

function __findCol0_(idx0, candidates) {
  for (const name of candidates) {
    const h = __normHeader_(name);
    if (idx0[h] !== undefined) return idx0[h];
  }
  return undefined;
}

function __toBool_(v) {
  if (v === true) return true;
  const s = String(v ?? '').trim().toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'YES';
}

function __applyCellUpdates_(sheet, updates) {
  if (!updates || updates.length === 0) return;
  for (const u of updates) {
    if (!u) continue;
    const r = u.r, c = u.c;
    if (!Number.isFinite(r) || r < 1) throw new Error(`invalid row: ${r}`);
    if (!Number.isFinite(c) || c < 1) throw new Error(`invalid col: ${c}`);
    sheet.getRange(r, c, 1, 1).setValue(u.v);
  }
}

/* =========================================================
 * テンプレ決定（厳密ルール）
 * ======================================================= */

function decideTemplate_(curStatus, doneFlag, TPL_TSUTO, TPL_WORK) {
  const st = String(curStatus || '').trim();
  const f = Number(doneFlag || 0);

  if (f === 1 && st === '疎通完了') return TPL_TSUTO;
  if (f === 2 && st === '完了') return TPL_WORK;

  return '';
}

/* =========================================================
 * 申請一覧：送信対象の申請番号SetでMap化
 * Map<申請番号, { mail, company, person }>
 * ======================================================= */

function buildApplyInfoMapDemo20_(ss, targetIds) {
  const apply = __loadSheetByHeaderFlex_(ss, '申請一覧_DEMO20', ['申請番号'], 5);
  const a申請番号 = __mustCol0_(apply.idx0, '申請番号');

  const aMail = __findCol0_(apply.idx0, ['工事会社メールアドレス', 'メールアドレス', 'メール']);
  const aCompany = __findCol0_(apply.idx0, ['工事会社名', '会社名']);
  const aPerson = __findCol0_(apply.idx0, ['工事会社担当者名', '担当者名']);

  const map = new Map();

  for (let r = 1; r < apply.valuesAll.length; r++) {
    const row = apply.valuesAll[r];
    const appId = String(row[a申請番号] ?? '').trim();
    if (!appId) continue;
    if (!targetIds.has(appId)) continue;

    const mail = (aMail !== undefined) ? String(row[aMail] ?? '').trim() : '';
    const company = (aCompany !== undefined) ? String(row[aCompany] ?? '').trim() : '';
    const person = (aPerson !== undefined) ? String(row[aPerson] ?? '').trim() : '';

    map.set(appId, { mail, company, person });
  }

  return map;
}

/* =========================================================
 * テンプレロード
 * MailTemplate:
 *  template_key, template_name, subject, body
 * ======================================================= */

function loadTemplatesDemo20_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = __getSheetOrThrow_(ss, 'MailTemplate');
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return new Map();

  const headers = values[0].map(__normHeader_);
  const col = (name) => {
    const i = headers.indexOf(__normHeader_(name));
    if (i < 0) throw new Error(`MailTemplate missing header: ${name}`);
    return i;
  };

  const cKey = col('template_key');
  const cName = col('template_name');
  const cSub = col('subject');
  const cBody = col('body');

  const map = new Map();
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const key = String(row[cKey] ?? '').trim();
    if (!key) continue;
    map.set(key, {
      name: String(row[cName] ?? '').trim(),
      subject: String(row[cSub] ?? ''),
      body: String(row[cBody] ?? ''),
    });
  }
  return map;
}

function replaceTokens_(text, dict) {
  let out = String(text ?? '');
  for (const [k, v] of Object.entries(dict)) {
    out = out.replaceAll(`{{${k}}}`, String(v ?? ''));
  }
  return out;
}

function mergeSentHistory_(history, tplKey) {
  const h = String(history ?? '').trim();
  if (!h) return tplKey;
  const parts = h.split(',').map(s => s.trim()).filter(Boolean);
  if (parts.includes(tplKey)) return h;
  parts.push(tplKey);
  return parts.join(',');
}
