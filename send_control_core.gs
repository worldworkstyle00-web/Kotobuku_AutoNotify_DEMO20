/***********************
 * send_control_core.gs
 * - Menu (Kotobuku / SendControl)
 * - Schema Candidate/Company
 * - Diff Engine (Snapshot)
 *
 * Note:
 * - v3由来の中核を集約
 * - 参照先不明の関数は安全スタブで落ちないようにする
 ***********************/

/**
 * ===== Global Constants =====
 */
const SHEETS = {
  MAIN: '申請一覧',
  SCHEMA_COMPANY: 'Schema_Company',
  SCHEMA_CANDIDATE: 'Schema_Candidate',
  SCHEMA_LOG: 'Schema_Log', // optional
  SNAPSHOT: 'Snapshot_申請一覧',
  DIFF_RUN_LOG: 'Diff_Run_Log',
};

const SCHEMA = {
  // Candidate/Company 共通の列構造（A-H）
  COLS: {
    SOURCE_SHEET: 1,  // A
    SOURCE_COLUMN: 2, // B
    PROP_NAME: 3,     // C
    LABEL: 4,         // D
    DATA_TYPE: 5,     // E
    FIELD_TYPE: 6,    // F
    ADD_TO_SCHEMA: 7, // G (checkbox)
    NOTE: 8,          // H
  },
  COMPANY_CREATE_COL_NAME: 'create_in_hubspot', // 末尾に追加されるチェック列
};

const DIFF_CONFIG = {
  sheetMain: SHEETS.MAIN,
  sheetSnapshot: SHEETS.SNAPSHOT,
  sheetRunLog: SHEETS.DIFF_RUN_LOG,

  headerRow: 1,

  // 申請一覧のヘッダー名（完全一致）
  colApplicationNo: '申請番号',
  colSerialNo: 'シリアル番号',

  colStatusPrev: '前日ステータス',
  colStatusCurr: '当日ステータス',

  colStatusUpdatedAt: 'ステータス更新日',
  colLastUpdatedAt: '最終更新日',

  colLogId: 'ログID',

  trackedColumns: [
    '申請番号',
    'シリアル番号',
    '前日ステータス',
    '当日ステータス',
  ],

  logIdDateFormat: 'yyyyMMddHHmmss',
  timeZone: 'Asia/Tokyo',
};

/**
 * ===== Menu =====
 * - v3の構成を踏襲しつつ、二重メニュー追加を排除
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Kotobuku メニュー
  ui.createMenu('Kotobuku')
    .addItem('申請一覧 → Schema候補生成', 'generateSchemaCandidate')
    .addItem('✅ Candidate → Company に反映', 'applyCandidateToSchemaCompany')
    .addSeparator()
    .addItem('差分検知 → 更新日付与 → スナップショット更新', 'runDiffAndUpdate')
    .addToUi();

  // 既存構成に合わせて “枠” は残す（中身が別ファイルにある/削除される可能性があるため）
  const hubSpotSyncMenu = ui.createMenu('HubSpotSync')
    .addItem('同期：申請一覧 → HubSpot(Companies)', 'menu_syncCompanies')
    .addSeparator()
    .addItem('初期設定：Companyカスタム項目を作成/確認', 'menu_ensureCompanyProps');

  const sendTodayMenu = ui.createMenu('Send_Today')
    .addItem('①送信対象を自動生成（send_today更新）', 'generateSendToday')
    .addItem('②選択行を送信対象から除外', 'excludeSelectedRowsFromSend')
    .addSeparator()
    .addItem('③送信実行（A案）', 'runSendBatch_A');

  ui.createMenu('SendControl')
    .addItem('送信制御列の整備（列名統一/型/チェック）', 'initSendControlColumns')
    .addSeparator()
    .addSubMenu(hubSpotSyncMenu)
    .addSubMenu(sendTodayMenu)
    .addToUi();
    
    // DEMO20 メニューも追加（demo20_mailer.gs 側の関数を呼ぶ）
  try {
    if (typeof onOpen_demo20_mailer_ === 'function') onOpen_demo20_mailer_();
  } catch (e) {
    Logger.log(`DEMO20 menu init failed: ${e}`);
  }

}

/**
 * ===== Schema: Candidate =====
 * - 申請一覧の列名をスキャン
 * - Schema_Company に既に登録済みの列は除外
 * - Schema_Candidate を「最新状態」に作り直す（肥大化防止）
 * - ただし add_to_schema チェックは復元する
 */
function generateSchemaCandidate() {
  const ss = SpreadsheetApp.getActive();
  const sourceSheet = mustGetSheet_(ss, SHEETS.MAIN);
  const companySheet = mustGetSheet_(ss, SHEETS.SCHEMA_COMPANY);
  const candidateSheet = mustGetSheet_(ss, SHEETS.SCHEMA_CANDIDATE);

  // ① Candidateのチェック状態を退避（labelベース）
  const checkMap = snapshotCandidateChecks_(candidateSheet);

  // ② Candidateを全消し（ヘッダー以外）
  clearSheetBody_(candidateSheet);

  // ③ 申請一覧ヘッダー取得
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0]
    .map(v => String(v || '').trim())
    .filter(v => v);

  // ④ Schema_Company に既にある source_column を取得（B列）
  const registeredColumns = getRegisteredSourceColumns_(companySheet); // Set<string>

  // ⑤ 未登録列だけ Candidate に再生成
  const rowsToAppend = [];
  sourceHeaders.forEach(col => {
    if (registeredColumns.has(col)) return;
    rowsToAppend.push([
      SHEETS.MAIN,            // source_sheet
      col,                   // source_column
      toPropertyName_(col),  // suggested_property_name
      col,                   // label
      'string',              // data_type (仮)
      'text',                // field_type (仮)
      false,                 // add_to_schema (checkbox)
      '',                    // note
    ]);
  });

  if (rowsToAppend.length > 0) {
    candidateSheet.getRange(candidateSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);

    // checkbox適用（G列）
    const startRow = 2;
    const endRow = candidateSheet.getLastRow();
    candidateSheet.getRange(startRow, SCHEMA.COLS.ADD_TO_SCHEMA, endRow - 1, 1).insertCheckboxes();
  }

  // ⑥ チェック復元
  restoreCandidateChecks_(candidateSheet, checkMap);

  SpreadsheetApp.getUi().alert(`Schema_Candidate を更新しました（候補 ${rowsToAppend.length} 件）。`);
}

/**
 * ===== Schema: Company =====
 * - Candidate の add_to_schema ✅ のものだけ Company に追加
 * - 追加した行に対して create_in_hubspot チェック列を自動付与（列がなければ作る）
 * - 重複は source_column(B列) で防止
 */
function applyCandidateToSchemaCompany() {
  const ss = SpreadsheetApp.getActive();
  const candidateSheet = mustGetSheet_(ss, SHEETS.SCHEMA_CANDIDATE);
  const companySheet = mustGetSheet_(ss, SHEETS.SCHEMA_COMPANY);
  const sourceSheet = mustGetSheet_(ss, SHEETS.MAIN);
  const logSheet = ss.getSheetByName(SHEETS.SCHEMA_LOG); // optional

  // --- 0) create_in_hubspot 列を保証（列番号取得） ---
  const createCol = ensureCreateInHubspotColumn_(companySheet);

  // --- 1) 申請一覧の「現在ヘッダー」をSet化（これに存在しないsource_columnは掃除対象） ---
  const currentHeaders = sourceSheet
    .getRange(1, 1, 1, sourceSheet.getLastColumn())
    .getValues()[0]
    .map(v => String(v || '').trim())
    .filter(v => v);
  const currentHeaderSet = new Set(currentHeaders);

  // --- 2) Schema_Company 現状の行を読み取り、チェックを退避しつつ「生存行」だけ残す ---
  const companyLastRow = companySheet.getLastRow();
  const keptRows = []; // A-H の行だけ（create列は別で付ける）
  const createCheckMap = {}; // key -> boolean

  if (companyLastRow >= 2) {
    const lastCol = companySheet.getLastColumn();
    const values = companySheet.getRange(2, 1, companyLastRow - 1, lastCol).getValues();

    values.forEach(row => {
      const sourceSheetName = String(row[SCHEMA.COLS.SOURCE_SHEET - 1] || '').trim();
      const sourceColumn = String(row[SCHEMA.COLS.SOURCE_COLUMN - 1] || '').trim();

      if (!sourceColumn) return;

      // create_in_hubspot のチェックを退避（列が末尾でもOK）
      const checked = row[createCol - 1] === true;
      const mapKey = `${sourceSheetName}|${sourceColumn}`;
      if (checked) createCheckMap[mapKey] = true;

      // ★掃除ルール：申請一覧に存在しない source_column は落とす（＝旧列名を削除）
      if (!currentHeaderSet.has(sourceColumn)) return;

      // A-H だけ保持（Schema_Companyの本体部分）
      keptRows.push(row.slice(0, SCHEMA.COLS.NOTE));
    });
  }

  // --- 3) Candidate ✅ を取り込み（keptRows の source_column と重複は除外） ---
  const existing = new Set(keptRows.map(r => String(r[SCHEMA.COLS.SOURCE_COLUMN - 1] || '').trim()).filter(Boolean));

  const cLastRow = candidateSheet.getLastRow();
  if (cLastRow < 2) {
    uiAlert_('Schema_Candidate にデータがありません。');
    return;
  }
  const cValues = candidateSheet.getRange(2, 1, cLastRow - 1, SCHEMA.COLS.NOTE).getValues();

  const now = new Date();
  const logRows = [];

  cValues.forEach(r => {
    const sourceSheetName = String(r[SCHEMA.COLS.SOURCE_SHEET - 1] || '').trim();
    const sourceColumn = String(r[SCHEMA.COLS.SOURCE_COLUMN - 1] || '').trim();
    const propName = r[SCHEMA.COLS.PROP_NAME - 1];
    const label = r[SCHEMA.COLS.LABEL - 1];
    const dataType = r[SCHEMA.COLS.DATA_TYPE - 1];
    const fieldType = r[SCHEMA.COLS.FIELD_TYPE - 1];
    const addFlag = r[SCHEMA.COLS.ADD_TO_SCHEMA - 1];
    const note = r[SCHEMA.COLS.NOTE - 1];

    if (addFlag !== true) return;
    if (!sourceColumn) return;

    // 申請一覧に存在しない列は今回は採用しない（掃除方針と整合）
    if (!currentHeaderSet.has(sourceColumn)) return;

    if (existing.has(sourceColumn)) return;

    keptRows.push([
      sourceSheetName,
      sourceColumn,
      propName,
      label,
      dataType,
      fieldType,
      true, // add_to_schema は採用済みとして true
      note || ''
    ]);
    existing.add(sourceColumn);

    if (logSheet) logRows.push([now, 'ADD', sourceSheetName, sourceColumn, propName, label]);
  });

  if (keptRows.length === 0) {
    uiAlert_('Schema_Company に残す/追加する行がありません（申請一覧ヘッダーと一致するものが0件）。');
    return;
  }

  // --- 4) Schema_Company をクリーンにして再構築（ヘッダーは残す） ---
  const maxRows = companySheet.getMaxRows();
  const maxCols = companySheet.getLastColumn();
  if (maxRows > 1) companySheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();

  // A-H を書き戻し
  companySheet.getRange(2, 1, keptRows.length, keptRows[0].length).setValues(keptRows);

  // --- 5) create_in_hubspot 列を false + checkbox にし、退避チェックを復元 ---
  const createRange = companySheet.getRange(2, createCol, keptRows.length, 1);
  const createVals = Array(keptRows.length).fill([false]);

  for (let i = 0; i < keptRows.length; i++) {
    const sourceSheetName = String(keptRows[i][SCHEMA.COLS.SOURCE_SHEET - 1] || '').trim();
    const sourceColumn = String(keptRows[i][SCHEMA.COLS.SOURCE_COLUMN - 1] || '').trim();
    const key = `${sourceSheetName}|${sourceColumn}`;
    if (createCheckMap[key] === true) createVals[i][0] = true;
  }

  createRange.setValues(createVals);
  createRange.insertCheckboxes();

  // --- 6) ログ追記（任意） ---
  if (logSheet && logRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
  }

  uiAlert_(`Schema_Company を再構築しました（残し+追加 合計 ${keptRows.length} 行）。`);
}

/**
 * ===== Diff Engine =====
 * - Snapshot_申請一覧 と比較して差分検知
 * - 変更があった行の「最終更新日」を更新
 * - ステータス変更があった行は「ステータス更新日」と「ログID(秒まで)」を更新
 * - Snapshot を全差し替え
 * - Diff_Run_Log に実行ログ追記
 */
function runDiffAndUpdate() {
  const ss = SpreadsheetApp.getActive();
  const main = mustGetSheet_(ss, DIFF_CONFIG.sheetMain);
  const snap = mustGetSheet_(ss, DIFF_CONFIG.sheetSnapshot);
  const log = mustGetSheet_(ss, DIFF_CONFIG.sheetRunLog);

  const now = new Date();
  const runId = Utilities.getUuid();

  const header = main.getRange(DIFF_CONFIG.headerRow, 1, 1, main.getLastColumn()).getValues()[0];
  const colIndex = headerIndexMap_(header);

  // 必須チェック
  const required = [
    DIFF_CONFIG.colApplicationNo,
    DIFF_CONFIG.colSerialNo,
    DIFF_CONFIG.colStatusPrev,
    DIFF_CONFIG.colStatusCurr,
    DIFF_CONFIG.colStatusUpdatedAt,
    DIFF_CONFIG.colLastUpdatedAt,
    DIFF_CONFIG.colLogId,
  ];
  for (const h of required) {
    if (!(h in colIndex)) return uiAlert_(`申請一覧のヘッダーが見つかりません: 「${h}」`);
  }
  for (const h of DIFF_CONFIG.trackedColumns) {
    if (!(h in colIndex)) return uiAlert_(`trackedColumns に指定されたヘッダーが見つかりません: 「${h}」`);
  }

  const lastRow = main.getLastRow();
  if (lastRow <= DIFF_CONFIG.headerRow) return uiAlert_('申請一覧にデータ行がありません。');

  const values = main.getRange(DIFF_CONFIG.headerRow + 1, 1, lastRow - DIFF_CONFIG.headerRow, main.getLastColumn()).getValues();
  const snapMap = loadSnapshotMap_(snap);

  let rowsChanged = 0;
  let statusChanged = 0;

  const idxStatusUpdatedAt = colIndex[DIFF_CONFIG.colStatusUpdatedAt];
  const idxLastUpdatedAt = colIndex[DIFF_CONFIG.colLastUpdatedAt];
  const idxLogId = colIndex[DIFF_CONFIG.colLogId];

  const newSnapshotRows = [];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];

    const appNo = safeStr_(row[colIndex[DIFF_CONFIG.colApplicationNo]]);
    const serial = safeStr_(row[colIndex[DIFF_CONFIG.colSerialNo]]);
    if (!appNo || !serial) continue;

    const key = `${appNo}__${serial}`;

    const rowHash = makeRowHash_(row, colIndex, DIFF_CONFIG.trackedColumns);

    const prevStatus = safeStr_(row[colIndex[DIFF_CONFIG.colStatusPrev]]);
    const currStatus = safeStr_(row[colIndex[DIFF_CONFIG.colStatusCurr]]);
    const isStatusChanged = prevStatus !== currStatus && prevStatus !== '' && currStatus !== '';

    const prevSnap = snapMap.get(key);
    const isRowChanged = !prevSnap || prevSnap.row_hash !== rowHash;

    if (isRowChanged) {
      rowsChanged++;
      row[idxLastUpdatedAt] = now;

      if (isStatusChanged) {
        statusChanged++;
        row[idxStatusUpdatedAt] = now;

        const ts = Utilities.formatDate(now, DIFF_CONFIG.timeZone, DIFF_CONFIG.logIdDateFormat);
        row[idxLogId] = `${appNo}-${ts}-${serial}`;
      }
    }

    newSnapshotRows.push([
      now,
      key,
      rowHash,
      prevStatus,
      currStatus,
      isStatusChanged ? now : (prevSnap ? prevSnap.status_changed_at : ''),
      isRowChanged ? now : (prevSnap ? prevSnap.last_updated_at : ''),
    ]);
  }

  // 書き戻し（ここは “申請一覧” 全体 setValues なので、数式列がある場合は注意）
  // v3では想定済み。もし数式列が混在しているなら trackedColumns & 更新列を別にして部分更新へ移行が必要。
  main.getRange(DIFF_CONFIG.headerRow + 1, 1, values.length, values[0].length).setValues(values);

  rewriteSnapshot_(snap, newSnapshotRows);

  log.appendRow([
    now,
    runId,
    values.length,
    rowsChanged,
    statusChanged,
    `updated_at=${Utilities.formatDate(now, DIFF_CONFIG.timeZone, 'yyyy-MM-dd HH:mm:ss')}`,
  ]);

  uiAlert_(`完了: 全${values.length}行 / 変更${rowsChanged}行 / ステータス変更${statusChanged}行`);
}

/**
 * ===== Utilities (Schema) =====
 */
function getRegisteredSourceColumns_(companySheet) {
  const lastRow = companySheet.getLastRow();
  const set = new Set();

  if (lastRow < 2) return set;

  // B列（source_column）だけ取れば十分
  const values = companySheet.getRange(2, SCHEMA.COLS.SOURCE_COLUMN, lastRow - 1, 1).getValues();
  values.forEach(v => {
    const s = String(v[0] || '').trim();
    if (s) set.add(s);
  });
  return set;
}

function ensureCreateInHubspotColumn_(companySheet) {
  const headerRow = 1;
  const lastCol = companySheet.getLastColumn();
  const headers = companySheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim());

  let idx = headers.indexOf(SCHEMA.COMPANY_CREATE_COL_NAME) + 1;
  if (idx === 0) {
    idx = lastCol + 1;
    companySheet.getRange(headerRow, idx).setValue(SCHEMA.COMPANY_CREATE_COL_NAME);
  }
  return idx; // 1-based
}

function toPropertyName_(text) {
  return String(text || '')
    .trim()
    .toLowerCase()
    .replace(/[^\w]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

/**
 * ===== Utilities (Candidate Check Snapshot/Restore) =====
 */
function snapshotCandidateChecks_(candidateSheet) {
  const lastRow = candidateSheet.getLastRow();
  const lastCol = candidateSheet.getLastColumn();
  if (lastRow < 2) return {};

  const values = candidateSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = {}; // key -> boolean

  // key = source_sheet|label
  values.forEach(row => {
    const sourceSheet = String(row[SCHEMA.COLS.SOURCE_SHEET - 1] || '').trim();
    const label = String(row[SCHEMA.COLS.LABEL - 1] || '').trim();
    const checked = row[SCHEMA.COLS.ADD_TO_SCHEMA - 1] === true;
    if (!sourceSheet || !label) return;
    const key = `${sourceSheet}|${label}`;
    map[key] = map[key] || checked;
  });

  return map;
}

function restoreCandidateChecks_(candidateSheet, checkMap) {
  const lastRow = candidateSheet.getLastRow();
  const lastCol = candidateSheet.getLastColumn();
  if (lastRow < 2) return;

  const rng = candidateSheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = rng.getValues();

  for (let i = 0; i < values.length; i++) {
    const sourceSheet = String(values[i][SCHEMA.COLS.SOURCE_SHEET - 1] || '').trim();
    const label = String(values[i][SCHEMA.COLS.LABEL - 1] || '').trim();
    if (!sourceSheet || !label) continue;

    const key = `${sourceSheet}|${label}`;
    if (checkMap[key] === true) values[i][SCHEMA.COLS.ADD_TO_SCHEMA - 1] = true;
  }
  rng.setValues(values);

  // checkbox再適用（G列）
  candidateSheet.getRange(2, SCHEMA.COLS.ADD_TO_SCHEMA, lastRow - 1, 1).insertCheckboxes();
}

function clearSheetBody_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
}

/**
 * ===== Utilities (Diff Helpers) =====
 */
function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シートが見つかりません: ${name}`);
  return sh;
}

function uiAlert_(msg) {
  SpreadsheetApp.getUi().alert(msg);
}

function headerIndexMap_(headerRowValues) {
  const map = {};
  for (let i = 0; i < headerRowValues.length; i++) {
    const h = String(headerRowValues[i] || '').trim();
    if (h) map[h] = i;
  }
  return map;
}

function safeStr_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function makeRowHash_(row, colIndex, trackedHeaders) {
  const obj = {};
  trackedHeaders.forEach(h => {
    obj[h] = row[colIndex[h]];
  });
  const json = JSON.stringify(obj);
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, json)
  );
}

function loadSnapshotMap_(snapSheet) {
  const lastRow = snapSheet.getLastRow();
  const map = new Map();
  if (lastRow < 2) return map;

  const header = snapSheet.getRange(1, 1, 1, snapSheet.getLastColumn()).getValues()[0];
  const idx = headerIndexMap_(header);

  const required = ['key', 'row_hash', 'status_changed_at', 'last_updated_at'];
  for (const h of required) {
    if (!(h in idx)) throw new Error(`Snapshot_申請一覧 のヘッダーに「${h}」が見つかりません。`);
  }

  const values = snapSheet.getRange(2, 1, lastRow - 1, snapSheet.getLastColumn()).getValues();
  for (const r of values) {
    const key = safeStr_(r[idx['key']]);
    if (!key) continue;
    map.set(key, {
      row_hash: safeStr_(r[idx['row_hash']]),
      status_changed_at: r[idx['status_changed_at']],
      last_updated_at: r[idx['last_updated_at']],
    });
  }
  return map;
}

function rewriteSnapshot_(snapSheet, rows) {
  const maxRows = snapSheet.getMaxRows();
  const maxCols = snapSheet.getLastColumn();
  if (maxRows > 1) snapSheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
  if (rows.length === 0) return;
  snapSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * ===== Dev Helper =====
 */
function diagnoseSheetSize() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('申請一覧');

  const info = {
    lastCol: sh.getLastColumn(),
    lastRow: sh.getLastRow(),
    maxRows: sh.getMaxRows(),
    maxCols: sh.getMaxColumns(),
  };
  Logger.log(JSON.stringify(info));

  return;
}

/**
 * ===== Safety Stubs =====
 * v3にメニューとして存在するが、実装が別ファイルだった可能性が高いもの。
 * 2+1に統合後、クリックしても落ちないための保険。
 */
function __notIncluded_(name) {
  SpreadsheetApp.getUi().alert(`このメニュー機能（${name}）は、今回の2+1構成には含めていません。`);
}
function menu_syncCompanies() { __notIncluded_('同期：申請一覧 → HubSpot(Companies)'); }
function menu_ensureCompanyProps() { __notIncluded_('初期設定：Companyカスタム項目を作成/確認'); }
function initSendControlColumns() { __notIncluded_('送信制御列の整備'); }
function generateSendToday() { __notIncluded_('①送信対象を自動生成（send_today更新）'); }
function excludeSelectedRowsFromSend() { __notIncluded_('②選択行を送信対象から除外'); }
function runSendBatch_A() { __notIncluded_('③送信実行（A案）'); }
