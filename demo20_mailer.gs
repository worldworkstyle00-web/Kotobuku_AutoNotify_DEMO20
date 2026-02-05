/************************************************
 * demo20_mailer.gs
 * - DEMO20 Mailer v4（公開関数 + メニュー）
 * - 実装本体のユーティリティは demo20_lib.gs に分離
 ************************************************/

function onOpen_demo20_mailer_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('DEMO20')
    .addItem('①キュー生成（prepare_send_queue_demo20）', 'prepare_send_queue_demo20')
    .addSeparator()
    .addItem('②送信（DRYRUN）run_send_demo20(true)', 'run_send_demo20_dryrun')
    .addItem('③送信（本番）run_send_demo20(false)', 'run_send_demo20_prod')
    .addToUi();
}

/** 既存のonOpenが別ファイルにあるので、干渉しないようラッパ */
function run_send_demo20_dryrun() { run_send_demo20(true); }
function run_send_demo20_prod() { run_send_demo20(false); }

/**
 * v4のCFGは lib 側に置く案もあるが、設定値は “見える場所” に置いた方がレビュー向きなのでここに残す
 * （内部ユーティリティは demo20_lib.gs）
 */
const CFG_DEMO20 = {
  SHEET_APPLY: '申請一覧_DEMO20',
  SHEET_SEND:  '送信管理_DEMO20',
  SHEET_TPL:   'MailTemplate',

  WORK_RELATED_STATUS: new Set(['疎通完了','工事依頼','工事中','完了']),
  HEADER_SCAN_ROWS: 5,
  RETRY_MINUTES: 10,
};

/* =========================================================
 * キュー生成：prepare_send_queue_demo20（v4そのまま）
 * ======================================================= */

function prepare_send_queue_demo20() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 申請一覧ロード（ヘッダー確定：申請番号/ステータス）
  const app = __loadSheetByHeaderFlex_(ss, CFG_DEMO20.SHEET_APPLY, ['申請番号', 'ステータス'], CFG_DEMO20.HEADER_SCAN_ROWS);
  const a申請番号 = __mustCol0_(app.idx0, '申請番号');
  const aステータス = __mustCol0_(app.idx0, 'ステータス');

  // id -> status
  const idToStatus = new Map();
  for (let r = 1; r < app.valuesAll.length; r++) {
    const row = app.valuesAll[r];
    const id = row[a申請番号];
    if (id === '' || id == null) continue;
    const st = String(row[aステータス] ?? '').trim();
    idToStatus.set(String(id), st);
  }

  // 送信管理ロード（必須）
  const send = __loadSheetByHeaderFlex_(ss, CFG_DEMO20.SHEET_SEND,
    ['申請番号', '発報対象', 'ステータス変更', '完了フラグ', 'キュー送信', '送信済テンプレ履歴', 'キューテンプレ'],
    CFG_DEMO20.HEADER_SCAN_ROWS
  );

  const s申請番号 = __mustCol0_(send.idx0, '申請番号');
  const s発報対象 = __mustCol0_(send.idx0, '発報対象');
  const sステータス変更 = __mustCol0_(send.idx0, 'ステータス変更');
  const s完了フラグ = __mustCol0_(send.idx0, '完了フラグ');
  const sキュー送信 = __mustCol0_(send.idx0, 'キュー送信');
  const s送信済履歴 = __mustCol0_(send.idx0, '送信済テンプレ履歴');
  const sキューテンプレ = __mustCol0_(send.idx0, 'キューテンプレ');

  const TPL_TSUTO = 'SO_TSUTO_COMPLETE';
  const TPL_WORK  = 'SO_WORK_COMPLETE';

  let queued = 0;
  let skipped = 0;

  const numRows = send.valuesAll.length - 1;
  if (numRows <= 0) return;

  const queueCol = [];
  const tplCol = [];

  for (let r = 1; r < send.valuesAll.length; r++) {
    const row = send.valuesAll[r];

    const appId = row[s申請番号];
    const notify = __toBool_(row[s発報対象]);
    const statusChanged = __toBool_(row[sステータス変更]);
    const doneFlag = Number(row[s完了フラグ] || 0);

    let nextQueue = false;
    let nextTpl = '';

    if (notify && statusChanged) {
      const curStatus = idToStatus.get(String(appId)) || '';
      const tpl = decideTemplate_(curStatus, doneFlag, TPL_TSUTO, TPL_WORK);

      if (tpl) {
        const sentHist = String(row[s送信済履歴] || '');
        if (!sentHist.includes(tpl)) {
          nextQueue = true;
          nextTpl = tpl;
          queued++;
        } else {
          skipped++;
        }
      } else {
        skipped++;
      }
    } else {
      skipped++;
    }

    queueCol.push([nextQueue]);
    tplCol.push([nextTpl]);
  }

  // 書き戻し：該当列だけ（ヘッダーの次行から）
  const headerRow1 = send.headerRowIndex0 + 1;  // 1-based
  const startRow = headerRow1 + 1;              // データ開始行

  send.sheet.getRange(startRow, sキュー送信 + 1, numRows, 1).setValues(queueCol);
  send.sheet.getRange(startRow, sキューテンプレ + 1, numRows, 1).setValues(tplCol);

  Logger.log(`[prepare_send_queue_demo20] queued=${queued}, skipped=${skipped}`);
}

/* =========================================================
 * 送信：run_send_demo20(dryRun)（v4そのまま）
 * ======================================================= */

function run_send_demo20(dryRun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const batchId = Utilities.getUuid();

  const send = __loadSheetByHeaderFlex_(ss, CFG_DEMO20.SHEET_SEND,
    ['申請番号', '発報対象', 'キュー送信', 'キューテンプレ', '送信済テンプレ履歴'],
    CFG_DEMO20.HEADER_SCAN_ROWS
  );

  // 必須
  const s申請番号 = __mustCol0_(send.idx0, '申請番号');
  const s発報対象 = __mustCol0_(send.idx0, '発報対象');
  const sキュー送信 = __mustCol0_(send.idx0, 'キュー送信');
  const sキューテンプレ = __mustCol0_(send.idx0, 'キューテンプレ');
  const s送信済履歴 = __mustCol0_(send.idx0, '送信済テンプレ履歴');

  // 任意（あれば更新）
  const sバッチID = __findCol0_(send.idx0, ['バッチID', '送信バッチID']);
  const s最終送信日時 = __findCol0_(send.idx0, ['最終送信日時']);
  const s最終送信ステータス = __findCol0_(send.idx0, ['最終送信ステータス']);
  const s送信結果 = __findCol0_(send.idx0, ['送信結果']);
  const s送信エラー = __findCol0_(send.idx0, ['送信エラー']);
  const sエラー回数 = __findCol0_(send.idx0, ['エラー回数']);
  const s最終エラー日時 = __findCol0_(send.idx0, ['最終エラー日時']);
  const s次回再試行日時 = __findCol0_(send.idx0, ['次回再試行日時']);

  // 送信対象ID Set
  const targetIds = new Set();
  for (let r = 1; r < send.valuesAll.length; r++) {
    const row = send.valuesAll[r];
    const appId = String(row[s申請番号] ?? '').trim();
    if (!appId) continue;

    if (!__toBool_(row[s発報対象])) continue;
    if (!__toBool_(row[sキュー送信])) continue;

    const tplKey = String(row[sキューテンプレ] ?? '').trim();
    if (!tplKey) continue;

    // 送信済履歴に入ってたらスキップ（再送防止）
    const hist = String(row[s送信済履歴] ?? '');
    if (hist.includes(tplKey)) continue;

    targetIds.add(appId);
  }

  if (targetIds.size === 0) {
    Logger.log(`[run_send_demo20] no targets. dryRun=${dryRun}`);
    return;
  }

  // 申請一覧の必要情報 Map（最強方式）
  const applyInfo = buildApplyInfoMapDemo20_(ss, targetIds);

  // テンプレ
  const templates = loadTemplatesDemo20_();

  // 更新（必要セルだけ）
  const updates = [];
  const headerRow1 = send.headerRowIndex0 + 1; // 1-based
  const dataStartRow = headerRow1 + 1;

  const setIf = (rowNum, col0, value) => {
    if (col0 === undefined) return;
    updates.push({ r: rowNum, c: col0 + 1, v: value });
  };

  let processed = 0;
  let ok = 0;
  let ng = 0;

  for (let r = 1; r < send.valuesAll.length; r++) {
    const row = send.valuesAll[r];
    const rowNum = dataStartRow + (r - 1);

    const appId = String(row[s申請番号] ?? '').trim();
    if (!appId) continue;

    if (!__toBool_(row[s発報対象])) continue;
    if (!__toBool_(row[sキュー送信])) continue;

    const tplKey = String(row[sキューテンプレ] ?? '').trim();
    if (!tplKey) continue;

    const hist = String(row[s送信済履歴] ?? '');
    if (hist.includes(tplKey)) continue;

    const info = applyInfo.get(appId) || { mail: '', company: '', person: '' };
    const mail = info.mail;

    // テンプレ存在チェック
    const tpl = templates.get(tplKey);
    if (!tpl) {
      const errMsg = `template not found: ${tplKey}`;
      setIf(rowNum, sバッチID, batchId);
      setIf(rowNum, s最終送信日時, now);
      setIf(rowNum, s最終送信ステータス, 'NG');
      setIf(rowNum, s送信結果, `NG tpl=${tplKey}`);
      setIf(rowNum, s送信エラー, errMsg);
      setIf(rowNum, s最終エラー日時, now);
      if (sエラー回数 !== undefined) {
        const cur = Number(row[sエラー回数] ?? 0);
        setIf(rowNum, sエラー回数, (Number.isFinite(cur) ? cur : 0) + 1);
      }
      if (s次回再試行日時 !== undefined) {
        setIf(rowNum, s次回再試行日時, new Date(now.getTime() + CFG_DEMO20.RETRY_MINUTES * 60 * 1000));
      }
      ng++; processed++;
      continue;
    }

    // 宛先が無い：NG
    if (!mail) {
      const errMsg = `mail address missing for appId=${appId}`;
      setIf(rowNum, sバッチID, batchId);
      setIf(rowNum, s最終送信日時, now);
      setIf(rowNum, s最終送信ステータス, 'NG');
      setIf(rowNum, s送信結果, `NG no-mail tpl=${tplKey}`);
      setIf(rowNum, s送信エラー, errMsg);
      setIf(rowNum, s最終エラー日時, now);
      if (sエラー回数 !== undefined) {
        const cur = Number(row[sエラー回数] ?? 0);
        setIf(rowNum, sエラー回数, (Number.isFinite(cur) ? cur : 0) + 1);
      }
      if (s次回再試行日時 !== undefined) {
        setIf(rowNum, s次回再試行日時, new Date(now.getTime() + CFG_DEMO20.RETRY_MINUTES * 60 * 1000));
      }
      ng++; processed++;
      continue;
    }

    // トークン
    const dict = {
      '申請番号': appId,
      '工事会社名': info.company,
      '工事会社担当者名': info.person,
      '工事会社メールアドレス': info.mail,
      '自社名': '合同会社コトブク', // ←これを追加（またはCFG/シートから取得）
    };

    const subject = replaceTokens_(tpl.subject, dict);
    const body = replaceTokens_(tpl.body, dict);

    try {
      if (dryRun) {
        setIf(rowNum, sバッチID, batchId);
        setIf(rowNum, s最終送信日時, now);
        setIf(rowNum, s最終送信ステータス, 'DRYRUN_OK');
        setIf(rowNum, s送信結果, `dryrun to=${mail} tpl=${tplKey}`);
        setIf(rowNum, s送信エラー, '');
        ok++; processed++;
        continue;
      }

      GmailApp.sendEmail(mail, subject, body);

      setIf(rowNum, sバッチID, batchId);
      setIf(rowNum, s最終送信日時, now);
      setIf(rowNum, s最終送信ステータス, 'SENT_OK');
      setIf(rowNum, s送信結果, `sent to=${mail} tpl=${tplKey}`);
      setIf(rowNum, s送信エラー, '');
      setIf(rowNum, s最終エラー日時, '');
      setIf(rowNum, s次回再試行日時, '');

      // 履歴追記
      const nextHist = mergeSentHistory_(hist, tplKey);
      updates.push({ r: rowNum, c: s送信済履歴 + 1, v: nextHist });

      // キューOFF
      updates.push({ r: rowNum, c: sキュー送信 + 1, v: false });
      updates.push({ r: rowNum, c: sキューテンプレ + 1, v: '' });

      ok++; processed++;
    } catch (e) {
      const errMsg = String(e && e.message ? e.message : e);

      setIf(rowNum, sバッチID, batchId);
      setIf(rowNum, s最終送信日時, now);
      setIf(rowNum, s最終送信ステータス, 'NG');
      setIf(rowNum, s送信結果, `NG to=${mail} tpl=${tplKey}`);
      setIf(rowNum, s送信エラー, errMsg);
      setIf(rowNum, s最終エラー日時, now);

      if (sエラー回数 !== undefined) {
        const cur = Number(row[sエラー回数] ?? 0);
        setIf(rowNum, sエラー回数, (Number.isFinite(cur) ? cur : 0) + 1);
      }
      if (s次回再試行日時 !== undefined) {
        setIf(rowNum, s次回再試行日時, new Date(now.getTime() + CFG_DEMO20.RETRY_MINUTES * 60 * 1000));
      }

      ng++; processed++;
    }
  }

  __applyCellUpdates_(send.sheet, updates);

  Logger.log(`[run_send_demo20] finished. dryRun=${dryRun}, processed=${processed}, ok=${ok}, ng=${ng}, batchId=${batchId}`);
}
