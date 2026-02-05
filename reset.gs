/**
 * DEMO20: 送信管理_DEMO20 の「送信系カラム」を一括リセット（DRYRUN用）
 * - 列位置ではなく「ヘッダー名」で処理する（列ズレ耐性）
 * - 申請番号〜発報対象(H)は触らない
 * - ログID(I)以降の送信・エラー関連を初期化
 */
function reset_demo20_dryrun_v2() {
  const SHEET_NAME = '送信管理_DEMO20';

  // リセット対象（ヘッダー名）
  const RESET_COLS = [
    'ログID',
    'キュー送信',
    '送信済テンプレ履歴',
    'キューテンプレ',
    'バッチID',
    '最終送信日時',
    '最終送信ステータス',
    '送信結果',
    '送信エラー',
    'エラー回数',
    '最終エラー日時',
    '次回再試行日時',
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    Logger.log('データ行がないため処理不要');
    return;
  }

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = buildHeaderIndex_(header);

  // 存在チェック（無いときに静かに壊れるのを防ぐ）
  const missing = RESET_COLS.filter(name => !(name in idx));
  if (missing.length) {
    throw new Error(
      `ヘッダーが見つかりません: ${missing.join(', ')}\n` +
      `（シート1行目の見出しと、コード内ヘッダー名が完全一致しているか確認してください）`
    );
  }

  const numRows = lastRow - 1;

  // 一括で値を作る（行×列）
  // - チェックボックス列: FALSE
  // - それ以外: 空文字（""）
  const values = Array.from({ length: numRows }, () =>
    RESET_COLS.map(name => (name === 'キュー送信' ? false : ''))
  );

  // まとめて書き込み（列が飛び飛びなので、列ごとに書く）
  RESET_COLS.forEach((name, i) => {
    const col = idx[name] + 1; // 0-based -> 1-based
    const colValues = values.map(r => [r[i]]);
    sh.getRange(2, col, numRows, 1).setValues(colValues);
  });

  Logger.log(`reset_demo20_dryrun 完了: ${numRows} 行`);
}

/** ヘッダー行を {ヘッダー名: index(0-based)} にする */
function buildHeaderIndex_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i;
  });
  return map;
}

