/**
 * スプレッドシート管理
 * シートの作成、データの読み書きを管理する
 */

/**
 * シートを確実に作成する（破壊的）
 * @param {string} name シート名
 * @param {Array} headers ヘッダー配列
 */
function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
}

/**
 * シートが存在しない場合のみ作成する（非破壊的）
 * @param {string} name シート名
 * @param {Array} headers ヘッダー配列
 */
function ensureSheetIfMissing(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }
}

/**
 * ルールを読み込む
 * @return {Array} ルール配列 [{enabled, keyword, template, matchType, priority, weight}]
 */
function loadRules() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.RULES);
  const last = sh.getLastRow();
  if (last < 2) return [];
  
  // 列数を確認して適切に読み込み
  const colCount = sh.getLastColumn();
  const values = sh.getRange(2, 1, last - 1, colCount).getValues();
  
  return values.map(([en, kw, tp, mt, pr, wt]) => ({
    enabled: String(en).toLowerCase() === 'true' || String(en) === '有効',
    keyword: String(kw || '').trim(),
    template: String(tp || '').trim(),
    matchType: String(mt || '部分一致').trim(),
    priority: parseInt(pr) || 5,
    weight: parseInt(wt) || 100
  }));
}

/**
 * 処理済みコメントIDを読み込む
 * @return {Set} 処理済みコメントIDのSet
 */
function loadProcessedIds() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.PROCESSED);
  const last = sh.getLastRow();
  const set = new Set();
  if (last < 2) return set;
  const values = sh.getRange(2, 1, last - 1, 1).getValues().flat(); // comment_id
  values.forEach(v => { if (v) set.add(String(v).trim()); });
  return set;
}

/**
 * ログから「返信済み」のコメントID集合を作成
 * @return {Set<string>} repliedコメントIDのSet
 */
function loadRepliedCommentIdsSet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.LOGS);
  const set = new Set();
  if (!sh) return set;
  const last = sh.getLastRow();
  if (last < 2) return set;
  const values = sh.getRange(2, 1, last - 1, 9).getValues();
  // 列構成: 日時, 投稿ID, コメントID, 投稿者名, コメント内容, マッチキーワード, 返信内容, ステータス, エラー
  for (const row of values) {
    const commentId = String(row[2] || '').trim();
    const status = String(row[7] || '').trim();
    if (commentId && status === 'replied') set.add(commentId);
  }
  return set;
}

/**
 * 処理済みコメントIDを追加する
 * @param {string} commentId コメントID
 * @param {string} createdTime 作成日時
 */
function appendProcessed(commentId, createdTime) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.PROCESSED);
  sh.appendRow([commentId, createdTime || '']);
}

/**
 * ログを記録する
 * @param {string} postId 投稿ID
 * @param {string} commentId コメントID
 * @param {string} name コメント投稿者名
 * @param {string} message コメント内容
 * @param {string} matchedKeyword マッチしたキーワード
 * @param {string} replyText 返信内容
 * @param {string} status 処理ステータス
 * @param {string} error エラーメッセージ
 */
function logComment(postId, commentId, name, message, matchedKeyword, replyText, status, error) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.LOGS);
  sh.appendRow([
    new Date(), postId, commentId, name, message, matchedKeyword, replyText, status, error || ''
  ]);
}

/**
 * シートを初期化する
 */
function initializeSheets() {
  ensureSheet(SHEET.SETTINGS, ['キー', '値', '説明']);
  ensureSheet(SHEET.RULES, [
    '有効', 'キーワード', '自動返信内容', 'マッチタイプ', '優先順位', '重み'
  ]);
  ensureSheet(SHEET.LOGS, [
    '日時', '投稿ID', 'コメントID', '投稿者名', 'コメント内容',
    'マッチキーワード', '返信内容', 'ステータス', 'エラー'
  ]);
  ensureSheet(SHEET.PROCESSED, ['コメントID', '作成日時']);
  ensureSheet(SHEET.POSTS, ['投稿ID', 'URL', '作成日時']);
  ensureSheet(SHEET.SCHEDULED, [
    '有効', '投稿本文', '日付', '時', '分', 'タイムゾーン',
    '状態', '投稿ID', 'URL', 'エラー'
  ]);
  
  // ルールを初期投入（空の場合のみ）
  seedDefaultRulesIfEmpty();
  
  // ルールシートにプルダウンバリデーションを設定
  setupRuleSheetValidation();

  // 予約投稿シートのプルダウンバリデーション設定
  setupScheduledSheetValidation();

  // デザイン適用（ヘッダー/交互行/見やすさ調整）
  applyDesignStyles();
}

/**
 * ルールシートが空の場合、既定の5行を投入
 * 有効=有効、マッチタイプ=部分一致、優先順位=1..5、重み=100
 */
function seedDefaultRulesIfEmpty() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.RULES);
  if (!sh) return;
  const last = sh.getLastRow();
  if (last >= 2) return; // 既にデータあり

  // 設定からデフォルト値を取得
  const settings = getSettings();
  const defaultMatchType = settings.get('デフォルトマッチタイプ') || '部分一致';
  const defaultWeight = settings.get('デフォルト重み') || '100';

  const rows = [];
  for (let i = 1; i <= 5; i++) {
    rows.push(['有効', '', '', defaultMatchType, i, defaultWeight]);
  }
  sh.getRange(2, 1, rows.length, 6).setValues(rows);
}

/**
 * ルールシートにプルダウンバリデーションを設定する
 */
function setupRuleSheetValidation() {
  const ss = SpreadsheetApp.getActive();
  const rulesSheet = ss.getSheetByName(SHEET.RULES);
  if (!rulesSheet) return;

  // ヘッダー行(1行目) + データ行1000行分を確保
  const requiredTotalRows = 1001;
  const currentMaxRows = rulesSheet.getMaxRows();
  if (currentMaxRows < requiredTotalRows) {
    rulesSheet.insertRowsAfter(currentMaxRows, requiredTotalRows - currentMaxRows);
  }

  const numRows = requiredTotalRows - 1; // ヘッダー下の行数

  // 有効列（A列）のバリデーション
  const enabledRange = rulesSheet.getRange(2, 1, numRows, 1);
  const enabledRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['有効', '無効'], true)
    .setAllowInvalid(false)
    .setHelpText('有効または無効を選択してください')
    .build();
  enabledRange.setDataValidation(enabledRule);

  // マッチタイプ列（D列）のバリデーション
  const matchTypeRange = rulesSheet.getRange(2, 4, numRows, 1);
  const matchTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['完全一致', '部分一致', '前方一致', '後方一致'], true)
    .setAllowInvalid(false)
    .setHelpText('マッチタイプを選択してください')
    .build();
  matchTypeRange.setDataValidation(matchTypeRule);

  // 優先順位列（E列）のバリデーション
  const priorityRange = rulesSheet.getRange(2, 5, numRows, 1);
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'], true)
    .setAllowInvalid(false)
    .setHelpText('優先順位を選択してください（1が最優先）')
    .build();
  priorityRange.setDataValidation(priorityRule);

  // 重み列（F列）のバリデーション（代表値のプルダウン）
  const weightRange = rulesSheet.getRange(2, 6, numRows, 1);
  const weightOptions = ['100','90','80','70','60','50','40','30','20','10'];
  const weightRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(weightOptions, true)
    .setAllowInvalid(false)
    .setHelpText('重みを選択してください（デフォルト100。必要に応じて拡張可能）')
    .build();
  weightRange.setDataValidation(weightRule);

  console.log('ルールシートのプルダウンバリデーションを設定しました（ヘッダー下1000行に適用）');

  // 自動返信内容（C列）を広めにし、折り返しを有効化
  try { rulesSheet.setColumnWidth(3, 500); } catch (e) {}
  rulesSheet.getRange(2, 3, numRows, 1).setWrap(true);
}

/**
 * 予約投稿シートにプルダウンバリデーションを設定する
 */
function setupScheduledSheetValidation() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET.SCHEDULED);
  if (!sh) return;

  // ヘッダー行(1行目) + データ行1000
  const requiredTotalRows = 1001;
  const currentMaxRows = sh.getMaxRows();
  if (currentMaxRows < requiredTotalRows) {
    sh.insertRowsAfter(currentMaxRows, requiredTotalRows - currentMaxRows);
  }
  const numRows = requiredTotalRows - 1;

  // 有効
  const enabledRange = sh.getRange(2, 1, numRows, 1);
  const enabledRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['有効', '無効'], true)
    .setAllowInvalid(false)
    .setHelpText('有効/無効を選択')
    .build();
  enabledRange.setDataValidation(enabledRule);

  // 時（0-23）
  const hours = Array.from({ length: 24 }, (_, i) => String(i));
  sh.getRange(2, 4, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(hours, true).setAllowInvalid(false).setHelpText('0-23 を選択').build()
  );

  // 分（0,5,10,...,55）
  const minutes = Array.from({ length: 12 }, (_, i) => String(i * 5));
  sh.getRange(2, 5, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(minutes, true).setAllowInvalid(false).setHelpText('5分刻み を選択').build()
  );

  // タイムゾーン（Asia/Tokyo に固定）
  sh.getRange(2, 6, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Asia/Tokyo'], true).setAllowInvalid(false).build()
  );
  // 既定値を1000行分セット（固定化）
  const tzValues = Array.from({ length: numRows }, () => ['Asia/Tokyo']);
  sh.getRange(2, 6, numRows, 1).setValues(tzValues);

  // 状態
  sh.getRange(2, 7, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['予約中', '送信済み', '失敗', '無効'], true).setAllowInvalid(true).build()
  );

  // 日付（カレンダー入力許可）: データバリデーション + 表示形式
  sh.getRange(2, 3, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('日付を選択').build()
  );
  sh.getRange(2, 3, numRows, 1).setNumberFormat('yyyy/m/d');

  // 投稿本文列: 幅を広げて折り返し
  try { sh.setColumnWidth(2, 380); } catch (e) {}
  sh.getRange(2, 2, numRows, 1).setWrap(true);
}

/**
 * 予約投稿を処理する（時間到来分を投稿）
 */
function processScheduledPosts() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET.SCHEDULED);
  if (!sh) return;

  const last = sh.getLastRow();
  if (last < 2) return { posted: 0, failed: 0 };

  let token = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
  if (!token) throw new Error("ページアクセストークンが未設定です");

  // トークンの有効性をチェックし、必要に応じて更新
  try {
    if (!isTokenValid(token)) {
      console.log("トークンが無効または期限切れのため、自動更新を試行します...");
      const refreshSuccess = refreshToken();
      if (refreshSuccess) {
        token = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
        console.log("トークンが正常に更新されました");
      } else {
        throw new Error("トークンの自動更新に失敗しました。手動でトークンを更新してください。");
      }
    }
  } catch (error) {
    console.error("トークン有効性チェックでエラー:", error);
    throw new Error(`トークンの有効性確認に失敗しました: ${error.message}`);
  }
  const now = new Date();
  const rows = sh.getRange(2, 1, last - 1, 10).getValues();
  let posted = 0;
  let failed = 0;

  for (let i = 0; i < rows.length; i++) {
    const [enabled, body, dateVal, hourVal, minVal, tz, status, postId, url, err] = rows[i];

    if (String(enabled) !== '有効') continue;
    if (String(status) !== '予約中') continue;
    if (!body) continue;
    if (!dateVal || isNaN(new Date(dateVal).getTime())) continue;

    const hour = parseInt(hourVal, 10);
    const minute = parseInt(minVal, 10);
    if (isNaN(hour) || isNaN(minute)) continue;

    // 予約日時（スクリプトのタイムゾーン前提）
    const scheduled = new Date(dateVal);
    scheduled.setHours(hour, minute, 0, 0);

    if (now.getTime() >= scheduled.getTime()) {
      try {
        const result = createPagePost(String(body), token);
        // 書き戻し
        sh.getRange(2 + i, 7, 1, 4).setValues([[
          '送信済み',
          result.postId || '',
          result.permalink || '',
          ''
        ]]);
        posted++;
      } catch (e) {
        sh.getRange(2 + i, 7, 1, 4).setValues([[
          '失敗',
          '',
          '',
          String(e && e.message ? e.message : e)
        ]]);
        failed++;
      }
      Utilities.sleep(300);
    }
  }
  return { posted, failed };
}

/**
 * 指定ハンドラーの時間主導トリガーを分単位で作成（存在しない場合のみ）
 * @param {string} handlerName
 * @param {number} minutes
 */
function ensureMinutesTrigger(handlerName, minutes) {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === handlerName);
  if (!exists) {
    ScriptApp.newTrigger(handlerName).timeBased().everyMinutes(minutes).create();
  }
}

/**
 * 予約投稿の時間主導トリガーを作成（5分毎）
 */
function setupScheduledPostsTrigger() {
  ensureMinutesTrigger('processScheduledPosts', 5);
}

/**
 * コメント自動返信の時間主導トリガーを作成（30分毎）
 */
function setupAutoReplyTrigger() {
  ensureMinutesTrigger('fetchAndRespond', 30);
}


/**
 * シートをアップデートする（非破壊的）
 */
function applySheetStructureUpdates() {
  // 旧シート名からの移行（投稿 -> 取得した投稿）
  const ss = SpreadsheetApp.getActive();
  const legacyPosts = ss.getSheetByName('投稿');
  if (legacyPosts && !ss.getSheetByName(SHEET.POSTS)) {
    legacyPosts.setName(SHEET.POSTS);
  }

  ensureSheetIfMissing(SHEET.POSTS, ['投稿ID', 'URL', '作成日時']);

  // 予約投稿シートを追加/更新（非破壊）
  ensureSheetIfMissing(SHEET.SCHEDULED, [
    '有効', '投稿本文', '日付', '時', '分', 'タイムゾーン',
    '状態', '投稿ID', 'URL', 'エラー'
  ]);

  // 古い「作成日時」列があれば削除し、ヘッダーを最新化
  const scheduledSheet = ss.getSheetByName(SHEET.SCHEDULED);
  if (scheduledSheet) {
    const headerVals = scheduledSheet.getRange(1, 1, 1, scheduledSheet.getLastColumn()).getValues()[0];
    const oldIndex = headerVals.indexOf('作成日時');
    if (oldIndex >= 0) {
      scheduledSheet.deleteColumn(oldIndex + 1);
    }
    const schedHeaders = ['有効', '投稿本文', '日付', '時', '分', 'タイムゾーン', '状態', '投稿ID', 'URL', 'エラー'];
    scheduledSheet.getRange(1, 1, 1, schedHeaders.length).setValues([schedHeaders]);
  }

  // 既存のルールシートをアップデート
  const rulesSheet = ss.getSheetByName(SHEET.RULES);
  if (rulesSheet) {
    // ヘッダーを更新（名称変更含む）
    const newHeaders = ['有効', 'キーワード', '自動返信内容', 'マッチタイプ', '優先順位', '重み'];
    rulesSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);

    // 既存列が不足している場合は追加
    const currentHeaderCount = rulesSheet.getLastColumn();
    if (currentHeaderCount < newHeaders.length) {
      const lastCol = currentHeaderCount;
      const addCount = newHeaders.length - currentHeaderCount;
      rulesSheet.getRange(1, lastCol + 1, 1, addCount).setValues([newHeaders.slice(lastCol)]);
    }

    // プルダウンバリデーションを再適用
    setupRuleSheetValidation();
  }

  // 予約投稿シートのバリデーションを再適用
  setupScheduledSheetValidation();

  // デザイン適用（ヘッダー/交互行/見やすさ調整）
  applyDesignStyles();
}

/**
 * デザイン適用（2025モダン配色）
 * - ヘッダー: 背景#F3F4F6, 太字
 * - 交互行: 白/#F9FAFB のバンディング
 * - 個別調整: ルールC列ワイド+折り返し、予約投稿B列ワイド+折り返し、状態の条件付き書式
 */
function applyDesignStyles() {
  const ss = SpreadsheetApp.getActive();
  const headerBg = '#F3F4F6';
  const band1 = '#FFFFFF';
  const band2 = '#F9FAFB';

  const targetSheets = [
    SHEET.SETTINGS,
    SHEET.RULES,
    SHEET.LOGS,
    SHEET.PROCESSED,
    SHEET.POSTS,
    SHEET.SCHEDULED
  ];

  targetSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastCol = Math.max(1, sh.getLastColumn());
    const headerRange = sh.getRange(1, 1, 1, lastCol);
    headerRange.setBackground(headerBg).setFontWeight('bold');

    // 既存バンディングを削除して再適用
    try {
      (sh.getBandings() || []).forEach(b => b.remove());
    } catch (e) {}

    const maxRows = Math.max(2, sh.getMaxRows());
    const bandRange = sh.getRange(1, 1, Math.min(1001, maxRows), lastCol);
    const band = bandRange.applyRowBanding();
    try { band.setHeaderRowColor(headerBg); } catch (e) {}
    try { band.setFirstBandColor(band1); } catch (e) {}
    try { band.setSecondBandColor(band2); } catch (e) {}

    // 全セルの折り返し（ヘッダー除く）
    const wrapRows = Math.max(0, Math.min(1000, maxRows - 1));
    if (wrapRows > 0) {
      sh.getRange(2, 1, wrapRows, lastCol).setWrap(true);
    }
  });

  // ルール: 自動返信内容（C列）を広めに、折り返し
  const rulesSheet = ss.getSheetByName(SHEET.RULES);
  if (rulesSheet) {
    try { rulesSheet.setColumnWidth(3, 500); } catch (e) {}
    const numRows = Math.max(1, Math.min(1000, rulesSheet.getMaxRows() - 1));
    if (numRows > 0) rulesSheet.getRange(2, 3, numRows, 1).setWrap(true);
  }

  // 予約投稿: 投稿本文（B列）を広めに、折り返し、状態の条件付き書式
  const sched = ss.getSheetByName(SHEET.SCHEDULED);
  if (sched) {
    try { sched.setColumnWidth(2, 380); } catch (e) {}
    const numRows = Math.max(1, Math.min(1000, sched.getMaxRows() - 1));
    if (numRows > 0) sched.getRange(2, 2, numRows, 1).setWrap(true);

    // 条件付き書式（状態: G列）
    const startRow = 2;
    const statusRange = sched.getRange(startRow, 7, Math.max(0, sched.getMaxRows() - (startRow - 1)), 1);
    const rules = [];
    const cf = SpreadsheetApp.newConditionalFormatRule;
    rules.push(cf().whenTextContains('予約中').setBackground('#FEF3C7').setFontColor('#92400E').setRanges([statusRange]).build());
    rules.push(cf().whenTextContains('送信済み').setBackground('#DCFCE7').setFontColor('#065F46').setRanges([statusRange]).build());
    rules.push(cf().whenTextContains('失敗').setBackground('#FEE2E2').setFontColor('#991B1B').setRanges([statusRange]).build());
    rules.push(cf().whenTextContains('無効').setBackground('#F3F4F6').setFontColor('#374151').setRanges([statusRange]).build());

    // 既存ルールに追加する形で適用（上書きで良い場合は setConditionalFormatRules(rules)）
    try {
      const existing = sched.getConditionalFormatRules() || [];
      // 旧状態ルールを除去（G列対象）
      const filtered = existing.filter(r => {
        const rngs = r.getRanges && r.getRanges();
        if (!rngs || !rngs.length) return true;
        return !rngs.some(rg => rg.getColumn() === 7);
      });
      sched.setConditionalFormatRules(filtered.concat(rules));
    } catch (e) {
      // フォールバック: 単独適用
      sched.setConditionalFormatRules(rules);
    }
  }

  // 設定: 列幅に余裕を持たせる（キー/値/説明）
  const settingsSheet = ss.getSheetByName(SHEET.SETTINGS);
  if (settingsSheet) {
    try {
      settingsSheet.setColumnWidth(1, 220); // キー
      settingsSheet.setColumnWidth(2, 320); // 値
      settingsSheet.setColumnWidth(3, 460); // 説明
    } catch (e) {}
  }
}
