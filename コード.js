/**
 * Facebook コメント自動返信（MVP + 自動投稿検出版）
 * 目的:
 *  - ページの長期有効アクセストークンから直近の投稿を自動検出
 *  - コメントを取得→キーワード（部分一致）で判定→返信→ログ記録
 * 必要権限:
 *  - Facebook Graph API: pages_read_engagement, pages_manage_posts
 *  - GAS: UrlFetchApp, SpreadsheetApp, PropertiesService, HtmlService
 */

// ========== メニュー ==========
function onOpen() {
  ensureAutomationTriggers();
  createMenu();
}

/**
 * 時間主導トリガーを必要に応じて作成
 */
function ensureAutomationTriggers() {
  try {
    setupScheduledPostsTrigger();
    setupAutoReplyTrigger();
  } catch (error) {
    console.error('トリガー設定エラー:', error);
  }
}

/**
 * 予約投稿・自動返信トリガーを全て削除
 */
function resetAutomationTriggers() {
  try {
    const targetHandlers = new Set(['processScheduledPosts', 'fetchAndRespond']);
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      const handler = trigger.getHandlerFunction && trigger.getHandlerFunction();
      if (handler && targetHandlers.has(handler)) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  } catch (error) {
    console.error('トリガー削除エラー:', error);
    throw error;
  }
}


/**
 * メニューを作成する（手動実行用）
 */
function createMenu() {
  try {
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('ツール設定')
      .addItem('1) 初期化: シート生成', 'initSheets')
      .addSeparator()
      .addSubMenu(ui.createMenu('トークン管理')
        .addItem('トークン設定', 'setupFromManualInput')
        .addItem('トークン情報を確認', 'showTokenInfo')
        .addItem('トークンを更新', 'refreshTokenManually'))
      .addSubMenu(ui.createMenu('画像管理')
        .addItem('ImgBB APIキー設定', 'showImageUploadSettings')
        .addItem('複数画像一括アップロード', 'showBatchImageUploadDialog')
        .addItem('画像アップロードテスト', 'testImageUpload'))
      .addSeparator()
      .addItem('コメント取得→返信 実行', 'fetchAndRespond')
        .addItem('予約投稿（手動実行）', 'runScheduledPostsNow')
        .addItem('未返信に一括返信（直近12時間）', 'runAutoReplyForLast12hUnreplied')
        .addItem('シートを最新の状態に更新', 'updateSheetsToLatest')
        .addSeparator()
        .addItem('⚠️ 全てのシートを再構成', 'rebuildAllSheets')
      .addToUi();

    ui.createMenu('クイックセットアップ')
      .addItem('アカウント確認とトリガー初期化', 'runQuickSetup')
      .addToUi();
  } catch (error) {
    console.error('メニュー作成エラー:', error);
  }
}

/**
 * アカウント情報の確認とトリガー初期化をまとめて実行
 */
function runQuickSetup() {
  const ui = SpreadsheetApp.getUi();
  try {
    const tokenInfo = getTokenInfo();

    resetAutomationTriggers();
    ensureAutomationTriggers();

    const pageName = tokenInfo.pageName || '未設定';
    const pageId = tokenInfo.pageId || '未設定';
    const expiresAt = tokenInfo.expiresAt ? tokenInfo.expiresAt.toLocaleString('ja-JP') : '不明';
    const status = tokenInfo.isValid ? '有効' : '無効または期限切れ';

    ui.alert(
      'クイックセットアップが完了しました。\n\n' +
      `ページ名: ${pageName}\n` +
      `ページID: ${pageId}\n` +
      `トークン有効期限: ${expiresAt}\n` +
      `トークン状態: ${status}\n\n` +
      '予約投稿トリガー（5分間隔）と自動返信トリガー（30分間隔）を再作成しました。'
    );
  } catch (error) {
    ui.alert(`❌ クイックセットアップでエラーが発生しました: ${error && error.message ? error.message : error}`);
  }
}

/**
 * 直近12時間の未返信コメントに一括返信
 */
function runAutoReplyForLast12hUnreplied() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = replyUnrepliedCommentsLast12h();
    ui.alert(
      '✅ 未返信への一括返信が完了しました。\n' +
      `対象コメント: ${result.total} 件\n` +
      `返信: ${result.replied} 件\n` +
      `失敗: ${result.failed} 件\n` +
      `無視（条件不一致など）: ${result.skipped} 件`
    );
  } catch (e) {
    ui.alert(`❌ エラー: ${e && e.message ? e.message : e}`);
  }
}

/**
 * 直近12時間の未返信コメントを検出し、ルールに基づいて返信
 * @return {{total:number,replied:number,failed:number,skipped:number}}
 */
function replyUnrepliedCommentsLast12h() {
  const token = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
  if (!token) throw new Error('アクセストークンが未設定です。');

  const settings = getSettings();
  const postIds = getTargetPostIds(token, settings);
  if (postIds.length === 0) throw new Error('対象投稿がありません。');

  const rules = loadRules();
  const repliedSet = loadRepliedCommentIdsSet();
  const sinceMs = Date.now() - (12 * 60 * 60 * 1000);

  let total = 0;
  let replied = 0;
  let failed = 0;
  let skipped = 0;

  for (const postId of postIds) {
    const comments = fetchCommentsSince(postId, token, sinceMs, 200) || [];
    for (const c of comments) {
      const commentId = String(c.id || '').trim();
      if (!commentId) { skipped++; continue; }

      // 二重送信防止（ログでreplied済み）
      if (repliedSet.has(commentId)) { skipped++; continue; }

      // 念のため時刻再チェック
      const created = c.created_time ? new Date(c.created_time).getTime() : 0;
      if (!created || created < sinceMs) { skipped++; continue; }

      total++;
      const message = String(c.message || '');
      const name = c.from && c.from.name ? c.from.name : '';

      const rule = findMatchingRule(message, rules);
      if (!rule) { skipped++; continue; }

      const replyText = generateReply(rule.template, name);
      try {
        postReply(commentId, replyText, token);
        appendProcessed(commentId, c.created_time || '');
        logComment(postId, commentId, name, message, rule.keyword, replyText, 'replied', '');
        replied++;
        Utilities.sleep(300);
      } catch (err) {
        logComment(postId, commentId, name, message, rule.keyword, replyText, 'error', String(err && err.message ? err.message : err));
        failed++;
      }
    }
  }

  return { total, replied, failed, skipped };
}

// ========== 初期化 ==========
function initSheets() {
  // シートを初期化
  initializeSheets();
  
  // デフォルト設定を初期化
  initializeDefaultSettings();

  SpreadsheetApp.getUi().alert(
    '✅ 初期化が完了しました！\n\n' +
    '設定シートに以下の項目が追加されました：\n' +
    '• 基本設定（投稿ID、取得件数など）\n' +
    '• トークン関連（空欄で初期化）\n' +
    '• ルール関連（デフォルト値設定）\n' +
    '• システム設定（API設定など）\n\n' +
    '次の順で進めてください：\n' +
    '1) 「トークン管理」→「トークン設定」でページトークン設定\n' +
    '2) ルールシートにキーワードと自動返信内容を入力\n' +
    '3) 「コメント取得→返信 実行」で検証'
  );
}


// ========== メイン: 取得→判定→返信 ==========
function fetchAndRespond() {
  const token = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
  if (!token) throw new Error('アクセストークンが未設定です。メニューから「トークン管理」→「トークン設定」を実行してください。');

  const settings = getSettings();
  const fetchLimit = parseInt(settings.get('取得件数') || '50', 10);

  // 対象投稿IDを取得
  const postIds = getTargetPostIds(token, settings);
  if (postIds.length === 0) {
    throw new Error('対象投稿がありません。投稿IDを指定するか、自動検出=trueで直近投稿があることを確認してください。');
  }

  const rules = loadRules();
  const processedSet = loadProcessedIds();

  // 各投稿のコメントを処理
  for (const postId of postIds) {
    processComments(postId, token, fetchLimit, rules, processedSet);
  }
}

// ========== トークン管理 ==========

/**
 * 手動入力でトークン設定を行う（改良版）
 */
function setupFromManualInput() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 500px;">
      <h3>Facebook トークン設定</h3>
      <p style="margin-bottom: 20px;">Facebook Graph API Debugger などで取得した長期ユーザートークンを入力してください。</p>

      <div style="margin-bottom: 20px;">
        <label for="longToken" style="display: block; margin-bottom: 5px; font-weight: bold;">長期ユーザートークン</label>
        <textarea id="longToken" style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; height: 80px;" placeholder="EAAG..."></textarea>
      </div>

      <div style="text-align: center;">
        <button onclick="submitForm()" style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px;">設定実行</button>
        <button onclick="google.script.host.close()" style="background-color: #f44336; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">キャンセル</button>
      </div>
    </div>

    <script>
      function submitForm() {
        const longToken = document.getElementById('longToken').value.trim();

        if (!longToken) {
          alert('長期ユーザートークンを入力してください。');
          return;
        }

        google.script.run
          .withSuccessHandler(onSuccess)
          .withFailureHandler(onFailure)
          .processTokenInput(longToken);
      }

      function onSuccess(result) {
        if (result.success) {
          alert('✅ トークン設定が完了しました！\\n\\n' +
                'ページ名: ' + result.pageName + '\\n' +
                'ページID: ' + result.pageId + '\\n' +
                'トークン種別: ' + result.tokenType + '\\n' +
                '有効期限: ' + result.expiresAt + '\\n\\n' +
                '✅ 長期トークンを保存しました\\n' +
                '✅ 設定シートに保存完了');
          google.script.host.close();
        } else {
          alert('❌ エラー: ' + result.error + '\\n\\n' +
                '以下の点を確認してください：\\n' +
                '• 長期トークンが正しくコピーされているか\\n' +
                '• ページの管理者権限があるか\\n' +
                '• トークンが有効期限内か');
        }
      }

      function onFailure(error) {
        alert('❌ エラーが発生しました: ' + error.message);
      }
    </script>
  `)
  .setWidth(540)
  .setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Facebook トークン設定');
}

/**
 * HTMLダイアログから呼び出される処理関数
 */
function processTokenInput(longToken) {
  try {
    // トークン設定を実行
    const result = setupTokensFromLongToken(longToken);

    if (result.success) {
      // 設定シートに長期トークン情報を追記
      saveTokenToSettingsSheet(result);

      return {
        success: true,
        pageName: result.pageName,
        pageId: result.pageId,
        tokenType: result.tokenType,
        expiresAt: result.expiresAt ? result.expiresAt.toLocaleString('ja-JP') : '不明'
      };
    } else {
      return {
        success: false,
        error: result.error
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 設定シートにトークン情報を保存する
 */
function saveTokenToSettingsSheet(tokenInfo) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
    if (!sh) {
      console.error('設定シートが見つかりません');
      return;
    }
    
    // 既存のトークン設定を更新または追加
    upsertSetting('長期トークン', tokenInfo.accessToken || '', 'Facebook長期アクセストークン');
    upsertSetting('ページトークン', tokenInfo.pageAccessToken || '', 'Facebookページアクセストークン');
    upsertSetting('ページ名', tokenInfo.pageName || '', 'Facebookページ名');
    upsertSetting('ページID', tokenInfo.pageId || '', 'FacebookページID');
    upsertSetting('トークン有効期限', tokenInfo.expiresAt ? tokenInfo.expiresAt.toLocaleString('ja-JP') : '', 'トークンの有効期限');
    upsertSetting('最終更新日時', new Date().toLocaleString('ja-JP'), '最後にトークンを更新した日時');
    
    console.log('設定シートにトークン情報を保存しました');
  } catch (error) {
    console.error('設定シートへの保存でエラー:', error);
  }
}


/**
 * トークン情報を表示（デバッグ情報付き）
 */
function showTokenInfo() {
  const ui = SpreadsheetApp.getUi();
  const tokenInfo = getTokenInfo();
  
  // デバッグ情報を取得
  let debugInfo = '';
  try {
    const props = PropertiesService.getScriptProperties();
    const userToken = props.getProperty('USER_ACCESS_TOKEN');
    const expiresAt = props.getProperty('TOKEN_EXPIRES_AT');
    
    if (userToken) {
      const tokenDetails = getTokenInfoFromToken(userToken);
      debugInfo = `\n\n【デバッグ情報】\n` +
        `API有効期限: ${tokenDetails.expires_in ? Math.round(tokenDetails.expires_in / 86400) + '日' : '不明'}\n` +
        `保存された有効期限: ${expiresAt ? new Date(parseInt(expiresAt)).toLocaleString('ja-JP') : '不明'}`;
    }
  } catch (error) {
    debugInfo = `\n\n【デバッグ情報】\nエラー: ${error.message}`;
  }
  
  const message = 
    `現在のトークン情報:\n\n` +
    `ページ名: ${tokenInfo.pageName}\n` +
    `ページID: ${tokenInfo.pageId}\n` +
    `有効期限: ${tokenInfo.expiresAt ? tokenInfo.expiresAt.toLocaleString('ja-JP') : '不明'}\n` +
    `状態: ${tokenInfo.isValid ? '有効' : '無効または期限切れ'}` +
    debugInfo;
  
  ui.alert(message);
}

/**
 * 手動でトークンを更新
 */
function refreshTokenManually() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const success = refreshToken();
    if (success) {
      const tokenInfo = getTokenInfo();
      ui.alert(
        'トークンが正常に更新されました！\n\n' +
        `ページ名: ${tokenInfo.pageName}\n` +
        `有効期限: ${tokenInfo.expiresAt ? tokenInfo.expiresAt.toLocaleString('ja-JP') : '不明'}`
      );
  } else {
      ui.alert('トークンの更新に失敗しました。アプリ設定を確認してください。');
    }
  } catch (error) {
    ui.alert(`エラーが発生しました: ${error.message}`);
  }
}

// ========== シート管理 ==========

// （MVPの簡素化に伴い、非必須のユーティリティ関数は削除）

/**
 * 全てのシートを再構成する
 * シート1以外の全てのシートを削除して、テンプレート状態のシートを新規作成
 */
function rebuildAllSheets() {
  const ui = SpreadsheetApp.getUi();
  
  // 確認ダイアログ
  const response = ui.alert(
    '⚠️ 全てのシートを再構成',
    'この操作は以下のシートのデータを完全に削除します：\n\n' +
    '• 設定\n' +
    '• ルール\n' +
    '• ログ\n' +
    '• 処理済み\n' +
    '• 取得した投稿\n\n' +
    'シート1は保持されます。\n\n' +
    '本当に実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('操作がキャンセルされました。');
      return;
  }
  
  try {
    const spreadsheet = SpreadsheetApp.getActive();
    const sheets = spreadsheet.getSheets();
    
    // シート1以外の全てのシートを削除
    for (let i = sheets.length - 1; i >= 0; i--) {
      const sheet = sheets[i];
      // シート1（最初のシート）は削除しない
      if (i > 0) {
        spreadsheet.deleteSheet(sheet);
      }
    }
    
    // テンプレート状態のシートを新規作成
    initializeSheets();
    initializeDefaultSettings();
    
    ui.alert(
      '✅ シートの再構成が完了しました！\n\n' +
      '以下のシートが新規作成されました：\n' +
      '• 設定（デフォルト設定付き）\n' +
      '• ルール\n' +
      '• ログ\n' +
      '• 処理済み\n' +
      '• 取得した投稿\n\n' +
      'シート1は保持されています。'
    );
    
  } catch (error) {
    ui.alert(`エラーが発生しました: ${error.message}`);
  }
}

/**
 * 編集時のデフォルト補完（予約投稿シート）
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (!sh || sh.getName() !== SHEET.SCHEDULED) return;
    const row = e.range.getRow();
    if (row < 2) return; // ヘッダー除外

    // 対象行の値を取得
    const values = sh.getRange(row, 1, 1, 10).getValues()[0];
    let [enabled, body, dateVal, hourVal, minVal, tz, status] = values;

    const updates = [];

    // デフォルト: 有効
    if (!enabled) {
      enabled = '有効';
      updates.push({ col: 1, val: enabled });
    }

    // タイムゾーンは固定（Asia/Tokyo）
    if (tz !== 'Asia/Tokyo') {
      tz = 'Asia/Tokyo';
      updates.push({ col: 6, val: tz });
    }

    // デフォルト: 状態
    if (!status) {
      status = '予約中';
      updates.push({ col: 7, val: status });
    }

    // まとめて書き込み
    if (updates.length) {
      updates.forEach(u => sh.getRange(row, u.col).setValue(u.val));
    }
  } catch (err) {
    console.error('onEdit デフォルト補完エラー:', err);
  }
}
/**
 * シートを最新の状態に更新（非破壊）
 */
function updateSheetsToLatest() {
  try {
    // 構造的な更新（ルール/投稿/予約投稿）
    applySheetStructureUpdates();

    // デフォルト設定を補完
    initializeDefaultSettings();

    SpreadsheetApp.getUi().alert(
      '✅ シートを最新の状態に更新しました！\n\n' +
      '実施内容：\n' +
      '• ルール: ヘッダー名称の更新とプルダウン再適用\n' +
      '• 取得した投稿: シートの存在確認\n' +
      '• 予約投稿: シートの追加/確認とプルダウン適用\n' +
      '• 設定: デフォルト値の補完\n' +
      '• デザイン: ヘッダー配色/交互行/状態の色分けを適用'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ 更新に失敗しました: ${e && e.message ? e.message : e}`);
  }
}
/**
 * 予約投稿を手動実行（予約中かつ期限到来分）
 */
function runScheduledPostsNow() {
  try {
    const res = processScheduledPosts();
    const posted = res && typeof res.posted === 'number' ? res.posted : 0;
    const failed = res && typeof res.failed === 'number' ? res.failed : 0;
    SpreadsheetApp.getUi().alert(
      '手動実行が完了しました。\n' +
      `投稿: ${posted} 件 / 失敗: ${failed} 件`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ 手動実行でエラー: ${e && e.message ? e.message : e}`);
  }
}
