/**
 * 設定管理
 * 定数と設定値の管理を行う
 */

// シート名の定義
const SHEET = {
  SETTINGS: '設定',
  RULES: 'ルール',
  LOGS: 'ログ',
  PROCESSED: '処理済み',
  POSTS: '取得した投稿',
  SCHEDULED: '予約投稿'
};

// プロパティキーの定義
const PROP_KEYS = {
  PAGE_ACCESS_TOKEN: 'FB_PAGE_ACCESS_TOKEN',
  USER_ACCESS_TOKEN: 'FB_USER_ACCESS_TOKEN',
  PAGE_ID: 'FB_PAGE_ID',
  PAGE_NAME: 'FB_PAGE_NAME',
  TOKEN_EXPIRES_AT: 'FB_TOKEN_EXPIRES_AT',
  APP_ID: 'FB_APP_ID',
  APP_SECRET: 'FB_APP_SECRET'
};

// Facebook API設定
const FB = {
  BASE: 'https://graph.facebook.com/v23.0'
};

// デフォルト設定値
const DEFAULT_SETTINGS = {
  // 基本設定
  '投稿ID': '',
  '取得件数': '50',
  '自動検出': 'true',
  '検索日数': '3',
  '最大投稿数': '10',
  
  // トークン関連（空欄で初期化）
  '長期トークン': '',
  'ページトークン': '',
  'ページ名': '',
  'ページID': '',
  'トークン有効期限': '',
  '最終更新日時': '',
  'スクリプトID': '',
  'アプリID': '',
  'アプリシークレット': '',
  
  // ルール関連
  'デフォルトマッチタイプ': '部分一致',
  'デフォルト優先順位': '5',
  'デフォルト重み': '100',
  
  // システム設定
  'Facebook API バージョン': 'v23.0',
  'レート制限待機時間': '300',
  'ログ保持日数': '30',
  'エラー通知設定': 'false'
};

// 設定の説明
const SETTING_DESCRIPTIONS = {
  // 基本設定
  '投稿ID': 'カンマ区切りで投稿IDを入力（空なら自動検出が有効）',
  '取得件数': '1回のコメント取得件数（MVPでは1ページのみ）',
  '自動検出': '投稿IDが空なら直近投稿を自動収集',
  '検索日数': '自動収集する過去日数（整数）',
  '最大投稿数': '1回に対象とする最大投稿数（整数）',
  
  // トークン関連
  '長期トークン': 'Facebook長期アクセストークン（自動設定）',
  'ページトークン': 'Facebookページアクセストークン（自動設定）',
  'ページ名': 'Facebookページ名（自動設定）',
  'ページID': 'FacebookページID（自動設定）',
  'トークン有効期限': 'トークンの有効期限（自動設定）',
  '最終更新日時': '最後にトークンを更新した日時（自動設定）',
  'スクリプトID': 'このGoogle Apps ScriptのスクリプトID（自動設定）',
  'アプリID': 'FacebookアプリID（手動設定）',
  'アプリシークレット': 'Facebookアプリシークレット（手動設定）',
  
  // ルール関連
  'デフォルトマッチタイプ': '新規ルールのデフォルトマッチタイプ',
  'デフォルト優先順位': '新規ルールのデフォルト優先順位',
  'デフォルト重み': '新規ルールのデフォルト重み',
  
  // システム設定
  'Facebook API バージョン': '使用するFacebook Graph APIのバージョン',
  'レート制限待機時間': 'API呼び出し間の待機時間（ミリ秒）',
  'ログ保持日数': 'ログを保持する日数',
  'エラー通知設定': 'エラー発生時の通知設定（true/false）'
};

/**
 * 設定値を取得する
 * @return {Map} 設定のMap
 */
function getSettings() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  const last = sh.getLastRow();
  const map = new Map();
  if (last < 2) return map;
  const values = sh.getRange(2, 1, last - 1, 2).getValues(); // key, value
  values.forEach(([k, v]) => {
    if (k) map.set(String(k).trim(), String(v || '').trim());
  });
  return map;
}

/**
 * 設定値を更新または追加する
 * @param {string} key 設定キー
 * @param {string} value 設定値
 * @param {string} notes 説明
 */
function upsertSetting(key, value, notes) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  const last = sh.getLastRow();
  if (last >= 2) {
    const keys = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || '').trim());
    const idx = keys.findIndex(k => k === key);
    if (idx >= 0) {
      sh.getRange(2 + idx, 2).setValue(value);
      if (notes !== undefined) sh.getRange(2 + idx, 3).setValue(notes);
      return;
    }
  }
  sh.appendRow([key, value, notes || '']);
}

/**
 * 設定値が存在しない場合のみ追加する
 * @param {string} key 設定キー
 * @param {string} value 設定値
 * @param {string} notes 説明
 */
function upsertSettingIfMissing(key, value, notes) {
  const settings = getSettings();
  if (!settings.has(key)) upsertSetting(key, value, notes || '');
}

/**
 * デフォルト設定を初期化する
 */
function initializeDefaultSettings() {
  Object.keys(DEFAULT_SETTINGS).forEach(key => {
    upsertSettingIfMissing(key, DEFAULT_SETTINGS[key], SETTING_DESCRIPTIONS[key]);
  });

  const scriptId = typeof ScriptApp !== 'undefined' && ScriptApp.getScriptId ? ScriptApp.getScriptId() : '';
  if (scriptId) {
    upsertSetting('スクリプトID', scriptId, SETTING_DESCRIPTIONS['スクリプトID']);
  } else {
    upsertSettingIfMissing('スクリプトID', '', SETTING_DESCRIPTIONS['スクリプトID']);
  }
}
