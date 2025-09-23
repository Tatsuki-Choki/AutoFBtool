/**
 * トークン管理（/debug_token エンドポイント対応版）
 * Facebook トークンの取得、交換、更新を管理する
 * 短期アクセストークンは使用しない
 */

// Facebook Graph API エンドポイント
const FB_TOKEN_ENDPOINTS = {
  EXCHANGE_TOKEN: 'https://graph.facebook.com/v23.0/oauth/access_token',
  GET_PAGE_TOKEN: 'https://graph.facebook.com/v23.0/me/accounts',
  REFRESH_TOKEN: 'https://graph.facebook.com/v23.0/oauth/access_token',
  DEBUG_TOKEN: 'https://graph.facebook.com/v23.0/debug_token'
};

/**
 * /debug_token エンドポイントを使用してトークン情報を取得する
 * @param {string} inputToken 確認したいアクセストークン
 * @param {string} accessToken アプリトークン（{app-id}|{app-secret}形式）またはユーザートークン
 * @return {Object} 詳細なトークン情報
 */
function getTokenInfoFromDebugToken(inputToken, accessToken = null) {
  try {
    // アクセストークンが指定されていない場合は、入力トークン自体を使用
    const debugAccessToken = accessToken || inputToken;
    
    const url = `${FB_TOKEN_ENDPOINTS.DEBUG_TOKEN}` +
      `?input_token=${encodeURIComponent(inputToken)}` +
      `&access_token=${encodeURIComponent(debugAccessToken)}`;

    console.log('デバッグトークンURL:', url.replace(debugAccessToken, '***TOKEN***'));

    const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const code = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`デバッグトークンレスポンス: ${code}`);
    console.log('レスポンス内容:', responseText);
    
    if (code < 200 || code >= 300) {
      if (code === 400) {
        throw new Error('トークンが無効です。正しいアクセストークンを入力してください。');
      } else if (code === 401) {
        throw new Error('トークンの認証に失敗しました。トークンが期限切れまたは権限が不足しています。');
      } else {
        throw new Error(`トークン情報の取得に失敗しました: ${code} ${responseText}`);
      }
    }

    const result = JSON.parse(responseText);
    if (!result.data) {
      throw new Error('トークン情報の解析に失敗しました');
    }

    const data = result.data;
    console.log('取得したデバッグトークン情報:', data);

    return {
      // 基本情報
      appId: data.app_id,
      type: data.type,
      application: data.application,
      isValid: data.is_valid,
      userId: data.user_id,
      
      // 有効期限情報
      expiresAt: data.expires_at, // UNIXタイムスタンプ（0の場合は無期限）
      dataAccessExpiresAt: data.data_access_expires_at,
      
      // 権限情報
      scopes: data.scopes || [],
      
      // 計算された値
      expiresIn: data.expires_at > 0 ? Math.max(0, data.expires_at - Math.floor(Date.now() / 1000)) : 0,
      isLongLived: data.expires_at === 0 || (data.expires_at > 0 && (data.expires_at - Math.floor(Date.now() / 1000)) > 24 * 60 * 60),
      
      // レガシー対応
      expires_in: data.expires_at > 0 ? Math.max(0, data.expires_at - Math.floor(Date.now() / 1000)) : 0,
      app_id: data.app_id
    };
  } catch (error) {
    console.error('デバッグトークン取得でエラー:', error);
    throw error;
  }
}

/**
 * アプリトークンを生成する
 * @param {string} appId アプリID
 * @param {string} appSecret アプリシークレット
 * @return {string} アプリトークン（{app-id}|{app-secret}形式）
 */
function createAppToken(appId, appSecret) {
  return `${appId}|${appSecret}`;
}

/**
 * トークン情報を取得する（/debug_token対応版）
 * @param {string} token アクセストークン
 * @param {string} appId アプリID（オプション）
 * @param {string} appSecret アプリシークレット（オプション）
 * @return {Object} トークン情報
 */
function getTokenInfoFromToken(token, appId = null, appSecret = null) {
  try {
    let accessToken = token;
    
    // アプリIDとシークレットが提供されている場合はアプリトークンを使用
    if (appId && appSecret) {
      accessToken = createAppToken(appId, appSecret);
    }
    
    return getTokenInfoFromDebugToken(token, accessToken);
  } catch (error) {
    console.error('トークン情報取得でエラー:', error);
    
    // フォールバックとして旧エンドポイントを試行
    try {
      console.log('フォールバック: /oauth/access_token_info を使用');
      const url = `${FB.BASE}/oauth/access_token_info?access_token=${encodeURIComponent(token)}`;
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
      const code = response.getResponseCode();
      
      if (code >= 200 && code < 300) {
        const data = JSON.parse(response.getContentText());
        return {
          appId: data.app_id,
          type: data.type || 'USER',
          isValid: true,
          expiresAt: 0, // 旧エンドポイントでは詳細な期限情報は取得できない
          expiresIn: data.expires_in || 0,
          expires_in: data.expires_in || 0,
          app_id: data.app_id,
          scopes: [],
          isLongLived: (data.expires_in || 0) > 24 * 60 * 60
        };
      }
    } catch (fallbackError) {
      console.error('フォールバックも失敗:', fallbackError);
    }
    
    throw error;
  }
}

/**
 * トークンの有効期限をチェックする（改良版）
 * @param {string} token 確認するトークン（オプション）
 * @return {boolean} トークンが有効かどうか
 */
function isTokenValid(token = null) {
  try {
    // トークンが指定されている場合は、APIで直接確認
    if (token) {
      const tokenInfo = getTokenInfoFromToken(token);
      console.log('APIによるトークン有効性確認:', tokenInfo);
      
      // is_validフィールドで確認
      if (typeof tokenInfo.isValid === 'boolean') {
        return tokenInfo.isValid;
      }
      
      // 有効期限で確認
      if (tokenInfo.expiresAt > 0) {
        const now = Math.floor(Date.now() / 1000);
        const bufferTime = 24 * 60 * 60; // 24時間のバッファ
        return (tokenInfo.expiresAt - now) > bufferTime;
      }
      
      // 無期限トークンの場合
      return true;
    }
    
    // 保存された有効期限で確認（従来の方法）
    const expiresAt = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.TOKEN_EXPIRES_AT);
    if (!expiresAt) return false;
    
    const now = new Date().getTime();
    const expires = parseInt(expiresAt, 10);
    
    // 1日前に更新するよう余裕を持たせる
    const bufferTime = 24 * 60 * 60 * 1000; // 24時間
    return (expires - now) > bufferTime;
  } catch (error) {
    console.error('トークン有効性確認でエラー:', error);
    return false;
  }
}

/**
 * 詳細なトークン情報を表示する
 * @param {string} token 確認するトークン
 * @return {Object} 詳細情報
 */
function getDetailedTokenInfo(token = null) {
  try {
    const targetToken = token || PropertiesService.getScriptProperties().getProperty(PROP_KEYS.USER_ACCESS_TOKEN);
    if (!targetToken) {
      throw new Error('トークンが設定されていません');
    }
    
    const appId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.APP_ID);
    const appSecret = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.APP_SECRET);
    
    const tokenInfo = getTokenInfoFromToken(targetToken, appId, appSecret);
    
    return {
      // 基本情報
      アプリID: tokenInfo.appId,
      トークンタイプ: tokenInfo.type,
      アプリケーション名: tokenInfo.application || '不明',
      ユーザーID: tokenInfo.userId || '不明',
      
      // 有効性
      有効状態: tokenInfo.isValid ? '有効' : '無効',
      長期トークン: tokenInfo.isLongLived ? 'はい' : 'いいえ',
      
      // 有効期限
      有効期限: tokenInfo.expiresAt > 0 ? 
        new Date(tokenInfo.expiresAt * 1000).toLocaleString('ja-JP') : '無期限',
      残り時間: tokenInfo.expiresIn > 0 ? 
        `${Math.floor(tokenInfo.expiresIn / 86400)}日 ${Math.floor((tokenInfo.expiresIn % 86400) / 3600)}時間` : '無期限',
      データアクセス期限: tokenInfo.dataAccessExpiresAt > 0 ? 
        new Date(tokenInfo.dataAccessExpiresAt * 1000).toLocaleString('ja-JP') : '不明',
      
      // 権限
      権限スコープ: tokenInfo.scopes.join(', ') || '不明',
      
      // 生データ
      raw: tokenInfo
    };
  } catch (error) {
    console.error('詳細トークン情報取得でエラー:', error);
    return {
      エラー: error.message,
      raw: null
    };
  }
}

/**
 * ページトークンを取得する
 * @param {string} userToken ユーザートークン
 * @return {Object} ページ情報とトークン
 */
function getPageTokens(userToken) {
  const url = `${FB_TOKEN_ENDPOINTS.GET_PAGE_TOKEN}` +
    `?access_token=${encodeURIComponent(userToken)}`;

  const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = response.getResponseCode();
  
  if (code < 200 || code >= 300) {
    throw new Error(`ページトークン取得に失敗しました: ${code} ${response.getContentText()}`);
  }

  const data = JSON.parse(response.getContentText());
  if (!data.data || !Array.isArray(data.data)) {
    throw new Error('ページ情報の取得に失敗しました');
  }

  return data.data;
}

/**
 * トークンを更新する（改良版）
 * @return {boolean} 更新が成功したかどうか
 */
function refreshToken() {
  try {
    const userToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.USER_ACCESS_TOKEN);
    const appId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.APP_ID);
    const appSecret = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.APP_SECRET);
    
    console.log('トークン更新を開始します...');
    console.log('ユーザートークン:', userToken ? userToken.substring(0, 20) + '...' : '未設定');
    console.log('アプリID:', appId || '未設定');
    console.log('アプリシークレット:', appSecret ? '設定済み' : '未設定');
    
    if (!userToken) {
      throw new Error('ユーザートークンが設定されていません');
    }
    
    if (!appId || !appSecret) {
      console.warn('アプリIDまたはアプリシークレットが設定されていません。既存のトークンを使用します。');
      // アプリ情報が設定されていない場合は、既存のトークンをそのまま使用
      const pages = getPageTokens(userToken);
      if (pages.length === 0) {
        throw new Error('管理可能なページが見つかりません');
      }
      
      const page = pages[0];
      console.log('既存のトークンを使用してページ情報を更新します');
      
      // 詳細トークン情報を取得して有効期限を正しく計算
      let expiresAtMs = null;
      try {
        const info = getTokenInfoFromToken(userToken, appId, appSecret);
        if (info && info.dataAccessExpiresAt > 0) {
          expiresAtMs = info.dataAccessExpiresAt * 1000;
        } else if (info && info.expiresAt > 0) {
          expiresAtMs = info.expiresAt * 1000;
        }
      } catch (e) {
        console.warn('デバッグトークン情報の取得に失敗（既存トークン分）:', e && e.message ? e.message : e);
      }

      PropertiesService.getScriptProperties().setProperties({
        [PROP_KEYS.PAGE_ACCESS_TOKEN]: page.access_token,
        [PROP_KEYS.PAGE_ID]: page.id,
        [PROP_KEYS.PAGE_NAME]: page.name,
        [PROP_KEYS.TOKEN_EXPIRES_AT]: expiresAtMs ? String(expiresAtMs) : (PropertiesService.getScriptProperties().getProperty(PROP_KEYS.TOKEN_EXPIRES_AT) || '')
      });
      
      // 設定シートを更新
      try {
        if (page && page.id) upsertSetting('ページID', page.id, 'FacebookページID（自動設定）');
        if (page && page.name) upsertSetting('ページ名', page.name, 'Facebookページ名（自動設定）');
        upsertSetting('ページトークン', page.access_token || '', 'Facebookページアクセストークン（自動設定）');
        if (expiresAtMs) upsertSetting('トークン有効期限', new Date(expiresAtMs).toLocaleString('ja-JP'), 'トークンの有効期限（自動設定）');
        upsertSetting('最終更新日時', new Date().toLocaleString('ja-JP'), '最後にトークンを更新した日時（自動設定）');
      } catch (e) {
        console.warn('設定シート更新に失敗（既存トークン分）:', e && e.message ? e.message : e);
      }

      return true;
    }

    // 長期トークンを再取得
    const tokenInfo = exchangeToLongLivedTokenWithCredentials(userToken, appId, appSecret);
    
    // ページトークンを再取得
    const pages = getPageTokens(tokenInfo.accessToken);
    if (pages.length === 0) {
      throw new Error('管理可能なページが見つかりません');
    }

    // 最初のページを使用（複数ページの場合は選択機能を追加可能）
    const page = pages[0];
    
    // /debug_token から正しい有効期限を取得（data_access_expires_at 優先）
    let expiresAtMs = null;
    try {
      const info = getTokenInfoFromToken(tokenInfo.accessToken, appId, appSecret);
      if (info && info.dataAccessExpiresAt > 0) {
        expiresAtMs = info.dataAccessExpiresAt * 1000;
      } else if (info && info.expiresAt > 0) {
        expiresAtMs = info.expiresAt * 1000;
      }
    } catch (e) {
      console.warn('デバッグトークン情報の取得に失敗（更新後）:', e && e.message ? e.message : e);
    }
    if (!expiresAtMs && tokenInfo.expiresIn) {
      expiresAtMs = new Date().getTime() + (tokenInfo.expiresIn * 1000);
    }

    // プロパティを更新
    PropertiesService.getScriptProperties().setProperties({
      [PROP_KEYS.USER_ACCESS_TOKEN]: tokenInfo.accessToken,
      [PROP_KEYS.PAGE_ACCESS_TOKEN]: page.access_token,
      [PROP_KEYS.PAGE_ID]: page.id,
      [PROP_KEYS.PAGE_NAME]: page.name,
      [PROP_KEYS.TOKEN_EXPIRES_AT]: expiresAtMs ? String(expiresAtMs) : ''
    });

    // 設定シートを更新
    try {
      upsertSetting('長期トークン', tokenInfo.accessToken || '', 'Facebook長期アクセストークン（自動設定）');
      upsertSetting('ページトークン', page.access_token || '', 'Facebookページアクセストークン（自動設定）');
      if (page && page.id) upsertSetting('ページID', page.id, 'FacebookページID（自動設定）');
      if (page && page.name) upsertSetting('ページ名', page.name, 'Facebookページ名（自動設定）');
      if (expiresAtMs) upsertSetting('トークン有効期限', new Date(expiresAtMs).toLocaleString('ja-JP'), 'トークンの有効期限（自動設定）');
      upsertSetting('最終更新日時', new Date().toLocaleString('ja-JP'), '最後にトークンを更新した日時（自動設定）');
    } catch (e) {
      console.warn('設定シート更新に失敗（更新後）:', e && e.message ? e.message : e);
    }

    console.log('トークン更新が完了しました');
    return true;
  } catch (error) {
    console.error('トークン更新エラー:', error);
    return false;
  }
}

/**
 * アプリ設定を保存する
 * @param {string} appId アプリID
 * @param {string} appSecret アプリシークレット
 */
function saveAppSettings(appId, appSecret) {
  PropertiesService.getScriptProperties().setProperties({
    [PROP_KEYS.APP_ID]: appId,
    [PROP_KEYS.APP_SECRET]: appSecret
  });
}

/**
 * 手動入力でトークン設定を行う（長期トークン・改良版）
 * @param {string} longLivedToken 長期ユーザートークン
 * @return {Object} 設定結果
 */
function setupTokensFromLongToken(longLivedToken) {
  try {
    if (!longLivedToken) {
      throw new Error('長期ユーザートークンが未入力です。');
    }

    // トークン情報を取得（/debug_token使用）
    const tokenInfo = getTokenInfoFromToken(longLivedToken);
    console.log('入力されたユーザートークン情報:', tokenInfo);

    // ページトークンを取得
    const pages = getPageTokens(longLivedToken);
    if (pages.length === 0) {
      throw new Error('管理可能なページが見つかりません。ページの管理者権限があることを確認してください。');
    }

    // 最初のページを使用
    const page = pages[0];
    console.log('選択されたページ:', page);

    // 有効期限を計算（改良版）
    let expiresIn = tokenInfo.expiresIn;
    let expiresAtMs;
    
    if (tokenInfo.expiresAt > 0) {
      // APIから取得した実際の有効期限を使用
      expiresAtMs = tokenInfo.expiresAt * 1000;
      console.log(`APIから取得した有効期限: ${new Date(expiresAtMs).toLocaleString('ja-JP')}`);
    } else if (expiresIn > 0) {
      // 相対的な有効期限から計算
      expiresAtMs = new Date().getTime() + (expiresIn * 1000);
      console.log(`計算された有効期限: ${new Date(expiresAtMs).toLocaleString('ja-JP')}`);
    } else {
      // 無期限または不明な場合
      console.log('無期限トークンまたは有効期限が不明です。');
      expiresIn = 60 * 24 * 60 * 60; // 60日間をデフォルト設定
      expiresAtMs = new Date().getTime() + (expiresIn * 1000);
    }

    const expiresAt = new Date(expiresAtMs);
    console.log(`保存する有効期限: ${expiresAt.toLocaleString('ja-JP')}`);

    // すべてのトークン情報を保存
    PropertiesService.getScriptProperties().setProperties({
      [PROP_KEYS.USER_ACCESS_TOKEN]: longLivedToken,
      [PROP_KEYS.PAGE_ACCESS_TOKEN]: page.access_token,
      [PROP_KEYS.PAGE_ID]: page.id,
      [PROP_KEYS.PAGE_NAME]: page.name,
      [PROP_KEYS.TOKEN_EXPIRES_AT]: expiresAtMs.toString(),
      [PROP_KEYS.APP_ID]: tokenInfo.appId || '',
      [PROP_KEYS.APP_SECRET]: '' // シークレットは保存されない
    });

    const tokenType = tokenInfo.isLongLived ? '長期' : '短期';

    return {
      success: true,
      pageName: page.name,
      pageId: page.id,
      pageAccessToken: page.access_token,
      accessToken: longLivedToken,
      expiresAt,
      tokenType,
      appIdFound: true,
      detailedInfo: tokenInfo
    };
  } catch (error) {
    console.error('手動入力トークン設定エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * アプリIDとアプリシークレットを使用して長期トークンに交換する
 * @param {string} userToken ユーザートークン
 * @param {string} appId アプリID
 * @param {string} appSecret アプリシークレット
 * @return {Object} 長期トークン情報
 */
function exchangeToLongLivedTokenWithCredentials(userToken, appId, appSecret) {
  const url = `${FB_TOKEN_ENDPOINTS.EXCHANGE_TOKEN}` +
    `?grant_type=fb_exchange_token` +
    `&client_id=${encodeURIComponent(appId)}` +
    `&client_secret=${encodeURIComponent(appSecret)}` +
    `&fb_exchange_token=${encodeURIComponent(userToken)}`;

  console.log('長期トークン交換URL:', url.replace(appSecret, '***SECRET***'));
  
  const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = response.getResponseCode();
  const responseText = response.getContentText();
  
  console.log(`長期トークン交換レスポンス: ${code}`);
  console.log('レスポンス内容:', responseText);
  
  if (code < 200 || code >= 300) {
    console.error(`長期トークン交換エラー: ${code} ${responseText}`);
    
    if (code === 400) {
      throw new Error('アプリIDまたはアプリシークレットが正しくありません。');
    } else if (code === 401) {
      throw new Error('ユーザートークンが無効または期限切れです。');
    } else {
      throw new Error(`長期トークン交換に失敗しました: ${code} ${responseText}`);
    }
  }

  const data = JSON.parse(responseText);
  if (!data.access_token) {
    throw new Error('長期トークンの取得に失敗しました');
  }

  console.log('長期トークン交換成功:');
  console.log('- アクセストークン:', data.access_token.substring(0, 20) + '...');
  console.log('- 有効期限（秒）:', data.expires_in);
  console.log('- トークンタイプ:', data.token_type);

  return {
    accessToken: data.access_token,
    expiresIn: data.expires_in || 0,
    tokenType: data.token_type || 'bearer'
  };
}

/**
 * 現在のトークン情報を取得する（改良版）
 * @return {Object} トークン情報
 */
function getTokenInfo() {
  const props = PropertiesService.getScriptProperties();
  const expiresAt = props.getProperty(PROP_KEYS.TOKEN_EXPIRES_AT);
  const userToken = props.getProperty(PROP_KEYS.USER_ACCESS_TOKEN);
  
  let detailedInfo = null;
  if (userToken) {
    try {
      detailedInfo = getDetailedTokenInfo(userToken);
    } catch (error) {
      console.warn('詳細トークン情報の取得に失敗:', error.message);
    }
  }
  
  return {
    pageName: props.getProperty(PROP_KEYS.PAGE_NAME) || '未設定',
    pageId: props.getProperty(PROP_KEYS.PAGE_ID) || '未設定',
    expiresAt: expiresAt ? new Date(parseInt(expiresAt, 10)) : null,
    isValid: isTokenValid(),
    detailedInfo: detailedInfo
  };
}

/**
 * 定期実行用のトークン更新チェック（改良版）
 * 30日に一度実行されることを想定
 */
function checkAndRefreshToken() {
  try {
    const userToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.USER_ACCESS_TOKEN);
    
    // APIで直接トークンの有効性を確認
    const isValidByApi = userToken ? isTokenValid(userToken) : false;
    const isValidByStorage = isTokenValid();
    
    console.log(`トークン有効性確認 - API: ${isValidByApi}, ストレージ: ${isValidByStorage}`);
    
    if (!isValidByApi || !isValidByStorage) {
      const success = refreshToken();
      if (success) {
        console.log('トークンが正常に更新されました');
      } else {
        console.error('トークンの更新に失敗しました');
      }
    } else {
      console.log('トークンは有効です');
    }
  } catch (error) {
    console.error('トークン更新チェックでエラー:', error);
  }
}