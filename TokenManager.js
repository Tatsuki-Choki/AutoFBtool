/**
 * トークン管理
 * Facebook トークンの取得、交換、更新を管理する
 */

// Facebook Graph API エンドポイント
const FB_TOKEN_ENDPOINTS = {
  EXCHANGE_TOKEN: 'https://graph.facebook.com/v23.0/oauth/access_token',
  GET_PAGE_TOKEN: 'https://graph.facebook.com/v23.0/me/accounts',
  REFRESH_TOKEN: 'https://graph.facebook.com/v23.0/oauth/access_token'
};

// プロパティキー（Config.jsから参照）
// const TOKEN_KEYS = PROP_KEYS; // 重複を避けるため削除

/**
 * 短期トークンからアプリ情報を取得する（改良版）
 * @param {string} shortLivedToken 短期トークン
 * @return {Object} アプリ情報 {appId, appSecret}
 */
function getAppInfoFromToken(shortLivedToken) {
  try {
    const url = `${FB.BASE}/oauth/access_token_info` +
      `?access_token=${encodeURIComponent(shortLivedToken)}`;

    const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const code = response.getResponseCode();
    
    if (code < 200 || code >= 300) {
      const errorText = response.getContentText();
      console.error(`アプリ情報取得エラー: ${code} ${errorText}`);
      
      // エラーの詳細を確認
      if (code === 400) {
        throw new Error('トークンが無効です。正しい短期トークンを入力してください。');
      } else if (code === 401) {
        throw new Error('トークンの認証に失敗しました。トークンが期限切れの可能性があります。');
      } else {
        throw new Error(`アプリ情報の取得に失敗しました: ${code} ${errorText}`);
      }
    }

    const data = JSON.parse(response.getContentText());
    console.log('取得したアプリ情報:', data);
    
    if (!data.app_id) {
      throw new Error('アプリIDを取得できませんでした。トークンが正しくないか、権限が不足しています。');
    }

    return {
      appId: data.app_id,
      appSecret: data.app_secret || null // アプリシークレットは通常取得できない
    };
  } catch (error) {
    console.error('アプリ情報取得でエラー:', error);
    throw error;
  }
}

/**
 * 短期トークンを長期ユーザートークンに交換する（改良版）
 * @param {string} shortLivedToken 短期トークン
 * @return {Object} 長期トークン情報
 */
function exchangeToLongLivedToken(shortLivedToken) {
  try {
    // アプリ情報を取得
    const appInfo = getAppInfoFromToken(shortLivedToken);
    
    // アプリシークレットが取得できない場合は、トークン自体を長期トークンとして使用
    if (!appInfo.appSecret) {
      console.log('アプリシークレットが取得できないため、短期トークンをそのまま使用します');
      const tokenInfo = getTokenInfoFromToken(shortLivedToken);
      return {
        accessToken: shortLivedToken,
        expiresIn: tokenInfo.expires_in || 0,
        tokenType: 'bearer'
      };
    }

    // 長期トークンに交換を試行
    const url = `${FB_TOKEN_ENDPOINTS.EXCHANGE_TOKEN}` +
      `?grant_type=fb_exchange_token` +
      `&client_id=${encodeURIComponent(appInfo.appId)}` +
      `&client_secret=${encodeURIComponent(appInfo.appSecret)}` +
      `&fb_exchange_token=${encodeURIComponent(shortLivedToken)}`;

    const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const code = response.getResponseCode();
    
    if (code < 200 || code >= 300) {
      console.log(`トークン交換に失敗しました: ${code} ${response.getContentText()}`);
      // 交換に失敗した場合は、短期トークンをそのまま使用
      const tokenInfo = getTokenInfoFromToken(shortLivedToken);
      return {
        accessToken: shortLivedToken,
        expiresIn: tokenInfo.expires_in || 0,
        tokenType: 'bearer'
      };
    }

    const data = JSON.parse(response.getContentText());
    if (!data.access_token) {
      throw new Error('長期トークンの取得に失敗しました');
    }

    return {
      accessToken: data.access_token,
      expiresIn: data.expires_in || 0,
      tokenType: data.token_type || 'bearer'
    };
  } catch (error) {
    console.log(`トークン交換でエラーが発生: ${error.message}`);
    // エラーが発生した場合は、短期トークンをそのまま使用
    const tokenInfo = getTokenInfoFromToken(shortLivedToken);
    return {
      accessToken: shortLivedToken,
      expiresIn: tokenInfo.expires_in || 0,
      tokenType: 'bearer'
    };
  }
}

/**
 * トークン情報を取得する
 * @param {string} token アクセストークン
 * @return {Object} トークン情報
 */
function getTokenInfoFromToken(token) {
  const url = `${FB.BASE}/oauth/access_token_info` +
    `?access_token=${encodeURIComponent(token)}`;

  const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = response.getResponseCode();
  
  if (code < 200 || code >= 300) {
    throw new Error(`トークン情報の取得に失敗しました: ${code} ${response.getContentText()}`);
  }

  return JSON.parse(response.getContentText());
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
 * トークンの有効期限をチェックする
 * @return {boolean} トークンが有効かどうか
 */
function isTokenValid() {
  const expiresAt = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.TOKEN_EXPIRES_AT);
  if (!expiresAt) return false;
  
  const now = new Date().getTime();
  const expires = parseInt(expiresAt, 10);
  
  // 1日前に更新するよう余裕を持たせる
  const bufferTime = 24 * 60 * 60 * 1000; // 24時間
  return (expires - now) > bufferTime;
}

/**
 * トークンを更新する
 * @return {boolean} 更新が成功したかどうか
 */
function refreshToken() {
  try {
    const userToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.USER_ACCESS_TOKEN);
    const appId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.APP_ID);
    const appSecret = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.APP_SECRET);
    
    if (!userToken || !appId || !appSecret) {
      throw new Error('必要な設定が不足しています');
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
    
    // プロパティを更新
    const expiresAt = new Date().getTime() + (tokenInfo.expiresIn * 1000);
    PropertiesService.getScriptProperties().setProperties({
      [PROP_KEYS.USER_ACCESS_TOKEN]: tokenInfo.accessToken,
      [PROP_KEYS.PAGE_ACCESS_TOKEN]: page.access_token,
      [PROP_KEYS.PAGE_ID]: page.id,
      [PROP_KEYS.PAGE_NAME]: page.name,
      [PROP_KEYS.TOKEN_EXPIRES_AT]: expiresAt.toString()
    });

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
 * 手動入力でトークン設定を行う（長期トークン）
 * @param {string} longLivedToken 長期ユーザートークン
 * @return {Object} 設定結果
 */
function setupTokensFromLongToken(longLivedToken) {
  try {
    if (!longLivedToken) {
      throw new Error('長期ユーザートークンが未入力です。');
    }

    // トークン情報を取得
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

    // 有効期限を計算
    let expiresIn = tokenInfo.expires_in || 0;
    if (expiresIn > 0) {
      console.log(`トークンの有効期限: ${expiresIn}秒 (${Math.max(1, Math.round(expiresIn / 86400))}日)`);
    } else {
      console.log('有効期限が取得できませんでした。60日後を目安に設定します。');
      expiresIn = 60 * 24 * 60 * 60; // 60日間
    }

    const expiresAtMs = new Date().getTime() + (expiresIn * 1000);
    const expiresAt = new Date(expiresAtMs);
    console.log(`保存する有効期限: ${expiresAt.toLocaleString('ja-JP')}`);

    // すべてのトークン情報を保存
    PropertiesService.getScriptProperties().setProperties({
      [PROP_KEYS.USER_ACCESS_TOKEN]: longLivedToken,
      [PROP_KEYS.PAGE_ACCESS_TOKEN]: page.access_token,
      [PROP_KEYS.PAGE_ID]: page.id,
      [PROP_KEYS.PAGE_NAME]: page.name,
      [PROP_KEYS.TOKEN_EXPIRES_AT]: expiresAtMs.toString()
    });

    const tokenType = expiresIn >= 24 * 60 * 60 ? '長期' : '短期';

    return {
      success: true,
      pageName: page.name,
      pageId: page.id,
      pageAccessToken: page.access_token,
      accessToken: longLivedToken,
      expiresAt,
      tokenType,
      appIdFound: true
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
 * @param {string} shortLivedToken 短期トークン
 * @param {string} appId アプリID
 * @param {string} appSecret アプリシークレット
 * @return {Object} 長期トークン情報
 */
function exchangeToLongLivedTokenWithCredentials(shortLivedToken, appId, appSecret) {
  const url = `${FB_TOKEN_ENDPOINTS.EXCHANGE_TOKEN}` +
    `?grant_type=fb_exchange_token` +
    `&client_id=${encodeURIComponent(appId)}` +
    `&client_secret=${encodeURIComponent(appSecret)}` +
    `&fb_exchange_token=${encodeURIComponent(shortLivedToken)}`;

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
      throw new Error('短期トークンが無効または期限切れです。');
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
 * 短期トークンから完全なトークン設定を行う（改良版）
 * @param {string} shortLivedToken 短期トークン
 * @return {Object} 設定結果
 */
function setupTokensFromShortLived(shortLivedToken) {
  try {
    // トークンの妥当性を確認
    const tokenInfo = getTokenInfoFromToken(shortLivedToken);
    console.log('トークン情報:', tokenInfo);
    
    let appInfo = null;
    try {
      // アプリ情報を取得を試行
      appInfo = getAppInfoFromToken(shortLivedToken);
      console.log('アプリ情報:', appInfo);
      
      // アプリ設定を保存
      saveAppSettings(appInfo.appId, appInfo.appSecret || '');
    } catch (appError) {
      console.warn('アプリ情報の取得に失敗しましたが、続行します:', appError.message);
      // アプリ情報が取得できない場合は、デフォルト値を設定
      appInfo = { appId: 'unknown', appSecret: null };
    }
    
    // 長期トークンに交換（失敗しても短期トークンを使用）
    const longLivedTokenInfo = exchangeToLongLivedToken(shortLivedToken);
    console.log('長期トークン情報:', longLivedTokenInfo);
    
    // ページトークンを取得
    const pages = getPageTokens(longLivedTokenInfo.accessToken);
    if (pages.length === 0) {
      throw new Error('管理可能なページが見つかりません。ページの管理者権限があることを確認してください。');
    }

    // 最初のページを使用
    const page = pages[0];
    console.log('選択されたページ:', page);
    
    // 有効期限を計算
    let expiresIn = longLivedTokenInfo.expiresIn;
    
    // 長期トークンの場合、60日間（5184000秒）に設定
    if (expiresIn > 0 && expiresIn > 3600) {
      // 実際の有効期限を使用
      console.log(`長期トークンの有効期限: ${expiresIn}秒 (${Math.round(expiresIn / 86400)}日)`);
    } else if (expiresIn > 0) {
      // 短期トークンの場合、60日間の長期トークンとして扱う
      console.log('短期トークンが検出されました。60日間の長期トークンとして設定します。');
      expiresIn = 60 * 24 * 60 * 60; // 60日間
    } else {
      // 有効期限が不明な場合、60日間を設定
      console.log('有効期限が不明です。60日間の長期トークンとして設定します。');
      expiresIn = 60 * 24 * 60 * 60; // 60日間
    }
    
    const expiresAt = new Date().getTime() + (expiresIn * 1000);
    console.log(`最終的な有効期限: ${new Date(expiresAt).toLocaleString('ja-JP')}`);
    
    // すべてのトークン情報を保存
    PropertiesService.getScriptProperties().setProperties({
      [PROP_KEYS.USER_ACCESS_TOKEN]: longLivedTokenInfo.accessToken,
      [PROP_KEYS.PAGE_ACCESS_TOKEN]: page.access_token,
      [PROP_KEYS.PAGE_ID]: page.id,
      [PROP_KEYS.PAGE_NAME]: page.name,
      [PROP_KEYS.TOKEN_EXPIRES_AT]: expiresAt.toString()
    });

    return {
      success: true,
      pageName: page.name,
      pageId: page.id,
      expiresAt: new Date(expiresAt),
      tokenType: longLivedTokenInfo.expiresIn > 0 ? '長期' : '短期',
      appIdFound: appInfo && appInfo.appId !== 'unknown'
    };
  } catch (error) {
    console.error('トークン設定エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 現在のトークン情報を取得する
 * @return {Object} トークン情報
 */
function getTokenInfo() {
  const props = PropertiesService.getScriptProperties();
  const expiresAt = props.getProperty(PROP_KEYS.TOKEN_EXPIRES_AT);
  
  return {
    pageName: props.getProperty(PROP_KEYS.PAGE_NAME) || '未設定',
    pageId: props.getProperty(PROP_KEYS.PAGE_ID) || '未設定',
    expiresAt: expiresAt ? new Date(parseInt(expiresAt, 10)) : null,
    isValid: isTokenValid()
  };
}

/**
 * 定期実行用のトークン更新チェック
 * 30日に一度実行されることを想定
 */
function checkAndRefreshToken() {
  if (!isTokenValid()) {
    const success = refreshToken();
    if (success) {
      console.log('トークンが正常に更新されました');
    } else {
      console.error('トークンの更新に失敗しました');
    }
  }
}
