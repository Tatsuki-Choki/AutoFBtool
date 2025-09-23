/**
 * Facebook API操作
 * Facebook Graph APIとの通信を管理する
 */

/**
 * コメントを取得する
 * @param {string} postId 投稿ID
 * @param {string} token アクセストークン
 * @param {number} limit 取得件数
 * @return {Array} コメント配列
 */
function fetchComments(postId, token, limit) {
  // クエリ値を必ずURLエンコード（特に from{name} の { } をエンコード）
  const fields = encodeURIComponent('id,from{name},message,created_time');
  const url = `${FB.BASE}/${encodeURIComponent(postId)}/comments` +
    `?fields=${fields}` +
    `&filter=${encodeURIComponent('stream')}` +
    `&limit=${encodeURIComponent(String(Math.min(500, Number(limit) || 500)))}` +
    `&access_token=${encodeURIComponent(token)}`;

  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code < 200 || code >= 300) {
    console.error(`GET comments failed: ${code} ${responseText}`);
    
    // トークン期限切れエラーの場合、自動更新を試行
    if (code === 400 && responseText.includes("Session has expired")) {
      console.log("トークンが期限切れのため、自動更新を試行します...");
      try {
        const refreshSuccess = refreshToken();
        if (refreshSuccess) {
          const newToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
          console.log("トークンが更新されました。再試行します...");
          return fetchComments(postId, newToken, limit); // 再帰的に再試行
        } else {
          throw new Error("トークンの自動更新に失敗しました。手動でトークンを更新してください。");
        }
      } catch (refreshError) {
        console.error("トークン更新エラー:", refreshError);
        throw new Error(`トークンの自動更新に失敗しました: ${refreshError.message}`);
      }
    }
    
    throw new Error(`GET comments failed: ${code} ${responseText}`);
  }
  
  const json = JSON.parse(responseText);
  return Array.isArray(json.data) ? json.data : [];
}

/**
 * 直近のコメントを取得（since指定）
 * @param {string} postId 投稿ID
 * @param {string} token アクセストークン
 * @param {number} sinceMs 取得開始のUNIXミリ秒
 * @param {number} limit 取得件数（任意、既定200）
 * @return {Array} コメント配列
 */
function fetchCommentsSince(postId, token, sinceMs, limit) {
  const fields = encodeURIComponent('id,from{name},message,created_time');
  const sinceSec = Math.floor(Number(sinceMs) / 1000);
  const url = `${FB.BASE}/${encodeURIComponent(postId)}/comments` +
    `?fields=${fields}` +
    `&filter=${encodeURIComponent('stream')}` +
    `&since=${encodeURIComponent(String(sinceSec))}` +
    `&limit=${encodeURIComponent(String(Math.min(500, Number(limit) || 500)))}` +
    `&access_token=${encodeURIComponent(token)}`;

  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code < 200 || code >= 300) {
    console.error(`GET comments(since) failed: ${code} ${responseText}`);
    
    // トークン期限切れエラーの場合、自動更新を試行
    if (code === 400 && responseText.includes("Session has expired")) {
      console.log("トークンが期限切れのため、自動更新を試行します...");
      try {
        const refreshSuccess = refreshToken();
        if (refreshSuccess) {
          const newToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
          console.log("トークンが更新されました。再試行します...");
          return fetchCommentsSince(postId, newToken, sinceMs, limit); // 再帰的に再試行
        } else {
          throw new Error("トークンの自動更新に失敗しました。手動でトークンを更新してください。");
        }
      } catch (refreshError) {
        console.error("トークン更新エラー:", refreshError);
        throw new Error(`トークンの自動更新に失敗しました: ${refreshError.message}`);
      }
    }
    
    throw new Error(`GET comments(since) failed: ${code} ${responseText}`);
  }
  
  const json = JSON.parse(responseText);
  return Array.isArray(json.data) ? json.data : [];
}

/**
 * コメントに返信を投稿する
 * @param {string} commentId コメントID
 * @param {string} message 返信メッセージ
 * @param {string} token アクセストークン
 * @return {string} 返信ID
 */
function postReply(commentId, message, token) {
  const url = `${FB.BASE}/${encodeURIComponent(commentId)}/comments`;
  const payload = { message: message, access_token: token };
  const res = UrlFetchApp.fetch(url, { method: 'post', payload: payload, muteHttpExceptions: true });
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code < 200 || code >= 300) {
    console.error(`POST reply failed: ${code} ${responseText}`);
    
    // トークン期限切れエラーの場合、自動更新を試行
    if (code === 400 && responseText.includes("Session has expired")) {
      console.log("トークンが期限切れのため、自動更新を試行します...");
      try {
        const refreshSuccess = refreshToken();
        if (refreshSuccess) {
          const newToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
          console.log("トークンが更新されました。再試行します...");
          return postReply(commentId, message, newToken); // 再帰的に再試行
        } else {
          throw new Error("トークンの自動更新に失敗しました。手動でトークンを更新してください。");
        }
      } catch (refreshError) {
        console.error("トークン更新エラー:", refreshError);
        throw new Error(`トークンの自動更新に失敗しました: ${refreshError.message}`);
      }
    }
    
    throw new Error(`POST reply failed: ${code} ${responseText}`);
  }
  
  const json = JSON.parse(responseText);
  if (!json || !json.id) throw new Error(`POST reply unexpected response: ${responseText}`);
  return json.id;
}

/**
 * ページに新規投稿を作成する
 * @param {string} message 投稿本文
 * @param {string} token ページアクセストークン
 * @return {Object} {postId, permalink}
 */
function createPagePost(message, token) {
  const url = `${FB.BASE}/me/feed`;
  const payload = { message: message, access_token: token };
  const res = UrlFetchApp.fetch(url, { method: 'post', payload: payload, muteHttpExceptions: true });
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code < 200 || code >= 300) {
    console.error(`POST feed failed: ${code} ${responseText}`);
    
    // トークン期限切れエラーの場合、自動更新を試行
    if (code === 400 && responseText.includes("Session has expired")) {
      console.log("トークンが期限切れのため、自動更新を試行します...");
      try {
        const refreshSuccess = refreshToken();
        if (refreshSuccess) {
          const newToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
          console.log("トークンが更新されました。再試行します...");
          return createPagePost(message, newToken); // 再帰的に再試行
        } else {
          throw new Error("トークンの自動更新に失敗しました。手動でトークンを更新してください。");
        }
      } catch (refreshError) {
        console.error("トークン更新エラー:", refreshError);
        throw new Error(`トークンの自動更新に失敗しました: ${refreshError.message}`);
      }
    }
    
    throw new Error(`POST feed failed: ${code} ${responseText}`);
  }
  
  const json = JSON.parse(responseText);
  if (!json || !json.id) throw new Error(`POST feed unexpected response: ${responseText}`);

  // パーマリンクを取得
  let permalink = '';
  try {
    const detailUrl = `${FB.BASE}/${encodeURIComponent(json.id)}?fields=permalink_url&access_token=${encodeURIComponent(token)}`;
    const detailRes = UrlFetchApp.fetch(detailUrl, { method: 'get', muteHttpExceptions: true });
    if (detailRes.getResponseCode() >= 200 && detailRes.getResponseCode() < 300) {
      const d = JSON.parse(detailRes.getContentText());
      permalink = d.permalink_url || '';
    }
  } catch (e) {
    // 取得失敗は致命的でないため握りつぶす
  }

  return { postId: String(json.id), permalink };
}

/**
 * ページトークンからページ情報を取得する（妥当性チェック）
 * @param {string} token アクセストークン
 * @return {Object} ページ情報 {pageId, pageName}
 */
function getPageInfo(token) {
  const url = `${FB.BASE}/me?fields=id,name&access_token=${encodeURIComponent(token)}`;
  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code < 200 || code >= 300) {
    console.error(`GET /me failed: ${code} ${responseText}`);
    
    // トークン期限切れエラーの場合、自動更新を試行
    if (code === 400 && responseText.includes("Session has expired")) {
      console.log("トークンが期限切れのため、自動更新を試行します...");
      try {
        const refreshSuccess = refreshToken();
        if (refreshSuccess) {
          const newToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
          console.log("トークンが更新されました。再試行します...");
          return getPageInfo(newToken); // 再帰的に再試行
        } else {
          throw new Error("トークンの自動更新に失敗しました。手動でトークンを更新してください。");
        }
      } catch (refreshError) {
        console.error("トークン更新エラー:", refreshError);
        throw new Error(`トークンの自動更新に失敗しました: ${refreshError.message}`);
      }
    }
    
    throw new Error(`GET /me failed: ${code} ${responseText}`);
  }
  
  const json = JSON.parse(responseText);
  if (!json || !json.id) throw new Error('ページIDを取得できませんでした。アクセストークンが「ページ用」か確認してください。');
  return { pageId: String(json.id), pageName: json.name || '' };
}

/**
 * 直近の投稿を取得する
 * @param {string} token アクセストークン
 * @param {number} lookbackDays 過去何日分を取得するか
 * @param {number} maxCount 最大取得件数
 * @return {Array} 投稿ID配列
 */
function fetchRecentPosts(token, lookbackDays, maxCount) {
  const sinceTs = new Date(Date.now() - lookbackDays * 24 * 60 * 60 * 1000);
  const sinceParam = Math.floor(sinceTs.getTime() / 1000);
  const url = `${FB.BASE}/me/posts` +
    `?fields=id,created_time,permalink_url` +
    `&since=${sinceParam}` +
    `&limit=${encodeURIComponent(String(maxCount))}` +
    `&access_token=${encodeURIComponent(token)}`;
  
  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code < 200 || code >= 300) {
    console.error(`GET /me/posts failed: ${code} ${responseText}`);
    
    // トークン期限切れエラーの場合、自動更新を試行
    if (code === 400 && responseText.includes("Session has expired")) {
      console.log("トークンが期限切れのため、自動更新を試行します...");
      try {
        const refreshSuccess = refreshToken();
        if (refreshSuccess) {
          const newToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
          console.log("トークンが更新されました。再試行します...");
          return fetchRecentPosts(newToken, lookbackDays, maxCount); // 再帰的に再試行
        } else {
          throw new Error("トークンの自動更新に失敗しました。手動でトークンを更新してください。");
        }
      } catch (refreshError) {
        console.error("トークン更新エラー:", refreshError);
        throw new Error(`トークンの自動更新に失敗しました: ${refreshError.message}`);
      }
    }
    
    throw new Error(`GET /me/posts failed: ${code} ${responseText}`);
  }
  
  const json = JSON.parse(responseText);
  const data = Array.isArray(json.data) ? json.data : [];
  
  // Postsシートに記録
  if (data.length) {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.POSTS);
    const rows = data.map(p => [p.id || '', p.permalink_url || '', p.created_time || '']);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  }
  
  return data.map(p => p.id).filter(Boolean);
}

/**
 * 直近のリールを取得する（自身が投稿したもの）
 * @param {string} token アクセストークン（ページ）
 * @param {number} lookbackDays 過去何日分を取得するか
 * @param {number} maxCount 最大取得件数
 * @return {Array} リールID配列
 */
function fetchRecentReels(token, lookbackDays, maxCount) {
  try {
    const { pageId } = getPageInfo(token);
    const sinceTs = new Date(Date.now() - lookbackDays * 24 * 60 * 60 * 1000);
    const sinceParam = Math.floor(sinceTs.getTime() / 1000);
    const url = `${FB.BASE}/${encodeURIComponent(pageId)}/video_reels` +
      `?fields=id,created_time,permalink_url` +
      `&since=${sinceParam}` +
      `&limit=${encodeURIComponent(String(maxCount))}` +
      `&access_token=${encodeURIComponent(token)}`;

    const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const code = res.getResponseCode();
    const responseText = res.getContentText();

    if (code < 200 || code >= 300) {
      console.error(`GET /{page-id}/video_reels failed: ${code} ${responseText}`);

      // トークン期限切れエラーの場合、自動更新を試行
      if (code === 400 && responseText.includes("Session has expired")) {
        console.log("トークンが期限切れのため、自動更新を試行します...");
        try {
          const refreshSuccess = refreshToken();
          if (refreshSuccess) {
            const newToken = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.PAGE_ACCESS_TOKEN);
            console.log("トークンが更新されました。再試行します...");
            return fetchRecentReels(newToken, lookbackDays, maxCount); // 再帰的に再試行
          } else {
            throw new Error("トークンの自動更新に失敗しました。手動でトークンを更新してください。");
          }
        } catch (refreshError) {
          console.error("トークン更新エラー:", refreshError);
          throw new Error(`トークンの自動更新に失敗しました: ${refreshError.message}`);
        }
      }

      // 許可されていない・未サポートの場合は空配列で返却（ソフトフォールバック）
      return [];
    }

    const json = JSON.parse(responseText);
    const data = Array.isArray(json.data) ? json.data : [];

    // Postsシートに記録
    if (data.length) {
      const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.POSTS);
      const rows = data.map(p => [p.id || '', p.permalink_url || '', p.created_time || '']);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    }

    return data.map(p => p.id).filter(Boolean);
  } catch (e) {
    console.warn('リール取得に失敗（フォールバックで続行）:', e && e.message ? e.message : e);
    return [];
  }
}
