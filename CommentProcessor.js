/**
 * コメント処理
 * コメントの取得、判定、返信処理を管理する
 */

/**
 * コメントを処理する
 * @param {string} postId 投稿ID
 * @param {string} token アクセストークン
 * @param {number} fetchLimit 取得件数
 * @param {Array} rules ルール配列
 * @param {Set} processedSet 処理済みコメントIDのSet
 */
function processComments(postId, token, fetchLimit, rules, processedSet) {
  const comments = fetchComments(postId, token, fetchLimit);
  
  for (const comment of comments) {
    const commentId = comment.id;
    if (processedSet.has(commentId)) continue; // 既読スキップ

    const message = (comment.message || '').toString();
    const name = (comment.from && comment.from.name) ? comment.from.name : '';

    // キーワードマッチング
    const matchedRule = findMatchingRule(message, rules);
    if (!matchedRule) {
      appendProcessed(commentId, comment.created_time);
      logComment(postId, commentId, name, message, '', '', 'no_match', '');
      continue;
    }

    // 返信処理
    const replyText = generateReply(matchedRule.template, name);
    try {
      postReply(commentId, replyText, token);
      appendProcessed(commentId, comment.created_time);
      logComment(postId, commentId, name, message, matchedRule.keyword, replyText, 'replied', '');
      Utilities.sleep(300); // レート制限緩和
    } catch (error) {
      logComment(postId, commentId, name, message, matchedRule.keyword, replyText, 'error', String(error && error.message ? error.message : error));
    }
  }
}

/**
 * メッセージにマッチするルールを検索する
 * @param {string} message コメントメッセージ
 * @param {Array} rules ルール配列
 * @return {Object|null} マッチしたルール、なければnull
 */
function findMatchingRule(message, rules) {
  return rules.find(rule => 
    rule.enabled && 
    rule.keyword && 
    message.toLowerCase().includes(rule.keyword.toLowerCase())
  );
}

/**
 * 返信テキストを生成する
 * @param {string} template テンプレート
 * @param {string} name コメント投稿者名
 * @return {string} 生成された返信テキスト
 */
function generateReply(template, name) {
  return String(template || '').replaceAll('{name}', name || '');
}

/**
 * 対象投稿IDを取得する
 * @param {string} token アクセストークン
 * @param {Map} settings 設定Map
 * @return {Array} 投稿ID配列
 */
function getTargetPostIds(token, settings) {
  const manualPostIds = (settings.get('投稿ID') || '').split(',').map(s => s.trim()).filter(Boolean);
  const autoDiscovery = (settings.get('自動検出') || 'true').toLowerCase() === 'true';
  const lookbackDays = parseInt(settings.get('検索日数') || '3', 10);
  const postMaxCount = parseInt(settings.get('最大投稿数') || '10', 10);

  // 手動指定の投稿IDがある場合はそれを使用
  if (manualPostIds.length > 0) {
    return manualPostIds;
  }

  // 自動検出が有効な場合は直近投稿を取得
  if (autoDiscovery) {
    // ページトークンの妥当性を確認
    getPageInfo(token);
    const posts = fetchRecentPosts(token, lookbackDays, postMaxCount);
    const reels = fetchRecentReels(token, lookbackDays, postMaxCount);
    // 先に通常投稿、続いてリール。重複排除
    const seen = new Set();
    const merged = [];
    for (const id of [...posts, ...reels]) {
      if (!seen.has(id)) { seen.add(id); merged.push(id); }
    }
    return merged;
  }

  return [];
}
