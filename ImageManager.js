/**
 * 画像管理機能
 * ImgBB APIを使用した画像アップロードとURL管理
 */

/**
 * ImgBB APIキーを設定する
 * @param {string} apiKey ImgBB APIキー
 */
function setImgBBApiKey(apiKey) {
  if (!apiKey || !apiKey.trim()) {
    throw new Error('APIキーが入力されていません');
  }
  
  // プロパティに保存
  PropertiesService.getScriptProperties().setProperty(PROP_KEYS.IMGBB_API_KEY, apiKey.trim());
  
  // 設定シートにも保存
  upsertSetting('ImgBB APIキー', apiKey.trim(), SETTING_DESCRIPTIONS['ImgBB APIキー']);
  
  console.log('ImgBB APIキーを設定しました');
}

/**
 * ImgBB APIキーを取得する
 * @return {string} APIキー
 */
function getImgBBApiKey() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEYS.IMGBB_API_KEY) || '';
}

/**
 * 画像をImgBBにアップロードする
 * @param {Blob} imageBlob 画像データ
 * @param {string} filename ファイル名（オプション）
 * @return {Object} アップロード結果 {success, url, error}
 */
function uploadImageToImgBB(imageBlob, filename = '') {
  const apiKey = getImgBBApiKey();
  if (!apiKey) {
    return {
      success: false,
      error: 'ImgBB APIキーが設定されていません。設定画面でAPIキーを入力してください。'
    };
  }
  
  try {
    // Base64エンコード
    const base64Data = Utilities.base64Encode(imageBlob.getBytes());
    
    // APIリクエストの準備
    const url = 'https://api.imgbb.com/1/upload';
    const payload = {
      'key': apiKey,
      'image': base64Data
    };
    
    if (filename) {
      payload['name'] = filename;
    }
    
    // API呼び出し
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      payload: payload,
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      console.error('ImgBB API Error:', responseCode, responseText);
      return {
        success: false,
        error: `アップロードに失敗しました (HTTP ${responseCode})`
      };
    }
    
    const result = JSON.parse(responseText);
    
    if (!result.success) {
      return {
        success: false,
        error: result.error?.message || 'アップロードに失敗しました'
      };
    }
    
    return {
      success: true,
      url: result.data.url,
      displayUrl: result.data.display_url,
      deleteUrl: result.data.delete_url,
      imageId: result.data.id
    };
    
  } catch (error) {
    console.error('画像アップロードエラー:', error);
    return {
      success: false,
      error: 'アップロード中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 画像URLをクリップボードにコピーするためのHTMLを生成
 * @param {string} imageUrl 画像URL
 * @return {string} HTML文字列
 */
function generateImageUrlDisplayHtml(imageUrl) {
  return `
    <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
      <h3>画像アップロード完了</h3>
      
      <div style="margin-bottom: 20px;">
        <label style="display: block; margin-bottom: 5px; font-weight: bold;">画像URL:</label>
        <div style="display: flex; align-items: center; gap: 10px;">
          <input type="text" id="imageUrl" value="${imageUrl}" 
                 style="flex: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-family: monospace;" 
                 readonly>
          <button onclick="copyToClipboard()" 
                  style="background-color: #2196F3; color: white; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; display: flex; align-items: center; gap: 5px;">
            📋 コピー
          </button>
        </div>
      </div>
      
      <div style="text-align: center; margin-top: 20px;">
        <a href="https://api.imgbb.com/" target="_blank" 
           style="color: #2196F3; text-decoration: none;">
          ImgBB APIキーの取得はこちら
        </a>
      </div>
      
      <div style="text-align: center; margin-top: 20px;">
        <button onclick="google.script.host.close()" 
                style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
          閉じる
        </button>
      </div>
    </div>

    <script>
      function copyToClipboard() {
        const input = document.getElementById('imageUrl');
        input.select();
        input.setSelectionRange(0, 99999); // モバイル対応
        
        try {
          document.execCommand('copy');
          alert('画像URLをクリップボードにコピーしました！');
        } catch (err) {
          // フォールバック: 選択状態を維持
          input.focus();
          alert('URLを選択しました。Ctrl+C（Cmd+C）でコピーしてください。');
        }
      }
    </script>
  `;
}

/**
 * 画像アップロード設定ダイアログを表示
 */
function showImageUploadSettings() {
  const currentApiKey = getImgBBApiKey();
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 500px;">
      <h3>画像アップロード設定</h3>
      <p style="margin-bottom: 20px;">ImgBB APIキーを設定して画像アップロード機能を使用できます。</p>

      <div style="margin-bottom: 20px;">
        <label for="apiKey" style="display: block; margin-bottom: 5px; font-weight: bold;">ImgBB APIキー</label>
        <input type="text" id="apiKey" value="${currentApiKey}" 
               style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;" 
               placeholder="APIキーを入力してください">
      </div>
      
      <div style="margin-bottom: 20px; padding: 10px; background-color: #f5f5f5; border-radius: 4px;">
        <p style="margin: 0; font-size: 14px;">
          <strong>APIキーの取得方法:</strong><br>
          1. <a href="https://api.imgbb.com/" target="_blank" style="color: #2196F3;">ImgBB</a> にアクセス<br>
          2. アカウントを作成またはログイン<br>
          3. APIキーを取得して入力
        </p>
      </div>

      <div style="text-align: center;">
        <button onclick="saveApiKey()" 
                style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px;">
          保存
        </button>
        <button onclick="google.script.host.close()" 
                style="background-color: #f44336; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
          キャンセル
        </button>
      </div>
    </div>

    <script>
      function saveApiKey() {
        const apiKey = document.getElementById('apiKey').value.trim();

        if (!apiKey) {
          alert('APIキーを入力してください。');
          return;
        }

        google.script.run
          .withSuccessHandler(onSuccess)
          .withFailureHandler(onFailure)
          .setImgBBApiKey(apiKey);
      }

      function onSuccess() {
        alert('✅ APIキーを保存しました！');
        google.script.host.close();
      }

      function onFailure(error) {
        alert('❌ エラー: ' + error.message);
      }
    </script>
  `);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '画像アップロード設定');
}

/**
 * 複数画像を一括でアップロードする
 * @param {Array<Blob>} imageBlobs 画像データの配列（最大10枚）
 * @return {Object} アップロード結果 {success, results, errors}
 */
function uploadMultipleImagesToImgBB(imageBlobs) {
  const apiKey = getImgBBApiKey();
  if (!apiKey) {
    return {
      success: false,
      error: 'ImgBB APIキーが設定されていません。設定画面でAPIキーを入力してください。'
    };
  }
  
  if (!imageBlobs || imageBlobs.length === 0) {
    return {
      success: false,
      error: 'アップロードする画像がありません。'
    };
  }
  
  if (imageBlobs.length > 10) {
    return {
      success: false,
      error: '一度にアップロードできる画像は最大10枚までです。'
    };
  }
  
  const results = [];
  const errors = [];
  let successCount = 0;
  
  console.log(`複数画像アップロード開始: ${imageBlobs.length}枚`);
  
  // 各画像を順次アップロード
  for (let i = 0; i < imageBlobs.length; i++) {
    const imageBlob = imageBlobs[i];
    const filename = `image_${i + 1}.${getFileExtension(imageBlob)}`;
    
    try {
      console.log(`画像 ${i + 1}/${imageBlobs.length} をアップロード中...`);
      const result = uploadImageToImgBB(imageBlob, filename);
      
      if (result.success) {
        results.push({
          index: i + 1,
          filename: filename,
          url: result.url,
          displayUrl: result.displayUrl,
          deleteUrl: result.deleteUrl,
          imageId: result.imageId
        });
        successCount++;
        console.log(`画像 ${i + 1} アップロード成功: ${result.url}`);
      } else {
        errors.push({
          index: i + 1,
          filename: filename,
          error: result.error
        });
        console.error(`画像 ${i + 1} アップロード失敗: ${result.error}`);
      }
      
      // API制限を考慮して少し待機
      if (i < imageBlobs.length - 1) {
        Utilities.sleep(500); // 0.5秒待機
      }
      
    } catch (error) {
      errors.push({
        index: i + 1,
        filename: filename,
        error: error.message
      });
      console.error(`画像 ${i + 1} アップロードエラー:`, error);
    }
  }
  
  console.log(`アップロード完了: 成功 ${successCount}/${imageBlobs.length}枚`);
  
  return {
    success: successCount > 0,
    successCount: successCount,
    totalCount: imageBlobs.length,
    results: results,
    errors: errors
  };
}

/**
 * ファイル拡張子を取得する
 * @param {Blob} blob ファイルBlob
 * @return {string} 拡張子
 */
function getFileExtension(blob) {
  const contentType = blob.getContentType();
  if (contentType.includes('jpeg') || contentType.includes('jpg')) return 'jpg';
  if (contentType.includes('png')) return 'png';
  if (contentType.includes('gif')) return 'gif';
  if (contentType.includes('webp')) return 'webp';
  return 'jpg'; // デフォルト
}

/**
 * 複数画像アップロード結果を表示するHTMLを生成
 * @param {Object} uploadResult アップロード結果
 * @return {string} HTML文字列
 */
function generateBatchUploadResultHtml(uploadResult) {
  const { successCount, totalCount, results, errors } = uploadResult;
  
  let html = `
    <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 800px;">
      <h3>一括画像アップロード完了</h3>
      
      <div style="margin-bottom: 20px; padding: 15px; background-color: ${successCount > 0 ? '#e8f5e8' : '#ffe8e8'}; border-radius: 4px;">
        <p style="margin: 0; font-weight: bold; color: ${successCount > 0 ? '#2e7d32' : '#d32f2f'};">
          ${successCount > 0 ? '✅' : '❌'} アップロード結果: ${successCount}/${totalCount}枚成功
        </p>
      </div>
  `;
  
  // 成功した画像のURL一覧
  if (results.length > 0) {
    html += `
      <div style="margin-bottom: 20px;">
        <h4 style="color: #2e7d32; margin-bottom: 10px;">✅ アップロード成功 (${results.length}枚)</h4>
        <div style="max-height: 300px; overflow-y: auto; border: 1px solid #ddd; border-radius: 4px;">
    `;
    
    results.forEach((result, index) => {
      html += `
        <div style="padding: 10px; border-bottom: 1px solid #eee; ${index % 2 === 0 ? 'background-color: #f9f9f9;' : ''}">
          <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 5px;">
            <span style="font-weight: bold; color: #666; min-width: 60px;">${result.index}.</span>
            <span style="font-weight: bold;">${result.filename}</span>
            <button onclick="copyUrl('${result.url}')" 
                    style="background-color: #2196F3; color: white; padding: 4px 8px; border: none; border-radius: 3px; cursor: pointer; font-size: 12px;">
              📋 コピー
            </button>
          </div>
          <input type="text" id="url_${result.index}" value="${result.url}" 
                 style="width: 100%; padding: 5px; border: 1px solid #ccc; border-radius: 3px; font-family: monospace; font-size: 12px;" 
                 readonly>
        </div>
      `;
    });
    
    html += `
        </div>
      </div>
    `;
  }
  
  // エラーがあった場合
  if (errors.length > 0) {
    html += `
      <div style="margin-bottom: 20px;">
        <h4 style="color: #d32f2f; margin-bottom: 10px;">❌ アップロード失敗 (${errors.length}枚)</h4>
        <div style="max-height: 200px; overflow-y: auto; border: 1px solid #ffcdd2; border-radius: 4px; background-color: #ffebee;">
    `;
    
    errors.forEach((error, index) => {
      html += `
        <div style="padding: 10px; border-bottom: 1px solid #ffcdd2;">
          <div style="font-weight: bold; color: #d32f2f;">
            ${error.index}. ${error.filename}
          </div>
          <div style="color: #666; font-size: 12px; margin-top: 2px;">
            ${error.error}
          </div>
        </div>
      `;
    });
    
    html += `
        </div>
      </div>
    `;
  }
  
  // 全URLコピーボタン
  if (results.length > 0) {
    const allUrls = results.map(r => r.url).join('\n');
    html += `
      <div style="margin-bottom: 20px; text-align: center;">
        <button onclick="copyAllUrls()" 
                style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;">
          📋 全URLをコピー (${results.length}件)
        </button>
      </div>
    `;
  }
  
  // リンクと閉じるボタン
  html += `
      <div style="text-align: center; margin-top: 20px;">
        <a href="https://api.imgbb.com/" target="_blank" 
           style="color: #2196F3; text-decoration: none; margin-right: 20px;">
          ImgBB APIキーの取得はこちら
        </a>
        <button onclick="google.script.host.close()" 
                style="background-color: #666; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
          閉じる
        </button>
      </div>
    </div>

    <script>
      function copyUrl(url) {
        const textArea = document.createElement('textarea');
        textArea.value = url;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
        alert('URLをコピーしました！');
      }
      
      function copyAllUrls() {
        const urls = [${results.map(r => `'${r.url}'`).join(',')}];
        const textArea = document.createElement('textarea');
        textArea.value = urls.join('\\n');
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
        alert('全URLをコピーしました！');
      }
    </script>
  `;
  
  return html;
}

/**
 * 複数画像アップロードダイアログを表示
 */
function showBatchImageUploadDialog() {
  const apiKey = getImgBBApiKey();
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('ImgBB APIキーが設定されていません。\n設定画面でAPIキーを入力してください。');
    return;
  }
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
      <h3>複数画像一括アップロード</h3>
      <p style="margin-bottom: 20px;">最大10枚の画像を同時にアップロードできます。</p>

      <div style="margin-bottom: 20px;">
        <label for="imageFiles" style="display: block; margin-bottom: 5px; font-weight: bold;">画像ファイルを選択</label>
        <input type="file" id="imageFiles" multiple accept="image/*" 
               style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;">
        <p style="font-size: 12px; color: #666; margin-top: 5px;">
          ※ 複数選択可能（最大10枚、対応形式: JPG, PNG, GIF, WebP）
        </p>
      </div>
      
      <div id="fileList" style="margin-bottom: 20px; display: none;">
        <h4>選択されたファイル:</h4>
        <div id="fileItems" style="max-height: 200px; overflow-y: auto; border: 1px solid #ddd; border-radius: 4px; padding: 10px;"></div>
      </div>

      <div style="text-align: center;">
        <button onclick="uploadImages()" id="uploadBtn" disabled
                style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px;">
          アップロード開始
        </button>
        <button onclick="google.script.host.close()" 
                style="background-color: #f44336; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
          キャンセル
        </button>
      </div>
    </div>

    <script>
      let selectedFiles = [];
      
      document.getElementById('imageFiles').addEventListener('change', function(e) {
        selectedFiles = Array.from(e.target.files);
        updateFileList();
      });
      
      function updateFileList() {
        const fileList = document.getElementById('fileList');
        const fileItems = document.getElementById('fileItems');
        const uploadBtn = document.getElementById('uploadBtn');
        
        if (selectedFiles.length === 0) {
          fileList.style.display = 'none';
          uploadBtn.disabled = true;
          return;
        }
        
        if (selectedFiles.length > 10) {
          alert('一度にアップロードできる画像は最大10枚までです。');
          selectedFiles = selectedFiles.slice(0, 10);
        }
        
        fileList.style.display = 'block';
        uploadBtn.disabled = false;
        
        fileItems.innerHTML = selectedFiles.map((file, index) => 
          \`<div style="padding: 5px; border-bottom: 1px solid #eee; display: flex; align-items: center; gap: 10px;">
            <span style="font-weight: bold; color: #666; min-width: 30px;">\${index + 1}.</span>
            <span style="flex: 1;">\${file.name}</span>
            <span style="color: #666; font-size: 12px;">(\${(file.size / 1024 / 1024).toFixed(2)} MB)</span>
          </div>\`
        ).join('');
      }
      
      function uploadImages() {
        if (selectedFiles.length === 0) {
          alert('画像ファイルを選択してください。');
          return;
        }
        
        const uploadBtn = document.getElementById('uploadBtn');
        uploadBtn.disabled = true;
        uploadBtn.textContent = 'アップロード中...';
        
        // ファイルをBase64に変換して送信
        const promises = selectedFiles.map(file => {
          return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = function(e) {
              const base64 = e.target.result.split(',')[1];
              const blob = Utilities.newBlob(Utilities.base64Decode(base64), file.type, file.name);
              resolve(blob);
            };
            reader.readAsDataURL(file);
          });
        });
        
        Promise.all(promises).then(blobs => {
          google.script.run
            .withSuccessHandler(onUploadSuccess)
            .withFailureHandler(onUploadFailure)
            .uploadMultipleImagesToImgBB(blobs);
        });
      }
      
      function onUploadSuccess(result) {
        google.script.run
          .withSuccessHandler(showUploadResult)
          .generateBatchUploadResultHtml(result);
      }
      
      function showUploadResult(html) {
        const newDialog = HtmlService.createHtmlOutput(html);
        SpreadsheetApp.getUi().showModalDialog(newDialog, 'アップロード完了');
      }
      
      function onUploadFailure(error) {
        alert('アップロードに失敗しました: ' + error.message);
        const uploadBtn = document.getElementById('uploadBtn');
        uploadBtn.disabled = false;
        uploadBtn.textContent = 'アップロード開始';
      }
    </script>
  `);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '複数画像一括アップロード');
}

/**
 * 画像アップロード機能のテスト
 */
function testImageUpload() {
  const apiKey = getImgBBApiKey();
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('ImgBB APIキーが設定されていません。\n設定画面でAPIキーを入力してください。');
    return;
  }
  
  // テスト用の小さな画像データを作成（1x1ピクセルの透明PNG）
  const testImageData = Utilities.base64Decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==');
  const testBlob = Utilities.newBlob(testImageData, 'image/png', 'test.png');
  
  const result = uploadImageToImgBB(testBlob, 'test.png');
  
  if (result.success) {
    const htmlOutput = HtmlService.createHtmlOutput(generateImageUrlDisplayHtml(result.url));
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'テストアップロード完了');
  } else {
    SpreadsheetApp.getUi().alert('テストアップロードに失敗しました:\n' + result.error);
  }
}
