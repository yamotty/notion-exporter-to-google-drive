/**
 * 画像処理に特化した関数群
 * DocsConverter.gsから画像処理部分を分離
 */

/**
 * Notionのブロックにある画像を処理する関数
 * @param {Object} block - Notionブロック
 * @param {Body} body - ドキュメントのbody
 * @param {Folder} imageFolder - 画像フォルダ
 * @param {string} pageId - ページID
 */
function processImageBlock(block, body, imageFolder, pageId) {
  let imageUrl = '';
  let imageId = '';
  let imageCaption = '';
  
  try {
    if (block.image.type === 'external') {
      imageUrl = block.image.external.url;
      
      // 外部画像のURLからハッシュを生成
      const urlHash = Utilities.base64EncodeWebSafe(
        Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, imageUrl)
      ).replace(/[^a-zA-Z0-9]/g, '').substring(0, 16);
      
      imageId = `external_${urlHash}`;
    } else if (block.image.type === 'file') {
      imageUrl = block.image.file.url;
      
      // Notionのファイル構造から一意のIDを抽出
      const urlParts = imageUrl.split('?')[0].split('/');
      const rawImageId = urlParts[urlParts.length - 1];
      
      // 画像IDとしてURLのユニークな部分を使用
      imageId = `notion_${rawImageId.replace(/[^a-zA-Z0-9]/g, '_')}`;
    } else {
      // 他の種類の画像はサポートしていない
      Logger.log(`未サポートの画像タイプ: ${block.image.type}`);
      body.appendParagraph('[サポートされていない画像タイプ]').setItalic(true);
      return;
    }
    
    // キャプションを処理
    if (block.image.caption && block.image.caption.length > 0) {
      imageCaption = block.image.caption.map(textObj => textObj.plain_text).join('');
    }
    
    if (!imageUrl) {
      Logger.log('画像URLが見つかりません');
      body.appendParagraph('[画像URLが見つかりません]').setItalic(true);
      return;
    }
    
    // 現在のブロックIDを画像識別に使用（より一意性を高めるため）
    const blockId = block.id || Utilities.getUuid();
    
    // 画像のMIMEタイプを推測
    let mimeType = guessMimeTypeFromUrl(imageUrl);
    
    // 画像の拡張子を取得
    const extension = mimeTypeToExtension(mimeType);
    
    // 最終的なファイル名を作成（ID + 拡張子）
    const fileName = `${imageId}.${extension}`;
    
    // 詳細なログ出力
    Logger.log(`画像処理: ブロックID=${blockId}, URL=${imageUrl.substring(0, 50)}...`);
    
    // 画像の挿入処理を改善
    const imageBlob = fetchAndCacheImage(imageUrl, fileName, imageFolder, pageId);
    
    if (imageBlob) {
      // 画像サイズ情報をログに記録
      const imageSize = imageBlob.getBytes().length;
      Logger.log(`画像サイズ: ${(imageSize / 1024).toFixed(2)} KB`);
      
      // 画像を挿入
      body.appendImage(imageBlob);
      
      // キャプションがあれば追加
      if (imageCaption) {
        const captionPara = body.appendParagraph(imageCaption);
        captionPara.setAttributes({
          [DocumentApp.Attribute.ITALIC]: true,
          [DocumentApp.Attribute.FONT_SIZE]: 10,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
        });
      }
      
      Logger.log(`画像を正常に挿入しました: ${fileName} (ページID: ${pageId.substring(0, 8)})`);
    } else {
      body.appendParagraph('[画像の挿入に失敗しました]').setItalic(true);
      Logger.log(`画像の取得に失敗: ${imageUrl}`);
    }
  } catch (imageError) {
    body.appendParagraph('[画像の挿入に失敗しました]').setItalic(true);
    Logger.log(`画像の挿入でエラー: ${imageError.message}`);
    
    // スタックトレースがあれば記録
    if (imageError.stack) {
      Logger.log(`エラースタックトレース: ${imageError.stack}`);
    }
  }
}

/**
 * 不要な画像キャッシュをクリーンアップする
 * @param {Folder} imageFolder - 画像フォルダ
 * @param {number} daysToKeep - 保持する日数（デフォルト30日）
 */
function cleanupImageCache(imageFolder, daysToKeep = 30) {
  if (!imageFolder) return;
  
  try {
    const files = imageFolder.getFiles();
    const now = new Date();
    let cleanedCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      const lastUpdated = file.getLastUpdated();
      const daysSinceUpdate = (now - lastUpdated) / (1000 * 60 * 60 * 24);
      
      if (daysSinceUpdate > daysToKeep) {
        file.setTrashed(true);
        cleanedCount++;
      }
    }
    
    Logger.log(`画像キャッシュのクリーンアップ完了: ${cleanedCount}ファイルを削除`);
  } catch (error) {
    Logger.log(`キャッシュクリーンアップ中にエラー: ${error.message}`);
  }
}

/**
 * 共通画像フォルダを取得または作成する
 * @param {string} parentFolderId - 親フォルダID
 * @return {Folder} 共通画像フォルダ
 */
function getOrCreateCommonImageFolder(parentFolderId) {
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const folderName = "NotionImages"; // 固定の共通フォルダ名
  
  // 既存のフォルダを検索
  const folderIterator = parentFolder.getFoldersByName(folderName);
  
  if (folderIterator.hasNext()) {
    // 既存のフォルダを返す
    return folderIterator.next();
  } else {
    // 新しいフォルダを作成
    return parentFolder.createFolder(folderName);
  }
}

/**
 * URLからMIMEタイプを推測する（拡張版）
 * @param {string} url - 画像URL
 * @return {string} MIMEタイプ
 */
function guessMimeTypeFromUrl(url) {
  // URLからクエリ部分を除去
  const baseUrl = url.split('?')[0];
  
  // ファイル拡張子を取得
  const parts = baseUrl.split('.');
  if (parts.length < 2) {
    return 'image/jpeg'; // デフォルト
  }
  
  const extension = parts[parts.length - 1].toLowerCase();
  
  // 拡張子からMIMEタイプを決定
  switch (extension) {
    case 'jpg':
    case 'jpeg':
    case 'jpe':
      return 'image/jpeg';
    case 'png':
      return 'image/png';
    case 'gif':
      return 'image/gif';
    case 'webp':
      return 'image/webp';
    case 'svg':
      return 'image/svg+xml';
    case 'bmp':
      return 'image/bmp';
    case 'tiff':
    case 'tif':
      return 'image/tiff';
    case 'ico':
      return 'image/x-icon';
    case 'heic':
      return 'image/heic';
    case 'heif':
      return 'image/heif';
    case 'avif':
      return 'image/avif';
    default:
      // 拡張子が不明の場合、URLパターンで判断を試みる
      if (url.includes('notion.so') || url.includes('amazonaws.com')) {
        return 'image/jpeg'; // Notionのファイルは多くの場合JPEG
      }
      return 'image/jpeg'; // デフォルトはJPEG
    }
  }
  
  /**
   * 画像をフェッチしてキャッシュする（改善版）
   * @param {string} imageUrl - 画像URL
   * @param {string} fileName - 保存するファイル名のベース
   * @param {Folder} imageFolder - 画像を保存するフォルダ
   * @param {string} pageId - ページID（一意性を確保するため）
   * @return {Blob} 画像Blob
   */
  function fetchAndCacheImage(imageUrl, fileName, imageFolder, pageId) {
    try {
      // より高い一意性を確保するためにURLのクエリパラメータを使用
      const urlWithoutParams = imageUrl.split('?')[0];
      const urlParams = imageUrl.includes('?') ? imageUrl.split('?')[1] : '';
      
      // URLからタイムスタンプやシグネチャを抽出（例: expiryのような有効期限パラメータを使用）
      const hasExpiryParam = urlParams.includes('expiry=') || urlParams.includes('Expires=');
      
      // 完全に一意の識別子を生成（ページID + URL特有の部分 + ファイル名）
      const urlIdentifier = Utilities.base64EncodeWebSafe(
        Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, urlWithoutParams)
      ).substring(0, 12);
      
      // 完全なユニークファイル名を生成
      const uniqueFileName = `${pageId.substring(0, 8)}_${urlIdentifier}_${fileName}`;
      
      // キャッシュ制御フラグ - 有効期限パラメータがある場合はURLが新しくなる可能性がある
      const shouldRefreshCache = hasExpiryParam;
      
      // 既にキャッシュされた画像があるか確認
      if (imageFolder && !shouldRefreshCache) {
        const cachedImageIterator = imageFolder.getFilesByName(uniqueFileName);
        
        if (cachedImageIterator.hasNext()) {
          const cachedFile = cachedImageIterator.next();
          // 最終更新日をチェック - 24時間以内ならキャッシュを使用
          const lastUpdated = cachedFile.getLastUpdated();
          const now = new Date();
          const hoursSinceUpdate = (now - lastUpdated) / (1000 * 60 * 60);
          
          if (hoursSinceUpdate < 24) {
            Logger.log(`キャッシュ済みの画像を使用（${hoursSinceUpdate.toFixed(2)}時間前に更新）: ${uniqueFileName}`);
            return cachedFile.getBlob();
          } else {
            Logger.log(`キャッシュが古いため更新（${hoursSinceUpdate.toFixed(2)}時間経過）: ${uniqueFileName}`);
            // 古いキャッシュは削除せず、上書きする
          }
        }
      }
      
      // 画像をURLからフェッチ
      Logger.log(`画像をフェッチ: ${imageUrl}`);
      const response = UrlFetchApp.fetch(imageUrl, {
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: false
      });
      
      // レスポンスコードをチェック
      const responseCode = response.getResponseCode();
      if (responseCode !== 200) {
        Logger.log(`画像の取得に失敗: HTTP ${responseCode}, URL: ${imageUrl}`);
        
        // キャッシュから古いバージョンを探してみる（フォールバック）
        if (imageFolder) {
          const cachedFallbackIterator = imageFolder.getFilesByName(uniqueFileName);
          if (cachedFallbackIterator.hasNext()) {
            Logger.log(`最新の取得に失敗しましたが、キャッシュから使用: ${uniqueFileName}`);
            return cachedFallbackIterator.next().getBlob();
          }
        }
        
        return null;
      }
      
      // Blobを取得
      const imageBlob = response.getBlob();
      
      // MIMEタイプを確認して画像かどうかを判断
      const contentType = imageBlob.getContentType();
      // データサイズもチェック
      const imageSize = imageBlob.getBytes().length;
      
      if (!contentType.startsWith('image/') || imageSize < 100) { // 最低100バイト以上を期待
        Logger.log(`無効な画像: タイプ=${contentType}, サイズ=${imageSize}バイト, URL: ${imageUrl}`);
        return null;
      }
      
      // 正しいMIMEタイプに基づいてファイル名の拡張子を修正
      const correctExt = mimeTypeToExtension(contentType);
      const finalFileName = uniqueFileName.replace(/\.[^.]+$/, `.${correctExt}`);
      
      // 画像をキャッシュに保存
      if (imageFolder) {
        try {
          // 既存のファイルがあれば削除 (名前の衝突を避けるため)
          const existingFiles = imageFolder.getFilesByName(finalFileName);
          while (existingFiles.hasNext()) {
            existingFiles.next().setTrashed(true);
          }
          
          // キャッシュに保存（一意的なファイル名を使用）
          imageFolder.createFile(imageBlob.setName(finalFileName));
          Logger.log(`画像をキャッシュに保存: ${finalFileName} (${imageSize} バイト)`);
        } catch (cacheError) {
          // キャッシュに保存できなくても続行
          Logger.log(`画像のキャッシュに失敗: ${cacheError.message}`);
        }
      }
      
      return imageBlob;
    } catch (error) {
      Logger.log(`画像フェッチ中のエラー: ${error.message}, URL: ${imageUrl}`);
      return null;
    }
  }
  
  /**
   * MIMEタイプから拡張子を取得
   * @param {string} mimeType - MIMEタイプ
   * @return {string} ファイル拡張子
   */
  function mimeTypeToExtension(mimeType) {
    switch (mimeType) {
      case 'image/jpeg':
        return 'jpg';
      case 'image/png':
        return 'png';
      case 'image/gif':
        return 'gif';
      case 'image/webp':
        return 'webp';
      case 'image/svg+xml':
        return 'svg';
      case 'image/bmp':
        return 'bmp';
      case 'image/tiff':
        return 'tiff';
      default:
        return 'jpg'; // デフォルト
    }
  }