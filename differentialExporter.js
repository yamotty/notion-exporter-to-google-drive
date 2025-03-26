/**
 * 差分エクスポート機能 - 新規/更新されたNotionページのみを処理
 */

/**
 * 差分エクスポート用のプロパティキー
 */
const DIFF_PROPS = {
    LAST_EXPORT_TIMESTAMP: 'LAST_EXPORT_TIMESTAMP',
    PROCESSED_PAGES: 'PROCESSED_PAGES_INFO'
  };
  
  /**
   * 差分エクスポートの実行関数
   * 新規または更新されたページのみをエクスポート
   */
  function exportDifferentialNotionPages() {
    // 設定を読み込む
    if (!loadSettings()) {
      return;
    }
    
    // APIキーが設定されているか確認
    if (!NOTION_API_KEY) {
      showAlert('APIキーが設定されていません', '「API Keyを設定」メニューからNotionのAPIキーを設定してください。');
      return;
    }
    
    // 処理結果を格納する配列
    const results = [];
    
    // プロパティサービスから前回のエクスポート情報を取得
    const properties = PropertiesService.getScriptProperties();
    let lastExportTimestamp = properties.getProperty(DIFF_PROPS.LAST_EXPORT_TIMESTAMP);
    let processedPagesInfo = properties.getProperty(DIFF_PROPS.PROCESSED_PAGES);
    
    // 処理済みページ情報のマップを作成（ページID -> {lastModified, docId}）
    let processedPages = {};
    if (processedPagesInfo) {
      try {
        processedPages = JSON.parse(processedPagesInfo);
      } catch (e) {
        Logger.log('処理済みページ情報の解析に失敗しました: ' + e.message);
        processedPages = {};
      }
    }
    
    // Notionデータベースからすべてのページのリストを取得
    const pages = getNotionDatabasePages(DATABASE_ID);
    
    if (!pages || pages.length === 0) {
      showAlert('データが見つかりません', 'データベースからページが見つかりませんでした。データベースIDを確認してください。');
      return;
    }
    
    // 現在の時刻（新しいタイムスタンプとして使用）
    const currentTimestamp = new Date().toISOString();
    
    // 処理する必要があるページを特定（新規または更新されたページ）
    const pagesToProcess = [];
    const skippedPages = [];
    
    for (const page of pages) {
      const pageId = page.id;
      const lastEditedTime = page.last_edited_time;
      
      // このページが前回の処理後に更新されたか、または新規ページかチェック
      const isNewOrUpdated = !processedPages[pageId] || 
                             (lastEditedTime && processedPages[pageId].lastModified < lastEditedTime);
      
      if (isNewOrUpdated) {
        pagesToProcess.push(page);
      } else {
        skippedPages.push({
          pageId: pageId,
          pageTitle: getPageTitle(page),
          status: 'Skipped',
          message: '前回のエクスポート以降更新がないためスキップされました',
          timestamp: new Date().toISOString(),
          fileId: processedPages[pageId].docId
        });
      }
    }
    
    // 進捗状況表示用のインジケータを作成
    try {
      const ui = SpreadsheetApp.getUi();
      showAlert('差分エクスポートを開始します', 
                `全${pages.length}ページ中、${pagesToProcess.length}ページを処理します（${skippedPages.length}ページはスキップ）。`);
    } catch (e) {
      Logger.log('UIが利用できないため、進捗表示はスキップします');
    }
    
    Logger.log(`${pagesToProcess.length}ページを処理します（${skippedPages.length}ページはスキップ）...`);
    
    // 各ページを処理
    for (const page of pagesToProcess) {
      try {
        // ページタイトルを取得
        const pageId = page.id;
        const pageTitle = getPageTitle(page);
        const lastEditedTime = page.last_edited_time;
        
        // 処理結果オブジェクトを初期化
        const result = {
          pageId: pageId,
          pageTitle: pageTitle,
          status: 'Processing',
          message: '',
          timestamp: new Date().toISOString()
        };
        
        try {
          // ページのブロック（コンテンツ）を取得
          const blocks = getPageBlocks(pageId);
          
          // 既存のドキュメントIDがあるか確認（更新の場合）
          let existingDocId = processedPages[pageId] ? processedPages[pageId].docId : null;
          
          // ブロックからGoogle Docsを生成して保存（既存のドキュメントがあれば更新）
          const saveResult = convertBlocksToGoogleDocs(blocks, pageTitle, existingDocId);
          
          if (saveResult.success) {
            result.status = 'Success';
            result.message = saveResult.message;
            result.fileId = saveResult.fileId;
            
            // 処理済みページの情報を更新
            processedPages[pageId] = {
              lastModified: lastEditedTime || currentTimestamp,
              docId: saveResult.fileId,
              title: pageTitle
            };
          } else {
            result.status = 'Fail';
            result.message = saveResult.message;
          }
        } catch (error) {
          result.status = 'Fail';
          result.message = `エラー: ${error.message}`;
        }
        
        // 結果を配列に追加
        results.push(result);
        
        // ログに出力
        Logger.log(`ページ "${pageTitle}": ${result.status} - ${result.message}`);
      } catch (error) {
        // ページタイトル取得時のエラー処理
        results.push({
          pageId: page.id,
          pageTitle: `Unknown (ID: ${page.id})`,
          status: 'Fail',
          message: `ページ情報の取得中にエラーが発生: ${error.message}`,
          timestamp: new Date().toISOString()
        });
        
        Logger.log(`ページID ${page.id} の処理中にエラーが発生しました: ${error.message}`);
      }
    }
    
    // スキップしたページの情報も結果に追加
    results.push(...skippedPages);
    
    // 処理済みページの情報を保存
    properties.setProperty(DIFF_PROPS.PROCESSED_PAGES, JSON.stringify(processedPages));
    properties.setProperty(DIFF_PROPS.LAST_EXPORT_TIMESTAMP, currentTimestamp);
    
    // 全体の処理結果をログに出力
    const successCount = results.filter(r => r.status === 'Success').length;
    const failCount = results.filter(r => r.status === 'Fail').length;
    const skippedCount = results.filter(r => r.status === 'Skipped').length;
    
    Logger.log('===== 差分エクスポート結果サマリー =====');
    Logger.log(`処理総数: ${results.length}`);
    Logger.log(`成功: ${successCount}`);
    Logger.log(`失敗: ${failCount}`);
    Logger.log(`スキップ: ${skippedCount}`);
    Logger.log('======================================');
    
    // スプレッドシートに結果を記録
    recordResultsToSpreadsheet(results);
    
    // 処理完了のアラート
    try {
      showAlert(
        '差分エクスポート完了',
        `全体: ${results.length}ページ\n` +
        `成功: ${successCount}ページ\n` +
        `失敗: ${failCount}ページ\n` +
        `スキップ: ${skippedCount}ページ\n\n` +
        `詳細は「エクスポート結果」シートをご確認ください。`
      );
      
      // 結果シートをアクティブにする
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const resultSheet = ss.getSheetByName('エクスポート結果');
      if (resultSheet) {
        ss.setActiveSheet(resultSheet);
      }
    } catch (e) {
      Logger.log('UIが利用できないため、完了アラートはスキップします');
    }
    
    return results;
  }
  
  /**
   * 差分エクスポートの情報をリセット
   * すべてのページを新規として扱うようにする
   */
  function resetDifferentialExportData() {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty(DIFF_PROPS.LAST_EXPORT_TIMESTAMP);
    properties.deleteProperty(DIFF_PROPS.PROCESSED_PAGES);
    
    showAlert('差分エクスポート情報をリセットしました', 
             '次回の差分エクスポート実行時、すべてのページが新規として処理されます。');
  }