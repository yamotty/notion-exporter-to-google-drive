/**
 * バッチ処理のためのファイル
 * Google Apps Scriptの実行時間制限（6分）に対処するためのバッチ処理機能
 */

/**
 * バッチ処理のためのプロパティキー
 */
const BATCH_PROPS = {
    BATCH_IN_PROGRESS: 'BATCH_IN_PROGRESS',
    CURRENT_BATCH_INDEX: 'CURRENT_BATCH_INDEX',
    TOTAL_PAGES: 'TOTAL_PAGES',
    PROCESSED_PAGES: 'PROCESSED_PAGES',
    BATCH_PAGE_IDS: 'BATCH_PAGE_IDS',
    BATCH_STARTED_AT: 'BATCH_STARTED_AT',
    BATCH_RESULTS: 'BATCH_RESULTS'
  };
  
  /**
   * バッチサイズ - 1回の実行で処理するページ数
   * 実行環境や複雑さに応じて調整してください
   */
  const BATCH_SIZE = 5;
  
  /**
   * バッチ処理のメイン関数
   * バッチごとに処理を行い、必要に応じてトリガーを設定して次のバッチを処理
   */
  function processBatchExport() {
    const properties = PropertiesService.getScriptProperties();
    
    // 設定を読み込む
    if (!loadSettings()) {
      properties.deleteAllProperties(); // バッチ処理をキャンセル
      showAlert('設定エラー', 'バッチ処理をキャンセルしました。設定を確認してください。');
      return;
    }
    
    // バッチ処理の状態を確認
    let batchInProgress = properties.getProperty(BATCH_PROPS.BATCH_IN_PROGRESS) === 'true';
    let currentBatchIndex = parseInt(properties.getProperty(BATCH_PROPS.CURRENT_BATCH_INDEX) || '0');
    let totalPages = parseInt(properties.getProperty(BATCH_PROPS.TOTAL_PAGES) || '0');
    let processedPages = parseInt(properties.getProperty(BATCH_PROPS.PROCESSED_PAGES) || '0');
    let batchStartedAt = properties.getProperty(BATCH_PROPS.BATCH_STARTED_AT);
    let batchResults = properties.getProperty(BATCH_PROPS.BATCH_RESULTS);
    batchResults = batchResults ? JSON.parse(batchResults) : [];
    
    // 初回実行か、継続実行かを判断
    if (!batchInProgress) {
      // 初回実行: バッチ処理を初期化
      const pages = getNotionDatabasePages(DATABASE_ID);
      
      if (!pages || pages.length === 0) {
        showAlert('データが見つかりません', 'データベースからページが見つかりませんでした。データベースIDを確認してください。');
        properties.deleteAllProperties();
        return;
      }
      
      // ページIDのリストを保存
      const pageIds = pages.map(page => page.id);
      totalPages = pageIds.length;
      
      // バッチ処理のステータスを初期化
      properties.setProperty(BATCH_PROPS.BATCH_IN_PROGRESS, 'true');
      properties.setProperty(BATCH_PROPS.CURRENT_BATCH_INDEX, '0');
      properties.setProperty(BATCH_PROPS.TOTAL_PAGES, totalPages.toString());
      properties.setProperty(BATCH_PROPS.PROCESSED_PAGES, '0');
      properties.setProperty(BATCH_PROPS.BATCH_PAGE_IDS, JSON.stringify(pageIds));
      properties.setProperty(BATCH_PROPS.BATCH_STARTED_AT, new Date().toISOString());
      properties.setProperty(BATCH_PROPS.BATCH_RESULTS, JSON.stringify([]));
      
      // ステータスを表示
      showAlert('バッチ処理を開始します', `全${totalPages}ページを${BATCH_SIZE}ページずつ処理します。\n処理が完了するまでスプレッドシートを開いたままにしてください。`);
      
      // ユーザーへのフィードバック用にバッチステータスシートを作成または更新
      updateBatchStatusSheet(0, totalPages, []);
      
      // バッチ変数を更新
      currentBatchIndex = 0;
      batchStartedAt = new Date().toISOString();
      batchResults = [];
    }
    
    // ページIDのリストを取得
    const pageIds = JSON.parse(properties.getProperty(BATCH_PROPS.BATCH_PAGE_IDS));
    
    // 現在のバッチのインデックス範囲を計算
    const startIndex = currentBatchIndex * BATCH_SIZE;
    const endIndex = Math.min(startIndex + BATCH_SIZE, totalPages);
    
    // このバッチで処理するページID
    const batchPageIds = pageIds.slice(startIndex, endIndex);
    
    // このバッチの結果を保存する配列
    const results = [];
    
    // 各ページを処理
    for (const [index, pageId] of batchPageIds.entries()) {
      try {
        // ページ情報を取得
        const page = getNotionPageById(pageId);
        if (!page) {
          const result = {
            pageId: pageId,
            pageTitle: `Unknown (ID: ${pageId})`,
            status: 'Fail',
            message: 'ページ情報を取得できませんでした',
            timestamp: new Date().toISOString()
          };
          results.push(result);
          
          // ページごとにステータスを更新
          processedPages++;
          properties.setProperty(BATCH_PROPS.PROCESSED_PAGES, processedPages.toString());
          
          // ステータスシートを更新 (最新の処理結果のみを渡す)
          updateBatchStatusSheet(processedPages, totalPages, [result]);
          continue;
        }
        
        const pageTitle = getPageTitle(page);
        
        // 処理結果オブジェクトを初期化
        const result = {
          pageId: pageId,
          pageTitle: pageTitle,
          status: 'Processing',
          message: '',
          timestamp: new Date().toISOString()
        };
        
        // 処理中のステータスを表示
        updateBatchStatusSheet(processedPages, totalPages, [{...result, message: '処理中...'}]);
        
        try {
          // ページのブロック（コンテンツ）を取得
          const blocks = getPageBlocks(pageId);
          
          // ブロックからGoogle Docsを直接生成して保存
          const saveResult = convertBlocksToGoogleDocs(blocks, pageTitle);
          
          if (saveResult.success) {
            result.status = 'Success';
            result.message = saveResult.message;
            result.fileId = saveResult.fileId;
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
        processedPages++;
        
        // プロパティをリアルタイムで更新
        properties.setProperty(BATCH_PROPS.PROCESSED_PAGES, processedPages.toString());
        
        // ステータスシートを更新 (最新の処理結果のみを渡す)
        updateBatchStatusSheet(processedPages, totalPages, [result]);
        
        // ログに出力
        Logger.log(`ページ "${pageTitle}": ${result.status} - ${result.message}`);
      } catch (error) {
        const result = {
          pageId: pageId,
          pageTitle: `Unknown (ID: ${pageId})`,
          status: 'Fail',
          message: `ページ処理中にエラーが発生: ${error.message}`,
          timestamp: new Date().toISOString()
        };
        results.push(result);
        processedPages++;
        
        // プロパティをリアルタイムで更新
        properties.setProperty(BATCH_PROPS.PROCESSED_PAGES, processedPages.toString());
        
        // ステータスシートを更新 (最新の処理結果のみを渡す)
        updateBatchStatusSheet(processedPages, totalPages, [result]);
        
        Logger.log(`ページID ${pageId} の処理中にエラーが発生しました: ${error.message}`);
      }
    }
    
    // バッチ結果を結合
    batchResults = batchResults.concat(results);
    
    // プロパティを更新
    currentBatchIndex++;
    properties.setProperty(BATCH_PROPS.CURRENT_BATCH_INDEX, currentBatchIndex.toString());
    properties.setProperty(BATCH_PROPS.PROCESSED_PAGES, processedPages.toString());
    properties.setProperty(BATCH_PROPS.BATCH_RESULTS, JSON.stringify(batchResults));
    
    // スプレッドシートに結果を記録
    recordResultsToSpreadsheet(results);
    
    // 全てのバッチが完了したかチェック
    if (endIndex >= totalPages) {
      // バッチ処理完了
      completeAndSummary(batchResults, batchStartedAt);
      
      // バッチ処理のプロパティをクリア
      properties.deleteProperty(BATCH_PROPS.BATCH_IN_PROGRESS);
      properties.deleteProperty(BATCH_PROPS.CURRENT_BATCH_INDEX);
      properties.deleteProperty(BATCH_PROPS.TOTAL_PAGES);
      properties.deleteProperty(BATCH_PROPS.PROCESSED_PAGES);
      properties.deleteProperty(BATCH_PROPS.BATCH_PAGE_IDS);
      properties.deleteProperty(BATCH_PROPS.BATCH_STARTED_AT);
      properties.deleteProperty(BATCH_PROPS.BATCH_RESULTS);
      
      // トリガーを削除
      deleteTriggers();
    } else {
      // 次のバッチのトリガーを設定
      createTriggerForNextBatch();
    }
  }
  
  /**
   * バッチ処理のステータスをシートに表示
   * @param {number} processed - 処理済みページ数
   * @param {number} total - 合計ページ数
   * @param {Array} latestResults - 最新の処理結果
   */
  function updateBatchStatusSheet(processed, total, latestResults) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Logger.log('スプレッドシートが見つかりません。');
        return;
      }
      
      // 「バッチステータス」という名前のシートを探す、なければ作成
      let sheet = ss.getSheetByName('バッチステータス');
      if (!sheet) {
        sheet = ss.insertSheet('バッチステータス');
        
        // 初期設定
        sheet.getRange('A1:B1').setValues([['バッチエクスポートステータス', '']]).setFontWeight('bold');
        sheet.getRange('A2:B2').setValues([['開始時間', new Date().toLocaleString()]]);
        sheet.getRange('A3:B3').setValues([['処理状況', `0 / ${total} ページ (0%)`]]);
        sheet.getRange('A4:B4').setValues([['進行状況', '░'.repeat(20)]]);
        
        // テーブルヘッダーを設定
        sheet.getRange('A6:E6').setValues([['ページタイトル', 'ステータス', 'メッセージ', 'ドキュメントリンク', 'タイムスタンプ']]).setFontWeight('bold');
        
        // 書式設定
        sheet.autoResizeColumns(1, 5);
        sheet.setFrozenRows(6); // ヘッダー行を固定
      }
      
      // 現在時刻を取得
      const currentTime = new Date().toLocaleString();
      
      // 処理状況の更新
      const percent = Math.round(processed / total * 100);
      sheet.getRange('B2').setValue(currentTime);
      sheet.getRange('B3').setValue(`${processed} / ${total} ページ (${percent}%)`);
      
      // プログレスバーを更新
      const progressBarWidth = 20; // プログレスバーの長さ
      const progressFilled = Math.round(processed / total * progressBarWidth);
      const progressBar = '▓'.repeat(progressFilled) + '░'.repeat(progressBarWidth - progressFilled);
      sheet.getRange('B4').setValue(progressBar);
      
      // 最新の処理結果を追加
      if (latestResults && latestResults.length > 0) {
        // 現在の結果の行数を取得
        const lastRow = Math.max(sheet.getLastRow(), 6);
        
        // 結果データを準備
        const resultRows = latestResults.map((result) => {
          let link = '';
          if (result.fileId) {
            link = `https://docs.google.com/document/d/${result.fileId}/edit`;
          }
          
          return [
            result.pageTitle,
            result.status,
            result.message,
            link ? '=HYPERLINK("' + link + '","開く")' : '',
            new Date(result.timestamp || new Date()).toLocaleString()
          ];
        });
        
        // 最新のデータを追加
        const startRow = lastRow + 1;
        sheet.getRange(startRow, 1, resultRows.length, 5).setValues(resultRows);
        
        // 条件付き書式を設定
        const statusColumn = 2; // 「ステータス」列のインデックス
        const dataRange = sheet.getRange(7, 1, Math.max(startRow + resultRows.length - 7, 1), 5);
        const statusRange = sheet.getRange(7, statusColumn, Math.max(startRow + resultRows.length - 7, 1), 1);
        
        // 既存の条件付き書式をクリア
        sheet.clearConditionalFormatRules();
        
        // 条件付き書式を追加
        const successRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Success')
          .setBackground('#b7e1cd')
          .setRanges([statusRange])
          .build();
        
        const processingRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Processing')
          .setBackground('#fff2cc')
          .setRanges([statusRange])
          .build();
        
        const failRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Fail')
          .setBackground('#f4c7c3')
          .setRanges([statusRange])
          .build();
        
        sheet.setConditionalFormatRules([successRule, processingRule, failRule]);
        
        // シートを見やすく整形
        sheet.autoResizeColumns(1, 5);
      }
      
    } catch (error) {
      Logger.log('バッチステータスシートの更新に失敗しました: ' + error.message);
    }
  }
  
  /**
   * バッチ処理の開始関数
   * メインメニューから呼び出される
   */
  function startBatchExport() {
    // APIキーが設定されているか確認
    if (!NOTION_API_KEY) {
      showAlert('APIキーが設定されていません', '「API Keyを設定」メニューからNotionのAPIキーを設定してください。');
      return;
    }
    
    // 設定を読み込む
    if (!loadSettings()) {
      return;
    }
    
    // 既存のバッチ処理をクリア（万が一の場合のため）
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty(BATCH_PROPS.BATCH_IN_PROGRESS);
    
    // 既存のトリガーを削除
    deleteTriggers();
    
    // 初回バッチを即時実行
    processBatchExport();
  }
  
  /**
   * NotionページIDからページ情報を取得
   * @param {string} pageId - NotionページID
   * @return {Object} ページ情報
   */
  function getNotionPageById(pageId) {
    try {
      const url = `https://api.notion.com/v1/pages/${pageId}`;
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${NOTION_API_KEY}`,
          'Notion-Version': '2022-06-28'
        }
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseData = JSON.parse(response.getContentText());
      
      return responseData;
    } catch (error) {
      Logger.log(`ページ情報の取得中にエラーが発生しました: ${error.message}`);
      return null;
    }
  }
  
  /**
   * 次のバッチ処理のためのトリガーを作成
   */
  function createTriggerForNextBatch() {
    // 既存のトリガーを削除
    deleteTriggers();
    
    // 1分後に実行するトリガーを作成
    ScriptApp.newTrigger('processBatchExport')
      .timeBased()
      .after(1 * 60 * 1000) // 1分後
      .create();
  }
  
  /**
   * 既存のトリガーを全て削除
   */
  function deleteTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'processBatchExport') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  }
  
  /**
   * バッチ処理が完了した際の集計とサマリー表示
   * @param {Array} results - 全バッチの結果
   * @param {string} startedAt - 開始時刻のISOフォーマット文字列
   */
  function completeAndSummary(results, startedAt) {
    // 成功・失敗の集計
    const successCount = results.filter(r => r.status === 'Success').length;
    const failCount = results.filter(r => r.status === 'Fail').length;
    
    // 処理時間の計算
    const startTime = new Date(startedAt);
    const endTime = new Date();
    const durationMs = endTime - startTime;
    const durationMin = Math.floor(durationMs / 60000);
    const durationSec = Math.floor((durationMs % 60000) / 1000);
    
    // ログに出力
    Logger.log('===== バッチエクスポート結果サマリー =====');
    Logger.log(`処理総数: ${results.length}`);
    Logger.log(`成功: ${successCount}`);
    Logger.log(`失敗: ${failCount}`);
    Logger.log(`処理時間: ${durationMin}分${durationSec}秒`);
    Logger.log('=========================================');
    
    // アラート表示
    try {
      showAlert(
        'バッチエクスポート完了',
        `処理総数: ${results.length}\n` +
        `成功: ${successCount}\n` +
        `失敗: ${failCount}\n` +
        `処理時間: ${durationMin}分${durationSec}秒\n\n` +
        `詳細は「エクスポート結果」シートをご確認ください。`
      );
      
      // バッチステータスシートに完了メッセージを追加
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const statusSheet = ss.getSheetByName('バッチステータス');
        if (statusSheet) {
          const lastRow = statusSheet.getLastRow() + 2;
          statusSheet.getRange(lastRow, 1, 1, 5).merge();
          statusSheet.getRange(lastRow, 1).setValue(
            `完了: 合計${results.length}ページ、成功${successCount}、失敗${failCount}、処理時間${durationMin}分${durationSec}秒`
          ).setFontWeight('bold').setHorizontalAlignment('center');
        }
      } catch (e) {
        Logger.log('完了メッセージの追加に失敗しました: ' + e.message);
      }
      
      // 結果シートをアクティブにする
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const resultSheet = ss.getSheetByName('エクスポート結果');
      if (resultSheet) {
        ss.setActiveSheet(resultSheet);
      }
    } catch (e) {
      Logger.log('UIが利用できないため、完了アラートはスキップします');
    }
  }
  
  /**
   * バッチ処理の強制停止
   */
  function cancelBatchExport() {
    // トリガーを削除
    deleteTriggers();
    
    // バッチ処理のプロパティをクリア
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty(BATCH_PROPS.BATCH_IN_PROGRESS);
    properties.deleteProperty(BATCH_PROPS.CURRENT_BATCH_INDEX);
    properties.deleteProperty(BATCH_PROPS.TOTAL_PAGES);
    properties.deleteProperty(BATCH_PROPS.PROCESSED_PAGES);
    properties.deleteProperty(BATCH_PROPS.BATCH_PAGE_IDS);
    properties.deleteProperty(BATCH_PROPS.BATCH_STARTED_AT);
    properties.deleteProperty(BATCH_PROPS.BATCH_RESULTS);
    
    showAlert('バッチ処理を停止しました', '現在の実行は停止されました。既に処理されたページの結果は「エクスポート結果」シートで確認できます。');
  }