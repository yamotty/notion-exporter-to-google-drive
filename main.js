/**
 * NotionデータベースのドキュメントをGoogle Docs形式でGoogle Driveに保存するスクリプト
 * スプレッドシートと連携して動作し、設定や結果の管理を行います
 * 
 * 使用前に以下の設定が必要です:
 * 1. Notion APIの統合をセットアップしてAPIキーを取得
 * 2. このスクリプトのプロパティに'NOTION_API_KEY'としてAPIキーを保存
 * 3. スプレッドシートの「設定」シートに対象のNotionデータベースIDと保存先フォルダIDを設定
 */

// スクリプトのプロパティからNotion APIキーを取得
const NOTION_API_KEY = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
// 設定シートから情報を読み込み
let DATABASE_ID = '';
let DRIVE_FOLDER_ID = '';

/**
 * スプレッドシートが開かれたときに実行される関数
 * カスタムメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Notion Export')
    .addItem('API Keyを設定', 'setApiKey')
    .addItem('必要なサービスを確認', 'checkRequiredServices')
    .addItem('設定シートを初期化', 'initSettingsSheet')
    .addSeparator()
    .addItem('全ページをエクスポート', 'exportNotionToGoogleDrive')
    .addItem('差分エクスポート（新規/更新ページのみ）', 'exportDifferentialNotionPages')
    .addItem('差分エクスポート情報をリセット', 'resetDifferentialExportData')
    .addSeparator()
    .addItem('バッチ処理でエクスポート（大量データ用）', 'startBatchExport')
    .addItem('バッチ処理を停止', 'cancelBatchExport')
    .addItem('デバッグ: 最初の1アイテムだけエクスポート', 'exportFirstNotionItem')
    .addSeparator()
    .addItem('エクスポート結果を表示', 'showExportResults')
    .addItem('使い方ガイド', 'showHelpGuide')
    .addToUi();
}

/**
 * メイン実行関数
 * 処理結果のログを表示する
 */
function exportNotionToGoogleDrive() {
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
  
  // Notionデータベースからページのリストを取得
  const pages = getNotionDatabasePages(DATABASE_ID);
  
  if (!pages || pages.length === 0) {
    showAlert('データが見つかりません', 'データベースからページが見つかりませんでした。データベースIDを確認してください。');
    return;
  }
  
  // 進捗状況表示用のインジケータを作成
  try {
    const ui = SpreadsheetApp.getUi();
    const totalPages = pages.length;
    showAlert('処理を開始します', `${totalPages}ページをエクスポートします。処理中はスプレッドシートを閉じないでください。`);
  } catch (e) {
    Logger.log('UIが利用できないため、進捗表示はスキップします');
  }
  
  Logger.log(`${pages.length}ページをエクスポートします...`);
  
  // 各ページを処理
  for (const page of pages) {
    try {
      // ページタイトルを取得（通常はNameやTitleプロパティに格納されています）
      const pageTitle = getPageTitle(page);
      
      // 処理結果オブジェクトを初期化
      const result = {
        pageId: page.id,
        pageTitle: pageTitle,
        status: 'Processing',
        message: '',
        timestamp: new Date().toISOString()
      };
      
      try {
        // ページのブロック（コンテンツ）を取得
        const blocks = getPageBlocks(page.id);
        
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
  
  // 全体の処理結果をログに出力
  const successCount = results.filter(r => r.status === 'Success').length;
  const failCount = results.filter(r => r.status === 'Fail').length;
  
  Logger.log('===== エクスポート結果サマリー =====');
  Logger.log(`処理総数: ${results.length}`);
  Logger.log(`成功: ${successCount}`);
  Logger.log(`失敗: ${failCount}`);
  Logger.log('=================================');
  
  // 詳細ログ出力
  Logger.log('===== 詳細結果 =====');
  results.forEach((result, index) => {
    Logger.log(`[${index + 1}] ${result.pageTitle}: ${result.status}`);
    if (result.status === 'Fail') {
      Logger.log(`    エラー詳細: ${result.message}`);
    }
  });
  
  // スプレッドシートに結果を記録
  recordResultsToSpreadsheet(results);
  
  // 処理完了のアラート
  try {
    showAlert(
      'エクスポート完了',
      `処理総数: ${results.length}\n成功: ${successCount}\n失敗: ${failCount}\n\n詳細は「エクスポート結果」シートをご確認ください。`
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
 * デバッグ用：NotionデータベースのFirstアイテムだけを処理する
 */
function exportFirstNotionItem() {
  // 設定を読み込む
  if (!loadSettings()) {
    return;
  }
  
  // APIキーが設定されているか確認
  if (!NOTION_API_KEY) {
    showAlert('APIキーが設定されていません', '「API Keyを設定」メニューからNotionのAPIキーを設定してください。');
    return;
  }
  
  // Notionデータベースからページのリストを取得
  const pages = getNotionDatabasePages(DATABASE_ID);
  
  if (!pages || pages.length === 0) {
    showAlert('データが見つかりません', 'データベースからページが見つかりませんでした。データベースIDを確認してください。');
    return;
  }
  
  // 最初のページだけを処理
  const page = pages[0];
  const result = {
    pageId: page.id,
    pageTitle: '',
    status: 'Processing',
    message: '',
    timestamp: new Date().toISOString()
  };
  
  try {
    // ページタイトルを取得
    const pageTitle = getPageTitle(page);
    result.pageTitle = pageTitle;
    
    Logger.log(`デバッグモード: "${pageTitle}" を処理します...`);
    
    // ページのブロック（コンテンツ）を取得
    const blocks = getPageBlocks(page.id);
    Logger.log(`ブロック数: ${blocks.length}`);
    
    // ブロックからGoogle Docsを直接生成
    const saveResult = convertBlocksToGoogleDocs(blocks, pageTitle);
    
    if (saveResult.success) {
      result.status = 'Success';
      result.message = saveResult.message;
      result.fileId = saveResult.fileId;
      
      // ファイルへのリンクを生成
      const fileLink = `https://docs.google.com/document/d/${saveResult.fileId}/edit`;
      Logger.log(`成功: ${saveResult.message}`);
      Logger.log(`ファイルリンク: ${fileLink}`);
    } else {
      result.status = 'Fail';
      result.message = saveResult.message;
      Logger.log(`失敗: ${saveResult.message}`);
      if (saveResult.error) {
        Logger.log(`エラー詳細: ${saveResult.error.stack || saveResult.error.toString()}`);
      }
    }
  } catch (error) {
    result.status = 'Fail';
    result.message = `エラー: ${error.message}`;
    Logger.log(`処理中にエラーが発生: ${error.message}`);
    Logger.log(`エラータイプ: ${error.name}`);
    Logger.log(`エラースタック: ${error.stack}`);
  }
  
  // スプレッドシートに結果を記録
  recordDebugResultToSpreadsheet(result);
  
  // 処理完了のアラート
  try {
    showAlert(
      'デバッグ実行完了',
      `ページ: ${result.pageTitle}\n` +
      `ステータス: ${result.status}\n` +
      `メッセージ: ${result.message}\n\n` +
      `詳細はログを確認してください。`
    );
  } catch (e) {
    Logger.log('UIが利用できないため、完了アラートはスキップします');
    Logger.log(`デバッグ実行完了: ${result.pageTitle} (${result.status})`);
  }
  
  return result;
}