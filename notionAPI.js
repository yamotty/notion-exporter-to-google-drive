/**
 * Notionデータベースからページのリストを取得
 * @param {string} databaseId - NotionデータベースID
 * @return {Array} ページの配列
 */
function getNotionDatabasePages(databaseId) {
    try {
      // データベースクエリのエンドポイント
      const url = `https://api.notion.com/v1/databases/${databaseId}/query`;
      
      // APIリクエストのオプション
      const options = {
        method: 'post',
        headers: {
          'Authorization': `Bearer ${NOTION_API_KEY}`,
          'Notion-Version': '2022-06-28', // 現在のAPI版数に合わせて更新してください
          'Content-Type': 'application/json'
        },
        // 必要に応じてフィルタやソートを追加できます
        payload: JSON.stringify({
          page_size: 100 // 一度に取得するページ数を調整
        })
      };
      
      // APIリクエストを実行
      const response = UrlFetchApp.fetch(url, options);
      const responseData = JSON.parse(response.getContentText());
      
      return responseData.results;
    } catch (error) {
      Logger.log(`データベースからページの取得中にエラーが発生しました: ${error.message}`);
      Logger.log(`エラータイプ: ${error.name}`);
      Logger.log(`エラースタック: ${error.stack}`);
      return [];
    }
  }
  
  /**
   * Notionページからブロック（コンテンツ）を取得
   * @param {string} pageId - NotionページID
   * @return {Array} ブロックの配列
   */
  function getPageBlocks(pageId) {
    try {
      const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${NOTION_API_KEY}`,
          'Notion-Version': '2022-06-28'
        }
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseData = JSON.parse(response.getContentText());
      
      // デバッグ情報を追加
      Logger.log(`ブロック総数: ${responseData.results.length}`);
      if (responseData.results.length > 0) {
        Logger.log(`最初のブロックタイプ: ${responseData.results[0].type}`);
      }
      
      return responseData.results;
    } catch (error) {
      Logger.log(`ページブロックの取得中にエラーが発生しました: ${error.message}`);
      Logger.log(`エラースタック: ${error.stack}`);
      return [];
    }
  }
  
  /**
   * NotionページからページタイトルまたはName/Titleプロパティを取得
   * @param {Object} page - Notionページオブジェクト
   * @return {string} ページタイトル
   */
  function getPageTitle(page) {
    try {
      // ページのプロパティを確認
      const properties = page.properties;
      
      // デバッグ情報
      Logger.log(`ページプロパティのキー: ${Object.keys(properties).join(', ')}`);
      
      // 一般的にName, Title, または他のタイトルプロパティを探す
      for (const propName in properties) {
        const prop = properties[propName];
        Logger.log(`プロパティ ${propName} のタイプ: ${prop.type}`);
        
        if (prop.type === 'title' && prop.title && prop.title.length > 0) {
          const title = prop.title.map(textObj => textObj.plain_text).join('');
          Logger.log(`タイトルを見つけました: ${title}`);
          return title;
        }
      }
      
      // タイトルが見つからない場合はページIDを使用
      const fallbackTitle = `Page_${page.id.replace(/-/g, '').substring(0, 8)}`;
      Logger.log(`タイトルが見つからないため代替タイトルを使用: ${fallbackTitle}`);
      return fallbackTitle;
    } catch (error) {
      Logger.log(`ページタイトルの取得中にエラーが発生しました: ${error.message}`);
      Logger.log(`エラースタック: ${error.stack}`);
      return `Page_${page.id.replace(/-/g, '').substring(0, 8)}`;
    }
  }
  
  /**
   * スクリプトプロパティを設定するヘルパー関数
   */
  function setApiKey() {
    const ui = SpreadsheetApp.getUi();
    const prompt = ui.prompt('Notion API Key', 'Notion API Keyを入力してください:', ui.ButtonSet.OK_CANCEL);
    
    if (prompt.getSelectedButton() === ui.Button.OK) {
      const apiKey = prompt.getResponseText();
      PropertiesService.getScriptProperties().setProperty('NOTION_API_KEY', apiKey);
      ui.alert('API Keyが保存されました。', 'Notion APIキーがスクリプトプロパティに保存されました。', ui.ButtonSet.OK);
    }
  }