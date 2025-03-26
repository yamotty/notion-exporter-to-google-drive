/**
 * スプレッドシートから設定を読み込む
 * 必要なシートがない場合は作成する
 * @return {boolean} 設定の読み込みに成功したかどうか
 */
function loadSettings() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      showAlert('エラー', 'スプレッドシートが見つかりません。');
      return false;
    }
    
    // 設定シートを探す、なければ作成
    let settingsSheet = ss.getSheetByName('設定');
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('設定');
      
      // 設定シートの初期設定
      settingsSheet.getRange('A1:B1').setValues([['設定項目', '値']]).setFontWeight('bold');
      settingsSheet.getRange('A2:A3').setValues([['NotionデータベースID'], ['Google DriveフォルダID']]);
      settingsSheet.getRange('B2:B3').setValues([['ここにデータベースIDを入力'], ['ここにフォルダIDを入力']]);
      settingsSheet.autoResizeColumns(1, 2);
      
      showAlert('設定シートを作成しました', '「設定」シートにNotionデータベースIDとGoogle DriveフォルダIDを入力してください。');
      ss.setActiveSheet(settingsSheet);
      return false;
    }
    
    // 設定を読み込む
    const settings = settingsSheet.getRange('A1:B10').getValues();
    
    // 設定値を探す
    for (let i = 0; i < settings.length; i++) {
      const key = settings[i][0];
      const value = settings[i][1];
      
      if (key === 'NotionデータベースID' && value && value !== 'ここにデータベースIDを入力') {
        DATABASE_ID = value;
      } else if (key === 'Google DriveフォルダID' && value && value !== 'ここにフォルダIDを入力') {
        DRIVE_FOLDER_ID = value;
      }
    }
    
    // 必要な設定が揃っているか確認
    if (!DATABASE_ID || !DRIVE_FOLDER_ID) {
      showAlert('設定が不足しています', '「設定」シートにNotionデータベースIDとGoogle DriveフォルダIDを入力してください。');
      ss.setActiveSheet(settingsSheet);
      return false;
    }
    
    return true;
  }
  
  /**
   * 設定シートを初期化する
   */
  function initSettingsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      showAlert('エラー', 'スプレッドシートが見つかりません。');
      return;
    }
    
    // 既存の設定シートを削除
    const existingSheet = ss.getSheetByName('設定');
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // 設定シートを新規作成
    const settingsSheet = ss.insertSheet('設定');
    
    // 設定シートの初期設定
    settingsSheet.getRange('A1:B1').setValues([['設定項目', '値']]).setFontWeight('bold');
    settingsSheet.getRange('A2:A3').setValues([['NotionデータベースID'], ['Google DriveフォルダID']]);
    settingsSheet.getRange('B2:B3').setValues([['ここにデータベースIDを入力'], ['ここにフォルダIDを入力']]);
    
    // 書式設定
    settingsSheet.getRange('B2:B3').setBackground('#f3f3f3');
    settingsSheet.getRange('A6:A7').setFontStyle('italic');
    settingsSheet.getRange('B6:B7').setFontStyle('italic').setFontColor('#666666');
    
    // 列幅の自動調整
    settingsSheet.autoResizeColumns(1, 2);
    
    // シートをアクティブにする
    ss.setActiveSheet(settingsSheet);
    
    showAlert('設定シートを初期化しました', '「設定」シートにNotionデータベースIDとGoogle DriveフォルダIDを入力してください。');
  }
  
  /**
   * スクリプトを実行する前に必要なサービスを確認
   */
  function checkRequiredServices() {
    // Drive APIを有効化するためのメッセージを表示
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Google Drive APIの有効化',
        'このスクリプトを使用するには、Google Drive APIを有効化する必要があります。\n\n' +
        '1. Apps Scriptエディタで「サービス」アイコンをクリック\n' + 
        '2. 「サービスを追加」から「Drive API」を選択\n' +
        '3. 「追加」ボタンをクリック\n\n' +
        '設定後、スクリプトを保存してください。',
        ui.ButtonSet.OK
      );
    } catch (e) {
      Logger.log('UIが利用できないため、メッセージを表示できません。Google Drive APIが有効化されていることを確認してください。');
    }
  }