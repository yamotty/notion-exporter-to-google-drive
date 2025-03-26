/**
 * コンテキストに応じてアラートを表示またはログに記録する
 * @param {string} title - アラートのタイトル
 * @param {string} message - アラートのメッセージ
 */
function showAlert(title, message) {
    // まずログに記録（これは常に動作する）
    Logger.log(`${title}: ${message}`);
    
    try {
      // UIが利用可能か試みる
      const ui = SpreadsheetApp.getUi();
      ui.alert(title, message, ui.ButtonSet.OK);
    } catch (e) {
      // UIが利用できない場合は単にログに記録するだけ
      Logger.log(`UIアラートを表示できません: ${e.message}`);
    }
  }
  
  /**
   * 最新のエクスポート結果を表示する
   */
  function showExportResults() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultSheet = ss.getSheetByName('エクスポート結果');
    
    if (!resultSheet) {
      showAlert('結果がありません', 'エクスポートが実行されていないか、結果シートが見つかりません。');
      return;
    }
    
    // 結果シートをアクティブにする
    ss.setActiveSheet(resultSheet);
    
    // サマリーを計算
    const lastRow = resultSheet.getLastRow();
    if (lastRow <= 1) {
      showAlert('データがありません', 'エクスポート結果のデータがありません。');
      return;
    }
    
    const statusColumn = 4; // 「ステータス」列のインデックス
    const statusRange = resultSheet.getRange(2, statusColumn, lastRow - 1, 1);
    const statusValues = statusRange.getValues();
    
    let successCount = 0;
    let failCount = 0;
    
    statusValues.forEach(row => {
      if (row[0] === 'Success') successCount++;
      if (row[0] === 'Fail') failCount++;
    });
    
    // 最新の実行日時を取得（最初の列）
    const dateColumn = 1;
    const dateRange = resultSheet.getRange(2, dateColumn, lastRow - 1, 1);
    const dateValues = dateRange.getValues();
    const latestDate = new Date(Math.max(...dateValues.map(d => new Date(d[0]).getTime())));
    
    // フィルタを適用して最新の実行結果のみを表示
    // フィルタがあれば削除
    if (resultSheet.getFilter()) {
      resultSheet.getFilter().remove();
    }
    
    // 現在の日付で時間まで一致するレコードをカウント
    let latestRunCount = 0;
    const latestDateStr = latestDate.toISOString().split('T')[0]; // YYYY-MM-DD
    dateValues.forEach(dateCell => {
      const cellDate = new Date(dateCell[0]);
      const cellDateStr = cellDate.toISOString().split('T')[0];
      if (cellDateStr === latestDateStr) {
        latestRunCount++;
      }
    });
    
    // ヘッダーを含む全データ範囲にフィルタを設定
    const dataRange = resultSheet.getRange(1, 1, lastRow, resultSheet.getLastColumn());
    const filter = dataRange.createFilter();
    
    // 日付列にフィルタ条件を設定（当日のデータのみ表示）
    const filterCriteria = SpreadsheetApp.newFilterCriteria()
      .whenDateEqualTo(latestDate)
      .build();
    
    filter.setColumnFilterCriteria(dateColumn, filterCriteria);
    
    showAlert(
      'エクスポート結果サマリー',
      `最新実行: ${latestDate.toLocaleString()}\n` +
      `総処理数: ${latestRunCount}\n` +
      `成功: ${successCount}\n` +
      `失敗: ${failCount}\n\n` +
      `詳細はエクスポート結果シートをご確認ください。\n` +
      `（フィルタを適用して最新の実行結果のみ表示しています）`
    );
  }
  
  /**
   * 使い方ガイドを表示する
   */
  function showHelpGuide() {
    showAlert(
      'Notion DB Exporter の使い方',
      '【準備】\n' +
      '1. 「API Keyを設定」メニューからNotion APIキーを設定\n' +
      '2. 「必要なサービスを確認」から必要な設定を確認\n' +
      '3. 「設定シート」にNotionデータベースIDとGoogle DriveフォルダIDを入力\n\n' +
      '【実行】\n' +
      '1. 「NotionからGoogle Docsとしてエクスポート」を実行\n' +
      '2. 処理が完了すると結果が表示されます\n' +
      '3. 「エクスポート結果」シートで詳細な結果を確認できます\n\n' +
      '【トラブルシューティング】\n' +
      '• APIキーが正しく設定されているか確認\n' +
      '• データベースIDが正しいか確認\n' +
      '• フォルダへのアクセス権限があるか確認\n\n' +
      '詳細は「エクスポート結果」シートのエラーメッセージを参照してください。'
    );
  }
  
  /**
   * 処理結果をスプレッドシートに記録する
   * @param {Array} results - 処理結果の配列
   */
  function recordResultsToSpreadsheet(results) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Logger.log('スプレッドシートが見つかりません。');
        return;
      }
      
      // 「エクスポート結果」という名前のシートを探す、なければ作成
      let sheet = ss.getSheetByName('エクスポート結果');
      if (!sheet) {
        sheet = ss.insertSheet('エクスポート結果');
      }
      
      // ヘッダー行を設定
      const headers = ['実行日時', 'ページID', 'ページタイトル', 'ステータス', 'メッセージ', 'ファイルID', 'ファイルリンク'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      
      // 今回の実行バッチのグループIDを生成（タイムスタンプベース）
      const batchId = new Date().toISOString();
      
      // データ行を準備
      const rows = results.map(result => {
        const timestamp = new Date();
        let fileLink = '';
        
        // ファイルIDがある場合はリンクを作成
        if (result.fileId) {
          fileLink = `https://docs.google.com/document/d/${result.fileId}/edit`;
        }
        
        return [
          timestamp,
          result.pageId,
          result.pageTitle,
          result.status,
          result.message,
          result.fileId || '',
          fileLink
        ];
      });
      
      // 最大行数を超えないようにする（古いログを削除）
      const MAX_ROWS = 1000; // 履歴として保持する最大行数
      const currentRows = sheet.getLastRow();
      
      if (currentRows + rows.length > MAX_ROWS && currentRows > 1) {
        // 保持する行数を計算
        const rowsToKeep = Math.min(MAX_ROWS - rows.length, currentRows - 1);
        
        // 新しいデータ用に空きを作る
        if (currentRows - rowsToKeep > 0) {
          sheet.deleteRows(2, currentRows - rowsToKeep);
        }
      }
      
      // 既存のデータの下に新しいデータを追加
      if (rows.length > 0) {
        const lastRow = Math.max(sheet.getLastRow(), 1);
        sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
      }
      
      // 列幅の自動調整
      sheet.autoResizeColumns(1, headers.length);
      
      // ステータス列に条件付き書式を設定
      const statusColumn = 4; // 「ステータス」列のインデックス
      const range = sheet.getRange(2, statusColumn, sheet.getLastRow() - 1, 1);
      
      // 既存の条件付き書式をクリア
      const rules = sheet.getConditionalFormatRules();
      sheet.clearConditionalFormatRules();
      
      // 成功の場合は緑、失敗の場合は赤で表示
      const successRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Success')
        .setBackground('#b7e1cd')
        .setRanges([range])
        .build();
      
      const failRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Fail')
        .setBackground('#f4c7c3')
        .setRanges([range])
        .build();
      
      sheet.setConditionalFormatRules([successRule, failRule]);
      
      // リンク列をハイパーリンクとして設定
      const linkColumn = 7; // 「ファイルリンク」列のインデックス
      if (rows.length > 0) {
        const lastRow = sheet.getLastRow();
        const firstRow = lastRow - rows.length + 1;
        
        // 各行に対してリンクを設定
        for (let i = 0; i < rows.length; i++) {
          const row = firstRow + i;
          const link = rows[i][6]; // リンク列の値
          
          if (link) {
            // リンクテキストを設定
            sheet.getRange(row, linkColumn).setValue('ファイルを開く');
            
            // セルにハイパーリンクを設定
            sheet.getRange(row, linkColumn).setFormula(
              `=HYPERLINK("${link}", "ファイルを開く")`
            );
          }
        }
      }
      
      Logger.log('処理結果をスプレッドシートに記録しました。');
      
    } catch (error) {
      Logger.log('スプレッドシートへの結果記録に失敗しました: ' + error.message);
    }
  }
  
  /**
   * デバッグ結果をスプレッドシートに記録する
   * @param {Object} result - 処理結果
   */
  function recordDebugResultToSpreadsheet(result) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Logger.log('スプレッドシートが見つかりません。');
        return;
      }
      
      // 「デバッグ結果」という名前のシートを探す、なければ作成
      let sheet = ss.getSheetByName('デバッグ結果');
      if (!sheet) {
        sheet = ss.insertSheet('デバッグ結果');
      }
      
      // ヘッダー行を設定
      const headers = ['実行日時', 'ページID', 'ページタイトル', 'ステータス', 'メッセージ', 'ファイルID', 'ファイルリンク'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      
      // データ行を準備
      const timestamp = new Date();
      let fileLink = '';
      
      // ファイルIDがある場合はリンクを作成
      if (result.fileId) {
        fileLink = `https://docs.google.com/document/d/${result.fileId}/edit`;
      }
      
      const row = [
        timestamp,
        result.pageId,
        result.pageTitle,
        result.status,
        result.message,
        result.fileId || '',
        fileLink
      ];
      
      // 既存のデータを一旦クリア（ヘッダーは残す）
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
      
      // 新しいデータを追加
      sheet.getRange(2, 1, 1, headers.length).setValues([row]);
      
      // 列幅の自動調整
      sheet.autoResizeColumns(1, headers.length);
      
      // ステータス列に条件付き書式を設定
      const statusColumn = 4; // 「ステータス」列のインデックス
      const range = sheet.getRange(2, statusColumn, 1, 1);
      
      // 既存の条件付き書式をクリア
      sheet.clearConditionalFormatRules();
      
      // 成功の場合は緑、失敗の場合は赤で表示
      const successRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Success')
        .setBackground('#b7e1cd')
        .setRanges([range])
        .build();
      
      const failRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Fail')
        .setBackground('#f4c7c3')
        .setRanges([range])
        .build();
      
      sheet.setConditionalFormatRules([successRule, failRule]);
      
      // リンク列をハイパーリンクとして設定
      const linkColumn = 7; // 「ファイルリンク」列のインデックス
      if (fileLink) {
        // リンクテキストを設定
        sheet.getRange(2, linkColumn).setValue('ファイルを開く');
        
        // セルにハイパーリンクを設定
        sheet.getRange(2, linkColumn).setFormula(
          `=HYPERLINK("${fileLink}", "ファイルを開く")`
        );
      }
      
      // シートをアクティブにする
      ss.setActiveSheet(sheet);
      
      Logger.log('デバッグ結果をスプレッドシートに記録しました。');
      
    } catch (error) {
      Logger.log('スプレッドシートへの結果記録に失敗しました: ' + error.message);
    }
  }