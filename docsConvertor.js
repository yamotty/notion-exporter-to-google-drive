// @ts-nocheck
/**
 * NotionのブロックをGoogle Docs形式に変換して保存 (リファクタリング版)
 * @param {Array} blocks - Notionブロックの配列
 * @param {string} pageTitle - ページタイトル
 * @param {string} existingDocId - 既存のGoogle DocsのID（更新する場合）
 * @return {Object} {success: boolean, message: string, fileId: string} 形式のステータス
 */
function convertBlocksToGoogleDocs(blocks, pageTitle, existingDocId = null) {
    try {
      // 安全なファイル名を作成
      const safeFileName = pageTitle.replace(/[\\/:*?"<>|]/g, '_');
      
      let doc;
      let docFile;
      let isUpdate = false;
      
      // 既存のドキュメントがある場合は更新、なければ新規作成
      if (existingDocId) {
        try {
          doc = DocumentApp.openById(existingDocId);
          docFile = DriveApp.getFileById(existingDocId);
          isUpdate = true;
          
          // ドキュメントの内容をクリア（最初の段落を除く）
          const body = doc.getBody();
          body.clear();
          
          Logger.log(`既存のドキュメント "${safeFileName}" (ID: ${existingDocId}) を更新します`);
        } catch (e) {
          Logger.log(`既存のドキュメントを開けませんでした: ${e.message}。新規作成します。`);
          existingDocId = null; // 既存ドキュメントが見つからない場合は新規作成モードに
        }
      }
      
      // 新規作成モード
      if (!existingDocId) {
        // 同名のファイルが存在するか確認
        const existingFiles = DriveApp.getFolderById(DRIVE_FOLDER_ID)
          .getFilesByName(safeFileName);
        
        // 既存のファイルがある場合は削除
        if (existingFiles.hasNext()) {
          const existingFile = existingFiles.next();
          try {
            existingFile.setTrashed(true);
            Logger.log(`既存のファイル "${safeFileName}" を削除しました`);
          } catch (e) {
            Logger.log(`既存ファイルの削除に失敗: ${e.message}`);
          }
        }
        
        // 新しいGoogle Docsを作成
        doc = DocumentApp.create(safeFileName);
        docFile = DriveApp.getFileById(doc.getId());
      }
      
      // 共通画像フォルダを取得または作成
      let imageFolder;
      try {
        // 固定の名前で画像フォルダを取得または作成
        imageFolder = getOrCreateCommonImageFolder(DRIVE_FOLDER_ID);
      } catch (e) {
        Logger.log(`共通画像フォルダの取得に失敗: ${e.message}`);
        // フォルダ作成に失敗しても処理は続行
      }
      
      // ドキュメントのbodyを取得
      const body = doc.getBody();
      
      // タイトルを設定
      body.appendParagraph(pageTitle)
          .setHeading(DocumentApp.ParagraphHeading.HEADING1)
          .setAttributes({
            [DocumentApp.Attribute.FONT_SIZE]: 18,
            [DocumentApp.Attribute.BOLD]: true
          });
      
      // ページIDを取得（最初のブロックから取得可能）
      let pageId = "";
      if (blocks.length > 0 && blocks[0].parent) {
        pageId = blocks[0].parent.page_id || "";
      }
      // ページIDがない場合はランダムなIDを生成
      if (!pageId) {
        pageId = Utilities.getUuid();
      }
      
      // ブロックを処理してGoogle Docsに変換
      let currentListItems = [];
      let currentListType = null;
      
      blocks.forEach((block, index) => {
        try {
          const blockType = block.type;
          
          // リスト以外のブロックが来たらリストをフラッシュ
          if (blockType !== 'bulleted_list_item' && blockType !== 'numbered_list_item' && currentListItems.length > 0) {
            appendListItems(body, currentListItems, currentListType);
            currentListItems = [];
            currentListType = null;
          }
          
          // ブロックタイプに応じて処理
          switch (blockType) {
            case 'paragraph':
              if (block.paragraph && block.paragraph.rich_text) {
                appendRichTextToDoc(body.appendParagraph(''), block.paragraph.rich_text);
              }
              break;
              
            case 'heading_1':
              if (block.heading_1 && block.heading_1.rich_text) {
                const heading = body.appendParagraph('');
                heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
                appendRichTextToDoc(heading, block.heading_1.rich_text);
              }
              break;
              
            case 'heading_2':
              if (block.heading_2 && block.heading_2.rich_text) {
                const heading = body.appendParagraph('');
                heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
                appendRichTextToDoc(heading, block.heading_2.rich_text);
              }
              break;
              
            case 'heading_3':
              if (block.heading_3 && block.heading_3.rich_text) {
                const heading = body.appendParagraph('');
                heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
                appendRichTextToDoc(heading, block.heading_3.rich_text);
              }
              break;
              
            case 'bulleted_list_item':
              if (block.bulleted_list_item && block.bulleted_list_item.rich_text) {
                // リストアイテムをバッファに追加
                currentListType = 'BULLET';
                currentListItems.push({
                  text: block.bulleted_list_item.rich_text,
                  type: 'BULLET'
                });
              }
              break;
              
            case 'numbered_list_item':
              if (block.numbered_list_item && block.numbered_list_item.rich_text) {
                // リストアイテムをバッファに追加
                currentListType = 'NUMBER';
                currentListItems.push({
                  text: block.numbered_list_item.rich_text,
                  type: 'NUMBER'
                });
              }
              break;
              
            case 'to_do':
              if (block.to_do && block.to_do.rich_text) {
                const todoText = block.to_do.checked ? '☑ ' : '☐ ';
                const paragraph = body.appendParagraph(todoText);
                appendRichTextToDoc(paragraph, block.to_do.rich_text);
              }
              break;
              
            case 'toggle':
              if (block.toggle && block.toggle.rich_text) {
                const toggleText = '▶ ';
                const paragraph = body.appendParagraph(toggleText);
                appendRichTextToDoc(paragraph, block.toggle.rich_text);
                // 子要素があればインデントして追加する（実際のToggleはサポートできないので視覚的な表現）
                body.appendParagraph('    [Toggle content - collapsed]').setItalic(true);
              }
              break;
              
            case 'callout':
              // Calloutブロックを処理
              processCalloutBlock(block, body);
              break;
              
            case 'code':
              if (block.code && block.code.rich_text) {
                const codeBlock = body.appendParagraph('');
                appendRichTextToDoc(codeBlock, block.code.rich_text);
                codeBlock.setAttributes({
                  [DocumentApp.Attribute.FONT_FAMILY]: 'Courier New',
                  [DocumentApp.Attribute.BACKGROUND_COLOR]: '#f6f8fa'
                });
                if (block.code.language) {
                  body.appendParagraph(`Language: ${block.code.language}`).setItalic(true);
                }
              }
              break;
              
            case 'quote':
              if (block.quote && block.quote.rich_text) {
                const quoteBlock = body.appendParagraph('');
                appendRichTextToDoc(quoteBlock, block.quote.rich_text);
                quoteBlock.setAttributes({
                  [DocumentApp.Attribute.INDENT_START]: 30,
                  [DocumentApp.Attribute.INDENT_FIRST_LINE]: 30,
                  [DocumentApp.Attribute.ITALIC]: true,
                  [DocumentApp.Attribute.FOREGROUND_COLOR]: '#6a737d'
                });
              }
              break;
              
            case 'divider':
              body.appendHorizontalRule();
              break;
              
            case 'image':
              // 画像処理を改善した関数を呼び出す
              processImageBlock(block, body, imageFolder, pageId);
              break;
              
            case 'table':
              // テーブルは完全なサポートが難しいため、プレースホルダを表示
              body.appendParagraph('[Table content - テーブルは完全にサポートされていません]').setItalic(true);
              break;
              
            default:
              Logger.log(`未サポートのブロックタイプ: ${blockType}`);
              body.appendParagraph(`[${blockType} - このブロックタイプはサポートされていません]`).setItalic(true);
          }
        } catch (error) {
          Logger.log(`ブロック ${index} (${block.type || 'unknown'}) の処理中にエラー: ${error.message}`);
          body.appendParagraph(`[Error processing ${block.type || 'unknown'} block: ${error.message}]`).setItalic(true);
        }
      });
      
      // 残っているリストアイテムがあればフラッシュ
      if (currentListItems.length > 0) {
        appendListItems(body, currentListItems, currentListType);
      }
      
      // ドキュメントを保存
      doc.saveAndClose();
      
      // 新規作成の場合は、作成したドキュメントを指定フォルダに移動
      if (!isUpdate) {
        const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        folder.addFile(docFile);
        DriveApp.getRootFolder().removeFile(docFile);
      }
      
      return {
        success: true,
        message: isUpdate 
          ? `ファイル "${safeFileName}" を更新しました`
          : `ファイル "${safeFileName}" をGoogle Docs形式で作成しました`,
        fileId: doc.getId()
      };
    } catch (error) {
      Logger.log(`Google Docs変換中にエラー: ${error.message}`);
      Logger.log(`エラースタック: ${error.stack}`);
      return {
        success: false,
        message: `変換中にエラーが発生しました: ${error.message}`,
        error: error
      };
    }
  }
  
  /**
   * NotionのCalloutブロックを処理する関数
   * @param {Object} block - Notionブロック
   * @param {Body} body - ドキュメントのbody
   */
  function processCalloutBlock(block, body) {
    if (!block.callout || !block.callout.rich_text) {
      return;
    }
    
    try {
      // Calloutのアイコンを取得（テキストとして使用）
      let iconText = "📝 "; // デフォルトアイコン
      
      if (block.callout.icon) {
        if (block.callout.icon.type === "emoji" && block.callout.icon.emoji) {
          iconText = block.callout.icon.emoji + " ";
        } else if (block.callout.icon.type === "external" && block.callout.icon.external && block.callout.icon.external.url) {
          // 外部アイコンの場合は記号に置き換え
          iconText = "🔗 ";
        } else if (block.callout.icon.type === "file" && block.callout.icon.file && block.callout.icon.file.url) {
          // ファイルアイコンの場合も記号に置き換え
          iconText = "📎 ";
        }
      }
      
      // Calloutの背景色に応じてスタイルを変更
      let backgroundColor = '#f1f1f1'; // デフォルトの背景色
      let textColor = '#000000'; // デフォルトのテキスト色
      
      // Notionの背景色に基づいて色を設定
      if (block.callout.color) {
        switch (block.callout.color) {
          case 'blue_background':
            backgroundColor = '#d3e5ef';
            break;
          case 'brown_background':
            backgroundColor = '#e9e5e3';
            break;
          case 'gray_background':
            backgroundColor = '#e9e8e8';
            break;
          case 'green_background':
            backgroundColor = '#ddedea';
            break;
          case 'orange_background':
            backgroundColor = '#f9e2d2';
            break;
          case 'pink_background':
            backgroundColor = '#f4dfeb';
            break;
          case 'purple_background':
            backgroundColor = '#e8def8';
            break;
          case 'red_background':
            backgroundColor = '#fbe4e4';
            break;
          case 'yellow_background':
            backgroundColor = '#fbf3db';
            break;
          // テキスト色が背景色のケース
          case 'blue':
            textColor = '#0b6bcb';
            break;
          case 'brown':
            textColor = '#64473a';
            break;
          case 'gray':
            textColor = '#787774';
            break;
          case 'green':
            textColor = '#0f7b6c';
            break;
          case 'orange':
            textColor = '#d9730d';
            break;
          case 'pink':
            textColor = '#ad1a72';
            break;
          case 'purple':
            textColor = '#9b51e0';
            break;
          case 'red':
            textColor = '#e03e3e';
            break;
          case 'yellow':
            textColor = '#dfab01';
            break;
          case 'default':
          default:
            // デフォルト値をそのまま使用
            break;
        }
      }
      
      // アイコンとテキストを連結したパラグラフを作成
      const paragraph = body.appendParagraph(iconText);
      
      // テキスト内容を追加
      appendRichTextToDoc(paragraph, block.callout.rich_text);
      
      // スタイル設定
      paragraph.setAttributes({
        [DocumentApp.Attribute.BACKGROUND_COLOR]: backgroundColor,
        [DocumentApp.Attribute.FOREGROUND_COLOR]: textColor,
        [DocumentApp.Attribute.INDENT_START]: 20,
        [DocumentApp.Attribute.INDENT_END]: 20,
        [DocumentApp.Attribute.SPACING_BEFORE]: 10,
        [DocumentApp.Attribute.SPACING_AFTER]: 10
      });
      
      // 区切り線を引くことでCalloutを際立たせる
      body.appendParagraph('').setAttributes({
        [DocumentApp.Attribute.SPACING_AFTER]: 10
      });
      
      Logger.log('Calloutブロックを処理しました');
    } catch (error) {
      Logger.log(`Calloutブロックの処理中にエラー: ${error.message}`);
      body.appendParagraph('[Calloutの処理に失敗しました]').setItalic(true);
    }
  }

  /**
   * リッチテキストをGoogle Docsに適用
   * @param {Paragraph} paragraph - 段落要素
   * @param {Array} richTextArray - リッチテキスト配列
   */
  function appendRichTextToDoc(paragraph, richTextArray) {
    // リッチテキスト配列が空かnullの場合は何もしない
    if (!richTextArray || richTextArray.length === 0) {
      return;
    }
    
    let textOffset = 0;
    
    richTextArray.forEach(textObj => {
      const content = textObj.plain_text || '';
      paragraph.appendText(content);
      
      // スタイルがある場合は適用
      if (textObj.annotations) {
        const textRange = paragraph.editAsText();
        const endOffset = textOffset + content.length;
        
        // スタイル情報を取得
        const { bold, italic, strikethrough, underline, code, color } = textObj.annotations;
        
        // Google Docsの属性に変換して適用
        const attributes = {};
        
        if (bold) attributes[DocumentApp.Attribute.BOLD] = true;
        if (italic) attributes[DocumentApp.Attribute.ITALIC] = true;
        if (strikethrough) attributes[DocumentApp.Attribute.STRIKETHROUGH] = true;
        if (underline) attributes[DocumentApp.Attribute.UNDERLINE] = true;
        
        // コードブロックの場合はコードスタイルを適用
        if (code) {
          attributes[DocumentApp.Attribute.FONT_FAMILY] = 'Courier New';
          attributes[DocumentApp.Attribute.BACKGROUND_COLOR] = '#f6f8fa';
        }
        
        // 色を設定
        if (color && color !== 'default') {
          attributes[DocumentApp.Attribute.FOREGROUND_COLOR] = color;
        }
        
        // スタイルを適用
        if (Object.keys(attributes).length > 0 && content.length > 0) {
          textRange.setAttributes(textOffset, endOffset - 1, attributes);
        }
        
        // リンクがある場合は設定
        if (textObj.href && content.length > 0) {
          textRange.setLinkUrl(textOffset, endOffset - 1, textObj.href);
        }
      }
      
      textOffset += content.length;
    });
  }
  
  /**
   * リストアイテムをGoogle Docsに追加
   * @param {Body} body - ドキュメントのボディ
   * @param {Array} items - リストアイテムの配列
   * @param {string} type - リストタイプ ('BULLET' or 'NUMBER')
   */
  function appendListItems(body, items, type) {
    // Google Apps Scriptでは ListType は GlyphType として定義されています
    const listType = type === 'NUMBER' 
      ? DocumentApp.GlyphType.NUMBER 
      : DocumentApp.GlyphType.BULLET;
    
    // リスト処理
    items.forEach(item => {
      const listItem = body.appendListItem('');
      appendRichTextToDoc(listItem, item.text);
      listItem.setGlyphType(listType);
    });
  }