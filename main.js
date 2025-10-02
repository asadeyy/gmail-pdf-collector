function savePdfFromGmail() {
  // ===== 設定項目 =====
  const TARGET_YEAR = 0; // 対象年（0の場合は当年）
  const TARGET_MONTH = 0; // 対象月（0の場合は前月、1-12）
  const FOLDER_ID = ""; // 保存先フォルダのID
  
  // ===== メイン処理 =====
  try {
    // 年月の設定（0の場合の処理）
    const today = new Date();
    let targetYear, targetMonth;
    
    if (TARGET_YEAR === 0 && TARGET_MONTH === 0) {
      // 両方0の場合は前月を計算
      const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      targetYear = lastMonth.getFullYear();
      targetMonth = lastMonth.getMonth() + 1;
    } else {
      // 年が0の場合は今年、月が0の場合は前月
      targetYear = TARGET_YEAR === 0 ? today.getFullYear() : TARGET_YEAR;
      targetMonth = TARGET_MONTH === 0 ? today.getMonth() : TARGET_MONTH;
    }
      
    Logger.log(`処理対象: ${targetYear}年${targetMonth}月`);
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // mainシートからメールタイトルのリストを取得
    const mainSheet = ss.getSheetByName("main");
    if (!mainSheet) {
      throw new Error("mainシートが見つかりません");
    }
    
    // 2行目から最終行までのA列とB列を取得
    const lastRow = mainSheet.getLastRow();
    if (lastRow < 2) {
      throw new Error("mainシートにメールタイトルが設定されていません");
    }
    
    const dataRange = mainSheet.getRange(2, 1, lastRow - 1, 2);
    const dataValues = dataRange.getValues();
    
    // タイトルとPDF化フラグのペアを作成
    const emailSettings = dataValues
      .filter(row => row[0] !== "")
      .map(row => ({
        title: row[0],
        convertToPdf: row[1] === true || row[1] === "TRUE" || row[1] === "true" || row[1] === "✓"
      }));
    
    if (emailSettings.length === 0) {
      throw new Error("有効なメールタイトルが見つかりません");
    }
    
    Logger.log(`処理対象のタイトル数: ${emailSettings.length}`);
    
    // ログシートを作成または取得
    const logSheetName = `${targetYear}${String(targetMonth).padStart(2, '0')}`;
    let logSheet = ss.getSheetByName(logSheetName);
    
    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      // ヘッダー行を作成
      logSheet.getRange(1, 1, 1, 7).setValues([[
        "処理日時", "メールタイトル", "メール受信日", "添付ファイル名", "保存ファイル名", "PDF化", "ステータス"
      ]]);
      logSheet.getRange(1, 1, 1, 7).setFontWeight("bold");
      logSheet.setFrozenRows(1);
    }
    
    // 保存先フォルダを取得
    const folder = DriveApp.getFolderById(FOLDER_ID);
    
    // 検索期間の設定
    const startDate = new Date(targetYear, targetMonth - 1, 1);
    const endDate = new Date(targetYear, targetMonth, 0);
    
    let totalSavedCount = 0;
    const processTime = new Date();
    
    // 各タイトルごとに処理
    emailSettings.forEach(setting => {
      const emailTitle = setting.title;
      const convertToPdf = setting.convertToPdf;
      
      try {
        Logger.log(`処理中: ${emailTitle} (PDF化: ${convertToPdf})`);
        
        // Gmail検索クエリを構築
        // PDF化する場合は本文検索、それ以外は件名検索
        let query;
        if (convertToPdf) {
          query = `"${emailTitle}" after:${formatDate(startDate)} before:${formatDate(endDate)}`;
        } else {
          query = `subject:"${emailTitle}" after:${formatDate(startDate)} before:${formatDate(endDate)} has:attachment filename:pdf`;
        }
        
        // メールスレッドを検索
        const threads = GmailApp.search(query);
        
        if (threads.length === 0) {
          // メールが見つからなかった場合もログに記録
          logSheet.appendRow([
            processTime,
            emailTitle,
            "",
            "",
            "",
            convertToPdf ? "はい" : "いいえ",
            "該当メールなし"
          ]);
          Logger.log(`該当メールなし: ${emailTitle}`);
          return;
        }
        
        let savedCount = 0;
        
        // 各スレッドを処理
        threads.forEach(thread => {
          const messages = thread.getMessages();
          
          messages.forEach(message => {
            const messageDate = message.getDate();
            const dateStr = Utilities.formatDate(messageDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
            
            // メール本文をPDF化する場合
            if (convertToPdf) {
              try {
                const subject = message.getSubject();
                const body = message.getBody();
                const fileName = `${dateStr}_${subject}.pdf`;
                
                // HTMLをPDFに変換
                const blob = Utilities.newBlob(body, "text/html", "temp.html");
                const pdfBlob = blob.getAs("application/pdf");
                pdfBlob.setName(fileName);
                
                // Google Driveに保存
                const file = folder.createFile(pdfBlob);
                
                // ログに記録
                logSheet.appendRow([
                  processTime,
                  emailTitle,
                  messageDate,
                  "メール本文",
                  fileName,
                  "はい",
                  "保存成功"
                ]);
                
                savedCount++;
                totalSavedCount++;
                Logger.log(`メールPDF保存成功: ${fileName}`);
                
              } catch (e) {
                logSheet.appendRow([
                  processTime,
                  emailTitle,
                  messageDate,
                  "メール本文",
                  "",
                  "はい",
                  `エラー: ${e.message}`
                ]);
                Logger.log(`メールPDF化エラー: ${e.message}`);
              }
            } else {
              // 添付PDFを保存する場合
              const attachments = message.getAttachments();
              
              if (attachments.length === 0) {
                logSheet.appendRow([
                  processTime,
                  emailTitle,
                  messageDate,
                  "",
                  "",
                  "いいえ",
                  "PDF添付なし"
                ]);
                return;
              }
              
              attachments.forEach(attachment => {
                try {
                  // PDFファイルのみを処理
                  if (attachment.getContentType() === "application/pdf") {
                    const originalName = attachment.getName();
                    const fileName = `${dateStr}_${originalName}`;
                    
                    // Google Driveに保存
                    const file = folder.createFile(attachment);
                    file.setName(fileName);
                    
                    // ログに記録
                    logSheet.appendRow([
                      processTime,
                      emailTitle,
                      messageDate,
                      originalName,
                      fileName,
                      "いいえ",
                      "保存成功"
                    ]);
                    
                    savedCount++;
                    totalSavedCount++;
                    Logger.log(`保存成功: ${fileName}`);
                  }
                } catch (e) {
                  // 添付ファイル処理エラー
                  logSheet.appendRow([
                    processTime,
                    emailTitle,
                    messageDate,
                    attachment.getName(),
                    "",
                    "いいえ",
                    `エラー: ${e.message}`
                  ]);
                  Logger.log(`添付ファイル保存エラー: ${e.message}`);
                }
              });
            }
          });
        });
        
        Logger.log(`${emailTitle}: ${savedCount}個のPDFを保存`);
        
      } catch (e) {
        // タイトルごとの処理エラー
        logSheet.appendRow([
          processTime,
          emailTitle,
          "",
          "",
          "",
          convertToPdf ? "はい" : "いいえ",
          `エラー: ${e.message}`
        ]);
        Logger.log(`エラー (${emailTitle}): ${e.message}`);
      }
    });
    
    // 列幅を自動調整
    logSheet.autoResizeColumns(1, 7);
    
    Logger.log(`=== 処理完了 ===`);
    Logger.log(`合計 ${totalSavedCount}個のPDFを保存しました`);
    Logger.log(`ログは「${logSheetName}」シートを確認してください`);
    
  } catch (e) {
    Logger.log(`致命的なエラーが発生しました: ${e.message}`);
    Browser.msgBox("エラー", e.message, Browser.Buttons.OK);
  }
}

/**
 * 日付をGmail検索用フォーマットに変換
 * @param {Date} date - 変換する日付
 * @return {string} yyyy/mm/dd形式の文字列
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  return `${year}/${month}/${day}`;
}

/**
 * mainシートのセットアップ
 * 初回実行時にmainシートを作成し、サンプルデータを入力
 */
function setupMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mainSheet = ss.getSheetByName("main");
  
  if (!mainSheet) {
    mainSheet = ss.insertSheet("main");
  }
  
  // ヘッダーとサンプルデータを設定
  mainSheet.getRange("A1").setValue("メールタイトル（部分一致）");
  mainSheet.getRange("B1").setValue("メール本文をPDF化");
  mainSheet.getRange("A1:B1").setFontWeight("bold");
  
  mainSheet.getRange("A2").setValue("請求書");
  mainSheet.getRange("B2").setValue(false);
  mainSheet.getRange("A3").setValue("領収書");
  mainSheet.getRange("B3").setValue(false);
  mainSheet.getRange("A4").setValue("重要なお知らせ");
  mainSheet.getRange("B4").setValue(true);
  
  mainSheet.setColumnWidth(1, 300);
  mainSheet.setColumnWidth(2, 150);
  mainSheet.setFrozenRows(1);
  
  Browser.msgBox("セットアップ完了", "mainシートを作成しました。\nA列: メールタイトルを入力\nB列: メール本文をPDF化する場合はチェック", Browser.Buttons.OK);
}