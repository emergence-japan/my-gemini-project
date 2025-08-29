// 領収書発行システム Apps Script コード
// エマージェンス・ジャパン合同会社 PDF領収書自動発行

// 設定読み込み関数
function loadConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');
  
  if (!configSheet) {
    throw new Error('configシートが見つかりません');
  }
  
  const data = configSheet.getDataRange().getValues();
  const config = {};
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      config[data[i][0]] = data[i][1];
    }
  }
  
  return config;
}

// HMAC-SHA256署名生成
function generateSignature(payload, secretKey) {
  const signature = Utilities.computeHmacSha256Signature(payload, secretKey);
  return Utilities.base64EncodeWebSafe(signature);
}

// トークン生成
function generateToken(orderData, secretKey) {
  const payload = {
    order_id: orderData.order_id,
    email: orderData.email,
    exp: Date.now() + (7 * 24 * 60 * 60 * 1000) // 7日後
  };
  
  const payloadJson = JSON.stringify(payload);
  const payloadB64 = Utilities.base64EncodeWebSafe(Utilities.newBlob(payloadJson).getBytes());
  const signature = generateSignature(payloadB64, secretKey);
  
  return payloadB64 + '.' + signature;
}

// トークン検証
function verifyToken(token, secretKey) {
  try {
    const parts = token.split('.');
    if (parts.length !== 2) return null;
    
    const [payloadB64, signature] = parts;
    const expectedSignature = generateSignature(payloadB64, secretKey);
    
    if (signature !== expectedSignature) return null;
    
    const payloadJson = Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString();
    const payload = JSON.parse(payloadJson);
    
    if (Date.now() > payload.exp) return null;
    
    return payload;
  } catch (e) {
    return null;
  }
}

// 管理者用URL生成メニュー
function generateReceiptURL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = ss.getSheetByName('ledger');
  const activeRange = ledger.getActiveRange();
  
  if (!activeRange || activeRange.getRow() < 2) {
    SpreadsheetApp.getUi().alert('台帳の行を選択してください');
    return;
  }
  
  const row = activeRange.getRow();
  const rowData = ledger.getRange(row, 1, 1, 11).getValues()[0];
  
  if (rowData[4]) { // issued列がtrueの場合
    SpreadsheetApp.getUi().alert('この注文は既に発行済みです');
    return;
  }
  
  const config = loadConfig();
  const orderData = {
    order_id: rowData[0],
    email: rowData[1]
  };
  
  const token = generateToken(orderData, config.secret_key);
  const expiry = new Date(Date.now() + (7 * 24 * 60 * 60 * 1000));
  
  // 台帳更新
  ledger.getRange(row, 3).setValue(token);
  ledger.getRange(row, 4).setValue(expiry);
  
  const url = ScriptApp.getService().getUrl() + '?token=' + token;
  
  // メール送信
  const subject = '領収書発行フォームのご案内 - エマージェンス・ジャパン合同会社';
  const body = `
お疲れさまです。

生成AIセミナーのお支払いありがとうございました。
下記のリンクより領収書をダウンロードできます。

【領収書発行フォーム】
${url}

※このリンクは7日間有効です。
※宛名をご入力いただくと、即座にPDF領収書がダウンロードできます。

何かご不明な点がございましたら、お気軽にお問い合わせください。

--
エマージェンス・ジャパン合同会社
Email: ${config.from_email}
TEL: ${config.tel}
  `;
  
  GmailApp.sendEmail(
    rowData[1],
    subject,
    body,
    {
      name: config.from_name,
      replyTo: config.from_email
    }
  );
  
  SpreadsheetApp.getUi().alert('領収書発行URLを送信しました:\n' + url);
}

// Web側 GET処理（フォーム表示）
function doGet(e) {
  const token = e && e.parameter ? e.parameter.token : null;
  
  if (!token) {
    return HtmlService.createHtmlOutput('<h1>エラー</h1><p>無効なリンクです。</p>');
  }
  
  const config = loadConfig();
  const payload = verifyToken(token, config.secret_key);
  
  if (!payload) {
    return HtmlService.createHtmlOutput('<h1>エラー</h1><p>リンクが無効または期限切れです。</p>');
  }
  
  // 台帳確認
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = ss.getSheetByName('ledger');
  const data = ledger.getDataRange().getValues();
  
  let orderRow = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === payload.order_id) {
      orderRow = data[i];
      break;
    }
  }
  
  if (!orderRow) {
    return HtmlService.createHtmlOutput('<h1>エラー</h1><p>注文が見つかりません。</p>');
  }
  
  const htmlTemplate = HtmlService.createTemplateFromFile('form');
  htmlTemplate.token = token;
  htmlTemplate.amount = config.amount_taxin;
  htmlTemplate.description = config.description;
  htmlTemplate.company = config.company;
  
  return htmlTemplate.evaluate()
    .setTitle('領収書発行フォーム - エマージェンス・ジャパン合同会社')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Web側 POST処理（PDF生成・送信）
function doPost(e) {
  try {
    const token = e && e.parameter ? e.parameter.token : null;
    const billTo = e && e.parameter ? e.parameter.bill_to : null;
    const customerEmail = e && e.parameter ? (e.parameter.customer_email || '') : '';
    
    if (!token || !billTo) {
      throw new Error('必要な情報が不足しています');
    }
    
    const config = loadConfig();
    const payload = verifyToken(token, config.secret_key);
    
    if (!payload) {
      throw new Error('トークンが無効または期限切れです');
    }
    
    // 台帳検索・更新
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ledger = ss.getSheetByName('ledger');
    const data = ledger.getDataRange().getValues();
    
    let orderRowIndex = -1;
    let orderRow = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.order_id) {
        orderRowIndex = i + 1;
        orderRow = data[i];
        break;
      }
    }
    
    if (!orderRow) {
      throw new Error('注文が見つかりません');
    }
    
    // 領収書番号生成
    const now = new Date();
    const year = now.getFullYear();
    let receiptNo = orderRow[6]; // 既存の領収書番号
    let isReissue = false;
    
    if (!receiptNo) {
      // 新規発行
      const lastNo = getLastReceiptNumber(ledger, year);
      receiptNo = `EJ-${year}-${String(lastNo + 1).padStart(5, '0')}`;
    } else {
      // 再発行
      isReissue = true;
    }
    
    // PDF生成
    const pdfBlob = generateReceiptPDF({
      receiptNo: receiptNo,
      billTo: billTo,
      payDate: orderRow[8] || now,
      isReissue: isReissue,
      config: config
    });
    
    // Googleドライブに保存
    const folder = DriveApp.getFolderById(config.pdf_folder_id);
    const fileName = `${receiptNo}_${billTo.replace(/[\/\\:*?"<>|]/g, '_')}.pdf`;
    const file = folder.createFile(pdfBlob.setName(fileName));
    
    // 台帳更新
    const updateRange = ledger.getRange(orderRowIndex, 1, 1, 11);
    const updateData = [
      orderRow[0], // order_id
      orderRow[1], // email
      orderRow[2], // token
      orderRow[3], // token_expiry
      true,        // issued
      now,         // issued_at
      receiptNo,   // receipt_no
      billTo,      // bill_to
      orderRow[8], // pay_date
      isReissue,   // reissue
      (orderRow[10] || 0) + 1 // downloads
    ];
    updateRange.setValues([updateData]);
    
    // 顧客にメール送信
    if (customerEmail || orderRow[1]) {
      const targetEmail = customerEmail || orderRow[1];
      const subject = `領収書発行完了 - ${receiptNo} - エマージェンス・ジャパン合同会社`;
      const body = `
お疲れさまです。

領収書を発行いたしました。

【領収書情報】
領収書番号: ${receiptNo}
宛名: ${billTo}
金額: ¥${Number(config.amount_taxin).toLocaleString()}（税込）
但し書き: ${config.description}

領収書PDFを添付いたします。

--
エマージェンス・ジャパン合同会社
Email: ${config.from_email}
TEL: ${config.tel}
      `;
      
      GmailApp.sendEmail(
        targetEmail,
        subject,
        body,
        {
          name: config.from_name,
          replyTo: config.from_email,
          attachments: [pdfBlob]
        }
      );
    }
    
    // 成功レスポンス
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      receiptNo: receiptNo,
      downloadUrl: file.getDownloadUrl(),
      message: isReissue ? '領収書を再発行いたしました。' : '領収書を発行いたしました。'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('PDF生成エラー: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 最後の領収書番号を取得
function getLastReceiptNumber(ledger, year) {
  const data = ledger.getDataRange().getValues();
  let maxNo = 0;
  
  for (let i = 1; i < data.length; i++) {
    const receiptNo = data[i][6];
    if (receiptNo && receiptNo.startsWith(`EJ-${year}-`)) {
      const no = parseInt(receiptNo.split('-')[2]);
      if (no > maxNo) {
        maxNo = no;
      }
    }
  }
  
  return maxNo;
}

// PDF生成
function generateReceiptPDF(params) {
  const config = params.config;
  const template = DriveApp.getFileById(config.doc_template_id);
  
  // テンプレート複製
  const copy = template.makeCopy(`temp_${params.receiptNo}_${Date.now()}`);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  
  // プレースホルダー置換
  const payDate = new Date(params.payDate);
  const issueDate = new Date();
  
  body.replaceText('{{RECEIPT_NO}}', params.receiptNo);
  body.replaceText('{{BILL_TO}}', params.billTo);
  body.replaceText('{{AMOUNT_TAXIN}}', '¥' + Number(config.amount_taxin).toLocaleString());
  body.replaceText('{{AMOUNT_TAXOUT}}', '¥' + Number(config.amount_taxout).toLocaleString());
  body.replaceText('{{TAX}}', '¥' + Number(config.tax).toLocaleString());
  body.replaceText('{{DESCRIPTION}}', config.description);
  body.replaceText('{{PAY_DATE}}', payDate.getFullYear() + '年' + (payDate.getMonth() + 1) + '月' + payDate.getDate() + '日');
  body.replaceText('{{ISSUE_DATE}}', issueDate.getFullYear() + '年' + (issueDate.getMonth() + 1) + '月' + issueDate.getDate() + '日');
  body.replaceText('{{COMPANY}}', config.company);
  body.replaceText('{{ADDRESS}}', config.address);
  body.replaceText('{{TEL}}', config.tel);
  body.replaceText('{{EMAIL}}', config.from_email);
  body.replaceText('{{INVOICE_NO}}', config.invoice_no);
  body.replaceText('{{PAY_METHOD}}', config.pay_method);
  
  if (params.isReissue) {
    body.replaceText('{{REISSUE_MARK}}', '【再発行】');
  } else {
    body.replaceText('{{REISSUE_MARK}}', '');
  }
  
  doc.saveAndClose();
  
  // PDF変換
  const pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
  
  // 一時ファイル削除
  DriveApp.getFileById(copy.getId()).setTrashed(true);
  
  return pdfBlob;
}

// メニュー作成（スプレッドシート用）
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('領収書管理')
    .addItem('選択行のURL生成', 'generateReceiptURL')
    .addToUi();
}