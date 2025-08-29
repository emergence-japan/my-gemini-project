// 領収書発行システム Apps Script コード（最終版）
// エマージェンス・ジャパン合同会社 PDF領収書自動発行

// 設定読み込み関数
function loadConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('config');
    
    if (!configSheet) {
      throw new Error('configシートが見つかりません');
    }
    
    const data = configSheet.getDataRange().getValues();
    const config = {};
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][1] !== undefined) {
        config[data[i][0]] = data[i][1];
      }
    }
    
    // 必須設定の確認
    const required = ['secret_key', 'pdf_folder_id', 'doc_template_id', 'company', 'amount_taxin'];
    for (const key of required) {
      if (!config[key]) {
        throw new Error(`必須設定が不足しています: ${key}`);
      }
    }
    
    return config;
  } catch (error) {
    Logger.log('Config読み込みエラー: ' + error.toString());
    throw error;
  }
}

// HMAC-SHA256署名生成
function generateSignature(payload, secretKey) {
  try {
    const signature = Utilities.computeHmacSha256Signature(payload, secretKey);
    return Utilities.base64EncodeWebSafe(signature);
  } catch (error) {
    Logger.log('署名生成エラー: ' + error.toString());
    throw error;
  }
}

// トークン生成（修正版）
function generateToken(orderData, secretKey) {
  try {
    const payload = {
      order_id: orderData.order_id,
      email: orderData.email,
      exp: Date.now() + (7 * 24 * 60 * 60 * 1000) // 7日後
    };
    
    const payloadJson = JSON.stringify(payload);
    const payloadB64 = Utilities.base64EncodeWebSafe(payloadJson);
    const signature = generateSignature(payloadB64, secretKey);
    
    return payloadB64 + '.' + signature;
  } catch (error) {
    Logger.log('トークン生成エラー: ' + error.toString());
    throw error;
  }
}

// トークン検証（修正版）
function verifyToken(token, secretKey) {
  try {
    if (!token || typeof token !== 'string') {
      Logger.log('無効なトークン形式');
      return null;
    }
    
    const parts = token.split('.');
    if (parts.length !== 2) {
      Logger.log('トークン形式エラー: 部分数が不正');
      return null;
    }
    
    const [payloadB64, signature] = parts;
    
    // 署名検証
    const expectedSignature = generateSignature(payloadB64, secretKey);
    if (signature !== expectedSignature) {
      Logger.log('署名検証失敗');
      return null;
    }
    
    // ペイロード復号
    const payloadJson = Utilities.base64DecodeWebSafe(payloadB64);
    const payloadStr = Utilities.newBlob(payloadJson).getDataAsString('UTF-8');
    const payload = JSON.parse(payloadStr);
    
    // 有効期限チェック
    if (Date.now() > payload.exp) {
      Logger.log('トークン期限切れ');
      return null;
    }
    
    return payload;
  } catch (error) {
    Logger.log('トークン検証エラー: ' + error.toString());
    return null;
  }
}

// 管理者用URL生成メニュー
function generateReceiptURL() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ledger = ss.getSheetByName('ledger');
    const activeRange = ledger.getActiveRange();
    
    if (!activeRange || activeRange.getRow() < 2) {
      SpreadsheetApp.getUi().alert('台帳の行を選択してください');
      return;
    }
    
    const row = activeRange.getRow();
    const rowData = ledger.getRange(row, 1, 1, 11).getValues()[0];
    
    // 必要なデータの確認
    if (!rowData[0] || !rowData[1]) {
      SpreadsheetApp.getUi().alert('order_idまたはemailが設定されていません');
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
    
    // **最新のデプロイURLに修正**
    const url = 'https://script.google.com/macros/s/AKfycbw6dCOp2r_WnmECfOxZ66Ck0e3WXpXpEMAzeIokDHK-jPX4rXpz96URL0AXwpFMS1QE/exec?token=' + encodeURIComponent(token);
    
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
Email: ${config.from_email || 'info@emergence-japan.com'}
TEL: ${config.tel || '090-1575-5068'}
    `;
    
    GmailApp.sendEmail(
      rowData[1],
      subject,
      body,
      {
        name: config.from_name || config.company,
        replyTo: config.from_email || 'info@emergence-japan.com'
      }
    );
    
    SpreadsheetApp.getUi().alert('領収書発行URLを送信しました:\n' + url);
    Logger.log('URL生成完了: ' + url);
    
  } catch (error) {
    Logger.log('URL生成エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

// Web側 GET処理（フォーム表示）
function doGet(e) {
  try {
    Logger.log('doGet開始: ' + JSON.stringify(e.parameter));
    
    const token = e && e.parameter ? e.parameter.token : null;
    
    if (!token) {
      Logger.log('トークンが提供されていません');
      return HtmlService.createHtmlOutput(`
        <h1>エラー</h1>
        <p>無効なリンクです。</p>
        <p>お問い合わせ: info@emergence-japan.com</p>
      `);
    }
    
    const config = loadConfig();
    const payload = verifyToken(token, config.secret_key);
    
    if (!payload) {
      Logger.log('トークン検証失敗');
      return HtmlService.createHtmlOutput(`
        <h1>エラー</h1>
        <p>リンクが無効または期限切れです。</p>
        <p>新しいリンクが必要な場合は、お問い合わせください。</p>
        <p>Email: info@emergence-japan.com</p>
      `);
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
      Logger.log('注文が見つかりません: ' + payload.order_id);
      return HtmlService.createHtmlOutput(`
        <h1>エラー</h1>
        <p>注文が見つかりません。</p>
        <p>お問い合わせ: info@emergence-japan.com</p>
      `);
    }
    
    const htmlTemplate = HtmlService.createTemplateFromFile('form');
    htmlTemplate.token = token;
    htmlTemplate.amount = config.amount_taxin || 11000;
    htmlTemplate.description = config.description || '生成ＡＩセミナー料として';
    htmlTemplate.company = config.company || 'エマージェンス・ジャパン合同会社';
    
    Logger.log('フォーム表示成功');
    
    return htmlTemplate.evaluate()
      .setTitle('領収書発行フォーム - エマージェンス・ジャパン合同会社')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    Logger.log('doGetエラー: ' + error.toString());
    return HtmlService.createHtmlOutput(`
      <h1>システムエラー</h1>
      <p>一時的な問題が発生しています。</p>
      <p>しばらく時間をおいて再試行してください。</p>
      <p>問題が続く場合は、お問い合わせください。</p>
      <p>Email: info@emergence-japan.com</p>
    `);
  }
}

// Web側 POST処理（PDF生成・送信）
function doPost(e) {
  try {
    Logger.log('doPost開始: ' + JSON.stringify(e.parameter));
    
    const token = e && e.parameter ? e.parameter.token : null;
    const billTo = e && e.parameter ? e.parameter.bill_to : null;
    const customerEmail = e && e.parameter ? (e.parameter.customer_email || '') : '';
    
    if (!token || !billTo) {
      throw new Error('必要な情報が不足しています（トークンまたは宛名）');
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
      throw new Error('注文が見つかりません: ' + payload.order_id);
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
      orderRow[8] || now, // pay_date
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
Email: ${config.from_email || 'info@emergence-japan.com'}
TEL: ${config.tel || '090-1575-5068'}
      `;
      
      GmailApp.sendEmail(
        targetEmail,
        subject,
        body,
        {
          name: config.from_name || config.company,
          replyTo: config.from_email || 'info@emergence-japan.com',
          attachments: [pdfBlob]
        }
      );
    }
    
    Logger.log('PDF生成・送信完了: ' + receiptNo);
    
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
  try {
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
  } catch (error) {
    Logger.log('領収書番号取得エラー: ' + error.toString());
    return 0;
  }
}

// PDF生成
function generateReceiptPDF(params) {
  try {
    const config = params.config;
    const template = DriveApp.getFileById(config.doc_template_id);
    
    // テンプレート複製
    const copy = template.makeCopy(`temp_${params.receiptNo}_${Date.now()}`);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    
    // プレースホルダー置換
    const payDate = new Date(params.payDate);
    const issueDate = new Date();
    
    const replacements = {
      '{{RECEIPT_NO}}': params.receiptNo,
      '{{BILL_TO}}': params.billTo,
      '{{AMOUNT_TAXIN}}': '¥' + Number(config.amount_taxin || 11000).toLocaleString(),
      '{{AMOUNT_TAXOUT}}': '¥' + Number(config.amount_taxout || 10000).toLocaleString(),
      '{{TAX}}': '¥' + Number(config.tax || 1000).toLocaleString(),
      '{{DESCRIPTION}}': config.description || '生成ＡＩセミナー料として',
      '{{PAY_DATE}}': payDate.getFullYear() + '年' + (payDate.getMonth() + 1) + '月' + payDate.getDate() + '日',
      '{{ISSUE_DATE}}': issueDate.getFullYear() + '年' + (issueDate.getMonth() + 1) + '月' + issueDate.getDate() + '日',
      '{{COMPANY}}': config.company || 'エマージェンス・ジャパン合同会社',
      '{{ADDRESS}}': config.address || '〒550-0014 大阪府大阪市西区北堀江４−４−６',
      '{{TEL}}': config.tel || '090-1575-5068',
      '{{EMAIL}}': config.from_email || 'info@emergence-japan.com',
      '{{INVOICE_NO}}': config.invoice_no || 'T7120003014136',
      '{{PAY_METHOD}}': config.pay_method || '銀行振込',
      '{{REISSUE_MARK}}': params.isReissue ? '【再発行】' : ''
    };
    
    // 一括置換
    for (const [placeholder, value] of Object.entries(replacements)) {
      body.replaceText(placeholder, value);
    }
    
    doc.saveAndClose();
    
    // PDF変換
    const pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
    
    // 一時ファイル削除
    DriveApp.getFileById(copy.getId()).setTrashed(true);
    
    return pdfBlob;
  } catch (error) {
    Logger.log('PDF生成エラー: ' + error.toString());
    throw error;
  }
}

// メニュー作成（スプレッドシート用）
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('領収書管理')
    .addItem('選択行のURL生成', 'generateReceiptURL')
    .addItem('設定確認', 'checkConfig')
    .addToUi();
}

// 設定確認用関数（新規追加）
function checkConfig() {
  try {
    const config = loadConfig();
    const ui = SpreadsheetApp.getUi();
    
    let message = '設定確認結果:\n\n';
    const requiredFields = {
      'secret_key': 'シークレットキー',
      'pdf_folder_id': 'PDFフォルダID',
      'doc_template_id': 'ドキュメントテンプレートID',
      'company': '会社名',
      'amount_taxin': '税込金額'
    };
    
    for (const [key, label] of Object.entries(requiredFields)) {
      message += `${label}: ${config[key] ? '✓設定済み' : '✗未設定'}\n`;
    }
    
    ui.alert('設定確認', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('設定確認エラー', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}