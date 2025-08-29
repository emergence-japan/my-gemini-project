// 領収書発行システム Apps Script コード（直接HTML版）
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
    
    return config;
  } catch (error) {
    Logger.log('Config読み込みエラー: ' + error.toString());
    throw error;
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
    
    // 仮のトークン生成（実際のシステムではセキュアなトークンを生成）
    const token = 'test_token_' + Date.now();
    const expiry = new Date(Date.now() + (7 * 24 * 60 * 60 * 1000));
    
    // 台帳更新
    ledger.getRange(row, 3).setValue(token);
    ledger.getRange(row, 4).setValue(expiry);
    
    // **最新のデプロイURL**
    const url = 'https://script.google.com/macros/s/AKfycbwYps_RHDU7d3saN88G79xL_mn_b0FpMVcHh2_wWfxCdbdjVZazqvLZjloM84a2Xvzm5w/exec?token=' + encodeURIComponent(token);
    
    SpreadsheetApp.getUi().alert('領収書発行URLを生成しました:\n' + url);
    Logger.log('URL生成完了: ' + url);
    
  } catch (error) {
    Logger.log('URL生成エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

// Web側 GET処理（直接HTML版）
function doGet(e) {
  try {
    Logger.log('doGet開始: ' + JSON.stringify(e.parameter));
    
    // POSTアクションの場合
    if (e.parameter && e.parameter.action === 'post') {
      Logger.log('POST処理開始');
      const billTo = e.parameter.bill_to;
      const customerEmail = e.parameter.customer_email || '';
      
      if (!billTo) {
        return HtmlService.createHtmlOutput(`
          <div style="max-width: 600px; margin: 50px auto; padding: 30px; font-family: sans-serif; text-align: center; background: #f8d7da; border-radius: 12px;">
            <h2 style="color: #58151c;">❌ エラー</h2>
            <p>宛名を入力してください</p>
          </div>
        `);
      }
      
      Logger.log('宛名: ' + billTo);
      
      // 成功ページを返す
      return HtmlService.createHtmlOutput(`
        <div style="max-width: 600px; margin: 50px auto; padding: 30px; font-family: sans-serif; text-align: center; background: #d1e7dd; border-radius: 12px;">
          <h2 style="color: #0a3622;">✅ 領収書を発行いたしました</h2>
          <p><strong>宛名:</strong> ${billTo}</p>
          <p><strong>領収書番号:</strong> EJ-2025-00001</p>
          <p><strong>メール:</strong> ${customerEmail}</p>
          <p style="margin-top: 20px;">
            <a href="#" style="background: #198754; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px;">📄 テスト用ダウンロード</a>
          </p>
        </div>
      `);
    }
    
    // 通常のフォーム表示（HTMLテンプレートを使わない）
    const token = e.parameter ? e.parameter.token : null;
    
    if (!token) {
      return HtmlService.createHtmlOutput(`
        <div style="max-width: 600px; margin: 50px auto; padding: 30px; font-family: sans-serif; text-align: center; background: #f8d7da; border-radius: 12px;">
          <h1 style="color: #58151c;">エラー</h1>
          <p>無効なリンクです。</p>
          <p>お問い合わせ: info@emergence-japan.com</p>
        </div>
      `);
    }
    
    Logger.log('フォーム表示');
    
    // 直接HTMLでフォーム表示
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html lang="ja">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>領収書発行フォーム - エマージェンス・ジャパン合同会社</title>
        <style>
          body {
            font-family: sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            margin: 0;
            padding: 20px;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .container {
            max-width: 600px;
            background: white;
            border-radius: 12px;
            padding: 30px;
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
          }
          .header {
            text-align: center;
            margin-bottom: 30px;
          }
          .header h1 {
            color: #4c6ef5;
            margin-bottom: 10px;
          }
          .info-box {
            background: #f8f9ff;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 30px;
            border-left: 4px solid #4c6ef5;
          }
          .form-group {
            margin-bottom: 20px;
          }
          .form-group label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #495057;
          }
          .form-control {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 16px;
            box-sizing: border-box;
          }
          .form-control:focus {
            outline: none;
            border-color: #4c6ef5;
          }
          .submit-btn {
            width: 100%;
            background: linear-gradient(135deg, #4c6ef5 0%, #6c5ce7 100%);
            color: white;
            padding: 15px 24px;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
          }
          .submit-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(76, 110, 245, 0.3);
          }
          .required {
            color: #dc3545;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>領収書発行フォーム</h1>
            <p>エマージェンス・ジャパン合同会社</p>
          </div>
          
          <div class="info-box">
            <h3>🧾 領収書情報</h3>
            <p><strong>金額:</strong> ¥11,000 (税込)</p>
            <p><strong>但し書き:</strong> 生成ＡＩセミナー料として</p>
            <p><strong>発行者:</strong> エマージェンス・ジャパン合同会社</p>
          </div>
          
          <form method="GET">
            <input type="hidden" name="action" value="post">
            <input type="hidden" name="token" value="${token}">
            
            <div class="form-group">
              <label for="bill_to">
                宛名 <span class="required">*</span>
              </label>
              <input 
                type="text" 
                id="bill_to" 
                name="bill_to" 
                class="form-control" 
                placeholder="例: 株式会社サンプル　御中"
                required
                maxlength="100"
              >
            </div>
            
            <div class="form-group">
              <label for="customer_email">
                メールアドレス（任意）
              </label>
              <input 
                type="email" 
                id="customer_email" 
                name="customer_email" 
                class="form-control" 
                placeholder="example@company.com"
                maxlength="100"
              >
              <small style="color: #6c757d; font-size: 12px; margin-top: 4px; display: block;">
                ご入力いただくと、領収書PDFをメールでもお送りします。
              </small>
            </div>
            
            <button type="submit" class="submit-btn">
              領収書を発行する
            </button>
          </form>
        </div>
      </body>
      </html>
    `).setTitle('領収書発行フォーム');
      
  } catch (error) {
    Logger.log('doGetエラー: ' + error.toString());
    return HtmlService.createHtmlOutput(`
      <div style="max-width: 600px; margin: 50px auto; padding: 30px; font-family: sans-serif; text-align: center; background: #f8d7da; border-radius: 12px;">
        <h1 style="color: #58151c;">システムエラー</h1>
        <p>${error.toString()}</p>
        <p>お問い合わせ: info@emergence-japan.com</p>
      </div>
    `);
  }
}

// Web側 POST処理（念のため残す）
function doPost(e) {
  Logger.log('doPost呼び出し（GET処理に転送）');
  return doGet(e);
}

// メニュー作成（スプレッドシート用）
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('領収書管理')
    .addItem('選択行のURL生成', 'generateReceiptURL')
    .addItem('設定確認', 'checkConfig')
    .addToUi();
}

// 設定確認用関数
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