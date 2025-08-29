// 領収書発行システム Apps Script コード（超シンプル版）
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

// Web側 GET処理（フォーム表示）+ POST判定
function doGet(e) {
  try {
    Logger.log('=== doGet開始 ===');
    Logger.log('Parameters: ' + JSON.stringify(e.parameter));
    
    // POSTアクションの場合
    if (e.parameter && e.parameter.action === 'post') {
      Logger.log('POST action detected');
      return handlePost(e);
    }
    
    const token = e && e.parameter ? e.parameter.token : null;
    
    if (!token) {
      Logger.log('トークンが提供されていません');
      return HtmlService.createHtmlOutput(`
        <div style="max-width: 600px; margin: 50px auto; padding: 30px; font-family: sans-serif; text-align: center; background: #f8d7da; border-radius: 12px;">
          <h1 style="color: #58151c;">エラー</h1>
          <p>無効なリンクです。</p>
          <p>お問い合わせ: info@emergence-japan.com</p>
        </div>
      `);
    }
    
    const config = loadConfig();
    
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
      <div style="max-width: 600px; margin: 50px auto; padding: 30px; font-family: sans-serif; text-align: center; background: #f8d7da; border-radius: 12px;">
        <h1 style="color: #58151c;">システムエラー</h1>
        <p>一時的な問題が発生しています。</p>
        <p>お問い合わせ: info@emergence-japan.com</p>
      </div>
    `);
  }
}

// POST処理ロジック（超シンプル版）
function handlePost(e) {
  try {
    Logger.log('=== handlePost開始 ===');
    Logger.log('POST parameters: ' + JSON.stringify(e.parameter));
    
    const billTo = e && e.parameter ? e.parameter.bill_to : null;
    
    if (!billTo) {
      throw new Error('宛名が入力されていません');
    }
    
    Logger.log('宛名: ' + billTo);
    
    // 超シンプルな成功ページを返す
    return HtmlService.createHtmlOutput(`
      <div style="max-width: 600px; margin: 50px auto; padding: 30px; font-family: sans-serif; text-align: center; background: #d1e7dd; border-radius: 12px;">
        <h2 style="color: #0a3622;">✅ 領収書を発行いたしました</h2>
        <p><strong>宛名:</strong> ${billTo}</p>
        <p><strong>領収書番号:</strong> EJ-2025-00001</p>
        <p style="margin-top: 20px;">
          <a href="#" style="background: #198754; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px;">📄 テスト用リンク</a>
        </p>
      </div>
    `);
    
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    return HtmlService.createHtmlOutput(`
      <div style="max-width: 600px; margin: 50px auto; padding: 30px; font-family: sans-serif; text-align: center; background: #f8d7da; border-radius: 12px;">
        <h2 style="color: #58151c;">❌ エラーが発生しました</h2>
        <p>${error.toString()}</p>
      </div>
    `);
  }
}

// Web側 POST処理（従来版・念のため残す）
function doPost(e) {
  Logger.log('=== doPost従来版が呼ばれました ===');
  return handlePost(e);
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