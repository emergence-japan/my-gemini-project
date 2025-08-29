// 領収書発行システム - 新プロジェクト用
// 最小限の機能で確実に動作

function doGet(e) {
  // フォーム送信の場合
  if (e.parameter && e.parameter.bill_to) {
    const billTo = e.parameter.bill_to;
    const email = e.parameter.customer_email || '';
    
    // 成功ページを表示
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html lang="ja">
      <head>
        <meta charset="UTF-8">
        <title>領収書発行完了</title>
        <style>
          body {
            font-family: sans-serif;
            max-width: 600px;
            margin: 50px auto;
            padding: 30px;
            background: #d1e7dd;
            border-radius: 12px;
            text-align: center;
          }
          h1 { color: #0a3622; }
          .info { margin: 20px 0; }
          .download-btn {
            display: inline-block;
            background: #198754;
            color: white;
            padding: 15px 30px;
            text-decoration: none;
            border-radius: 8px;
            margin-top: 20px;
          }
        </style>
      </head>
      <body>
        <h1>✅ 領収書を発行いたしました</h1>
        <div class="info">
          <p><strong>宛名:</strong> ${billTo}</p>
          <p><strong>領収書番号:</strong> EJ-2025-00001</p>
          <p><strong>金額:</strong> ¥11,000（税込）</p>
          <p><strong>但し書き:</strong> 生成ＡＩセミナー料として</p>
          ${email ? `<p><strong>メール:</strong> ${email}</p>` : ''}
        </div>
        <a href="#" class="download-btn">📄 領収書PDFをダウンロード</a>
        <p style="margin-top: 20px; font-size: 12px; color: #666;">
          エマージェンス・ジャパン合同会社
        </p>
      </body>
      </html>
    `);
  }
  
  // 通常のフォーム表示
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="UTF-8">
      <title>領収書発行フォーム</title>
      <style>
        body {
          font-family: sans-serif;
          max-width: 600px;
          margin: 50px auto;
          padding: 30px;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
        }
        .container {
          background: white;
          padding: 40px;
          border-radius: 12px;
          box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
        }
        h1 {
          color: #4c6ef5;
          text-align: center;
          margin-bottom: 30px;
        }
        .info-box {
          background: #f8f9ff;
          padding: 20px;
          border-radius: 8px;
          margin-bottom: 30px;
          border-left: 4px solid #4c6ef5;
        }
        .form-group {
          margin-bottom: 20px;
        }
        label {
          display: block;
          margin-bottom: 8px;
          font-weight: 600;
          color: #495057;
        }
        input {
          width: 100%;
          padding: 12px 16px;
          border: 2px solid #e9ecef;
          border-radius: 8px;
          font-size: 16px;
          box-sizing: border-box;
        }
        input:focus {
          outline: none;
          border-color: #4c6ef5;
        }
        button {
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
        button:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 25px rgba(76, 110, 245, 0.3);
        }
        .required { color: #dc3545; }
        small { color: #6c757d; font-size: 12px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>領収書発行フォーム</h1>
        <p style="text-align: center; margin-bottom: 30px; color: #666;">
          エマージェンス・ジャパン合同会社
        </p>
        
        <div class="info-box">
          <h3>🧾 領収書情報</h3>
          <p><strong>金額:</strong> ¥11,000（税込）</p>
          <p><strong>但し書き:</strong> 生成ＡＩセミナー料として</p>
          <p><strong>発行者:</strong> エマージェンス・ジャパン合同会社</p>
        </div>
        
        <form method="GET">
          <div class="form-group">
            <label for="bill_to">
              宛名 <span class="required">*</span>
            </label>
            <input 
              type="text" 
              id="bill_to" 
              name="bill_to" 
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
              placeholder="example@company.com"
              maxlength="100"
            >
            <small>ご入力いただくと、領収書PDFをメールでもお送りします。</small>
          </div>
          
          <button type="submit">
            領収書を発行する
          </button>
        </form>
      </div>
    </body>
    </html>
  `);
}