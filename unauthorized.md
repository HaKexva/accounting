---
layout: default
title: Unauthorized
---

<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>未授權 - Accounting System</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft JhengHei", Roboto, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      justify-content: center;
      align-items: center;
      margin: 0;
      padding: 20px;
      position: relative;
      overflow: hidden;
    }
    
    body::before {
      content: '';
      position: absolute;
      top: -50%;
      left: -50%;
      width: 200%;
      height: 200%;
      background: radial-gradient(circle, rgba(255,255,255,0.05) 1px, transparent 1px);
      background-size: 50px 50px;
      animation: float 20s linear infinite;
    }
    
    @keyframes float {
      0% { transform: translate(0, 0); }
      100% { transform: translate(50px, 50px); }
    }
    
    .unauthorized-container {
      background: rgba(255, 255, 255, 0.98);
      border-radius: 20px;
      box-shadow: 0 30px 80px rgba(0,0,0,0.2), 0 10px 40px rgba(0,0,0,0.1);
      padding: 50px;
      max-width: 520px;
      width: 100%;
      text-align: center;
      position: relative;
      backdrop-filter: blur(10px);
      animation: slideUp 0.5s ease-out;
    }
    
    @keyframes slideUp {
      from {
        opacity: 0;
        transform: translateY(30px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }
    
    .error-icon {
      width: 100px;
      height: 100px;
      background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
      border-radius: 50%;
      margin: 0 auto 25px;
      display: flex;
      align-items: center;
      justify-content: center;
      position: relative;
      animation: pulse 2s ease-in-out infinite;
    }
    
    @keyframes pulse {
      0%, 100% { transform: scale(1); }
      50% { transform: scale(1.05); }
    }
    
    .error-icon::before {
      content: '';
      position: absolute;
      width: 100%;
      height: 100%;
      border-radius: 50%;
      background: inherit;
      filter: blur(20px);
      opacity: 0.4;
      z-index: -1;
    }
    
    .error-icon svg {
      width: 50px;
      height: 50px;
      fill: white;
    }
    
    .error-title {
      font-size: 32px;
      font-weight: 700;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
      margin-bottom: 20px;
      letter-spacing: -0.5px;
    }
    
    .error-message {
      font-size: 17px;
      color: #64748b;
      margin-bottom: 35px;
      line-height: 1.7;
      font-weight: 400;
    }
    
    .contact-info {
      background: linear-gradient(135deg, #f6f8fb 0%, #f1f5f9 100%);
      padding: 25px;
      border-radius: 16px;
      margin: 25px 0 35px;
      border: 1px solid rgba(0,0,0,0.05);
      position: relative;
    }
    
    .contact-info::before {
      content: '📧';
      position: absolute;
      top: -15px;
      left: 50%;
      transform: translateX(-50%);
      background: white;
      padding: 0 10px;
      font-size: 24px;
    }
    
    .contact-info h3 {
      margin: 10px 0 15px 0;
      color: #1e293b;
      font-size: 19px;
      font-weight: 600;
    }
    
    .contact-info p {
      margin: 8px 0;
      color: #64748b;
      font-size: 15px;
    }
    
    .contact-info p.email {
      font-weight: 600;
      color: #667eea;
      font-size: 16px;
      margin: 12px 0;
    }
    
    .back-button {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 14px 32px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      text-decoration: none;
      border-radius: 12px;
      font-weight: 600;
      font-size: 16px;
      transition: all 0.3s ease;
      box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
      position: relative;
      overflow: hidden;
    }
    
    .back-button::before {
      content: '';
      position: absolute;
      top: 0;
      left: -100%;
      width: 100%;
      height: 100%;
      background: rgba(255, 255, 255, 0.2);
      transition: left 0.3s ease;
    }
    
    .back-button:hover {
      transform: translateY(-2px);
      box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
    }
    
    .back-button:hover::before {
      left: 100%;
    }
    
    .back-button svg {
      width: 18px;
      height: 18px;
      fill: currentColor;
    }
    
    @media (max-width: 600px) {
      .unauthorized-container {
        padding: 35px 25px;
      }
      
      .error-title {
        font-size: 26px;
      }
      
      .error-message {
        font-size: 15px;
      }
    }
  </style>
</head>
<body>
  <div class="unauthorized-container">
    <div class="error-icon">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
        <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 15v-2h2v2h-2zm0-4V7h2v6h-2z"/>
      </svg>
    </div>
    <h1 class="error-title">訪問被拒絕</h1>
    <p class="error-message">
      很抱歉，您的 Google 帳號未被授權訪問此記帳系統。<br>
      只有經過管理員授權的用戶才能使用此系統。
    </p>
    
    <div class="contact-info">
      <h3>需要訪問權限？</h3>
      <p>請聯繫系統管理員申請授權</p>
      <p class="email">📨 ray120424@gmail.com</p>
      <p>請提供您的 Google 帳號郵箱地址</p>
    </div>
    
    <a href="/accounting/login.html" class="back-button">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
        <path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/>
      </svg>
      返回登入頁面
    </a>
  </div>

  <script>
    // Clear any existing session
    localStorage.removeItem('auth_session');
    
    // Disable Google auto-select
    if (typeof google !== 'undefined' && google.accounts && google.accounts.id) {
      google.accounts.id.disableAutoSelect();
    }
  </script>
</body>
</html>