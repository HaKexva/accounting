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
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);
      min-height: 100vh;
      display: flex;
      justify-content: center;
      align-items: center;
      margin: 0;
      padding: 20px;
    }
    
    .unauthorized-container {
      background: white;
      border-radius: 12px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
      padding: 40px;
      max-width: 500px;
      width: 100%;
      text-align: center;
    }
    
    .error-icon {
      width: 80px;
      height: 80px;
      background: #e74c3c;
      border-radius: 50%;
      margin: 0 auto 20px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 36px;
      color: white;
    }
    
    .error-title {
      font-size: 28px;
      font-weight: 600;
      color: #e74c3c;
      margin-bottom: 15px;
    }
    
    .error-message {
      font-size: 16px;
      color: #7f8c8d;
      margin-bottom: 30px;
      line-height: 1.6;
    }
    
    .contact-info {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 8px;
      margin: 20px 0;
    }
    
    .contact-info h3 {
      margin: 0 0 10px 0;
      color: #2c3e50;
      font-size: 18px;
    }
    
    .contact-info p {
      margin: 5px 0;
      color: #555;
    }
    
    .back-button {
      display: inline-block;
      padding: 12px 30px;
      background: #3498db;
      color: white;
      text-decoration: none;
      border-radius: 6px;
      font-weight: 500;
      transition: background 0.3s;
    }
    
    .back-button:hover {
      background: #2980b9;
    }
  </style>
</head>
<body>
  <div class="unauthorized-container">
    <div class="error-icon">🚫</div>
    <h1 class="error-title">訪問被拒絕</h1>
    <p class="error-message">
      很抱歉，您的 Google 帳號未被授權訪問此記帳系統。<br>
      只有經過管理員授權的用戶才能使用此系統。
    </p>
    
    <div class="contact-info">
      <h3>需要訪問權限？</h3>
      <p>請聯繫系統管理員</p>
      <p>Email: ray120424@gmail.com</p>
      <p>請提供您的 Google 帳號郵箱地址</p>
    </div>
    
    <a href="/accounting/login.html" class="back-button">返回登入頁面</a>
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