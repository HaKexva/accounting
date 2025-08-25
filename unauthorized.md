---
layout: default
title: 未授權 - Accounting System
---
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