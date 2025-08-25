---
layout: default
title: 未授權 - Accounting System
---
很抱歉，您的 Google 帳號未被授權訪問此記帳系統。<br>
只有經過管理員授權的用戶才能使用此系統。

需要訪問權限？
請聯繫系統管理員－－Email: ray120424@gmail.com
請提供您的 Google 帳號郵箱地址
<a href="/accounting/login.html" class="back-button">返回登入頁面</a>

<script>
  // Clear any existing session
  localStorage.removeItem('auth_session');
  
  // Disable Google auto-select
  if (typeof google !== 'undefined' && google.accounts && google.accounts.id) {
    google.accounts.id.disableAutoSelect();
  }
</script>