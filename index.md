---
layout: page
title: ""
---

<style>
  .home-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: flex-start;
    min-height: 70vh;
    padding: 30px 20px;
    gap: 40px;
    width: 100%;
    max-width: 100%;
    box-sizing: border-box;
  }
  
  .home-title {
    font-family: 'Noto Sans TC', 'PingFang TC', sans-serif;
    font-size: 4rem;
    font-weight: 900;
    color: #2c3e50;
    text-align: center;
    margin: 0;
    padding: 0;
    letter-spacing: 0.3em;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
  }
  
  .buttons-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 25px;
    width: 100%;
    max-width: 400px;
    box-sizing: border-box;
  }
  
  .home-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 100%;
    padding: 50px 30px;
    font-size: 1.8rem;
    font-weight: 700;
    text-decoration: none;
    border-radius: 20px;
    transition: all 0.3s ease;
    box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    cursor: pointer;
    border: none;
    letter-spacing: 0.1em;
    box-sizing: border-box;
  }
  
  .home-btn:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 35px rgba(0,0,0,0.2);
    text-decoration: none;
  }
  
  .home-btn:active {
    transform: translateY(-2px);
  }
  
  .btn-expense {
    background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);
    color: white;
  }
  
  .btn-expense:hover {
    background: linear-gradient(135deg, #c0392b 0%, #a93226 100%);
    color: white;
  }
  
  .btn-budget {
    background: linear-gradient(135deg, #27ae60 0%, #1e8449 100%);
    color: white;
  }
  
  .btn-budget:hover {
    background: linear-gradient(135deg, #1e8449 0%, #196f3d 100%);
    color: white;
  }
  
  .btn-settings {
    background: linear-gradient(135deg, #7f8c8d 0%, #616a6b 100%);
    color: white;
  }
  
  .btn-settings:hover {
    background: linear-gradient(135deg, #616a6b 0%, #515a5a 100%);
    color: white;
  }
  
  .btn-icon {
    margin-right: 15px;
    font-size: 2rem;
  }
  
  /* 隱藏預設的頁面標題 */
  .post-title, .page-heading {
    display: none !important;
  }
  
  /* 手機版調整 */
  @media (max-width: 600px) {
    .home-title {
      font-size: 3rem;
      letter-spacing: 0.2em;
    }
    
    .home-btn {
      padding: 40px 20px;
      font-size: 1.5rem;
    }
    
    .btn-icon {
      font-size: 1.8rem;
      margin-right: 12px;
    }
  }
</style>

<div id="user-info" style="position: fixed; top: 10px; right: 10px; z-index: 1000;"></div>

<div class="home-container">
  <h1 class="home-title">記帳</h1>
  
  <div class="buttons-container">
    <a href="{{ '/expense/' | relative_url }}" class="home-btn btn-expense">
      <span class="btn-icon">💸</span>
      支出填寫
    </a>
    
    <a href="{{ '/budget_table/' | relative_url }}" class="home-btn btn-budget">
      <span class="btn-icon">📊</span>
      預算填寫
    </a>
    
    <a href="{{ '/settings/' | relative_url }}" class="home-btn btn-settings">
      <span class="btn-icon">⚙️</span>
      設定
    </a>
  </div>
</div>