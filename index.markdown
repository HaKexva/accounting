---
layout: home
---

# 記帳資料

<div id="data-container">
  <p>正在載入記帳資料...</p>
</div>

<style>
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    line-height: 1.6;
    color: #333;
  }
  
  #data-container {
    margin: 20px 0;
    padding: 20px;
    border-radius: 10px;
    background-color: #f8f9fa;
  }
  
  /* 響應式設計 */
  @media (max-width: 768px) {
    #data-container {
      padding: 10px;
    }
    
    table {
      font-size: 14px;
    }
    
    th, td {
      padding: 6px 8px !important;
    }
  }
  
  /* 表格響應式 */
  @media (max-width: 600px) {
    table {
      display: block;
      overflow-x: auto;
      white-space: nowrap;
    }
  }
</style>
