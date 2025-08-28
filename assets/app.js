
const WEB_APP_URL =
  'https://script.google.com/macros/s/AKfycbzZ7D3ZlMmWDujXCTN7xE4wbM3om5aJXkjtt9kJpGfN9wpycwXHff7P87BmvvtcrMZ11A/exec';
let LAST_DATA = null;
let IS_EDIT_MODE = true;

// New: per-section configuration
const SECTION_CONFIG = {
  '當月收入': { editable: true },
  '當月支出預算': { editable: true },
  '隔月預計支出': { editable: true },

};

// debounce helper for autosave
function debounce(fn, wait) {
  let t;
  return function(...args) {
    clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), wait);
  };
}

// 等待 DOM 載入完成後執行
document.addEventListener('DOMContentLoaded', function() {
  fetchData();
});

async function sendSectionUpdate(sectionTitle, headers, rows) {
  // pick target section if remapped
  const cfg = SECTION_CONFIG[sectionTitle] || {};
  const target = cfg.targetSection || sectionTitle;
  const payload = {
    action: 'updateSection',
    section: target,
    headers,
    rows,
  };
  try {
    const resp = await fetch(WEB_APP_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      keepalive: true,
    });
    if (!resp.ok) throw new Error(`HTTP error! status: ${resp.status}`);
    return await resp.json();
  } catch (err) {
    // 後援：避免 CORS/Preflight 問題，送出 opaque 請求
    try {
      await fetch(WEB_APP_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload),
        keepalive: true,
      });
      return { ok: true, needsRefetch: true };
    } catch (e2) {
      throw err;
    }
  }
}

function fetchData() {
  const container = document.getElementById('data-container');
  if (!container) {
    console.error('找不到 data-container 元素');
    return;
  }

  container.innerHTML = '<p>正在載入記帳資料...</p>';

  fetch(WEB_APP_URL)
    .then((response) => {
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return response.json();
    })
    .then((result) => {
      if (result && result.data) {
        LAST_DATA = result.data;
        displayAccountingData(LAST_DATA);
      } else {
        throw new Error('資料格式不正確');
      }
    })
    .catch((error) => {
      console.error('載入資料時發生錯誤:', error);
      container.innerHTML = `
        <div style="color: red; padding: 10px; background-color: #ffe6e6; border-radius: 5px;">
          <h3>載入失敗</h3>
          <p>無法載入記帳資料，請稍後再試。</p>
          <p>錯誤訊息: ${error.message}</p>
        </div>
      `;
    });
}

function displayAccountingData(data) {
  const container = document.getElementById('data-container');
  if (!container) return;

  container.innerHTML = '';

  const mainTitle = document.createElement('h1');
  mainTitle.textContent = '記帳資料總覽';
  mainTitle.style.textAlign = 'center';
  mainTitle.style.marginBottom = '30px';
  mainTitle.style.color = '#2c3e50';
  container.appendChild(mainTitle);

  // 顯示當月收入
  if (data['當月收入'] && data['當月收入'].length > 0) {
    displaySection(container, '當月收入', data['當月收入'], 'income');
  }
  // 顯示當月支出
  if (data['當月支出預算'] && data['當月支出預算'].length > 0) {
    displaySection(container, '當月支出預算', data['當月支出預算'], 'expense');
  }
  // 顯示隔月預計支出
  if (data['隔月預計支出'] && data['隔月預計支出'].length > 0) {
    displaySection(container, '隔月預計支出', data['隔月預計支出'], 'future');
  }
}

function displaySection(container, title, items, type) {
  const cfg = SECTION_CONFIG[title] || { editable: true };

  const section = document.createElement('div');
  section.className = 'accounting-section';
  section.style.marginBottom = '30px';
  section.style.padding = '20px';
  section.style.borderRadius = '8px';
  section.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';

  if (type === 'income') {
    section.style.backgroundColor = '#e8f5e8';
    section.style.borderLeft = '4px solid #27ae60';
  } else if (type === 'expense') {
    section.style.backgroundColor = '#ffeaea';
    section.style.borderLeft = '4px solid #e74c3c';
  } else {
    section.style.backgroundColor = '#f0f8ff';
    section.style.borderLeft = '4px solid #3498db';
  }

  const sectionTitle = document.createElement('h2');
  sectionTitle.textContent = title + ' ▼';
  sectionTitle.style.marginBottom = '15px';
  sectionTitle.style.color = '#2c3e50';
  sectionTitle.style.cursor = 'pointer';
  sectionTitle.setAttribute('tabindex', '0');
  sectionTitle.setAttribute('role', 'button');
  sectionTitle.setAttribute('aria-expanded', 'true');
  section.appendChild(sectionTitle);

  const contentDiv = document.createElement('div');
  contentDiv.style.display = 'block';

  // 工具列
  const controlsDiv = document.createElement('div');
  controlsDiv.style.display = 'flex';
  controlsDiv.style.gap = '8px';
  controlsDiv.style.margin = '8px 0 12px 0';
  controlsDiv.style.flexWrap = 'wrap';

  const undoBtn = document.createElement('button');
  undoBtn.textContent = 'Undo';
  undoBtn.style.padding = '6px 10px';
  undoBtn.style.border = '1px solid #aaa';
  undoBtn.style.background = '#f1f1f1';
  undoBtn.style.borderRadius = '6px';
  undoBtn.style.cursor = 'pointer';

  const redoBtn = document.createElement('button');
  redoBtn.textContent = 'Redo';
  redoBtn.style.padding = '6px 10px';
  redoBtn.style.border = '1px solid #3498db';
  redoBtn.style.background = '#e3f2fd';
  redoBtn.style.borderRadius = '6px';
  redoBtn.style.cursor = 'pointer';

  const addRowBtn = document.createElement('button');
  addRowBtn.textContent = '新增列';
  addRowBtn.style.padding = '6px 10px';
  addRowBtn.style.border = '1px solid #3498db';
  addRowBtn.style.background = '#e3f2fd';
  addRowBtn.style.borderRadius = '6px';
  addRowBtn.style.cursor = 'pointer';

  const deleteRowBtn = document.createElement('button');
  deleteRowBtn.textContent = '刪除列';
  deleteRowBtn.style.padding = '6px 10px';
  deleteRowBtn.style.border = '1px solid #e74c3c';
  deleteRowBtn.style.background = '#ffebee';
  deleteRowBtn.style.borderRadius = '6px';
  deleteRowBtn.style.cursor = 'pointer';

  const autosaveHint = document.createElement('span');
  autosaveHint.textContent = '';
  autosaveHint.style.alignSelf = 'center';
  autosaveHint.style.color = '#666';

  // 只有可編輯區塊顯示增刪、取消、儲存
  if (cfg.editable) {
    controlsDiv.appendChild(undoBtn);
    controlsDiv.appendChild(redoBtn);
    controlsDiv.appendChild(addRowBtn);
    controlsDiv.appendChild(deleteRowBtn);
    controlsDiv.appendChild(autosaveHint);
  }
  contentDiv.appendChild(controlsDiv);

  // 當月支出預算：移除表單呈現，僅顯示紀錄區（唯讀）

  // 表格容器
  const tableContainer = document.createElement('div');
  tableContainer.style.width = '100%';
  tableContainer.style.overflowX = 'auto';
  tableContainer.style.border = '1px solid #ddd';
  tableContainer.style.borderRadius = '4px';
  
  // 手機端觸控優化 - 將在表格容器添加到DOM後執行
  let touchHint = null;
  if (window.innerWidth < 768) {
    tableContainer.className = 'table-container';
    
    // 創建觸控滾動提示，但暫時不添加到DOM
    touchHint = document.createElement('div');
    touchHint.style.cssText = `
      text-align: center;
      color: #666;
      font-size: 12px;
      margin: 8px 0;
      padding: 6px;
      background: #f8f9fa;
      border-radius: 4px;
      display: none;
    `;
    touchHint.textContent = '👆 雙指縮放或左右滑動查看更多內容';
  }

  const table = document.createElement('table');
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';
  table.style.marginTop = '0';
  table.style.tableLayout = 'fixed';
  table.style.minWidth = '800px';

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  const headers = Object.keys(items[0] || {});
  headers.forEach(header => {
    const th = document.createElement('th');
    th.textContent = header;
    th.style.padding = '12px 15px';
    th.style.textAlign = 'left';
    th.style.borderBottom = '2px solid #ddd';
    th.style.backgroundColor = '#f8f9fa';
    th.style.whiteSpace = 'nowrap';
    th.style.minWidth = '80px';
    if (header.includes('金額') || header.includes('預算') || header.includes('實際消費金額')) {
      th.style.width = '100px'; th.style.textAlign = 'right';
    } else if (header.includes('項目')) {
      th.style.width = '150px';
    } else if (header.includes('細節') || header.includes('備註')) {
      th.style.width = '200px'; th.style.whiteSpace = 'normal'; th.style.wordWrap = 'break-word';
    } else if (header.includes('日期')) {
      th.style.width = '100px';
    } else if (header.includes('類別')) {
      th.style.width = '120px';
    } else if (header.includes('方式')) {
      th.style.width = '100px';
    } else {
      th.style.width = '120px';
    }
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  items.forEach((item, index) => {
    const row = document.createElement('tr');
    row.style.backgroundColor = index % 2 === 0 ? '#ffffff' : '#f8f9fa';
    row.dataset.rowIndex = index;
    headers.forEach(header => {
      const td = document.createElement('td');
      td.textContent = item[header] || '';
      td.style.padding = '10px 15px';
      td.style.borderBottom = '1px solid #ddd';
      td.style.verticalAlign = 'top';
      td.style.fontSize = '14px';
      td.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
      if (cfg.editable && IS_EDIT_MODE) {
        td.contentEditable = 'true';
        td.style.outline = '1px dashed rgba(0,0,0,0.2)';
        td.style.backgroundColor = 'rgba(255,255,0,0.06)';
      }
      if (header.includes('金額') || header.includes('預算') || header.includes('實際消費金額')) {
        td.style.fontWeight = 'bold';
        td.style.textAlign = 'right';
        td.style.fontFamily = 'monospace';
        if (type === 'income') {
          td.style.color = '#27ae60';
        } else if (type === 'expense') {
          td.style.color = '#e74c3c';
        }
      } else if (header.includes('細節') || header.includes('備註')) {
        td.style.whiteSpace = 'normal';
        td.style.wordWrap = 'break-word';
        td.style.maxWidth = '200px';
        td.style.lineHeight = '1.4';
      } else if (header.includes('項目')) {
        td.style.fontWeight = '500';
        td.style.color = '#2c3e50';
      } else if (header.includes('日期')) {
        td.style.fontFamily = 'monospace';
        td.style.color = '#2c3e50';
      }
      row.appendChild(td);
    });
    tbody.appendChild(row);
  });
  table.appendChild(tbody);

  const tfoot = document.createElement('tfoot');
  const totalRow = document.createElement('tr');
  totalRow.style.backgroundColor = '#f8f9fa';
  totalRow.style.fontWeight = 'bold';
  headers.forEach((header, i) => {
    const td = document.createElement('td');
    td.style.padding = '10px 15px';
    td.style.borderTop = '2px solid #aaa';
    td.style.fontWeight = 'bold';
    td.style.fontSize = '14px';
    td.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    if (i === 0) {
      td.textContent = '總計';
    } else if (header.includes('金額') || header.includes('預算') || header.includes('實際消費金額')) {
      td.style.textAlign = 'right';
      td.style.fontFamily = 'monospace';
      td.dataset.totalFor = header;
    } else {
      td.textContent = '';
    }
    totalRow.appendChild(td);
  });
  tfoot.appendChild(totalRow);
  table.appendChild(tfoot);

  function recalcTotals() {
    headers.forEach((header) => {
      if (header.includes('金額') || header.includes('預算') || header.includes('實際消費金額')) {
        let sum = 0;
        Array.from(tbody.querySelectorAll('tr')).forEach(tr => {
          const idx = headers.indexOf(header);
          const cell = tr.children[idx];
          const num = parseFloat((cell?.innerText || '').replace(/[^\d.-]/g, '')) || 0;
          sum += num;
        });
        const totalCell = totalRow.children[headers.indexOf(header)];
        if (totalCell) totalCell.textContent = sum.toLocaleString();
      }
    });
  }
  setTimeout(recalcTotals, 0);
  tbody.addEventListener('input', recalcTotals);

  // History (Undo/Redo)
  let historyStack = [];
  let futureStack = [];
  let lastSnapshot = getSnapshot();

  function getSnapshot() {
    const rows = [];
    Array.from(tbody.querySelectorAll('tr')).forEach(tr => {
      const obj = {};
      const cells = Array.from(tr.querySelectorAll('td'));
      headers.forEach((h, i) => {
        obj[h] = (cells[i]?.innerText || '').trim();
      });
      rows.push(obj);
    });
    return rows;
  }
  function applySnapshot(snapshot) {
    // Rebuild tbody to match snapshot length
    tbody.innerHTML = '';
    snapshot.forEach((rowObj, idx) => {
      const tr = document.createElement('tr');
      tr.style.backgroundColor = idx % 2 === 0 ? '#ffffff' : '#f8f9fa';
      tr.dataset.rowIndex = idx;
      headers.forEach(h => {
        const td = document.createElement('td');
        td.textContent = rowObj[h] || '';
        td.style.padding = '10px 15px';
        td.style.borderBottom = '1px solid #ddd';
        td.style.verticalAlign = 'top';
        td.style.fontSize = '14px';
        td.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
        if (cfg.editable && IS_EDIT_MODE) {
          td.contentEditable = 'true';
          td.style.outline = '1px dashed rgba(0,0,0,0.2)';
          td.style.backgroundColor = 'rgba(255,255,0,0.06)';
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    recalcTotals();
  }

  // Input: push previous state to history, clear future, autosave
  if (cfg.editable) {
    tbody.addEventListener('input', () => {
      historyStack.push(lastSnapshot);
      futureStack = [];
      lastSnapshot = getSnapshot();
    });

    undoBtn.addEventListener('click', () => {
      if (historyStack.length === 0) return;
      const current = getSnapshot();
      const prev = historyStack.pop();
      futureStack.push(current);
      applySnapshot(prev);
      lastSnapshot = getSnapshot();
      autosaveHint.textContent = '自動儲存中...';
      debouncedAutosave();
    });

    redoBtn.addEventListener('click', () => {
      if (futureStack.length === 0) return;
      const current = getSnapshot();
      const next = futureStack.pop();
      historyStack.push(current);
      applySnapshot(next);
      lastSnapshot = getSnapshot();
      autosaveHint.textContent = '自動儲存中...';
      debouncedAutosave();
    });
  }

  tableContainer.appendChild(table);
  contentDiv.appendChild(tableContainer);

  // 手機端觸控優化 - 在表格容器添加到DOM後執行
  if (window.innerWidth < 768 && touchHint) {
    try {
      // 確保表格容器已經有父節點
      if (tableContainer.parentNode) {
        tableContainer.parentNode.insertBefore(touchHint, tableContainer);
        
        // 檢測觸控滾動
        let isScrolling = false;
        tableContainer.addEventListener('scroll', () => {
          if (!isScrolling) {
            isScrolling = true;
            touchHint.style.display = 'block';
            setTimeout(() => {
              touchHint.style.display = 'none';
              isScrolling = false;
            }, 2000);
          }
        });
        
        // 添加觸控手勢支援
        let startX = 0;
        let startY = 0;
        
        tableContainer.addEventListener('touchstart', (e) => {
          startX = e.touches[0].clientX;
          startY = e.touches[0].clientY;
        });
        
        tableContainer.addEventListener('touchmove', (e) => {
          if (!startX || !startY) return;
          
          const deltaX = e.touches[0].clientX - startX;
          const deltaY = e.touches[0].clientY - startY;
          
          // 水平滑動優先
          if (Math.abs(deltaX) > Math.abs(deltaY)) {
            e.preventDefault();
          }
        });
      }
    } catch (error) {
      console.warn('觸控優化功能載入失敗:', error);
    }
  }

  // Mobile-friendly: card view toggle for narrow screens
  if (window.innerWidth < 768) {
    const toggle = document.createElement('button');
    toggle.textContent = '切換卡片/表格視圖';
    toggle.style.margin = '10px 0';
    toggle.style.padding = '6px 10px';
    toggle.style.border = '1px solid #3498db';
    toggle.style.background = '#e3f2fd';
    toggle.style.borderRadius = '6px';
    toggle.style.cursor = 'pointer';
    let isCard = false;
    toggle.addEventListener('click', () => {
      isCard = !isCard;
      if (isCard) {
        tableContainer.style.display = 'none';
        renderCards();
      } else {
        tableContainer.style.display = 'block';
        if (cardsDiv) cardsDiv.remove();
      }
    });
    contentDiv.appendChild(toggle);

    var cardsDiv = null;
    function renderCards() {
      if (cardsDiv) cardsDiv.remove();
      cardsDiv = document.createElement('div');
      cardsDiv.className = 'card-view';
      
      items.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = 'card';
        
        headers.forEach(h => {
          const row = document.createElement('div');
          row.className = 'card-row';
          
          const k = document.createElement('div');
          k.className = 'card-label';
          k.textContent = h;
          
          const v = document.createElement('div');
          v.className = 'card-value';
          v.textContent = item[h] || '';
          
          // 為金額欄位添加特殊樣式
          if (h.includes('金額') || h.includes('預算') || h.includes('實際消費金額')) {
            v.classList.add('amount');
            if (type === 'income') {
              v.classList.add('income');
            } else if (type === 'expense') {
              v.classList.add('expense');
            }
          }
          
          row.appendChild(k);
          row.appendChild(v);
          card.appendChild(row);
        });
        
        // 為卡片添加點擊效果
        card.style.cursor = 'pointer';
        card.addEventListener('click', () => {
          // 高亮選中的卡片
          cardsDiv.querySelectorAll('.card').forEach(c => c.style.border = '1px solid #ddd');
          card.style.border = '2px solid #3498db';
          card.style.boxShadow = '0 4px 12px rgba(52, 152, 219, 0.3)';
        });
        
        cardsDiv.appendChild(card);
      });
      
      contentDiv.appendChild(cardsDiv);
    }
    
    // 監聽視窗大小變化，自動切換視圖
    window.addEventListener('resize', () => {
      if (window.innerWidth >= 768 && isCard) {
        isCard = false;
        tableContainer.style.display = 'block';
        if (cardsDiv) cardsDiv.remove();
      }
    });
  }

  section.appendChild(contentDiv);
  container.appendChild(section);

  // 事件：新增列
  addRowBtn.addEventListener('click', () => {
    if (!cfg.editable) return;
    const newRow = document.createElement('tr');
    const rowIndex = tbody.children.length;
    newRow.style.backgroundColor = rowIndex % 2 === 0 ? '#ffffff' : '#f8f9fa';
    newRow.dataset.rowIndex = rowIndex;
    headers.forEach(header => {
      const td = document.createElement('td');
      td.textContent = '';
      td.style.padding = '10px 15px';
      td.style.borderBottom = '1px solid #ddd';
      td.style.verticalAlign = 'top';
      td.style.fontSize = '14px';
      td.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
      td.contentEditable = 'true';
      td.style.outline = '1px dashed rgba(0,0,0,0.2)';
      td.style.backgroundColor = 'rgba(255,255,0,0.06)';
      newRow.appendChild(td);
    });
    // push history before modifying DOM snapshot reference
    historyStack.push(lastSnapshot);
    futureStack = [];
    tbody.appendChild(newRow);
    recalcTotals();
    lastSnapshot = getSnapshot();
  });

  // 事件：刪除列
  deleteRowBtn.addEventListener('click', () => {
    if (!cfg.editable) return;
    const rows = tbody.querySelectorAll('tr');
    if (rows.length > 0) {
      historyStack.push(lastSnapshot);
      futureStack = [];
      const lastRow = rows[rows.length - 1];
      lastRow.remove();
      recalcTotals();
      lastSnapshot = getSnapshot();
    }
  });

  // Autosave on idle (1.5s debounce)
  const debouncedAutosave = debounce(() => {
    autosaveHint.textContent = '自動儲存中...';
    // reuse sendSectionUpdate using current snapshot
    const rows = getSnapshot();
    sendSectionUpdate(title, headers, rows).then(() => {
      autosaveHint.textContent = '已自動儲存';
      setTimeout(() => (autosaveHint.textContent = ''), 1500);
    }).catch(() => {
      autosaveHint.textContent = '自動儲存失敗';
      setTimeout(() => (autosaveHint.textContent = ''), 2000);
    });
  }, 1500);
  if (cfg.editable) {
    tbody.addEventListener('input', debouncedAutosave);
  }

  // 收合/展開
  sectionTitle.addEventListener('click', function() {
    if (contentDiv.style.display === 'none') {
      contentDiv.style.display = 'block';
      sectionTitle.textContent = title + ' ▼';
      sectionTitle.setAttribute('aria-expanded', 'true');
    } else {
      contentDiv.style.display = 'none';
      sectionTitle.textContent = title + ' ▲';
      sectionTitle.setAttribute('aria-expanded', 'false');
    }
  });
  sectionTitle.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' || e.key === ' ') {
      sectionTitle.click();
    }
  });
}


