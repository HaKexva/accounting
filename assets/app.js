// get
const base = "https://script.google.com/macros/s/AKfycbyqDB4AE3jPpjT0-Py-yB6uYIQVs3zHNcMWJO31CkKNOex_DfhVmRi7Sq59LTq_S6QpjQ/exec";
const createBtn = document.createElement('button');
  createBtn.textContent = 'Create Data';
  createBtn.style.padding = '6px 10px';
  createBtn.style.border = '1px solid #aaa';
  createBtn.style.background = '#f1f1f1';
  createBtn.style.borderRadius = '6px';
  createBtn.style.cursor = 'pointer';
// const res = await fetch(
//   base,
//   {
//     method: "POST",
//     redirect: "follow",
//     keepalive: true,
//     headers: {
//       "Content-Type": "text/plain;charset=utf-8",
//     },
//     body: JSON.stringify({ 
//       name: "Create Tab", 
//     })
//   }
// );
// const data = await res.json();



// post
// const params = { name: "Show Tab Data", sheet: 2 };
// const url = `${base}?${new URLSearchParams(params)}`;
// res = await fetch(url);
// data = await res.json();






// const WEB_APP_URL =
//   'https://script.google.com/macros/s/AKfycby857iH5s40pc-1qNWqhn-76r0ZJCDlhzA9e4nG98htCtcHumCtGEEWa4CW5FpU_6nDTg/exec';
// let LAST_DATA = null;

// // New: per-section configuration
// const SECTION_CONFIG = {
//   '當月收入': {},
//   '當月支出預算': {},
//   '隔月預計支出': {},
// };

// const SECTION_HEADERS = {
//   '當月收入': ['項目', '金額', '備註'],
//   '當月支出預算': ['項目', '細節', '預算', '備註'],
//   '隔月預計支出': ['項目', '金額', '備註'],
// };


// // debounce helper for autosave
// function debounce(fn, wait) {
//   let t;
//   return function(...args) {
//     clearTimeout(t);
//     t = setTimeout(() => fn.apply(this, args), wait);
//   };
// }

// // 等待 DOM 載入完成後執行
// document.addEventListener('DOMContentLoaded', function() {
//   fetchData();
// });

// async function sendSectionUpdate(sectionTitle, headers, rows) {
//   // pick target section if remapped
//   const cfg = SECTION_CONFIG[sectionTitle];
//   const target = cfg.targetSection || sectionTitle;
//   const payload = {
//     action: 'updateSection',
//     section: target,
//     headers,
//     rows,
//   };
//   try {
//     const resp = await WEB_APP_URL, {
//       method: 'POST',
//       headers: { 'Content-Type': 'application/json' },
//       body: JSON.stringify(payload),
//       keepalive: true,
//     });
//     if (!resp.ok) throw new Error(`HTTP error! status: ${resp.status}`);
//     return await resp.json();
//   } catch (err) {
//     // 後援：避免 CORS/Preflight 問題，送出 opaque 請求
//     try {
//       await fetch(WEB_APP_URL, {
//         method: 'POST',
//         mode: 'no-cors',
//         headers: { 'Content-Type': 'application/json' },
//         body: JSON.stringify(payload),
//         keepalive: true,
//       });
//       return { ok: true, needsRefetch: true };
//     } catch (e2) {
//       throw err;
//     }
//   }
// }

// function fetchData() {
//   const container = document.getElementById('data-container');
//   if (!container) {
//     console.error('找不到 data-container 元素');
//     return;
//   }

//   container.innerHTML = '<p>正在載入記帳資料...</p>';

//   fetch(WEB_APP_URL)
//     .then((response) => {
//       if (!response.ok) {
//         throw new Error(`HTTP error! status: ${response.status}`);
//       }
//       return response.json();
//     })
//     .then((result) => {
//       if (result && result.data) {
//         LAST_DATA = result.data;
//         displayAccountingData(LAST_DATA);
//       } else {
//         throw new Error('資料格式不正確');
//       }
//     })
//     .catch((error) => {
//       console.error('載入資料時發生錯誤:', error);
//       container.innerHTML = `
//         <div style="color: red; padding: 10px; background-color: #ffe6e6; border-radius: 5px;">
//           <h3>載入失敗</h3>
//           <p>無法載入記帳資料，請稍後再試。</p>
//           <p>錯誤訊息: ${error.message}</p>
//         </div>
//       `;
//     });
// }

// function displayAccountingData(data) {
//   const container = document.getElementById('data-container');
//   if (!container) return;

//   container.innerHTML = '';

//   const mainTitle = document.createElement('h1');
//   mainTitle.textContent = '記帳資料總覽';
//   mainTitle.style.textAlign = 'center';
//   mainTitle.style.marginBottom = '30px';
//   mainTitle.style.color = '#2c3e50';
//   container.appendChild(mainTitle);

//   // 顯示當月收入
//   if (data['當月收入'] && data['當月收入'].length > 0) {
//     displaySection(container, '當月收入', data['當月收入'], 'income');
//   }
//   // 顯示當月支出
//   if (data['當月支出預算'] && data['當月支出預算'].length > 0) {
//     displaySection(container, '當月支出預算', data['當月支出預算'], 'expense');
//   }
//   // 顯示隔月預計支出
//   if (data['隔月預計支出'] && data['隔月預計支出'].length > 0) {
//     displaySection(container, '隔月預計支出', data['隔月預計支出'], 'future');
//   }
// }

// function displaySection(container, title, items, type) {
//   const cfg = SECTION_CONFIG[title];

//   const section = document.createElement('div');
//   section.className = 'accounting-section';
//   section.style.marginBottom = '30px';
//   section.style.padding = '20px';
//   section.style.borderRadius = '8px';
//   section.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';

//   if (type === 'income') {
//     section.style.backgroundColor = '#e8f5e8';
//     section.style.borderLeft = '4px solid #27ae60';
//   } else if (type === 'expense') {
//     section.style.backgroundColor = '#ffeaea';
//     section.style.borderLeft = '4px solid #e74c3c';
//   } else {
//     section.style.backgroundColor = '#f0f8ff';
//     section.style.borderLeft = '4px solid #3498db';
//   }

//   const sectionTitle = document.createElement('h2');
//   sectionTitle.textContent = title + ' ▼';
//   sectionTitle.style.marginBottom = '15px';
//   sectionTitle.style.color = '#2c3e50';
//   sectionTitle.style.cursor = 'pointer';
//   sectionTitle.setAttribute('tabindex', '0');
//   sectionTitle.setAttribute('role', 'button');
//   sectionTitle.setAttribute('aria-expanded', 'true');
//   section.appendChild(sectionTitle);

//   const contentDiv = document.createElement('div');
//   contentDiv.style.display = 'block';

//   // 工具列
//   const controlsDiv = document.createElement('div');
//   controlsDiv.style.display = 'flex';
//   controlsDiv.style.gap = '8px';
//   controlsDiv.style.margin = '8px 0 12px 0';
//   controlsDiv.style.flexWrap = 'wrap';

//   const undoBtn = document.createElement('button');
//   undoBtn.textContent = 'Undo';
//   undoBtn.style.padding = '6px 10px';
//   undoBtn.style.border = '1px solid #aaa';
//   undoBtn.style.background = '#f1f1f1';
//   undoBtn.style.borderRadius = '6px';
//   undoBtn.style.cursor = 'pointer';

//   const redoBtn = document.createElement('button');
//   redoBtn.textContent = 'Redo';
//   redoBtn.style.padding = '6px 10px';
//   redoBtn.style.border = '1px solid #3498db';
//   redoBtn.style.background = '#e3f2fd';
//   redoBtn.style.borderRadius = '6px';
//   redoBtn.style.cursor = 'pointer';

//   const autosaveHint = document.createElement('span');
//   autosaveHint.textContent = '';
//   autosaveHint.style.alignSelf = 'center';
//   autosaveHint.style.color = '#666';

//   // 只有可編輯區塊顯示增刪、取消、儲存
//   controlsDiv.appendChild(undoBtn);
//   controlsDiv.appendChild(redoBtn);
//   controlsDiv.appendChild(autosaveHint);
//   contentDiv.appendChild(controlsDiv);


//    // 卡片容器
//    const cardContainer = document.createElement('div');
//    cardContainer.style.width = '100%';
//    cardContainer.style.overflowX = 'auto';
//    cardContainer.style.border = '1px solid #ddd';
//    cardContainer.style.borderRadius = '4px';

// //   // const thead = document.createElement('thead');
// //   // const headerRow = document.createElement('tr');
// //   const headers = (items.length > 0)
// //   ? Object.keys(items[0])
// //   : (SECTION_HEADERS[title]);

// //   // headers.forEach(header => {
// //   //   const th = document.createElement('th');
// //   //   th.textContent = header;
// //   //   th.style.padding = '12px 15px';
// //   //   th.style.textAlign = 'left';
// //   //   th.style.borderBottom = '2px solid #ddd';
// //   //   th.style.backgroundColor = '#f8f9fa';
// //   //   th.style.whiteSpace = 'nowrap';
// //   //   th.style.minWidth = '80px';
// //   //   if (header.includes('金額') || header.includes('預算')) {
// //   //     th.style.width = '100px'; th.style.textAlign = 'right';
// //   //   } else if (header.includes('項目')) {
// //   //     th.style.width = '150px';
// //   //   } else if (header.includes('細節') || header.includes('備註')) {
// //   //     th.style.width = '200px'; th.style.whiteSpace = 'normal';
// //   //   } else {
// //   //     th.style.width = '120px';
// //   //   }
// //   //   headerRow.appendChild(th);
// //   // });
// //   // thead.appendChild(headerRow);


// //   // const tbody = document.createElement('tbody');
// //   let contentRows = []
// //   items.forEach((item) => {
// //     const row = [];
// //     headers.forEach(header => {
// //       const cell = {};
// //       cell['textContent'] = item[header] || '';
// //       cell['style'] = {
// //         padding: '10px 15px',
// //         borderBottom: '1px solid #ddd',
// //         verticalAlign: 'top',
// //         fontSize: '14px',
// //         fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
// //         contentEditable: 'true',
// //         outline: '1px dashed rgba(0,0,0,0.2)',
// //         backgroundColor: 'rgba(255,255,0,0.06)',
// //       };
// //       if (header.includes('金額') || header.includes('預算')) {
// //         cell['style']['fontWeight'] = 'bold';
// //         cell['style']['textAlign'] = 'right';
// //         cell['style']['fontFamily'] = 'monospace';
// //         if (type === 'income') {
// //           cell['style']['color'] = '#27ae60';
// //         } else if (type === 'expense') {
// //           cell['style']['color'] = '#e74c3c';
// //         }
// //       } else if (header.includes('細節') || header.includes('備註')) {
// //         cell['style']['whiteSpace'] = 'normal';
// //         cell['style']['maxWidth'] = '200px';
// //         cell['style']['lineHeight'] = '1.4';
// //       }
// //       row.push(cell);
// //     });
// //     contentRows.push(row);
// //   });

// //   // const tfoot = document.createElement('tfoot');
// //   // const totalRow = document.createElement('tr');
// //   // totalRow.style.backgroundColor = '#f8f9fa';
// //   // totalRow.style.fontWeight = 'bold';
// //   // headers.forEach((header, i) => {
// //   //   const td = document.createElement('td');
// //   //   td.style.padding = '10px 15px';
// //   //   td.style.borderTop = '2px solid #aaa';
// //   //   td.style.fontWeight = 'bold';
// //   //   td.style.fontSize = '14px';
// //   //   td.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
// //   //   if (i === 0) {
// //   //     td.textContent = '總計';
// //   //   } else if (header.includes('金額') || header.includes('預算')) {
// //   //     td.style.textAlign = 'right';
// //   //     td.style.fontFamily = 'monospace';
// //   //     td.dataset.totalFor = header;
// //   //   } else {
// //   //     td.textContent = '';
// //   //   }
// //   //   totalRow.appendChild(td);
// //   // });
// //   // tfoot.appendChild(totalRow);

// //   // function recalcTotals() {
// //   //   headers.forEach((header) => {
// //   //     if (header.includes('金額') || header.includes('預算')) {
// //   //       let sum = 0;
// //   //       contentRows.forEach(row => {
// //   //         const idx = headers.indexOf(header);
// //   //         const cell = row[idx];
// //   //         const num = parseFloat((cell?.innerText || '').replace(/[^\d.-]/g, '')) || 0;
// //   //         sum += num;
// //   //       });
// //   //       // const totalCell = totalRow.children[headers.indexOf(header)];
// //   //       // if (totalCell) totalCell.textContent = sum.toLocaleString();
// //   //     }
// //   //   });
// //   // }
// //   // setTimeout(recalcTotals, 0);

// //   // History (Undo/Redo)
// //   let historyStack = [];
// //   let futureStack = [];
// //   let lastSnapshot = getSnapshot();

// //   function getSnapshot() {
// //     const rows = [];
// //     contentRows.forEach(row => {
// //       const obj = {};
// //       const cells = row;
// //       headers.forEach((h, i) => {
// //         obj[h] = (cells[i]?.innerText || '').trim();
// //       });
// //       rows.push(obj);
// //     });
// //     return rows;
// //   }
// //   function applySnapshot(snapshot) {
// //     // Rebuild tbody to match snapshot length
// //     contentRows = [];
// //     snapshot.forEach((rowObj) => {
// //       const row = [];
// //       headers.forEach(h => {
// //         const cell = {};
// //         cell['textContent'] = rowObj[h] || '';
// //         cell['style'] = {
// //           padding: '10px 15px',
// //           borderBottom: '1px solid #ddd',
// //           verticalAlign: 'top',
// //           fontSize: '14px',
// //           fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
// //           contentEditable: 'true',
// //           outline: '1px dashed rgba(0,0,0,0.2)',
// //           backgroundColor: 'rgba(255,255,0,0.06)',
// //         };  
// //         row.push(cell);
// //       });
// //       contentRows.push(row);
// //     });
// //     // recalcTotals();
// //   }

// //   // Input: push previous state to history, clear future, autosave
// //   cardContainer.addEventListener('input', () => {
// //     historyStack.push(lastSnapshot);
// //     futureStack = [];
// //     lastSnapshot = getSnapshot();
// //   });

// //   undoBtn.addEventListener('click', () => {
// //     if (historyStack.length === 0) return;
// //     const current = getSnapshot();
// //     const prev = historyStack.pop();
// //     futureStack.push(current);
// //     applySnapshot(prev);
// //     lastSnapshot = getSnapshot();
// //     autosaveHint.textContent = '自動儲存中...';
// //     debouncedAutosave();
// //   });

// //   redoBtn.addEventListener('click', () => {
// //     if (futureStack.length === 0) return;
// //     const current = getSnapshot();
// //     const next = futureStack.pop();
// //     historyStack.push(current);
// //     applySnapshot(next);
// //     lastSnapshot = getSnapshot();
// //     autosaveHint.textContent = '自動儲存中...';
// //     debouncedAutosave();
// //   });

// //   contentDiv.appendChild(cardContainer);

// //   // 手機端觸控優化 - 在表格容器添加到DOM後執行
// //   if (window.innerWidth < 768 && touchHint) {
// //     try {
// //       // 確保表格容器已經有父節點
// //       if (cardContainer.parentNode) {
// //         cardContainer.parentNode.insertBefore(touchHint, cardContainer);
        
// //         // 檢測觸控滾動
// //         let isScrolling = false;
// //         cardContainer.addEventListener('scroll', () => {
// //           if (!isScrolling) {
// //             isScrolling = true;
// //             touchHint.style.display = 'block';
// //             setTimeout(() => {
// //               touchHint.style.display = 'none';
// //               isScrolling = false;
// //             }, 2000);
// //           }
// //         });
        
// //         // 添加觸控手勢支援
// //         let startX = 0;
// //         let startY = 0;
        
// //         cardContainer.addEventListener('touchstart', (e) => {
// //           startX = e.touches[0].clientX;
// //           startY = e.touches[0].clientY;
// //         });
        
// //         cardContainer.addEventListener('touchmove', (e) => {
// //           if (!startX || !startY) return;
          
// //           const deltaX = e.touches[0].clientX - startX;
// //           const deltaY = e.touches[0].clientY - startY;
          
// //           // 水平滑動優先
// //           if (Math.abs(deltaX) > Math.abs(deltaY)) {
// //             e.preventDefault();
// //           }
// //         });
// //       }
// //     } catch (error) {
// //       console.warn('觸控優化功能載入失敗:', error);
// //     }
// //   }

// //   // 一律使用卡片視圖：隱藏表格容器並建立卡片視圖（與表格資料同步）
// //   cardContainer.style.display = 'none';

// //   let cardsDiv = null;
// //   function renderCardsFromSnapshot() {
// //     if (cardsDiv) cardsDiv.remove();
// //     cardsDiv = document.createElement('div');
// //     cardsDiv.className = 'card-view';

// //     const snapshot = getSnapshot();
// //     snapshot.forEach((rowObj, rowIndex) => {
// //       const card = document.createElement('div');
// //       card.className = 'card';

// //       headers.forEach((h) => {
// //         const row = document.createElement('div');
// //         row.className = 'card-row';

// //         const k = {};
// //         k.className = 'card-label';
// //         k.textContent = h;

// //         const v = {};
// //         v.className = 'card-value';
// //         v.textContent = rowObj[h] || '';

// //         // 允許編輯並同步回表格
// //         v.contentEditable = 'true';
// //         v.addEventListener('input', () => {
// //           const cellIndex = headers.indexOf(h);
// //           const targetRow = contentRows[rowIndex];
// //           if (targetRow && targetRow[cellIndex]) {
// //             // Redo the input event since we no longer have table structure
// //           }
// //         });

// //         if (h.includes('金額') || h.includes('預算')) {
// //           v.classList.add('amount');
// //           if (type === 'income') v.classList.add('income');
// //           if (type === 'expense') v.classList.add('expense');
// //         }

// //         row.push(k);
// //         row.push(v);
// //         card.appendChild(row);
// //       });

// //       cardsDiv.appendChild(card);
// //     });

// //     contentDiv.appendChild(cardsDiv);
// //   }

// //   // 初次渲染卡片
// //   renderCardsFromSnapshot();

// //   section.appendChild(contentDiv);
// //   container.appendChild(section);

// //   // 事件：新增列
// //   addBudgetBtn.addEventListener('click', () => {
// //   const newRow = [];
// //   const headersForSection = SECTION_HEADERS[title] || headers;
// //   headersForSection.forEach(header => {
// //     const cell = {};
// //     cell['textContent'] = '';
// //     cell['style'] = {
// //       padding: '10px 15px',
// //       borderBottom: '1px solid #ddd',
// //       verticalAlign: 'top',
// //       fontSize: '14px',
// //       fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
// //       contentEditable: 'true',
// //       outline: '1px dashed rgba(0,0,0,0.2)',
// //       backgroundColor: 'rgba(255,255,0,0.06)',
// //     };
// //     newRow.push(cell);
// //   });

// //     // push history before modifying DOM snapshot reference
// //     historyStack.push(lastSnapshot);
// //     futureStack = [];
// //     lastSnapshot = getSnapshot();
// //   });

// //   // 事件：刪除列
// //   deleteBudgetBtn.addEventListener('click', () => {
// //     const rows = contentRows;
// //     if (rows.length > 0) {
// //       historyStack.push(lastSnapshot);
// //       futureStack = [];
// //       const lastRow = rows[rows.length - 1];
// //       lastRow.forEach(cell => {
// //         cell.remove();
// //       });
// //       // recalcTotals();
// //       lastSnapshot = getSnapshot();
// //     }
// //   });

// //   // Autosave on idle (1.5s debounce)
// //   const debouncedAutosave = debounce(() => {
// //     autosaveHint.textContent = '自動儲存中...';
// //     // reuse sendSectionUpdate using current snapshot
// //     const rows = getSnapshot();
// //     sendSectionUpdate(title, headers, rows).then(() => {
// //       autosaveHint.textContent = '已自動儲存';
// //       setTimeout(() => (autosaveHint.textContent = ''), 1500);
// //     }).catch(() => {
// //       autosaveHint.textContent = '自動儲存失敗';
// //       setTimeout(() => (autosaveHint.textContent = ''), 2000);
// //     });
// //   }, 1500);
// //   cardContainer.addEventListener('input', debouncedAutosave);

// //   // 收合/展開
// //   sectionTitle.addEventListener('click', function() {
// //     if (contentDiv.style.display === 'none') {
// //       contentDiv.style.display = 'block';
// //       sectionTitle.textContent = title + ' ▼';
// //       sectionTitle.setAttribute('aria-expanded', 'true');
// //     } else {
// //       contentDiv.style.display = 'none';
// //       sectionTitle.textContent = title + ' ▲';
// //       sectionTitle.setAttribute('aria-expanded', 'false');
// //     }
// //   });
// //   sectionTitle.addEventListener('keydown', function(e) {
// //     if (e.key === 'Enter' || e.key === ' ') {
// //       sectionTitle.click();
// //     }
// //   });
// }


