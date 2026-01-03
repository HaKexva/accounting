---
layout: page
title: 支出
permalink: /expense/
---

<link rel="stylesheet" href="{{ '/assets/common.css' | relative_url }}?v={{ site.time | date: '%s' }}">
<link rel="stylesheet" href="{{ '/assets/expense-table.css' | relative_url }}?v={{ site.time | date: '%s' }}">

<div id="user-info"></div>

<script>

// 支出表 / 預算表 GAS Web App URL（最新）
// 支出用「新的」 URL，預算維持原本的 URL
const baseExpense = "https://script.google.com/macros/s/AKfycbxpBh0QVSVTjylhh9cj7JG9d6aJi7L7y6pQPW88EbAsNtcd5ckucLagH8XpSAGa8IZt/exec";
const baseBudget  = "https://script.google.com/macros/s/AKfycbxkOKU5YxZWP1XTCFCF7a62Ar71fUz4Qw7tjF3MvMGkLTt6QzzhGLnDsD7wVI_cgpAR/exec";
// 目前選擇的試算表分頁索引（對應 Apps Script 中的 getSheets()[index]，2 代表第三個分頁）
let currentSheetIndex = 2;
// 從 Show Tab Name 取得的所有月份分頁名稱
// 如果前兩個是無效項目（空白表、下拉選單），則會被過濾掉
let sheetNames = [];
// 記錄原始數據是否包含前兩個無效項目，用於正確計算 sheetIndex
let hasInvalidFirstTwoSheets = false;

// 預先載入的所有月份資料（key: sheetIndex, value: { data: {...}, total: [...] }）
let allMonthsData = {}; // 儲存每個月份的資料和總計

// 純新增模式，不需要記錄索引
let allRecords = []; // 用於歷史紀錄列表
let filteredRecords = []; // 過濾後的記錄
let currentRecordIndex = 0; // 當前記錄索引

// 根據類型過濾記錄（支出頁面只顯示支出）
const filterRecordsByType = (type) => {
  filteredRecords = allRecords.filter(r => r.type === type);
};

// ===== 下拉選單選項（以「下拉選單」sheet=1 為主，以下為預設值／後備值）=====
let EXPENSE_CATEGORY_OPTIONS = [
  { value: '生活花費：食', text: '生活花費：食' },
  { value: '生活花費：衣與外貌', text: '生活花費：衣與外貌' },
  { value: '生活花費：住、居家裝修、衛生用品、次月繳納帳單', text: '生活花費：住、居家裝修、衛生用品、次月繳納帳單' },
  { value: '生活花費：行', text: '生活花費：行' },
  { value: '生活花費：育', text: '生活花費：育' },
  { value: '生活花費：樂', text: '生活花費：樂' },
  { value: '生活花費：健（醫療）', text: '生活花費：健（醫療）' },
  { value: '生活花費：帳單', text: '生活花費：帳單' },
  { value: '儲蓄：退休金、醫療預備金、過年紅包支出', text: '儲蓄：退休金、醫療預備金、過年紅包支出' },
  { value: '家人：過年紅包、紀念日', text: '家人：過年紅包、紀念日' }
];
let PAYMENT_METHOD_OPTIONS = [
  { value: '現金', text: '現金' },
  { value: '信用卡', text: '信用卡' },
  { value: '轉帳', text: '轉帳' },
  { value: '存款或儲值的支出：LINE BANK / 悠遊付 / mos card / 髮果 等', text: '存款或儲值的支出：LINE BANK / 悠遊付 / mos card / 髮果 等' }
];
let CREDIT_CARD_PAYMENT_OPTIONS = [
  { value: '分期付款', text: '分期付款' },
  { value: '一次支付', text: '一次支付' }
];
let MONTH_PAYMENT_OPTIONS = [
  { value: '本月支付', text: '本月支付' },
  { value: '次月支付', text: '次月支付' }
];
let PAYMENT_PLATFORM_OPTIONS = [
  { value: 'LINE BANK', text: 'LINE BANK' },
  { value: '悠遊付', text: '悠遊付' },
  { value: 'mos card', text: 'mos card' },
  { value: '髮果', text: '髮果' }
];

// ===== 支付方式關鍵字檢測 =====
// 檢測是否為信用卡類型的支付方式（包含「信用卡」等關鍵字）
const isCreditCardPayment = (paymentMethod) => {
  if (!paymentMethod) return false;
  const keywords = ['信用卡', '刷卡', 'credit card', 'creditcard'];
  const lowerPayment = paymentMethod.toLowerCase();
  return keywords.some(keyword => lowerPayment.includes(keyword.toLowerCase()));
};

// 檢測是否為存款或儲值類型的支付方式（包含「存款」、「儲值」等關鍵字）
const isStoredValuePayment = (paymentMethod) => {
  if (!paymentMethod) return false;
  const keywords = ['存款', '儲值', '儲值的支出', '預付'];
  const lowerPayment = paymentMethod.toLowerCase();
  return keywords.some(keyword => lowerPayment.includes(keyword.toLowerCase()));
};

// ===== 使用共用快取模組 (SyncStatus) =====
// 使用 SyncStatus 模組的快取功能 (定義在 assets/sync-status.js)
const getFromIDB = (key) => SyncStatus.getFromCache(key);
const setToIDB = (key, value) => SyncStatus.setToCache(key, value);
const getCacheTimestamp = (key) => SyncStatus.getCacheTimestamp(key);

// 背景同步：從 API 載入最新資料
const syncFromAPI = async () => {
  SyncStatus.startSync();
  try {
    // 載入最新的月份列表
    await loadMonthNames();
    setToIDB('sheetNames', sheetNames).catch(() => {});
    setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});

    // 檢查是否需要建立新月份
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const currentMonthStr = `${year}${month}`;
    if (!sheetNames.includes(currentMonthStr)) {
      try {
        await callAPI({ name: "Create Tab" });
        await loadMonthNames();
        setToIDB('sheetNames', sheetNames).catch(() => {});
        setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});
      } catch (e) {
        // 建立失敗，稍後再試
      }
    }

    // Recalculate current month index
    const closestSheetIndex = findClosestMonth();
    currentSheetIndex = closestSheetIndex;

    // 載入當前月份的最新資料
    const currentMonthData = await loadMonthData(currentSheetIndex);
    allMonthsData[currentSheetIndex] = currentMonthData;
    setToIDB(`monthData_${currentSheetIndex}`, currentMonthData).catch(() => {});

    // 更新顯示（使用者可能已經在看頁面）
    processDataFromResponse(currentMonthData.data, true);
    updateTotalDisplay();

    // 載入當前月份預算
    await loadBudgetForMonth(currentSheetIndex);
    updateTotalDisplay();
    setToIDB('budgetTotals', budgetTotals).catch(() => {});

    // 同步完成
    SyncStatus.endSync(true);

    // 背景預載其他月份
    preloadAllMonthsData()
      .then(() => {
        updateTotalDisplay();
        Object.keys(allMonthsData).forEach(idx => {
          setToIDB(`monthData_${idx}`, allMonthsData[idx]).catch(() => {});
        });
      })
      .catch(() => {});

    // Background preload budgets for other months
    const budgetSheetIndices = sheetNames.map((name, idx) => idx + 2);
    const otherBudgetIndices = budgetSheetIndices.filter(sheetIndex => sheetIndex !== currentSheetIndex);
    const budgetPromises = otherBudgetIndices.map(sheetIndex => loadBudgetForMonth(sheetIndex).catch(() => {}));

    Promise.all(budgetPromises).then(() => {
      setToIDB('budgetTotals', budgetTotals).catch(() => {});
    });
  } catch (e) {
    SyncStatus.endSync(false);
  }
};

// 從「下拉選單」sheet=1 載入最新選項（只打一小次 API，速度很快）
async function loadDropdownOptions() {
  try {
    const params = { name: "Show Tab Data", sheet: 1, _t: Date.now() };
    const url = `${baseExpense}?${new URLSearchParams(params)}`;
    const res = await fetch(url, {
      method: "GET",
      redirect: "follow",
      mode: "cors",
      cache: "no-store"
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    const responseData = await res.json();
    // 處理不同的資料格式
    let data = null;

    // 如果是陣列，直接使用
    if (Array.isArray(responseData)) {
      data = responseData;
    }
    // 如果是物件，可能是命名範圍的格式，嘗試找到第一個陣列值
    else if (typeof responseData === 'object' && responseData !== null) {
      // 尋找第一個值是陣列的鍵
      for (const key in responseData) {
        if (Array.isArray(responseData[key]) && responseData[key].length > 0) {
          data = responseData[key];
          break;
        }
      }
    }

    if (!data || !Array.isArray(data) || data.length === 0) {
      return;
    }

    const headerRow = data[0];
    // 找對應欄位
    const colCategory   = findHeaderColumn(headerRow, ['消費類別', '類別']);
    const colPayment    = findHeaderColumn(headerRow, ['支付方式']);
    const colCreditCard = findHeaderColumn(headerRow, ['信用卡支付方式']);
    const colMonthPay   = findHeaderColumn(headerRow, ['本月／次月支付']);
    const colPlatform   = findHeaderColumn(headerRow, ['支付平台', '平台']);

    const readColumn = (col) => {
      const arr = [];
      if (col < 0) return arr;
      const seen = new Set();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        const raw = row[col];
        if (raw === undefined || raw === null) continue;
        const val = raw.toString().trim();
        if (!val || seen.has(val)) continue;
        seen.add(val);
        arr.push({ value: val, text: val });
      }
      return arr;
    };

    if (colCategory >= 0) {
      EXPENSE_CATEGORY_OPTIONS = readColumn(colCategory);
    }
    if (colPayment >= 0) {
      PAYMENT_METHOD_OPTIONS = readColumn(colPayment);
    }
    if (colCreditCard >= 0) {
      CREDIT_CARD_PAYMENT_OPTIONS = readColumn(colCreditCard);
    }
    if (colMonthPay >= 0) {
      MONTH_PAYMENT_OPTIONS = readColumn(colMonthPay);
    }
    if (colPlatform >= 0) {
      PAYMENT_PLATFORM_OPTIONS = readColumn(colPlatform);
    }

  } catch (err) {
  }
}

// 根據新的選項更新指定 select + 自訂顯示（如果存在）
function updateSelectOptions(selectId, options) {
  const select = document.getElementById(selectId);
  if (!select) {
    return;
  }

  if (!options || options.length === 0) {
    return;
  }

  const currentValue = select.value;
  select.innerHTML = '';

  options.forEach(opt => {
    const o = document.createElement('option');
    o.value = opt.value;
    o.textContent = opt.text;
    select.appendChild(o);
  });

  // 恢復之前選擇的值（如果還存在）
  if (currentValue && options.some(o => o.value === currentValue)) {
    select.value = currentValue;
  }

  // 更新自訂下拉選單的顯示文字
  const container = select.parentElement;
  if (container) {
    const display = container.querySelector('.select-display');
    if (display) {
      const textEl = display.querySelector('.select-text');
      if (textEl) {
        const selectedOpt = select.options[select.selectedIndex];
        textEl.textContent = selectedOpt ? selectedOpt.textContent : '';
      }

      // 更新下拉選單的選項列表
      const dropdown = container.querySelector('.select-dropdown');
      if (dropdown) {
        dropdown.innerHTML = '';
        options.forEach(opt => {
          const item = document.createElement('div');
          item.className = 'select-option';
          item.textContent = opt.text;
          item.onclick = () => {
            select.value = opt.value;
            if (textEl) textEl.textContent = opt.text;
            dropdown.style.display = 'none';
            const arrow = container.querySelector('.select-arrow');
            if (arrow) arrow.style.transform = 'rotate(0deg)';
            // 觸發 change 事件
            select.dispatchEvent(new Event('change'));
          };
          dropdown.appendChild(item);
        });
      }
    }
  }
}

function findHeaderColumn(headerRow, keywords) {
  for (let c = 0; c < headerRow.length; c++) {
    const headerText = (headerRow[c] || '').toString().trim();
    if (!headerText) continue;
    if (keywords.some(k => headerText.includes(k))) return c;
  }
  return -1;
}

// 清空表單，準備新增
function clearForm() {
    const itemInput = document.getElementById('item-input');
  const expenseCategorySelect = document.getElementById('expense-category-select');
  const paymentMethodSelect = document.getElementById('payment-method-select');
  const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
  const monthPaymentSelect = document.getElementById('month-payment-select');
  const paymentPlatformSelect = document.getElementById('payment-platform-select');
  const actualCostInput = document.getElementById('actual-cost-input');
  const recordCostInput = document.getElementById('record-cost-input');
    const noteInput = document.getElementById('note-input');

  if (itemInput) itemInput.value = '';
  if (expenseCategorySelect && expenseCategorySelect.options.length > 0) {
    expenseCategorySelect.value = expenseCategorySelect.options[0].value;
    const selectContainer = expenseCategorySelect.parentElement;
        if (selectContainer) {
      const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
        const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
          selectText.textContent = expenseCategorySelect.options[0].textContent;
        }
      }
    }
  }
  // 支付方式不自動重置，保持用戶之前的選擇
  // if (paymentMethodSelect && paymentMethodSelect.options.length > 0) {
  //   paymentMethodSelect.value = paymentMethodSelect.options[0].value;
  //   const selectContainer = paymentMethodSelect.parentElement;
  //   if (selectContainer) {
  //     const selectDisplay = selectContainer.querySelector('.select-display');
  //     if (selectDisplay) {
  //       const selectText = selectDisplay.querySelector('.select-text');
  //       if (selectText) {
  //         selectText.textContent = paymentMethodSelect.options[0].textContent;
  //       }
  //     }
  //   }
  //   paymentMethodSelect.dispatchEvent(new Event('change'));
  // }
  if (creditCardPaymentSelect) creditCardPaymentSelect.value = '';
  if (monthPaymentSelect) monthPaymentSelect.value = '';
  if (paymentPlatformSelect) paymentPlatformSelect.value = '';
  if (actualCostInput) actualCostInput.value = '';
  if (recordCostInput) recordCostInput.value = '';
    if (noteInput) noteInput.value = '';

  // 日期已移除
}

// 取得現在日期並格式化為 YYYY/MM/DD（不包含時間）
function getNowFormattedDateTime() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

// 將各種時間字串格式統一轉為 YYYY/MM/DD（不包含時間）
function formatRecordDateTime(raw) {
  if (!raw) return '';
  const dt = new Date(raw);
  if (Number.isNaN(dt.getTime())) {
    if (typeof raw === 'string' && raw.includes('/')) {
      return raw.split(' ')[0];
    }
    return raw;
  }
  const year = dt.getFullYear();
  const month = String(dt.getMonth() + 1).padStart(2, '0');
  const day = String(dt.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

// 統一的 API 調用函數
async function callAPI(postData) {
  const response = await fetch(baseExpense, {
    method: "POST",
    redirect: "follow",
    mode: "cors",
    keepalive: true,
    body: JSON.stringify(postData)
  });

  const responseText = await response.text();
  // 部分 GAS 可能回傳空字串；視為成功（無額外資料）
  if (!responseText || responseText.trim() === '') {
    return { success: true, data: null, total: null };
  }

  let result;
  try {
    result = JSON.parse(responseText);
  } catch (e) {
    throw new Error('後端響應格式錯誤: ' + responseText.substring(0, 100));
  }

  if (!response.ok || !result.success) {
    throw new Error(result.message || result.error || '操作失敗');
  }

  // 不再清除快取 - 呼叫端會使用 result.data 更新快取
  // 這避免了不必要的重新載入，提升效能

  return result;
}

// ===== 歷史紀錄列表刷新（統一邏輯，避免重複）=====
function refreshHistoryList() {
  const historyModal = document.querySelector('.history-modal');
  if (!historyModal) return;

  const newRecords = loadHistoryListFromCache(currentSheetIndex);
  const listElement = historyModal.querySelector('.history-list');
  if (!listElement) return;

  listElement.innerHTML = '';
  const displayRecords = newRecords.filter(r => !isHeaderRecord(r));
  if (displayRecords.length === 0) {
    const emptyMsg = document.createElement('div');
    emptyMsg.textContent = '尚無歷史紀錄';
    emptyMsg.style.cssText = 'text-align: center; padding: 40px; color: #999;';
    listElement.appendChild(emptyMsg);
  } else {
    displayRecords.forEach(record => {
      listElement.appendChild(createHistoryItem(record, listElement));
    });
  }
}

// 統一的更新暫存區和歷史記錄列表
// 如果提供了 data 和 total，直接使用，否則重新載入
async function updateCacheAndHistory(resultData = null, resultTotal = null) {
  // 如果有回傳的資料，直接使用，否則重新載入
  if (resultData && resultTotal) {
    // 處理回傳的資料並更新暫存區（後端返回的 data 是陣列格式）
    processDataFromResponse(resultData, false, currentSheetIndex);
    allMonthsData[currentSheetIndex] = { data: resultData, total: resultTotal };

    // 保存到 IndexedDB
    setToIDB(`monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]).catch(() => {});
  } else {
    // 沒有回傳資料，重新載入
    try {
      const monthData = await loadMonthData(currentSheetIndex);
      allMonthsData[currentSheetIndex] = monthData;
    } catch (error) {
    }
  }

  // 更新歷史記錄列表
  refreshHistoryList();
}

// 填充表單欄位（用於編輯模式）
function fillForm(row) {
  if (!row) return;

  // 根據 Apps Script 結構：[時間(0), item(1), category(2), spendWay(3), creditCard(4), monthIndex(5), actualCost(6), payment(7), recordCost(8), note(9)]
  setTimeout(() => {
    const itemInput = document.getElementById('item-input');
    const expenseCategorySelect = document.getElementById('expense-category-select');
    const paymentMethodSelect = document.getElementById('payment-method-select');
    const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
    const monthPaymentSelect = document.getElementById('month-payment-select');
    const paymentPlatformSelect = document.getElementById('payment-platform-select');
    const actualCostInput = document.getElementById('actual-cost-input');
    const recordCostInput = document.getElementById('record-cost-input');
    const noteInput = document.getElementById('note-input');

    // 項目
    if (itemInput) {
      itemInput.value = row[1] || '';
    }

    // 類別
    if (expenseCategorySelect) {
      expenseCategorySelect.value = row[2] || '';
      // 同步更新自訂下拉顯示文字
      const selectContainer = expenseCategorySelect.parentElement;
      if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
        if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
          if (selectText) {
            const selectedOption = expenseCategorySelect.options[expenseCategorySelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[2] || '';
          }
        }
      }
    }

    // 支付方式
    if (paymentMethodSelect) {
      paymentMethodSelect.value = row[3] || '';
      // 同步更新自訂下拉顯示文字
      const selectContainer = paymentMethodSelect.parentElement;
      if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
        if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
          if (selectText) {
            const selectedOption = paymentMethodSelect.options[paymentMethodSelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[3] || '';
          }
        }
      }
      // 觸發支付方式變更，以顯示/隱藏相關欄位
      paymentMethodSelect.dispatchEvent(new Event('change'));
    }

    // 信用卡支付方式（如果支付方式是信用卡類型）
    if (creditCardPaymentSelect && isCreditCardPayment(row[3])) {
      creditCardPaymentSelect.value = row[4] || '';
      // 同步更新自訂下拉顯示文字
      const selectContainer = creditCardPaymentSelect.parentElement;
        if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
            const selectedOption = creditCardPaymentSelect.options[creditCardPaymentSelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[4] || '';
          }
        }
      }
    }

    // 本月/次月支付（如果支付方式是信用卡類型）
    if (monthPaymentSelect && isCreditCardPayment(row[3])) {
      monthPaymentSelect.value = row[5] || '';
      // 同步更新自訂下拉顯示文字
      const selectContainer = monthPaymentSelect.parentElement;
        if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
            const selectedOption = monthPaymentSelect.options[monthPaymentSelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[5] || '';
          }
        }
      }
    }

    // 支付平台（如果支付方式是存款或儲值類型）
    if (paymentPlatformSelect && isStoredValuePayment(row[3])) {
      paymentPlatformSelect.value = row[7] || '';
      // 同步更新自訂下拉顯示文字
      const selectContainer = paymentPlatformSelect.parentElement;
      if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('.select-display');
        if (selectDisplay) {
          const selectText = selectDisplay.querySelector('.select-text');
          if (selectText) {
            const selectedOption = paymentPlatformSelect.options[paymentPlatformSelect.selectedIndex];
            selectText.textContent = selectedOption ? selectedOption.textContent : row[7] || '';
          }
        }
      }
    }

    // 實際消費金額
    if (actualCostInput) {
      actualCostInput.value = row[6] || '';
    }

    // 列帳消費金額
    if (recordCostInput) {
      recordCostInput.value = row[8] || '';
    }

    // 備註
    if (noteInput) {
      noteInput.value = row[9] || '';
    }

  }, 150);
}

// 判斷一列是否為標題列（例如：時間 / 日期, 項目 等），用於歷史紀錄顯示時略過
function isHeaderRecord(record) {
  if (!record || !record.row) return false;
  const row = record.row;
  const c0 = (row[0] || '').toString().trim();
  const c1 = (row[1] || '').toString().trim();
  if (!c0 && !c1) return false;
  const headerWords0 = ['時間', '日期'];
  const headerWords1 = ['項目', '品項', '標題'];
  const isHeader0 = headerWords0.some(w => c0.includes(w));
  const isHeader1 = headerWords1.some(w => c1.includes(w));
  return isHeader0 && isHeader1;
}

// ===== 預算快取與預先彙總（從預算表載入） =====

// 每個月份對應的「預算支出彙總」，key: sheetIndex, value: { [category: string]: number }
const budgetTotals = {};

// 載入預算表中指定月份的資料，並先把每個類別的預算加總好存起來
async function loadBudgetForMonth(sheetIndex) {
  // Return from cache if available
  if (budgetTotals[sheetIndex]) {
    return budgetTotals[sheetIndex];
  }

  const params = { name: "Show Tab Data", sheet: sheetIndex, _t: Date.now() };
  const url = `${baseBudget}?${new URLSearchParams(params)}`;
  let res;
  try {
    res = await fetch(url, {
      method: "GET",
      redirect: "follow",
      mode: "cors",
      cache: "no-store"
    });
  } catch (err) {
    throw new Error(`無法連接預算表伺服器: ${err.message}`);
  }

  if (!res.ok) {
    throw new Error(`載入預算資料失敗: HTTP ${res.status} ${res.statusText}`);
  }

  let data;
  try {
    data = await res.json();
  } catch (jsonErr) {
    const text = await res.text();
    throw new Error('預算表回應格式錯誤');
  }

  const categoryTotals = {};
  let processedRowsCount = 0;
  let skippedRowsCount = 0;
  let skippedRowsReasons = {
    notExpenseBudget: 0,
    emptyRow: 0,
    headerOrTotal: 0,
    noCategory: 0,
    invalidCost: 0
  };

  if (data && typeof data === 'object') {
    Object.keys(data).forEach(key => {
      const rows = data[key] || [];

      // 只處理包含「支出」字樣的命名範圍（例如：當月支出預算202512）
      const isExpenseBudget = key.includes('支出');
      if (!isExpenseBudget) {
        skippedRowsReasons.notExpenseBudget++;
        return;
      }

      rows.forEach((row, rowIndex) => {
        if (!row || row.length === 0) {
          skippedRowsReasons.emptyRow++;
          return;
        }

        const firstCell = row[0];
        const firstCellStr = String(firstCell || '').trim();

        // 跳過標題列與總計列
        // 預算格式：[編號, 時間, category, item, cost, note]
        // 標題行的 firstCell 可能是 "編號" 或其他標題文字
        // 總計行的 firstCell 可能是 "總計" 或其他總計標記
        // 數據行的 firstCell 通常是數字（編號）或空字串
        if (firstCellStr === '編號' || firstCellStr === '總計' ||
            firstCellStr.toLowerCase() === '編號' || firstCellStr.toLowerCase() === '總計' ||
            firstCellStr.includes('編號') || firstCellStr.includes('總計')) {
          skippedRowsReasons.headerOrTotal++;
          return;
        }

        // 如果 firstCell 是數字，表示這是有效的資料行（編號）
        // 如果 firstCell 是空字串或 null/undefined，也可能是有效的資料行（允許空編號）
        // 繼續處理

        // 真正的資料列：允許 firstCell 為空，只要有類別和金額即可
        // 類別欄位：優先用第 3 欄(row[2])，如果是空的就用第 2 欄(row[1])，以支援「生活花費：食」這種寫法
        // 支出預算格式：[編號, 時間, category, item, cost, note]
        // 重要：如果 category 相同，預算要相加，這樣才能正確顯示該 category 有多少預算
        let category = (row[2] || '').toString().trim();
        if (!category) {
          category = (row[1] || '').toString().trim();
        }
        const item = (row[3] || '').toString().trim(); // item 在第 3 欄（索引 3），僅用於日誌

        const costRaw = row[4];
        const cost = parseFloat(costRaw);

        if (!category) {
          skippedRowsReasons.noCategory++;
          return;
        }
        if (!Number.isFinite(cost)) {
          skippedRowsReasons.invalidCost++;
          return;
        }

        // 重要：如果 category 相同，預算要相加，而不是取第一筆
        // 使用 category 作為 key，確保相同 category 的預算會相加
        // 這樣在支出表中才能正確顯示該 category 有多少預算
        const budgetKey = category; // 只用 category 作為 key
        const oldTotal = categoryTotals[budgetKey] || 0;
        categoryTotals[budgetKey] = oldTotal + cost;
        processedRowsCount++;
      });
    });
  }

  budgetTotals[sheetIndex] = categoryTotals;

  return categoryTotals;
}

// 處理從 Apps Script 回傳的資料（用於更新 allRecords）
const processDataFromResponse = (data, shouldFilter = true, sheetIndexForContext = null) => {
  // 先清空目前的記錄
  allRecords = [];

  // 如果數據是陣列格式，需要先轉換
  if (Array.isArray(data)) {
    const convertedData = {};
    let expenseRows = [];

    data.forEach((row, rowIndex) => {
      if (!row || row.length === 0) return;

      // 跳過標題行（第一行通常是標題）
      if (rowIndex === 0) {
        const firstCell = String(row[0] || '').trim().toLowerCase();
        if (firstCell === '交易日期' || firstCell === '時間' || firstCell === '日期' ||
            firstCell.includes('項目') || firstCell.includes('金額')) {
          return; // 跳過標題行
        }
      }

      // 跳過總計行
      const firstCell = String(row[0] || '').trim();
      if (firstCell === '總計' || firstCell === 'Total' || firstCell === '') {
        return;
      }

      // 根據 Apps Script 結構：[時間(0), item(1), category(2), spendWay(3), creditCard(4), monthIndex(5), actualCost(6), payment(7), recordCost(8), note(9)]
      // 支出頁面處理所有符合結構的行（10欄）
      if (row.length >= 10) {
        // 檢查第一欄是否為時間格式或有效資料
        const timeValue = row[0];
        // 如果第一欄是時間格式（包含 / 或 -），或者是有效的日期字串，或者是非空值，則認為是有效記錄
        // 排除明顯的標題行（如 "交易日期"）
        const timeStr = String(timeValue || '').trim();
        const isHeader = timeStr.toLowerCase() === '交易日期' ||
                         timeStr.toLowerCase() === '時間' ||
                         timeStr.toLowerCase() === '日期';

        if (!isHeader && timeValue !== null && timeValue !== undefined && timeStr !== '') {
          // 進一步檢查：如果是日期格式或非空字串，則加入
          if (timeStr.includes('/') || timeStr.includes('-') ||
              !isNaN(Date.parse(timeValue)) || timeStr.length > 0) {
          expenseRows.push(row);
          }
        }
      }
    });

    // key 需要包含月份名稱以便後續過濾
    const effectiveSheetIndex = (sheetIndexForContext !== undefined && sheetIndexForContext !== null) ? sheetIndexForContext : currentSheetIndex;
    const monthIdx = effectiveSheetIndex - 2;
    const monthName = (monthIdx >= 0 && monthIdx < sheetNames.length) ? sheetNames[monthIdx] : '';
    if (expenseRows.length > 0) convertedData[`當月支出預算${monthName}`] = expenseRows;

    data = convertedData;
  }

  if (data && typeof data === 'object' && !Array.isArray(data)) {
    // 獲取當前月份名稱（允許覆蓋 sheetIndex，避免顯示錯月）
    const effectiveSheetIndex = (sheetIndexForContext !== undefined && sheetIndexForContext !== null) ? sheetIndexForContext : currentSheetIndex;
    const monthIndex = effectiveSheetIndex - 2;
    const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length && sheetNames.length > 0) ? sheetNames[monthIndex] : '';

    // 使用 Set 來追蹤已處理的記錄，避免重複添加
    const processedRowKeys = new Set();

    Object.keys(data).forEach(key => {
      const rows = data[key] || [];

      // 根據命名範圍名稱判斷類型
      // 命名範圍格式：當月收入202506 或 當月支出預算202506
      const isIncome = key.includes('收入');
      const isExpense = key.includes('支出');

      if (!isIncome && !isExpense) {
        return; // 跳過不是收入或支出的資料
      }

      // 只處理當前月份的資料（命名範圍名稱應該包含當前月份）
      // 修正：更嚴格的月份檢查，避免不同月份資料混合
      let isCurrentMonth = false;
      
      // 如果沒有月份名稱（初始化階段），只接受完全匹配的通用 key
      if (!currentMonthName || currentMonthName === '') {
        // 只接受純通用 key（沒有月份數字的）
        isCurrentMonth = key === '當月支出預算' || key === '當月收入';
      } else {
        // 有月份名稱時，必須精確匹配當前月份
        // 例如 currentMonthName = '202602'，key 必須是 '當月支出預算202602'
        const exactMatchKey = key === `當月支出預算${currentMonthName}` || key === `當月收入${currentMonthName}`;
        const keyContainsMonth = key.includes(currentMonthName);
        isCurrentMonth = exactMatchKey || keyContainsMonth;
      }
      
      if (!isCurrentMonth) {
        return; // 跳過不是當前月份的資料
      }

      const type = isIncome ? '收入' : '支出';

      rows.forEach(row => {
        if (!row || row.length === 0) return;

        // 跳過總計行和空行
        const firstCell = (row[0] || '').toString().trim();
        if (firstCell === '交易日期' || firstCell === '總計' || firstCell === 'Total' || firstCell === '') {
          return;
        }

        // 對於支出類型，第一欄是時間格式，不是數字編號
        // 所以不需要檢查是否為數字，只需要確保不是標題行
        const isHeader = firstCell.toLowerCase() === '交易日期' ||
                         firstCell.toLowerCase() === '時間' ||
                         firstCell.toLowerCase() === '日期' ||
                         firstCell.includes('項目') ||
                         firstCell.includes('金額');
        if (isHeader) {
          return;
        }

        // 使用時間、項目和金額作為唯一標識，避免重複添加
        const rowKey = `${row[0] || ''}_${row[1] || ''}_${row[8] || ''}`;
        if (processedRowKeys.has(rowKey)) {
          return; // 跳過重複記錄
        }
        processedRowKeys.add(rowKey);

        allRecords.push({ type, row });
      });
    });
  }

  // 根據目前選擇的類型過濾記錄（預設顯示支出）
  if (shouldFilter) {
    filterRecordsByType('支出'); // 現在只有支出

    // 支出頁面只有新增模式，不需要顯示歷史記錄
    if (filteredRecords.length >= 1) {
      isNewMode = true;
      currentRecordIndex = 0;
    }
  }
};

// 重新計算上方「預算 / 支出 / 餘額」
// 預算：根據當前選擇的類別，從預算表讀取該類別的預算
// 支出：根據當前選擇的類別，從當前月份的所有歷史記錄加總該類別的支出
const updateTotalDisplay = () => {
  const recordCostInput = document.getElementById('record-cost-input');
  const expenseCategorySelect = document.getElementById('expense-category-select');

  // 獲取當前選擇的類別
  const selectedCategory = expenseCategorySelect ? expenseCategorySelect.value : '';
  // 如果沒有選擇類別，顯示 0
  if (!selectedCategory) {
    incomeAmount.textContent = '0';
    expenseAmount.textContent = '0';
    totalAmount.textContent = '0';
    updateTotalColor(0);
    return;
  }

  // 1. 預算：從預算表的 budgetTotals 讀取當前類別的預算
  let budget = 0;
  const budgetData = budgetTotals[currentSheetIndex];
  if (budgetData && typeof budgetData === 'object') {
    budget = parseFloat(budgetData[selectedCategory] || 0) || 0;
  }

  // 2. 支出：從當前月份的所有歷史記錄加總當前類別的支出
  let historyExpense = 0;
  let records = [];
  const monthData = allMonthsData[currentSheetIndex];
  if (monthData && monthData.data) {
    records = loadHistoryListFromCache(currentSheetIndex);
  }

  if (Array.isArray(records) && records.length > 0) {
    const processedRecordKeys = new Set();

    historyExpense = records.reduce((sum, r) => {
      const row = r.row || [];

      // 檢查類別是否匹配（類別在 index 2）
      const recordCategory = (row[2] || '').toString().trim();
      if (recordCategory !== selectedCategory) {
        return sum; // 跳過不同類別的記錄
      }

      // 使用列帳金額（index 8）
      const raw = row[8] !== undefined && row[8] !== null && row[8] !== '' ? row[8] : 0;
      const num = parseFloat(raw) || 0;

      // 使用時間、項目和金額作為唯一標識，避免重複計算
      const recordKey = `${row[0] || ''}_${row[1] || ''}_${raw}`;
      if (processedRecordKeys.has(recordKey)) {
        return sum;
      }
      processedRecordKeys.add(recordKey);

      return sum + num;
    }, 0);
  }

  // 即時輸入的列帳金額（只有當前選擇的類別才計算）
  let liveInput = 0;
  if (recordCostInput && recordCostInput.value) {
    // 檢查當前輸入的類別是否與選擇的類別一致
    const currentInputCategory = expenseCategorySelect ? expenseCategorySelect.value : '';
    if (currentInputCategory === selectedCategory) {
      liveInput = parseFloat(recordCostInput.value) || 0;
    }
  }

  const expense = historyExpense + liveInput;

  // 3. 餘額：預算 - 支出
  const remain = budget - expense;

  // 格式化顯示（使用千分位符號）
  incomeAmount.textContent = budget.toLocaleString('zh-TW');
  expenseAmount.textContent = expense.toLocaleString('zh-TW');
  totalAmount.textContent = remain.toLocaleString('zh-TW');
  updateTotalColor(remain);

  // 隱藏總計區域載入覆蓋層
  const overlay = document.querySelector('.summary-loading-overlay');
  if (overlay) {
    overlay.remove();
  }
};

// 載入單個月份的資料和總計
const loadMonthData = async (sheetIndex) => {
  // 驗證 sheetIndex 是否有效
  if (!Number.isFinite(sheetIndex) || sheetIndex < 2) {
    throw new Error(`無效的 sheet 索引: ${sheetIndex}`);
  }

  // 從試算表抓出「當月收入 / 支出」等資料 - 添加時間戳避免快取
  const dataParams = { name: "Show Tab Data", sheet: sheetIndex, _t: Date.now() };
  const dataUrl = `${baseExpense}?${new URLSearchParams(dataParams)}`;
  let res;
  try {
    res = await fetch(dataUrl, {
    method: "GET",
    redirect: "follow",
    mode: "cors",
    cache: "no-store" // 強制不使用快取
  });
  } catch (fetchError) {
    throw new Error(`無法連接到伺服器: ${fetchError.message}。請檢查網路連接或 CORS 設定。`);
  }

  if (!res.ok) {
    throw new Error(`載入資料失敗: HTTP ${res.status} ${res.statusText}`);
  }

  let data;
  try {
    data = await res.json();
  } catch (jsonError) {
    const text = await res.text();
    throw new Error(`伺服器回應格式錯誤: ${jsonError.message}`);
  }

  // 如果數據是陣列格式（Apps Script ShowTabData 返回 getValues()），需要轉換為物件格式
  if (Array.isArray(data)) {
    // 將陣列轉換為物件格式，以便與 processDataFromResponse 兼容
    // 假設陣列包含所有數據行，我們需要根據實際情況區分收入和支出
    // 由於支出頁面只處理支出數據，我們將其轉換為物件格式
    const convertedData = {};

    // 支出記錄通常是：編號, 時間, 類別, 項目, 金額, 備註, ...
    // 我們將所有數據都放到 "當月支出" 鍵下
    // 如果有收入數據，可能需要根據欄位數量或其他標識來區分
    let expenseRows = [];
    let incomeRows = [];

    data.forEach((row, index) => {
      if (!row || row.length === 0) return;

      // 跳過標題行或空行
      const firstCell = row[0];
      if (firstCell === '' || firstCell === null || firstCell === undefined) return;

      // 跳過「總計」行
      if (firstCell === '總計' || firstCell === 'Total' || firstCell.toString().trim() === '總計') {
        return;
      }

      // 根據 Apps Script 結構：[時間(0), item(1), category(2), spendWay(3), creditCard(4), monthIndex(5), actualCost(6), payment(7), recordCost(8), note(9)]
      // 支出頁面主要處理支出，所有行都當作支出處理（因為結構相同）
      if (row.length >= 10) {
        // 確保是有效數據行（時間欄位不為空）
        if (firstCell !== '' && firstCell !== null && firstCell !== undefined) {
          expenseRows.push(row);
        }
      }
    });

    // 構建物件格式，key 需要包含月份名稱以便後續過濾
    const monthIndex = sheetIndex - 2;
    const monthName = (monthIndex >= 0 && monthIndex < sheetNames.length) ? sheetNames[monthIndex] : '';
    if (expenseRows.length > 0) {
      // key 包含月份名稱，例如 "當月支出預算202601"
      convertedData[`當月支出預算${monthName}`] = expenseRows;
    }
    if (incomeRows.length > 0) {
      convertedData[`當月收入${monthName}`] = incomeRows;
    }
    data = convertedData; // 將轉換後的數據賦值回去
  } else {
  // 檢查每個 key 的資料長度
  Object.keys(data).forEach(key => {
    const rows = data[key] || [];
  });
  }

  // 根據資料計算總計（本地計算比 API 更快更可靠）
  let calculatedIncome = 0;
  let calculatedExpense = 0;
  let incomeCount = 0;
  let expenseCount = 0;

  // 使用 Set 來追蹤已處理的記錄，避免重複計算
  const processedIncomeRecords = new Set();
  const processedExpenseRecords = new Set();

  if (data && typeof data === 'object') {
    // 獲取當前月份名稱（從 sheetNames 中查找，sheetIndex 是實際的 sheet 索引，需要減 2）
    const monthIndex = sheetIndex - 2;
    const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length && sheetNames.length > 0) ? sheetNames[monthIndex] : '';

    Object.keys(data).forEach(key => {
      const rows = data[key] || [];

      // 根據命名範圍名稱判斷類型
      // 命名範圍格式：當月收入202506 或 當月支出預算202506
      const isIncome = key.includes('收入');
      const isExpense = key.includes('支出');

      if (!isIncome && !isExpense) {
        return; // 跳過不是收入或支出的資料
      }

      // 只處理當前月份的資料（命名範圍名稱應該包含當前月份）
      const isCurrentMonth = currentMonthName && key.includes(currentMonthName);
      if (!isCurrentMonth) {
        return; // 跳過不是當前月份的資料
      }

      rows.forEach((row, rowIndex) => {
        if (!row || row.length === 0) return;

        // 檢查是否為空行（所有欄位都是空）
        const isEmptyRow = row.every(cell => cell === '' || cell === null || cell === undefined);
        if (isEmptyRow) {
          return;
        }

        // 跳過總計行（第一欄是 "總計" / "Total" 或空字串）
        const firstCell = row[0];
        if (firstCell === '交易日期' || firstCell === '總計' || firstCell === 'Total' || firstCell === '' || firstCell === null || firstCell === undefined) {
          return;
        }

        // 檢查是否為有效記錄（第一欄應該是數字編號）
        const num = parseInt(firstCell, 10);
        if (!Number.isFinite(num) || num <= 0) {
          return;
        }

        // 檢查是否已經處理過這筆記錄（使用編號+時間作為唯一標識）
        const recordKey = `${num}_${row[1] || ''}`;
        if (isIncome) {
          if (processedIncomeRecords.has(recordKey)) {
            return;
          }
          processedIncomeRecords.add(recordKey);
        } else if (isExpense) {
          if (processedExpenseRecords.has(recordKey)) {
            return;
          }
          processedExpenseRecords.add(recordKey);
        }

        // 收入：[編號, 時間, item, cost, note] - cost 在索引 3 (D欄)
        // 支出：[編號, 時間, category, item, cost, note] - cost 在索引 4 (K欄)
        const costIndex = isIncome ? 3 : (isExpense ? 4 : -1);
        if (costIndex >= 0 && row[costIndex] !== undefined && row[costIndex] !== null && row[costIndex] !== '') {
          const cost = parseFloat(row[costIndex]);
          if (Number.isFinite(cost) && cost !== 0) { // 允許負數，但不累加0
            if (isIncome) {
              calculatedIncome += cost;
              incomeCount++;
            } else if (isExpense) {
              calculatedExpense += cost;
              expenseCount++;
            }
          }
        }
      });

    });
  }

  const calculatedTotal = calculatedIncome - calculatedExpense;
  const calculatedTotalData = [calculatedIncome, calculatedExpense, calculatedTotal];
  // 使用計算的總計，而不是 API 返回的總計
  return { data, total: calculatedTotalData };
};

// 預先載入所有月份的資料
const preloadAllMonthsData = async () => {
  if (sheetNames.length === 0) return;

  // 計算需要載入的月份總數（排除當前月份，因為已經載入過了）
  const monthsToLoad = sheetNames.filter((name, idx) => {
    const sheetIndex = idx + 2;
    return sheetIndex !== currentSheetIndex;
  });

  const totalMonths = monthsToLoad.length;
  if (totalMonths === 0) return; // 如果沒有需要預載的月份，直接返回

  let loadedCount = 0;
  const baseProgress = 2; // 已經載入了月份列表(1)和當前月份(1)
  const totalProgress = sheetNames.length + 1; // 總進度 = 月份列表(1) + 所有月份

  // 由新到舊排序（假設月份字串如 202512），並行送出所有請求
  // 先放目前選擇的月份，其餘按年月由新到舊
  const allMonths = sheetNames.map((name, idx) => ({ name, sheetIndex: idx + 2 }));
  const current = allMonths.find(m => m.sheetIndex === currentSheetIndex);
  const others = allMonths.filter(m => m.sheetIndex !== currentSheetIndex)
    .sort((a, b) => (parseInt(b.name, 10) || 0) - (parseInt(a.name, 10) || 0));
  const orderedMonths = current ? [current, ...others] : others;

  const tasks = orderedMonths.map(async ({ name, sheetIndex }) => {
    if (sheetIndex === currentSheetIndex) {
      return null; // 跳過當前月份，因為已經載入過了
    }

    // 先檢查是否已經有預先存取的資料
    if (allMonthsData[sheetIndex]) {
      loadedCount++;
      updateProgress(baseProgress + loadedCount, totalProgress, '載入月份（從快取）');
      return { sheetIndex, name, success: true, fromCache: true };
    }

    try {
      const monthData = await loadMonthData(sheetIndex);
      allMonthsData[sheetIndex] = monthData;

      // 保存到 IndexedDB
      setToIDB(`monthData_${sheetIndex}`, monthData).catch(() => {});

      // 更新進度條（baseProgress + 已載入的月份數）
      loadedCount++;
      updateProgress(baseProgress + loadedCount, totalProgress, '載入月份');

      return { sheetIndex, name, success: true, fromCache: false };
    } catch (error) {
      // 即使失敗也更新進度
      loadedCount++;
      updateProgress(baseProgress + loadedCount, totalProgress, '載入月份');

      // 載入失敗，繼續處理其他月份
      return { sheetIndex, name, success: false, error: error.message || error.toString() };
    }
  });

  const results = await Promise.all(tasks);

  // 更新進度條到 100%
  updateProgress(totalProgress, totalProgress, '載入完成');

  Object.keys(allMonthsData).forEach(key => {
    const monthData = allMonthsData[key];
    const total = Array.isArray(monthData.total) ? monthData.total : 'N/A';
  });

  // 延遲一點後隱藏進度條，讓用戶看到「載入完成」提示
  // 注意：hideSpinner 內部已經有 800ms 延遲顯示「已完成」，所以這裡不需要額外延遲
  setTimeout(() => {
    hideSpinner();
  }, 100);
};

// 從記憶體載入當前月份的資料（不發送請求）
const loadContentFromMemory = async () => {
  // 先清空目前的記錄（確保不同月份的資料不會混在一起）
  allRecords = [];
  filteredRecords = [];
  currentRecordIndex = 0;

  // 驗證 currentSheetIndex 是否有效
  if (!Number.isFinite(currentSheetIndex) || currentSheetIndex < 2) {
    currentSheetIndex = 2; // 預設為第三個分頁
  }

  // 先從記憶體讀取資料
  let monthData = allMonthsData[currentSheetIndex];

  // 如果記憶體中沒有，嘗試從快取讀取
  if (!monthData) {
    try {
      const storedData = await getFromIDB(`monthData_${currentSheetIndex}`);
      if (storedData) {
        monthData = storedData;
        // 明確標記為從快取載入
        monthData._fromCache = true;
        // 同時載入到記憶體
        allMonthsData[currentSheetIndex] = monthData;
      }
    } catch (e) {
      // 快取可能不可用或數據損壞，忽略錯誤
    }
  }

  if (!monthData) {
    return false; // 表示需要重新載入
  }

  // 確認載入的資料
  const totalPreview = Array.isArray(monthData.total) ? monthData.total : 'N/A';

  // 處理資料（會自動過濾並顯示記錄）
  processDataFromResponse(monthData.data, true);

  // 更新總計顯示（使用預算表快取 + 當前類別支出）
  updateTotalDisplay();

  // 支出頁面只有新增模式
  isNewMode = true;
  currentRecordIndex = 0;

  return true; // 表示成功從記憶體載入
};

// 載入當前月份的資料（優先從記憶體讀取，如果沒有則發送請求）
const loadContent = async (forceReload = false) => {
  // 如果不強制重新載入，先嘗試從記憶體讀取
  if (!forceReload && await loadContentFromMemory()) {
    return; // 成功從記憶體載入，直接返回
  }

  // 如果記憶體中沒有資料，或需要強制重新載入，則發送請求
  try {
    const monthData = await loadMonthData(currentSheetIndex);

    // 更新記憶體中的資料
    allMonthsData[currentSheetIndex] = monthData;

    // 保存到 IndexedDB
    setToIDB(`monthData_${currentSheetIndex}`, monthData).catch(() => {});

    // 處理資料
    processDataFromResponse(monthData.data);

    // 更新總計顯示（使用預算表快取 + 當前類別支出）
    updateTotalDisplay();

    // 支出頁面只有新增模式
    isNewMode = true;
    currentRecordIndex = 0;
  } catch (error) {
    throw error;
  }
};


const loadTotal = async () => {
  // 驗證 currentSheetIndex 是否有效
  if (!Number.isFinite(currentSheetIndex) || currentSheetIndex < 2) {
    currentSheetIndex = 2; // 預設為第三個分頁
  }

  // 優先從預算快取讀取（預算表來源）
  if (budgetTotals[currentSheetIndex]) {
    updateTotalDisplay();
    return;
  }

  // 如果快取中沒有，則向「預算表」發送請求並彙總同類別預算
  try {
    await loadBudgetForMonth(currentSheetIndex);
    updateTotalDisplay();
  } catch (error) {
    // 不拋出錯誤，只記錄，避免影響其他功能
  }
};

// 進度條動畫變數
let progressAnimationTimer = null;
let progressAnimationStartTime = null;
let progressAnimationTarget = 99; // 自動動畫的目標百分比（99%）

// 更新進度條（實際進度，會覆蓋自動動畫）
const updateProgress = (current, total, text = '載入中...') => {
  const progressContainer = document.getElementById('loading-progress');
  if (!progressContainer) return;

  const percentage = total > 0 ? Math.min(99, Math.round((current / total) * 99)) : 0;
  const progressBar = progressContainer.querySelector('.progress-bar');

  if (progressBar) {
    progressBar.style.width = `${percentage}%`;
    progressAnimationTarget = percentage; // 更新目標，但不會超過99%
  }
};

// 啟動進度條自動動畫（10秒內從0%到99%）
const startProgressAnimation = () => {
  // 清除之前的動畫
  if (progressAnimationTimer) {
    clearInterval(progressAnimationTimer);
    progressAnimationTimer = null;
  }

  const progressContainer = document.getElementById('loading-progress');
  if (!progressContainer) return;

  const progressBar = progressContainer.querySelector('.progress-bar');
  if (!progressBar) return;

  progressAnimationStartTime = Date.now();
  const duration = 10000; // 10秒
  const startPercentage = 0;
  const endPercentage = 99;

  progressAnimationTimer = setInterval(() => {
    const elapsed = Date.now() - progressAnimationStartTime;
    const progress = Math.min(elapsed / duration, 1);
    const currentPercentage = startPercentage + (endPercentage - startPercentage) * progress;

    // 確保不超過實際進度目標
    const finalPercentage = Math.min(currentPercentage, progressAnimationTarget);

    progressBar.style.width = `${finalPercentage}%`;

    // 如果已經達到99%或超過，停止動畫
    if (progress >= 1 || finalPercentage >= 99) {
      clearInterval(progressAnimationTimer);
      progressAnimationTimer = null;
    }
  }, 16); // 約60fps
};

const showSpinner = (coverHeader = false) => {
  // 如果已經存在，先移除
  const existingOverlay = document.getElementById('loading-overlay');
  if (existingOverlay) {
    existingOverlay.remove();
  }

  // 檢查是否有歷史紀錄模態框打開（z-index: 10000）
  const historyModal = document.querySelector('.history-modal');
  const isHistoryModalOpen = historyModal && window.getComputedStyle(historyModal).display !== 'none';

  // 決定是否覆蓋頁首
  const shouldCoverHeader = coverHeader || isHistoryModalOpen;
  // z-index: 覆蓋頁首時要高於 header (2000)，模態框打開時要高於模態框 (10000)
  const zIndexValue = isHistoryModalOpen ? 10001 : (coverHeader ? 2001 : 1500);

  // 創建全屏遮罩
  const overlay = document.createElement('div');
  overlay.id = 'loading-overlay';
  overlay.style.cssText = `
    position: fixed;
    top: ${shouldCoverHeader ? '0' : '60px'};
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(255, 255, 255, 0.95);
    z-index: ${zIndexValue};
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    cursor: not-allowed;
  `;

  const progressContainer = document.createElement('div');
  progressContainer.id = 'loading-progress';
  const wrapperDiv = document.createElement('div');
  wrapperDiv.style.cssText = 'width: 300px; text-align: center;';
  const progressText = document.createElement('div');
  progressText.textContent = '載入中...';
  progressText.style.cssText = 'font-size: 16px; margin-bottom: 15px; color: #333;';
  const bgDiv = document.createElement('div');
  bgDiv.style.cssText = 'width: 100%; height: 8px; background-color: #e0e0e0; border-radius: 4px; overflow: hidden;';
  const progressBar = document.createElement('div');
  progressBar.className = 'progress-bar';
  progressBar.style.width = '0%';
  bgDiv.appendChild(progressBar);
  wrapperDiv.appendChild(progressText);
  wrapperDiv.appendChild(bgDiv);
  progressContainer.appendChild(wrapperDiv);

  overlay.appendChild(progressContainer);
  document.body.appendChild(overlay);

  // 啟動自動進度條動畫（10秒內從0%到99%）
  setTimeout(() => {
    startProgressAnimation();
  }, 50);
};

const hideSpinner = () => {
  // 清除自動動畫計時器
  if (progressAnimationTimer) {
    clearInterval(progressAnimationTimer);
    progressAnimationTimer = null;
  }

  const overlay = document.getElementById('loading-overlay');
  if (overlay) {
    const progressContainer = document.getElementById('loading-progress');
    if (progressContainer) {
      const progressBar = progressContainer.querySelector('.progress-bar');
      // 先跳到100%，然後再移除
      if (progressBar) {
        progressBar.style.width = '100%';
      }
    }

    // 稍微延遲後移除遮罩，讓用戶看到100%
    setTimeout(() => {
      overlay.remove();
      const spinner = document.getElementById('loading-spinner');
      if (spinner) {
        spinner.remove();
      }
      const progress = document.getElementById('loading-progress');
      if (progress) {
        progress.remove();
      }
      // 重置目標百分比
      progressAnimationTarget = 99;
    }, 200);
  } else {
    const spinner = document.getElementById('loading-spinner');
    if (spinner) {
      spinner.remove();
    }
    const progress = document.getElementById('loading-progress');
    if (progress) {
      progress.remove();
    }
  }
};

const createInputRow = (labelText, inputId, inputType = 'text') => {
  const row = document.createElement('div');
  row.className = 'input-row';

  const label = document.createElement('label');
  label.textContent = labelText;
  label.htmlFor = inputId; // 關聯到 input

  const input = document.createElement('input');
  input.id = inputId;
  input.name = inputId; // 添加 name 屬性以支持自動填充
  input.type = inputType;

  row.appendChild(label);
  row.appendChild(input);
  return row;
};

const createTextareaRow = (labelText, textareaId, rows = 3) => {
  const row = document.createElement('div');
  row.className = 'input-row';

  const label = document.createElement('label');
  label.textContent = labelText;
  label.htmlFor = textareaId; // 關聯到 textarea

  const textarea = document.createElement('textarea');
  textarea.id = textareaId;
  textarea.name = textareaId; // 添加 name 屬性以支持自動填充
  textarea.rows = rows;

  row.appendChild(label);
  row.appendChild(textarea);
  return row;
};

const createSelectRow = (labelText, selectId, options) => {
  const row = document.createElement('div');
  row.className = 'select-row';

  const label = document.createElement('label');
  label.textContent = labelText;
  label.htmlFor = selectId; // 關聯到 select

  const selectContainer = document.createElement('div');
  selectContainer.className = 'select-container';

  const selectDisplay = document.createElement('div');
  selectDisplay.className = 'select-display';

  // 處理空選項的情況
  const safeOptions = (options && options.length > 0) ? options : [{ value: '', text: '無選項' }];

  const selectText = document.createElement('div');
  selectText.className = 'select-text';
  selectText.textContent = safeOptions[0].text;

  const selectArrow = document.createElement('div');
  selectArrow.className = 'select-arrow';
  selectArrow.textContent = '▼';

  selectDisplay.appendChild(selectText);
  selectDisplay.appendChild(selectArrow);

  const hiddenSelect = document.createElement('select');
  hiddenSelect.id = selectId;
  hiddenSelect.name = selectId; // 添加 name 屬性以支持自動填充
  hiddenSelect.style.display = 'none';
  hiddenSelect.value = safeOptions[0].value;

  const dropdown = document.createElement('div');
  dropdown.className = 'select-dropdown';

  safeOptions.forEach(opt => {
    const option = document.createElement('div');
    option.className = 'select-option';
    option.textContent = opt.text;
    option.dataset.value = opt.value;

    option.addEventListener('click', function() {
      selectText.textContent = opt.text;
      hiddenSelect.value = opt.value;
      dropdown.style.display = 'none';
      selectArrow.style.transform = 'rotate(0deg)';
      hiddenSelect.dispatchEvent(new Event('change'));
    });

    dropdown.appendChild(option);
    const hiddenOption = document.createElement('option');
    hiddenOption.value = opt.value;
    hiddenOption.textContent = opt.text;
    hiddenSelect.appendChild(hiddenOption);
  });

  selectDisplay.addEventListener('click', function(e) {
    e.stopPropagation();
    const isOpen = dropdown.style.display === 'block';

    // 關閉所有其他下拉選單
    document.querySelectorAll('.select-dropdown').forEach(otherDropdown => {
      if (otherDropdown !== dropdown) {
        otherDropdown.style.display = 'none';
        const otherContainer = otherDropdown.closest('.select-container');
        if (otherContainer) {
          const otherArrow = otherContainer.querySelector('.select-arrow');
          if (otherArrow) {
            otherArrow.style.transform = 'rotate(0deg)';
          }
        }
      }
    });

    dropdown.style.display = isOpen ? 'none' : 'block';
    selectArrow.style.transform = isOpen ? 'rotate(0deg)' : 'rotate(180deg)';
  });

  document.addEventListener('click', function(e) {
    if (!selectContainer.contains(e.target)) {
      dropdown.style.display = 'none';
      selectArrow.style.transform = 'rotate(0deg)';
    }
  });

  selectContainer.appendChild(selectDisplay);
  selectContainer.appendChild(dropdown);
  selectContainer.appendChild(hiddenSelect);

  row.appendChild(label);
  row.appendChild(selectContainer);
  return row;
};

// （已移至上方：callAPI 與 updateCacheAndHistory）

// 獲取表單數據
function getFormData(prefix = '') {
  const itemInput = document.getElementById(prefix ? `${prefix}-item-input` : 'item-input');
  const dateInput = document.getElementById(prefix ? `${prefix}-date-input` : 'date-input');
  const expenseCategorySelect = document.getElementById(prefix ? `${prefix}-expense-category-select` : 'expense-category-select');
  const paymentMethodSelect = document.getElementById(prefix ? `${prefix}-payment-method-select` : 'payment-method-select');
  const creditCardPaymentSelect = document.getElementById(prefix ? `${prefix}-credit-card-payment-select` : 'credit-card-payment-select');
  const monthPaymentSelect = document.getElementById(prefix ? `${prefix}-month-payment-select` : 'month-payment-select');
  const paymentPlatformSelect = document.getElementById(prefix ? `${prefix}-payment-platform-select` : 'payment-platform-select');
  const actualCostInput = document.getElementById(prefix ? `${prefix}-actual-cost-input` : 'actual-cost-input');
  const recordCostInput = document.getElementById(prefix ? `${prefix}-record-cost-input` : 'record-cost-input');
  const noteInput = document.getElementById(prefix ? `${prefix}-note-input` : 'note-input');

  if (!itemInput || !expenseCategorySelect || !paymentMethodSelect || !actualCostInput || !recordCostInput) {
    throw new Error('請等待表單載入完成');
  }

  const item = itemInput.value.trim();
  // 獲取日期並轉換為 YYYY/MM/DD 格式
  let date = '';
  if (dateInput && dateInput.value) {
    const dateValue = dateInput.value; // 格式：YYYY-MM-DD
    const [year, month, day] = dateValue.split('-');
    date = `${year}/${month}/${day}`;
  } else {
    // 如果沒有選擇日期，使用今天
    date = getNowFormattedDateTime();
  }

  const category = expenseCategorySelect.value;
  const spendWay = paymentMethodSelect.value;
  const actualCostValue = actualCostInput.value.trim();
  const recordCostValue = recordCostInput.value.trim();
  const note = noteInput ? noteInput.value.trim() : '';

  if (!item) throw new Error('請輸入項目');
  if (!category) throw new Error('請選擇類別');
  if (!spendWay) throw new Error('請選擇支付方式');

  const actualCost = parseFloat(actualCostValue) || 0;
  const recordCost = parseFloat(recordCostValue) || 0;

  if (!actualCostValue || isNaN(actualCost) || actualCost <= 0) {
    throw new Error('請輸入有效實際消費金額');
  }

  if (!recordCostValue || isNaN(recordCost) || recordCost <= 0) {
    throw new Error('請輸入有效列帳消費金額');
  }

  let creditCard = '';
  let monthIndex = '';
  let payment = '';

  if (isCreditCardPayment(spendWay)) {
    creditCard = creditCardPaymentSelect ? creditCardPaymentSelect.value : '';
    monthIndex = monthPaymentSelect ? monthPaymentSelect.value : '';
  } else if (isStoredValuePayment(spendWay)) {
    payment = paymentPlatformSelect ? paymentPlatformSelect.value : '';
  }

  return { date, item, category, spendWay, creditCard, monthIndex, actualCost, payment, recordCost, note };
}

const updateDivVisibility = (forceType = null) => {
  // 如果提供了類型參數，使用它；否則嘗試從 DOM 獲取最新元素的值
  let categoryValue = forceType;
  if (categoryValue === null) {
    // 先嘗試從全局變數獲取
    if (typeof categorySelect !== 'undefined' && categorySelect.value) {
      categoryValue = categorySelect.value;
    } else {
      // 如果全局變數不可用，從 DOM 獲取最新元素
      const categorySelectElement = document.getElementById('category-select');
      if (categorySelectElement) {
        categoryValue = categorySelectElement.value;
      } else {
        categoryValue = '支出'; // 默認值
      }
    }
  }

  div2.innerHTML = '';
  div3.innerHTML = '';
  div4.innerHTML = '';

  // 添加日期輸入字段（在所有類型中都顯示）
  const dateRow = createInputRow('日期：', 'date-input', 'date');
  const dateInput = dateRow.querySelector('#date-input');
  if (dateInput) {
    // 設置默認值為今天
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    dateInput.value = `${year}-${month}-${day}`;
  }

  if (categoryValue === '支出') {
    const categoryRow = createSelectRow('類別：', 'expense-category-select', [
      { value: '生活花費：食', text: '生活花費：食' },
      { value: '生活花費：衣與外貌', text: '生活花費：衣與外貌' },
      { value: '生活花費：住、居家裝修、衛生用品、次月繳納帳單', text: '生活花費：住、居家裝修、衛生用品、次月繳納帳單' },
      { value: '生活花費：行', text: '生活花費：行' },
      { value: '生活花費：育', text: '生活花費：育' },
      { value: '生活花費：樂', text: '生活花費：樂' },
      { value: '生活花費：健（醫療）', text: '生活花費：健（醫療）' },
      { value: '生活花費：帳單', text: '生活花費：帳單' },
      { value: '儲蓄：退休金、醫療預備金、過年紅包支出', text: '儲蓄：退休金、醫療預備金、過年紅包支出' },
      { value: '家人：過年紅包、紀念日', text: '家人：過年紅包、紀念日' }
    ]);
    const costRow = createInputRow('金額：', 'cost-input', 'number');
    const noteRow = createTextareaRow('備註：', 'note-input', 3);
    noteRow.style.marginBottom = '0px';

    div2.appendChild(dateRow);
    div2.appendChild(categoryRow);
    div3.appendChild(costRow);
    div4.appendChild(noteRow);

    itemContainer.style.display = 'flex';
    div2.style.display = 'flex';
    div3.style.display = 'flex';
    div4.style.display = 'flex';
  } else if (categoryValue === '收入') {
    const costRow = createInputRow('金額：', 'cost-input', 'number');
    const noteRow = createTextareaRow('備註：', 'note-input', 3);
    noteRow.style.marginBottom = '0px';

    div2.appendChild(dateRow);
    div2.appendChild(costRow);
    div3.appendChild(noteRow);

    itemContainer.style.display = 'flex';
    div2.style.display = 'flex';
    div3.style.display = 'flex';
    div4.style.display = 'none';
  }
};

const saveData = async () => {
  // 鎖定整個頁面，等待後端回傳
  showSpinner();
  saveButton.textContent = '儲存中...';
  saveButton.disabled = true;
  saveButton.style.opacity = '0.6';
  saveButton.style.cursor = 'not-allowed';

  // 禁用所有輸入和按鈕
  const itemInput = document.getElementById('item-input');
  const expenseCategorySelect = document.getElementById('expense-category-select');
  const paymentMethodSelect = document.getElementById('payment-method-select');
  const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
  const monthPaymentSelect = document.getElementById('month-payment-select');
  const paymentPlatformSelect = document.getElementById('payment-platform-select');
  const actualCostInput = document.getElementById('actual-cost-input');
  const recordCostInput = document.getElementById('record-cost-input');
  const noteInput = document.getElementById('note-input');
  if (itemInput) itemInput.disabled = true;
  if (expenseCategorySelect) expenseCategorySelect.disabled = true;
  if (paymentMethodSelect) paymentMethodSelect.disabled = true;
  if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = true;
  if (monthPaymentSelect) monthPaymentSelect.disabled = true;
  if (paymentPlatformSelect) paymentPlatformSelect.disabled = true;
  if (actualCostInput) actualCostInput.disabled = true;
  if (recordCostInput) recordCostInput.disabled = true;
  if (noteInput) noteInput.disabled = true;
  if (historyButton) historyButton.disabled = true;

  try {
    const formData = getFormData();
    const monthIndex = currentSheetIndex - 2;
    const currentMonthName = (monthIndex >= 0 && monthIndex < sheetNames.length) ? sheetNames[monthIndex] : '';
    const result = await callAPI({
      name: "Upsert Data",
      sheet: currentSheetIndex,
      ...formData
    });

    // 等待後端回傳後才顯示成功訊息
    alert('資料已成功儲存！');

    // 使用回傳資料更新暫存
    if (result && result.data) {
      allMonthsData[currentSheetIndex] = { data: result.data, total: result.total };
      processDataFromResponse(result.data, false, currentSheetIndex);
      // 同步刷新歷史紀錄列表
      refreshHistoryList();

      // 保存到 IndexedDB
      setToIDB(`monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]).catch(() => {});
    }

    // 更新總計顯示（使用最新資料重新計算）
    updateTotalDisplay();

    // 儲存完成後清空表單，準備下一筆新增
    clearForm();
  } catch (error) {
    alert('儲存失敗: ' + error.message);
  } finally {
    // 恢復所有按鈕和輸入
    hideSpinner();
    saveButton.textContent = '儲存';
    saveButton.disabled = false;
    saveButton.style.opacity = '1';
    saveButton.style.cursor = 'pointer';

    if (itemInput) itemInput.disabled = false;
    if (expenseCategorySelect) expenseCategorySelect.disabled = false;
    if (paymentMethodSelect) paymentMethodSelect.disabled = false;
    if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = false;
    if (monthPaymentSelect) monthPaymentSelect.disabled = false;
    if (paymentPlatformSelect) paymentPlatformSelect.disabled = false;
    if (actualCostInput) actualCostInput.disabled = false;
    if (recordCostInput) recordCostInput.disabled = false;
    if (noteInput) noteInput.disabled = false;
    if (historyButton) historyButton.disabled = false;
  }
};


const totalContainer = document.createElement('div');
totalContainer.className = 'total-container';

const budgetCardsContainer = document.createElement('div');
budgetCardsContainer.className = 'budget-cards-container';

// 日期已移除


// 刪除功能已移除（純新增模式不需要）
// 刪除功能改為在歷史紀錄彈出視窗中處理

// 歷史紀錄按鈕（時鐘圖標）
// 歷史紀錄按鈕（文字按鈕）
const historyButton = document.createElement('button');
historyButton.className = 'history-button';
historyButton.textContent = '歷史紀錄';
historyButton.style.cssText = `
  display: inline-block;
  margin-left: 20px;
  padding: 4px 12px;
  border: 1px solid #ddd;
  border-radius: 4px;
  background-color: #fff;
  color: #333;
  font-size: 14px;
  cursor: pointer;
  transition: all 0.2s;
`;
historyButton.onmouseenter = () => {
  historyButton.style.backgroundColor = '#f5f5f5';
};
historyButton.onmouseleave = () => {
  historyButton.style.backgroundColor = '#fff';
};

// 防止重複點擊的標誌
let isHistoryModalOpen = false;

// 月份選擇容器
let monthSelectContainer = null;

// 載入月份列表
async function loadMonthNames() {
  try {
    const params = { name: "Show Tab Name" };
    const url = `${baseExpense}?${new URLSearchParams(params)}&_t=${Date.now()}`;
    const response = await fetch(url, {
      method: "GET",
      redirect: "follow",
      mode: "cors",
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(`載入失敗: HTTP ${response.status}`);
    }

    const data = await response.json();

    if (Array.isArray(data)) {
      // 檢查前兩個項目是否是無效項目（空白表、下拉選單）
      const invalidItems = ['空白表', '下拉選單', '', null, undefined];
      const firstItem = data[0];
      const secondItem = data[1];

      const shouldSkipFirst = invalidItems.includes(firstItem);
      const shouldSkipSecond = invalidItems.includes(secondItem);

      if (shouldSkipFirst && shouldSkipSecond && data.length > 2) {
        // 如果前兩個都是無效項目，跳過它們
        sheetNames = data.slice(2);
        hasInvalidFirstTwoSheets = true;
      } else {
        // 如果前兩個不是無效項目，使用所有數據（都是有效的月份）
        sheetNames = data;
        hasInvalidFirstTwoSheets = false;
      }

      // 保存到 IndexedDB
      setToIDB('sheetNames', sheetNames).catch(() => {});
      setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});

      return sheetNames;
    } else {
      return [];
    }
  } catch (error) {
    return [];
  }
}

// 顯示月份選擇
async function showMonthSelect() {
  // 如果已經打開，不重複打開
  if (isHistoryModalOpen) {
    return;
  }

  isHistoryModalOpen = true;
  historyButton.disabled = true;

  // 如果已經存在月份選擇容器，先移除
  if (monthSelectContainer) {
    monthSelectContainer.remove();
  }

  // 每次打開都重新載入月份列表，確保數據是最新的
  const months = await loadMonthNames();
  // 如果載入失敗或為空，顯示錯誤信息
  if (!months || months.length === 0) {
    alert('無法載入月份列表，請稍後再試');
    isHistoryModalOpen = false;
    historyButton.disabled = false;
    return;
  }

  // 創建彈出視窗
  const modal = document.createElement('div');
  modal.className = 'month-select-modal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: rgba(0, 0, 0, 0.5);
    z-index: 10000;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  `;

  // 創建內容容器
  const content = document.createElement('div');
  content.className = 'month-select-modal-content';
  content.style.cssText = `
    background-color: #fff;
    border-radius: 12px;
    padding: 30px;
    max-width: 400px;
    width: 100%;
    position: relative;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
  `;

  // 關閉按鈕（右上角）
  const closeBtn = document.createElement('button');
  closeBtn.innerHTML = '✕';
  closeBtn.style.cssText = `
    position: absolute;
    top: 10px;
    right: 10px;
    width: 30px;
    height: 30px;
    border: none;
    background: transparent;
    font-size: 24px;
    cursor: pointer;
    color: #666;
    line-height: 1;
  `;
  closeBtn.onclick = () => {
    modal.remove();
    isHistoryModalOpen = false;
    historyButton.disabled = false;
  };

  // 標題
  const title = document.createElement('h2');
  title.textContent = '選擇月份';
  title.style.cssText = `
    margin: 0 0 20px 0;
    font-size: 20px;
    font-weight: 600;
  `;

  // 創建下拉選單容器
  const selectContainer = document.createElement('div');
  selectContainer.style.cssText = `
    margin-bottom: 20px;
  `;

  // 創建下拉選單
  const select = document.createElement('select');
  select.id = 'month-select';
  select.name = 'month-select';
  select.style.cssText = `
    width: 100%;
    padding: 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    background-color: #fff;
    color: #333;
    font-size: 16px;
    cursor: pointer;
  `;

  // 添加選項
  if (months.length === 0) {
    const option = document.createElement('option');
    option.value = '';
    option.textContent = '載入中...';
    option.disabled = true;
    select.appendChild(option);
  } else {
    months.forEach((month, index) => {
      const option = document.createElement('option');
      // 修正：直接使用陣列索引計算 sheetIndex
      // sheetNames 已跳過前兩個無效項目（空白表、下拉選單）
      // 所以 sheetIndex = index + 2
      const sheetIndex = index + 2;

      option.value = sheetIndex;
      option.textContent = month;
      option.dataset.monthName = month; // 儲存月份名稱以便調試
      if (sheetIndex === currentSheetIndex) {
        option.selected = true;
      }
      select.appendChild(option);
    });
  }

  // 選擇月份後載入該月份的歷史紀錄
  select.addEventListener('change', async (e) => {
    const selectedSheetIndex = parseInt(e.target.value, 10);
    if (isNaN(selectedSheetIndex)) {
      return;
    }

    const selectedOption = e.target.options[e.target.selectedIndex];
    const selectedMonthName = selectedOption ? selectedOption.dataset.monthName || selectedOption.textContent : '';
    currentSheetIndex = selectedSheetIndex;

    // 關閉月份選擇彈出視窗
    modal.remove();
    isHistoryModalOpen = false;

    // 顯示歷史紀錄彈出視窗
    await showHistoryModal();
  });

  selectContainer.appendChild(select);

  content.appendChild(closeBtn);
  content.appendChild(title);
  content.appendChild(selectContainer);
  modal.appendChild(content);

  // 點擊背景關閉
  modal.onclick = (e) => {
    if (e.target === modal) {
      modal.remove();
      isHistoryModalOpen = false;
      historyButton.disabled = false;
    }
  };

  document.body.appendChild(modal);
}

historyButton.onclick = showHistoryModal;

// 創建歷史記錄項目的輔助函數
function createHistoryItem(record, listElement) {
  const item = document.createElement('div');
  item.className = 'history-item';
  item.style.cssText = `
    padding: 12px;
    border-bottom: 1px solid #eee;
    cursor: pointer;
    transition: background-color 0.2s;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: relative;
  `;
  item.onmouseenter = () => item.style.backgroundColor = '#f5f5f5';
  item.onmouseleave = () => item.style.backgroundColor = 'transparent';

  const itemContent = document.createElement('div');
  itemContent.style.cssText = 'flex: 1; min-width: 0;';

  const date = document.createElement('div');
  date.textContent = formatRecordDateTime(record.row[0] || '');
  date.style.cssText = 'font-size: 14px; color: #666; margin-bottom: 4px;';

  const itemTitle = document.createElement('div');
  itemTitle.textContent = record.row[1] || '(無標題)';
  itemTitle.style.cssText = 'font-size: 16px; font-weight: 500; color: #333;';

  itemContent.appendChild(date);
  itemContent.appendChild(itemTitle);

  // 刪除按鈕
  const deleteBtn = document.createElement('button');
  deleteBtn.textContent = '刪除';
  deleteBtn.style.cssText = `
    padding: 4px 12px;
    border: 1px solid #dc3545;
    border-radius: 4px;
    background-color: #fff;
    color: #dc3545;
    font-size: 12px;
    cursor: pointer;
    margin-left: 10px;
    transition: all 0.2s;
    flex-shrink: 0;
  `;
  deleteBtn.onmouseenter = () => {
    deleteBtn.style.backgroundColor = '#dc3545';
    deleteBtn.style.color = '#fff';
  };
  deleteBtn.onmouseleave = () => {
    deleteBtn.style.backgroundColor = '#fff';
    deleteBtn.style.color = '#dc3545';
  };
  deleteBtn.onclick = async (e) => {
    e.stopPropagation();
    if (!confirm('確定要刪除這筆記錄嗎？')) return;

    // 鎖定整個頁面，等待後端回傳
    showSpinner();
    deleteBtn.disabled = true;

    // 禁用所有輸入和按鈕
    const itemInput = document.getElementById('item-input');
    const expenseCategorySelect = document.getElementById('expense-category-select');
    const paymentMethodSelect = document.getElementById('payment-method-select');
    const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
    const monthPaymentSelect = document.getElementById('month-payment-select');
    const paymentPlatformSelect = document.getElementById('payment-platform-select');
    const actualCostInput = document.getElementById('actual-cost-input');
    const recordCostInput = document.getElementById('record-cost-input');
    const noteInput = document.getElementById('note-input');
    if (itemInput) itemInput.disabled = true;
    if (expenseCategorySelect) expenseCategorySelect.disabled = true;
    if (paymentMethodSelect) paymentMethodSelect.disabled = true;
    if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = true;
    if (monthPaymentSelect) monthPaymentSelect.disabled = true;
    if (paymentPlatformSelect) paymentPlatformSelect.disabled = true;
    if (actualCostInput) actualCostInput.disabled = true;
    if (recordCostInput) recordCostInput.disabled = true;
    if (noteInput) noteInput.disabled = true;
    if (saveButton) saveButton.disabled = true;
    if (historyButton) historyButton.disabled = true;

    try {
      await deleteRecord(record);
      // 等待後端回傳後才顯示成功訊息
      alert('記錄已成功刪除！');
      // 從畫面上移除這一筆
      if (listElement && item.parentNode === listElement) {
        listElement.removeChild(item);
      } else {
        item.remove();
      }
    } catch (error) {
      alert('刪除失敗: ' + error.message);
    } finally {
      // 恢復所有按鈕和輸入
      hideSpinner();
      deleteBtn.disabled = false;
      if (itemInput) itemInput.disabled = false;
      if (expenseCategorySelect) expenseCategorySelect.disabled = false;
      if (paymentMethodSelect) paymentMethodSelect.disabled = false;
      if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = false;
      if (monthPaymentSelect) monthPaymentSelect.disabled = false;
      if (paymentPlatformSelect) paymentPlatformSelect.disabled = false;
      if (actualCostInput) actualCostInput.disabled = false;
      if (recordCostInput) recordCostInput.disabled = false;
      if (noteInput) noteInput.disabled = false;
      if (saveButton) saveButton.disabled = false;
      if (historyButton) historyButton.disabled = false;
    }
  };

  item.appendChild(itemContent);
  item.appendChild(deleteBtn);

  item.onclick = (e) => {
    if (e.target === deleteBtn || deleteBtn.contains(e.target)) {
      return;
    }
    showEditModal(record);
  };

  return item;
}

// 從暫存區載入歷史紀錄列表（不發送請求，僅從記憶體讀取）
function loadHistoryListFromCache(sheetIndex) {
  // 先清空目前的記錄
  allRecords = [];

  // 從記憶體讀取資料（不使用 IndexedDB，避免 async 複雜性）
  let monthData = allMonthsData[sheetIndex];

  if (!monthData || !monthData.data) {
    return [];
  }

  // 處理資料（不進行過濾，因為只需要記錄列表）
  // 傳入 sheetIndex 確保正確處理月份資料
  processDataFromResponse(monthData.data, false, sheetIndex);

  // 為每一筆記錄標註對應的試算表列號
  // 注意：由於 processDataFromResponse 已經過濾了標題和總計行，
  // 我們使用簡單的索引計算（第 1 列是標題，所以從第 2 列開始）
  // 如果資料是陣列格式，ShowTabData 返回的資料中第一行（索引 0）是標題，資料從第二行（索引 1）開始
  // 但由於我們已經過濾了標題行，這裡的 index 對應的是過濾後的記錄索引
  // 實際的 sheet 行號需要考慮：標題行（第 1 行）+ 過濾掉的總計行
  // 為了簡化，我們假設記錄是按順序的，使用 index + 2（第 1 行是標題，第 2 行開始是資料）
  allRecords.forEach((r, index) => {
    r.sheetRowIndex = index + 2;
  });
  return allRecords;
}

// 找到最接近的月份（當前月份或最新月份）
// 修正：直接從 sheetNames 陣列查找，避免硬編碼參考點造成的錯誤
function findClosestMonth() {
  // 根據現在的年月選擇最接近的月份
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;
  const currentMonthStr = `${currentYear}${String(currentMonth).padStart(2, '0')}`;

  // 如果 sheetNames 有資料，從陣列中查找
  if (sheetNames.length > 0) {
    // 嘗試找到當前月份
    const currentIndex = sheetNames.findIndex(name => name === currentMonthStr);
    if (currentIndex !== -1) {
      // sheetNames 已經跳過了前兩個無效項目（空白表、下拉選單）
      // 所以實際的 sheet 索引是 currentIndex + 2
      return currentIndex + 2;
    }

    // 如果當前月份不存在，找最接近的月份（優先選擇最新的）
    // 將月份字串轉為數字進行比較
    const monthNumbers = sheetNames.map(name => parseInt(name, 10)).filter(n => !isNaN(n));
    const currentMonthNum = parseInt(currentMonthStr, 10);

    // 找到小於等於當前月份的最大值（最接近的過去或當前月份）
    let closestMonth = null;
    let closestIndex = -1;
    for (let i = 0; i < sheetNames.length; i++) {
      const monthNum = parseInt(sheetNames[i], 10);
      if (!isNaN(monthNum) && monthNum <= currentMonthNum) {
        if (closestMonth === null || monthNum > closestMonth) {
          closestMonth = monthNum;
          closestIndex = i;
        }
      }
    }

    // 如果找到了，返回對應的 sheet 索引
    if (closestIndex !== -1) {
      return closestIndex + 2;
    }

    // 如果都沒找到，返回最後一個（最新的）月份
    return sheetNames.length - 1 + 2;
  }

  // 如果 sheetNames 沒有資料，返回預設值
  return 2;
}

// 顯示歷史紀錄彈出視窗（包含月份選擇）
async function showHistoryModal() {
  // 檢查是否已經有現有的歷史記錄彈出視窗
  const existingModal = document.querySelector('.history-modal');
  if (existingModal && isHistoryModalOpen) {
    return;
  }

  // 如果有現有的彈出視窗，先移除
  if (existingModal) {
    existingModal.remove();
  }

  isHistoryModalOpen = true;
  historyButton.disabled = true; // 禁用按鈕，防止重複點擊

  // 創建彈出視窗（先顯示，準備顯示載入中）
  const modal = document.createElement('div');
  modal.className = 'history-modal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: rgba(0, 0, 0, 0.5);
    z-index: 10000;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  `;

  // 創建內容容器
  const content = document.createElement('div');
  content.className = 'history-modal-content';
  content.style.cssText = `
    background-color: #fff;
    border-radius: 12px;
    padding: 0 20px 20px 20px; /* 上緣貼齊，移除頂部內距 */
    max-width: 600px;
    width: 100%;
    max-height: 80vh;
    overflow-y: auto;
    position: relative;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
  `;

  // 設定內容容器最小高度，確保載入中顯示正確
  content.style.minHeight = '300px';

  // 創建載入中顯示（使用 CSS spinner）
  const loadingDiv = document.createElement('div');
  loadingDiv.id = 'history-loading-spinner';
  loadingDiv.style.cssText = `
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 250px;
    width: 100%;
  `;

  const spinnerIcon = document.createElement('div');
  spinnerIcon.style.cssText = `
    width: 48px;
    height: 48px;
    border: 4px solid #e0e0e0;
    border-top-color: #3498db;
    border-radius: 50%;
    animation: spin 1s linear infinite;
    margin-bottom: 20px;
  `;

  const loadingText = document.createElement('div');
  loadingText.textContent = '載入中...';
  loadingText.style.cssText = `
    font-size: 16px;
    color: #666;
  `;

  loadingDiv.appendChild(spinnerIcon);
  loadingDiv.appendChild(loadingText);

  // 找到最接近的月份
  const closestSheetIndex = findClosestMonth();
  currentSheetIndex = closestSheetIndex;

  // 先檢查是否有快取資料（記憶體或 IndexedDB）
  let hasCachedData = false;
  if (!allMonthsData[currentSheetIndex]) {
    try {
      const storedData = await getFromIDB(`monthData_${currentSheetIndex}`);
      if (storedData) {
        allMonthsData[currentSheetIndex] = storedData;
        hasCachedData = true;
      }
    } catch (e) {}
  } else {
    hasCachedData = true;
  }

  // 如果有快取，直接顯示（不顯示 spinner）
  // 如果沒有快取，顯示 spinner 並載入資料
  if (!hasCachedData) {
    content.appendChild(loadingDiv);
  }
  modal.appendChild(content);
  document.body.appendChild(modal);

  // 如果沒有快取，需要載入資料
  if (!hasCachedData) {
    // 如果月份列表還沒載入，先載入
    if (sheetNames.length === 0) {
      await loadMonthNames();
    }

    try {
      const monthData = await loadMonthData(currentSheetIndex);
      allMonthsData[currentSheetIndex] = monthData;

      // 保存到 IndexedDB
      setToIDB(`monthData_${currentSheetIndex}`, monthData).catch(() => {});
    } catch (e) {
    }

    // 移除 spinner
    loadingDiv.remove();
  }

  // 從暫存區載入歷史記錄列表
  const records = loadHistoryListFromCache(currentSheetIndex);

  // 手機版本添加右側 padding
  if (window.innerWidth <= 768) {
    content.style.paddingRight = '20px';
    content.style.paddingLeft = '20px';
  }

  // 關閉按鈕（放在頁首，與標題同一列，一起 sticky）
  const closeBtn = document.createElement('button');
  closeBtn.innerHTML = '✕';
  closeBtn.style.cssText = `
    width: 28px;
    height: 28px;
    border: none;
    background: transparent;
    font-size: 20px;
    cursor: pointer;
    color: #666;
    line-height: 1;
  `;
  closeBtn.onclick = () => {
    modal.remove();
    isHistoryModalOpen = false;
    historyButton.disabled = false; // 重新啟用按鈕
  };

  // 更新記錄列表顯示的函數
  const updateHistoryList = (listElement, recordsToShow) => {
    listElement.innerHTML = '';

    // 顯示所有記錄（過濾標題行）
    const displayRecords = recordsToShow.filter(r => !isHeaderRecord(r));

    if (displayRecords.length === 0) {
      const emptyMsg = document.createElement('div');
      emptyMsg.textContent = '尚無歷史紀錄';
      emptyMsg.style.cssText = 'text-align: center; padding: 40px; color: #999;';
      listElement.appendChild(emptyMsg);
  } else {
      displayRecords.forEach((record, index) => {
        const item = document.createElement('div');
        item.className = 'history-item';
        item.style.cssText = `
          padding: 12px;
          padding-right: ${window.innerWidth <= 768 ? '20px' : '12px'};
          border-bottom: 1px solid #eee;
          cursor: pointer;
          transition: background-color 0.2s;
          display: flex;
          align-items: center;
          justify-content: space-between;
          position: relative;
        `;
        listElement.appendChild(createHistoryItem(record, listElement));
      });
    }
  };

  // 標題和月份選擇容器（固定浮現）
  const headerContainer = document.createElement('div');
  headerContainer.style.cssText = `
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 20px;
    position: sticky;
    top: 0;
    background-color: #fff;
    padding: 10px 0;
    z-index: 10;
    border-bottom: 1px solid #eee;
  `;

  const title = document.createElement('h2');
  title.textContent = '歷史紀錄';
  title.id = 'history-modal-title';
  title.style.cssText = `
    margin: 0;
    font-size: 20px;
    font-weight: 600;
  `;

  // 月份選擇下拉選單
  monthSelectContainer = document.createElement('div');
  monthSelectContainer.style.cssText = `
    display: flex;
    align-items: center;
    gap: 10px;
  `;

  const monthLabel = document.createElement('label');
  monthLabel.textContent = '月份：';
  monthLabel.htmlFor = 'history-month-select'; // 關聯到 select
  monthLabel.style.cssText = 'font-size: 14px; color: #666;';

  const monthSelect = document.createElement('select');
  monthSelect.id = 'history-month-select';
  monthSelect.name = 'history-month-select';
  monthSelect.style.cssText = `
    padding: 6px 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    background-color: #fff;
    color: #333;
    font-size: 14px;
    cursor: pointer;
  `;

  // 添加月份選項
  sheetNames.forEach((month, index) => {
    const option = document.createElement('option');
    const sheetIndex = index + 2;
    option.value = sheetIndex;
    option.textContent = month;
    option.dataset.monthName = month; // 儲存月份名稱以便調試
    if (sheetIndex === currentSheetIndex) {
      option.selected = true;
    }
    monthSelect.appendChild(option);
  });

  // 月份選擇變更事件
  monthSelect.addEventListener('change', async (e) => {
    const selectedSheetIndex = parseInt(e.target.value, 10);
    if (isNaN(selectedSheetIndex)) {
      return;
    }

    const selectedOption = e.target.options[e.target.selectedIndex];
    const selectedMonthName = selectedOption ? selectedOption.dataset.monthName || selectedOption.textContent : '';
    currentSheetIndex = selectedSheetIndex;

    // 如果該月份的資料還沒載入，顯示載入中
    if (!allMonthsData[currentSheetIndex]) {
      // 創建載入中顯示（使用 CSS spinner）
      const loadingDiv = document.createElement('div');
      loadingDiv.style.cssText = `
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 300px;
        width: 100%;
      `;

      const spinnerIcon = document.createElement('div');
      spinnerIcon.style.cssText = `
        width: 48px;
        height: 48px;
        border: 4px solid #e0e0e0;
        border-top-color: #3498db;
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin-bottom: 20px;
      `;

      const loadingText = document.createElement('div');
      loadingText.textContent = '載入中...';
      loadingText.style.cssText = `
        font-size: 16px;
        color: #666;
      `;

      loadingDiv.appendChild(spinnerIcon);
      loadingDiv.appendChild(loadingText);
      list.innerHTML = '';
      list.appendChild(loadingDiv);

      // 載入該月份的資料
      try {
        const monthData = await loadMonthData(currentSheetIndex);
        allMonthsData[currentSheetIndex] = monthData;

        // 保存到 IndexedDB
        setToIDB(`monthData_${currentSheetIndex}`, monthData).catch(() => {});

        // 移除載入中顯示
        loadingDiv.remove();
      } catch (e) {
        // 載入失敗，顯示錯誤訊息
        list.innerHTML = '<div style="text-align: center; padding: 40px; color: #999;">載入失敗，請稍後再試</div>';
        return;
      }
    }

    // 從暫存區重新載入該月份的記錄
    const newRecords = loadHistoryListFromCache(currentSheetIndex);

    // 更新記錄列表顯示
    updateHistoryList(list, newRecords);
  });

  monthSelectContainer.appendChild(monthLabel);
  monthSelectContainer.appendChild(monthSelect);

  // 左側：標題 + 月份選擇；右側：叉叉關閉
  const headerLeftContainer = document.createElement('div');
  headerLeftContainer.style.cssText = `
    display: flex;
    align-items: center;
    gap: 16px;
  `;
  headerLeftContainer.appendChild(title);
  headerLeftContainer.appendChild(monthSelectContainer);

  headerContainer.appendChild(headerLeftContainer);
  headerContainer.appendChild(closeBtn);

  // 記錄列表
  const list = document.createElement('div');
  list.className = 'history-list';

  // 初始化記錄列表顯示
  updateHistoryList(list, records);

  // 只添加一次：headerContainer 已經包含 closeBtn
  content.appendChild(headerContainer);
  content.appendChild(list);
  modal.appendChild(content);

  // 點擊背景關閉
  modal.onclick = (e) => {
    if (e.target === modal) {
      modal.remove();
      isHistoryModalOpen = false;
      historyButton.disabled = false; // 重新啟用按鈕
    }
  };

  document.body.appendChild(modal);
}

// 顯示編輯彈出視窗
function showEditModal(record) {
  // 先從暫存區載入該月份的完整記錄列表
  const allRecordsForMonth = loadHistoryListFromCache(currentSheetIndex);
  // 根據記錄的唯一標識（時間和項目）找到對應的完整記錄
  const recordTime = record.row[0] || '';
  const recordItem = record.row[1] || '';
  const fullRecord = allRecordsForMonth.find(r => {
    const rTime = r.row[0] || '';
    const rItem = r.row[1] || '';
    return rTime === recordTime && rItem === recordItem;
  });

  if (!fullRecord) {
    alert('無法載入完整的記錄數據，請重新選擇');
    return;
  }
  const modal = document.createElement('div');
  modal.className = 'edit-modal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: rgba(0, 0, 0, 0.5);
    z-index: 10001;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  `;

  const content = document.createElement('div');
  content.className = 'edit-modal-content';
  content.style.cssText = `
    background-color: #fff;
    border-radius: 12px;
    padding: 30px;
    max-width: 800px;
    width: 100%;
    max-height: 90vh;
    overflow-y: auto;
    position: relative;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
  `;

  // 關閉按鈕（右上角）
  const closeBtn = document.createElement('button');
  closeBtn.innerHTML = '✕';
  closeBtn.style.cssText = `
    position: absolute;
    top: 10px;
    right: 10px;
    width: 30px;
    height: 30px;
    border: none;
    background: transparent;
    font-size: 24px;
    cursor: pointer;
    color: #666;
    line-height: 1;
    z-index: 12;
  `;
  // 關閉按鈕：只關閉編輯視窗，回到歷史記錄列表
  closeBtn.onclick = () => {
    modal.remove();
    // 不關閉歷史記錄列表，讓用戶可以繼續選擇其他記錄
  };

  // 標題
  const title = document.createElement('h2');
  title.textContent = '編輯記錄';
  title.style.cssText = `
    margin: 0 0 20px 0;
    font-size: 20px;
    font-weight: 600;
  `;

  // 創建表單容器（複製表單結構）
  const formContainer = document.createElement('div');
  formContainer.className = 'edit-form-container';

  // 重新創建表單元素（避免ID衝突）
  const itemRow = createInputRow('項目：', 'edit-item-input');
  const expenseCategoryRow = createSelectRow('類別：', 'edit-expense-category-select', EXPENSE_CATEGORY_OPTIONS);

  const paymentMethodRow = createSelectRow('支付方式：', 'edit-payment-method-select', PAYMENT_METHOD_OPTIONS);

  const creditCardPaymentRow = createSelectRow('信用卡支付方式：', 'edit-credit-card-payment-select', CREDIT_CARD_PAYMENT_OPTIONS);

  const monthPaymentRow = createSelectRow('本月／次月支付：', 'edit-month-payment-select', MONTH_PAYMENT_OPTIONS);

  const paymentPlatformRow = createSelectRow('支付平台：', 'edit-payment-platform-select', PAYMENT_PLATFORM_OPTIONS);

  const actualCostRow = createInputRow('實際消費金額：', 'edit-actual-cost-input', 'number');
  const recordCostRow = createInputRow('列帳消費金額：', 'edit-record-cost-input', 'number');

  // 備註使用 textarea
  const noteRow = document.createElement('div');
  noteRow.className = 'input-row';
  const noteLabel = document.createElement('label');
  noteLabel.textContent = '備註：';
  noteLabel.htmlFor = 'edit-note-input';
  const noteInput = document.createElement('textarea');
  noteInput.id = 'edit-note-input';
  noteInput.name = 'edit-note-input';
  noteInput.rows = 3;
  noteRow.appendChild(noteLabel);
  noteRow.appendChild(noteInput);

  // 設置條件顯示的初始狀態（預設隱藏）
  creditCardPaymentRow.style.display = 'none';
  monthPaymentRow.style.display = 'none';
  paymentPlatformRow.style.display = 'none';

  // 添加支付方式變更事件處理，控制條件欄位的顯示/隱藏
  const paymentMethodSelect = paymentMethodRow.querySelector('#edit-payment-method-select');
  if (paymentMethodSelect) {
    paymentMethodSelect.addEventListener('change', () => {
      const paymentMethod = paymentMethodSelect.value;
      if (isCreditCardPayment(paymentMethod)) {
        creditCardPaymentRow.style.display = 'flex';
        monthPaymentRow.style.display = 'flex';
        paymentPlatformRow.style.display = 'none';
      } else if (isStoredValuePayment(paymentMethod)) {
        creditCardPaymentRow.style.display = 'none';
        monthPaymentRow.style.display = 'none';
        paymentPlatformRow.style.display = 'flex';
  } else {
        creditCardPaymentRow.style.display = 'none';
        monthPaymentRow.style.display = 'none';
        paymentPlatformRow.style.display = 'none';
      }
    });
  }

  formContainer.appendChild(itemRow);
  formContainer.appendChild(expenseCategoryRow);
  formContainer.appendChild(paymentMethodRow);
  formContainer.appendChild(creditCardPaymentRow);
  formContainer.appendChild(monthPaymentRow);
  formContainer.appendChild(paymentPlatformRow);
  formContainer.appendChild(actualCostRow);
  formContainer.appendChild(recordCostRow);
  formContainer.appendChild(noteRow);

  // 先將表單添加到 modal，確保元素在 DOM 中
  content.appendChild(closeBtn);
  content.appendChild(title);
  content.appendChild(formContainer);
  modal.appendChild(content);
  document.body.appendChild(modal);

  // 等待 DOM 更新後再填充數據
  setTimeout(() => {
    // 使用完整記錄數據填充表單
    const row = fullRecord.row;

  const itemInput = document.getElementById('edit-item-input');
  const expenseCategorySelect = document.getElementById('edit-expense-category-select');
  // paymentMethodSelect 已經在第 1858 行聲明，不需要重複聲明
  const creditCardPaymentSelect = document.getElementById('edit-credit-card-payment-select');
  const monthPaymentSelect = document.getElementById('edit-month-payment-select');
  const paymentPlatformSelect = document.getElementById('edit-payment-platform-select');
  const actualCostInput = document.getElementById('edit-actual-cost-input');
  const recordCostInput = document.getElementById('edit-record-cost-input');
  const noteInput = document.getElementById('edit-note-input'); // 重新獲取，確保元素已存在

  // 填充項目（確保可以編輯）
  if (itemInput) {
    itemInput.value = row[1] || '';
    itemInput.readOnly = false; // 確保可以編輯
    itemInput.disabled = false; // 確保可以編輯
  } else {
  }

  // 填充類別
  if (expenseCategorySelect) {
    expenseCategorySelect.value = row[2] || '';
    const selectContainer = expenseCategorySelect.parentElement;
    if (selectContainer) {
      const selectDisplay = selectContainer.querySelector('.select-display');
      if (selectDisplay) {
        const selectText = selectDisplay.querySelector('.select-text');
        if (selectText) {
          const selectedOption = expenseCategorySelect.options[expenseCategorySelect.selectedIndex];
          selectText.textContent = selectedOption ? selectedOption.textContent : row[2] || '';
        }
      }
    }
  }

  // 填充支付方式（先設置值，觸發 change 事件以顯示/隱藏條件欄位）
  if (paymentMethodSelect) {
    paymentMethodSelect.value = row[3] || '';
    const selectContainer = paymentMethodSelect.parentElement;
    if (selectContainer) {
      const selectDisplay = selectContainer.querySelector('.select-display');
      if (selectDisplay) {
        const selectText = selectDisplay.querySelector('.select-text');
        if (selectText) {
          const selectedOption = paymentMethodSelect.options[paymentMethodSelect.selectedIndex];
          selectText.textContent = selectedOption ? selectedOption.textContent : row[3] || '';
        }
      }
    }
    // 手動設置條件欄位的顯示狀態
    const paymentMethod = row[3] || '';
    if (isCreditCardPayment(paymentMethod)) {
      creditCardPaymentRow.style.display = 'flex';
      monthPaymentRow.style.display = 'flex';
      paymentPlatformRow.style.display = 'none';
    } else if (isStoredValuePayment(paymentMethod)) {
      creditCardPaymentRow.style.display = 'none';
      monthPaymentRow.style.display = 'none';
      paymentPlatformRow.style.display = 'flex';
    } else {
      creditCardPaymentRow.style.display = 'none';
      monthPaymentRow.style.display = 'none';
      paymentPlatformRow.style.display = 'none';
    }

    // 等待 DOM 更新後再填充條件欄位
    setTimeout(() => {
      // 填充信用卡支付方式（如果支付方式是信用卡類型）
      if (isCreditCardPayment(paymentMethod) && creditCardPaymentSelect) {
        creditCardPaymentSelect.value = row[4] || '';
        const selectContainer = creditCardPaymentSelect.parentElement;
        if (selectContainer) {
          const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
            const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
              const selectedOption = creditCardPaymentSelect.options[creditCardPaymentSelect.selectedIndex];
              selectText.textContent = selectedOption ? selectedOption.textContent : row[4] || '';
            }
          }
        }
      }

      // 填充本月/次月支付（如果支付方式是信用卡類型）
      if (isCreditCardPayment(paymentMethod) && monthPaymentSelect) {
        monthPaymentSelect.value = row[5] || '';
        const selectContainer = monthPaymentSelect.parentElement;
        if (selectContainer) {
          const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
            const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
              const selectedOption = monthPaymentSelect.options[monthPaymentSelect.selectedIndex];
              selectText.textContent = selectedOption ? selectedOption.textContent : row[5] || '';
            }
          }
        }
      }

      // 填充支付平台（如果支付方式是存款或儲值類型）
      if (isStoredValuePayment(paymentMethod) && paymentPlatformSelect) {
        paymentPlatformSelect.value = row[7] || '';
        const selectContainer = paymentPlatformSelect.parentElement;
        if (selectContainer) {
          const selectDisplay = selectContainer.querySelector('.select-display');
          if (selectDisplay) {
            const selectText = selectDisplay.querySelector('.select-text');
            if (selectText) {
              const selectedOption = paymentPlatformSelect.options[paymentPlatformSelect.selectedIndex];
              selectText.textContent = selectedOption ? selectedOption.textContent : row[7] || '';
            }
          }
        }
      }
    }, 50);
  }

  // 填充金額和備註（在 setTimeout 內部，確保元素已存在）
  if (actualCostInput) {
    actualCostInput.value = row[6] || '';
  } else {
  }
  if (recordCostInput) {
    recordCostInput.value = row[8] || '';
  } else {
  }
  if (noteInput) {
    noteInput.value = row[9] || '';
  } else {
  }
  }, 100); // 等待 DOM 完全渲染後再填充

  // 儲存按鈕
  const saveBtn = document.createElement('button');
  saveBtn.textContent = '儲存';
  saveBtn.className = 'save-button';
  saveBtn.style.cssText = `
    margin-top: 20px;
    width: 100%;
  `;

  let isSaving = false;
  saveBtn.onclick = async () => {
    if (isSaving) return;
    isSaving = true;

    // 鎖定整個頁面，等待後端回傳
    showSpinner();
    saveBtn.textContent = '儲存中...';
    saveBtn.disabled = true;

    // 禁用所有輸入和按鈕
    const editItemInput = document.getElementById('edit-item-input');
    const editExpenseCategorySelect = document.getElementById('edit-expense-category-select');
    const editPaymentMethodSelect = document.getElementById('edit-payment-method-select');
    const editCreditCardPaymentSelect = document.getElementById('edit-credit-card-payment-select');
    const editMonthPaymentSelect = document.getElementById('edit-month-payment-select');
    const editPaymentPlatformSelect = document.getElementById('edit-payment-platform-select');
    const editActualCostInput = document.getElementById('edit-actual-cost-input');
    const editRecordCostInput = document.getElementById('edit-record-cost-input');
    const editNoteInput = document.getElementById('edit-note-input');
    if (editItemInput) editItemInput.disabled = true;
    if (editExpenseCategorySelect) editExpenseCategorySelect.disabled = true;
    if (editPaymentMethodSelect) editPaymentMethodSelect.disabled = true;
    if (editCreditCardPaymentSelect) editCreditCardPaymentSelect.disabled = true;
    if (editMonthPaymentSelect) editMonthPaymentSelect.disabled = true;
    if (editPaymentPlatformSelect) editPaymentPlatformSelect.disabled = true;
    if (editActualCostInput) editActualCostInput.disabled = true;
    if (editRecordCostInput) editRecordCostInput.disabled = true;
    if (editNoteInput) editNoteInput.disabled = true;

    // 禁用主頁面的所有輸入和按鈕
    const itemInput = document.getElementById('item-input');
    const expenseCategorySelect = document.getElementById('expense-category-select');
    const paymentMethodSelect = document.getElementById('payment-method-select');
    const creditCardPaymentSelect = document.getElementById('credit-card-payment-select');
    const monthPaymentSelect = document.getElementById('month-payment-select');
    const paymentPlatformSelect = document.getElementById('payment-platform-select');
    const actualCostInput = document.getElementById('actual-cost-input');
    const recordCostInput = document.getElementById('record-cost-input');
    const noteInput = document.getElementById('note-input');
    if (itemInput) itemInput.disabled = true;
    if (expenseCategorySelect) expenseCategorySelect.disabled = true;
    if (paymentMethodSelect) paymentMethodSelect.disabled = true;
    if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = true;
    if (monthPaymentSelect) monthPaymentSelect.disabled = true;
    if (paymentPlatformSelect) paymentPlatformSelect.disabled = true;
    if (actualCostInput) actualCostInput.disabled = true;
    if (recordCostInput) recordCostInput.disabled = true;
    if (noteInput) noteInput.disabled = true;
    if (saveButton) saveButton.disabled = true;
    if (historyButton) historyButton.disabled = true;

    try {
      await saveDataForEdit(fullRecord);
      // 等待後端回傳後才顯示成功訊息
      alert('修改成功！');
      modal.remove();
      // 關閉編輯視窗後，重新載入歷史記錄列表以顯示最新數據
      // 找到歷史記錄列表的 modal
      const historyModal = document.querySelector('.history-modal');
      if (historyModal) {
        // 重新載入當前月份的記錄列表（從暫存區重新載入，因為已經更新）
        const newRecords = loadHistoryListFromCache(currentSheetIndex);
        const listElement = historyModal.querySelector('.history-list');
        if (listElement) {
          // 使用 updateHistoryList 函數更新列表
          listElement.innerHTML = '';
          const displayRecords = newRecords;
          if (displayRecords.length === 0) {
            const emptyMsg = document.createElement('div');
            emptyMsg.textContent = '尚無歷史紀錄';
            emptyMsg.style.cssText = 'text-align: center; padding: 40px; color: #999;';
            listElement.appendChild(emptyMsg);
          } else {
            displayRecords.forEach(record => {
              listElement.appendChild(createHistoryItem(record, listElement));
            });
          }
        }
      }
    } catch (error) {
      alert('儲存失敗: ' + error.message);
    } finally {
      // 恢復所有按鈕和輸入
      hideSpinner();
      isSaving = false;
      saveBtn.textContent = '儲存';
      saveBtn.disabled = false;

      // 恢復編輯 modal 的輸入
      if (editItemInput) editItemInput.disabled = false;
      if (editExpenseCategorySelect) editExpenseCategorySelect.disabled = false;
      if (editPaymentMethodSelect) editPaymentMethodSelect.disabled = false;
      if (editCreditCardPaymentSelect) editCreditCardPaymentSelect.disabled = false;
      if (editMonthPaymentSelect) editMonthPaymentSelect.disabled = false;
      if (editPaymentPlatformSelect) editPaymentPlatformSelect.disabled = false;
      if (editActualCostInput) editActualCostInput.disabled = false;
      if (editRecordCostInput) editRecordCostInput.disabled = false;
      if (editNoteInput) editNoteInput.disabled = false;

      // 恢復主頁面的輸入和按鈕
      if (itemInput) itemInput.disabled = false;
      if (expenseCategorySelect) expenseCategorySelect.disabled = false;
      if (paymentMethodSelect) paymentMethodSelect.disabled = false;
      if (creditCardPaymentSelect) creditCardPaymentSelect.disabled = false;
      if (monthPaymentSelect) monthPaymentSelect.disabled = false;
      if (paymentPlatformSelect) paymentPlatformSelect.disabled = false;
      if (actualCostInput) actualCostInput.disabled = false;
      if (recordCostInput) recordCostInput.disabled = false;
      if (noteInput) noteInput.disabled = false;
      if (saveButton) saveButton.disabled = false;
      if (historyButton) historyButton.disabled = false;
    }
  };

  content.appendChild(closeBtn);
  content.appendChild(title);
  content.appendChild(formContainer);
  content.appendChild(saveBtn);
  modal.appendChild(content);

  // 點擊背景關閉編輯視窗，返回歷史記錄列表
  modal.onclick = (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  };

  document.body.appendChild(modal);
}

// 刪除記錄的函數
async function deleteRecord(record) {
  // 優先使用預先標註好的 sheetRowIndex（在 loadHistoryListFromCache 中設定）
  const row = record.sheetRowIndex;

  if (!row) {
    throw new Error('無法找到要刪除的記錄位置（缺少列號資訊）');
  }

  const result = await callAPI({
    name: "Delete Data",
    sheet: currentSheetIndex,
    // 後端統一使用 updateRow 作為列號參數（新增/修改/刪除共用）
    updateRow: row
  });

  // 更新總計顯示（使用最新資料重新計算）
  updateTotalDisplay();

  // 更新暫存區和歷史記錄（使用回傳的 data）
  if (result.data) {
    processDataFromResponse(result.data, false, currentSheetIndex);
    allMonthsData[currentSheetIndex] = { data: result.data, total: result.total };

    // 保存到 IndexedDB
    setToIDB(`monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]).catch(() => {});

    refreshHistoryList();
  }
}

// 編輯模式的儲存函數
async function saveDataForEdit(record) {
  const formData = getFormData('edit');

  const allRecordsForMonth = loadHistoryListFromCache(currentSheetIndex);
  const recordTime = formatRecordDateTime(record.row[0] || '');
  const recordItem = (record.row[1] || '').trim();
  const recordCategory = (record.row[2] || '').trim();
  const originalRecordCost = (record.row[6] || '').toString().trim();

  const recordIndex = allRecordsForMonth.findIndex(r => {
    const rTime = formatRecordDateTime(r.row[0] || '');
    const rItem = (r.row[1] || '').trim();
    const rCategory = (r.row[2] || '').trim();
    const rCost = (r.row[6] || '').toString().trim();
    return rTime === recordTime && rItem === recordItem && rCategory === recordCategory && rCost === originalRecordCost;
  });

  if (recordIndex === -1) {
    throw new Error('無法找到對應的記錄位置');
  }

  const result = await callAPI({
    name: "Upsert Data",
    sheet: currentSheetIndex,
    ...formData,
    // sheet 第 1 列是標題，第 2 列才是第一筆資料，所以要 +2
    updateRow: recordIndex + 2
  });

  // 更新總計顯示（使用最新資料重新計算）
  updateTotalDisplay();

  // 更新暫存區和歷史記錄（使用回傳的 data）
  if (result.data) {
    processDataFromResponse(result.data, false, currentSheetIndex);
    allMonthsData[currentSheetIndex] = { data: result.data, total: result.total };

    // 保存到 IndexedDB
    setToIDB(`monthData_${currentSheetIndex}`, allMonthsData[currentSheetIndex]).catch(() => {});

    refreshHistoryList();
  }
}

// 切換收入/支出類型
function switchType(targetType) {
    // 確保 categorySelect 元素存在
    const categorySelectElement = document.getElementById('category-select');
    if (!categorySelectElement) {
      return; // 如果元素不存在，直接返回
    }

    const currentType = categorySelectElement.value || '支出';

    // 如果目標類型與當前類型不同，則切換
    if (currentType !== targetType) {
      // 先更新全局變數，這樣 updateDivVisibility 才能讀取到正確的值
      if (typeof categorySelect !== 'undefined') {
        categorySelect.value = targetType;
      }

      // 更新 DOM 元素
      categorySelectElement.value = targetType;

      // 更新顯示文字
      const selectContainer = categorySelectElement.parentElement;
      if (selectContainer) {
        const selectDisplay = selectContainer.querySelector('div');
        if (selectDisplay) {
          const selectText = selectDisplay.querySelector('div');
          if (selectText) {
            selectText.textContent = targetType;
          }
        }
      }

      // 先更新 UI 元素顯示（特別是支出/收入的欄位切換）
      if (typeof updateDivVisibility === 'function') {
        updateDivVisibility();
      }

      // 然後過濾記錄並更新 UI（filterRecordsByType 會自動處理相同編號的查找）
      // 使用 setTimeout 確保 updateDivVisibility 完成後再執行
      setTimeout(() => {
      // updateDivVisibility 會重新創建 expense-category-select 元素，需要重新獲取並更新
      const newCategorySelectElement = document.getElementById('expense-category-select');
        if (newCategorySelectElement) {
          // 更新新創建的元素的值（如果目標類型是支出，設置第一個選項；如果是收入，不需要設置）
          if (targetType === '支出' && newCategorySelectElement.options.length > 0) {
            newCategorySelectElement.value = newCategorySelectElement.options[0].value;
            // 同步更新自訂下拉顯示文字
            const newSelectContainer = newCategorySelectElement.parentElement;
            if (newSelectContainer) {
              const newSelectDisplay = newSelectContainer.querySelector('div');
              if (newSelectDisplay) {
                const newSelectText = newSelectDisplay.querySelector('div');
                if (newSelectText) {
                  newSelectText.textContent = newCategorySelectElement.options[0].textContent;
                }
              }
            }
          }

          // 更新全局變數的 value 屬性（雖然元素已更換，但我們可以通過更新屬性來保持一致性）
          // 實際上，由於 categorySelect 是 const，我們需要確保後續代碼使用 getElementById 獲取最新元素
          // 但為了兼容性，我們也更新全局變數的 value（如果元素還存在的話）
          if (typeof categorySelect !== 'undefined' && categorySelect.parentNode) {
            categorySelect.value = targetType;
          }
        }

        if (typeof filterRecordsByType === 'function') {
          filterRecordsByType(targetType);
        }

        // 更新相關 UI 元素
        if (typeof updateDeleteButton === 'function') {
          updateDeleteButton();
        }
      }, 200);
  }
}

// 鍵盤事件已簡化，移除導航功能

// 觸摸滑動事件已移除（純新增模式不需要）

// 移除類別選擇（支出/收入），改為新的表單欄位結構

const itemContainer = document.createElement('div');
itemContainer.className = 'item-container';

// 0. 日期（所有類型都要填寫）
const dateRow = createInputRow('日期：', 'date-input', 'date');
const dateInput = dateRow.querySelector('#date-input');
if (dateInput) {
  // 預設帶入今天（YYYY-MM-DD），可以手動修改
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  dateInput.value = `${year}-${month}-${day}`;

  // 當日期變更時，切換到對應月份並更新 summary
  dateInput.addEventListener('change', async () => {
    const selectedDate = dateInput.value; // YYYY-MM-DD
    if (!selectedDate) return;

    const [yearStr, monthStr] = selectedDate.split('-');
    const targetMonthStr = `${yearStr}${monthStr}`; // e.g., "202501"

    // 找到對應月份的 sheetIndex
    const targetIndex = sheetNames.findIndex(name => name === targetMonthStr);
    if (targetIndex !== -1) {
      const newSheetIndex = targetIndex + 2; // sheetIndex = index + 2
      if (newSheetIndex !== currentSheetIndex) {
        currentSheetIndex = newSheetIndex;
        // 載入該月份資料並更新 summary
        if (!allMonthsData[currentSheetIndex]) {
          try {
            const monthData = await loadMonthData(currentSheetIndex);
            allMonthsData[currentSheetIndex] = monthData;
            setToIDB(`monthData_${currentSheetIndex}`, monthData).catch(() => {});
          } catch (e) {
          }
        }
        updateTotalDisplay();
        // Update month indicator if function exists
        if (typeof updateMonthIndicator === 'function') {
          updateMonthIndicator();
        }
      }
    }
  });
}

// 1. 項目
const itemRow = createInputRow('項目：', 'item-input', 'text');
const itemInput = itemRow.querySelector('#item-input');
itemInput.placeholder = '輸入項目名稱...';

// 2. 類別（消費類別）- 使用統一常數
const categoryRow = createSelectRow('類別：', 'expense-category-select', EXPENSE_CATEGORY_OPTIONS);
const expenseCategorySelect = categoryRow.querySelector('#expense-category-select');

// 類別變更時即時更新上方「預算 / 支出 / 餘額」
if (expenseCategorySelect) {
  expenseCategorySelect.addEventListener('change', () => {
    updateTotalDisplay();
  });
}

// 3. 支付方式 - 使用統一常數
const paymentMethodRow = createSelectRow('支付方式：', 'payment-method-select', PAYMENT_METHOD_OPTIONS);
const paymentMethodSelect = paymentMethodRow.querySelector('#payment-method-select');

// 4. 信用卡支付方式（條件顯示：支付方式是信用卡）- 使用統一常數
const creditCardPaymentRow = createSelectRow('信用卡支付方式：', 'credit-card-payment-select', CREDIT_CARD_PAYMENT_OPTIONS);
creditCardPaymentRow.style.display = 'none'; // 預設隱藏
const creditCardPaymentSelect = creditCardPaymentRow.querySelector('#credit-card-payment-select');

// 5. 本月／次月支付（條件顯示：支付方式是信用卡）- 使用統一常數
const monthPaymentRow = createSelectRow('本月／次月支付：', 'month-payment-select', MONTH_PAYMENT_OPTIONS);
monthPaymentRow.style.display = 'none'; // 預設隱藏
const monthPaymentSelect = monthPaymentRow.querySelector('#month-payment-select');

// 6. 支付平台（條件顯示：支付方式是存款或儲值的支出）- 使用統一常數
const paymentPlatformRow = createSelectRow('支付平台：', 'payment-platform-select', PAYMENT_PLATFORM_OPTIONS);
paymentPlatformRow.style.display = 'none'; // 預設隱藏
const paymentPlatformSelect = paymentPlatformRow.querySelector('#payment-platform-select');

// 7. 實際消費金額
const actualCostRow = createInputRow('實際消費金額：', 'actual-cost-input', 'number');
const actualCostInput = actualCostRow.querySelector('#actual-cost-input');

// 8. 列帳消費金額
const recordCostRow = createInputRow('列帳消費金額：', 'record-cost-input', 'number');
const recordCostInput = recordCostRow.querySelector('#record-cost-input');

// Debounce helper for input handlers
let updateTotalDebounceTimer = null;
const debouncedUpdateTotal = (immediate = false) => {
  if (updateTotalDebounceTimer) {
    clearTimeout(updateTotalDebounceTimer);
  }
  if (immediate) {
    updateTotalDisplay();
  } else {
    updateTotalDebounceTimer = setTimeout(() => {
      updateTotalDisplay();
    }, 150); // 150ms debounce
  }
};

// 實際金額 / 列帳金額 / 類別變更時即時更新上方「預算 / 支出 / 餘額」
if (actualCostInput) {
  actualCostInput.addEventListener('input', () => {
    debouncedUpdateTotal();
  });
}
if (recordCostInput) {
  recordCostInput.addEventListener('input', () => {
    debouncedUpdateTotal();
  });

  // Sticky summary when record-cost input is focused and summary is not visible
  const placeholder = document.createElement('div');
  placeholder.className = 'total-container-placeholder';
  placeholder.style.cssText = 'display: none; width: 100%;';

  let stickyActive = false;
  let inputFocused = false;

  // Check if summary is visible in viewport
  const isSummaryVisible = () => {
    if (!placeholder.parentNode) return true; // If no placeholder, check original position
    const rect = placeholder.getBoundingClientRect();
    const vv = window.visualViewport;
    const viewportTop = vv ? vv.offsetTop : 0;
    const navHeight = 62;
    // Summary is visible if its top is below the nav bar
    return rect.top >= (viewportTop + navHeight);
  };

  // Update sticky state based on scroll position and visibility
  const updateStickyState = () => {
    if (!inputFocused || !totalContainer) return;

    const shouldBeSticky = !isSummaryVisible();

    if (shouldBeSticky && !stickyActive) {
      // Make sticky
      totalContainer.classList.add('sticky-active');
      stickyActive = true;
      updateStickyPosition();
    } else if (!shouldBeSticky && stickyActive) {
      // Remove sticky
      totalContainer.classList.remove('sticky-active');
      totalContainer.style.top = '';
      stickyActive = false;
    }
  };

  // Update sticky position based on visual viewport (for keyboard)
  const updateStickyPosition = () => {
    if (!stickyActive || !totalContainer) return;

    const vv = window.visualViewport;
    if (vv) {
      // Check if keyboard is likely open (viewport height significantly reduced)
      const keyboardOpen = vv.height < window.innerHeight * 0.75;

      if (keyboardOpen) {
        // When keyboard is open, position at top of visual viewport
        totalContainer.style.top = vv.offsetTop + 'px';
      } else {
        // Normal: position below nav bar
        totalContainer.style.top = '62px';
      }
    }
  };

  // Listen to visual viewport changes (keyboard show/hide)
  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', updateStickyPosition);
    window.visualViewport.addEventListener('scroll', updateStickyPosition);
  }

  // Listen to scroll for checking visibility
  window.addEventListener('scroll', updateStickyState, { passive: true });

  recordCostInput.addEventListener('focus', () => {
    if (totalContainer && totalContainer.parentNode) {
      inputFocused = true;
      // Insert placeholder before totalContainer if not already there
      if (!placeholder.parentNode) {
        totalContainer.parentNode.insertBefore(placeholder, totalContainer);
      }
      placeholder.style.height = totalContainer.offsetHeight + 'px';
      placeholder.style.display = 'block';
      // Check if should be sticky immediately
      updateStickyState();
    }
  });

  recordCostInput.addEventListener('blur', () => {
    inputFocused = false;
    if (totalContainer) {
      totalContainer.classList.remove('sticky-active');
      totalContainer.style.top = '';
      placeholder.style.display = 'none';
      stickyActive = false;
    }
  });
}

// 9. 備註
const noteRow = document.createElement('div');
noteRow.className = 'input-row';
const noteLabel = document.createElement('label');
noteLabel.textContent = '備註：';
noteLabel.htmlFor = 'note-input';
const noteInput = document.createElement('textarea');
noteInput.id = 'note-input';
noteInput.name = 'note-input';
noteInput.rows = 3;
noteRow.appendChild(noteLabel);
noteRow.appendChild(noteInput);

// 條件顯示邏輯：根據支付方式顯示/隱藏相關欄位
const updatePaymentFieldsVisibility = () => {
  const paymentMethod = paymentMethodSelect.value;

  // 如果支付方式名稱含有「信用卡」等關鍵字，顯示信用卡支付方式和本月/次月支付
  if (isCreditCardPayment(paymentMethod)) {
    creditCardPaymentRow.style.display = 'flex';
    monthPaymentRow.style.display = 'flex';
    paymentPlatformRow.style.display = 'none';
  }
  // 如果支付方式名稱含有「存款」、「儲值」等關鍵字，顯示支付平台
  else if (isStoredValuePayment(paymentMethod)) {
    creditCardPaymentRow.style.display = 'none';
    monthPaymentRow.style.display = 'none';
    paymentPlatformRow.style.display = 'flex';
  }
  // 其他情況隱藏所有條件欄位
  else {
    creditCardPaymentRow.style.display = 'none';
    monthPaymentRow.style.display = 'none';
    paymentPlatformRow.style.display = 'none';
  }
};

// 監聽支付方式變更
paymentMethodSelect.addEventListener('change', updatePaymentFieldsVisibility);

// 將所有欄位添加到 itemContainer（日期放在項目名稱下面）
itemContainer.appendChild(itemRow);
itemContainer.appendChild(dateRow);
itemContainer.appendChild(categoryRow);
itemContainer.appendChild(paymentMethodRow);
itemContainer.appendChild(creditCardPaymentRow);
itemContainer.appendChild(monthPaymentRow);
itemContainer.appendChild(paymentPlatformRow);
itemContainer.appendChild(actualCostRow);
itemContainer.appendChild(recordCostRow);
itemContainer.appendChild(noteRow);

const saveButton = document.createElement('button');
saveButton.textContent = '儲存';
saveButton.className = 'save-button';

const columnsContainer = document.createElement('div');
columnsContainer.className = 'columns-container';
columnsContainer.style.position = 'relative';

// 總計區域載入中覆蓋層
const summaryLoadingOverlay = document.createElement('div');
summaryLoadingOverlay.className = 'summary-loading-overlay';
summaryLoadingOverlay.innerHTML = '<span class="summary-spinner"></span>';
columnsContainer.appendChild(summaryLoadingOverlay);

// Month indicator in summary section
const monthIndicator = document.createElement('div');
monthIndicator.className = 'summary-month-indicator';
monthIndicator.style.cssText = 'font-size: 14px; color: #666; margin-bottom: 8px; text-align: center; font-weight: 500;';

// Function to update month indicator based on currentSheetIndex
const updateMonthIndicator = () => {
  if (sheetNames.length > 0 && currentSheetIndex >= 2) {
    const monthName = sheetNames[currentSheetIndex - 2]; // e.g., "202501"
    if (monthName && monthName.length >= 6) {
      const year = monthName.substring(0, 4);
      const month = monthName.substring(4, 6);
      monthIndicator.textContent = `${year}年${parseInt(month)}月`;
    }
  } else {
    // Default to current month
    const now = new Date();
    monthIndicator.textContent = `${now.getFullYear()}年${now.getMonth() + 1}月`;
  }
};
updateMonthIndicator();

const incomeColumn = document.createElement('div');
incomeColumn.className = 'income-column';

const incomeTitle = document.createElement('h3');
incomeTitle.className = 'income-title';
incomeTitle.textContent = '預算：';

const incomeAmount = document.createElement('div');
incomeAmount.className = 'income-amount';
incomeAmount.textContent = '---';

const expenseColumn = document.createElement('div');
expenseColumn.className = 'expense-column';

const expenseTitle = document.createElement('h3');
expenseTitle.className = 'expense-title';
expenseTitle.textContent = '支出：';

const expenseAmount = document.createElement('div');
expenseAmount.className = 'expense-amount';
expenseAmount.textContent = '---';

const totalColumn = document.createElement('div');
totalColumn.className = 'total-column';

const totalTitle = document.createElement('h3');
totalTitle.className = 'total-title';
totalTitle.textContent = '餘額：';

const totalAmount = document.createElement('div');
totalAmount.className = 'total-amount';
totalAmount.textContent = '---';

const updateTotalColor = (value) => {
  const numValue = parseFloat(value) || 0;
  totalAmount.classList.remove('positive', 'negative');
  totalTitle.classList.remove('positive', 'negative');
  if (numValue > 0) {
    totalAmount.classList.add('positive');
    totalTitle.classList.add('positive');
  } else if (numValue < 0) {
    totalAmount.classList.add('negative');
    totalTitle.classList.add('negative');
  }
};


// 移除圓餅圖區域

const submitContainer = document.createElement('div');
submitContainer.style.width = '100%';
submitContainer.style.display = 'flex';
submitContainer.style.justifyContent = 'center';
submitContainer.style.padding = '0';



totalContainer.appendChild(monthIndicator);
totalContainer.appendChild(columnsContainer);
      columnsContainer.appendChild(incomeColumn);
        incomeColumn.appendChild(incomeTitle);
        incomeColumn.appendChild(incomeAmount);
      columnsContainer.appendChild(expenseColumn);
        expenseColumn.appendChild(expenseTitle);
        expenseColumn.appendChild(expenseAmount);
      columnsContainer.appendChild(totalColumn);
        totalColumn.appendChild(totalTitle);
        totalColumn.appendChild(totalAmount);

budgetCardsContainer.appendChild(itemContainer);
budgetCardsContainer.appendChild(submitContainer);
  submitContainer.appendChild(saveButton);

saveButton.addEventListener('click', saveData);
saveButton.addEventListener('click', loadTotal);

document.addEventListener('DOMContentLoaded', async function() {
  // 顯示進度條載入下拉選單（第一個請求）
  showSpinner();

  // 一進到頁面就發送 Create Tab
  try {
    await callAPI({ name: "Create Tab" });
    // 發送 Create Tab 後，重新載入月份列表
    await loadMonthNames();
    setToIDB('sheetNames', sheetNames).catch(() => {});
    setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});
    // 如果月份選擇下拉選單已經打開，更新它
    const existingModal = document.querySelector('.month-select-modal');
    if (existingModal) {
      // 如果月份選擇彈出視窗已經打開，重新載入月份列表並更新選項
      const select = existingModal.querySelector('#month-select');
      if (select) {
        select.innerHTML = '';
        sheetNames.forEach((month, index) => {
          const option = document.createElement('option');
          // 修正：直接使用陣列索引計算 sheetIndex
          // sheetNames 已跳過前兩個無效項目（空白表、下拉選單）
          // 所以 sheetIndex = index + 2
          const sheetIndex = index + 2;

          option.value = sheetIndex;
          option.textContent = month;
          option.dataset.monthName = month;
          if (sheetIndex === currentSheetIndex) {
            option.selected = true;
          }
          select.appendChild(option);
        });
      }
    }
  } catch (e) {
    // 建立失敗，忽略錯誤（可能已經存在）
  }

  // ===== 新的載入流程：先從 IndexedDB 快取載入，再背景同步 =====
  let loadedFromCache = false;
  let cachedTimestamp = null;

  try {
    // 嘗試從 IndexedDB 載入快取資料
    const cachedSheetNames = await getFromIDB('sheetNames');
    const cachedHasInvalid = await getFromIDB('hasInvalidFirstTwoSheets');
    const cachedBudgetTotals = await getFromIDB('budgetTotals');

    if (cachedSheetNames && cachedSheetNames.length > 0) {
      sheetNames = cachedSheetNames;
      hasInvalidFirstTwoSheets = cachedHasInvalid || false;

      // 根據目前日期，設定對應的 sheet（找最接近的月份）
      const closestSheetIndex = findClosestMonth();
      currentSheetIndex = closestSheetIndex;

      // 載入當前月份的快取資料
      const cachedMonthData = await getFromIDB(`monthData_${currentSheetIndex}`);
      cachedTimestamp = await getCacheTimestamp(`monthData_${currentSheetIndex}`);

      if (cachedMonthData) {
        allMonthsData[currentSheetIndex] = cachedMonthData;

        // Load budget cache - but verify it has correct month data
        if (cachedBudgetTotals && cachedBudgetTotals[currentSheetIndex]) {
          const monthName = sheetNames[currentSheetIndex - 2] || '';
          // Clear stale cache - force fresh load from API
          // budgetTotals[currentSheetIndex] = cachedBudgetTotals[currentSheetIndex];
        }

        loadedFromCache = true;
      }
    }
  } catch (e) {
    // IndexedDB 不可用，繼續使用 API 載入
  }

  try {
    // 確保 post-content 元素存在
    const postContentElements = document.getElementsByClassName('post-content');
    if (!postContentElements || postContentElements.length === 0) {
      // 嘗試等待一下再重試
      setTimeout(() => {
        const retryElements = document.getElementsByClassName('post-content');
        if (retryElements && retryElements.length > 0) {
          retryElements[0].appendChild(totalContainer);
          retryElements[0].appendChild(budgetCardsContainer);
        } else {
        }
      }, 500);
      return;
    }

    const postContent = postContentElements[0];
    // 添加總計容器和表單容器
    postContent.appendChild(totalContainer);
    postContent.appendChild(budgetCardsContainer);

    // 第一個請求：載入下拉選單選項（阻塞式，確保表單可用）
    try {
      await loadDropdownOptions();
      updateSelectOptions('expense-category-select', EXPENSE_CATEGORY_OPTIONS);
      updateSelectOptions('payment-method-select', PAYMENT_METHOD_OPTIONS);
      updateSelectOptions('credit-card-payment-select', CREDIT_CARD_PAYMENT_OPTIONS);
      updateSelectOptions('month-payment-select', MONTH_PAYMENT_OPTIONS);
      updateSelectOptions('payment-platform-select', PAYMENT_PLATFORM_OPTIONS);
    } catch (err) {
      console.error('[支出表] 載入下拉選單失敗:', err);
    }

    // 下拉選單載入完成，隱藏進度條
    hideSpinner();

    // 非阻塞重新載入下拉選單的函數（用於監聽更新）
    const refreshDropdowns = () => {
      loadDropdownOptions().then(() => {
        updateSelectOptions('expense-category-select', EXPENSE_CATEGORY_OPTIONS);
        updateSelectOptions('payment-method-select', PAYMENT_METHOD_OPTIONS);
        updateSelectOptions('credit-card-payment-select', CREDIT_CARD_PAYMENT_OPTIONS);
        updateSelectOptions('month-payment-select', MONTH_PAYMENT_OPTIONS);
        updateSelectOptions('payment-platform-select', PAYMENT_PLATFORM_OPTIONS);
      }).catch(err => {
      });
    };

    // 監聽設定頁的更新通知（當設定頁更新下拉選單後，自動重新載入）
    // 使用 capture 階段確保能捕獲到事件
    window.addEventListener('storage', (e) => {
      if (e.key === 'dropdownUpdated') {
        refreshDropdowns();
      }
    }, true);

    // 監聽同頁面的自定義事件（當同一頁面觸發更新時）
    window.addEventListener('dropdownUpdated', () => {
      refreshDropdowns();
    });

    // 也監聽同頁面的 storage 事件（因為 storage 事件只在其他標籤頁觸發）
    let lastUpdateTime = localStorage.getItem('dropdownUpdated');
    const checkInterval = setInterval(() => {
      const current = localStorage.getItem('dropdownUpdated');
      if (current && current !== lastUpdateTime) {
        lastUpdateTime = current;
        refreshDropdowns();
      }
    }, 500); // 每500毫秒檢查一次，更頻繁地檢查

    // 頁面卸載時清理定時器
    window.addEventListener('beforeunload', () => {
      clearInterval(checkInterval);
    });

    // 將歷史紀錄按鈕添加到頁面標題右側（留空隙）
    const pageTitle = document.querySelector('h1.post-title, h1, .post-title');

    if (pageTitle) {
      pageTitle.style.cssText = 'display: flex; align-items: center; justify-content: space-between;';
      // 在標題和按鈕之間留空隙
      const spacer = document.createElement('div');
      spacer.style.cssText = 'flex: 1;';
      pageTitle.appendChild(spacer);
      pageTitle.appendChild(historyButton);
    } else {
      // 如果找不到標題，創建一個標題容器
      const titleContainer = document.createElement('div');
      titleContainer.style.cssText = 'display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px;';
      const title = document.createElement('h1');
      title.textContent = '支出';
      titleContainer.appendChild(title);
      const spacer = document.createElement('div');
      spacer.style.cssText = 'flex: 1;';
      titleContainer.appendChild(spacer);
      titleContainer.appendChild(historyButton);
      postContent.insertBefore(titleContainer, postContent.firstChild);
    }

    // ===== 快取優先載入邏輯 =====
    if (loadedFromCache) {
      try {
        // 從快取立即顯示資料
        processDataFromResponse(allMonthsData[currentSheetIndex].data, true);
        updateTotalDisplay();

        // 背景同步：從 API 載入最新資料
        syncFromAPI().catch(e => {
          console.error('[支出表] 背景同步失敗:', e);
        });
      } catch (cacheError) {
        // 快取載入失敗，清除快取標記，讓下面的 else 分支處理
        loadedFromCache = false;
      }
    }

    // 監聽手動同步請求（從導航列的同步圖示觸發）
    window.addEventListener('syncRequested', () => {
      syncFromAPI().catch(e => {
        console.error('[支出表] 手動同步失敗:', e);
      });
    });

    if (!loadedFromCache) {
      // 沒有快取，從 API 載入
      try {
        await loadMonthNames();

        // 儲存 sheetNames 到 IndexedDB
        setToIDB('sheetNames', sheetNames).catch(() => {});
        setToIDB('hasInvalidFirstTwoSheets', hasInvalidFirstTwoSheets).catch(() => {});

        // 一開始就顯示總數
        const totalMonths = sheetNames.length;
        const totalProgress = totalMonths + 1;
        updateProgress(0, totalProgress, '載入月份列表');
        updateProgress(1, totalProgress, '載入月份列表');

        // 檢查當前月份是否已有表格
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const currentMonthStr = `${year}${month}`;
        const hasCurrentMonth = sheetNames.includes(currentMonthStr);

        if (!hasCurrentMonth) {
          try {
            const createResult = await callAPI({ name: "Create Tab" });
            alert(createResult.message || `已建立新分頁：${currentMonthStr}`);
            await loadMonthNames();
            setToIDB('sheetNames', sheetNames).catch(() => {});
            const newTotalMonths = sheetNames.length;
            const newTotalProgress = newTotalMonths + 1;
            updateProgress(1, newTotalProgress, '載入月份列表');
          } catch (createError) {
            alert('無法自動建立本月表格，請稍後再試或手動建立。');
          }
        }

        if (sheetNames.length > 0) {
          const closestSheetIndex = findClosestMonth();
          currentSheetIndex = closestSheetIndex;

          updateProgress(2, totalProgress, '載入當前月份');
          const currentMonthData = await loadMonthData(currentSheetIndex);
          allMonthsData[currentSheetIndex] = currentMonthData;

          // 儲存到 IndexedDB
          setToIDB(`monthData_${currentSheetIndex}`, currentMonthData).catch(() => {});

          processDataFromResponse(currentMonthData.data, true);
          updateTotalDisplay();

          // 載入預算
          loadBudgetForMonth(currentSheetIndex)
            .then(() => {
              updateTotalDisplay();
              // 儲存預算到 IndexedDB
              setToIDB('budgetTotals', budgetTotals).catch(() => {});
            })
            .catch(err => {
              console.error('[支出表] 載入預算資料失敗:', err);
            });

          // 背景預載其他月份
          preloadAllMonthsData()
            .then(() => {
              updateTotalDisplay();
              // 儲存所有月份資料到 IndexedDB
              Object.keys(allMonthsData).forEach(idx => {
                setToIDB(`monthData_${idx}`, allMonthsData[idx]).catch(() => {});
              });
            })
            .catch(e => {
              console.error('[支出表] 預載所有月份支出資料失敗:', e);
            });

          // 背景預載其他月份的預算（並行載入）
          const budgetPromises = sheetNames
            .map((name, idx) => idx + 2)
            .filter(sheetIndex => sheetIndex !== currentSheetIndex)
            .map(sheetIndex => loadBudgetForMonth(sheetIndex).catch(() => {}));

          Promise.all(budgetPromises).then(() => {
            setToIDB('budgetTotals', budgetTotals).catch(() => {});
          });
        }
      } catch (e) {
        console.error('[支出表] 初始化失敗:', e);
      }
    }

    // 等待 DOM 更新後再初始化表單，但避免清掉使用者已經輸入的內容
    setTimeout(() => {
      const itemInput = document.getElementById('item-input');
      const actualCostInput = document.getElementById('actual-cost-input');
      const recordCostInput = document.getElementById('record-cost-input');
      const noteInput = document.getElementById('note-input');

      const hasUserInput = !!(
        (itemInput && itemInput.value && itemInput.value.trim() !== '') ||
        (actualCostInput && actualCostInput.value && actualCostInput.value.trim() !== '') ||
        (recordCostInput && recordCostInput.value && recordCostInput.value !== '') ||
        (noteInput && noteInput.value && noteInput.value.trim() !== '')
      );

      if (!hasUserInput) {
        // 只有在表單目前是空的時候才執行初始化，避免清掉使用者正在輸入的內容
      clearForm();
      } else {
      }
    }, 100);

    // 載入總計
    try {
      await loadTotal();
    } catch (e) {
    }
  } catch (error) {
  }
});
</script>
