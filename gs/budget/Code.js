function doGet(e) {
  var name = e.parameter.name;
  var sheet = e.parameter.sheet;
  if (name === "Show Tab Name") {
    var tab = ShowTabName();
    return _json(tab);
  } else if (name === "Show Tab Data") {
    var data = ShowTabData(sheet);
    return _json(data);
  } else if (name === "Show Total"){
    var data = GetSummary(sheet)
    return _json(data)
  }
}

function doPost(e) {
  var contents = JSON.parse(e.postData.contents)
  var name = contents.name;
  var sheet = contents.sheet;
  var range = contents.range;
  var sheetName = contents.sheetName;
  var category = contents.category;
  var item = contents.item;
  var cost = contents.cost;
  var note = contents.note;
  var number = contents.number
  var updateRow = contents.updateRow;
  var oldValue = contents.oldValue
  var newValue = contents.newValue
  var itemId = contents.itemId
  var action = contents.action
  // 新增：接收拖拽排序的索引
  var oldIndex = contents.oldIndex;
  var newIndex = contents.newIndex;
  // 新增：接收批次更新的資料
  var originalData = contents.originalData;
  var newData = contents.newData;

  var result = { success: false, message: "" };

  try {
    if (name === "Upsert Data") {
      result = UpsertData(sheet,range,category,item,cost,note,updateRow);
    } else if (name === "Create Tab") {
      result = CreateNewTab();
    } else if (name === "Delete Data") {
      result = DeleteData(sheet,range, number);
    } else if (name === "Delete Tab") {
      result = DeleteTab(sheet);
    } else if (name === "Change Tab Name") {
      result = ChangeTabName(sheet,sheetName);
    } else if (name === "Update Dropdown"){
      result = UpdateDropdown(action, itemId, oldValue, newValue, oldIndex, newIndex);
    } else if (name === "Batch Update Dropdown") {
      result = BatchUpdateDropdown(itemId, originalData, newData);
    } else {
      result.message = "未知的操作類型: " + e.parameter;
    }
  } catch (error) {
    result.success = false;
    result.message = "操作失敗: " + error.toString();
  }

  return _json(result);
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Cache helper functions
var CACHE_EXPIRATION = 300; // 5 minutes

function getCacheKey(prefix, sheetIndex) {
  return prefix + '_' + (sheetIndex || 'all');
}

function getFromCache(key) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      return null;
    }
  }
  return null;
}

function setToCache(key, data) {
  var cache = CacheService.getScriptCache();
  try {
    var jsonStr = JSON.stringify(data);
    // Only cache if data is less than 100KB (CacheService limit)
    if (jsonStr.length < 100000) {
      cache.put(key, jsonStr, CACHE_EXPIRATION);
    }
  } catch (e) {
    // Ignore cache errors
  }
}

function invalidateCache(sheetIndex) {
  var cache = CacheService.getScriptCache();
  cache.remove(getCacheKey('tabData', sheetIndex));
  cache.remove(getCacheKey('summary', sheetIndex));
}

function ShowTabName() {
  var cacheKey = getCacheKey('tabNames');
  var cached = getFromCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var sheetNames = [];
  
  // 從索引 2 開始（跳過前兩個「空白表」、「下拉選單」）
  // slice(2) 會包含從索引 2 開始到最後的所有元素
  var sheetsToProcess = allSheets.slice(2);
  sheetsToProcess.forEach(function(sheet){
    sheetNames.push(sheet.getSheetName());
  });

  setToCache(cacheKey, sheetNames);
  return sheetNames;
}

// 批次更新下拉選單
function BatchUpdateDropdown(itemId, originalData, newData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dropdownSheet = ss.getSheets()[1];
  var data = dropdownSheet.getDataRange().getValues();

  if (data.length === 0) {
    return { success: false, message: '下拉選單表是空的' };
  }

  var headerRow = data[0];
  var colIndex = -1;

  // 找到對應欄位
  if (itemId === 'category') {
    colIndex = FindHeaderColumn(headerRow, ['支出－項目', '支出-項目', '消費類別', '類別']);
  } else if (itemId === 'payment') {
    colIndex = FindHeaderColumn(headerRow, ['支付方式']);
  } else if (itemId === 'platform') {
    colIndex = FindHeaderColumn(headerRow, ['支付平台', '平台']);
  }

  if (colIndex < 0) {
    return { success: false, message: '找不到對應欄位' };
  }

  // 驗證資料
  if (!Array.isArray(originalData) || !Array.isArray(newData)) {
    return { success: false, message: 'originalData 和 newData 必須是陣列' };
  }

  // 找出被刪除的項目（在原始資料中但不在新資料中）
  var removedItems = originalData.filter(function(item) {
    return newData.indexOf(item) === -1;
  });

  // 找出新增的項目（在新資料中但不在原始資料中）
  var addedItems = newData.filter(function(item) {
    return originalData.indexOf(item) === -1;
  });

  // 偵測重新命名：如果刪除和新增數量相同，假設是重新命名
  var renames = [];
  if (removedItems.length === addedItems.length && removedItems.length > 0) {
    for (var i = 0; i < removedItems.length; i++) {
      renames.push({
        oldValue: removedItems[i],
        newValue: addedItems[i]
      });
    }
  }

  // 1. 更新下拉選單表：清除該欄位並寫入新資料
  var maxRows = dropdownSheet.getMaxRows();
  if (maxRows > 1) {
    dropdownSheet.getRange(2, colIndex + 1, maxRows - 1, 1).clearContent();
  }

  // 寫入新資料
  if (newData.length > 0) {
    var writeData = newData.map(function(item) { return [item]; });
    dropdownSheet.getRange(2, colIndex + 1, newData.length, 1).setValues(writeData);
  }

  // 2. 如果是 category，更新所有月份表格中的歷史資料
  var historicalUpdates = 0;
  if (itemId === 'category' && renames.length > 0) {
    var allSheets = ss.getSheets();
    var categoryColumnIndex = 8; // I 欄（類別）
    var expenseStartColumnIndex = 6; // G 欄（支出記錄的起始欄位，編號）

    for (var s = 2; s < allSheets.length; s++) {
      var monthSheet = allSheets[s];
      var monthData = monthSheet.getDataRange().getValues();

      if (monthData.length < 2) continue;

      for (var row = 1; row < monthData.length; row++) {
        if (monthData[row][0] === '總計' || monthData[row][0] === '') {
          continue;
        }

        var expenseNumber = monthData[row][expenseStartColumnIndex];
        if (expenseNumber !== '' && expenseNumber !== null && expenseNumber !== undefined) {
          var categoryValue = monthData[row][categoryColumnIndex];

          for (var r = 0; r < renames.length; r++) {
            if (categoryValue === renames[r].oldValue) {
              monthSheet.getRange(row + 1, categoryColumnIndex + 1).setValue(renames[r].newValue);
              historicalUpdates++;
              break;
            }
          }
        }
      }
    }
  }

  // 組裝回傳訊息
  var message = '下拉選單已更新';
  if (renames.length > 0) {
    message += '，重新命名 ' + renames.length + ' 項';
    if (historicalUpdates > 0) {
      message += '，歷史記錄更新 ' + historicalUpdates + ' 筆';
    }
  }
  if (removedItems.length > addedItems.length) {
    message += '，刪除 ' + (removedItems.length - addedItems.length) + ' 項';
  }
  if (addedItems.length > removedItems.length) {
    message += '，新增 ' + (addedItems.length - removedItems.length) + ' 項';
  }

  return {
    success: true,
    message: message,
    renames: renames,
    historicalUpdates: historicalUpdates
  };
}

function UpdateDropdown(action, itemId, oldValue, newValue, oldIndex, newIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dropdownSheet = ss.getSheets()[1];
  var data = dropdownSheet.getDataRange().getValues();

  if (data.length === 0) {
    return { success: false, message: '下拉選單表是空的' };
  }

  var headerRow = data[0];
  var colIndex = -1;

  // 在下拉選單表中，category 對應的欄位名稱是「支出－項目」
  if (itemId === 'category') {
    colIndex = FindHeaderColumn(headerRow, ['支出－項目', '支出-項目', '消費類別', '類別']);
  } else if (itemId === 'payment') {
    colIndex = FindHeaderColumn(headerRow, ['支付方式']);
  } else if (itemId === 'platform') {
    colIndex = FindHeaderColumn(headerRow, ['支付平台', '平台']);
  }

  if (colIndex < 0) {
    return { success: false, message: '找不到對應欄位' };
  }

  // === 新增：拖拽排序功能 ===
  if (action === 'reorder') {
    // 檢查索引參數
    if (typeof oldIndex !== 'number' || typeof newIndex !== 'number') {
      return { success: false, message: '排序時 oldIndex 和 newIndex 必須是數字' };
    }

    // 收集該欄位的所有值（跳過標題行）
    var values = [];
    for (var r = 1; r < data.length; r++) {
      var val = data[r][colIndex];
      if (val !== '' && val !== null && val !== undefined) {
        values.push(val);
      }
    }

    // 檢查索引範圍
    if (oldIndex < 0 || oldIndex >= values.length || newIndex < 0 || newIndex >= values.length) {
      return { success: false, message: '索引超出範圍' };
    }

    // 執行排序：將 oldIndex 位置的項移到 newIndex 位置
    var movedItem = values.splice(oldIndex, 1)[0];
    values.splice(newIndex, 0, movedItem);

    // 更新下拉選單表（從第2行開始寫入）
    var writeRow = 2;
    for (var i = 0; i < values.length; i++) {
      dropdownSheet.getRange(writeRow, colIndex + 1).setValue(values[i]);
      writeRow++;
    }

    // 清除後面多餘的單元格
    if (writeRow <= data.length) {
      dropdownSheet.getRange(writeRow, colIndex + 1, data.length - writeRow + 1, 1).clearContent();
    }

    return { success: true, message: '排序成功，已更新下拉選單順序' };
  }

  var updatedCount = 0;

  // === 新增：只更新下拉選單表 ===
  if (action === 'add') {
    if (!newValue) {
      return { success: false, message: '新增時 newValue 必填' };
    }

    // 找到該欄位的最後一個有值的行（從下往上找）
    var lastRowInColumn = 1; // 預設為標題行（索引1，對應第2行）
    for (var r = data.length - 1; r >= 1; r--) { // 從最後一行往上找，跳過標題行
      var val = data[r][colIndex];
      if (val !== '' && val !== null && val !== undefined) {
        lastRowInColumn = r + 1; // 找到最後一個有值的行（+1 因為要轉換為行號）
        break;
      }
    }

    // 在該欄位的最後一列的下一個位置新增
    dropdownSheet.getRange(lastRowInColumn + 1, colIndex + 1).setValue(newValue);
    return { success: true, message: '已新增到下拉選單', updated: 1 };
  }

  // === 刪除：只更新下拉選單表 ===
  if (action === 'delete') {
    if (!oldValue) {
      return { success: false, message: '刪除時 oldValue 必填（不能是空白）' };
    }
    var writeRow = 2;
    for (var r = 1; r < data.length; r++) {
      var val = data[r][colIndex];
      if (val === oldValue) {
        updatedCount++;
        continue;
      }
      dropdownSheet.getRange(writeRow, colIndex + 1).setValue(val);
      writeRow++;
    }
    if (writeRow <= data.length) {
      dropdownSheet.getRange(writeRow, colIndex + 1, data.length - writeRow + 1, 1).clearContent();
    }
    return { success: true, message: '已從下拉選單刪除', updated: updatedCount };
  }

  // === 編輯（合併）：更新下拉選單表 + 所有月份表格中的歷史資料 ===
  if (action === 'edit') {
    if (!oldValue || !newValue || oldValue === newValue) {
      return { success: false, message: '編輯時 oldValue / newValue 不正確（oldValue 不可空白）' };
    }

    // 1. 更新下拉選單表（欄位名稱：支出－項目）
    for (var r2 = 1; r2 < data.length; r2++) {
      var val2 = data[r2][colIndex];
      if (val2 === oldValue) {
        data[r2][colIndex] = newValue;
        updatedCount++;
      }
    }
    dropdownSheet.getDataRange().setValues(data);

    // 2. 更新所有月份表格中的歷史資料（只更新 range = 0 的支出記錄的「類別」欄位）
    // 只有當 itemId === 'category' 時才更新月份表格
    if (itemId === 'category') {
      var allSheets = ss.getSheets();
      var monthSheetsUpdated = 0;
      var totalRecordsUpdated = 0;

      // 預算表中，支出記錄從 G 欄開始（索引 6）
      // G 欄（索引 6）：編號
      // H 欄（索引 7）：時間
      // I 欄（索引 8）：類別（category）- 這是我們要更新的欄位
      var categoryColumnIndex = 8; // I 欄（類別）
      var expenseStartColumnIndex = 6;  // G 欄（支出記錄的起始欄位，編號）

      for (var s = 2; s < allSheets.length; s++) { // 從索引 2 開始（跳過前兩個 sheet）
        var monthSheet = allSheets[s];
        var monthData = monthSheet.getDataRange().getValues();

        if (monthData.length < 2) continue; // 跳過空表格

        var monthUpdated = false;

        // 從第2行開始（索引1），跳過標題行
        for (var row = 1; row < monthData.length; row++) {
          // 檢查是否為「總計」行或空行
          if (monthData[row][0] === '總計' || monthData[row][0] === '') {
            continue;
          }

          // 只處理 range = 0 的支出記錄（從 G 欄開始的記錄）
          // 檢查該行的 G 欄是否有值（編號），如果有則表示這是支出記錄
          var expenseNumber = monthData[row][expenseStartColumnIndex]; // G 欄（編號）

          // 如果 G 欄有值（是支出記錄），才檢查 I 欄的類別
          if (expenseNumber !== '' && expenseNumber !== null && expenseNumber !== undefined) {
            var categoryValue = monthData[row][categoryColumnIndex]; // I 欄（類別）

            if (categoryValue === oldValue) {
              monthSheet.getRange(row + 1, categoryColumnIndex + 1).setValue(newValue);
              monthUpdated = true;
              totalRecordsUpdated++;
            }
          }
        }

        if (monthUpdated) {
          monthSheetsUpdated++;
        }
      }

      return {
        success: true,
        message: '已合併完成，下拉選單更新：' + updatedCount + ' 筆，月份表格更新：' + totalRecordsUpdated + ' 筆（' + monthSheetsUpdated + ' 個月份）'
      };
    } else {
      // 如果不是 category，只更新下拉選單表
      return {
        success: true,
        message: '已更新下拉選單：' + updatedCount + ' 筆'
      };
    }
  }

  return { success: false, message: '未知的 action: ' + action };
}

function ShowTabData(sheet) {
  var cacheKey = getCacheKey('tabData', sheet);
  var cached = getFromCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var ranges = targetSheet.getNamedRanges();
  var res = {};
  ranges.forEach(function(range) {
    var title = range.getName();
    var rng = range.getRange();
    res[title] = cleanValues(rng.getValues());
  });

  function isEmpty(v) {
    return v === "" || v === undefined || v === null;
  }

  function cleanValues(values) {
    var res = [];
    values.forEach(function(value){
      if(!value.every(isEmpty)) {
        res.push(value);
      }
    });
    return res;
  }

  setToCache(cacheKey, res);
  return res;
}

function FindHeaderColumn(headerRow, keywords) {
  for (var c = 0; c < headerRow.length; c++) {
    var headerText = (headerRow[c] || '').toString().trim();
    if (!headerText) continue;
    for (var k = 0; k < keywords.length; k++) {
      if (headerText.indexOf(keywords[k]) !== -1) {
        return c;
      }
    }
  }
  return -1;
}

function CreateNewTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var targetSheet = ss.getSheets()[0];
  var incomeValue = [["總計",,,"0"]]
  targetSheet.getRange("A3:D3").setValues(incomeValue)
  var expenseValue = [["總計",,,,"0"]]
  targetSheet.getRange("G3:K3").setValues(expenseValue)
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth()+1;
  var sheetNames = [];
  allSheets.forEach(function(sheet){
    sheetNames.push(sheet.getSheetName());
  });
  if (month < 10) {
    month = '0' + month;
  }
  if (sheetNames.includes(year.toString() + month.toString()) && parseInt(month)+1 <= 12) {
    month = parseInt(month) + 1;
  } else if (sheetNames.includes(year.toString() + month.toString()) && parseInt(month)+1 > 12){
    year = parseInt(year) + 1;
    month = 1;
  }
  if (parseInt(month) < 10) {
    month = '0' + parseInt(month);
  }
  if (sheetNames.includes(year.toString() + month.toString())) {
    return { success: false, message: '隔月試算表已新增，請勿繼續新增' };
  } else {
    var destination = SpreadsheetApp.openById('1G-c8fzN38vH02Eu8izCc9TpGAiPQBMYH-z3tUB_4tVM');
    targetSheet.copyTo(destination);
    var copyName = ss.getSheetByName('「空白表」的副本');
    copyName.setName(year + month.toString());
  }
  month = month.toString();
  var targetSheet = ss.getSheetByName(year.toString()+month.toString());
  var range = targetSheet.getRange('A2:E');
  var name = '當月收入' + year.toString() + month.toString();
  ss.setNamedRange(name,range);
  var range = targetSheet.getRange('G2:L');
  var name = '當月支出預算' + year.toString() + month.toString();
  ss.setNamedRange(name,range);
  return { success: true, message: '新分頁已成功建立: ' + year + month };
}


function UpsertData(sheetIndex, rangeType, category, item, cost, note, updateRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];
  var startRow = 2; // 從第二行開始
  var column, totalColumn;

  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1;
  var date = now.getDate();
  var hour = now.getHours();
  var minute = now.getMinutes();
  var timeOutput = year + '/' + month + '/' + date + ' ' + hour + ':' + minute;

  // ===== 先刪掉舊總計 =====
  RemoveSummaryRow(sheet, startRow, rangeType === 0 ? 7 : 1);

  if (updateRow === undefined) { // 新增
    if (rangeType === 0) { // 支出
      column = 6; // 編號~備註
      totalColumn = 5; // 計算總和的欄

      // 找最後一筆資料列
      var lastDataRow = sheet.getLastRow();
      if (lastDataRow < startRow) lastDataRow = startRow - 1;

      // 取最後編號
      var prev = sheet.getRange(lastDataRow, 7).getValue(); // G欄編號
      var lastNumber = Number(prev);
      if (isNaN(lastNumber)) lastNumber = 0;

      var row = lastDataRow + 1;
      // 確保 cost 是數字類型
      var costValue = cost;
      if (typeof costValue === 'string') {
        costValue = parseFloat(costValue) || 0;
      } else if (costValue === null || costValue === undefined) {
        costValue = 0;
      }
      var values = [lastNumber + 1, timeOutput, category, item, costValue, note];
      sheet.getRange(row, 7, 1, column).setValues([values]);

      // 新增總計
      sheet.getRange(row + 1, 7).setValue('總計');
      sheet.getRange(row + 1, 11).setFormula(`=SUM(K${startRow}:K${row})`);

    } else { // 收入
      column = 5;
      totalColumn = 4;

      var lastDataRow = sheet.getLastRow();
      if (lastDataRow < startRow) lastDataRow = startRow - 1;

      var prev = sheet.getRange(lastDataRow, 1).getValue(); // A欄編號
      var lastNumber = Number(prev);
      if (isNaN(lastNumber)) lastNumber = 0;

      var row = lastDataRow + 1;
      // 確保 cost 是數字類型
      var costValue = cost;
      if (typeof costValue === 'string') {
        costValue = parseFloat(costValue) || 0;
      } else if (costValue === null || costValue === undefined) {
        costValue = 0;
      }
      var values = [lastNumber + 1, timeOutput, item, costValue, note];
      sheet.getRange(row, 1, 1, column).setValues([values]);

      // 新增總計
      sheet.getRange(row + 1, 1).setValue('總計');
      sheet.getRange(row + 1, totalColumn).setFormula(`=SUM(D${startRow}:D${row})`);
    }

  } else { // 修改資料
    if (rangeType === 0) { // 支出
      column = 6;
      totalColumn = 11; // K欄是支出金額總計
      
      // 確保 cost 是數字類型
      var costValue = cost;
      if (typeof costValue === 'string') {
        costValue = parseFloat(costValue) || 0;
      } else if (costValue === null || costValue === undefined) {
        costValue = 0;
      }
      
      // 先更新資料（在刪除總計行之前，避免行號變化）
      var values = [updateRow - 2, timeOutput, category, item, costValue, note];
      sheet.getRange(updateRow, 7, 1, column).setValues([values]);
      
      // 更新總計行（在更新資料之後）
      RemoveSummaryRow(sheet, startRow, 7); // 刪除舊總計行（G欄）
      
      // 找到最後一筆支出資料行（跳過總計行）
      var lastDataRow = sheet.getLastRow();
      while (lastDataRow >= startRow) {
        var cellValue = sheet.getRange(lastDataRow, 7).getValue(); // G欄
        if (cellValue !== '總計' && cellValue !== '' && cellValue !== null && cellValue !== undefined) {
          // 檢查是否是數字編號（有效的資料行）
          var numValue = Number(cellValue);
          if (!isNaN(numValue) && numValue > 0) {
            break; // 找到最後一筆資料行
          }
        }
        lastDataRow--;
      }
      
      if (lastDataRow < startRow) {
        // 沒有資料就設總計為 0
        sheet.getRange(startRow, 7).setValue('總計');
        sheet.getRange(startRow, totalColumn).setValue(0);
      } else {
        sheet.getRange(lastDataRow + 1, 7).setValue('總計');
        sheet.getRange(lastDataRow + 1, totalColumn).setFormula(`=SUM(K${startRow}:K${lastDataRow})`);
      }
    } else { // 收入
      column = 5;
      totalColumn = 4; // D欄是收入金額總計
      
      // 確保 cost 是數字類型
      var costValue = cost;
      if (typeof costValue === 'string') {
        costValue = parseFloat(costValue) || 0;
      } else if (costValue === null || costValue === undefined) {
        costValue = 0;
      }
      
      // 先更新資料（在刪除總計行之前，避免行號變化）
      var values = [updateRow - 2, timeOutput, item, costValue, note];
      sheet.getRange(updateRow, 1, 1, column).setValues([values]);
      
      // 更新總計行（在更新資料之後）
      RemoveSummaryRow(sheet, startRow, 1); // 刪除舊總計行（A欄）
      
      // 找到最後一筆收入資料行（跳過總計行）
      var lastDataRow = sheet.getLastRow();
      while (lastDataRow >= startRow) {
        var cellValue = sheet.getRange(lastDataRow, 1).getValue(); // A欄
        if (cellValue !== '總計' && cellValue !== '' && cellValue !== null && cellValue !== undefined) {
          // 檢查是否是數字編號（有效的資料行）
          var numValue = Number(cellValue);
          if (!isNaN(numValue) && numValue > 0) {
            break; // 找到最後一筆資料行
          }
        }
        lastDataRow--;
      }
      
      if (lastDataRow < startRow) {
        // 沒有資料就設總計為 0
        sheet.getRange(startRow, 1).setValue('總計');
        sheet.getRange(startRow, totalColumn).setValue(0);
      } else {
        sheet.getRange(lastDataRow + 1, 1).setValue('總計');
        sheet.getRange(lastDataRow + 1, totalColumn).setFormula(`=SUM(D${startRow}:D${lastDataRow})`);
      }
    }
  }

  // Invalidate cache after data modification
  invalidateCache(sheetIndex);

  return { success: true, message: '資料已成功新增', data: ShowTabData(sheetIndex), total: GetSummary(sheetIndex) };
}

// ===== 補充：刪除舊總計行 =====
function RemoveSummaryRow(sheet, startRow, labelCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  if (sheet.getRange(lastRow, labelCol).getValue() === '總計') {
    sheet.deleteRow(lastRow);
  }
}


function DeleteTab(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var sheetName = targetSheet.getSheetName();
  ss.deleteSheet(targetSheet);

  return { success: true, message: '分頁已成功刪除: ' + sheetName };
}

function DeleteData(sheetIndex, rangeType, number) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];
  var startRow = 2; // 從第二行開始
  var column, totalColumn;

  if (rangeType === 0) { // 支出
    column = 6;       // 編號~備註
    totalColumn = 11; // K欄是支出金額總計
    var numberCol = 7; // G欄是編號
  } else { // 收入
    column = 5;       // 編號~備註
    totalColumn = 4;  // D欄是收入金額總計
    var numberCol = 1; // A欄是編號
  }

  // 找到要刪除的行
  var lastRow = sheet.getLastRow();
  var targetRow = -1;
  for (var r = startRow; r <= lastRow; r++) {
    var cellVal = sheet.getRange(r, numberCol).getValue();
    if (cellVal === number) {
      targetRow = r;
      break;
    }
  }

  if (targetRow === -1) {
    return { success: false, message: '找不到對應編號資料' };
  }

  // 刪除該行
  sheet.deleteRow(targetRow);

  // 更新總計
  lastRow = sheet.getLastRow();
  // 刪掉舊總計行（如果存在）
  if (sheet.getRange(lastRow, numberCol).getValue() === '總計') {
    sheet.deleteRow(lastRow);
    lastRow--;
  }

  // 設定新的總計行
  var firstDataRow = startRow;
  var lastDataRow = lastRow;
  if (lastDataRow < firstDataRow) {
    // 沒有資料就設總計為 0
    sheet.getRange(firstDataRow, numberCol).setValue('總計');
    sheet.getRange(firstDataRow, totalColumn).setValue(0);
  } else {
    sheet.getRange(lastDataRow + 1, numberCol).setValue('總計');
    sheet.getRange(lastDataRow + 1, totalColumn).setFormula(
      `=SUM(${sheet.getRange(firstDataRow, totalColumn).getA1Notation()}:${sheet.getRange(lastDataRow, totalColumn).getA1Notation()})`
    );
  }

  // Invalidate cache after data modification
  invalidateCache(sheetIndex);

  return {
    success: true,
    message: '資料已成功刪除，總計已更新',
    data: ShowTabData(sheetIndex),
    total: GetSummary(sheetIndex)
  };
}

function ChangeTabName(sheet,name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var oldName = targetSheet.getSheetName();
  targetSheet.setName(name);

  return { success: true, message: '分頁名稱已從 "' + oldName + '" 更改為 "' + name + '"' };
}

function GetSummary(sheet){
  var cacheKey = getCacheKey('summary', sheet);
  var cached = getFromCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var income = 0;
  var expense = 0;
  var incomeFound = false;
  var expenseFound = false;

  // Only read rows that have data (not entire column)
  var lastRow = targetSheet.getLastRow();
  if (lastRow > 0) {
    // Read both income (A:D) and expense (G:K) in one call
    var data = targetSheet.getRange(1, 1, lastRow, 11).getValues();

    // 從後往前查找，找到最後一個"總計"行（應該是最新的）
    for (var i = data.length - 1; i >= 0; i--) {
      // Find income total: look for "總計" in column A, get value from column D
      if (data[i][0] === '總計' && !incomeFound) {
        // 直接從 D 欄讀取值（索引 3）
        var incomeCell = targetSheet.getRange(i + 1, 4); // D 欄是第 4 列
        var incomeValue = incomeCell.getValue();
        // 確保值是數字類型
        if (typeof incomeValue === 'number') {
          income = incomeValue;
        } else if (typeof incomeValue === 'string' && incomeValue.trim() !== '') {
          income = parseFloat(incomeValue) || 0;
        } else {
          income = 0;
        }
        incomeFound = true;
      }
      // Find expense total: look for "總計" in column G, get value from column K
      if (data[i][6] === '總計' && !expenseFound) {
        // 直接從 K 欄讀取值（索引 10，對應第 11 列）
        var expenseCell = targetSheet.getRange(i + 1, 11); // K 欄是第 11 列
        var expenseValue = expenseCell.getValue();
        var cellFormula = expenseCell.getFormula();
        
        // 如果有公式，強制重新計算並讀取顯示值
        if (cellFormula && cellFormula.trim() !== '') {
          // 使用 getDisplayValue() 獲取公式計算後的顯示值
          var displayValue = expenseCell.getDisplayValue();
          if (displayValue && displayValue.trim() !== '') {
            // 移除千分位符號等格式字符
            var cleanValue = displayValue.replace(/,/g, '').trim();
            expenseValue = parseFloat(cleanValue);
            if (isNaN(expenseValue)) {
              expenseValue = expenseCell.getValue(); // 回退到 getValue()
            }
          } else {
            // 如果顯示值為空，嘗試使用 getValue()
            expenseValue = expenseCell.getValue();
          }
        }
        
        // 確保值是數字類型
        if (typeof expenseValue === 'number' && !isNaN(expenseValue)) {
          expense = expenseValue;
        } else if (typeof expenseValue === 'string' && expenseValue.trim() !== '') {
          // 移除千分位符號等格式字符
          var cleanValue = expenseValue.replace(/,/g, '').trim();
          expense = parseFloat(cleanValue) || 0;
        } else {
          expense = 0;
        }
        
        // 如果讀取到的值是 0 但公式存在，且總計行之前有數據行，嘗試手動計算
        if (expense === 0 && cellFormula && cellFormula.trim() !== '' && i + 1 > 2) {
          // 檢查是否有支出記錄（從第 2 行開始到總計行之前）
          var hasExpenseData = false;
          for (var checkRow = 2; checkRow < i + 1; checkRow++) {
            var checkNumber = targetSheet.getRange(checkRow, 7).getValue(); // G 欄編號
            if (typeof checkNumber === 'number' && checkNumber > 0) {
              hasExpenseData = true;
              break;
            }
          }
          
          // 如果有支出記錄但總計是 0，手動計算
          if (hasExpenseData) {
            var expenseStartRow = 2; // 資料從第 2 行開始
            var expenseSum = 0;
            for (var r = expenseStartRow; r < i + 1; r++) {
              var expenseRowData = targetSheet.getRange(r, 7, 1, 6).getValues()[0]; // G 到 L 欄
              // 檢查是否是有效的支出記錄（G 欄應該是數字編號）
              var rowNumber = expenseRowData[0];
              if (typeof rowNumber === 'number' && rowNumber > 0) {
                // K 欄是金額（索引 4）
                var rowCost = expenseRowData[4];
                if (typeof rowCost === 'number' && !isNaN(rowCost)) {
                  expenseSum += rowCost;
                } else if (typeof rowCost === 'string' && rowCost.trim() !== '') {
                  var numCost = parseFloat(rowCost.replace(/,/g, '')) || 0;
                  expenseSum += numCost;
                }
              }
            }
            if (expenseSum !== 0) {
              expense = expenseSum;
            }
          }
        }
        
        expenseFound = true;
      }
      // Stop if both found
      if (incomeFound && expenseFound) break;
    }
  }

  var total = income - expense;
  var result = [income, expense, total];
  setToCache(cacheKey, result);
  return result;
}

function arrayMatch(a1, a2) {
  if (!Array.isArray(a1) || !Array.isArray(a2)) return false;
  if (a1.length !== a2.length) return false;
  for (let i = 0; i < a1.length; i++) {
    const val1 = a1[i];
    const val2 = a2[i];
    const bothArrays = Array.isArray(val1) && Array.isArray(val2);
    if (bothArrays) {
      if (!arrayMatch(val1, val2)) return false;
    } else if (val1 !== val2) {
      return false;
    }
  }
  return true;
}

