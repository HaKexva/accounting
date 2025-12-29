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
  // 新增：接收日期欄位（可選）
  var dateValue = contents.date;   // e.g. "2025/01/31" 或 "2025-01-31"
  var item = contents.item;
  var category = contents.category;
  var spendWay = contents.spendWay;
  var creditCard = contents.creditCard
  var month = contents.month
  var actualCost = contents.actualCost
  var payment = contents.payment
  var recordCost = contents.recordCost
  var note = contents.note;
  var updateRow = contents.updateRow;
  var oldValue = contents.oldValue
  var newValue = contents.newValue
  var itemId = contents.itemId
  var action = contents.action
  // 新增：接收拖拽排序的索引
  var oldIndex = contents.oldIndex;
  var newIndex = contents.newIndex;

  var result = { success: false, message: "" };

  try {
    if (name === "Upsert Data") {
      // 傳入 dateValue，讓 UpsertData 可以使用指定日期
      result = UpsertData(sheet, dateValue, item, category, spendWay, creditCard, month, actualCost, payment, recordCost, note, updateRow);
    } else if (name === "Create Tab") {
      result = CreateNewTab();
    } else if (name === "Delete Data") {
      result = DeleteData(sheet,updateRow);
    } else if (name === "Update Dropdown"){
      result = UpdateDropdown(action, itemId, oldValue, newValue, oldIndex, newIndex);
    } else {
      result.message = "未知的操作類型: " + name;
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

function ShowTabName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var sheetNames = [];
  allSheets.slice(2).forEach(function(sheet){
    sheetNames.push(sheet.getSheetName());
  });
  console.log(sheetNames);
  return sheetNames;
}

function ShowTabData(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var lastRow = targetSheet.getRange('A' + targetSheet.getMaxRows() + ':J' + targetSheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow()
  var output = targetSheet.getRange('A1:J' + lastRow).getValues()
  console.log(lastRow)
  console.log(output)
  return output
}

function CreateNewTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var targetSheet = ss.getSheets()[0];
  var value = [["總計",,,,,,"0",,"0"]]
  targetSheet.getRange("A2:I2").setValues(value)
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
    console.log('隔月試算表已新增，請勿繼續新增');
    return { success: false, message: '隔月試算表已新增，請勿繼續新增' };
  } else {
    var destination = SpreadsheetApp.openById('1R-r3cwuVVDP0seQfkyDB4D7D1youEc_-fQlGYRqtkDk');
    targetSheet.copyTo(destination);
    var copyName = ss.getSheetByName('「空白表」的副本');
    copyName.setName(year + month.toString());
  }
  return { success: true, message: '新分頁已成功建立: ' + year + month };
}

function CleanupSummary(sheetIndex) {  // 改為接收 sheetIndex
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];  // 使用 sheetIndex 獲取 sheet
  var column = 10
  var row = sheet.getRange('A' + sheet.getMaxRows() + ':J' + sheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  if (sheet.getRange(row,1).getValue() === '總計') {
    sheet.getRange(row, 1, 1, column).clear();
  }
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

  if (itemId === 'category') {
    colIndex = FindHeaderColumn(headerRow, ['消費類別', '類別']);
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

  // === 新增：找到該欄位的最後一列，在下一行新增 ===
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

    // 1. 更新下拉選單表
    for (var r2 = 1; r2 < data.length; r2++) {
      var val2 = data[r2][colIndex];
      if (val2 === oldValue) {
        data[r2][colIndex] = newValue;
        updatedCount++;
      }
    }
    dropdownSheet.getDataRange().setValues(data);

    // 2. 更新所有月份表格中的歷史資料（從索引 2 開始）
    var allSheets = ss.getSheets();
    var monthSheetsUpdated = 0;
    var totalRecordsUpdated = 0;

    // 確定要更新的欄位索引（根據 UpsertData 的欄位順序）
    // 支出表欄位順序：[日期, 項目, 類別, 支付方式, 信用卡支付方式, 本月/次月支付, 實際消費金額, 支付平台, 列帳消費金額, 備註]
    var targetColumnIndex = -1;
    if (itemId === 'category') {
      targetColumnIndex = 2; // 第3欄（索引2）：類別
    } else if (itemId === 'payment') {
      targetColumnIndex = 3; // 第4欄（索引3）：支付方式
    } else if (itemId === 'platform') {
      targetColumnIndex = 7; // 第8欄（索引7）：支付平台
    }

    if (targetColumnIndex >= 0) {
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

          var cellValue = monthData[row][targetColumnIndex];
          if (cellValue === oldValue) {
            monthSheet.getRange(row + 1, targetColumnIndex + 1).setValue(newValue);
            monthUpdated = true;
            totalRecordsUpdated++;
          }
        }

        if (monthUpdated) {
          monthSheetsUpdated++;
        }
      }
    }

    return {
      success: true,
      message: '已合併完成，下拉選單更新：' + updatedCount + ' 筆，月份表格更新：' + totalRecordsUpdated + ' 筆（' + monthSheetsUpdated + ' 個月份）'
    };
  }

  return { success: false, message: '未知的 action: ' + action };
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

// 調整參數順序：多一個 dateValue
function UpsertData(sheetIndex, dateValue, item, category, spendWay, creditCard, monthIndex, actualCost, payment, recordCost, note, updateRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];
  var startRow = 2;
  var values;

  // 預設為「今天」的日期字串
  var now = new Date();
  var year = now.getFullYear().toString();
  var month = (now.getMonth() + 1).toString();
  var day = now.getDate().toString();
  var defaultDateString = year + '/' + month + '/' + day;

  // 如果前端有傳 date，就用傳進來的；否則用 defaultDateString
  var timeOutput = dateValue ? dateValue : defaultDateString;

  if (updateRow === undefined) { // 新增
    CleanupSummary(sheetIndex);
    var lastDataRow = getLastDataRow(sheet, startRow);
    var insertRow = lastDataRow + 1;

    // 第一欄是日期
    values = [timeOutput, item, category, spendWay, creditCard, monthIndex, actualCost, payment, recordCost, note];
    sheet.getRange(insertRow, 1, 1, 10).setValues([values]);

    var totalRow = insertRow + 1;
    sheet.getRange(totalRow, 1).setValue('總計');

    var firstCell = sheet.getRange(startRow, 7).getA1Notation();
    var lastCell = sheet.getRange(insertRow, 7).getA1Notation();
    sheet.getRange(totalRow, 7).setFormula(`=SUM(${firstCell}:${lastCell})`);

    var firstCell2 = sheet.getRange(startRow, 9).getA1Notation();
    var lastCell2 = sheet.getRange(insertRow, 9).getA1Notation();
    sheet.getRange(totalRow, 9).setFormula(`=SUM(${firstCell2}:${lastCell2})`);

  } else { // 修改
    // 如果前端有傳新的 dateValue，就更新日期；否則沿用原本日期。
    var existingDate = sheet.getRange(updateRow, 1).getValue();
    var finalDate = dateValue ? dateValue : existingDate;

    values = [finalDate, item, category, spendWay, creditCard, monthIndex, actualCost, payment, recordCost, note];
    sheet.getRange(updateRow, 1, 1, 10).setValues([values]);
  }

  return { success: true, message: '資料已成功新增', data: ShowTabData(sheetIndex), total: GetSummary(sheetIndex) };
}

function DeleteData(sheetIndex, updateRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[sheetIndex];
  var startRow = 2;

  // 避免刪到標題或不合法列
  if (!updateRow || updateRow < startRow || updateRow > sheet.getLastRow()) {
    return { success: false, message: '刪除列不合法: ' + updateRow };
  }

  CleanupSummary(sheetIndex);
  sheet.deleteRow(updateRow);

  var lastDataRow = getLastDataRow(sheet, startRow);
  var totalRow = lastDataRow + 1;
  sheet.getRange(totalRow, 1).setValue('總計');

  if (lastDataRow >= startRow) {
    var firstCell = sheet.getRange(startRow, 7).getA1Notation();
    var lastCell = sheet.getRange(lastDataRow, 7).getA1Notation();
    sheet.getRange(totalRow, 7).setFormula(`=SUM(${firstCell}:${lastCell})`);

    var firstCell2 = sheet.getRange(startRow, 9).getA1Notation();
    var lastCell2 = sheet.getRange(lastDataRow, 9).getA1Notation();
    sheet.getRange(totalRow, 9).setFormula(`=SUM(${firstCell2}:${lastCell2})`);
  }

  return {
    success: true,
    message: '資料已成功刪除',
    data: ShowTabData(sheetIndex),
    total: GetSummary(sheetIndex)
  };
}

function getLastDataRow(sheet, startRow) {
  var lastRow = sheet.getLastRow();

  // 完全沒資料（只有標題）
  if (lastRow < startRow) {
    return startRow - 1;
  }

  for (var row = startRow; row <= lastRow; row++) {
    var val = sheet.getRange(row, 1).getValue();
    if (val === '' || val === '總計' || val === null) {
      return row - 1;
    }
  }

  return lastRow;
}

function GetSummary(sheet){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheets()[sheet];
  var lastRow = targetSheet.getRange('A900').getNextDataCell(SpreadsheetApp.Direction.UP).getRow()-1
  var income1 = targetSheet.getRange('A:J').getCell(lastRow,7).getValues()[0][0]
  var income2 = targetSheet.getRange('A:J').getCell(lastRow,9).getValues()[0][0]
  return [income1,income2]
}

