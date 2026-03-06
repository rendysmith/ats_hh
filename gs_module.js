/**
 * Возвращает максимальное число из указанного столбца вкладки
 *
 * @param {string} tab_name - имя листа.
 * @return {string} Имя лучшего кандидата и ссылка на него.
 * @customfunction
 */
function getTheBestName(sheetName, position) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return "Лист не найден";
  }

  var datas = getData(sheetName);
  var data = datas[1];

  Logger.log(data);

  if (data.length == 1) {
    return "No datas"
  }

  Logger.log(data.length);

  var data_list = data[0];

  var pos_match = data_list.indexOf('perc_match');
  var pos_name = data_list.indexOf('name');

  var numbers = [];
  for (let i = 1; i < data.length; i++) {
    numbers.push(data[i][pos_match]);
  }

  var max_number = Math.max.apply(null, numbers);

  for (var i = 1; i < data.length; i++) {
    if (data[i][pos_match] == max_number) {
      var bestName = data[i][pos_name];

      if (position == 1) {
        return bestName;
      }

      break;
    }
  }

  var gid = sheet.getSheetId();
  var link = "#gid=" + gid + "&range=B" + (i + 1);
  //Logger.log(link)
  return link
  // var nameLink = SpreadsheetApp.newRichTextValue()
  //   .setText(bestName)
  //   .setLinkUrl(link)
  //   .build();

  //Logger.log(nameLink);
  //return nameLink;
}

function getData(tab_name) {
  // Открываем таблицу по ID
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tab_name);

  // Проверяем, найдена ли вкладка
  if (!sheet) {
    Logger.log(Utilities.formatString('Вкладка "%s" не найдена!', tab_name));
    return;
  }

  // Получаем все данные из вкладки (начиная с первой строки и первого столбца)
  var data = sheet.getDataRange().getValues();

  // Выводим полученные данные в лог (для проверки)
  return [sheet, data];
}

function writeDataByColumnNameAndRowNumber(sheet, columnName, rowNumber, value) {
  // Получаем первую строку листа, чтобы найти название столбца
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnNumber = -1;

  // Ищем индекс столбца по названию
  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === columnName) {
      columnNumber = i + 1; // Индексы столбцов начинаются с 1
      break;
    }
  }

  // Проверяем, был ли найден столбец
  if (columnNumber !== -1) {
    // Записываем значение в указанную ячейку
    sheet.getRange(rowNumber, columnNumber).setValue(value);
    Logger.log("Значение '" + value + "' записано в столбец '" + columnName + "', строка " + rowNumber);
  } else {
    Logger.log("Столбец с названием '" + columnName + "' не найден.");
  }
}

function appendDataByColumnName(sheet, columnName, value) {
  // Get the header row to find the column name
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnNumber = -1;

  // Find the column index by name
  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === columnName) {
      columnNumber = i + 1; // Column indices start from 1
      break;
    }
  }

  // Check if the column was found
  if (columnNumber !== -1) {
    // Determine the next empty row
    var newRowNumber = sheet.getLastRow() + 1;

    // Write the value to the specified cell in the new row
    sheet.getRange(newRowNumber, columnNumber).setValue(value);
    Logger.log("Value '" + value + "' written to column '" + columnName + "', new row " + newRowNumber);
  } else {
    Logger.log("Column with name '" + columnName + "' not found.");
  }
}

function appendRowByColumnNamesOLD(sheet, rowData) {
  // Get the header row to find column numbers
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Determine the next empty row
  var newRowNumber = sheet.getLastRow() + 1;
  
  // Create an array to hold the values for the new row, initially filled with empty strings
  var newRowValues = new Array(headerRow.length).fill("");

  var columnsFound = 0; // To track if any relevant columns were found

  // Iterate over the header row to find column indices and populate newRowValues
  for (var i = 0; i < headerRow.length; i++) {
    var columnName = headerRow[i];
    if (rowData.hasOwnProperty(columnName)) {
      newRowValues[i] = rowData[columnName];
      columnsFound++;
    }
  }

  // Check if any specified columns were found in the header
  if (columnsFound > 0) {
    // Write the entire row of values to the new row
    sheet.getRange(newRowNumber, 1, 1, newRowValues.length).setValues([newRowValues]);
    Logger.log("Data written to new row " + newRowNumber + ". Data: " + JSON.stringify(rowData));
  } else {
    Logger.log("No matching columns found for the provided data in rowData: " + JSON.stringify(rowData));
  }
}

function appendRowByColumnNames(sheet, rowData) {
  var lastCol = sheet.getLastColumn();
  var headerRow = [];
  
  // Если лист пустой — создаём заголовки
  if (lastCol === 0) {
    headerRow = Object.keys(rowData);
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  } else {
    headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }

  // Проверяем новые ключи, которых нет в заголовках
  var newKeys = Object.keys(rowData).filter(function(key) {
    return headerRow.indexOf(key) === -1;
  });

  if (newKeys.length > 0) {
    // Добавляем новые заголовки справа
    sheet.getRange(1, headerRow.length + 1, 1, newKeys.length).setValues([newKeys]);
    headerRow = headerRow.concat(newKeys);
  }

  // Определяем номер новой строки
  var newRowNumber = sheet.getLastRow() + 1;
  var newRowValues = new Array(headerRow.length).fill("");

  // Заполняем строку значениями
  for (var i = 0; i < headerRow.length; i++) {
    var columnName = headerRow[i];
    if (rowData.hasOwnProperty(columnName)) {
      newRowValues[i] = rowData[columnName];
    }
  }

  // Записываем в таблицу
  sheet.getRange(newRowNumber, 1, 1, newRowValues.length).setValues([newRowValues]);
  Logger.log("✅ Row " + newRowNumber + " added: " + JSON.stringify(rowData));
}

function appendDataToColumn(sheet, columnName, value) {
  // Получаем заголовки (первую строку)
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnNumber = -1;

  // Ищем столбец по имени
  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === columnName) {
      columnNumber = i + 1; // Нумерация столбцов с 1
      break;
    }
  }

  if (columnNumber === -1) {
    Logger.log("Столбец с названием '" + columnName + "' не найден.");
    return;
  }

  // Находим последнюю непустую строку в нужном столбце
  var lastRow = sheet.getLastRow();
  var valuesInColumn = sheet.getRange(1, columnNumber, lastRow, 1).getValues();
  
  // Ищем первую пустую ячейку снизу
  var targetRow = lastRow + 1;
  for (var r = lastRow; r >= 1; r--) {
    if (valuesInColumn[r - 1][0] !== "") {
      targetRow = r + 1;
      break;
    }
  }

  // Записываем значение
  sheet.getRange(targetRow, columnNumber).setValue(value);
  Logger.log("Значение '" + value + "' добавлено в столбец '" + columnName + "' на строку " + targetRow);
}

function test_gs() {
  getTheBestName('QA-инженер (тестировщик)');
  getTheBestName('Системный аналитик');
}

