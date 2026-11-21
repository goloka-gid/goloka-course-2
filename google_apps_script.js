// === APPS SCRIPT ДЛЯ КУРСА 2 ===
// 
// ИНСТРУКЦИЯ ПО ПОДКЛЮЧЕНИЮ (Вариант B — отдельный скрипт, та же таблица):
// 
// 1. Открой ту же Google-таблицу, где сейчас лежат ответы Курса 1
// 2. Меню: Extensions → Apps Script
// 3. В открывшемся редакторе создай НОВЫЙ проект (File → New project)
//    или просто скопируй этот код в новый файл проекта
// 4. Назови проект, например, "Course2_Backend"
// 5. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Скопируй полученный URL вида https://script.google.com/macros/s/AKfycb.../exec
// 7. Вставь его в config.js (точнее — в app.js, замени GOOGLE_APP_SCRIPT_URL)
// 
// ВАЖНО: Существующий скрипт Курса 1 НЕ ТРОГАЕМ. Этот скрипт работает
// в той же таблице, но пишет в отдельный лист "Homeworks_Level2".
// При первой отправке скрипт автоматически создаст этот лист с нужными колонками.

// Имя листа для этого курса (можно переименовать, если хочешь)
const SHEET_NAME = "Homeworks_Level2";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    
    if (action === 'submitHomework') {
      return submitHomework(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      "success": false, 
      "error": "Unknown action: " + action
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "success": false, 
      "error": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var action = e.parameter.action;
  
  if (action === 'getProgress') {
    return getProgress(e.parameter.email);
  }
  
  // Health check endpoint
  return ContentService.createTextOutput(JSON.stringify({
    "status": "ok",
    "course": "level2",
    "sheet": SHEET_NAME
  })).setMimeType(ContentService.MimeType.JSON);
}

function submitHomework(data) {
  var email = data.email || 'Неизвестный';
  var day = data.day || 'Неизвестный день';
  var text = data.text || '';
  var date = new Date();
  
  var sheet = getSheet();
  
  // Если уже есть отправка для этого email+day — обновляем её, иначе добавляем новую
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 6);
    var values = dataRange.getValues();
    
    for (var i = 0; i < values.length; i++) {
      if (values[i][1] === email && String(values[i][2]) === String(day)) {
        // Обновляем существующую строку
        sheet.getRange(i + 2, 1, 1, 1).setValue(date);     // Колонка A: Дата
        sheet.getRange(i + 2, 4, 1, 1).setValue(text);     // Колонка D: Текст
        sheet.getRange(i + 2, 5, 1, 1).setValue("На проверке");  // Колонка E: Статус
        sheet.getRange(i + 2, 6, 1, 1).setValue("");       // Колонка F: Комментарий
        return ContentService.createTextOutput(JSON.stringify({"success": true}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  
  // Добавляем новую строку
  sheet.appendRow([date, email, day, text, "На проверке", ""]);
  return ContentService.createTextOutput(JSON.stringify({"success": true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getProgress(email) {
  var sheet = getSheet();
  var progress = {};
  
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 6);
    var values = dataRange.getValues();
    
    for (var i = 0; i < values.length; i++) {
      if (values[i][1] === email) {
        var day = String(values[i][2]);
        var status = values[i][4] || "";
        var comment = values[i][5] || "";
        
        // Если есть несколько отправок для одного дня — берём последнюю
        progress[day] = {
          status: status,
          comment: comment
        };
      }
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    progress: progress
  })).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    // Автоматически создаём лист с заголовками
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Дата", "Email", "День", "Текст задания", "Статус", "Комментарий"]);
    
    // Жирный шрифт для заголовка
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    
    // Закрепляем первую строку
    sheet.setFrozenRows(1);
    
    // Автоширина колонок
    sheet.autoResizeColumns(1, 6);
  }
  
  return sheet;
}

// === АДМИН-ФУНКЦИИ (вызывай вручную из редактора для управления) ===

// Одобрить домашку для email+day
function approveHomework(email, day, comment) {
  comment = comment || "";
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow >= 2) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 6);
    var values = dataRange.getValues();
    
    for (var i = 0; i < values.length; i++) {
      if (values[i][1] === email && String(values[i][2]) === String(day)) {
        sheet.getRange(i + 2, 5, 1, 1).setValue("Одобрено");
        sheet.getRange(i + 2, 6, 1, 1).setValue(comment);
        return "Approved: " + email + " day " + day;
      }
    }
  }
  return "Not found: " + email + " day " + day;
}

// Попросить переделать
function requestResubmit(email, day, comment) {
  comment = comment || "Попробуйте дополнить ответ.";
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow >= 2) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 6);
    var values = dataRange.getValues();
    
    for (var i = 0; i < values.length; i++) {
      if (values[i][1] === email && String(values[i][2]) === String(day)) {
        sheet.getRange(i + 2, 5, 1, 1).setValue("Переделать");
        sheet.getRange(i + 2, 6, 1, 1).setValue(comment);
        return "Resubmit requested: " + email + " day " + day;
      }
    }
  }
  return "Not found: " + email + " day " + day;
}

// Пример вызова (раскомментируй и запусти в редакторе):
// approveHomework("test@example.com", 1, "Отличная работа!");
// requestResubmit("test@example.com", 2, "Попробуй описать свои чувства подробнее");
