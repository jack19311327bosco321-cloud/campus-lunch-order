// ============================
// 校園訂餐系統 - Google Apps Script (v2)
// 功能：接收訂單 → 按學校分Sheet → 每個學生一行(Mon-Fri)
// ============================

var NOTIFY_EMAIL = 'jack19311327bosco321@gmail.com';

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var school = data.school || '未知學校';
    var name = data.name || '';
    var className = data['class'] || '';
    var studentId = data.studentId || '';
    var orders = data.orders || [];

    // 取得或建立該學校的 Sheet
    var sheet = getOrCreateSchoolSheet(school);

    // 找到該學生是否已有行（用「班別+學號」做唯一識別）
    var studentKey = className + '_' + studentId;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var existRow = findStudentRow(sheet, studentKey);

    // 整理每天資料：代號 + 時菜
    var dayMap = {};
    var totalPrice = 0;
    for (var i = 0; i < orders.length; i++) {
      var o = orders[i];
      var dayKey = o.day; // '星期一', '星期二', etc.
      var cellVal = '';
      if (o.mealCode) {
        cellVal = o.mealCode;
      } else {
        cellVal = o.meal.substring(0, 4);
      }
      if (o.hasVeg && o.vegName) {
        cellVal += '+' + o.vegName;
      }
      dayMap[dayKey] = cellVal;
      totalPrice += (o.price || 0);
    }

    // 欄位順序
    var rowData = [
      new Date(),       // 下單時間
      name,             // 姓名
      className,        // 班別
      studentId,        // 學號
      dayMap['星期一'] || '-',   // 星期一
      dayMap['星期二'] || '-',   // 星期二
      dayMap['星期三'] || '-',   // 星期三
      dayMap['星期四'] || '-',   // 星期四
      dayMap['星期五'] || '-',   // 星期五
      totalPrice        // 總價
    ];

    if (existRow > 0) {
      // 更新已有行
      sheet.getRange(existRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // 新增行
      sheet.appendRow(rowData);
    }

    // 自動調整欄寬
    for (var c = 1; c <= 10; c++) {
      sheet.autoResizeColumn(c);
    }

    // 發送 Email 通知
    sendEmailNotificationV2(data, orders, school);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 取得或建立學校專用 Sheet
function getOrCreateSchoolSheet(schoolName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 清理學校名稱（Sheet名不能超過31字元，不能含特殊字元）
  var safeName = schoolName.replace(/[^a-zA-Z0-9\u4e00-\u9fff]/g, '').substring(0, 31);

  var sheet = ss.getSheetByName(safeName);
  if (sheet) return sheet;

  // 建立新 Sheet
  sheet = ss.insertSheet(safeName);

  // 寫入標題列
  var headers = ['下單時間', '姓名', '班別', '學號', '星期一', '星期二', '星期三', '星期四', '星期五', '總價'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a90d9')
    .setFontColor('#fff')
    .setHorizontalAlignment('center');

  // 凍結首行
  sheet.setFrozenRows(1);

  // 刪除預設的 Sheet（如果存在且不是學校 Sheet）
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  return sheet;
}

// 用班別+學號找學生已有行
function findStudentRow(sheet, studentKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var data = sheet.getRange(2, 3, lastRow - 1, 2).getValues(); // C=班別, D=學號
  for (var r = 0; r < data.length; r++) {
    var key = data[r][0] + '_' + data[r][1];
    if (key === studentKey) {
      return r + 2; // +2 因為從第2行開始，且getValues是0-based
    }
  }
  return -1;
}

// Email 通知
function sendEmailNotificationV2(data, orders, school) {
  var subject = '【新訂單】' + school + ' - ' + data.name;

  var body = '收到新訂單！\n\n' +
    '學校：' + school + '\n' +
    '姓名：' + data.name + '\n' +
    '班別：' + data['class'] + '\n' +
    '學號：' + data.studentId + '\n' +
    '─────────────\n';

  var totalPrice = 0;
  for (var i = 0; i < orders.length; i++) {
    var o = orders[i];
    var line = o.day + '：' + (o.mealCode ? '[' + o.mealCode + '] ' : '') + o.meal;
    if (o.hasVeg) line += ' + ' + (o.vegName || '時菜');
    line += ' ($' + o.price + ')';
    body += line + '\n';
    totalPrice += o.price;
  }

  body += '─────────────\n' +
    '總價：$' + totalPrice + '\n\n' +
    '時間：' + new Date().toLocaleString('zh-HK') + '\n\n' +
    '— 校園訂餐系統自動通知';

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: subject,
    body: body
  });
}
