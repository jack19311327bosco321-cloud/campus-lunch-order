// ============================
// 校園訂餐系統 - Google Apps Script (v2)
// 功能：接收訂單 → 按學校分Sheet → 每天一行
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

    // 每天一行
    for (var i = 0; i < orders.length; i++) {
      var o = orders[i];
      var mealText = '';
      if (o.mealCode) {
        mealText = o.mealCode;
      } else {
        mealText = o.meal.substring(0, 4);
      }

      var vegText = '';
      if (o.hasVeg && o.vegName) {
        vegText = o.vegName;
      }

      sheet.appendRow([
        new Date(),
        name,
        className,
        studentId,
        o.day || '',
        mealText,
        vegText
      ]);
    }

    // 自動調整欄寬
    for (var c = 1; c <= 7; c++) {
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
  var safeName = schoolName.replace(/[^a-zA-Z0-9\u4e00-\u9fff]/g, '').substring(0, 31);

  var sheet = ss.getSheetByName(safeName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(safeName);

  var headers = ['時間', '姓名', '班別', '學號', '日期', '餐點', '時菜'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a90d9')
    .setFontColor('#fff')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  return sheet;
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
