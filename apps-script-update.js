// 校園訂餐系統 - Google Apps Script (v3)
// 一個學生一行，Mon-Fri 合併顯示

var NOTIFY_EMAIL = 'jack19311327bosco321@gmail.com';
var SHEET_ID = '1K19uUQ6oDxY2Zc3qdDyYASKyss6NJpgfIGO0A4t5i7U';

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

    var sheet = getOrCreateSchoolSheet(school);

    // 建立每天欄位：代號 + 時菜
    var mon = '', tue = '', wed = '', thu = '', fri = '';
    var totalPrice = 0;
    for (var i = 0; i < orders.length; i++) {
      var o = orders[i];
      var mealText = o.mealCode || '';
      if (o.hasVeg && o.vegName) mealText += '+' + o.vegName;
      if (o.day === '星期一') mon = mealText;
      else if (o.day === '星期二') tue = mealText;
      else if (o.day === '星期三') wed = mealText;
      else if (o.day === '星期四') thu = mealText;
      else if (o.day === '星期五') fri = mealText;
      totalPrice += (o.price || 0);
    }

    sheet.appendRow([
      new Date(), name, className, studentId,
      mon, tue, wed, thu, fri, totalPrice
    ]);

    for (var c = 1; c <= 10; c++) sheet.autoResizeColumn(c);
    sendEmailNotificationV2(data, orders, school);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSchoolSheet(schoolName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var safeName = schoolName.replace(/[^a-zA-Z0-9\u4e00-\u9fff]/g, '').substring(0, 31);

  var sheet = ss.getSheetByName(safeName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(safeName);
  var headers = ['時間', '姓名', '班別', '學號', '星期一', '星期二', '星期三', '星期四', '星期五', '總價'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#4a90d9')
    .setFontColor('#fff').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }
  return sheet;
}

function sendEmailNotificationV2(data, orders, school) {
  var subject = '【新訂單】' + school + ' - ' + data.name;
  var body = '收到新訂單！\n\n學校：' + school + '\n姓名：' + data.name +
    '\n班別：' + data['class'] + '\n學號：' + data.studentId + '\n─────────────\n';
  var totalPrice = 0;
  for (var i = 0; i < orders.length; i++) {
    var o = orders[i];
    var line = o.day + '：' + (o.mealCode ? '[' + o.mealCode + '] ' : '') + o.meal;
    if (o.hasVeg) line += ' + ' + (o.vegName || '時菜');
    line += ' ($' + o.price + ')';
    body += line + '\n';
    totalPrice += o.price;
  }
  body += '─────────────\n總價：$' + totalPrice + '\n\n時間：' +
    new Date().toLocaleString('zh-HK') + '\n\n— 校園訂餐系統自動通知';
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, body: body });
}
