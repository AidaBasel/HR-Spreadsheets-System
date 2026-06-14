/**
 * وظيفة العرض الرئيسية - تبديل الصفحات بناءً على الرابط
 */
function doGet(e) {
  var page = e.parameter.page;
  if (page == 'manager') {
    return HtmlService.createTemplateFromFile('Manager').evaluate()
        .setTitle('لوحة اعتماد المدير - FINTEK')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('تقديم طلب إجازة - FINTEK')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * جلب بيانات الموظف (البحث بالرقم الوظيفي)
 */
function getEmployeeDetails(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("بيانات الموظفين"); 
  if (!sheet) return null;
  
  var data = sheet.getDataRange().getValues();
  var searchId = id.toString().trim();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === searchId) {
      return { name: data[i][1], department: data[i][2] };
    } 
  }
  return null;
}

/**
 * استقبال الطلب من الموظف وحفظه في "الطلبـات_المعلقة" وإرسال إيميل
 */
function processLeaveRequest(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("الطلبات_المعلقة");
  if (!sheet) return "خطأ: ورقة الطلبات_المعلقة غير موجودة";
  
  var fileUrl = "لا يوجد مرفق"; 
  if (data.attachment) {
    var folder = DriveApp.getFolderById("17vvJ204wHYVTe-p5OqkSd2qJcKQm0puz");
    var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(data.attachment.base64), data.attachment.mimeType, data.attachment.fileName));
    fileUrl = file.getUrl();
  }
  
  // حفظ البيانات في شيت المعلق
  sheet.appendRow([new Date(), data.empId, data.name, data.type, data.fromDate, data.toDate, fileUrl, "Pending"]);
  
  // إرسال الإيميل للمدير
  try {
    sendEmailToManager(data);
  } catch(e) {
    console.log("فشل إرسال الإيميل: " + e.message);
  }
  
  return "✅ تم إرسال طلبك بنجاح وجاري إشعار المدير.";
}

/**
 * وظيفة إرسال الإيميل للمدير
 */
function sendEmailToManager(data) {
  try {
    var managerEmail = "aidaqudar97@gmail.com"; 
    // الحصول على رابط النشر الحالي
    var appUrl = ScriptApp.getService().getUrl() + "?page=manager";
    
    var subject = "🚨 طلب إجازة جديد: " + data.name;
    var body = "مرحباً سيادة المدير،\n\n" +
               "هناك طلب إجازة جديد ينتظر موافقتك:\n" +
               "----------------------------------\n" +
               "- الموظف: " + data.name + "\n" +
               "- النوع: " + data.type + "\n" +
               "- الفترة: من " + data.fromDate + " إلى " + data.toDate + "\n" +
               "----------------------------------\n\n" +
               "للاتخاذ القرار (قبول/رفض)، يرجى الضغط على الرابط أدناه:\n" + appUrl;
    
    MailApp.sendEmail(managerEmail, subject, body);
    console.log("تم إرسال الإيميل بنجاح إلى: " + managerEmail);
  } catch (e) {
    console.error("فشل إرسال الإيميل. السبب: " + e.message);
  }
}

/**
 * جلب الطلبات المعلقة للوحة المدير
 */
function getPendingRequests() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("الطلبات_المعلقة");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] == "Pending") {
      results.push({
        rowNum: i + 1,
        empId: data[i][1],
        name: data[i][2],
        type: data[i][3],
        from: Utilities.formatDate(new Date(data[i][4]), "GMT+3", "yyyy-MM-dd"),
        to: Utilities.formatDate(new Date(data[i][5]), "GMT+3", "yyyy-MM-dd"),
        file: data[i][6]
      });
    }
  }
  return results;
}

/**
 * اعتماد المدير: نقل الطلب لـ "نظام الاجازات" وتحديث الحالة
 */
function finalizeRequest(rowNum, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = ss.getSheetByName("الطلبات_المعلقة");
  var finalSheet = ss.getSheetByName("نظام الاجازات");
  
  var data = tempSheet.getRange(rowNum, 1, 1, 7).getValues()[0];
  
  // إضافة البيانات للشيت الرئيسي
  finalSheet.appendRow([data[0], data[1], data[2], data[3], data[4], data[5], data[6], status, status]);
  
  // تحديث الحالة في المعلق ليختفي من اللوحة
  tempSheet.getRange(rowNum, 8).setValue(status);
  
  return "تم تحديث الطلب بنجاح إلى: " + status;
}
