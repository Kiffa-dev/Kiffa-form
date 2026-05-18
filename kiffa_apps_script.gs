// =====================================================
// كِفّة — Google Apps Script
// انسخ هذا الكود كاملاً في محرر Apps Script
// =====================================================

var SHEET_NAME = "ردود الموظفين";  // اسم تبويب الشيت

function doPost(e) {
  try {
    var sheet = getOrCreateSheet();
    var data  = JSON.parse(e.postData.contents);

    // أضف رأس الأعمدة إذا كان الشيت فارغاً
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "التاريخ والوقت",
        "الاسم الكامل",
        "رقم الجوال",
        "الفرع",
        "TikTok",
        "Instagram",
        "Snapchat",
        "المسموح به (المختار)",
        "الإقرار",
        "التوقيع بالاسم"
      ]);
      // تنسيق صف الرأس
      var header = sheet.getRange(1, 1, 1, 10);
      header.setBackground("#5cc4b9");
      header.setFontColor("#ffffff");
      header.setFontWeight("bold");
      header.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // أضف بيانات الموظف
    sheet.appendRow([
      new Date().toLocaleString("ar-SA", {timeZone: "Asia/Riyadh"}),
      data.fname     || "",
      data.phone     || "",
      data.branch    || "",
      data.tiktok    || "—",
      data.insta     || "—",
      data.snap      || "—",
      data.allowed   || "لا شيء محدد",
      data.pledge    || "",
      data.sig       || ""
    ]);

    // ضبط عرض الأعمدة تلقائياً
    sheet.autoResizeColumns(1, 10);

    return ContentService
      .createTextOutput(JSON.stringify({status: "ok"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

// دالة اختبار — شغّلها مرة للتأكد أن كل شيء يعمل
function testPost() {
  var fakeData = {
    postData: {
      contents: JSON.stringify({
        fname:   "أحمد محمد",
        phone:   "0512345678",
        branch:  "فرع الرياض — العليا",
        tiktok:  "@ahmed",
        insta:   "@ahmed.ksa",
        snap:    "—",
        allowed: "المنتجات، أجواء الفرع",
        pledge:  "أقر وأتعهد",
        sig:     "أحمد محمد"
      })
    }
  };
  Logger.log(doPost(fakeData));
}
