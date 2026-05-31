// ================================================================
// CleanOps — Apps Script (Updated for Full Form Submission)
// Replace your ENTIRE existing script with this.
// Deploy → New Deployment → Web App → Anyone → Deploy
// ================================================================

function doGet(e) {

  // ── FORM SUBMISSION ──────────────────────────────────────────
  if (e && e.parameter && e.parameter.action === 'submit') {
    return handleSubmit(e.parameter);
  }

  // ── DASHBOARD DATA (existing) ─────────────────────────────────
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonOut({data:[],headers:[],total:0});
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var v = data[i][j];
      row[headers[j]] = (v instanceof Date) ? v.toISOString() : v;
    }
    rows.push(row);
  }
  return jsonOut({data:rows, headers:headers, total:rows.length});
}

// ── SUBMIT: append row matching exact column order ────────────
function handleSubmit(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  // Auto-create headers on first use
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'طابع زمني','النتيجة','اسم الموقع','اسم المشرف',
      'تم تنظيف الأرضيات','تم تعقيم الـ WC','تم تنظيف الأحواض والمرايات',
      'تم تنظيف مقابض الأبواب ونقاط اللمس','تم تفريغ سلال المهملات',
      'تم تغيير أكياس القمامة','تم التأكد من توفر الصابون والمناديل والمعطرات',
      'هل توجد روائح غير طبيعية؟','هل توجد تسريبات مياه؟',
      'هل توجد أعطال تحتاج صيانة؟','هل تم استخدام مواد التعقيم المعتمدة؟',
      'رابط صورة بعد التنظيف','ملاحظات العامل','ملاحظات المشرف',
      'تقييم الحالة العامة','هل تم تنفيذ المهمة بالكامل؟',
      'في حالة عدم التنفيذ، اذكر السبب',' عدد العمال في القسم ؟ [الصف 1]',
      'التزام العمال بزي الشركه ؟','تم تنظيف الحوائط الشبابيك ؟',
      'تم تنظيف الاسنسير و السلم ؟','تم تنظيف الغرف و المكاتب ؟',
      'غسيل و فحص بساكت المخلفات ؟',' عدد العمال في القسم ؟',
      'العمود 28','رقم الدور ؟ [الصف 1]'
    ]);
  }

  // Append row in exact column order
  sheet.appendRow([
    p.timestamp   || new Date().toLocaleString(),
    p.result      || '',
    p.site        || '',
    p.supervisor  || '',
    p.floors      || 'FALSE',
    p.wc          || 'FALSE',
    p.basins      || 'FALSE',
    p.handles     || 'FALSE',
    p.bins        || 'FALSE',
    p.bags        || 'FALSE',
    p.supplies    || 'FALSE',
    p.odors       || '—',
    p.leaks       || '—',
    p.failures    || '—',
    p.sanitizers  || 'FALSE',
    p.photo       || '',
    p.workerNotes || '',
    p.supNotes    || '',
    p.rating      || '',
    p.completed   || '',
    p.reason      || '',
    p.workers     || '',
    p.uniform     || 'FALSE',
    p.walls       || 'FALSE',
    p.elevator    || 'FALSE',
    p.rooms       || 'FALSE',
    p.baskets     || 'FALSE',
    p.workers2    || '',
    p.col28       || '',
    p.floor       || ''
  ]);

  return jsonOut({success: true});
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
