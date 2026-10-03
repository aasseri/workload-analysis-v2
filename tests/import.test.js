/**
 * import.test.js — اختبارات تحويل النص الملصوق من Excel وWord إلى مهام.
 */
(function () {
  'use strict';
  var WL = window.WL, cfg = WL.config, TI = WL.taskImport;
  var T = '\t';

  function items(text, existing) {
    var p = TI.analyze(text, cfg);
    return TI.buildItems(p, p.mapping, existing || [], cfg);
  }
  function titles(list) { return list.map(function (i) { return i.title; }).join('|'); }

  window.__runImportTests = function (check) {
    // ── التنظيف
    check('تنظيف: ترقيم عربي بشرطة', TI.cleanTitle('١- إعداد التقارير'), 'إعداد التقارير');
    check('تنظيف: نقطة Word', TI.cleanTitle('•   مراجعة الطلبات  '), 'مراجعة الطلبات');
    check('تنظيف: حرف ترقيم', TI.cleanTitle('أ- استقبال المعاملات'), 'استقبال المعاملات');
    check('تنظيف: رقم بين قوسين', TI.cleanTitle('(3) أرشفة الملفات'), 'أرشفة الملفات');
    check('تنظيف: رقم جزء من النص يبقى', TI.cleanTitle('5 تقارير أسبوعية للإدارة'), '5 تقارير أسبوعية للإدارة');
    check('تنظيف: «القيام بـ» تبقى كما هي', TI.cleanTitle('القيام بإعداد التقرير.'), 'القيام بإعداد التقرير.');

    // ── Word: أسطر مرقمة ونقاط وأسطر فارغة
    var word = '1.' + T + 'استقبال طلبات الموظفين\n\n2.' + T + 'مراجعة الطلبات وتدقيقها\n•' + T + 'أرشفة الملفات\nإعداد التقرير الشهري\n';
    var w = items(word);
    check('Word: عدد المهام', w.length, 4);
    check('Word: النصوص', titles(w), 'استقبال طلبات الموظفين|مراجعة الطلبات وتدقيقها|أرشفة الملفات|إعداد التقرير الشهري');
    check('Word: الحقول الكمية فارغة', w.every(function (i) { return !i.frequencyKey && !i.repetitions && !i.durationMinutes; }), true);
    check('Word: الكل محدد', w.every(function (i) { return i.selected; }), true);

    var withHeading = items('المهام الرئيسية:\n- متابعة البريد الوارد\n- إعداد المحاضر');
    check('Word: سطر العنوان غير محدد', withHeading[0].heading + '/' + withHeading[0].selected, 'true/false');

    // ── Excel: عمود واحد بعنوان
    var col = items('المهام\nإدخال البيانات في النظام\nتحديث السجلات');
    check('Excel عمود: العنوان غير محدد', col[0].title + '/' + col[0].selected, 'المهام/false');
    check('Excel عمود: المهام', col.filter(function (i) { return i.selected; }).length, 2);

    // ── Excel: جدول بعناوين وأعمدة
    var table = 'م' + T + 'المهمة' + T + 'التردد' + T + 'التكرار' + T + 'المدة بالدقائق\n' +
      '1' + T + 'استقبال الطلبات' + T + 'يوميًا' + T + '20' + T + '10\n' +
      '2' + T + 'إعداد تقرير الإجازات' + T + 'شهري' + T + '3' + T + '180\n' +
      '3' + T + 'تحديث الهيكل' + T + 'كل فترة' + T + 'غير معروف' + T + '0\n';
    var pt = TI.analyze(table, cfg);
    check('جدول: اكتشاف العناوين', pt.hasHeader, true);
    check('جدول: استبعاد عمود «م»', pt.columns.map(function (c) { return c.name; }).join('|'), 'المهمة|التردد|التكرار|المدة بالدقائق');
    check('جدول: ربط الأعمدة تلقائيًا', [pt.mapping.title, pt.mapping.frequency, pt.mapping.repetitions, pt.mapping.duration].join(','), '1,2,3,4');
    var ti = TI.buildItems(pt, pt.mapping, [], cfg);
    check('جدول: عدد المهام', ti.length, 3);
    check('جدول: «يوميًا» ← يومي', ti[0].frequencyKey + '/' + ti[0].repetitions + '/' + ti[0].durationMinutes, 'daily/20/10');
    check('جدول: شهري', ti[1].frequencyKey, 'monthly');
    check('جدول: قيم غير صالحة تبقى فارغة', ti[2].frequencyKey + '|' + ti[2].repetitions + '|' + ti[2].durationMinutes, '||');

    // عمود المسمى الفعلي في جدول بعناوين
    var withPos = 'المهمة' + T + 'المسمى الفعلي' + T + 'التردد\n' +
      'تدقيق المستندات المالية' + T + 'مدقق   حسابات' + T + 'يومي\n' +
      'إعداد محاضر الاجتماعات' + T + 'مساعد إداري' + T + 'أسبوعي';
    var pp = TI.analyze(withPos, cfg);
    check('جدول: ربط عمود المسمى الفعلي', pp.mapping.position + '/' + pp.mapping.title, '1/0');
    var pi = TI.buildItems(pp, pp.mapping, [], cfg);
    check('جدول: نقل المسمى (مع حذف المسافات الزائدة)', pi[0].positionTitle + '|' + pi[1].positionTitle, 'مدقق حسابات|مساعد إداري');
    // عمود القسم / الوحدة
    var withSec = 'المهمة' + T + 'القسم / الوحدة' + T + 'المسمى الفعلي' + T + 'التردد\n' +
      'تدقيق المستندات المالية' + T + 'وحدة  الرواتب' + T + 'مدقق حسابات' + T + 'يومي';
    var sp = TI.analyze(withSec, cfg);
    check('جدول: ربط عمود القسم والمسمى', sp.mapping.section + '/' + sp.mapping.position, '1/2');
    check('جدول: نقل اسم القسم كما هو', TI.buildItems(sp, sp.mapping, [], cfg)[0].sectionName, 'وحدة الرواتب');
    check('قائمة Word: المسمى فارغ', items('مراجعة الطلبات الواردة')[0].positionTitle, '');

    // جدول بلا عناوين: يُكتشف عمود المهمة (الأطول) وعمود التردد (من قيمه)
    var noHead = 'مراجعة ملفات الموظفين وتحديثها' + T + 'أسبوعي\nإعداد مسير الرواتب ومطابقته' + T + 'شهري';
    var pn = TI.analyze(noHead, cfg);
    var ni = TI.buildItems(pn, pn.mapping, [], cfg);
    check('بلا عناوين: المهمة والتردد', ni[0].title + '/' + ni[0].frequencyKey + '/' + ni[1].frequencyKey, 'مراجعة ملفات الموظفين وتحديثها/weekly/monthly');

    // خلية فيها سطر جديد (Excel يضعها بين علامتي تنصيص)
    var quoted = '"مراجعة العقود' + '\n' + 'وتدقيقها"' + T + 'شهري\nأرشفة الملفات' + T + 'أسبوعي';
    var qi = items(quoted);
    check('Excel: خلية متعددة الأسطر', qi.length + '|' + qi[0].title, '2|مراجعة العقود وتدقيقها');

    // ── المكرر
    var dup = items('أرشفة الملفات\nإعداد التقارير\nأرشفة  الملفات', [{ title: 'إعداد التقارير' }]);
    check('مكرر: مع مهمة موجودة', dup[1].duplicate + '/' + dup[1].selected, 'existing/false');
    check('مكرر: داخل النص الملصوق', dup[2].duplicate + '/' + dup[2].selected, 'pasted/false');
    check('مكرر: الأصل محدد', dup[0].selected, true);

    // ── مطابقة التردد
    check('تردد: أسبوعيا بلا همزة', TI.matchFrequency('اسبوعيا', cfg), 'weekly');
    check('تردد: نصف سنوي', TI.matchFrequency('نصف سنوي', cfg), 'semiannual');
    check('تردد: قيمة غير معتمدة لا تُخمَّن', TI.matchFrequency('فصلي', cfg), '');
    check('نص فارغ', TI.analyze('   \n  ', cfg).rows.length, 0);
  };
})();
