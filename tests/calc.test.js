/**
 * calc.test.js — اختبارات محرك الحساب والتحقق، ثم توليد ملفات Excel للاختبار في Excel الحقيقي.
 * تُشغَّل من tests/run-tests.html عبر خادم الاختبار (tests/test_server.py).
 */
(function () {
  'use strict';
  var WL = window.WL, cfg = WL.config, U = WL.utils, M = WL.model, V = WL.validation, Calc = WL.calc;
  var results = [];

  function check(name, actual, expected, tol) {
    var ok = typeof expected === 'number'
      ? Math.abs(actual - expected) <= (tol || 1e-9)
      : actual === expected;
    results.push({ name: name, ok: ok, actual: actual, expected: expected });
  }

  function project(actual, tasks) {
    var p = M.createProject();
    p.org.entity = 'جهة اختبار';
    p.org.actualCount = String(actual);
    tasks.forEach(function (t) {
      M.addTask(p, { title: t[0], frequencyKey: t[1], repetitions: String(t[2]), durationMinutes: String(t[3]) });
    });
    return p;
  }

  // ── 1. الترددات
  var expectedFreq = { daily: 240, weekly: 48, semimonthly: 24, monthly: 12, quarterly: 4, semiannual: 2, annual: 1 };
  Object.keys(expectedFreq).forEach(function (k) {
    check('التردد ' + k, Calc.findFrequency(cfg, k).perYear, expectedFreq[k]);
  });

  // ── 2. ساعات المهمة (مثال المواصفات: يومي × 5 × 10 دقائق = 200)
  check('مثال المواصفات 240×5×10÷60', Calc.taskHours({ frequencyKey: 'daily', repetitions: '5', durationMinutes: '10' }, cfg), 200);
  check('أرقام عربية ٥ و ١٠', Calc.taskHours({ frequencyKey: 'daily', repetitions: '٥', durationMinutes: '١٠' }, cfg), 200);
  check('مهمة ناقصة = null', Calc.taskHours({ frequencyKey: '', repetitions: '5', durationMinutes: '10' }, cfg), null);

  // ── 3. البيانات التجريبية (العدد الفعلي 6)
  var sample = WL.sampleProject();
  var a = Calc.analyze(sample, cfg);
  check('إجمالي الساعات للبيانات التجريبية', a.totalHours, 5224);
  check('الاحتياج الدقيق', a.exactNeed, 5224 / 1023);
  check('الاحتياج المحسوب (5.107 ← 5)', a.calcNeed, 5);
  check('الفجوة = العدد الفعلي 6 − الاحتياج 5 = +1', a.gap, 1);
  check('الحالة', a.status, 'فائض');
  check('مجموع النسب = 100%', a.tasks.reduce(function (s, r) { return s + r.share; }, 0), 1, 1e-12);
  check('عدد الترددات المستخدمة = 7', a.byFrequency.filter(function (f) { return f.count > 0; }).length, 7);

  // ── 3ب. الاحتياج حسب المسمى الفعلي
  var pos = {};
  a.byPosition.forEach(function (g) { pos[g.title] = g; });
  check('مسمى: مساعد إداري 2200 ساعة ← 2', pos['مساعد إداري'].hours + '/' + pos['مساعد إداري'].need, '2200/2');
  check('مسمى: أخصائي موارد بشرية 2172 ساعة ← 2', pos['أخصائي موارد بشرية'].hours + '/' + pos['أخصائي موارد بشرية'].need, '2172/2');
  check('مسمى: محاسب رواتب 852 ساعة ← 1', pos['محاسب رواتب'].hours + '/' + pos['محاسب رواتب'].need, '852/1');
  check('مسمى: المجموع = الاحتياج الإجمالي في البيانات التجريبية', a.positionsNeedSum, 5);

  // أمثلة المستخدم: 1000 ← 1، 1500 ← 1، 1600 ← 2، 2000 ← 2، 3000 ← 3
  function hoursProject(list) {
    var p = M.createProject(); p.org.entity = 'س'; p.org.actualCount = '1';
    M.addSection(p, 'قسم واحد');   // قسم واحد: يُربط تلقائيًا بكل المهام
    list.forEach(function (x) { M.addTask(p, { title: 'مهمة ' + x[0], positionTitle: x[0], frequencyKey: 'annual', repetitions: '1', durationMinutes: String(x[1] * 60) }); });
    return Calc.analyze(p, cfg);
  }
  var ex = hoursProject([['مدقق أ', 1000], ['مدقق ب', 1500], ['مدقق ج', 1600], ['مدقق د', 2000], ['مساعد إداري', 3000]]);
  check('أمثلة المستخدم (1000/1500/1600/2000/3000)', ex.byPosition.map(function (g) { return g.need; }).join(','), '1,1,2,2,3');
  var mix = hoursProject([['أ', 1432.2], ['ب', 1432.2]]); // 1.4 + 1.4
  check('اختلاف التقريب: مجموع المسميات 2 والإجمالي 3', mix.positionsNeedSum + '/' + mix.calcNeed, '2/3');
  var spaced = hoursProject([['مساعد   إداري ', 500], ['مساعد إداري', 500]]);
  check('المسافات الزائدة لا تُنشئ مسمى جديدًا', spaced.byPosition.length + '/' + spaced.byPosition[0].hours, '1/1000');

  // ── 3ج. الأقسام / الوحدات (الإصدار 2): الاحتياج لكل قسم × مسمى
  check('عدد الأقسام في البيانات التجريبية', sample.org.sections.length, 3);
  var bs = a.bySection.map(function (sec) {
    return sec.name + ':' + sec.positions.map(function (g) { return g.title + '=' + g.hours + '/' + g.need; }).join(',') + ':' + sec.needSum;
  }).join(' | ');
  check('قسم × مسمى للبيانات التجريبية', bs,
    'وحدة الخدمات الوظيفية:مساعد إداري=2200/2,أخصائي موارد بشرية=912/1:3 | وحدة الرواتب:محاسب رواتب=852/1:1 | وحدة التخطيط والتطوير:أخصائي موارد بشرية=1260/1:1');
  check('اسم القسم في صف المهمة', a.tasks[0].sectionName, 'وحدة الخدمات الوظيفية');

  // نفس المسمى في قسمين يُقرَّب في كل قسم على حدة
  var ps = M.createProject(); ps.org.entity = 'س'; ps.org.actualCount = '1';
  var sA = M.addSection(ps, 'قسم أ').section.id, sB = M.addSection(ps, 'قسم ب').section.id;
  M.addTask(ps, { title: 'م1', sectionId: sA, positionTitle: 'مدقق', frequencyKey: 'annual', repetitions: '1', durationMinutes: String(1432.2 * 60) });
  M.addTask(ps, { title: 'م2', sectionId: sB, positionTitle: 'مدقق', frequencyKey: 'annual', repetitions: '1', durationMinutes: String(1432.2 * 60) });
  var psa = Calc.analyze(ps, cfg);
  check('مسمى واحد في قسمين: 1 + 1 والإجمالي 3', psa.bySection.map(function (x) { return x.needSum; }).join('+') + '/' + psa.calcNeed, '1+1/3');

  // واجهة الأقسام
  var pq = M.createProject();
  check('قسم فارغ مرفوض', M.addSection(pq, '   ').error, 'اكتب اسم القسم أو الوحدة أولًا.');
  var q1 = M.addSection(pq, 'وحدة  الرواتب ').section;
  check('حذف المسافات الزائدة من اسم القسم', q1.name, 'وحدة الرواتب');
  check('قسم مكرر (همزة/تاء مربوطة) مرفوض', M.addSection(pq, 'وحده الرواتب').error, 'هذا القسم مضاف مسبقًا.');
  var tAuto = M.addTask(pq, { title: 'س' });
  check('قسم واحد: يُختار تلقائيًا للمهمة', tAuto.sectionId, q1.id);
  var q2 = M.addSection(pq, 'وحدة التوظيف').section;
  check('أكثر من قسم: لا اختيار تلقائي', M.addTask(pq, { title: 'ص' }).sectionId, '');
  check('إعادة تسمية لاسم موجود مرفوضة', M.renameSection(pq, q2.id, 'وحدة الرواتب').error, 'يوجد قسم آخر بهذا الاسم.');
  M.renameSection(pq, q1.id, 'وحدة الأجور');
  check('إعادة التسمية تنعكس على المهمة', Calc.analyze(pq, cfg).tasks[0].sectionName, 'وحدة الأجور');
  check('حذف قسم يلغي ارتباط مهامه', M.removeSection(pq, q1.id) + '/' + tAuto.sectionId, '1/');

  // ترحيل مشروع من الإصدار الأول (القسم نص واحد)
  var v1 = M.deserialize({ org: { entity: 'ج', section: 'وحدة قديمة', actualCount: '2' }, tasks: [{ id: 'x1', title: 'م', frequencyKey: 'daily' }] });
  check('ترحيل الإصدار الأول: القسم يصبح قائمة', v1.org.sections.length + '/' + v1.org.sections[0].name, '1/وحدة قديمة');
  check('ترحيل الإصدار الأول: المهام تُربط بالقسم', v1.tasks[0].sectionId, v1.org.sections[0].id);
  var dangling = M.deserialize({ org: { sections: [{ id: 's1', name: 'أ' }] }, tasks: [{ title: 'م', sectionId: 'missing' }] });
  check('قسم محذوف في الملف: يُلغى الارتباط', dangling.tasks[0].sectionId, '');

  // ── 4. التقريب (1.5 ← 2 ، 1.3 ← 1)
  check('تقريب 1.5', U.roundHalfUp(1.5, 0), 2);
  check('تقريب 1.3', U.roundHalfUp(1.3, 0), 1);
  check('تقريب 2.5', U.roundHalfUp(2.5, 0), 3);
  check('تقريب 0.4999', U.roundHalfUp(0.4999, 0), 0);
  // مثل Excel: القيمة تُقرأ بدقة 15 رقمًا، فخطأ الفاصلة العائمة في 1.5 لا يقلب التقريب
  check('تقريب 1.4999999999999998 مثل Excel', U.roundHalfUp(1.4999999999999998, 0), 2);

  // احتياج دقيق = 1.5 بالضبط → 2 → عجز 1 مع عدد فعلي 1
  var half = Calc.analyze(project(1, [['م', 'annual', 1, 1023 * 60 * 1.5]]), cfg);
  check('احتياج 1.5 يصبح 2', half.calcNeed, 2);
  check('حالة عجز', half.status, 'عجز');
  check('فجوة −1 (فعلي 1 − احتياج 2)', half.gap, -1);

  // احتياج 1.3 → 1 = العدد الفعلي → متوازن
  var bal = Calc.analyze(project(1, [['م', 'annual', 1, 1023 * 60 * 1.3]]), cfg);
  check('احتياج 1.3 يصبح 1', bal.calcNeed, 1);
  check('حالة متوازن', bal.status, 'متوازن');

  // ── 5. التحقق
  var empty = V.validateProject(M.createProject(), cfg);
  var msgs = empty.errors.map(function (e) { return e.message; });
  check('مشروع فارغ غير صالح', empty.valid, false);
  check('رسالة اسم الجهة', msgs.indexOf('يرجى إدخال اسم الجهة.') >= 0, true);
  check('رسالة العدد الفعلي', msgs.indexOf('يرجى إدخال العدد الفعلي.') >= 0, true);
  check('رسالة لا توجد مهام', msgs.indexOf('يرجى إضافة مهمة واحدة على الأقل.') >= 0, true);

  var bad = project('-2', [['', '', '0', 'abc']]);
  var badMsgs = V.validateProject(bad, cfg).errors.map(function (e) { return e.message; });
  check('عدد فعلي سالب', badMsgs.indexOf('يجب ألا يقل العدد الفعلي عن صفر.') >= 0, true);
  check('تردد مفقود', badMsgs.indexOf('المهمة رقم (1): لم يتم تحديد التردد.') >= 0, true);
  check('تكرار صفر', badMsgs.indexOf('المهمة رقم (1): يجب أن يكون عدد التكرارات أكبر من صفر.') >= 0, true);
  check('المسمى الفعلي إلزامي', badMsgs.indexOf('المهمة رقم (1): يرجى إدخال المسمى الفعلي.') >= 0, true);
  check('القسم إلزامي لكل مهمة', badMsgs.indexOf('المهمة رقم (1): يرجى اختيار القسم / الوحدة.') >= 0, true);
  check('قسم واحد على الأقل في البيانات الأساسية', badMsgs.indexOf('يرجى إضافة قسم أو وحدة واحدة على الأقل.') >= 0, true);
  check('مدة غير رقمية', badMsgs.indexOf('المهمة رقم (1): المدة بالدقائق يجب أن تكون رقمًا.') >= 0, true);
  check('لا رسائل تقنية', badMsgs.join(' ').match(/NaN|undefined|null/) === null, true);
  check('البيانات التجريبية صالحة', V.validateProject(sample, cfg).valid, true);

  // ── 6. اسم الملف
  var p2 = M.createProject(); p2.org.entity = 'إدارة: "الموارد/البشرية"?';
  check('تنظيف اسم الملف', /[\\/:*?"<>|]/.test(WL.excel.buildFileName(p2, cfg)), false);

  // ── 7. نموذج البيانات
  var p3 = WL.sampleProject();
  var firstId = p3.tasks[0].id;
  M.moveTask(p3, firstId, 2);
  check('نقل مهمة', p3.tasks[2].id, firstId);
  var copy = M.duplicateTask(p3, firstId);
  check('نسخ مهمة بعد الأصل', p3.tasks[3].id, copy.id);
  check('الحفظ والاسترجاع', M.serialize(M.deserialize(M.serialize(p3))), M.serialize(p3));

  window.__calcResults = results;
})();
