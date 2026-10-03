/**
 * excel-export.js — محرك Excel.
 * يبني ملف XLSX حقيقيًا: البيانات المدخلة كقيم، وكل النتائج كمعادلات Excel فعلية
 * تعتمد على أسماء معرّفة (Named Ranges) وعلى ورقة الإعدادات.
 * القيمة المحسوبة في JavaScript تُخزن فقط كـ "نتيجة مخزنة" للمعادلة، ويعيد Excel الحساب عند الفتح.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  // ───────────── أدوات مساعدة ─────────────

  function f(formula, result) {
    var v = { formula: formula };
    if (result !== undefined && result !== null && !(typeof result === 'number' && !isFinite(result))) {
      v.result = result;
    }
    return v;
  }

  function sheetRef(ctx, key, addr) {
    return "'" + ctx.cfg.excel.sheetNames[key] + "'!" + addr;
  }

  function nameRange(ws, col, fromRow, toRow, name) {
    for (var r = fromRow; r <= toRow; r++) ws.getCell(col + r).addName(name);
  }

  function mergeStyled(ws, range, value, styler) {
    ws.mergeCells(range);
    var cell = ws.getCell(range.split(':')[0]);
    if (value !== undefined) cell.value = value;
    if (styler) styler(cell);
    return cell;
  }

  function needFormat(digits) {
    return digits > 0 ? '#,##0.' + new Array(digits + 1).join('0') : '#,##0';
  }

  function gapFormat(digits) {
    var n = needFormat(digits);
    return '+' + n + ';-' + n + ';0';
  }

  function blankToNull(v) {
    return v === undefined || v === null || String(v).trim() === '' ? null : v;
  }

  function estimateRowHeight(text, colWidth) {
    var len = String(text || '').length;
    var perLine = Math.max(10, Math.floor(colWidth * 1.15));
    var lines = Math.max(1, Math.ceil(len / perLine));
    return Math.max(22, lines * 16 + 6);
  }

  function titleBlock(ctx, ws, lastCol, title, withEntity) {
    var T = ctx.T;
    mergeStyled(ws, 'A1:' + lastCol + '1', title, T.style.title);
    ws.getRow(1).height = 34;
    if (withEntity) {
      var N = T.NAMES;
      mergeStyled(ws, 'A2:' + lastCol + '2',
        f('"الجهة: "&' + N.entity, 'الجهة: ' + ctx.project.org.entity),
        T.style.subtitle);
      ws.getRow(2).height = 22;
    }
  }

  function numberValidation(kind, operator, value, message) {
    return {
      type: kind, operator: operator, allowBlank: true, formulae: [value],
      showErrorMessage: true, errorStyle: 'stop',
      errorTitle: 'قيمة غير صحيحة', error: message
    };
  }

  // ───────────── ورقة الإعدادات ─────────────

  function buildSettings(ctx) {
    var ws = ctx.sheets.settings, T = ctx.T, N = T.NAMES, L = T.LAYOUT.settings, cfg = ctx.cfg;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'C', 'الإعدادات والقيم المرجعية', false);
    mergeStyled(ws, 'A2:C2', 'جميع معادلات الملف تعتمد على القيم في هذه الورقة. الخلية الملونة بالأصفر الفاتح قابلة للتعديل.', T.style.note);

    ['الإعداد', 'القيمة', 'ملاحظة'].forEach(function (h, i) {
      var c = ws.getRow(3).getCell(i + 1); c.value = h; T.style.header(c);
    });

    var r = L.annualHoursRow;
    ws.getCell('A' + r).value = 'ساعات العمل الفعلية السنوية للموظف'; T.style.label(ws.getCell('A' + r));
    var ah = ws.getCell('B' + r);
    ah.value = cfg.annualWorkHours;
    T.style.input(ah, T.NUM.integer, 'center');
    ah.addName(N.annualHours);
    ah.dataValidation = numberValidation('decimal', 'greaterThan', 0, 'ساعات العمل السنوية يجب أن تكون رقمًا أكبر من صفر.');
    ws.getCell('C' + r).value = 'يستخدم لحساب الاحتياج: إجمالي الساعات ÷ هذه القيمة';
    T.style.note(ws.getCell('C' + r));

    r = L.needDecimalsRow;
    ws.getCell('A' + r).value = 'منازل تقريب الاحتياج المحسوب'; T.style.label(ws.getCell('A' + r));
    var nd = ws.getCell('B' + r);
    nd.value = cfg.needRounding.digits;
    T.style.formula(nd, T.NUM.integer);
    nd.addName(N.needDecimals);
    ws.getCell('C' + r).value = 'تقريب لأقرب عدد صحيح: 0.5 فأعلى يُجبر للأعلى (1.5 ← 2 ، 1.3 ← 1)';
    T.style.note(ws.getCell('C' + r));

    var hr = L.freqHeaderRow;
    mergeStyled(ws, 'A' + (hr - 1) + ':C' + (hr - 1), 'جدول الترددات السنوية (قيم معتمدة)', T.style.section);

    ws.addTable({
      name: T.LAYOUT.settingsFreqTableName,
      ref: 'A' + hr,
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleLight1', showRowStripes: false },
      columns: [{ name: 'التردد', filterButton: false }, { name: 'المرات السنوية', filterButton: false }],
      rows: cfg.frequencies.map(function (fr) { return [fr.label, fr.perYear]; })
    });

    var first = hr + 1, last = hr + cfg.frequencies.length;
    [ws.getCell('A' + hr), ws.getCell('B' + hr)].forEach(T.style.header);
    for (var i = first; i <= last; i++) {
      T.style.formula(ws.getCell('A' + i), null, { align: 'right' });
      T.style.formula(ws.getCell('B' + i), T.NUM.integer);
      ws.getRow(i).height = 22;
    }
    [3, L.annualHoursRow, L.needDecimalsRow, hr - 1, hr].forEach(function (n) { ws.getRow(n).height = 24; });
    nameRange(ws, 'A', first, last, N.freqNames);
    nameRange(ws, 'B', first, last, N.freqValues);

    ctx.freqRows = { first: first, last: last };
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:C' + last, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── ورقة تحليل المهام ─────────────

  function buildTasks(ctx) {
    var ws = ctx.sheets.tasks, T = ctx.T, N = T.NAMES, L = T.LAYOUT.tasks, cfg = ctx.cfg, a = ctx.analysis;
    var cols = L.columns;
    /** حرف عمود Excel من مفتاحه في LAYOUT.tasks.columns (لا أحرف ثابتة في المعادلات). */
    function C(key) {
      for (var i = 0; i < cols.length; i++) if (cols[i].key === key) return String.fromCharCode(65 + i);
      throw new Error('UNKNOWN_COLUMN ' + key);
    }
    function idx(key) { return C(key).charCodeAt(0) - 64; }
    var LAST = C(cols[cols.length - 1].key);
    var BEFORE_HOURS = String.fromCharCode(C('hours').charCodeAt(0) - 1);

    T.setColumnWidths(ws, cols.map(function (c) { return c.width; }));
    titleBlock(ctx, ws, LAST, 'تحليل المهام', true);

    var hr = L.headerRow;
    var first = hr + 1;
    var count = a.tasks.length + cfg.excel.spareTaskRows;
    var last = hr + count;

    var rows = [];
    for (var i = 0; i < count; i++) {
      var r = first + i;
      var t = a.tasks[i];
      var fq = C('freq') + r, py = C('perYear') + r, rp = C('reps') + r, du = C('duration') + r, hh = C('hours') + r;
      var noF = 'IF(' + C('title') + r + '="","",ROW()-' + hr + ')';
      var perYearF = 'IF(' + fq + '="","",IF(COUNTIF(' + N.freqNames + ',' + fq + ')=0,"",INDEX(' + N.freqValues + ',MATCH(' + fq + ',' + N.freqNames + ',0))))';
      var hoursF = 'IF(OR(' + py + '="",' + rp + '="",' + du + '=""),"",' + py + '*' + rp + '*' + du + '/' + cfg.minutesPerHour + ')';
      var shareF = 'IF(OR(' + hh + '="",' + N.totalHours + '=0),"",' + hh + '/' + N.totalHours + ')';
      var values = {
        no: f(noF, t ? i + 1 : ''),
        title: t ? t.title : null,
        section: t ? (t.sectionName || null) : null,
        position: t ? (t.positionTitle || null) : null,
        freq: t ? t.frequencyLabel : null,
        perYear: f(perYearF, t ? t.perYear : ''),
        reps: t ? t.repetitions : null,
        duration: t ? t.durationMinutes : null,
        hours: f(hoursF, t ? t.hours : ''),
        share: f(shareF, t ? t.share : '')
      };
      rows.push(cols.map(function (c) { return values[c.key]; }));
    }

    ws.addTable({
      name: L.tableName,
      ref: 'A' + hr,
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleLight1', showRowStripes: false },
      columns: cols.map(function (c) { return { name: c.header, filterButton: true }; }),
      rows: rows
    });

    cols.forEach(function (c, i) { T.style.header(ws.getRow(hr).getCell(i + 1)); });
    ws.getRow(hr).height = 30;

    var freqError = 'يرجى اختيار التردد من القائمة المنسدلة.';
    for (var k = 0; k < count; k++) {
      var row = ws.getRow(first + k);
      var stripe = k % 2 === 1 ? T.COLORS.stripe : T.COLORS.white;
      T.style.formula(row.getCell(idx('no')), null, { fill: stripe });
      T.style.input(row.getCell(idx('title')));
      T.style.input(row.getCell(idx('section')));
      T.style.input(row.getCell(idx('position')));
      T.style.input(row.getCell(idx('freq')), null, 'center');
      T.style.formula(row.getCell(idx('perYear')), T.NUM.integer, { fill: stripe });
      T.style.input(row.getCell(idx('reps')), T.NUM.general, 'center');
      T.style.input(row.getCell(idx('duration')), T.NUM.general, 'center');
      T.style.formula(row.getCell(idx('hours')), T.NUM.hours, { fill: stripe, bold: true });
      T.style.formula(row.getCell(idx('share')), T.NUM.percent, { fill: stripe });

      row.getCell(idx('freq')).dataValidation = {
        type: 'list', allowBlank: true, formulae: [N.freqNames],
        showErrorMessage: true, errorStyle: 'stop', errorTitle: 'تردد غير صحيح', error: freqError,
        showInputMessage: true, promptTitle: 'التردد', prompt: 'اختر التردد من القائمة'
      };
      row.getCell(idx('section')).dataValidation = {
        type: 'list', allowBlank: true, formulae: [N.sectionNames],
        showErrorMessage: true, errorStyle: 'stop', errorTitle: 'قسم غير صحيح', error: 'يرجى اختيار القسم / الوحدة من القائمة (تُعرّف الأقسام في ورقة البيانات الأساسية).',
        showInputMessage: true, promptTitle: 'القسم / الوحدة', prompt: 'اختر القسم أو الوحدة من القائمة'
      };
      row.getCell(idx('reps')).dataValidation = numberValidation('decimal', 'greaterThan', 0, 'يجب أن يكون عدد التكرارات رقمًا أكبر من صفر.');
      row.getCell(idx('duration')).dataValidation = numberValidation('decimal', 'greaterThan', 0, 'يجب أن تكون المدة بالدقائق رقمًا أكبر من صفر.');

      var task = a.tasks[k];
      row.height = task ? estimateRowHeight(task.title, cols[idx('title') - 1].width) : 20;
    }

    nameRange(ws, C('title'), first, last, N.taskTitles);
    nameRange(ws, C('section'), first, last, N.taskSections);
    nameRange(ws, C('position'), first, last, N.taskPositions);
    nameRange(ws, C('freq'), first, last, N.taskFreq);
    nameRange(ws, C('hours'), first, last, N.taskHours);

    // تنبيهات داخل Excel: تردد غير معروف، أو مهمة بلا قسم أو مسمى أو تردد أو تكرار/مدة، أو قسم غير معرّف
    var fT = C('title') + first, fF = C('freq') + first;
    T.addWarningRule(ws, C('freq') + first + ':' + C('freq') + last, 'AND($' + fF + '<>"",$' + C('perYear') + first + '="")');
    T.addWarningRule(ws, C('freq') + first + ':' + C('freq') + last, 'AND($' + fT + '<>"",$' + fF + '="")');
    T.addWarningRule(ws, C('section') + first + ':' + C('section') + last, 'AND($' + fT + '<>"",' + C('section') + first + '="")');
    T.addWarningRule(ws, C('section') + first + ':' + C('section') + last, 'AND(' + C('section') + first + '<>"",COUNTIF(' + N.sectionNames + ',' + C('section') + first + ')=0)');
    T.addWarningRule(ws, C('position') + first + ':' + C('position') + last, 'AND($' + fT + '<>"",' + C('position') + first + '="")');
    T.addWarningRule(ws, C('reps') + first + ':' + C('duration') + last, 'AND($' + fT + '<>"",' + C('reps') + first + '="")');

    // الإجماليات أعلى الجدول (تبقى ظاهرة مع تجميد الأجزاء)
    var tr = L.totalRow, cr = L.countRow;
    mergeStyled(ws, 'A' + tr + ':' + BEFORE_HOURS + tr, 'إجمالي ساعات العمل السنوية', T.style.label);
    var total = ws.getCell(C('hours') + tr);
    total.value = f('SUM(' + N.taskHours + ')', a.totalHours);
    T.style.total(total, T.NUM.hours);
    total.addName(N.totalHours);
    var shareTotal = ws.getCell(C('share') + tr);
    shareTotal.value = f('SUM(' + C('share') + first + ':' + C('share') + last + ')', a.totalHours > 0 ? 1 : 0);
    T.style.total(shareTotal, T.NUM.percent);

    mergeStyled(ws, 'A' + cr + ':' + BEFORE_HOURS + cr, 'عدد المهام', T.style.label);
    var cnt = ws.getCell(C('hours') + cr);
    cnt.value = f('COUNTA(' + N.taskTitles + ')', a.taskCount);
    T.style.total(cnt, T.NUM.integer);
    cnt.addName(N.taskCount);
    T.style.total(ws.getCell(C('share') + cr));
    ws.getRow(tr).height = 24;
    ws.getRow(cr).height = 24;

    mergeStyled(ws, 'A' + L.noteRow + ':' + LAST + L.noteRow,
      'الخلايا الملونة بالأصفر الفاتح للإدخال. اختر القسم والتردد من القائمة وتُحسب بقية الأعمدة تلقائيًا. يوجد ' +
      cfg.excel.spareTaskRows + ' صفًا فارغًا جاهزًا أسفل المهام لإضافة مهام جديدة، ويتسع نطاق الطباعة تلقائيًا ليشملها.', T.style.note);
    ws.getRow(L.noteRow).height = 20;

    var paper = a.tasks.length + hr > cfg.excel.a3RowThreshold ? 'A3' : 'A4';
    T.applyPrint(ws, {
      paper: paper, landscape: true, printArea: 'A1:' + LAST + (hr + a.tasks.length),
      titleRows: hr + ':' + hr, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText
    });

    // نطاق طباعة ديناميكي: حتى آخر صف يحتوي مهمة (بدل طباعة الصفوف الاحتياطية الفارغة)
    var sn = "'" + cfg.excel.sheetNames.tasks + "'!";
    var titles = sn + '$' + C('title') + '$' + first + ':$' + C('title') + '$' + last;
    ctx.dynamicPrintAreas.push({
      sheetKey: 'tasks',
      formula: 'OFFSET(' + sn + '$A$1,0,0,IF(COUNTA(' + titles + ')=0,' + hr + ',LOOKUP(2,1/(' + titles + '<>""),ROW(' + titles + '))),' + cols.length + ')'
    });
    ctx.taskRows = { first: first, last: last };
  }

  // ───────────── ورقة الاحتياج حسب القسم / الوحدة × المسمى الفعلي ─────────────

  /**
   * لكل (قسم، مسمى): عدد المهام وإجمالي الساعات (SUMPRODUCT على القسم والمسمى معًا)،
   * والاحتياج قبل التقريب وبعده (نفس قاعدة الاحتياج العام)، مع مجموع لكل قسم.
   * صفوف احتياطية فارغة لإضافة (قسم، مسمى) يدويًا، وسطر تحقق يكشف ساعات مهام غير مدرجة في الجدول.
   */
  function buildPositions(ctx) {
    var ws = ctx.sheets.positions, T = ctx.T, N = T.NAMES, L = T.LAYOUT.positions, cfg = ctx.cfg, a = ctx.analysis;
    var digits = cfg.needRounding.digits;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'G', 'الاحتياج حسب القسم / الوحدة والمسمى الفعلي', true);
    mergeStyled(ws, 'A3:G3', 'الاحتياج لكل مسمى داخل كل قسم = إجمالي ساعات مهامه ÷ ساعات العمل الفعلية السنوية، مقربًا لأقرب عدد صحيح. ' +
      'الصفوف الصفراء الفارغة في الأسفل لإضافة قسم ومسمى يدويًا.', T.style.note);
    ws.getRow(3).height = 20;

    var hr = L.headerRow;
    ['م', 'القسم / الوحدة', 'المسمى الفعلي', 'عدد المهام', 'إجمالي الساعات', 'الاحتياج قبل التقريب', 'الاحتياج'].forEach(function (h, i) {
      var c = ws.getRow(hr).getCell(i + 1); c.value = h; T.style.header(c);
    });
    ws.getRow(hr).height = 30;

    /** صف تفصيلي (قسم، مسمى): القسم والمسمى للإدخال، وبقية الأعمدة معادلات. */
    function detailRow(r, no, sec, g, stripe) {
      var row = ws.getRow(r);
      var key = 'B' + r + '="",C' + r + '=""';
      row.getCell(1).value = no;
      row.getCell(2).value = sec;
      row.getCell(3).value = g ? g.title : null;
      row.getCell(4).value = f('IF(OR(' + key + '),"",SUMPRODUCT(--(' + N.taskSections + '=B' + r + '),--(' + N.taskPositions + '=C' + r + ')))', g ? g.count : '');
      row.getCell(5).value = f('IF(OR(' + key + '),"",SUMPRODUCT(--(' + N.taskSections + '=B' + r + '),--(' + N.taskPositions + '=C' + r + '),' + N.taskHours + '))', g ? g.hours : '');
      row.getCell(6).value = f('IF(OR(E' + r + '="",' + N.annualHours + '=0),"",E' + r + '/' + N.annualHours + ')', g ? g.exactNeed : '');
      row.getCell(7).value = f('IF(F' + r + '="","",ROUND(F' + r + ',' + N.needDecimals + '))', g ? g.need : '');
      T.style.formula(row.getCell(1), null, { fill: stripe });
      T.style.input(row.getCell(2));
      T.style.input(row.getCell(3));
      T.style.formula(row.getCell(4), T.NUM.integer, { fill: stripe });
      T.style.formula(row.getCell(5), T.NUM.hours, { fill: stripe });
      T.style.formula(row.getCell(6), T.NUM.decimal, { fill: stripe });
      T.style.formula(row.getCell(7), needFormat(digits), { fill: stripe, bold: true, size: 12 });
      row.getCell(2).dataValidation = {
        type: 'list', allowBlank: true, formulae: [N.sectionNames],
        showErrorMessage: true, errorStyle: 'stop', errorTitle: 'قسم غير صحيح', error: 'يرجى اختيار القسم / الوحدة من القائمة.'
      };
      row.height = 24;
    }

    var r = hr + 1, no = 0;
    var subtotalRows = [];
    a.bySection.filter(function (sec) { return sec.positions.length; }).forEach(function (sec) {
      var from = r;
      sec.positions.forEach(function (g, i) {
        no += 1;
        detailRow(r, no, sec.name, g, i % 2 === 1 ? T.COLORS.stripe : T.COLORS.white);
        r += 1;
      });
      var to = r - 1;
      var sRow = ws.getRow(r);
      mergeStyled(ws, 'A' + r + ':C' + r, 'مجموع ' + sec.name, T.style.label);
      sRow.getCell(4).value = f('SUM(D' + from + ':D' + to + ')', sec.count);
      sRow.getCell(5).value = f('SUM(E' + from + ':E' + to + ')', sec.hours);
      sRow.getCell(7).value = f('SUM(G' + from + ':G' + to + ')', sec.needSum);
      T.style.label(sRow.getCell(4)); sRow.getCell(4).numFmt = T.NUM.integer;
      T.style.label(sRow.getCell(5)); sRow.getCell(5).numFmt = T.NUM.hours;
      T.style.label(sRow.getCell(6));
      T.style.label(sRow.getCell(7)); sRow.getCell(7).numFmt = needFormat(digits);
      [4, 5, 7].forEach(function (c) { sRow.getCell(c).alignment = { horizontal: 'center', vertical: 'middle' }; });
      sRow.height = 24;
      subtotalRows.push(r);
      r += 1;
    });

    // صفوف احتياطية: تُحتسب في الإجمالي مباشرة (ليست ضمن مجموع أي قسم)
    var spareFirst = r;
    for (var k = 0; k < L.spareRows; k++) {
      detailRow(r, null, null, null, k % 2 === 1 ? T.COLORS.stripe : T.COLORS.white);
      r += 1;
    }
    var spareLast = r - 1;

    /** الإجمالي = مجاميع الأقسام + الصفوف الاحتياطية (دون احتساب الصفوف التفصيلية مرتين). */
    function grand(col) {
      return subtotalRows.map(function (x) { return col + x; }).concat(['SUM(' + col + spareFirst + ':' + col + spareLast + ')']).join('+');
    }
    var secs = a.bySection.filter(function (sec) { return sec.positions.length; });
    var tr = r;
    var tRow = ws.getRow(tr);
    mergeStyled(ws, 'A' + tr + ':C' + tr, 'الإجمالي', T.style.total);
    tRow.getCell(4).value = f(grand('D'), secs.reduce(function (s, x) { return s + x.count; }, 0));
    tRow.getCell(5).value = f(grand('E'), secs.reduce(function (s, x) { return s + x.hours; }, 0));
    tRow.getCell(7).value = f(grand('G'), a.positionsNeedSum);
    T.style.total(tRow.getCell(4), T.NUM.integer);
    T.style.total(tRow.getCell(5), T.NUM.hours);
    T.style.total(tRow.getCell(6));
    T.style.total(tRow.getCell(7), needFormat(digits));
    tRow.height = 26;

    // مقارنة مع الاحتياج الإجمالي، وتحقق من الساعات غير الموزعة على (قسم، مسمى) مدرج
    var sumNeedAddr = 'G' + tr, sumHoursAddr = 'E' + tr;
    var checks = [
      { label: 'الاحتياج الإجمالي للجهة (إجمالي الساعات ÷ ساعات العمل السنوية)', v: f(N.calcNeed, a.calcNeed), fmt: needFormat(digits),
        note: f('IF(' + sumNeedAddr + '=' + N.calcNeed + ',"مطابق لمجموع الأقسام والمسميات","يختلف عن المجموع بسبب تقريب كل مسمى داخل كل قسم على حدة")',
          a.positionsNeedSum === a.calcNeed ? 'مطابق لمجموع الأقسام والمسميات' : 'يختلف عن المجموع بسبب تقريب كل مسمى داخل كل قسم على حدة') },
      { label: 'ساعات مهام لقسم أو مسمى غير مدرج في الجدول أعلاه', v: f(N.totalHours + '-' + sumHoursAddr, 0), fmt: T.NUM.hours, warn: true,
        note: f('IF(ROUND(' + N.totalHours + '-' + sumHoursAddr + ',2)=0,"كل الساعات موزعة على الأقسام والمسميات","أضف القسم والمسمى الناقص في صف فارغ أعلاه")',
          'كل الساعات موزعة على الأقسام والمسميات') }
    ];
    checks.forEach(function (c, i) {
      var rr = tr + 2 + i;
      mergeStyled(ws, 'A' + rr + ':D' + rr, c.label, T.style.label);
      var vc = ws.getCell('E' + rr);
      vc.value = c.v;
      T.style.formula(vc, c.fmt, { bold: true });
      var nc = mergeStyled(ws, 'F' + rr + ':G' + rr, c.note);
      T.style.formula(nc, null, { align: 'right' });
      nc.font = T.font({ size: 10, color: T.COLORS.muted });
      if (c.warn) {
        ws.addConditionalFormatting({ ref: 'E' + rr, rules: [{ type: 'expression', priority: 1, formulae: ['ROUND(E' + rr + ',2)<>0'],
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: T.COLORS.deficitFill } }, font: { bold: true, color: { argb: T.COLORS.deficitFont } } } }] });
      }
      ws.getRow(rr).height = 30;
    });

    var lastRow = tr + 2 + checks.length - 1;
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:G' + lastRow, titleRows: hr + ':' + hr,
      reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── ورقة البيانات الأساسية ─────────────

  function buildBasic(ctx) {
    var ws = ctx.sheets.basic, T = ctx.T, N = T.NAMES, L = T.LAYOUT.basic, cfg = ctx.cfg;
    var org = ctx.project.org, a = ctx.analysis;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'B', 'البيانات الأساسية', false);
    mergeStyled(ws, 'A2:B2', cfg.app.reportTitle, T.style.subtitle);

    var r = L.infoHeaderRow;
    mergeStyled(ws, 'A' + r + ':B' + r, 'بيانات الجهة', T.style.section);

    var actual = U.parseNumber(org.actualCount);
    var fields = [
      { label: 'اسم الجهة',          value: blankToNull(org.entity),     name: N.entity },
      { label: 'الإدارة',            value: blankToNull(org.department), name: N.department }
    ];
    // الأقسام / الوحدات: صف لكل قسم + صفوف فارغة، وتغذي القائمة المنسدلة في ورقة تحليل المهام
    var secCount = org.sections.length + L.spareSections;
    for (var si = 0; si < secCount; si++) {
      fields.push({ label: 'القسم / الوحدة (' + (si + 1) + ')', value: org.sections[si] ? org.sections[si].name : null, name: N.sectionNames });
    }
    fields = fields.concat([
      { label: 'العدد الفعلي',       value: U.isValidNumber(actual) ? actual : null, name: N.actualCount, numFmt: T.NUM.integer,
        dv: numberValidation('whole', 'greaterThanOrEqual', 0, 'العدد الفعلي يجب أن يكون عددًا صحيحًا لا يقل عن صفر.') },
      { label: 'تاريخ الإعداد',      value: ctx.now, numFmt: T.NUM.date }   // تاريخ التصدير تلقائيًا
    ]);

    fields.forEach(function (fd, i) {
      var row = r + 1 + i;
      var lc = ws.getCell('A' + row);
      lc.value = fd.label; T.style.label(lc);
      var vc = ws.getCell('B' + row);
      vc.value = fd.value;
      T.style.input(vc, fd.numFmt, fd.numFmt ? 'center' : 'right');
      if (fd.name) vc.addName(fd.name);
      if (fd.dv) vc.dataValidation = fd.dv;
      ws.getRow(row).height = 24;
      if (fd.name === N.sectionNames) ctx.sectionCells.push('B' + row);
    });

    var rr = r + fields.length + 1 + L.resultGapRows; // قسم النتائج بعد الحقول مباشرة مع سطر فاصل
    mergeStyled(ws, 'A' + rr + ':B' + rr, 'نتائج التحليل (محسوبة تلقائيًا)', T.style.section);
    var results = [
      { label: 'إجمالي ساعات العمل السنوية', v: f(N.totalHours, a.totalHours), fmt: T.NUM.hours },
      { label: 'الاحتياج المحسوب',           v: f(N.calcNeed, a.calcNeed), fmt: needFormat(cfg.needRounding.digits) },
      { label: 'الفجوة',                      v: f(N.gap, a.gap), fmt: gapFormat(cfg.needRounding.digits) },
      { label: 'الحالة',                      v: f(N.status, a.status), status: true }
    ];
    results.forEach(function (it, i) {
      var row = rr + 1 + i;
      var lc = ws.getCell('A' + row);
      lc.value = it.label; T.style.label(lc);
      var vc = ws.getCell('B' + row);
      vc.value = it.v;
      T.style.formula(vc, it.fmt, { bold: true });
      if (it.status) T.addStatusFormatting(ws, 'B' + row, cfg.status);
      ws.getRow(row).height = 24;
    });

    var lastRow = rr + results.length;
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:B' + lastRow, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── ورقة تحليل عبء العمل ─────────────

  function buildWorkload(ctx) {
    var ws = ctx.sheets.workload, T = ctx.T, N = T.NAMES, L = T.LAYOUT.workload, cfg = ctx.cfg, a = ctx.analysis;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'E', 'تحليل عبء العمل حسب التردد', true);

    var hr = L.headerRow;
    ['التردد', 'المرات السنوية', 'عدد المهام', 'إجمالي الساعات', 'نسبة الساعات'].forEach(function (h, i) {
      var c = ws.getRow(hr).getCell(i + 1); c.value = h; T.style.header(c);
    });
    ws.getRow(hr).height = 28;

    var first = hr + 1;
    a.byFrequency.forEach(function (bf, i) {
      var r = first + i;
      var stripe = i % 2 === 1 ? T.COLORS.stripe : T.COLORS.white;
      var row = ws.getRow(r);
      row.getCell(1).value = f('INDEX(' + N.freqNames + ',' + (i + 1) + ')', bf.label);
      row.getCell(2).value = f('INDEX(' + N.freqValues + ',' + (i + 1) + ')', bf.perYear);
      row.getCell(3).value = f('COUNTIF(' + N.taskFreq + ',A' + r + ')', bf.count);
      row.getCell(4).value = f('SUMIF(' + N.taskFreq + ',A' + r + ',' + N.taskHours + ')', bf.hours);
      row.getCell(5).value = f('IF(' + N.totalHours + '=0,0,D' + r + '/' + N.totalHours + ')', bf.share);
      T.style.formula(row.getCell(1), null, { fill: stripe, bold: true });
      T.style.formula(row.getCell(2), T.NUM.integer, { fill: stripe });
      T.style.formula(row.getCell(3), T.NUM.integer, { fill: stripe });
      T.style.formula(row.getCell(4), T.NUM.hours, { fill: stripe });
      T.style.formula(row.getCell(5), T.NUM.percent, { fill: stripe });
      row.height = 22;
    });

    var last = first + a.byFrequency.length - 1;
    var tr = last + 1;
    var tRow = ws.getRow(tr);
    tRow.getCell(1).value = 'الإجمالي';
    tRow.getCell(3).value = f('SUM(C' + first + ':C' + last + ')', a.taskCount);
    tRow.getCell(4).value = f('SUM(D' + first + ':D' + last + ')', a.totalHours);
    tRow.getCell(5).value = f('SUM(E' + first + ':E' + last + ')', a.totalHours > 0 ? 1 : 0);
    T.style.total(tRow.getCell(1));
    T.style.total(tRow.getCell(2));
    T.style.total(tRow.getCell(3), T.NUM.integer);
    T.style.total(tRow.getCell(4), T.NUM.hours);
    T.style.total(tRow.getCell(5), T.NUM.percent);
    tRow.height = 24;

    ws.autoFilter = 'A' + hr + ':E' + last;
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:E' + tr, titleRows: hr + ':' + hr, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── ورقة الاحتياج والفجوة ─────────────

  function buildNeed(ctx) {
    var ws = ctx.sheets.need, T = ctx.T, N = T.NAMES, L = T.LAYOUT.need, cfg = ctx.cfg, a = ctx.analysis;
    var digits = cfg.needRounding.digits;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'D', 'الاحتياج والفجوة', true);

    var hr = L.headerRow;
    ['الخطوة', 'البيان', 'القيمة', 'طريقة الحساب'].forEach(function (h, i) {
      var c = ws.getRow(hr).getCell(i + 1); c.value = h; T.style.header(c);
    });
    ws.getRow(hr).height = 28;

    var st = cfg.status;
    var steps = [
      { label: 'إجمالي ساعات العمل السنوية للمهام', v: f(N.totalHours, a.totalHours), fmt: T.NUM.hours,
        how: 'مجموع (المرات السنوية × التكرار × المدة بالدقائق ÷ ' + cfg.minutesPerHour + ') لجميع المهام' },
      { label: 'ساعات العمل الفعلية السنوية للموظف', v: f(N.annualHours, a.annualWorkHours), fmt: T.NUM.integer,
        how: 'من ورقة الإعدادات' },
      { label: 'الاحتياج قبل التقريب (للاطلاع)', name: N.exactNeed, fmt: T.NUM.decimal,
        v: f('IF(' + N.annualHours + '=0,0,' + N.totalHours + '/' + N.annualHours + ')', a.exactNeed),
        how: 'إجمالي الساعات ÷ ساعات العمل الفعلية السنوية' },
      { label: 'الاحتياج المحسوب', name: N.calcNeed, fmt: needFormat(digits), bold: true,
        v: f('ROUND(' + N.exactNeed + ',' + N.needDecimals + ')', a.calcNeed),
        how: 'تقريب لأقرب عدد صحيح: 0.5 فأعلى يُجبر للأعلى' },
      { label: 'العدد الفعلي', v: f(N.actualCount, a.actualCount), fmt: T.NUM.integer,
        how: 'من ورقة البيانات الأساسية' },
      { label: 'الفجوة', name: N.gap, fmt: gapFormat(digits), bold: true,
        v: f(N.actualCount + '-' + N.calcNeed, a.gap),
        how: 'العدد الفعلي − الاحتياج المحسوب' },
      { label: 'الحالة', name: N.status, bold: true, status: true,
        v: f('IF(' + N.gap + '>0,"' + st.surplus + '",IF(' + N.gap + '<0,"' + st.deficit + '","' + st.balanced + '"))', a.status),
        how: 'موجبة = ' + st.surplus + ' ، سالبة = ' + st.deficit + ' ، صفر = ' + st.balanced }
    ];

    steps.forEach(function (s, i) {
      var r = hr + 1 + i;
      var row = ws.getRow(r);
      row.getCell(1).value = i + 1;
      row.getCell(2).value = s.label;
      row.getCell(3).value = s.v;
      row.getCell(4).value = s.how;
      T.style.formula(row.getCell(1), null, { fill: T.COLORS.primaryLight, bold: true });
      T.style.label(row.getCell(2));
      T.style.formula(row.getCell(3), s.fmt, { bold: s.bold, size: s.bold ? 12 : 11 });
      T.style.formula(row.getCell(4), null, { align: 'right' });
      row.getCell(4).font = T.font({ size: 10, color: T.COLORS.muted });
      if (s.name) row.getCell(3).addName(s.name);
      if (s.status) T.addStatusFormatting(ws, 'C' + r, st);
      row.height = 26;
    });

    var lastRow = hr + steps.length;
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:D' + lastRow, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── الملخص التنفيذي ─────────────

  function narrativeText(ctx) {
    var a = ctx.analysis;
    var txt = 'بناءً على تحليل ' + a.taskCount + ' مهمة بإجمالي ' + U.formatNumber(a.totalHours, 2) +
      ' ساعة عمل سنويًا، وبواقع ' + a.annualWorkHours + ' ساعة عمل فعلية سنوية للموظف، يبلغ الاحتياج المحسوب ' +
      a.calcNeed + ' موظف مقابل عدد فعلي ' + a.actualCount + '، وبذلك تكون الحالة: ' + a.status;
    return txt + (a.gap === 0 ? '.' : ' بمقدار ' + Math.abs(a.gap) + ' موظف.');
  }

  function buildSummary(ctx) {
    var ws = ctx.sheets.summary, T = ctx.T, N = T.NAMES, L = T.LAYOUT.summary, cfg = ctx.cfg, a = ctx.analysis;
    var digits = cfg.needRounding.digits, st = cfg.status;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'H', 'الملخص التنفيذي', false);
    mergeStyled(ws, 'A2:H2', cfg.app.reportTitle, T.style.subtitle);

    var basic = function (addr) { return sheetRef(ctx, 'basic', addr); };
    var org = ctx.project.org;
    // قائمة الأقسام كمعادلة حية: كل قسم غير فارغ يُسبق بفاصلة ثم تُحذف الفاصلة الأولى
    var secJoin = ctx.sectionCells.map(function (c) { return 'IF(' + basic(c) + '<>"","، "&' + basic(c) + ',"")'; }).join('&');
    var secText = org.sections.map(function (x) { return x.name; }).join('، ');
    var info = [
      { label: 'الجهة', v: f(N.entity, org.entity) },
      { label: 'الإدارة', v: f('IF(' + N.department + '="","—",' + N.department + ')', String(org.department || '').trim() || '—') },
      { label: 'الأقسام / الوحدات', v: f('IF(LEN(' + secJoin + ')=0,"—",MID(' + secJoin + ',3,1000))', secText || '—') }
    ];
    info.forEach(function (it, i) {
      var r = 4 + i;
      mergeStyled(ws, 'A' + r + ':B' + r, it.label, T.style.label);
      var vc = mergeStyled(ws, 'C' + r + ':H' + r, it.v);
      T.style.formula(vc, null, { align: 'right', bold: true });
      ws.getRow(r).height = i === 2 ? estimateRowHeight(secText, 90) : 24;
    });

    /** صف مؤشرات على 8 أعمدة؛ كل مؤشر يأخذ عمودين افتراضيًا أو span أعمدة. */
    function kpiRow(labelRow, items) {
      var col = 0;
      items.forEach(function (it) {
        var span = it.span || 2;
        var c1 = String.fromCharCode(65 + col), c2 = String.fromCharCode(65 + col + span - 1);
        col += span;
        mergeStyled(ws, c1 + labelRow + ':' + c2 + labelRow, it.label, T.style.kpiLabel);
        var vc = mergeStyled(ws, c1 + (labelRow + 1) + ':' + c2 + (labelRow + 1), it.v);
        T.style.kpiValue(vc, it.fmt);
        if (it.status) T.addStatusFormatting(ws, c1 + (labelRow + 1), st);
      });
      ws.getRow(labelRow).height = 22;
      ws.getRow(labelRow + 1).height = 40;
    }

    kpiRow(8, [
      { label: 'عدد المهام', v: f(N.taskCount, a.taskCount), fmt: T.NUM.integer },
      { label: 'إجمالي ساعات العمل السنوية', v: f(N.totalHours, a.totalHours), fmt: T.NUM.hours },
      { label: 'ساعات العمل الفعلية السنوية', v: f(N.annualHours, a.annualWorkHours), fmt: T.NUM.integer },
      { label: 'الاحتياج المحسوب', v: f(N.calcNeed, a.calcNeed), fmt: needFormat(digits) }
    ]);
    kpiRow(11, [
      { label: 'العدد الفعلي', v: f(N.actualCount, a.actualCount), fmt: T.NUM.integer, span: 2 },
      { label: 'الفجوة', v: f(N.gap, a.gap), fmt: gapFormat(digits), span: 3 },
      { label: 'الحالة', v: f(N.status, a.status), status: true, span: 3 }
    ]);

    mergeStyled(ws, 'A14:H14', 'الخلاصة', T.style.section);
    var narrativeF =
      '"بناءً على تحليل "&' + N.taskCount + '&" مهمة بإجمالي "&FIXED(' + N.totalHours + ',2)&' +
      '" ساعة عمل سنويًا، وبواقع "&' + N.annualHours + '&" ساعة عمل فعلية سنوية للموظف، يبلغ الاحتياج المحسوب "&' +
      N.calcNeed + '&" موظف مقابل عدد فعلي "&' + N.actualCount + '&"، وبذلك تكون الحالة: "&' + N.status +
      '&IF(' + N.gap + '=0,"."," بمقدار "&ABS(' + N.gap + ')&" موظف.")';
    var p = mergeStyled(ws, 'A15:H17', f(narrativeF, narrativeText(ctx)), T.style.paragraph);
    ws.getRow(15).height = 24; ws.getRow(16).height = 24; ws.getRow(17).height = 24;
    p.protection = { locked: true };

    mergeStyled(ws, 'A19:H19', 'الاعتماد', T.style.section);
    var sign = [
      ['معد البيانات', null, 'التوقيع'],
      ['المعتمد', null, 'التوقيع'],
      ['التاريخ', null, '']
    ];
    sign.forEach(function (s, i) {
      var r = 20 + i;
      mergeStyled(ws, 'A' + r + ':B' + r, s[0], T.style.label);
      T.style.input(mergeStyled(ws, 'C' + r + ':D' + r, s[1]));
      if (s[2]) {
        mergeStyled(ws, 'E' + r + ':F' + r, s[2], T.style.label);
        T.style.input(mergeStyled(ws, 'G' + r + ':H' + r));
      }
      ws.getRow(r).height = 30;
    });

    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:H22', reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── نطاق الطباعة الديناميكي ─────────────

  /**
   * ExcelJS لا يدعم نطاق طباعة مبنيًا على معادلة، لذا نمدد نموذج الأسماء المعرّفة لهذا المصنف فقط
   * (المكتبة مثبتة محليًا بإصدار 4.4.0). إذا لم تتوفر البنية المتوقعة يبقى نطاق الطباعة الثابت كما هو.
   */
  function applyDynamicPrintAreas(ctx) {
    var list = ctx.dynamicPrintAreas;
    if (!list.length) return;
    var dn = ctx.wb.definedNames;
    var desc = dn && Object.getOwnPropertyDescriptor(Object.getPrototypeOf(dn), 'model');
    if (!desc || typeof desc.get !== 'function') return;

    var order = ctx.cfg.excel.sheetOrder;
    var extra = list.map(function (p) {
      return { name: '_xlnm.Print_Area', localSheetId: order.indexOf(p.sheetKey), ranges: [p.formula] };
    });
    Object.defineProperty(dn, 'model', {
      configurable: true,
      get: function () { return desc.get.call(this).concat(extra); },
      set: function (v) { if (desc.set) desc.set.call(this, v); }
    });
    list.forEach(function (p) { ctx.sheets[p.sheetKey].pageSetup.printArea = undefined; });
  }

  // ───────────── التجميع ─────────────

  function protectionOptions() {
    return {
      selectLockedCells: true, selectUnlockedCells: true,
      formatColumns: true, formatRows: true, formatCells: false,
      insertRows: false, deleteRows: false,
      autoFilter: true, sort: true
    };
  }

  /**
   * يبني مصنف Excel كاملًا.
   * @returns {Promise<ExcelJS.Workbook>}
   */
  async function buildWorkbook(project, cfg, T, ExcelJSLib) {
    var ExcelJS = ExcelJSLib || root.ExcelJS;
    if (!ExcelJS) throw new Error('EXCELJS_MISSING');

    var wb = new ExcelJS.Workbook();
    var now = new Date();
    wb.creator = cfg.app.title;
    wb.title = cfg.app.reportTitle;
    wb.created = now;
    wb.modified = now;
    wb.calcProperties.fullCalcOnLoad = true;

    var ctx = {
      wb: wb, cfg: cfg, T: T, project: project,
      analysis: WL.calc.analyze(project, cfg),
      now: new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())),
      dateText: U.todayISO(),
      sheets: {},
      dynamicPrintAreas: [],
      sectionCells: []
    };

    var freeze = { tasks: T.LAYOUT.tasks.headerRow, workload: T.LAYOUT.workload.headerRow, need: T.LAYOUT.need.headerRow,
      positions: T.LAYOUT.positions.headerRow };
    cfg.excel.sheetOrder.forEach(function (key) {
      ctx.sheets[key] = wb.addWorksheet(cfg.excel.sheetNames[key], T.sheetOptions(freeze[key]));
    });

    // الترتيب مهم: الإعدادات أولًا ثم المهام، لأن باقي الأوراق تعتمد على أسمائهما
    buildSettings(ctx);
    buildTasks(ctx);
    buildPositions(ctx);
    buildBasic(ctx);   // قبل الملخص: يحدد خلايا الأقسام التي يقرؤها
    buildWorkload(ctx);
    buildNeed(ctx);
    buildSummary(ctx);
    applyDynamicPrintAreas(ctx);

    if (cfg.excel.protection.enabled) {
      var pwd = cfg.excel.protection.password || '';
      for (var i = 0; i < cfg.excel.sheetOrder.length; i++) {
        await ctx.sheets[cfg.excel.sheetOrder[i]].protect(pwd, protectionOptions());
      }
    }
    return wb;
  }

  function buildFileName(project, cfg) {
    var entity = U.sanitizeFileName(project.org.entity) || 'جهة';
    return cfg.excel.fileNamePrefix + '_' + entity + '_' + U.todayISO() + '.xlsx';
  }

  async function exportToBuffer(project, cfg, T, ExcelJSLib) {
    var wb = await buildWorkbook(project, cfg, T, ExcelJSLib);
    return wb.xlsx.writeBuffer();
  }

  function downloadBuffer(buffer, fileName) {
    var blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  WL.excel = {
    buildWorkbook: buildWorkbook,
    exportToBuffer: exportToBuffer,
    buildFileName: buildFileName,
    downloadBuffer: downloadBuffer
  };
})(typeof window !== 'undefined' ? window : globalThis);
