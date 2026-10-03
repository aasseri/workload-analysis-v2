/**
 * template.js — محرك القالب: الهوية البصرية لملف Excel، خريطة مواقع الخلايا،
 * الأسماء المعرّفة، الطباعة، والتنسيق الشرطي.
 * تغيير شكل التقرير يتم هنا فقط دون المساس بمنطق الحساب أو بناء المعادلات.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});

  var C = {
    primary:      'FF1B5E4B',
    primaryDark:  'FF123F33',
    primaryLight: 'FFE8F1ED',
    stripe:       'FFF6F9F8',
    border:       'FFB9CCC4',
    text:         'FF1F2937',
    muted:        'FF5B6770',
    white:        'FFFFFFFF',
    inputFill:    'FFFFFBEA',   // خلايا الإدخال القابلة للتعديل
    warnFill:     'FFFDE8E8',
    deficitFill:  'FFFBE0DF', deficitFont:  'FF9B1C1C',
    surplusFill:  'FFFFF0D4', surplusFont:  'FF8A5300',
    balancedFill: 'FFECEFF1', balancedFont: 'FF374151'
  };

  var FONT = 'Arial';

  var NUM = {
    hours:   '#,##0.00',
    decimal: '#,##0.00',
    integer: '#,##0',
    general: 'General',
    percent: '0.00%',
    date:    'yyyy/mm/dd'
  };

  /** الأسماء المعرّفة داخل ملف Excel — تُستخدم في كل المعادلات بدل العناوين المباشرة. */
  var NAMES = {
    annualHours:  'AnnualHours',
    needDecimals: 'NeedDecimals',
    freqNames:    'FreqNames',
    freqValues:   'FreqValues',
    taskTitles:   'TaskTitles',
    taskPositions:'TaskPositions',
    taskSections: 'TaskSections',
    sectionNames: 'SectionNames',
    department:   'Department',
    taskFreq:     'TaskFreq',
    taskHours:    'TaskHours',
    totalHours:   'TotalHours',
    taskCount:    'TaskCount',
    entity:       'Entity',
    actualCount:  'ActualCount',
    exactNeed:    'ExactNeed',
    calcNeed:     'CalcNeed',
    gap:          'Gap',
    status:       'Status'
  };

  /** خريطة مواقع الخلايا والأعمدة لكل ورقة. */
  var LAYOUT = {
    settings: {
      columns: [38, 18, 46],
      annualHoursRow: 4,
      needDecimalsRow: 5,
      freqHeaderRow: 8
    },
    tasks: {
      totalRow: 3,
      countRow: 4,
      noteRow: 5,
      headerRow: 6,
      columns: [
        { key: 'no',       header: 'م',               width: 6 },
        { key: 'title',    header: 'المهمة',          width: 48 },
        { key: 'section',  header: 'القسم / الوحدة',  width: 22 },
        { key: 'position', header: 'المسمى الفعلي',   width: 22 },
        { key: 'freq',     header: 'التردد',          width: 14 },
        { key: 'perYear',  header: 'المرات السنوية',  width: 14 },
        { key: 'reps',     header: 'التكرار',         width: 11 },
        { key: 'duration', header: 'المدة بالدقائق',  width: 14 },
        { key: 'hours',    header: 'إجمالي الساعات',  width: 16 },
        { key: 'share',    header: 'نسبة الساعات',    width: 14 }
      ],
      tableName: 'tblTasks'
    },
    basic: {
      columns: [34, 52],
      infoHeaderRow: 4,
      spareSections: 5,  // صفوف فارغة لإضافة أقسام يدويًا داخل Excel
      resultGapRows: 1   // أسطر فارغة بين بيانات الجهة ونتائج التحليل
    },
    workload: {
      columns: [18, 16, 14, 18, 16],
      headerRow: 4
    },
    need: {
      columns: [8, 40, 18, 52],
      headerRow: 4
    },
    summary: {
      columns: [15, 15, 15, 15, 15, 15, 15, 15]
    },
    positions: {   // الاحتياج حسب القسم / الوحدة × المسمى الفعلي
      columns: [6, 28, 28, 11, 15, 17, 12],
      headerRow: 4,
      spareRows: 10   // صفوف فارغة لإضافة (قسم، مسمى) يدويًا داخل Excel
    },
    settingsFreqTableName: 'tblFreq'
  };

  // ───────────── أنماط الخلايا ─────────────

  function thinBorder(color) {
    var s = { style: 'thin', color: { argb: color || C.border } };
    return { top: s, left: s, bottom: s, right: s };
  }

  function fill(argb) {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } };
  }

  function font(opts) {
    var o = opts || {};
    return {
      name: FONT,
      size: o.size || 11,
      bold: !!o.bold,
      color: { argb: o.color || C.text }
    };
  }

  var ALIGN_RIGHT = { horizontal: 'right', vertical: 'middle', wrapText: true, readingOrder: 'rtl' };
  var ALIGN_CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true, readingOrder: 'rtl' };

  var style = {
    title: function (cell) {
      cell.font = font({ size: 16, bold: true, color: C.white });
      cell.fill = fill(C.primary);
      cell.alignment = ALIGN_CENTER;
    },
    subtitle: function (cell) {
      cell.font = font({ size: 11, color: C.muted });
      cell.alignment = ALIGN_CENTER;
    },
    note: function (cell) {
      cell.font = font({ size: 9, color: C.muted });
      cell.alignment = ALIGN_RIGHT;
    },
    section: function (cell) {
      cell.font = font({ size: 12, bold: true, color: C.primaryDark });
      cell.fill = fill(C.primaryLight);
      cell.alignment = ALIGN_RIGHT;
      cell.border = thinBorder();
    },
    header: function (cell) {
      cell.font = font({ bold: true, color: C.white });
      cell.fill = fill(C.primary);
      cell.alignment = ALIGN_CENTER;
      cell.border = thinBorder(C.primaryDark);
    },
    label: function (cell) {
      cell.font = font({ bold: true, color: C.primaryDark });
      cell.fill = fill(C.primaryLight);
      cell.alignment = ALIGN_RIGHT;
      cell.border = thinBorder();
    },
    /** خلية إدخال: مفتوحة للتعديل عند حماية الورقة. */
    input: function (cell, numFmt, align) {
      cell.font = font();
      cell.fill = fill(C.inputFill);
      cell.alignment = align === 'center' ? ALIGN_CENTER : ALIGN_RIGHT;
      cell.border = thinBorder();
      cell.protection = { locked: false };
      if (numFmt) cell.numFmt = numFmt;
    },
    /** خلية معادلة: مقفلة عند حماية الورقة. */
    formula: function (cell, numFmt, opts) {
      var o = opts || {};
      cell.font = font({ bold: !!o.bold, size: o.size });
      cell.fill = fill(o.fill || C.white);
      cell.alignment = o.align === 'right' ? ALIGN_RIGHT : ALIGN_CENTER;
      cell.border = thinBorder();
      cell.protection = { locked: true };
      if (numFmt) cell.numFmt = numFmt;
    },
    total: function (cell, numFmt) {
      cell.font = font({ bold: true, color: C.primaryDark });
      cell.fill = fill(C.primaryLight);
      cell.alignment = ALIGN_CENTER;
      cell.border = thinBorder(C.primary);
      cell.protection = { locked: true };
      if (numFmt) cell.numFmt = numFmt;
    },
    kpiLabel: function (cell) {
      cell.font = font({ size: 10, bold: true, color: C.muted });
      cell.fill = fill(C.primaryLight);
      cell.alignment = ALIGN_CENTER;
      cell.border = thinBorder();
    },
    kpiValue: function (cell, numFmt) {
      cell.font = font({ size: 18, bold: true, color: C.primaryDark });
      cell.fill = fill(C.white);
      cell.alignment = ALIGN_CENTER;
      cell.border = thinBorder();
      if (numFmt) cell.numFmt = numFmt;
    },
    paragraph: function (cell) {
      cell.font = font({ size: 12 });
      cell.alignment = { horizontal: 'right', vertical: 'top', wrapText: true, readingOrder: 'rtl' };
      cell.border = thinBorder();
    }
  };

  // ───────────── إعداد الورقة والطباعة ─────────────

  function sheetOptions(freezeRows) {
    var view = { rightToLeft: true, showGridLines: false, zoomScale: 100 };
    if (freezeRows) {
      view.state = 'frozen';
      view.ySplit = freezeRows;
      view.xSplit = 0;
      view.topLeftCell = 'A' + (freezeRows + 1);
    }
    return { views: [view], properties: { defaultRowHeight: 20 } };
  }

  function setColumnWidths(ws, widths) {
    widths.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
  }

  /**
   * @param {object} o { paper:'A4'|'A3', landscape, printArea, titleRows, reportTitle, dateText }
   */
  function applyPrint(ws, o) {
    ws.pageSetup = {
      paperSize: o.paper === 'A3' ? 8 : 9,
      orientation: o.landscape ? 'landscape' : 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      margins: { left: 0.4, right: 0.4, top: 0.7, bottom: 0.7, header: 0.3, footer: 0.3 },
      printArea: o.printArea
    };
    if (o.titleRows) ws.pageSetup.printTitlesRow = o.titleRows;
    ws.headerFooter = {
      oddHeader: '&C&"Arial,Bold"&11' + o.reportTitle,
      oddFooter: '&Rصفحة &P من &N&L' + (o.dateText || '')
    };
  }

  // ───────────── التنسيق الشرطي ─────────────

  function dxf(fillArgb, fontArgb) {
    return {
      fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: fillArgb } },
      font: { bold: true, color: { argb: fontArgb } }
    };
  }

  /** تنسيق الحالة مبني على قيمة الخلية (وليس لونًا ثابتًا). */
  function addStatusFormatting(ws, ref, statusLabels) {
    ws.addConditionalFormatting({
      ref: ref,
      rules: [
        { type: 'cellIs', operator: 'equal', priority: 1, formulae: ['"' + statusLabels.deficit + '"'],  style: dxf(C.deficitFill, C.deficitFont) },
        { type: 'cellIs', operator: 'equal', priority: 2, formulae: ['"' + statusLabels.surplus + '"'],  style: dxf(C.surplusFill, C.surplusFont) },
        { type: 'cellIs', operator: 'equal', priority: 3, formulae: ['"' + statusLabels.balanced + '"'], style: dxf(C.balancedFill, C.balancedFont) }
      ]
    });
  }

  /** يلوّن الخلية عند تحقق شرط (يستخدم لتنبيه المستخدم لبيانات ناقصة داخل Excel). */
  function addWarningRule(ws, ref, formula) {
    ws.addConditionalFormatting({
      ref: ref,
      rules: [{
        type: 'expression', priority: 10, formulae: [formula],
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.warnFill } }, border: thinBorder('FFE08A8A') }
      }]
    });
  }

  WL.template = {
    COLORS: C,
    FONT: FONT,
    NUM: NUM,
    NAMES: NAMES,
    LAYOUT: LAYOUT,
    style: style,
    fill: fill,
    font: font,
    thinBorder: thinBorder,
    sheetOptions: sheetOptions,
    setColumnWidths: setColumnWidths,
    applyPrint: applyPrint,
    addStatusFormatting: addStatusFormatting,
    addWarningRule: addWarningRule
  };
})(typeof window !== 'undefined' ? window : globalThis);
