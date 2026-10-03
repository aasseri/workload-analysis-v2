/**
 * task-import.js — تحويل نص ملصوق من Excel أو Word إلى قائمة مهام.
 * منطق بحت بلا واجهة: يقسّم النص إلى صفوف وأعمدة، ويكتشف العناوين والأعمدة،
 * وينظف الترقيم والرموز من بداية كل مهمة، ويحدد المهام المكررة. لا يعدّل المشروع.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  /** رمز أو رقم ترقيم وحده في خلية (عمود «م» في Excel أو «1.» من Word). */
  var MARKER_ONLY = /^\s*(?:[0-9٠-٩]+\s*[-.)٫]?|[(（][0-9٠-٩]+[)）]|[-–—•*·▪●◦►✓✔]|[أ-ي]\s*[-.)])\s*$/;

  /** ترقيم أو رمز في بداية النص. */
  var LEADING_MARKER = /^\s*(?:[-–—•*·▪●◦►✓✔]+|[0-9٠-٩]+\s*[-.)٫]|[(（][0-9٠-٩]+[)）]|[أ-ي]\s*[-.)](?=\s))\s*/;

  /** يحذف الترقيم والرموز من بداية الوصف والمسافات الزائدة فقط، ولا يغير نص المهمة. */
  function cleanTitle(text) {
    var t = String(text || '').replace(/\s+/g, ' ').trim();
    var prev;
    do { prev = t; t = t.replace(LEADING_MARKER, '').trim(); } while (t !== prev);
    return t;
  }

  /** تطبيع عربي للمقارنة فقط (التشكيل، التطويل، الهمزات، التاء المربوطة، الألف المقصورة). */
  function normalizeArabic(text) {
    return String(text || '')
      .replace(/[ً-ْـ]/g, '')
      .replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي')
      .replace(/[^؀-ۿa-zA-Z0-9]+/g, ' ')
      .replace(/\s+/g, ' ').trim();
  }

  /**
   * يقسم نصًا ملصوقًا إلى صفوف وخلايا (صيغة Excel: خلايا بعلامة Tab، وخلية بها سطر جديد تكون بين علامتي تنصيص).
   * @returns {string[][]}
   */
  function parseTable(text) {
    var s = String(text || '').replace(/\r\n?/g, '\n');
    var rows = [], row = [], cell = '', i = 0, inQuotes = false;
    while (i < s.length) {
      var ch = s[i];
      if (inQuotes) {
        if (ch === '"' && s[i + 1] === '"') { cell += '"'; i += 2; continue; }
        if (ch === '"') { inQuotes = false; i += 1; continue; }
        cell += ch; i += 1; continue;
      }
      if (ch === '"' && cell === '') { inQuotes = true; i += 1; continue; }
      if (ch === '\t') { row.push(cell); cell = ''; i += 1; continue; }
      if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; i += 1; continue; }
      cell += ch; i += 1;
    }
    row.push(cell);
    rows.push(row);
    return rows
      .map(function (r) { return r.map(function (c) { return c.replace(/\s+/g, ' ').trim(); }); })
      .filter(function (r) { return r.some(function (c) { return c; }); });
  }

  /** يطابق نص التردد مع القيم المعتمدة فقط (يقبل الصيغ مثل «يوميًا» و«شهريا»)، وإلا يعيد ''. */
  function matchFrequency(value, cfg) {
    var v = normalizeArabic(value).replace(/ا$/, '');
    if (!v) return '';
    for (var i = 0; i < cfg.frequencies.length; i++) {
      if (normalizeArabic(cfg.frequencies[i].label) === v) return cfg.frequencies[i].key;
    }
    return '';
  }

  function positiveNumber(value) {
    var n = U.parseNumber(value);
    return U.isValidNumber(n) && n > 0 ? n : null;
  }

  /** كلمات عناوين الأعمدة (بعد التطبيع)، وتُطابق ككلمات كاملة لا كأجزاء من كلمات. */
  var HEADER_WORDS = {
    title: ['المهمه', 'المهام', 'مهمه', 'مهام', 'الوصف', 'وصف', 'النشاط', 'نشاط', 'الانشطه', 'البيان'],
    frequency: ['التردد', 'تردد', 'الدوريه', 'دوريه'],
    repetitions: ['التكرار', 'تكرار', 'التكرارات', 'العدد', 'عدد'],
    duration: ['المده', 'مده', 'الزمن', 'زمن', 'الوقت', 'وقت', 'الدقائق', 'دقائق', 'بالدقائق', 'دقيقه'],
    position: ['المسمي', 'مسمي', 'المسميات', 'الوظيفه', 'وظيفه'],
    section: ['القسم', 'قسم', 'الاقسام', 'الوحده', 'وحده', 'الوحدات']
  };

  function headerField(name) {
    var words = normalizeArabic(name).split(' ');
    return Object.keys(HEADER_WORDS).filter(function (k) {
      return words.some(function (w) { return HEADER_WORDS[k].indexOf(w) >= 0; });
    })[0] || null;
  }

  /** هل الصف الأول صف عناوين؟ (خلايا قصيرة، وفيها كلمة عنوان معروفة، ولا تحتوي قيمة تردد مثل «شهري») */
  function looksLikeHeader(row, cfg) {
    if (!row.every(function (c) { return String(c).length <= 30; })) return false;
    if (row.some(function (c) { return matchFrequency(c, cfg); })) return false;
    return row.some(function (c) { return headerField(c) || normalizeArabic(c) === 'م'; });
  }

  /**
   * يحلل النص الملصوق.
   * @returns {{ columns: Array<{index, name, sample}>, hasHeader: boolean, rows: string[][], mapping: {title, frequency, repetitions, duration} }}
   *   mapping: رقم العمود لكل حقل، أو -1 إن لم يوجد
   */
  function analyze(text, cfg) {
    var rows = parseTable(text);
    if (!rows.length) return { columns: [], hasHeader: false, rows: [], mapping: { title: -1, section: -1, position: -1, frequency: -1, repetitions: -1, duration: -1 } };

    var width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    rows = rows.map(function (r) { while (r.length < width) r.push(''); return r; });

    // قائمة (Word أو عمود واحد): كل سطر فيه نص واحد فقط بعد استبعاد رمز الترقيم في أوله (مثل «1.<Tab>النص»)
    var listMode = rows.every(function (r) { return contentCells(r).length <= 1; });
    if (listMode) {
      var first = rows.map(textCell).filter(Boolean)[0] || '';
      // عنوان عمود مثل «المهام» أو «وصف المهمة» في السطر الأول
      var firstIsHeading = rows.length > 1 && first.split(/\s+/).length <= 2 && headerField(first) === 'title';
      return { listMode: true, firstIsHeading: firstIsHeading, columns: [{ index: -1, name: 'النص', sample: first }],
        hasHeader: false, rows: rows, mapping: { title: -1, section: -1, position: -1, frequency: -1, repetitions: -1, duration: -1 } };
    }

    var hasHeader = width > 1 && rows.length > 1 && looksLikeHeader(rows[0], cfg);
    var header = hasHeader ? rows[0] : null;
    var data = hasHeader ? rows.slice(1) : rows;

    // عمود الترقيم (مثل «م») لا يُعرض ولا يُستخدم
    var keep = [];
    for (var c = 0; c < width; c++) {
      var values = data.map(function (r) { return r[c]; }).filter(Boolean);
      if (values.length && !isNumberingColumn(values, header && header[c])) keep.push(c);
    }
    if (!keep.length) keep = [0];

    var columns = keep.map(function (ci, i) {
      var sample = (data.filter(function (r) { return r[ci]; })[0] || [])[ci] || '';
      return { index: ci, name: header && header[ci] ? header[ci] : 'العمود ' + (i + 1), sample: sample };
    });

    var mapping = { title: -1, section: -1, position: -1, frequency: -1, repetitions: -1, duration: -1 };
    function colBy(test) {
      var hit = columns.filter(function (col) { return test(col); })[0];
      return hit ? hit.index : -1;
    }
    function share(ci, pred) {
      var vals = data.map(function (r) { return r[ci]; }).filter(Boolean);
      return vals.length ? vals.filter(pred).length / vals.length : 0;
    }
    function used(ci) { return Object.keys(mapping).some(function (k) { return mapping[k] === ci; }); }

    if (header) {
      ['frequency', 'repetitions', 'duration', 'section', 'position', 'title'].forEach(function (k) {
        mapping[k] = colBy(function (col) { return !used(col.index) && headerField(col.name) === k; });
      });
    }
    // أعمدة الأرقام تُقبل فقط إذا كانت معظم قيمها أرقامًا أكبر من صفر
    ['repetitions', 'duration'].forEach(function (k) {
      if (mapping[k] >= 0 && share(mapping[k], function (v) { return positiveNumber(v) !== null; }) < 0.6) mapping[k] = -1;
    });
    if (mapping.frequency < 0 && columns.length > 1) {
      mapping.frequency = colBy(function (col) {
        return !used(col.index) && share(col.index, function (v) { return !!matchFrequency(v, cfg); }) >= 0.6;
      });
    }
    if (mapping.title < 0) {
      // العمود ذو النص الأطول في المتوسط
      var best = -1, bestLen = -1;
      columns.forEach(function (col) {
        if (used(col.index)) return;
        var vals = data.map(function (r) { return r[col.index]; }).filter(Boolean);
        var avg = vals.reduce(function (s, v) { return s + v.length; }, 0) / Math.max(1, vals.length);
        if (avg > bestLen) { bestLen = avg; best = col.index; }
      });
      mapping.title = best;
    }
    // ملاحظة: أعمدة التكرار والمدة تُربط تلقائيًا فقط عند وجود عنوان صريح، لتجنب الخلط بينهما

    return { listMode: false, columns: columns, hasHeader: hasHeader, rows: data, mapping: mapping };
  }

  /** خلايا المحتوى في السطر: كل الخلايا غير الفارغة عدا رمز الترقيم إن كان في أول خلية. */
  function contentCells(r) {
    return r.filter(function (c, i) { return c && !(i === 0 && MARKER_ONLY.test(c)); });
  }

  /** النص الوحيد في سطر القائمة. */
  function textCell(r) {
    return contentCells(r)[0] || '';
  }

  /**
   * عمود ترقيم: عنوانه «م» أو «الرقم»، أو قيمه رموز ترقيم (1. ، • ، (1))، أو أرقام متسلسلة 1، 2، 3...
   * (أعمدة الأرقام العادية مثل المدة أو التكرار ليست متسلسلة فلا تُحذف)
   */
  function isNumberingColumn(values, headerName) {
    var h = normalizeArabic(headerName);
    if (h === 'م' || h === 'الرقم' || h === 'رقم' || h === 'التسلسل') return true;
    if (!values.every(function (v) { return MARKER_ONLY.test(v); })) return false;
    var plain = values.map(function (v) { return U.parseNumber(v); });
    if (!plain.every(function (n) { return U.isValidNumber(n) && Math.floor(n) === n; })) return true; // «1.» أو «•»
    return plain.every(function (n, i) { return i === 0 || n === plain[i - 1] + 1; });
  }

  /**
   * يبني عناصر المعاينة حسب ربط الأعمدة، ويحدد المكرر مع المهام الحالية أو داخل النص الملصوق.
   * @returns {Array<{title, frequencyKey, repetitions, durationMinutes, duplicate: ''|'existing'|'pasted', heading, selected}>}
   */
  function buildItems(parsed, mapping, existingTasks, cfg) {
    var existing = {};
    (existingTasks || []).forEach(function (t) { existing[normalizeArabic(t.title)] = true; });
    var seen = {};
    var items = [];
    parsed.rows.forEach(function (r, rowIndex) {
      var raw = parsed.listMode ? textCell(r) : (mapping.title >= 0 ? r[mapping.title] : '');
      var title = cleanTitle(raw);
      if (!title) return;
      // سطر عنوان مثل «المهام:» يظهر في المعاينة غير محدد
      var heading = /[:：]$/.test(title) || (rowIndex === 0 && !!parsed.firstIsHeading);
      title = title.replace(/[:：]+$/, '').trim();
      if (!title) return;
      var key = normalizeArabic(title);
      var duplicate = existing[key] ? 'existing' : seen[key] ? 'pasted' : '';
      seen[key] = true;
      var reps = mapping.repetitions >= 0 ? positiveNumber(r[mapping.repetitions]) : null;
      var dur = mapping.duration >= 0 ? positiveNumber(r[mapping.duration]) : null;
      items.push({
        title: title.slice(0, cfg.ui.maxTaskTitleLength),
        sectionName: mapping.section >= 0 ? String(r[mapping.section] || '').replace(/\s+/g, ' ').trim() : '',
        positionTitle: mapping.position >= 0 ? String(r[mapping.position] || '').replace(/\s+/g, ' ').trim() : '',
        frequencyKey: mapping.frequency >= 0 ? matchFrequency(r[mapping.frequency], cfg) : '',
        repetitions: reps === null ? '' : String(reps),
        durationMinutes: dur === null ? '' : String(dur),
        duplicate: duplicate,
        heading: heading,
        selected: !duplicate && !heading
      });
    });
    return items;
  }

  WL.taskImport = {
    cleanTitle: cleanTitle,
    normalizeArabic: normalizeArabic,
    parseTable: parseTable,
    matchFrequency: matchFrequency,
    analyze: analyze,
    buildItems: buildItems
  };
})(typeof window !== 'undefined' ? window : globalThis);
