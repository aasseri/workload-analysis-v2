/**
 * calculations.js — محرك الحساب.
 * يطبق نفس القواعد التي تُكتب كمعادلات في Excel، ويُستخدم للعرض والمعاينة
 * ولتعبئة "النتيجة المخزنة" بجوار كل معادلة. لا يُصدَّر أي ناتج منه كقيمة ثابتة.
 *
 *   ساعات المهمة     = المرات السنوية × التكرار × المدة ÷ 60
 *   إجمالي الساعات   = Σ ساعات المهام
 *   الاحتياج الدقيق   = إجمالي الساعات ÷ ساعات العمل السنوية
 *   الاحتياج المحسوب = تقريب(الاحتياج الدقيق)  ← 0.5 فأعلى للأعلى
 *   الفجوة           = العدد الفعلي − الاحتياج المحسوب
 *   الحالة           = فائض (>0) | عجز (<0) | متوازن (=0)
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  function findFrequency(cfg, key) {
    for (var i = 0; i < cfg.frequencies.length; i++) {
      if (cfg.frequencies[i].key === key) return cfg.frequencies[i];
    }
    return null;
  }

  /** @returns {number|null} null إذا كانت بيانات المهمة غير مكتملة */
  function taskHours(task, cfg) {
    var freq = findFrequency(cfg, task.frequencyKey);
    var rep = U.parseNumber(task.repetitions);
    var dur = U.parseNumber(task.durationMinutes);
    if (!freq || !U.isValidNumber(rep) || !U.isValidNumber(dur)) return null;
    return (freq.perYear * rep * dur) / cfg.minutesPerHour;
  }

  function statusOf(gap, cfg) {
    if (!U.isValidNumber(gap)) return { key: null, label: '—' };
    // الفجوة = العدد الفعلي − الاحتياج: موجبة = موظفون أكثر من الحاجة (فائض)، سالبة = أقل (عجز)
    if (gap > 0) return { key: 'surplus', label: cfg.status.surplus };
    if (gap < 0) return { key: 'deficit', label: cfg.status.deficit };
    return { key: 'balanced', label: cfg.status.balanced };
  }

  function analyze(project, cfg) {
    var rows = project.tasks.map(function (t) {
      var freq = findFrequency(cfg, t.frequencyKey);
      return {
        id: t.id,
        title: t.title,
        sectionId: t.sectionId || '',
        sectionName: sectionNameOf(project, t.sectionId),
        positionTitle: positionKey(t.positionTitle),
        frequencyKey: t.frequencyKey,
        frequencyLabel: freq ? freq.label : '',
        perYear: freq ? freq.perYear : null,
        repetitions: U.parseNumber(t.repetitions),
        durationMinutes: U.parseNumber(t.durationMinutes),
        hours: taskHours(t, cfg),
        share: null
      };
    });

    var totalHours = rows.reduce(function (sum, r) {
      return sum + (r.hours === null ? 0 : r.hours);
    }, 0);

    rows.forEach(function (r) {
      r.share = r.hours === null ? null : (totalHours === 0 ? 0 : r.hours / totalHours);
    });

    var byFrequency = cfg.frequencies.map(function (f) {
      var matched = rows.filter(function (r) { return r.frequencyKey === f.key; });
      var hours = matched.reduce(function (s, r) { return s + (r.hours || 0); }, 0);
      return {
        key: f.key,
        label: f.label,
        perYear: f.perYear,
        count: matched.length,
        hours: hours,
        share: totalHours === 0 ? 0 : hours / totalHours
      };
    });

    var annual = cfg.annualWorkHours;
    var exactNeed = annual > 0 ? totalHours / annual : 0;
    var calcNeed = U.roundHalfUp(exactNeed, cfg.needRounding.digits);

    // الاحتياج حسب المسمى الفعلي: نفس القاعدة (ساعات المسمى ÷ ساعات العمل السنوية، ثم التقريب) لكل مسمى
    var byPosition = [];
    rows.forEach(function (r) {
      if (!r.positionTitle) return;
      var g = byPosition.filter(function (p) { return p.title === r.positionTitle; })[0];
      if (!g) { g = { title: r.positionTitle, count: 0, hours: 0 }; byPosition.push(g); }
      g.count += 1;
      g.hours += r.hours || 0;
    });
    byPosition.forEach(function (g) {
      g.exactNeed = annual > 0 ? g.hours / annual : 0;
      g.need = U.roundHalfUp(g.exactNeed, cfg.needRounding.digits);
    });
    // الاحتياج حسب القسم / الوحدة × المسمى الفعلي: كل (قسم، مسمى) يُحسب ويُقرَّب مستقلًا، مع مجموع لكل قسم.
    // ترتيب الأقسام كما عُرّفت في البيانات الأساسية، والمسميات داخل القسم بترتيب ظهورها.
    var bySection = (project.org.sections || []).map(function (s) {
      return { id: s.id, name: s.name, positions: [], count: 0, hours: 0, needSum: 0 };
    });
    rows.forEach(function (r) {
      if (!r.positionTitle) return;
      var sec = bySection.filter(function (s) { return s.id === r.sectionId; })[0];
      if (!sec) return;
      var g = sec.positions.filter(function (p) { return p.title === r.positionTitle; })[0];
      if (!g) { g = { title: r.positionTitle, count: 0, hours: 0 }; sec.positions.push(g); }
      g.count += 1;
      g.hours += r.hours || 0;
    });
    bySection.forEach(function (sec) {
      sec.positions.forEach(function (g) {
        g.exactNeed = annual > 0 ? g.hours / annual : 0;
        g.need = U.roundHalfUp(g.exactNeed, cfg.needRounding.digits);
        sec.count += g.count;
        sec.hours += g.hours;
        sec.needSum += g.need;
      });
    });
    var positionsNeedSum = bySection.reduce(function (s, sec) { return s + sec.needSum; }, 0);

    var actual = U.parseNumber(project.org.actualCount);
    var gap = U.isValidNumber(actual) ? actual - calcNeed : null;
    var status = statusOf(gap, cfg);

    return {
      tasks: rows,
      taskCount: rows.length,
      totalHours: totalHours,
      annualWorkHours: annual,
      exactNeed: exactNeed,
      calcNeed: calcNeed,
      actualCount: U.isValidNumber(actual) ? actual : null,
      gap: gap,
      statusKey: status.key,
      status: status.label,
      byFrequency: byFrequency,
      byPosition: byPosition,          // على مستوى الجهة (للاطلاع)
      bySection: bySection,            // القسم × المسمى (المعتمد في المراجعة والتصدير)
      positionsNeedSum: positionsNeedSum
    };
  }

  function sectionNameOf(project, id) {
    var s = (project.org.sections || []).filter(function (x) { return x.id === id; })[0];
    return s ? s.name : '';
  }

  /** صيغة موحّدة للمسمى الفعلي (بلا مسافات زائدة)؛ تطابق ما يُكتب في Excel حتى يعمل SUMIF بدقة. */
  function positionKey(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
  }

  WL.calc = {
    positionKey: positionKey,
    findFrequency: findFrequency,
    taskHours: taskHours,
    statusOf: statusOf,
    analyze: analyze
  };
})(typeof window !== 'undefined' ? window : globalThis);
