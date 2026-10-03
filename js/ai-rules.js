/**
 * ai-rules.js — الفحص المحلي للمساعد الذكي (يعمل دون إنترنت).
 * ينتج «نتيجة مهمة» بنفس بنية نتائج نموذج الذكاء الاصطناعي، ولا يعدّل المشروع.
 *
 * نتيجة المهمة (TaskResult):
 *   { id, source: 'rules'|'llm', taskId, snapshot,
 *     quality: 'good'|'needs_improvement'|'not_measurable'|'',
 *     clarity, measurability, outputClarity, jobRelevance,          // تقييم وصفي (نص عربي أو فارغ)
 *     improvedTask, improvementReason, output, unit,
 *     suggestedFrequencyKey, frequencyConfidence, frequencyReason,
 *     missing: [], question, valueWarnings: [{text, source}], composite, splitSuggestions: [],
 *     recommendations: [], decisions: {title: 'pending'|'accepted', frequency: ...}, dismissed }
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  var seq = 0;

  function snapshotOf(task) {
    return { title: String(task.title || ''), frequencyKey: task.frequencyKey || '',
      repetitions: String(task.repetitions), durationMinutes: String(task.durationMinutes) };
  }

  /** نتيجة فارغة لمهمة. */
  function emptyResult(task, source) {
    seq += 1;
    return {
      id: 'r' + Date.now().toString(36) + seq,
      source: source || 'rules',
      taskId: task.id,
      snapshot: snapshotOf(task),
      quality: '', clarity: '', measurability: '', outputClarity: '', jobRelevance: '',
      improvedTask: '', improvementReason: '', output: '', unit: '',
      suggestedFrequencyKey: '', frequencyConfidence: '', frequencyReason: '',
      missing: [], question: '', valueWarnings: [], composite: false, splitSuggestions: [],
      recommendations: [], decisions: { title: 'pending', frequency: 'pending' }, dismissed: false
    };
  }

  // ───────────── الصياغة ─────────────

  var LEADING_MARKERS = /^\s*(?:[-–—•*·▪●◦]+|[0-9٠-٩]+\s*[-.)٫]|[(（][0-9٠-٩]+[)）])\s*/;
  var DOER_PREFIX = /^\s*(?:يقوم\s+(?:الموظف|شاغل\s+الوظيفة)?\s*ب|القيام\s+ب)\s*(?:ـ\s*)?/;

  /** إصلاحات آلية آمنة لا تغيّر المعنى. */
  function mechanicalFix(title) {
    var t = String(title || '');
    t = t.replace(LEADING_MARKERS, '');
    var withoutDoer = t.replace(DOER_PREFIX, '');
    if (withoutDoer.trim().length >= 3) t = withoutDoer;
    return t.replace(/\s+/g, ' ').replace(/[\s.:؛،,]+$/, '').trim();
  }

  // ───────────── تحليل مهمة واحدة ─────────────

  function analyzeTask(task, row, analysis, cfg) {
    var r = emptyResult(task, 'rules');
    var L = cfg.ai.limits;
    var title = String(task.title || '').trim();
    var fixed = mechanicalFix(title);
    var words = fixed.split(/\s+/).filter(Boolean).length;
    var vague = cfg.ai.vagueStarts.filter(function (v) { return fixed.indexOf(v) === 0 || title.indexOf(v) === 0; })[0];
    var tooShort = words < L.minTitleWords;

    if (fixed && fixed !== title) {
      r.improvedTask = fixed;
      r.improvementReason = 'تنظيف الوصف من الترقيم أو عبارة «القيام بـ» أو علامات الترقيم الزائدة، ليبدأ بالعمل مباشرة.';
    }

    if (tooShort) {
      r.clarity = 'غير واضح';
      r.measurability = 'غير قابل للقياس حاليًا';
      r.recommendations.push('وصف المهمة مختصر جدًا (' + words + ' كلمة). حدّد النشاط وموضوعه ومخرجه.');
      r.question = 'ما المخرج النهائي لهذه المهمة؟';
    } else if (vague) {
      r.clarity = 'يحتاج تحسين';
      r.measurability = 'يحتاج تحسين';
      r.recommendations.push('الصياغة تبدأ بـ «' + vague + '» ولا تحدد نشاطًا يمكن قياسه. حدّد النشاط والمخرج.');
      r.question = 'ما النشاط المحدد الذي يتم تنفيذه، وما مخرجه؟';
    }

    if (!task.frequencyKey) r.missing.push('التردد');
    if (!String(task.repetitions).trim()) r.missing.push('عدد التكرارات في كل فترة');
    if (!String(task.durationMinutes).trim()) r.missing.push('المدة بالدقائق للمرة الواحدة');

    if (row && row.hours !== null) {
      var freq = task.frequencyKey;
      var perPeriod = freq === 'daily' ? 'يوميًا' : freq === 'weekly' ? 'أسبوعيًا' : '';
      if ((freq === 'daily' || freq === 'weekly') && row.durationMinutes > L.maxOccurrenceMinutes) {
        r.valueWarnings.push({ source: 'rules', text: 'مدة المرة الواحدة ' + U.formatNumber(row.durationMinutes / 60, 1) +
          ' ساعة لمهمة تتكرر ' + perPeriod + ' تتجاوز يوم عمل كامل. تحقق: هل المدة للمرة الواحدة أم لمجموع المرات؟' });
      }
      if (freq === 'daily' && row.repetitions > L.maxDailyRepetitions) {
        r.valueWarnings.push({ source: 'rules', text: 'التكرار اليومي ' + row.repetitions +
          ' مرة مرتفع بشكل غير معتاد. تأكد أنه متوسط يوم عمل فعلي وليس مجموع أسبوع أو شهر.' });
      }
      if (analysis.taskCount >= 3 && row.share > L.dominantShare) {
        r.valueWarnings.push({ source: 'rules', text: 'هذه المهمة وحدها تمثل ' + U.formatPercent(row.share, 1) +
          ' من إجمالي الساعات. توجد قيمة مرتفعة مقارنة بباقي المهام، ويُنصح بالتحقق من حجم العمل والزمن الفعلي قبل اعتمادها.' });
      }
    }

    r.quality = tooShort ? 'not_measurable'
      : (vague || r.improvedTask || r.valueWarnings.length || r.missing.length) ? 'needs_improvement' : 'good';
    if (!r.clarity) r.clarity = 'جيد';
    if (!r.measurability) r.measurability = r.missing.length ? 'يحتاج تحسين' : 'جيد';
    return r;
  }

  // ───────────── التداخل (مقارنة نصية) ─────────────

  function tokens(title) {
    return String(title || '')
      .replace(/[ً-ْـ]/g, '')          // التشكيل والتطويل
      .replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي')
      .replace(/[^؀-ۿa-zA-Z0-9\s]/g, ' ')
      .split(/\s+/)
      .map(function (w) { return w.replace(/^(وال|بال|لل|ال|و)(?=\S{3,})/, ''); })
      .map(function (w) { return w.replace(/(ها|هم|ه)$/, function (m, s, i) { return i >= 4 ? '' : m; }); })
      .filter(function (w) { return w.length > 2; });
  }

  function uniq(a) { return Object.keys(a.reduce(function (m, w) { m[w] = 1; return m; }, {})); }

  /**
   * @returns {{overlap: number, jaccard: number}}
   *   overlap = المشترك ÷ أقصر الوصفين (يلتقط «إعداد التقارير» و«إعداد التقارير الشهرية»)
   *   jaccard = المشترك ÷ الاتحاد (للتمييز بين التكرار الكامل والتشابه الجزئي)
   */
  function similarity(a, b) {
    var A = uniq(tokens(a)), B = uniq(tokens(b));
    if (A.length < 2 || B.length < 2) return { overlap: 0, jaccard: 0 };
    var common = A.filter(function (w) { return B.indexOf(w) >= 0; }).length;
    return { overlap: common / Math.min(A.length, B.length), jaccard: common / uniq(A.concat(B)).length };
  }

  /** @returns {Array<{taskIds: string[], type, note, source}>} */
  function findOverlaps(project, onlyTaskId) {
    var out = [];
    var tasks = project.tasks.filter(function (t) { return String(t.title || '').trim(); });
    for (var i = 0; i < tasks.length; i++) {
      for (var j = i + 1; j < tasks.length; j++) {
        if (onlyTaskId && tasks[i].id !== onlyTaskId && tasks[j].id !== onlyTaskId) continue;
        var s = similarity(tasks[i].title, tasks[j].title);
        if (s.overlap >= 0.75) {
          out.push({
            taskIds: [tasks[i].id, tasks[j].id],
            type: s.jaccard >= 0.9 ? 'تكرار كامل' : 'تشابه جزئي',
            note: 'تشابه مرتفع في الصياغة. قد يؤدي احتساب المهمتين معًا إلى تكرار قياس نفس النشاط؛ راجع إمكانية دمجهما إذا كانتا تمثلان نفس النشاط فعليًا.',
            source: 'rules'
          });
        }
      }
    }
    return out;
  }

  /**
   * الفحص المحلي.
   * @param {object} opts { taskId? } لتحليل مهمة واحدة
   */
  function analyze(project, cfg, opts) {
    var only = opts && opts.taskId;
    var analysis = WL.calc.analyze(project, cfg);
    var rows = {};
    analysis.tasks.forEach(function (r) { rows[r.id] = r; });
    var results = project.tasks
      .filter(function (t) { return String(t.title || '').trim() && (!only || t.id === only); })
      .map(function (t) { return analyzeTask(t, rows[t.id], analysis, cfg); });
    return { tasks: results, overlaps: findOverlaps(project, only) };
  }

  WL.aiRules = {
    snapshotOf: snapshotOf,
    emptyResult: emptyResult,
    mechanicalFix: mechanicalFix,
    similarity: similarity,
    findOverlaps: findOverlaps,
    analyze: analyze
  };
})(typeof window !== 'undefined' ? window : globalThis);
