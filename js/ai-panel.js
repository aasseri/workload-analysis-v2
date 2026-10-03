/**
 * ai-panel.js — واجهة المساعد الذكي (لوحة جانبية في صفحة المهام).
 * ستة أزرار: خمسة للمهمة المختارة وزر لمراجعة جميع المهام.
 * لا يُطبَّق أي تعديل إلا بضغط المستخدم على «اعتماد» (الصياغة والتردد فقط)؛
 * التقسيم والدمج والمهام المحتملة توصيات نصية ينفذها المستخدم بنفسه.
 */
(function () {
  'use strict';
  var WL = window.WL;
  var cfg = WL.config, U = WL.utils, M = WL.model, V = WL.validation;
  var e = U.escapeHtml;

  var MODE_LABELS = {
    improve: '✨ تحسين المهمة', analyze: '🔍 تحليل المهمة', measure: '📏 جعلها قابلة للقياس',
    recommend: '💡 اقتراح تحسين', overlap: '🔄 فحص التداخل', review: '🧠 مراجعة جميع المهام'
  };
  var QUALITY = {
    good: { icon: '🟢', label: 'جيدة', cls: 'q-good' },
    needs_improvement: { icon: '🟡', label: 'تحتاج تحسين', cls: 'q-mid' },
    not_measurable: { icon: '🔴', label: 'غير قابلة للقياس حاليًا', cls: 'q-bad' }
  };
  /** الأقسام المعروضة لكل زر (القسم 33: لا تعرض أكثر من الضروري). */
  var SECTIONS = {
    improve: ['current', 'improved', 'measurability', 'missing', 'frequency', 'repsDur', 'recommendations'],
    analyze: ['current', 'dimensions', 'valueWarnings', 'composite', 'recommendations'],
    measure: ['current', 'unit', 'missing', 'question', 'frequency', 'repsDur', 'composite'],
    recommend: ['current', 'recommendations', 'question', 'valueWarnings'],
    overlap: ['current'],
    review: ['improved', 'missing', 'question', 'frequency', 'valueWarnings', 'composite', 'recommendations']
  };

  var ui = { taskId: null, running: null, result: null, editing: null, notes: {} };

  function $(sel, rootEl) { return (rootEl || document).querySelector(sel); }
  function project() { return WL.app.getState().project; }
  function taskIndex(id) { return M.indexOfTask(project(), id); }
  function freqLabel(key) { var f = WL.calc.findFrequency(cfg, key); return f ? f.label : ''; }

  // ───────────── فتح وإغلاق ─────────────

  function open(taskId) {
    document.body.classList.add('ai-open');
    $('#aiDrawer').hidden = false;
    $('#aiBackdrop').hidden = false;
    if (taskId) ui.taskId = taskId;
    renderEngine();
    renderTaskSelect();
    render();
    $('#aiDrawer').focus();
  }

  function close() {
    document.body.classList.remove('ai-open');
    $('#aiDrawer').hidden = true;
    $('#aiBackdrop').hidden = true;
  }

  // ───────────── حالة المحرك وإعدادات المزوّد ─────────────

  function currentConn() {
    var st = WL.ai.getSettings(cfg);
    return st ? WL.ai.resolve(cfg, st) : null;
  }

  function renderEngine() {
    var conn = currentConn();
    $('#aiEngine').innerHTML =
      '<span class="engine-dot ' + (conn ? 'on' : '') + '"></span>' +
      '<span>' + (conn
        ? 'التحليل الذكي <b>مفعّل</b> عبر ' + e(conn.label) + ' <span dir="ltr">(' + e(conn.model) + ')</span> + الفحص المحلي'
        : 'الفحص المحلي فقط (دون إنترنت). فعّل الذكاء الاصطناعي للتحليل الكامل وفق التعليمات.') + '</span>' +
      '<button type="button" class="link-btn" data-ai="toggle-key">' + (conn ? 'الإعدادات' : 'تفعيل الذكاء الاصطناعي') + '</button>';
    $('#aiKeyRemove').hidden = !conn;
  }

  function fillSettingsForm() {
    var saved = WL.ai.getSettings(cfg);
    var sel = $('#aiProvider');
    sel.innerHTML = Object.keys(cfg.ai.providers).map(function (k) {
      return '<option value="' + k + '">' + e(cfg.ai.providers[k].label) + '</option>';
    }).join('');
    sel.value = saved ? saved.provider : cfg.ai.defaultProvider;
    $('#aiKeyInput').value = '';
    $('#aiModelInput').value = saved ? saved.model : '';
    $('#aiBaseUrlInput').value = saved ? saved.baseUrl : '';
    $('#aiKeyError').textContent = '';
    updateProviderFields();
  }

  function updateProviderFields() {
    var key = $('#aiProvider').value;
    var p = cfg.ai.providers[key];
    var saved = WL.ai.getSettings(cfg);
    var keepKey = saved && saved.provider === key;
    $('#aiKeyInput').placeholder = keepKey ? 'محفوظ. اتركه فارغًا للإبقاء عليه' : p.keyHint;
    $('#aiBaseUrlField').hidden = key !== 'custom';
    $('#aiModelInput').placeholder = p.defaultModel || 'اسم النموذج';
    $('#aiModelHint').textContent = p.defaultModel ? 'اتركه فارغًا لاستخدام النموذج الافتراضي.' : 'اكتب اسم النموذج كما يحدده المزوّد.';
    var link = $('#aiKeysLink');
    link.hidden = !p.keysUrl;
    if (p.keysUrl) { link.href = p.keysUrl; link.textContent = 'الحصول على مفتاح من ' + p.label + ' ↗'; }
    $('#aiPrivacy').textContent = 'تُحفظ الإعدادات في هذا المتصفح فقط. عند التحليل الذكي تُرسل نصوص المهام وأرقامها وبيانات الجهة إلى ' +
      (key === 'custom' ? 'المزوّد الذي تحدده' : p.label) + ' عبر الإنترنت.';
  }

  function onKeyInput() {
    var detected = WL.ai.detectProvider($('#aiKeyInput').value);
    if (detected && detected !== $('#aiProvider').value && $('#aiProvider').value !== 'custom') {
      $('#aiProvider').value = detected;
      updateProviderFields();
    }
  }

  function saveKey(ev) {
    ev.preventDefault();
    var err = $('#aiKeyError');
    var provider = $('#aiProvider').value;
    var saved = WL.ai.getSettings(cfg);
    var key = $('#aiKeyInput').value.trim() || (saved && saved.provider === provider ? saved.apiKey : '');
    var settings = { provider: provider, apiKey: key, model: $('#aiModelInput').value, baseUrl: $('#aiBaseUrlInput').value };
    var msg = WL.ai.validateSettings(cfg, settings);
    if (msg) { err.textContent = msg; return; }
    if (!WL.ai.saveSettings(cfg, settings)) { err.textContent = 'تعذر حفظ الإعدادات في هذا المتصفح.'; return; }
    err.textContent = '';
    $('#aiKeyInput').value = '';
    $('#aiKeyForm').hidden = true;
    renderEngine();
    WL.app.toast('تم تفعيل التحليل الذكي عبر ' + WL.ai.resolve(cfg, settings).label + '.', 'success');
  }

  function removeKey() {
    WL.ai.clearSettings(cfg);
    $('#aiKeyForm').hidden = true;
    renderEngine();
    WL.app.toast('تم حذف المفتاح والإعدادات من هذا المتصفح.', 'success');
  }

  // ───────────── اختيار المهمة ─────────────

  function renderTaskSelect() {
    var p = project();
    var sel = $('#aiTaskSelect');
    var withTitle = p.tasks.filter(function (t) { return String(t.title || '').trim(); });
    if (!ui.taskId || taskIndex(ui.taskId) < 0 || !String(p.tasks[taskIndex(ui.taskId)].title || '').trim()) {
      ui.taskId = withTitle.length ? withTitle[0].id : null;
    }
    sel.innerHTML = withTitle.length
      ? p.tasks.map(function (t, i) {
        if (!String(t.title || '').trim()) return '';
        var title = String(t.title).trim();
        return '<option value="' + e(t.id) + '"' + (t.id === ui.taskId ? ' selected' : '') + '>' +
          (i + 1) + ' — ' + e(title.length > 70 ? title.slice(0, 70) + '…' : title) + '</option>';
      }).join('')
      : '<option value="">لا توجد مهام بعد</option>';
    sel.disabled = !withTitle.length;
    Array.prototype.forEach.call(document.querySelectorAll('#aiDrawer [data-mode]'), function (b) {
      b.disabled = !!ui.running || !withTitle.length;
      b.classList.toggle('is-running', ui.running === b.dataset.mode);
    });
  }

  // ───────────── التشغيل ─────────────

  async function runMode(mode) {
    if (ui.running) return;
    var p = project();
    if (mode !== 'review' && (!ui.taskId || taskIndex(ui.taskId) < 0)) {
      WL.app.toast('اختر مهمة أولًا.', 'error');
      return;
    }
    ui.running = mode;
    ui.editing = null;
    ui.notes = {};
    renderTaskSelect();
    render();
    try {
      ui.result = await WL.ai.run(p, cfg, { mode: mode, taskId: ui.taskId });
    } catch (err) {
      console.error(err);
      ui.result = null;
      WL.app.toast(WL.ai.MESSAGES.GENERIC, 'error');
    } finally {
      ui.running = null;
      renderTaskSelect();
      render();
    }
  }

  // ───────────── عرض الأقسام ─────────────

  function taskNo(taskId) { var i = taskIndex(taskId); return i < 0 ? '—' : i + 1; }

  function srcTag(source, r) {
    return source === 'llm'
      ? '<span class="tag tag-ai">اقتراح AI' + (r && r.providerLabel ? ' · ' + e(r.providerLabel) : '') + '</span>'
      : '<span class="tag tag-src">فحص محلي</span>';
  }

  function qualityBadge(q) {
    var d = QUALITY[q];
    return d ? '<span class="qbadge ' + d.cls + '">' + d.icon + ' ' + d.label + '</span>' : '';
  }

  function block(title, inner, extraCls) {
    return '<div class="rsec ' + (extraCls || '') + '"><h4>' + title + '</h4>' + inner + '</div>';
  }

  function list(items) {
    return '<ul class="rlist">' + items.map(function (x) { return '<li>' + e(x) + '</li>'; }).join('') + '</ul>';
  }

  function pendingTitle(r) { return !!r.improvedTask && r.decisions.title === 'pending'; }
  function pendingFreq(r) { return !!r.suggestedFrequencyKey && r.decisions.frequency === 'pending'; }

  var RENDER = {
    current: function (r, t) {
      return block('المهمة الحالية <span class="tag tag-user">بيانات المستخدم</span>', '<p class="rtext">' + e(t.title) + '</p>');
    },
    improved: function (r) {
      if (!r.improvedTask) return '';
      var status = r.decisions.title === 'accepted' ? '<p class="sug-status ok">✓ تم اعتماد الصياغة</p>' : '';
      return block('المهمة المقترحة ' + srcTag(r.improvedSource || r.source, r),
        '<p class="rtext ins">' + e(r.improvedTask) + '</p>' +
        (r.improvementReason ? '<p class="rnote"><b>سبب التعديل:</b> ' + e(r.improvementReason) + '</p>' : '') + status, 'sec-improved');
    },
    measurability: function (r) {
      return r.measurability ? block('قابلية القياس', '<p class="rtext">' + e(r.measurability) + '</p>') : '';
    },
    dimensions: function (r) {
      var rows = [['التصنيف', QUALITY[r.quality] ? QUALITY[r.quality].icon + ' ' + QUALITY[r.quality].label : ''],
        ['وضوح المهمة', r.clarity], ['قابلية القياس', r.measurability], ['وجود مخرج واضح', r.outputClarity],
        ['الارتباط الوظيفي', r.jobRelevance]].filter(function (x) { return x[1]; });
      if (!rows.length) return '';
      return block('تقييم وصفي', '<dl class="dims">' + rows.map(function (x) {
        return '<dt>' + x[0] + '</dt><dd>' + e(x[1]) + '</dd>';
      }).join('') + '</dl>');
    },
    unit: function (r) {
      var inner = '';
      if (r.output) inner += '<p class="rtext"><b>المخرج:</b> ' + e(r.output) + '</p>';
      if (r.unit) inner += '<p class="rtext"><b>وحدة القياس:</b> ' + e(r.unit) + '</p>';
      return inner ? block('طريقة القياس', inner) : '';
    },
    missing: function (r) {
      return r.missing.length ? block('ما يحتاج إلى استكمال', list(r.missing), 'sec-missing') : '';
    },
    question: function (r) {
      return r.question ? block('سؤال', '<p class="rtext q">❓ ' + e(r.question) + '</p>') : '';
    },
    frequency: function (r, t) {
      if (!r.suggestedFrequencyKey) return '';
      var status = r.decisions.frequency === 'accepted' ? '<p class="sug-status ok">✓ تم اعتماد التردد</p>' : '';
      return block('التردد المقترح ' + srcTag('llm', r),
        '<p class="rtext"><del>' + e(freqLabel(r.snapshot.frequencyKey) || 'غير محدد') + '</del> ← <ins>' +
        e(freqLabel(r.suggestedFrequencyKey)) + '</ins>' +
        (r.frequencyConfidence ? ' <span class="conf">الثقة: ' + e(r.frequencyConfidence) + '</span>' : '') + '</p>' +
        (r.frequencyReason ? '<p class="rnote"><b>السبب:</b> ' + e(r.frequencyReason) + '</p>' : '') +
        '<p class="rnote muted">تردد مقترح وليس قيمة معتمدة؛ يُعتمد التردد الفعلي من المستخدم.</p>' + status);
    },
    repsDur: function (r, t) {
      function val(v, unit) {
        return String(v).trim() ? '<b>' + e(v) + '</b> ' + unit + ' <span class="tag tag-user">بيانات المستخدم</span>'
          : '<span class="need">يحتاج إدخال المستخدم</span>';
      }
      return block('التكرار والمدة',
        '<p class="rtext">التكرار: ' + val(t.repetitions, '') + '</p>' +
        '<p class="rtext">المدة: ' + val(t.durationMinutes, 'دقيقة') + '</p>');
    },
    valueWarnings: function (r) {
      if (!r.valueWarnings.length) return '';
      return block('⚠️ ملاحظات تحقق على القيم', '<ul class="rlist warn">' + r.valueWarnings.map(function (w) {
        return '<li>' + e(w.text) + ' ' + srcTag(w.source, r) + '</li>';
      }).join('') + '</ul>');
    },
    composite: function (r) {
      if (!r.composite) return '';
      return block('مهمة مركبة', '<p class="rtext">تبدو المهمة مركبة من عدة أنشطة يمكن أن تؤثر على دقة قياس الزمن.</p>' +
        (r.splitSuggestions.length ? '<p class="rnote"><b>فصل مقترح (توصية):</b></p><ol class="rlist">' +
          r.splitSuggestions.map(function (x) { return '<li>' + e(x) + '</li>'; }).join('') + '</ol>' : '') +
        '<p class="rnote muted">لا يتم الفصل تلقائيًا؛ عدّل المهام من الجدول إذا وافقت.</p>');
    },
    recommendations: function (r) {
      return r.recommendations.length ? block('التوصية', list(r.recommendations)) : '';
    }
  };

  function actionsHtml(r) {
    var p = project();
    var staleT = pendingTitle(r) && WL.ai.isStale(p, r, 'title');
    var staleF = pendingFreq(r) && WL.ai.isStale(p, r, 'frequencyKey');
    if (staleT || staleF) return '<p class="sug-status warn">تغيّرت المهمة بعد التحليل. أعد التحليل للحصول على اقتراح محدّث.</p>';
    var btns = '';
    if (pendingTitle(r)) btns += '<button type="button" class="btn btn-primary btn-sm" data-ai="accept-title">اعتماد الصياغة</button>';
    if (pendingFreq(r)) btns += '<button type="button" class="btn btn-primary btn-sm" data-ai="accept-freq">اعتماد التردد المقترح</button>';
    btns += '<button type="button" class="btn btn-outline btn-sm" data-ai="edit">تعديل</button>';
    btns += '<button type="button" class="btn btn-ghost btn-sm" data-ai="dismiss">إلغاء</button>';
    return '<div class="sug-actions">' + btns + '</div>';
  }

  function editHtml(r, t) {
    var title = pendingTitle(r) ? r.improvedTask : t.title;
    var fk = pendingFreq(r) ? r.suggestedFrequencyKey : t.frequencyKey;
    return '<div class="sug-edit">' +
      '<label for="ed-title-' + r.id + '">وصف المهمة</label>' +
      '<textarea id="ed-title-' + r.id + '" data-edit="title" rows="3" maxlength="' + cfg.ui.maxTaskTitleLength + '">' + e(title) + '</textarea>' +
      '<label for="ed-freq-' + r.id + '">التردد</label>' +
      '<select id="ed-freq-' + r.id + '" data-edit="frequencyKey"><option value="">اختر التردد</option>' +
      cfg.frequencies.map(function (f) {
        return '<option value="' + f.key + '"' + (f.key === fk ? ' selected' : '') + '>' + e(f.label) + '</option>';
      }).join('') + '</select>' +
      '<small class="err" data-edit-err></small>' +
      '<div class="sug-actions"><button type="button" class="btn btn-primary btn-sm" data-ai="apply-edit">اعتماد</button>' +
      '<button type="button" class="btn btn-outline btn-sm" data-ai="cancel-edit">إلغاء</button></div></div>';
  }

  function cardHtml(r, mode) {
    var p = project();
    var i = taskIndex(r.taskId);
    if (i < 0 || r.dismissed) return '';
    var t = p.tasks[i];
    var body = SECTIONS[mode].map(function (s) { return RENDER[s](r, t); }).join('');
    if (mode === 'improve' && !r.improvedTask) {
      body = RENDER.current(r, t) + '<p class="rnote">الصياغة الحالية مناسبة ولا تحتاج تعديلًا.</p>' +
        SECTIONS.improve.slice(2).map(function (s) { return RENDER[s](r, t); }).join('');
    }
    if (mode === 'review' && !body) body = '<p class="rnote">المهمة واضحة وقابلة للقياس. يمكنك استكمال التردد والتكرار والزمن.</p>';
    if (mode !== 'review' && mode !== 'overlap' && body === RENDER.current(r, t)) {
      body += '<p class="rnote">' + (r.source === 'rules'
        ? 'لا توجد ملاحظات من الفحص المحلي لهذا الزر. فعّل الذكاء الاصطناعي للحصول على تحليل أعمق.'
        : 'المهمة واضحة وقابلة للقياس. يمكنك استكمال التردد والتكرار والزمن.') + '</p>';
    }
    var note = ui.notes[r.id] ? '<p class="sug-status warn">' + e(ui.notes[r.id]) + '</p>' : '';
    return '<article class="sug rcard ' + (QUALITY[r.quality] ? QUALITY[r.quality].cls : '') + '" data-res="' + r.id + '">' +
      '<header class="sug-head">' +
        '<button type="button" class="sug-task" data-ai="goto">المهمة ' + (i + 1) + '</button>' +
        qualityBadge(r.quality) + srcTag(r.source, r) +
      '</header>' +
      (mode === 'review' ? '<p class="sug-title">' + e(t.title) + '</p>' : '') +
      body + note +
      (ui.editing === r.id ? editHtml(r, t) : actionsHtml(r)) +
      '</article>';
  }

  function overlapsHtml(overlaps, title) {
    var items = overlaps.filter(function (o) { return o.taskIds.every(function (id) { return taskIndex(id) >= 0; }); });
    if (!items.length) return '';
    return '<section class="rbox"><h4>' + title + '</h4><ul class="ovl">' + items.map(function (o) {
      return '<li><div class="ovl-head">' + o.taskIds.map(function (id) {
        return '<button type="button" class="sug-task" data-ai="goto-id" data-id="' + e(id) + '">المهمة ' + taskNo(id) + '</button>';
      }).join(' ↔ ') + ' <span class="tag tag-warn">' + e(o.type) + '</span> ' + srcTag(o.source) + '</div>' +
        (o.note ? '<p class="rnote">' + e(o.note) + '</p>' : '') + '</li>';
    }).join('') + '</ul><p class="rnote muted">لا يتم الدمج أو الحذف تلقائيًا؛ راجع المهام من الجدول.</p></section>';
  }

  function reportHtml(res) {
    var counts = { good: 0, needs_improvement: 0, not_measurable: 0 };
    res.tasks.forEach(function (r) { if (counts[r.quality] !== undefined) counts[r.quality] += 1; });
    var top = res.topRecommendations.length ? '<h5>أهم التوصيات</h5><ol class="rlist">' +
      res.topRecommendations.map(function (x) { return '<li>' + e(x) + '</li>'; }).join('') + '</ol>' : '';
    var rows = res.tasks.map(function (r) {
      var problem = r.missing[0] || (r.valueWarnings[0] && r.valueWarnings[0].text) ||
        (r.improvedTask ? 'الصياغة تحتاج تحسين' : '') ||
        (r.quality !== 'good' && r.measurability !== 'جيد' ? 'قابلية القياس: ' + r.measurability : '') || '—';
      var rec = r.recommendations[0] || r.question || (r.improvedTask ? 'مراجعة الصياغة المقترحة' : '—');
      return '<tr data-jump="' + r.id + '"><td>' + taskNo(r.taskId) + '</td><td>' +
        (QUALITY[r.quality] ? QUALITY[r.quality].icon : '') + '</td><td>' + e(problem) + '</td><td>' + e(rec) + '</td></tr>';
    }).join('');
    return '<section class="rbox report"><h4>ملخص جودة البيانات</h4>' +
      '<div class="rstats">' +
        '<div><b>' + res.tasks.length + '</b><span>عدد المهام</span></div>' +
        '<div class="q-good"><b>' + counts.good + '</b><span>🟢 جيدة</span></div>' +
        '<div class="q-mid"><b>' + counts.needs_improvement + '</b><span>🟡 تحتاج تحسين</span></div>' +
        '<div class="q-bad"><b>' + counts.not_measurable + '</b><span>🔴 غير قابلة للقياس حاليًا</span></div>' +
        '<div><b>' + res.overlaps.length + '</b><span>التداخل المحتمل</span></div>' +
      '</div>' + top +
      '<table class="mini rtable"><thead><tr><th>المهمة</th><th>الجودة</th><th>المشكلة</th><th>التوصية</th></tr></thead><tbody>' +
      rows + '</tbody></table></section>';
  }

  function render() {
    var body = $('#aiBody');
    if (ui.running) {
      body.innerHTML = '<div class="ai-loading"><span class="spinner"></span><p>' + e(MODE_LABELS[ui.running]) + '…</p>' +
        (WL.ai.isEnabled(cfg) ? '<small>التحليل الذكي قد يستغرق حتى دقيقة.</small>' : '') + '</div>';
      return;
    }
    var res = ui.result;
    if (!res) {
      body.innerHTML = '<div class="ai-empty"><p>اختر مهمة ثم أحد الأزرار، أو «مراجعة جميع المهام» لتقرير جودة القائمة كاملة.</p>' +
        '<p class="muted">لن يتغير أي شيء في المهام إلا بعد اعتمادك.</p></div>';
      return;
    }

    var html = '<div class="rmode">' + e(MODE_LABELS[res.mode]) +
      (res.mode !== 'review' && res.taskId ? ' — المهمة ' + taskNo(res.taskId) : '') + '</div>';
    if (res.llmError) html += '<div class="ai-banner warn">' + e(res.llmError) + ' عُرضت نتائج الفحص المحلي فقط.</div>';
    else if (!res.llmUsed) html += '<div class="ai-banner info">نتيجة الفحص المحلي. فعّل الذكاء الاصطناعي للتحليل الكامل وفق التعليمات.</div>';

    if (res.mode === 'review') {
      html += reportHtml(res);
      html += overlapsHtml(res.overlaps, '🔄 التداخل المحتمل');
      if (res.missingTasks.length) {
        html += '<section class="rbox"><h4>مهام محتملة للمراجعة</h4>' + list(res.missingTasks) +
          '<p class="rnote muted">توصيات فقط؛ لا تُضاف أي مهمة تلقائيًا.</p></section>';
      }
      var cards = res.tasks.filter(function (r) { return r.quality !== 'good' || r.improvedTask || r.valueWarnings.length; })
        .map(function (r) { return cardHtml(r, 'review'); }).join('');
      html += cards ? '<h4 class="rsub">تفاصيل المهام التي تحتاج مراجعة</h4>' + cards : '';
    } else {
      html += res.tasks.map(function (r) { return cardHtml(r, res.mode); }).join('');
      if (res.mode === 'overlap') {
        html += overlapsHtml(res.overlaps, 'التداخل مع المهام الأخرى') ||
          '<div class="ai-empty ok"><p>✓ لا يوجد تداخل ظاهر مع المهام الأخرى.</p></div>';
      }
    }
    body.innerHTML = html;
  }

  // ───────────── القرارات ─────────────

  function findRes(id) {
    return ui.result ? ui.result.tasks.filter(function (r) { return r.id === id; })[0] : null;
  }

  /** يطبق التعديل بعد التحقق منه بنفس قواعد النظام. */
  function applyChanges(r, changes) {
    var p = project();
    var i = taskIndex(r.taskId);
    if (i < 0) return 'المهمة لم تعد موجودة.';
    var candidate = U.deepClone(p.tasks[i]);
    Object.keys(changes).forEach(function (k) { candidate[k] = changes[k]; });
    var errs = V.validateTask(candidate, i, cfg).filter(function (er) { return changes.hasOwnProperty(er.field); });
    if (errs.length) return errs[0].short;
    M.updateTask(p, r.taskId, changes);
    Object.keys(changes).forEach(function (k) { r.snapshot[k] = String(changes[k]); });
    WL.app.renderTasks();
    return null;
  }

  function flashRow(taskId) {
    var row = document.querySelector('#taskBody tr[data-id="' + taskId + '"]');
    if (!row) return;
    row.scrollIntoView({ block: 'center', behavior: 'smooth' });
    row.classList.remove('flash'); void row.offsetWidth; row.classList.add('flash');
  }

  function onClick(ev) {
    var modeBtn = ev.target.closest('[data-mode]');
    if (modeBtn) { if (!modeBtn.disabled) runMode(modeBtn.dataset.mode); return; }
    var jump = ev.target.closest('[data-jump]');
    if (jump) {
      var target = document.querySelector('[data-res="' + jump.dataset.jump + '"]');
      if (target) target.scrollIntoView({ block: 'start', behavior: 'smooth' });
      return;
    }
    var btn = ev.target.closest('[data-ai]');
    if (!btn) return;
    var card = btn.closest('[data-res]');
    var r = card ? findRes(card.dataset.res) : null;
    var err;

    switch (btn.dataset.ai) {
      case 'close': close(); break;
      case 'toggle-key':
        $('#aiKeyForm').hidden = !$('#aiKeyForm').hidden;
        if (!$('#aiKeyForm').hidden) { fillSettingsForm(); $('#aiProvider').focus(); }
        break;
      case 'remove-key': removeKey(); break;
      case 'goto': if (r) flashRow(r.taskId); break;
      case 'goto-id': flashRow(btn.dataset.id); break;
      case 'accept-title':
        if (WL.ai.isStale(project(), r, 'title')) { render(); break; }
        err = applyChanges(r, { title: r.improvedTask });
        if (err) ui.notes[r.id] = err;
        else { r.decisions.title = 'accepted'; delete ui.notes[r.id]; WL.app.toast('تم اعتماد الصياغة.', 'success'); }
        render();
        break;
      case 'accept-freq':
        if (WL.ai.isStale(project(), r, 'frequencyKey')) { render(); break; }
        err = applyChanges(r, { frequencyKey: r.suggestedFrequencyKey });
        if (err) ui.notes[r.id] = err;
        else { r.decisions.frequency = 'accepted'; delete ui.notes[r.id]; WL.app.toast('تم اعتماد التردد.', 'success'); }
        render();
        break;
      case 'edit': ui.editing = r.id; render(); focusEdit(r.id); break;
      case 'cancel-edit': ui.editing = null; render(); break;
      case 'apply-edit': {
        var changes = {};
        Array.prototype.forEach.call(card.querySelectorAll('[data-edit]'), function (c) {
          changes[c.dataset.edit] = c.dataset.edit === 'title' ? c.value.trim() : c.value;
        });
        err = applyChanges(r, changes);
        if (err) { card.querySelector('[data-edit-err]').textContent = err; break; }
        ui.editing = null;
        if (r.improvedTask) r.decisions.title = 'accepted';
        if (r.suggestedFrequencyKey) r.decisions.frequency = 'accepted';
        WL.app.toast('تم تعديل المهمة.', 'success');
        render();
        break;
      }
      case 'dismiss': r.dismissed = true; render(); break;
    }
  }

  function focusEdit(id) {
    var el = document.querySelector('[data-res="' + id + '"] [data-edit]');
    if (el) el.focus();
  }

  /** يحدّث اللوحة بعد تعديل المهام من الجدول (أرقام المهام، قائمة الاختيار، حالة الاقتراحات القديمة). */
  function refreshIfIdle() {
    var d = document.getElementById('aiDrawer');
    if (d && !d.hidden && !ui.running && !ui.editing) {
      renderTaskSelect();
      render();
    }
  }

  function init() {
    $('#aiDrawer').addEventListener('click', onClick);
    $('#aiBackdrop').addEventListener('click', close);
    $('#aiKeyForm').addEventListener('submit', saveKey);
    $('#aiProvider').addEventListener('change', updateProviderFields);
    $('#aiKeyInput').addEventListener('input', onKeyInput);
    $('#aiTaskSelect').addEventListener('change', function (ev) { ui.taskId = ev.target.value; });
    document.addEventListener('keydown', function (ev) {
      if (ev.key === 'Escape' && !$('#aiDrawer').hidden && !document.querySelector('dialog[open]')) close();
    });
  }

  WL.aiPanel = { open: open, close: close, runMode: runMode, render: render, refreshIfIdle: refreshIfIdle };
  document.addEventListener('DOMContentLoaded', init);
})();
