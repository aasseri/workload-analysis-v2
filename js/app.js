/**
 * app.js — طبقة الواجهة: عرض البيانات، استقبال الإدخال، والتنقل بين المراحل.
 * لا تحتوي على منطق حساب أو تحقق أو بناء Excel؛ تستدعي الوحدات المختصة فقط.
 */
(function () {
  'use strict';
  var WL = window.WL;
  var cfg = WL.config, U = WL.utils, M = WL.model, V = WL.validation, Calc = WL.calc;

  var LAST_STEP = 2;
  var state = {
    project: null,
    step: 0,
    showAllErrors: false,
    tasksGateShown: false,   // حاول المستخدم الانتقال للمراجعة والمهام ناقصة: تظهر رسالة النواقص تحت الجدول
    touched: {},
    analysis: null,
    validation: null
  };

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  // ───────────── بدء فارغ دائمًا ─────────────
  // النظام يفتح دون بيانات في كل مرة (لا مسودة تلقائية). للاحتفاظ بالعمل: «حفظ المشروع» من القائمة.

  /** يحذف المسودة التي كانت تحفظها الإصدارات السابقة في المتصفح. */
  function clearLegacyDraft() {
    try { localStorage.removeItem(cfg.app.legacyDraftKey); } catch (e) { /* التخزين غير متاح */ }
  }

  function projectHasData(p) {
    var o = p.org;
    return p.tasks.length > 0 || o.sections.length > 0 || !!(o.entity || o.department || String(o.actualCount).trim());
  }

  /** تنبيه المتصفح عند إغلاق الصفحة أو تحديثها وفيها بيانات غير محفوظة. */
  function warnBeforeLeaving(e) {
    if (!projectHasData(state.project)) return;
    e.preventDefault();
    e.returnValue = '';
  }

  // ───────────── التنسيق للعرض ─────────────

  var needDigits = cfg.needRounding.digits;
  function fmtHours(v) { return U.formatNumber(v, 2); }
  function fmtNeed(v) { return U.formatNumber(v, needDigits); }
  function fmtGap(v) {
    if (!U.isValidNumber(v)) return '—';
    var s = U.formatNumber(Math.abs(v), needDigits);
    return v > 0 ? '+' + s : v < 0 ? '−' + s : s;
  }

  // ───────────── إشعارات وحوار التأكيد ─────────────

  function toast(message, kind) {
    // الرسالة نفسها ظاهرة بالفعل (نقرات متكررة): لا تُكرر
    if ($$('#toasts .toast').some(function (t) { return t.textContent === message; })) return;
    var el = document.createElement('div');
    el.className = 'toast ' + (kind || '');
    el.textContent = message;
    $('#toasts').appendChild(el);
    setTimeout(function () { el.remove(); }, 3800);
  }

  function confirmDialog(title, text, okLabel) {
    var dlg = $('#confirmDialog');
    $('#confirmTitle').textContent = title;
    $('#confirmText').textContent = text;
    $('#confirmOk').textContent = okLabel || 'تأكيد';
    return new Promise(function (resolve) {
      function finish(ok) {
        dlg.removeEventListener('click', onClick);
        dlg.removeEventListener('cancel', onCancel);
        if (dlg.open) dlg.close();
        resolve(ok);
      }
      function onClick(e) {
        var btn = e.target.closest('[data-confirm]');
        if (btn) finish(btn.dataset.confirm === 'ok');
      }
      function onCancel(e) { e.preventDefault(); finish(false); } // مفتاح Esc
      dlg.addEventListener('click', onClick);
      dlg.addEventListener('cancel', onCancel);
      dlg.showModal();
      $('[data-confirm="cancel"]', dlg).focus();
    });
  }

  // ───────────── البيانات الأساسية ─────────────

  function fillOrgInputs() {
    $$('[data-org]').forEach(function (input) {
      var v = state.project.org[input.dataset.org];
      input.value = v === null || v === undefined ? '' : v;
    });
    $('#f-newSection').value = '';
    renderSections();
  }

  // ───────────── الأقسام / الوحدات (تُعرّف هنا وتُختار لكل مهمة) ─────────────

  function renderSections() {
    var p = state.project;
    $('#sectionList').innerHTML = p.org.sections.map(function (s) {
      var n = M.tasksInSection(p, s.id);
      return '<li data-sec-id="' + U.escapeHtml(s.id) + '">' +
        '<input type="text" class="sec-name" value="' + U.escapeHtml(s.name) + '" aria-label="اسم القسم" autocomplete="off">' +
        '<span class="sec-count">' + (n ? tasksCountText(n) : 'لا مهام') + '</span>' +
        '<button type="button" class="icon-btn danger" data-sec-del title="حذف القسم" aria-label="حذف القسم"><svg class="ic"><use href="#i-trash"/></svg></button>' +
        '</li>';
    }).join('');
  }

  function addSectionFromInput() {
    var input = $('#f-newSection');
    var res = M.addSection(state.project, input.value);
    if (res.error) { toast(res.error, 'error'); input.focus(); return; }
    // أول قسم يُعرّف: يُربط تلقائيًا بالمهام التي ليس لها قسم
    if (state.project.org.sections.length === 1) {
      state.project.tasks.forEach(function (t) { if (!t.sectionId) t.sectionId = res.section.id; });
    }
    input.value = '';
    input.focus();
    renderSections();
    renderTasks();
  }

  function bindSections() {
    $('#f-newSection').addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); addSectionFromInput(); }
    });
    $('#sectionList').addEventListener('change', function (e) {
      if (!e.target.classList.contains('sec-name')) return;
      var id = e.target.closest('[data-sec-id]').dataset.secId;
      var res = M.renameSection(state.project, id, e.target.value);
      if (res.error) { toast(res.error, 'error'); e.target.value = M.findSection(state.project, id).name; return; }
      e.target.value = res.section.name;
      renderTasks();
    });
    $('#sectionList').addEventListener('click', async function (e) {
      if (!e.target.closest('[data-sec-del]')) return;
      var id = e.target.closest('[data-sec-id]').dataset.secId;
      var s = M.findSection(state.project, id);
      var n = M.tasksInSection(state.project, id);
      if (n) {
        var ok = await confirmDialog('حذف القسم', 'القسم «' + s.name + '» مرتبط بـ ' + tasksCountText(n) +
          '. عند حذفه يجب اختيار قسم آخر لهذه المهام. هل تريد المتابعة؟', 'حذف');
        if (!ok) return;
      }
      M.removeSection(state.project, id);
      renderSections();
      renderTasks();
    });
  }

  function sectionOptions(selected) {
    var html = '<option value="">اختر القسم / الوحدة</option>';
    state.project.org.sections.forEach(function (s) {
      html += '<option value="' + U.escapeHtml(s.id) + '"' + (s.id === selected ? ' selected' : '') + '>' + U.escapeHtml(s.name) + '</option>';
    });
    return html;
  }

  function bindOrgInputs() {
    $$('[data-org]').forEach(function (input) {
      input.addEventListener('input', function () {
        state.project.org[input.dataset.org] = input.value;
        refresh();
      });
      input.addEventListener('blur', function () {
        state.touched['org:' + input.dataset.org] = true;
        refresh();
      });
    });
  }

  // ───────────── جدول المهام ─────────────

  function frequencyOptions(selected) {
    var html = '<option value="">اختر التردد</option>';
    cfg.frequencies.forEach(function (f) {
      html += '<option value="' + f.key + '"' + (f.key === selected ? ' selected' : '') + '>' + U.escapeHtml(f.label) + '</option>';
    });
    return html;
  }

  function taskRowHtml(t, i, count) {
    var e = U.escapeHtml;
    return '' +
      '<tr data-id="' + e(t.id) + '">' +
        '<td class="c-order"><div class="order-tools">' +
          '<button type="button" class="icon-btn" data-row-action="up" title="تحريك للأعلى" aria-label="تحريك للأعلى"' + (i === 0 ? ' disabled' : '') + '><svg class="ic"><use href="#i-up"/></svg></button>' +
          '<span class="icon-btn grip" data-grip title="اسحب لإعادة الترتيب"><svg class="ic"><use href="#i-grip"/></svg></span>' +
          '<button type="button" class="icon-btn" data-row-action="down" title="تحريك للأسفل" aria-label="تحريك للأسفل"' + (i === count - 1 ? ' disabled' : '') + '><svg class="ic"><use href="#i-down"/></svg></button>' +
        '</div></td>' +
        '<td class="c-no"><span class="row-no">' + (i + 1) + '</span></td>' +
        '<td class="c-title" data-label="المهمة ' + (i + 1) + '">' +
          '<textarea rows="1" data-task-field="title" maxlength="' + cfg.ui.maxTaskTitleLength + '" placeholder="اكتب وصف المهمة" aria-label="وصف المهمة">' + e(t.title) + '</textarea>' +
          '<small class="err" data-err="title"></small></td>' +
        '<td class="c-sec" data-label="القسم / الوحدة">' +
          '<select data-task-field="sectionId" aria-label="القسم / الوحدة">' + sectionOptions(t.sectionId) + '</select>' +
          '<small class="err" data-err="sectionId"></small></td>' +
        '<td class="c-pos" data-label="المسمى الفعلي">' +
          '<input type="text" list="positionList" data-task-field="positionTitle" value="' + e(t.positionTitle) + '" maxlength="120" placeholder="مثال: مساعد إداري" aria-label="المسمى الفعلي" autocomplete="off">' +
          '<small class="err" data-err="positionTitle"></small></td>' +
        '<td class="c-freq" data-label="التردد">' +
          '<select data-task-field="frequencyKey" aria-label="التردد">' + frequencyOptions(t.frequencyKey) + '</select>' +
          '<small class="err" data-err="frequencyKey"></small></td>' +
        '<td class="c-num" data-label="التكرار">' +
          '<input type="text" class="num" inputmode="decimal" data-task-field="repetitions" value="' + e(t.repetitions) + '" aria-label="عدد التكرارات">' +
          '<small class="err" data-err="repetitions"></small></td>' +
        '<td class="c-num" data-label="المدة بالدقائق">' +
          '<input type="text" class="num" inputmode="decimal" data-task-field="durationMinutes" value="' + e(t.durationMinutes) + '" aria-label="المدة بالدقائق">' +
          '<small class="err" data-err="durationMinutes"></small></td>' +
        '<td class="c-num" data-label="إجمالي الساعات"><span class="calc hours" data-out="hours"></span></td>' +
        '<td class="c-share" data-label="نسبة الساعات"><span class="calc" data-out="share"></span><div class="share-bar"><i data-out="bar"></i></div></td>' +
        '<td class="c-act"><div class="row-actions">' +
          '<button type="button" class="icon-btn ai" data-row-action="ai" title="فتح المساعد الذكي لهذه المهمة" aria-label="فتح المساعد الذكي لهذه المهمة"><svg class="ic"><use href="#i-spark"/></svg></button>' +
          '<button type="button" class="icon-btn" data-row-action="duplicate" title="نسخ المهمة" aria-label="نسخ المهمة"><svg class="ic"><use href="#i-copy"/></svg></button>' +
          '<button type="button" class="icon-btn danger" data-row-action="delete" title="حذف المهمة" aria-label="حذف المهمة"><svg class="ic"><use href="#i-trash"/></svg></button>' +
        '</div></td>' +
      '</tr>';
  }

  function autoGrow(textarea) {
    textarea.style.height = 'auto';
    textarea.style.height = textarea.scrollHeight + 2 + 'px';
  }

  function renderTasks(focusId) {
    var tasks = state.project.tasks;
    var body = $('#taskBody');
    body.innerHTML = tasks.map(function (t, i) { return taskRowHtml(t, i, tasks.length); }).join('');
    $('#tasksEmpty').hidden = tasks.length > 0;
    $('#tasksTable thead').hidden = tasks.length === 0;
    $$('textarea', body).forEach(autoGrow);
    refreshPositionList();
    refresh();
    if (focusId) {
      var row = body.querySelector('tr[data-id="' + focusId + '"]');
      if (row) {
        row.classList.add('flash');
        var ta = row.querySelector('textarea');
        ta.focus();
        row.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      }
    }
  }

  function rowId(el) {
    var tr = el.closest('tr[data-id]');
    return tr ? tr.dataset.id : null;
  }

  // ───────────── المسمى الفعلي: توحيد الكتابة وقائمة الاقتراحات ─────────────

  /** المسميات المستخدمة بترتيب ظهورها (دون تكرار). */
  function usedPositions(exceptId) {
    var seen = {}, list = [];
    state.project.tasks.forEach(function (t) {
      var v = WL.calc.positionKey(t.positionTitle);
      if (!v || t.id === exceptId || seen[v]) return;
      seen[v] = true;
      list.push(v);
    });
    return list;
  }

  function refreshPositionList() {
    $('#positionList').innerHTML = usedPositions().map(function (v) {
      return '<option value="' + U.escapeHtml(v) + '"></option>';
    }).join('');
  }

  /**
   * يحذف المسافات الزائدة، ويوحّد الكتابة مع مسمى موجود يطابقه بعد تجاهل الهمزات والتاء المربوطة والتشكيل،
   * حتى لا يُحسب «مساعد اداري» و«مساعد إداري» مسميين مختلفين في ورقة الاحتياج حسب المسمى.
   */
  function unifyPosition(id, input) {
    var raw = WL.calc.positionKey(input.value);
    var key = WL.taskImport.normalizeArabic(raw);
    var match = key ? usedPositions(id).filter(function (v) {
      return v !== raw && WL.taskImport.normalizeArabic(v) === key;
    })[0] : null;
    var value = match || raw;
    if (value !== input.value) {
      input.value = value;
      M.updateTask(state.project, id, { positionTitle: value });
      if (match) toast('تم توحيد كتابة المسمى مع «' + match + '».', 'success');
    }
    refreshPositionList();
  }

  function addTask() {
    var t = M.addTask(state.project);
    renderTasks(t.id);
  }

  async function handleRowAction(action, id) {
    var p = state.project;
    if (action === 'up' || action === 'down') {
      if (M.moveTask(p, id, action === 'up' ? -1 : 1)) renderTasks();
      var btn = $('#taskBody tr[data-id="' + id + '"] [data-row-action="' + action + '"]');
      if (btn && !btn.disabled) btn.focus();
    } else if (action === 'ai') {
      WL.aiPanel.open(id);
    } else if (action === 'duplicate') {
      var copy = M.duplicateTask(p, id);
      renderTasks(copy.id);
      toast('تم نسخ المهمة.', 'success');
    } else if (action === 'delete') {
      var task = p.tasks[M.indexOfTask(p, id)];
      var hasContent = task && (String(task.title).trim() || task.frequencyKey || String(task.repetitions).trim() || String(task.durationMinutes).trim());
      if (hasContent) {
        var ok = await confirmDialog('حذف المهمة', 'سيتم حذف المهمة «' + (String(task.title).trim() || 'بدون وصف') + '». هل تريد المتابعة؟', 'حذف');
        if (!ok) return;
      }
      M.removeTask(p, id);
      renderTasks();
    }
  }

  function bindTaskTable() {
    var body = $('#taskBody');

    body.addEventListener('input', function (e) {
      var field = e.target.dataset.taskField;
      if (!field) return;
      var change = {};
      change[field] = e.target.value;
      M.updateTask(state.project, rowId(e.target), change);
      if (e.target.tagName === 'TEXTAREA') autoGrow(e.target);
      refresh();
    });

    body.addEventListener('focusout', function (e) {
      var field = e.target.dataset && e.target.dataset.taskField;
      if (!field) return;
      if (field === 'positionTitle') unifyPosition(rowId(e.target), e.target);
      state.touched[rowId(e.target) + ':' + field] = true;
      refresh();
    });

    body.addEventListener('click', function (e) {
      var btn = e.target.closest('[data-row-action]');
      if (btn) handleRowAction(btn.dataset.rowAction, rowId(btn));
    });

    // السحب والإفلات لإعادة الترتيب (من المقبض فقط)
    var dragId = null;
    function clearDropMarks() { $$('.drop-before, .drop-after', body).forEach(function (r) { r.classList.remove('drop-before', 'drop-after'); }); }

    body.addEventListener('mousedown', function (e) {
      if (e.target.closest('[data-grip]')) e.target.closest('tr').draggable = true;
    });
    body.addEventListener('dragstart', function (e) {
      var tr = e.target.closest('tr[data-id]');
      if (!tr || !tr.draggable) return;
      dragId = tr.dataset.id;
      tr.classList.add('is-dragging');
      e.dataTransfer.effectAllowed = 'move';
      try { e.dataTransfer.setData('text/plain', dragId); } catch (err) { /* بعض المتصفحات */ }
    });
    body.addEventListener('dragover', function (e) {
      if (!dragId) return;
      var tr = e.target.closest('tr[data-id]');
      if (!tr) return;
      e.preventDefault();
      clearDropMarks();
      if (tr.dataset.id === dragId) return;
      var rect = tr.getBoundingClientRect();
      tr.classList.add(e.clientY < rect.top + rect.height / 2 ? 'drop-before' : 'drop-after');
    });
    body.addEventListener('drop', function (e) {
      if (!dragId) return;
      e.preventDefault();
      var tr = e.target.closest('tr[data-id]');
      if (tr && tr.dataset.id !== dragId) {
        var p = state.project;
        var from = M.indexOfTask(p, dragId);
        var to = M.indexOfTask(p, tr.dataset.id) + (tr.classList.contains('drop-after') ? 1 : 0);
        if (from < to) to -= 1;
        if (M.moveTaskTo(p, dragId, to)) renderTasks();
      }
      clearDropMarks();
    });
    body.addEventListener('dragend', function () {
      $$('tr', body).forEach(function (r) { r.draggable = false; r.classList.remove('is-dragging'); });
      clearDropMarks();
      dragId = null;
    });
  }

  // ───────────── التحديث الحي للنتائج والأخطاء ─────────────

  function errorVisible(key) { return state.showAllErrors || !!state.touched[key]; }

  function applyFieldError(control, errEl, message) {
    control.classList.toggle('is-invalid', !!message);
    control.setAttribute('aria-invalid', message ? 'true' : 'false');
    if (errEl) errEl.textContent = message || '';
  }

  function refresh() {
    var p = state.project;
    var a = (state.analysis = Calc.analyze(p, cfg));
    var v = (state.validation = V.validateProject(p, cfg));
    var errs = V.indexErrors(v.errors);

    // أخطاء البيانات الأساسية
    $$('[data-org]').forEach(function (input) {
      var key = 'org:' + input.dataset.org;
      applyFieldError(input, input.parentNode.querySelector('.err'), errorVisible(key) ? errs[key] : '');
    });
    applyFieldError($('#f-newSection'), $('#sectionsErr'), errorVisible('org:sections') ? errs['org:sections'] : '');
    // عدد المهام بجوار كل قسم
    $$('#sectionList [data-sec-id]').forEach(function (li) {
      var n = M.tasksInSection(p, li.dataset.secId);
      li.querySelector('.sec-count').textContent = n ? tasksCountText(n) : 'لا مهام';
    });

    // صفوف المهام: القيم المحسوبة والأخطاء
    var byId = {};
    a.tasks.forEach(function (r) { byId[r.id] = r; });
    $$('#taskBody tr[data-id]').forEach(function (tr) {
      var id = tr.dataset.id, r = byId[id];
      if (!r) return;
      setCalc(tr, 'hours', r.hours === null ? null : fmtHours(r.hours));
      setCalc(tr, 'share', r.share === null ? null : U.formatPercent(r.share, 1));
      tr.querySelector('[data-out="bar"]').style.width = r.share ? Math.min(100, r.share * 100).toFixed(1) + '%' : '0';
      $$('[data-task-field]', tr).forEach(function (ctrl) {
        var key = id + ':' + ctrl.dataset.taskField;
        applyFieldError(ctrl, tr.querySelector('[data-err="' + ctrl.dataset.taskField + '"]'), errorVisible(key) ? errs[key] : '');
      });
    });
    $('#tasksMsg').textContent = (state.showAllErrors || state.tasksGateShown) ? tasksGateMessage(v) : '';

    renderStepMarks(v);
    if (state.step === 2) renderReview();
    if (WL.aiPanel) WL.aiPanel.refreshIfIdle();
  }

  function setCalc(tr, key, text) {
    var el = tr.querySelector('[data-out="' + key + '"]');
    el.textContent = text === null ? '—' : text;
    el.classList.toggle('is-empty', text === null);
  }

  function renderStepMarks(v) {
    var orgErr = v.errors.some(function (e) { return e.scope === 'org'; });
    var taskErr = v.errors.some(function (e) { return e.scope !== 'org'; });
    var marks = [orgErr, taskErr, !v.valid];
    // لا انتقال إلى المهام قبل اكتمال البيانات الأساسية، ولا إلى المراجعة قبل اكتمال كل المهام
    var orgHint = 'أكمل الحقول الإلزامية في البيانات الأساسية أولًا';
    var taskHint = 'أكمل القسم والمسمى الفعلي والتردد والتكرار والمدة لجميع المهام أولًا';
    $$('.step').forEach(function (btn) {
      var n = Number(btn.dataset.step);
      btn.classList.toggle('is-active', n === state.step);
      btn.classList.toggle('has-error', (state.showAllErrors || (n === 1 && state.tasksGateShown)) && n < 2 && marks[n]);
      btn.classList.toggle('is-done', n < 2 && !marks[n] && n !== state.step);
      if (n > 0 && orgErr) setLocked(btn, true, orgHint);
      else setLocked(btn, n === 2 && taskErr, taskHint);
    });
    if (state.step === 0) setLocked($('#nextBtn'), orgErr, orgHint);
    else setLocked($('#nextBtn'), state.step === 1 && taskErr, taskHint);
  }

  /** عدد بصيغة عربية صحيحة: مهمة واحدة، مهمتان، 3 مهام، 11 مهمة. */
  function tasksCountText(n) {
    if (n === 1) return 'مهمة واحدة';
    if (n === 2) return 'مهمتان';
    return n + (n >= 3 && n <= 10 ? ' مهام' : ' مهمة');
  }

  /** رسالة ثابتة تحت جدول المهام بما يمنع الانتقال إلى المراجعة (فارغة إذا اكتملت المهام). */
  function tasksGateMessage(v) {
    if (!state.project.tasks.length) return 'يرجى إضافة مهمة واحدة على الأقل للانتقال إلى المرحلة التالية.';
    var ids = {};
    v.errors.forEach(function (e) { if (e.scope === 'task') ids[e.taskId] = true; });
    var n = Object.keys(ids).length;
    if (!n) return '';
    return (n === 1 ? 'توجد مهمة واحدة غير مكتملة' : 'توجد ' + tasksCountText(n) + ' غير مكتملة') +
      '. أكمل القسم والمسمى الفعلي والتردد والتكرار والمدة لكل مهمة للانتقال إلى المرحلة التالية.';
  }

  function setLocked(el, isLocked, hint) {
    el.setAttribute('aria-disabled', isLocked ? 'true' : 'false');
    if (isLocked) el.title = hint; else el.removeAttribute('title');
  }

  // ───────────── المراجعة ─────────────

  function kpi(label, value, note, extraClass) {
    return '<div class="kpi ' + (extraClass || '') + '"><div class="kpi-label">' + label + '</div>' +
      '<div class="kpi-value">' + value + '</div>' + (note ? '<div class="kpi-note">' + note + '</div>' : '') + '</div>';
  }

  function renderReview() {
    var a = state.analysis, v = state.validation, org = state.project.org;
    var e = U.escapeHtml;

    $('#reviewSubtitle').textContent = [org.entity, org.department].filter(function (x) { return String(x || '').trim(); }).join(' — ') ||
      'راجع النتائج ثم صدّر ملف Excel.';

    // الأخطاء
    var panel = $('#errorPanel');
    panel.hidden = v.valid;
    $('#errorList').innerHTML = v.errors.map(function (err, i) {
      return '<li><a data-error-index="' + i + '">' + e(err.message) + '</a></li>';
    }).join('');
    $('#exportBtn').disabled = !v.valid;

    // المؤشرات
    $('#kpis').innerHTML =
      kpi('العدد الفعلي', a.actualCount === null ? '—' : a.actualCount, '') +
      kpi('الاحتياج المحسوب', fmtNeed(a.calcNeed), 'قبل التقريب: ' + U.formatNumber(a.exactNeed, 2)) +
      kpi('إجمالي ساعات العمل', fmtHours(a.totalHours), 'ساعة سنويًا') +
      kpi('الفجوة', '<span dir="ltr">' + fmtGap(a.gap) + '</span>', 'العدد الفعلي − الاحتياج المحسوب') +
      kpi('الحالة', e(a.status), '', a.statusKey ? 'status-' + a.statusKey : '') +
      kpi('عدد المهام', a.taskCount, '');

    // التوزيع حسب التردد
    var rows = a.byFrequency.map(function (f) {
      var muted = f.count === 0 ? ' class="muted"' : '';
      return '<tr' + muted + '><td>' + e(f.label) + '</td><td class="num">' + f.perYear + '</td><td class="num">' + f.count +
        '</td><td class="num">' + fmtHours(f.hours) + '</td><td class="num">' + U.formatPercent(f.share, 1) + '</td></tr>';
    }).join('');
    $('#freqTable').innerHTML =
      '<thead><tr><th>التردد</th><th>المرات السنوية</th><th>عدد المهام</th><th>الساعات</th><th>النسبة</th></tr></thead>' +
      '<tbody>' + rows + '</tbody>' +
      '<tfoot><tr><td>الإجمالي</td><td></td><td class="num">' + a.taskCount + '</td><td class="num">' + fmtHours(a.totalHours) +
      '</td><td class="num">' + U.formatPercent(a.totalHours > 0 ? 1 : 0, 1) + '</td></tr></tfoot>';

    // الاحتياج حسب القسم / الوحدة × المسمى الفعلي
    var secs = a.bySection.filter(function (sec) { return sec.positions.length; });
    var posRows = secs.map(function (sec) {
      return '<tr class="sec-head"><td colspan="5">' + e(sec.name) + '</td></tr>' +
        sec.positions.map(function (g) {
          return '<tr><td class="pos-title">' + e(g.title) + '</td><td class="num">' + g.count + '</td><td class="num">' + fmtHours(g.hours) +
            '</td><td class="num">' + U.formatNumber(g.exactNeed, 2) + '</td><td class="num"><b>' + fmtNeed(g.need) + '</b></td></tr>';
        }).join('') +
        '<tr class="sec-total"><td>مجموع ' + e(sec.name) + '</td><td class="num">' + sec.count + '</td><td class="num">' + fmtHours(sec.hours) +
          '</td><td></td><td class="num">' + fmtNeed(sec.needSum) + '</td></tr>';
    }).join('');
    $('#posTable').innerHTML =
      '<thead><tr><th>القسم / الوحدة — المسمى الفعلي</th><th>عدد المهام</th><th>إجمالي الساعات</th><th>قبل التقريب</th><th>الاحتياج</th></tr></thead>' +
      '<tbody>' + (posRows || '<tr><td colspan="5" class="muted">لا توجد مسميات بعد.</td></tr>') + '</tbody>' +
      (posRows ? '<tfoot><tr><td>الإجمالي</td><td class="num">' + secs.reduce(function (s, x) { return s + x.count; }, 0) +
        '</td><td class="num">' + fmtHours(secs.reduce(function (s, x) { return s + x.hours; }, 0)) +
        '</td><td></td><td class="num">' + fmtNeed(a.positionsNeedSum) + '</td></tr></tfoot>' : '');
    var diff = secs.length && a.positionsNeedSum !== a.calcNeed;
    $('#posNote').hidden = !diff;
    $('#posNote').textContent = diff
      ? 'مجموع الاحتياج حسب القسم والمسمى (' + fmtNeed(a.positionsNeedSum) + ') يختلف عن الاحتياج الإجمالي للجهة (' + fmtNeed(a.calcNeed) +
        ') لأن كل مسمى داخل كل قسم يُقرَّب على حدة، بينما يُحسب الاحتياج الإجمالي من مجموع الساعات.'
      : '';
  }

  function jumpToError(err) {
    var selector, step;
    if (err.scope === 'org') { step = 0; selector = err.field === 'sections' ? '#f-newSection' : '[data-org="' + err.field + '"]'; }
    else if (err.scope === 'task') { step = 1; selector = 'tr[data-id="' + err.taskId + '"] [data-task-field="' + err.field + '"]'; }
    else { step = 1; selector = '[data-action="add-task"]'; }
    goToStep(step);
    var el = $(selector);
    if (el) { el.focus(); el.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
  }

  // ───────────── التنقل ─────────────

  /** اكتملت الحقول الإلزامية في البيانات الأساسية (شرط الانتقال إلى المهام والمراجعة). */
  function orgComplete() {
    return V.validateOrg(state.project.org).length === 0;
  }

  /** عند محاولة التقدم قبل اكتمال البيانات الأساسية: إظهار الأخطاء والانتقال لأول حقل ناقص. */
  function explainBlockedNavigation() {
    $$('[data-org]').forEach(function (i) { state.touched['org:' + i.dataset.org] = true; });
    state.touched['org:sections'] = true;
    if (state.step !== 0) goToStep(0, true);
    refresh();
    var first = V.validateOrg(state.project.org)[0];
    var input = first && (first.field === 'sections' ? $('#f-newSection') : $('[data-org="' + first.field + '"]'));
    if (input) { input.focus(); input.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
    toast('أكمل الحقول الإلزامية في البيانات الأساسية أولًا.', 'error');
  }

  /** أخطاء المهام (مهمة ناقصة أو لا توجد مهام) التي تمنع الانتقال إلى المراجعة. */
  function taskErrors() {
    return V.validateProject(state.project, cfg).errors.filter(function (e) { return e.scope !== 'org'; });
  }

  /** عند محاولة الانتقال إلى المراجعة قبل اكتمال المهام: إظهار النواقص والانتقال لأول حقل ناقص. */
  function explainBlockedTasks() {
    state.tasksGateShown = true;
    state.project.tasks.forEach(function (t) {
      ['title', 'sectionId', 'positionTitle', 'frequencyKey', 'repetitions', 'durationMinutes'].forEach(function (f) { state.touched[t.id + ':' + f] = true; });
    });
    if (state.step !== 1) goToStep(1, true);
    refresh();
    var first = taskErrors()[0];
    var el = first && first.scope === 'task'
      ? $('#taskBody tr[data-id="' + first.taskId + '"] [data-task-field="' + first.field + '"]')
      : $('[data-action="add-task"]');
    if (el) { el.focus(); el.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
    toast(first && first.scope === 'task'
      ? 'أكمل القسم والمسمى الفعلي والتردد والتكرار والمدة لجميع المهام للانتقال إلى المرحلة التالية.'
      : 'أضف مهمة واحدة على الأقل للانتقال إلى المرحلة التالية.', 'error');
  }

  function goToStep(n, silent) {
    if (n > 0 && !orgComplete()) { explainBlockedNavigation(); return; }
    if (n === LAST_STEP && taskErrors().length) { explainBlockedTasks(); return; }
    state.step = Math.max(0, Math.min(LAST_STEP, n));
    if (state.step === LAST_STEP) state.showAllErrors = true;
    $$('[data-panel]').forEach(function (p) { p.hidden = Number(p.dataset.panel) !== state.step; });
    $('#prevBtn').disabled = state.step === 0;
    $('#nextBtn').hidden = state.step === LAST_STEP;
    if (state.step === 1) $$('#taskBody textarea').forEach(autoGrow);
    if (!silent) window.scrollTo({ top: 0, behavior: 'smooth' });
    refresh();
  }

  // ───────────── التصدير ─────────────

  async function exportExcel() {
    refresh();
    if (!state.validation.valid) {
      state.showAllErrors = true;
      goToStep(2);
      toast('لا يمكن التصدير: توجد بيانات ناقصة أو غير صحيحة.', 'error');
      return;
    }
    var btn = $('#exportBtn');
    var label = btn.querySelector('span');
    btn.classList.add('is-loading');
    label.textContent = 'جاري إنشاء الملف…';
    try {
      var buffer = await WL.excel.exportToBuffer(state.project, cfg, WL.template);
      var name = WL.excel.buildFileName(state.project, cfg);
      WL.excel.downloadBuffer(buffer, name);
      toast('تم إنشاء ملف Excel: ' + name, 'success');
    } catch (err) {
      console.error(err);
      toast(err && err.message === 'EXCELJS_MISSING'
        ? 'تعذر تحميل مكتبة Excel. تأكد من وجود مجلد lib بجوار الصفحة.'
        : 'تعذر إنشاء ملف Excel. يرجى المحاولة مرة أخرى.', 'error');
    } finally {
      btn.classList.remove('is-loading');
      label.textContent = 'تصدير Excel';
    }
  }

  // ───────────── إدارة المشروع ─────────────

  function setProject(project) {
    state.project = project;
    state.touched = {};
    state.showAllErrors = false;
    state.tasksGateShown = false;
    fillOrgInputs();
    renderTasks();
    goToStep(0);
  }

  async function confirmReplace(title) {
    if (!projectHasData(state.project)) return true;
    return confirmDialog(title, 'سيتم استبدال البيانات الحالية. احفظ المشروع أولًا إذا أردت الاحتفاظ بها.', 'متابعة');
  }

  function saveProjectFile() {
    var blob = new Blob([M.serialize(state.project)], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'مشروع_عبء_العمل_' + (U.sanitizeFileName(state.project.org.entity) || 'جديد') + '_' + U.todayISO() + '.json';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    toast('تم حفظ ملف المشروع.', 'success');
  }

  function openProjectFile(file) {
    var reader = new FileReader();
    reader.onload = async function () {
      var project;
      try { project = M.deserialize(String(reader.result)); }
      catch (e) { toast('الملف المختار ليس ملف مشروع صالحًا.', 'error'); return; }
      if (await confirmReplace('فتح مشروع')) {
        setProject(project);
        toast('تم فتح المشروع.', 'success');
      }
    };
    reader.onerror = function () { toast('تعذر قراءة الملف.', 'error'); };
    reader.readAsText(file, 'utf-8');
  }

  // ───────────── القائمة ووضع الذكاء الاصطناعي ─────────────

  function setMenuOpen(open) {
    $('#appMenu').hidden = !open;
    $('#menuBtn').setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) $('#appMenu [role^="menuitem"]').focus();
  }

  function bindMenu() {
    $('#menuBtn').addEventListener('click', function () { setMenuOpen($('#appMenu').hidden); });
    document.addEventListener('click', function (e) {
      if (!$('#appMenu').hidden && !e.target.closest('.menu-wrap')) setMenuOpen(false);
    });
    $('#appMenu').addEventListener('keydown', function (e) {
      var items = $$('#appMenu [role^="menuitem"]');
      var i = items.indexOf(document.activeElement);
      if (e.key === 'Escape') { setMenuOpen(false); $('#menuBtn').focus(); }
      else if (e.key === 'ArrowDown') { e.preventDefault(); items[(i + 1) % items.length].focus(); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); items[(i - 1 + items.length) % items.length].focus(); }
    });
  }

  /** وضع الذكاء الاصطناعي متوقف افتراضيًا، ولا يعمل إلا بعد تفعيله من القائمة. */
  var aiModeSession = false; // احتياطي إن تعذّر التخزين في المتصفح (التصفح الخاص مثلًا)
  function isAiModeOn() {
    try {
      var v = localStorage.getItem(cfg.app.aiModeStorageKey);
      return v === null ? aiModeSession : v === 'on';
    } catch (e) { return aiModeSession; }
  }

  /** يُظهر أو يُخفي كل نقاط الدخول للمساعد الذكي (الزر، أزرار ✦ في المهام، اللوحة). */
  function applyAiMode(on) {
    document.body.classList.toggle('ai-disabled', !on);
    $('#aiModeToggle').setAttribute('aria-checked', on ? 'true' : 'false');
    if (!on && WL.aiPanel) WL.aiPanel.close();
  }

  function toggleAiMode() {
    var on = !isAiModeOn();
    aiModeSession = on;
    try { localStorage.setItem(cfg.app.aiModeStorageKey, on ? 'on' : 'off'); } catch (e) { /* يبقى للجلسة الحالية */ }
    applyAiMode(on);
    toast(on ? 'تم تفعيل وضع الذكاء الاصطناعي.' : 'تم إيقاف وضع الذكاء الاصطناعي.', 'success');
  }

  // ───────────── ربط الأحداث العامة ─────────────

  function bindGlobalActions() {
    document.addEventListener('click', async function (e) {
      var errLink = e.target.closest('[data-error-index]');
      if (errLink) { jumpToError(state.validation.errors[Number(errLink.dataset.errorIndex)]); return; }

      var stepBtn = e.target.closest('.step[data-step]');
      if (stepBtn) { goToStep(Number(stepBtn.dataset.step)); return; }

      var el = e.target.closest('[data-action]');
      if (!el) return;
      // الإجراءات من القائمة تغلقها، ما عدا مفتاح وضع الذكاء الاصطناعي ليرى المستخدم حالته
      if (el.closest('#appMenu') && el.dataset.action !== 'toggle-ai-mode') setMenuOpen(false);
      switch (el.dataset.action) {
        case 'toggle-ai-mode': toggleAiMode(); break;
        case 'add-task': addTask(); break;
        case 'add-section': addSectionFromInput(); break;
        case 'paste-tasks': WL.pasteDialog.open(); break;
        case 'ai-open': if (isAiModeOn()) WL.aiPanel.open(); break;
        case 'next':
          if (state.step === 0) {
            $$('[data-org]').forEach(function (i) { state.touched['org:' + i.dataset.org] = true; });
            state.touched['org:sections'] = true;
          }
          goToStep(state.step + 1); break;
        case 'prev': goToStep(state.step - 1); break;
        case 'goto-step': goToStep(Number(el.dataset.target)); break;
        case 'export': exportExcel(); break;
        case 'save-project': saveProjectFile(); break;
        case 'open-project': $('#projectFile').click(); break;
        case 'new-project':
          if (await confirmReplace('مشروع جديد')) setProject(M.createProject());
          break;
        case 'load-sample':
          if (await confirmReplace('تحميل بيانات تجريبية')) { setProject(WL.sampleProject()); toast('تم تحميل بيانات تجريبية.', 'success'); }
          break;
      }
    });

    $('#projectFile').addEventListener('change', function (e) {
      var file = e.target.files && e.target.files[0];
      if (file) openProjectFile(file);
      e.target.value = '';
    });
  }

  // ───────────── التشغيل ─────────────

  /** رأس جدول المهام يلتصق أسفل شريط المراحل مباشرة مهما كان ارتفاعه. */
  function syncStickyOffset() {
    document.documentElement.style.setProperty('--stepper-h', $('.stepper').offsetHeight + 'px');
  }

  function init() {
    document.title = 'تحليل عبء العمل';
    syncStickyOffset();
    window.addEventListener('resize', syncStickyOffset);
    $('#appTitle').textContent = cfg.app.title;
    clearLegacyDraft();
    state.project = M.createProject();
    window.addEventListener('beforeunload', warnBeforeLeaving);
    bindOrgInputs();
    bindSections();
    bindTaskTable();
    bindGlobalActions();
    bindMenu();
    applyAiMode(isAiModeOn());
    fillOrgInputs();
    renderTasks();
    goToStep(0, true);
  }

  // واجهة صغيرة للاختبار الآلي
  WL.app = {
    getState: function () { return state; },
    setProject: setProject,
    goToStep: goToStep,
    renderTasks: renderTasks,
    toast: toast,
    confirm: confirmDialog,
    /** يُظهر أخطاء الحقول الناقصة لمهام محددة (بعد لصقها) دون انتظار خروج المستخدم من كل حقل. */
    touchTasks: function (ids) {
      ids.forEach(function (id) {
        ['title', 'sectionId', 'positionTitle', 'frequencyKey', 'repetitions', 'durationMinutes'].forEach(function (f) { state.touched[id + ':' + f] = true; });
      });
      refresh();
    }
  };

  document.addEventListener('DOMContentLoaded', init);
})();
