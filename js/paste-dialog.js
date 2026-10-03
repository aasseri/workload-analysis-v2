/**
 * paste-dialog.js — نافذة «لصق مهام»: لصق مهام من Excel أو Word، ومعاينتها وتعديلها، ثم إضافتها للجدول.
 * التحليل في task-import.js؛ هذه الطبقة للواجهة فقط. لا يُضاف شيء قبل ضغط «إضافة».
 */
(function () {
  'use strict';
  var WL = window.WL;
  var cfg = WL.config, U = WL.utils, M = WL.model, TI = WL.taskImport;
  var e = U.escapeHtml;

  var FIELDS = [
    { key: 'title', label: 'عمود المهمة', required: true },
    { key: 'section', label: 'عمود القسم / الوحدة' },
    { key: 'position', label: 'عمود المسمى الفعلي' },
    { key: 'frequency', label: 'عمود التردد' },
    { key: 'repetitions', label: 'عمود التكرار' },
    { key: 'duration', label: 'عمود المدة (دقائق)' }
  ];

  var st = { parsed: null, mapping: null, items: [], timer: null };

  /** عدد المهام بصيغة عربية صحيحة: مهمة واحدة، مهمتان، 3 مهام، 11 مهمة. */
  function tasksText(n) {
    if (n === 1) return 'مهمة واحدة';
    if (n === 2) return 'مهمتين';
    return n + (n >= 3 && n <= 10 ? ' مهام' : ' مهمة');
  }

  function $(sel) { return document.querySelector(sel); }
  function dlg() { return $('#pasteDialog'); }
  function project() { return WL.app.getState().project; }
  function mode() { var r = document.querySelector('input[name="pasteMode"]:checked'); return r ? r.value : 'append'; }

  function open() {
    st = { parsed: null, mapping: null, items: [], timer: null };
    $('#pasteInput').value = '';
    var hasTasks = project().tasks.length > 0;
    $('#pasteModeBox').hidden = !hasTasks;
    document.querySelector('input[name="pasteMode"][value="append"]').checked = true;
    render();
    dlg().showModal();
    $('#pasteInput').focus();
  }

  function close() { if (dlg().open) dlg().close(); }

  // ───────────── التحليل والمعاينة ─────────────

  function reparse() {
    var text = $('#pasteInput').value;
    st.parsed = text.trim() ? TI.analyze(text, cfg) : null;
    st.mapping = st.parsed ? Object.assign({}, st.parsed.mapping) : null;
    rebuildItems();
  }

  /** يعيد بناء عناصر المعاينة (عند تغيير النص أو ربط الأعمدة). */
  function rebuildItems() {
    var existing = mode() === 'replace' ? [] : project().tasks;
    st.items = st.parsed ? TI.buildItems(st.parsed, st.mapping, existing, cfg) : [];
    render();
  }

  /** عند تغيير طريقة الإضافة: يُحدَّث وسم «مكررة» دون المساس بما عدّله المستخدم يدويًا. */
  function refreshDuplicates() {
    var existing = {};
    if (mode() !== 'replace') project().tasks.forEach(function (t) { existing[TI.normalizeArabic(t.title)] = true; });
    var seen = {};
    st.items.forEach(function (it) {
      var key = TI.normalizeArabic(it.title);
      it.duplicate = existing[key] ? 'existing' : seen[key] ? 'pasted' : '';
      seen[key] = true;
      if (!it.userToggled) it.selected = !it.duplicate && !it.heading;
    });
    render();
  }

  function mappingHtml() {
    var p = st.parsed;
    if (!p || p.listMode || p.columns.length < 2) return '';
    return FIELDS.map(function (f) {
      var opts = (f.required ? '' : '<option value="-1">لا يوجد</option>') + p.columns.map(function (c) {
        var sample = c.sample.length > 28 ? c.sample.slice(0, 28) + '…' : c.sample;
        return '<option value="' + c.index + '"' + (st.mapping[f.key] === c.index ? ' selected' : '') + '>' +
          e(c.name) + (sample && c.name !== sample ? ' — ' + e(sample) : '') + '</option>';
      }).join('');
      return '<label class="pd-map"><span>' + f.label + (f.required ? ' <span class="req">*</span>' : '') + '</span>' +
        '<select data-map="' + f.key + '">' + opts + '</select></label>';
    }).join('') + '<p class="pd-note">التردد يُنقل فقط إذا طابق القيم المعتمدة (يومي، أسبوعي...). والتكرار والمدة فقط إذا كانت أرقامًا أكبر من صفر.</p>';
  }

  /**
   * القسم المطابق لعنصر ملصوق: بالاسم (بعد تجاهل الهمزات والتاء المربوطة) من الأقسام المعرّفة فقط —
   * لا يُنشأ قسم جديد تلقائيًا. إذا كان هناك قسم واحد فقط يُختار هو.
   */
  function sectionFor(it) {
    var p = project();
    var s = it.sectionName ? M.findSectionByName(p, it.sectionName) : null;
    if (!s && p.org.sections.length === 1) s = p.org.sections[0];
    return s;
  }

  function itemHtml(it, i) {
    var f = it.frequencyKey ? WL.calc.findFrequency(cfg, it.frequencyKey) : null;
    var meta = [];
    var sec = sectionFor(it);
    meta.push(sec ? '<span class="chip ok">' + e(sec.name) + '</span>'
      : it.sectionName ? '<span class="chip warn">القسم «' + e(it.sectionName) + '» غير معرّف: يُختار لاحقًا</span>'
      : '<span class="chip need">القسم: يُختار لاحقًا</span>');
    meta.push(it.positionTitle ? '<span class="chip ok">' + e(it.positionTitle) + '</span>' : '<span class="chip need">المسمى: يُدخل لاحقًا</span>');
    meta.push(f ? '<span class="chip ok">' + e(f.label) + '</span>' : '<span class="chip need">التردد: يُدخل لاحقًا</span>');
    if (it.repetitions) meta.push('<span class="chip ok">التكرار ' + e(it.repetitions) + '</span>');
    if (it.durationMinutes) meta.push('<span class="chip ok">' + e(it.durationMinutes) + ' دقيقة</span>');
    if (!it.repetitions || !it.durationMinutes) {
      meta.push('<span class="chip need">' + (!it.repetitions && !it.durationMinutes ? 'التكرار والمدة' : !it.repetitions ? 'التكرار' : 'المدة') + ': يُدخل لاحقًا</span>');
    }
    if (it.duplicate === 'existing') meta.push('<span class="chip warn">مكررة مع مهمة موجودة</span>');
    if (it.duplicate === 'pasted') meta.push('<span class="chip warn">مكررة في النص الملصوق</span>');
    if (it.heading) meta.push('<span class="chip warn">يبدو عنوانًا وليس مهمة</span>');
    return '<li class="pd-item' + (it.selected ? '' : ' is-off') + '">' +
      '<input type="checkbox" data-sel="' + i + '"' + (it.selected ? ' checked' : '') + ' aria-label="تضمين المهمة ' + (i + 1) + '">' +
      '<div class="pd-body"><input type="text" class="pd-title" data-title="' + i + '" value="' + e(it.title) + '" maxlength="' +
      cfg.ui.maxTaskTitleLength + '" aria-label="وصف المهمة ' + (i + 1) + '">' +
      '<div class="pd-meta">' + meta.join('') + '</div></div></li>';
  }

  function selectedItems() {
    return st.items.filter(function (it) { return it.selected && it.title.trim(); });
  }

  function render() {
    $('#pasteMapping').innerHTML = mappingHtml();
    $('#pasteMapping').hidden = !$('#pasteMapping').innerHTML;
    var n = selectedItems().length;
    $('#pastePreviewBox').hidden = !st.items.length;
    $('#pasteEmpty').hidden = !($('#pasteInput').value.trim() && !st.items.length);
    $('#pasteCount').textContent = st.items.length
      ? 'سيتم إضافة ' + n + ' من ' + tasksText(st.items.length)
      : '';
    $('#pastePreview').innerHTML = st.items.map(itemHtml).join('');
    var add = $('#pasteAdd');
    add.disabled = n === 0;
    add.textContent = n ? (mode() === 'replace' ? 'استبدال بـ ' + tasksText(n) : 'إضافة ' + tasksText(n)) : 'إضافة';
  }

  // ───────────── الإضافة ─────────────

  async function commit() {
    var chosen = selectedItems();
    if (!chosen.length) return;
    var p = project();
    var replace = mode() === 'replace' && p.tasks.length > 0;
    if (replace) {
      var ok = await WL.app.confirm('استبدال المهام', 'سيتم حذف ' + tasksText(p.tasks.length) + ' حالية واستبدالها بـ ' +
        tasksText(chosen.length) + ' ملصوقة. هل تريد المتابعة؟', 'استبدال');
      if (!ok) return;
      p.tasks = [];
    }
    // توحيد كتابة المسميات مع الموجودة (مثل «مساعد اداري» و«مساعد إداري») حتى لا تُحسب مسميين
    var canonical = {};
    p.tasks.concat(chosen.map(function (it) { return { positionTitle: it.positionTitle }; })).forEach(function (t) {
      var v = WL.calc.positionKey(t.positionTitle);
      var k = TI.normalizeArabic(v);
      if (v && !canonical[k]) canonical[k] = v;
    });
    var ids = chosen.map(function (it) {
      var pos = WL.calc.positionKey(it.positionTitle);
      var sec = sectionFor(it);
      return M.addTask(p, { title: it.title.trim(), sectionId: sec ? sec.id : '', positionTitle: pos ? canonical[TI.normalizeArabic(pos)] : '',
        frequencyKey: it.frequencyKey, repetitions: it.repetitions, durationMinutes: it.durationMinutes }).id;
    });
    close();
    WL.app.renderTasks();
    WL.app.touchTasks(ids);   // تظهر الحقول الناقصة مباشرة لتُستكمل
    var incomplete = chosen.filter(function (it) {
      return !sectionFor(it) || !it.positionTitle || !it.frequencyKey || !it.repetitions || !it.durationMinutes;
    }).length;
    WL.app.toast((replace ? 'تم استبدال المهام بـ ' : 'تمت إضافة ') + tasksText(chosen.length) +
      (incomplete ? '. أكمل القسم والمسمى الفعلي والتردد والتكرار والمدة.' : '.'), 'success');
    var firstRow = document.querySelector('#taskBody tr[data-id="' + ids[0] + '"]');
    if (firstRow) {
      firstRow.scrollIntoView({ block: 'center', behavior: 'smooth' });
      var firstGap = !sectionFor(chosen[0]) ? 'sectionId' : !chosen[0].positionTitle ? 'positionTitle' : !chosen[0].frequencyKey ? 'frequencyKey' : 'repetitions';
      var focusField = firstRow.querySelector('[data-task-field="' + firstGap + '"]');
      if (focusField) focusField.focus({ preventScroll: true });
    }
  }

  // ───────────── الأحداث ─────────────

  function init() {
    var d = dlg();
    $('#pasteInput').addEventListener('input', function () {
      clearTimeout(st.timer);
      st.timer = setTimeout(reparse, 200);
    });
    d.addEventListener('change', function (ev) {
      var t = ev.target;
      if (t.dataset.map) { st.mapping[t.dataset.map] = Number(t.value); rebuildItems(); }
      else if (t.dataset.sel !== undefined) {
        var it = st.items[Number(t.dataset.sel)];
        it.selected = t.checked; it.userToggled = true;
        t.closest('.pd-item').classList.toggle('is-off', !t.checked);
        render();
      } else if (t.name === 'pasteMode') refreshDuplicates();
    });
    d.addEventListener('input', function (ev) {
      if (ev.target.dataset.title !== undefined) {
        st.items[Number(ev.target.dataset.title)].title = ev.target.value;
        var n = selectedItems().length;
        $('#pasteCount').textContent = 'سيتم إضافة ' + n + ' من ' + tasksText(st.items.length);
        $('#pasteAdd').disabled = n === 0;
      }
    });
    d.addEventListener('click', function (ev) {
      var b = ev.target.closest('[data-paste]');
      if (!b) return;
      if (b.dataset.paste === 'cancel') close();
      else if (b.dataset.paste === 'add') commit();
      else if (b.dataset.paste === 'all' || b.dataset.paste === 'none') {
        st.items.forEach(function (it) { it.selected = b.dataset.paste === 'all'; it.userToggled = true; });
        render();
      }
    });
  }

  WL.pasteDialog = { open: open, close: close };
  document.addEventListener('DOMContentLoaded', init);
})();
