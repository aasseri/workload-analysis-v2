/**
 * validation.js — طبقة التحقق. تعيد قائمة أخطاء بلغة عربية واضحة، ولا تعدّل البيانات.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  var MSG = {
    entityRequired: 'يرجى إدخال اسم الجهة.',
    actualRequired: 'يرجى إدخال العدد الفعلي.',
    actualNotNumber: 'العدد الفعلي يجب أن يكون رقمًا.',
    actualNegative: 'يجب ألا يقل العدد الفعلي عن صفر.',
    actualNotInteger: 'العدد الفعلي يجب أن يكون عددًا صحيحًا (بدون كسور).',
    noTasks: 'يرجى إضافة مهمة واحدة على الأقل.',
    taskTitle: 'يرجى كتابة وصف المهمة.',
    taskPosition: 'يرجى إدخال المسمى الفعلي.',
    taskSection: 'يرجى اختيار القسم / الوحدة.',
    sectionsRequired: 'يرجى إضافة قسم أو وحدة واحدة على الأقل.',
    taskFrequency: 'لم يتم تحديد التردد.',
    repRequired: 'يرجى إدخال عدد التكرارات.',
    repNotNumber: 'عدد التكرارات يجب أن يكون رقمًا.',
    repPositive: 'يجب أن يكون عدد التكرارات أكبر من صفر.',
    durRequired: 'يرجى إدخال المدة بالدقائق.',
    durNotNumber: 'المدة بالدقائق يجب أن تكون رقمًا.',
    durPositive: 'يجب أن تكون المدة بالدقائق أكبر من صفر.'
  };

  function isBlank(v) {
    return v === null || v === undefined || String(v).trim() === '';
  }

  function validateOrg(org) {
    var errors = [];
    function add(field, message) { errors.push({ scope: 'org', field: field, message: message }); }

    if (isBlank(org.entity)) add('entity', MSG.entityRequired);
    if (!Array.isArray(org.sections) || !org.sections.length) add('sections', MSG.sectionsRequired);

    var actual = U.parseNumber(org.actualCount);
    if (actual === null) add('actualCount', MSG.actualRequired);
    else if (!U.isValidNumber(actual)) add('actualCount', MSG.actualNotNumber);
    else if (actual < 0) add('actualCount', MSG.actualNegative);
    else if (Math.floor(actual) !== actual) add('actualCount', MSG.actualNotInteger);

    return errors;
  }

  function validatePositive(value, msgs) {
    var n = U.parseNumber(value);
    if (n === null) return msgs[0];
    if (!U.isValidNumber(n)) return msgs[1];
    if (n <= 0) return msgs[2];
    return null;
  }

  /** @param {Array} sections أقسام المشروع (اختياري؛ للتحقق من أن قسم المهمة ما زال موجودًا) */
  function validateTask(task, index, cfg, sections) {
    var errors = [];
    function add(field, message) {
      errors.push({
        scope: 'task', field: field, taskId: task.id, taskIndex: index,
        message: 'المهمة رقم (' + (index + 1) + '): ' + message,
        short: message
      });
    }
    var known = cfg.frequencies.some(function (f) { return f.key === task.frequencyKey; });

    if (isBlank(task.title)) add('title', MSG.taskTitle);
    var secOk = !isBlank(task.sectionId) &&
      (!sections || sections.some(function (s) { return s.id === task.sectionId; }));
    if (!secOk) add('sectionId', MSG.taskSection);
    if (isBlank(task.positionTitle)) add('positionTitle', MSG.taskPosition);
    if (!known) add('frequencyKey', MSG.taskFrequency);

    var rep = validatePositive(task.repetitions, [MSG.repRequired, MSG.repNotNumber, MSG.repPositive]);
    if (rep) add('repetitions', rep);
    var dur = validatePositive(task.durationMinutes, [MSG.durRequired, MSG.durNotNumber, MSG.durPositive]);
    if (dur) add('durationMinutes', dur);
    return errors;
  }

  /** @returns {{valid:boolean, errors:Array}} */
  function validateProject(project, cfg) {
    var errors = validateOrg(project.org);
    if (!project.tasks.length) {
      errors.push({ scope: 'tasks', field: 'tasks', message: MSG.noTasks });
    }
    project.tasks.forEach(function (t, i) {
      errors = errors.concat(validateTask(t, i, cfg, project.org.sections));
    });
    return { valid: errors.length === 0, errors: errors };
  }

  /** فهرس سريع: مفتاح الحقل ← الرسالة المختصرة، لعرض الخطأ بجانب الحقل. */
  function indexErrors(errors) {
    var map = {};
    errors.forEach(function (e) {
      var key = e.scope === 'task' ? e.taskId + ':' + e.field : e.scope + ':' + e.field;
      if (!map[key]) map[key] = e.short || e.message;
    });
    return map;
  }

  WL.validation = {
    MESSAGES: MSG,
    validateOrg: validateOrg,
    validateTask: validateTask,
    validateProject: validateProject,
    indexErrors: indexErrors
  };
})(typeof window !== 'undefined' ? window : globalThis);
