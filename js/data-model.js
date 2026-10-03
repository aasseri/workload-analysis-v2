/**
 * data-model.js — بنية بيانات المشروع وعمليات التعديل عليها.
 * القيم المشتقة (المرات السنوية، الساعات، النسب) لا تُخزّن هنا أبدًا؛ تُحسب في calculations.js.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  var SCHEMA_VERSION = 2;   // 2: الأقسام / الوحدات قائمة، وكل مهمة مرتبطة بقسم

  var ORG_FIELDS = [
    'entity', 'department', 'jobTitle',
    'actualCount'
  ];

  function createProject() {
    return {
      schemaVersion: SCHEMA_VERSION,
      org: {
        entity: '',
        department: '',
        sections: [],          // الأقسام / الوحدات: [{ id, name }] — تُعرّف في البيانات الأساسية وتُختار لكل مهمة
        jobTitle: '',          // مخفي حاليًا في الواجهة؛ سياق اختياري للمساعد الذكي فقط، لا يُصدَّر إلى Excel
        actualCount: ''    // يُخزن كما أدخله المستخدم، ويُقرأ رقميًا عند التحقق/الحساب
        // تاريخ الإعداد لا يُخزن: يُكتب في ملف Excel بتاريخ التصدير تلقائيًا
      },
      tasks: [],
      employees: [] // محجوز للتوسع المستقبلي
    };
  }

  function createTask(partial) {
    var p = partial || {};
    return {
      id: U.uid(),
      title: p.title || '',
      sectionId: p.sectionId || '',           // القسم / الوحدة (معرّف من org.sections)
      positionTitle: p.positionTitle || '',   // المسمى الفعلي الذي ينفّذ المهمة
      frequencyKey: p.frequencyKey || '',
      repetitions: p.repetitions === undefined ? '' : p.repetitions,
      durationMinutes: p.durationMinutes === undefined ? '' : p.durationMinutes
    };
  }

  // ───────────── الأقسام / الوحدات ─────────────

  function sectionName(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
  }

  function findSection(project, id) {
    return project.org.sections.filter(function (s) { return s.id === id; })[0] || null;
  }

  function findSectionByName(project, name) {
    var key = WL.taskImport ? WL.taskImport.normalizeArabic(name) : sectionName(name);
    return project.org.sections.filter(function (s) {
      return (WL.taskImport ? WL.taskImport.normalizeArabic(s.name) : s.name) === key;
    })[0] || null;
  }

  /** يضيف قسمًا. @returns {{section}|{error}} */
  function addSection(project, name) {
    var n = sectionName(name);
    if (!n) return { error: 'اكتب اسم القسم أو الوحدة أولًا.' };
    if (findSectionByName(project, n)) return { error: 'هذا القسم مضاف مسبقًا.' };
    var s = { id: 's' + U.uid(), name: n };
    project.org.sections.push(s);
    return { section: s };
  }

  /** يعيد تسمية قسم (المهام مرتبطة بالمعرّف فتتبع الاسم الجديد تلقائيًا). */
  function renameSection(project, id, name) {
    var s = findSection(project, id);
    var n = sectionName(name);
    if (!s) return { error: 'القسم غير موجود.' };
    if (!n) return { error: 'اسم القسم لا يمكن أن يكون فارغًا.' };
    var other = findSectionByName(project, n);
    if (other && other.id !== id) return { error: 'يوجد قسم آخر بهذا الاسم.' };
    s.name = n;
    return { section: s };
  }

  /** يحذف قسمًا ويلغي ارتباط المهام به. @returns عدد المهام التي أُلغي ارتباطها */
  function removeSection(project, id) {
    project.org.sections = project.org.sections.filter(function (s) { return s.id !== id; });
    var n = 0;
    project.tasks.forEach(function (t) { if (t.sectionId === id) { t.sectionId = ''; n += 1; } });
    return n;
  }

  function tasksInSection(project, id) {
    return project.tasks.filter(function (t) { return t.sectionId === id; }).length;
  }

  // ───────────── المهام ─────────────

  function indexOfTask(project, id) {
    for (var i = 0; i < project.tasks.length; i++) {
      if (project.tasks[i].id === id) return i;
    }
    return -1;
  }

  function addTask(project, partial) {
    var task = createTask(partial);
    // قسم واحد فقط معرّف: يُختار تلقائيًا
    if (!task.sectionId && project.org.sections.length === 1) task.sectionId = project.org.sections[0].id;
    project.tasks.push(task);
    return task;
  }

  function updateTask(project, id, changes) {
    var i = indexOfTask(project, id);
    if (i < 0) return null;
    Object.keys(changes).forEach(function (k) {
      if (k !== 'id') project.tasks[i][k] = changes[k];
    });
    return project.tasks[i];
  }

  function duplicateTask(project, id) {
    var i = indexOfTask(project, id);
    if (i < 0) return null;
    var copy = U.deepClone(project.tasks[i]);
    copy.id = U.uid();
    project.tasks.splice(i + 1, 0, copy);
    return copy;
  }

  function removeTask(project, id) {
    var i = indexOfTask(project, id);
    if (i < 0) return null;
    return project.tasks.splice(i, 1)[0];
  }

  /** ينقل المهمة إلى موضع جديد (فهرس يبدأ من 0). */
  function moveTaskTo(project, id, newIndex) {
    var i = indexOfTask(project, id);
    if (i < 0) return false;
    var target = Math.max(0, Math.min(project.tasks.length - 1, newIndex));
    if (target === i) return false;
    var task = project.tasks.splice(i, 1)[0];
    project.tasks.splice(target, 0, task);
    return true;
  }

  function moveTask(project, id, delta) {
    return moveTaskTo(project, id, indexOfTask(project, id) + delta);
  }

  function serialize(project) {
    return JSON.stringify(project, null, 2);
  }

  /**
   * يقرأ مشروعًا محفوظًا ويكمل أي حقل ناقص بقيمه الافتراضية. يرمي خطأ إذا كان الملف غير صالح.
   * المشاريع من الإصدار الأول (حقل قسم نصي واحد) تُحوَّل إلى قائمة أقسام بقسم واحد، وتُربط مهامها به.
   */
  function deserialize(json) {
    var raw = typeof json === 'string' ? JSON.parse(json) : json;
    if (!raw || typeof raw !== 'object' || !raw.org || !Array.isArray(raw.tasks)) {
      throw new Error('INVALID_PROJECT');
    }
    var project = createProject();
    ORG_FIELDS.forEach(function (f) {
      if (raw.org[f] !== undefined && raw.org[f] !== null) project.org[f] = raw.org[f];
    });
    if (Array.isArray(raw.org.sections)) {
      raw.org.sections.forEach(function (s) {
        var n = sectionName(s && s.name);
        if (n && !findSectionByName(project, n)) project.org.sections.push({ id: String(s.id || 's' + U.uid()), name: n });
      });
    }
    var legacyId = '';
    if (!project.org.sections.length && sectionName(raw.org.section)) {
      legacyId = addSection(project, raw.org.section).section.id;
    }
    project.tasks = raw.tasks.map(function (t) {
      var task = createTask(t);
      if (t.id) task.id = String(t.id);
      if (task.sectionId && !findSection(project, task.sectionId)) task.sectionId = '';
      if (!task.sectionId && legacyId) task.sectionId = legacyId;
      return task;
    });
    project.employees = Array.isArray(raw.employees) ? raw.employees : [];
    return project;
  }

  WL.model = {
    SCHEMA_VERSION: SCHEMA_VERSION,
    ORG_FIELDS: ORG_FIELDS,
    createProject: createProject,
    createTask: createTask,
    sectionName: sectionName,
    findSection: findSection,
    findSectionByName: findSectionByName,
    addSection: addSection,
    renameSection: renameSection,
    removeSection: removeSection,
    tasksInSection: tasksInSection,
    indexOfTask: indexOfTask,
    addTask: addTask,
    updateTask: updateTask,
    duplicateTask: duplicateTask,
    removeTask: removeTask,
    moveTask: moveTask,
    moveTaskTo: moveTaskTo,
    serialize: serialize,
    deserialize: deserialize
  };
})(typeof window !== 'undefined' ? window : globalThis);
