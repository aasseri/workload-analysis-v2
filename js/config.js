/**
 * config.js — مصدر وحيد لكل القيم المرجعية وقواعد النظام.
 * لا توضع أي قيمة مرجعية (ترددات، ساعات سنوية، أسماء أوراق...) في أي ملف آخر.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});

  WL.config = {
    app: {
      title: 'نظام تحليل عبء العمل والاحتياج الوظيفي',
      reportTitle: 'تقرير تحليل عبء العمل والاحتياج الوظيفي',
      legacyDraftKey: 'wl.project.v1',   // مسودة الإصدارات السابقة: تُحذف عند الفتح (النظام يبدأ فارغًا دائمًا)
      aiModeStorageKey: 'wl.ui.aiMode'   // تشغيل/إيقاف وضع الذكاء الاصطناعي (متوقف افتراضيًا)
    },

    /** ساعات العمل الفعلية السنوية للموظف الواحد (تُكتب في ورقة الإعدادات وتُرجع إليها كل المعادلات). */
    annualWorkHours: 1023,

    /** عدد الدقائق في الساعة المستخدم في معادلة ساعات المهمة. */
    minutesPerHour: 60,

    /**
     * تقريب الاحتياج المحسوب: لأقرب عدد صحيح، 0.5 فأعلى يُجبر للأعلى (1.5 ← 2 ، 1.3 ← 1).
     * يطابق دالة ROUND في Excel. digits = عدد المنازل العشرية بعد التقريب.
     */
    needRounding: { digits: 0 },

    /** جدول الترددات السنوية (إلزامي، غير قابل للإدخال اليدوي). الترتيب هنا هو ترتيب العرض. */
    frequencies: [
      { key: 'daily',       label: 'يومي',      perYear: 240 },
      { key: 'weekly',      label: 'أسبوعي',    perYear: 48 },
      { key: 'semimonthly', label: 'نصف شهري',  perYear: 24 },
      { key: 'monthly',     label: 'شهري',      perYear: 12 },
      { key: 'quarterly',   label: 'ربع سنوي',  perYear: 4 },
      { key: 'semiannual',  label: 'نصف سنوي',  perYear: 2 },
      { key: 'annual',      label: 'سنوي',      perYear: 1 }
    ],

    status: {
      deficit:  'عجز',
      surplus:  'فائض',
      balanced: 'متوازن'
    },

    excel: {
      fileNamePrefix: 'تحليل_عبء_العمل',
      sheetNames: {
        basic:    'البيانات الأساسية',
        tasks:    'تحليل المهام',
        workload: 'تحليل عبء العمل',
        need:     'الاحتياج والفجوة',
        positions:'الاحتياج حسب القسم والمسمى',
        summary:  'الملخص التنفيذي',
        settings: 'الإعدادات'
      },
      /** ترتيب الأوراق في الملف الناتج. */
      sheetOrder: ['summary', 'basic', 'tasks', 'workload', 'need', 'positions', 'settings'],
      /** صفوف فارغة جاهزة بالمعادلات أسفل المهام ليضيف المستخدم مهامًا داخل Excel. */
      spareTaskRows: 50,
      /** إذا تجاوز عدد صفوف جدول المهام هذا الحد تُطبع الورقة على A3 بدل A4. */
      a3RowThreshold: 60,
      protection: {
        enabled: true,
        password: '' // بدون كلمة مرور: حماية من التعديل الخاطئ فقط
      }
    },

    ui: {
      maxTaskTitleLength: 500
    },

    /**
     * المساعد الذكي: قواعد محلية تعمل دائمًا دون إنترنت، ويُضاف تحليل Claude عند حفظ مفتاح API.
     * المساعد يقترح فقط؛ لا يُطبَّق أي تعديل إلا بقبول المستخدم، ولا يغيّر قواعد الحساب.
     */
    ai: {
      settingsStorageKey: 'wl.ai.settings',
      legacyKeyStorageKey: 'wl.ai.apiKey',   // مفتاح Claude المحفوظ في الإصدار السابق
      defaultProvider: 'anthropic',
      /**
       * المزوّدون المدعومون. أي مزوّد يدعم واجهة OpenAI Chat Completions يعمل عبر «متوافق مع OpenAI»
       * بإدخال عنوان الخدمة واسم النموذج.
       */
      providers: {
        anthropic: {
          label: 'Claude (Anthropic)',
          api: 'anthropic',
          defaultModel: 'claude-opus-5',
          sdkUrl: 'https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@0.128.0/+esm',
          maxTokens: 16000,
          keyPattern: '^sk-ant-\\S{10,}$',
          keyHint: 'sk-ant-...',
          keysUrl: 'https://console.anthropic.com/settings/keys'
        },
        openai: {
          label: 'OpenAI',
          api: 'openai-chat',
          baseUrl: 'https://api.openai.com/v1',
          defaultModel: 'gpt-6-astra',
          keyPattern: '^sk-(?!ant-)\\S{10,}$',
          keyHint: 'sk-...',
          keysUrl: 'https://platform.openai.com/api-keys'
        },
        custom: {
          label: 'مزوّد آخر متوافق مع OpenAI',
          api: 'openai-chat',
          baseUrl: '',
          defaultModel: '',
          keyPattern: '^\\S{8,}$',
          keyHint: 'المفتاح',
          keysUrl: ''
        }
      },
      /** حدود فحص معقولية الأرقام في القواعد المحلية (تنبيهات فقط، لا تمنع الحفظ). */
      limits: {
        maxOccurrenceMinutes: 480,   // مدة المرة الواحدة لمهمة متكررة أكثر من يوم عمل كامل
        maxDailyRepetitions: 200,    // عدد مرات يومي مرتفع بشكل غير معتاد
        dominantShare: 0.4,          // مهمة واحدة تستهلك أكثر من 40% من إجمالي الساعات
        minTitleWords: 3             // وصف مختصر جدًا
      },
      /** بدايات صياغة عامة تحتاج تحديد المخرج أو الموضوع. */
      vagueStarts: ['القيام ب', 'المشاركة في', 'المساهمة في', 'المساعدة في', 'أي مهام', 'أعمال أخرى', 'مهام أخرى', 'ما يكلف به', 'تنفيذ ما']
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
