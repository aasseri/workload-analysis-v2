/**
 * ai-assistant.js — طبقة المساعد الذكي (AI Task Assistant).
 * تقع بين واجهة الإدخال ومحرك الحساب: تحلل المهام وتقترح تحسينات، ولا تعدّل المشروع بنفسها
 * ولا تمس قواعد الحساب.
 *
 * - التعليمات الأساسية وتعليمات كل زر في ai-prompt.js (قابلة للتعديل).
 * - الفحص المحلي (ai-rules.js) يعمل دائمًا دون إنترنت ويُدمج مع نتيجة النموذج.
 * - المزوّدون: Claude عبر مكتبة Anthropic الرسمية، أو OpenAI/متوافق عبر Chat Completions.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var Rules = WL.aiRules;

  // ───────────── الإعدادات (محفوظة في هذا المتصفح فقط) ─────────────

  function storageGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function storageSet(k, v) { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } }
  function storageDel(k) { try { localStorage.removeItem(k); } catch (e) { /* لا شيء */ } }

  /** @returns {{provider, apiKey, model, baseUrl}|null} */
  function getSettings(cfg) {
    var raw = storageGet(cfg.ai.settingsStorageKey);
    if (raw) {
      try {
        var s = JSON.parse(raw);
        if (s && s.apiKey && cfg.ai.providers[s.provider]) return s;
      } catch (e) { /* إعدادات تالفة */ }
    }
    var legacy = storageGet(cfg.ai.legacyKeyStorageKey);
    if (legacy) return { provider: 'anthropic', apiKey: legacy, model: '', baseUrl: '' };
    return null;
  }

  function saveSettings(cfg, s) {
    storageDel(cfg.ai.legacyKeyStorageKey);
    return storageSet(cfg.ai.settingsStorageKey, JSON.stringify({
      provider: s.provider, apiKey: String(s.apiKey).trim(),
      model: String(s.model || '').trim(), baseUrl: String(s.baseUrl || '').trim().replace(/\/+$/, '')
    }));
  }

  function clearSettings(cfg) {
    storageDel(cfg.ai.settingsStorageKey);
    storageDel(cfg.ai.legacyKeyStorageKey);
  }

  /** الإعدادات الفعلية بعد تطبيق القيم الافتراضية للمزوّد. */
  function resolve(cfg, s) {
    var p = cfg.ai.providers[s.provider];
    return {
      provider: s.provider, def: p, apiKey: s.apiKey,
      model: s.model || p.defaultModel,
      baseUrl: (s.baseUrl || p.baseUrl || '').replace(/\/+$/, ''),
      label: s.provider === 'custom' && s.model ? s.model : p.label
    };
  }

  /** يتحقق من الإعدادات قبل الحفظ ويعيد رسالة خطأ عربية أو null. */
  function validateSettings(cfg, s) {
    var p = cfg.ai.providers[s.provider];
    if (!p) return 'اختر مزوّد الخدمة.';
    var key = String(s.apiKey || '').trim();
    if (!key) return 'يرجى إدخال المفتاح.';
    if (!new RegExp(p.keyPattern).test(key)) {
      return s.provider === 'anthropic' ? 'صيغة المفتاح غير صحيحة. يبدأ مفتاح Anthropic بـ ⁦sk-ant-⁩'
        : s.provider === 'openai' ? 'صيغة المفتاح غير صحيحة. يبدأ مفتاح OpenAI بـ ⁦sk-⁩'
          : 'صيغة المفتاح غير صحيحة.';
    }
    if (s.provider === 'custom') {
      if (!/^https:\/\/\S+$/.test(String(s.baseUrl || '').trim())) return 'أدخل عنوان الخدمة ويبدأ بـ ⁦https://⁩';
      if (!String(s.model || '').trim()) return 'أدخل اسم النموذج.';
    }
    return null;
  }

  /** يقترح المزوّد من شكل المفتاح الملصوق. */
  function detectProvider(key) {
    var k = String(key || '').trim();
    if (/^sk-ant-/.test(k)) return 'anthropic';
    if (/^sk-/.test(k)) return 'openai';
    return null;
  }

  // ───────────── الأخطاء ─────────────

  var MESSAGES = {
    SDK_LOAD: 'تعذر تحميل مكتبة المزوّد. تحقق من الاتصال بالإنترنت.',
    AUTH: 'مفتاح API غير صحيح أو منتهي. راجع المفتاح في إعدادات المساعد.',
    PERMISSION: 'المفتاح لا يملك صلاحية استخدام النموذج.',
    MODEL: 'النموذج المحدد غير متاح لهذا المفتاح. راجع اسم النموذج في الإعدادات.',
    QUOTA: 'رصيد حساب المزوّد غير كافٍ. أضف رصيدًا من لوحة المزوّد.',
    RATE: 'تم تجاوز حد الطلبات المسموح. حاول بعد دقيقة.',
    OVERLOADED: 'خدمة الذكاء الاصطناعي مشغولة حاليًا. حاول بعد قليل.',
    CONNECTION: 'تعذر الاتصال بخدمة الذكاء الاصطناعي. تحقق من الاتصال بالإنترنت.',
    CONNECTION_OR_KEY: 'تعذر إكمال الطلب. تحقق من الاتصال بالإنترنت ومن صحة المفتاح.',
    BAD_REQUEST: 'رفض المزوّد الطلب. تأكد أن النموذج يدعم المخرجات المنظمة (JSON Schema).',
    REFUSAL: 'لم يتمكن النموذج من معالجة هذا الطلب.',
    TRUNCATED: 'قائمة المهام طويلة جدًا لتحليلها دفعة واحدة. حلّل المهام على دفعات.',
    NO_RESULT: 'لم يُرجع النموذج نتيجة قابلة للقراءة. حاول مرة أخرى.',
    GENERIC: 'حدث خطأ أثناء التحليل الذكي. حاول مرة أخرى.'
  };

  function aiError(code) {
    var e = new Error(MESSAGES[code] || MESSAGES.GENERIC);
    e.code = code;
    return e;
  }

  // ───────────── المخطط والرسائل (مشتركة بين المزوّدين) ─────────────

  var TOOL_NAME = 'submit_task_analysis';
  var MODES = ['improve', 'analyze', 'measure', 'recommend', 'overlap', 'review'];
  var QUALITY = ['good', 'needs_improvement', 'not_measurable'];
  var CONFIDENCE = ['', 'مرتفع', 'متوسط', 'منخفض'];
  var OVERLAP_TYPES = ['تكرار كامل', 'تشابه جزئي', 'يمكن دمجها'];

  function str(desc) { return { type: 'string', description: desc }; }
  function strList(desc) { return { type: 'array', items: { type: 'string' }, description: desc }; }
  function strEnum(values, desc) { return { type: 'string', enum: values, description: desc }; }

  function analysisSchema(cfg) {
    var labels = cfg.frequencies.map(function (f) { return f.label; });
    var taskProps = {
      task_number: { type: 'integer', description: 'رقم المهمة كما ورد في القائمة' },
      quality: strEnum(QUALITY, 'تصنيف جودة المهمة'),
      clarity: strEnum(['جيد', 'يحتاج تحسين', 'غير واضح'], 'وضوح المهمة'),
      measurability: strEnum(['جيد', 'يحتاج تحسين', 'غير قابل للقياس حاليًا'], 'قابلية القياس'),
      output_clarity: strEnum(['واضح', 'غير مكتمل', 'غير محدد'], 'وجود مخرج واضح'),
      job_relevance: strEnum(['مرتبط', 'يحتاج تحقق', 'غير واضح'], 'الارتباط بالمسمى الوظيفي والسياق'),
      improved_task: str('الصياغة المقترحة أو نص فارغ'),
      improvement_reason: str('سبب التعديل أو نص فارغ'),
      output: str('المخرج أو النتيجة التي يمكن التحقق منها، أو نص فارغ'),
      measurement_unit: str('وحدة القياس (ما الذي يُعد) أو نص فارغ'),
      suggested_frequency: strEnum([''].concat(labels), 'تردد مقترح أو نص فارغ'),
      frequency_confidence: strEnum(CONFIDENCE, 'مستوى الثقة في التردد المقترح'),
      frequency_reason: str('سبب التردد المقترح أو نص فارغ'),
      missing_information: strList('البيانات الناقصة'),
      question: str('أهم سؤال واحد قصير أو نص فارغ'),
      value_warnings: strList('ملاحظات تحقق على قيم المستخدم المرتفعة أو غير المنطقية'),
      composite: { type: 'boolean', description: 'هل المهمة مركبة من عدة أنشطة مستقلة' },
      split_suggestions: strList('الأنشطة المقترح فصلها'),
      recommendations: strList('توصيات قصيرة')
    };
    return {
      type: 'object',
      properties: {
        tasks: {
          type: 'array',
          items: { type: 'object', properties: taskProps, required: Object.keys(taskProps), additionalProperties: false }
        },
        overlaps: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              task_numbers: { type: 'array', items: { type: 'integer' } },
              type: strEnum(OVERLAP_TYPES, 'نوع التداخل'),
              note: str('ملاحظة قصيرة')
            },
            required: ['task_numbers', 'type', 'note'],
            additionalProperties: false
          }
        },
        possible_missing_tasks: strList('مهام محتملة للمراجعة بصيغة «قد تكون...»'),
        top_recommendations: strList('أهم التوصيات للقائمة كاملة (3 كحد أقصى)')
      },
      required: ['tasks', 'overlaps', 'possible_missing_tasks', 'top_recommendations'],
      additionalProperties: false
    };
  }

  /** أداة Claude (مخرجات منظمة عبر strict tool use). */
  function buildTool(cfg) {
    return {
      name: TOOL_NAME,
      description: 'Submit the structured analysis of the workload tasks for the requested operation.',
      strict: true,
      input_schema: analysisSchema(cfg)
    };
  }

  /** @param {string} outputMode 'tool' (Claude) أو 'json' (مزوّدو OpenAI) */
  function buildSystem(cfg, outputMode) {
    return WL.aiPrompt.master + '\n\n' + WL.aiPrompt.systemAppendix(cfg, outputMode);
  }

  function buildUserMessage(project, cfg, mode, focusIndex) {
    var o = project.org;
    var analysis = WL.calc.analyze(project, cfg);
    var ctxLines = [
      ['الجهة (المجال الإداري)', o.entity], ['الإدارة', o.department], ['الأقسام / الوحدات', (o.sections || []).map(function (x) { return x.name; }).join('، ')],
      ['المسمى الوظيفي', o.jobTitle],
      ['العدد الفعلي', o.actualCount]
    ].map(function (p) { return p[0] + ': ' + (String(p[1] || '').trim() || 'غير متوفر'); });

    var list = project.tasks.map(function (t, i) {
      var f = WL.calc.findFrequency(cfg, t.frequencyKey);
      var row = analysis.tasks[i];
      return {
        task_number: i + 1,
        title: t.title,
        section: row ? row.sectionName : '',
        actual_job_title: String(t.positionTitle || ''),
        frequency: f ? f.label : '',
        repetitions: String(t.repetitions),
        duration_minutes: String(t.durationMinutes),
        annual_hours_computed: row && row.hours !== null ? Math.round(row.hours * 100) / 100 : null
      };
    });

    var instruction = WL.aiPrompt.modes[mode].replace(/\{n\}/g, String(focusIndex + 1));
    return 'بيانات الوظيفة:\n' + ctxLines.join('\n') + '\n\n' +
      'قائمة المهام (' + list.length + ') — القيم كما أدخلها المستخدم، والساعات السنوية محسوبة بمحرك النظام:\n' +
      JSON.stringify(list, null, 1) + '\n\n' + instruction;
  }

  // ───────────── تحويل رد النموذج إلى نتائج ─────────────

  function cleanList(a) {
    return Array.isArray(a) ? a.map(function (x) { return String(x || '').trim(); }).filter(Boolean) : [];
  }

  function toResult(raw, project, cfg, mode, focusTaskId, providerLabel) {
    var byLabel = {};
    cfg.frequencies.forEach(function (f) { byLabel[f.label] = f.key; });
    var tasks = project.tasks;
    var out = { tasks: [], overlaps: [], missingTasks: [], topRecommendations: [] };

    (raw && Array.isArray(raw.tasks) ? raw.tasks : []).forEach(function (t) {
      var task = tasks[Number(t.task_number) - 1];
      if (!task) return;
      if (focusTaskId && task.id !== focusTaskId) return;
      if (out.tasks.some(function (r) { return r.taskId === task.id; })) return;
      var r = Rules.emptyResult(task, 'llm');
      r.providerLabel = providerLabel;
      r.quality = QUALITY.indexOf(t.quality) >= 0 ? t.quality : '';
      r.clarity = String(t.clarity || '');
      r.measurability = String(t.measurability || '');
      r.outputClarity = String(t.output_clarity || '');
      r.jobRelevance = String(t.job_relevance || '');
      var improved = String(t.improved_task || '').trim().slice(0, cfg.ui.maxTaskTitleLength);
      if (improved && improved !== String(task.title).trim()) {
        r.improvedTask = improved;
        r.improvementReason = String(t.improvement_reason || '').trim();
      }
      r.output = String(t.output || '').trim();
      r.unit = String(t.measurement_unit || '').trim();
      var fk = byLabel[t.suggested_frequency];
      if (fk && fk !== task.frequencyKey) {
        r.suggestedFrequencyKey = fk;
        r.frequencyConfidence = CONFIDENCE.indexOf(t.frequency_confidence) > 0 ? t.frequency_confidence : '';
        r.frequencyReason = String(t.frequency_reason || '').trim();
      }
      r.missing = cleanList(t.missing_information);
      r.question = String(t.question || '').trim();
      r.valueWarnings = cleanList(t.value_warnings).map(function (x) { return { text: x, source: 'llm' }; });
      r.composite = !!t.composite;
      r.splitSuggestions = r.composite ? cleanList(t.split_suggestions) : [];
      r.recommendations = cleanList(t.recommendations);
      out.tasks.push(r);
    });

    (raw && Array.isArray(raw.overlaps) ? raw.overlaps : []).forEach(function (ov) {
      var ids = (Array.isArray(ov.task_numbers) ? ov.task_numbers : [])
        .map(function (n) { return tasks[Number(n) - 1]; }).filter(Boolean)
        .map(function (t) { return t.id; })
        .filter(function (id, i, a) { return a.indexOf(id) === i; });
      if (ids.length < 2) return;
      if (focusTaskId && ids.indexOf(focusTaskId) < 0) return;
      out.overlaps.push({ taskIds: ids, type: OVERLAP_TYPES.indexOf(ov.type) >= 0 ? ov.type : 'تشابه جزئي',
        note: String(ov.note || '').trim(), source: 'llm' });
    });

    if (mode === 'review') {
      out.missingTasks = cleanList(raw && raw.possible_missing_tasks);
      out.topRecommendations = cleanList(raw && raw.top_recommendations).slice(0, 3);
    }
    return out;
  }

  // ───────────── محوّل Claude (مكتبة Anthropic الرسمية) ─────────────

  var sdkPromise = null;
  function loadAnthropicSdk(url) {
    if (!sdkPromise) {
      sdkPromise = import(url)
        .then(function (mod) { return mod.default || mod.Anthropic; })
        .catch(function () { sdkPromise = null; throw aiError('SDK_LOAD'); });
    }
    return sdkPromise;
  }

  function mapAnthropicError(Anthropic, err) {
    if (err && err.code && MESSAGES[err.code]) return err;
    if (Anthropic) {
      if (err instanceof Anthropic.AuthenticationError) return aiError('AUTH');
      if (err instanceof Anthropic.PermissionDeniedError) return aiError('PERMISSION');
      if (err instanceof Anthropic.NotFoundError) return aiError('MODEL');
      if (err instanceof Anthropic.RateLimitError) return aiError('RATE');
      if (err instanceof Anthropic.APIConnectionError) return aiError('CONNECTION');
      if (err instanceof Anthropic.APIError && err.status >= 500) return aiError('OVERLOADED');
    }
    return aiError('GENERIC');
  }

  async function runAnthropic(conn, cfg, userMessage, client) {
    var Anthropic = null;
    if (!client) {
      Anthropic = await loadAnthropicSdk(conn.def.sdkUrl);
      client = new Anthropic({ apiKey: conn.apiKey, dangerouslyAllowBrowser: true });
    }
    var response;
    try {
      response = await client.beta.messages.create({
        model: conn.model,
        max_tokens: conn.def.maxTokens,
        // التعليمات الأساسية ثابتة بين الطلبات، فتُخزَّن مؤقتًا لتقليل التكلفة
        system: [{ type: 'text', text: buildSystem(cfg, 'tool'), cache_control: { type: 'ephemeral' } }],
        messages: [{ role: 'user', content: userMessage }],
        tools: [buildTool(cfg)],
        tool_choice: { type: 'auto' },
        // عند رفض الطلب من مصنفات الأمان يُعاد تلقائيًا على النموذج البديل الموصى به
        betas: ['server-side-fallback-2026-07-01'],
        fallbacks: 'default'
      });
    } catch (err) {
      throw mapAnthropicError(Anthropic, err);
    }
    if (response.stop_reason === 'refusal') throw aiError('REFUSAL');
    var call = (response.content || []).filter(function (b) { return b.type === 'tool_use' && b.name === TOOL_NAME; })[0];
    if (!call) throw aiError(response.stop_reason === 'max_tokens' ? 'TRUNCATED' : 'NO_RESULT');
    return call.input;
  }

  // ───────────── محوّل OpenAI ومتوافقاته (Chat Completions) ─────────────

  function openAiStatusError(status, body) {
    var code = body && body.error && body.error.code;
    if (status === 401) return aiError('AUTH');
    if (status === 403) return aiError('PERMISSION');
    if (status === 404) return aiError('MODEL');
    if (status === 429) return aiError(code === 'insufficient_quota' ? 'QUOTA' : 'RATE');
    if (status === 400 && code === 'model_not_found') return aiError('MODEL');
    if (status === 400) return aiError('BAD_REQUEST');
    if (status >= 500) return aiError('OVERLOADED');
    return aiError('GENERIC');
  }

  /** يستخرج كائن JSON من نص الرد (يتحمل الأسوار ```json التي تضيفها بعض النماذج). */
  function parseJsonContent(text) {
    var t = String(text || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
    try { return JSON.parse(t); } catch (e) { /* محاولة ثانية */ }
    var a = t.indexOf('{'), b = t.lastIndexOf('}');
    if (a >= 0 && b > a) {
      try { return JSON.parse(t.slice(a, b + 1)); } catch (e2) { /* ليس JSON */ }
    }
    return null;
  }

  async function runOpenAiChat(conn, cfg, userMessage, fetchImpl) {
    var doFetch = fetchImpl || root.fetch.bind(root);
    var body = {
      model: conn.model,
      messages: [
        { role: 'system', content: buildSystem(cfg, 'json') },
        { role: 'user', content: userMessage }
      ],
      response_format: {
        type: 'json_schema',
        json_schema: { name: 'task_analysis', strict: true, schema: analysisSchema(cfg) }
      }
    };

    var res;
    try {
      res = await doFetch(conn.baseUrl + '/chat/completions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + conn.apiKey },
        body: JSON.stringify(body)
      });
    } catch (err) {
      // بعض المزوّدين (منهم OpenAI) لا يرسلون ترويسة CORS مع ردود الخطأ،
      // فيظهر المفتاح الخاطئ في المتصفح كفشل اتصال؛ لذلك تذكر الرسالة الاحتمالين.
      throw aiError('CONNECTION_OR_KEY');
    }

    var data = null;
    try { data = await res.json(); } catch (e) { /* رد غير JSON */ }
    if (!res.ok) throw openAiStatusError(res.status, data);

    var choice = data && data.choices && data.choices[0];
    if (!choice || !choice.message) throw aiError('NO_RESULT');
    if (choice.message.refusal) throw aiError('REFUSAL');
    if (choice.finish_reason === 'length') throw aiError('TRUNCATED');
    var parsed = parseJsonContent(choice.message.content);
    if (!parsed) throw aiError('NO_RESULT');
    return parsed;
  }

  // ───────────── التحليل ─────────────

  /**
   * يطلب التحليل من النموذج.
   * @param {object} opts { mode, taskId?, settings?, client? (عميل Anthropic وهمي), fetch? (fetch وهمي) }
   */
  async function llmAnalyze(project, cfg, opts) {
    var settings = opts.settings || getSettings(cfg);
    var conn = resolve(cfg, settings);
    var focusIndex = opts.taskId ? WL.model.indexOfTask(project, opts.taskId) : 0;
    var userMessage = buildUserMessage(project, cfg, opts.mode, focusIndex);
    var raw = conn.def.api === 'anthropic'
      ? await runAnthropic(conn, cfg, userMessage, opts.client)
      : await runOpenAiChat(conn, cfg, userMessage, opts.fetch);
    return toResult(raw, project, cfg, opts.mode, opts.taskId || null, conn.label);
  }

  function overlapKey(ov) { return ov.taskIds.slice().sort().join('|'); }

  /** يدمج الفحص المحلي مع نتيجة النموذج: النموذج أساس، وتُضاف ملاحظات القيم والتداخلات المحلية غير المكررة. */
  function merge(rulesRes, llmRes) {
    var byTask = {};
    rulesRes.tasks.forEach(function (r) { byTask[r.taskId] = r; });
    llmRes.tasks.forEach(function (r) {
      var local = byTask[r.taskId];
      if (!local) return;
      local.valueWarnings.forEach(function (w) { r.valueWarnings.push(w); });
      if (!r.improvedTask && local.improvedTask) {
        r.improvedTask = local.improvedTask;
        r.improvementReason = local.improvementReason;
        r.improvedSource = 'rules';
      }
    });
    var seen = {};
    llmRes.overlaps.forEach(function (o) { seen[overlapKey(o)] = true; });
    rulesRes.overlaps.forEach(function (o) { if (!seen[overlapKey(o)]) llmRes.overlaps.push(o); });
    return llmRes;
  }

  /**
   * ينفذ عملية أحد الأزرار الستة.
   * @param {object} opts { mode: 'improve'|'analyze'|'measure'|'recommend'|'overlap'|'review', taskId? }
   * @returns {Promise<{mode, taskId, tasks, overlaps, missingTasks, topRecommendations, llmUsed, llmError, providerLabel}>}
   */
  async function run(project, cfg, opts) {
    var o = opts || {};
    if (MODES.indexOf(o.mode) < 0) throw new Error('UNKNOWN_MODE');
    var taskId = o.mode === 'review' ? null : o.taskId;
    var rulesRes = Rules.analyze(project, cfg, { taskId: taskId });
    var base = { mode: o.mode, taskId: taskId || null, missingTasks: [], topRecommendations: [],
      llmUsed: false, llmError: null, providerLabel: '' };

    var settings = o.settings || getSettings(cfg);
    if (!settings) return Object.assign(base, rulesRes);
    try {
      var llmRes = await llmAnalyze(project, cfg, Object.assign({}, o, { taskId: taskId, settings: settings }));
      var merged = merge(rulesRes, llmRes);
      // مهمة لم يُرجعها النموذج تبقى بنتيجة الفحص المحلي
      rulesRes.tasks.forEach(function (r) {
        if (!merged.tasks.some(function (x) { return x.taskId === r.taskId; })) merged.tasks.push(r);
      });
      var order = {};
      project.tasks.forEach(function (t, i) { order[t.id] = i; });
      merged.tasks.sort(function (a, b) { return order[a.taskId] - order[b.taskId]; });
      return Object.assign(base, merged, { llmUsed: true, providerLabel: resolve(cfg, settings).label });
    } catch (err) {
      return Object.assign(base, rulesRes, { llmError: err.message || MESSAGES.GENERIC });
    }
  }

  /** هل تغيّر حقل المهمة الذي بُني عليه الاقتراح منذ التحليل؟ field: 'title' | 'frequencyKey' */
  function isStale(project, result, field) {
    var i = WL.model.indexOfTask(project, result.taskId);
    if (i < 0) return true;
    return String(project.tasks[i][field]) !== String(result.snapshot[field]);
  }

  WL.ai = {
    MESSAGES: MESSAGES,
    MODES: MODES,
    getSettings: getSettings,
    saveSettings: saveSettings,
    clearSettings: clearSettings,
    validateSettings: validateSettings,
    detectProvider: detectProvider,
    resolve: resolve,
    isEnabled: function (cfg) { return !!getSettings(cfg); },
    analysisSchema: analysisSchema,
    buildTool: buildTool,
    buildSystem: buildSystem,
    buildUserMessage: buildUserMessage,
    parseJsonContent: parseJsonContent,
    toResult: toResult,
    run: run,
    isStale: isStale
  };
})(typeof window !== 'undefined' ? window : globalThis);
