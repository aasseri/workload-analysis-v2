/**
 * ai.test.js — اختبارات المساعد الذكي: الفحص المحلي، والتعليمات، ومحوّلا Claude وOpenAI (بعميل وهمي دون استدعاء فعلي).
 */
(function () {
  'use strict';
  var WL = window.WL, cfg = WL.config, M = WL.model;

  function projectWith(tasks) {
    var p = M.createProject();
    p.org.entity = 'جهة اختبار';
    p.org.jobTitle = 'أخصائي شؤون موظفين';
    p.org.actualCount = '3';
    tasks.forEach(function (t) {
      M.addTask(p, { title: t[0], frequencyKey: t[1], repetitions: String(t[2]), durationMinutes: String(t[3]) });
    });
    return p;
  }

  function mockClient(reply, capture) {
    return { beta: { messages: { create: async function (params) {
      if (capture) capture.params = params;
      if (reply instanceof Error) throw reply;
      return reply;
    } } } };
  }
  function toolReply(input) {
    return { stop_reason: 'tool_use', content: [{ type: 'tool_use', id: 'tu_1', name: 'submit_task_analysis', input: input }] };
  }
  function mockFetch(status, json, capture) {
    return async function (url, init) {
      if (capture) { capture.url = url; capture.init = init; }
      if (status === 'throw') throw new TypeError('Failed to fetch');
      return { ok: status >= 200 && status < 300, status: status, json: async function () { return json; } };
    };
  }
  function llmTask(n, extra) {
    var t = { task_number: n, quality: 'needs_improvement', clarity: 'يحتاج تحسين', measurability: 'يحتاج تحسين',
      output_clarity: 'غير مكتمل', job_relevance: 'مرتبط', improved_task: '', improvement_reason: '', output: '',
      measurement_unit: '', suggested_frequency: '', frequency_confidence: '', frequency_reason: '',
      missing_information: [], question: '', value_warnings: [], composite: false, split_suggestions: [], recommendations: [] };
    Object.keys(extra || {}).forEach(function (k) { t[k] = extra[k]; });
    return t;
  }
  function analysisInput(tasks, extra) {
    var o = { tasks: tasks, overlaps: [], possible_missing_tasks: [], top_recommendations: [] };
    Object.keys(extra || {}).forEach(function (k) { o[k] = extra[k]; });
    return o;
  }

  window.__runAiTests = async function (check) {
    var R = WL.aiRules, AI = WL.ai;

    // ── الفحص المحلي
    check('تنظيف: ترقيم + القيام بـ + نقطة', R.mechanicalFix('1- القيام بإعداد التقرير الشهري.'), 'إعداد التقرير الشهري');
    check('تنظيف: يقوم الموظف بـ', R.mechanicalFix('يقوم الموظف بمراجعة العقود'), 'مراجعة العقود');
    check('تنظيف: لا يمس رقمًا جزءًا من المعنى', R.mechanicalFix('5 تقارير أسبوعية للإدارة'), '5 تقارير أسبوعية للإدارة');

    var p = projectWith([
      ['متابعة', 'daily', 5, 10],
      ['المشاركة في الاجتماعات الدورية للإدارة', 'weekly', 1, 60],
      ['1- القيام بإعداد التقرير الشهري.', 'monthly', 1, 120],
      ['مراجعة ملفات الموظفين وتحديثها', 'daily', 2, 600],
      ['الرد على المكالمات الهاتفية الواردة', 'daily', 250, 1],
      ['مراجعة ملفات الموظفين وتحديث بياناتها', 'weekly', 1, 30]
    ]);
    var loc = R.analyze(p, cfg);
    function lr(i) { return loc.tasks.filter(function (r) { return r.taskId === p.tasks[i].id; })[0]; }
    check('محلي: وصف مختصر = غير قابلة للقياس', lr(0).quality, 'not_measurable');
    check('محلي: سؤال للمهمة المختصرة', !!lr(0).question, true);
    check('محلي: صياغة عامة = تحتاج تحسين', lr(1).quality, 'needs_improvement');
    check('محلي: تنظيف الصياغة', lr(2).improvedTask, 'إعداد التقرير الشهري');
    check('محلي: مدة 600 دقيقة يومية', lr(3).valueWarnings.length > 0, true);
    check('محلي: تكرار يومي 250', lr(4).valueWarnings.some(function (w) { return /250/.test(w.text); }), true);
    check('محلي: لا اقتراح أرقام إطلاقًا', loc.tasks.every(function (r) { return !r.suggestedFrequencyKey; }), true);
    check('محلي: تداخل بين مهمتين متشابهتين', loc.overlaps.some(function (o) {
      return o.taskIds.indexOf(p.tasks[3].id) >= 0 && o.taskIds.indexOf(p.tasks[5].id) >= 0;
    }), true);
    check('تشابه: مثال البرومبت (التقارير الشهرية)', R.similarity('إعداد التقارير الشهرية', 'إعداد التقارير الشهرية للإدارة').overlap >= 0.75, true);
    check('تشابه: مهمتان مختلفتان', R.similarity('إعداد مسير الرواتب ومطابقته', 'إعداد تقرير الإجازات والغياب').overlap < 0.75, true);
    var sample = R.analyze(WL.sampleProject(), cfg);
    check('محلي: كل مهام البيانات التجريبية جيدة', sample.tasks.filter(function (r) { return r.quality === 'good'; }).length, 16);
    check('محلي: لا تداخل في البيانات التجريبية', sample.overlaps.length, 0);
    check('محلي: مهمة واحدة فقط', R.analyze(p, cfg, { taskId: p.tasks[4].id }).tasks.length, 1);

    var noKey = await AI.run(p, cfg, { mode: 'improve', taskId: p.tasks[2].id, settings: null });
    check('بدون مفتاح: نتيجة محلية', noKey.llmUsed + '/' + noKey.tasks.length + '/' + noKey.tasks[0].source, 'false/1/rules');

    // ── التعليمات
    var sys = AI.buildSystem(cfg, 'tool');
    check('تعليمات: تبدأ ببرومبت المستخدم', sys.indexOf('# MASTER PROMPT') === 0, true);
    check('تعليمات: القسم 41 موجود', sys.indexOf('# 41. القاعدة النهائية') > 0, true);
    check('تعليمات: قيمة 1023 من الإعدادات', sys.indexOf('ساعات العمل الفعلية السنوية للموظف: 1023') > 0, true);
    check('تعليمات: اسم الأداة', sys.indexOf('submit_task_analysis') > 0, true);
    check('تعليمات: وضع JSON لمزوّدي OpenAI', AI.buildSystem(cfg, 'json').indexOf('أعد كائن JSON واحدًا') > 0, true);
    var um = AI.buildUserMessage(p, cfg, 'measure', 1);
    check('رسالة: تعليمات الزر برقم المهمة', um.indexOf('جعل المهمة رقم 2 قابلة للقياس') > 0, true);
    check('رسالة: المسمى الوظيفي', um.indexOf('المسمى الوظيفي: أخصائي شؤون موظفين') > 0, true);
    check('رسالة: الساعات المحسوبة', /"annual_hours_computed": 200/.test(um), true);

    // سلامة المخطط الصارم: كل كائن required = كل الحقول و additionalProperties = false
    var schemaOk = true;
    (function walk(s) {
      if (s.type === 'object') {
        var keys = Object.keys(s.properties).sort().join();
        if (s.additionalProperties !== false || s.required.slice().sort().join() !== keys) schemaOk = false;
        Object.keys(s.properties).forEach(function (k) { walk(s.properties[k]); });
      } else if (s.type === 'array') walk(s.items);
    })(AI.analysisSchema(cfg));
    check('مخطط: صالح للوضع الصارم', schemaOk, true);

    // ── Claude (عميل وهمي)
    var ANT = { provider: 'anthropic', apiKey: 'sk-ant-test-0000000000', model: '', baseUrl: '' };
    var cap = {};
    var input = analysisInput([
      llmTask(4, { improved_task: 'مراجعة ملفات الموظفين الورقية وتحديث بياناتها في النظام', improvement_reason: 'تحديد النطاق',
        suggested_frequency: 'أسبوعي', frequency_confidence: 'متوسط', frequency_reason: 'نشاط دوري',
        missing_information: ['متوسط عدد الملفات'], question: 'كم ملفًا يُراجع أسبوعيًا؟', value_warnings: ['المدة مرتفعة'] }),
      llmTask(1, { improved_task: 'مهمة أخرى' })
    ]);
    var res = await AI.run(p, cfg, { mode: 'improve', taskId: p.tasks[3].id, settings: ANT, client: mockClient(toolReply(input), cap) });
    var prm = cap.params || {};
    check('Claude: النموذج', prm.model, 'claude-opus-5');
    check('Claude: التعليمات مخزنة مؤقتًا', prm.system && prm.system[0].cache_control && prm.system[0].cache_control.type, 'ephemeral');
    check('Claude: أداة strict', prm.tools[0].strict + '/' + prm.tools[0].name, 'true/submit_task_analysis');
    check('Claude: fallbacks default', prm.fallbacks + '/' + (prm.betas || []).join(), 'default/server-side-fallback-2026-07-01');
    check('Claude: تعليمات الزر في الرسالة', prm.messages[0].content.indexOf('تحسين المهمة رقم 4') > 0, true);
    check('Claude: نجاح', res.llmUsed, true);
    check('Claude: نتيجة المهمة المختارة فقط', res.tasks.length + '/' + res.tasks[0].taskId, '1/' + p.tasks[3].id);
    var r0 = res.tasks[0];
    check('Claude: الصياغة المقترحة', r0.improvedTask, 'مراجعة ملفات الموظفين الورقية وتحديث بياناتها في النظام');
    check('Claude: التردد المقترح إلى مفتاح', r0.suggestedFrequencyKey + '/' + r0.frequencyConfidence, 'weekly/متوسط');
    check('Claude: السؤال والنواقص', r0.question + '|' + r0.missing.length, 'كم ملفًا يُراجع أسبوعيًا؟|1');
    check('Claude: دمج ملاحظة القيم المحلية', r0.valueWarnings.some(function (w) { return w.source === 'rules'; }) &&
      r0.valueWarnings.some(function (w) { return w.source === 'llm'; }), true);
    check('Claude: اسم المزوّد', r0.providerLabel, 'Claude (Anthropic)');

    var same = await AI.run(p, cfg, { mode: 'improve', taskId: p.tasks[1].id, settings: ANT,
      client: mockClient(toolReply(analysisInput([llmTask(2, { suggested_frequency: 'أسبوعي', frequency_confidence: 'مرتفع' })]))) });
    check('Claude: تردد مطابق للحالي لا يُقترح', same.tasks[0].suggestedFrequencyKey, '');

    var rev = await AI.run(p, cfg, { mode: 'review', settings: ANT, client: mockClient(toolReply(analysisInput(
      [llmTask(1, { quality: 'not_measurable' }), llmTask(2)],
      { overlaps: [{ task_numbers: [4, 6], type: 'تشابه جزئي', note: 'نفس النشاط' }, { task_numbers: [9, 1], type: 'تكرار كامل', note: 'x' }],
        possible_missing_tasks: ['قد تكون هناك مهمة أرشفة'], top_recommendations: ['أ', 'ب', 'ج', 'د'] }))) });
    check('مراجعة: كل المهام (النموذج + محلي لما لم يُرجعه)', rev.tasks.length, 6);
    check('مراجعة: ترتيب حسب الجدول', rev.tasks[0].taskId, p.tasks[0].id);
    check('مراجعة: المهام غير المرجعة من الفحص المحلي', rev.tasks[2].source, 'rules');
    check('مراجعة: تداخل صالح فقط + عدم تكرار المحلي', rev.overlaps.length, 1);
    check('مراجعة: مهام محتملة', rev.missingTasks[0], 'قد تكون هناك مهمة أرشفة');
    check('مراجعة: 3 توصيات كحد أقصى', rev.topRecommendations.length, 3);

    var r2 = await AI.run(p, cfg, { mode: 'analyze', taskId: p.tasks[0].id, settings: ANT, client: mockClient({ stop_reason: 'refusal', content: [] }) });
    check('خطأ: الرفض مع نتيجة محلية', r2.llmError + '|' + r2.tasks.length, AI.MESSAGES.REFUSAL + '|1');
    var r3 = await AI.run(p, cfg, { mode: 'review', settings: ANT, client: mockClient({ stop_reason: 'max_tokens', content: [] }) });
    check('خطأ: قطع الإخراج', r3.llmError, AI.MESSAGES.TRUNCATED);
    var r4 = await AI.run(p, cfg, { mode: 'recommend', taskId: p.tasks[0].id, settings: ANT, client: mockClient(new Error('x')) });
    check('خطأ: عام برسالة عربية', r4.llmError, AI.MESSAGES.GENERIC);

    // ── OpenAI ومتوافقاته (fetch وهمي)
    var OAI = { provider: 'openai', apiKey: 'sk-proj-test000000000', model: '', baseUrl: '' };
    var content = JSON.stringify(analysisInput([llmTask(1, { improved_task: 'متابعة المعاملات الواردة وإحالتها للأقسام المختصة' })]));
    function reply200(msg, finish) { return { choices: [{ finish_reason: finish || 'stop', message: msg }] }; }
    function oRun(settings, f, mode) { return AI.run(p, cfg, { mode: mode || 'improve', taskId: p.tasks[0].id, settings: settings, fetch: f }); }
    var ocap = {};
    var ro = await oRun(OAI, mockFetch(200, reply200({ content: content, refusal: null }), ocap));
    var ob = JSON.parse(ocap.init.body);
    check('OpenAI: العنوان', ocap.url, 'https://api.openai.com/v1/chat/completions');
    check('OpenAI: التفويض', ocap.init.headers.Authorization, 'Bearer sk-proj-test000000000');
    check('OpenAI: النموذج الافتراضي', ob.model, 'gpt-6-astra');
    check('OpenAI: مخطط صارم', ob.response_format.type + '/' + ob.response_format.json_schema.strict, 'json_schema/true');
    check('OpenAI: التعليمات كاملة', ob.messages[0].content.indexOf('# MASTER PROMPT') === 0, true);
    check('OpenAI: نجاح', ro.llmUsed + '|' + ro.tasks[0].improvedTask, 'true|متابعة المعاملات الواردة وإحالتها للأقسام المختصة');
    check('OpenAI: اسم المزوّد', ro.tasks[0].providerLabel, 'OpenAI');
    var fence = String.fromCharCode(96, 96, 96);
    check('OpenAI: JSON داخل أسوار', (await oRun(OAI, mockFetch(200, reply200({ content: fence + 'json\n' + content + '\n' + fence })))).llmUsed, true);
    check('OpenAI: 401', (await oRun(OAI, mockFetch(401, { error: { code: 'invalid_api_key' } }))).llmError, AI.MESSAGES.AUTH);
    check('OpenAI: رصيد غير كافٍ', (await oRun(OAI, mockFetch(429, { error: { code: 'insufficient_quota' } }))).llmError, AI.MESSAGES.QUOTA);
    check('OpenAI: حد الطلبات', (await oRun(OAI, mockFetch(429, { error: { code: 'rate_limit_exceeded' } }))).llmError, AI.MESSAGES.RATE);
    check('OpenAI: نموذج غير موجود', (await oRun(OAI, mockFetch(404, { error: { code: 'model_not_found' } }))).llmError, AI.MESSAGES.MODEL);
    check('OpenAI: اتصال أو مفتاح', (await oRun(OAI, mockFetch('throw'))).llmError, AI.MESSAGES.CONNECTION_OR_KEY);
    check('OpenAI: رفض', (await oRun(OAI, mockFetch(200, reply200({ content: null, refusal: 'no' })))).llmError, AI.MESSAGES.REFUSAL);
    check('OpenAI: قطع', (await oRun(OAI, mockFetch(200, reply200({ content: '{"t' }, 'length')))).llmError, AI.MESSAGES.TRUNCATED);
    check('OpenAI: نص غير JSON', (await oRun(OAI, mockFetch(200, reply200({ content: 'مرحبا' })))).llmError, AI.MESSAGES.NO_RESULT);
    var ccap = {};
    await oRun({ provider: 'custom', apiKey: 'AIzaTestKey123', model: 'my-model-1', baseUrl: 'https://example.com/v1beta/openai/' },
      mockFetch(200, reply200({ content: content }), ccap));
    check('متوافق: العنوان والنموذج', ccap.url + '|' + JSON.parse(ccap.init.body).model, 'https://example.com/v1beta/openai/chat/completions|my-model-1');

    // ── القِدم
    var st = ro.tasks[0];
    check('قِدم: صالح قبل التعديل', AI.isStale(p, st, 'title'), false);
    M.updateTask(p, p.tasks[0].id, { repetitions: '9' });
    check('قِدم: تعديل التكرار لا يُبطل الصياغة', AI.isStale(p, st, 'title'), false);
    M.updateTask(p, p.tasks[0].id, { title: 'متابعة البريد' });
    check('قِدم: تعديل الوصف يُبطل الصياغة', AI.isStale(p, st, 'title'), true);

    // ── الإعدادات
    check('تعرّف: Anthropic / OpenAI / آخر', [AI.detectProvider('sk-ant-api03-x'), AI.detectProvider('sk-proj-a'), AI.detectProvider('AIza1')].join(), 'anthropic,openai,');
    check('تحقق: مفتاح OpenAI سليم', AI.validateSettings(cfg, OAI), null);
    check('تحقق: مفتاح Anthropic في خانة OpenAI مرفوض', AI.validateSettings(cfg, { provider: 'openai', apiKey: 'sk-ant-xxxxxxxxxxxxxxxx' }) !== null, true);
    check('تحقق: متوافق بلا عنوان مرفوض', AI.validateSettings(cfg, { provider: 'custom', apiKey: 'abcdefghij', model: 'm', baseUrl: '' }) !== null, true);
    var testCfg = JSON.parse(JSON.stringify(cfg));
    testCfg.ai.settingsStorageKey = 'wl.ai.test.settings';
    testCfg.ai.legacyKeyStorageKey = 'wl.ai.test.legacy';
    localStorage.setItem(testCfg.ai.legacyKeyStorageKey, 'sk-ant-legacy000000000');
    var mig = AI.getSettings(testCfg);
    check('ترحيل: المفتاح القديم يُقرأ كـ Claude', mig && mig.provider, 'anthropic');
    AI.saveSettings(testCfg, OAI);
    check('ترحيل: الحفظ يحذف القديم', localStorage.getItem(testCfg.ai.legacyKeyStorageKey), null);
    AI.clearSettings(testCfg);
    check('حذف: لا إعدادات', AI.getSettings(testCfg), null);
  };
})();
