/**
 * sample-data.js — بيانات تجريبية واقعية لتجربة النظام واختباره.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});

  WL.sampleProject = function () {
    var M = WL.model;
    var p = M.createProject();
    p.org.entity = 'جامعة ........';   // اسم غير حقيقي: البيانات التجريبية لا تنتمي لجهة معروفة
    p.org.department = 'الإدارة العامة للموارد البشرية';
    p.org.actualCount = '6';

    var SRV = M.addSection(p, 'وحدة الخدمات الوظيفية').section.id;
    var PAY = M.addSection(p, 'وحدة الرواتب').section.id;
    var DEV = M.addSection(p, 'وحدة التخطيط والتطوير').section.id;

    var ADMIN = 'مساعد إداري', HR = 'أخصائي موارد بشرية', PAYROLL = 'محاسب رواتب';
    [
      ['استقبال ومراجعة طلبات الموظفين الواردة عبر النظام', SRV, ADMIN, 'daily', 20, 10],
      ['إدخال وتحديث بيانات الموظفين في نظام الموارد البشرية', DEV, HR, 'daily', 15, 12],
      ['الرد على استفسارات الموظفين الهاتفية والبريدية', SRV, ADMIN, 'daily', 25, 6],
      ['إعداد خطابات التعريف والشهادات الوظيفية', SRV, ADMIN, 'daily', 8, 15],
      ['التدقيق على قرارات شؤون الموظفين قبل اعتمادها', SRV, HR, 'daily', 6, 20],
      ['حفظ وأرشفة الملفات الورقية والإلكترونية', SRV, ADMIN, 'daily', 10, 8],
      ['مراجعة سجلات الحضور والانصراف ومعالجة الملاحظات', PAY, PAYROLL, 'weekly', 10, 45],
      ['متابعة إجراءات التعيين والمباشرة', SRV, HR, 'weekly', 6, 90],
      ['معالجة طلبات الترقيات والعلاوات', DEV, HR, 'weekly', 5, 60],
      ['إعداد مسيرات العمل الإضافي والانتدابات', PAY, PAYROLL, 'semimonthly', 4, 120],
      ['إعداد تقرير الإجازات والغياب', PAY, PAYROLL, 'monthly', 3, 180],
      ['إعداد مسير الرواتب ومطابقته', PAY, PAYROLL, 'monthly', 2, 480],
      ['إعداد تقارير الأداء الربعية', DEV, HR, 'quarterly', 3, 600],
      ['تحديث الهيكل التنظيمي وبطاقات الوصف الوظيفي', DEV, HR, 'semiannual', 2, 900],
      ['إعداد الخطة السنوية للاحتياج التدريبي', DEV, HR, 'annual', 1, 2400],
      ['تقييم الأداء الوظيفي السنوي ومتابعة اعتماده', DEV, HR, 'annual', 1, 4800]
    ].forEach(function (t) {
      M.addTask(p, { title: t[0], sectionId: t[1], positionTitle: t[2], frequencyKey: t[3],
        repetitions: String(t[4]), durationMinutes: String(t[5]) });
    });
    return p;
  };
})(typeof window !== 'undefined' ? window : globalThis);
