/**
 * utilities.js — دوال مساعدة عامة لا تعرف شيئًا عن منطق النظام.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});

  var ARABIC_DIGITS = '٠١٢٣٤٥٦٧٨٩';
  var PERSIAN_DIGITS = '۰۱۲۳۴۵۶۷۸۹';

  /** يحوّل الأرقام العربية/الفارسية والفاصلة العربية إلى صيغة قياسية. */
  function normalizeDigits(text) {
    return String(text)
      .replace(/[٠-٩]/g, function (d) { return String(ARABIC_DIGITS.indexOf(d)); })
      .replace(/[۰-۹]/g, function (d) { return String(PERSIAN_DIGITS.indexOf(d)); })
      .replace(/[٫,]/g, '.')
      .replace(/\s+/g, '');
  }

  /**
   * يقرأ قيمة رقمية من إدخال المستخدم.
   * يعيد null إذا كان الحقل فارغًا، و NaN إذا كان النص غير رقمي.
   */
  function parseNumber(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === 'number') return isFinite(value) ? value : NaN;
    var text = normalizeDigits(value);
    if (text === '') return null;
    if (!/^[-+]?(\d+\.?\d*|\.\d+)$/.test(text)) return NaN;
    return Number(text);
  }

  function isValidNumber(n) {
    return typeof n === 'number' && isFinite(n);
  }

  /** تقريب مطابق لدالة ROUND في Excel (0.5 يُجبر بعيدًا عن الصفر). */
  function roundHalfUp(value, digits) {
    var f = Math.pow(10, digits || 0);
    var sign = value < 0 ? -1 : 1;
    var abs = Number((Math.abs(value) * f).toPrecision(15));
    return (sign * Math.round(abs)) / f;
  }

  var numberFormatters = {};
  /** تنسيق للعرض فقط — لا يغيّر القيمة الداخلية. */
  function formatNumber(value, digits) {
    if (!isValidNumber(value)) return '—';
    var key = String(digits);
    if (!numberFormatters[key]) {
      numberFormatters[key] = new Intl.NumberFormat('en-US', {
        minimumFractionDigits: digits,
        maximumFractionDigits: digits
      });
    }
    return numberFormatters[key].format(value);
  }

  function formatPercent(value, digits) {
    if (!isValidNumber(value)) return '—';
    return formatNumber(value * 100, digits === undefined ? 1 : digits) + '%';
  }

  function todayISO() {
    var d = new Date();
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }

  function pad2(n) { return n < 10 ? '0' + n : String(n); }

  /** يحوّل 'YYYY-MM-DD' إلى Date بتوقيت UTC (لتجنب إزاحة اليوم داخل Excel). */
  function isoToUtcDate(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    if (!m) return null;
    return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  }

  /** ينظف النص ليكون صالحًا كاسم ملف في ويندوز وماك. */
  function sanitizeFileName(text) {
    return String(text || '')
      .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '')
      .replace(/\s+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^[_.]+|[_.]+$/g, '')
      .slice(0, 80);
  }

  function uid() {
    return 't' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function escapeHtml(text) {
    return String(text === null || text === undefined ? '' : text)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  WL.utils = {
    normalizeDigits: normalizeDigits,
    parseNumber: parseNumber,
    isValidNumber: isValidNumber,
    roundHalfUp: roundHalfUp,
    formatNumber: formatNumber,
    formatPercent: formatPercent,
    todayISO: todayISO,
    isoToUtcDate: isoToUtcDate,
    sanitizeFileName: sanitizeFileName,
    uid: uid,
    deepClone: deepClone,
    escapeHtml: escapeHtml
  };
})(typeof window !== 'undefined' ? window : globalThis);
