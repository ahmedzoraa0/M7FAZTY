export const APP_VERSION = '2.3.0';
export const SCRIPT_VERSION = '2.1.0';
export const DEFAULT_GITHUB_REPO = 'ahmedzoraa0/M7FAZTY';

export const RELEASE_NOTES: string[] = [
  'تحديث كامل لواجهة الملف المستقل: شريط تنقل سفلي ثابت على الهاتف بدلاً من الزر العائم لمنع التداخل',
  'قائمة جانبية ثابتة دائمة على شاشات الكمبيوتر لسهولة التنقل المباشر بين الأقسام',
  'تجميع سجل الحركات حسب التاريخ (اليوم، أمس، التواريخ) مع تنبيه واضح للطلبات المعلّقة',
  'تأكيد حذف مباشر داخل بطاقات المستحقات والتنبيهات وتفاصيل الحركات بدون نوافذ متصفح منبثقة',
  'شبكة متجاوبة لأزرار المبالغ السريعة ودعم سطرين لأسماء الحسابات الطويلة وتحسين بطاقة الرصيد الرئيسية',
  'بدء نظيف فارغ تماماً + وضع تجربة معزول ببيانات افتراضية وتحسينات كاملة لإتاحة الوصول',
];

/**
 * Normalizes a GitHub repository URL or "owner/repo" string into "owner/repo".
 * Returns empty string if input is empty or invalid.
 */
export function normalizeGithubRepo(input: string): string {
  const raw = String(input || '').trim();
  if (!raw) return '';
  const cleaned = raw
    .replace(/^https?:\/\/(www\.)?github\.com\//i, '')
    .replace(/^github\.com\//i, '')
    .replace(/\.git$/i, '')
    .replace(/^\/+|\/+$/g, '');
  const parts = cleaned.split('/').filter(Boolean);
  if (parts.length < 2) return '';
  const owner = parts[0].replace(/[^a-zA-Z0-9._-]/g, '');
  const repo = parts[1].replace(/[^a-zA-Z0-9._-]/g, '');
  return owner && repo ? `${owner}/${repo}` : '';
}
