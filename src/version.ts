export const APP_VERSION = '2.2.0';
export const SCRIPT_VERSION = '2.1.0';
export const DEFAULT_GITHUB_REPO = 'ahmedzoraa0/M7FAZTY';

export const RELEASE_NOTES: string[] = [
  'بدء نظيف فارغ تماماً بدون أي بيانات مسبقة أو اتصالات شبكة قبل الربط',
  'إضافة وضع تجربة صريح ببيانات افتراضية معزولة تماماً عن جداول جوجل',
  'تحسين عرض الهاتف: منع تداخل الزر العائم، شبكة كاملة لأزرار المبالغ السريعة، ودعم سطرين لأسماء الحسابات',
  'تحسين إتاحة الوصول: السماح بتكبير الشاشة، مؤشرات التركيز للوحة المفاتيح، ودعم تقليل الحركة',
  'إضافة تنبيه أمني واضح حول طبيعة رمز القفل داخل شاشة الإعدادات',
  'تسريع التحميل وتقليل الخطوط المحمّلة وحل تعارض تثبيت الحزم',
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
