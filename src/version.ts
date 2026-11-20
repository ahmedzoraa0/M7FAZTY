export const APP_VERSION = '2.1.0';
export const SCRIPT_VERSION = '2.1.0';

export const RELEASE_NOTES: string[] = [
  '10 ثيمات كاملة الهوية تغيّر الألوان والخطوط والأيقونات والحواف (CRT، ماين كرافت، ليجو، كلاسيكي تراثي، أركيد وغيرها)',
  'إعادة ترتيب شاشة الإعدادات مع شبكة ثيمات مصغّرة وأيقونات متناسقة',
  'سداد الديون والتنبيهات مباشرة عبر شاشة الحركة مع تحديث الشيت تلقائياً',
  'إحصائيات شهرية وسنوية كاملة مع مقارنة الشهور والتحكم في إظهار الأقسام',
  'إصلاح أنواع الحسابات + محرر مؤمّن بـ 3 خطوات مع قائمة منسدلة في الشيت',
  'نظام فحص التحديثات والتحديث المباشر من GitHub مع الحفاظ على ربط الشيت والإعدادات',
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
