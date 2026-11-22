import { FinanceState } from '../types';

export const DEFAULT_SCRIPT_URL = '';

export const DEFAULT_SPREADSHEET_ID = '';

export function cleanNumber(val: unknown): number {
  if (typeof val === 'number') return Number.isFinite(val) ? val : 0;
  if (!val) return 0;
  const str = String(val)
    .replace(/٫/g, '.')
    .replace(/[^\d.-]/g, '');
  const n = parseFloat(str);
  return Number.isNaN(n) ? 0 : n;
}

export function makeRequestQuery(paramsObj: Record<string, string | number | undefined>): string {
  const params = new URLSearchParams();
  Object.entries(paramsObj).forEach(([k, v]) => {
    if (v !== undefined && v !== null) params.set(k, String(v));
  });
  if (!params.get('requestId')) {
    const id =
      typeof crypto !== 'undefined' && crypto.randomUUID
        ? crypto.randomUUID()
        : `req-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    params.set('requestId', id);
  }
  return params.toString();
}

export function fetchViaAppsScript(scriptUrl: string, token = ''): Promise<FinanceState> {
  return new Promise((resolve, reject) => {
    if (!scriptUrl || scriptUrl.startsWith('PASTE_NEW_')) {
      return reject(new Error('رابط Apps Script غير مضبوط'));
    }
    if (!navigator.onLine) {
      return reject(new Error('الجهاز غير متصل بالإنترنت'));
    }

    const cbName = `__mfRead_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const params = new URLSearchParams({
      callback: cbName,
      _: String(Date.now()),
    });
    if (token) params.set('token', token);

    const script = document.createElement('script');
    let settled = false;

    const cleanup = () => {
      script.remove();
      try {
        delete (window as unknown as Record<string, unknown>)[cbName];
      } catch {
        (window as unknown as Record<string, unknown>)[cbName] = undefined;
      }
    };

    const timer = window.setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(new Error('انتهت مهلة انتظار رد Google Apps Script'));
    }, 15000);

    (window as unknown as Record<string, (data: unknown) => void>)[cbName] = (raw: unknown) => {
      if (settled) return;
      settled = true;
      window.clearTimeout(timer);
      cleanup();
      const data = raw as Record<string, unknown> | null;
      if (!data || data.status !== 'success') {
        return reject(new Error(String(data?.message || 'رفض السكربت القراءة')));
      }

      const accounts = Array.isArray(data.accounts)
        ? data.accounts.map((a: Record<string, unknown>) => ({
            row: typeof a.row === 'number' ? a.row : undefined,
            name: String(a.name || ''),
            kind: (a.kind as FinanceState['accounts'][0]['kind']) || undefined,
            opening: a.opening !== undefined ? cleanNumber(a.opening) : undefined,
            current: cleanNumber(a.current),
          }))
        : [];

      const cats = (data.categories as Record<string, unknown>) || {};
      const expense = Array.isArray(cats.expense) ? cats.expense.map(String) : [];
      const income = Array.isArray(cats.income) ? cats.income.map(String) : [];
      const paymentMethods = Array.isArray(data.paymentMethods)
        ? data.paymentMethods.map(String)
        : accounts.map((a) => a.name);

      resolve({
        accounts,
        categories: { expense, income },
        paymentMethods,
        recent: Array.isArray(data.recent)
          ? data.recent.map((t: Record<string, unknown>) => ({
              row: typeof t.row === 'number' ? t.row : undefined,
              date: String(t.date || ''),
              type: String(t.type || ''),
              category: String(t.category || ''),
              desc: String(t.desc || ''),
              fromAcc: String(t.fromAcc || ''),
              toAcc: String(t.toAcc || ''),
              amount: cleanNumber(t.amount),
              notes: String(t.notes || ''),
            }))
          : [],
        debts: Array.isArray(data.debts)
          ? data.debts.map((d: Record<string, unknown>) => ({
              row: typeof d.row === 'number' ? d.row : undefined,
              id: String(d.id || ''),
              date: String(d.date || ''),
              type: String(d.type || ''),
              desc: String(d.desc || ''),
              amount: cleanNumber(d.amount),
              status: String(d.status || 'مستحق'),
              notes: String(d.notes || ''),
            }))
          : [],
        reminders: Array.isArray(data.reminders)
          ? data.reminders.map((r: Record<string, unknown>) => ({
              row: typeof r.row === 'number' ? r.row : undefined,
              id: String(r.id || ''),
              dueDate: String(r.dueDate || ''),
              desc: String(r.desc || ''),
              amount: cleanNumber(r.amount),
              repeat: String(r.repeat || 'شهري'),
              status: String(r.status || 'قادم'),
              notes: String(r.notes || ''),
            }))
          : [],
      });
    };

    script.onerror = () => {
      if (settled) return;
      settled = true;
      window.clearTimeout(timer);
      cleanup();
      reject(new Error('تعذر الوصول إلى رابط Google Apps Script'));
    };

    script.src = `${scriptUrl}?${params.toString()}`;
    script.async = true;
    document.body.appendChild(script);
  });
}

export function writeViaAppsScript(
  scriptUrl: string,
  bodyQuery: string,
  token = ''
): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    if (!scriptUrl || scriptUrl.startsWith('PASTE_NEW_')) {
      return reject(new Error('ضع رابط Apps Script أولاً'));
    }
    if (!navigator.onLine) {
      return reject(new Error('الجهاز غير متصل بالإنترنت'));
    }

    const params = new URLSearchParams(bodyQuery);
    if (token) params.set('token', token);
    const cbName = `__mfWrite_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    params.set('callback', cbName);
    params.set('_', String(Date.now()));

    const script = document.createElement('script');
    let settled = false;

    const finish = (err: Error | null, data?: Record<string, unknown>) => {
      if (settled) return;
      settled = true;
      window.clearTimeout(timer);
      script.remove();
      try {
        delete (window as unknown as Record<string, unknown>)[cbName];
      } catch {
        (window as unknown as Record<string, unknown>)[cbName] = undefined;
      }
      if (err) reject(err);
      else resolve(data || {});
    };

    (window as unknown as Record<string, (data: unknown) => void>)[cbName] = (raw: unknown) => {
      const data = raw as Record<string, unknown> | null;
      if (!data || data.status !== 'success') {
        const apiError = new Error(String(data?.message || 'رفض الشيت تنفيذ العملية')) as Error & {
          apiError?: boolean;
        };
        apiError.apiError = true;
        return finish(apiError);
      }
      finish(null, data);
    };

    script.onerror = () => finish(new Error('تعذر الوصول إلى Google Apps Script'));
    const timer = window.setTimeout(
      () => finish(new Error('انتهت مهلة انتظار رد الشيت')),
      15000
    );

    script.src = `${scriptUrl}?${params.toString()}`;
    script.async = true;
    document.body.appendChild(script);
  });
}
