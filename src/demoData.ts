import { FinanceState } from './types';

/**
 * Clean empty state for fresh installations before linking a Google Sheet.
 * Contains zero personal accounts, zero balances, and zero transactions.
 */
export const EMPTY_FINANCE_STATE: FinanceState = {
  accounts: [],
  categories: {
    expense: [
      'مأكولات ومشروبات',
      'مواصلات وبنزين',
      'تسوق ومستلزمات',
      'فواتير شهرية',
      'صحة وأدوية',
      'ترفيه وخروجات',
      'أخرى',
    ],
    income: ['راتب شهري', 'عمل إضافي', 'أخرى'],
  },
  paymentMethods: [],
  recent: [],
  debts: [],
  reminders: [],
};

function offsetDateKey(daysDelta: number): string {
  const d = new Date();
  d.setDate(d.getDate() + daysDelta);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/**
 * Completely synthetic placeholder data used ONLY when the user explicitly
 * activates the labeled "وضع تجريبي" (Demo Mode). Isolated from any Google Sheet.
 */
export function createSyntheticDemoState(): FinanceState {
  const today = offsetDateKey(0);
  const yesterday = offsetDateKey(-1);
  const threeDaysAgo = offsetDateKey(-3);
  const lastWeek = offsetDateKey(-7);
  const lastMonth = offsetDateKey(-32);
  const inFiveDays = offsetDateKey(5);
  const inTwelveDays = offsetDateKey(12);

  return {
    accounts: [
      { name: 'محفظة تجريبية', kind: 'محفظة', opening: 1500, current: 2450 },
      { name: 'حساب بنكي تجريبي للادخار', kind: 'بنك', opening: 10000, current: 14800 },
      { name: 'خزنة نقدية تجريبية', kind: 'خزنة', opening: 2000, current: 1650 },
      { name: 'بطاقة دفع تجريبية', kind: 'بطاقة', opening: 1000, current: 820 },
    ],
    categories: {
      expense: [
        'مأكولات ومشروبات',
        'مواصلات وبنزين',
        'تسوق ومستلزمات',
        'فواتير شهرية',
        'صحة وأدوية',
        'ترفيه وخروجات',
        'أخرى',
      ],
      income: ['راتب شهري', 'عمل إضافي', 'أخرى'],
    },
    paymentMethods: [
      'محفظة تجريبية',
      'حساب بنكي تجريبي للادخار',
      'خزنة نقدية تجريبية',
      'بطاقة دفع تجريبية',
    ],
    recent: [
      {
        row: 101,
        date: today,
        type: 'صرف',
        category: 'مأكولات ومشروبات',
        desc: 'مشتريات بقالة تجريبية',
        fromAcc: 'خزنة نقدية تجريبية',
        toAcc: '',
        amount: 350,
        notes: 'بيان تجريبي',
      },
      {
        row: 102,
        date: today,
        type: 'صرف',
        category: 'مواصلات وبنزين',
        desc: 'مواصلات يومية تجريبية',
        fromAcc: 'محفظة تجريبية',
        toAcc: '',
        amount: 120,
        notes: '',
      },
      {
        row: 103,
        date: yesterday,
        type: 'تحويل بين حسابات',
        category: 'تحويل',
        desc: 'تحويل تجريبي بين الحسابات',
        fromAcc: 'حساب بنكي تجريبي للادخار',
        toAcc: 'محفظة تجريبية',
        amount: 1000,
        notes: '',
      },
      {
        row: 104,
        date: threeDaysAgo,
        type: 'قبض',
        category: 'راتب شهري',
        desc: 'دخل شهري تجريبي',
        fromAcc: 'حساب بنكي تجريبي للادخار',
        toAcc: '',
        amount: 8500,
        notes: '',
      },
      {
        row: 105,
        date: lastWeek,
        type: 'صرف',
        category: 'فواتير شهرية',
        desc: 'فاتورة كهرباء تجريبية',
        fromAcc: 'بطاقة دفع تجريبية',
        toAcc: '',
        amount: 480,
        notes: '',
      },
      {
        row: 106,
        date: lastMonth,
        type: 'قبض',
        category: 'عمل إضافي',
        desc: 'عائد مشروع تجريبي',
        fromAcc: 'حساب بنكي تجريبي للادخار',
        toAcc: '',
        amount: 3200,
        notes: '',
      },
      {
        row: 107,
        date: lastMonth,
        type: 'صرف',
        category: 'تسوق ومستلزمات',
        desc: 'مستلزمات منزلية تجريبية',
        fromAcc: 'بطاقة دفع تجريبية',
        toAcc: '',
        amount: 950,
        notes: '',
      },
    ],
    debts: [
      {
        row: 2,
        id: 'DEMO-D1',
        date: threeDaysAgo,
        type: 'ليا',
        desc: 'شخص تجريبي (أ)',
        amount: 600,
        status: 'مستحق',
        notes: 'بيان تجريبي',
      },
      {
        row: 3,
        id: 'DEMO-D2',
        date: lastWeek,
        type: 'عليا',
        desc: 'جهة تجريبية (ب)',
        amount: 450,
        status: 'مستحق',
        notes: 'بيان تجريبي',
      },
    ],
    reminders: [
      {
        row: 2,
        id: 'DEMO-R1',
        dueDate: inFiveDays,
        desc: 'فاتورة إنترنت تجريبية',
        amount: 320,
        repeat: 'شهري',
        status: 'قادم',
        notes: '',
      },
      {
        row: 3,
        id: 'DEMO-R2',
        dueDate: inTwelveDays,
        desc: 'اشتراك خدمات تجريبي',
        amount: 250,
        repeat: 'شهري',
        status: 'قادم',
        notes: '',
      },
    ],
  };
}
