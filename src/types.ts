export type TransactionType = 'صرف' | 'قبض' | 'تحويل بين حسابات';

export type AccountKind = 'محفظة' | 'بنك' | 'خزنة' | 'بطاقة' | 'أخرى';

export interface AccountItem {
  row?: number;
  name: string;
  kind?: AccountKind;
  opening?: number;
  incoming?: number;
  outgoing?: number;
  current: number;
}

export interface TransactionItem {
  row?: number;
  date: string;
  type: string;
  category: string;
  desc: string;
  fromAcc: string;
  toAcc: string;
  amount: number;
  notes?: string;
}

export interface DebtItem {
  row?: number;
  id: string;
  date: string;
  type: 'ليا' | 'عليا' | string;
  desc: string;
  amount: number;
  status: 'مستحق' | 'تم السداد' | string;
  notes?: string;
}

export interface ReminderItem {
  row?: number;
  id: string;
  dueDate: string;
  desc: string;
  amount: number;
  repeat: string;
  status: 'قادم' | 'تم الدفع' | string;
  notes?: string;
}

export interface CategoriesState {
  expense: string[];
  income: string[];
}

export interface FinanceState {
  accounts: AccountItem[];
  categories: CategoriesState;
  paymentMethods: string[];
  recent: TransactionItem[];
  debts: DebtItem[];
  reminders: ReminderItem[];
}

export interface SheetInspectionReport {
  spreadsheetTitle: string;
  spreadsheetUrl: string;
  sheetNames: string[];
  missingSheets: string[];
  accountCount: number;
  firstTxRow: number;
  totalRow: number;
  txCount: number;
  expenseCatCount: number;
  incomeCatCount: number;
  payMethodCount: number;
  debtCount: number;
  reminderCount: number;
}
