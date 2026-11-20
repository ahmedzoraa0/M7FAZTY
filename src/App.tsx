import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  BarChart3,
  Bell,
  CheckCircle2,
  Clock,
  CreditCard,
  Database,
  Download,
  ExternalLink,
  Eye,
  EyeOff,
  FileSpreadsheet,
  FileText,
  Home,
  KeyRound,
  Landmark,
  Lock,
  Menu,
  Palette,
  Plus,
  PlusCircle,
  RefreshCw,
  Send,
  Settings,
  ShieldCheck,
  User,
  Vault,
  Wallet,
  X,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import {
  AccountItem,
  AccountKind,
  DebtItem,
  FinanceState,
  ReminderItem,
  SheetInspectionReport,
  TransactionItem,
  TransactionType,
} from './types';
import { THEMES, applyThemeToDocument } from './themes';
import {
  DEFAULT_SCRIPT_URL,
  DEFAULT_SPREADSHEET_ID,
  cleanNumber,
  fetchViaAppsScript,
  makeRequestQuery,
  writeViaAppsScript,
} from './services/appsScriptClient';
import {
  clearCachedToken,
  getAccessToken,
  googleSignIn,
  initAuth,
  logout,
} from './services/googleAuth';
import {
  DriveSheetMatch,
  addAccountDirect,
  addCategoryDirect,
  addTransactionDirect,
  buildM7faztySheetName,
  createNewFinanceSpreadsheet,
  extractSpreadsheetId,
  findExistingM7faztySheets,
  formatAndConfigureWorkbook,
  inspectSpreadsheet,
  updateAccountKindDirect,
  updateDebtStatusDirect,
  updateReminderStatusDirect,
} from './services/sheetsApi';
import { AddAccountModal } from './components/AddAccountModal';
import { AccountKindManagerModal } from './components/AccountKindManagerModal';
import { AnalyticsSection } from './components/AnalyticsSection';
import { PinLockOverlay } from './components/PinLockOverlay';
import { SheetStudioPanel } from './components/SheetStudioPanel';

type TabKey =
  | 'home'
  | 'add'
  | 'debts'
  | 'reminders'
  | 'history'
  | 'stats'
  | 'sheets'
  | 'settings';

const DEFAULT_STATE: FinanceState = {
  accounts: [
    { name: 'كاش', kind: 'خزنة', current: 4779.0 },
    { name: 'المشرق نيو', kind: 'بنك', current: 697.73 },
    { name: 'المشرق ادخار', kind: 'بنك', current: 25078.36 },
    { name: 'اتصالات كاش', kind: 'محفظة', current: 757.2 },
    { name: 'ميزة', kind: 'بطاقة', current: 915.6 },
    { name: 'ايزي باي', kind: 'بطاقة', current: 0 },
    { name: 'اخرى', kind: 'أخرى', current: 0 },
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
      'طباعة وتصميمات',
      'تبرعات',
    ],
    income: ['راتب شهري', 'تصميم', 'أخرى'],
  },
  paymentMethods: [
    'كاش',
    'المشرق نيو',
    'المشرق ادخار',
    'اتصالات كاش',
    'ميزة',
    'ايزي باي',
    'اخرى',
  ],
  recent: [],
  debts: [],
  reminders: [],
};

const fmt = (n: number) =>
  Number(n || 0).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

function localDateKey(date = new Date()): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function txMonthKey(d: string): string {
  const s = String(d || '').trim();
  let m = s.match(/^(\d{4})[-\/](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}`;
  const dt = new Date(s);
  return Number.isNaN(dt.getTime())
    ? ''
    : `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`;
}

const isDebtLya = (t: string) => ['ليا', 'مدين (ليا)', 'مدين'].includes(String(t || '').trim());
const isDebtAlya = (t: string) => ['عليا', 'داين (عليا)', 'داين'].includes(String(t || '').trim());

function accountKindIcon(kind?: AccountKind) {
  switch (kind) {
    case 'بنك':
      return <Landmark className="w-3.5 h-3.5" />;
    case 'خزنة':
      return <Vault className="w-3.5 h-3.5" />;
    case 'بطاقة':
      return <CreditCard className="w-3.5 h-3.5" />;
    default:
      return <Wallet className="w-3.5 h-3.5" />;
  }
}

export default function App() {
  // Theme state
  const [themePref, setThemePref] = useState<string>(
    () => localStorage.getItem('mf_theme') || 'ivory-gold'
  );
  const [resolvedThemeId, setResolvedThemeId] = useState<string>('ivory-gold');

  // Navigation & Drawer
  const [activeTab, setActiveTab] = useState<TabKey>('home');
  const [drawerOpen, setDrawerOpen] = useState(false);

  // Profile / Wallet name -> M7FAZTY_<Name>
  const [profileName, setProfileName] = useState<string>(
    () => localStorage.getItem('mf_profile_name') || ''
  );

  // Connectivity & Sync
  const [scriptUrl, setScriptUrl] = useState<string>(
    () => localStorage.getItem('mf_script_url') || DEFAULT_SCRIPT_URL
  );
  const [spreadsheetId, setSpreadsheetId] = useState<string>(
    () => localStorage.getItem('mf_sheet_id') || DEFAULT_SPREADSHEET_ID
  );

  const sheetStorageKey = useMemo(() => {
    const sid = extractSpreadsheetId(spreadsheetId);
    return sid || 'default';
  }, [spreadsheetId]);

  // Finance State (isolated per spreadsheetId)
  const [state, setState] = useState<FinanceState>(() => {
    const initialSid = extractSpreadsheetId(localStorage.getItem('mf_sheet_id') || DEFAULT_SPREADSHEET_ID) || 'default';
    try {
      const cached =
        localStorage.getItem(`mf_cache_v11::${initialSid}`) ||
        (initialSid === DEFAULT_SPREADSHEET_ID ? localStorage.getItem('mf_cache_v11') : null);
      if (cached) return JSON.parse(cached) as FinanceState;
    } catch {
      // ignore
    }
    return DEFAULT_STATE;
  });

  const [connStatus, setConnStatus] = useState<'online' | 'offline' | 'syncing'>('syncing');
  const [connLabel, setConnLabel] = useState<string>('جارٍ الاتصال');
  const [queue, setQueue] = useState<string[]>(() => {
    const initialSid = extractSpreadsheetId(localStorage.getItem('mf_sheet_id') || DEFAULT_SPREADSHEET_ID) || 'default';
    try {
      const raw = JSON.parse(
        localStorage.getItem(`mf_tx_queue::${initialSid}`) ||
          (initialSid === DEFAULT_SPREADSHEET_ID ? localStorage.getItem('mf_tx_queue') : null) ||
          '[]'
      );
      return Array.isArray(raw) ? raw : [];
    } catch {
      return [];
    }
  });

  // Google Sheets OAuth state
  const [googleEmail, setGoogleEmail] = useState<string | null>(null);
  const [googleToken, setGoogleToken] = useState<string | null>(() => getAccessToken());
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [inspectionReport, setInspectionReport] = useState<SheetInspectionReport | null>(null);
  const [isInspecting, setIsInspecting] = useState(false);
  const [isTestingScript, setIsTestingScript] = useState(false);
  const [existingSheets, setExistingSheets] = useState<DriveSheetMatch[]>([]);
  const [isSearchingDrive, setIsSearchingDrive] = useState(false);

  // Privacy toggles
  const [globalVisible, setGlobalVisible] = useState(false);
  const [unmaskedAccounts, setUnmaskedAccounts] = useState<Record<number, boolean>>({});
  const [overviewVisible, setOverviewVisible] = useState({ expense: false, income: false });

  // PIN Lock state (never auto-forces creating a new PIN on refresh)
  const [storedPin, setStoredPin] = useState<string | null>(() =>
    localStorage.getItem('app_user_pin')
  );
  const [pinOverlay, setPinOverlay] = useState<{
    open: boolean;
    mode: 'unlock' | 'setup' | 'change';
  }>(() => ({ open: Boolean(localStorage.getItem('app_user_pin')), mode: 'unlock' }));
  const [autoLockMins, setAutoLockMins] = useState<string>(
    () => localStorage.getItem('mf_autolock') || '0'
  );

  // Add Account Modal & Safe Account Kind Manager Modal
  const [addAccountOpen, setAddAccountOpen] = useState(false);
  const [accountKindModalOpen, setAccountKindModalOpen] = useState(false);

  // Add Transaction Form State
  const [txType, setTxType] = useState<TransactionType>('صرف');
  const [txAmount, setTxAmount] = useState('');
  const [fromAcc, setFromAcc] = useState('');
  const [toAcc, setToAcc] = useState('');
  const [txCat, setTxCat] = useState('');
  const [txDesc, setTxDesc] = useState('');
  const [txDate, setTxDate] = useState(() => localDateKey());
  const [txSaving, setTxSaving] = useState(false);
  const [editTx, setEditTx] = useState<TransactionItem | null>(null);
  const [pendingReminderPay, setPendingReminderPay] = useState<{
    id: string;
    row?: number;
    desc: string;
    amount: number;
  } | null>(null);
  const [pendingDebtSettle, setPendingDebtSettle] = useState<{
    id: string;
    row?: number;
    desc: string;
    debtType: string;
    amount: number;
  } | null>(null);

  // Inline Add Category
  const [newCatOpen, setNewCatOpen] = useState(false);
  const [newCatName, setNewCatName] = useState('');
  const [newCatSaving, setNewCatSaving] = useState(false);

  // Debts State
  const [debtType, setDebtType] = useState<'ليا' | 'عليا'>('ليا');
  const [debtDesc, setDebtDesc] = useState('');
  const [debtAmount, setDebtAmount] = useState('');
  const [debtDate, setDebtDate] = useState(() => localDateKey());
  const [debtFilter, setDebtFilter] = useState<'all' | 'مستحق' | 'ليا' | 'عليا' | 'تم السداد'>('all');
  const [debtSaving, setDebtSaving] = useState(false);

  // Reminders State
  const [remDesc, setRemDesc] = useState('');
  const [remAmount, setRemAmount] = useState('');
  const [remDueDate, setRemDueDate] = useState(() => localDateKey());
  const [remRepeat, setRemRepeat] = useState('شهري');
  const [remNotes, setRemNotes] = useState('');
  const [remSaving, setRemSaving] = useState(false);

  // History & Analytics State
  const [historySearch, setHistorySearch] = useState('');
  const [historyFilter, setHistoryFilter] = useState<'all' | 'صرف' | 'قبض' | 'تحويل'>('all');
  const [selectedTx, setSelectedTx] = useState<TransactionItem | null>(null);
  const [budgets, setBudgets] = useState<Record<string, number>>(() => {
    const initialSid = extractSpreadsheetId(localStorage.getItem('mf_sheet_id') || DEFAULT_SPREADSHEET_ID) || 'default';
    try {
      return JSON.parse(
        localStorage.getItem(`mf_budgets::${initialSid}`) || localStorage.getItem('mf_budgets') || '{}'
      );
    } catch {
      return {};
    }
  });
  const [budgetModalCat, setBudgetModalCat] = useState<string | null>(null);
  const [budgetModalVal, setBudgetModalVal] = useState('');

  // Toast
  const [toastMsg, setToastMsg] = useState<string | null>(null);
  const notify = useCallback((msg: string) => {
    setToastMsg(msg);
  }, []);

  useEffect(() => {
    if (!toastMsg) return;
    const t = window.setTimeout(() => setToastMsg(null), 2800);
    return () => window.clearTimeout(t);
  }, [toastMsg]);

  // Apply theme across all screens
  useEffect(() => {
    localStorage.setItem('mf_theme', themePref);
    const resolved = applyThemeToDocument(themePref);
    setResolvedThemeId(resolved.id);
  }, [themePref]);

  // Persist per-sheet cache & queue so switching spreadsheets never leaks data
  useEffect(() => {
    localStorage.setItem('mf_profile_name', profileName);
  }, [profileName]);

  useEffect(() => {
    try {
      localStorage.setItem(`mf_cache_v11::${sheetStorageKey}`, JSON.stringify(state));
    } catch {
      // ignore
    }
  }, [state, sheetStorageKey]);

  useEffect(() => {
    localStorage.setItem('mf_script_url', scriptUrl);
  }, [scriptUrl]);

  useEffect(() => {
    localStorage.setItem('mf_sheet_id', spreadsheetId);
  }, [spreadsheetId]);

  useEffect(() => {
    localStorage.setItem(`mf_tx_queue::${sheetStorageKey}`, JSON.stringify(queue));
  }, [queue, sheetStorageKey]);

  // When active spreadsheetId changes, restore that specific sheet's queue/budgets/cache cleanly
  useEffect(() => {
    try {
      const qRaw = JSON.parse(localStorage.getItem(`mf_tx_queue::${sheetStorageKey}`) || '[]');
      setQueue(Array.isArray(qRaw) ? qRaw : []);
    } catch {
      setQueue([]);
    }
    try {
      const bRaw = JSON.parse(localStorage.getItem(`mf_budgets::${sheetStorageKey}`) || '{}');
      setBudgets(bRaw && typeof bRaw === 'object' ? bRaw : {});
    } catch {
      setBudgets({});
    }
  }, [sheetStorageKey]);

  // Initialize Firebase Auth listener for Google Sheets OAuth
  useEffect(() => {
    const unsub = initAuth(
      (user, token) => {
        setGoogleEmail(user.email || user.displayName || 'حساب جوجل');
        setGoogleToken(token);
      },
      (user) => {
        setGoogleEmail(user?.email || null);
        setGoogleToken(null);
      }
    );
    return () => unsub();
  }, []);

  // Ensure default dropdown values stay valid
  const accountOptions = useMemo(() => {
    return Array.from(
      new Set([
        ...(state.paymentMethods || []),
        ...(state.accounts || []).map((a) => a.name),
      ].filter(Boolean))
    );
  }, [state.paymentMethods, state.accounts]);

  useEffect(() => {
    if (accountOptions.length > 0) {
      if (!fromAcc || !accountOptions.includes(fromAcc)) setFromAcc(accountOptions[0]);
      if (!toAcc || !accountOptions.includes(toAcc)) {
        setToAcc(accountOptions[1] || accountOptions[0]);
      }
    }
  }, [accountOptions, fromAcc, toAcc]);

  const activeCategoryList = useMemo(
    () => (txType === 'قبض' ? state.categories.income : state.categories.expense),
    [txType, state.categories]
  );

  useEffect(() => {
    if (txType !== 'تحويل بين حسابات' && activeCategoryList.length > 0) {
      if (!txCat || !activeCategoryList.includes(txCat)) {
        setTxCat(activeCategoryList[0]);
      }
    }
  }, [txType, activeCategoryList, txCat]);

  // Load data from Google Sheets (prefer direct OAuth token if available, fallback to Apps Script)
  const fetchData = useCallback(async () => {
    setConnStatus('syncing');
    setConnLabel('جارٍ المزامنة');

    if (googleToken) {
      try {
        const { report, state: directState } = await inspectSpreadsheet(
          spreadsheetId,
          googleToken
        );
        setInspectionReport(report);
        setState((prev) => ({
          accounts: directState.accounts.length ? directState.accounts : prev.accounts,
          categories: {
            expense: directState.categories.expense.length
              ? directState.categories.expense
              : prev.categories.expense,
            income: directState.categories.income.length
              ? directState.categories.income
              : prev.categories.income,
          },
          paymentMethods: directState.paymentMethods.length
            ? directState.paymentMethods
            : prev.paymentMethods,
          recent: directState.recent,
          debts: directState.debts,
          reminders: directState.reminders,
        }));
        setConnStatus('online');
        setConnLabel('متصل (Sheets API)');
        return;
      } catch (err) {
        const e = err as Error & { status?: number };
        if (e.status === 401) clearCachedToken();
      }
    }

    try {
      const data = await fetchViaAppsScript(scriptUrl);
      setState((prev) => ({
        accounts: data.accounts.length ? data.accounts : prev.accounts,
        categories: {
          expense: data.categories.expense.length
            ? data.categories.expense
            : prev.categories.expense,
          income: data.categories.income.length
            ? data.categories.income
            : prev.categories.income,
        },
        paymentMethods: data.paymentMethods.length
          ? data.paymentMethods
          : prev.paymentMethods,
        recent: data.recent,
        debts: data.debts,
        reminders: data.reminders,
      }));
      setConnStatus('online');
      setConnLabel('متصل بالشيت');
    } catch {
      setConnStatus('offline');
      setConnLabel('وضع محلي');
    }
  }, [googleToken, spreadsheetId, scriptUrl]);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  // Queue helper
  const sendOrQueue = useCallback(
    async (paramsObj: Record<string, string | number | undefined>) => {
      const req = makeRequestQuery(paramsObj);
      try {
        await writeViaAppsScript(scriptUrl, req);
        return { ok: true as const };
      } catch (err) {
        const e = err as Error & { apiError?: boolean };
        if (e.apiError) {
          return { ok: false as const, permanent: true as const, error: e };
        }
        const reqId = new URLSearchParams(req).get('requestId');
        setQueue((prev) => {
          if (reqId && prev.some((x) => new URLSearchParams(x).get('requestId') === reqId)) {
            return prev;
          }
          return [...prev.slice(-299), req];
        });
        setConnStatus('offline');
        setConnLabel('بانتظار الإرسال');
        return { ok: false as const, permanent: false as const, error: e };
      }
    },
    [scriptUrl]
  );

  const flushQueue = useCallback(async () => {
    if (!queue.length || !navigator.onLine) return;
    const left: string[] = [];
    for (let i = 0; i < queue.length; i++) {
      try {
        await writeViaAppsScript(scriptUrl, queue[i]);
      } catch (err) {
        const e = err as Error & { apiError?: boolean };
        if (e.apiError) continue;
        left.push(...queue.slice(i));
        break;
      }
    }
    setQueue(left);
    if (left.length < queue.length) {
      notify('تم إرسال الطلبات المعلّقة للشيت');
      window.setTimeout(fetchData, 600);
    }
  }, [queue, scriptUrl, notify, fetchData]);

  // Auto-lock listener
  useEffect(() => {
    let hiddenAt = 0;
    const onVis = () => {
      if (document.hidden) {
        hiddenAt = Date.now();
        return;
      }
      const mins = parseInt(autoLockMins || '0', 10);
      if (mins > 0 && hiddenAt && Date.now() - hiddenAt > mins * 60000 && storedPin) {
        setPinOverlay({ open: true, mode: 'unlock' });
      }
      flushQueue();
    };
    document.addEventListener('visibilitychange', onVis);
    return () => document.removeEventListener('visibilitychange', onVis);
  }, [autoLockMins, storedPin, flushQueue]);

  // Google Sign-In & Sheet Studio handlers
  const handleSearchExistingSheets = useCallback(
    async (nameOverride?: string): Promise<DriveSheetMatch[]> => {
      const token = googleToken || getAccessToken();
      if (!token) return [];
      setIsSearchingDrive(true);
      try {
        const res = await findExistingM7faztySheets(
          token,
          nameOverride !== undefined ? nameOverride : profileName
        );
        setExistingSheets(res.matches);
        return res.matches;
      } catch (err) {
        const e = err as Error & { status?: number };
        if (e.status === 401) clearCachedToken();
        notify('تعذر فحص Google Drive: ' + e.message);
        return [];
      } finally {
        setIsSearchingDrive(false);
      }
    },
    [googleToken, profileName, notify]
  );

  const handleSelectExistingSheet = useCallback(
    async (targetSheetId: string) => {
      const token = googleToken || getAccessToken();
      if (!token) return;
      const cleanId = extractSpreadsheetId(targetSheetId);
      setIsInspecting(true);
      try {
        const { report, state: directState } = await inspectSpreadsheet(cleanId, token);
        setSpreadsheetId(cleanId);
        setInspectionReport(report);
        setState({
          accounts: directState.accounts,
          categories: directState.categories,
          paymentMethods: directState.paymentMethods,
          recent: directState.recent,
          debts: directState.debts,
          reminders: directState.reminders,
        });
        // Extract profileName suffix if sheet title starts with M7FAZTY_
        const mTitle = String(report.spreadsheetTitle || '').trim();
        if (mTitle.toUpperCase().startsWith('M7FAZTY_')) {
          const derivedName = mTitle.slice('M7FAZTY_'.length).trim();
          if (derivedName) setProfileName(derivedName);
        }
        setConnStatus('online');
        setConnLabel('متصل (Sheets API)');
        notify(`تم تشغيل الشيت "${report.spreadsheetTitle}" وتحميل بياناته ✅`);
      } catch (err) {
        notify('تعذر تحميل بيانات الشيت: ' + (err as Error).message);
      } finally {
        setIsInspecting(false);
      }
    },
    [googleToken, notify]
  );

  const handleGoogleLogin = async () => {
    setIsLoggingIn(true);
    try {
      const res = await googleSignIn();
      if (res) {
        setGoogleEmail(res.user.email || res.user.displayName || 'حساب جوجل');
        setGoogleToken(res.accessToken);
        notify('تم الربط بحساب جوجل بنجاح');
        setIsInspecting(true);
        try {
          const [{ report, state: directState }, driveRes] = await Promise.all([
            inspectSpreadsheet(spreadsheetId, res.accessToken),
            findExistingM7faztySheets(res.accessToken, profileName).catch(() => ({
              matches: [] as DriveSheetMatch[],
              exactMatch: null,
              targetName: buildM7faztySheetName(profileName),
            })),
          ]);
          setExistingSheets(driveRes.matches);
          setInspectionReport(report);
          setState((prev) => ({
            ...prev,
            accounts: directState.accounts.length ? directState.accounts : prev.accounts,
            categories: {
              expense: directState.categories.expense.length
                ? directState.categories.expense
                : prev.categories.expense,
              income: directState.categories.income.length
                ? directState.categories.income
                : prev.categories.income,
            },
            paymentMethods: directState.paymentMethods.length
              ? directState.paymentMethods
              : prev.paymentMethods,
            recent: directState.recent,
            debts: directState.debts,
            reminders: directState.reminders,
          }));
          setConnStatus('online');
          setConnLabel('متصل (Sheets API)');
          if (driveRes.matches.length > 0) {
            notify(`تم العثور على ${driveRes.matches.length} شيت M7FAZTY في حسابك`);
          }
        } finally {
          setIsInspecting(false);
        }
      }
    } catch (err) {
      notify('تعذر تسجيل الدخول: ' + (err as Error).message);
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleGoogleLogout = async () => {
    await logout();
    setGoogleToken(null);
    setGoogleEmail(null);
    setInspectionReport(null);
    setExistingSheets([]);
    notify('تم تسجيل الخروج من حساب جوجل');
  };

  const handleInspectSheet = async () => {
    if (!googleToken) return;
    setIsInspecting(true);
    try {
      const { report, state: directState } = await inspectSpreadsheet(
        spreadsheetId,
        googleToken
      );
      setInspectionReport(report);
      setState((prev) => ({
        ...prev,
        accounts: directState.accounts.length ? directState.accounts : prev.accounts,
        categories: {
          expense: directState.categories.expense.length
            ? directState.categories.expense
            : prev.categories.expense,
          income: directState.categories.income.length
            ? directState.categories.income
            : prev.categories.income,
        },
        paymentMethods: directState.paymentMethods.length
          ? directState.paymentMethods
          : prev.paymentMethods,
        recent: directState.recent,
        debts: directState.debts,
        reminders: directState.reminders,
      }));
      notify('تم فحص وتحديث البيانات من Google Sheets مباشرة');
    } catch (err) {
      notify('تعذر قراءة الشيت: ' + (err as Error).message);
    } finally {
      setIsInspecting(false);
    }
  };

  const handleFormatWorkbook = async () => {
    if (!googleToken) return;
    try {
      const report = await formatAndConfigureWorkbook(spreadsheetId, googleToken);
      setInspectionReport(report);
      notify('تم إصلاح المعادلات والتصنيفات وتنسيق الشيت بالكامل ✅');
      await fetchData();
    } catch (err) {
      notify('فشل تنسيق الشيت: ' + (err as Error).message);
    }
  };

  const handleCreateNewSheet = async (title: string) => {
    if (!googleToken) return null;
    try {
      const finalTitle = title.trim() || buildM7faztySheetName(profileName);
      const created = await createNewFinanceSpreadsheet(googleToken, { title: finalTitle });
      setSpreadsheetId(created.spreadsheetId);
      setInspectionReport(created.report);
      notify(`تم إنشاء شيت "${finalTitle}" وتنسيقه بالكامل ✅`);
      const { state: directState } = await inspectSpreadsheet(created.spreadsheetId, googleToken);
      setState(directState);
      handleSearchExistingSheets(profileName);
      return {
        spreadsheetId: created.spreadsheetId,
        spreadsheetUrl: created.spreadsheetUrl,
      };
    } catch (err) {
      notify('تعذر إنشاء الشيت الجديد: ' + (err as Error).message);
      return null;
    }
  };

  const handleTestScriptConnection = async () => {
    if (!scriptUrl.trim()) {
      notify('أدخل رابط Apps Script أولاً');
      return;
    }
    setIsTestingScript(true);
    try {
      await writeViaAppsScript(scriptUrl.trim(), 'action=ping');
      notify('تم الاتصال برابط Apps Script بنجاح ✅');
      await fetchData();
    } catch (err) {
      notify('فشل الاتصال بـ Apps Script: ' + (err as Error).message);
    } finally {
      setIsTestingScript(false);
    }
  };

  // Add Account / Wallet / Bank / Safe
  const handleAddAccount = async (acc: { name: string; kind: AccountKind; opening: number }) => {
    if (state.accounts.some((a) => a.name.trim() === acc.name.trim())) {
      notify('هذا الحساب موجود بالفعل');
      return;
    }

    const newAccount: AccountItem = {
      name: acc.name,
      kind: acc.kind,
      opening: acc.opening,
      current: acc.opening,
    };

    setState((prev) => ({
      ...prev,
      accounts: [...prev.accounts, newAccount],
      paymentMethods: Array.from(new Set([...prev.paymentMethods, acc.name])),
    }));

    if (googleToken) {
      try {
        await addAccountDirect(spreadsheetId, googleToken, acc);
        notify(`تمت إضافة "${acc.name}" وربطه بمعادلات الشيت مباشرة`);
        window.setTimeout(fetchData, 500);
        return;
      } catch (err) {
        notify('تعذر الإضافة عبر Sheets API، جارٍ الإرسال عبر Apps Script...');
      }
    }

    const res = await sendOrQueue({
      action: 'add_account',
      accountName: acc.name,
      accountKind: acc.kind,
      openingBalance: acc.opening,
    });

    if (res.ok) {
      notify(`تمت إضافة "${acc.name}" وربطه بالشيت بنجاح`);
      window.setTimeout(fetchData, 600);
    } else if (res.permanent) {
      setState((prev) => ({
        ...prev,
        accounts: prev.accounts.filter((a) => a.name !== acc.name),
      }));
      notify('رفض الشيت إضافة الحساب: ' + res.error.message);
    }
  };

  // Safe Update Account Kind (from 3-step protected modal in Settings)
  const handleUpdateAccountKind = async (accountName: string, newKind: AccountKind) => {
    setState((prev) => ({
      ...prev,
      accounts: prev.accounts.map((a) => (a.name === accountName ? { ...a, kind: newKind } : a)),
    }));

    if (googleToken) {
      try {
        await updateAccountKindDirect(spreadsheetId, googleToken, accountName, newKind);
        notify(`تم تحديث نوع "${accountName}" إلى «${newKind}» في الشيت ✅`);
        window.setTimeout(fetchData, 500);
        return;
      } catch {
        // fallback to Apps Script
      }
    }

    const res = await sendOrQueue({
      action: 'update_account_kind',
      accountName,
      accountKind: newKind,
    });
    if (res.ok) {
      notify(`تم تحديث نوع "${accountName}" إلى «${newKind}» في الشيت ✅`);
      window.setTimeout(fetchData, 600);
    } else if (res.permanent) {
      notify('رفض الشيت تعديل النوع: ' + res.error.message);
      window.setTimeout(fetchData, 400);
    } else {
      notify(`تم حفظ التعديل محلياً وسيُرسل للشيت تلقائياً`);
    }
  };

  // Add Category
  const handleSaveNewCategory = async () => {
    const clean = newCatName.trim();
    if (!clean) {
      notify('اكتب اسم التصنيف الجديد أولاً');
      return;
    }
    const catType: 'صرف' | 'قبض' = txType === 'قبض' ? 'قبض' : 'صرف';
    const currentList = catType === 'قبض' ? state.categories.income : state.categories.expense;

    if (currentList.some((c) => c.trim() === clean)) {
      setTxCat(clean);
      setNewCatOpen(false);
      setNewCatName('');
      notify('التصنيف موجود بالفعل وتم اختياره');
      return;
    }

    setNewCatSaving(true);
    setState((prev) => ({
      ...prev,
      categories: {
        ...prev.categories,
        expense: catType === 'صرف' ? [...prev.categories.expense, clean] : prev.categories.expense,
        income: catType === 'قبض' ? [...prev.categories.income, clean] : prev.categories.income,
      },
    }));
    setTxCat(clean);
    setNewCatOpen(false);
    setNewCatName('');

    try {
      if (googleToken) {
        try {
          await addCategoryDirect(spreadsheetId, googleToken, catType, clean);
          notify(`تمت إضافة تصنيف "${clean}" في الشيت بنجاح`);
          window.setTimeout(fetchData, 500);
          return;
        } catch {
          // fallback to Apps Script
        }
      }

      const res = await sendOrQueue({
        action: 'add_category',
        catType,
        categoryName: clean,
      });

      if (res.ok) {
        notify(`تم حفظ تصنيف "${clean}" في الشيت بنجاح`);
        window.setTimeout(fetchData, 600);
      } else if (res.permanent) {
        notify('رفض الشيت إضافة التصنيف: ' + res.error.message);
      }
    } finally {
      setNewCatSaving(false);
    }
  };

  // Submit Transaction
  const handleTxSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const amt = parseFloat(txAmount);
    if (!(amt > 0)) {
      notify('أدخل مبلغاً صحيحاً');
      return;
    }
    if (txType === 'تحويل بين حسابات' && fromAcc === toAcc) {
      notify('اختر حسابين مختلفين');
      return;
    }

    setTxSaving(true);
    try {
      if (editTx?.row) {
        const isTransfer = txType === 'تحويل بين حسابات';
        const fp = [
          String(editTx.type || '').trim(),
          cleanNumber(editTx.amount),
          String(editTx.desc || '').trim(),
        ].join('|');
        const res = await sendOrQueue({
          action: 'update',
          row: editTx.row,
          fp,
          date: txDate,
          type: txType,
          category: isTransfer ? 'تحويل' : txCat,
          description: txDesc || (isTransfer ? 'تحويل' : 'معاملة'),
          fromAccount: fromAcc,
          toAccount: isTransfer ? toAcc : '',
          amount: amt,
          notes: '',
        });
        if (res.ok) {
          notify('تم تعديل الحركة في الشيت');
          setEditTx(null);
          setTxAmount('');
          setTxDesc('');
          setActiveTab('history');
          window.setTimeout(fetchData, 500);
        } else if (res.permanent) {
          notify('رفض الشيت التعديل: ' + res.error.message);
        }
        return;
      }

      const isTransfer = txType === 'تحويل بين حسابات';
      const categoryVal = isTransfer ? 'تحويل' : txCat;
      const descVal = txDesc.trim() || (isTransfer ? 'تحويل' : 'معاملة');
      const newTx: TransactionItem = {
        date: txDate,
        type: txType,
        category: categoryVal,
        desc: descVal,
        fromAcc,
        toAcc: isTransfer ? toAcc : '',
        amount: amt,
        notes: '',
      };

      const snapshotAccounts = state.accounts.map((a) => ({ ...a }));
      setState((prev) => ({
        ...prev,
        recent: [newTx, ...prev.recent],
        accounts: prev.accounts.map((a) => {
          if (txType === 'صرف' && a.name === fromAcc) return { ...a, current: a.current - amt };
          if (txType === 'قبض' && a.name === fromAcc) return { ...a, current: a.current + amt };
          if (isTransfer) {
            if (a.name === fromAcc) return { ...a, current: a.current - amt };
            if (a.name === toAcc) return { ...a, current: a.current + amt };
          }
          return a;
        }),
      }));

      setTxAmount('');
      setTxDesc('');

      let synced = false;
      if (googleToken) {
        try {
          await addTransactionDirect(spreadsheetId, googleToken, newTx);
          synced = true;
          notify('تم تسجيل الحركة في الشيت مباشرة');
        } catch {
          // fallback to Apps Script
        }
      }

      if (!synced) {
        const res = await sendOrQueue({
          action: 'add',
          date: newTx.date,
          type: newTx.type,
          category: newTx.category,
          description: newTx.desc,
          fromAccount: newTx.fromAcc,
          toAccount: newTx.toAcc,
          amount: newTx.amount,
          notes: '',
        });
        if (res.ok) {
          notify('تم تسجيل الحركة في الشيت بنجاح');
          synced = true;
        } else if (res.permanent) {
          setState((prev) => ({
            ...prev,
            accounts: snapshotAccounts,
            recent: prev.recent.filter((t) => t !== newTx),
          }));
          notify('رفض الشيت التسجيل: ' + res.error.message);
          return;
        }
      }

      if (pendingReminderPay) {
        const remTarget = pendingReminderPay;
        setPendingReminderPay(null);
        setState((prev) => ({
          ...prev,
          reminders: prev.reminders.map((r) =>
            (remTarget.id && r.id === remTarget.id) || (remTarget.row && r.row === remTarget.row)
              ? { ...r, status: 'تم الدفع' }
              : r
          ),
        }));
        let remUpdated = false;
        if (googleToken) {
          try {
            await updateReminderStatusDirect(
              spreadsheetId,
              googleToken,
              { id: remTarget.id, row: remTarget.row },
              'تم الدفع'
            );
            remUpdated = true;
          } catch {
            // fallback to Apps Script
          }
        }
        if (!remUpdated) {
          await sendOrQueue({
            action: 'update_reminder_status',
            id: remTarget.id || '',
            row: remTarget.id ? '' : remTarget.row,
            status: 'تم الدفع',
          });
        }
        notify('تم تسجيل الدفع وتحديث حالة الالتزام في الشيت ✅');
      }

      if (pendingDebtSettle) {
        const debtTarget = pendingDebtSettle;
        setPendingDebtSettle(null);
        setState((prev) => ({
          ...prev,
          debts: prev.debts.map((d) =>
            (debtTarget.id && d.id === debtTarget.id) || (debtTarget.row && d.row === debtTarget.row)
              ? { ...d, status: 'تم السداد' }
              : d
          ),
        }));
        let debtUpdated = false;
        if (googleToken) {
          try {
            await updateDebtStatusDirect(
              spreadsheetId,
              googleToken,
              { id: debtTarget.id, row: debtTarget.row },
              'تم السداد'
            );
            debtUpdated = true;
          } catch {
            // fallback to Apps Script
          }
        }
        if (!debtUpdated) {
          await sendOrQueue({
            action: 'update_debt_status',
            id: debtTarget.id || '',
            row: debtTarget.id ? '' : debtTarget.row,
            status: 'تم السداد',
          });
        }
        notify('تم تسجيل السداد وتحديث حالة المستحق في الشيت ✅');
      }

      if (synced) window.setTimeout(fetchData, 600);
    } finally {
      setTxSaving(false);
    }
  };

  const handleDeleteTx = async () => {
    if (!selectedTx?.row) return;
    const fp = [
      String(selectedTx.type || '').trim(),
      cleanNumber(selectedTx.amount),
      String(selectedTx.desc || '').trim(),
    ].join('|');
    const row = selectedTx.row;
    setSelectedTx(null);
    const res = await sendOrQueue({ action: 'delete', row, fp });
    if (res.ok) {
      notify('تم حذف الحركة من الشيت');
      window.setTimeout(fetchData, 500);
    } else if (res.permanent) {
      notify('رفض الشيت الحذف: ' + res.error.message);
    }
  };

  // Debts handlers
  const handleDebtSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const amt = parseFloat(debtAmount);
    if (!(amt > 0) || !debtDesc.trim()) return;
    setDebtSaving(true);
    try {
      const item: DebtItem = {
        id: 'D-' + Date.now(),
        date: debtDate,
        type: debtType,
        desc: debtDesc.trim(),
        amount: amt,
        status: 'مستحق',
        notes: '',
      };
      setState((prev) => ({ ...prev, debts: [item, ...prev.debts] }));
      setDebtDesc('');
      setDebtAmount('');
      const res = await sendOrQueue({
        action: 'add_debt',
        id: item.id,
        date: item.date,
        type: item.type,
        desc: item.desc,
        amount: item.amount,
        status: 'مستحق',
        notes: '',
      });
      if (res.ok) {
        notify('تم حفظ المستحق في الشيت');
        window.setTimeout(fetchData, 500);
      }
    } finally {
      setDebtSaving(false);
    }
  };

  const handleStartDebtSettle = (item: DebtItem) => {
    const isLya = isDebtLya(item.type);
    setEditTx(null);
    setPendingReminderPay(null);
    setPendingDebtSettle({
      id: item.id,
      row: item.row,
      desc: item.desc,
      debtType: item.type,
      amount: cleanNumber(item.amount),
    });
    setTxType(isLya ? 'قبض' : 'صرف');
    setTxAmount(cleanNumber(item.amount).toFixed(2));
    setTxDesc(isLya ? `تحصيل مستحق: ${item.desc}` : `سداد مستحق: ${item.desc}`);
    setTxDate(localDateKey());
    setActiveTab('add');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    notify('اختر الحساب والتصنيف واضغط تسجيل الحركة لإتمام السداد');
  };

  const toggleDebtStatus = async (item: DebtItem) => {
    const newStatus = item.status === 'تم السداد' ? 'مستحق' : 'تم السداد';
    setState((prev) => ({
      ...prev,
      debts: prev.debts.map((d) => (d === item ? { ...d, status: newStatus } : d)),
    }));
    if (googleToken) {
      try {
        await updateDebtStatusDirect(spreadsheetId, googleToken, { id: item.id, row: item.row }, newStatus);
        window.setTimeout(fetchData, 500);
        return;
      } catch {
        // fallback to Apps Script
      }
    }
    const res = await sendOrQueue({
      action: 'update_debt_status',
      id: item.id || '',
      row: item.id ? '' : item.row,
      status: newStatus,
    });
    if (res.ok) window.setTimeout(fetchData, 500);
  };

  const deleteDebt = async (item: DebtItem) => {
    setState((prev) => ({ ...prev, debts: prev.debts.filter((d) => d !== item) }));
    const res = await sendOrQueue({
      action: 'delete_debt',
      id: item.id || '',
      row: item.id ? '' : item.row,
    });
    if (res.ok) {
      notify('تم حذف المستحق');
      window.setTimeout(fetchData, 500);
    }
  };

  // Reminders handlers
  const handleReminderSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const amt = parseFloat(remAmount);
    if (!(amt > 0) || !remDesc.trim()) return;
    setRemSaving(true);
    try {
      const item: ReminderItem = {
        id: 'R-' + Date.now(),
        dueDate: remDueDate,
        desc: remDesc.trim(),
        amount: amt,
        repeat: remRepeat,
        status: 'قادم',
        notes: remNotes.trim(),
      };
      setState((prev) => ({ ...prev, reminders: [...prev.reminders, item] }));
      setRemDesc('');
      setRemAmount('');
      setRemNotes('');
      const res = await sendOrQueue({
        action: 'add_reminder',
        id: item.id,
        dueDate: item.dueDate,
        desc: item.desc,
        amount: item.amount,
        repeat: item.repeat,
        status: 'قادم',
        notes: item.notes,
      });
      if (res.ok) {
        notify('تم حفظ التنبيه في الشيت');
        window.setTimeout(fetchData, 500);
      }
    } finally {
      setRemSaving(false);
    }
  };

  const deleteReminder = async (item: ReminderItem) => {
    setState((prev) => ({ ...prev, reminders: prev.reminders.filter((r) => r !== item) }));
    const res = await sendOrQueue({
      action: 'delete_reminder',
      id: item.id || '',
      row: item.id ? '' : item.row,
    });
    if (res.ok) {
      notify('تم حذف التنبيه');
      window.setTimeout(fetchData, 500);
    }
  };

  // Computed metrics
  const totalBalance = useMemo(
    () => state.accounts.reduce((s, a) => s + cleanNumber(a.current), 0),
    [state.accounts]
  );

  const curMonth = useMemo(() => {
    const n = new Date();
    return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, '0')}`;
  }, []);

  const currentMonthTxs = useMemo(
    () => state.recent.filter((t) => txMonthKey(t.date) === curMonth),
    [state.recent, curMonth]
  );

  const monthStats = useMemo(() => {
    let expense = 0;
    let income = 0;
    currentMonthTxs.forEach((t) => {
      const a = cleanNumber(t.amount);
      if (t.type === 'صرف') expense += a;
      else if (t.type === 'قبض' || t.type === 'دخل') income += a;
    });
    return { expense, income, net: income - expense };
  }, [currentMonthTxs]);

  const debtSummary = useMemo(() => {
    let lya = 0;
    let alya = 0;
    let pending = 0;
    state.debts.forEach((d) => {
      if (d.status !== 'تم السداد') {
        pending++;
        const amt = cleanNumber(d.amount);
        if (isDebtLya(d.type)) lya += amt;
        else if (isDebtAlya(d.type)) alya += amt;
      }
    });
    return { lya, alya, net: lya - alya, pending };
  }, [state.debts]);

  const reminderSummary = useMemo(() => {
    const upcoming = state.reminders.filter((r) => r.status !== 'تم الدفع');
    const sum = upcoming.reduce((s, r) => s + cleanNumber(r.amount), 0);
    return { count: upcoming.length, sum };
  }, [state.reminders]);

  const filteredHistory = useMemo(() => {
    const q = historySearch.trim().toLowerCase();
    return state.recent.filter((t) => {
      if (!(cleanNumber(t.amount) > 0) || t.type === 'النوع') return false;
      if (historyFilter !== 'all') {
        if (historyFilter === 'تحويل') {
          if (!t.type.includes('تحويل')) return false;
        } else if (t.type !== historyFilter) {
          return false;
        }
      }
      if (!q) return true;
      return (
        t.desc.toLowerCase().includes(q) ||
        t.category.toLowerCase().includes(q) ||
        t.fromAcc.toLowerCase().includes(q) ||
        t.toAcc.toLowerCase().includes(q)
      );
    });
  }, [state.recent, historyFilter, historySearch]);

  const filteredDebts = useMemo(() => {
    if (debtFilter === 'all') return state.debts;
    if (debtFilter === 'مستحق') return state.debts.filter((d) => d.status !== 'تم السداد');
    if (debtFilter === 'تم السداد') return state.debts.filter((d) => d.status === 'تم السداد');
    if (debtFilter === 'ليا') return state.debts.filter((d) => isDebtLya(d.type));
    return state.debts.filter((d) => isDebtAlya(d.type));
  }, [state.debts, debtFilter]);

  // Export functions
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const txSheet = XLSX.utils.aoa_to_sheet([
      ['التاريخ', 'النوع', 'التصنيف', 'البيان', 'من حساب', 'إلى حساب', 'المبلغ', 'ملاحظات'],
      ...state.recent
        .slice()
        .reverse()
        .map((t) => [t.date, t.type, t.category, t.desc, t.fromAcc, t.toAcc, t.amount, t.notes || '']),
    ]);
    XLSX.utils.book_append_sheet(wb, txSheet, 'المعاملات');
    const accSheet = XLSX.utils.aoa_to_sheet([
      ['الحساب', 'النوع', 'الرصيد الحالي'],
      ...state.accounts.map((a) => [a.name, a.kind || 'محفظة', a.current]),
      ['الإجمالي', '', totalBalance],
    ]);
    XLSX.utils.book_append_sheet(wb, accSheet, 'الأرصدة');
    wb.Workbook = { Views: [{ RTL: true }] };
    XLSX.writeFile(wb, `${buildM7faztySheetName(profileName)}_${localDateKey()}.xlsx`);
    notify('تم تصدير ملف Excel بنجاح');
  };

  const handleExportCsv = () => {
    let csv = '\uFEFFالتاريخ,النوع,التصنيف,البيان,من حساب,إلى حساب,المبلغ\n';
    state.recent.forEach((t) => {
      csv += `"${t.date}","${t.type}","${t.category}","${String(t.desc || '').replace(/"/g, '""')}","${t.fromAcc}","${t.toAcc}","${t.amount}"\n`;
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
    a.download = `${buildM7faztySheetName(profileName)}_سجل_${localDateKey()}.csv`;
    a.click();
    notify('تم تصدير ملف CSV بنجاح');
  };

  const handleExportBackup = () => {
    const payload = {
      version: 'M7FAZTY_v11',
      profileName,
      exportedAt: new Date().toISOString(),
      ...state,
      budgets,
      theme: themePref,
    };
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' }));
    a.download = `${buildM7faztySheetName(profileName)}_Backup_${localDateKey()}.json`;
    a.click();
    notify('تم حفظ النسخة الاحتياطية');
  };

  const handleImportBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(String(ev.target?.result || '{}'));
        if (!Array.isArray(data.accounts) && !Array.isArray(data.recent)) {
          notify('ملف النسخة الاحتياطية غير صالح');
          return;
        }
        setState((prev) => ({
          accounts: data.accounts || prev.accounts,
          categories: data.categories || prev.categories,
          paymentMethods: data.paymentMethods || prev.paymentMethods,
          recent: data.recent || prev.recent,
          debts: data.debts || prev.debts,
          reminders: data.reminders || prev.reminders,
        }));
        if (data.budgets) {
          setBudgets(data.budgets);
          localStorage.setItem('mf_budgets', JSON.stringify(data.budgets));
        }
        if (data.theme) setThemePref(data.theme);
        notify('تمت استعادة النسخة الاحتياطية بنجاح');
      } catch {
        notify('تعذر قراءة الملف');
      } finally {
        e.target.value = '';
      }
    };
    reader.readAsText(file);
  };

  const navItems: Array<{ key: TabKey; label: string; icon: React.ReactNode; count?: number }> = [
    { key: 'home', label: 'الرئيسية', icon: <Home className="w-5 h-5" /> },
    { key: 'add', label: 'إضافة حركة', icon: <PlusCircle className="w-5 h-5" /> },
    { key: 'debts', label: 'المستحقات', icon: <Wallet className="w-5 h-5" />, count: debtSummary.pending },
    { key: 'reminders', label: 'التنبيهات', icon: <Bell className="w-5 h-5" />, count: reminderSummary.count },
    { key: 'history', label: 'السجل', icon: <FileText className="w-5 h-5" /> },
    { key: 'stats', label: 'التحليلات', icon: <BarChart3 className="w-5 h-5" /> },
    { key: 'sheets', label: 'استوديو الشيت والتصدير', icon: <FileSpreadsheet className="w-5 h-5" /> },
    { key: 'settings', label: 'الإعدادات والثيمات', icon: <Settings className="w-5 h-5" /> },
  ];

  const screenTitle = navItems.find((n) => n.key === activeTab)?.label || 'الرئيسية';
  const todayLabel = new Date().toLocaleDateString('ar-EG', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });

  const navigateTo = (tab: TabKey) => {
    setActiveTab(tab);
    setDrawerOpen(false);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  return (
    <div className="min-h-screen pb-20">
      {/* شاشة القفل المربوطة بالثيم */}
      <PinLockOverlay
        open={pinOverlay.open}
        mode={pinOverlay.mode}
        storedPin={storedPin}
        onUnlock={() => setPinOverlay((p) => ({ ...p, open: false }))}
        onSavePin={(newPin) => {
          localStorage.setItem('app_user_pin', newPin);
          setStoredPin(newPin);
          setPinOverlay({ open: false, mode: 'unlock' });
          notify('تم حفظ رمز PIN بنجاح');
        }}
        onCancel={
          pinOverlay.mode !== 'unlock'
            ? () => setPinOverlay((p) => ({ ...p, open: false }))
            : undefined
        }
        onNotify={notify}
      />

      {/* القائمة الجانبية */}
      {drawerOpen && (
        <div
          className="fixed inset-0 z-40 bg-black/45 backdrop-blur-xs"
          onClick={() => setDrawerOpen(false)}
        />
      )}
      <aside
        className={`fixed top-0 right-0 bottom-0 z-50 w-72 max-w-[85vw] p-4 border-l flex flex-col transition-transform duration-200 overflow-y-auto ${
          drawerOpen ? 'translate-x-0' : 'translate-x-full'
        }`}
        style={{
          background: 'var(--card-bg)',
          borderColor: 'var(--border)',
          boxShadow: 'var(--shadow-md)',
        }}
        aria-label="القائمة الرئيسية"
      >
        <div
          className="flex items-center justify-between pb-4 mb-3 border-b"
          style={{ borderColor: 'var(--border)' }}
        >
          <div className="flex items-center gap-2.5">
            <div
              className="w-10 h-10 rounded-xl flex items-center justify-center"
              style={{
                background: 'var(--accent-gradient)',
                color: 'var(--accent-contrast)',
              }}
            >
              <Wallet className="w-5 h-5" />
            </div>
            <div>
              <div className="text-base font-extrabold">M7FAZTY</div>
              <div className="text-xs font-bold truncate" style={{ color: 'var(--text-muted)' }}>
                {profileName.trim() ? `M7FAZTY_${profileName.trim()}` : 'الإدارة المالية المتكاملة'}
              </div>
            </div>
          </div>
          <button
            type="button"
            onClick={() => setDrawerOpen(false)}
            className="w-9 h-9 rounded-xl border flex items-center justify-center cursor-pointer"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text-muted)',
            }}
            aria-label="إغلاق القائمة"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <nav className="flex flex-col gap-1 flex-1">
          {navItems.map((item) => {
            const active = activeTab === item.key;
            return (
              <button
                key={item.key}
                type="button"
                onClick={() => navigateTo(item.key)}
                className="flex items-center justify-between px-3.5 py-2.5 rounded-xl border text-sm font-bold cursor-pointer transition-colors"
                style={{
                  background: active ? 'var(--surface-soft)' : 'transparent',
                  borderColor: active ? 'var(--primary-accent)' : 'transparent',
                  color: active ? 'var(--primary-accent)' : 'var(--text)',
                }}
              >
                <span className="flex items-center gap-3">
                  {item.icon}
                  <span>{item.label}</span>
                </span>
                {item.count !== undefined && item.count > 0 && (
                  <span className="text-xs font-extrabold tabular-nums" style={{ color: 'var(--expense)' }}>
                    {item.count}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        <div className="mt-4 pt-3 border-t space-y-2" style={{ borderColor: 'var(--border)' }}>
          <button
            type="button"
            onClick={() => {
              setDrawerOpen(false);
              setPinOverlay({ open: true, mode: storedPin ? 'unlock' : 'setup' });
            }}
            className="w-full py-2.5 rounded-xl border text-xs font-bold flex items-center justify-center gap-2 cursor-pointer"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
            }}
          >
            <Lock className="w-3.5 h-3.5" />
            <span>{storedPin ? 'قفل الشاشة' : 'تعيين رمز PIN'}</span>
          </button>
        </div>
      </aside>

      {/* الهيكل الرئيسي */}
      <div className="max-w-5xl mx-auto px-3 md:px-6 pt-3 space-y-4">
        {/* رأس الصفحة */}
        <header
          className="sticky top-2.5 z-30 flex items-center justify-between gap-3 px-3.5 py-2.5 rounded-2xl border backdrop-blur-md"
          style={{
            background: 'var(--header-bg)',
            borderColor: 'var(--border)',
            boxShadow: 'var(--shadow-sm)',
          }}
        >
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={() => setDrawerOpen(true)}
              className="w-11 h-11 rounded-xl border flex items-center justify-center cursor-pointer shrink-0"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
              aria-label="فتح القائمة"
            >
              <Menu className="w-5 h-5" />
            </button>
            <div className="min-w-0">
              <h1 className="text-lg font-extrabold truncate">{screenTitle}</h1>
              <p className="text-xs truncate" style={{ color: 'var(--text-muted)' }}>
                {todayLabel}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <span
              className="hidden sm:inline text-xs font-bold whitespace-nowrap"
              style={{
                color:
                  connStatus === 'online'
                    ? 'var(--income)'
                    : connStatus === 'offline'
                    ? 'var(--warning)'
                    : 'var(--text-muted)',
              }}
            >
              {connLabel}
            </span>
            <button
              type="button"
              onClick={() => navigateTo('sheets')}
              className="h-11 px-3 rounded-xl border text-xs font-bold hidden md:flex items-center gap-1.5 cursor-pointer whitespace-nowrap"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            >
              <FileSpreadsheet className="w-4 h-4" style={{ color: 'var(--primary-accent)' }} />
              <span>الشيت والتصدير</span>
            </button>
            <button
              type="button"
              onClick={fetchData}
              title="تحديث البيانات"
              className="w-11 h-11 rounded-xl border flex items-center justify-center cursor-pointer shrink-0"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            >
              <RefreshCw className={`w-4 h-4 ${connStatus === 'syncing' ? 'animate-spin' : ''}`} />
            </button>
          </div>
        </header>

        {/* شاشة الرئيسية */}
        {activeTab === 'home' && (
          <div className="space-y-4">
            {/* كارت إجمالي الأرصدة */}
            <div
              onClick={() => setGlobalVisible((v) => !v)}
              className="rounded-3xl p-5 border cursor-pointer select-none relative overflow-hidden"
              style={{
                background: 'var(--hero-bg)',
                borderColor: 'var(--hero-border)',
                boxShadow: 'var(--hero-shadow)',
              }}
            >
              <div className="flex items-center justify-between mb-3">
                <div className="flex items-center gap-2.5">
                  <div
                    className="w-9 h-9 rounded-xl border flex items-center justify-center"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--hero-border)',
                      color: 'var(--hero-accent)',
                    }}
                  >
                    <Wallet className="w-4 h-4" />
                  </div>
                  <span className="text-sm font-extrabold" style={{ color: 'var(--hero-text)' }}>
                    إجمالي الأرصدة
                  </span>
                </div>
                <span
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl border text-xs font-extrabold"
                  style={{
                    background: 'var(--surface-soft)',
                    borderColor: 'var(--hero-border)',
                    color: 'var(--hero-accent)',
                  }}
                >
                  {globalVisible ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                  <span>{globalVisible ? 'إخفاء' : 'إظهار'}</span>
                </span>
              </div>

              <div className="text-3xl sm:text-4xl font-extrabold font-mono-num text-right" dir="ltr" style={{ color: 'var(--hero-text)' }}>
                {globalVisible ? (
                  <>
                    {fmt(totalBalance)}{' '}
                    <span className="text-base font-bold" style={{ color: 'var(--hero-accent)' }}>
                      ج.م
                    </span>
                  </>
                ) : (
                  <span className="tracking-widest text-2xl" style={{ color: 'var(--text-muted)' }}>
                    •••• ••••
                  </span>
                )}
              </div>

              <div
                className="flex items-center justify-between pt-3 mt-3 border-t text-xs font-bold"
                style={{ borderColor: 'var(--hero-border)', color: 'var(--text-muted)' }}
              >
                <span>المحافظ النشطة: {state.accounts.length} حساب</span>
                <span dir="ltr" style={{ color: 'var(--hero-accent)' }} className="font-mono-num">
                  {globalVisible
                    ? `صافي الشهر: ${monthStats.net > 0 ? '+' : ''}${fmt(monthStats.net)} ج.م`
                    : 'اضغط للكشف السريع'}
                </span>
              </div>
            </div>

            {/* ملخص الشهر */}
            <div className="grid grid-cols-2 gap-3">
              <div
                onClick={() => setOverviewVisible((o) => ({ ...o, expense: !o.expense }))}
                className="rounded-2xl p-4 border cursor-pointer select-none"
                style={{
                  background: 'var(--card-bg)',
                  borderColor: 'var(--border)',
                  boxShadow: 'var(--shadow-md)',
                }}
              >
                <div className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                  صرف الشهر الحالي
                </div>
                <div
                  className="text-lg font-extrabold mt-1.5 text-right font-mono-num"
                  dir="ltr"
                  style={{ color: globalVisible || overviewVisible.expense ? 'var(--expense)' : 'var(--text-muted)' }}
                >
                  {globalVisible || overviewVisible.expense ? `${fmt(monthStats.expense)} ج.م` : '••••••'}
                </div>
              </div>

              <div
                onClick={() => setOverviewVisible((o) => ({ ...o, income: !o.income }))}
                className="rounded-2xl p-4 border cursor-pointer select-none"
                style={{
                  background: 'var(--card-bg)',
                  borderColor: 'var(--border)',
                  boxShadow: 'var(--shadow-md)',
                }}
              >
                <div className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                  قبض الشهر الحالي
                </div>
                <div
                  className="text-lg font-extrabold mt-1.5 text-right font-mono-num"
                  dir="ltr"
                  style={{ color: globalVisible || overviewVisible.income ? 'var(--income)' : 'var(--text-muted)' }}
                >
                  {globalVisible || overviewVisible.income ? `${fmt(monthStats.income)} ج.م` : '••••••'}
                </div>
              </div>
            </div>

            {/* شبكة الحسابات + آخر العمليات */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 items-start">
              <div className="lg:col-span-7 space-y-2.5">
                <div className="flex items-center justify-between">
                  <h2 className="text-sm font-extrabold" style={{ color: 'var(--text-muted)' }}>
                    الأرصدة والمحافظ والبنوك
                  </h2>
                  <button
                    type="button"
                    onClick={() => setAddAccountOpen(true)}
                    className="px-3 py-2 rounded-xl border text-xs font-extrabold flex items-center gap-1.5 cursor-pointer"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--primary-accent)',
                      color: 'var(--primary-accent)',
                    }}
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>إضافة محفظة / حساب</span>
                  </button>
                </div>

                <div className="grid grid-cols-2 gap-2.5">
                  {state.accounts.map((acc, idx) => {
                    const visible = globalVisible || Boolean(unmaskedAccounts[idx]);
                    const negative = acc.current < 0;
                    return (
                      <div
                        key={acc.name}
                        onClick={() =>
                          setUnmaskedAccounts((u) => ({ ...u, [idx]: !u[idx] }))
                        }
                        className={`rounded-2xl p-3.5 border cursor-pointer select-none transition-colors ${
                          idx === state.accounts.length - 1 && state.accounts.length % 2 === 1
                            ? 'col-span-2'
                            : ''
                        }`}
                        style={{
                          background: 'var(--card-bg)',
                          borderColor: 'var(--border)',
                          boxShadow: 'var(--shadow-md)',
                        }}
                      >
                        <div className="flex items-center justify-between gap-2">
                          <span className="text-sm font-extrabold truncate">{acc.name}</span>
                          <span
                            className="flex items-center gap-1 text-[11px] font-bold shrink-0"
                            style={{ color: 'var(--text-muted)' }}
                          >
                            {accountKindIcon(acc.kind)}
                            <span>{acc.kind || 'محفظة'}</span>
                          </span>
                        </div>
                        <div
                          className="text-lg font-extrabold mt-2 text-right font-mono-num"
                          dir="ltr"
                          style={{
                            color: !visible
                              ? 'var(--text-muted)'
                              : negative
                              ? 'var(--expense)'
                              : 'var(--text)',
                          }}
                        >
                          {visible ? (
                            <>
                              {negative ? '-' : ''}
                              {fmt(Math.abs(acc.current))}{' '}
                              <span className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                                ج.م
                              </span>
                            </>
                          ) : (
                            '••••••'
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div
                className="lg:col-span-5 rounded-2xl p-4 border"
                style={{
                  background: 'var(--card-bg)',
                  borderColor: 'var(--border)',
                  boxShadow: 'var(--shadow-md)',
                }}
              >
                <div className="flex items-center justify-between mb-2">
                  <h2 className="text-sm font-extrabold" style={{ color: 'var(--text-muted)' }}>
                    آخر العمليات
                  </h2>
                  <button
                    type="button"
                    onClick={() => navigateTo('history')}
                    className="text-xs font-extrabold cursor-pointer"
                    style={{ color: 'var(--primary-accent)' }}
                  >
                    عرض الكل ➔
                  </button>
                </div>

                {state.recent.length === 0 ? (
                  <div className="text-center py-8 text-xs" style={{ color: 'var(--text-muted)' }}>
                    لا توجد معاملات مسجلة بعد
                  </div>
                ) : (
                  <div className="divide-y" style={{ borderColor: 'var(--border)' }}>
                    {state.recent.slice(0, 6).map((t, idx) => (
                      <TxRow
                        key={`${t.row || idx}-${t.date}`}
                        tx={t}
                        onClick={() => t.row && setSelectedTx(t)}
                      />
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* شاشة الإضافة */}
        {activeTab === 'add' && (
          <div
            className="max-w-2xl mx-auto rounded-2xl p-5 border"
            style={{
              background: 'var(--card-bg)',
              borderColor: 'var(--border)',
              boxShadow: 'var(--shadow-md)',
            }}
          >
            <div
              className="flex p-1 rounded-xl border gap-1 mb-4"
              style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
            >
              {(['صرف', 'قبض', 'تحويل بين حسابات'] as TransactionType[]).map((t) => {
                const active = txType === t;
                const activeColor =
                  t === 'صرف'
                    ? 'var(--expense)'
                    : t === 'قبض'
                    ? 'var(--income)'
                    : 'var(--transfer)';
                return (
                  <button
                    key={t}
                    type="button"
                    onClick={() => {
                      setTxType(t);
                      setNewCatOpen(false);
                    }}
                    className="flex-1 py-2.5 rounded-lg text-sm font-bold cursor-pointer transition-all"
                    style={{
                      background: active ? 'var(--card-bg)' : 'transparent',
                      color: active ? activeColor : 'var(--text-muted)',
                      boxShadow: active ? 'var(--shadow-sm)' : 'none',
                    }}
                  >
                    {t === 'تحويل بين حسابات' ? 'تحويل' : t}
                  </button>
                );
              })}
            </div>

            {(pendingDebtSettle || pendingReminderPay) && (
              <div
                className="mb-4 p-3 rounded-xl border flex items-center justify-between gap-2.5"
                style={{
                  background: 'var(--surface-soft)',
                  borderColor: 'var(--primary-accent)',
                }}
              >
                <div className="flex items-center gap-2.5 min-w-0">
                  <div className="theme-icon-box">
                    <CheckCircle2 className="w-4 h-4" />
                  </div>
                  <div className="min-w-0">
                    <div className="text-xs font-extrabold truncate" style={{ color: 'var(--primary-accent)' }}>
                      {pendingDebtSettle
                        ? `سداد مستحق (${isDebtLya(pendingDebtSettle.debtType) ? 'مدين ليا — قبض' : 'داين عليا — صرف'}): ${pendingDebtSettle.desc}`
                        : `دفع التزام: ${pendingReminderPay?.desc}`}
                    </div>
                    <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                      اختر الحساب والتصنيف ثم اضغط تسجيل ليتم التحديث في الشيت تلقائياً
                    </div>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    setPendingDebtSettle(null);
                    setPendingReminderPay(null);
                  }}
                  className="px-2.5 py-1.5 rounded-lg border text-[11px] font-bold cursor-pointer shrink-0"
                  style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', color: 'var(--text-muted)' }}
                >
                  إلغاء الربط
                </button>
              </div>
            )}

            <form onSubmit={handleTxSubmit} className="space-y-3.5">
              <div>
                <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
                  المبلغ (ج.م)
                </label>
                <input
                  type="number"
                  step="0.01"
                  inputMode="decimal"
                  required
                  value={txAmount}
                  onChange={(e) => setTxAmount(e.target.value)}
                  placeholder="0.00"
                  className="w-full min-h-[52px] px-4 rounded-xl border text-2xl font-extrabold font-mono-num outline-none"
                  style={{
                    background: 'var(--surface-soft)',
                    borderColor: 'var(--border)',
                    color: 'var(--text)',
                  }}
                />
              </div>

              <div className="flex gap-1.5 overflow-x-auto no-scrollbar pb-1">
                {[1, 5, 10, 50, 100, 200, 500, 1000].map((v) => (
                  <button
                    key={v}
                    type="button"
                    onClick={() =>
                      setTxAmount((prev) => ((parseFloat(prev) || 0) + v).toFixed(2))
                    }
                    className="px-3 py-1.5 rounded-lg border text-xs font-bold cursor-pointer shrink-0"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--border)',
                      color: 'var(--text)',
                    }}
                  >
                    +{v}
                  </button>
                ))}
                <button
                  type="button"
                  onClick={() => setTxAmount('')}
                  className="px-3 py-1.5 rounded-lg border text-xs font-bold cursor-pointer shrink-0"
                  style={{
                    background: 'var(--surface-soft)',
                    borderColor: 'var(--border)',
                    color: 'var(--text-muted)',
                  }}
                >
                  مسح
                </button>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                      {txType === 'قبض' ? 'إيداع إلى' : txType === 'صرف' ? 'طريقة الدفع' : 'من حساب'}
                    </label>
                    <button
                      type="button"
                      onClick={() => setAddAccountOpen(true)}
                      className="text-[11px] font-extrabold px-2 py-0.5 rounded-md border cursor-pointer"
                      style={{
                        background: 'var(--surface-soft)',
                        borderColor: 'var(--primary-accent)',
                        color: 'var(--primary-accent)',
                      }}
                    >
                      + حساب جديد
                    </button>
                  </div>
                  <select
                    value={fromAcc}
                    onChange={(e) => setFromAcc(e.target.value)}
                    className="w-full min-h-[46px] px-3 rounded-xl border text-sm font-bold outline-none"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--border)',
                      color: 'var(--text)',
                    }}
                  >
                    {accountOptions.map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                </div>

                {txType !== 'تحويل بين حسابات' ? (
                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <label className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                        التصنيف
                      </label>
                      <button
                        type="button"
                        onClick={() => setNewCatOpen((v) => !v)}
                        className="text-[11px] font-extrabold px-2 py-0.5 rounded-md border cursor-pointer"
                        style={{
                          background: 'var(--surface-soft)',
                          borderColor: 'var(--primary-accent)',
                          color: 'var(--primary-accent)',
                        }}
                      >
                        + تصنيف جديد
                      </button>
                    </div>
                    <select
                      value={txCat}
                      onChange={(e) => {
                        if (e.target.value === '__ADD_NEW__') {
                          setNewCatOpen(true);
                        } else {
                          setTxCat(e.target.value);
                        }
                      }}
                      className="w-full min-h-[46px] px-3 rounded-xl border text-sm font-bold outline-none"
                      style={{
                        background: 'var(--surface-soft)',
                        borderColor: 'var(--border)',
                        color: 'var(--text)',
                      }}
                    >
                      {activeCategoryList.map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                      <option value="__ADD_NEW__">+ إضافة تصنيف جديد...</option>
                    </select>
                  </div>
                ) : (
                  <div>
                    <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
                      إلى حساب
                    </label>
                    <select
                      value={toAcc}
                      onChange={(e) => setToAcc(e.target.value)}
                      className="w-full min-h-[46px] px-3 rounded-xl border text-sm font-bold outline-none"
                      style={{
                        background: 'var(--surface-soft)',
                        borderColor: 'var(--border)',
                        color: 'var(--text)',
                      }}
                    >
                      {accountOptions.map((m) => (
                        <option key={m} value={m}>
                          {m}
                        </option>
                      ))}
                    </select>
                  </div>
                )}
              </div>

              {newCatOpen && (
                <div
                  className="p-3 rounded-xl border border-dashed space-y-2"
                  style={{
                    background: 'var(--surface-soft)',
                    borderColor: 'var(--primary-accent)',
                  }}
                >
                  <div className="text-xs font-extrabold" style={{ color: 'var(--primary-accent)' }}>
                    إضافة تصنيف ({txType === 'قبض' ? 'قبض' : 'صرف'}) جديد إلى الشيت فوراً
                  </div>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      value={newCatName}
                      onChange={(e) => setNewCatName(e.target.value)}
                      placeholder="اسم التصنيف..."
                      className="flex-1 min-h-[42px] px-3 rounded-xl border text-sm outline-none"
                      style={{
                        background: 'var(--card-bg)',
                        borderColor: 'var(--border)',
                        color: 'var(--text)',
                      }}
                    />
                    <button
                      type="button"
                      disabled={newCatSaving}
                      onClick={handleSaveNewCategory}
                      className="px-4 rounded-xl text-xs font-extrabold cursor-pointer"
                      style={{
                        background: 'var(--accent-gradient)',
                        color: 'var(--accent-contrast)',
                      }}
                    >
                      {newCatSaving ? '...' : 'حفظ'}
                    </button>
                    <button
                      type="button"
                      onClick={() => setNewCatOpen(false)}
                      className="px-3 rounded-xl border text-xs font-bold cursor-pointer"
                      style={{
                        background: 'var(--card-bg)',
                        borderColor: 'var(--border)',
                        color: 'var(--text)',
                      }}
                    >
                      ✕
                    </button>
                  </div>
                </div>
              )}

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="sm:col-span-2">
                  <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
                    البيان / الوصف
                  </label>
                  <input
                    type="text"
                    value={txDesc}
                    onChange={(e) => setTxDesc(e.target.value)}
                    list="descSuggestions"
                    className="w-full min-h-[46px] px-3.5 rounded-xl border text-sm outline-none"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--border)',
                      color: 'var(--text)',
                    }}
                  />
                  <datalist id="descSuggestions">
                    {Array.from(new Set(state.recent.map((r) => r.desc).filter(Boolean)))
                      .slice(0, 60)
                      .map((d) => (
                        <option key={d} value={d} />
                      ))}
                  </datalist>
                </div>
                <div>
                  <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
                    التاريخ
                  </label>
                  <input
                    type="date"
                    required
                    value={txDate}
                    onChange={(e) => setTxDate(e.target.value)}
                    className="w-full min-h-[46px] px-3 rounded-xl border text-sm outline-none"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--border)',
                      color: 'var(--text)',
                    }}
                  />
                </div>
              </div>

              <button
                type="submit"
                disabled={txSaving}
                className="w-full min-h-[50px] rounded-xl text-sm font-extrabold cursor-pointer disabled:opacity-50"
                style={{
                  background: 'var(--accent-gradient)',
                  color: 'var(--accent-contrast)',
                  boxShadow: '0 10px 24px -8px var(--accent-shadow)',
                }}
              >
                {txSaving
                  ? 'جاري الحفظ...'
                  : editTx
                  ? 'حفظ التعديل في الشيت'
                  : 'تسجيل الحركة في الشيت'}
              </button>

              {editTx && (
                <button
                  type="button"
                  onClick={() => {
                    setEditTx(null);
                    setTxAmount('');
                    setTxDesc('');
                  }}
                  className="w-full py-2.5 rounded-xl border text-xs font-bold cursor-pointer"
                  style={{
                    background: 'var(--surface-soft)',
                    borderColor: 'var(--border)',
                    color: 'var(--text)',
                  }}
                >
                  إلغاء التعديل
                </button>
              )}
            </form>
          </div>
        )}

        {/* شاشة المستحقات */}
        {activeTab === 'debts' && (
          <div className="max-w-3xl mx-auto space-y-4">
            <div className="grid grid-cols-3 gap-2.5">
              <div
                className="rounded-2xl p-3.5 border"
                style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
              >
                <div className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                  مدين (ليا)
                </div>
                <div className="text-base font-extrabold mt-1 font-mono-num" dir="ltr" style={{ color: 'var(--income)' }}>
                  {fmt(debtSummary.lya)}
                </div>
              </div>
              <div
                className="rounded-2xl p-3.5 border"
                style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
              >
                <div className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                  داين (عليا)
                </div>
                <div className="text-base font-extrabold mt-1 font-mono-num" dir="ltr" style={{ color: 'var(--expense)' }}>
                  {fmt(debtSummary.alya)}
                </div>
              </div>
              <div
                className="rounded-2xl p-3.5 border"
                style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
              >
                <div className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                  الصافي
                </div>
                <div
                  className="text-base font-extrabold mt-1 font-mono-num"
                  dir="ltr"
                  style={{ color: debtSummary.net >= 0 ? 'var(--income)' : 'var(--expense)' }}
                >
                  {debtSummary.net >= 0 ? '+' : ''}
                  {fmt(debtSummary.net)}
                </div>
              </div>
            </div>

            <form
              onSubmit={handleDebtSubmit}
              className="rounded-2xl p-5 border space-y-3"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
            >
              <div
                className="flex p-1 rounded-xl border gap-1"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
              >
                {(['ليا', 'عليا'] as const).map((t) => (
                  <button
                    key={t}
                    type="button"
                    onClick={() => setDebtType(t)}
                    className="flex-1 py-2 rounded-lg text-sm font-bold cursor-pointer"
                    style={{
                      background: debtType === t ? 'var(--card-bg)' : 'transparent',
                      color:
                        debtType === t
                          ? t === 'ليا'
                            ? 'var(--income)'
                            : 'var(--expense)'
                          : 'var(--text-muted)',
                    }}
                  >
                    {t === 'ليا' ? 'مدين (ليا)' : 'داين (عليا)'}
                  </button>
                ))}
              </div>

              <input
                type="text"
                required
                value={debtDesc}
                onChange={(e) => setDebtDesc(e.target.value)}
                placeholder={debtType === 'ليا' ? 'اسم المدين' : 'اسم الدائن'}
                className="w-full min-h-[46px] px-3.5 rounded-xl border text-sm outline-none"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
              />

              <div className="grid grid-cols-2 gap-3">
                <input
                  type="number"
                  step="0.01"
                  required
                  value={debtAmount}
                  onChange={(e) => setDebtAmount(e.target.value)}
                  placeholder="المبلغ"
                  className="w-full min-h-[46px] px-3.5 rounded-xl border text-base font-extrabold font-mono-num outline-none"
                  style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                />
                <input
                  type="date"
                  required
                  value={debtDate}
                  onChange={(e) => setDebtDate(e.target.value)}
                  className="w-full min-h-[46px] px-3 rounded-xl border text-sm outline-none"
                  style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                />
              </div>

              <button
                type="submit"
                disabled={debtSaving}
                className="w-full min-h-[48px] rounded-xl text-sm font-extrabold cursor-pointer"
                style={{ background: 'var(--accent-gradient)', color: 'var(--accent-contrast)' }}
              >
                {debtSaving ? 'جاري الحفظ...' : 'حفظ المستحق في الشيت'}
              </button>
            </form>

            <div
              className="rounded-2xl p-4 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
            >
              <div className="flex gap-1.5 overflow-x-auto no-scrollbar mb-3">
                {(['all', 'مستحق', 'ليا', 'عليا', 'تم السداد'] as const).map((f) => (
                  <button
                    key={f}
                    type="button"
                    onClick={() => setDebtFilter(f)}
                    className="px-3 py-1.5 rounded-lg border text-xs font-bold cursor-pointer shrink-0"
                    style={{
                      background: debtFilter === f ? 'var(--accent-gradient)' : 'var(--surface-soft)',
                      borderColor: debtFilter === f ? 'transparent' : 'var(--border)',
                      color: debtFilter === f ? 'var(--accent-contrast)' : 'var(--text-muted)',
                    }}
                  >
                    {f === 'all' ? 'الكل' : f === 'مستحق' ? 'معلّق' : f === 'ليا' ? 'مدين (ليا)' : f === 'عليا' ? 'داين (عليا)' : 'تم السداد'}
                  </button>
                ))}
              </div>

              {filteredDebts.length === 0 ? (
                <div className="text-center py-6 text-xs" style={{ color: 'var(--text-muted)' }}>
                  لا توجد مستحقات مطابقة
                </div>
              ) : (
                <div className="divide-y" style={{ borderColor: 'var(--border)' }}>
                  {filteredDebts.map((d, idx) => {
                    const isLya = isDebtLya(d.type);
                    const paid = d.status === 'تم السداد';
                    return (
                      <div
                        key={d.id || idx}
                        className="py-3 flex items-center justify-between gap-2"
                        style={{ opacity: paid ? 0.6 : 1 }}
                      >
                        <div>
                          <div className="text-sm font-bold">{d.desc}</div>
                          <div className="text-xs mt-0.5" style={{ color: 'var(--text-muted)' }}>
                            <span style={{ color: isLya ? 'var(--income)' : 'var(--expense)' }} className="font-bold">
                              {isLya ? 'مدين (ليا)' : 'داين (عليا)'}
                            </span>{' '}
                            ·{' '}
                            <span style={{ color: paid ? 'var(--info)' : 'var(--warning)' }} className="font-bold">
                              {paid ? 'تم السداد' : 'مستحق'}
                            </span>{' '}
                            · {d.date}
                          </div>
                        </div>
                        <div className="flex items-center gap-2 shrink-0">
                          <span
                            className="text-sm font-extrabold font-mono-num"
                            dir="ltr"
                            style={{ color: isLya ? 'var(--income)' : 'var(--expense)' }}
                          >
                            {fmt(d.amount)}
                          </span>
                          <button
                            type="button"
                            onClick={() => (paid ? toggleDebtStatus(d) : handleStartDebtSettle(d))}
                            className="px-2.5 py-1.5 rounded-lg border text-xs font-bold cursor-pointer"
                            style={{
                              background: 'var(--surface-soft)',
                              borderColor: paid ? 'var(--border)' : 'var(--primary-accent)',
                              color: paid ? 'var(--text-muted)' : 'var(--income)',
                            }}
                          >
                            {paid ? 'إعادة فتح' : 'سداد ←'}
                          </button>
                          <button
                            type="button"
                            onClick={() => deleteDebt(d)}
                            className="px-2.5 py-1.5 rounded-lg border text-xs font-bold cursor-pointer"
                            style={{
                              background: 'var(--surface-soft)',
                              borderColor: 'var(--border)',
                              color: 'var(--expense)',
                            }}
                          >
                            ✕
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        )}

        {/* شاشة التنبيهات والالتزامات */}
        {activeTab === 'reminders' && (
          <div className="max-w-3xl mx-auto space-y-4">
            <div className="grid grid-cols-2 gap-3">
              <div
                className="rounded-2xl p-4 border"
                style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
              >
                <div className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                  التزامات قادمة
                </div>
                <div className="text-lg font-extrabold mt-1" style={{ color: 'var(--primary-accent)' }}>
                  {reminderSummary.count} التزام
                </div>
              </div>
              <div
                className="rounded-2xl p-4 border"
                style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
              >
                <div className="text-xs font-bold" style={{ color: 'var(--text-muted)' }}>
                  إجمالي المطلوب سداده
                </div>
                <div className="text-lg font-extrabold mt-1 font-mono-num" dir="ltr" style={{ color: 'var(--expense)' }}>
                  {fmt(reminderSummary.sum)} ج.م
                </div>
              </div>
            </div>

            <form
              onSubmit={handleReminderSubmit}
              className="rounded-2xl p-5 border space-y-3"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
            >
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <input
                  type="text"
                  required
                  value={remDesc}
                  onChange={(e) => setRemDesc(e.target.value)}
                  placeholder="البيان / الالتزام (مثال: فاتورة إنترنت)"
                  className="sm:col-span-2 min-h-[46px] px-3.5 rounded-xl border text-sm outline-none"
                  style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                />
                <input
                  type="number"
                  step="0.01"
                  required
                  value={remAmount}
                  onChange={(e) => setRemAmount(e.target.value)}
                  placeholder="المبلغ"
                  className="min-h-[46px] px-3.5 rounded-xl border text-base font-extrabold font-mono-num outline-none"
                  style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <input
                  type="date"
                  required
                  value={remDueDate}
                  onChange={(e) => setRemDueDate(e.target.value)}
                  className="min-h-[46px] px-3 rounded-xl border text-sm outline-none"
                  style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                />
                <select
                  value={remRepeat}
                  onChange={(e) => setRemRepeat(e.target.value)}
                  className="min-h-[46px] px-3 rounded-xl border text-sm font-bold outline-none"
                  style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                >
                  <option value="شهري">شهري</option>
                  <option value="لمرة واحدة">لمرة واحدة</option>
                  <option value="أسبوعي">أسبوعي</option>
                  <option value="سنوي">سنوي</option>
                </select>
              </div>

              <input
                type="text"
                value={remNotes}
                onChange={(e) => setRemNotes(e.target.value)}
                placeholder="ملاحظات إضافية (اختياري)"
                className="w-full min-h-[44px] px-3.5 rounded-xl border text-sm outline-none"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
              />

              <button
                type="submit"
                disabled={remSaving}
                className="w-full min-h-[48px] rounded-xl text-sm font-extrabold cursor-pointer"
                style={{ background: 'var(--accent-gradient)', color: 'var(--accent-contrast)' }}
              >
                {remSaving ? 'جاري الحفظ...' : 'حفظ التنبيه في الشيت'}
              </button>
            </form>

            <div
              className="rounded-2xl p-4 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
            >
              {state.reminders.length === 0 ? (
                <div className="text-center py-6 text-xs" style={{ color: 'var(--text-muted)' }}>
                  لا توجد تنبيهات مسجلة بعد
                </div>
              ) : (
                <div className="divide-y" style={{ borderColor: 'var(--border)' }}>
                  {[...state.reminders]
                    .sort((a, b) => (a.dueDate || '').localeCompare(b.dueDate || ''))
                    .map((r, idx) => {
                      const paid = r.status === 'تم الدفع';
                      return (
                        <div
                          key={r.id || idx}
                          className="py-3 flex items-center justify-between gap-2"
                          style={{ opacity: paid ? 0.6 : 1 }}
                        >
                          <div>
                            <div className="text-sm font-bold">{r.desc}</div>
                            <div className="text-xs mt-0.5" style={{ color: 'var(--text-muted)' }}>
                              <span style={{ color: paid ? 'var(--info)' : 'var(--warning)' }} className="font-bold">
                                {paid ? 'تم الدفع' : 'قادم'}
                              </span>{' '}
                              · استحقاق {r.dueDate} · {r.repeat}
                            </div>
                          </div>
                          <div className="flex items-center gap-2 shrink-0">
                            <span className="text-sm font-extrabold font-mono-num" dir="ltr" style={{ color: 'var(--expense)' }}>
                              {fmt(r.amount)}
                            </span>
                            {!paid && (
                              <button
                                type="button"
                                onClick={() => {
                                  setEditTx(null);
                                  setPendingDebtSettle(null);
                                  setActiveTab('add');
                                  setTxType('صرف');
                                  setTxAmount(cleanNumber(r.amount).toFixed(2));
                                  setTxDesc(r.desc);
                                  setTxDate(localDateKey());
                                  setPendingReminderPay({
                                    id: r.id,
                                    row: r.row,
                                    desc: r.desc,
                                    amount: cleanNumber(r.amount),
                                  });
                                  window.scrollTo({ top: 0, behavior: 'smooth' });
                                  notify('اختر الحساب والتصنيف واضغط تسجيل الحركة');
                                }}
                                className="px-2.5 py-1.5 rounded-lg border text-xs font-bold cursor-pointer"
                                style={{
                                  background: 'var(--surface-soft)',
                                  borderColor: 'var(--primary-accent)',
                                  color: 'var(--income)',
                                }}
                              >
                                سداد ←
                              </button>
                            )}
                            <button
                              type="button"
                              onClick={() => deleteReminder(r)}
                              className="px-2.5 py-1.5 rounded-lg border text-xs font-bold cursor-pointer"
                              style={{
                                background: 'var(--surface-soft)',
                                borderColor: 'var(--border)',
                                color: 'var(--expense)',
                              }}
                            >
                              ✕
                            </button>
                          </div>
                        </div>
                      );
                    })}
                </div>
              )}
            </div>
          </div>
        )}

        {/* شاشة السجل */}
        {activeTab === 'history' && (
          <div
            className="max-w-3xl mx-auto rounded-2xl p-4 border space-y-3"
            style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
          >
            <input
              type="text"
              value={historySearch}
              onChange={(e) => setHistorySearch(e.target.value)}
              placeholder="ابحث بالبيان أو التصنيف أو الحساب..."
              className="w-full min-h-[46px] px-3.5 rounded-xl border text-sm outline-none"
              style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
            />

            <div className="flex gap-1.5 overflow-x-auto no-scrollbar">
              {(['all', 'صرف', 'قبض', 'تحويل'] as const).map((f) => (
                <button
                  key={f}
                  type="button"
                  onClick={() => setHistoryFilter(f)}
                  className="px-3.5 py-1.5 rounded-lg border text-xs font-bold cursor-pointer shrink-0"
                  style={{
                    background: historyFilter === f ? 'var(--accent-gradient)' : 'var(--surface-soft)',
                    borderColor: historyFilter === f ? 'transparent' : 'var(--border)',
                    color: historyFilter === f ? 'var(--accent-contrast)' : 'var(--text-muted)',
                  }}
                >
                  {f === 'all' ? 'الكل' : f === 'تحويل' ? 'تحويلات' : f}
                </button>
              ))}
            </div>

            {filteredHistory.length === 0 ? (
              <div className="text-center py-8 text-xs" style={{ color: 'var(--text-muted)' }}>
                لا توجد معاملات مطابقة
              </div>
            ) : (
              <div className="divide-y" style={{ borderColor: 'var(--border)' }}>
                {filteredHistory.map((t, idx) => (
                  <TxRow
                    key={`${t.row || idx}-${t.date}`}
                    tx={t}
                    onClick={() => t.row && setSelectedTx(t)}
                  />
                ))}
              </div>
            )}
          </div>
        )}

        {/* شاشة التحليلات */}
        {activeTab === 'stats' && (
          <AnalyticsSection
            transactions={state.recent}
            budgets={budgets}
            storageKey={sheetStorageKey}
            onSelectBudgetCat={(cat, curBud) => {
              setBudgetModalCat(cat);
              setBudgetModalVal(curBud ? String(curBud) : '');
            }}
          />
        )}

        {/* شاشة استوديو Google Sheets وتنزيل HTML / Apps Script */}
        {activeTab === 'sheets' && (
          <div className="max-w-3xl mx-auto">
            <SheetStudioPanel
              googleUserEmail={googleEmail}
              hasGoogleToken={Boolean(googleToken)}
              isLoggingIn={isLoggingIn}
              onGoogleLogin={handleGoogleLogin}
              onGoogleLogout={handleGoogleLogout}
              profileName={profileName}
              onProfileNameChange={setProfileName}
              spreadsheetIdInput={spreadsheetId}
              onSpreadsheetIdChange={setSpreadsheetId}
              scriptUrlInput={scriptUrl}
              onScriptUrlChange={setScriptUrl}
              inspectionReport={inspectionReport}
              isInspecting={isInspecting}
              onInspectSheet={handleInspectSheet}
              onFormatWorkbook={handleFormatWorkbook}
              existingSheets={existingSheets}
              isSearchingDrive={isSearchingDrive}
              onSearchExistingSheets={handleSearchExistingSheets}
              onSelectExistingSheet={handleSelectExistingSheet}
              onCreateNewSheet={handleCreateNewSheet}
              onTestScriptConnection={handleTestScriptConnection}
              isTestingScript={isTestingScript}
              onNotify={notify}
            />
          </div>
        )}

        {/* شاشة الإعدادات والثيمات */}
        {activeTab === 'settings' && (
          <div className="max-w-3xl mx-auto space-y-4">
            {/* 1. الثيمات المصغرة المتكاملة (خط + شكل + أيقونات + ألوان) */}
            <div
              className="rounded-2xl p-4 sm:p-5 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
            >
              <div className="flex items-center justify-between gap-2 mb-3.5">
                <div className="flex items-center gap-2.5">
                  <div className="theme-icon-box">
                    <Palette className="w-4 h-4" />
                  </div>
                  <div>
                    <h2 className="text-sm font-extrabold">ثيمات التطبيق</h2>
                    <p className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                      تُغيّر الألوان والخطوط وشكل الأيقونات والبطاقات بالكامل
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    setThemePref('auto');
                    notify('تم تفعيل الثيم التلقائي');
                  }}
                  className="px-3 py-1.5 rounded-xl border text-[11px] font-bold cursor-pointer shrink-0"
                  style={{
                    background: themePref === 'auto' ? 'var(--accent-gradient)' : 'var(--surface-soft)',
                    borderColor: themePref === 'auto' ? 'transparent' : 'var(--border)',
                    color: themePref === 'auto' ? 'var(--accent-contrast)' : 'var(--text)',
                  }}
                >
                  تلقائي
                </button>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2">
                {THEMES.map((t) => {
                  const active = resolvedThemeId === t.id && themePref !== 'auto';
                  return (
                    <button
                      key={t.id}
                      type="button"
                      onClick={() => {
                        setThemePref(t.id);
                        notify(`تم تفعيل ثيم ${t.name}`);
                      }}
                      className="p-2 rounded-xl border-2 cursor-pointer text-right transition-all flex flex-col gap-1.5"
                      style={{
                        background: active ? 'var(--card-bg)' : 'var(--surface-soft)',
                        borderColor: active ? 'var(--primary-accent)' : 'var(--border)',
                        fontFamily: t.vars['--font-main'],
                      }}
                    >
                      <div
                        className="h-6 w-full rounded-md overflow-hidden flex border border-black/10"
                        style={{ borderRadius: t.vars['--radius-btn'] }}
                      >
                        <span className="flex-1 h-full" style={{ background: t.previewSwatches[0] }} />
                        <span className="flex-1 h-full" style={{ background: t.previewSwatches[1] }} />
                        <span className="flex-1 h-full" style={{ background: t.previewSwatches[2] }} />
                      </div>
                      <div className="flex items-center justify-between gap-1">
                        <span className="text-[11px] font-extrabold truncate">{t.name}</span>
                        <span className="text-[9px] font-bold shrink-0" style={{ color: 'var(--text-muted)' }}>
                          {t.badge}
                        </span>
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* 2. الهوية والربط بشيت جوجل */}
            <div
              className="rounded-2xl p-4 sm:p-5 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
            >
              <div className="text-xs font-extrabold mb-2" style={{ color: 'var(--text-muted)' }}>
                الهوية والربط بشيت جوجل
              </div>
              <div className="divide-y" style={{ borderColor: 'var(--border)' }}>
                <div className="py-3 flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="theme-icon-box">
                      <User className="w-4 h-4" />
                    </div>
                    <div className="min-w-0">
                      <div className="text-sm font-bold">اسمك / اسم المحفظة</div>
                      <div className="text-[11px] truncate font-mono-num" style={{ color: 'var(--text-muted)' }}>
                        {buildM7faztySheetName(profileName)}
                      </div>
                    </div>
                  </div>
                  <input
                    type="text"
                    value={profileName}
                    onChange={(e) => setProfileName(e.target.value)}
                    placeholder="أدخل اسمك..."
                    className="h-9 px-3 rounded-xl border text-xs font-bold outline-none w-44"
                    style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                  />
                </div>

                <div className="py-3 flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="theme-icon-box">
                      <FileSpreadsheet className="w-4 h-4" />
                    </div>
                    <div className="min-w-0">
                      <div className="text-sm font-bold">استوديو Google Sheets والنسخة المستقلة</div>
                      <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                        فحص الشيت، إصلاح المعادلات، وتنزيل HTML و Code.gs
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => navigateTo('sheets')}
                      className="h-9 px-3 rounded-xl border text-xs font-bold cursor-pointer"
                      style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
                    >
                      فتح الاستوديو
                    </button>
                    <a
                      href={`https://docs.google.com/spreadsheets/d/${extractSpreadsheetId(spreadsheetId)}/edit`}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="h-9 px-3 rounded-xl border text-xs font-bold inline-flex items-center gap-1"
                      style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                    >
                      <span>الشيت</span>
                      <ExternalLink className="w-3.5 h-3.5" />
                    </a>
                  </div>
                </div>
              </div>
            </div>

            {/* 3. الأمان والخصوصية وإدارة أنواع الحسابات */}
            <div
              className="rounded-2xl p-4 sm:p-5 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
            >
              <div className="text-xs font-extrabold mb-2" style={{ color: 'var(--text-muted)' }}>
                الأمان وإدارة الحسابات
              </div>
              <div className="divide-y" style={{ borderColor: 'var(--border)' }}>
                <div className="py-3 flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="theme-icon-box">
                      <ShieldCheck className="w-4 h-4" />
                    </div>
                    <div className="min-w-0">
                      <div className="text-sm font-bold">تعديل نوع الحسابات بأمان</div>
                      <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                        تصحيح نوع حساب (بنك / خزنة / محفظة / بطاقة) في الشيت بـ 3 خطوات مؤمّنة
                      </div>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => setAccountKindModalOpen(true)}
                    className="h-9 px-3.5 rounded-xl border text-xs font-bold cursor-pointer"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--primary-accent)',
                      color: 'var(--primary-accent)',
                    }}
                  >
                    إدارة الأنواع
                  </button>
                </div>

                <div className="py-3 flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="theme-icon-box">
                      <KeyRound className="w-4 h-4" />
                    </div>
                    <div className="min-w-0">
                      <div className="text-sm font-bold">رمز PIN الحماية (اختياري)</div>
                      <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                        {storedPin ? 'مُفعّل — يمكنك تغييره أو تعطيله في أي وقت' : 'غير مُفعّل — يمكنك تعيين رمز من 4 أرقام'}
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => setPinOverlay({ open: true, mode: storedPin ? 'change' : 'setup' })}
                      className="h-9 px-3 rounded-xl border text-xs font-bold cursor-pointer"
                      style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
                    >
                      {storedPin ? 'تغيير' : 'تعيين PIN'}
                    </button>
                    {storedPin && (
                      <>
                        <button
                          type="button"
                          onClick={() => setPinOverlay({ open: true, mode: 'unlock' })}
                          className="h-9 px-3 rounded-xl border text-xs font-bold cursor-pointer"
                          style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
                        >
                          قفل
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            localStorage.removeItem('app_user_pin');
                            setStoredPin(null);
                            notify('تم إلغاء رمز PIN بنجاح');
                          }}
                          className="h-9 px-3 rounded-xl border text-xs font-bold cursor-pointer"
                          style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--expense)' }}
                        >
                          إلغاء
                        </button>
                      </>
                    )}
                  </div>
                </div>

                <div className="py-3 flex items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="theme-icon-box">
                      <Clock className="w-4 h-4" />
                    </div>
                    <div className="min-w-0">
                      <div className="text-sm font-bold">القفل التلقائي</div>
                      <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                        يُقفل التطبيق عند ترك التبويب لفترة
                      </div>
                    </div>
                  </div>
                  <select
                    value={autoLockMins}
                    onChange={(e) => {
                      setAutoLockMins(e.target.value);
                      localStorage.setItem('mf_autolock', e.target.value);
                    }}
                    className="h-9 px-3 rounded-xl border text-xs font-bold outline-none"
                    style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
                  >
                    <option value="0">بدون</option>
                    <option value="1">بعد دقيقة</option>
                    <option value="5">بعد 5 دقائق</option>
                  </select>
                </div>
              </div>
            </div>

            {/* 4. البيانات والنسخ الاحتياطي والمزامنة */}
            <div
              className="rounded-2xl p-4 sm:p-5 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
            >
              <div className="text-xs font-extrabold mb-2" style={{ color: 'var(--text-muted)' }}>
                البيانات والنسخ الاحتياطي
              </div>
              <div className="divide-y" style={{ borderColor: 'var(--border)' }}>
                <div className="py-3 flex items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="theme-icon-box">
                      <Download className="w-4 h-4" />
                    </div>
                    <div className="min-w-0">
                      <div className="text-sm font-bold">تصدير السجل والأرصدة</div>
                      <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                        ملف Excel كامل أو CSV
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={handleExportExcel}
                      className="h-9 px-3 rounded-xl border text-xs font-bold cursor-pointer"
                      style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
                    >
                      Excel
                    </button>
                    <button
                      type="button"
                      onClick={handleExportCsv}
                      className="h-9 px-3 rounded-xl border text-xs font-bold cursor-pointer"
                      style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
                    >
                      CSV
                    </button>
                  </div>
                </div>

                <div className="py-3 flex items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="theme-icon-box">
                      <Database className="w-4 h-4" />
                    </div>
                    <div className="min-w-0">
                      <div className="text-sm font-bold">نسخة احتياطية كاملة (JSON)</div>
                      <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                        حفظ أو استعادة البيانات المحلية
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={handleExportBackup}
                      className="h-9 px-3 rounded-xl border text-xs font-bold cursor-pointer"
                      style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
                    >
                      حفظ
                    </button>
                    <label
                      className="h-9 px-3 rounded-xl border text-xs font-bold cursor-pointer inline-flex items-center"
                      style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
                    >
                      استعادة
                      <input type="file" accept="application/json,.json" hidden onChange={handleImportBackup} />
                    </label>
                  </div>
                </div>

                <div className="py-3 flex items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="theme-icon-box">
                      <Send className="w-4 h-4" />
                    </div>
                    <div className="min-w-0">
                      <div className="text-sm font-bold">حركات بانتظار الإرسال</div>
                      <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                        {queue.length} طلب بانتظار التأكيد
                      </div>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={flushQueue}
                    className="h-9 px-3.5 rounded-xl border text-xs font-bold cursor-pointer"
                    style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
                  >
                    إرسال الآن
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* زر الإضافة السريعة العائم */}
      {activeTab !== 'add' && (
        <button
          type="button"
          onClick={() => navigateTo('add')}
          title="إضافة حركة جديدة"
          aria-label="إضافة حركة جديدة"
          className="fixed bottom-5 left-5 w-14 h-14 rounded-2xl flex items-center justify-center cursor-pointer z-30 transition-transform hover:-translate-y-0.5 active:scale-95"
          style={{
            background: 'var(--accent-gradient)',
            color: 'var(--accent-contrast)',
            boxShadow: '0 12px 28px -6px var(--accent-shadow)',
          }}
        >
          <Plus className="w-6 h-6" />
        </button>
      )}

      {/* مودال إضافة محفظة / بنك / خزنة */}
      <AddAccountModal
        open={addAccountOpen}
        onClose={() => setAddAccountOpen(false)}
        onSubmit={handleAddAccount}
      />

      {/* مودال تعديل نوع الحساب المؤمّن بـ 3 خطوات */}
      <AccountKindManagerModal
        open={accountKindModalOpen}
        accounts={state.accounts}
        storedPin={storedPin}
        sheetUrl={`https://docs.google.com/spreadsheets/d/${extractSpreadsheetId(spreadsheetId)}/edit`}
        onClose={() => setAccountKindModalOpen(false)}
        onConfirmChange={handleUpdateAccountKind}
      />

      {/* مودال تفاصيل وحذف وتعديل معاملة */}
      {selectedTx && (
        <div
          className="fixed inset-0 z-50 flex items-end md:items-center justify-center p-3 bg-black/50 backdrop-blur-sm"
          onClick={(e) => {
            if (e.target === e.currentTarget) setSelectedTx(null);
          }}
        >
          <div
            className="w-full max-w-md rounded-2xl p-5 border space-y-3"
            style={{ background: 'var(--card-bg)', borderColor: 'var(--border)', boxShadow: 'var(--shadow-md)' }}
          >
            <div className="text-base font-extrabold">{selectedTx.desc || selectedTx.category}</div>
            <div className="text-xs leading-relaxed" style={{ color: 'var(--text-muted)' }}>
              {selectedTx.type} · {selectedTx.category} · {selectedTx.fromAcc}
              {selectedTx.toAcc ? ` ← ${selectedTx.toAcc}` : ''}
              <br />
              {selectedTx.date} · <span className="font-mono-num font-bold">{fmt(selectedTx.amount)} ج.م</span>
            </div>
            <div className="grid grid-cols-2 gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => {
                  const t = selectedTx;
                  setSelectedTx(null);
                  setEditTx(t);
                  setTxType((t.type as TransactionType) || 'صرف');
                  setFromAcc(t.fromAcc);
                  if (t.toAcc) setToAcc(t.toAcc);
                  if (t.type !== 'تحويل بين حسابات') setTxCat(t.category);
                  setTxDesc(t.desc);
                  setTxAmount(cleanNumber(t.amount).toFixed(2));
                  if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) setTxDate(t.date);
                  setActiveTab('add');
                }}
                className="py-2.5 rounded-xl border text-xs font-bold cursor-pointer"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
              >
                تعديل
              </button>
              <button
                type="button"
                onClick={handleDeleteTx}
                className="py-2.5 rounded-xl border text-xs font-bold cursor-pointer"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--expense)' }}
              >
                حذف
              </button>
            </div>
            <button
              type="button"
              onClick={() => setSelectedTx(null)}
              className="w-full py-2 text-xs font-bold cursor-pointer"
              style={{ color: 'var(--text-muted)' }}
            >
              إغلاق
            </button>
          </div>
        </div>
      )}

      {/* مودال تحديد ميزانية تصنيف */}
      {budgetModalCat && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm"
          onClick={(e) => {
            if (e.target === e.currentTarget) setBudgetModalCat(null);
          }}
        >
          <div
            className="w-full max-w-sm rounded-2xl p-5 border space-y-3"
            style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
          >
            <h3 className="text-sm font-extrabold">الميزانية الشهرية لتصنيف «{budgetModalCat}»</h3>
            <input
              type="number"
              step="0.01"
              value={budgetModalVal}
              onChange={(e) => setBudgetModalVal(e.target.value)}
              placeholder="اتركه فارغاً لحذف الميزانية"
              className="w-full min-h-[44px] px-3 rounded-xl border text-sm font-bold font-mono-num outline-none"
              style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)', color: 'var(--text)' }}
            />
            <div className="flex gap-2 justify-end">
              <button
                type="button"
                onClick={() => setBudgetModalCat(null)}
                className="px-4 py-2 rounded-xl border text-xs font-bold cursor-pointer"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
              >
                إلغاء
              </button>
              <button
                type="button"
                onClick={() => {
                  const val = cleanNumber(budgetModalVal);
                  const next = { ...budgets };
                  if (val > 0) next[budgetModalCat] = val;
                  else delete next[budgetModalCat];
                  setBudgets(next);
                  localStorage.setItem(`mf_budgets::${sheetStorageKey}`, JSON.stringify(next));
                  setBudgetModalCat(null);
                  notify('تم تحديث الميزانية');
                }}
                className="px-4 py-2 rounded-xl text-xs font-extrabold cursor-pointer"
                style={{ background: 'var(--accent-gradient)', color: 'var(--accent-contrast)' }}
              >
                حفظ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toastMsg && (
        <div
          className="fixed bottom-20 left-1/2 -translate-x-1/2 px-5 py-2.5 rounded-full text-xs font-bold z-[10000] shadow-lg max-w-[90vw] text-center"
          style={{ background: 'var(--text)', color: 'var(--bg)' }}
        >
          {toastMsg}
        </div>
      )}
    </div>
  );
}

function TxRow({ tx, onClick }: { tx: TransactionItem; onClick?: () => void }) {
  const isIncome = tx.type === 'قبض' || tx.type === 'دخل';
  const isTransfer = tx.type.includes('تحويل');
  const color = isIncome ? 'var(--income)' : isTransfer ? 'var(--transfer)' : 'var(--expense)';
  const sign = isIncome ? '+' : isTransfer ? '' : '-';
  const accText = isTransfer ? `${tx.fromAcc} ← ${tx.toAcc}` : tx.fromAcc;

  return (
    <div
      onClick={onClick}
      className={`py-3 flex items-center justify-between gap-2 ${onClick ? 'cursor-pointer' : ''}`}
    >
      <div className="min-w-0">
        <div className="text-sm font-bold truncate">{tx.desc || tx.category || tx.type}</div>
        <div className="text-xs mt-0.5 truncate" style={{ color: 'var(--text-muted)' }}>
          {tx.category} · {accText} · {tx.date}
        </div>
      </div>
      <span className="text-sm font-extrabold font-mono-num shrink-0" dir="ltr" style={{ color }}>
        {sign}
        {fmt(cleanNumber(tx.amount))}
      </span>
    </div>
  );
}
