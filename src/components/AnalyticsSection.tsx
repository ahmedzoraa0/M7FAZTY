import React, { useEffect, useMemo, useState } from 'react';
import {
  ArrowDownRight,
  ArrowUpRight,
  BarChart3,
  Calendar,
  ChevronLeft,
  ChevronRight,
  Eye,
  EyeOff,
  PieChart,
  TrendingUp,
} from 'lucide-react';
import { TransactionItem } from '../types';
import { cleanNumber } from '../services/appsScriptClient';

const ARABIC_MONTHS = [
  'يناير',
  'فبراير',
  'مارس',
  'أبريل',
  'مايو',
  'يونيو',
  'يوليو',
  'أغسطس',
  'سبتمبر',
  'أكتوبر',
  'نوفمبر',
  'ديسمبر',
];

const fmt = (n: number) =>
  Number(n || 0).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const fmtCompact = (n: number) =>
  Number(n || 0).toLocaleString('en-US', {
    maximumFractionDigits: 0,
  });

function parseTxYearMonth(dateStr: string): { year: number; month: number } | null {
  const s = String(dateStr || '').trim();
  let m = s.match(/^(\d{4})[-\/](\d{1,2})/);
  if (m) return { year: parseInt(m[1], 10), month: parseInt(m[2], 10) };
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (m) return { year: parseInt(m[3], 10), month: parseInt(m[2], 10) };
  const dt = new Date(s);
  if (Number.isNaN(dt.getTime())) return null;
  return { year: dt.getFullYear(), month: dt.getMonth() + 1 };
}

interface AnalyticsPanelsState {
  summary: boolean;
  yearlyChart: boolean;
  expenseCats: boolean;
  incomeCats: boolean;
}

const DEFAULT_PANELS: AnalyticsPanelsState = {
  summary: true,
  yearlyChart: true,
  expenseCats: true,
  incomeCats: false,
};

interface AnalyticsSectionProps {
  transactions: TransactionItem[];
  budgets: Record<string, number>;
  storageKey: string;
  onSelectBudgetCat: (cat: string, currentBudget: number) => void;
}

export const AnalyticsSection: React.FC<AnalyticsSectionProps> = ({
  transactions,
  budgets,
  storageKey,
  onSelectBudgetCat,
}) => {
  const now = useMemo(() => new Date(), []);
  const [selectedYear, setSelectedYear] = useState<number>(now.getFullYear());
  // 0 = full year, 1..12 = specific month
  const [selectedMonth, setSelectedMonth] = useState<number>(now.getMonth() + 1);

  const [panels, setPanels] = useState<AnalyticsPanelsState>(() => {
    try {
      const saved = localStorage.getItem(`mf_analytics_panels::${storageKey}`);
      if (saved) return { ...DEFAULT_PANELS, ...JSON.parse(saved) };
    } catch {
      // ignore
    }
    return DEFAULT_PANELS;
  });

  useEffect(() => {
    try {
      localStorage.setItem(`mf_analytics_panels::${storageKey}`, JSON.stringify(panels));
    } catch {
      // ignore
    }
  }, [panels, storageKey]);

  const togglePanel = (key: keyof AnalyticsPanelsState) => {
    setPanels((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  // All available years from transactions + current year
  const availableYears = useMemo(() => {
    const set = new Set<number>([now.getFullYear()]);
    transactions.forEach((t) => {
      const ym = parseTxYearMonth(t.date);
      if (ym && ym.year >= 2000 && ym.year <= 2100) set.add(ym.year);
    });
    return Array.from(set).sort((a, b) => b - a);
  }, [transactions, now]);

  // 12-month series for selectedYear
  const monthlySeries = useMemo(() => {
    const months = Array.from({ length: 12 }, (_, idx) => ({
      month: idx + 1,
      name: ARABIC_MONTHS[idx],
      expense: 0,
      income: 0,
      net: 0,
      count: 0,
    }));

    transactions.forEach((t) => {
      const ym = parseTxYearMonth(t.date);
      if (!ym || ym.year !== selectedYear) return;
      const amt = cleanNumber(t.amount);
      if (!(amt > 0)) return;
      const bucket = months[ym.month - 1];
      if (!bucket) return;
      if (t.type === 'صرف') {
        bucket.expense += amt;
        bucket.count++;
      } else if (t.type === 'قبض' || t.type === 'دخل') {
        bucket.income += amt;
        bucket.count++;
      }
    });

    months.forEach((m) => {
      m.net = m.income - m.expense;
    });

    return months;
  }, [transactions, selectedYear]);

  // Helper to get expense/income for ANY (year, month) for cross-year prev/next comparison
  const getMonthTotals = useMemo(() => {
    const map: Record<string, { expense: number; income: number }> = {};
    transactions.forEach((t) => {
      const ym = parseTxYearMonth(t.date);
      if (!ym) return;
      const key = `${ym.year}-${ym.month}`;
      if (!map[key]) map[key] = { expense: 0, income: 0 };
      const amt = cleanNumber(t.amount);
      if (t.type === 'صرف') map[key].expense += amt;
      else if (t.type === 'قبض' || t.type === 'دخل') map[key].income += amt;
    });
    return (y: number, m: number) => map[`${y}-${m}`] || { expense: 0, income: 0 };
  }, [transactions]);

  // Filtered transactions for selectedYear & selectedMonth (or whole year if 0)
  const scopedTxs = useMemo(() => {
    return transactions.filter((t) => {
      const ym = parseTxYearMonth(t.date);
      if (!ym || ym.year !== selectedYear) return false;
      if (selectedMonth > 0 && ym.month !== selectedMonth) return false;
      return true;
    });
  }, [transactions, selectedYear, selectedMonth]);

  // Period metrics
  const periodStats = useMemo(() => {
    let expense = 0;
    let income = 0;
    const expMap: Record<string, number> = {};
    const incMap: Record<string, number> = {};

    scopedTxs.forEach((t) => {
      const amt = cleanNumber(t.amount);
      if (!(amt > 0)) return;
      const cat = t.category || 'أخرى';
      if (t.type === 'صرف') {
        expense += amt;
        expMap[cat] = (expMap[cat] || 0) + amt;
      } else if (t.type === 'قبض' || t.type === 'دخل') {
        income += amt;
        incMap[cat] = (incMap[cat] || 0) + amt;
      }
    });

    if (selectedMonth > 0) {
      Object.keys(budgets).forEach((c) => {
        if (!(c in expMap)) expMap[c] = 0;
      });
    }

    const expenseList = Object.entries(expMap).sort((a, b) => b[1] - a[1]);
    const incomeList = Object.entries(incMap).sort((a, b) => b[1] - a[1]);

    // Days divisor for daily average
    let days = 30;
    if (selectedMonth === 0) {
      days = selectedYear === now.getFullYear()
        ? Math.max(1, Math.ceil((now.getTime() - new Date(selectedYear, 0, 1).getTime()) / 86400000))
        : 365;
    } else if (selectedYear === now.getFullYear() && selectedMonth === now.getMonth() + 1) {
      days = Math.max(1, now.getDate());
    } else {
      days = new Date(selectedYear, selectedMonth, 0).getDate();
    }

    return {
      expense,
      income,
      net: income - expense,
      dailyAvg: expense / days,
      expenseList,
      incomeList,
    };
  }, [scopedTxs, budgets, selectedMonth, selectedYear, now]);

  // Previous & Next Month Comparison (when a specific month is selected)
  const adjacentComparison = useMemo(() => {
    if (selectedMonth === 0) {
      const prevYearTxs = transactions.filter((t) => parseTxYearMonth(t.date)?.year === selectedYear - 1);
      let prevExp = 0;
      let prevInc = 0;
      prevYearTxs.forEach((t) => {
        const a = cleanNumber(t.amount);
        if (t.type === 'صرف') prevExp += a;
        else if (t.type === 'قبض' || t.type === 'دخل') prevInc += a;
      });
      return {
        prevLabel: `سنة ${selectedYear - 1}`,
        prev: { expense: prevExp, income: prevInc },
        nextLabel: `سنة ${selectedYear + 1}`,
        next: { expense: 0, income: 0 },
      };
    }

    const prevMonth = selectedMonth === 1 ? 12 : selectedMonth - 1;
    const prevYear = selectedMonth === 1 ? selectedYear - 1 : selectedYear;
    const nextMonth = selectedMonth === 12 ? 1 : selectedMonth + 1;
    const nextYear = selectedMonth === 12 ? selectedYear + 1 : selectedYear;

    return {
      prevLabel: `${ARABIC_MONTHS[prevMonth - 1]} ${prevYear}`,
      prev: getMonthTotals(prevYear, prevMonth),
      nextLabel: `${ARABIC_MONTHS[nextMonth - 1]} ${nextYear}`,
      next: getMonthTotals(nextYear, nextMonth),
    };
  }, [selectedMonth, selectedYear, transactions, getMonthTotals]);

  const maxChartVal = useMemo(() => {
    const m = Math.max(...monthlySeries.map((x) => Math.max(x.expense, x.income)), 1);
    return m;
  }, [monthlySeries]);

  const stepMonth = (delta: number) => {
    if (selectedMonth === 0) {
      setSelectedYear((y) => y + delta);
      return;
    }
    let m = selectedMonth + delta;
    let y = selectedYear;
    if (m < 1) {
      m = 12;
      y -= 1;
    } else if (m > 12) {
      m = 1;
      y += 1;
    }
    setSelectedYear(y);
    setSelectedMonth(m);
  };

  const calcDeltaPct = (curr: number, prev: number) => {
    if (prev <= 0 && curr <= 0) return null;
    if (prev <= 0) return 100;
    return Math.round(((curr - prev) / prev) * 100);
  };

  const expDelta = calcDeltaPct(periodStats.expense, adjacentComparison.prev.expense);
  const incDelta = calcDeltaPct(periodStats.income, adjacentComparison.prev.income);

  const periodTitle =
    selectedMonth === 0
      ? `سنة ${selectedYear} كاملة`
      : `${ARABIC_MONTHS[selectedMonth - 1]} ${selectedYear}`;

  return (
    <div className="max-w-4xl mx-auto space-y-4">
      {/* شريط التحكم في السنة والشهر وإظهار/إخفاء الأقسام */}
      <div
        className="rounded-2xl p-4 border space-y-3"
        style={{
          background: 'var(--card-bg)',
          borderColor: 'var(--border)',
          boxShadow: 'var(--shadow-md)',
        }}
      >
        <div className="flex flex-wrap items-center justify-between gap-2.5">
          <div className="flex items-center gap-2.5">
            <div className="theme-icon-box">
              <Calendar className="w-4 h-4" />
            </div>
            <div>
              <h2 className="text-sm font-extrabold">الفترة الزمنية للتحليل</h2>
              <p className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                {periodTitle} · {scopedTxs.length} حركة مسجلة
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-1.5">
            <button
              type="button"
              onClick={() => stepMonth(-1)}
              title="الفترة السابقة"
              className="w-9 h-9 rounded-xl border flex items-center justify-center cursor-pointer"
              style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
            >
              <ChevronRight className="w-4 h-4" />
            </button>

            <select
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(Number(e.target.value))}
              className="h-9 px-2.5 rounded-xl border text-xs font-bold outline-none cursor-pointer"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            >
              <option value={0}>السنة كاملة (كل الشهور)</option>
              {ARABIC_MONTHS.map((name, idx) => (
                <option key={name} value={idx + 1}>
                  {String(idx + 1).padStart(2, '0')} - {name}
                </option>
              ))}
            </select>

            <select
              value={selectedYear}
              onChange={(e) => setSelectedYear(Number(e.target.value))}
              className="h-9 px-2.5 rounded-xl border text-xs font-bold font-mono-num outline-none cursor-pointer"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            >
              {availableYears.map((yr) => (
                <option key={yr} value={yr}>
                  {yr}
                </option>
              ))}
            </select>

            <button
              type="button"
              onClick={() => stepMonth(1)}
              title="الفترة التالية"
              className="w-9 h-9 rounded-xl border flex items-center justify-center cursor-pointer"
              style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* أزرار إظهار / إخفاء الأقسام لمنع زحمة الشاشة */}
        <div
          className="pt-2.5 border-t flex flex-wrap items-center gap-1.5"
          style={{ borderColor: 'var(--border)' }}
        >
          <span className="text-[11px] font-bold ml-1" style={{ color: 'var(--text-muted)' }}>
            عرض الأقسام:
          </span>
          {(
            [
              { key: 'summary', label: 'ملخص ومقارنة' },
              { key: 'yearlyChart', label: 'شهور السنة (1-12)' },
              { key: 'expenseCats', label: 'تصنيف المصروفات' },
              { key: 'incomeCats', label: 'تصنيف القبض' },
            ] as const
          ).map((item) => {
            const on = panels[item.key];
            return (
              <button
                key={item.key}
                type="button"
                onClick={() => togglePanel(item.key)}
                className="px-2.5 py-1.5 rounded-lg border text-[11px] font-bold flex items-center gap-1 cursor-pointer transition-colors"
                style={{
                  background: on ? 'var(--accent-gradient)' : 'var(--surface-soft)',
                  borderColor: on ? 'transparent' : 'var(--border)',
                  color: on ? 'var(--accent-contrast)' : 'var(--text-muted)',
                }}
              >
                {on ? <Eye className="w-3 h-3" /> : <EyeOff className="w-3 h-3" />}
                <span>{item.label}</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* القسم 1: ملخص الفترة والمقارنة بالشهر السابق والتالي */}
      {panels.summary && (
        <div className="space-y-3">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
            <div
              className="rounded-2xl p-3.5 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
            >
              <div className="text-[11px] font-bold" style={{ color: 'var(--text-muted)' }}>
                إجمالي الصرف
              </div>
              <div
                className="text-base font-extrabold mt-1 font-mono-num"
                dir="ltr"
                style={{ color: 'var(--expense)' }}
              >
                {fmt(periodStats.expense)}
              </div>
              {expDelta !== null && (
                <div
                  className="text-[10px] font-bold mt-1 flex items-center gap-0.5"
                  style={{ color: expDelta > 0 ? 'var(--expense)' : 'var(--income)' }}
                >
                  {expDelta > 0 ? (
                    <ArrowUpRight className="w-3 h-3" />
                  ) : (
                    <ArrowDownRight className="w-3 h-3" />
                  )}
                  <span>
                    {expDelta > 0 ? '+' : ''}
                    {expDelta}% عن {adjacentComparison.prevLabel}
                  </span>
                </div>
              )}
            </div>

            <div
              className="rounded-2xl p-3.5 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
            >
              <div className="text-[11px] font-bold" style={{ color: 'var(--text-muted)' }}>
                إجمالي القبض
              </div>
              <div
                className="text-base font-extrabold mt-1 font-mono-num"
                dir="ltr"
                style={{ color: 'var(--income)' }}
              >
                {fmt(periodStats.income)}
              </div>
              {incDelta !== null && (
                <div
                  className="text-[10px] font-bold mt-1 flex items-center gap-0.5"
                  style={{ color: incDelta >= 0 ? 'var(--income)' : 'var(--expense)' }}
                >
                  {incDelta >= 0 ? (
                    <ArrowUpRight className="w-3 h-3" />
                  ) : (
                    <ArrowDownRight className="w-3 h-3" />
                  )}
                  <span>
                    {incDelta > 0 ? '+' : ''}
                    {incDelta}% عن {adjacentComparison.prevLabel}
                  </span>
                </div>
              )}
            </div>

            <div
              className="rounded-2xl p-3.5 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
            >
              <div className="text-[11px] font-bold" style={{ color: 'var(--text-muted)' }}>
                صافي الفترة
              </div>
              <div
                className="text-base font-extrabold mt-1 font-mono-num"
                dir="ltr"
                style={{ color: periodStats.net >= 0 ? 'var(--income)' : 'var(--expense)' }}
              >
                {periodStats.net > 0 ? '+' : ''}
                {fmt(periodStats.net)}
              </div>
              <div className="text-[10px] mt-1" style={{ color: 'var(--text-muted)' }}>
                {periodStats.net >= 0 ? 'فائض ادخار' : 'عجز عن الدخل'}
              </div>
            </div>

            <div
              className="rounded-2xl p-3.5 border"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
            >
              <div className="text-[11px] font-bold" style={{ color: 'var(--text-muted)' }}>
                متوسط الصرف اليومي
              </div>
              <div
                className="text-base font-extrabold mt-1 font-mono-num"
                dir="ltr"
                style={{ color: 'var(--primary-accent)' }}
              >
                {fmt(periodStats.dailyAvg)}
              </div>
              <div className="text-[10px] mt-1 truncate" style={{ color: 'var(--text-muted)' }}>
                الأعلى: {periodStats.expenseList[0]?.[0] || 'لا يوجد'}
              </div>
            </div>
          </div>

          {/* شريط مقارنة سريع: الشهر السابق vs الحالي vs التالي */}
          {selectedMonth > 0 && (
            <div
              className="rounded-2xl p-3.5 border grid grid-cols-1 sm:grid-cols-3 gap-2.5"
              style={{ background: 'var(--card-bg)', borderColor: 'var(--border)' }}
            >
              <div
                className="p-2.5 rounded-xl border"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
              >
                <div className="text-[11px] font-bold" style={{ color: 'var(--text-muted)' }}>
                  الشهر السابق · {adjacentComparison.prevLabel}
                </div>
                <div className="flex justify-between text-xs font-bold mt-1.5 font-mono-num">
                  <span style={{ color: 'var(--expense)' }}>
                    صرف: {fmtCompact(adjacentComparison.prev.expense)}
                  </span>
                  <span style={{ color: 'var(--income)' }}>
                    قبض: {fmtCompact(adjacentComparison.prev.income)}
                  </span>
                </div>
              </div>

              <div
                className="p-2.5 rounded-xl border"
                style={{
                  background: 'var(--surface-soft)',
                  borderColor: 'var(--primary-accent)',
                }}
              >
                <div className="text-[11px] font-extrabold" style={{ color: 'var(--primary-accent)' }}>
                  الشهر المختار · {periodTitle}
                </div>
                <div className="flex justify-between text-xs font-extrabold mt-1.5 font-mono-num">
                  <span style={{ color: 'var(--expense)' }}>
                    صرف: {fmtCompact(periodStats.expense)}
                  </span>
                  <span style={{ color: 'var(--income)' }}>
                    قبض: {fmtCompact(periodStats.income)}
                  </span>
                </div>
              </div>

              <div
                className="p-2.5 rounded-xl border"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
              >
                <div className="text-[11px] font-bold" style={{ color: 'var(--text-muted)' }}>
                  الشهر التالي · {adjacentComparison.nextLabel}
                </div>
                <div className="flex justify-between text-xs font-bold mt-1.5 font-mono-num">
                  <span style={{ color: 'var(--expense)' }}>
                    صرف: {fmtCompact(adjacentComparison.next.expense)}
                  </span>
                  <span style={{ color: 'var(--income)' }}>
                    قبض: {fmtCompact(adjacentComparison.next.income)}
                  </span>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* القسم 2: مقارنة شهور السنة كاملة (يناير إلى ديسمبر) */}
      {panels.yearlyChart && (
        <div
          className="rounded-2xl p-4 sm:p-5 border space-y-4"
          style={{
            background: 'var(--card-bg)',
            borderColor: 'var(--border)',
            boxShadow: 'var(--shadow-md)',
          }}
        >
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2.5">
              <div className="theme-icon-box">
                <BarChart3 className="w-4 h-4" />
              </div>
              <div>
                <h3 className="text-sm font-extrabold">
                  مقارنة شهور سنة {selectedYear} (يناير — ديسمبر)
                </h3>
                <p className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                  اضغط على أي شهر لعرض تفصيله بالأسفل
                </p>
              </div>
            </div>
            <div className="flex items-center gap-3 text-[11px] font-bold">
              <span className="flex items-center gap-1" style={{ color: 'var(--income)' }}>
                <span className="w-2.5 h-2.5 rounded-xs inline-block" style={{ background: 'var(--income)' }} />
                قبض
              </span>
              <span className="flex items-center gap-1" style={{ color: 'var(--expense)' }}>
                <span className="w-2.5 h-2.5 rounded-xs inline-block" style={{ background: 'var(--expense)' }} />
                صرف
              </span>
            </div>
          </div>

          {/* مخطط الأعمدة لـ 12 شهراً */}
          <div className="overflow-x-auto pb-1">
            <div className="grid grid-cols-12 gap-1.5 min-w-[540px] items-end h-40 pt-3">
              {monthlySeries.map((m) => {
                const incH = Math.round((m.income / maxChartVal) * 100);
                const expH = Math.round((m.expense / maxChartVal) * 100);
                const isSelected = selectedMonth === m.month;
                return (
                  <button
                    key={m.month}
                    type="button"
                    onClick={() => setSelectedMonth(m.month)}
                    className="flex flex-col items-center gap-1.5 p-1 rounded-xl border cursor-pointer transition-colors"
                    style={{
                      background: isSelected ? 'var(--surface-soft)' : 'transparent',
                      borderColor: isSelected ? 'var(--primary-accent)' : 'transparent',
                    }}
                    title={`${m.name}: قبض ${fmt(m.income)} | صرف ${fmt(m.expense)}`}
                  >
                    <div className="w-full h-24 flex items-end justify-center gap-1">
                      <div
                        className="w-2.5 rounded-t-sm transition-all"
                        style={{
                          height: `${Math.max(incH, m.income > 0 ? 6 : 2)}%`,
                          background: 'var(--income)',
                          opacity: m.income > 0 ? 1 : 0.18,
                        }}
                      />
                      <div
                        className="w-2.5 rounded-t-sm transition-all"
                        style={{
                          height: `${Math.max(expH, m.expense > 0 ? 6 : 2)}%`,
                          background: 'var(--expense)',
                          opacity: m.expense > 0 ? 1 : 0.18,
                        }}
                      />
                    </div>
                    <span
                      className="text-[10px] font-bold truncate"
                      style={{ color: isSelected ? 'var(--primary-accent)' : 'var(--text-muted)' }}
                    >
                      {m.name}
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* جدول مقارنة شهري مضغوط للشهر المختار والشهور النشطة */}
          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr
                  className="border-b text-[11px]"
                  style={{ borderColor: 'var(--border)', color: 'var(--text-muted)' }}
                >
                  <th className="py-2 text-right font-bold">الشهر</th>
                  <th className="py-2 text-left font-bold">القبض</th>
                  <th className="py-2 text-left font-bold">الصرف</th>
                  <th className="py-2 text-left font-bold">الصافي</th>
                </tr>
              </thead>
              <tbody className="divide-y" style={{ borderColor: 'var(--border)' }}>
                {monthlySeries
                  .filter((m) => m.income > 0 || m.expense > 0 || m.month === selectedMonth)
                  .map((m) => {
                    const active = selectedMonth === m.month;
                    return (
                      <tr
                        key={m.month}
                        onClick={() => setSelectedMonth(m.month)}
                        className="cursor-pointer transition-colors"
                        style={{
                          background: active ? 'var(--surface-soft)' : 'transparent',
                        }}
                      >
                        <td className="py-2 font-bold">
                          {m.name} {selectedYear}
                        </td>
                        <td className="py-2 text-left font-mono-num font-bold" dir="ltr" style={{ color: 'var(--income)' }}>
                          {fmt(m.income)}
                        </td>
                        <td className="py-2 text-left font-mono-num font-bold" dir="ltr" style={{ color: 'var(--expense)' }}>
                          {fmt(m.expense)}
                        </td>
                        <td
                          className="py-2 text-left font-mono-num font-extrabold"
                          dir="ltr"
                          style={{ color: m.net >= 0 ? 'var(--income)' : 'var(--expense)' }}
                        >
                          {m.net > 0 ? '+' : ''}
                          {fmt(m.net)}
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* القسم 3: تصنيف المصروفات والميزانية */}
      {panels.expenseCats && (
        <div
          className="rounded-2xl p-4 sm:p-5 border space-y-3"
          style={{
            background: 'var(--card-bg)',
            borderColor: 'var(--border)',
            boxShadow: 'var(--shadow-md)',
          }}
        >
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2.5">
              <div className="theme-icon-box">
                <PieChart className="w-4 h-4" />
              </div>
              <div>
                <h3 className="text-sm font-extrabold">
                  تصنيف المصروفات — {periodTitle}
                </h3>
                <p className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                  اضغط على أي تصنيف لتحديد ميزانيته الشهرية
                </p>
              </div>
            </div>
            <span className="text-xs font-extrabold font-mono-num" dir="ltr" style={{ color: 'var(--expense)' }}>
              {fmt(periodStats.expense)} ج.م
            </span>
          </div>

          {periodStats.expenseList.length === 0 ? (
            <div className="text-center py-6 text-xs" style={{ color: 'var(--text-muted)' }}>
              لا توجد مصروفات مسجلة في {periodTitle}
            </div>
          ) : (
            <div className="space-y-3">
              {periodStats.expenseList.map(([cat, amt]) => {
                const bud = selectedMonth > 0 ? budgets[cat] : 0;
                const base = bud || periodStats.expense;
                const pct = base ? Math.round((amt / base) * 100) : 0;
                const barColor = bud
                  ? pct >= 100
                    ? 'var(--expense)'
                    : pct >= 80
                    ? 'var(--warning)'
                    : 'var(--income)'
                  : 'var(--primary-accent)';

                return (
                  <button
                    key={cat}
                    type="button"
                    onClick={() => onSelectBudgetCat(cat, budgets[cat] || 0)}
                    className="w-full text-right cursor-pointer"
                  >
                    <div className="flex justify-between text-xs font-bold mb-1">
                      <span>{cat}</span>
                      <span dir="ltr" className="font-mono-num">
                        {bud
                          ? `${fmt(amt)} / ${fmt(bud)} ج.م (${pct}%)`
                          : `${fmt(amt)} ج.م (${pct}%)`}
                      </span>
                    </div>
                    <div
                      className="w-full h-2 rounded-full overflow-hidden"
                      style={{ background: 'var(--surface-soft)' }}
                    >
                      <div
                        className="h-full rounded-full transition-all"
                        style={{ width: `${Math.min(pct, 100)}%`, background: barColor }}
                      />
                    </div>
                  </button>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* القسم 4: تصنيف الدخل / القبض (اختياري) */}
      {panels.incomeCats && (
        <div
          className="rounded-2xl p-4 sm:p-5 border space-y-3"
          style={{
            background: 'var(--card-bg)',
            borderColor: 'var(--border)',
            boxShadow: 'var(--shadow-md)',
          }}
        >
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2.5">
              <div className="theme-icon-box">
                <TrendingUp className="w-4 h-4" />
              </div>
              <div>
                <h3 className="text-sm font-extrabold">
                  تصنيف الدخل / القبض — {periodTitle}
                </h3>
                <p className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                  توزيع مصادر الدخل خلال الفترة المختارة
                </p>
              </div>
            </div>
            <span className="text-xs font-extrabold font-mono-num" dir="ltr" style={{ color: 'var(--income)' }}>
              {fmt(periodStats.income)} ج.م
            </span>
          </div>

          {periodStats.incomeList.length === 0 ? (
            <div className="text-center py-6 text-xs" style={{ color: 'var(--text-muted)' }}>
              لا توجد حركات قبض في {periodTitle}
            </div>
          ) : (
            <div className="space-y-3">
              {periodStats.incomeList.map(([cat, amt]) => {
                const pct = periodStats.income ? Math.round((amt / periodStats.income) * 100) : 0;
                return (
                  <div key={cat}>
                    <div className="flex justify-between text-xs font-bold mb-1">
                      <span>{cat}</span>
                      <span dir="ltr" className="font-mono-num">
                        {fmt(amt)} ج.م ({pct}%)
                      </span>
                    </div>
                    <div
                      className="w-full h-2 rounded-full overflow-hidden"
                      style={{ background: 'var(--surface-soft)' }}
                    >
                      <div
                        className="h-full rounded-full transition-all"
                        style={{ width: `${Math.min(pct, 100)}%`, background: 'var(--income)' }}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
};
