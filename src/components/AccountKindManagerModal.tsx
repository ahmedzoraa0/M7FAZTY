import React, { useEffect, useState } from 'react';
import {
  CheckCircle2,
  CreditCard,
  ExternalLink,
  Landmark,
  ShieldCheck,
  Sliders,
  Vault,
  Wallet,
  X,
} from 'lucide-react';
import { AccountItem, AccountKind } from '../types';

const KINDS: Array<{ value: AccountKind; label: string; icon: React.ReactNode }> = [
  { value: 'محفظة', label: 'محفظة إلكترونية', icon: <Wallet className="w-4 h-4" /> },
  { value: 'بنك', label: 'حساب بنكي', icon: <Landmark className="w-4 h-4" /> },
  { value: 'خزنة', label: 'خزنة / كاش', icon: <Vault className="w-4 h-4" /> },
  { value: 'بطاقة', label: 'بطاقة / فيزا', icon: <CreditCard className="w-4 h-4" /> },
  { value: 'أخرى', label: 'أخرى', icon: <Sliders className="w-4 h-4" /> },
];

interface AccountKindManagerModalProps {
  open: boolean;
  accounts: AccountItem[];
  storedPin: string | null;
  sheetUrl: string;
  onClose: () => void;
  onConfirmChange: (accountName: string, newKind: AccountKind) => Promise<void>;
}

export const AccountKindManagerModal: React.FC<AccountKindManagerModalProps> = ({
  open,
  accounts,
  storedPin,
  sheetUrl,
  onClose,
  onConfirmChange,
}) => {
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [pinInput, setPinInput] = useState('');
  const [safetyChecked, setSafetyChecked] = useState(false);
  const [confirmWord, setConfirmWord] = useState('');
  const [stepError, setStepError] = useState('');

  const [selectedAccount, setSelectedAccount] = useState<string>('');
  const [newKind, setNewKind] = useState<AccountKind>('محفظة');
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (open) {
      setStep(1);
      setPinInput('');
      setSafetyChecked(false);
      setConfirmWord('');
      setStepError('');
      setSaving(false);
      if (accounts.length > 0) {
        setSelectedAccount(accounts[0].name);
        setNewKind(accounts[0].kind || 'محفظة');
      }
    }
  }, [open, accounts]);

  if (!open) return null;

  const currentAccObj = accounts.find((a) => a.name === selectedAccount) || accounts[0];
  const currentKind: AccountKind = currentAccObj?.kind || 'محفظة';

  const handleUnlockStep1 = () => {
    setStepError('');
    if (storedPin) {
      if (pinInput.trim() !== storedPin) {
        setStepError('رمز الـ PIN غير صحيح');
        return;
      }
    } else {
      if (!safetyChecked || confirmWord.trim() !== 'تأكيد') {
        setStepError('فعّل خانة الموافقة واكتب كلمة «تأكيد» للمتابعة بأمان');
        return;
      }
    }
    setStep(2);
  };

  const handleProceedToReview = () => {
    setStepError('');
    if (!currentAccObj) {
      setStepError('اختر حساباً أولاً');
      return;
    }
    if (newKind === currentKind) {
      setStepError('اختر نوعاً مختلفاً عن النوع الحالي');
      return;
    }
    setStep(3);
  };

  const handleFinalSave = async () => {
    if (!currentAccObj) return;
    setSaving(true);
    try {
      await onConfirmChange(currentAccObj.name, newKind);
      onClose();
    } catch (err) {
      setStepError((err as Error).message || 'تعذر حفظ التعديل');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/55 backdrop-blur-sm"
      onClick={(e) => {
        if (e.target === e.currentTarget && !saving) onClose();
      }}
    >
      <div
        className="w-full max-w-md rounded-2xl p-5 border space-y-4"
        style={{
          background: 'var(--card-bg)',
          borderColor: 'var(--border)',
          boxShadow: 'var(--shadow-md)',
        }}
      >
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-2.5">
            <div className="theme-icon-box">
              <ShieldCheck className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-extrabold">تعديل نوع الحساب بأمان</h3>
              <p className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                خطوة {step} من 3 · مرتبط بعمود F في ورقة القوائم والإعدادات
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            disabled={saving}
            className="w-8 h-8 rounded-lg border flex items-center justify-center cursor-pointer"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text-muted)',
            }}
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* مؤشر الخطوات */}
        <div className="grid grid-cols-3 gap-1.5">
          {[1, 2, 3].map((s) => (
            <div
              key={s}
              className="h-1.5 rounded-full transition-colors"
              style={{
                background: step >= s ? 'var(--primary-accent)' : 'var(--surface-soft)',
              }}
            />
          ))}
        </div>

        {/* الخطوة 1: بوابة التحقق الأمني */}
        {step === 1 && (
          <div className="space-y-3">
            <div
              className="p-3 rounded-xl border text-xs leading-relaxed"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            >
              تعديل تصنيف الحساب (بنك / محفظة / خزنة / بطاقة) يُحدّث ورقة{' '}
              <strong>«القوائم والإعدادات»</strong> في شيت جوجل دون المساس بالأرصدة أو المعادلات.
              يمكنك أيضاً تعديله يدوياً مباشرة من شيت جوجل.
            </div>

            {storedPin ? (
              <div>
                <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
                  أدخل رمز الـ PIN الحالي للمتابعة
                </label>
                <input
                  type="password"
                  inputMode="numeric"
                  maxLength={4}
                  value={pinInput}
                  onChange={(e) => setPinInput(e.target.value)}
                  placeholder="••••"
                  className="w-full min-h-[44px] px-3 rounded-xl border text-center text-lg font-mono-num tracking-widest outline-none"
                  style={{
                    background: 'var(--surface-soft)',
                    borderColor: 'var(--border)',
                    color: 'var(--text)',
                  }}
                />
              </div>
            ) : (
              <div className="space-y-2.5">
                <label className="flex items-start gap-2 text-xs cursor-pointer">
                  <input
                    type="checkbox"
                    checked={safetyChecked}
                    onChange={(e) => setSafetyChecked(e.target.checked)}
                    className="mt-0.5"
                  />
                  <span>أؤكد رغبتي في تعديل نوع أحد الحسابات في شيت Google Sheets</span>
                </label>
                <div>
                  <label className="block text-[11px] font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
                    اكتب كلمة «تأكيد» لفك قفل التعديل
                  </label>
                  <input
                    type="text"
                    value={confirmWord}
                    onChange={(e) => setConfirmWord(e.target.value)}
                    placeholder="تأكيد"
                    className="w-full min-h-[42px] px-3 rounded-xl border text-sm outline-none"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--border)',
                      color: 'var(--text)',
                    }}
                  />
                </div>
              </div>
            )}

            {stepError && (
              <div className="text-xs font-bold" style={{ color: 'var(--expense)' }}>
                {stepError}
              </div>
            )}

            <div className="flex items-center justify-between gap-2 pt-1">
              <a
                href={sheetUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="px-3 py-2 rounded-xl border text-xs font-bold inline-flex items-center gap-1.5"
                style={{
                  background: 'var(--surface-soft)',
                  borderColor: 'var(--border)',
                  color: 'var(--text)',
                }}
              >
                <ExternalLink className="w-3.5 h-3.5" />
                <span>تعديل من الشيت</span>
              </a>
              <button
                type="button"
                onClick={handleUnlockStep1}
                className="px-4 py-2 rounded-xl text-xs font-extrabold cursor-pointer"
                style={{
                  background: 'var(--accent-gradient)',
                  color: 'var(--accent-contrast)',
                }}
              >
                التالي (اختيار الحساب) ←
              </button>
            </div>
          </div>
        )}

        {/* الخطوة 2: اختيار الحساب والنوع الجديد */}
        {step === 2 && (
          <div className="space-y-3">
            <div>
              <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
                اختر الحساب المراد تعديل نوعه
              </label>
              <select
                value={selectedAccount}
                onChange={(e) => {
                  const name = e.target.value;
                  setSelectedAccount(name);
                  const found = accounts.find((a) => a.name === name);
                  if (found?.kind) setNewKind(found.kind);
                  setStepError('');
                }}
                className="w-full min-h-[44px] px-3 rounded-xl border text-sm font-bold outline-none"
                style={{
                  background: 'var(--surface-soft)',
                  borderColor: 'var(--border)',
                  color: 'var(--text)',
                }}
              >
                {accounts.map((a) => (
                  <option key={a.name} value={a.name}>
                    {a.name} — (الحالي: {a.kind || 'محفظة'})
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-bold mb-1.5" style={{ color: 'var(--text-muted)' }}>
                اختر النوع الجديد
              </label>
              <div className="grid grid-cols-2 gap-2">
                {KINDS.map((k) => {
                  const active = newKind === k.value;
                  return (
                    <button
                      key={k.value}
                      type="button"
                      onClick={() => {
                        setNewKind(k.value);
                        setStepError('');
                      }}
                      className="flex items-center gap-2 p-2.5 rounded-xl border text-xs font-bold cursor-pointer"
                      style={{
                        background: active ? 'var(--card-bg)' : 'var(--surface-soft)',
                        borderColor: active ? 'var(--primary-accent)' : 'var(--border)',
                        color: active ? 'var(--primary-accent)' : 'var(--text)',
                      }}
                    >
                      {k.icon}
                      <span>{k.label}</span>
                    </button>
                  );
                })}
              </div>
            </div>

            {stepError && (
              <div className="text-xs font-bold" style={{ color: 'var(--expense)' }}>
                {stepError}
              </div>
            )}

            <div className="flex justify-between gap-2 pt-1">
              <button
                type="button"
                onClick={() => setStep(1)}
                className="px-3.5 py-2 rounded-xl border text-xs font-bold cursor-pointer"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
              >
                رجوع
              </button>
              <button
                type="button"
                onClick={handleProceedToReview}
                className="px-4 py-2 rounded-xl text-xs font-extrabold cursor-pointer"
                style={{
                  background: 'var(--accent-gradient)',
                  color: 'var(--accent-contrast)',
                }}
              >
                مراجعة التعديل ←
              </button>
            </div>
          </div>
        )}

        {/* الخطوة 3: المراجعة والتأكيد النهائي */}
        {step === 3 && currentAccObj && (
          <div className="space-y-3">
            <div
              className="p-3.5 rounded-xl border space-y-2 text-xs"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--primary-accent)',
              }}
            >
              <div className="flex items-center gap-2 font-extrabold" style={{ color: 'var(--primary-accent)' }}>
                <CheckCircle2 className="w-4 h-4" />
                <span>مراجعة التعديل قبل الحفظ في الشيت</span>
              </div>
              <div className="flex justify-between py-1 border-b" style={{ borderColor: 'var(--border)' }}>
                <span style={{ color: 'var(--text-muted)' }}>الحساب:</span>
                <span className="font-extrabold">{currentAccObj.name}</span>
              </div>
              <div className="flex justify-between py-1 border-b" style={{ borderColor: 'var(--border)' }}>
                <span style={{ color: 'var(--text-muted)' }}>النوع السابق:</span>
                <span className="font-bold">{currentKind}</span>
              </div>
              <div className="flex justify-between py-1">
                <span style={{ color: 'var(--text-muted)' }}>النوع الجديد:</span>
                <span className="font-extrabold" style={{ color: 'var(--income)' }}>
                  {newKind}
                </span>
              </div>
              <div className="text-[11px] pt-1" style={{ color: 'var(--text-muted)' }}>
                لن يتم تغيير أي رصيد أو معادلة في ورقة «المعاملات اليومية».
              </div>
            </div>

            {stepError && (
              <div className="text-xs font-bold" style={{ color: 'var(--expense)' }}>
                {stepError}
              </div>
            )}

            <div className="flex justify-between gap-2 pt-1">
              <button
                type="button"
                disabled={saving}
                onClick={() => setStep(2)}
                className="px-3.5 py-2 rounded-xl border text-xs font-bold cursor-pointer"
                style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
              >
                تعديل الاختيار
              </button>
              <button
                type="button"
                disabled={saving}
                onClick={handleFinalSave}
                className="px-4 py-2 rounded-xl text-xs font-extrabold cursor-pointer disabled:opacity-50"
                style={{
                  background: 'var(--accent-gradient)',
                  color: 'var(--accent-contrast)',
                }}
              >
                {saving ? 'جاري الحفظ...' : 'تأكيد وحفظ في الشيت'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
