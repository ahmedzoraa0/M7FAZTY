import React, { useEffect, useState } from 'react';
import { Lock, Delete } from 'lucide-react';

interface PinLockOverlayProps {
  open: boolean;
  mode: 'unlock' | 'setup' | 'change';
  storedPin: string | null;
  onUnlock: () => void;
  onSavePin: (newPin: string) => void;
  onCancel?: () => void;
  onNotify: (msg: string) => void;
}

export const PinLockOverlay: React.FC<PinLockOverlayProps> = ({
  open,
  mode,
  storedPin,
  onUnlock,
  onSavePin,
  onCancel,
  onNotify,
}) => {
  const [input, setInput] = useState('');
  const [step, setStep] = useState<'verify_old' | 'enter_new' | 'confirm_new' | 'unlock'>('unlock');
  const [tempPin, setTempPin] = useState('');
  const [fails, setFails] = useState(0);
  const [lockUntil, setLockUntil] = useState(0);

  useEffect(() => {
    if (!open) return;
    setInput('');
    setTempPin('');
    if (mode === 'setup') setStep('enter_new');
    else if (mode === 'change') setStep(storedPin ? 'verify_old' : 'enter_new');
    else setStep('unlock');
  }, [open, mode, storedPin]);

  const handleComplete = (code: string) => {
    if (step === 'unlock') {
      if (code === storedPin) {
        setFails(0);
        setInput('');
        onUnlock();
      } else {
        const nextFails = fails + 1;
        setFails(nextFails);
        setInput('');
        if (nextFails >= 5) {
          const waitMs = 30000 * Math.min(nextFails - 4, 6);
          setLockUntil(Date.now() + waitMs);
          onNotify(`محاولات كثيرة، انتظر ${Math.round(waitMs / 1000)} ثانية`);
        } else {
          onNotify(`رمز PIN غير صحيح (متبقي ${5 - nextFails} محاولات)`);
        }
      }
      return;
    }

    if (step === 'verify_old') {
      if (code === storedPin) {
        setInput('');
        setStep('enter_new');
      } else {
        setInput('');
        onNotify('رمز PIN الحالي غير صحيح');
      }
      return;
    }

    if (step === 'enter_new') {
      setTempPin(code);
      setInput('');
      setStep('confirm_new');
      return;
    }

    if (step === 'confirm_new') {
      if (code === tempPin) {
        setInput('');
        onSavePin(code);
      } else {
        setInput('');
        setTempPin('');
        setStep('enter_new');
        onNotify('الرمز غير متطابق، أعد المحاولة');
      }
    }
  };

  const pressDigit = (d: string) => {
    if (!open || Date.now() < lockUntil) return;
    if (input.length >= 4) return;
    const next = input + d;
    setInput(next);
    if (next.length === 4) {
      window.setTimeout(() => handleComplete(next), 140);
    }
  };

  const handleBackspace = () => {
    setInput((prev) => prev.slice(0, -1));
  };

  useEffect(() => {
    if (!open) return;
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key >= '0' && e.key <= '9') pressDigit(e.key);
      else if (e.key === 'Backspace') handleBackspace();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  });

  if (!open) return null;

  let title = 'أدخل رمز PIN';
  let desc = 'لفتح ومتابعة محفظتك المالية';
  if (step === 'verify_old') {
    title = 'تغيير رمز PIN';
    desc = 'أدخل رمز PIN الحالي أولاً للمتابعة';
  } else if (step === 'enter_new') {
    title = 'تعيين رمز PIN جديد';
    desc = 'أدخل رمزاً مكوناً من 4 أرقام لحماية التطبيق';
  } else if (step === 'confirm_new') {
    title = 'تأكيد رمز PIN';
    desc = 'أعد إدخال نفس الرمز للتأكيد';
  }

  return (
    <div
      className="fixed inset-0 z-[9999] flex flex-col items-center justify-center px-4 py-6 select-none"
      style={{ background: 'var(--lock-bg)', color: 'var(--lock-text)' }}
    >
      <div
        className="w-16 h-16 rounded-2xl flex items-center justify-center mb-4"
        style={{
          background: 'var(--accent-gradient)',
          color: 'var(--accent-contrast)',
          boxShadow: '0 12px 30px -8px var(--accent-shadow)',
        }}
      >
        <Lock className="w-7 h-7" />
      </div>

      <h2 className="text-xl font-extrabold mb-1.5" style={{ color: 'var(--lock-text)' }}>
        {title}
      </h2>
      <p className="text-sm mb-7 text-center" style={{ color: 'var(--lock-muted)' }}>
        {desc}
      </p>

      <div className="flex gap-3.5 mb-8" dir="ltr">
        {[0, 1, 2, 3].map((i) => {
          const filled = i < input.length;
          return (
            <div
              key={i}
              className="w-4 h-4 rounded-full transition-all duration-150"
              style={{
                border: '2px solid var(--primary-accent)',
                background: filled ? 'var(--primary-accent)' : 'transparent',
                transform: filled ? 'scale(1.15)' : 'scale(1)',
                boxShadow: filled ? '0 0 12px var(--accent-shadow)' : 'none',
              }}
            />
          );
        })}
      </div>

      <div className="grid grid-cols-3 gap-3.5 w-full max-w-[264px]" dir="ltr">
        {['1', '2', '3', '4', '5', '6', '7', '8', '9'].map((digit) => (
          <button
            key={digit}
            type="button"
            onClick={() => pressDigit(digit)}
            className="aspect-square rounded-full text-2xl font-bold flex items-center justify-center transition-transform active:scale-90 cursor-pointer"
            style={{
              background: 'var(--lock-surface)',
              border: '1px solid var(--lock-border)',
              color: 'var(--lock-text)',
            }}
          >
            {digit}
          </button>
        ))}
        <div />
        <button
          type="button"
          onClick={() => pressDigit('0')}
          className="aspect-square rounded-full text-2xl font-bold flex items-center justify-center transition-transform active:scale-90 cursor-pointer"
          style={{
            background: 'var(--lock-surface)',
            border: '1px solid var(--lock-border)',
            color: 'var(--lock-text)',
          }}
        >
          0
        </button>
        <button
          type="button"
          onClick={handleBackspace}
          title="مسح"
          className="aspect-square rounded-full flex items-center justify-center transition-transform active:scale-90 cursor-pointer"
          style={{ color: 'var(--lock-muted)' }}
        >
          <Delete className="w-6 h-6" />
        </button>
      </div>

      {onCancel && (
        <button
          type="button"
          onClick={onCancel}
          className="mt-6 px-4 py-2 text-sm font-bold cursor-pointer"
          style={{ color: 'var(--lock-muted)' }}
        >
          إلغاء والعودة
        </button>
      )}
    </div>
  );
};
