import React, { useState } from 'react';
import { X, Wallet, Landmark, Vault, CreditCard } from 'lucide-react';
import { AccountKind } from '../types';

interface AddAccountModalProps {
  open: boolean;
  onClose: () => void;
  onSubmit: (account: { name: string; kind: AccountKind; opening: number }) => Promise<void>;
}

const KINDS: Array<{ id: AccountKind; label: string; icon: React.ReactNode }> = [
  { id: 'محفظة', label: 'محفظة إلكترونية', icon: <Wallet className="w-4 h-4" /> },
  { id: 'بنك', label: 'حساب بنكي', icon: <Landmark className="w-4 h-4" /> },
  { id: 'خزنة', label: 'خزنة / كاش', icon: <Vault className="w-4 h-4" /> },
  { id: 'بطاقة', label: 'بطاقة بنكية', icon: <CreditCard className="w-4 h-4" /> },
];

export const AddAccountModal: React.FC<AddAccountModalProps> = ({
  open,
  onClose,
  onSubmit,
}) => {
  const [name, setName] = useState('');
  const [kind, setKind] = useState<AccountKind>('محفظة');
  const [opening, setOpening] = useState('0');
  const [saving, setSaving] = useState(false);

  if (!open) return null;

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    const cleanName = name.trim();
    if (!cleanName) return;
    setSaving(true);
    try {
      await onSubmit({
        name: cleanName,
        kind,
        opening: parseFloat(opening) || 0,
      });
      setName('');
      setOpening('0');
      onClose();
    } finally {
      setSaving(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-end md:items-center justify-center p-3 bg-black/50 backdrop-blur-sm"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        className="w-full max-w-md rounded-2xl p-5 border"
        style={{
          background: 'var(--card-bg)',
          borderColor: 'var(--border)',
          color: 'var(--text)',
          boxShadow: 'var(--shadow-md)',
        }}
      >
        <div className="flex items-center justify-between mb-4">
          <div>
            <h3 className="text-base font-extrabold">إضافة محفظة / بنك / خزنة</h3>
            <p className="text-xs mt-0.5" style={{ color: 'var(--text-muted)' }}>
              يُضاف فوراً في الرئيسية ويُربط بمعادلات الشيت وقائمة الدفع
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="w-9 h-9 rounded-xl flex items-center justify-center border cursor-pointer"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text-muted)',
            }}
            aria-label="إغلاق"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSave} className="space-y-3.5">
          <div>
            <label className="block text-xs font-bold mb-1.5" style={{ color: 'var(--text-muted)' }}>
              اسم الحساب أو المحفظة
            </label>
            <input
              type="text"
              required
              autoFocus
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="مثال: البنك الأهلي، فودافون كاش، خزنة البيت"
              className="w-full min-h-[46px] px-3.5 rounded-xl border text-sm font-medium outline-none"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            />
          </div>

          <div>
            <label className="block text-xs font-bold mb-1.5" style={{ color: 'var(--text-muted)' }}>
              النوع
            </label>
            <div className="grid grid-cols-2 gap-2">
              {KINDS.map((item) => {
                const active = kind === item.id;
                return (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => setKind(item.id)}
                    className="flex items-center gap-2 px-3 py-2.5 rounded-xl border text-xs font-bold cursor-pointer transition-colors"
                    style={{
                      background: active ? 'var(--card-bg)' : 'var(--surface-soft)',
                      borderColor: active ? 'var(--primary-accent)' : 'var(--border)',
                      color: active ? 'var(--primary-accent)' : 'var(--text)',
                    }}
                  >
                    {item.icon}
                    <span>{item.label}</span>
                  </button>
                );
              })}
            </div>
          </div>

          <div>
            <label className="block text-xs font-bold mb-1.5" style={{ color: 'var(--text-muted)' }}>
              الرصيد الافتتاحي (ج.م)
            </label>
            <input
              type="number"
              step="0.01"
              inputMode="decimal"
              value={opening}
              onChange={(e) => setOpening(e.target.value)}
              className="w-full min-h-[46px] px-3.5 rounded-xl border text-base font-extrabold font-mono-num outline-none"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            />
          </div>

          <button
            type="submit"
            disabled={saving}
            className="w-full min-h-[48px] rounded-xl font-extrabold text-sm cursor-pointer disabled:opacity-50"
            style={{
              background: 'var(--accent-gradient)',
              color: 'var(--accent-contrast)',
              boxShadow: '0 10px 24px -8px var(--accent-shadow)',
            }}
          >
            {saving ? 'جاري الإضافة والربط بالشيت...' : 'إضافة وربط بالشيت فوراً'}
          </button>
        </form>
      </div>
    </div>
  );
};
