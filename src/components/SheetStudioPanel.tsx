import React, { useState } from 'react';
import {
  AlertTriangle,
  Check,
  Copy,
  Download,
  ExternalLink,
  FilePlus2,
  FileSpreadsheet,
  FolderSearch,
  LogOut,
  PlugZap,
  RefreshCw,
  ShieldAlert,
  Sparkles,
  UserCheck,
} from 'lucide-react';
import { SheetInspectionReport } from '../types';
import { generateAppsScriptCode } from '../templates/appsScriptTemplate';
import { generateStandaloneHtml } from '../templates/standaloneHtmlGenerator';
import {
  DriveSheetMatch,
  buildM7faztySheetName,
  extractSpreadsheetId,
} from '../services/sheetsApi';

interface SheetStudioPanelProps {
  googleUserEmail: string | null;
  hasGoogleToken: boolean;
  isLoggingIn: boolean;
  onGoogleLogin: () => Promise<void>;
  onGoogleLogout: () => Promise<void>;
  profileName: string;
  onProfileNameChange: (val: string) => void;
  spreadsheetIdInput: string;
  onSpreadsheetIdChange: (val: string) => void;
  scriptUrlInput: string;
  onScriptUrlChange: (val: string) => void;
  inspectionReport: SheetInspectionReport | null;
  isInspecting: boolean;
  onInspectSheet: () => Promise<void>;
  onFormatWorkbook: () => Promise<void>;
  existingSheets: DriveSheetMatch[];
  isSearchingDrive: boolean;
  onSearchExistingSheets: (nameOverride?: string) => Promise<DriveSheetMatch[]>;
  onSelectExistingSheet: (sheetId: string) => Promise<void>;
  onCreateNewSheet: (title: string) => Promise<{ spreadsheetId: string; spreadsheetUrl: string } | null>;
  onTestScriptConnection: () => Promise<void>;
  isTestingScript: boolean;
  onNotify: (msg: string) => void;
}

export const SheetStudioPanel: React.FC<SheetStudioPanelProps> = ({
  googleUserEmail,
  hasGoogleToken,
  isLoggingIn,
  onGoogleLogin,
  onGoogleLogout,
  profileName,
  onProfileNameChange,
  spreadsheetIdInput,
  onSpreadsheetIdChange,
  scriptUrlInput,
  onScriptUrlChange,
  inspectionReport,
  isInspecting,
  onInspectSheet,
  onFormatWorkbook,
  existingSheets,
  isSearchingDrive,
  onSearchExistingSheets,
  onSelectExistingSheet,
  onCreateNewSheet,
  onTestScriptConnection,
  isTestingScript,
  onNotify,
}) => {
  const [showFormatConfirm, setShowFormatConfirm] = useState(false);
  const [showCreateConfirm, setShowCreateConfirm] = useState(false);
  const [newSheetTitle, setNewSheetTitle] = useState(() => buildM7faztySheetName(profileName));
  const [driveMatchesInModal, setDriveMatchesInModal] = useState<DriveSheetMatch[]>([]);
  const [checkingBeforeCreate, setCheckingBeforeCreate] = useState(false);
  const [creatingSheet, setCreatingSheet] = useState(false);
  const [selectingSheetId, setSelectingSheetId] = useState<string | null>(null);
  const [formatting, setFormatting] = useState(false);
  const [copiedGs, setCopiedGs] = useState(false);

  const cleanSheetId = extractSpreadsheetId(spreadsheetIdInput);
  const expectedSheetName = buildM7faztySheetName(profileName);
  const appsScriptCode = generateAppsScriptCode('');

  const handleConfirmFormat = async () => {
    setFormatting(true);
    try {
      await onFormatWorkbook();
      setShowFormatConfirm(false);
    } finally {
      setFormatting(false);
    }
  };

  const handleOpenCreateFlow = async () => {
    setNewSheetTitle(buildM7faztySheetName(profileName));
    setShowCreateConfirm(true);
    setCheckingBeforeCreate(true);
    try {
      const found = await onSearchExistingSheets(profileName);
      setDriveMatchesInModal(found);
    } finally {
      setCheckingBeforeCreate(false);
    }
  };

  const handleUseExisting = async (sheetId: string) => {
    setSelectingSheetId(sheetId);
    try {
      await onSelectExistingSheet(sheetId);
      setShowCreateConfirm(false);
    } finally {
      setSelectingSheetId(null);
    }
  };

  const handleConfirmCreateSheet = async () => {
    setCreatingSheet(true);
    try {
      const created = await onCreateNewSheet(newSheetTitle.trim() || expectedSheetName);
      if (created) {
        setShowCreateConfirm(false);
      }
    } finally {
      setCreatingSheet(false);
    }
  };

  const handleDownloadHtml = (starterMode = false) => {
    const cleanProfile = profileName.trim();
    const html = generateStandaloneHtml(
      starterMode
        ? { starterMode: true, profileName: cleanProfile }
        : {
            scriptUrl: scriptUrlInput.trim(),
            spreadsheetId: cleanSheetId,
            profileName: cleanProfile,
          }
    );
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const safeSlug = cleanProfile.replace(/[^\w\u0600-\u06FF-]+/g, '_').replace(/^_+|_+$/g, '');
    a.download = starterMode
      ? 'M7FAZTY_Setup.html'
      : safeSlug
      ? `M7FAZTY_${safeSlug}.html`
      : 'M7FAZTY.html';
    a.click();
    URL.revokeObjectURL(url);
    onNotify(
      starterMode
        ? 'تم تحميل نسخة M7FAZTY_Setup.html الفارغة بمعالج البدء لأول مرة'
        : `تم تحميل ملف ${a.download} مربوط بشيتك وجاهز للعمل`
    );
  };

  const handleDownloadAppsScript = () => {
    const blob = new Blob([appsScriptCode], { type: 'text/javascript;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Code.gs';
    a.click();
    URL.revokeObjectURL(url);
    onNotify('تم تحميل ملف Code.gs المحدّث');
  };

  const handleCopyAppsScript = async () => {
    try {
      await navigator.clipboard.writeText(appsScriptCode);
      setCopiedGs(true);
      onNotify('تم نسخ كود Apps Script بالكامل');
      window.setTimeout(() => setCopiedGs(false), 2000);
    } catch {
      onNotify('تعذر النسخ التلقائي، يمكنك تحميل ملف Code.gs');
    }
  };

  return (
    <div className="space-y-4">
      {/* 1. هوية المحفظة + الربط المباشر مع Google Sheets */}
      <div
        className="rounded-2xl p-5 border"
        style={{
          background: 'var(--card-bg)',
          borderColor: 'var(--border)',
          boxShadow: 'var(--shadow-md)',
        }}
      >
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
          <div>
            <h2 className="text-base font-extrabold flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5" style={{ color: 'var(--primary-accent)' }} />
              <span>إنشاء وربط شيت M7FAZTY مع Google Sheets</span>
            </h2>
            <p className="text-xs mt-1" style={{ color: 'var(--text-muted)' }}>
              سجّل الدخول بجوجل لفحص وجود شيت <strong className="font-mono-num">{expectedSheetName}</strong> أولاً والعمل عليه تلقائياً أو إنشاء شيت جديد مهيأ بالكامل.
            </p>
          </div>

          {!hasGoogleToken ? (
            <button
              type="button"
              className="gsi-material-button"
              disabled={isLoggingIn}
              onClick={onGoogleLogin}
            >
              <div className="gsi-material-button-icon">
                <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
                  <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z" />
                  <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z" />
                  <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z" />
                  <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z" />
                </svg>
              </div>
              <span>{isLoggingIn ? 'جاري الربط...' : 'Sign in with Google'}</span>
            </button>
          ) : (
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold" style={{ color: 'var(--income)' }}>
                متصل: {googleUserEmail || 'حساب جوجل'}
              </span>
              <button
                type="button"
                onClick={onGoogleLogout}
                className="px-3 py-2 rounded-xl border text-xs font-bold flex items-center gap-1.5 cursor-pointer"
                style={{
                  background: 'var(--surface-soft)',
                  borderColor: 'var(--border)',
                  color: 'var(--text)',
                }}
              >
                <LogOut className="w-3.5 h-3.5" />
                <span>خروج</span>
              </button>
            </div>
          )}
        </div>

        {/* اسم المستخدم / المحفظة */}
        <div
          className="rounded-xl p-3.5 border mb-4"
          style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}
        >
          <div className="flex flex-wrap items-center justify-between gap-2 mb-1.5">
            <label className="text-xs font-extrabold flex items-center gap-1.5" style={{ color: 'var(--text)' }}>
              <UserCheck className="w-4 h-4" style={{ color: 'var(--primary-accent)' }} />
              <span>اسمك أو اسم المحفظة (يُضاف تلقائياً لاسم الشيت والملف)</span>
            </label>
            <span className="text-xs font-bold font-mono-num" dir="ltr" style={{ color: 'var(--primary-accent)' }}>
              {expectedSheetName}
            </span>
          </div>
          <input
            type="text"
            value={profileName}
            onChange={(e) => onProfileNameChange(e.target.value)}
            placeholder="مثال: أحمد (سيكون اسم الشيت M7FAZTY_أحمد)"
            className="w-full min-h-[42px] px-3 rounded-xl border text-sm font-bold outline-none"
            style={{
              background: 'var(--card-bg)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
            }}
          />
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
          <div>
            <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
              معرّف أو رابط Google Sheet
            </label>
            <input
              type="text"
              dir="ltr"
              value={spreadsheetIdInput}
              onChange={(e) => onSpreadsheetIdChange(e.target.value)}
              className="w-full min-h-[42px] px-3 rounded-xl border text-xs font-mono-num outline-none"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            />
          </div>
          <div>
            <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
              رابط نشر Google Apps Script (Web App URL)
            </label>
            <input
              type="url"
              dir="ltr"
              value={scriptUrlInput}
              onChange={(e) => onScriptUrlChange(e.target.value)}
              className="w-full min-h-[42px] px-3 rounded-xl border text-xs font-mono-num outline-none"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            />
          </div>
        </div>

        <div className="flex flex-wrap gap-2.5">
          <button
            type="button"
            disabled={!hasGoogleToken || isSearchingDrive}
            onClick={handleOpenCreateFlow}
            className="px-4 py-2.5 rounded-xl text-xs font-extrabold flex items-center gap-2 cursor-pointer disabled:opacity-45"
            style={{
              background: 'var(--accent-gradient)',
              color: 'var(--accent-contrast)',
              boxShadow: '0 8px 20px -8px var(--accent-shadow)',
            }}
          >
            <FilePlus2 className="w-4 h-4" />
            <span>إنشاء أو ربط شيت {expectedSheetName}</span>
          </button>

          <button
            type="button"
            disabled={!hasGoogleToken || isSearchingDrive}
            onClick={() => onSearchExistingSheets(profileName)}
            className="px-4 py-2.5 rounded-xl border text-xs font-bold flex items-center gap-2 cursor-pointer disabled:opacity-45"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
            }}
          >
            <FolderSearch className={`w-4 h-4 ${isSearchingDrive ? 'animate-pulse' : ''}`} />
            <span>{isSearchingDrive ? 'جاري البحث في Drive...' : 'فحص شيتات M7FAZTY الموجودة'}</span>
          </button>

          <button
            type="button"
            disabled={!hasGoogleToken || isInspecting}
            onClick={onInspectSheet}
            className="px-4 py-2.5 rounded-xl border text-xs font-bold flex items-center gap-2 cursor-pointer disabled:opacity-45"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
            }}
          >
            <RefreshCw className={`w-4 h-4 ${isInspecting ? 'animate-spin' : ''}`} />
            <span>قراءة الشيت المحدد</span>
          </button>

          <button
            type="button"
            disabled={!hasGoogleToken}
            onClick={() => setShowFormatConfirm(true)}
            className="px-4 py-2.5 rounded-xl border text-xs font-bold flex items-center gap-2 cursor-pointer disabled:opacity-45"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--primary-accent)',
              color: 'var(--primary-accent)',
            }}
          >
            <Sparkles className="w-4 h-4" />
            <span>إصلاح المعادلات وتنسيق الشيت</span>
          </button>

          <button
            type="button"
            disabled={isTestingScript || !scriptUrlInput.trim()}
            onClick={onTestScriptConnection}
            className="px-4 py-2.5 rounded-xl border text-xs font-bold flex items-center gap-2 cursor-pointer disabled:opacity-45"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
            }}
          >
            <PlugZap className="w-4 h-4" />
            <span>{isTestingScript ? 'جارٍ الاختبار...' : 'اختبار Apps Script'}</span>
          </button>

          {cleanSheetId && (
            <a
              href={`https://docs.google.com/spreadsheets/d/${cleanSheetId}/edit`}
              target="_blank"
              rel="noreferrer"
              className="px-4 py-2.5 rounded-xl border text-xs font-bold flex items-center gap-1.5 no-underline"
              style={{
                background: 'var(--surface-soft)',
                borderColor: 'var(--border)',
                color: 'var(--text)',
              }}
            >
              <span>فتح الشيت</span>
              <ExternalLink className="w-3.5 h-3.5" />
            </a>
          )}
        </div>

        {/* شيتات M7FAZTY المكتشفة على Drive */}
        {hasGoogleToken && existingSheets.length > 0 && (
          <div
            className="mt-4 pt-4 border-t space-y-2"
            style={{ borderColor: 'var(--border)' }}
          >
            <div className="text-xs font-extrabold flex items-center justify-between" style={{ color: 'var(--primary-accent)' }}>
              <span>شيتات M7FAZTY الموجودة على Google Drive ({existingSheets.length})</span>
              <span style={{ color: 'var(--text-muted)' }}>اضغط للعمل على أي شيت مباشرة بدون إنشاء نسخ مكررة</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              {existingSheets.map((item) => {
                const isActive = item.id === cleanSheetId;
                return (
                  <div
                    key={item.id}
                    className="rounded-xl p-3 border flex items-center justify-between gap-2"
                    style={{
                      background: isActive ? 'var(--card-bg)' : 'var(--surface-soft)',
                      borderColor: isActive ? 'var(--primary-accent)' : 'var(--border)',
                    }}
                  >
                    <div className="min-w-0">
                      <div className="text-xs font-extrabold truncate flex items-center gap-1.5">
                        <span className="truncate">{item.name}</span>
                        {item.isExactTarget && (
                          <span className="text-[10px] font-bold" style={{ color: 'var(--income)' }}>
                            (مطابق للاسم)
                          </span>
                        )}
                      </div>
                      <div className="text-[11px] font-mono-num truncate" dir="ltr" style={{ color: 'var(--text-muted)' }}>
                        {item.modifiedTime ? new Date(item.modifiedTime).toLocaleDateString('ar-EG') : item.id}
                      </div>
                    </div>
                    <button
                      type="button"
                      disabled={selectingSheetId === item.id}
                      onClick={() => handleUseExisting(item.id)}
                      className="px-3 py-1.5 rounded-lg text-[11px] font-extrabold shrink-0 cursor-pointer"
                      style={{
                        background: isActive ? 'var(--income-bg)' : 'var(--accent-gradient)',
                        color: isActive ? 'var(--income)' : 'var(--accent-contrast)',
                      }}
                    >
                      {selectingSheetId === item.id
                        ? '...'
                        : isActive
                        ? 'مُفعّل حالياً ✓'
                        : 'تشغيل هذا الشيت'}
                    </button>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {inspectionReport && (
          <div
            className="mt-4 pt-4 border-t grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs"
            style={{ borderColor: 'var(--border)' }}
          >
            <div>
              <div style={{ color: 'var(--text-muted)' }}>اسم الملف</div>
              <div className="font-bold mt-0.5 truncate">{inspectionReport.spreadsheetTitle}</div>
            </div>
            <div>
              <div style={{ color: 'var(--text-muted)' }}>الحسابات وبداية المعاملات</div>
              <div className="font-bold mt-0.5 tabular-nums">
                {inspectionReport.accountCount} حساب · صف {inspectionReport.firstTxRow}
              </div>
            </div>
            <div>
              <div style={{ color: 'var(--text-muted)' }}>التصنيفات</div>
              <div className="font-bold mt-0.5 tabular-nums">
                {inspectionReport.expenseCatCount} صرف · {inspectionReport.incomeCatCount} قبض
              </div>
            </div>
            <div>
              <div style={{ color: 'var(--text-muted)' }}>السجلات المقروءة</div>
              <div className="font-bold mt-0.5 tabular-nums">
                {inspectionReport.txCount} حركة · {inspectionReport.debtCount} مستحق
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 2. دليل التشغيل لأول مرة وربط Apps Script + تحميل ملفات HTML */}
      <div
        className="rounded-2xl p-5 border space-y-4"
        style={{
          background: 'var(--card-bg)',
          borderColor: 'var(--border)',
          boxShadow: 'var(--shadow-md)',
        }}
      >
        <div>
          <h3 className="text-base font-extrabold">تصدير ملفات M7FAZTY المستقلة وكود Apps Script</h3>
          <p className="text-xs mt-1 leading-relaxed" style={{ color: 'var(--text-muted)' }}>
            كل ملف HTML يتم تصديره أو ربطه يعمل بمساحة تخزين مستقلة تماماً بناءً على رابط الشيت/السكربت الخاص به — يمكنك فتح أكثر من ملف (مثل M7FAZTY_1 و M7FAZTY_2) في نفس المتصفح بدون أي تداخل.
          </p>
        </div>

        <ol className="text-xs space-y-2 pr-5 list-decimal leading-relaxed" style={{ color: 'var(--text)' }}>
          <li>
            أدخل اسمك بالأعلى، واضغط <span className="font-bold">«إنشاء أو ربط شيت {expectedSheetName}»</span> (إذا وُجد شيت سابق سيشغّله تلقائياً بدلاً من تكراره).
          </li>
          <li>انسخ أو حمّل كود <span className="font-bold">Code.gs</span> المحدّث من الأسفل.</li>
          <li>
            من داخل الشيت افتح: <span className="font-bold">الإضافات (Extensions) ← Apps Script</span>، الصق الكود واحفظه (Ctrl+S).
          </li>
          <li>
            اضغط <span className="font-bold">نشر (Deploy) ← نشر جديد ← تطبيق ويب (Web app)</span>، واجعل التنفيذ <span className="font-bold">أنا (Me)</span> والوصول <span className="font-bold">أي شخص (Anyone)</span> ثم انسخ الرابط المنتهي بـ <code className="font-mono-num">/exec</code>.
          </li>
          <li>
            الصق الرابط في خانة Web App URL، اختبره، ثم حمّل ملف الـ HTML الخاص بهذه المحفظة.
          </li>
        </ol>

        <div className="flex flex-wrap gap-2.5">
          <button
            type="button"
            onClick={() => handleDownloadHtml(false)}
            className="px-4 py-2.5 rounded-xl text-xs font-extrabold flex items-center gap-2 cursor-pointer"
            style={{
              background: 'var(--accent-gradient)',
              color: 'var(--accent-contrast)',
              boxShadow: '0 8px 20px -8px var(--accent-shadow)',
            }}
          >
            <Download className="w-4 h-4" />
            <span>تحميل {profileName.trim() ? `M7FAZTY_${profileName.trim()}.html` : 'M7FAZTY.html'} (مربوط)</span>
          </button>

          <button
            type="button"
            onClick={() => handleDownloadHtml(true)}
            className="px-4 py-2.5 rounded-xl border text-xs font-bold flex items-center gap-2 cursor-pointer"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--primary-accent)',
              color: 'var(--primary-accent)',
            }}
          >
            <Download className="w-4 h-4" />
            <span>تحميل M7FAZTY_Setup.html (معالج لأول مرة)</span>
          </button>

          <button
            type="button"
            onClick={handleCopyAppsScript}
            className="px-4 py-2.5 rounded-xl border text-xs font-bold flex items-center gap-2 cursor-pointer"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
            }}
          >
            {copiedGs ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
            <span>{copiedGs ? 'تم النسخ' : 'نسخ كود Code.gs'}</span>
          </button>

          <button
            type="button"
            onClick={handleDownloadAppsScript}
            className="px-4 py-2.5 rounded-xl border text-xs font-bold flex items-center gap-2 cursor-pointer"
            style={{
              background: 'var(--surface-soft)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
            }}
          >
            <Download className="w-4 h-4" />
            <span>تحميل Code.gs</span>
          </button>
        </div>

        <details className="rounded-xl border p-3" style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}>
          <summary className="text-xs font-bold cursor-pointer select-none">
            معاينة كود Apps Script المحدّث (Code.gs)
          </summary>
          <pre
            dir="ltr"
            className="p-3 mt-3 rounded-lg border text-[11px] leading-relaxed overflow-x-auto max-h-72 font-mono-num"
            style={{
              background: 'var(--card-bg)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
            }}
          >
            {appsScriptCode}
          </pre>
        </details>
      </div>

      {/* Confirmation dialog before creating or reusing an existing M7FAZTY Google Sheet */}
      {showCreateConfirm && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/55 backdrop-blur-sm"
          onClick={(e) => {
            if (e.target === e.currentTarget && !creatingSheet && !selectingSheetId) {
              setShowCreateConfirm(false);
            }
          }}
        >
          <div
            className="w-full max-w-lg rounded-2xl p-5 border space-y-4"
            style={{
              background: 'var(--card-bg)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
              boxShadow: 'var(--shadow-md)',
            }}
          >
            {checkingBeforeCreate ? (
              <div className="py-6 text-center space-y-2">
                <FolderSearch className="w-8 h-8 mx-auto animate-pulse" style={{ color: 'var(--primary-accent)' }} />
                <div className="text-sm font-extrabold">جاري فحص Google Drive بحثاً عن شيت {expectedSheetName}...</div>
                <div className="text-xs" style={{ color: 'var(--text-muted)' }}>
                  نتأكد أولاً إن كان لديك شيت موجود مسبقاً لتشغيله وتحميل بياناته بدون إنشاء نسخ مكررة.
                </div>
              </div>
            ) : driveMatchesInModal.length > 0 ? (
              <>
                <div className="flex items-start gap-3">
                  <div
                    className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0"
                    style={{ background: 'var(--warning-bg)', color: 'var(--warning)' }}
                  >
                    <AlertTriangle className="w-5 h-5" />
                  </div>
                  <div>
                    <h4 className="text-base font-extrabold">تم العثور على شيت M7FAZTY موجود مسبقاً!</h4>
                    <p className="text-xs mt-1 leading-relaxed" style={{ color: 'var(--text-muted)' }}>
                      يوجد بالفعل {driveMatchesInModal.length} شيت على حسابك. يمكنك تشغيل الشيت الموجود وتحميل بياناته فوراً بدون إنشاء شيت جديد:
                    </p>
                  </div>
                </div>

                <div className="space-y-2 max-h-52 overflow-y-auto">
                  {driveMatchesInModal.map((m) => (
                    <div
                      key={m.id}
                      className="rounded-xl p-3 border flex items-center justify-between gap-2"
                      style={{
                        background: 'var(--surface-soft)',
                        borderColor: m.isExactTarget ? 'var(--primary-accent)' : 'var(--border)',
                      }}
                    >
                      <div className="min-w-0">
                        <div className="text-xs font-extrabold truncate">{m.name}</div>
                        <div className="text-[11px]" style={{ color: 'var(--text-muted)' }}>
                          {m.isExactTarget ? 'مطابق تماماً للاسم المطلوب' : 'شيت M7FAZTY آخر على حسابك'}
                        </div>
                      </div>
                      <button
                        type="button"
                        disabled={Boolean(selectingSheetId) || creatingSheet}
                        onClick={() => handleUseExisting(m.id)}
                        className="px-3.5 py-2 rounded-xl text-xs font-extrabold shrink-0 cursor-pointer"
                        style={{
                          background: 'var(--accent-gradient)',
                          color: 'var(--accent-contrast)',
                        }}
                      >
                        {selectingSheetId === m.id ? 'جاري التحميل...' : 'العمل على هذا الشيت'}
                      </button>
                    </div>
                  ))}
                </div>

                <details className="rounded-xl border p-3" style={{ background: 'var(--surface-soft)', borderColor: 'var(--border)' }}>
                  <summary className="text-xs font-bold cursor-pointer" style={{ color: 'var(--text-muted)' }}>
                    أو إنشاء شيت جديد منفصل رغم وجود الشيت الحالي
                  </summary>
                  <div className="mt-3 space-y-2.5">
                    <input
                      type="text"
                      value={newSheetTitle}
                      onChange={(e) => setNewSheetTitle(e.target.value)}
                      className="w-full min-h-[40px] px-3 rounded-xl border text-sm font-bold outline-none"
                      style={{
                        background: 'var(--card-bg)',
                        borderColor: 'var(--border)',
                        color: 'var(--text)',
                      }}
                    />
                    <button
                      type="button"
                      disabled={creatingSheet || Boolean(selectingSheetId)}
                      onClick={handleConfirmCreateSheet}
                      className="w-full py-2.5 rounded-xl border text-xs font-extrabold cursor-pointer"
                      style={{
                        background: 'var(--card-bg)',
                        borderColor: 'var(--primary-accent)',
                        color: 'var(--primary-accent)',
                      }}
                    >
                      {creatingSheet ? 'جاري الإنشاء...' : `إنشاء شيت جديد باسم "${newSheetTitle.trim() || expectedSheetName}"`}
                    </button>
                  </div>
                </details>

                <div className="flex justify-end">
                  <button
                    type="button"
                    disabled={creatingSheet || Boolean(selectingSheetId)}
                    onClick={() => setShowCreateConfirm(false)}
                    className="px-4 py-2 rounded-xl border text-xs font-bold cursor-pointer"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--border)',
                      color: 'var(--text)',
                    }}
                  >
                    إغلاق
                  </button>
                </div>
              </>
            ) : (
              <>
                <div className="flex items-start gap-3">
                  <div
                    className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0"
                    style={{ background: 'var(--income-bg)', color: 'var(--income)' }}
                  >
                    <FilePlus2 className="w-5 h-5" />
                  </div>
                  <div>
                    <h4 className="text-base font-extrabold">إنشاء شيت {expectedSheetName} جديد</h4>
                    <p className="text-xs mt-1 leading-relaxed" style={{ color: 'var(--text-muted)' }}>
                      لم يُعثر على شيت بهذا الاسم في حسابك. سيتم إنشاء شيت جديد مهيأ بالكامل بـ 4 أوراق ومعادلات الأرصدة والتصنيفات المصلحة.
                    </p>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold mb-1" style={{ color: 'var(--text-muted)' }}>
                    اسم ملف الشيت الجديد في Google Drive
                  </label>
                  <input
                    type="text"
                    value={newSheetTitle}
                    onChange={(e) => setNewSheetTitle(e.target.value)}
                    className="w-full min-h-[42px] px-3 rounded-xl border text-sm font-bold outline-none"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--border)',
                      color: 'var(--text)',
                    }}
                  />
                </div>

                <div className="flex items-center justify-end gap-2">
                  <button
                    type="button"
                    disabled={creatingSheet}
                    onClick={() => setShowCreateConfirm(false)}
                    className="px-4 py-2.5 rounded-xl border text-xs font-bold cursor-pointer"
                    style={{
                      background: 'var(--surface-soft)',
                      borderColor: 'var(--border)',
                      color: 'var(--text)',
                    }}
                  >
                    إلغاء
                  </button>
                  <button
                    type="button"
                    disabled={creatingSheet}
                    onClick={handleConfirmCreateSheet}
                    className="px-4 py-2.5 rounded-xl text-xs font-extrabold cursor-pointer disabled:opacity-50"
                    style={{
                      background: 'var(--accent-gradient)',
                      color: 'var(--accent-contrast)',
                    }}
                  >
                    {creatingSheet ? 'جاري الإنشاء والتنسيق...' : 'تأكيد إنشاء الشيت الجديد'}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* Confirmation dialog required before mutating/formatting an existing Google Sheet */}
      {showFormatConfirm && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/55 backdrop-blur-sm"
          onClick={(e) => {
            if (e.target === e.currentTarget && !formatting) setShowFormatConfirm(false);
          }}
        >
          <div
            className="w-full max-w-lg rounded-2xl p-5 border"
            style={{
              background: 'var(--card-bg)',
              borderColor: 'var(--border)',
              color: 'var(--text)',
              boxShadow: 'var(--shadow-md)',
            }}
          >
            <div className="flex items-start gap-3 mb-3">
              <div
                className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0"
                style={{ background: 'var(--warning-bg)', color: 'var(--warning)' }}
              >
                <ShieldAlert className="w-5 h-5" />
              </div>
              <div>
                <h4 className="text-base font-extrabold">تأكيد إصلاح المعادلات وتنسيق الشيت</h4>
                <p className="text-xs mt-1 leading-relaxed" style={{ color: 'var(--text-muted)' }}>
                  سيتم تطبيق التعديلات التالية على الملف{' '}
                  <span className="font-bold" style={{ color: 'var(--text)' }}>
                    {inspectionReport?.spreadsheetTitle || cleanSheetId}
                  </span>
                  :
                </p>
              </div>
            </div>

            <ul className="text-xs space-y-1.5 pr-4 list-disc mb-5" style={{ color: 'var(--text-muted)' }}>
              <li>تنظيف تصنيفات الصرف والقبض من التكرار والمسافات الزائدة وإصلاح معادلة E2 بدون #REF!.</li>
              <li>تحديث معادلات SUMIFS للأرصدة (لتشمل قبض، دخل، وتحويل*) وصف الإجمالي.</li>
              <li>ضبط اتجاه الأوراق (RTL) وتمديد القوائم المنسدلة إلى 500 تصنيف وحساب.</li>
              <li>تطبيق التنسيق الشرطي بالألوان وتنسيق الأرقام والتواريخ.</li>
            </ul>

            <div className="flex items-center justify-end gap-2">
              <button
                type="button"
                disabled={formatting}
                onClick={() => setShowFormatConfirm(false)}
                className="px-4 py-2.5 rounded-xl border text-xs font-bold cursor-pointer"
                style={{
                  background: 'var(--surface-soft)',
                  borderColor: 'var(--border)',
                  color: 'var(--text)',
                }}
              >
                إلغاء
              </button>
              <button
                type="button"
                disabled={formatting}
                onClick={handleConfirmFormat}
                className="px-4 py-2.5 rounded-xl text-xs font-extrabold cursor-pointer disabled:opacity-50"
                style={{
                  background: 'var(--accent-gradient)',
                  color: 'var(--accent-contrast)',
                }}
              >
                {formatting ? 'جاري الإصلاح والتنسيق...' : 'تأكيد وتنفيذ الإصلاح'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

