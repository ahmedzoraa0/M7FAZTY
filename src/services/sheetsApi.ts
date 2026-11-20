import {
  AccountItem,
  AccountKind,
  DebtItem,
  FinanceState,
  ReminderItem,
  SheetInspectionReport,
  TransactionItem,
} from '../types';
import { cleanNumber } from './appsScriptClient';

const BASE_URL = 'https://sheets.googleapis.com/v4/spreadsheets';
const DRIVE_FILES_URL = 'https://www.googleapis.com/drive/v3/files';

export const SHEET_NAMES = {
  TX: 'المعاملات اليومية',
  SETTINGS: 'القوائم والإعدادات',
  DEBTS: 'المستحقات',
  REMINDERS: 'التنبيهات والالتزامات',
} as const;

export function buildM7faztySheetName(profileName?: string): string {
  const clean = String(profileName || '').trim();
  return clean ? `M7FAZTY_${clean}` : 'M7FAZTY';
}

export interface DriveSheetMatch {
  id: string;
  name: string;
  modifiedTime?: string;
  webViewLink?: string;
  isExactTarget?: boolean;
}

export function extractSpreadsheetId(input: string): string {
  const trimmed = (input || '').trim();
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : trimmed;
}

export async function findExistingM7faztySheets(
  accessToken: string,
  profileName?: string
): Promise<{
  targetName: string;
  matches: DriveSheetMatch[];
  exactMatch: DriveSheetMatch | null;
}> {
  const targetName = buildM7faztySheetName(profileName);
  const query =
    "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false and (name contains 'M7FAZTY' or name contains 'm7fazty')";
  const params = new URLSearchParams({
    q: query,
    orderBy: 'modifiedTime desc',
    pageSize: '25',
    fields: 'files(id,name,modifiedTime,webViewLink)',
    supportsAllDrives: 'true',
    includeItemsFromAllDrives: 'true',
  });

  const data = await sheetsFetch<{
    files?: Array<{
      id: string;
      name: string;
      modifiedTime?: string;
      webViewLink?: string;
    }>;
  }>(`${DRIVE_FILES_URL}?${params.toString()}`, accessToken);

  const normalizedTarget = targetName.trim().toLowerCase();
  const matches: DriveSheetMatch[] = (data.files || []).map((f) => ({
    id: f.id,
    name: f.name,
    modifiedTime: f.modifiedTime,
    webViewLink: f.webViewLink || `https://docs.google.com/spreadsheets/d/${f.id}/edit`,
    isExactTarget: f.name.trim().toLowerCase() === normalizedTarget,
  }));

  matches.sort((a, b) => Number(Boolean(b.isExactTarget)) - Number(Boolean(a.isExactTarget)));
  const exactMatch = matches.find((m) => m.isExactTarget) || null;

  return { targetName, matches, exactMatch };
}

async function sheetsFetch<T>(
  url: string,
  accessToken: string,
  options: RequestInit = {}
): Promise<T> {
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const errMsg = data?.error?.message || `HTTP ${res.status}`;
    const err = new Error(errMsg) as Error & { status?: number };
    err.status = res.status;
    throw err;
  }
  return data as T;
}

interface SpreadsheetMeta {
  spreadsheetId: string;
  spreadsheetUrl: string;
  properties?: { title?: string };
  sheets?: Array<{
    properties: {
      sheetId: number;
      title: string;
      index: number;
      gridProperties?: { rowCount?: number; columnCount?: number; frozenRowCount?: number };
    };
  }>;
}

interface ValueRangeResponse {
  range: string;
  values?: unknown[][];
}

interface BatchGetResponse {
  valueRanges?: ValueRangeResponse[];
}

function findSummaryLayout(txRows: unknown[][]) {
  let totalRow = -1;
  let txHeaderRow = -1;
  const accounts: AccountItem[] = [];

  for (let i = 1; i < Math.min(txRows.length, 40); i++) {
    const row = txRows[i] || [];
    const colA = String(row[0] ?? '').trim();
    const colB = String(row[1] ?? '').trim();

    if (colA === 'التاريخ' && colB === 'النوع') {
      txHeaderRow = i + 1; // 1-indexed
      break;
    }

    if (colA === 'الإجمالي') {
      totalRow = i + 1;
      continue;
    }

    if (totalRow === -1 && colA && colA !== 'الحساب' && colA !== 'المحفظة') {
      accounts.push({
        row: i + 1,
        name: colA,
        opening: cleanNumber(row[1]),
        incoming: cleanNumber(row[2]),
        outgoing: cleanNumber(row[3]),
        current: cleanNumber(row[4]),
      });
    }
  }

  const firstTxRow = txHeaderRow > 0 ? txHeaderRow + 1 : 11;
  if (totalRow === -1) {
    totalRow = accounts.length > 0 ? accounts[accounts.length - 1].row! + 1 : 9;
  }

  return { accounts, totalRow, txHeaderRow, firstTxRow };
}

export async function inspectSpreadsheet(
  spreadsheetId: string,
  accessToken: string
): Promise<{ report: SheetInspectionReport; state: FinanceState }> {
  const id = extractSpreadsheetId(spreadsheetId);
  const meta = await sheetsFetch<SpreadsheetMeta>(
    `${BASE_URL}/${encodeURIComponent(id)}?fields=spreadsheetId,spreadsheetUrl,properties.title,sheets.properties`,
    accessToken
  );

  const sheetNames = (meta.sheets || []).map((s) => s.properties.title);
  const required = Object.values(SHEET_NAMES);
  const missingSheets = required.filter((name) => !sheetNames.includes(name));

  const rangesToFetch: string[] = [];
  if (sheetNames.includes(SHEET_NAMES.TX)) {
    rangesToFetch.push(`'${SHEET_NAMES.TX}'!A1:H1000`);
  }
  if (sheetNames.includes(SHEET_NAMES.SETTINGS)) {
    rangesToFetch.push(`'${SHEET_NAMES.SETTINGS}'!A1:F500`);
  }
  if (sheetNames.includes(SHEET_NAMES.DEBTS)) {
    rangesToFetch.push(`'${SHEET_NAMES.DEBTS}'!A1:G300`);
  }
  if (sheetNames.includes(SHEET_NAMES.REMINDERS)) {
    rangesToFetch.push(`'${SHEET_NAMES.REMINDERS}'!A1:G300`);
  }

  let txRows: unknown[][] = [];
  let settingsRows: unknown[][] = [];
  let debtsRows: unknown[][] = [];
  let remindersRows: unknown[][] = [];

  if (rangesToFetch.length > 0) {
    const query = rangesToFetch.map((r) => `ranges=${encodeURIComponent(r)}`).join('&');
    const batchRes = await sheetsFetch<BatchGetResponse>(
      `${BASE_URL}/${encodeURIComponent(id)}/values:batchGet?${query}`,
      accessToken
    );

    const mapRange = (targetSheet: string) => {
      const found = (batchRes.valueRanges || []).find((vr) =>
        vr.range.includes(targetSheet)
      );
      return found?.values || [];
    };

    txRows = mapRange(SHEET_NAMES.TX);
    settingsRows = mapRange(SHEET_NAMES.SETTINGS);
    debtsRows = mapRange(SHEET_NAMES.DEBTS);
    remindersRows = mapRange(SHEET_NAMES.REMINDERS);
  }

  const layout = findSummaryLayout(txRows);

  // Parse settings (categories & payment methods & account types)
  const expenseCats: string[] = [];
  const incomeCats: string[] = [];
  const payMethods: string[] = [];
  const kindMap: Record<string, AccountKind> = {};

  for (let i = 1; i < settingsRows.length; i++) {
    const r = settingsRows[i] || [];
    const exp = String(r[1] ?? '').trim();
    const inc = String(r[2] ?? '').trim();
    const pay = String(r[3] ?? '').trim();
    const kind = String(r[5] ?? '').trim() as AccountKind;
    if (exp) expenseCats.push(exp);
    if (inc) incomeCats.push(inc);
    if (pay) {
      payMethods.push(pay);
      if (kind) kindMap[pay] = kind;
    }
  }

  const accounts: AccountItem[] = layout.accounts.map((a) => ({
    ...a,
    kind: kindMap[a.name] || a.kind || inferAccountKind(a.name),
  }));

  // Parse transactions
  const recent: TransactionItem[] = [];
  for (let i = layout.firstTxRow - 1; i < txRows.length; i++) {
    const r = txRows[i] || [];
    const amount = cleanNumber(r[6]);
    const type = String(r[1] ?? '').trim();
    if (amount > 0 && type && type !== 'النوع') {
      recent.push({
        row: i + 1,
        date: String(r[0] ?? '').trim(),
        type,
        category: String(r[2] ?? '').trim(),
        desc: String(r[3] ?? '').trim(),
        fromAcc: String(r[4] ?? '').trim(),
        toAcc: String(r[5] ?? '').trim(),
        amount,
        notes: String(r[7] ?? '').trim(),
      });
    }
  }
  recent.reverse();

  // Parse debts
  const debts: DebtItem[] = [];
  for (let i = 1; i < debtsRows.length; i++) {
    const r = debtsRows[i] || [];
    const desc = String(r[3] ?? '').trim();
    const amount = cleanNumber(r[4]);
    if (desc || amount > 0) {
      debts.push({
        row: i + 1,
        id: String(r[0] ?? ''),
        date: String(r[1] ?? '').trim(),
        type: String(r[2] ?? '').trim(),
        desc,
        amount,
        status: String(r[5] ?? '').trim() || 'مستحق',
        notes: String(r[6] ?? '').trim(),
      });
    }
  }
  debts.reverse();

  // Parse reminders
  const reminders: ReminderItem[] = [];
  for (let i = 1; i < remindersRows.length; i++) {
    const r = remindersRows[i] || [];
    const desc = String(r[2] ?? '').trim();
    const amount = cleanNumber(r[3]);
    if (desc || amount > 0) {
      reminders.push({
        row: i + 1,
        id: String(r[0] ?? ''),
        dueDate: String(r[1] ?? '').trim(),
        desc,
        amount,
        repeat: String(r[4] ?? '').trim() || 'شهري',
        status: String(r[5] ?? '').trim() || 'قادم',
        notes: String(r[6] ?? '').trim(),
      });
    }
  }

  const unique = (arr: string[]) => Array.from(new Set(arr));

  const state: FinanceState = {
    accounts,
    categories: {
      expense: unique(expenseCats),
      income: unique(incomeCats),
    },
    paymentMethods: unique(payMethods.length ? payMethods : accounts.map((a) => a.name)),
    recent,
    debts,
    reminders,
  };

  const report: SheetInspectionReport = {
    spreadsheetTitle: meta.properties?.title || 'جدول البيانات',
    spreadsheetUrl: meta.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${id}/edit`,
    sheetNames,
    missingSheets,
    accountCount: accounts.length,
    firstTxRow: layout.firstTxRow,
    totalRow: layout.totalRow,
    txCount: recent.length,
    expenseCatCount: state.categories.expense.length,
    incomeCatCount: state.categories.income.length,
    payMethodCount: state.paymentMethods.length,
    debtCount: debts.length,
    reminderCount: reminders.length,
  };

  return { report, state };
}

export async function addCategoryDirect(
  spreadsheetId: string,
  accessToken: string,
  catType: 'صرف' | 'قبض',
  categoryName: string
): Promise<void> {
  const id = extractSpreadsheetId(spreadsheetId);
  const cleanCat = categoryName.trim();
  if (!cleanCat) return;
  const colLetter = catType === 'قبض' ? 'C' : 'B';
  const res = await sheetsFetch<ValueRangeResponse>(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!${colLetter}2:${colLetter}500`)}`,
    accessToken
  );

  const vals = (res.values || []).map((r) => String(r[0] ?? '').trim());
  if (!vals.includes(cleanCat)) {
    let targetRow = 2 + vals.length;
    for (let i = 0; i < vals.length; i++) {
      if (!vals[i]) {
        targetRow = i + 2;
        break;
      }
    }

    await sheetsFetch(
      `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!${colLetter}${targetRow}`)}?valueInputOption=USER_ENTERED`,
      accessToken,
      {
        method: 'PUT',
        body: JSON.stringify({ values: [[cleanCat]] }),
      }
    );
  }

  // Ensure E2 dynamic categories formula is intact and free of #REF!
  const e2Check = await sheetsFetch<ValueRangeResponse>(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!E2`)}?valueRenderOption=FORMULA`,
    accessToken
  );
  const currentE2 = String(e2Check.values?.[0]?.[0] ?? '').trim();
  if (!currentE2.startsWith('=') || !currentE2.includes('B2:B500')) {
    await sheetsFetch(
      `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!E2:E500`)}:clear`,
      accessToken,
      { method: 'POST', body: JSON.stringify({}) }
    );
    await sheetsFetch(
      `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!E2`)}?valueInputOption=USER_ENTERED`,
      accessToken,
      {
        method: 'PUT',
        body: JSON.stringify({
          values: [['=IFERROR(UNIQUE(FILTER({B2:B500;C2:C500}, {B2:B500;C2:C500}<>"")), "")']],
        }),
      }
    );
  }
}

export async function addAccountDirect(
  spreadsheetId: string,
  accessToken: string,
  account: { name: string; kind: AccountKind; opening: number }
): Promise<void> {
  const id = extractSpreadsheetId(spreadsheetId);
  const meta = await sheetsFetch<SpreadsheetMeta>(
    `${BASE_URL}/${encodeURIComponent(id)}?fields=sheets.properties`,
    accessToken
  );
  const txSheetMeta = (meta.sheets || []).find((s) => s.properties.title === SHEET_NAMES.TX);
  if (!txSheetMeta) throw new Error(`لم يتم العثور على ورقة "${SHEET_NAMES.TX}"`);

  const txSheetId = txSheetMeta.properties.sheetId;

  const [txRes, settingsRes] = await Promise.all([
    sheetsFetch<ValueRangeResponse>(
      `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.TX}'!A1:H45`)}`,
      accessToken
    ),
    sheetsFetch<ValueRangeResponse>(
      `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!D2:F200`)}`,
      accessToken
    ),
  ]);

  const txRows = txRes.values || [];
  const layout = findSummaryLayout(txRows);

  if (layout.accounts.some((a) => a.name.trim() === account.name.trim())) {
    throw new Error('هذا الحساب موجود بالفعل في الشيت');
  }

  // Insert 1 row right before the totalRow (0-indexed: totalRow - 1)
  const insertAtZeroIndex = Math.max(1, layout.totalRow - 1);
  const newAccRow = insertAtZeroIndex + 1; // 1-indexed
  const newTotalRow = layout.totalRow + 1;
  const newFirstTxRow = layout.firstTxRow + 1;

  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}:batchUpdate`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({
        requests: [
          {
            insertDimension: {
              range: {
                sheetId: txSheetId,
                dimension: 'ROWS',
                startIndex: insertAtZeroIndex,
                endIndex: insertAtZeroIndex + 1,
              },
              inheritFromBefore: true,
            },
          },
        ],
      }),
    }
  );

  // Build updated summary formulas for all account rows + new row + total row
  const allAccCount = layout.accounts.length + 1;
  const summaryValues: unknown[][] = [];

  for (let idx = 0; idx < allAccCount; idx++) {
    const rNum = 2 + idx;
    const isNew = idx === allAccCount - 1;
    const accName = isNew ? account.name.trim() : layout.accounts[idx].name;
    const openingVal = isNew ? account.opening : layout.accounts[idx].opening ?? 0;

    summaryValues.push([
      accName,
      openingVal,
      `=SUMIFS($G$${newFirstTxRow}:$G,$B$${newFirstTxRow}:$B,"قبض",$E$${newFirstTxRow}:$E,A${rNum})+SUMIFS($G$${newFirstTxRow}:$G,$B$${newFirstTxRow}:$B,"دخل",$E$${newFirstTxRow}:$E,A${rNum})+SUMIFS($G$${newFirstTxRow}:$G,$B$${newFirstTxRow}:$B,"تحويل*",$F$${newFirstTxRow}:$F,A${rNum})`,
      `=SUMIFS($G$${newFirstTxRow}:$G,$B$${newFirstTxRow}:$B,"صرف",$E$${newFirstTxRow}:$E,A${rNum})+SUMIFS($G$${newFirstTxRow}:$G,$B$${newFirstTxRow}:$B,"تحويل*",$E$${newFirstTxRow}:$E,A${rNum})`,
      `=B${rNum}+C${rNum}-D${rNum}`,
    ]);
  }

  const lastAccRow = 1 + allAccCount;
  summaryValues.push([
    'الإجمالي',
    `=SUM(B2:B${lastAccRow})`,
    `=SUM(C2:C${lastAccRow})`,
    `=SUM(D2:D${lastAccRow})`,
    `=SUM(E2:E${lastAccRow})`,
  ]);

  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.TX}'!A2:E${newTotalRow}`)}?valueInputOption=USER_ENTERED`,
    accessToken,
    {
      method: 'PUT',
      body: JSON.stringify({ values: summaryValues }),
    }
  );

  // Add to القوائم والإعدادات D (paymentMethods) and F (account kind)
  const settingsVals = settingsRes.values || [];
  const existingMethods = settingsVals.map((r) => String(r[0] ?? '').trim());
  if (!existingMethods.includes(account.name.trim())) {
    let targetSettingsRow = 2 + existingMethods.length;
    for (let i = 0; i < existingMethods.length; i++) {
      if (!existingMethods[i]) {
        targetSettingsRow = i + 2;
        break;
      }
    }
    await sheetsFetch(
      `${BASE_URL}/${encodeURIComponent(id)}/values:batchUpdate`,
      accessToken,
      {
        method: 'POST',
        body: JSON.stringify({
          valueInputOption: 'USER_ENTERED',
          data: [
            {
              range: `'${SHEET_NAMES.SETTINGS}'!D${targetSettingsRow}`,
              values: [[account.name.trim()]],
            },
            {
              range: `'${SHEET_NAMES.SETTINGS}'!F${targetSettingsRow}`,
              values: [[account.kind]],
            },
          ],
        }),
      }
    );
  }

  // Ensure data validation on E and F points to Settings D2:D500
  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}:batchUpdate`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({
        requests: [
          {
            setDataValidation: {
              range: {
                sheetId: txSheetId,
                startRowIndex: newFirstTxRow - 1,
                endRowIndex: 1000,
                startColumnIndex: 4,
                endColumnIndex: 6,
              },
              rule: {
                condition: {
                  type: 'ONE_OF_RANGE',
                  values: [{ userEnteredValue: `='${SHEET_NAMES.SETTINGS}'!$D$2:$D$500` }],
                },
                showCustomUi: true,
                strict: false,
              },
            },
          },
          {
            repeatCell: {
              range: {
                sheetId: txSheetId,
                startRowIndex: newAccRow - 1,
                endRowIndex: newAccRow,
                startColumnIndex: 1,
                endColumnIndex: 5,
              },
              cell: {
                userEnteredFormat: {
                  numberFormat: { type: 'NUMBER', pattern: '#,##0.00' },
                },
              },
              fields: 'userEnteredFormat.numberFormat',
            },
          },
        ],
      }),
    }
  );
}

export function inferAccountKind(name: string): AccountKind {
  const n = String(name || '').trim().toLowerCase();
  if (!n) return 'محفظة';
  if (/كاش|خزنة|نقد|جيب|cash|safe/.test(n) && !/فودافون|اتصالات|اورنج|وي|vodafone|etisalat|orange/.test(n)) {
    return 'خزنة';
  }
  if (/بنك|المشرق|الأهلي|الاهلي|مصر|ادخار|توفير|qnb|cib|hsbc|nbe|bank|mashreq/.test(n)) {
    return 'بنك';
  }
  if (/فيزا|ميزة|بطاقة|كارت|ايزي باي|إيزي باي|كريدت|ماستر|visa|master|card|meeza|easypay/.test(n)) {
    return 'بطاقة';
  }
  if (/محفظة|فودافون|اتصالات|اورنج|أورنج|وي باي|انستا|إنستا|telda|instapay|wallet|pay/.test(n)) {
    return 'محفظة';
  }
  if (/أخرى|اخرى|other/.test(n)) {
    return 'أخرى';
  }
  return 'محفظة';
}

export async function updateAccountKindDirect(
  spreadsheetId: string,
  accessToken: string,
  accountName: string,
  newKind: AccountKind
): Promise<void> {
  const id = extractSpreadsheetId(spreadsheetId);
  const cleanName = accountName.trim();
  if (!cleanName) throw new Error('اسم الحساب غير صالح');

  const settingsRes = await sheetsFetch<ValueRangeResponse>(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!D2:F500`)}`,
    accessToken
  );
  const rows = settingsRes.values || [];
  let targetRow = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i]?.[0] ?? '').trim() === cleanName) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = 2 + rows.length;
    for (let i = 0; i < rows.length; i++) {
      if (!String(rows[i]?.[0] ?? '').trim()) {
        targetRow = i + 2;
        break;
      }
    }
    await sheetsFetch(
      `${BASE_URL}/${encodeURIComponent(id)}/values:batchUpdate`,
      accessToken,
      {
        method: 'POST',
        body: JSON.stringify({
          valueInputOption: 'USER_ENTERED',
          data: [
            { range: `'${SHEET_NAMES.SETTINGS}'!D${targetRow}`, values: [[cleanName]] },
            { range: `'${SHEET_NAMES.SETTINGS}'!F${targetRow}`, values: [[newKind]] },
          ],
        }),
      }
    );
  } else {
    await sheetsFetch(
      `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!F${targetRow}`)}?valueInputOption=USER_ENTERED`,
      accessToken,
      {
        method: 'PUT',
        body: JSON.stringify({ values: [[newKind]] }),
      }
    );
  }
}

export async function updateDebtStatusDirect(
  spreadsheetId: string,
  accessToken: string,
  identifier: { id?: string; row?: number },
  status: string
): Promise<void> {
  const id = extractSpreadsheetId(spreadsheetId);
  const res = await sheetsFetch<ValueRangeResponse>(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.DEBTS}'!A2:A500`)}`,
    accessToken
  );
  let targetRow = -1;
  const cleanId = String(identifier.id || '').trim();
  const rows = res.values || [];
  if (cleanId) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i]?.[0] ?? '').trim() === cleanId) {
        targetRow = i + 2;
        break;
      }
    }
  }
  if (targetRow === -1 && identifier.row && identifier.row >= 2) {
    targetRow = identifier.row;
  }
  if (targetRow < 2) return;
  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.DEBTS}'!F${targetRow}`)}?valueInputOption=USER_ENTERED`,
    accessToken,
    {
      method: 'PUT',
      body: JSON.stringify({ values: [[status]] }),
    }
  );
}

export async function updateReminderStatusDirect(
  spreadsheetId: string,
  accessToken: string,
  identifier: { id?: string; row?: number },
  status: string
): Promise<void> {
  const id = extractSpreadsheetId(spreadsheetId);
  const res = await sheetsFetch<ValueRangeResponse>(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.REMINDERS}'!A2:A500`)}`,
    accessToken
  );
  let targetRow = -1;
  const cleanId = String(identifier.id || '').trim();
  const rows = res.values || [];
  if (cleanId) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i]?.[0] ?? '').trim() === cleanId) {
        targetRow = i + 2;
        break;
      }
    }
  }
  if (targetRow === -1 && identifier.row && identifier.row >= 2) {
    targetRow = identifier.row;
  }
  if (targetRow < 2) return;
  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.REMINDERS}'!F${targetRow}`)}?valueInputOption=USER_ENTERED`,
    accessToken,
    {
      method: 'PUT',
      body: JSON.stringify({ values: [[status]] }),
    }
  );
}

export async function addTransactionDirect(
  spreadsheetId: string,
  accessToken: string,
  tx: TransactionItem
): Promise<void> {
  const id = extractSpreadsheetId(spreadsheetId);
  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.TX}'!A:H`)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({
        values: [[tx.date, tx.type, tx.category, tx.desc, tx.fromAcc, tx.toAcc || '', tx.amount, tx.notes || '']],
      }),
    }
  );
}

export async function formatAndConfigureWorkbook(
  spreadsheetId: string,
  accessToken: string
): Promise<SheetInspectionReport> {
  const id = extractSpreadsheetId(spreadsheetId);
  const meta = await sheetsFetch<SpreadsheetMeta>(
    `${BASE_URL}/${encodeURIComponent(id)}?fields=spreadsheetId,spreadsheetUrl,properties.title,sheets.properties`,
    accessToken
  );

  const existingTitles = (meta.sheets || []).map((s) => s.properties.title);
  const missingTitles = Object.values(SHEET_NAMES).filter((n) => !existingTitles.includes(n));

  if (missingTitles.length > 0) {
    await sheetsFetch(`${BASE_URL}/${encodeURIComponent(id)}:batchUpdate`, accessToken, {
      method: 'POST',
      body: JSON.stringify({
        requests: missingTitles.map((title) => ({
          addSheet: {
            properties: {
              title,
              rightToLeft: true,
            },
          },
        })),
      }),
    });
  }

  const refreshedMeta = await sheetsFetch<SpreadsheetMeta>(
    `${BASE_URL}/${encodeURIComponent(id)}?fields=spreadsheetId,spreadsheetUrl,properties.title,sheets.properties`,
    accessToken
  );
  const sheetIdByName: Record<string, number> = {};
  (refreshedMeta.sheets || []).forEach((s) => {
    sheetIdByName[s.properties.title] = s.properties.sheetId;
  });

  // Read TX summary & Settings sheet columns B/C/D/F to clean categories and build accurate formulas
  const [txRes, settingsRes] = await Promise.all([
    sheetsFetch<ValueRangeResponse>(
      `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.TX}'!A1:H50`)}`,
      accessToken
    ),
    sheetsFetch<ValueRangeResponse>(
      `${BASE_URL}/${encodeURIComponent(id)}/values/${encodeURIComponent(`'${SHEET_NAMES.SETTINGS}'!A1:F500`)}`,
      accessToken
    ),
  ]);
  const layout = findSummaryLayout(txRes.values || []);

  // Normalize/trim & deduplicate expense (col B) and income (col C) categories to fix broken/duplicate categories
  const sRows = settingsRes.values || [];
  const rawExp: string[] = [];
  const rawInc: string[] = [];
  const rawMethods: string[] = [];
  for (let i = 1; i < sRows.length; i++) {
    const r = sRows[i] || [];
    const exp = String(r[1] ?? '').trim();
    const inc = String(r[2] ?? '').trim();
    const pm = String(r[3] ?? '').trim();
    if (exp) rawExp.push(exp);
    if (inc) rawInc.push(inc);
    if (pm) rawMethods.push(pm);
  }

  const dedupeTrimmed = (arr: string[]) => Array.from(new Set(arr.map((x) => x.trim()).filter(Boolean)));
  const cleanExp = dedupeTrimmed(rawExp);
  const cleanInc = dedupeTrimmed(rawInc);

  const finalExp =
    cleanExp.length > 0
      ? cleanExp
      : ['مأكولات ومشروبات', 'مواصلات وبنزين', 'تسوق ومستلزمات', 'فواتير شهرية', 'صحة وأدوية', 'ترفيه وخروجات', 'أخرى'];
  const finalInc =
    cleanInc.length > 0 ? cleanInc : ['راتب شهري', 'تصميم', 'عمل إضافي', 'أخرى'];

  // Clear B2:C500 & E2:E500 first to eliminate leftover cells that cause #REF! array expansion failures
  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}/values:batchClear`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({
        ranges: [
          `'${SHEET_NAMES.SETTINGS}'!B2:B500`,
          `'${SHEET_NAMES.SETTINGS}'!C2:C500`,
          `'${SHEET_NAMES.SETTINGS}'!E2:E500`,
        ],
      }),
    }
  );

  // Ensure headers exist on all 4 sheets
  const valueUpdates: Array<{ range: string; values: unknown[][] }> = [
    {
      range: `'${SHEET_NAMES.TX}'!A1:E1`,
      values: [['الحساب / المحفظة', 'الرصيد الافتتاحي', 'إجمالي الداخل', 'إجمالي الخارج', 'الرصيد الحالي']],
    },
    {
      range: `'${SHEET_NAMES.TX}'!A${layout.txHeaderRow > 0 ? layout.txHeaderRow : layout.firstTxRow - 1}:H${layout.txHeaderRow > 0 ? layout.txHeaderRow : layout.firstTxRow - 1}`,
      values: [['التاريخ', 'النوع', 'التصنيف', 'البيان / الوصف', 'من حساب', 'إلى حساب', 'المبلغ', 'ملاحظات']],
    },
    {
      range: `'${SHEET_NAMES.SETTINGS}'!A1:F1`,
      values: [['النوع', 'تصنيفات الصرف', 'تصنيفات القبض', 'طرق الدفع والحسابات', 'كل التصنيفات', 'نوع الحساب']],
    },
    {
      range: `'${SHEET_NAMES.SETTINGS}'!A2:A4`,
      values: [['صرف'], ['قبض'], ['تحويل بين حسابات']],
    },
    {
      range: `'${SHEET_NAMES.SETTINGS}'!B2:B${1 + finalExp.length}`,
      values: finalExp.map((c) => [c]),
    },
    {
      range: `'${SHEET_NAMES.SETTINGS}'!C2:C${1 + finalInc.length}`,
      values: finalInc.map((c) => [c]),
    },
    {
      range: `'${SHEET_NAMES.SETTINGS}'!E2`,
      values: [['=IFERROR(UNIQUE(FILTER({B2:B500;C2:C500}, {B2:B500;C2:C500}<>"")), "")']],
    },
    {
      range: `'${SHEET_NAMES.DEBTS}'!A1:G1`,
      values: [['المعرف', 'التاريخ', 'النوع', 'الاسم / الجهة', 'المبلغ', 'الحالة', 'ملاحظات']],
    },
    {
      range: `'${SHEET_NAMES.REMINDERS}'!A1:G1`,
      values: [['المعرف', 'تاريخ الاستحقاق', 'البيان / الالتزام', 'المبلغ', 'التكرار', 'الحالة', 'ملاحظات']],
    },
  ];

  // Sync payment methods (col D) and non-destructively fill missing account kinds (col F)
  const existingKindByMethod: Record<string, string> = {};
  for (let i = 1; i < sRows.length; i++) {
    const r = sRows[i] || [];
    const pm = String(r[3] ?? '').trim();
    const kd = String(r[5] ?? '').trim();
    if (pm && kd) existingKindByMethod[pm] = kd;
  }
  const mergedMethods = dedupeTrimmed([...rawMethods, ...layout.accounts.map((a) => a.name)]);
  if (mergedMethods.length > 0) {
    valueUpdates.push({
      range: `'${SHEET_NAMES.SETTINGS}'!D2:D${1 + mergedMethods.length}`,
      values: mergedMethods.map((m) => [m]),
    });
    valueUpdates.push({
      range: `'${SHEET_NAMES.SETTINGS}'!F2:F${1 + mergedMethods.length}`,
      values: mergedMethods.map((m) => [existingKindByMethod[m] || inferAccountKind(m)]),
    });
  }

  // Refresh dynamic SUMIFS formulas on summary rows
  if (layout.accounts.length > 0) {
    const ftx = layout.firstTxRow;
    const summaryRowsValues = layout.accounts.map((acc, idx) => {
      const r = 2 + idx;
      return [
        acc.name,
        acc.opening ?? 0,
        `=SUMIFS($G$${ftx}:$G,$B$${ftx}:$B,"قبض",$E$${ftx}:$E,A${r})+SUMIFS($G$${ftx}:$G,$B$${ftx}:$B,"دخل",$E$${ftx}:$E,A${r})+SUMIFS($G$${ftx}:$G,$B$${ftx}:$B,"تحويل*",$F$${ftx}:$F,A${r})`,
        `=SUMIFS($G$${ftx}:$G,$B$${ftx}:$B,"صرف",$E$${ftx}:$E,A${r})+SUMIFS($G$${ftx}:$G,$B$${ftx}:$B,"تحويل*",$E$${ftx}:$E,A${r})`,
        `=B${r}+C${r}-D${r}`,
      ];
    });
    const lastAccRow = 1 + layout.accounts.length;
    summaryRowsValues.push([
      'الإجمالي',
      `=SUM(B2:B${lastAccRow})`,
      `=SUM(C2:C${lastAccRow})`,
      `=SUM(D2:D${lastAccRow})`,
      `=SUM(E2:E${lastAccRow})`,
    ]);
    valueUpdates.push({
      range: `'${SHEET_NAMES.TX}'!A2:E${layout.totalRow}`,
      values: summaryRowsValues,
    });
  }

  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}/values:batchUpdate`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: valueUpdates,
      }),
    }
  );

  // Visual formatting + RTL + Data Validation + Conditional Formatting
  const txId = sheetIdByName[SHEET_NAMES.TX];
  const setId = sheetIdByName[SHEET_NAMES.SETTINGS];
  const debtId = sheetIdByName[SHEET_NAMES.DEBTS];
  const remId = sheetIdByName[SHEET_NAMES.REMINDERS];

  const headerBg = { red: 0.11, green: 0.09, blue: 0.07 };
  const goldAccentText = { red: 0.94, green: 0.87, blue: 0.68 };
  const softRowBg = { red: 0.98, green: 0.96, blue: 0.92 };
  const totalRowBg = { red: 0.95, green: 0.91, blue: 0.82 };
  const borderColor = { red: 0.85, green: 0.81, blue: 0.74 };

  const headerRow0 = (layout.txHeaderRow > 0 ? layout.txHeaderRow : layout.firstTxRow - 1) - 1;

  const requests: unknown[] = [
    // Set RTL & font family on all 4 sheets
    ...[txId, setId, debtId, remId].map((sid) => ({
      updateSheetProperties: {
        properties: { sheetId: sid, rightToLeft: true },
        fields: 'rightToLeft',
      },
    })),
    // Freeze top rows on Settings, Debts, Reminders
    ...[setId, debtId, remId].map((sid) => ({
      updateSheetProperties: {
        properties: { sheetId: sid, gridProperties: { frozenRowCount: 1 } },
        fields: 'gridProperties.frozenRowCount',
      },
    })),
    // Style Headers on Settings, Debts, Reminders + TX Summary Header + TX Table Header
    ...[
      { sid: txId, r0: 0, c1: 5 },
      { sid: txId, r0: headerRow0, c1: 8 },
      { sid: setId, r0: 0, c1: 6 },
      { sid: debtId, r0: 0, c1: 7 },
      { sid: remId, r0: 0, c1: 7 },
    ].map(({ sid, r0, c1 }) => ({
      repeatCell: {
        range: { sheetId: sid, startRowIndex: r0, endRowIndex: r0 + 1, startColumnIndex: 0, endColumnIndex: c1 },
        cell: {
          userEnteredFormat: {
            backgroundColor: headerBg,
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE',
            textFormat: { foregroundColor: goldAccentText, bold: true, fontSize: 11, fontFamily: 'Tajawal' },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)',
      },
    })),
    // Style Total Row in TX sheet
    {
      repeatCell: {
        range: {
          sheetId: txId,
          startRowIndex: layout.totalRow - 1,
          endRowIndex: layout.totalRow,
          startColumnIndex: 0,
          endColumnIndex: 5,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: totalRowBg,
            textFormat: { bold: true, fontSize: 11, fontFamily: 'Tajawal' },
            numberFormat: { type: 'NUMBER', pattern: '#,##0.00' },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)',
      },
    },
    // Style account summary body rows
    {
      repeatCell: {
        range: {
          sheetId: txId,
          startRowIndex: 1,
          endRowIndex: Math.max(2, layout.totalRow - 1),
          startColumnIndex: 1,
          endColumnIndex: 5,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: softRowBg,
            numberFormat: { type: 'NUMBER', pattern: '#,##0.00' },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,numberFormat,horizontalAlignment)',
      },
    },
    // Borders on summary table
    {
      updateBorders: {
        range: {
          sheetId: txId,
          startRowIndex: 0,
          endRowIndex: layout.totalRow,
          startColumnIndex: 0,
          endColumnIndex: 5,
        },
        top: { style: 'SOLID', width: 1, color: borderColor },
        bottom: { style: 'SOLID', width: 1, color: borderColor },
        left: { style: 'SOLID', width: 1, color: borderColor },
        right: { style: 'SOLID', width: 1, color: borderColor },
        innerHorizontal: { style: 'SOLID', width: 1, color: borderColor },
        innerVertical: { style: 'SOLID', width: 1, color: borderColor },
      },
    },
    // Number & Date formatting on TX rows
    {
      repeatCell: {
        range: { sheetId: txId, startRowIndex: headerRow0 + 1, endRowIndex: 1000, startColumnIndex: 0, endColumnIndex: 1 },
        cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-MM-dd' }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
      },
    },
    {
      repeatCell: {
        range: { sheetId: txId, startRowIndex: headerRow0 + 1, endRowIndex: 1000, startColumnIndex: 6, endColumnIndex: 7 },
        cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0.00' }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
      },
    },
    // Data Validations on TX Sheet: Type, Category, FromAcc, ToAcc
    {
      setDataValidation: {
        range: { sheetId: txId, startRowIndex: headerRow0 + 1, endRowIndex: 1000, startColumnIndex: 1, endColumnIndex: 2 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'صرف' }, { userEnteredValue: 'قبض' }, { userEnteredValue: 'تحويل بين حسابات' }] },
          showCustomUi: true,
          strict: false,
        },
      },
    },
    {
      setDataValidation: {
        range: { sheetId: txId, startRowIndex: headerRow0 + 1, endRowIndex: 1000, startColumnIndex: 2, endColumnIndex: 3 },
        rule: {
          condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='${SHEET_NAMES.SETTINGS}'!$E$2:$E$500` }] },
          showCustomUi: true,
          strict: false,
        },
      },
    },
    {
      setDataValidation: {
        range: { sheetId: txId, startRowIndex: headerRow0 + 1, endRowIndex: 1000, startColumnIndex: 4, endColumnIndex: 6 },
        rule: {
          condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='${SHEET_NAMES.SETTINGS}'!$D$2:$D$500` }] },
          showCustomUi: true,
          strict: false,
        },
      },
    },
    // Data Validations on Debts & Reminders
    {
      setDataValidation: {
        range: { sheetId: debtId, startRowIndex: 1, endRowIndex: 500, startColumnIndex: 2, endColumnIndex: 3 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'ليا' }, { userEnteredValue: 'عليا' }] },
          showCustomUi: true,
          strict: false,
        },
      },
    },
    {
      setDataValidation: {
        range: { sheetId: debtId, startRowIndex: 1, endRowIndex: 500, startColumnIndex: 5, endColumnIndex: 6 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'مستحق' }, { userEnteredValue: 'تم السداد' }] },
          showCustomUi: true,
          strict: false,
        },
      },
    },
    {
      setDataValidation: {
        range: { sheetId: remId, startRowIndex: 1, endRowIndex: 500, startColumnIndex: 4, endColumnIndex: 5 },
        rule: {
          condition: {
            type: 'ONE_OF_LIST',
            values: [{ userEnteredValue: 'شهري' }, { userEnteredValue: 'لمرة واحدة' }, { userEnteredValue: 'أسبوعي' }, { userEnteredValue: 'سنوي' }],
          },
          showCustomUi: true,
          strict: false,
        },
      },
    },
    {
      setDataValidation: {
        range: { sheetId: remId, startRowIndex: 1, endRowIndex: 500, startColumnIndex: 5, endColumnIndex: 6 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'قادم' }, { userEnteredValue: 'تم الدفع' }] },
          showCustomUi: true,
          strict: false,
        },
      },
    },
    // Data Validation on Settings column F (Account Kind) so user can safely change it directly in Google Sheets
    {
      setDataValidation: {
        range: { sheetId: setId, startRowIndex: 1, endRowIndex: 500, startColumnIndex: 5, endColumnIndex: 6 },
        rule: {
          condition: {
            type: 'ONE_OF_LIST',
            values: [
              { userEnteredValue: 'محفظة' },
              { userEnteredValue: 'بنك' },
              { userEnteredValue: 'خزنة' },
              { userEnteredValue: 'بطاقة' },
              { userEnteredValue: 'أخرى' },
            ],
          },
          showCustomUi: true,
          strict: false,
        },
      },
    },
    // Conditional formatting on TX Type column
    {
      addConditionalFormatRule: {
        index: 0,
        rule: {
          ranges: [{ sheetId: txId, startRowIndex: headerRow0 + 1, endRowIndex: 1000, startColumnIndex: 1, endColumnIndex: 2 }],
          booleanRule: {
            condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'قبض' }] },
            format: { backgroundColor: { red: 0.86, green: 0.96, blue: 0.91 }, textFormat: { foregroundColor: { red: 0.06, green: 0.45, blue: 0.28 }, bold: true } },
          },
        },
      },
    },
    {
      addConditionalFormatRule: {
        index: 1,
        rule: {
          ranges: [{ sheetId: txId, startRowIndex: headerRow0 + 1, endRowIndex: 1000, startColumnIndex: 1, endColumnIndex: 2 }],
          booleanRule: {
            condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'صرف' }] },
            format: { backgroundColor: { red: 0.99, green: 0.91, blue: 0.92 }, textFormat: { foregroundColor: { red: 0.72, green: 0.16, blue: 0.22 }, bold: true } },
          },
        },
      },
    },
    {
      addConditionalFormatRule: {
        index: 2,
        rule: {
          ranges: [{ sheetId: txId, startRowIndex: headerRow0 + 1, endRowIndex: 1000, startColumnIndex: 1, endColumnIndex: 2 }],
          booleanRule: {
            condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'تحويل بين حسابات' }] },
            format: { backgroundColor: { red: 0.93, green: 0.92, blue: 0.99 }, textFormat: { foregroundColor: { red: 0.36, green: 0.28, blue: 0.78 }, bold: true } },
          },
        },
      },
    },
    // Auto-resize columns nicely
    ...[txId, setId, debtId, remId].map((sid) => ({
      autoResizeDimensions: {
        dimensions: { sheetId: sid, dimension: 'COLUMNS', startIndex: 0, endIndex: 8 },
      },
    })),
  ];

  await sheetsFetch(
    `${BASE_URL}/${encodeURIComponent(id)}:batchUpdate`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({ requests }),
    }
  );

  const finalCheck = await inspectSpreadsheet(id, accessToken);
  return finalCheck.report;
}

export async function createNewFinanceSpreadsheet(
  accessToken: string,
  options: {
    title: string;
    accounts?: Array<{ name: string; kind: AccountKind; opening: number }>;
  }
): Promise<{
  spreadsheetId: string;
  spreadsheetUrl: string;
  report: SheetInspectionReport;
}> {
  const starterAccounts: Array<{ name: string; kind: AccountKind; opening: number }> =
    options.accounts && options.accounts.length > 0
      ? options.accounts
      : [
          { name: 'كاش', kind: 'خزنة', opening: 0 },
          { name: 'محفظة إلكترونية', kind: 'محفظة', opening: 0 },
          { name: 'حساب بنكي', kind: 'بنك', opening: 0 },
        ];

  const created = await sheetsFetch<SpreadsheetMeta>(BASE_URL, accessToken, {
    method: 'POST',
    body: JSON.stringify({
      properties: {
        title: options.title.trim() || 'M7FAZTY',
        locale: 'ar_EG',
        timeZone: 'Africa/Cairo',
      },
      sheets: Object.values(SHEET_NAMES).map((title) => ({
        properties: {
          title,
          rightToLeft: true,
        },
      })),
    }),
  });

  const newId = created.spreadsheetId;
  const totalRow = 2 + starterAccounts.length;
  const txHeaderRow = totalRow + 2;

  const accountRows = starterAccounts.map((acc) => [acc.name.trim(), acc.opening || 0]);
  const settingsPaymentRows = starterAccounts.map((acc) => [
    acc.name.trim(),
    '',
    acc.kind || 'محفظة',
  ]);

  const expenseCats = [
    ['مأكولات ومشروبات'],
    ['مواصلات وبنزين'],
    ['تسوق ومستلزمات'],
    ['فواتير شهرية'],
    ['صحة وأدوية'],
    ['ترفيه وخروجات'],
    ['أخرى'],
  ];
  const incomeCats = [['راتب شهري'], ['تصميم'], ['عمل إضافي'], ['أخرى']];

  await sheetsFetch(`${BASE_URL}/${encodeURIComponent(newId)}/values:batchUpdate`, accessToken, {
    method: 'POST',
    body: JSON.stringify({
      valueInputOption: 'USER_ENTERED',
      data: [
        {
          range: `'${SHEET_NAMES.TX}'!A1:E1`,
          values: [['الحساب / المحفظة', 'الرصيد الافتتاحي', 'إجمالي الداخل', 'إجمالي الخارج', 'الرصيد الحالي']],
        },
        {
          range: `'${SHEET_NAMES.TX}'!A2:B${1 + starterAccounts.length}`,
          values: accountRows,
        },
        {
          range: `'${SHEET_NAMES.TX}'!A${totalRow}`,
          values: [['الإجمالي']],
        },
        {
          range: `'${SHEET_NAMES.TX}'!A${txHeaderRow}:H${txHeaderRow}`,
          values: [['التاريخ', 'النوع', 'التصنيف', 'البيان / الوصف', 'من حساب', 'إلى حساب', 'المبلغ', 'ملاحظات']],
        },
        {
          range: `'${SHEET_NAMES.SETTINGS}'!B2:B${1 + expenseCats.length}`,
          values: expenseCats,
        },
        {
          range: `'${SHEET_NAMES.SETTINGS}'!C2:C${1 + incomeCats.length}`,
          values: incomeCats,
        },
        {
          range: `'${SHEET_NAMES.SETTINGS}'!D2:F${1 + settingsPaymentRows.length}`,
          values: settingsPaymentRows,
        },
      ],
    }),
  });

  const report = await formatAndConfigureWorkbook(newId, accessToken);
  return {
    spreadsheetId: newId,
    spreadsheetUrl:
      created.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${newId}/edit`,
    report,
  };
}

