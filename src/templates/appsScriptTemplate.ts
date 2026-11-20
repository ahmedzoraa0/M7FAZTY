export function generateAppsScriptCode(token = ''): string {
  return `// ============================================================================
// M7FAZTY - Google Apps Script المحدّث (v11)
// يدعم:
// 1. القراءة والكتابة الديناميكية للأرصدة والمحافظ (عدد غير محدود من الحسابات)
// 2. إضافة محفظة / بنك / خزنة جديدة من التطبيق مباشرة مع تحديث معادلات SUMIFS والقوائم
// 3. إضافة تصنيفات صرف وقبض جديدة مع إصلاح معادلة E2 ومنع #REF! والتكرار
// 4. قائمة داخل الشيت لتنسيق وتظبيط الجداول والقوائم بضغطة واحدة (setupAndFormatWorkbook)
// ============================================================================

const TOKEN = ${JSON.stringify(token)};
const TX_SHEET = "المعاملات اليومية";
const SETTINGS_SHEET = "القوائم والإعدادات";
const DEBTS_SHEET = "المستحقات";
const REMINDERS_SHEET = "التنبيهات والالتزامات";
const LOG_SHEET = "__API_Log";
const MAX_ROWS = 600;
const TYPES = ["صرف", "قبض", "تحويل بين حسابات"];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💰 أدوات M7FAZTY")
    .addItem("تنسيق وتظبيط الشيت والقوائم بالكامل", "setupAndFormatWorkbook")
    .addItem("تحديث معادلات الأرصدة وقوائم التحقق", "refreshFormulasAndValidationsMenu")
    .addToUi();
}

function out_(obj, cb) {
  const json = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

const num_ = v => Number(String(v || "").replace(/٫/g, ".").replace(/[^0-9.-]/g, "")) || 0;
const fp_ = (type, amount, desc) => [String(type).trim(), num_(amount), String(desc).trim()].join("|");
const ACCOUNT_KINDS = ["محفظة", "بنك", "خزنة", "بطاقة", "أخرى"];

function inferAccountKind_(name) {
  const n = String(name || "").trim().toLowerCase();
  if (!n) return "محفظة";
  if (/كاش|خزنة|نقد|جيب|cash|safe/.test(n) && !/فودافون|اتصالات|اورنج|وي|vodafone|etisalat|orange/.test(n)) return "خزنة";
  if (/بنك|المشرق|الأهلي|الاهلي|مصر|ادخار|توفير|qnb|cib|hsbc|nbe|bank|mashreq/.test(n)) return "بنك";
  if (/فيزا|ميزة|بطاقة|كارت|ايزي باي|إيزي باي|كريدت|ماستر|visa|master|card|meeza|easypay/.test(n)) return "بطاقة";
  if (/محفظة|فودافون|اتصالات|اورنج|أورنج|وي باي|انستا|إنستا|telda|instapay|wallet|pay/.test(n)) return "محفظة";
  if (/أخرى|اخرى|other/.test(n)) return "أخرى";
  return "محفظة";
}

// اكتشاف صفوف الحسابات وصف الإجمالي وصف بداية المعاملات ديناميكياً
function detectTxLayout_(sheet) {
  const scanRows = Math.min(Math.max(sheet.getLastRow(), 15), 50);
  const topVals = sheet.getRange(1, 1, scanRows, 2).getValues();
  let totalRow = -1;
  let txHeaderRow = -1;
  const accountRows = [];

  for (let i = 1; i < topVals.length; i++) {
    const colA = String(topVals[i][0] || "").trim();
    const colB = String(topVals[i][1] || "").trim();

    if (colA === "التاريخ" && colB === "النوع") {
      txHeaderRow = i + 1;
      break;
    }
    if (colA === "الإجمالي") {
      totalRow = i + 1;
      continue;
    }
    if (totalRow === -1 && colA && colA !== "الحساب" && colA !== "المحفظة" && colA !== "الحساب / المحفظة") {
      accountRows.push(i + 1);
    }
  }

  if (txHeaderRow === -1) txHeaderRow = (totalRow > 0 ? totalRow + 1 : 10);
  const firstTxRow = txHeaderRow + 1;
  if (totalRow === -1) {
    totalRow = accountRows.length ? accountRows[accountRows.length - 1] + 1 : 9;
  }

  return {
    accountRows: accountRows,
    totalRow: totalRow,
    txHeaderRow: txHeaderRow,
    firstTxRow: firstTxRow
  };
}

// ==========================================
// 1. نقطة الدخول للقراءة والكتابة عبر JSONP (doGet)
// ==========================================
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback;
  if (TOKEN && p.token !== TOKEN) return out_({ status: "error", message: "unauthorized" }, cb);

  if (p.action === "ping") {
    return out_({ status: "success", message: "Apps Script متصل ويعمل بنجاح ✅" }, cb);
  }

  if (p.action || (p.amount && p.type)) {
    return handleWrite_(p, cb);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TX_SHEET);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  const debtsSheet = ss.getSheetByName(DEBTS_SHEET);
  const remindersSheet = ss.getSheetByName(REMINDERS_SHEET);

  // أ. قراءة القوائم والأنواع من القوائم والإعدادات
  const expenseCats = [], incomeCats = [], payMethods = [], kindMap = {};
  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    const s = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < s.length; i++) {
      if (s[i][1]) expenseCats.push(String(s[i][1]).trim());
      if (s[i][2]) incomeCats.push(String(s[i][2]).trim());
      if (s[i][3]) {
        const accName = String(s[i][3]).trim();
        payMethods.push(accName);
        if (s[i][5]) kindMap[accName] = String(s[i][5]).trim();
      }
    }
  }

  // ب. قراءة أرصدة الحسابات ديناميكياً (بدون حد 7 صفوف)
  const layout = detectTxLayout_(sheet);
  const accounts = [];
  if (layout.accountRows.length > 0) {
    const startAcc = layout.accountRows[0];
    const endAcc = layout.accountRows[layout.accountRows.length - 1];
    const accVals = sheet.getRange(startAcc, 1, endAcc - startAcc + 1, 5).getValues();
    accVals.forEach((r, idx) => {
      const name = String(r[0]).trim();
      if (!name || name === "الإجمالي") return;
      accounts.push({
        row: startAcc + idx,
        name: name,
        kind: kindMap[name] || inferAccountKind_(name),
        opening: num_(r[1]),
        incoming: num_(r[2]),
        outgoing: num_(r[3]),
        current: num_(r[4])
      });
    });
  }

  // ج. قراءة آخر المعاملات المالية
  let recent = [];
  const lastRow = sheet.getLastRow();
  if (lastRow >= layout.firstTxRow) {
    const startRow = Math.max(layout.firstTxRow, lastRow - MAX_ROWS + 1);
    const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();
    recent = values.map((r, i) => ({
      row: startRow + i,
      date: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[0]).trim(),
      type: String(r[1]).trim(),
      category: String(r[2]).trim(),
      desc: String(r[3]).trim(),
      fromAcc: String(r[4]).trim(),
      toAcc: String(r[5]).trim(),
      amount: num_(r[6]),
      notes: String(r[7]).trim()
    })).filter(t => t.amount > 0 && t.type !== "النوع").reverse();
  }

  // د. قراءة المستحقات
  let debts = [];
  if (debtsSheet && debtsSheet.getLastRow() >= 2) {
    const dValues = debtsSheet.getRange(2, 1, debtsSheet.getLastRow() - 1, 7).getValues();
    debts = dValues.map((r, i) => ({
      row: i + 2,
      id: String(r[0]),
      date: r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[1]).trim(),
      type: String(r[2]).trim(),
      desc: String(r[3]).trim(),
      amount: num_(r[4]),
      status: String(r[5]).trim() || "مستحق",
      notes: String(r[6]).trim()
    })).filter(d => d.desc || d.amount > 0).reverse();
  }

  // هـ. قراءة التنبيهات والالتزامات
  let reminders = [];
  if (remindersSheet && remindersSheet.getLastRow() >= 2) {
    const remValues = remindersSheet.getRange(2, 1, remindersSheet.getLastRow() - 1, 7).getValues();
    reminders = remValues.map((r, i) => ({
      row: i + 2,
      id: String(r[0]),
      dueDate: r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[1]).trim(),
      desc: String(r[2]).trim(),
      amount: num_(r[3]),
      repeat: String(r[4]).trim() || "شهري",
      status: String(r[5]).trim() || "قادم",
      notes: String(r[6]).trim()
    })).filter(rem => rem.desc || rem.amount > 0);
  }

  const finalPayMethods = payMethods.length ? payMethods : accounts.map(a => a.name);

  return out_({
    status: "success",
    spreadsheetId: ss.getId(),
    spreadsheetTitle: ss.getName(),
    accounts: accounts,
    categories: { expense: [...new Set(expenseCats)], income: [...new Set(incomeCats)] },
    paymentMethods: [...new Set(finalPayMethods)],
    recent: recent,
    debts: debts,
    reminders: reminders
  }, cb);
}

// ==========================================
// 2. نقطة الدخول للكتابة عبر POST (doPost)
// ==========================================
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (TOKEN && p.token !== TOKEN) return out_({ status: "error", message: "unauthorized" });
  return handleWrite_(p, null);
}

// ==========================================
// 3. معالج عمليات الكتابة والتعديل والحذف الموحد
// ==========================================
function handleWrite_(p, cb) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(10000);
    locked = true;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reqId = String(p.requestId || "").trim();

    let logSheet = ss.getSheetByName(LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET);
      logSheet.appendRow(["Request ID", "Created At", "Response JSON"]);
      logSheet.hideSheet();
    }
    if (reqId && logSheet.getLastRow() >= 2) {
      const logRows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 3).getValues();
      for (let i = logRows.length - 1; i >= 0; i--) {
        if (String(logRows[i][0]).trim() === reqId) {
          const prev = JSON.parse(String(logRows[i][2]));
          return out_(prev, cb);
        }
      }
    }

    const action = p.action || "add";
    let responseObj = { status: "success" };

    // 1. إضافة حساب / محفظة / بنك / خزنة جديدة وربطها بالشيت والمعادلات والقوائم
    if (action === "add_account") {
      const accName = String(p.accountName || "").trim();
      const accKind = String(p.accountKind || "محفظة").trim();
      const opening = num_(p.openingBalance);
      if (!accName) throw new Error("اسم المحفظة أو الحساب فارغ");

      const sheet = ss.getSheetByName(TX_SHEET);
      const sSheet = ss.getSheetByName(SETTINGS_SHEET);
      const layout = detectTxLayout_(sheet);

      // التحقق من عدم التكرار
      if (layout.accountRows.length > 0) {
        const existingNames = sheet.getRange(layout.accountRows[0], 1, layout.accountRows.length, 1)
          .getValues().map(r => String(r[0]).trim());
        if (existingNames.indexOf(accName) !== -1) {
          throw new Error("الحساب موجود بالفعل في الشيت");
        }
      }

      // إدراج صف جديد قبل صف "الإجمالي"
      sheet.insertRowBefore(layout.totalRow);
      const newAccRow = layout.totalRow;
      const newTotalRow = layout.totalRow + 1;
      const newFirstTxRow = layout.firstTxRow + 1;

      sheet.getRange(newAccRow, 1, 1, 2).setValues([[accName, opening]]);
      sheet.getRange(newAccRow, 2, 1, 4).setNumberFormat("#,##0.00");

      // تحديث معادلات كل الحسابات وصف الإجمالي بحيث تشير إلى newFirstTxRow
      refreshSummaryFormulas_(sheet, 2, newTotalRow - 1, newTotalRow, newFirstTxRow);

      // إضافة الحساب إلى ورقة القوائم والإعدادات (العمود D لطرق الدفع، والعمود F لنوع الحساب)
      const maxS = Math.max(sSheet.getLastRow(), 2);
      const dVals = sSheet.getRange(2, 4, maxS - 1, 1).getValues().map(r => String(r[0]).trim());
      if (dVals.indexOf(accName) === -1) {
        let targetRow = 2;
        for (let i = 0; i < dVals.length; i++) {
          if (dVals[i] !== "") targetRow = i + 3;
          else { targetRow = i + 2; break; }
        }
        sSheet.getRange(targetRow, 4).setValue(accName);
        sSheet.getRange(targetRow, 6).setValue(accKind);
      }

      applyTxValidations_(sheet, sSheet, newFirstTxRow);
      responseObj = { status: "success", account: accName, kind: accKind, opening: opening };
    }

    // 1.5. تعديل نوع حساب (محفظة / بنك / خزنة / بطاقة / أخرى) في عمود F من ورقة القوائم والإعدادات بأمان
    else if (action === "update_account_kind") {
      const accName = String(p.accountName || "").trim();
      const accKind = String(p.accountKind || "").trim();
      if (!accName) throw new Error("اسم الحساب فارغ");
      if (ACCOUNT_KINDS.indexOf(accKind) === -1) throw new Error("نوع الحساب غير صحيح");

      const sSheet = ss.getSheetByName(SETTINGS_SHEET);
      const maxS = Math.max(sSheet.getLastRow(), 2);
      const dVals = sSheet.getRange(2, 4, maxS - 1, 1).getValues().map(r => String(r[0] || "").trim());
      const existingIdx = dVals.indexOf(accName);
      if (existingIdx !== -1) {
        sSheet.getRange(existingIdx + 2, 6).setValue(accKind);
      } else {
        let targetRow = 2;
        for (let i = 0; i < dVals.length; i++) {
          if (dVals[i] !== "") targetRow = i + 3;
          else { targetRow = i + 2; break; }
        }
        sSheet.getRange(targetRow, 4).setValue(accName);
        sSheet.getRange(targetRow, 6).setValue(accKind);
      }
      responseObj = { status: "success", account: accName, kind: accKind };
    }

    // 2. إضافة تصنيف جديد في ورقة "القوائم والإعدادات"
    else if (action === "add_category") {
      const catName = String(p.categoryName || "").trim();
      const catType = String(p.catType || "صرف").trim();
      if (!catName) throw new Error("اسم التصنيف فارغ");

      const sSheet = ss.getSheetByName(SETTINGS_SHEET);
      const txSheet = ss.getSheetByName(TX_SHEET);
      const colIdx = catType === "قبض" ? 3 : 2;
      const maxR = Math.max(sSheet.getLastRow(), 2);
      const colVals = sSheet.getRange(2, colIdx, maxR - 1, 1).getValues().map(r => String(r[0]).trim());

      if (colVals.indexOf(catName) === -1) {
        let targetRow = 2;
        for (let i = 0; i < colVals.length; i++) {
          if (colVals[i] !== "") targetRow = i + 3;
          else { targetRow = i + 2; break; }
        }
        sSheet.getRange(targetRow, colIdx).setValue(catName);
      }

      const e2Formula = String(sSheet.getRange("E2").getFormula() || "");
      const e2Display = String(sSheet.getRange("E2").getDisplayValue() || "");
      if (!e2Formula || e2Formula.indexOf("B2:B500") === -1 || e2Display.indexOf("#") === 0) {
        sSheet.getRange("E2:E500").clearContent();
        sSheet.getRange("E2").setFormula('=IFERROR(UNIQUE(FILTER({B2:B500;C2:C500}, {B2:B500;C2:C500}<>"")), "")');
      }

      const layout = detectTxLayout_(txSheet);
      applyTxValidations_(txSheet, sSheet, layout.firstTxRow);
      responseObj = { status: "success", category: catName, catType: catType };
    }

    // 3. تظبيط وتنسيق الشيت بالكامل عن بُعد
    else if (action === "setup_sheet") {
      const desiredName = String(p.sheetName || "").trim();
      const curName = String(ss.getName() || "").trim();
      if (desiredName && (p.forceRename === "1" || curName.indexOf("Untitled") !== -1 || curName.indexOf("بلا عنوان") !== -1)) {
        ss.rename(desiredName);
      }
      setupAndFormatWorkbook();
      responseObj = {
        status: "success",
        spreadsheetId: ss.getId(),
        spreadsheetTitle: ss.getName(),
        message: "تم تنسيق الشيت والقوائم بنجاح"
      };
    }

    // 4. إجراءات المستحقات
    else if (action === "add_debt") {
      const dSheet = ss.getSheetByName(DEBTS_SHEET);
      const newId = p.id || ("D-" + Date.now());
      dSheet.appendRow([newId, p.date, p.type, p.desc, num_(p.amount), p.status || "مستحق", p.notes || ""]);
      responseObj = { status: "success", id: newId };
    }
    else if (action === "update_debt_status") {
      const dSheet = ss.getSheetByName(DEBTS_SHEET);
      const targetRow = findRowByIdOrIndex_(dSheet, p.id, p.row);
      if (targetRow >= 2) dSheet.getRange(targetRow, 6).setValue(p.status);
    }
    else if (action === "delete_debt") {
      const dSheet = ss.getSheetByName(DEBTS_SHEET);
      const targetRow = findRowByIdOrIndex_(dSheet, p.id, p.row);
      if (targetRow >= 2) dSheet.deleteRow(targetRow);
    }

    // 5. إجراءات التنبيهات والالتزامات
    else if (action === "add_reminder") {
      const remSheet = ss.getSheetByName(REMINDERS_SHEET);
      const newId = p.id || ("R-" + Date.now());
      remSheet.appendRow([newId, p.dueDate, p.desc, num_(p.amount), p.repeat || "شهري", p.status || "قادم", p.notes || ""]);
      responseObj = { status: "success", id: newId };
    }
    else if (action === "update_reminder_status") {
      const remSheet = ss.getSheetByName(REMINDERS_SHEET);
      const targetRow = findRowByIdOrIndex_(remSheet, p.id, p.row);
      if (targetRow >= 2) remSheet.getRange(targetRow, 6).setValue(p.status);
    }
    else if (action === "delete_reminder") {
      const remSheet = ss.getSheetByName(REMINDERS_SHEET);
      const targetRow = findRowByIdOrIndex_(remSheet, p.id, p.row);
      if (targetRow >= 2) remSheet.deleteRow(targetRow);
    }

    // 6. إجراءات المعاملات اليومية
    else {
      const sheet = ss.getSheetByName(TX_SHEET);
      const layout = detectTxLayout_(sheet);
      const rowValues = () => {
        if (TYPES.indexOf(p.type) === -1) throw new Error("نوع الحركة غير صحيح");
        if (!(Number(p.amount) > 0)) throw new Error("المبلغ يجب أن يكون أكبر من صفر");
        return [p.date, p.type, p.category, p.description, p.fromAccount, p.toAccount || "", Number(p.amount)];
      };

      if (action === "add") {
        sheet.appendRow(rowValues().concat([p.notes || ""]));
      } else if (action === "update" || action === "delete") {
        const row = Number(p.row);
        if (!(row >= layout.firstTxRow) || row > sheet.getLastRow()) throw new Error("رقم الصف غير صحيح");
        const cur = sheet.getRange(row, 1, 1, 8).getValues()[0];
        if (p.fp && fp_(cur[1], cur[6], cur[3]) !== p.fp) {
          throw new Error("الصف تغير في الشيت، حدّث التطبيق وحاول مرة أخرى");
        }
        if (action === "delete") sheet.deleteRow(row);
        else sheet.getRange(row, 1, 1, 8).setValues([rowValues().concat([cur[7]])]);
      } else {
        throw new Error("أمر غير معروف: " + action);
      }
    }

    if (reqId) {
      logSheet.appendRow([reqId, new Date(), JSON.stringify(responseObj)]);
    }
    return out_(responseObj, cb);

  } catch (err) {
    return out_({ status: "error", message: err.toString() }, cb);
  } finally {
    if (locked) lock.releaseLock();
  }
}

function refreshSummaryFormulas_(sheet, firstAccRow, lastAccRow, totalRow, firstTxRow) {
  if (lastAccRow >= firstAccRow) {
    const formulas = [];
    for (let r = firstAccRow; r <= lastAccRow; r++) {
      formulas.push([
        \`=SUMIFS($G$\${firstTxRow}:$G,$B$\${firstTxRow}:$B,"قبض",$E$\${firstTxRow}:$E,A\${r})+SUMIFS($G$\${firstTxRow}:$G,$B$\${firstTxRow}:$B,"دخل",$E$\${firstTxRow}:$E,A\${r})+SUMIFS($G$\${firstTxRow}:$G,$B$\${firstTxRow}:$B,"تحويل*",$F$\${firstTxRow}:$F,A\${r})\`,
        \`=SUMIFS($G$\${firstTxRow}:$G,$B$\${firstTxRow}:$B,"صرف",$E$\${firstTxRow}:$E,A\${r})+SUMIFS($G$\${firstTxRow}:$G,$B$\${firstTxRow}:$B,"تحويل*",$E$\${firstTxRow}:$E,A\${r})\`,
        \`=B\${r}+C\${r}-D\${r}\`
      ]);
    }
    sheet.getRange(firstAccRow, 3, lastAccRow - firstAccRow + 1, 3).setFormulas(formulas);

    sheet.getRange(totalRow, 1).setValue("الإجمالي");
    sheet.getRange(totalRow, 2, 1, 4).setFormulas([[
      \`=SUM(B\${firstAccRow}:B\${lastAccRow})\`,
      \`=SUM(C\${firstAccRow}:C\${lastAccRow})\`,
      \`=SUM(D\${firstAccRow}:D\${lastAccRow})\`,
      \`=SUM(E\${firstAccRow}:E\${lastAccRow})\`
    ]]);
  }
}

function applyTxValidations_(txSheet, settingsSheet, firstTxRow) {
  const maxTxRows = Math.max(txSheet.getMaxRows() - firstTxRow + 1, 200);
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).setAllowInvalid(true).build();
  txSheet.getRange(firstTxRow, 2, maxTxRows, 1).setDataValidation(typeRule);

  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingsSheet.getRange("E2:E500"), true)
    .setAllowInvalid(true)
    .build();
  txSheet.getRange(firstTxRow, 3, maxTxRows, 1).setDataValidation(catRule);

  const accRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingsSheet.getRange("D2:D500"), true)
    .setAllowInvalid(true)
    .build();
  txSheet.getRange(firstTxRow, 5, maxTxRows, 2).setDataValidation(accRule);

  const kindRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ACCOUNT_KINDS, true)
    .setAllowInvalid(true)
    .build();
  settingsSheet.getRange("F2:F500").setDataValidation(kindRule);
}

function refreshFormulasAndValidationsMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName(TX_SHEET);
  const sSheet = ss.getSheetByName(SETTINGS_SHEET);
  const layout = detectTxLayout_(txSheet);
  if (layout.accountRows.length > 0) {
    refreshSummaryFormulas_(
      txSheet,
      layout.accountRows[0],
      layout.accountRows[layout.accountRows.length - 1],
      layout.totalRow,
      layout.firstTxRow
    );
  }
  applyTxValidations_(txSheet, sSheet, layout.firstTxRow);
  SpreadsheetApp.getUi().alert("تم تحديث المعادلات والقوائم المنسدلة بنجاح ✅");
}

// ==========================================
// 4. أداة تظبيط وتنسيق ملف الإكسيل بالكامل (وتأسيس شيت جديد من الصفر)
// ==========================================
function setupAndFormatWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getOrCreate = (name) => ss.getSheetByName(name) || ss.insertSheet(name);

  const txSheet = getOrCreate(TX_SHEET);
  const sSheet = getOrCreate(SETTINGS_SHEET);
  const dSheet = getOrCreate(DEBTS_SHEET);
  const rSheet = getOrCreate(REMINDERS_SHEET);

  // إزالة الورقة الافتراضية الفارغة إن وُجدت في شيت جديد
  ["Sheet1", "ورقة1", "ورقة 1"].forEach(defName => {
    const sh = ss.getSheetByName(defName);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });

  [txSheet, sSheet, dSheet, rSheet].forEach(sh => sh.setRightToLeft(true));

  let layout = detectTxLayout_(txSheet);

  // إذا كان الشيت جديداً تماماً بدون حسابات، نُنشئ حسابات مبدئية متكاملة
  if (layout.accountRows.length === 0) {
    txSheet.getRange("A1:E1").setValues([["الحساب / المحفظة", "الرصيد الافتتاحي", "إجمالي الداخل", "إجمالي الخارج", "الرصيد الحالي"]]);
    txSheet.getRange("A2:B4").setValues([
      ["كاش", 0],
      ["محفظة إلكترونية", 0],
      ["حساب بنكي", 0]
    ]);
    txSheet.getRange("A5").setValue("الإجمالي");
    txSheet.getRange("A7:H7").setValues([["التاريخ", "النوع", "التصنيف", "البيان / الوصف", "من حساب", "إلى حساب", "المبلغ", "ملاحظات"]]);
    layout = detectTxLayout_(txSheet);
  }

  // 1. ترويسات ورقة المعاملات اليومية
  txSheet.getRange("A1:E1")
    .setValues([["الحساب / المحفظة", "الرصيد الافتتاحي", "إجمالي الداخل", "إجمالي الخارج", "الرصيد الحالي"]])
    .setBackground("#1c1711").setFontColor("#f4ede0").setFontWeight("bold").setHorizontalAlignment("center");

  txSheet.getRange(layout.txHeaderRow, 1, 1, 8)
    .setValues([["التاريخ", "النوع", "التصنيف", "البيان / الوصف", "من حساب", "إلى حساب", "المبلغ", "ملاحظات"]])
    .setBackground("#1c1711").setFontColor("#e0bb6a").setFontWeight("bold").setHorizontalAlignment("center");

  if (layout.accountRows.length > 0) {
    const firstA = layout.accountRows[0];
    const lastA = layout.accountRows[layout.accountRows.length - 1];
    refreshSummaryFormulas_(txSheet, firstA, lastA, layout.totalRow, layout.firstTxRow);
    txSheet.getRange(firstA, 1, lastA - firstA + 1, 5)
      .setBackground("#f9f5ec").setBorder(true, true, true, true, true, true, "#d6cbb5", SpreadsheetApp.BorderStyle.SOLID);
    txSheet.getRange(firstA, 2, lastA - firstA + 1, 4).setNumberFormat("#,##0.00");
  }

  txSheet.getRange(layout.totalRow, 1, 1, 5)
    .setBackground("#efe0c1").setFontWeight("bold").setNumberFormat("#,##0.00");

  // 2. ورقة القوائم والإعدادات (تنظيف التصنيفات وإصلاح معادلة E2 بدون #REF!)
  sSheet.getRange("A1:F1")
    .setValues([["النوع", "تصنيفات الصرف", "تصنيفات القبض", "طرق الدفع والحسابات", "كل التصنيفات", "نوع الحساب"]])
    .setBackground("#1c1711").setFontColor("#f4ede0").setFontWeight("bold").setHorizontalAlignment("center");
  sSheet.getRange("A2:A4").setValues([["صرف"], ["قبض"], ["تحويل بين حسابات"]]);

  const rawExpCats = sSheet.getRange("B2:B500").getValues().map(r => String(r[0] || "").trim()).filter(Boolean);
  const rawIncCats = sSheet.getRange("C2:C500").getValues().map(r => String(r[0] || "").trim()).filter(Boolean);
  const cleanExpCats = rawExpCats.length ? [...new Set(rawExpCats)] : [
    "مأكولات ومشروبات", "مواصلات وبنزين", "تسوق ومستلزمات", "فواتير شهرية", "صحة وأدوية", "ترفيه وخروجات", "أخرى"
  ];
  const cleanIncCats = rawIncCats.length ? [...new Set(rawIncCats)] : [
    "راتب شهري", "تصميم", "عمل إضافي", "أخرى"
  ];

  sSheet.getRange("B2:B500").clearContent();
  sSheet.getRange("C2:C500").clearContent();
  sSheet.getRange("E2:E500").clearContent();

  sSheet.getRange(2, 2, cleanExpCats.length, 1).setValues(cleanExpCats.map(c => [c]));
  sSheet.getRange(2, 3, cleanIncCats.length, 1).setValues(cleanIncCats.map(c => [c]));

  const existingDF = sSheet.getRange("D2:F500").getValues();
  const existingMethods = [];
  const existingKindMap = {};
  existingDF.forEach(r => {
    const nm = String(r[0] || "").trim();
    const kd = String(r[2] || "").trim();
    if (nm) {
      existingMethods.push(nm);
      if (kd) existingKindMap[nm] = kd;
    }
  });
  const txAccNames = layout.accountRows.length > 0
    ? txSheet.getRange(layout.accountRows[0], 1, layout.accountRows.length, 1).getValues().map(r => String(r[0] || "").trim()).filter(Boolean)
    : ["كاش", "محفظة إلكترونية", "حساب بنكي"];
  const mergedMethods = [...new Set(existingMethods.concat(txAccNames))];
  if (mergedMethods.length > 0) {
    sSheet.getRange(2, 4, mergedMethods.length, 1).setValues(mergedMethods.map(m => [m]));
    sSheet.getRange(2, 6, mergedMethods.length, 1).setValues(mergedMethods.map(m => [existingKindMap[m] || inferAccountKind_(m)]));
  }

  sSheet.getRange("E2").setFormula('=IFERROR(UNIQUE(FILTER({B2:B500;C2:C500}, {B2:B500;C2:C500}<>"")), "")');
  sSheet.setFrozenRows(1);

  // 3. ورقة المستحقات
  dSheet.getRange("A1:G1")
    .setValues([["المعرف", "التاريخ", "النوع", "الاسم / الجهة", "المبلغ", "الحالة", "ملاحظات"]])
    .setBackground("#1c1711").setFontColor("#f4ede0").setFontWeight("bold").setHorizontalAlignment("center");
  dSheet.setFrozenRows(1);
  dSheet.getRange("E2:E500").setNumberFormat("#,##0.00");

  // 4. ورقة التنبيهات والالتزامات
  rSheet.getRange("A1:G1")
    .setValues([["المعرف", "تاريخ الاستحقاق", "البيان / الالتزام", "المبلغ", "التكرار", "الحالة", "ملاحظات"]])
    .setBackground("#1c1711").setFontColor("#f4ede0").setFontWeight("bold").setHorizontalAlignment("center");
  rSheet.setFrozenRows(1);
  rSheet.getRange("D2:D500").setNumberFormat("#,##0.00");

  // 5. القوائم المنسدلة والتنسيق الشرطي
  applyTxValidations_(txSheet, sSheet, layout.firstTxRow);
  txSheet.getRange(layout.firstTxRow, 1, 500, 1).setNumberFormat("yyyy-MM-dd").setHorizontalAlignment("center");
  txSheet.getRange(layout.firstTxRow, 7, 500, 1).setNumberFormat("#,##0.00").setHorizontalAlignment("center");

  const typeCol = txSheet.getRange(layout.firstTxRow, 2, 500, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("قبض").setBackground("#e7f5ee").setFontColor("#0f8a5f").setBold(true).setRanges([typeCol]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("صرف").setBackground("#fbebec").setFontColor("#c43d4b").setBold(true).setRanges([typeCol]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("تحويل بين حسابات").setBackground("#efecfb").setFontColor("#6b5bd6").setBold(true).setRanges([typeCol]).build()
  ];
  txSheet.setConditionalFormatRules(rules);

  [txSheet, sSheet, dSheet, rSheet].forEach(sh => sh.autoResizeColumns(1, 8));
}

function findRowByIdOrIndex_(sheet, id, fallbackRow) {
  const cleanId = String(id || "").trim();
  if (cleanId && sheet.getLastRow() >= 2) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === cleanId) return i + 2;
    }
  }
  const r = Number(fallbackRow);
  return (r >= 2 && r <= sheet.getLastRow()) ? r : -1;
}
`;
}
