// ===== الإعدادات الأساسية =====
const TOKEN = "";
const TX_SHEET = "المعاملات اليومية";
const SETTINGS_SHEET = "القوائم والإعدادات";
const DEBTS_SHEET = "المستحقات";
const REMINDERS_SHEET = "التنبيهات والالتزامات";
const LOG_SHEET = "__API_Log";
const FIRST_TX_ROW = 11;
const MAX_ROWS = 500;
const TYPES = ["صرف", "قبض", "تحويل بين حسابات"];

function out_(obj, cb) {
  const json = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

const num_ = v => Number(String(v || "").replace(/[^0-9.-]/g, "")) || 0;
const fp_ = (type, amount, desc) => [String(type).trim(), num_(amount), String(desc).trim()].join("|");

// ==========================================
// 1. نقطة الدخول للقراءة والكتابة عبر JSONP (doGet)
// ==========================================
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback;
  if (TOKEN && p.token !== TOKEN) return out_({ status: "error", message: "unauthorized" }, cb);

  // إذا كان الطلب يحتوي على action أو بيانات حركة، نفّذ عملية الكتابة
  if (p.action || (p.amount && p.type)) {
    return handleWrite_(p, cb);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TX_SHEET);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  const debtsSheet = ss.getSheetByName(DEBTS_SHEET);
  const remindersSheet = ss.getSheetByName(REMINDERS_SHEET);

  // أ. قراءة أرصدة الحسابات
  const accounts = [];
  sheet.getRange("A2:E8").getValues().forEach(r => {
    const name = String(r[0]).trim();
    if (!name || name === "الإجمالي") return;
    accounts.push({ name: name, current: num_(r[4]) });
  });

  // ب. قراءة التصنيفات وطرق الدفع
  const expenseCats = [], incomeCats = [], payMethods = [];
  const s = settingsSheet.getDataRange().getValues();
  for (let i = 1; i < s.length; i++) {
    if (s[i][1]) expenseCats.push(String(s[i][1]).trim());
    if (s[i][2]) incomeCats.push(String(s[i][2]).trim());
    if (s[i][3]) payMethods.push(String(s[i][3]).trim());
  }

  // ج. قراءة آخر المعاملات المالية
  let recent = [];
  const lastRow = sheet.getLastRow();
  if (lastRow >= FIRST_TX_ROW) {
    const startRow = Math.max(FIRST_TX_ROW, lastRow - MAX_ROWS + 1);
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

  return out_({
    status: "success",
    accounts: accounts,
    categories: { expense: [...new Set(expenseCats)], income: [...new Set(incomeCats)] },
    paymentMethods: [...new Set(payMethods)],
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

    // التحقق من عدم تكرار نفس الطلب في __API_Log
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

    // 1. إضافة تصنيف جديد في ورقة "القوائم والإعدادات"
    if (action === "add_category") {
      const catName = String(p.categoryName || "").trim();
      const catType = String(p.catType || "صرف").trim();
      if (!catName) throw new Error("اسم التصنيف فارغ");

      const sSheet = ss.getSheetByName(SETTINGS_SHEET);
      const colIdx = catType === "قبض" ? 3 : 2; // العمود B للصرف (2) والعمود C للقبض (3)
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

      // تحديث عمود "كل التصنيفات" (E) إذا لم يكن معادلة تلقائية
      const e2Formula = sSheet.getRange("E2").getFormula();
      if (!e2Formula) {
        const allVals = sSheet.getRange(2, 5, Math.max(sSheet.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim());
        if (allVals.indexOf(catName) === -1) {
          let eTargetRow = 2;
          for (let i = 0; i < allVals.length; i++) {
            if (allVals[i] !== "") eTargetRow = i + 3;
            else { eTargetRow = i + 2; break; }
          }
          sSheet.getRange(eTargetRow, 5).setValue(catName);
        }
      }
      responseObj = { status: "success", category: catName, catType: catType };
    }

    // 2. إجراءات المستحقات
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

    // 3. إجراءات التنبيهات والالتزامات
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

    // 4. إجراءات المعاملات اليومية
    else {
      const sheet = ss.getSheetByName(TX_SHEET);
      const rowValues = () => {
        if (TYPES.indexOf(p.type) === -1) throw new Error("نوع الحركة غير صحيح");
        if (!(Number(p.amount) > 0)) throw new Error("المبلغ يجب أن يكون أكبر من صفر");
        return [p.date, p.type, p.category, p.description, p.fromAccount, p.toAccount || "", Number(p.amount)];
      };

      if (action === "add") {
        sheet.appendRow(rowValues().concat([p.notes || ""]));
      } else if (action === "update" || action === "delete") {
        const row = Number(p.row);
        if (!(row >= FIRST_TX_ROW) || row > sheet.getLastRow()) throw new Error("رقم الصف غير صحيح");
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