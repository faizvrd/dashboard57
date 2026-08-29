/**
 * ============================================================================
 *  ESR - AROHERA  |  Backend (Google Apps Script)
 * ============================================================================
 *  Cara pakai:
 *  1. Buka Google Sheet "ESR_Arohera_Database.xlsx" (import ke Google Sheets).
 *  2. Extensions > Apps Script, hapus isi default, tempel file ini sebagai Code.gs.
 *  3. Deploy > New deployment > Web app.
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  4. Copy URL /exec yang dihasilkan, tempel ke variabel API_BASE_URL / apiUrl
 *     project pada file index.html (lihat getProjectById / PROJECTS).
 *
 *  Sheet yang dipakai (nama harus sama persis):
 *    PO_Tracking, FAT_Schedule, Shipment_Tracking, Action_Log,
 *    User_Management, Risk_Register, Currency_Rates
 *
 *  Kolom "header" (baris 1) tidak wajib lengkap 100% — jika frontend
 *  mengirim field baru yang belum ada sebagai kolom, backend ini otomatis
 *  menambahkan kolom baru di sheet terkait (lihat getOrCreateColumnIndex_).
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// KONFIGURASI
// ---------------------------------------------------------------------------

// Kosongkan (biarkan "") jika script ini "container-bound" (dibuka lewat
// Extensions > Apps Script dari dalam Google Sheet). Isi dengan Spreadsheet
// ID jika script berdiri sendiri (standalone).
const SPREADSHEET_ID = "";

const KEY_COLUMNS = {
  PO_Tracking: "poTrackingId",
  FAT_Schedule: "FAT_ID",
  Shipment_Tracking: "Shipment_ID",
  Action_Log: "Action_ID",
  Risk_Register: "Risk_ID",
  User_Management: "User_ID"
};

const ID_PREFIX = {
  PO_Tracking: "PO",
  FAT_Schedule: "FAT",
  Shipment_Tracking: "SHP",
  Action_Log: "ACT",
  Risk_Register: "RSK",
  User_Management: "USR"
};

// Folder Google Drive (opsional) untuk menyimpan file upload (foto FAT,
// dokumen unpriced PO, dll). Kosongkan untuk simpan di root Drive akun.
const UPLOAD_FOLDER_ID = "";

// ---------------------------------------------------------------------------
// ENTRY POINTS
// ---------------------------------------------------------------------------

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action;
    if (!action) {
      return jsonOutput_({ success: false, error: "Missing 'action' parameter." });
    }
    const result = routeAction_(action, params);
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = body.action;
    const payload = body.payload || {};
    if (!action) {
      return jsonOutput_({ success: false, error: "Missing 'action' in body." });
    }
    const result = routeAction_(action, payload);
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err && err.message || err) });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// ROUTER
// ---------------------------------------------------------------------------

function routeAction_(action, p) {
  switch (action) {
    case "login":
      return doLogin_(p.username, p.password);
    case "getAllUsers":
      return getAll_("User_Management");
    case "getAllPO":
      return getAll_("PO_Tracking");
    case "getAllFAT":
      return getAll_("FAT_Schedule");
    case "getAllShipments":
      return getAll_("Shipment_Tracking");
    case "getAllActions":
      return getAll_("Action_Log");
    case "getAllMilestones":
      return getAllMilestones_();
    case "getDashboardData":
      return getDashboardData_(Number(p.upcomingDays) || 30);
    case "getCurrencyRates":
      return getCurrencyRates_();
    case "addRow":
      return addRow_(p.sheet, p.rowData || {});
    case "updateRow":
      return updateRow_(p.sheet, p.keyColumn, p.keyValue, p.newData || {});
    case "deleteRow":
      return deleteRow_(p.sheet, p.keyColumn, p.keyValue);
    case "closeAction":
      return closeAction_(p.actionId);
    case "uploadFileToDrive":
      return uploadFileToDrive_(p);
    default:
      return { success: false, error: "Unknown action: " + action };
  }
}

// ---------------------------------------------------------------------------
// SPREADSHEET HELPERS
// ---------------------------------------------------------------------------

function getSpreadsheet_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error("Sheet '" + name + "' tidak ditemukan.");
  return sheet;
}

function getHeaders_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

/** Ubah baris sheet menjadi array of object memakai header baris 1. */
function sheetToObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values
    .filter(row => row.some(cell => cell !== "" && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        if (!h) return;
        obj[h] = formatCell_(row[i]);
      });
      return obj;
    });
}

function formatCell_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
  }
  return v;
}

function getAll_(sheetName) {
  const sheet = getSheet_(sheetName);
  return { success: true, data: sheetToObjects_(sheet) };
}

/** Cari/pastikan kolom header ada; jika belum ada, tambahkan kolom baru. */
function getOrCreateColumnIndex_(sheet, headerName) {
  const headers = getHeaders_(sheet);
  let idx = headers.indexOf(headerName);
  if (idx === -1) {
    const newCol = Math.max(sheet.getLastColumn(), headers.length) + 1;
    sheet.getRange(1, newCol).setValue(headerName);
    idx = newCol - 1;
  }
  return idx; // 0-based
}

function generateId_(sheetName, sheet) {
  const prefix = ID_PREFIX[sheetName] || "ID";
  const keyCol = KEY_COLUMNS[sheetName];
  const existing = sheetToObjects_(sheet).map(r => String(r[keyCol] || ""));
  let n = existing.length + 1;
  let candidate;
  do {
    candidate = prefix + "-" + String(n).padStart(4, "0");
    n++;
  } while (existing.indexOf(candidate) !== -1);
  return candidate;
}

// ---------------------------------------------------------------------------
// GENERIC CRUD (dipakai oleh semua sheet: PO_Tracking, FAT_Schedule,
// Shipment_Tracking, Action_Log, User_Management, Risk_Register, dll)
// ---------------------------------------------------------------------------

function addRow_(sheetName, rowData) {
  const sheet = getSheet_(sheetName);
  const keyCol = KEY_COLUMNS[sheetName];

  if (keyCol && !rowData[keyCol]) {
    rowData[keyCol] = generateId_(sheetName, sheet);
  }
  if (sheetName === "Action_Log" && !rowData.Date) {
    rowData.Date = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
  }
  if (sheetName === "Action_Log" && !rowData.Status) {
    rowData.Status = "OPEN";
  }

  // pastikan semua kolom ada (auto-extend header jika perlu)
  Object.keys(rowData).forEach(k => getOrCreateColumnIndex_(sheet, k));

  const headers = getHeaders_(sheet);
  const newRow = headers.map(h => (h && rowData.hasOwnProperty(h)) ? rowData[h] : "");
  sheet.appendRow(newRow);

  return { success: true, message: "Data berhasil disimpan.", id: keyCol ? rowData[keyCol] : null };
}

function findRowIndexByKey_(sheet, keyColumn, keyValue) {
  const headers = getHeaders_(sheet);
  const colIdx = headers.indexOf(keyColumn);
  if (colIdx === -1) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, colIdx + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(keyValue)) return i + 2; // 1-based row number
  }
  return -1;
}

function updateRow_(sheetName, keyColumn, keyValue, newData) {
  const sheet = getSheet_(sheetName);
  const rowNum = findRowIndexByKey_(sheet, keyColumn, keyValue);
  if (rowNum === -1) return { success: false, error: "Data dengan " + keyColumn + " = " + keyValue + " tidak ditemukan." };

  Object.keys(newData).forEach(k => getOrCreateColumnIndex_(sheet, k));
  const headers = getHeaders_(sheet);

  headers.forEach((h, i) => {
    if (!h) return;
    if (newData.hasOwnProperty(h)) {
      // jangan menimpa password kosong saat edit user tanpa isi password baru
      if (sheetName === "User_Management" && h === "Password" && !newData[h]) return;
      sheet.getRange(rowNum, i + 1).setValue(newData[h]);
    }
  });

  return { success: true, message: "Data berhasil diupdate." };
}

function deleteRow_(sheetName, keyColumn, keyValue) {
  const sheet = getSheet_(sheetName);
  const rowNum = findRowIndexByKey_(sheet, keyColumn, keyValue);
  if (rowNum === -1) return { success: false, error: "Data tidak ditemukan." };
  sheet.deleteRow(rowNum);
  return { success: true, message: "Data berhasil dihapus." };
}

// ---------------------------------------------------------------------------
// ACTION LOG SPECIFIC
// ---------------------------------------------------------------------------

function closeAction_(actionId) {
  const sheet = getSheet_("Action_Log");
  const rowNum = findRowIndexByKey_(sheet, "Action_ID", actionId);
  if (rowNum === -1) return { success: false, error: "Action tidak ditemukan." };
  const headers = getHeaders_(sheet);
  const statusCol = getOrCreateColumnIndex_(sheet, "Status") + 1;
  const closedDateCol = getOrCreateColumnIndex_(sheet, "Closed_Date") + 1;
  sheet.getRange(rowNum, statusCol).setValue("CLOSED");
  sheet.getRange(rowNum, closedDateCol).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd"));
  return { success: true, message: "Action ditutup." };
}

// ---------------------------------------------------------------------------
// LOGIN
// ---------------------------------------------------------------------------

function doLogin_(username, password) {
  if (!username || !password) return { success: false, error: "Username dan password wajib diisi." };
  const sheet = getSheet_("User_Management");
  const users = sheetToObjects_(sheet);
  const user = users.find(u =>
    String(u.Username || "").toLowerCase() === String(username).toLowerCase() &&
    String(u.Password || "") === String(password)
  );
  if (!user) return { success: false, error: "Username atau password salah." };
  if (String(user.Status || "").toLowerCase() === "inactive") {
    return { success: false, error: "Akun tidak aktif. Hubungi admin." };
  }
  return {
    success: true,
    role: user.Role || "User",
    name: user.Name || username,
    project: user.Project_ID || ""
  };
}

// ---------------------------------------------------------------------------
// CURRENCY RATES
// ---------------------------------------------------------------------------

function getCurrencyRates_() {
  try {
    const sheet = getSheet_("Currency_Rates");
    const rows = sheetToObjects_(sheet);
    const rates = {};
    rows.forEach(r => {
      if (r.Code) rates[r.Code] = Number(r.Rate) || 0;
    });
    if (Object.keys(rates).length) return { success: true, data: rates };
  } catch (err) {
    // sheet mungkin belum dibuat, pakai default di bawah
  }
  return { success: true, data: { USD: 1, IDR: 16300, JPY: 157 } };
}

// ---------------------------------------------------------------------------
// MILESTONES (dipakai halaman Reports)
// ---------------------------------------------------------------------------

function getAllMilestones_() {
  const po = sheetToObjects_(getSheet_("PO_Tracking"));
  const milestoneFields = [
    ["kickOff", "Kick Off Meeting"],
    ["keyDocApproval", "Key Document Approval"],
    ["materialOrdered", "Material Ordered"],
    ["materialReceiptVendor", "Material Receipt by Vendor"],
    ["fabricationStart", "Fabrication Start"],
    ["fabricationCompletion", "Fabrication Completion"],
    ["packingDispatch", "Packing and Dispatch"],
    ["materialReceivedSite", "Material Received at Site"]
  ];
  const out = [];
  po.forEach(row => {
    milestoneFields.forEach(([key, label]) => {
      const plan = row[key + "Plan"];
      const actual = row[key + "Actual"];
      if (plan || actual) {
        out.push({
          poTrackingId: row.poTrackingId,
          Item: row.itemDescription,
          Milestone: label,
          Plan: plan || "",
          Actual: actual || ""
        });
      }
    });
  });
  return { success: true, data: out };
}

// ---------------------------------------------------------------------------
// DASHBOARD  (KPI ringkasan + Calendar badge count)
// ---------------------------------------------------------------------------

function getDashboardData_(upcomingDays) {
  const po = sheetToObjects_(getSheet_("PO_Tracking"));
  const fat = sheetToObjects_(getSheet_("FAT_Schedule"));
  const actions = sheetToObjects_(getSheet_("Action_Log"));
  const shipments = sheetToObjects_(getSheet_("Shipment_Tracking"));
  let risks = [];
  try { risks = sheetToObjects_(getSheet_("Risk_Register")); } catch (e) { risks = []; }

  const today = stripTime_(new Date());
  const horizon = new Date(today.getTime() + upcomingDays * 86400000);

  const poPlan = po.length;
  const poActual = po.filter(r => String(r.orderStatus).toLowerCase() === "close").length;

  const criticalItems = po
    .filter(r => r.category === "At Risk" || r.category === "Delay")
    .map(r => ({
      Item_Description: r.itemDescription,
      Vendor: r.supplier,
      RDD: r.deliveryForecast || r.deliveryPlan,
      Delay_Days: daysBetween_(r.deliveryPlan, r.deliveryForecast),
      Status: r.category === "Delay" ? "DELAY" : "AT RISK"
    }));

  const openActions = actions.filter(a => a.Status === "OPEN" || a.Status === "PENDING").length;
  const closedActions = actions.filter(a => a.Status === "CLOSED").length;

  const upcomingFATRows = fat.filter(f => {
    const d = parseDate_(f.FAT_Date);
    return d && d >= today && d <= horizon;
  });

  const overdueActions = actions.filter(a => {
    const d = parseDate_(a.Due_Date);
    return d && d < today && a.Status !== "CLOSED";
  }).length;
  const overduePO = po.filter(r => {
    const d = parseDate_(r.deliveryForecast);
    return d && d < today && String(r.orderStatus).toLowerCase() !== "close";
  }).length;

  const progressValues = po.map(r => Number(r.cumulativeProgress) || 0);
  const avgProgress = progressValues.length ? Math.round(100 * progressValues.reduce((a, b) => a + b, 0) / progressValues.length) : 0;

  return {
    success: true,
    kpi: {
      poPlan: poPlan,
      poActual: poActual,
      criticalItems: criticalItems.length,
      criticalPlan: 0,
      criticalActual: 0,
      overdueItems: overdueActions + overduePO,
      openActions: openActions,
      closedActions: closedActions,
      upcomingFAT: upcomingFATRows.length,
      fatPlan: fat.length,
      fatActual: fat.filter(f => f.Status === "Complete").length
    },
    progress: { actual: avgProgress, variance: 0 },
    sCurve: [],
    criticalItems: criticalItems,
    risks: risks,
    warnings: overdueActions || overduePO ? [
      (overdueActions + overduePO) + " item(s) melewati due date/delivery forecast."
    ] : [],
    fatSchedule: fat.slice(0, 5),
    actions: actions.filter(a => a.Status !== "CLOSED").slice(0, 5),
    shipments: shipments.slice(0, 5)
  };
}

function stripTime_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function parseDate_(v) {
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d) ? null : stripTime_(d);
}

function daysBetween_(planStr, forecastStr) {
  const plan = parseDate_(planStr);
  const forecast = parseDate_(forecastStr);
  if (!plan || !forecast) return 0;
  return Math.max(0, Math.round((forecast - plan) / 86400000));
}

// ---------------------------------------------------------------------------
// FILE UPLOAD (foto FAT, dokumen PO, dll)
// payload: { sheet, keyColumn, keyValue, column, fileName, mimeType, base64Data }
// ---------------------------------------------------------------------------

function uploadFileToDrive_(payload) {
  const { sheet: sheetName, keyColumn, keyValue, column, fileName, mimeType, base64Data } = payload;
  if (!base64Data) return { success: false, error: "File kosong." };

  const folder = UPLOAD_FOLDER_ID ? DriveApp.getFolderById(UPLOAD_FOLDER_ID) : DriveApp.getRootFolder();
  const bytes = Utilities.base64Decode(base64Data.split(",").pop());
  const blob = Utilities.newBlob(bytes, mimeType || "application/octet-stream", fileName || "upload");
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = "https://drive.google.com/file/d/" + file.getId() + "/view";

  if (sheetName && keyColumn && keyValue && column) {
    const sh = getSheet_(sheetName);
    const rowNum = findRowIndexByKey_(sh, keyColumn, keyValue);
    if (rowNum !== -1) {
      const colIdx = getOrCreateColumnIndex_(sh, column) + 1;
      sh.getRange(rowNum, colIdx).setValue(url);
    }
  }
  return { success: true, url: url, message: "File berhasil diupload." };
}
