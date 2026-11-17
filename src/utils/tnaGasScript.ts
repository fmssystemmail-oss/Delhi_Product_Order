import { TNA_SPREADSHEET_ID, TNA_K2_SHEET_FORMULA } from '../types/tna';

export function generateTnaGasScript(
  firebaseUrl: string = 'https://delhi-product-order-default-rtdb.firebaseio.com',
  firebaseSecret: string = 'iIElINFgR124ssaG3QWjExz2lT72cRh5g502oP6J'
): string {
  const cleanUrl = firebaseUrl.replace(/\/+$/, '');
  const escapedFormula = TNA_K2_SHEET_FORMULA.replace(/\\/g, '\\\\').replace(/`/g, '\\`');

  return `/**
 * ==============================================================================
 * GINZA FMS - ZERO-QUOTA TNA (TIME & ACTION) 2-WAY GOOGLE SHEET SYNC SCRIPT
 * Target Google Sheet ID : ${TNA_SPREADSHEET_ID}
 * Source FMS Sheet ID    : 15jjr5tjxaWRPiv_nh6L9EZyFHNFZyQ3PQ6Jd89cV4k0
 * Firebase Realtime DB   : ${cleanUrl}
 *
 * KEY FEATURES:
 *   1. ZERO UrlFetchApp calls in doPost(e): Web App sends rows directly in JSON
 *      so updates NEVER fail with "premium urlfetch" daily quota errors!
 *   2. Direct Sheet-to-Sheet fallback: Can read Open orders (AN = "Open")
 *      directly from "Order Automation 1" without needing Firebase HTTP calls.
 *   3. Exact K2 =LET() Order Status formula & P (M-O), U (R-T), AC (Y-Z) Delays.
 *   4. True Date serial conversion so DD/MM/YYYY calculates Delay Days cleanly.
 *   5. Auto-removes Closed orders from the TNA Google Sheet when AN != "Open".
 * ==============================================================================
 */

const TNA_CFG = {
  TARGET_SPREADSHEET_ID: "${TNA_SPREADSHEET_ID}",
  SOURCE_AUTOMATION_SHEET_ID: "15jjr5tjxaWRPiv_nh6L9EZyFHNFZyQ3PQ6Jd89cV4k0",
  DATABASE_URL: "${cleanUrl}",
  DATABASE_SECRET: "${firebaseSecret}",
  TNA_TAB_NAME: "TNA",
  FIRST_DATA_ROW: 3,
  TOTAL_COLS: 29 // A (1) to AC (29)
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 TNA Live Sync")
    .addItem("1. Setup Headers, K2 Formula & 5-Min Auto Sync", "setupTnaTwoWaySync")
    .addItem("2. Sync Now (From Automation / Firebase -> TNA Sheet)", "syncTnaNow")
    .addItem("3. Apply K2 & Delay Formulas Only", "applyTnaFormulasOnly")
    .addToUi();
}

function setupTnaTwoWaySync() {
  const ss = SpreadsheetApp.openById(TNA_CFG.TARGET_SPREADSHEET_ID);
  const sheet = getOrCreateTnaSheet_(ss);
  setupTnaHeaders_(sheet, true);
  syncTnaNow();

  ScriptApp.getProjectTriggers().forEach(function(trig) {
    const fn = trig.getHandlerFunction();
    if (fn === "syncTnaTwoWayWithFirebase" || fn === "syncTnaNow") {
      ScriptApp.deleteTrigger(trig);
    }
  });
  ScriptApp.newTrigger("syncTnaNow")
    .timeBased()
    .everyMinutes(5)
    .create();

  ss.toast("TNA Headers, Formulas & 5-Minute Sync Activated!", "TNA Ready", 6);
}

function getOrCreateTnaSheet_(ss) {
  let sheet =
    ss.getSheetByName(TNA_CFG.TNA_TAB_NAME) ||
    ss.getSheetByName("TNA ") ||
    ss.getSheets()[0];
  if (!sheet) {
    sheet = ss.insertSheet(TNA_CFG.TNA_TAB_NAME);
  }
  return sheet;
}

function setupTnaHeaders_(sheet, forceRestyle) {
  if (sheet.getMaxColumns() < TNA_CFG.TOTAL_COLS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), TNA_CFG.TOTAL_COLS - sheet.getMaxColumns());
  }

  const d2 = String(sheet.getRange(2, 4).getDisplayValue() || "").trim().toLowerCase();
  const l1 = String(sheet.getRange(1, 12).getDisplayValue() || "").trim().toLowerCase();
  if (!forceRestyle && d2 === "item name" && l1.indexOf("lapdip") !== -1) {
    return;
  }

  // Break any conflicting merges in Row 1 first
  sheet.getRange(1, 1, 1, TNA_CFG.TOTAL_COLS).breakApart();

  const row1 = [
    "", "", "", "", "", "", "", "", "", "", "",
    "Step 1. Lapdip", "", "", "",
    "Delay Days",
    "Step 2. FOB", "", "", "",
    "Delay Days",
    "Step 3. Dispatch", "",
    "Step 4. Balance fabric ( If goods pending )", "", "", "", "",
    "Delay Days"
  ];

  const row2 = [
    "Order Date",
    "Order No",
    "Customer Name",
    "Item Name",
    "Color",
    "Width",
    "Unit",
    "Qty",
    "Unit Of Item",
    "Delivery Date",
    "Order Status",
    "Commited Date",
    "Actual  Date",
    "Rejection Date",
    "Approval Date",
    "",
    "Commited Date",
    "Actual Date",
    "Rejection Date",
    "Approval Date",
    "",
    "Actual Dispatch Date",
    "Total Dispatched Qty",
    "Pending Qty",
    "Commited balance date",
    "Actual balance dispatch Date",
    "Dispatch qty",
    "Order Remarks",
    ""
  ];

  sheet.getRange(1, 1, 1, TNA_CFG.TOTAL_COLS).setValues([row1]);
  sheet.getRange(2, 1, 1, TNA_CFG.TOTAL_COLS).setValues([row2]);

  sheet.getRange(1, 12, 1, 4).merge().setBackground("#312e81").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(1, 16, 1, 1).setBackground("#9f1239").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(1, 17, 1, 4).merge().setBackground("#0c4a6e").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(1, 21, 1, 1).setBackground("#9f1239").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(1, 22, 1, 2).merge().setBackground("#065f46").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(1, 24, 1, 5).merge().setBackground("#92400e").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(1, 29, 1, 1).setBackground("#9f1239").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange(2, 1, 1, TNA_CFG.TOTAL_COLS).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setFrozenRows(2);
}

function applyTnaFormulasOnly(targetSheet, explicitRowCount) {
  const sheet = targetSheet || getOrCreateTnaSheet_(SpreadsheetApp.openById(TNA_CFG.TARGET_SPREADSHEET_ID));

  // Clear any leftover manual values in K3:K so K2 =LET(...) array spill never hits #REF!
  const clearRows = Math.max(sheet.getLastRow(), explicitRowCount ? 2 + explicitRowCount : 3, 3) - 2;
  if (clearRows > 0) {
    sheet.getRange(3, 11, clearRows, 1).clearContent();
  }

  const k2Formula = \`${escapedFormula}\`;
  sheet.getRange("K2").setFormula(k2Formula);

  const lastRow = explicitRowCount ? (2 + explicitRowCount) : sheet.getLastRow();
  const count = Math.max(0, lastRow - 2);
  if (count > 0) {
    const pFormulas = [];
    const uFormulas = [];
    const acFormulas = [];
    for (let i = 0; i < count; i++) {
      const r = 3 + i;
      pFormulas.push(['=IF(AND(ISNUMBER(M' + r + '),ISNUMBER(O' + r + ')), M' + r + '-O' + r + ', "")']);
      uFormulas.push(['=IF(AND(ISNUMBER(R' + r + '),ISNUMBER(T' + r + ')), R' + r + '-T' + r + ', "")']);
      acFormulas.push(['=IF(AND(ISNUMBER(Y' + r + '),ISNUMBER(Z' + r + ')), Y' + r + '-Z' + r + ', "")']);
    }
    sheet.getRange(3, 16, count, 1).setFormulas(pFormulas).setNumberFormat("0");   // Col P (16)
    sheet.getRange(3, 21, count, 1).setFormulas(uFormulas).setNumberFormat("0");   // Col U (21)
    sheet.getRange(3, 29, count, 1).setFormulas(acFormulas).setNumberFormat("0");  // Col AC (29)
  }
}

/**
 * Converts DD/MM/YYYY or YYYY-MM-DD strings into native Date objects
 * so M-O, R-T, Y-Z Delay formulas evaluate properly in any Google Sheet locale.
 */
function tnaParseDateValue_(val) {
  if (val instanceof Date) return val;
  const s = String(val || "").trim();
  if (!s) return "";

  const iso = s.match(/^(\\d{4})[\\/\\-](\\d{1,2})[\\/\\-](\\d{1,2})(?:[T\\s].*)?$/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }

  const dmy = s.match(/^(\\d{1,2})[\\/\\.\\-](\\d{1,2})[\\/\\.\\-](\\d{2,4})(?:\\s+.*)?$/);
  if (dmy) {
    const p1 = Number(dmy[1]);
    const p2 = Number(dmy[2]);
    let y = Number(dmy[3]);
    if (y < 100) y += 2000;
    let d = p1, m = p2;
    if (p2 > 12 && p1 <= 12) {
      m = p1;
      d = p2;
    }
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return new Date(y, m - 1, d);
    }
  }
  return s;
}

function tnaSanitizeKey_(raw) {
  return String(raw || "TNA_ROW")
    .trim()
    .toLowerCase()
    .replace(/[.#$/\\[\\]\\s]+/g, "_")
    .replace(/[^a-z0-9_|-]/g, "")
    .substring(0, 170);
}

function tnaFbGetSafe_(path) {
  try {
    const auth = TNA_CFG.DATABASE_SECRET ? "?auth=" + encodeURIComponent(TNA_CFG.DATABASE_SECRET) : "";
    const res = UrlFetchApp.fetch(TNA_CFG.DATABASE_URL + "/" + path + ".json" + auth, { muteHttpExceptions: true });
    if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
      return JSON.parse(res.getContentText() || "{}") || {};
    }
  } catch (e) {}
  return null;
}

/**
 * Fallback direct Sheet-to-Sheet reader (ZERO UrlFetchApp quota)
 * Reads ONLY rows from "Order Automation 1" where Col AN == "Open" and Item Name (G) != "".
 */
function readOpenRowsDirectlyFromAutomationSheet_() {
  try {
    const srcSs = SpreadsheetApp.openById(TNA_CFG.SOURCE_AUTOMATION_SHEET_ID);
    const srcSheet =
      srcSs.getSheetByName("Order Automation 1") ||
      srcSs.getSheetByName("Order Automation 1 ") ||
      srcSs.getSheetByName("Product Automation") ||
      srcSs.getSheets()[0];
    if (!srcSheet) return [];

    const lastRow = srcSheet.getLastRow();
    if (lastRow < 2) return [];
    const maxCols = Math.max(srcSheet.getLastColumn(), 40);
    const disp = srcSheet.getRange(2, 1, lastRow - 1, maxCols).getDisplayValues();
    const raw = srcSheet.getRange(2, 1, lastRow - 1, maxCols).getValues();

    const rows = [];
    const occCounter = {};
    let curOrderDate = "";
    let curOrderNo = "";
    let curCustomer = "";
    let curStatusAN = "Open";

    for (let i = 0; i < disp.length; i++) {
      const r = disp[i];
      const hasOrderHeader = Boolean(String(r[2] || "").trim() || String(r[4] || "").trim() || String(r[23] || "").trim());
      if (String(r[2] || "").trim()) curCustomer = String(r[2] || "").trim();
      if (String(r[4] || "").trim()) curOrderDate = String(r[4] || "").trim();
      if (String(r[23] || "").trim()) curOrderNo = String(r[23] || "").trim();
      if (String(r[39] || "").trim()) {
        curStatusAN = String(r[39] || "").trim();
      } else if (hasOrderHeader) {
        curStatusAN = "Open";
      }

      const itemName = String(r[6] || "").trim();
      if (!itemName) continue;

      const statusAN = String(r[39] || "").trim() || curStatusAN || "Open";
      if (statusAN !== "Open") continue;

      const color = String(r[7] || "").trim();
      const width = String(r[8] || "").trim();
      const qty = parseFloat(raw[i][10]) || 0;
      const base = (curOrderNo || curOrderDate).toLowerCase() + "|" + curCustomer.toLowerCase() + "|" + itemName.toLowerCase() + "|" + color.toLowerCase() + "|" + width.toLowerCase() + "|" + qty;
      const occ = (occCounter[base] || 0) + 1;
      occCounter[base] = occ;

      rows.push({
        tnaKey: tnaSanitizeKey_(base + "|#" + occ),
        altKey: tnaSanitizeKey_(base),
        orderDate: curOrderDate,
        orderNo: curOrderNo,
        customerName: curCustomer,
        itemName: itemName,
        color: color,
        width: width,
        unit: String(r[5] || "").trim(),
        qty: qty,
        unitOfItem: String(r[9] || "").trim(),
        deliveryDate: String(r[13] || "").trim()
      });
    }
    return rows;
  } catch (e) {
    return [];
  }
}

/**
 * Core Batch Writer: Writes Open TNA rows (A..J, L..O, Q..T, V..AB) to Sheet
 * and removes any leftover rows when orders move from Open -> Closed!
 */
function writeTnaRowsToSheet_(snapshotRows, fbRecords, preferWebApp) {
  const ss = SpreadsheetApp.openById(TNA_CFG.TARGET_SPREADSHEET_ID);
  const sheet = getOrCreateTnaSheet_(ss);
  setupTnaHeaders_(sheet, false);

  if (!snapshotRows || snapshotRows.length === 0) {
    const prevLast = sheet.getLastRow();
    if (prevLast >= TNA_CFG.FIRST_DATA_ROW) {
      sheet.getRange(TNA_CFG.FIRST_DATA_ROW, 1, prevLast - 2, TNA_CFG.TOTAL_COLS).clearContent();
    }
    applyTnaFormulasOnly(sheet, 0);
    return { ok: true, count: 0 };
  }

  // Read existing sheet entries (L:O, Q:T, V:AB) so direct sheet edits aren't lost
  const sheetLastRow = sheet.getLastRow();
  const sheetManualByKey = {};
  if (sheetLastRow >= TNA_CFG.FIRST_DATA_ROW) {
    const existingVals = sheet.getRange(TNA_CFG.FIRST_DATA_ROW, 1, sheetLastRow - 2, TNA_CFG.TOTAL_COLS).getDisplayValues();
    const occCounter = {};
    for (let i = 0; i < existingVals.length; i++) {
      const r = existingVals[i];
      const itemName = String(r[3] || "").trim();
      if (!itemName) continue;
      const ordNo = String(r[1] || "").trim().toLowerCase();
      const ordDate = String(r[0] || "").trim().toLowerCase();
      const cust = String(r[2] || "").trim().toLowerCase();
      const color = String(r[4] || "").trim().toLowerCase();
      const width = String(r[5] || "").trim().toLowerCase();
      const qty = String(parseFloat(r[7]) || 0);
      const base = (ordNo || ordDate) + "|" + cust + "|" + itemName.toLowerCase() + "|" + color + "|" + width + "|" + qty;
      const occ = (occCounter[base] || 0) + 1;
      occCounter[base] = occ;
      const k = tnaSanitizeKey_(base + "|#" + occ);

      sheetManualByKey[k] = {
        labdipCommittedDate: String(r[11] || "").trim(),
        labdipActualDate: String(r[12] || "").trim(),
        labdipRejectionDate: String(r[13] || "").trim(),
        labdipApprovalDate: String(r[14] || "").trim(),
        fobCommittedDate: String(r[16] || "").trim(),
        fobActualDate: String(r[17] || "").trim(),
        fobRejectionDate: String(r[18] || "").trim(),
        fobApprovalDate: String(r[19] || "").trim(),
        actualDispatchDate: String(r[21] || "").trim(),
        totalDispatchedQty: String(r[22] || "").trim(),
        pendingQty: String(r[23] || "").trim(),
        committedBalanceDate: String(r[24] || "").trim(),
        actualBalanceDispatchDate: String(r[25] || "").trim(),
        balanceDispatchQty: String(r[26] || "").trim(),
        orderRemarks: String(r[27] || "").trim()
      };
    }
  }

  const outAtoJ = [];
  const outLtoO = [];
  const outQtoT = [];
  const outVtoAB = [];

  for (let i = 0; i < snapshotRows.length; i++) {
    const src = snapshotRows[i];
    const key = src.tnaKey;
    const inlineManual = src.manual || null;
    const fbMan = inlineManual || (fbRecords && (fbRecords[key] || fbRecords[src.altKey])) || null;
    const shMan = sheetManualByKey[key] || {};

    const pick = function(field) {
      const fbVal = fbMan && fbMan[field] !== undefined ? String(fbMan[field] || "").trim() : "";
      const shVal = shMan[field] !== undefined ? String(shMan[field] || "").trim() : "";
      if (preferWebApp && fbMan) {
        return fbVal;
      }
      return fbVal !== "" ? fbVal : shVal;
    };

    outAtoJ.push([
      tnaParseDateValue_(src.orderDate || ""),
      src.orderNo || "",
      src.customerName || "",
      src.itemName || "",
      src.color || "",
      src.width || "",
      src.unit || "",
      Number(src.qty) || 0,
      src.unitOfItem || "",
      tnaParseDateValue_(src.deliveryDate || "")
    ]);

    outLtoO.push([
      tnaParseDateValue_(pick("labdipCommittedDate")),
      tnaParseDateValue_(pick("labdipActualDate")),
      tnaParseDateValue_(pick("labdipRejectionDate")),
      tnaParseDateValue_(pick("labdipApprovalDate"))
    ]);

    outQtoT.push([
      tnaParseDateValue_(pick("fobCommittedDate")),
      tnaParseDateValue_(pick("fobActualDate")),
      tnaParseDateValue_(pick("fobRejectionDate")),
      tnaParseDateValue_(pick("fobApprovalDate"))
    ]);

    outVtoAB.push([
      tnaParseDateValue_(pick("actualDispatchDate")),
      pick("totalDispatchedQty"),
      pick("pendingQty"),
      tnaParseDateValue_(pick("committedBalanceDate")),
      tnaParseDateValue_(pick("actualBalanceDispatchDate")),
      pick("balanceDispatchQty"),
      pick("orderRemarks")
    ]);
  }

  const neededLastRow = 2 + outAtoJ.length;
  if (sheet.getMaxRows() < neededLastRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededLastRow - sheet.getMaxRows());
  }

  sheet.getRange(3, 1, outAtoJ.length, 10).setValues(outAtoJ);
  sheet.getRange(3, 1, outAtoJ.length, 1).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(3, 10, outAtoJ.length, 1).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(3, 12, outLtoO.length, 4).setValues(outLtoO).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(3, 17, outQtoT.length, 4).setValues(outQtoT).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(3, 22, outVtoAB.length, 7).setValues(outVtoAB);
  sheet.getRange(3, 22, outVtoAB.length, 1).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(3, 25, outVtoAB.length, 2).setNumberFormat("dd/MM/yyyy");

  // Remove any leftover rows below neededLastRow (e.g. when orders close in Order Automation!)
  const prevLast = sheet.getLastRow();
  if (prevLast > neededLastRow) {
    sheet.getRange(neededLastRow + 1, 1, prevLast - neededLastRow, TNA_CFG.TOTAL_COLS).clearContent();
  }

  applyTnaFormulasOnly(sheet, outAtoJ.length);
  return { ok: true, count: outAtoJ.length, version: "ZERO_QUOTA_V2", sheetManualByKey: sheetManualByKey };
}

/**
 * Triggered manually from menu or 5-min trigger (reads Firebase if available, else direct Sheet-to-Sheet).
 */
function syncTnaNow() {
  const openSnap = tnaFbGetSafe_("tna_open_snapshot");
  const fbRecords = tnaFbGetSafe_("tna_records") || {};
  let rows = (openSnap && Array.isArray(openSnap.rows)) ? openSnap.rows : [];
  if (rows.length === 0) {
    rows = readOpenRowsDirectlyFromAutomationSheet_();
  }
  const result = writeTnaRowsToSheet_(rows, fbRecords, false);
  try {
    SpreadsheetApp.openById(TNA_CFG.TARGET_SPREADSHEET_ID).toast(
      "Synced " + result.count + " Open TNA rows!",
      "TNA Sync Complete",
      4
    );
  } catch (e) {}
  return result;
}

function syncTnaTwoWayWithFirebase(preferWebApp) {
  return syncTnaNow();
}

function doGet(e) {
  let rowCount = 0;
  let tabName = TNA_CFG.TNA_TAB_NAME;
  try {
    const ss = SpreadsheetApp.openById(TNA_CFG.TARGET_SPREADSHEET_ID);
    const sh = getOrCreateTnaSheet_(ss);
    tabName = sh.getName();
    rowCount = Math.max(0, sh.getLastRow() - 2);
  } catch (err) {}
  return ContentService.createTextOutput(
    JSON.stringify({
      ok: true,
      service: "GINZA_TNA_SYNC",
      version: "ZERO_QUOTA_V2",
      spreadsheetId: TNA_CFG.TARGET_SPREADSHEET_ID,
      tabName: tabName,
      rowCount: rowCount
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Instant Zero-Quota Webhook:
 * Receives full Open TNA rows + manual edits directly in POST body from Web App
 * (ZERO UrlFetchApp calls needed — never hits daily UrlFetch quota!).
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }

    if (payload && Array.isArray(payload.rows)) {
      const result = writeTnaRowsToSheet_(
        payload.rows,
        payload.tnaRecords || null,
        Boolean(payload.preferWebApp)
      );
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const fallbackResult = syncTnaNow();
    return ContentService.createTextOutput(JSON.stringify(fallbackResult))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
`;
}
