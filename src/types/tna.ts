import { GroupedOrder, OrderItem } from './fms';

export const TNA_SPREADSHEET_ID = '1Wdu7PtyAEEnOY3R3OgvNM-Hnu9LnTnw1WKJQf06z-Qo';
export const TNA_SHEET_URL = `https://docs.google.com/spreadsheets/d/${TNA_SPREADSHEET_ID}/edit`;

/**
 * Exact Google Sheets formula requested by user for Cell K2 ("Order Status").
 */
export const TNA_K2_SHEET_FORMULA = `=LET(
  lastR, IFERROR(XMATCH(TRUE, ARRAYFORMULA(D3:D<>""), 0, -1), 0),
  IF(lastR=0, "Order Status",
    LET(
      rng, L3:INDEX(AA:AA, 2 + lastR),
      ok, LAMBDA(c, LET(t, UPPER(TRIM(c&"")), (t<>"")*(t<>"0")*(t<>"-")*(t<>"—")*(t<>"N/A")*(t<>"NA"))),
      VSTACK(
        "Order Status",
        BYROW(rng, LAMBDA(r,
          IF(INDEX(D:D, ROW(INDEX(r,1,1)))="", "",
            IF(ok(INDEX(r,1,13))+ok(INDEX(r,1,14))+ok(INDEX(r,1,15))>0, "Balance Fabric",
            IF(ok(INDEX(r,1,11))+ok(INDEX(r,1,12))>0, "Dispatch",
            IF(ok(INDEX(r,1,6))+ok(INDEX(r,1,7))+ok(INDEX(r,1,8))+ok(INDEX(r,1,9))>0, "FOB",
            IF(ok(INDEX(r,1,1))+ok(INDEX(r,1,2))+ok(INDEX(r,1,3))+ok(INDEX(r,1,4))>0, "Labdip", ""))))
          )
        ))
      )
    )
  )
)`;

export interface TNAManualFields {
  // Step 1. Lapdip (Cols L..O)
  labdipCommittedDate: string; // Col L (11)
  labdipActualDate: string;    // Col M (12)
  labdipRejectionDate: string; // Col N (13)
  labdipApprovalDate: string;  // Col O (14)
  // Col P (15): Delay Days = M - O (computed)

  // Step 2. FOB (Cols Q..T)
  fobCommittedDate: string;    // Col Q (16)
  fobActualDate: string;       // Col R (17)
  fobRejectionDate: string;    // Col S (18)
  fobApprovalDate: string;     // Col T (19)
  // Col U (20): Delay Days = R - T (computed)

  // Step 3. Dispatch (Cols V..W)
  actualDispatchDate: string;  // Col V (21)
  totalDispatchedQty: string;  // Col W (22)

  // Step 4. Balance fabric ( If goods pending ) (Cols X..AB)
  pendingQty: string;                 // Col X (23)
  committedBalanceDate: string;       // Col Y (24)
  actualBalanceDispatchDate: string;  // Col Z (25)
  balanceDispatchQty: string;         // Col AA (26)
  orderRemarks: string;               // Col AB (27)
  // Col AC (28): Delay Days = Y - Z (computed)

  updatedAt?: string;
}

export const EMPTY_TNA_MANUAL_FIELDS: TNAManualFields = {
  labdipCommittedDate: '',
  labdipActualDate: '',
  labdipRejectionDate: '',
  labdipApprovalDate: '',
  fobCommittedDate: '',
  fobActualDate: '',
  fobRejectionDate: '',
  fobApprovalDate: '',
  actualDispatchDate: '',
  totalDispatchedQty: '',
  pendingQty: '',
  committedBalanceDate: '',
  actualBalanceDispatchDate: '',
  balanceDispatchQty: '',
  orderRemarks: ''
};

export type TNAStageStatus = 'Balance Fabric' | 'Dispatch' | 'FOB' | 'Labdip' | '';

export interface TNARow {
  tnaKey: string;
  altKey: string;
  orderKey: string;
  itemId: string;
  // Cols A..J (carried from Order Automation where Status AN === 'Open')
  orderDate: string;     // A <- E (row[4])
  orderNo: string;       // B <- X (row[23])
  customerName: string;  // C <- C (row[2])
  itemName: string;      // D <- G (row[6])
  color: string;         // E <- H (row[7])
  width: string;         // F <- I (row[8])
  unit: string;          // G <- F (row[5])
  qty: number;           // H <- K (row[10])
  unitOfItem: string;    // I <- J (row[9])
  deliveryDate: string;  // J <- N (row[13])
  // Col K (computed)
  orderStatus: TNAStageStatus;
  // Manual fields L..AB
  manual: TNAManualFields;
  // Computed Delay Days
  labdipDelayDays: number | null;  // Col P: M - O
  fobDelayDays: number | null;     // Col U: R - T
  balanceDelayDays: number | null; // Col AC: Y - Z
}

export interface TNAColumnDef {
  colLetter: string;
  index: number; // 0..28
  groupLabel: string;
  groupTheme: 'order' | 'labdip' | 'delay' | 'fob' | 'dispatch' | 'balance';
  header: string;
  field?: keyof TNAManualFields | keyof Pick<TNARow, 'orderDate' | 'orderNo' | 'customerName' | 'itemName' | 'color' | 'width' | 'unit' | 'qty' | 'unitOfItem' | 'deliveryDate'>;
  kind: 'source' | 'status' | 'date' | 'number' | 'text' | 'delay';
  sourceRule?: string;
}

export const TNA_COLUMNS: TNAColumnDef[] = [
  { colLetter: 'A', index: 0, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Order Date', field: 'orderDate', kind: 'source', sourceRule: 'Col E (row[4])' },
  { colLetter: 'B', index: 1, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Order No', field: 'orderNo', kind: 'source', sourceRule: 'Col X (row[23])' },
  { colLetter: 'C', index: 2, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Customer Name', field: 'customerName', kind: 'source', sourceRule: 'Col C (row[2])' },
  { colLetter: 'D', index: 3, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Item Name', field: 'itemName', kind: 'source', sourceRule: 'Col G (row[6])' },
  { colLetter: 'E', index: 4, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Color', field: 'color', kind: 'source', sourceRule: 'Col H (row[7])' },
  { colLetter: 'F', index: 5, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Width', field: 'width', kind: 'source', sourceRule: 'Col I (row[8])' },
  { colLetter: 'G', index: 6, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Unit', field: 'unit', kind: 'source', sourceRule: 'Col F (row[5])' },
  { colLetter: 'H', index: 7, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Qty', field: 'qty', kind: 'source', sourceRule: 'Col K (row[10])' },
  { colLetter: 'I', index: 8, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Unit Of Item', field: 'unitOfItem', kind: 'source', sourceRule: 'Col J (row[9])' },
  { colLetter: 'J', index: 9, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Delivery Date', field: 'deliveryDate', kind: 'source', sourceRule: 'Col N (row[13])' },
  { colLetter: 'K', index: 10, groupLabel: 'Order Details (Open)', groupTheme: 'order', header: 'Order Status', kind: 'status', sourceRule: 'LET() Priority Formula' },

  { colLetter: 'L', index: 11, groupLabel: 'Step 1. Lapdip', groupTheme: 'labdip', header: 'Commited Date', field: 'labdipCommittedDate', kind: 'date' },
  { colLetter: 'M', index: 12, groupLabel: 'Step 1. Lapdip', groupTheme: 'labdip', header: 'Actual  Date', field: 'labdipActualDate', kind: 'date' },
  { colLetter: 'N', index: 13, groupLabel: 'Step 1. Lapdip', groupTheme: 'labdip', header: 'Rejection Date', field: 'labdipRejectionDate', kind: 'date' },
  { colLetter: 'O', index: 14, groupLabel: 'Step 1. Lapdip', groupTheme: 'labdip', header: 'Approval Date', field: 'labdipApprovalDate', kind: 'date' },
  { colLetter: 'P', index: 15, groupLabel: 'Delay Days', groupTheme: 'delay', header: 'Delay Days (M-O)', kind: 'delay', sourceRule: 'M - O' },

  { colLetter: 'Q', index: 16, groupLabel: 'Step 2. FOB', groupTheme: 'fob', header: 'Commited Date', field: 'fobCommittedDate', kind: 'date' },
  { colLetter: 'R', index: 17, groupLabel: 'Step 2. FOB', groupTheme: 'fob', header: 'Actual Date', field: 'fobActualDate', kind: 'date' },
  { colLetter: 'S', index: 18, groupLabel: 'Step 2. FOB', groupTheme: 'fob', header: 'Rejection Date', field: 'fobRejectionDate', kind: 'date' },
  { colLetter: 'T', index: 19, groupLabel: 'Step 2. FOB', groupTheme: 'fob', header: 'Approval Date', field: 'fobApprovalDate', kind: 'date' },
  { colLetter: 'U', index: 20, groupLabel: 'Delay Days', groupTheme: 'delay', header: 'Delay Days (R-T)', kind: 'delay', sourceRule: 'R - T' },

  { colLetter: 'V', index: 21, groupLabel: 'Step 3. Dispatch', groupTheme: 'dispatch', header: 'Actual Dispatch Date', field: 'actualDispatchDate', kind: 'date' },
  { colLetter: 'W', index: 22, groupLabel: 'Step 3. Dispatch', groupTheme: 'dispatch', header: 'Total Dispatched Qty', field: 'totalDispatchedQty', kind: 'number' },

  { colLetter: 'X', index: 23, groupLabel: 'Step 4. Balance fabric ( If goods pending )', groupTheme: 'balance', header: 'Pending Qty', field: 'pendingQty', kind: 'number' },
  { colLetter: 'Y', index: 24, groupLabel: 'Step 4. Balance fabric ( If goods pending )', groupTheme: 'balance', header: 'Commited balance date', field: 'committedBalanceDate', kind: 'date' },
  { colLetter: 'Z', index: 25, groupLabel: 'Step 4. Balance fabric ( If goods pending )', groupTheme: 'balance', header: 'Actual balance dispatch Date', field: 'actualBalanceDispatchDate', kind: 'date' },
  { colLetter: 'AA', index: 26, groupLabel: 'Step 4. Balance fabric ( If goods pending )', groupTheme: 'balance', header: 'Dispatch qty', field: 'balanceDispatchQty', kind: 'number' },
  { colLetter: 'AB', index: 27, groupLabel: 'Step 4. Balance fabric ( If goods pending )', groupTheme: 'balance', header: 'Order Remarks', field: 'orderRemarks', kind: 'text' },
  { colLetter: 'AC', index: 28, groupLabel: 'Delay Days', groupTheme: 'delay', header: 'Delay Days (Y-Z)', kind: 'delay', sourceRule: 'Y - Z' }
];

export const TNA_GROUP_SPANS: Array<{
  id: string;
  label: string;
  colSpan: number;
  theme: TNAColumnDef['groupTheme'];
}> = [
  { id: 'order', label: 'Order Details (Order Automation — Open Only)', colSpan: 11, theme: 'order' },
  { id: 'labdip', label: 'Step 1. Lapdip', colSpan: 4, theme: 'labdip' },
  { id: 'delay_p', label: 'Delay Days', colSpan: 1, theme: 'delay' },
  { id: 'fob', label: 'Step 2. FOB', colSpan: 4, theme: 'fob' },
  { id: 'delay_u', label: 'Delay Days', colSpan: 1, theme: 'delay' },
  { id: 'dispatch', label: 'Step 3. Dispatch', colSpan: 2, theme: 'dispatch' },
  { id: 'balance', label: 'Step 4. Balance fabric ( If goods pending )', colSpan: 5, theme: 'balance' },
  { id: 'delay_ac', label: 'Delay Days', colSpan: 1, theme: 'delay' }
];

/**
 * Matches the Google Sheets LAMBDA `ok`:
 * (t<>"")*(t<>"0")*(t<>"-")*(t<>"—")*(t<>"N/A")*(t<>"NA")
 */
export function isTnaValueOk(val: unknown): boolean {
  const t = String(val ?? '').trim().toUpperCase();
  return t !== '' && t !== '0' && t !== '-' && t !== '—' && t !== 'N/A' && t !== 'NA';
}

/**
 * Exact implementation of the user's K2 LET/BYROW Stage Priority Formula:
 * Range L3:AA3 (1-based indices inside L..AA):
 *   1..4  (L..O)  => Labdip
 *   6..9  (Q..T)  => FOB
 *   11..12(V..W)  => Dispatch
 *   13..15(X..Z)  => Balance Fabric (checked first: highest priority!)
 */
export function computeTnaOrderStatus(manual: TNAManualFields): TNAStageStatus {
  // 1. Balance Fabric: indices 13 (X), 14 (Y), 15 (Z) in L..AA
  if (
    isTnaValueOk(manual.pendingQty) ||
    isTnaValueOk(manual.committedBalanceDate) ||
    isTnaValueOk(manual.actualBalanceDispatchDate)
  ) {
    return 'Balance Fabric';
  }
  // 2. Dispatch: indices 11 (V), 12 (W)
  if (
    isTnaValueOk(manual.actualDispatchDate) ||
    isTnaValueOk(manual.totalDispatchedQty)
  ) {
    return 'Dispatch';
  }
  // 3. FOB: indices 6 (Q), 7 (R), 8 (S), 9 (T)
  if (
    isTnaValueOk(manual.fobCommittedDate) ||
    isTnaValueOk(manual.fobActualDate) ||
    isTnaValueOk(manual.fobRejectionDate) ||
    isTnaValueOk(manual.fobApprovalDate)
  ) {
    return 'FOB';
  }
  // 4. Labdip: indices 1 (L), 2 (M), 3 (N), 4 (O)
  if (
    isTnaValueOk(manual.labdipCommittedDate) ||
    isTnaValueOk(manual.labdipActualDate) ||
    isTnaValueOk(manual.labdipRejectionDate) ||
    isTnaValueOk(manual.labdipApprovalDate)
  ) {
    return 'Labdip';
  }
  return '';
}

const MONTH_NAMES_MAP: Record<string, number> = {
  jan: 1, january: 1,
  feb: 2, february: 2,
  mar: 3, march: 3,
  apr: 4, april: 4,
  may: 5,
  jun: 6, june: 6,
  jul: 7, july: 7,
  aug: 8, august: 8,
  sep: 9, sept: 9, september: 9,
  oct: 10, october: 10,
  nov: 11, november: 11,
  dec: 12, december: 12
};

/**
 * Parses YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY, or DD-MMM-YYYY to UTC midnight ms.
 */
export function parseTnaDateToUtcMs(val?: string | null): number | null {
  const s = (val || '').trim();
  if (!s || !isTnaValueOk(s)) return null;

  // 1. YYYY-MM-DD (from <input type="date"> or ISO string)
  const iso = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:[T\s].*)?$/);
  if (iso) {
    const y = Number(iso[1]);
    const m = Number(iso[2]);
    const d = Number(iso[3]);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return Date.UTC(y, m - 1, d);
    }
  }

  // 2. DD/MM/YYYY or D/M/YYYY or DD-MM-YYYY (Indian Sheet standard)
  const dmy = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:\s+.*)?$/);
  if (dmy) {
    const p1 = Number(dmy[1]);
    const p2 = Number(dmy[2]);
    let y = Number(dmy[3]);
    if (y < 100) y += 2000;
    // Interpret as DD/MM/YYYY unless month > 12 (fallback MM/DD/YYYY)
    let day = p1;
    let month = p2;
    if (p2 > 12 && p1 <= 12) {
      month = p1;
      day = p2;
    }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return Date.UTC(y, month - 1, day);
    }
  }

  // 3. DD MMM YYYY or DD-MMM-YYYY
  const textMonth = s.match(/^(\d{1,2})[\s\-\/]+([A-Za-z]+)[\s\-\/,]+(\d{2,4})$/);
  if (textMonth) {
    const day = Number(textMonth[1]);
    const month = MONTH_NAMES_MAP[textMonth[2].toLowerCase()];
    let y = Number(textMonth[3]);
    if (y < 100) y += 2000;
    if (month && day >= 1 && day <= 31) {
      return Date.UTC(y, month - 1, day);
    }
  }

  return null;
}

/**
 * Computes integer difference in days: (leftDate - rightDate)
 * Used for:
 *   Col P = M - O
 *   Col U = R - T
 *   Col AC = Y - Z
 */
export function computeTnaDelayDays(leftDate?: string, rightDate?: string): number | null {
  const leftMs = parseTnaDateToUtcMs(leftDate);
  const rightMs = parseTnaDateToUtcMs(rightDate);
  if (leftMs === null || rightMs === null) return null;
  return Math.round((leftMs - rightMs) / 86400000);
}

/**
 * Formats any recognized date string as DD/MM/YYYY for Google Sheets consistency,
 * or converts to YYYY-MM-DD for <input type="date">.
 */
export function toTnaDisplayDate(val?: string): string {
  const s = (val || '').trim();
  if (!s) return '';
  const ms = parseTnaDateToUtcMs(s);
  if (ms === null) return s;
  const d = new Date(ms);
  return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
}

export function toTnaHtmlDate(val?: string): string {
  const ms = parseTnaDateToUtcMs(val);
  if (ms === null) return '';
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
}

export function fromTnaHtmlDateToDisplay(isoVal: string): string {
  if (!isoVal) return '';
  const m = isoVal.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return isoVal;
  return `${m[3]}/${m[2]}/${m[1]}`;
}

/**
 * Sanitize Firebase RTDB key (forbids . # $ / [ ]).
 */
export function sanitizeTnaFbKey(raw: string): string {
  return String(raw || 'TNA_ROW')
    .trim()
    .toLowerCase()
    .replace(/[.#$/[\]\s]+/g, '_')
    .replace(/[^a-z0-9_|-]/g, '')
    .substring(0, 170);
}

/**
 * Builds deterministic key for a TNA item row so edits persist across syncs & devices.
 */
export function makeTnaKeys(order: GroupedOrder, item: OrderItem, occIdx: number): {
  tnaKey: string;
  altKey: string;
} {
  const cleanOrderNo = (order.orderNo || '').trim().toLowerCase();
  const cleanCust = (order.customerName || '').trim().toLowerCase();
  const cleanDate = (order.orderDate || '').trim().toLowerCase();
  const cleanItem = (item.itemName || '').trim().toLowerCase();
  const cleanColor = (item.color || '').trim().toLowerCase();
  const cleanWidth = (item.width || '').trim().toLowerCase();
  const cleanQty = String(Number(item.qty) || 0);

  const primary = sanitizeTnaFbKey(
    `${cleanOrderNo || cleanDate}|${cleanCust}|${cleanItem}|${cleanColor}|${cleanWidth}|${cleanQty}|#${occIdx}`
  );
  const alt = sanitizeTnaFbKey(
    `${cleanOrderNo || cleanDate}|${cleanCust}|${cleanItem}|${cleanColor}|${cleanWidth}|${cleanQty}`
  );
  return { tnaKey: primary, altKey: alt };
}
