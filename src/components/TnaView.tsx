import React, { useEffect, useMemo, useState } from 'react';
import { GroupedOrder } from '../types/fms';
import {
  EMPTY_TNA_MANUAL_FIELDS,
  TNA_COLUMNS,
  TNA_GROUP_SPANS,
  TNA_K2_SHEET_FORMULA,
  TNA_SHEET_URL,
  TNA_SPREADSHEET_ID,
  TNAManualFields,
  TNARow,
  TNAStageStatus,
  computeTnaDelayDays,
  computeTnaOrderStatus,
  fromTnaHtmlDateToDisplay,
  makeTnaKeys,
  sanitizeTnaFbKey,
  toTnaDisplayDate,
  toTnaHtmlDate
} from '../types/tna';
import { MONTH_OPTIONS, parseOrderDateParts } from './ProductOrderView';
import { generateTnaGasScript } from '../utils/tnaGasScript';
import {
  ArrowDown,
  ArrowUp,
  Calendar,
  Check,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  Clock,
  Copy,
  Download,
  ExternalLink,
  Filter,
  Link2,
  Loader2,
  Maximize2,
  Minimize2,
  RefreshCw,
  Search,
  Sparkles,
  X
} from 'lucide-react';

interface TnaViewProps {
  openOrders: GroupedOrder[];
  tnaRecords: Record<string, TNAManualFields>;
  onUpdateTnaField: (
    tnaKey: string,
    altKey: string,
    field: keyof TNAManualFields,
    value: string
  ) => void;
  onBulkMergeTnaRecords: (records: Record<string, TNAManualFields>) => void;
  onSyncTnaAll: (rows: TNARow[], manualClick?: boolean, preferWebApp?: boolean) => Promise<void>;
  isSyncingTna: boolean;
  tnaSyncNotice: { type: 'success' | 'info' | 'error'; text: string } | null;
  tnaWebAppUrl: string;
  onSaveTnaWebAppUrl: (url: string) => void;
  firebaseUrl: string;
  firebaseSecret: string;
  isFullWidth: boolean;
  setIsFullWidth?: (val: boolean) => void;
  freezeColumns?: number;
  freezeRows?: number;
  autoFitColumns?: boolean;
  searchTerm: string;
  setSearchTerm: (val: string) => void;
}

function parseCsvRowsSimple(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];
    if (inQuotes) {
      if (ch === '"' && next === '"') {
        current += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        current += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(current);
      current = '';
    } else if (ch === '\r' && next === '\n') {
      row.push(current);
      rows.push(row);
      row = [];
      current = '';
      i++;
    } else if (ch === '\n' || ch === '\r') {
      row.push(current);
      rows.push(row);
      row = [];
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.length > 0 || row.length > 0) {
    row.push(current);
    rows.push(row);
  }
  return rows;
}

/**
 * Inline Spreadsheet Date Cell:
 * Displays DD/MM/YYYY cleanly and lets the user pick via calendar OR clear in 1 click.
 */
const TnaDateCell: React.FC<{
  value: string;
  onChange: (newDisplayDate: string) => void;
  accentClass?: string;
}> = ({ value, onChange, accentClass = '' }) => {
  const displayVal = toTnaDisplayDate(value);
  const htmlDateVal = toTnaHtmlDate(value);

  return (
    <div className={`relative flex items-center gap-0.5 w-full group ${accentClass}`}>
      <input
        type="date"
        value={htmlDateVal}
        onChange={(e) => {
          const v = e.target.value;
          onChange(v ? fromTnaHtmlDateToDisplay(v) : '');
        }}
        title={displayVal ? `Date: ${displayVal}` : 'Select date'}
        className={`w-full min-w-0 px-0.5 py-0.5 text-[10px] font-mono rounded border transition-colors cursor-pointer focus:outline-none focus:ring-1 focus:ring-indigo-500 ${
          displayVal
            ? 'bg-white border-slate-300 text-slate-900 font-semibold'
            : 'bg-slate-50/70 border-slate-200/90 text-slate-400 hover:bg-white hover:border-slate-300'
        }`}
      />
      {displayVal && (
        <button
          type="button"
          onClick={() => onChange('')}
          title="Clear date"
          className="p-0 text-slate-400 hover:text-rose-600 rounded opacity-60 group-hover:opacity-100 cursor-pointer shrink-0"
        >
          <X className="w-2.5 h-2.5" />
        </button>
      )}
    </div>
  );
};

export const TnaView: React.FC<TnaViewProps> = ({
  openOrders,
  tnaRecords,
  onUpdateTnaField,
  onBulkMergeTnaRecords,
  onSyncTnaAll,
  isSyncingTna,
  tnaSyncNotice,
  tnaWebAppUrl,
  onSaveTnaWebAppUrl,
  firebaseUrl,
  firebaseSecret,
  isFullWidth,
  setIsFullWidth,
  freezeColumns = 2,
  freezeRows = 1,
  autoFitColumns = true,
  searchTerm,
  setSearchTerm
}) => {
  const [stageFilter, setStageFilter] = useState<'ALL' | 'Labdip' | 'FOB' | 'Dispatch' | 'Balance Fabric' | 'BLANK'>('ALL');
  const [monthFilter, setMonthFilter] = useState<string>('ALL');
  const [yearFilter, setYearFilter] = useState<string>('ALL');
  const [page, setPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(50);
  const [colFilters, setColFilters] = useState<Record<string, string[]>>({});
  const [sortConfig, setSortConfig] = useState<{ colLetter: string; direction: 'asc' | 'desc' } | null>(null);
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [filterSearch, setFilterSearch] = useState<string>('');

  // Setup & Google Sheet Script Modal
  const [showSheetSetupModal, setShowSheetSetupModal] = useState(false);
  const [draftWebAppUrl, setDraftWebAppUrl] = useState(tnaWebAppUrl);
  const [copiedScript, setCopiedScript] = useState(false);
  const [copiedFormula, setCopiedFormula] = useState(false);
  const [isPullingSheetCsv, setIsPullingSheetCsv] = useState(false);
  const [csvPullMsg, setCsvPullMsg] = useState<string | null>(null);
  const [webhookHealth, setWebhookHealth] = useState<{
    status: 'idle' | 'checking' | 'ok' | 'outdated' | 'error';
    message: string;
  }>({ status: 'idle', message: '' });

  const testWebhookConnection = async (urlToTest: string) => {
    const clean = urlToTest.trim();
    if (!clean.startsWith('http')) {
      setWebhookHealth({ status: 'error', message: 'Enter a valid Apps Script /exec URL.' });
      return;
    }
    setWebhookHealth({ status: 'checking', message: 'Checking deployment...' });
    try {
      const resp = await fetch(clean, { method: 'GET' });
      const text = await resp.text();
      try {
        const json = JSON.parse(text);
        if (json && json.ok && json.service === 'GINZA_TNA_SYNC') {
          setWebhookHealth({
            status: 'ok',
            message: `Connected (${json.version || 'Active'}) · Sheet rows: ${json.rowCount ?? 0}`
          });
          return;
        }
      } catch {}
      if (text.includes('<!DOCTYPE') || text.includes('Script function not found')) {
        setWebhookHealth({
          status: 'outdated',
          message:
            'This URL is still running the OLD script (no doGet). In Apps Script: Save code -> Deploy -> Manage deployments -> Edit (pencil) -> Version: New version -> Deploy. (Or if you used New deployment, paste the NEW /exec URL above!)'
        });
      } else {
        setWebhookHealth({ status: 'error', message: `Unexpected response (HTTP ${resp.status}).` });
      }
    } catch (err: any) {
      setWebhookHealth({ status: 'error', message: err?.message || 'Unable to reach Web App URL.' });
    }
  };

  // Build flat 29-column TNA rows ONLY from Order Automation items whose Col AN === "Open"
  const allTnaRows: TNARow[] = useMemo(() => {
    const rows: TNARow[] = [];
    const occCounter: Record<string, number> = {};

    openOrders.forEach(order => {
      if (order.sheetSource === 'closed') return;
      order.items.forEach(item => {
        const itemStatusAN = item.statusAN || 'Open';
        if (itemStatusAN !== 'Open') return;
        const itemName = (item.itemName || '').trim();
        if (!itemName) return; // D (Item Name) must be non-blank to sync

        const baseSig = `${(order.orderNo || order.orderDate || '').trim().toLowerCase()}|${(order.customerName || '').trim().toLowerCase()}|${itemName.toLowerCase()}|${(item.color || '').trim().toLowerCase()}|${(item.width || '').trim().toLowerCase()}|${Number(item.qty) || 0}`;
        const occ = (occCounter[baseSig] || 0) + 1;
        occCounter[baseSig] = occ;

        const { tnaKey, altKey } = makeTnaKeys(order, item, occ);
        const manual: TNAManualFields = tnaRecords[tnaKey] || tnaRecords[altKey] || EMPTY_TNA_MANUAL_FIELDS;
        const orderStatus = computeTnaOrderStatus(manual);

        rows.push({
          tnaKey,
          altKey,
          orderKey: order.orderKey,
          itemId: item.id,
          orderDate: order.orderDate || '',     // A <- E
          orderNo: order.orderNo || '',         // B <- X
          customerName: order.customerName || '', // C <- C
          itemName,                             // D <- G
          color: item.color || '',              // E <- H
          width: item.width || '',              // F <- I
          unit: item.unit || '',                // G <- F
          qty: Number(item.qty) || 0,           // H <- K
          unitOfItem: item.unitOfItem || '',    // I <- J
          deliveryDate: item.deliveryDate || '', // J <- N
          orderStatus,                          // K <- dynamic LET()
          manual,
          labdipDelayDays: computeTnaDelayDays(manual.labdipActualDate, manual.labdipApprovalDate), // P = M - O
          fobDelayDays: computeTnaDelayDays(manual.fobActualDate, manual.fobApprovalDate),          // U = R - T
          balanceDelayDays: computeTnaDelayDays(manual.committedBalanceDate, manual.actualBalanceDispatchDate) // AC = Y - Z
        });
      });
    });

    return rows;
  }, [openOrders, tnaRecords]);

  // Helper to get string representation of any column A..AC for filtering/sorting/CSV export
  const getTnaCellValue = (row: TNARow, colLetter: string): string => {
    switch (colLetter) {
      case 'A': return row.orderDate;
      case 'B': return row.orderNo;
      case 'C': return row.customerName;
      case 'D': return row.itemName;
      case 'E': return row.color;
      case 'F': return row.width;
      case 'G': return row.unit;
      case 'H': return String(row.qty);
      case 'I': return row.unitOfItem;
      case 'J': return row.deliveryDate;
      case 'K': return row.orderStatus;
      case 'L': return toTnaDisplayDate(row.manual.labdipCommittedDate);
      case 'M': return toTnaDisplayDate(row.manual.labdipActualDate);
      case 'N': return toTnaDisplayDate(row.manual.labdipRejectionDate);
      case 'O': return toTnaDisplayDate(row.manual.labdipApprovalDate);
      case 'P': return row.labdipDelayDays !== null ? String(row.labdipDelayDays) : '';
      case 'Q': return toTnaDisplayDate(row.manual.fobCommittedDate);
      case 'R': return toTnaDisplayDate(row.manual.fobActualDate);
      case 'S': return toTnaDisplayDate(row.manual.fobRejectionDate);
      case 'T': return toTnaDisplayDate(row.manual.fobApprovalDate);
      case 'U': return row.fobDelayDays !== null ? String(row.fobDelayDays) : '';
      case 'V': return toTnaDisplayDate(row.manual.actualDispatchDate);
      case 'W': return row.manual.totalDispatchedQty;
      case 'X': return row.manual.pendingQty;
      case 'Y': return toTnaDisplayDate(row.manual.committedBalanceDate);
      case 'Z': return toTnaDisplayDate(row.manual.actualBalanceDispatchDate);
      case 'AA': return row.manual.balanceDispatchQty;
      case 'AB': return row.manual.orderRemarks;
      case 'AC': return row.balanceDelayDays !== null ? String(row.balanceDelayDays) : '';
      default: return '';
    }
  };

  // Stage summary counts
  const stageCounts = useMemo(() => {
    const counts = { total: allTnaRows.length, Labdip: 0, FOB: 0, Dispatch: 0, 'Balance Fabric': 0, BLANK: 0 };
    allTnaRows.forEach(r => {
      if (r.orderStatus === 'Labdip') counts.Labdip++;
      else if (r.orderStatus === 'FOB') counts.FOB++;
      else if (r.orderStatus === 'Dispatch') counts.Dispatch++;
      else if (r.orderStatus === 'Balance Fabric') counts['Balance Fabric']++;
      else counts.BLANK++;
    });
    return counts;
  }, [allTnaRows]);

  const availableYears = useMemo(() => {
    const set = new Set<string>();
    allTnaRows.forEach(r => {
      const p = parseOrderDateParts(r.orderDate);
      if (p) set.add(String(p.year));
    });
    return Array.from(set).sort((a, b) => Number(b) - Number(a));
  }, [allTnaRows]);

  const filteredRows = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    let list = allTnaRows.filter(row => {
      if (stageFilter === 'BLANK' && row.orderStatus !== '') return false;
      if (stageFilter !== 'ALL' && stageFilter !== 'BLANK' && row.orderStatus !== stageFilter) return false;

      if (monthFilter !== 'ALL' || yearFilter !== 'ALL') {
        const p = parseOrderDateParts(row.orderDate);
        if (!p) return false;
        if (monthFilter !== 'ALL' && String(p.month) !== monthFilter) return false;
        if (yearFilter !== 'ALL' && String(p.year) !== yearFilter) return false;
      }

      if (term) {
        const haystack = [
          row.orderDate,
          row.orderNo,
          row.customerName,
          row.itemName,
          row.color,
          row.width,
          row.unit,
          row.deliveryDate,
          row.orderStatus,
          row.manual.orderRemarks
        ]
          .join(' ')
          .toLowerCase();
        if (!haystack.includes(term)) return false;
      }

      for (const [col, allowed] of Object.entries(colFilters)) {
        if (!allowed || allowed.length === 0) continue;
        const val = getTnaCellValue(row, col);
        if (!allowed.includes(val)) return false;
      }

      return true;
    });

    if (sortConfig) {
      const { colLetter, direction } = sortConfig;
      list = [...list].sort((a, b) => {
        const va = getTnaCellValue(a, colLetter);
        const vb = getTnaCellValue(b, colLetter);
        const cmp = va.localeCompare(vb, undefined, { numeric: true, sensitivity: 'base' });
        return direction === 'asc' ? cmp : -cmp;
      });
    }

    return list;
  }, [allTnaRows, stageFilter, monthFilter, yearFilter, searchTerm, colFilters, sortConfig]);

  const totalPages = pageSize === 0 ? 1 : Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const paginatedRows = useMemo(() => {
    if (pageSize === 0) return filteredRows;
    const start = (safePage - 1) * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, pageSize, safePage]);

  const getUniqueValuesForCol = (colLetter: string): string[] => {
    const set = new Set<string>();
    allTnaRows.forEach(r => {
      const v = getTnaCellValue(r, colLetter);
      if (v !== '') set.add(v);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
  };

  // Pull existing manual TNA data directly from Google Sheet 1Wdu7PtyAEEnOY3R3OgvNM-Hnu9LnTnw1WKJQf06z-Qo via CSV
  const handlePullFromTnaGoogleSheet = async (silent: boolean = false) => {
    if (!silent) {
      setIsPullingSheetCsv(true);
      setCsvPullMsg(null);
    }
    const cacheBust = `&_t=${Date.now()}`;
    const candidateUrls = [
      `https://docs.google.com/spreadsheets/d/${TNA_SPREADSHEET_ID}/export?format=csv${cacheBust}`,
      `https://docs.google.com/spreadsheets/d/${TNA_SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent('TNA')}${cacheBust}`,
      `https://docs.google.com/spreadsheets/d/${TNA_SPREADSHEET_ID}/gviz/tq?tqx=out:csv${cacheBust}`
    ];

    try {
      let csvText = '';
      for (const url of candidateUrls) {
        try {
          const resp = await fetch(url);
          if (resp.ok) {
            const txt = await resp.text();
            if (txt && txt.trim().length > 0 && !txt.startsWith('<!DOCTYPE html>') && !txt.includes('<html')) {
              csvText = txt;
              break;
            }
          }
        } catch {}
      }

      if (!csvText) {
        if (!silent) {
          await onSyncTnaAll(allTnaRows, true, false);
        }
        return;
      }

      const parsed = parseCsvRowsSimple(csvText.trim());
      const importedMap: Record<string, TNAManualFields> = {};
      const occCounter: Record<string, number> = {};
      let matchedCount = 0;

      // Skip 2 header rows (or detect header rows by D header = Item Name)
      let startIdx = 0;
      for (let i = 0; i < Math.min(parsed.length, 4); i++) {
        const dVal = (parsed[i][3] || '').toLowerCase().trim();
        if (dVal.includes('item name') || dVal === '') {
          startIdx = i + 1;
        }
      }

      let lastDate = '';
      let lastOrdNo = '';
      let lastCust = '';

      for (let i = startIdx; i < parsed.length; i++) {
        const r = parsed[i];
        if (!r || r.length < 4) continue;
        const itemName = (r[3] || '').trim();
        if (!itemName) continue;

        const ordDate = (r[0] || '').trim() || lastDate;
        const ordNo = (r[1] || '').trim() || lastOrdNo;
        const cust = (r[2] || '').trim() || lastCust;
        if ((r[2] || '').trim()) {
          lastDate = ordDate;
          lastOrdNo = ordNo;
          lastCust = cust;
        }

        const color = (r[4] || '').trim();
        const width = (r[5] || '').trim();
        const qty = parseFloat(r[7] || '0') || 0;

        const baseSig = `${(ordNo || ordDate).toLowerCase()}|${cust.toLowerCase()}|${itemName.toLowerCase()}|${color.toLowerCase()}|${width.toLowerCase()}|${qty}`;
        const occ = (occCounter[baseSig] || 0) + 1;
        occCounter[baseSig] = occ;
        const key = sanitizeTnaFbKey(`${baseSig}|#${occ}`);
        const altKey = sanitizeTnaFbKey(baseSig);

        const rec: TNAManualFields = {
          labdipCommittedDate: (r[11] || '').trim(),
          labdipActualDate: (r[12] || '').trim(),
          labdipRejectionDate: (r[13] || '').trim(),
          labdipApprovalDate: (r[14] || '').trim(),
          fobCommittedDate: (r[16] || '').trim(),
          fobActualDate: (r[17] || '').trim(),
          fobRejectionDate: (r[18] || '').trim(),
          fobApprovalDate: (r[19] || '').trim(),
          actualDispatchDate: (r[21] || '').trim(),
          totalDispatchedQty: (r[22] || '').trim(),
          pendingQty: (r[23] || '').trim(),
          committedBalanceDate: (r[24] || '').trim(),
          actualBalanceDispatchDate: (r[25] || '').trim(),
          balanceDispatchQty: (r[26] || '').trim(),
          orderRemarks: (r[27] || '').trim(),
          updatedAt: new Date().toISOString()
        };

        const hasAnyData = Object.entries(rec).some(([k, v]) => k !== 'updatedAt' && String(v || '').trim() !== '');
        if (hasAnyData) {
          importedMap[key] = rec;
          importedMap[altKey] = rec;
          matchedCount++;
        }
      }

      if (matchedCount > 0) {
        onBulkMergeTnaRecords(importedMap);
      }
      if (!silent) {
        await onSyncTnaAll(allTnaRows, true, false);
        if (matchedCount > 0) {
          setCsvPullMsg(`Merged ${matchedCount} existing TNA rows from Sheet CSV.`);
        }
      }
    } catch (err: any) {
      if (!silent) {
        setCsvPullMsg(`Sheet pull error: ${err?.message || String(err)}`);
      }
    } finally {
      if (!silent) {
        setIsPullingSheetCsv(false);
      }
    }
  };

  useEffect(() => {
    handlePullFromTnaGoogleSheet(true);
  }, []);

  const handleExportCsv = () => {
    const esc = (v: string) => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const row1 = [
      '', '', '', '', '', '', '', '', '', '', '',
      'Step 1. Lapdip', '', '', '',
      'Delay Days',
      'Step 2. FOB', '', '', '',
      'Delay Days',
      'Step 3. Dispatch', '',
      'Step 4. Balance fabric ( If goods pending )', '', '', '', '',
      'Delay Days'
    ];
    const row2 = TNA_COLUMNS.map(c => (c.kind === 'delay' ? 'Delay Days' : c.header));
    const lines = [
      row1.map(esc).join(','),
      row2.map(esc).join(','),
      ...filteredRows.map(row => TNA_COLUMNS.map(c => esc(getTnaCellValue(row, c.colLetter))).join(','))
    ];
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `GINZA_TNA_Open_Orders_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const renderStageBadge = (status: TNAStageStatus) => {
    if (status === 'Balance Fabric') {
      return (
        <span className="inline-flex items-center px-2 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300 font-bold text-[10px] whitespace-nowrap">
          Balance Fabric
        </span>
      );
    }
    if (status === 'Dispatch') {
      return (
        <span className="inline-flex items-center px-2 py-0.5 rounded bg-emerald-100 text-emerald-900 border border-emerald-300 font-bold text-[10px] whitespace-nowrap">
          Dispatch
        </span>
      );
    }
    if (status === 'FOB') {
      return (
        <span className="inline-flex items-center px-2 py-0.5 rounded bg-sky-100 text-sky-900 border border-sky-300 font-bold text-[10px] whitespace-nowrap">
          FOB
        </span>
      );
    }
    if (status === 'Labdip') {
      return (
        <span className="inline-flex items-center px-2 py-0.5 rounded bg-indigo-100 text-indigo-900 border border-indigo-300 font-bold text-[10px] whitespace-nowrap">
          Labdip
        </span>
      );
    }
    return <span className="text-slate-300 text-[10px]">—</span>;
  };

  const renderDelayBadge = (days: number | null, formulaLabel: string) => {
    if (days === null) {
      return <span className="text-slate-300 text-[10px] font-mono" title={formulaLabel}>—</span>;
    }
    const isDelayed = days > 0;
    return (
      <span
        title={`${formulaLabel} = ${days} day(s)`}
        className={`inline-flex items-center justify-center min-w-[36px] px-1.5 py-0.5 rounded font-mono font-bold text-[11px] border ${
          isDelayed
            ? 'bg-rose-100 text-rose-800 border-rose-300'
            : days === 0
            ? 'bg-slate-100 text-slate-700 border-slate-300'
            : 'bg-emerald-100 text-emerald-800 border-emerald-300'
        }`}
      >
        {days > 0 ? `+${days}d` : `${days}d`}
      </span>
    );
  };

  const groupBannerTheme = (theme: (typeof TNA_GROUP_SPANS)[number]['theme']) => {
    switch (theme) {
      case 'order': return 'bg-slate-900 text-sky-200 border-r border-slate-700';
      case 'labdip': return 'bg-indigo-900 text-indigo-100 border-r border-indigo-700';
      case 'delay': return 'bg-rose-900 text-rose-100 border-r border-rose-700';
      case 'fob': return 'bg-sky-900 text-sky-100 border-r border-sky-700';
      case 'dispatch': return 'bg-emerald-900 text-emerald-100 border-r border-emerald-700';
      case 'balance': return 'bg-amber-900 text-amber-100 border-r border-amber-700';
    }
  };

  const colHeaderTheme = (theme: (typeof TNA_COLUMNS)[number]['groupTheme']) => {
    switch (theme) {
      case 'order': return 'bg-slate-800 text-slate-100 border-r border-slate-700';
      case 'labdip': return 'bg-indigo-950 text-indigo-100 border-r border-indigo-800';
      case 'delay': return 'bg-rose-950 text-rose-100 border-r border-rose-800';
      case 'fob': return 'bg-sky-950 text-sky-100 border-r border-sky-800';
      case 'dispatch': return 'bg-emerald-950 text-emerald-100 border-r border-emerald-800';
      case 'balance': return 'bg-amber-950 text-amber-100 border-r border-amber-800';
    }
  };

  const tnaGasCode = useMemo(
    () => generateTnaGasScript(firebaseUrl, firebaseSecret),
    [firebaseUrl, firebaseSecret]
  );

  return (
    <div className="flex flex-col flex-1 min-h-0">
      {/* TOP TNA TOOLBAR (Single-line scrollable) */}
      <div className="shrink-0 px-2.5 pt-1.5 pb-1.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-2 overflow-x-auto select-none fms-toolbar-scroll whitespace-nowrap">
        <div className="flex items-center gap-2 shrink-0">
          {/* Search */}
          <div className="relative w-48 sm:w-56 shrink-0">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              placeholder="Search TNA Order No, Customer, Item..."
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setPage(1);
              }}
              className="w-full pl-7 pr-6 py-1 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-1 focus:ring-indigo-500 text-slate-800 placeholder-slate-400"
            />
            {searchTerm && (
              <button
                type="button"
                onClick={() => setSearchTerm('')}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-3 h-3" />
              </button>
            )}
          </div>

          {/* Month Filter */}
          <div className="flex items-center gap-1 bg-white px-2 py-1 rounded-md border border-slate-300 shadow-2xs shrink-0">
            <Calendar className="w-3 h-3 text-indigo-600 shrink-0" />
            <span className="text-[11px] font-bold text-slate-600">Month:</span>
            <select
              value={monthFilter}
              onChange={(e) => {
                setMonthFilter(e.target.value);
                setPage(1);
              }}
              className="text-xs font-bold text-slate-800 bg-transparent focus:outline-none cursor-pointer pr-1"
            >
              {MONTH_OPTIONS.map(m => (
                <option key={m.value} value={m.value}>
                  {m.label}
                </option>
              ))}
            </select>
          </div>

          {/* Year Filter */}
          <div className="flex items-center gap-1 bg-white px-2 py-1 rounded-md border border-slate-300 shadow-2xs shrink-0">
            <span className="text-[11px] font-bold text-slate-600">Year:</span>
            <select
              value={yearFilter}
              onChange={(e) => {
                setYearFilter(e.target.value);
                setPage(1);
              }}
              className="text-xs font-bold text-slate-800 bg-transparent focus:outline-none cursor-pointer pr-1"
            >
              <option value="ALL">All Years</option>
              {availableYears.map(yr => (
                <option key={yr} value={yr}>{yr}</option>
              ))}
            </select>
          </div>

          {/* Stage Filter */}
          <div className="flex items-center gap-1 bg-white px-2 py-1 rounded-md border border-slate-300 shadow-2xs shrink-0">
            <Filter className="w-3 h-3 text-slate-500" />
            <span className="text-[11px] font-bold text-slate-600">Stage:</span>
            <select
              value={stageFilter}
              onChange={(e) => {
                setStageFilter(e.target.value as any);
                setPage(1);
              }}
              className="text-xs font-bold text-slate-800 bg-transparent focus:outline-none cursor-pointer pr-1"
            >
              <option value="ALL">All Stages ({stageCounts.total})</option>
              <option value="Labdip">Labdip ({stageCounts.Labdip})</option>
              <option value="FOB">FOB ({stageCounts.FOB})</option>
              <option value="Dispatch">Dispatch ({stageCounts.Dispatch})</option>
              <option value="Balance Fabric">Balance Fabric ({stageCounts['Balance Fabric']})</option>
              <option value="BLANK">Not Started / Blank ({stageCounts.BLANK})</option>
            </select>
          </div>

          {Object.keys(colFilters).length > 0 && (
            <button
              type="button"
              onClick={() => setColFilters({})}
              className="px-2 py-1 text-[11px] font-bold text-rose-700 bg-rose-50 border border-rose-200 rounded hover:bg-rose-100 flex items-center gap-1 cursor-pointer shrink-0"
            >
              <X className="w-3 h-3" />
              <span>Reset Col Filters ({Object.keys(colFilters).length})</span>
            </button>
          )}
        </div>

        {/* Right Controls: Google Sheet Sync + Export + Full Width */}
        <div className="flex items-center gap-1.5 shrink-0">
          <button
            type="button"
            onClick={() => handlePullFromTnaGoogleSheet(false)}
            disabled={isSyncingTna || isPullingSheetCsv}
            className="px-2.5 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded-md text-xs font-bold flex items-center gap-1.5 shadow-2xs disabled:opacity-60 cursor-pointer"
            title={`Sync TNA Tab <-> Google Sheet (${TNA_SPREADSHEET_ID}) & Firebase`}
          >
            {(isSyncingTna || isPullingSheetCsv) ? (
              <Loader2 className="w-3.5 h-3.5 animate-spin" />
            ) : (
              <RefreshCw className="w-3.5 h-3.5" />
            )}
            <span>Sync TNA Sheet</span>
          </button>

          <a
            href={TNA_SHEET_URL}
            target="_blank"
            rel="noopener noreferrer"
            className="px-2.5 py-1 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-md text-xs font-bold flex items-center gap-1 shadow-2xs"
            title={`Open Target Google Sheet ${TNA_SPREADSHEET_ID}`}
          >
            <ExternalLink className="w-3.5 h-3.5 text-emerald-700" />
            <span>Open TNA Sheet</span>
          </a>

          <button
            type="button"
            onClick={handleExportCsv}
            className="px-2.5 py-1 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-md text-xs font-bold flex items-center gap-1 shadow-2xs cursor-pointer"
            title="Download 29-column TNA CSV"
          >
            <Download className="w-3.5 h-3.5 text-slate-600" />
            <span>CSV</span>
          </button>

          {setIsFullWidth && (
            <button
              type="button"
              onClick={() => setIsFullWidth(!isFullWidth)}
              className="px-2 py-1 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-md text-xs font-bold flex items-center gap-1 cursor-pointer"
            >
              {isFullWidth ? <Minimize2 className="w-3.5 h-3.5" /> : <Maximize2 className="w-3.5 h-3.5" />}
            </button>
          )}
        </div>
      </div>

      {/* Compact Sync Notice (only on explicit action) */}
      {(tnaSyncNotice || csvPullMsg) && (
        <div
          className={`shrink-0 px-3 py-1 border-b text-[11px] flex items-center justify-between gap-2 ${
            tnaSyncNotice?.type === 'error'
              ? 'bg-rose-50 border-rose-200 text-rose-800'
              : 'bg-emerald-50 border-emerald-200 text-emerald-800'
          }`}
        >
          <div className="flex items-center gap-1.5">
            <CheckCircle2 className="w-3.5 h-3.5 shrink-0" />
            <span className="font-semibold">
              {tnaSyncNotice?.text || csvPullMsg}
            </span>
          </div>
          {csvPullMsg && !tnaSyncNotice && (
            <button
              type="button"
              onClick={() => setCsvPullMsg(null)}
              className="text-slate-400 hover:text-slate-600 cursor-pointer"
            >
              <X className="w-3 h-3" />
            </button>
          )}
        </div>
      )}

      {/* MAIN 29-COLUMN COMPACT TNA SPREADSHEET GRID */}
      <div className="flex-1 min-h-0 overflow-x-auto overflow-y-auto select-text relative fms-sheet-scroll border-b border-slate-300 bg-white">
        <table className="w-max min-w-full table-fixed border-collapse text-[11px] text-left">
          <colgroup>
            <col className="w-[34px]" />
            <col className="w-[74px]" />
            <col className="w-[92px]" />
            <col className="w-[122px]" />
            <col className="w-[124px]" />
            <col className="w-[72px]" />
            <col className="w-[48px]" />
            <col className="w-[68px]" />
            <col className="w-[54px]" />
            <col className="w-[48px]" />
            <col className="w-[74px]" />
            <col className="w-[84px]" />
            <col className="w-[104px]" />
            <col className="w-[104px]" />
            <col className="w-[104px]" />
            <col className="w-[104px]" />
            <col className="w-[50px]" />
            <col className="w-[104px]" />
            <col className="w-[104px]" />
            <col className="w-[104px]" />
            <col className="w-[104px]" />
            <col className="w-[50px]" />
            <col className="w-[104px]" />
            <col className="w-[60px]" />
            <col className="w-[58px]" />
            <col className="w-[104px]" />
            <col className="w-[104px]" />
            <col className="w-[58px]" />
            <col className="w-[110px]" />
            <col className="w-[50px]" />
          </colgroup>
          <thead className="sticky top-0 z-30 select-none">
            {/* ROW 1: COMPACT CENTER-ALIGNED STAGE GROUP BANNERS */}
            <tr className="text-[9px] font-bold tracking-wide border-b border-slate-700">
              <th className="px-1 py-0.5 text-center bg-slate-950 text-slate-400 border-r border-slate-800 leading-none">
                Row
              </th>
              {TNA_GROUP_SPANS.map(grp => (
                <th
                  key={grp.id}
                  colSpan={grp.colSpan}
                  className={`px-1 py-0.5 text-center uppercase whitespace-nowrap leading-none ${groupBannerTheme(grp.theme)}`}
                >
                  {grp.id === 'order' ? 'Order Details' : grp.label}
                </th>
              ))}
            </tr>

            {/* ROW 2: COMPACT CENTER-ALIGNED HEADERS (NO A..AC LETTERS, ONLY HEADING ROWS FROZEN) */}
            <tr className="text-[9.5px] font-semibold border-b-2 border-slate-950">
              <th className="px-1 py-0.5 text-center bg-slate-900 text-slate-300 border-r border-slate-700 font-mono leading-tight">
                #
              </th>

              {TNA_COLUMNS.map(col => {
                const isFiltered = colFilters[col.colLetter] && colFilters[col.colLetter].length > 0;
                const isSorted = sortConfig?.colLetter === col.colLetter;

                return (
                  <th
                    key={col.colLetter}
                    className={`px-1 py-0.5 text-center whitespace-normal leading-[1.1] align-middle ${colHeaderTheme(col.groupTheme)}`}
                    title={col.header}
                  >
                    <div className="flex items-center justify-center gap-0.5 text-center">
                      <span className="break-words text-center">
                        {col.kind === 'delay' ? 'Delay' : col.header}
                      </span>
                      {isSorted && (
                        <span className="text-emerald-300 font-bold text-[9px]">
                          {sortConfig?.direction === 'asc' ? '↑' : '↓'}
                        </span>
                      )}
                      <button
                        type="button"
                        onClick={() => setActiveFilterCol(activeFilterCol === col.colLetter ? null : col.colLetter)}
                        className={`p-0.5 rounded hover:bg-white/15 cursor-pointer shrink-0 ${
                          isFiltered ? 'text-rose-300 bg-rose-900/60' : 'text-white/55 hover:text-white'
                        }`}
                        title={`Filter or Sort ${col.header}`}
                      >
                        <Filter className="w-2.5 h-2.5" />
                      </button>
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>

          <tbody className="divide-y divide-slate-200 font-sans text-left">
            {paginatedRows.length === 0 ? (
              <tr>
                <td colSpan={30} className="py-12 text-center text-slate-500 bg-white">
                  <Clock className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                  <p className="font-bold text-slate-800 text-sm">No matching Open TNA orders found.</p>
                  <p className="text-xs text-slate-500 mt-1">
                    Only rows from Order Automation with Col AN = &quot;Open&quot; and non-blank Item Name appear here.
                  </p>
                </td>
              </tr>
            ) : (
              paginatedRows.map((row, idx) => {
                const globalRowNum = (safePage - 1) * (pageSize || filteredRows.length) + idx + 3; // Starts at Row 3 like Google Sheet!
                const prevRow = idx > 0 ? paginatedRows[idx - 1] : null;
                const isNewOrderGroup = !prevRow || prevRow.orderKey !== row.orderKey;
                const rowBg = isNewOrderGroup ? 'bg-white' : 'bg-slate-50/40';
                const topBorder = isNewOrderGroup ? 'border-t-2 border-slate-300' : '';

                const updateField = (field: keyof TNAManualFields, val: string) => {
                  onUpdateTnaField(row.tnaKey, row.altKey, field, val);
                };

                return (
                  <tr key={row.tnaKey} className={`${rowBg} ${topBorder} hover:bg-indigo-50/35 transition-colors`}>
                    {/* Sheet Row # (starts at 3) */}
                    <td className="px-1 py-0.5 text-left font-mono text-[10px] text-slate-500 bg-slate-100 border-r border-slate-200 font-semibold">
                      {globalRowNum}
                    </td>

                    {/* Order Date */}
                    <td
                      className="px-1 py-0.5 text-left font-mono text-[10px] text-slate-800 border-r border-slate-200 whitespace-nowrap truncate"
                      title={row.orderDate}
                    >
                      {row.orderDate || '—'}
                    </td>

                    {/* Order No */}
                    <td
                      className="px-1 py-0.5 text-left font-mono text-[10px] font-bold text-indigo-900 border-r border-slate-200 whitespace-nowrap truncate"
                      title={row.orderNo}
                    >
                      {row.orderNo || '—'}
                    </td>

                    {/* Customer Name */}
                    <td className="px-1.5 py-0.5 text-left text-[10.5px] font-semibold text-slate-900 border-r border-slate-200 whitespace-nowrap truncate" title={row.customerName}>
                      {row.customerName}
                    </td>

                    {/* Item Name */}
                    <td className="px-1.5 py-0.5 text-left text-[10.5px] font-medium text-slate-900 border-r border-slate-200 whitespace-nowrap truncate" title={row.itemName}>
                      {row.itemName}
                    </td>

                    {/* Color */}
                    <td className="px-1 py-0.5 text-left text-[10.5px] text-slate-700 border-r border-slate-200 whitespace-nowrap truncate" title={row.color}>
                      {row.color || '—'}
                    </td>

                    {/* Width */}
                    <td className="px-1 py-0.5 text-left text-[10.5px] font-mono text-slate-700 border-r border-slate-200 whitespace-nowrap truncate">
                      {row.width || '—'}
                    </td>

                    {/* Unit */}
                    <td className="px-1 py-0.5 text-left text-[10.5px] text-slate-700 border-r border-slate-200 whitespace-nowrap truncate" title={row.unit}>
                      {row.unit || '—'}
                    </td>

                    {/* Qty */}
                    <td className="px-1 py-0.5 text-left text-[10.5px] font-mono font-bold text-slate-900 border-r border-slate-200 whitespace-nowrap">
                      {row.qty}
                    </td>

                    {/* Unit Of Item */}
                    <td className="px-1 py-0.5 text-left text-[10px] text-slate-700 border-r border-slate-200 whitespace-nowrap truncate">
                      {row.unitOfItem || '—'}
                    </td>

                    {/* Delivery Date */}
                    <td className="px-1 py-0.5 text-left text-[10px] font-mono text-slate-700 border-r border-slate-200 whitespace-nowrap truncate">
                      {row.deliveryDate || '—'}
                    </td>

                    {/* Order Status (LET Formula Output) */}
                    <td className="px-1 py-0.5 text-left border-r border-slate-300 bg-slate-50/70 whitespace-nowrap">
                      {renderStageBadge(row.orderStatus)}
                    </td>

                    {/* Step 1. Lapdip */}
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-indigo-50/15">
                      <TnaDateCell value={row.manual.labdipCommittedDate} onChange={(v) => updateField('labdipCommittedDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-indigo-50/15">
                      <TnaDateCell value={row.manual.labdipActualDate} onChange={(v) => updateField('labdipActualDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-indigo-50/15">
                      <TnaDateCell value={row.manual.labdipRejectionDate} onChange={(v) => updateField('labdipRejectionDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-indigo-50/15">
                      <TnaDateCell value={row.manual.labdipApprovalDate} onChange={(v) => updateField('labdipApprovalDate', v)} />
                    </td>

                    {/* Delay Days (M - O) */}
                    <td className="px-1 py-0.5 text-left border-r border-slate-300 bg-rose-50/25 whitespace-nowrap">
                      {renderDelayBadge(row.labdipDelayDays, 'Labdip Delay = M - O')}
                    </td>

                    {/* Step 2. FOB */}
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-sky-50/15">
                      <TnaDateCell value={row.manual.fobCommittedDate} onChange={(v) => updateField('fobCommittedDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-sky-50/15">
                      <TnaDateCell value={row.manual.fobActualDate} onChange={(v) => updateField('fobActualDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-sky-50/15">
                      <TnaDateCell value={row.manual.fobRejectionDate} onChange={(v) => updateField('fobRejectionDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-sky-50/15">
                      <TnaDateCell value={row.manual.fobApprovalDate} onChange={(v) => updateField('fobApprovalDate', v)} />
                    </td>

                    {/* Delay Days (R - T) */}
                    <td className="px-1 py-0.5 text-left border-r border-slate-300 bg-rose-50/25 whitespace-nowrap">
                      {renderDelayBadge(row.fobDelayDays, 'FOB Delay = R - T')}
                    </td>

                    {/* Step 3. Dispatch */}
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-emerald-50/15">
                      <TnaDateCell value={row.manual.actualDispatchDate} onChange={(v) => updateField('actualDispatchDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-300 bg-emerald-50/15">
                      <input
                        type="number"
                        step="any"
                        placeholder="0"
                        value={row.manual.totalDispatchedQty}
                        onChange={(e) => updateField('totalDispatchedQty', e.target.value)}
                        className="w-full px-1 py-0.5 text-[10.5px] font-mono text-left bg-white border border-slate-300 rounded focus:outline-none focus:ring-1 focus:ring-emerald-500"
                      />
                    </td>

                    {/* Step 4. Balance fabric ( If goods pending ) */}
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-amber-50/15">
                      <input
                        type="number"
                        step="any"
                        placeholder="0"
                        value={row.manual.pendingQty}
                        onChange={(e) => updateField('pendingQty', e.target.value)}
                        className="w-full px-1 py-0.5 text-[10.5px] font-mono text-left bg-white border border-slate-300 rounded focus:outline-none focus:ring-1 focus:ring-amber-500"
                      />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-amber-50/15">
                      <TnaDateCell value={row.manual.committedBalanceDate} onChange={(v) => updateField('committedBalanceDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-amber-50/15">
                      <TnaDateCell value={row.manual.actualBalanceDispatchDate} onChange={(v) => updateField('actualBalanceDispatchDate', v)} />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-200 bg-amber-50/15">
                      <input
                        type="number"
                        step="any"
                        placeholder="0"
                        value={row.manual.balanceDispatchQty}
                        onChange={(e) => updateField('balanceDispatchQty', e.target.value)}
                        className="w-full px-1 py-0.5 text-[10.5px] font-mono text-left bg-white border border-slate-300 rounded focus:outline-none focus:ring-1 focus:ring-amber-500"
                      />
                    </td>
                    <td className="px-0.5 py-0.5 text-left border-r border-slate-300 bg-amber-50/15">
                      <input
                        type="text"
                        placeholder="Remarks..."
                        value={row.manual.orderRemarks}
                        onChange={(e) => updateField('orderRemarks', e.target.value)}
                        className="w-full px-1.5 py-0.5 text-[10.5px] text-left bg-white border border-slate-300 rounded focus:outline-none focus:ring-1 focus:ring-amber-500"
                      />
                    </td>

                    {/* Delay Days (Y - Z) */}
                    <td className="px-1 py-0.5 text-left bg-rose-50/25 whitespace-nowrap">
                      {renderDelayBadge(row.balanceDelayDays, 'Balance Delay = Y - Z')}
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* COMPACT PAGINATION BAR */}
      <div className="shrink-0 flex items-center justify-between px-2.5 py-1 bg-slate-100 border-t border-slate-200 text-[10px] text-slate-700 select-none">
        <div className="flex items-center gap-2">
          <span className="font-semibold text-slate-800">
            TNA Rows: <span className="font-bold text-indigo-700 font-mono">{filteredRows.length.toLocaleString()}</span>
            <span className="text-slate-500 font-normal ml-1">
              (of {allTnaRows.length.toLocaleString()} Open · Pg {safePage}/{totalPages})
            </span>
          </span>
          <div className="flex items-center gap-1 ml-1 pl-2 border-l border-slate-300">
            <span className="text-slate-500">Show:</span>
            <select
              value={pageSize}
              onChange={(e) => {
                setPageSize(Number(e.target.value));
                setPage(1);
              }}
              className="px-1 py-0 h-5 bg-white border border-slate-300 rounded text-[10px] font-semibold text-slate-700 cursor-pointer"
            >
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
              <option value={0}>All</option>
            </select>
          </div>
        </div>

        {pageSize > 0 && (
          <div className="flex items-center gap-1">
            <button
              type="button"
              onClick={() => setPage(1)}
              disabled={safePage <= 1}
              className="px-1.5 h-5 rounded bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 disabled:opacity-40 font-bold cursor-pointer"
            >
              |&laquo;
            </button>
            <button
              type="button"
              onClick={() => setPage(p => Math.max(1, p - 1))}
              disabled={safePage <= 1}
              className="px-1.5 h-5 rounded bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 disabled:opacity-40 flex items-center gap-0.5 font-semibold cursor-pointer"
            >
              <ChevronLeft className="w-3 h-3" /> Prev
            </button>
            <span className="px-2 font-mono font-bold text-slate-800">
              {safePage} / {totalPages}
            </span>
            <button
              type="button"
              onClick={() => setPage(p => Math.min(totalPages, p + 1))}
              disabled={safePage >= totalPages}
              className="px-1.5 h-5 rounded bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 disabled:opacity-40 flex items-center gap-0.5 font-semibold cursor-pointer"
            >
              Next <ChevronRight className="w-3 h-3" />
            </button>
            <button
              type="button"
              onClick={() => setPage(totalPages)}
              disabled={safePage >= totalPages}
              className="px-1.5 h-5 rounded bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 disabled:opacity-40 font-bold cursor-pointer"
            >
              &raquo;|
            </button>
          </div>
        )}
      </div>

      {/* COLUMN FILTER POPOVER */}
      {activeFilterCol && (() => {
        const colMeta = TNA_COLUMNS.find(c => c.colLetter === activeFilterCol);
        const allVals = getUniqueValuesForCol(activeFilterCol);
        const matchedVals = filterSearch.trim()
          ? allVals.filter(v => v.toLowerCase().includes(filterSearch.trim().toLowerCase()))
          : allVals;

        return (
          <div
            className="fixed inset-0 z-50 flex items-center justify-center bg-black/25"
            onClick={() => {
              setActiveFilterCol(null);
              setFilterSearch('');
            }}
          >
            <div
              className="bg-white rounded-xl shadow-2xl border border-slate-300 p-4 w-80 text-xs space-y-2.5"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-center justify-between pb-2 border-b border-slate-200">
                <span className="font-bold text-slate-900">
                  {colMeta?.header}
                </span>
                <button
                  type="button"
                  onClick={() => {
                    setActiveFilterCol(null);
                    setFilterSearch('');
                  }}
                  className="text-slate-400 hover:text-slate-700 cursor-pointer"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setSortConfig({ colLetter: activeFilterCol, direction: 'asc' });
                    setActiveFilterCol(null);
                  }}
                  className="flex-1 py-1 bg-slate-100 hover:bg-slate-200 rounded font-semibold flex items-center justify-center gap-1 cursor-pointer"
                >
                  <ArrowUp className="w-3 h-3 text-indigo-600" /> A → Z
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setSortConfig({ colLetter: activeFilterCol, direction: 'desc' });
                    setActiveFilterCol(null);
                  }}
                  className="flex-1 py-1 bg-slate-100 hover:bg-slate-200 rounded font-semibold flex items-center justify-center gap-1 cursor-pointer"
                >
                  <ArrowDown className="w-3 h-3 text-indigo-600" /> Z → A
                </button>
              </div>

              <div className="relative">
                <Search className="w-3 h-3 absolute left-2 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Search values..."
                  value={filterSearch}
                  onChange={(e) => setFilterSearch(e.target.value)}
                  className="w-full pl-6 pr-2 py-1 text-[11px] bg-slate-50 border border-slate-300 rounded"
                />
              </div>

              <div className="max-h-48 overflow-y-auto bg-slate-50 border border-slate-200 rounded p-2 space-y-1">
                {matchedVals.length === 0 ? (
                  <p className="text-center text-slate-400 italic py-2">No values in this column.</p>
                ) : (
                  matchedVals.slice(0, 150).map(val => {
                    const checked = !colFilters[activeFilterCol] || colFilters[activeFilterCol].includes(val);
                    return (
                      <label key={val} className="flex items-center gap-2 py-0.5 px-1 hover:bg-white rounded cursor-pointer">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => {
                            setColFilters(prev => {
                              const current = prev[activeFilterCol] || allVals;
                              const next = current.includes(val)
                                ? current.filter(v => v !== val)
                                : [...current, val];
                              if (next.length >= allVals.length) {
                                const copy = { ...prev };
                                delete copy[activeFilterCol];
                                return copy;
                              }
                              return { ...prev, [activeFilterCol]: next };
                            });
                            setPage(1);
                          }}
                          className="rounded text-indigo-600"
                        />
                        <span className="truncate">{val}</span>
                      </label>
                    );
                  })
                )}
              </div>

              <div className="flex items-center justify-between pt-1">
                <button
                  type="button"
                  onClick={() => {
                    setColFilters(prev => {
                      const copy = { ...prev };
                      delete copy[activeFilterCol];
                      return copy;
                    });
                    setActiveFilterCol(null);
                  }}
                  className="text-slate-500 underline cursor-pointer"
                >
                  Clear Filter
                </button>
                <button
                  type="button"
                  onClick={() => setActiveFilterCol(null)}
                  className="px-3 py-1 bg-indigo-700 text-white rounded font-bold cursor-pointer"
                >
                  Done
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* GOOGLE SHEET LIVE SYNC SETUP MODAL (1Wdu7PtyAEEnOY3R3OgvNM-Hnu9LnTnw1WKJQf06z-Qo) */}
      {showSheetSetupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col">
            <div className="px-5 py-3.5 bg-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-emerald-400" />
                <h3 className="font-bold text-sm">
                  TNA Live Sync Setup — Google Sheet ({TNA_SPREADSHEET_ID})
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setShowSheetSetupModal(false)}
                className="text-white/70 hover:text-white p-1 rounded cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-4 text-xs text-slate-700">
              <div className="bg-amber-50 border border-amber-300 rounded-xl p-3 space-y-1 text-amber-950">
                <p className="font-bold">
                  Why rows might not appear after deploying:
                </p>
                <ol className="list-decimal list-inside space-y-0.5 text-[11px]">
                  <li>Paste the complete script below into Google Sheet <strong>Extensions &gt; Apps Script</strong> and click <strong>Save</strong>.</li>
                  <li>Click <strong>Deploy &gt; Manage deployments &gt; Edit (pencil) &gt; Version: New version &gt; Deploy</strong> (using the same URL), OR if you used <strong>New deployment</strong>, paste the new <code className="bg-white px-1 rounded font-mono">/exec</code> URL below and click <strong>Save &amp; Sync</strong>.</li>
                  <li>You can also run <strong>TNA Live Sync &gt; 2. Sync Now</strong> directly from the menu inside the Google Sheet.</li>
                </ol>
              </div>

              {/* Web App URL & Live Connection Verification */}
              <div className="space-y-1.5">
                <label className="font-bold text-slate-900 flex items-center gap-1.5">
                  <Link2 className="w-3.5 h-3.5 text-indigo-600" />
                  <span>TNA Apps Script Web App URL</span>
                </label>
                <div className="flex gap-2">
                  <input
                    type="url"
                    value={draftWebAppUrl}
                    onChange={(e) => setDraftWebAppUrl(e.target.value)}
                    placeholder="https://script.google.com/macros/s/.../exec"
                    className="flex-1 px-3 py-1.5 border border-slate-300 rounded-lg font-mono text-[11px] focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                  <button
                    type="button"
                    onClick={() => testWebhookConnection(draftWebAppUrl)}
                    className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold rounded-lg border border-slate-300 cursor-pointer"
                  >
                    Test
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      const clean = draftWebAppUrl.trim();
                      onSaveTnaWebAppUrl(clean);
                      if (clean.startsWith('http')) {
                        testWebhookConnection(clean);
                      }
                      setShowSheetSetupModal(false);
                    }}
                    className="px-4 py-1.5 bg-indigo-700 hover:bg-indigo-800 text-white font-bold rounded-lg cursor-pointer"
                  >
                    Save &amp; Sync
                  </button>
                </div>

                {webhookHealth.status !== 'idle' && (
                  <div
                    className={`px-3 py-2 rounded-lg text-[11px] font-semibold border ${
                      webhookHealth.status === 'ok'
                        ? 'bg-emerald-50 border-emerald-300 text-emerald-900'
                        : webhookHealth.status === 'checking'
                        ? 'bg-slate-50 border-slate-300 text-slate-700'
                        : 'bg-rose-50 border-rose-300 text-rose-900'
                    }`}
                  >
                    {webhookHealth.message}
                  </div>
                )}

                <p className="text-[11px] text-slate-500">
                  Deploy as <strong>Web app (Execute as: Me, Who has access: Anyone)</strong>.
                </p>
              </div>

              {/* K2 LET Formula Copy */}
              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-900">Cell K2 Order Status Formula (=LET)</span>
                  <button
                    type="button"
                    onClick={() => {
                      navigator.clipboard.writeText(TNA_K2_SHEET_FORMULA).then(() => {
                        setCopiedFormula(true);
                        setTimeout(() => setCopiedFormula(false), 2000);
                      });
                    }}
                    className="px-2.5 py-1 bg-slate-800 hover:bg-slate-900 text-white rounded font-bold flex items-center gap-1 cursor-pointer"
                  >
                    {copiedFormula ? <Check className="w-3 h-3 text-emerald-400" /> : <Copy className="w-3 h-3" />}
                    <span>{copiedFormula ? 'Copied K2 Formula!' : 'Copy K2 Formula'}</span>
                  </button>
                </div>
                <pre className="p-2.5 bg-slate-900 text-emerald-300 rounded-lg font-mono text-[10px] overflow-x-auto max-h-36">
                  {TNA_K2_SHEET_FORMULA}
                </pre>
              </div>

              {/* Full TNA Apps Script Copy */}
              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-900">
                    Complete TNA Apps Script (Paste in Extensions &gt; Apps Script, then run <code className="bg-amber-100 px-1 rounded">setupTnaTwoWaySync</code>)
                  </span>
                  <button
                    type="button"
                    onClick={() => {
                      navigator.clipboard.writeText(tnaGasCode).then(() => {
                        setCopiedScript(true);
                        setTimeout(() => setCopiedScript(false), 2500);
                      });
                    }}
                    className="px-3 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded font-bold flex items-center gap-1 cursor-pointer"
                  >
                    {copiedScript ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                    <span>{copiedScript ? 'Copied Script!' : 'Copy Full TNA Script'}</span>
                  </button>
                </div>
                <pre className="p-2.5 bg-slate-950 text-slate-200 rounded-lg font-mono text-[10px] overflow-x-auto max-h-52">
                  {tnaGasCode}
                </pre>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
