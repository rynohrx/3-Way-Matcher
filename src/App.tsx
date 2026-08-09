import React, { useState, useEffect } from 'react';
import { 
  FileText, 
  Upload, 
  CheckCircle, 
  CheckCircle2,
  XCircle, 
  AlertTriangle, 
  Search, 
  RefreshCw, 
  Database, 
  User, 
  ShieldCheck, 
  Mail, 
  Check, 
  TrendingUp, 
  Info,
  Clock,
  FileSpreadsheet,
  FileCheck,
  FileX,
  Plus,
  ChevronRight,
  Trash2,
  FilePlus,
  Download,
  ListFilter
} from 'lucide-react';
import { 
  PurchaseOrder, 
  GoodsReceivedNote, 
  SupplierInvoice, 
  MatchReport,
  AuditedInvoiceItem 
} from './types';
import { GoogleSheetsDb } from './db-mock';
import { runThreeWayMatch, cleanNum, parseExcelDateOrString, compareSupplierNames, resolveQtyAndPrice } from './backend-matcher';
import * as XLSX from 'xlsx';
import { GoogleSheetsModal } from './components/GoogleSheetsModal';

// Sanitize database state to filter out header rows (e.g. PO-PO NUMBER) and duplicate IDs
function sanitizeDb(db: GoogleSheetsDb): GoogleSheetsDb {
  if (!db) return { purchaseOrders: [], goodsReceivedNotes: [], pastInvoices: [] };

  const isHeaderString = (str: string) => {
    if (!str) return true;
    const u = str.toUpperCase().trim();
    return (
      u.includes('NUMBER') ||
      u.includes('HEADER') ||
      u.includes('SUPPLIER') ||
      u.includes('DESCRIPTION') ||
      u.includes('PO NUMBER') ||
      u.includes('GRN NUMBER') ||
      u === 'PO' ||
      u === 'PO-' ||
      u === 'PO-PO' ||
      u === 'PO-PO NUMBER' ||
      u === 'PO-PO NO' ||
      u === 'PO-PO REF' ||
      u === 'PO-PO ID' ||
      u === 'PO-PO#' ||
      u === 'GRN-GRN NUMBER' ||
      u === 'GRN-GRN NO' ||
      u === 'GRN-GRN REF' ||
      u === 'GRN-GRN ID'
    );
  };

  const cleanPos: PurchaseOrder[] = [];
  const poSeen = new Set<string>();
  (db.purchaseOrders || []).forEach(p => {
    if (p && p.id && !isHeaderString(p.id) && !poSeen.has(p.id)) {
      poSeen.add(p.id);
      cleanPos.push(p);
    }
  });

  const cleanGrns: GoodsReceivedNote[] = [];
  const grnSeen = new Set<string>();
  (db.goodsReceivedNotes || []).forEach(g => {
    if (g && g.id && !isHeaderString(g.id) && !grnSeen.has(g.id)) {
      grnSeen.add(g.id);
      cleanGrns.push(g);
    }
  });

  const cleanInvoices: SupplierInvoice[] = [];
  const invSeen = new Set<string>();
  (db.pastInvoices || []).forEach(i => {
    if (i && i.id && !isHeaderString(i.id) && !invSeen.has(i.id)) {
      invSeen.add(i.id);
      cleanInvoices.push(i);
    }
  });

  return {
    purchaseOrders: cleanPos,
    goodsReceivedNotes: cleanGrns,
    pastInvoices: cleanInvoices
  };
}

// Smart Goods Received Note (GRN) row parser
function parseGrnFromRow(strRow: string[]): GoodsReceivedNote | null {
  if (!strRow || strRow.length === 0) return null;

  let rawGrnId = strRow.find(c => c.toUpperCase().startsWith('GRN-')) || strRow.find(c => c.toUpperCase().startsWith('GRN'));
  if (!rawGrnId && strRow[0] && !strRow[0].toUpperCase().startsWith('PO-')) {
    rawGrnId = strRow[0];
  }

  if (rawGrnId) {
    const u = rawGrnId.toUpperCase().trim();
    if (u.includes('NUMBER') || u.includes('HEADER') || u.includes('REF') || u.includes('QTY') || u.includes('DESCRIPTION') || u === 'GRN' || u === 'GRN#') {
      return null;
    }
  }

  let grnId = rawGrnId ? rawGrnId.trim() : '';
  if (grnId && !grnId.toUpperCase().startsWith('GRN-')) {
    grnId = 'GRN-' + grnId;
  }
  if (!grnId || grnId.toUpperCase() === 'GRN-GRN NUMBER') {
    grnId = `GRN-2026-${Math.floor(5000 + Math.random() * 4000)}`;
  }

  let poCell = strRow.find(c => c.toUpperCase().startsWith('PO-')) || '';
  if (!poCell) {
    const cand = strRow.find(c => c !== rawGrnId && (c.includes('100') || c.includes('2026')));
    if (cand && !cand.includes('-') && !isNaN(Number(cand))) {
      poCell = 'PO-' + cand;
    }
  }
  if (!poCell) poCell = 'PO-2026-1001';

  let dateCell = strRow.find(c => c.includes('-') && (c.length === 10 || c.length === 8 || c.length === 9));
  if (!dateCell) {
    const numDate = strRow.find(c => !isNaN(Number(c)) && Number(c) > 25000 && Number(c) < 70000);
    if (numDate) dateCell = numDate;
  }
  const dateReceived = parseExcelDateOrString(dateCell);

  let qtyVal = 1;
  const numCells = strRow.filter(c => !isNaN(Number(c)) && Number(c) > 0 && Number(c) <= 20000 && c !== dateCell);
  if (numCells.length > 0) {
    qtyVal = cleanNum(numCells[0], 1);
  }

  let desc = strRow.find(c => isNaN(Number(c)) && !c.toUpperCase().startsWith('GRN-') && !c.toUpperCase().startsWith('PO-') && !c.includes('-') && c.length > 2) || 'Hardware Item';

  return {
    id: grnId,
    poNumber: poCell,
    dateReceived: dateReceived,
    itemDescription: desc,
    quantityReceived: qtyVal
  };
}

// Smart Purchase Order row parser that handles unaligned Excel columns, dates (including Excel serial numbers like 46218), and vendor names
function parsePoFromRow(strRow: string[]): PurchaseOrder | null {
  if (!strRow || strRow.length === 0) return null;
  
  let rawPoId = strRow.find(c => c.toUpperCase().startsWith('PO-')) 
    || strRow.find(c => c.toUpperCase().startsWith('PO')) 
    || (strRow[0] && strRow[0].trim() !== '' ? strRow[0] : null);
  
  if (!rawPoId) return null;

  const upperRaw = rawPoId.toUpperCase().trim();
  if (
    upperRaw.includes('NUMBER') ||
    upperRaw.includes('HEADER') ||
    upperRaw.includes('SUPPLIER') ||
    upperRaw.includes('DESCRIPTION') ||
    upperRaw.includes('REF') ||
    upperRaw.includes('QTY') ||
    upperRaw.includes('PRICE') ||
    upperRaw === 'PO' ||
    upperRaw === 'PO#' ||
    upperRaw === 'PO NO' ||
    upperRaw === 'PO NUMBER' ||
    upperRaw === 'PO REF' ||
    upperRaw === 'PO ID'
  ) {
    return null;
  }

  let poId = rawPoId.trim();
  if (!poId.toUpperCase().startsWith('PO-')) {
    poId = 'PO-' + poId;
  }

  if (poId.toUpperCase() === 'PO-PO NUMBER' || poId.toUpperCase() === 'PO-PO') {
    return null;
  }

  const otherCells = strRow.filter(c => c !== rawPoId && c.trim() !== '');

  let dateVal = new Date().toISOString().split('T')[0];
  let supplier = 'Hardware Supplier';
  let desc = 'Hardware Item';

  const dates: string[] = [];
  const textStrings: string[] = [];
  const numValues: number[] = [];

  otherCells.forEach(c => {
    const trimmed = c.trim();
    const num = Number(trimmed);

    if (!isNaN(num) && num > 25000 && num < 70000) {
      dates.push(parseExcelDateOrString(trimmed));
    } else if (/^\d{4}-\d{2}-\d{2}/.test(trimmed) || /^\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2,4}/.test(trimmed)) {
      dates.push(trimmed);
    } else if (!isNaN(cleanNum(trimmed, NaN))) {
      numValues.push(cleanNum(trimmed));
    } else if (trimmed.length > 0) {
      textStrings.push(trimmed);
    }
  });

  if (dates.length > 0) {
    dateVal = dates[0];
  }

  if (textStrings.length >= 2) {
    const vendorIdx = textStrings.findIndex(s => /pte|ltd|co|corp|inc|supplies|hardware|vendor|enterprise|trading|supplier/i.test(s));
    if (vendorIdx !== -1) {
      supplier = textStrings[vendorIdx];
      desc = textStrings.filter((_, i) => i !== vendorIdx).join(' ');
    } else {
      supplier = textStrings[0];
      desc = textStrings.slice(1).join(' ');
    }
  } else if (textStrings.length === 1) {
    if (/pte|ltd|co|corp|inc|supplies|hardware|vendor|enterprise|trading|supplier/i.test(textStrings[0])) {
      supplier = textStrings[0];
    } else {
      desc = textStrings[0];
    }
  }

  let rawQ = 1, rawP = 10, rawT = 10;
  if (numValues.length >= 3) {
    rawQ = numValues[0];
    rawP = numValues[1];
    rawT = numValues[2];
  } else if (numValues.length === 2) {
    rawQ = numValues[0];
    rawP = numValues[1];
    rawT = Number((rawQ * rawP).toFixed(2));
  } else if (numValues.length === 1) {
    rawT = numValues[0];
    rawQ = 1;
    rawP = rawT;
  }

  const { qty, unitPrice, totalAmount } = resolveQtyAndPrice(rawQ, rawP, rawT);

  return {
    id: poId,
    supplierName: supplier,
    purchaseDate: dateVal,
    itemDescription: desc,
    quantityOrdered: qty,
    unitPrice: unitPrice,
    totalAmount: totalAmount
  };
}

export default function App() {
  // Live database state (synchronized from Express server)
  const [dbState, setDbState] = useState<GoogleSheetsDb>({
    purchaseOrders: [],
    goodsReceivedNotes: [],
    pastInvoices: []
  });
  
  // Active documents being audited
  const [invoice, setInvoice] = useState<SupplierInvoice | null>(null);
  const [po, setPo] = useState<PurchaseOrder | null>(null);
  const [grn, setGrn] = useState<GoodsReceivedNote | null>(null);
  const [matchReport, setMatchReport] = useState<MatchReport | null>(null);
  
  // System states
  const [loading, setLoading] = useState<boolean>(false);
  const [apiMode, setApiMode] = useState<"active" | "simulation">("simulation");
  const [uploadProgress, setUploadProgress] = useState<string | null>(null);
  const [activeFileName, setActiveFileName] = useState<string>("");

  // Database Tab & Search
  const [dbTab, setDbTab] = useState<'audited' | 'po' | 'grn' | 'past'>('audited');
  const [dbSearch, setDbSearch] = useState<string>("");
  const [dbUploadModal, setDbUploadModal] = useState<boolean>(false);

  // Database Upload / Entry Modals
  const [showPoModal, setShowPoModal] = useState<boolean>(false);
  const [showGrnModal, setShowGrnModal] = useState<boolean>(false);
  const [showBatchModal, setShowBatchModal] = useState<boolean>(false);
  const [showClearModal, setShowClearModal] = useState<boolean>(false);
  const [showGoogleSheetsModal, setShowGoogleSheetsModal] = useState<boolean>(false);
  const [batchRawText, setBatchRawText] = useState<string>("");

  // New PO Form State
  const [newPo, setNewPo] = useState<PurchaseOrder>({
    id: `PO-2026-${Math.floor(1000 + Math.random() * 9000)}`,
    supplierName: "Boon Huat Hardware Supplier",
    purchaseDate: new Date().toISOString().split('T')[0],
    itemDescription: "Brass Water Valves 3/4 Inch",
    quantityOrdered: 100,
    unitPrice: 18.50,
    totalAmount: 1850.00
  });

  // New GRN Form State
  const [newGrn, setNewGrn] = useState<GoodsReceivedNote>({
    id: `GRN-2026-${Math.floor(5000 + Math.random() * 9000)}`,
    poNumber: "",
    dateReceived: new Date().toISOString().split('T')[0],
    itemDescription: "Brass Water Valves 3/4 Inch",
    quantityReceived: 100
  });

  // Invoice Form State
  const [showInvoiceForm, setShowInvoiceForm] = useState<boolean>(false);
  const [manualInvoice, setManualInvoice] = useState<SupplierInvoice>({
    id: `INV-${Math.floor(10000 + Math.random() * 90000)}`,
    invoiceDate: new Date().toISOString().split('T')[0],
    supplierName: "Boon Huat Hardware Supplier",
    poNumber: "",
    itemDescription: "Brass Water Valves 3/4 Inch",
    quantityBilled: 100,
    unitPrice: 18.50,
    totalAmount: 1850.00
  });

  // Action status state
  const [actionStatus, setActionStatus] = useState<string | null>(null);
  const [actionTimestamp, setActionTimestamp] = useState<string | null>(null);

  // Email & Supervisor Modals
  const [showEmailModal, setShowEmailModal] = useState<boolean>(false);
  const [showForwardModal, setShowForwardModal] = useState<boolean>(false);
  const [forwardSupervisor, setForwardSupervisor] = useState<string>("Purchasing Dept (Mr. Tan)");
  const [disputeEmailContent, setDisputeEmailContent] = useState<string>("");
  const [copiedEmail, setCopiedEmail] = useState<boolean>(false);

  // Audited Invoices state for multi-invoice batch analysis & Madam Lim recommendations
  const [auditedInvoices, setAuditedInvoices] = useState<AuditedInvoiceItem[]>([]);
  const [selectedAuditedId, setSelectedAuditedId] = useState<string | null>(null);

  // Helper to construct audited item with 3-way match recommendations
  const createAuditedItem = (
    inv: SupplierInvoice, 
    currentDb: GoogleSheetsDb, 
    existingStatus?: 'APPROVED' | 'REJECTED' | 'REVIEW_SENT' | null
  ): AuditedInvoiceItem => {
    const q = cleanNum(inv.quantityBilled, 1);
    const p = cleanNum(inv.unitPrice, 0);
    let t = cleanNum(inv.totalAmount, 0);
    if (t === 0 && q > 0 && p > 0) {
      t = Number((q * p).toFixed(2));
    }
    const cleanInv: SupplierInvoice = {
      ...inv,
      quantityBilled: q,
      unitPrice: p,
      totalAmount: t
    };

    const isApproved = existingStatus === 'APPROVED';
    const report = runThreeWayMatch(cleanInv, currentDb, undefined, undefined, isApproved);
    
    const normPoNum = cleanInv.poNumber ? cleanInv.poNumber.replace(/[^A-Za-z0-9]/g, '').toUpperCase() : '';
    const referencedPo = normPoNum 
      ? currentDb.purchaseOrders.find(p => p.id.replace(/[^A-Za-z0-9]/g, '').toUpperCase() === normPoNum) || null
      : null;
      
    const normPoId = referencedPo ? referencedPo.id.replace(/[^A-Za-z0-9]/g, '').toUpperCase() : normPoNum;
    const referencedGrn = normPoId 
      ? currentDb.goodsReceivedNotes.find(g => g.poNumber.replace(/[^A-Za-z0-9]/g, '').toUpperCase() === normPoId) || null
      : null;

    let recommendation: 'ACCEPT' | 'REJECT' | 'FORWARD' = 'ACCEPT';
    if (report.overallResult === 'REJECT PAYMENT') {
      recommendation = 'REJECT';
    } else if (report.overallResult === 'MANUAL REVIEW REQUIRED') {
      recommendation = 'FORWARD';
    } else {
      recommendation = 'ACCEPT';
    }

    return {
      id: cleanInv.id,
      invoice: cleanInv,
      po: referencedPo,
      grn: referencedGrn,
      matchReport: report,
      recommendation,
      recommendationSummary: report.explanation,
      actionStatus: existingStatus || null
    };
  };

  const updateAuditedList = (newInvoices: SupplierInvoice[]) => {
    const items = newInvoices.map(inv => {
      const existing = auditedInvoices.find(a => a.id === inv.id);
      return createAuditedItem(inv, dbState, existing?.actionStatus);
    });
    setAuditedInvoices(prev => {
      const combined = [...items];
      prev.forEach(existing => {
        if (!combined.some(c => c.id === existing.id)) {
          combined.push(existing);
        }
      });
      return combined;
    });

    // Automatically select the first discrepancy item or first item for detailed 3-way view
    const priorityItem = items.find(i => i.recommendation === 'REJECT' || i.recommendation === 'FORWARD') || items[0];
    if (priorityItem) {
      selectAuditedInvoiceForView(priorityItem);
    }
  };

  // Re-evaluate 3-Way Match for all sent supplier invoices whenever the DB state (POs / GRNs) updates
  useEffect(() => {
    if (auditedInvoices.length > 0) {
      setAuditedInvoices(prev => 
        prev.map(item => createAuditedItem(item.invoice, dbState, item.actionStatus))
      );
    }
  }, [dbState]);

  const selectAuditedInvoiceForView = (item: AuditedInvoiceItem) => {
    setSelectedAuditedId(item.id);
    setInvoice(item.invoice);
    setPo(item.po);
    setGrn(item.grn);
    setMatchReport(item.matchReport);
    setActionStatus(item.actionStatus);
    setActiveFileName(item.invoice.id);

    setTimeout(() => {
      const elem = document.getElementById('3way-match-report-section');
      if (elem) elem.scrollIntoView({ behavior: 'smooth' });
    }, 100);
  };

  const handleBatchAction = async (item: AuditedInvoiceItem, action: 'approve' | 'dispute' | 'forward') => {
    selectAuditedInvoiceForView(item);
    if (action === 'approve') {
      try {
        const res = await fetch("/api/sheets/add", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ type: "pastInvoices", data: item.invoice })
        });
        const data = await res.json();
        setDbState(data.db);
        setActionStatus("APPROVED");
        setActionTimestamp(new Date().toLocaleTimeString());
        setAuditedInvoices(prev => prev.map(i => i.id === item.id ? { ...i, actionStatus: 'APPROVED' } : i));
        alert(`Payment for Invoice ${item.invoice.id} approved and logged into Boon Huat database!`);
      } catch (err) {
        console.error("Approve failed:", err);
      }
    } else if (action === 'dispute') {
      const inv = item.invoice;
      const poRef = item.po;
      const grnRef = item.grn;
      const report = item.matchReport;
      
      const emailText = `SUBJECT: Payment Withheld / Invoice Discrepancy Notice - ${inv.id}

Dear Accounts Department (${inv.supplierName}),

We are writing regarding Invoice ${inv.id} dated ${inv.invoiceDate} for the sum of $${inv.totalAmount.toFixed(2)}.

During our automated 3-Way Match verification at Boon Huat Hardware & Supplies Pte Ltd, the following discrepancies were identified:

${report.explanation}

${poRef ? `• Approved PO Ref: ${poRef.id} (${poRef.itemDescription}, Qty: ${poRef.quantityOrdered}, Unit Price: $${poRef.unitPrice.toFixed(2)})` : '• Purchase Order: Missing/Not Found'}
${grnRef ? `• Goods Received Note (GRN): ${grnRef.id} (Quantity Received: ${grnRef.quantityReceived})` : '• Goods Received Note: Missing/Not Recorded'}

RECOMMENDED ACTION:
${report.recommendedAction}

Please issue a revised invoice or credit note to align with our approved purchase order and warehouse receiving slip. Payment is currently on hold pending resolution.

Best regards,

Madam Lim (Accounts Executive)
Boon Huat Hardware & Supplies Pte Ltd
Singapore`;

      setDisputeEmailContent(emailText);
      setShowEmailModal(true);
      setActionStatus("REJECTED");
      setActionTimestamp(new Date().toLocaleTimeString());
      setAuditedInvoices(prev => prev.map(i => i.id === item.id ? { ...i, actionStatus: 'REJECTED' } : i));
    } else if (action === 'forward') {
      setShowForwardModal(true);
      setAuditedInvoices(prev => prev.map(i => i.id === item.id ? { ...i, actionStatus: 'REVIEW_SENT' } : i));
    }
  };

  // Fetch initial database & health check on mount
  useEffect(() => {
    fetchDatabase();
    checkHealth();
  }, []);

  const fetchDatabase = async () => {
    try {
      const res = await fetch("/api/sheets");
      const data = await res.json();
      setDbState(sanitizeDb(data));
    } catch (e) {
      console.error("Failed to load database:", e);
    }
  };

  const checkHealth = async () => {
    try {
      const res = await fetch("/api/health");
      const data = await res.json();
      setApiMode(data.geminiMode === "active" ? "active" : "simulation");
    } catch (e) {
      console.error("Failed to check health:", e);
    }
  };

  // Run 3-Way Match against backend
  const runMatchForInvoice = async (inv: SupplierInvoice, manualPo?: PurchaseOrder | null, manualGrn?: GoodsReceivedNote | null) => {
    setLoading(true);
    setInvoice(inv);
    setActionStatus(null);
    
    // Ensure this invoice is in the auditedInvoices list
    updateAuditedList([inv]);

    try {
      const res = await fetch("/api/match", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          invoice: inv,
          manualPo: manualPo ?? undefined,
          manualGrn: manualGrn ?? undefined
        })
      });
      const result = await res.json();
      if (result.success) {
        setPo(result.po);
        setGrn(result.grn);
        setMatchReport(result.report);
      }
    } catch (e) {
      console.error("Error running 3-way match:", e);
    } finally {
      setLoading(false);
    }
  };

  // Flexible helper to parse invoices from an Excel sheet (supporting all Excel formats, headered, headerless, title-prefixed, single invoice template, etc.)
  const parseInvoicesFromExcelSheet = (sheet: XLSX.WorkSheet): SupplierInvoice[] => {
    const extracted: SupplierInvoice[] = [];

    const isHeaderKeyword = (str: string) => {
      if (!str) return false;
      const u = str.toUpperCase().trim();
      return [
        'INVOICE ID', 'INVOICE NO', 'INVOICE NUMBER', 'INV NO', 'INV NUMBER', 'BILL NO', 'BILL NUMBER',
        'PO NUMBER', 'PO NO', 'PO REF', 'PO ID', 'SUPPLIER NAME', 'SUPPLIER', 'VENDOR', 'VENDOR NAME',
        'DESCRIPTION', 'ITEM DESCRIPTION', 'QTY', 'QUANTITY', 'UNIT PRICE', 'PRICE', 'TOTAL', 'TOTAL AMOUNT',
        'DATE', 'INVOICE DATE', 'HEADER', 'ID', 'NO', 'NUMBER', 'REF'
      ].includes(u);
    };

    // 1. Convert sheet to raw 2D array of rows
    const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rawRows || rawRows.length === 0) return extracted;

    // 2. Scan top 25 rows to find the best header row index
    let bestHeaderRowIdx = -1;
    let maxHeaderScore = 0;

    for (let r = 0; r < Math.min(25, rawRows.length); r++) {
      const row = rawRows[r];
      if (!Array.isArray(row)) continue;
      
      let score = 0;
      row.forEach(cell => {
        const str = String(cell || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (
          str.includes('invoice') ||
          str.includes('supplier') ||
          str.includes('vendor') ||
          str.includes('biller') ||
          str.includes('company') ||
          str.includes('ponumber') ||
          str.includes('po#') ||
          str.includes('poref') ||
          str.includes('description') ||
          str.includes('item') ||
          str.includes('qty') ||
          str.includes('quantity') ||
          str.includes('price') ||
          str.includes('unitprice') ||
          str.includes('total') ||
          str.includes('amount') ||
          str.includes('date') ||
          str === 'inv' ||
          str === 'invno' ||
          str === 'bill' ||
          str === 'no' ||
          str === 'id'
        ) {
          score++;
        }
      });

      if (score > maxHeaderScore) {
        maxHeaderScore = score;
        bestHeaderRowIdx = r;
      }
    }

    // 3. Object Row Parsing (Header-based)
    const rangeOpt = (bestHeaderRowIdx >= 0 && maxHeaderScore >= 1) ? { range: bestHeaderRowIdx } : {};
    const objectRows: any[] = XLSX.utils.sheet_to_json(sheet, rangeOpt);

    if (objectRows && objectRows.length > 0) {
      objectRows.forEach((obj, rowIdx) => {
        if (!obj || typeof obj !== 'object') return;
        const keys = Object.keys(obj);
        if (keys.length === 0) return;

        // Key classification helper
        const findVal = (keywords: string[]) => {
          for (const kw of keywords) {
            const kwClean = kw.toLowerCase().replace(/[^a-z0-9]/g, '');
            const exactKey = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === kwClean);
            if (exactKey !== undefined && obj[exactKey] !== undefined && String(obj[exactKey]).trim() !== '') {
              return obj[exactKey];
            }
          }
          for (const kw of keywords) {
            const kwClean = kw.toLowerCase().replace(/[^a-z0-9]/g, '');
            const subKey = keys.find(k => {
              const kClean = k.toLowerCase().replace(/[^a-z0-9]/g, '');
              if (kwClean.length <= 3) {
                return kClean === kwClean || kClean.startsWith(kwClean) || kClean.endsWith(kwClean);
              }
              return kClean.includes(kwClean);
            });
            if (subKey !== undefined && obj[subKey] !== undefined && String(obj[subKey]).trim() !== '') {
              return obj[subKey];
            }
          }
          return undefined;
        };

        const rawInvId = findVal(['invoicenumber', 'invoiceno', 'invoiceid', 'invno', 'invoicenum', 'billno', 'billnumber', 'billid', 'invnumber', 'docno', 'voucherno', 'invoice', 'inv', 'id', 'no', 'number', 'ref']);
        const rawSupplier = findVal(['suppliername', 'supplier', 'vendorname', 'vendor', 'companyname', 'company', 'biller', 'merchant', 'issuer', 'from', 'name']);
        const rawDate = findVal(['invoicedate', 'date', 'billdate', 'createddate', 'issuedate', 'invdate']);
        const rawDueDate = findVal(['duedate', 'paymentduedate', 'due', 'expiry', 'expdate', 'payby', 'paydate', 'paymentdue', 'termsdate', 'dueon']);
        const rawPoRef = findVal(['ponumber', 'poid', 'pono', 'poref', 'purchasenumber', 'po#', 'purchaseorder', 'po']);
        const rawDesc = findVal(['itemdescription', 'itemdesc', 'description', 'item', 'product', 'goods', 'particulars', 'details', 'descriptionofgoods', 'service', 'material']);
        const rawQty = findVal(['quantitybilled', 'quantityordered', 'billedqty', 'quantity', 'qty', 'units', 'count', 'pcs']);
        const rawPrice = findVal(['unitprice', 'priceperunit', 'rateperunit', 'unitrate', 'unitcost', 'itemprice', 'price', 'rate', 'cost']);
        const rawTotal = findVal(['totalamount', 'grandtotal', 'totalprice', 'totalcost', 'totalbill', 'netamount', 'totalvalue', 'total', 'amount']);

        // Check if row has any substantial data
        const hasSomeData = rawInvId !== undefined || rawSupplier !== undefined || rawTotal !== undefined || rawPrice !== undefined || rawDesc !== undefined;

        if (hasSomeData) {
          let invIdStr = rawInvId !== undefined ? String(rawInvId).trim() : '';

          if (invIdStr && isHeaderKeyword(invIdStr)) return;

          if (invIdStr && !invIdStr.toUpperCase().startsWith('INV') && !invIdStr.toUpperCase().startsWith('BILL') && /^\d+$/.test(invIdStr)) {
            invIdStr = 'INV-' + invIdStr;
          }

          if (!invIdStr) {
            invIdStr = `INV-${1000 + rowIdx}`;
          }

          let poRefStr = rawPoRef !== undefined ? String(rawPoRef).trim() : undefined;
          if (poRefStr && !poRefStr.toUpperCase().startsWith('PO-') && /^\d+$/.test(poRefStr)) {
            poRefStr = 'PO-' + poRefStr;
          }

          const { qty, unitPrice, totalAmount } = resolveQtyAndPrice(rawQty, rawPrice, rawTotal);

          if (totalAmount > 0 || unitPrice > 0 || (rawDesc && String(rawDesc).trim() !== '')) {
            if (!extracted.some(i => i.id === invIdStr)) {
              const parsedInvDate = parseExcelDateOrString(rawDate);
              const parsedDueDate = rawDueDate ? parseExcelDateOrString(rawDueDate) : undefined;
              extracted.push({
                id: invIdStr,
                supplierName: String(rawSupplier || 'Hardware Supplier').trim(),
                invoiceDate: parsedInvDate,
                dueDate: parsedDueDate,
                poNumber: poRefStr || undefined,
                itemDescription: String(rawDesc || 'Hardware Materials').trim(),
                quantityBilled: qty || 1,
                unitPrice: unitPrice || 0,
                totalAmount: totalAmount || 0
              });
            }
          }
        }
      });
    }

    // 4. Array Row Scanning (for unheadered rows or fallback)
    if (extracted.length === 0) {
      const startRow = Math.max(0, bestHeaderRowIdx >= 0 ? bestHeaderRowIdx + 1 : 0);
      for (let r = startRow; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!Array.isArray(row)) continue;
        const strRow = row.map(cell => String(cell || '').trim());
        if (strRow.every(c => c === '')) continue;

        if (strRow.some(c => c.toUpperCase().includes('INVOICE NO') || c.toUpperCase().includes('UNIT PRICE') || c.toUpperCase().includes('SUPPLIER NAME'))) continue;

        let invIdCell = strRow.find(c => c.toUpperCase().startsWith('INV-') || c.toUpperCase().startsWith('BILL-') || c.toUpperCase().startsWith('INV')) || strRow[0];
        if (invIdCell && isHeaderKeyword(invIdCell)) invIdCell = '';

        let invId = invIdCell ? invIdCell.trim() : '';
        if (invId && !invId.toUpperCase().startsWith('INV-') && !invId.toUpperCase().startsWith('BILL-') && /^\d+$/.test(invId)) {
          invId = 'INV-' + invId;
        }
        if (!invId) {
          invId = `INV-${1000 + r}`;
        }

        const poRef = strRow.find(c => c.toUpperCase().startsWith('PO-')) || (strRow[3]?.toUpperCase().startsWith('PO-') ? strRow[3] : undefined);

        const numTokens = strRow
          .map(c => cleanNum(c, -99999))
          .filter(val => val !== -99999);

        let qty = 1;
        let price = 10;
        let total = 10;

        if (numTokens.length >= 3) {
          qty = numTokens[0];
          price = numTokens[1];
          total = numTokens[2];
        } else if (numTokens.length === 2) {
          qty = numTokens[0];
          price = numTokens[1];
          total = Number((qty * price).toFixed(2));
        } else if (numTokens.length === 1) {
          total = numTokens[0];
          qty = 1;
          price = total;
        }

        const { qty: finalQ, unitPrice: finalP, totalAmount: finalT } = resolveQtyAndPrice(qty, price, total);

        const textTokens = strRow.filter(c => isNaN(Number(c)) && !c.toUpperCase().startsWith('INV') && !c.toUpperCase().startsWith('PO') && c.length > 1);
        const supplierName = textTokens.find(t => /pte|ltd|co|corp|inc|hardware|supplier|trading|vendor/i.test(t)) || textTokens[0] || 'Hardware Supplier';
        const itemDesc = textTokens.find(t => t !== supplierName) || 'Hardware Item';

        const dateCell = strRow.find(c => c.includes('-') || c.includes('/') || (!isNaN(Number(c)) && Number(c) > 25000 && Number(c) < 70000));

        if (finalT > 0 || finalP > 0) {
          if (!extracted.some(i => i.id === invId)) {
            extracted.push({
              id: invId,
              supplierName,
              invoiceDate: parseExcelDateOrString(dateCell),
              poNumber: poRef,
              itemDescription: itemDesc,
              quantityBilled: finalQ,
              unitPrice: finalP,
              totalAmount: finalT
            });
          }
        }
      }
    }

    // 5. Key-Value Cell Scan (Single Invoice Document Layout)
    if (extracted.length === 0) {
      let invId = '';
      let supplierName = '';
      let invoiceDate = '';
      let poNumber = '';
      let itemDescription = '';
      let totalAmount = 0;
      let unitPrice = 0;
      let quantityBilled = 1;

      for (let r = 0; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!Array.isArray(row)) continue;
        for (let c = 0; c < row.length; c++) {
          const val = String(row[c] || '').trim();
          const valLower = val.toLowerCase();
          const nextVal = String(row[c + 1] || '').trim();

          if (valLower.includes('invoice') && (valLower.includes('no') || valLower.includes('number') || valLower.includes('#') || valLower.includes('id'))) {
            invId = nextVal || val.split(/[:#]/)[1] || '';
          } else if (valLower.includes('supplier') || valLower.includes('vendor') || valLower.includes('biller')) {
            supplierName = nextVal || val.split(/[:]/)[1] || '';
          } else if (valLower.includes('date')) {
            invoiceDate = parseExcelDateOrString(nextVal || val.split(/[:]/)[1] || '');
          } else if (valLower.includes('po') && (valLower.includes('no') || valLower.includes('#') || valLower.includes('ref') || valLower.includes('number'))) {
            poNumber = nextVal || val.split(/[:#]/)[1] || '';
          } else if (valLower.includes('total') || valLower.includes('amount due') || valLower.includes('grand total')) {
            const num = cleanNum(nextVal || val.split(/[:$]/)[1], 0);
            if (num > 0) totalAmount = num;
          }
        }
      }

      if (totalAmount > 0 || supplierName || invId) {
        if (!invId) invId = `INV-2026-${Math.floor(1000 + Math.random() * 9000)}`;
        if (poNumber && !poNumber.toUpperCase().startsWith('PO-') && /^\d+$/.test(poNumber)) poNumber = 'PO-' + poNumber;
        const res = resolveQtyAndPrice(quantityBilled, unitPrice, totalAmount);
        extracted.push({
          id: invId,
          supplierName: supplierName || 'Hardware Supplier',
          invoiceDate: invoiceDate || new Date().toISOString().split('T')[0],
          poNumber: poNumber || undefined,
          itemDescription: itemDescription || 'Hardware Items',
          quantityBilled: res.qty,
          unitPrice: res.unitPrice,
          totalAmount: res.totalAmount
        });
      }
    }

    return extracted;
  };

  // Upload & Audit Supplier Invoices (Excel .xlsx / .xls, CSV, JSON, PDF, Images, Google Sheets export)
  const handleInvoiceUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;
    const files: File[] = Array.from(fileList);
    e.target.value = '';

    setLoading(true);
    setActionStatus(null);
    setUploadProgress(`Processing ${files.length} uploaded invoice file(s)...`);

    const extractedInvoices: SupplierInvoice[] = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const fileLower = file.name.toLowerCase();

      if (fileLower.endsWith('.xlsx') || fileLower.endsWith('.xls') || fileLower.endsWith('.csv')) {
        // Parse Excel sheet / CSV (containing invoices or multi-invoice sheets)
        try {
          const buffer = await file.arrayBuffer();
          const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });

          // Find "Approved Invoices" sheet tab if present
          const approvedSheet = workbook.SheetNames.find(s => s.trim().toLowerCase().includes('approved'));
          const sheetNamesToRead = approvedSheet ? [approvedSheet] : workbook.SheetNames;

          sheetNamesToRead.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const parsedInvs = parseInvoicesFromExcelSheet(sheet);
            parsedInvs.forEach(inv => {
              if (!extractedInvoices.some(e => e.id === inv.id)) {
                extractedInvoices.push(inv);
              }
            });
          });
        } catch (err) {
          console.error("Error reading Excel invoice sheet:", err);
        }

      } else if (fileLower.endsWith('.json') || fileLower.endsWith('.txt')) {
        try {
          const content = await file.text();
          if (fileLower.endsWith('.json')) {
            const parsed = JSON.parse(content);
            const invArray = Array.isArray(parsed) ? parsed : (Array.isArray(parsed.invoices) ? parsed.invoices : [parsed]);
            invArray.forEach((inv: any) => {
              if (inv && inv.id && !extractedInvoices.some(existing => existing.id === inv.id)) {
                const q = cleanNum(inv.quantityBilled, 1);
                const p = cleanNum(inv.unitPrice, 0);
                let t = cleanNum(inv.totalAmount, 0);
                const { qty, unitPrice, totalAmount } = resolveQtyAndPrice(q, p, t);
                extractedInvoices.push({
                  ...inv,
                  quantityBilled: qty,
                  unitPrice: unitPrice,
                  totalAmount: totalAmount
                });
              }
            });
          } else {
            const workbook = XLSX.read(content, { type: 'string' });
            workbook.SheetNames.forEach(sheetName => {
              const sheet = workbook.Sheets[sheetName];
              const parsedInvs = parseInvoicesFromExcelSheet(sheet);
              parsedInvs.forEach(inv => {
                if (!extractedInvoices.some(e => e.id === inv.id)) {
                  extractedInvoices.push(inv);
                }
              });
            });
          }
        } catch (err) {
          console.error("CSV/JSON reading failed:", err);
        }

      } else {
        // PDF / Image file -> API extract
        try {
          const base64 = await new Promise<string>((resolve) => {
            const reader = new FileReader();
            reader.onload = (ev) => resolve(ev.target?.result as string);
            reader.readAsDataURL(file);
          });

          const extractRes = await fetch("/api/extract", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              fileBase64: base64,
              fileName: file.name,
              mimeType: file.type,
              docType: "invoice"
            })
          });
          const extractData = await extractRes.json();
          if (extractData.success && extractData.data) {
            const rawInv = extractData.data;
            const { qty, unitPrice, totalAmount } = resolveQtyAndPrice(
              rawInv.quantityBilled,
              rawInv.unitPrice,
              rawInv.totalAmount
            );
            extractedInvoices.push({
              ...rawInv,
              quantityBilled: qty,
              unitPrice: unitPrice,
              totalAmount: totalAmount
            });
          }
        } catch (err) {
          console.error("PDF/Image extraction failed:", err);
        }
      }
    }

    if (e.target) {
      e.target.value = '';
    }

    if (extractedInvoices.length === 0) {
      alert("No valid supplier invoices could be detected in the uploaded file(s). Please ensure your spreadsheet or document contains invoice records with an Invoice ID, Supplier Name, and Total Billed.");
      setLoading(false);
      setUploadProgress(null);
      return;
    }

    // Run 3-Way Match across all extracted invoices against the Boon Huat database
    updateAuditedList(extractedInvoices);
    setDbTab('audited');
    setLoading(false);
    setUploadProgress(null);

    // Smooth scroll down to audited invoices section
    setTimeout(() => {
      const elem = document.getElementById('invoice-audit-section');
      if (elem) elem.scrollIntoView({ behavior: 'smooth' });
    }, 100);
  };

  // Import and audit invoices from Google Sheets 2D cell array
  const handleGoogleSheetsImport = (rows2D: any[][], fileName: string) => {
    if (!rows2D || rows2D.length === 0) return;
    setLoading(true);
    setActionStatus(null);
    setUploadProgress(`Processing Google Sheet "${fileName}"...`);

    try {
      const sheet = XLSX.utils.aoa_to_sheet(rows2D);
      const parsedInvs = parseInvoicesFromExcelSheet(sheet);

      if (parsedInvs.length === 0) {
        alert(`No valid supplier invoice records could be detected in Google Sheet "${fileName}". Please ensure your sheet contains headers like Invoice ID, Supplier Name, and Total Amount.`);
        setLoading(false);
        setUploadProgress(null);
        return;
      }

      updateAuditedList(parsedInvs);
      setDbTab('audited');
      setLoading(false);
      setUploadProgress(null);

      setTimeout(() => {
        const elem = document.getElementById('invoice-audit-section');
        if (elem) elem.scrollIntoView({ behavior: 'smooth' });
      }, 100);
    } catch (err) {
      console.error("Error importing Google Sheet:", err);
      alert("Failed to parse invoice data from Google Sheet.");
      setLoading(false);
      setUploadProgress(null);
    }
  };

  // Upload PO or GRN document into the database
  const handleDocUploadToDb = (e: React.ChangeEvent<HTMLInputElement>, docType: 'po' | 'grn') => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    setLoading(true);
    setUploadProgress(`Extracting & Saving ${docType.toUpperCase()} document to Boon Huat Database...`);

    const reader = new FileReader();
    reader.onload = async (event) => {
      const base64 = event.target?.result as string;
      try {
        const extractRes = await fetch("/api/extract", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            fileBase64: base64,
            fileName: file.name,
            mimeType: file.type,
            docType
          })
        });
        
        const extractData = await extractRes.json();
        if (extractData.db) {
          setDbState(sanitizeDb(extractData.db));
        } else {
          fetchDatabase();
        }
        alert(`Successfully imported ${docType.toUpperCase()} document into Boon Huat Database!`);
      } catch (err) {
        console.error("Database document upload failed:", err);
      } finally {
        setLoading(false);
        setUploadProgress(null);
      }
    };
    reader.readAsDataURL(file);
  };

  // Upload Database File (Excel .xlsx / .xls, CSV, JSON)
  const handleDatabaseFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    const fileLower = file.name.toLowerCase();

    if (fileLower.endsWith('.xlsx') || fileLower.endsWith('.xls')) {
      // Process Excel Workbook
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const excelBuffer = new Uint8Array(event.target?.result as ArrayBuffer);
          const workbook = XLSX.read(excelBuffer, { type: 'array' });
          
          let importedPos: PurchaseOrder[] = [];
          let importedGrns: GoodsReceivedNote[] = [];

          workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            
            // 1. Check object rows with headers FIRST (e.g. Columns: PO Number / Qty Ordered / Unit Price / Total Amount / Item Description)
            const objectRows: any[] = XLSX.utils.sheet_to_json(sheet);
            objectRows.forEach((obj, idx) => {
              const keys = Object.keys(obj);
              const getVal = (...matchKeys: string[]) => {
                const foundKey = keys.find(k => matchKeys.some(m => k.toLowerCase().replace(/[^a-z0-9]/g, '').includes(m)));
                return foundKey !== undefined ? obj[foundKey] : undefined;
              };
              const getValExcluding = (excludePatterns: string[], ...matchKeys: string[]) => {
                const foundKey = keys.find(k => {
                  const kClean = k.toLowerCase().replace(/[^a-z0-9]/g, '');
                  if (excludePatterns.some(ex => kClean.includes(ex))) return false;
                  return matchKeys.some(m => kClean.includes(m.toLowerCase().replace(/[^a-z0-9]/g, '')));
                });
                return foundKey !== undefined ? obj[foundKey] : undefined;
              };

              const isHeaderVal = (val: any) => {
                if (!val) return false;
                const u = String(val).toUpperCase().trim();
                return u.includes('NUMBER') || u.includes('HEADER') || u.includes('SUPPLIER') || u.includes('DESCRIPTION') || u.includes('REF') || u === 'PO' || u === 'GRN' || u === 'PO-PO' || u === 'PO-PO NUMBER' || u === 'GRN-GRN NUMBER';
              };

              const sheetNameClean = sheetName.toLowerCase().replace(/[^a-z0-9]/g, '');
              const isGrnSheet = sheetNameClean.includes('grn') || sheetNameClean.includes('goodsreceived') || sheetNameClean.includes('receiving') || sheetNameClean.includes('delivery') || sheetNameClean.includes('yard');

              const rawGrnId = getVal('grnnumber', 'grnid', 'grn', 'goodsreceived', 'receipt', 'deliverynote');
              const hasGrnKey = keys.some(k => {
                const kc = k.toLowerCase().replace(/[^a-z0-9]/g, '');
                return kc.includes('grn') || kc.includes('goodsreceived') || kc.includes('qtyreceived') || kc.includes('quantityreceived') || kc.includes('datereceived') || kc.includes('receivedqty');
              });

              const isGrnRecord = isGrnSheet || hasGrnKey || (rawGrnId && !isHeaderVal(rawGrnId));

              if (isGrnRecord) {
                let cleanGrnId = rawGrnId ? String(rawGrnId).trim() : `GRN-2026-${5000 + importedGrns.length + idx + 1}`;
                if (!cleanGrnId.toUpperCase().startsWith('GRN-')) {
                  cleanGrnId = 'GRN-' + cleanGrnId;
                }

                if (!isHeaderVal(cleanGrnId) && !importedGrns.some(g => g.id === cleanGrnId)) {
                  let poNum = String(getVal('ponumber', 'po', 'poref', 'purchasenumber', 'order') || '').trim();
                  if (poNum && !poNum.toUpperCase().startsWith('PO-') && !isNaN(Number(poNum))) {
                    poNum = 'PO-' + poNum;
                  }

                  importedGrns.push({
                    id: cleanGrnId,
                    poNumber: poNum,
                    dateReceived: parseExcelDateOrString(getVal('date', 'datereceived', 'receiveddate', 'deliverydate')),
                    itemDescription: String(getVal('itemdescription', 'itemdesc', 'description', 'item', 'product', 'goods', 'particulars') || 'Hardware Item').trim(),
                    quantityReceived: cleanNum(getVal('quantityreceived', 'qtyreceived', 'receivedqty', 'received', 'quantity', 'qty'), 1)
                  });
                }
              } else {
                const rawPoId = getVal('id', 'ponumber', 'po', 'order', 'purchasenumber', 'po#', 'ref', 'ordernumber', 'num', 'code');
                if ((rawPoId && !isHeaderVal(rawPoId)) || getVal('supplier', 'vendor', 'suppliername', 'item', 'description') || sheetNameClean.includes('po')) {
                  let cleanPoId = rawPoId ? String(rawPoId).trim() : `PO-2026-${1000 + importedPos.length + idx + 1}`;
                  if (!cleanPoId.toUpperCase().startsWith('PO-')) {
                    cleanPoId = 'PO-' + cleanPoId;
                  }

                  if (!isHeaderVal(cleanPoId) && !importedPos.some(p => p.id === cleanPoId)) {
                    const rawSup = String(getVal('supplier', 'vendor', 'suppliername', 'company') || '').trim();
                    const rawDate = String(getVal('date', 'purchasedate', 'podate') || '').trim();
                    
                    let cleanDate = parseExcelDateOrString(rawDate);
                    let cleanSup = rawSup;

                    if (!rawSup || !isNaN(Number(rawSup))) {
                      if (!isNaN(Number(rawSup)) && Number(rawSup) > 25000) {
                        cleanDate = parseExcelDateOrString(rawSup);
                      }
                      cleanSup = String(getVal('company', 'vendorname', 'suppliername', 'name') || 'Hardware Supplier').trim();
                      if (!cleanSup || !isNaN(Number(cleanSup))) {
                        cleanSup = 'Hardware Supplier';
                      }
                    }

                    const rawQ = getValExcluding(['price', 'rate', 'cost', 'total', 'amount'], 'quantityordered', 'qtyordered', 'orderedqty', 'quantity', 'qty', 'units', 'count');
                    const rawP = getValExcluding(['total', 'amount', 'grand', 'net', 'sum', 'qty', 'quantity'], 'unitprice', 'priceperunit', 'rateperunit', 'unitrate', 'unitcost', 'itemprice', 'price', 'rate', 'cost');
                    const rawT = getValExcluding(['unitprice', 'unitrate', 'qty', 'quantity'], 'totalamount', 'grandtotal', 'totalprice', 'totalcost', 'totalbill', 'netamount', 'totalvalue', 'total', 'amount');

                    const { qty, unitPrice, totalAmount } = resolveQtyAndPrice(rawQ, rawP, rawT);

                    importedPos.push({
                      id: cleanPoId,
                      supplierName: cleanSup,
                      purchaseDate: cleanDate,
                      itemDescription: String(getVal('itemdescription', 'itemdesc', 'description', 'item', 'product', 'goods', 'particulars') || 'Hardware Item').trim(),
                      quantityOrdered: qty,
                      unitPrice: unitPrice,
                      totalAmount: totalAmount
                    });
                  }
                }
              }
            });

            // 2. Check array rows as fallback for unheadered data
            const rawRows: any[] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            if (rawRows && rawRows.length > 0) {
              rawRows.forEach(row => {
                if (!Array.isArray(row)) return;
                const strRow = row.map(cell => String(cell || '').trim()).filter(Boolean);
                if (strRow.length === 0) return;

                const firstCell = strRow[0] || '';
                const sNameClean = sheetName.toLowerCase().replace(/[^a-z0-9]/g, '');
                const isGrnContext = sNameClean.includes('grn') || 
                  sNameClean.includes('goodsreceived') || 
                  sNameClean.includes('receiving') || 
                  sNameClean.includes('delivery') || 
                  sNameClean.includes('yard') || 
                  firstCell.toUpperCase().startsWith('GRN-') || 
                  strRow.some(c => c.toUpperCase().startsWith('GRN-'));

                if (isGrnContext) {
                  const parsedGrn = parseGrnFromRow(strRow);
                  if (parsedGrn && !importedGrns.some(g => g.id === parsedGrn.id)) {
                    importedGrns.push(parsedGrn);
                  }
                } else if (firstCell.toUpperCase().startsWith('PO-') || strRow.some(c => c.toUpperCase().startsWith('PO-')) || strRow.length >= 3) {
                  const parsedPo = parsePoFromRow(strRow);
                  if (parsedPo && !importedPos.some(p => p.id === parsedPo.id)) {
                    importedPos.push(parsedPo);
                  }
                }
              });
            }
          });

          if (importedPos.length === 0 && importedGrns.length === 0) {
            alert("Excel sheet read successfully, but no valid records were detected. Please ensure your Excel spreadsheet contains PO or GRN data.");
            return;
          }

          const excelRes = await fetch("/api/sheets/import", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              purchaseOrders: importedPos,
              goodsReceivedNotes: importedGrns,
              replace: false
            })
          });
          const excelData = await excelRes.json();
          setDbState(sanitizeDb(excelData.db));
          setDbTab(importedGrns.length > 0 && importedPos.length === 0 ? 'grn' : 'po');
          alert(`Successfully imported ${importedPos.length} Purchase Orders and ${importedGrns.length} Goods Received Notes into Boon Huat Database!`);
        } catch (err) {
          console.error("Failed to parse Excel file:", err);
          alert("Error reading Excel sheet. Please ensure it is a valid .xlsx or .xls file.");
        }
      };
      reader.readAsArrayBuffer(file);

    } else {
      // Text / CSV / JSON file
      const reader = new FileReader();
      reader.onload = async (event) => {
        const content = event.target?.result as string;
        try {
          let importedPos: PurchaseOrder[] = [];
          let importedGrns: GoodsReceivedNote[] = [];

          if (fileLower.endsWith('.json')) {
            const parsed = JSON.parse(content);
            if (Array.isArray(parsed.purchaseOrders)) importedPos = parsed.purchaseOrders;
            if (Array.isArray(parsed.goodsReceivedNotes)) importedGrns = parsed.goodsReceivedNotes;
          } else {
            // Parse CSV lines
            const lines = content.split('\n').map(l => l.trim()).filter(Boolean);
            lines.forEach(line => {
              const parts = line.split(',').map(p => p.trim());
              const isGrn = parts.some(p => p.toUpperCase().startsWith('GRN-') || p.toUpperCase().startsWith('GRN'));
              if (isGrn) {
                const parsedGrn = parseGrnFromRow(parts);
                if (parsedGrn) importedGrns.push(parsedGrn);
              } else if (parts.length >= 2) {
                const parsedPo = parsePoFromRow(parts);
                if (parsedPo) importedPos.push(parsedPo);
              }
            });
          }

          const res = await fetch("/api/sheets/import", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              purchaseOrders: importedPos,
              goodsReceivedNotes: importedGrns,
              replace: false
            })
          });
          const data = await res.json();
          setDbState(sanitizeDb(data.db));
          setDbTab(importedPos.length > 0 ? 'po' : (importedGrns.length > 0 ? 'grn' : 'po'));
          alert(`Successfully imported ${importedPos.length} Purchase Orders and ${importedGrns.length} Goods Received Notes into Boon Huat Database!`);
        } catch (err) {
          console.error("Failed to parse database file:", err);
          alert("Invalid file format. Please upload a valid Excel (.xlsx), CSV or JSON database file.");
        }
      };
      reader.readAsText(file);
    }
  };

  // Manual PO submission to database
  const handleAddPo = async (e: React.FormEvent) => {
    e.preventDefault();
    const calculatedTotal = Number((newPo.quantityOrdered * newPo.unitPrice).toFixed(2));
    const poToAdd = { ...newPo, totalAmount: calculatedTotal };

    try {
      const res = await fetch("/api/sheets/add", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: "purchaseOrders", data: poToAdd })
      });
      const data = await res.json();
      setDbState(sanitizeDb(data.db));
      setDbTab('po');
      setShowPoModal(false);
      // Reset form
      setNewPo({
        id: `PO-2026-${Math.floor(1000 + Math.random() * 9000)}`,
        supplierName: "Boon Huat Hardware Supplier",
        purchaseDate: new Date().toISOString().split('T')[0],
        itemDescription: "Hardware Fittings",
        quantityOrdered: 100,
        unitPrice: 10.00,
        totalAmount: 1000.00
      });
      alert(`Successfully added Purchase Order ${poToAdd.id} to Boon Huat Database!`);
    } catch (e) {
      console.error("Failed to add PO:", e);
    }
  };

  // Manual GRN submission to database
  const handleAddGrn = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      const res = await fetch("/api/sheets/add", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: "goodsReceivedNotes", data: newGrn })
      });
      const data = await res.json();
      setDbState(sanitizeDb(data.db));
      setDbTab('grn');
      setShowGrnModal(false);
      setNewGrn({
        id: `GRN-2026-${Math.floor(5000 + Math.random() * 9000)}`,
        poNumber: "",
        dateReceived: new Date().toISOString().split('T')[0],
        itemDescription: "Hardware Goods",
        quantityReceived: 100
      });
      alert(`Successfully added GRN ${newGrn.id} to Boon Huat Database!`);
    } catch (e) {
      console.error("Failed to add GRN:", e);
    }
  };

  // Batch paste text import to database
  const handleBatchTextImport = async () => {
    if (!batchRawText.trim()) return;

    const lines = batchRawText.split('\n').map(l => l.trim()).filter(Boolean);
    const pos: PurchaseOrder[] = [];
    const grns: GoodsReceivedNote[] = [];

    lines.forEach(line => {
      const parts = line.split(',').map(p => p.trim());
      const isGrn = parts.some(p => p.toUpperCase().startsWith('GRN-') || p.toUpperCase().startsWith('GRN'));
      if (isGrn) {
        const parsedGrn = parseGrnFromRow(parts);
        if (parsedGrn) grns.push(parsedGrn);
      } else if (parts.length >= 2) {
        const parsedPo = parsePoFromRow(parts);
        if (parsedPo) pos.push(parsedPo);
      }
    });

    try {
      const res = await fetch("/api/sheets/import", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ purchaseOrders: pos, goodsReceivedNotes: grns, replace: false })
      });
      const data = await res.json();
      setDbState(sanitizeDb(data.db));
      setDbTab(pos.length > 0 ? 'po' : (grns.length > 0 ? 'grn' : 'po'));
      setShowBatchModal(false);
      setBatchRawText("");
      alert(`Imported ${pos.length} Purchase Orders and ${grns.length} GRNs to Boon Huat Database!`);
    } catch (e) {
      console.error("Failed batch import:", e);
    }
  };

  // Clear Database
  const handleClearDatabase = () => {
    setShowClearModal(true);
  };

  const handleConfirmClearDatabase = async () => {
    try {
      const res = await fetch("/api/sheets/clear", { method: "POST" });
      const data = await res.json();
      setDbState(sanitizeDb(data.db || { purchaseOrders: [], goodsReceivedNotes: [], pastInvoices: [] }));
      setDbTab('po');
      setInvoice(null);
      setPo(null);
      setGrn(null);
      setMatchReport(null);
      setAuditedInvoices([]);
      setSelectedAuditedId(null);
      setActionStatus(null);
      setShowClearModal(false);
    } catch (e) {
      console.error("Failed to clear database:", e);
    }
  };

  const handleResetSampleDatabase = async () => {
    try {
      const res = await fetch("/api/sheets/reset", { method: "POST" });
      const data = await res.json();
      setDbState(sanitizeDb(data.db));
      setDbTab('po');
      setShowClearModal(false);
    } catch (e) {
      console.error("Failed to reset database:", e);
    }
  };

  // Allow user / Madam Lim to edit PO supplier name directly
  const handleEditPoSupplierName = async (targetPo: PurchaseOrder) => {
    const currentSup = targetPo.supplierName ? String(targetPo.supplierName).trim() : '';
    const defaultVal = (!currentSup || !isNaN(Number(currentSup))) ? (invoice ? invoice.supplierName : 'Lian Seng Hardware Supplies Pte Ltd') : currentSup;
    const newName = prompt(`Enter corrected Supplier Name for Purchase Order ${targetPo.id}:`, defaultVal);
    if (newName && newName.trim() !== '') {
      const updatedPo = { ...targetPo, supplierName: newName.trim() };
      
      const newDbPos = dbState.purchaseOrders.map(p => p.id === targetPo.id ? updatedPo : p);
      const updatedDbState = { ...dbState, purchaseOrders: newDbPos };
      setDbState(updatedDbState);

      if (po && po.id === targetPo.id) {
        setPo(updatedPo);
        if (invoice) {
          const report = runThreeWayMatch(invoice, updatedDbState, updatedPo, grn);
          setMatchReport(report);
        }
      }

      try {
        await fetch("/api/sheets/add", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ type: "purchaseOrders", data: updatedPo })
        });
      } catch (err) {
        console.error("Failed to update PO supplier name on backend:", err);
      }
    }
  };

  // Allow user / Madam Lim to edit PO unit price directly
  const handleEditPoUnitPrice = async (targetPo: PurchaseOrder) => {
    const currentPrice = targetPo.unitPrice || 0;
    const newPriceStr = prompt(`Enter corrected Unit Price ($) for Purchase Order ${targetPo.id}:`, String(currentPrice));
    if (newPriceStr !== null && newPriceStr.trim() !== '') {
      const newPrice = cleanNum(newPriceStr, currentPrice);
      const newTotal = Number((targetPo.quantityOrdered * newPrice).toFixed(2));
      const updatedPo = { ...targetPo, unitPrice: newPrice, totalAmount: newTotal };
      
      const newDbPos = dbState.purchaseOrders.map(p => p.id === targetPo.id ? updatedPo : p);
      const updatedDbState = { ...dbState, purchaseOrders: newDbPos };
      setDbState(updatedDbState);

      if (po && po.id === targetPo.id) {
        setPo(updatedPo);
        if (invoice) {
          const report = runThreeWayMatch(invoice, updatedDbState, updatedPo, grn);
          setMatchReport(report);
        }
      }

      try {
        await fetch("/api/sheets/add", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ type: "purchaseOrders", data: updatedPo })
        });
      } catch (err) {
        console.error("Failed to update PO unit price on backend:", err);
      }
    }
  };

  // Allow user / Madam Lim to edit PO quantity ordered directly
  const handleEditPoQuantity = async (targetPo: PurchaseOrder) => {
    const currentQty = targetPo.quantityOrdered || 1;
    const newQtyStr = prompt(`Enter corrected Quantity Ordered for Purchase Order ${targetPo.id}:`, String(currentQty));
    if (newQtyStr !== null && newQtyStr.trim() !== '') {
      const newQty = cleanNum(newQtyStr, currentQty);
      const newTotal = Number((newQty * targetPo.unitPrice).toFixed(2));
      const updatedPo = { ...targetPo, quantityOrdered: newQty, totalAmount: newTotal };
      
      const newDbPos = dbState.purchaseOrders.map(p => p.id === targetPo.id ? updatedPo : p);
      const updatedDbState = { ...dbState, purchaseOrders: newDbPos };
      setDbState(updatedDbState);

      if (po && po.id === targetPo.id) {
        setPo(updatedPo);
        if (invoice) {
          const report = runThreeWayMatch(invoice, updatedDbState, updatedPo, grn);
          setMatchReport(report);
        }
      }

      try {
        await fetch("/api/sheets/add", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ type: "purchaseOrders", data: updatedPo })
        });
      } catch (err) {
        console.error("Failed to update PO quantity on backend:", err);
      }
    }
  };

  // Export current database to Excel (.xlsx)
  const handleExportDatabase = () => {
    try {
      const wb = XLSX.utils.book_new();
      
      // Convert Purchase Orders
      const poSheetData = dbState.purchaseOrders.map(p => ({
        "PO Number": p.id,
        "Supplier Name": p.supplierName,
        "Purchase Date": p.purchaseDate,
        "Item Description": p.itemDescription,
        "Quantity Ordered": p.quantityOrdered,
        "Unit Price ($)": p.unitPrice,
        "Total Amount ($)": p.totalAmount
      }));
      const poWs = XLSX.utils.json_to_sheet(poSheetData.length > 0 ? poSheetData : [{ "PO Number": "PO-SAMPLE", "Supplier Name": "Sample Supplier", "Purchase Date": "2026-08-01", "Item Description": "Sample Item", "Quantity Ordered": 100, "Unit Price ($)": 10.00, "Total Amount ($)": 1000.00 }]);
      XLSX.utils.book_append_sheet(wb, poWs, "Purchase Orders");

      // Convert Goods Received Notes
      const grnSheetData = dbState.goodsReceivedNotes.map(g => ({
        "GRN Number": g.id,
        "PO Reference": g.poNumber,
        "Date Received": g.dateReceived,
        "Item Description": g.itemDescription,
        "Quantity Received": g.quantityReceived
      }));
      const grnWs = XLSX.utils.json_to_sheet(grnSheetData.length > 0 ? grnSheetData : [{ "GRN Number": "GRN-SAMPLE", "PO Reference": "PO-SAMPLE", "Date Received": "2026-08-01", "Item Description": "Sample Item", "Quantity Received": 100 }]);
      XLSX.utils.book_append_sheet(wb, grnWs, "Goods Received Notes");

      XLSX.writeFile(wb, `boon_huat_database_${new Date().toISOString().split('T')[0]}.xlsx`);
    } catch (err) {
      console.error("Export error:", err);
      // Fallback to JSON
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(dbState, null, 2));
      const downloadAnchor = document.createElement('a');
      downloadAnchor.setAttribute("href", dataStr);
      downloadAnchor.setAttribute("download", `boon_huat_database_${new Date().toISOString().split('T')[0]}.json`);
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
    }
  };

  // Submit manual invoice form to trigger 3-Way Match
  const handleManualInvoiceSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setShowInvoiceForm(false);
    setActiveFileName("Manual Entry");
    
    const calculatedTotal = Number((manualInvoice.quantityBilled * manualInvoice.unitPrice).toFixed(2));
    const invToMatch: SupplierInvoice = {
      ...manualInvoice,
      totalAmount: calculatedTotal
    };
    
    updateAuditedList([invToMatch]);
    setDbTab('audited');
    runMatchForInvoice(invToMatch);

    setTimeout(() => {
      const elem = document.getElementById('invoice-audit-section');
      if (elem) elem.scrollIntoView({ behavior: 'smooth' });
    }, 100);
  };

  // Audit specific PO from table by creating a matching test invoice
  const auditPoFromTable = (selectedPo: PurchaseOrder) => {
    const matchingGrn = dbState.goodsReceivedNotes.find(g => g.poNumber.toUpperCase() === selectedPo.id.toUpperCase());
    const testInvoice: SupplierInvoice = {
      id: `INV-${Math.floor(10000 + Math.random() * 90000)}`,
      invoiceDate: new Date().toISOString().split('T')[0],
      supplierName: selectedPo.supplierName,
      poNumber: selectedPo.id,
      itemDescription: selectedPo.itemDescription,
      quantityBilled: selectedPo.quantityOrdered,
      unitPrice: selectedPo.unitPrice,
      totalAmount: selectedPo.totalAmount
    };

    setActiveFileName(`Linked from DB: ${selectedPo.id}`);
    runMatchForInvoice(testInvoice, selectedPo, matchingGrn || null);
    window.scrollTo({ top: 600, behavior: 'smooth' });
  };

  // Action Handlers
  const handleApprove = async () => {
    if (!invoice) return;
    const timestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    setActionStatus("APPROVED");
    setActionTimestamp(timestamp);
    setAuditedInvoices(prev => prev.map(i => i.id === invoice.id ? { ...i, actionStatus: 'APPROVED' } : i));

    try {
      await fetch("/api/sheets/add", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "pastInvoices",
          data: invoice
        })
      });
      fetchDatabase();
    } catch (e) {
      console.error("Failed to update database:", e);
    }

    // Trigger Google Sheets Export modal after approval
    setTimeout(() => {
      setShowGoogleSheetsModal(true);
    }, 300);
  };

  const handleReject = () => {
    if (!invoice || !matchReport) return;
    
    const failedChecks = (Object.values(matchReport.checks) as any[]).filter(c => c.status === 'FAIL');
    const reasons = failedChecks.map(c => `• ${c.name}: ${c.explanation || c.details}`).join('\n');
    
    const emailBody = `Dear ${invoice.supplierName} Accounts Team,

We are writing regarding your invoice #${invoice.id} dated ${invoice.invoiceDate} for the total amount of $${invoice.totalAmount.toFixed(2)}.

During our 3-Way Matching audit against our Boon Huat Purchase Order database and receiving logs, the following discrepancies were flagged:

${reasons || '• Invoice details do not match authorized Purchase Order or Goods Received records.'}

As a result, payment for invoice #${invoice.id} has been placed on HOLD pending your clarification. Please review and provide an amended credit note or revised invoice.

Best regards,
Madam Lim
Accounts Payable Department
Boon Huat Hardware & Supplies Pte Ltd
Tel: +65 6748 1122`;

    setDisputeEmailContent(emailBody);
    setShowEmailModal(true);
    
    const timestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    setActionStatus("REJECTED");
    setActionTimestamp(timestamp);
    setAuditedInvoices(prev => prev.map(i => i.id === invoice.id ? { ...i, actionStatus: 'REJECTED' } : i));
  };

  const submitForwardReview = () => {
    setShowForwardModal(false);
    const timestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    setActionStatus("REVIEW_SENT");
    setActionTimestamp(timestamp);
    if (invoice) {
      setAuditedInvoices(prev => prev.map(i => i.id === invoice.id ? { ...i, actionStatus: 'REVIEW_SENT' } : i));
    }
  };

  // Filtered DB items for table
  const getFilteredDb = () => {
    const search = dbSearch.toLowerCase().trim();
    if (dbTab === 'audited') {
      return auditedInvoices.filter(item => 
        !search ||
        item.id.toLowerCase().includes(search) ||
        item.invoice.supplierName.toLowerCase().includes(search) ||
        item.invoice.itemDescription.toLowerCase().includes(search) ||
        (item.invoice.poNumber && item.invoice.poNumber.toLowerCase().includes(search)) ||
        item.recommendation.toLowerCase().includes(search)
      );
    } else if (dbTab === 'po') {
      return dbState.purchaseOrders.filter(p => 
        !search || 
        p.id.toLowerCase().includes(search) || 
        p.supplierName.toLowerCase().includes(search) || 
        p.itemDescription.toLowerCase().includes(search)
      );
    } else if (dbTab === 'grn') {
      return dbState.goodsReceivedNotes.filter(g => 
        !search || 
        g.id.toLowerCase().includes(search) || 
        g.poNumber.toLowerCase().includes(search) || 
        g.itemDescription.toLowerCase().includes(search)
      );
    } else {
      return dbState.pastInvoices.filter(i => 
        !search || 
        i.id.toLowerCase().includes(search) || 
        i.supplierName.toLowerCase().includes(search) || 
        i.itemDescription.toLowerCase().includes(search)
      );
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans flex flex-col selection:bg-slate-200">
      
      {/* BRAND & SECURITY HEADER */}
      <header className="bg-slate-900 text-white border-b border-slate-800 py-4 px-6 sticky top-0 z-40 shadow-sm">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          
          <div className="flex items-center gap-3">
            <div className="h-10 w-10 bg-amber-500 rounded flex items-center justify-center shadow-inner">
              <span className="font-display font-bold text-slate-900 text-xl tracking-tighter">BH</span>
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="font-display font-bold text-lg tracking-tight">Boon Huat Hardware & Supplies Pte Ltd</h1>
                <span className="text-xs bg-slate-800 text-slate-300 font-mono py-0.5 px-2 rounded border border-slate-700">ERP LINKER</span>
              </div>
              <p className="text-xs text-slate-400">AP A.I. 3-Way Matching Engine & PO Database Hub</p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3 text-xs">
            <div className="flex items-center gap-1.5 bg-slate-800 py-1.5 px-3 rounded border border-slate-700 text-slate-300">
              <Database className="w-3.5 h-3.5 text-indigo-400" />
              <span>Boon Huat Database: <strong>{dbState.purchaseOrders.length} POs | {dbState.goodsReceivedNotes.length} GRNs</strong></span>
            </div>
            <div className="flex items-center gap-1.5 bg-slate-800 py-1.5 px-3 rounded border border-slate-700 text-slate-300">
              <ShieldCheck className="w-3.5 h-3.5 text-amber-500" />
              <span>Engine: <strong>{apiMode === 'active' ? 'Gemini 3.5 Flash' : 'Standard Parser'}</strong></span>
            </div>
            <div className="flex items-center gap-1.5 bg-amber-500/10 text-amber-300 py-1.5 px-3 rounded border border-amber-500/20">
              <User className="w-3.5 h-3.5" />
              <span>Auditor: <strong>Madam Lim</strong></span>
            </div>
          </div>
        </div>
      </header>

      {/* MAIN CONTAINER */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-6 space-y-8">
        
        {/* SECTION 1: BOON HUAT DATABASE HUB */}
        <section className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
          <div className="bg-slate-900 text-white p-4 sm:p-5 flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-800">
            <div>
              <div className="flex items-center gap-2">
                <Database className="w-5 h-5 text-amber-400" />
                <h2 className="font-display font-bold text-base text-white">1. Boon Huat Purchase Order & GRN Database Repository</h2>
              </div>
              <p className="text-xs text-slate-300 mt-1">
                Upload or manage your official Purchase Orders (PO) and Goods Received Notes (GRN). When an invoice is sent, the system automatically links back to this database.
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                onClick={() => setShowPoModal(true)}
                className="text-xs font-semibold py-1.5 px-3 bg-amber-500 hover:bg-amber-400 text-slate-950 rounded flex items-center gap-1.5 transition-all shadow-xs"
              >
                <Plus className="w-3.5 h-3.5" />
                Add PO to DB
              </button>
              <button
                onClick={() => setShowGrnModal(true)}
                className="text-xs font-semibold py-1.5 px-3 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded flex items-center gap-1.5 transition-all"
              >
                <Plus className="w-3.5 h-3.5" />
                Add GRN to DB
              </button>
              <button
                onClick={() => setShowBatchModal(true)}
                className="text-xs font-semibold py-1.5 px-3 bg-indigo-600 hover:bg-indigo-500 text-white rounded flex items-center gap-1.5 transition-all"
              >
                <FilePlus className="w-3.5 h-3.5" />
                Paste / Batch CSV
              </button>
              <button
                onClick={handleExportDatabase}
                disabled={dbState.purchaseOrders.length === 0 && dbState.goodsReceivedNotes.length === 0}
                className="text-xs font-medium py-1.5 px-3 bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 rounded flex items-center gap-1 disabled:opacity-50"
              >
                <Download className="w-3.5 h-3.5" />
                Export DB
              </button>
              <button
                onClick={handleClearDatabase}
                className="text-xs font-medium py-1.5 px-3 bg-slate-800 hover:bg-red-950/60 text-red-400 hover:text-red-300 border border-slate-700 rounded flex items-center gap-1 transition-colors"
              >
                <Trash2 className="w-3.5 h-3.5" />
                Clear DB
              </button>
            </div>
          </div>

          <div className="p-5 space-y-5">
            
            {/* Database File Upload Card (Full width master database uploader) */}
            <div className="border border-indigo-200 rounded-lg p-4 bg-indigo-50/50 hover:bg-indigo-50 transition-all flex flex-col md:flex-row md:items-center justify-between gap-4 relative">
              <input 
                type="file" 
                accept=".xlsx,.xls,.csv,.json,.txt"
                onChange={handleDatabaseFileUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
              />
              <div className="flex items-start gap-3">
                <div className="p-2.5 bg-indigo-600 text-white rounded-lg shrink-0 mt-0.5 shadow-xs">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-sm text-indigo-950 flex flex-wrap items-center gap-2">
                    Upload Master Database Spreadsheet (Excel .xlsx / CSV / JSON)
                    <span className="text-[10px] bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded font-mono font-medium">All POs & GRNs Bunched Together</span>
                  </h3>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Upload your master Excel spreadsheet (.xlsx, .xls) or CSV containing all Boon Huat Purchase Orders (POs) and Goods Received Notes (GRNs). The system automatically parses both sheets/rows into the live database repository.
                  </p>
                </div>
              </div>
              <div className="shrink-0">
                <button className="text-xs font-semibold py-2 px-4 bg-indigo-600 hover:bg-indigo-500 text-white rounded shadow-xs flex items-center gap-1.5 transition-all">
                  <Upload className="w-4 h-4" />
                  Select Master Database Excel File
                </button>
              </div>
            </div>

            {/* LIVE DATABASE TABLE VIEW */}
            <div className="border border-slate-200 rounded-lg overflow-hidden">
              <div className="bg-slate-100 border-b border-slate-200 p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                
                {/* Tabs */}
                <div className="flex flex-wrap items-center gap-1 bg-white p-1 rounded border border-slate-200">
                  <button
                    onClick={() => setDbTab('audited')}
                    className={`px-3 py-1 rounded font-semibold transition-all flex items-center gap-1.5 ${
                      dbTab === 'audited' ? 'bg-amber-500 text-slate-950 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    <FileCheck className="w-3.5 h-3.5" />
                    Sent Supplier Invoices ({auditedInvoices.length})
                  </button>
                  <button
                    onClick={() => setDbTab('po')}
                    className={`px-3 py-1 rounded font-semibold transition-all ${
                      dbTab === 'po' ? 'bg-slate-900 text-white shadow-xs' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Purchase Orders ({dbState.purchaseOrders.length})
                  </button>
                  <button
                    onClick={() => setDbTab('grn')}
                    className={`px-3 py-1 rounded font-semibold transition-all ${
                      dbTab === 'grn' ? 'bg-slate-900 text-white shadow-xs' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Goods Received ({dbState.goodsReceivedNotes.length})
                  </button>
                  <button
                    onClick={() => setDbTab('past')}
                    className={`px-3 py-1 rounded font-semibold transition-all ${
                      dbTab === 'past' ? 'bg-slate-900 text-white shadow-xs' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Approved History ({dbState.pastInvoices.length})
                  </button>
                </div>

                {/* Search */}
                <div className="relative">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                  <input
                    type="text"
                    placeholder="Search database..."
                    value={dbSearch}
                    onChange={(e) => setDbSearch(e.target.value)}
                    className="pl-8 pr-3 py-1.5 bg-white border border-slate-300 rounded text-xs focus:ring-1 focus:ring-slate-800 outline-none w-full sm:w-48"
                  />
                </div>

              </div>

              {/* Table Data */}
              <div className="overflow-x-auto max-h-72">
                {dbTab === 'audited' ? (
                  auditedInvoices.length === 0 ? (
                    <div className="p-8 text-center text-slate-500 text-xs bg-slate-50">
                      <FileX className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                      <p className="font-semibold text-slate-700">No Sent Supplier Invoices Loaded Yet</p>
                      <p className="text-slate-400 mt-1 max-w-md mx-auto">
                        Upload a batch Excel / CSV spreadsheet in <strong>Section 2</strong> or enter an invoice manually to perform 3-Way Match against the Boon Huat database.
                      </p>
                    </div>
                  ) : (
                    <table className="w-full text-left text-xs text-slate-700">
                      <thead className="bg-slate-100 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider text-[10px] sticky top-0">
                        <tr>
                          <th className="p-2.5">Invoice ID</th>
                          <th className="p-2.5">Supplier Name</th>
                          <th className="p-2.5">Linked PO Ref</th>
                          <th className="p-2.5">Linked GRN Ref</th>
                          <th className="p-2.5">Billed Amount</th>
                          <th className="p-2.5">3-Way Match Recommendation</th>
                          <th className="p-2.5 text-right">Action</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                        {getFilteredDb().map((item: any, idx: number) => {
                          const isSelected = selectedAuditedId === item.id;
                          return (
                            <tr 
                              key={`audited-${item.id}-${idx}`} 
                              className={`transition-colors cursor-pointer ${
                                isSelected ? 'bg-amber-100/60 font-semibold' : 'hover:bg-slate-50'
                              }`}
                              onClick={() => selectAuditedInvoiceForView(item)}
                            >
                              <td className="p-2.5 font-bold text-slate-900 flex items-center gap-1.5">
                                <span className="font-mono">{item.id}</span>
                                {item.actionStatus === 'APPROVED' && (
                                  <span className="text-[9px] bg-emerald-100 text-emerald-800 px-1.5 py-0.2 rounded font-sans">Approved</span>
                                )}
                                {item.actionStatus === 'REJECTED' && (
                                  <span className="text-[9px] bg-rose-100 text-rose-800 px-1.5 py-0.2 rounded font-sans">Disputed</span>
                                )}
                              </td>
                              <td className="p-2.5 font-sans font-medium text-slate-800">{item.invoice.supplierName}</td>
                              <td className="p-2.5 text-indigo-600 font-bold">{item.invoice.poNumber || '—'}</td>
                              <td className="p-2.5 text-sky-700 font-bold">{item.grn ? item.grn.id : (item.invoice.poNumber ? 'GRN Pending' : '—')}</td>
                              <td className="p-2.5 font-bold text-slate-900">${item.invoice.totalAmount?.toFixed(2)}</td>
                              <td className="p-2.5 font-sans">
                                {item.recommendation === 'ACCEPT' && (
                                  <span className="inline-flex items-center gap-1 text-[10px] bg-emerald-100 text-emerald-800 font-bold py-0.5 px-2 rounded-full">
                                    <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                                    ACCEPT & PAY
                                  </span>
                                )}
                                {item.recommendation === 'REJECT' && (
                                  <span className="inline-flex items-center gap-1 text-[10px] bg-rose-100 text-rose-800 font-bold py-0.5 px-2 rounded-full">
                                    <XCircle className="w-3 h-3 text-rose-600" />
                                    REJECT PAYMENT
                                  </span>
                                )}
                                {item.recommendation === 'FORWARD' && (
                                  <span className="inline-flex items-center gap-1 text-[10px] bg-amber-100 text-amber-800 font-bold py-0.5 px-2 rounded-full">
                                    <AlertTriangle className="w-3 h-3 text-amber-600" />
                                    MANUAL REVIEW
                                  </span>
                                )}
                              </td>
                              <td className="p-2.5 text-right font-sans">
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    selectAuditedInvoiceForView(item);
                                  }}
                                  className="text-[10px] bg-slate-900 hover:bg-slate-800 text-white font-semibold py-1 px-2.5 rounded transition-all"
                                >
                                  Inspect Details
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  )
                ) : dbTab === 'po' ? (
                  dbState.purchaseOrders.length === 0 ? (
                    <div className="p-8 text-center bg-slate-50/80 rounded-lg border border-dashed border-slate-300 my-3 mx-4 space-y-3">
                      <Database className="w-9 h-9 text-slate-400 mx-auto" />
                      <div>
                        <h4 className="font-bold text-slate-800 text-sm">Purchase Order Database is Empty</h4>
                        <p className="text-slate-500 text-xs mt-0.5 max-w-md mx-auto">
                          The master database was cleared. You can manually create a Purchase Order, upload an Excel file, or restore the Boon Huat demo dataset.
                        </p>
                      </div>
                      <div className="flex flex-wrap items-center justify-center gap-2 pt-1">
                        <button
                          onClick={() => { setShowPoModal(true); setDbTab('po'); }}
                          className="text-xs font-semibold py-2 px-3.5 bg-amber-500 hover:bg-amber-400 text-slate-950 rounded shadow-xs flex items-center gap-1.5 transition-all"
                        >
                          <Plus className="w-4 h-4" />
                          + Add PO to DB
                        </button>
                        <button
                          onClick={() => setShowBatchModal(true)}
                          className="text-xs font-semibold py-2 px-3.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded shadow-xs flex items-center gap-1.5 transition-all"
                        >
                          <FilePlus className="w-4 h-4" />
                          Paste CSV / Text Lines
                        </button>
                        <button
                          onClick={handleResetSampleDatabase}
                          className="text-xs font-semibold py-2 px-3.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded shadow-xs flex items-center gap-1.5 border border-slate-700 transition-all"
                        >
                          <RefreshCw className="w-3.5 h-3.5 text-amber-400" />
                          Restore Sample Dataset
                        </button>
                      </div>
                    </div>
                  ) : (
                    <table className="w-full text-left text-xs text-slate-700">
                      <thead className="bg-slate-50 border-b border-slate-200 text-slate-500 font-bold uppercase tracking-wider text-[10px]">
                        <tr>
                          <th className="p-2.5">PO Number</th>
                          <th className="p-2.5">Supplier Name</th>
                          <th className="p-2.5">Purchase Date</th>
                          <th className="p-2.5">Item Description</th>
                          <th className="p-2.5">Qty Ordered</th>
                          <th className="p-2.5">Unit Price</th>
                          <th className="p-2.5">Total Amount</th>
                          <th className="p-2.5 text-right">Action</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                        {getFilteredDb().map((p: any, idx: number) => (
                          <tr key={`po-${p.id}-${idx}`} className="hover:bg-amber-50/40">
                            <td className="p-2.5 font-bold text-slate-900">{p.id}</td>
                            <td className="p-2.5 font-sans font-medium">
                              <div className="flex items-center gap-1.5">
                                <span>{p.supplierName}</span>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleEditPoSupplierName(p);
                                  }}
                                  className="text-[10px] text-blue-600 hover:text-blue-800 font-semibold hover:underline"
                                  title="Edit Supplier Name"
                                >
                                  (Edit)
                                </button>
                              </div>
                            </td>
                            <td className="p-2.5 text-slate-500">{p.purchaseDate}</td>
                            <td className="p-2.5 font-sans">{p.itemDescription}</td>
                            <td className="p-2.5">
                              <div className="flex items-center gap-1.5">
                                <span>{p.quantityOrdered}</span>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleEditPoQuantity(p);
                                  }}
                                  className="text-[10px] text-blue-600 hover:text-blue-800 font-semibold hover:underline"
                                  title="Edit Qty"
                                >
                                  (Edit)
                                </button>
                              </div>
                            </td>
                            <td className="p-2.5">
                              <div className="flex items-center gap-1.5">
                                <span>${p.unitPrice?.toFixed(2)}</span>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleEditPoUnitPrice(p);
                                  }}
                                  className="text-[10px] text-blue-600 hover:text-blue-800 font-semibold hover:underline"
                                  title="Edit Unit Price"
                                >
                                  (Edit)
                                </button>
                              </div>
                            </td>
                            <td className="p-2.5 font-bold text-slate-900">${p.totalAmount?.toFixed(2)}</td>
                            <td className="p-2.5 text-right font-sans">
                              <button
                                onClick={() => auditPoFromTable(p)}
                                className="text-[10px] bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold py-1 px-2 rounded"
                              >
                                Test Match Invoice
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )
                ) : dbTab === 'grn' ? (
                  dbState.goodsReceivedNotes.length === 0 ? (
                    <div className="p-8 text-center bg-slate-50/80 rounded-lg border border-dashed border-slate-300 my-3 mx-4 space-y-3">
                      <Database className="w-9 h-9 text-slate-400 mx-auto" />
                      <div>
                        <h4 className="font-bold text-slate-800 text-sm">Goods Received Note (GRN) Database is Empty</h4>
                        <p className="text-slate-500 text-xs mt-0.5 max-w-md mx-auto">
                          No receiving notes logged yet. Log a receiving note to pair with Purchase Orders or restore the Boon Huat demo dataset.
                        </p>
                      </div>
                      <div className="flex flex-wrap items-center justify-center gap-2 pt-1">
                        <button
                          onClick={() => { setShowGrnModal(true); setDbTab('grn'); }}
                          className="text-xs font-semibold py-2 px-3.5 bg-sky-600 hover:bg-sky-500 text-white rounded shadow-xs flex items-center gap-1.5 transition-all"
                        >
                          <Plus className="w-4 h-4" />
                          + Add GRN to DB
                        </button>
                        <button
                          onClick={() => setShowBatchModal(true)}
                          className="text-xs font-semibold py-2 px-3.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded shadow-xs flex items-center gap-1.5 transition-all"
                        >
                          <FilePlus className="w-4 h-4" />
                          Paste CSV / Text Lines
                        </button>
                        <button
                          onClick={handleResetSampleDatabase}
                          className="text-xs font-semibold py-2 px-3.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded shadow-xs flex items-center gap-1.5 border border-slate-700 transition-all"
                        >
                          <RefreshCw className="w-3.5 h-3.5 text-amber-400" />
                          Restore Sample Dataset
                        </button>
                      </div>
                    </div>
                  ) : (
                    <table className="w-full text-left text-xs text-slate-700">
                      <thead className="bg-slate-50 border-b border-slate-200 text-slate-500 font-bold uppercase tracking-wider text-[10px]">
                        <tr>
                          <th className="p-2.5">GRN Number</th>
                          <th className="p-2.5">Linked PO Ref</th>
                          <th className="p-2.5">Date Received</th>
                          <th className="p-2.5">Item Description</th>
                          <th className="p-2.5">Qty Received</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                        {getFilteredDb().map((g: any, idx: number) => (
                          <tr key={`grn-${g.id}-${idx}`} className="hover:bg-sky-50/40">
                            <td className="p-2.5 font-bold text-slate-900">{g.id}</td>
                            <td className="p-2.5 text-indigo-600 font-bold">{g.poNumber}</td>
                            <td className="p-2.5 text-slate-500">{g.dateReceived}</td>
                            <td className="p-2.5 font-sans">{g.itemDescription}</td>
                            <td className="p-2.5 font-bold text-sky-950">{g.quantityReceived}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )
                ) : (
                  dbState.pastInvoices.length === 0 ? (
                    <div className="p-8 text-center text-slate-400 text-xs">
                      No audited invoices logged in past history yet. Approved invoices will appear here.
                    </div>
                  ) : (
                    <table className="w-full text-left text-xs text-slate-700">
                      <thead className="bg-slate-50 border-b border-slate-200 text-slate-500 font-bold uppercase tracking-wider text-[10px]">
                        <tr>
                          <th className="p-2.5">Invoice ID</th>
                          <th className="p-2.5">Supplier Name</th>
                          <th className="p-2.5">PO Ref</th>
                          <th className="p-2.5">Invoice Date</th>
                          <th className="p-2.5">Item Description</th>
                          <th className="p-2.5">Total Paid</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                        {getFilteredDb().map((inv: any, idx: number) => (
                          <tr key={`past-${inv.id}-${idx}`} className="hover:bg-emerald-50/40">
                            <td className="p-2.5 font-bold text-slate-900">{inv.id}</td>
                            <td className="p-2.5 font-sans font-medium">{inv.supplierName}</td>
                            <td className="p-2.5 text-indigo-600">{inv.poNumber}</td>
                            <td className="p-2.5 text-slate-500">{inv.invoiceDate}</td>
                            <td className="p-2.5 font-sans">{inv.itemDescription}</td>
                            <td className="p-2.5 font-bold text-emerald-700">${inv.totalAmount?.toFixed(2)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )
                )}
              </div>

            </div>

          </div>
        </section>

        {/* SECTION 2: SUPPLIER INVOICES BATCH AUDIT & 3-WAY MATCH WORKSPACE */}
        <section className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden" id="invoice-audit-section">
          <div className="bg-slate-900 text-white p-4 sm:p-5 flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-800">
            <div>
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-5 h-5 text-indigo-400" />
                <h2 className="font-display font-bold text-base text-white">2. Send / Upload Supplier Invoices to Audit</h2>
              </div>
              <p className="text-xs text-slate-300 mt-1">
                Upload supplier invoices via Excel spreadsheet (.xlsx), Google Sheets export, CSV, or batch invoice files. Automatically links back to the PO & GRN database.
              </p>
            </div>
            {auditedInvoices.length > 0 && (
              <div className="flex items-center gap-2">
                <span className="text-xs bg-slate-800 text-slate-300 px-3 py-1 rounded border border-slate-700">
                  Audited: <strong>{auditedInvoices.length}</strong> Invoices
                </span>
                <button
                  onClick={() => setAuditedInvoices([])}
                  className="text-xs text-red-400 hover:text-red-300 py-1 px-2 border border-slate-700 rounded hover:bg-slate-800"
                >
                  Clear Results
                </button>
              </div>
            )}
          </div>

          <div className="p-5 space-y-6">
            
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
              
              {/* File Uploader supporting Excel, CSV, PDFs, Images, JSON */}
              <div className="border-2 border-dashed border-indigo-300 rounded-lg p-5 bg-indigo-50/30 hover:bg-indigo-50/60 text-center transition-all relative flex flex-col items-center justify-center group">
                <input 
                  type="file" 
                  accept=".xlsx,.xls,.csv,.json,.pdf,image/*,.txt"
                  multiple
                  onChange={handleInvoiceUpload}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                  id="invoice-file-input"
                />
                <div className="space-y-2">
                  <div className="inline-flex h-12 w-12 items-center justify-center rounded-full bg-indigo-100 text-indigo-700 border border-indigo-200 shadow-xs group-hover:scale-105 transition-transform">
                    <Upload className="w-6 h-6" />
                  </div>
                  <div>
                    <p className="font-bold text-sm text-slate-900">
                      Upload Supplier Invoice Files
                    </p>
                    <p className="text-xs text-slate-500 max-w-xs mx-auto mt-1">
                      Upload Excel (.xlsx), CSV, PDFs, images or JSON invoice batches.
                    </p>
                  </div>
                  <div className="pt-1">
                    <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-indigo-700 bg-indigo-100/80 px-3 py-1 rounded-full border border-indigo-200">
                      <FileSpreadsheet className="w-3.5 h-3.5" />
                      Select Excel or PDF Files
                    </span>
                  </div>
                </div>
              </div>

              {/* Google Sheets Direct Sync & Import Card */}
              <div className="border-2 border-dashed border-emerald-400 rounded-lg p-5 bg-emerald-50/40 hover:bg-emerald-50/80 text-center transition-all flex flex-col items-center justify-between group relative">
                <div className="space-y-2 flex flex-col items-center">
                  <div className="inline-flex h-12 w-12 items-center justify-center rounded-full bg-emerald-100 text-emerald-700 border border-emerald-300 shadow-xs group-hover:scale-105 transition-transform">
                    <FileSpreadsheet className="w-6 h-6 text-emerald-600" />
                  </div>
                  <div>
                    <div className="flex items-center justify-center gap-1.5">
                      <p className="font-bold text-sm text-slate-900">Google Sheets Integration</p>
                      <span className="text-[10px] bg-emerald-600 text-white font-bold px-1.5 py-0.2 rounded">LIVE</span>
                    </div>
                    <p className="text-xs text-slate-500 max-w-xs mx-auto mt-1">
                      Import invoice spreadsheets directly from Google Drive or export audit reports to Google Sheets.
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setShowGoogleSheetsModal(true)}
                  className="mt-3 text-xs font-bold py-2.5 px-4 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg shadow-sm flex items-center justify-center gap-2 transition-all w-full"
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  Open Google Sheets Sync
                </button>
              </div>

              {/* Manual Form Trigger */}
              <div className="border border-slate-200 rounded-lg p-5 bg-slate-50 flex flex-col justify-between">
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <Plus className="w-4 h-4 text-indigo-600" />
                    <h3 className="font-bold text-sm text-slate-900">Or Input Invoice Manually</h3>
                  </div>
                  <p className="text-xs text-slate-500">
                    Input custom invoice figures manually to test 3-way matching against your database records.
                  </p>
                </div>
                <button
                  onClick={() => setShowInvoiceForm(!showInvoiceForm)}
                  className="mt-3 text-xs font-semibold py-2 px-3 bg-white hover:bg-slate-100 text-slate-800 border border-slate-300 rounded shadow-xs flex items-center justify-center gap-1.5 self-start"
                >
                  {showInvoiceForm ? 'Hide Form' : 'Open Manual Invoice Input'}
                  <ChevronRight className="w-3.5 h-3.5 text-slate-400" />
                </button>
              </div>

            </div>

            {/* AUDITED INVOICES & MADAM LIM RECOMMENDATIONS TABLE */}
            {auditedInvoices.length > 0 && (
              <div className="space-y-4 pt-2">
                <div className="bg-slate-900 text-white p-4 rounded-lg flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-sm">
                  <div>
                    <div className="flex items-center gap-2">
                      <ShieldCheck className="w-5 h-5 text-emerald-400" />
                      <h3 className="font-bold text-sm text-white">3-Way Match Audit Verification Results & Recommendations for Madam Lim</h3>
                    </div>
                    <p className="text-xs text-slate-300 mt-1">
                      Evaluated every invoice against Boon Huat PO and GRN database records. Clear recommendations are provided below for immediate AP processing.
                    </p>
                  </div>

                  {/* Summary Pills & Export Button */}
                  <div className="flex flex-wrap items-center gap-2 text-xs font-bold">
                    <button
                      onClick={() => setShowGoogleSheetsModal(true)}
                      className="bg-emerald-600 hover:bg-emerald-500 text-white border border-emerald-500 px-3 py-1.5 rounded flex items-center gap-1.5 transition-colors shadow-xs"
                    >
                      <FileSpreadsheet className="w-4 h-4 text-white" />
                      <span>Export to Google Sheets</span>
                    </button>
                    <div className="bg-emerald-950/80 text-emerald-300 border border-emerald-800 px-3 py-1.5 rounded flex items-center gap-1.5">
                      <CheckCircle className="w-4 h-4 text-emerald-400" />
                      <span>ACCEPT & PAY: {auditedInvoices.filter(i => i.recommendation === 'ACCEPT').length}</span>
                    </div>
                    <div className="bg-red-950/80 text-red-300 border border-red-800 px-3 py-1.5 rounded flex items-center gap-1.5">
                      <XCircle className="w-4 h-4 text-red-400" />
                      <span>REJECT: {auditedInvoices.filter(i => i.recommendation === 'REJECT').length}</span>
                    </div>
                    <div className="bg-amber-950/80 text-amber-300 border border-amber-800 px-3 py-1.5 rounded flex items-center gap-1.5">
                      <AlertTriangle className="w-4 h-4 text-amber-400" />
                      <span>REVIEW / FORWARD: {auditedInvoices.filter(i => i.recommendation === 'FORWARD').length}</span>
                    </div>
                  </div>
                </div>

                {/* Table */}
                <div className="border border-slate-200 rounded-lg overflow-hidden shadow-xs bg-white">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-100 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider text-[10px]">
                        <tr>
                          <th className="p-3">Invoice Details</th>
                          <th className="p-3">Supplier Name</th>
                          <th className="p-3">PO & GRN Status</th>
                          <th className="p-3">Billed Amount</th>
                          <th className="p-3">Madam Lim Recommendation</th>
                          <th className="p-3">Audit Reason</th>
                          <th className="p-3 text-right">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 text-[11px]">
                        {auditedInvoices.map((item, idx) => {
                          const isSelected = selectedAuditedId === item.id;
                          return (
                            <tr 
                              key={`audited-inv-${item.id}-${idx}`} 
                              className={`transition-colors hover:bg-slate-50/80 ${
                                isSelected ? 'bg-indigo-50/60 border-l-4 border-indigo-600' : ''
                              }`}
                            >
                              <td className="p-3 font-mono">
                                <span className="font-bold text-slate-900 block">{item.invoice.id}</span>
                                <span className="text-[10px] text-slate-500">{item.invoice.invoiceDate}</span>
                              </td>
                              <td className="p-3 font-sans font-medium text-slate-800">
                                {item.invoice.supplierName}
                              </td>
                              <td className="p-3 font-mono space-y-1">
                                {item.po ? (
                                  <div className="flex items-center gap-1 text-emerald-700">
                                    <Check className="w-3 h-3 text-emerald-600 shrink-0" />
                                    <span>{item.po.id}</span>
                                  </div>
                                ) : (
                                  <div className="flex items-center gap-1 text-red-600 font-sans">
                                    <XCircle className="w-3 h-3 text-red-500 shrink-0" />
                                    <span>No PO Found</span>
                                  </div>
                                )}
                                {item.grn ? (
                                  <div className="flex items-center gap-1 text-sky-700">
                                    <Check className="w-3 h-3 text-sky-600 shrink-0" />
                                    <span>{item.grn.id} (Received {item.grn.quantityReceived})</span>
                                  </div>
                                ) : (
                                  <div className="flex items-center gap-1 text-amber-600 font-sans">
                                    <AlertTriangle className="w-3 h-3 text-amber-500 shrink-0" />
                                    <span>No GRN</span>
                                  </div>
                                )}
                              </td>
                              <td className="p-3 font-mono font-bold text-slate-900">
                                ${item.invoice.totalAmount.toFixed(2)}
                              </td>
                              <td className="p-3 font-sans">
                                {item.recommendation === 'ACCEPT' ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                                    <CheckCircle className="w-3.5 h-3.5 text-emerald-600" />
                                    RECOMMENDATION: ACCEPT PAYMENT
                                  </span>
                                ) : item.recommendation === 'REJECT' ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-red-100 text-red-800 border border-red-300">
                                    <XCircle className="w-3.5 h-3.5 text-red-600" />
                                    RECOMMENDATION: REJECT PAYMENT
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300">
                                    <AlertTriangle className="w-3.5 h-3.5 text-amber-600" />
                                    RECOMMENDATION: FORWARD REVIEW
                                  </span>
                                )}
                              </td>
                              <td className="p-3 font-sans text-[11px] text-slate-600 max-w-xs">
                                {item.recommendationSummary}
                              </td>
                              <td className="p-3 text-right font-sans space-x-1.5">
                                <button
                                  onClick={() => selectAuditedInvoiceForView(item)}
                                  className={`py-1 px-2.5 rounded text-[10px] font-semibold border transition-all ${
                                    isSelected 
                                      ? 'bg-indigo-600 text-white border-indigo-600' 
                                      : 'bg-white hover:bg-slate-100 text-slate-700 border-slate-300'
                                  }`}
                                >
                                  Inspect Details
                                </button>
                                {item.actionStatus === 'APPROVED' ? (
                                  <span className="text-[10px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-1 rounded">
                                    ✓ PAID & LOGGED
                                  </span>
                                ) : item.actionStatus === 'REJECTED' ? (
                                  <span className="text-[10px] font-bold text-red-700 bg-red-50 border border-red-200 px-2 py-1 rounded">
                                    ✕ REJECTED
                                  </span>
                                ) : item.actionStatus === 'REVIEW_SENT' ? (
                                  <span className="text-[10px] font-bold text-amber-700 bg-amber-50 border border-amber-200 px-2 py-1 rounded">
                                    FORWARDED
                                  </span>
                                ) : (
                                  <>
                                    <button
                                      onClick={() => handleBatchAction(item, 'approve')}
                                      className="py-1 px-2 rounded text-[10px] font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-xs"
                                    >
                                      Approve
                                    </button>
                                    <button
                                      onClick={() => handleBatchAction(item, 'dispute')}
                                      className="py-1 px-2 rounded text-[10px] font-bold bg-red-600 hover:bg-red-500 text-white shadow-xs"
                                    >
                                      Reject
                                    </button>
                                  </>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* Uploading progress indicator */}
            {loading && uploadProgress && (
              <div className="p-4 bg-slate-900 text-white rounded-lg flex items-center gap-3 animate-pulse border border-slate-800 shadow-md">
                <RefreshCw className="w-5 h-5 text-indigo-400 animate-spin shrink-0" />
                <div>
                  <p className="text-xs font-mono tracking-wider uppercase text-indigo-400">Processing Document</p>
                  <p className="text-sm font-medium">{uploadProgress}</p>
                </div>
              </div>
            )}

            {/* Manual Invoice Entry Form */}
            {showInvoiceForm && (
              <form onSubmit={handleManualInvoiceSubmit} className="bg-indigo-50/40 border border-indigo-100 p-4 rounded-lg space-y-3 animate-fade-in text-xs">
                <h4 className="font-bold text-slate-800 text-xs uppercase tracking-wider">Manual Invoice Entry Form</h4>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                  <div>
                    <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Invoice ID</label>
                    <input 
                      type="text" 
                      value={manualInvoice.id} 
                      onChange={e => setManualInvoice({...manualInvoice, id: e.target.value})}
                      className="w-full p-2 bg-white border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Supplier Name</label>
                    <input 
                      type="text" 
                      value={manualInvoice.supplierName} 
                      onChange={e => setManualInvoice({...manualInvoice, supplierName: e.target.value})}
                      className="w-full p-2 bg-white border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">PO Reference Number</label>
                    <input 
                      type="text" 
                      placeholder="e.g. PO-2026-1001"
                      value={manualInvoice.poNumber || ''} 
                      onChange={e => setManualInvoice({...manualInvoice, poNumber: e.target.value})}
                      className="w-full p-2 bg-white border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Invoice Date</label>
                    <input 
                      type="date" 
                      value={manualInvoice.invoiceDate} 
                      onChange={e => setManualInvoice({...manualInvoice, invoiceDate: e.target.value})}
                      className="w-full p-2 bg-white border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Payment Due Date</label>
                    <input 
                      type="date" 
                      value={manualInvoice.dueDate || ''} 
                      placeholder="e.g. 2026-09-06"
                      onChange={e => setManualInvoice({...manualInvoice, dueDate: e.target.value})}
                      className="w-full p-2 bg-white border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
                    />
                  </div>
                  <div className="sm:col-span-2">
                    <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Item Description</label>
                    <input 
                      type="text" 
                      value={manualInvoice.itemDescription} 
                      onChange={e => setManualInvoice({...manualInvoice, itemDescription: e.target.value})}
                      className="w-full p-2 bg-white border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Quantity Billed</label>
                    <input 
                      type="number" 
                      value={manualInvoice.quantityBilled} 
                      onChange={e => setManualInvoice({...manualInvoice, quantityBilled: Number(e.target.value)})}
                      className="w-full p-2 bg-white border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Unit Price ($)</label>
                    <input 
                      type="number" 
                      step="0.01" 
                      value={manualInvoice.unitPrice} 
                      onChange={e => setManualInvoice({...manualInvoice, unitPrice: Number(e.target.value)})}
                      className="w-full p-2 bg-white border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
                    />
                  </div>
                </div>
                <div className="pt-2 flex justify-end">
                  <button
                    type="submit"
                    className="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-2 px-4 rounded shadow-xs"
                  >
                    Send & Audit Against Database
                  </button>
                </div>
              </form>
            )}

          </div>
        </section>

        {/* COMPARISON WORKSPACE (PO vs GRN vs INVOICE) */}
        <div id="3way-match-report-section"></div>
        {invoice && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            
            {/* COLUMN 1: LINKED PURCHASE ORDER (PO) */}
            <div className="bg-white border border-slate-200 rounded-lg shadow-sm flex flex-col overflow-hidden">
              <div className="bg-slate-900 text-slate-100 border-b border-slate-800 py-3 px-4 flex items-center justify-between">
                <div className="flex items-center gap-1.5">
                  <FileText className="w-4 h-4 text-amber-500" />
                  <span className="font-display font-bold text-xs tracking-wide uppercase">1. Linked Database PO</span>
                </div>
                {po ? (
                  <span className="text-[10px] font-mono font-bold bg-amber-500 text-slate-950 py-0.5 px-2 rounded">
                    {po.id}
                  </span>
                ) : (
                  <span className="text-[10px] font-mono font-bold bg-red-500 text-white py-0.5 px-2 rounded">
                    NOT FOUND IN DB
                  </span>
                )}
              </div>
              
              <div className="p-4 flex-1 space-y-4 text-xs">
                {po ? (
                  <div className="space-y-3 relative">
                    <div className="absolute top-1 right-1 border-2 border-emerald-600 text-emerald-600 text-[10px] uppercase font-bold tracking-widest px-2 py-1 transform rotate-12 rounded opacity-40 select-none">
                      Authorized PO
                    </div>
                    
                    <div className="grid grid-cols-2 gap-2 border-b border-slate-100 pb-3">
                      <div>
                        <span className="text-slate-400 block text-[10px] uppercase tracking-wider">Approved Supplier</span>
                        <div className="flex items-center gap-1.5">
                          <span className="font-semibold text-slate-800">{po.supplierName}</span>
                          <button
                            onClick={() => handleEditPoSupplierName(po)}
                            className="text-[10px] text-blue-600 hover:text-blue-800 font-semibold hover:underline"
                            title="Edit PO Supplier Name"
                          >
                            (Edit)
                          </button>
                        </div>
                      </div>
                      <div>
                        <span className="text-slate-400 block text-[10px] uppercase tracking-wider">Purchase Date</span>
                        <span className="font-mono text-slate-800">{po.purchaseDate}</span>
                      </div>
                    </div>

                    <div className="bg-slate-50 p-3 rounded border border-slate-100">
                      <span className="text-slate-400 block text-[10px] uppercase tracking-wider mb-1">Item Ordered</span>
                      <p className="font-medium text-slate-900 text-sm">{po.itemDescription}</p>
                    </div>

                    <div className="grid grid-cols-3 gap-2 pt-1">
                      <div className="bg-slate-50 p-2 rounded text-center">
                        <span className="text-slate-400 block text-[9px] uppercase tracking-wider">Qty Ordered</span>
                        <div className="flex items-center justify-center gap-1">
                          <span className="font-mono font-semibold text-slate-800 text-sm">{po.quantityOrdered}</span>
                          <button
                            onClick={() => handleEditPoQuantity(po)}
                            className="text-[10px] text-blue-600 hover:text-blue-800 font-semibold hover:underline"
                            title="Edit PO Quantity"
                          >
                            (Edit)
                          </button>
                        </div>
                      </div>
                      <div className="bg-slate-50 p-2 rounded text-center">
                        <span className="text-slate-400 block text-[9px] uppercase tracking-wider">Unit Price</span>
                        <div className="flex items-center justify-center gap-1">
                          <span className="font-mono font-semibold text-slate-800 text-sm">${po.unitPrice.toFixed(2)}</span>
                          <button
                            onClick={() => handleEditPoUnitPrice(po)}
                            className="text-[10px] text-blue-600 hover:text-blue-800 font-semibold hover:underline"
                            title="Edit PO Unit Price"
                          >
                            (Edit)
                          </button>
                        </div>
                      </div>
                      <div className="bg-slate-900/5 p-2 rounded text-center border border-slate-100">
                        <span className="text-slate-400 block text-[9px] uppercase tracking-wider">Total PO</span>
                        <span className="font-mono font-bold text-slate-900 text-sm">${po.totalAmount.toFixed(2)}</span>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="h-full flex flex-col items-center justify-center text-center p-6 text-slate-400 space-y-2">
                    <AlertTriangle className="w-8 h-8 text-red-500" />
                    <p className="font-semibold text-slate-700">No PO Linked from Database</p>
                    <p className="text-[11px] text-slate-500">
                      The invoice references PO "{invoice.poNumber || 'N/A'}", which does not exist in your Boon Huat database. Add the PO to the database above to link.
                    </p>
                  </div>
                )}
              </div>
            </div>

            {/* COLUMN 2: LINKED GOODS RECEIVED NOTE (GRN) */}
            <div className="bg-white border border-slate-200 rounded-lg shadow-sm flex flex-col overflow-hidden">
              <div className="bg-slate-900 text-slate-100 border-b border-slate-800 py-3 px-4 flex items-center justify-between">
                <div className="flex items-center gap-1.5">
                  <Database className="w-4 h-4 text-sky-400" />
                  <span className="font-display font-bold text-xs tracking-wide uppercase">2. Linked Warehouse GRN</span>
                </div>
                {grn ? (
                  <span className="text-[10px] font-mono font-bold bg-sky-400 text-slate-950 py-0.5 px-2 rounded">
                    {grn.id}
                  </span>
                ) : (
                  <span className="text-[10px] font-mono font-bold bg-red-500 text-white py-0.5 px-2 rounded">
                    MISSING GRN
                  </span>
                )}
              </div>
              
              <div className="p-4 flex-1 space-y-4 text-xs">
                {grn ? (
                  <div className="space-y-3 relative">
                    <div className="absolute top-1 right-1 border-2 border-sky-600 text-sky-600 text-[10px] uppercase font-bold tracking-widest px-2 py-1 transform -rotate-12 rounded opacity-40 select-none">
                      Goods Logged
                    </div>
                    
                    <div className="grid grid-cols-2 gap-2 border-b border-slate-100 pb-3">
                      <div>
                        <span className="text-slate-400 block text-[10px] uppercase tracking-wider">PO Reference</span>
                        <span className="font-mono font-semibold text-slate-800">{grn.poNumber}</span>
                      </div>
                      <div>
                        <span className="text-slate-400 block text-[10px] uppercase tracking-wider">Date Received</span>
                        <span className="font-mono text-slate-800">{grn.dateReceived}</span>
                      </div>
                    </div>

                    <div className="bg-slate-50 p-3 rounded border border-slate-100">
                      <span className="text-slate-400 block text-[10px] uppercase tracking-wider mb-1">Items Unloaded</span>
                      <p className="font-medium text-slate-900 text-sm">{grn.itemDescription}</p>
                    </div>

                    <div className="grid grid-cols-2 gap-2 pt-1">
                      <div className="bg-sky-50 p-2.5 rounded border border-sky-100 text-center col-span-2">
                        <span className="text-slate-500 block text-[9px] uppercase tracking-wider font-medium">Quantity Received in Yard</span>
                        <span className="font-mono font-bold text-sky-950 text-lg">{grn.quantityReceived} units</span>
                        <span className="text-[10px] block text-sky-700 mt-0.5">(logged in Boon Huat database)</span>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="h-full flex flex-col items-center justify-center text-center p-6 text-slate-400 space-y-2">
                    <XCircle className="w-8 h-8 text-red-500" />
                    <p className="font-semibold text-slate-700">No GRN Delivery Record</p>
                    <p className="text-[11px] text-slate-500">
                      No Goods Received Note (GRN) found in the database for PO {po?.id || invoice.poNumber || 'N/A'}. Upload or log a receiving slip to verify delivery.
                    </p>
                  </div>
                )}
              </div>
            </div>

            {/* COLUMN 3: SUPPLIER INVOICE (AUDITED BILL) */}
            <div className="bg-white border border-slate-200 rounded-lg shadow-sm flex flex-col overflow-hidden">
              <div className="bg-slate-900 text-slate-100 border-b border-slate-800 py-3 px-4 flex items-center justify-between">
                <div className="flex items-center gap-1.5">
                  <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                  <span className="font-display font-bold text-xs tracking-wide uppercase">3. Sent Supplier Invoice</span>
                </div>
                <span className="text-[10px] font-mono font-bold bg-emerald-400 text-slate-950 py-0.5 px-2 rounded">
                  {invoice.id}
                </span>
              </div>
              
              <div className="p-4 flex-1 space-y-4 text-xs">
                <div className="space-y-3 relative">
                  {activeFileName && (
                    <div className="absolute -top-1 right-1 bg-indigo-100 text-indigo-800 border border-indigo-200 text-[9px] uppercase font-bold tracking-wider px-1.5 py-0.5 rounded">
                      {activeFileName}
                    </div>
                  )}
                  
                  <div className="grid grid-cols-3 gap-2 border-b border-slate-100 pb-3">
                    <div>
                      <span className="text-slate-400 block text-[10px] uppercase tracking-wider">Supplier Name</span>
                      <span className="font-semibold text-slate-800 truncate block">{invoice.supplierName}</span>
                    </div>
                    <div>
                      <span className="text-slate-400 block text-[10px] uppercase tracking-wider">Invoice Date</span>
                      <span className="font-mono text-slate-800">{invoice.invoiceDate}</span>
                    </div>
                    <div>
                      <span className="text-emerald-700 block text-[10px] uppercase tracking-wider font-semibold">Payment Due Date</span>
                      <span className="font-mono font-bold text-emerald-900">
                        {invoice.dueDate || new Date(new Date(invoice.invoiceDate).getTime() + 30*24*60*60*1000).toISOString().split('T')[0]}
                      </span>
                    </div>
                  </div>

                  <div className="bg-slate-50 p-3 rounded border border-slate-100">
                    <span className="text-slate-400 block text-[10px] uppercase tracking-wider mb-1">Item Billed</span>
                    <p className="font-medium text-slate-900 text-sm">{invoice.itemDescription}</p>
                  </div>

                  <div className="grid grid-cols-3 gap-2 pt-1">
                    <div className="bg-slate-50 p-2 rounded text-center">
                      <span className="text-slate-400 block text-[9px] uppercase tracking-wider">Qty Billed</span>
                      <span className="font-mono font-semibold text-slate-800 text-sm">{invoice.quantityBilled}</span>
                    </div>
                    <div className="bg-slate-50 p-2 rounded text-center">
                      <span className="text-slate-400 block text-[9px] uppercase tracking-wider">Unit Price</span>
                      <span className="font-mono font-semibold text-slate-800 text-sm">${invoice.unitPrice.toFixed(2)}</span>
                    </div>
                    <div className="bg-emerald-50 p-2 rounded text-center border border-emerald-100">
                      <span className="text-emerald-700 block text-[9px] uppercase tracking-wider font-semibold">Total Billed</span>
                      <span className="font-mono font-bold text-emerald-950 text-sm">${invoice.totalAmount.toFixed(2)}</span>
                    </div>
                  </div>

                  {invoice.poNumber && (
                    <div className="bg-slate-100 p-2 rounded flex items-center justify-between text-[11px] font-mono border border-slate-200">
                      <span className="text-slate-500">Stated PO Ref:</span>
                      <span className="font-semibold text-slate-800">{invoice.poNumber}</span>
                    </div>
                  )}
                </div>
              </div>
            </div>

          </div>
        )}

        {/* 3-WAY MATCH RESULTS & EXPLAINABILITY REPORT */}
        {invoice && matchReport && (
          <div className="bg-white border border-slate-200 rounded-lg shadow-md overflow-hidden animate-fade-in space-y-0">
            
            {/* OVERALL MATCHING STATUS BANNER */}
            <div className={`p-6 flex flex-col md:flex-row md:items-center justify-between gap-4 border-b ${
              matchReport.overallResult === 'APPROVED' 
                ? 'bg-emerald-500 text-white border-emerald-600' 
                : matchReport.overallResult === 'MANUAL REVIEW REQUIRED' 
                  ? 'bg-amber-400 text-slate-950 border-amber-500' 
                  : 'bg-red-500 text-white border-red-600'
            }`}>
              <div className="flex items-start gap-4">
                <div className={`h-12 w-12 rounded-full flex items-center justify-center shrink-0 shadow-md ${
                  matchReport.overallResult === 'APPROVED' 
                    ? 'bg-emerald-600' 
                    : matchReport.overallResult === 'MANUAL REVIEW REQUIRED' 
                      ? 'bg-amber-500' 
                      : 'bg-red-600'
                }`}>
                  {matchReport.overallResult === 'APPROVED' ? (
                    <CheckCircle className="w-7 h-7" />
                  ) : matchReport.overallResult === 'MANUAL REVIEW REQUIRED' ? (
                    <AlertTriangle className="w-7 h-7 text-slate-950" />
                  ) : (
                    <XCircle className="w-7 h-7" />
                  )}
                </div>
                <div>
                  <span className="text-xs uppercase font-bold tracking-wider opacity-90">Boon Huat AI AP Matcher Report</span>
                  <h3 className="text-2xl font-display font-extrabold tracking-tight mt-0.5">
                    {matchReport.overallResult === 'APPROVED' 
                      ? 'COMPLIANT: APPROVED' 
                      : matchReport.overallResult === 'MANUAL REVIEW REQUIRED' 
                        ? 'COMPLIANCE WARNING: MANUAL REVIEW REQUIRED' 
                        : 'COMPLIANCE FAILURE: REJECT PAYMENT'}
                  </h3>
                  <p className="text-xs opacity-90 mt-1 max-w-xl font-medium leading-relaxed">
                    {matchReport.explanation}
                  </p>
                </div>
              </div>

              {actionStatus ? (
                <div className="bg-slate-950/20 py-2 px-4 rounded-lg border border-white/20 flex items-center gap-2 text-xs md:self-center font-semibold">
                  <Check className="w-4 h-4 text-emerald-300 shrink-0" />
                  <div>
                    <p className="uppercase tracking-wider text-[10px] text-slate-200">Madam Lim Actioned</p>
                    <p className="text-white">
                      {actionStatus === 'APPROVED' ? 'Approved & Logged' : actionStatus === 'REJECTED' ? 'Rejected' : 'Forwarded for Review'} ({actionTimestamp})
                    </p>
                  </div>
                </div>
              ) : (
                <div className="bg-slate-950/10 py-1.5 px-3.5 rounded-lg border border-white/15 text-xs text-center md:self-center">
                  <Clock className="w-4 h-4 mx-auto mb-1 opacity-80" />
                  <span className="opacity-80">Pending Auditor Decision</span>
                </div>
              )}
            </div>

            {/* EXPLAINABILITY & RECOMMENDED ACTION */}
            <div className="grid grid-cols-1 md:grid-cols-2 border-b border-slate-200">
              <div className="p-6 border-b md:border-b-0 md:border-r border-slate-200 space-y-3">
                <div className="flex items-center gap-2 text-slate-700">
                  <Info className="w-5 h-5 text-slate-500" />
                  <h4 className="font-display font-bold text-sm">Plain-English Explainability</h4>
                </div>
                <div className="bg-slate-50 p-4 rounded-lg border border-slate-100 text-xs leading-relaxed text-slate-700 space-y-2">
                  <p className="font-semibold text-slate-900">Madam Lim's Audit Note:</p>
                  <p>
                    {matchReport.checks.duplicateInvoice.status === 'FAIL' 
                      ? matchReport.checks.duplicateInvoice.explanation 
                      : matchReport.checks.poNumber.status === 'FAIL'
                        ? matchReport.checks.poNumber.explanation
                        : matchReport.checks.grnPresence.status === 'FAIL'
                          ? matchReport.checks.grnPresence.explanation
                          : matchReport.checks.quantity.status === 'FAIL'
                            ? matchReport.checks.quantity.explanation
                            : matchReport.checks.unitPrice.status === 'FAIL'
                              ? matchReport.checks.unitPrice.explanation
                              : matchReport.checks.calculation.status === 'FAIL'
                                ? matchReport.checks.calculation.explanation
                                : "Documents completely align with the Boon Huat PO database and receiving log. Payment can proceed."}
                  </p>
                </div>
              </div>

              <div className="p-6 space-y-3 bg-slate-50/40">
                <div className="flex items-center gap-2 text-slate-700">
                  <TrendingUp className="w-5 h-5 text-indigo-500" />
                  <h4 className="font-display font-bold text-sm text-indigo-950">Recommended Action</h4>
                </div>
                <div className="bg-white p-4 rounded-lg border border-indigo-100 text-xs leading-relaxed text-slate-700 space-y-3 shadow-xs">
                  <p className="font-bold text-indigo-950">System Suggestion:</p>
                  <p className="text-indigo-900">{matchReport.recommendedAction}</p>
                  
                  <div className="flex flex-wrap gap-2 pt-2">
                    <button
                      disabled={actionStatus !== null}
                      onClick={handleApprove}
                      className={`text-xs font-semibold py-2 px-4 rounded shadow-sm transition-all ${
                        actionStatus === 'APPROVED'
                          ? 'bg-emerald-600 text-white cursor-not-allowed'
                          : actionStatus !== null
                            ? 'bg-slate-100 text-slate-400 cursor-not-allowed'
                            : 'bg-emerald-600 hover:bg-emerald-700 text-white'
                      }`}
                    >
                      {actionStatus === 'APPROVED' ? 'Approved ✓' : 'Approve Payment'}
                    </button>
                    
                    <button
                      disabled={actionStatus !== null}
                      onClick={handleReject}
                      className={`text-xs font-semibold py-2 px-4 rounded shadow-sm transition-all ${
                        actionStatus === 'REJECTED'
                          ? 'bg-red-600 text-white cursor-not-allowed'
                          : actionStatus !== null
                            ? 'bg-slate-100 text-slate-400 cursor-not-allowed'
                            : 'bg-red-600 hover:bg-red-700 text-white'
                      }`}
                    >
                      {actionStatus === 'REJECTED' ? 'Rejected ✓' : 'Reject'}
                    </button>
                    
                    <button
                      disabled={actionStatus !== null}
                      onClick={() => setShowForwardModal(true)}
                      className={`text-xs font-semibold py-2 px-4 rounded shadow-sm transition-all ${
                        actionStatus === 'REVIEW_SENT'
                          ? 'bg-slate-700 text-white cursor-not-allowed'
                          : actionStatus !== null
                            ? 'bg-slate-100 text-slate-400 cursor-not-allowed'
                            : 'bg-white text-slate-700 border border-slate-300 hover:bg-slate-50'
                      }`}
                    >
                      {actionStatus === 'REVIEW_SENT' ? 'Sent for Review ✓' : 'Send for Review'}
                    </button>
                  </div>

                  {/* Post-Approval Google Sheets Export Action Card */}
                  {actionStatus === 'APPROVED' && (
                    <div className="mt-3 p-3 bg-emerald-50 border border-emerald-300 rounded-lg flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 text-xs animate-fade-in shadow-xs">
                      <div className="space-y-0.5">
                        <div className="flex items-center gap-1.5 font-bold text-emerald-950">
                          <FileSpreadsheet className="w-4 h-4 text-emerald-600 shrink-0" />
                          <span>Approved Payment Logged</span>
                        </div>
                        <p className="text-[11px] text-emerald-800">
                          Payment Due Date: <strong className="font-mono text-emerald-950">{
                            invoice.dueDate || new Date(new Date(invoice.invoiceDate).getTime() + 30*24*60*60*1000).toISOString().split('T')[0]
                          }</strong>
                        </p>
                      </div>
                      <button
                        onClick={() => setShowGoogleSheetsModal(true)}
                        className="bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-2 px-3.5 rounded-lg shadow-sm flex items-center gap-1.5 text-xs transition-colors shrink-0"
                      >
                        <FileSpreadsheet className="w-4 h-4" />
                        <span>Export to Google Sheets</span>
                      </button>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* CHECKLIST GRID */}
            <div className="p-6 space-y-4">
              <h4 className="font-display font-bold text-sm text-slate-800">8-Point Matching Checklist against Boon Huat DB</h4>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                
                <div className={`p-3 rounded-lg border text-xs space-y-1 ${
                  matchReport.checks.poNumber.status === 'PASS' ? 'bg-emerald-50/40 border-emerald-100' : 'bg-red-50/40 border-red-100'
                }`}>
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">PO Registration</span>
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                      matchReport.checks.poNumber.status === 'PASS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
                    }`}>{matchReport.checks.poNumber.status}</span>
                  </div>
                  <p className="text-[11px] text-slate-500">{matchReport.checks.poNumber.details}</p>
                </div>

                <div className={`p-3 rounded-lg border text-xs space-y-1 ${
                  matchReport.checks.supplierName.status === 'PASS' ? 'bg-emerald-50/40 border-emerald-100' : 'bg-red-50/40 border-red-100'
                }`}>
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">Supplier Name Match</span>
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                      matchReport.checks.supplierName.status === 'PASS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
                    }`}>{matchReport.checks.supplierName.status}</span>
                  </div>
                  <p className="text-[11px] text-slate-500">{matchReport.checks.supplierName.details}</p>
                </div>

                <div className={`p-3 rounded-lg border text-xs space-y-1 ${
                  matchReport.checks.grnPresence.status === 'PASS' ? 'bg-emerald-50/40 border-emerald-100' : 'bg-red-50/40 border-red-100'
                }`}>
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">Goods Unloaded Log</span>
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                      matchReport.checks.grnPresence.status === 'PASS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
                    }`}>{matchReport.checks.grnPresence.status}</span>
                  </div>
                  <p className="text-[11px] text-slate-500">{matchReport.checks.grnPresence.details}</p>
                </div>

                <div className={`p-3 rounded-lg border text-xs space-y-1 ${
                  matchReport.checks.quantity.status === 'PASS' ? 'bg-emerald-50/40 border-emerald-100' : 'bg-red-50/40 border-red-100'
                }`}>
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">Quantity Verification</span>
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                      matchReport.checks.quantity.status === 'PASS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
                    }`}>{matchReport.checks.quantity.status}</span>
                  </div>
                  <p className="text-[11px] text-slate-500">{matchReport.checks.quantity.details}</p>
                </div>

                <div className={`p-3 rounded-lg border text-xs space-y-1 ${
                  matchReport.checks.unitPrice.status === 'PASS' ? 'bg-emerald-50/40 border-emerald-100' : 'bg-red-50/40 border-red-100'
                }`}>
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">Unit Price Validation</span>
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                      matchReport.checks.unitPrice.status === 'PASS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
                    }`}>{matchReport.checks.unitPrice.status}</span>
                  </div>
                  <p className="text-[11px] text-slate-500">{matchReport.checks.unitPrice.details}</p>
                </div>

                <div className={`p-3 rounded-lg border text-xs space-y-1 ${
                  matchReport.checks.calculation.status === 'PASS' ? 'bg-emerald-50/40 border-emerald-100' : 'bg-red-50/40 border-red-100'
                }`}>
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">Arithmetic Calculation</span>
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                      matchReport.checks.calculation.status === 'PASS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
                    }`}>{matchReport.checks.calculation.status}</span>
                  </div>
                  <p className="text-[11px] text-slate-500">{matchReport.checks.calculation.details}</p>
                </div>

                <div className={`p-3 rounded-lg border text-xs space-y-1 ${
                  matchReport.checks.duplicateInvoice.status === 'PASS' ? 'bg-emerald-50/40 border-emerald-100' : 'bg-red-50/40 border-red-100'
                }`}>
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">Duplicate Check</span>
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                      matchReport.checks.duplicateInvoice.status === 'PASS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
                    }`}>{matchReport.checks.duplicateInvoice.status}</span>
                  </div>
                  <p className="text-[11px] text-slate-500">{matchReport.checks.duplicateInvoice.details}</p>
                </div>

                <div className={`p-3 rounded-lg border text-xs space-y-1 ${
                  matchReport.checks.totalAmount.status === 'PASS' ? 'bg-emerald-50/40 border-emerald-100' : 'bg-red-50/40 border-red-100'
                }`}>
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">Total Amount Match</span>
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                      matchReport.checks.totalAmount.status === 'PASS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
                    }`}>{matchReport.checks.totalAmount.status}</span>
                  </div>
                  <p className="text-[11px] text-slate-500">{matchReport.checks.totalAmount.details}</p>
                </div>

              </div>
            </div>

          </div>
        )}

      </main>

      {/* FOOTER */}
      <footer className="bg-slate-900 text-slate-400 py-4 px-6 text-center text-xs border-t border-slate-800">
        <p>Boon Huat Hardware & Supplies Pte Ltd — Accounts Payable (AP) 3-Way Match System</p>
      </footer>

      {/* MODAL: ADD PO TO DATABASE */}
      {showPoModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in">
          <div className="bg-white border border-slate-200 rounded-lg shadow-xl max-w-lg w-full p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3">
              <h3 className="font-display font-bold text-sm text-slate-900 flex items-center gap-2">
                <FileText className="w-4 h-4 text-amber-500" />
                Add Purchase Order (PO) to Boon Huat Database
              </h3>
              <button onClick={() => setShowPoModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-sm">✕</button>
            </div>
            
            <form onSubmit={handleAddPo} className="space-y-3 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">PO Number</label>
                  <input 
                    type="text" 
                    required 
                    value={newPo.id} 
                    onChange={e => setNewPo({...newPo, id: e.target.value})}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Purchase Date</label>
                  <input 
                    type="date" 
                    required 
                    value={newPo.purchaseDate} 
                    onChange={e => setNewPo({...newPo, purchaseDate: e.target.value})}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Approved Supplier Name</label>
                <input 
                  type="text" 
                  required 
                  value={newPo.supplierName} 
                  onChange={e => setNewPo({...newPo, supplierName: e.target.value})}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                />
              </div>

              <div>
                <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Item Description</label>
                <input 
                  type="text" 
                  required 
                  value={newPo.itemDescription} 
                  onChange={e => setNewPo({...newPo, itemDescription: e.target.value})}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Quantity Ordered</label>
                  <input 
                    type="number" 
                    required 
                    value={newPo.quantityOrdered} 
                    onChange={e => setNewPo({...newPo, quantityOrdered: Number(e.target.value)})}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Approved Unit Price ($)</label>
                  <input 
                    type="number" 
                    step="0.01" 
                    required 
                    value={newPo.unitPrice} 
                    onChange={e => setNewPo({...newPo, unitPrice: Number(e.target.value)})}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                  />
                </div>
              </div>

              <div className="pt-3 border-t flex items-center justify-between">
                <span className="font-mono text-slate-600">Total: <strong>${(newPo.quantityOrdered * newPo.unitPrice).toFixed(2)}</strong></span>
                <div className="flex gap-2">
                  <button 
                    type="button" 
                    onClick={() => setShowPoModal(false)}
                    className="py-1.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded font-medium"
                  >
                    Cancel
                  </button>
                  <button 
                    type="submit"
                    className="py-1.5 px-4 bg-amber-500 hover:bg-amber-400 text-slate-950 rounded font-bold"
                  >
                    Save to Database
                  </button>
                </div>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: ADD GRN TO DATABASE */}
      {showGrnModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in">
          <div className="bg-white border border-slate-200 rounded-lg shadow-xl max-w-lg w-full p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3">
              <h3 className="font-display font-bold text-sm text-slate-900 flex items-center gap-2">
                <Database className="w-4 h-4 text-sky-500" />
                Add Goods Received Note (GRN) to Database
              </h3>
              <button onClick={() => setShowGrnModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-sm">✕</button>
            </div>
            
            <form onSubmit={handleAddGrn} className="space-y-3 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">GRN ID</label>
                  <input 
                    type="text" 
                    required 
                    value={newGrn.id} 
                    onChange={e => setNewGrn({...newGrn, id: e.target.value})}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Target PO Number</label>
                  <input 
                    type="text" 
                    required 
                    placeholder="e.g. PO-2026-1001"
                    value={newGrn.poNumber} 
                    onChange={e => setNewGrn({...newGrn, poNumber: e.target.value})}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Date Received in Yard</label>
                <input 
                  type="date" 
                  required 
                  value={newGrn.dateReceived} 
                  onChange={e => setNewGrn({...newGrn, dateReceived: e.target.value})}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                />
              </div>

              <div>
                <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Item Received Description</label>
                <input 
                  type="text" 
                  required 
                  value={newGrn.itemDescription} 
                  onChange={e => setNewGrn({...newGrn, itemDescription: e.target.value})}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                />
              </div>

              <div>
                <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Quantity Received</label>
                <input 
                  type="number" 
                  required 
                  value={newGrn.quantityReceived} 
                  onChange={e => setNewGrn({...newGrn, quantityReceived: Number(e.target.value)})}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none"
                />
              </div>

              <div className="pt-3 border-t flex justify-end gap-2">
                <button 
                  type="button" 
                  onClick={() => setShowGrnModal(false)}
                  className="py-1.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded font-medium"
                >
                  Cancel
                </button>
                <button 
                  type="submit"
                  className="py-1.5 px-4 bg-sky-600 hover:bg-sky-500 text-white rounded font-bold"
                >
                  Save GRN to Database
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: BATCH / TEXT CSV IMPORT */}
      {showBatchModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in">
          <div className="bg-white border border-slate-200 rounded-lg shadow-xl max-w-xl w-full p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3">
              <h3 className="font-display font-bold text-sm text-slate-900 flex items-center gap-2">
                <FilePlus className="w-4 h-4 text-indigo-600" />
                Batch Import CSV / Text Records to Boon Huat Database
              </h3>
              <button onClick={() => setShowBatchModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-sm">✕</button>
            </div>
            
            <div className="space-y-3 text-xs">
              <p className="text-slate-500">
                Paste CSV lines for POs and GRNs below. Format:
              </p>
              <div className="bg-slate-900 text-slate-300 font-mono text-[11px] p-3 rounded space-y-1">
                <p className="text-amber-400">PO-2026-1001, Supplier Name, 2026-07-01, Brass Valves, 100, 18.50, 1850.00</p>
                <p className="text-sky-400">GRN-2026-5001, PO-2026-1001, 2026-07-05, Brass Valves, 100</p>
              </div>

              <textarea 
                rows={6}
                value={batchRawText}
                onChange={e => setBatchRawText(e.target.value)}
                placeholder="Paste CSV lines here..."
                className="w-full p-3 font-mono text-xs border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 outline-none"
              />

              <div className="pt-2 flex justify-end gap-2">
                <button 
                  type="button" 
                  onClick={() => setShowBatchModal(false)}
                  className="py-1.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded font-medium"
                >
                  Cancel
                </button>
                <button 
                  type="button"
                  onClick={handleBatchTextImport}
                  className="py-1.5 px-4 bg-indigo-600 hover:bg-indigo-500 text-white rounded font-bold"
                >
                  Import All Lines to Database
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: DISPUTE EMAIL */}
      {showEmailModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in">
          <div className="bg-white border border-slate-200 rounded-lg shadow-xl max-w-xl w-full p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3">
              <h3 className="font-display font-bold text-sm text-slate-900 flex items-center gap-2">
                <Mail className="w-4 h-4 text-red-600" />
                Pre-Drafted Vendor Dispute Email
              </h3>
              <button onClick={() => setShowEmailModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-sm">✕</button>
            </div>
            
            <div className="space-y-3 text-xs">
              <p className="text-slate-500">
                Generated based on 3-Way Match findings. Madam Lim can copy and send directly to supplier accounts team.
              </p>

              <textarea 
                rows={10}
                value={disputeEmailContent}
                onChange={e => setDisputeEmailContent(e.target.value)}
                className="w-full p-3 font-mono text-xs border border-slate-300 rounded focus:ring-1 focus:ring-slate-800 outline-none bg-slate-50"
              />

              <div className="flex justify-between items-center pt-2">
                <span className="text-[11px] text-slate-400">Payment placed on HOLD in system</span>
                <div className="flex gap-2">
                  <button 
                    onClick={() => {
                      navigator.clipboard.writeText(disputeEmailContent);
                      setCopiedEmail(true);
                      setTimeout(() => setCopiedEmail(false), 2000);
                    }}
                    className="py-1.5 px-4 bg-slate-900 hover:bg-slate-800 text-white rounded font-bold"
                  >
                    {copiedEmail ? 'Copied to Clipboard ✓' : 'Copy Email Text'}
                  </button>
                  <button 
                    onClick={() => setShowEmailModal(false)}
                    className="py-1.5 px-3 bg-slate-100 text-slate-700 rounded font-medium"
                  >
                    Close
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: FORWARD TO SUPERVISOR */}
      {showForwardModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in">
          <div className="bg-white border border-slate-200 rounded-lg shadow-xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3">
              <h3 className="font-display font-bold text-sm text-slate-900 flex items-center gap-2">
                <User className="w-4 h-4 text-indigo-600" />
                Send Invoice for Internal Review
              </h3>
              <button onClick={() => setShowForwardModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-sm">✕</button>
            </div>
            
            <div className="space-y-3 text-xs">
              <div>
                <label className="text-[10px] text-slate-500 uppercase font-semibold block mb-1">Select Supervisor / Dept</label>
                <select 
                  value={forwardSupervisor} 
                  onChange={e => setForwardSupervisor(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded outline-none"
                >
                  <option value="Purchasing (Mr. Tan)">Purchasing Dept (Mr. Tan)</option>
                  <option value="Yard Receiving (Ah Huat)">Yard Receiving (Supervisor Ah Huat)</option>
                  <option value="Finance Director (Mr. Boon)">Finance Director (Mr. Boon)</option>
                </select>
              </div>

              <div className="pt-2 flex justify-end gap-2">
                <button 
                  onClick={() => setShowForwardModal(false)}
                  className="py-1.5 px-3 bg-slate-100 text-slate-700 rounded font-medium"
                >
                  Cancel
                </button>
                <button 
                  onClick={submitForwardReview}
                  className="py-1.5 px-4 bg-indigo-600 hover:bg-indigo-500 text-white rounded font-bold"
                >
                  Forward Review Case
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: CLEAR / RESET MASTER DATABASE */}
      {showClearModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in">
          <div className="bg-white border border-slate-200 rounded-lg shadow-xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3">
              <h3 className="font-display font-bold text-sm text-slate-900 flex items-center gap-2">
                <Trash2 className="w-4 h-4 text-red-600" />
                Clear Master Database Repository
              </h3>
              <button onClick={() => setShowClearModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-sm">✕</button>
            </div>
            
            <div className="space-y-3 text-xs">
              <p className="text-slate-600">
                Choose an option to manage the Boon Huat Purchase Order & GRN master database:
              </p>

              <div className="space-y-2.5 pt-1">
                <button
                  onClick={handleConfirmClearDatabase}
                  className="w-full text-left p-3.5 rounded-lg border border-red-200 bg-red-50 hover:bg-red-100/80 transition-colors flex items-start gap-3 group"
                >
                  <div className="p-2 bg-red-600 text-white rounded shrink-0 mt-0.5 shadow-xs">
                    <Trash2 className="w-4 h-4" />
                  </div>
                  <div>
                    <h4 className="font-bold text-red-950 text-xs group-hover:underline">Clear Database Completely (0 Records)</h4>
                    <p className="text-[11px] text-red-800 mt-0.5">Wipes all POs and GRNs from memory so you can upload a fresh master Excel sheet.</p>
                  </div>
                </button>

                <button
                  onClick={handleResetSampleDatabase}
                  className="w-full text-left p-3.5 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100/80 transition-colors flex items-start gap-3 group"
                >
                  <div className="p-2 bg-indigo-600 text-white rounded shrink-0 mt-0.5 shadow-xs">
                    <Database className="w-4 h-4" />
                  </div>
                  <div>
                    <h4 className="font-bold text-indigo-950 text-xs group-hover:underline">Reset to Demo Sample Data</h4>
                    <p className="text-[11px] text-indigo-800 mt-0.5">Restores the standard sample Boon Huat hardware PO and GRN records.</p>
                  </div>
                </button>
              </div>

              <div className="pt-2 flex justify-end">
                <button 
                  onClick={() => setShowClearModal(false)}
                  className="py-1.5 px-4 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded font-semibold text-xs"
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: GOOGLE SHEETS SYNC & EXPORT */}
      <GoogleSheetsModal
        isOpen={showGoogleSheetsModal}
        onClose={() => setShowGoogleSheetsModal(false)}
        onImportInvoices={handleGoogleSheetsImport}
        exportData={{
          auditedInvoices,
          actionTimestamp: actionTimestamp || new Date().toLocaleString('en-SG')
        }}
      />

    </div>
  );
}
