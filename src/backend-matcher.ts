import { 
  PurchaseOrder, 
  GoodsReceivedNote, 
  SupplierInvoice, 
  MatchReport, 
  MatchCheckItem, 
  MatchStatus, 
  OverallDecision 
} from './types';
import { GoogleSheetsDb } from './db-mock';

// Utility to normalize strings for comparison (lowercase, remove excess spacing)
function normalizeStr(str: string): string {
  return str.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

// Convert Excel serial date numbers (e.g. 46218) into YYYY-MM-DD date strings
export function parseExcelDateOrString(val: any, fallbackDate?: string): string {
  if (val === undefined || val === null || val === '') {
    return fallbackDate || new Date().toISOString().split('T')[0];
  }
  const str = String(val).trim();
  const num = Number(str);
  if (!isNaN(num) && num > 25000 && num < 70000) {
    const jsDate = new Date((num - (25567 + 2)) * 86400 * 1000);
    if (!isNaN(jsDate.getTime())) {
      return jsDate.toISOString().split('T')[0];
    }
  }
  return str;
}

// Robust number cleaner for currency strings ($1,200.00, 1,200, etc.)
export function cleanNum(val: any, fallback: number = 0): number {
  if (val === undefined || val === null || val === '') return fallback;
  if (typeof val === 'number') return isNaN(val) ? fallback : val;
  const str = String(val).trim();
  if (!str) return fallback;

  // Strip currency symbols, commas, spaces and common quantity/currency 3-letter words
  const stripCurrency = str.replace(/[$€£¥,]|\b(usd|sgd|eur|cad|aud|gbp|rmb|myr|idr|php|pcs|pc|units|unit|boxes|box|pkgs|pkg|ctn|ctns|ea)\b/gi, '').trim();
  // If the string still contains alphabetic characters (e.g. "Box of 100", "M12 Bolts"), it is a text description, not a standalone number!
  if (/[a-zA-Z]/.test(stripCurrency)) {
    return fallback;
  }

  const cleanStr = str.replace(/[^0-9.-]/g, '');
  if (!cleanStr || cleanStr === '-' || cleanStr === '.') return fallback;
  const parsed = parseFloat(cleanStr);
  return isNaN(parsed) ? fallback : parsed;
}

/**
 * Resolves Quantity vs Unit Price when reading columns or row tokens.
 * Prevents Quantity Ordered and Unit Price from being inverted/swapped.
 */
export function resolveQtyAndPrice(
  rawVal1: any, 
  rawVal2: any, 
  totalVal?: any
): { qty: number; unitPrice: number; totalAmount: number } {
  let v1 = cleanNum(rawVal1, 0);
  let v2 = cleanNum(rawVal2, 0);
  let total = cleanNum(totalVal, 0);

  let qty = v1;
  let unitPrice = v2;

  // Case 1: If v1 (qty position) has decimals (e.g. 18.50) while v2 is a whole integer (e.g. 100),
  // then v1 is actually the Unit Price ($18.50) and v2 is the Quantity (100). Swap them!
  if (v1 > 0 && v2 > 0) {
    const v1IsInteger = Number.isInteger(v1);
    const v2IsInteger = Number.isInteger(v2);

    if (!v1IsInteger && v2IsInteger) {
      qty = v2;
      unitPrice = v1;
    }
  }

  // Calculate or adjust total amount
  if (total === 0 && qty > 0 && unitPrice > 0) {
    total = Number((qty * unitPrice).toFixed(2));
  } else if (total > 0 && qty > 0 && unitPrice === 0) {
    unitPrice = Number((total / qty).toFixed(2));
  } else if (total > 0 && unitPrice > 0 && qty === 0) {
    qty = Math.round(total / unitPrice) || 1;
  }

  return { qty, unitPrice, totalAmount: total };
}

/**
 * Intelligent Supplier Name Matching Helper
 * Cleans corporate entity suffixes (Pte Ltd, Co, Corp, Inc, etc.), handles token overlaps,
 * and performs fuzzy string distance comparison to eliminate false mismatches.
 */
export function compareSupplierNames(name1: string, name2: string): { match: boolean; confidence: 'EXACT' | 'HIGH' | 'MEDIUM' | 'NONE'; details: string } {
  if (!name1 || !name2) {
    return { match: false, confidence: 'NONE', details: 'One or both supplier names missing' };
  }

  const raw1 = name1.trim();
  const raw2 = name2.trim();

  if (raw1.toLowerCase() === raw2.toLowerCase()) {
    return { match: true, confidence: 'EXACT', details: `Exact match: "${raw1}"` };
  }

  // Helper to remove corporate entity noise and normalize
  const clean = (str: string) => {
    let s = str.toLowerCase();
    // replace punctuation with spaces
    s = s.replace(/[^a-z0-9\s]/g, ' ');
    
    // Entity noise words/suffixes to strip
    const noiseWords = [
      'private limited', 'pte ltd', 'pte', 'ltd', 'limited',
      'co ltd', 'company', 'co', 'corp', 'corporation', 'inc', 'incorporated',
      'llc', 'sdn bhd', 'bhd', 'trading', 'enterprise', 'enterprises',
      'supplies', 'supply', 'hardware', 'services', 'singapore', 'asia'
    ];

    for (const word of noiseWords) {
      const reg = new RegExp(`\\b${word}\\b`, 'gi');
      s = s.replace(reg, ' ');
    }

    return s.replace(/\s+/g, ' ').trim();
  };

  const clean1 = clean(raw1);
  const clean2 = clean(raw2);

  if (clean1 && clean2) {
    if (clean1 === clean2) {
      return { match: true, confidence: 'HIGH', details: `Matched brand core: "${clean1.toUpperCase()}"` };
    }

    if (clean1.includes(clean2) || clean2.includes(clean1)) {
      return { match: true, confidence: 'HIGH', details: `Substring match between "${raw1}" and "${raw2}"` };
    }

    // Token comparison
    const tokens1 = clean1.split(' ').filter(t => t.length > 1);
    const tokens2 = clean2.split(' ').filter(t => t.length > 1);

    if (tokens1.length > 0 && tokens2.length > 0) {
      const commonTokens = tokens1.filter(t => tokens2.includes(t));
      const minTokens = Math.min(tokens1.length, tokens2.length);
      const maxTokens = Math.max(tokens1.length, tokens2.length);

      // If all tokens of shorter name are present in longer name
      if (commonTokens.length === minTokens && minTokens >= 1) {
        return { match: true, confidence: 'HIGH', details: `Matched core tokens: "${commonTokens.join(' ')}"` };
      }

      // If at least 50% of tokens match in both names
      const overlapRatio = commonTokens.length / maxTokens;
      if (overlapRatio >= 0.5) {
        return { match: true, confidence: 'MEDIUM', details: `Token overlap: ${commonTokens.join(', ')}` };
      }
    }

    // Edit distance check on cleaned strings
    const dist = levenshteinDistance(clean1, clean2);
    const maxLen = Math.max(clean1.length, clean2.length);
    const similarity = 1 - dist / maxLen;

    if (maxLen > 3 && similarity >= 0.65) {
      return { match: true, confidence: 'MEDIUM', details: `Similar supplier names (${Math.round(similarity * 100)}% match)` };
    }
  }

  return { match: false, confidence: 'NONE', details: `Invoice: "${raw1}" | PO: "${raw2}"` };
}

function levenshteinDistance(a: string, b: string): number {
  const matrix: number[][] = [];
  for (let i = 0; i <= b.length; i++) matrix[i] = [i];
  for (let j = 0; j <= a.length; j++) matrix[0][j] = j;

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

export function runThreeWayMatch(
  rawInvoice: SupplierInvoice,
  db: GoogleSheetsDb,
  manualPo?: PurchaseOrder | null,
  manualGrn?: GoodsReceivedNote | null,
  isAlreadyApproved?: boolean
): MatchReport {
  
  // Clean and sanitize all numerical values on the incoming invoice
  const rawInvQty = cleanNum(rawInvoice.quantityBilled, 1);
  const rawInvPrice = cleanNum(rawInvoice.unitPrice, 0);
  const rawInvTotal = cleanNum(rawInvoice.totalAmount, 0);

  const resolvedInv = resolveQtyAndPrice(rawInvQty, rawInvPrice, rawInvTotal);

  const invoice: SupplierInvoice = {
    ...rawInvoice,
    quantityBilled: resolvedInv.qty,
    unitPrice: resolvedInv.unitPrice,
    totalAmount: resolvedInv.totalAmount
  };
  
  // 1. Retrieve or use manual PO
  let po: PurchaseOrder | null = null;
  if (manualPo) {
    po = manualPo;
  } else if (invoice.poNumber) {
    const normInvPo = invoice.poNumber.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    po = db.purchaseOrders.find(p => p.id.replace(/[^A-Za-z0-9]/g, '').toUpperCase() === normInvPo) || null;
  }

  if (po) {
    const pQty = cleanNum(po.quantityOrdered, 1);
    const pPrice = cleanNum(po.unitPrice, 0);
    const pTotal = cleanNum(po.totalAmount, 0);
    
    const resolvedPo = resolveQtyAndPrice(pQty, pPrice, pTotal);

    po = {
      ...po,
      quantityOrdered: resolvedPo.qty,
      unitPrice: resolvedPo.unitPrice,
      totalAmount: resolvedPo.totalAmount
    };
  }

  // 2. Retrieve or use manual GRN
  let grn: GoodsReceivedNote | null = null;
  if (manualGrn) {
    grn = manualGrn;
  } else if (po) {
    const normPoId = po.id.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    grn = db.goodsReceivedNotes.find(g => g.poNumber.replace(/[^A-Za-z0-9]/g, '').toUpperCase() === normPoId) || null;
  } else if (invoice.poNumber) {
    const normInvPo = invoice.poNumber.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    grn = db.goodsReceivedNotes.find(g => g.poNumber.replace(/[^A-Za-z0-9]/g, '').toUpperCase() === normInvPo) || null;
  }

  if (grn) {
    grn = {
      ...grn,
      quantityReceived: cleanNum(grn.quantityReceived, 0)
    };
  }

  const failures: string[] = [];
  const warnings: string[] = [];

  // --- CHECK 1: PO Existence ---
  let poCheck: MatchCheckItem = {
    name: "Purchase Order Reference",
    status: "PASS",
    details: invoice.poNumber ? `Invoice references PO: ${invoice.poNumber}` : "No PO reference found on invoice."
  };
  if (!invoice.poNumber) {
    poCheck.status = "FAIL";
    poCheck.details = "No Purchase Order number specified on the invoice.";
    
    // Check if a candidate PO exists for this supplier
    const candidatePo = db.purchaseOrders.find(p => compareSupplierNames(p.supplierName, invoice.supplierName).match);
    if (candidatePo) {
      poCheck.explanation = `The invoice does not list a PO number. However, we found an existing Purchase Order (${candidatePo.id}) for supplier "${candidatePo.supplierName}" in our Google Sheets database. Madam Lim should verify if this invoice belongs to PO ${candidatePo.id}.`;
    } else {
      poCheck.explanation = "The invoice does not list any Purchase Order number. Madam Lim should consult with the purchasing department to verify if an authorized PO was ever created.";
    }
    failures.push("Missing PO reference");
  } else if (!po) {
    poCheck.status = "FAIL";
    poCheck.details = `PO "${invoice.poNumber}" could not be found in our Google Sheets records.`;
    
    // Check if supplier has other POs
    const candidatePo = db.purchaseOrders.find(p => compareSupplierNames(p.supplierName, invoice.supplierName).match);
    if (candidatePo) {
      poCheck.explanation = `The invoice references PO "${invoice.poNumber}", which does not exist in our system. However, supplier "${invoice.supplierName}" has PO ${candidatePo.id} on record. Please check for a typo in the PO number.`;
    } else {
      poCheck.explanation = `The invoice references PO number "${invoice.poNumber}", which does not exist in our system. Please check if there is a typo in the invoice or if the PO has not been logged in the sheet yet.`;
    }
    failures.push("PO not found");
  }

  // --- CHECK 2: Supplier Match ---
  let supplierCheck: MatchCheckItem = {
    name: "Supplier Verification",
    status: "PASS",
    details: `Invoice Supplier: "${invoice.supplierName}"`
  };
  if (po) {
    const rawPoSup = po.supplierName ? String(po.supplierName).trim() : '';
    const isPoSupNumericOrDate = !rawPoSup || !isNaN(Number(rawPoSup)) || rawPoSup.toLowerCase() === 'supplier' || rawPoSup.toLowerCase() === 'vendor';
    const isInvSupValid = invoice.supplierName && isNaN(Number(invoice.supplierName.trim()));

    if (isPoSupNumericOrDate && isInvSupValid) {
      // Repair PO supplier name to match the supplier name provided on the invoice
      po.supplierName = invoice.supplierName;
      supplierCheck.status = "PASS";
      supplierCheck.details = `Matched with PO: "${invoice.supplierName}" (PO supplier field "${rawPoSup}" resolved to "${invoice.supplierName}")`;
      supplierCheck.explanation = `The Purchase Order record previously contained an unparsed column value ("${rawPoSup}"). We auto-aligned the PO supplier name to match the supplier name provided on the invoice ("${invoice.supplierName}").`;
    } else {
      const comp = compareSupplierNames(invoice.supplierName, po.supplierName);
      if (!comp.match) {
        supplierCheck.status = "FAIL";
        supplierCheck.details = `Invoice Supplier: "${invoice.supplierName}" | PO Supplier: "${po.supplierName}"`;
        supplierCheck.explanation = `Supplier mismatch detected! The invoice is billed by "${invoice.supplierName}", but the corresponding Purchase Order "${po.id}" was authorized for "${po.supplierName}". This requires immediate investigation.`;
        failures.push("Supplier name mismatch");
      } else {
        if (comp.confidence === 'EXACT') {
          supplierCheck.details = `Matched with PO: "${po.supplierName}"`;
        } else {
          supplierCheck.details = `Matched with PO: "${po.supplierName}" (${comp.details})`;
        }
      }
    }
  } else {
    supplierCheck.status = "WARNING";
    supplierCheck.details = `Billed by "${invoice.supplierName}" (Cannot verify against PO)`;
    supplierCheck.explanation = "Since the Purchase Order was not found, we cannot verify whether this supplier is authorized for this purchase.";
    warnings.push("Cannot verify supplier");
  }

  // --- CHECK 3: GRN Presence ---
  let grnPresenceCheck: MatchCheckItem = {
    name: "Goods Received Note",
    status: "PASS",
    details: grn ? `Found matching GRN: ${grn.id}` : "No GRN found."
  };
  if (po && !grn) {
    grnPresenceCheck.status = "FAIL";
    grnPresenceCheck.details = `No Goods Received Note found for PO ${po.id}`;
    grnPresenceCheck.explanation = `Although Purchase Order "${po.id}" exists, there is no record of a Goods Received Note (GRN) in our database. The goods may still be in transit or the warehouse team has not yet processed the receiving slip.`;
    failures.push("GRN missing");
  }

  // --- CHECK 4: Quantity Match ---
  let quantityCheck: MatchCheckItem = {
    name: "Quantity Validation",
    status: "PASS",
    details: `Billed: ${invoice.quantityBilled}`
  };
  
  if (po && grn) {
    const invQty = invoice.quantityBilled;
    const poQty = po.quantityOrdered;
    const grnQty = grn.quantityReceived;

    quantityCheck.details = `Billed: ${invQty} | Ordered (PO): ${poQty} | Received (GRN): ${grnQty}`;

    if (invQty > grnQty) {
      quantityCheck.status = "FAIL";
      if (grnQty === 0) {
        quantityCheck.explanation = `Goods were not received. The invoice bills for ${invQty} units of "${invoice.itemDescription}", but the Goods Received Note shows 0 units were delivered. Payment should be withheld.`;
        failures.push("Goods not received");
      } else {
        quantityCheck.explanation = `Goods billed exceed goods received. The invoice requests payment for ${invQty} units, but the warehouse only recorded receiving ${grnQty} units. Madam Lim should contact the supplier to adjust the invoice, or check with the warehouse if more goods arrived later.`;
        failures.push("Quantity billed exceeds received");
      }
    } else if (invQty > poQty) {
      quantityCheck.status = "WARNING";
      quantityCheck.explanation = `The invoice bills for ${invQty} units, which is more than the approved PO quantity of ${poQty} units. However, the warehouse received ${grnQty} units. Madam Lim should verify if this over-delivery was authorized by management.`;
      warnings.push("Invoice quantity exceeds PO");
    } else if (grnQty < poQty && invQty === poQty) {
      quantityCheck.status = "FAIL";
      quantityCheck.explanation = `The supplier billed for the full ordered quantity (${invQty} units), but the warehouse only recorded receiving ${grnQty} units (short-delivery). Payment should be held or adjusted to reflect only the received quantity.`;
      failures.push("Billed for short-delivered goods");
    }
  } else if (po && !grn) {
    quantityCheck.status = "FAIL";
    quantityCheck.details = `Billed: ${invoice.quantityBilled} | Ordered: ${po.quantityOrdered} | Received: No GRN`;
    quantityCheck.explanation = `We cannot verify received quantities because no Goods Received Note is available in our database. Goods might not have arrived.`;
    failures.push("Quantity cannot be verified (no GRN)");
  } else {
    quantityCheck.status = "WARNING";
    quantityCheck.details = `Billed: ${invoice.quantityBilled} (No PO or GRN to compare)`;
    quantityCheck.explanation = "Without PO and GRN documents, we cannot verify if this billed quantity is correct or if the goods were received.";
    warnings.push("Quantity unverified");
  }

  // --- CHECK 5: Unit Price Match ---
  let priceCheck: MatchCheckItem = {
    name: "Unit Price Validation",
    status: "PASS",
    details: `Invoice Price: $${invoice.unitPrice.toFixed(2)}`
  };
  if (po) {
    const diff = invoice.unitPrice - po.unitPrice;
    if (Math.abs(diff) > 0.01) {
      priceCheck.status = "FAIL";
      priceCheck.details = `Invoice Price: $${invoice.unitPrice.toFixed(2)} | PO Approved Price: $${po.unitPrice.toFixed(2)}`;
      priceCheck.explanation = `Invoice price differs from approved purchase order. The invoice charges $${invoice.unitPrice.toFixed(2)} per unit, but the approved PO price was $${po.unitPrice.toFixed(2)}. This results in an unauthorized overcharge of $${diff.toFixed(2)} per unit. Madam Lim should ask the supplier for an amended invoice matching the PO.`;
      failures.push("Unit price mismatch");
    } else {
      priceCheck.details = `Matched with PO approved price ($${po.unitPrice.toFixed(2)})`;
    }
  } else {
    priceCheck.status = "WARNING";
    priceCheck.details = `Invoice Price: $${invoice.unitPrice.toFixed(2)} (No PO to compare)`;
    priceCheck.explanation = "We cannot verify if this unit price is correct because no authorized Purchase Order is available to compare against.";
    warnings.push("Price unverified");
  }

  // --- CHECK 6: Invoice Total Calculation ---
  let calcCheck: MatchCheckItem = {
    name: "Calculation Accuracy",
    status: "PASS",
    details: `Billed: ${invoice.quantityBilled} × $${invoice.unitPrice.toFixed(2)} = $${invoice.totalAmount.toFixed(2)}`
  };
  const calculatedTotal = invoice.quantityBilled * invoice.unitPrice;
  const calcDiff = Math.abs(invoice.totalAmount - calculatedTotal);
  if (calcDiff > 0.05) { // allow a 5-cent buffer for rounding
    calcCheck.status = "FAIL";
    calcCheck.details = `Invoice states: $${invoice.totalAmount.toFixed(2)} | Math says: ${invoice.quantityBilled} × $${invoice.unitPrice.toFixed(2)} = $${calculatedTotal.toFixed(2)}`;
    calcCheck.explanation = `An arithmetic error was found on the invoice. Billed quantity (${invoice.quantityBilled}) multiplied by unit price ($${invoice.unitPrice.toFixed(2)}) equals $${calculatedTotal.toFixed(2)}, but the invoice lists the total as $${invoice.totalAmount.toFixed(2)}. This invoice is invalid and must be corrected by the supplier.`;
    failures.push("Invoice calculation error");
  }

  // --- CHECK 7: Duplicate Invoice Check ---
  let duplicateCheck: MatchCheckItem = {
    name: "Duplicate Invoice Search",
    status: "PASS",
    details: "No previous invoice matches found."
  };

  if (isAlreadyApproved) {
    duplicateCheck.status = "PASS";
    duplicateCheck.details = "Approved & logged in Boon Huat database.";
    duplicateCheck.explanation = "This invoice has been approved by Madam Lim and logged into the database.";
  } else {
    // Search in Sheets (pastInvoices)
    const duplicateById = db.pastInvoices.find(inv => inv.id.toUpperCase() === invoice.id.toUpperCase());
    const duplicateByMetadata = db.pastInvoices.find(inv => 
      compareSupplierNames(inv.supplierName, invoice.supplierName).match &&
      Math.abs(inv.totalAmount - invoice.totalAmount) < 0.01 &&
      inv.invoiceDate === invoice.invoiceDate
    );

    if (duplicateById) {
      duplicateCheck.status = "FAIL";
      duplicateCheck.details = `An invoice with ID "${invoice.id}" already exists in our Google Sheets database (recorded on ${duplicateById.invoiceDate}).`;
      duplicateCheck.explanation = `Possible duplicate invoice. Invoice number "${invoice.id}" was already submitted and approved/paid in a past session. Madam Lim should verify if this is an accidental double-submission by the supplier.`;
      failures.push("Duplicate invoice ID");
    } else if (duplicateByMetadata) {
      duplicateCheck.status = "FAIL";
      duplicateCheck.details = `An invoice for $${invoice.totalAmount.toFixed(2)} from "${invoice.supplierName}" dated ${invoice.invoiceDate} already exists under ID "${duplicateByMetadata.id}".`;
      duplicateCheck.explanation = `Possible duplicate invoice. Although the invoice number is different, we found another invoice from "${invoice.supplierName}" on the same date (${invoice.invoiceDate}) for the exact same amount of $${invoice.totalAmount.toFixed(2)}. This could be a double billing error.`;
      failures.push("Duplicate supplier, amount, and date");
    }
  }

  // --- CHECK 8: Item Description Check ---
  let totalAmountCheck: MatchCheckItem = {
    name: "Total Amount Comparison",
    status: "PASS",
    details: `Invoice Total: $${invoice.totalAmount.toFixed(2)}`
  };
  if (po) {
    const diff = invoice.totalAmount - po.totalAmount;
    if (Math.abs(diff) > 0.05) {
      totalAmountCheck.status = diff > 0 ? "FAIL" : "WARNING";
      totalAmountCheck.details = `Invoice Total: $${invoice.totalAmount.toFixed(2)} | PO Approved Total: $${po.totalAmount.toFixed(2)}`;
      totalAmountCheck.explanation = diff > 0 
        ? `The invoice total of $${invoice.totalAmount.toFixed(2)} exceeds the PO approved amount of $${po.totalAmount.toFixed(2)} by $${diff.toFixed(2)}. Payment should not be released without an approved change order.`
        : `The invoice total of $${invoice.totalAmount.toFixed(2)} is less than the PO approved amount of $${po.totalAmount.toFixed(2)} (difference of $${Math.abs(diff).toFixed(2)}). This is acceptable as long as it matches the received goods.`;
      
      if (diff > 0) {
        failures.push("Invoice total exceeds approved PO amount");
      } else {
        warnings.push("Invoice total is less than PO total");
      }
    } else {
      totalAmountCheck.details = `Matched with PO approved total ($${po.totalAmount.toFixed(2)})`;
    }
  } else {
    totalAmountCheck.status = "WARNING";
    totalAmountCheck.details = `Invoice Total: $${invoice.totalAmount.toFixed(2)} (No PO to compare)`;
    totalAmountCheck.explanation = "We cannot verify this total amount against authorized purchasing records because the Purchase Order is missing.";
    warnings.push("Amount unverified");
  }

  // --- COMPUTE OVERALL DECISION ---
  let overallResult: OverallDecision = "APPROVED";
  let explanation = "";
  let recommendedAction = "";

  if (failures.length > 0) {
    overallResult = "REJECT PAYMENT";
    explanation = `Matching failed due to ${failures.length} critical issues, including: ${failures.join(", ")}.`;
    recommendedAction = `Payment should be REJECTED. Madam Lim should contact ${invoice.supplierName} immediately and ask for a corrected invoice, or follow up with the warehouse / purchasing department to resolve document mismatches.`;
  } else if (warnings.length > 0) {
    overallResult = "MANUAL REVIEW REQUIRED";
    explanation = `Documents match, but there are ${warnings.length} warning flags that need human verification: ${warnings.join(", ")}.`;
    recommendedAction = `Madam Lim should manually review the details before making a decision. Verification with the warehouse supervisor or the buyer is recommended.`;
  } else {
    overallResult = "APPROVED";
    explanation = `All matching checks passed successfully! The Supplier Invoice fully matches the Purchase Order and the Goods Received Note.`;
    recommendedAction = `Payment is safe to proceed. Madam Lim can confidently click 'Approve' to authorize payment.`;
  }

  // Specific explanations override for nice presets
  if (failures.includes("Goods not received") || failures.includes("Quantity billed exceeds received") || failures.includes("Billed for short-delivered goods")) {
    explanation = `The supplier has billed for ${invoice.quantityBilled} units, but our warehouse records show we only received ${grn ? grn.quantityReceived : 0} units. This violates the 3-Way Match rule.`;
    recommendedAction = `Madam Lim should reject the payment and instruct the supplier to issue a credit note or a revised invoice reflecting only the ${grn ? grn.quantityReceived : 0} units actually received.`;
  } else if (failures.includes("Unit price mismatch")) {
    explanation = `The invoice unit price of $${invoice.unitPrice.toFixed(2)} is higher than the approved Purchase Order price of $${po ? po.unitPrice.toFixed(2) : 0.00}. This represents an unauthorized overcharge.`;
    recommendedAction = `Reject payment and hold the transaction. Contact "${invoice.supplierName}" to clarify the pricing discrepancy and request a revised invoice matching our authorized PO rates.`;
  } else if (failures.includes("Duplicate invoice ID") || failures.includes("Duplicate supplier, amount, and date")) {
    explanation = `Our automated checks have flagged this document as a duplicate invoice. An invoice with this ID (${invoice.id}) or matching metadata already exists in our records as paid.`;
    recommendedAction = `REJECT payment. This is likely an accidental double-billing. Madam Lim should verify in the Past Supplier Invoices log and notify the vendor that this invoice has already been settled.`;
  } else if (failures.includes("GRN missing")) {
    explanation = `We have an approved PO for these goods, but we cannot find any Goods Received Note (GRN) in our files. The goods might not have been delivered to our yard yet.`;
    recommendedAction = `Mark as Send for Review or Hold. Madam Lim should double check with the receiving yard supervisor (e.g., Ah Huat) to verify if the physical shipment arrived before approving any payment.`;
  }

  return {
    overallResult,
    checks: {
      poNumber: poCheck,
      supplierName: supplierCheck,
      quantity: quantityCheck,
      unitPrice: priceCheck,
      totalAmount: totalAmountCheck,
      duplicateInvoice: duplicateCheck,
      calculation: calcCheck,
      grnPresence: grnPresenceCheck
    },
    explanation,
    recommendedAction,
    responsibleAIAudit: {
      supplierEquityChecked: true,
      evidenceBasedOnly: true,
      biasMitigationMessage: `Boon Huat AI operates strictly under evidence-based rules. No bias weighting or history adjustments are applied. Both long-term partners (like Sin Siah Metal Works) and new suppliers are checked using the exact same compliance metrics.`
    }
  };
}
