/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface PurchaseOrder {
  id: string; // PO Number
  supplierName: string;
  purchaseDate: string;
  itemDescription: string;
  quantityOrdered: number;
  unitPrice: number;
  totalAmount: number;
}

export interface GoodsReceivedNote {
  id: string; // GRN Number
  poNumber: string;
  dateReceived: string;
  itemDescription: string;
  quantityReceived: number;
}

export interface SupplierInvoice {
  id: string; // Invoice Number
  invoiceDate: string;
  supplierName: string;
  poNumber?: string;
  itemDescription: string;
  quantityBilled: number;
  unitPrice: number;
  totalAmount: number;
}

export type MatchStatus = 'PASS' | 'FAIL' | 'WARNING';

export interface MatchCheckItem {
  name: string;
  status: MatchStatus;
  details: string; // e.g. "PO: 100, GRN: 95, Invoice: 100"
  explanation?: string; // plain-English explanation for Madam Lim
}

export type OverallDecision = 'APPROVED' | 'MANUAL REVIEW REQUIRED' | 'REJECT PAYMENT';

export interface MatchReport {
  overallResult: OverallDecision;
  checks: {
    supplierName: MatchCheckItem;
    poNumber: MatchCheckItem;
    quantity: MatchCheckItem;
    unitPrice: MatchCheckItem;
    totalAmount: MatchCheckItem;
    duplicateInvoice: MatchCheckItem;
    calculation: MatchCheckItem;
    grnPresence: MatchCheckItem;
  };
  explanation: string;
  recommendedAction: string;
  responsibleAIAudit: {
    supplierEquityChecked: boolean;
    evidenceBasedOnly: boolean;
    biasMitigationMessage: string;
  };
}

export interface DemoScenario {
  id: string;
  title: string;
  description: string;
  po: PurchaseOrder | null;
  grn: GoodsReceivedNote | null;
  invoice: SupplierInvoice;
  expectedStatus: OverallDecision;
}

export interface AuditedInvoiceItem {
  id: string;
  invoice: SupplierInvoice;
  po: PurchaseOrder | null;
  grn: GoodsReceivedNote | null;
  matchReport: MatchReport;
  recommendation: 'ACCEPT' | 'REJECT' | 'FORWARD';
  recommendationSummary: string;
  actionStatus: 'APPROVED' | 'REJECTED' | 'REVIEW_SENT' | null;
}
