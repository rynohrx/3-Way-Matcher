import { PurchaseOrder, GoodsReceivedNote, SupplierInvoice } from './types';

export interface GoogleSheetsDb {
  purchaseOrders: PurchaseOrder[];
  goodsReceivedNotes: GoodsReceivedNote[];
  pastInvoices: SupplierInvoice[];
}

// Clean empty database - populated by user uploading Boon Huat PO/GRN database files or documents
export const INITIAL_DATABASE: GoogleSheetsDb = {
  purchaseOrders: [],
  goodsReceivedNotes: [],
  pastInvoices: []
};
