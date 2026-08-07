import express from "express";
import path from "path";
import dotenv from "dotenv";
import { GoogleGenAI, Type } from "@google/genai";
import { runThreeWayMatch } from "./src/backend-matcher";
import { INITIAL_DATABASE, GoogleSheetsDb } from "./src/db-mock";
import { SupplierInvoice, PurchaseOrder, GoodsReceivedNote } from "./src/types";

// Load environment variables
dotenv.config();

// Keep database in server memory for active session updates
let db: GoogleSheetsDb = { ...INITIAL_DATABASE };

const app = express();
const PORT = 3000;

// Set up JSON parsing with generous limit for document image uploads (Base64)
app.use(express.json({ limit: '20mb' }));

// Lazy initialize Gemini API client to prevent crashes if key is missing
let aiClient: GoogleGenAI | null = null;
function getGemini(): GoogleGenAI | null {
  if (!aiClient) {
    const key = process.env.GEMINI_API_KEY;
    if (key && key !== "MY_GEMINI_API_KEY" && key.trim() !== "") {
      aiClient = new GoogleGenAI({
        apiKey: key,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });
      console.log("Gemini client successfully initialized.");
    } else {
      console.warn("GEMINI_API_KEY is not configured or is placeholder. Running in Simulation/Demo mode.");
    }
  }
  return aiClient;
}

// ==========================================
// API ENDPOINTS
// ==========================================

// 1. Check server health & mode (Gemini vs Demo simulation)
app.get("/api/health", (req, res) => {
  const isGeminiAvailable = getGemini() !== null;
  res.json({
    status: "ok",
    geminiMode: isGeminiAvailable ? "active" : "simulation",
    currentTime: new Date().toISOString()
  });
});

function sanitizeServerDb(database: GoogleSheetsDb): GoogleSheetsDb {
  if (!database) return { purchaseOrders: [], goodsReceivedNotes: [], pastInvoices: [] };

  const isHeaderString = (str: string) => {
    if (!str) return true;
    const u = String(str).toUpperCase().trim();
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
  (database.purchaseOrders || []).forEach(p => {
    if (p && p.id && !isHeaderString(p.id) && !poSeen.has(p.id)) {
      poSeen.add(p.id);
      cleanPos.push(p);
    }
  });

  const cleanGrns: GoodsReceivedNote[] = [];
  const grnSeen = new Set<string>();
  (database.goodsReceivedNotes || []).forEach(g => {
    if (g && g.id && !isHeaderString(g.id) && !grnSeen.has(g.id)) {
      grnSeen.add(g.id);
      cleanGrns.push(g);
    }
  });

  const cleanInvoices: SupplierInvoice[] = [];
  const invSeen = new Set<string>();
  (database.pastInvoices || []).forEach(i => {
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

// 2. Fetch the "Google Sheets" database state
app.get("/api/sheets", (req, res) => {
  db = sanitizeServerDb(db);
  res.json(db);
});

// 3. Reset or Clear the "Google Sheets" database
app.post("/api/sheets/reset", (req, res) => {
  db = sanitizeServerDb(JSON.parse(JSON.stringify(INITIAL_DATABASE)));
  res.json({ message: "Database reset to initial state", db });
});

app.post("/api/sheets/clear", (req, res) => {
  db = { purchaseOrders: [], goodsReceivedNotes: [], pastInvoices: [] };
  res.json({ message: "Database cleared", db });
});

// 4. Update or Import into the "Google Sheets" database
app.post("/api/sheets/add", (req, res) => {
  if (!db) db = { purchaseOrders: [], goodsReceivedNotes: [], pastInvoices: [] };
  if (!Array.isArray(db.purchaseOrders)) db.purchaseOrders = [];
  if (!Array.isArray(db.goodsReceivedNotes)) db.goodsReceivedNotes = [];
  if (!Array.isArray(db.pastInvoices)) db.pastInvoices = [];

  const { type, data } = req.body;
  if (type === "pastInvoices" && data) {
    db.pastInvoices.unshift(data as SupplierInvoice);
  } else if (type === "purchaseOrders" && data) {
    db.purchaseOrders.unshift(data as PurchaseOrder);
  } else if (type === "goodsReceivedNotes" && data) {
    db.goodsReceivedNotes.unshift(data as GoodsReceivedNote);
  }
  db = sanitizeServerDb(db);
  res.json({ message: "Successfully logged to Google Sheet", db });
});

app.post("/api/sheets/import", (req, res) => {
  if (!db) db = { purchaseOrders: [], goodsReceivedNotes: [], pastInvoices: [] };
  if (!Array.isArray(db.purchaseOrders)) db.purchaseOrders = [];
  if (!Array.isArray(db.goodsReceivedNotes)) db.goodsReceivedNotes = [];
  if (!Array.isArray(db.pastInvoices)) db.pastInvoices = [];

  const { purchaseOrders, goodsReceivedNotes, pastInvoices, replace } = req.body;
  
  if (replace) {
    db = {
      purchaseOrders: Array.isArray(purchaseOrders) ? purchaseOrders : [],
      goodsReceivedNotes: Array.isArray(goodsReceivedNotes) ? goodsReceivedNotes : [],
      pastInvoices: Array.isArray(pastInvoices) ? pastInvoices : []
    };
  } else {
    if (Array.isArray(purchaseOrders) && purchaseOrders.length > 0) {
      db.purchaseOrders = [...purchaseOrders, ...db.purchaseOrders];
    }
    if (Array.isArray(goodsReceivedNotes) && goodsReceivedNotes.length > 0) {
      db.goodsReceivedNotes = [...goodsReceivedNotes, ...db.goodsReceivedNotes];
    }
    if (Array.isArray(pastInvoices) && pastInvoices.length > 0) {
      db.pastInvoices = [...pastInvoices, ...db.pastInvoices];
    }
  }
  
  db = sanitizeServerDb(db);
  res.json({ message: "Database successfully updated", db });
});

// 5. Extract fields from uploaded document using Gemini or Fallback simulator
app.post("/api/extract", async (req, res) => {
  try {
    const { fileBase64, fileName, mimeType, docType } = req.body;

    if (!docType) {
      return res.status(400).json({ error: "docType ('invoice' | 'po' | 'grn') is required" });
    }

    const ai = getGemini();

    if (ai) {
      // --- REAL GEMINI PARSING ---
      console.log(`Analyzing document with Gemini. Type: ${docType}, Name: ${fileName}`);

      // Strip metadata from base64 if present
      let rawBase64 = fileBase64;
      if (fileBase64.includes(";base64,")) {
        rawBase64 = fileBase64.split(";base64,")[1];
      }

      const mediaPart = {
        inlineData: {
          mimeType: mimeType || "image/jpeg",
          data: rawBase64
        }
      };

      let schema: any;
      let promptText = "";

      if (docType === "invoice") {
        promptText = `You are an expert accounts assistant for Boon Huat Hardware & Supplies Pte Ltd.
Analyze this invoice and extract the details. It may be a PDF, scanned document, or handwritten slip.
If fields like PO Number are not present, leave them blank. Ensure date formats are in YYYY-MM-DD format.`;
        
        schema = {
          type: Type.OBJECT,
          properties: {
            id: { type: Type.STRING, description: "The invoice number or invoice ID (e.g. INV-1234)" },
            invoiceDate: { type: Type.STRING, description: "The date of the invoice in YYYY-MM-DD format" },
            supplierName: { type: Type.STRING, description: "The name of the vendor/supplier" },
            poNumber: { type: Type.STRING, description: "The purchase order number referenced, if any (e.g. PO-2026-1001)" },
            itemDescription: { type: Type.STRING, description: "Brief description of the main items billed" },
            quantityBilled: { type: Type.INTEGER, description: "The total quantity billed" },
            unitPrice: { type: Type.NUMBER, description: "The unit price of the items" },
            totalAmount: { type: Type.NUMBER, description: "The total amount of the invoice" }
          },
          required: ["id", "invoiceDate", "supplierName", "itemDescription", "quantityBilled", "unitPrice", "totalAmount"]
        };
      } else if (docType === "po") {
        promptText = `Extract the Purchase Order details. Format date as YYYY-MM-DD.`;
        schema = {
          type: Type.OBJECT,
          properties: {
            id: { type: Type.STRING, description: "The Purchase Order number (e.g. PO-2026-1001)" },
            supplierName: { type: Type.STRING, description: "Name of the supplier" },
            purchaseDate: { type: Type.STRING, description: "The date of purchase in YYYY-MM-DD format" },
            itemDescription: { type: Type.STRING, description: "The description of items ordered" },
            quantityOrdered: { type: Type.INTEGER, description: "The ordered quantity" },
            unitPrice: { type: Type.NUMBER, description: "The agreed unit price" },
            totalAmount: { type: Type.NUMBER, description: "The total approved amount" }
          },
          required: ["id", "supplierName", "purchaseDate", "itemDescription", "quantityOrdered", "unitPrice", "totalAmount"]
        };
      } else {
        promptText = `Extract the Goods Received Note (GRN) details. Format date as YYYY-MM-DD.`;
        schema = {
          type: Type.OBJECT,
          properties: {
            id: { type: Type.STRING, description: "The GRN number (e.g. GRN-2026-5001)" },
            poNumber: { type: Type.STRING, description: "The referenced purchase order number" },
            dateReceived: { type: Type.STRING, description: "The date received in YYYY-MM-DD format" },
            itemDescription: { type: Type.STRING, description: "Description of the items received" },
            quantityReceived: { type: Type.INTEGER, description: "Quantity received/delivered" }
          },
          required: ["id", "poNumber", "dateReceived", "itemDescription", "quantityReceived"]
        };
      }

      try {
        const response = await ai.models.generateContent({
          model: "gemini-2.5-flash",
          contents: { parts: [mediaPart, { text: promptText }] },
          config: {
            responseMimeType: "application/json",
            responseSchema: schema,
            temperature: 0.1
          }
        });

        const resultText = response.text || "{}";
        const extractedData = JSON.parse(resultText);

        // Auto-save parsed PO / GRN documents into database if docType is po or grn
        if (docType === "po") {
          db.purchaseOrders.unshift(extractedData as PurchaseOrder);
        } else if (docType === "grn") {
          db.goodsReceivedNotes.unshift(extractedData as GoodsReceivedNote);
        }

        return res.json({ success: true, method: "gemini", data: extractedData, db });
      } catch (geminiErr: any) {
        console.warn("Gemini model extraction failed, falling back to smart simulation:", geminiErr?.message || geminiErr);
      }
    }

    // --- DEMO / SIMULATION FALLBACK ---
    console.log(`Fallback simulation extraction for: ${fileName} (${docType})`);
      
      let mockData: any = {};
      const fileClean = (fileName || "").replace(/\.[^/.]+$/, "");

      // If user has Purchase Orders in database, try matching filename to an existing PO in database
      const existingPo = db.purchaseOrders[0] || null;

      if (docType === "invoice") {
        mockData = {
          id: `INV-${Math.floor(10000 + Math.random() * 90000)}`,
          invoiceDate: new Date().toISOString().split("T")[0],
          supplierName: (existingPo && existingPo.supplierName && isNaN(Number(existingPo.supplierName.trim()))) ? existingPo.supplierName : "Lian Seng Hardware Supplies Pte Ltd",
          poNumber: existingPo ? existingPo.id : "PO-2026-1001",
          itemDescription: existingPo ? existingPo.itemDescription : (fileClean || "Industrial Hardware Supplies"),
          quantityBilled: existingPo ? existingPo.quantityOrdered : 100,
          unitPrice: existingPo ? existingPo.unitPrice : 15.00,
          totalAmount: existingPo ? existingPo.totalAmount : 1500.00
        };
      } else if (docType === "po") {
        mockData = {
          id: `PO-${new Date().getFullYear()}-${Math.floor(1000 + Math.random() * 9000)}`,
          supplierName: "Lian Seng Hardware Supplies Pte Ltd",
          purchaseDate: new Date().toISOString().split("T")[0],
          itemDescription: fileClean || "Hardware Materials",
          quantityOrdered: 100,
          unitPrice: 20.00,
          totalAmount: 2000.00
        };
        db.purchaseOrders.unshift(mockData as PurchaseOrder);
      } else {
        mockData = {
          id: `GRN-${new Date().getFullYear()}-${Math.floor(5000 + Math.random() * 9000)}`,
          poNumber: existingPo ? existingPo.id : "PO-2026-1001",
          dateReceived: new Date().toISOString().split("T")[0],
          itemDescription: existingPo ? existingPo.itemDescription : "Received Hardware Items",
          quantityReceived: existingPo ? existingPo.quantityOrdered : 100
        };
        db.goodsReceivedNotes.unshift(mockData as GoodsReceivedNote);
      }

      // Simulate network delay
      await new Promise(r => setTimeout(r, 800));
      return res.json({ success: true, method: "simulation", data: mockData, db });
  } catch (error: any) {
    console.error("Extraction error:", error);
    res.status(500).json({ error: error.message || "Failed to parse document" });
  }
});

// 6. Perform the 3-Way Match comparison
app.post("/api/match", (req, res) => {
  try {
    const { invoice, manualPo, manualGrn } = req.body;
    if (!invoice) {
      return res.status(400).json({ error: "Invoice data is required" });
    }

    const report = runThreeWayMatch(invoice, db, manualPo, manualGrn);
    
    // Look up referenced documents to return to client for side-by-side view with normalized matching
    const normInvPo = invoice.poNumber ? invoice.poNumber.replace(/[^A-Za-z0-9]/g, '').toUpperCase() : '';
    const referencedPo = manualPo || (normInvPo 
      ? db.purchaseOrders.find(p => p.id.replace(/[^A-Za-z0-9]/g, '').toUpperCase() === normInvPo) || null 
      : null);
    
    const normPoId = referencedPo ? referencedPo.id.replace(/[^A-Za-z0-9]/g, '').toUpperCase() : normInvPo;
    const referencedGrn = manualGrn || (normPoId 
      ? db.goodsReceivedNotes.find(g => g.poNumber.replace(/[^A-Za-z0-9]/g, '').toUpperCase() === normPoId) || null 
      : null);

    res.json({
      success: true,
      report,
      po: referencedPo,
      grn: referencedGrn,
      invoice
    });
  } catch (error: any) {
    console.error("Match error:", error);
    res.status(500).json({ error: error.message || "Failed to perform matching" });
  }
});

// ==========================================
// VITE OR STATIC SERVING MIDDLEWARE
// ==========================================

async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    console.log("Vite dev server integrating in middlewareMode...");
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    console.log("Serving static production files from dist/...");
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server fully running on http://localhost:${PORT}`);
  });
}

startServer();
