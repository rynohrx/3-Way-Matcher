import React, { useState, useEffect } from 'react';
import {
  X,
  FileSpreadsheet,
  CloudDownload,
  CloudUpload,
  Search,
  CheckCircle,
  AlertCircle,
  Loader2,
  ExternalLink,
  Table,
  PlusCircle,
  ArrowRight,
  ShieldCheck,
  LogOut,
  FolderOpen
} from 'lucide-react';
import { User } from 'firebase/auth';
import {
  googleSignIn,
  getAccessToken,
  initAuth,
  logout
} from '../lib/googleAuth';
import {
  DriveFile,
  SpreadsheetInfo,
  listDriveSpreadsheets,
  getSpreadsheetDetails,
  getSpreadsheetValues,
  extractSpreadsheetId,
  createGoogleSheet,
  appendGoogleSheetValues
} from '../lib/googleSheets';

interface GoogleSheetsModalProps {
  isOpen: boolean;
  onClose: () => void;
  onImportInvoices: (rows2D: any[][], fileName: string) => void;
  exportData?: {
    auditedInvoices: any[];
    actionTimestamp: string;
  };
  initialTab?: 'import' | 'export';
}

// Calculate Payment Due Date (Net 30 days from invoice date if custom due date is not specified)
function calculateDueDate(invoiceDateStr?: string, customDueDate?: string): string {
  if (customDueDate && customDueDate.trim()) return customDueDate.trim();
  if (!invoiceDateStr) return 'N/A';
  try {
    const d = new Date(invoiceDateStr);
    if (isNaN(d.getTime())) return invoiceDateStr;
    d.setDate(d.getDate() + 30); // Net 30 default payment terms
    return d.toISOString().split('T')[0];
  } catch {
    return invoiceDateStr;
  }
}

export const GoogleSheetsModal: React.FC<GoogleSheetsModalProps> = ({
  isOpen,
  onClose,
  onImportInvoices,
  exportData,
  initialTab = 'export'
}) => {
  const [activeTab, setActiveTab] = useState<'import' | 'export'>(initialTab);
  
  // Auth state
  const [user, setUser] = useState<User | null>(null);
  const [token, setToken] = useState<string | null>(null);
  const [isSigningIn, setIsSigningIn] = useState(false);
  
  // Drive & Sheets state
  const [driveFiles, setDriveFiles] = useState<DriveFile[]>([]);
  const [loadingDrive, setLoadingDrive] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  
  // Selected Sheet state
  const [sheetUrlInput, setSheetUrlInput] = useState('');
  const [selectedSpreadsheet, setSelectedSpreadsheet] = useState<SpreadsheetInfo | null>(null);
  const [selectedTabTitle, setSelectedTabTitle] = useState<string>('');
  const [previewRows, setPreviewRows] = useState<any[][]>([]);
  const [loadingSheetDetails, setLoadingSheetDetails] = useState(false);
  const [loadingValues, setLoadingValues] = useState(false);

  // Export state
  const [exportFilter, setExportFilter] = useState<'approved' | 'all'>('approved');
  const [exportMode, setExportMode] = useState<'new' | 'existing'>('new');
  const [newSheetTitle, setNewSheetTitle] = useState('Madam Lim AP Approved Invoices - ' + new Date().toISOString().slice(0, 10));
  const [exporting, setExporting] = useState(false);
  const [exportSuccessUrl, setExportSuccessUrl] = useState<string | null>(null);
  const [showConfirmExport, setShowConfirmExport] = useState(false);

  const [statusMsg, setStatusMsg] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);

  useEffect(() => {
    if (isOpen && initialTab) {
      setActiveTab(initialTab);
    }
  }, [isOpen, initialTab]);

  // Initialize Auth listener on mount
  useEffect(() => {
    const unsubscribe = initAuth(
      (authUser, authToken) => {
        setUser(authUser);
        setToken(authToken);
        fetchDriveFiles(authToken);
      },
      () => {
        setUser(null);
        setToken(null);
      }
    );
    return () => unsubscribe();
  }, []);

  const fetchDriveFiles = async (authToken: string) => {
    setLoadingDrive(true);
    setStatusMsg(null);
    try {
      const files = await listDriveSpreadsheets(authToken);
      setDriveFiles(files);
    } catch (err: any) {
      console.error('Error fetching Drive spreadsheets:', err);
      setStatusMsg({ type: 'error', text: err.message || 'Could not load Google Drive spreadsheets.' });
    } finally {
      setLoadingDrive(false);
    }
  };

  const handleSignIn = async () => {
    setIsSigningIn(true);
    setStatusMsg(null);
    try {
      const res = await googleSignIn();
      if (res) {
        setUser(res.user);
        setToken(res.accessToken);
        setStatusMsg({ type: 'success', text: `Signed in as ${res.user.displayName || res.user.email}` });
        fetchDriveFiles(res.accessToken);
      }
    } catch (err: any) {
      console.error('Sign In error:', err);
      setStatusMsg({ type: 'error', text: err.message || 'Sign in failed or popup was closed.' });
    } finally {
      setIsSigningIn(false);
    }
  };

  const handleSignOut = async () => {
    await logout();
    setUser(null);
    setToken(null);
    setSelectedSpreadsheet(null);
    setPreviewRows([]);
    setDriveFiles([]);
    setStatusMsg({ type: 'info', text: 'Logged out of Google account.' });
  };

  const handleLoadSpreadsheet = async (spreadsheetIdOrUrl: string) => {
    const id = extractSpreadsheetId(spreadsheetIdOrUrl);
    if (!id) {
      setStatusMsg({ type: 'error', text: 'Invalid Google Spreadsheet URL or ID.' });
      return;
    }

    const currentToken = token || getAccessToken();
    if (!currentToken) {
      setStatusMsg({ type: 'error', text: 'Please sign in with Google first.' });
      return;
    }

    setLoadingSheetDetails(true);
    setStatusMsg(null);
    setPreviewRows([]);

    try {
      const info = await getSpreadsheetDetails(currentToken, id);
      setSelectedSpreadsheet(info);
      if (info.sheets.length > 0) {
        // Look for tab named "Approved Invoices" (case insensitive)
        const approvedTab = info.sheets.find(s =>
          s.title.trim().toLowerCase().includes('approved')
        );

        if (approvedTab) {
          setSelectedTabTitle(approvedTab.title);
          fetchTabValues(currentToken, id, approvedTab.title);
        } else {
          // If no 'Approved Invoices' tab found, select the first tab but notify user of restriction
          const firstTab = info.sheets[0].title;
          setSelectedTabTitle(firstTab);
          fetchTabValues(currentToken, id, firstTab);
          setStatusMsg({
            type: 'error',
            text: `Notice: No "Approved Invoices" tab found in "${info.title}". Imports are restricted to reading ONLY from an "Approved Invoices" sheet tab.`
          });
        }
      }
    } catch (err: any) {
      console.error('Error getting sheet details:', err);
      setStatusMsg({ type: 'error', text: err.message || 'Unable to open Google Sheet. Check permissions.' });
    } finally {
      setLoadingSheetDetails(false);
    }
  };

  const fetchTabValues = async (authToken: string, spreadsheetId: string, tabName: string) => {
    setLoadingValues(true);
    const isApprovedTab = tabName.trim().toLowerCase().includes('approved');
    try {
      const rows = await getSpreadsheetValues(authToken, spreadsheetId, tabName);
      setPreviewRows(rows);
      if (!isApprovedTab) {
        setStatusMsg({
          type: 'error',
          text: `Restricted Tab: Selected tab "${tabName}" is not an "Approved Invoices" tab. Data can ONLY be imported from an "Approved Invoices" tab.`
        });
      } else if (rows.length === 0) {
        setStatusMsg({ type: 'info', text: 'Selected "Approved Invoices" tab is empty.' });
      } else {
        setStatusMsg({ type: 'success', text: `Loaded ${rows.length} invoice rows from "Approved Invoices" tab (${tabName}).` });
      }
    } catch (err: any) {
      console.error('Error fetching sheet rows:', err);
      setStatusMsg({ type: 'error', text: err.message || 'Failed to read rows from sheet tab.' });
    } finally {
      setLoadingValues(false);
    }
  };

  const handleTabChange = (newTab: string) => {
    setSelectedTabTitle(newTab);
    const currentToken = token || getAccessToken();
    if (selectedSpreadsheet && currentToken) {
      fetchTabValues(currentToken, selectedSpreadsheet.id, newTab);
    }
  };

  const handleExecuteImport = () => {
    const isApprovedTab = selectedTabTitle.trim().toLowerCase().includes('approved');
    if (!isApprovedTab) {
      setStatusMsg({
        type: 'error',
        text: 'Import Restricted: Invoice data can ONLY be read from the "Approved Invoices" sheet/tab. Please select an "Approved Invoices" sheet tab.'
      });
      return;
    }

    if (previewRows.length === 0) {
      setStatusMsg({ type: 'error', text: 'No invoice rows available to import in "Approved Invoices" tab.' });
      return;
    }
    const fileName = selectedSpreadsheet
      ? `${selectedSpreadsheet.title} (${selectedTabTitle})`
      : 'Google Sheet Import';
    onImportInvoices(previewRows, fileName);
    onClose();
  };

  const handleConfirmExport = async () => {
    const currentToken = token || getAccessToken();
    if (!currentToken) {
      setStatusMsg({ type: 'error', text: 'Please sign in with Google first.' });
      return;
    }

    if (!exportData || !exportData.auditedInvoices || exportData.auditedInvoices.length === 0) {
      setStatusMsg({ type: 'error', text: 'No audit records available to export.' });
      return;
    }

    setExporting(true);
    setStatusMsg(null);
    setExportSuccessUrl(null);

    // STRICT REQUIREMENT: Only export invoices that have been explicitly APPROVED
    const approvedInvoices = (exportData?.auditedInvoices || []).filter(
      (i) => i.actionStatus === 'APPROVED'
    );

    if (approvedInvoices.length === 0) {
      setStatusMsg({
        type: 'error',
        text: 'Export Restricted: You must approve at least one invoice in Madam Lim AP Auditor before exporting to Google Sheets.'
      });
      return;
    }

    setExporting(true);
    setStatusMsg(null);
    setExportSuccessUrl(null);

    const targetInvoices = approvedInvoices;

    const headers = [
      'Invoice ID',
      'Supplier Name',
      'Invoice Date',
      'Payment Due Date',
      'PO Reference',
      'Item Description',
      'Qty Billed',
      'Unit Price ($)',
      'Total Billed ($)',
      '3-Way Match Result',
      'Audit Recommendation',
      'Approval Status',
      'Price Variance',
      'Qty Variance',
      'Export & Audit Time'
    ];

    const rows = targetInvoices.map((inv) => {
      const dueDate = calculateDueDate(inv.invoice.invoiceDate, inv.invoice.dueDate);
      const statusText =
        inv.actionStatus === 'APPROVED'
          ? 'APPROVED & LOGGED'
          : inv.actionStatus === 'REJECTED'
          ? 'REJECTED'
          : 'PENDING / AUDITED';

      return [
        inv.invoice.id,
        inv.invoice.supplierName,
        inv.invoice.invoiceDate,
        dueDate,
        inv.invoice.poNumber || 'N/A',
        inv.invoice.itemDescription,
        inv.invoice.quantityBilled,
        inv.invoice.unitPrice,
        inv.invoice.totalAmount,
        inv.matchReport.overallResult,
        inv.recommendation,
        statusText,
        inv.matchReport.priceVariancePercent ? `${inv.matchReport.priceVariancePercent}%` : '0%',
        inv.matchReport.qtyVariancePercent ? `${inv.matchReport.qtyVariancePercent}%` : '0%',
        exportData.actionTimestamp
      ];
    });

    try {
      if (exportMode === 'new') {
        const result = await createGoogleSheet(currentToken, newSheetTitle, headers, rows);
        setExportSuccessUrl(result.spreadsheetUrl);
        setStatusMsg({
          type: 'success',
          text: `Successfully created Google Sheet "${newSheetTitle}" with ${rows.length} audit records!`
        });
      } else {
        if (!selectedSpreadsheet) {
          throw new Error('Please select an existing Google Sheet to append to.');
        }
        await appendGoogleSheetValues(
          currentToken,
          selectedSpreadsheet.id,
          selectedTabTitle || 'Sheet1',
          rows
        );
        const url = `https://docs.google.com/spreadsheets/d/${selectedSpreadsheet.id}/edit`;
        setExportSuccessUrl(url);
        setStatusMsg({
          type: 'success',
          text: `Successfully appended ${rows.length} audit records to "${selectedSpreadsheet.title}"!`
        });
      }
      setShowConfirmExport(false);
    } catch (err: any) {
      console.error('Export error:', err);
      setStatusMsg({ type: 'error', text: err.message || 'Failed to export audit report to Google Sheets.' });
    } finally {
      setExporting(false);
    }
  };

  if (!isOpen) return null;

  const filteredDriveFiles = driveFiles.filter((f) =>
    f.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 backdrop-blur-xs p-4">
      <div className="bg-slate-900 border border-slate-700 rounded-xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden text-slate-100">
        
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 bg-slate-800/80 border-b border-slate-700">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-emerald-900/50 border border-emerald-600/40 rounded-lg text-emerald-400">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-white flex items-center gap-2">
                Google Sheets Integration
                <span className="text-xs bg-emerald-900/80 text-emerald-300 border border-emerald-700 px-2 py-0.5 rounded-full font-medium">
                  Live Sync
                </span>
              </h2>
              <p className="text-xs text-slate-400">
                Import invoice datasets directly from Google Drive or export Madam Lim audit reports
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-700/60 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* User Auth Banner */}
        <div className="bg-slate-950/70 px-6 py-3 border-b border-slate-800 flex items-center justify-between flex-wrap gap-3">
          {user ? (
            <div className="flex items-center gap-3">
              {user.photoURL ? (
                <img src={user.photoURL} alt="Avatar" className="w-7 h-7 rounded-full border border-emerald-500/50" />
              ) : (
                <div className="w-7 h-7 rounded-full bg-emerald-700 text-white font-bold flex items-center justify-center text-xs">
                  {user.email?.charAt(0).toUpperCase() || 'G'}
                </div>
              )}
              <div>
                <p className="text-xs font-semibold text-emerald-300 flex items-center gap-1">
                  <ShieldCheck className="w-3.5 h-3.5" /> Connected to Google
                </p>
                <p className="text-[11px] text-slate-400">{user.email}</p>
              </div>
            </div>
          ) : (
            <div className="flex items-center gap-2">
              <span className="text-xs text-amber-300 flex items-center gap-1.5">
                <AlertCircle className="w-4 h-4" /> Sign in to access your Google Drive & Sheets
              </span>
            </div>
          )}

          <div>
            {user ? (
              <button
                onClick={handleSignOut}
                className="text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-1.5 rounded-md border border-slate-600 flex items-center gap-1.5 transition-colors"
              >
                <LogOut className="w-3.5 h-3.5" /> Sign Out
              </button>
            ) : (
              <button
                onClick={handleSignIn}
                disabled={isSigningIn}
                className="gsi-material-button bg-white hover:bg-slate-100 text-slate-900 px-4 py-2 rounded-lg text-xs font-semibold shadow-md flex items-center gap-2 transition-all disabled:opacity-50"
              >
                {isSigningIn ? (
                  <Loader2 className="w-4 h-4 animate-spin text-slate-700" />
                ) : (
                  <svg className="w-4 h-4" viewBox="0 0 48 48">
                    <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z" />
                    <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z" />
                    <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z" />
                    <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z" />
                  </svg>
                )}
                <span>Sign in with Google</span>
              </button>
            )}
          </div>
        </div>

        {/* Action Tabs */}
        <div className="flex border-b border-slate-800 bg-slate-900/50 px-6">
          <button
            onClick={() => setActiveTab('import')}
            className={`py-3 px-4 font-semibold text-xs border-b-2 transition-colors flex items-center gap-2 ${
              activeTab === 'import'
                ? 'border-emerald-500 text-emerald-400 bg-emerald-950/20'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <CloudDownload className="w-4 h-4" /> Import Invoices from Google Sheet
          </button>
          <button
            onClick={() => setActiveTab('export')}
            className={`py-3 px-4 font-semibold text-xs border-b-2 transition-colors flex items-center gap-2 ${
              activeTab === 'export'
                ? 'border-emerald-500 text-emerald-400 bg-emerald-950/20'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <CloudUpload className="w-4 h-4" /> Export Audit Reports to Google Sheets
          </button>
        </div>

        {/* Notifications & Status Banner */}
        {statusMsg && (
          <div
            className={`px-6 py-2.5 text-xs font-medium flex items-center gap-2 border-b ${
              statusMsg.type === 'success'
                ? 'bg-emerald-950/80 text-emerald-300 border-emerald-800'
                : statusMsg.type === 'error'
                ? 'bg-red-950/80 text-red-300 border-red-800'
                : 'bg-blue-950/80 text-blue-300 border-blue-800'
            }`}
          >
            {statusMsg.type === 'success' ? (
              <CheckCircle className="w-4 h-4 text-emerald-400 shrink-0" />
            ) : statusMsg.type === 'error' ? (
              <AlertCircle className="w-4 h-4 text-red-400 shrink-0" />
            ) : (
              <Loader2 className="w-4 h-4 text-blue-400 animate-spin shrink-0" />
            )}
            <span className="flex-1">{statusMsg.text}</span>
            {exportSuccessUrl && (
              <a
                href={exportSuccessUrl}
                target="_blank"
                rel="noreferrer"
                className="bg-emerald-600 hover:bg-emerald-500 text-white px-2.5 py-1 rounded text-xs font-bold flex items-center gap-1 transition-colors"
              >
                Open Google Sheet <ExternalLink className="w-3 h-3" />
              </a>
            )}
          </div>
        )}

        {/* Tab Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          {!user ? (
            <div className="text-center py-12 space-y-4 max-w-md mx-auto">
              <div className="w-16 h-16 rounded-2xl bg-slate-800 border border-slate-700 flex items-center justify-center mx-auto text-emerald-400">
                <FileSpreadsheet className="w-8 h-8" />
              </div>
              <h3 className="text-base font-bold text-white">Google Workspace Auth Required</h3>
              <p className="text-xs text-slate-400 leading-relaxed">
                Connect your Google account to browse spreadsheets directly from your Google Drive, import invoice rows into Madam Lim's 3-way auditor, or export verified audit logs.
              </p>
              <button
                onClick={handleSignIn}
                disabled={isSigningIn}
                className="bg-emerald-600 hover:bg-emerald-500 text-white px-6 py-2.5 rounded-lg text-xs font-bold shadow-lg flex items-center gap-2 mx-auto transition-colors disabled:opacity-50"
              >
                {isSigningIn ? <Loader2 className="w-4 h-4 animate-spin" /> : <ShieldCheck className="w-4 h-4" />}
                Sign in with Google Account
              </button>
            </div>
          ) : activeTab === 'import' ? (
            <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
              
              {/* Left Column: Drive File Picker or URL Input */}
              <div className="md:col-span-5 space-y-4">
                
                {/* Direct Sheet URL / ID Input */}
                <div className="bg-slate-800/60 border border-slate-700 rounded-xl p-4 space-y-3">
                  <label className="text-xs font-bold text-slate-200 block">
                    Paste Google Sheet URL or ID
                  </label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      value={sheetUrlInput}
                      onChange={(e) => setSheetUrlInput(e.target.value)}
                      placeholder="https://docs.google.com/spreadsheets/d/..."
                      className="flex-1 bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-xs text-white placeholder-slate-500 focus:outline-hidden focus:border-emerald-500"
                    />
                    <button
                      onClick={() => handleLoadSpreadsheet(sheetUrlInput)}
                      disabled={loadingSheetDetails || !sheetUrlInput.trim()}
                      className="bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-2 rounded-lg text-xs font-semibold disabled:opacity-50 flex items-center gap-1 transition-colors"
                    >
                      {loadingSheetDetails ? <Loader2 className="w-4 h-4 animate-spin" /> : 'Load'}
                    </button>
                  </div>
                </div>

                {/* Google Drive Recent Spreadsheets */}
                <div className="bg-slate-800/60 border border-slate-700 rounded-xl p-4 space-y-3 flex flex-col h-[320px]">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-bold text-slate-200 flex items-center gap-1.5">
                      <FolderOpen className="w-4 h-4 text-emerald-400" /> Spreadsheets in Google Drive
                    </label>
                    <button
                      onClick={() => token && fetchDriveFiles(token)}
                      className="text-[11px] text-emerald-400 hover:underline"
                    >
                      Refresh
                    </button>
                  </div>

                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      placeholder="Search Drive files..."
                      className="w-full bg-slate-900 border border-slate-700 rounded-lg pl-8 pr-3 py-1.5 text-xs text-white placeholder-slate-500 focus:outline-hidden focus:border-emerald-500"
                    />
                  </div>

                  <div className="flex-1 overflow-y-auto space-y-1.5 pr-1">
                    {loadingDrive ? (
                      <div className="text-center py-8 text-xs text-slate-400 flex items-center justify-center gap-2">
                        <Loader2 className="w-4 h-4 animate-spin text-emerald-400" /> Loading files from Drive...
                      </div>
                    ) : filteredDriveFiles.length === 0 ? (
                      <div className="text-center py-8 text-xs text-slate-500">
                        No Google Spreadsheets found in your Drive.
                      </div>
                    ) : (
                      filteredDriveFiles.map((file) => (
                        <button
                          key={file.id}
                          onClick={() => {
                            setSheetUrlInput(file.id);
                            handleLoadSpreadsheet(file.id);
                          }}
                          className={`w-full text-left p-2.5 rounded-lg border transition-all flex items-start gap-2.5 ${
                            selectedSpreadsheet?.id === file.id
                              ? 'bg-emerald-950/70 border-emerald-500/80 text-white'
                              : 'bg-slate-900/60 border-slate-700/60 hover:bg-slate-700/40 text-slate-300'
                          }`}
                        >
                          <FileSpreadsheet className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
                          <div className="flex-1 min-w-0">
                            <p className="text-xs font-semibold truncate text-slate-200">{file.name}</p>
                            <p className="text-[10px] text-slate-500 truncate">
                              Modified: {file.modifiedTime ? new Date(file.modifiedTime).toLocaleDateString() : 'N/A'}
                            </p>
                          </div>
                        </button>
                      ))
                    )}
                  </div>
                </div>

              </div>

              {/* Right Column: Sheet Tabs & Preview */}
              <div className="md:col-span-7 space-y-4 flex flex-col">
                {selectedSpreadsheet ? (
                  <div className="bg-slate-800/60 border border-slate-700 rounded-xl p-4 space-y-4 flex-1 flex flex-col">
                    <div className="flex items-center justify-between border-b border-slate-700 pb-3">
                      <div>
                        <h4 className="text-sm font-bold text-white flex items-center gap-2">
                          <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                          {selectedSpreadsheet.title}
                        </h4>
                        <p className="text-[11px] text-slate-400">
                          Spreadsheet ID: <code className="text-slate-300">{selectedSpreadsheet.id.slice(0, 16)}...</code>
                        </p>
                      </div>

                      {/* Sheet Tab Selector */}
                      <div className="flex items-center gap-2">
                        <span className="text-xs text-slate-400 font-medium">Sheet Tab:</span>
                        <select
                          value={selectedTabTitle}
                          onChange={(e) => handleTabChange(e.target.value)}
                          className={`border rounded-lg px-2.5 py-1 text-xs font-semibold focus:outline-hidden ${
                            selectedTabTitle.trim().toLowerCase().includes('approved')
                              ? 'bg-slate-900 border-emerald-500 text-emerald-300'
                              : 'bg-slate-900 border-amber-500/80 text-amber-300'
                          }`}
                        >
                          {selectedSpreadsheet.sheets.map((tab) => (
                            <option key={tab.sheetId} value={tab.title}>
                              {tab.title} {tab.title.trim().toLowerCase().includes('approved') ? '✓ (Approved)' : ''}
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>

                    {/* Tab Enforcement Banner */}
                    {!selectedTabTitle.trim().toLowerCase().includes('approved') ? (
                      <div className="bg-amber-950/80 border border-amber-600/80 p-2.5 rounded-lg text-xs text-amber-200 flex items-center gap-2">
                        <AlertCircle className="w-4 h-4 text-amber-400 shrink-0" />
                        <span>
                          <strong>Import Restricted:</strong> Invoice data can ONLY be read from an <strong>"Approved Invoices"</strong> sheet tab. Selected tab "<em>{selectedTabTitle}</em>" is restricted.
                        </span>
                      </div>
                    ) : (
                      <div className="bg-emerald-950/40 border border-emerald-800/60 p-2 rounded-lg text-[11px] text-emerald-300 flex items-center gap-2">
                        <CheckCircle className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
                        <span>
                          Reading verified invoice dataset from <strong>"Approved Invoices"</strong> sheet tab ({selectedTabTitle}).
                        </span>
                      </div>
                    )}

                    {/* Preview Table */}
                    <div className="flex-1 border border-slate-700 rounded-lg overflow-hidden bg-slate-950 flex flex-col min-h-[260px]">
                      <div className="bg-slate-900 px-3 py-2 border-b border-slate-800 text-xs font-semibold text-slate-300 flex items-center justify-between">
                        <span className="flex items-center gap-1.5">
                          <Table className="w-3.5 h-3.5 text-emerald-400" /> Preview ({previewRows.length} rows)
                        </span>
                        {loadingValues && <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-400" />}
                      </div>

                      <div className="flex-1 overflow-auto max-h-[280px] p-2">
                        {loadingValues ? (
                          <div className="text-center py-12 text-xs text-slate-400 flex items-center justify-center gap-2">
                            <Loader2 className="w-4 h-4 animate-spin text-emerald-400" /> Fetching rows from Google Sheet...
                          </div>
                        ) : previewRows.length === 0 ? (
                          <div className="text-center py-12 text-xs text-slate-500">
                            No row data found in this tab.
                          </div>
                        ) : (
                          <table className="w-full text-left text-[11px] border-collapse">
                            <tbody>
                              {previewRows.slice(0, 15).map((row, rIdx) => (
                                <tr
                                  key={rIdx}
                                  className={`border-b border-slate-800/60 ${
                                    rIdx === 0 ? 'bg-slate-800/80 font-bold text-emerald-300' : 'hover:bg-slate-800/30 text-slate-300'
                                  }`}
                                >
                                  <td className="px-2 py-1 text-[10px] text-slate-500 bg-slate-900 border-r border-slate-800 font-mono w-8 text-center">
                                    {rIdx + 1}
                                  </td>
                                  {Array.isArray(row) &&
                                    row.slice(0, 8).map((cell, cIdx) => (
                                      <td key={cIdx} className="px-2.5 py-1 border-r border-slate-800/40 truncate max-w-[120px]">
                                        {String(cell ?? '')}
                                      </td>
                                    ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        )}
                      </div>
                    </div>

                    {/* Import Button */}
                    <div className="pt-2 flex justify-end">
                      <button
                        onClick={handleExecuteImport}
                        disabled={previewRows.length === 0 || !selectedTabTitle.trim().toLowerCase().includes('approved')}
                        className={`font-bold px-5 py-2.5 rounded-lg text-xs shadow-lg flex items-center gap-2 transition-all disabled:opacity-50 ${
                          selectedTabTitle.trim().toLowerCase().includes('approved')
                            ? 'bg-emerald-600 hover:bg-emerald-500 text-white'
                            : 'bg-slate-800 text-slate-500 cursor-not-allowed border border-slate-700'
                        }`}
                      >
                        <ArrowRight className="w-4 h-4" />
                        {!selectedTabTitle.trim().toLowerCase().includes('approved')
                          ? 'Select "Approved Invoices" Tab to Import'
                          : 'Import Invoices from Approved Tab'}
                      </button>
                    </div>
                  </div>
                ) : (
                  <div className="bg-slate-800/30 border border-dashed border-slate-700 rounded-xl p-8 text-center text-slate-400 space-y-3 flex flex-col items-center justify-center min-h-[350px]">
                    <FileSpreadsheet className="w-12 h-12 text-slate-600" />
                    <p className="text-xs text-slate-300 font-medium">No Google Sheet Selected</p>
                    <p className="text-[11px] text-slate-500 max-w-xs">
                      Choose a spreadsheet from your Google Drive list on the left or paste a direct Google Sheet URL.
                    </p>
                  </div>
                )}
              </div>

            </div>
          ) : (
            /* EXPORT TAB */
            <div className="max-w-2xl mx-auto space-y-6">
              <div className="bg-slate-800/60 border border-slate-700 rounded-xl p-5 space-y-4">
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <CloudUpload className="w-4 h-4 text-emerald-400" /> Export Audit Results to Google Sheets
                </h3>
                <p className="text-xs text-slate-400 leading-relaxed">
                  Export verified 3-Way Match audit records (invoices, PO matching status, price/qty variances, Madam Lim recommendations) directly to Google Sheets for reporting and archive.
                </p>

                {/* Export Record Scope Selection */}
                <div className="space-y-2 pt-1">
                  <label className="text-xs font-semibold text-slate-300">Select Audit Records to Export:</label>
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      onClick={() => setExportFilter('approved')}
                      className={`p-2.5 rounded-lg border text-xs font-semibold flex items-center justify-between transition-all ${
                        exportFilter === 'approved'
                          ? 'bg-emerald-950/80 border-emerald-500 text-emerald-300 shadow-xs'
                          : 'bg-slate-900 border-slate-700 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      <span className="flex items-center gap-1.5">
                        <CheckCircle className="w-3.5 h-3.5 text-emerald-400" />
                        Approved Invoices Only
                      </span>
                      <span className="bg-emerald-900/60 text-emerald-300 px-1.5 py-0.5 rounded text-[10px] font-mono font-bold">
                        {exportData?.auditedInvoices?.filter(i => i.actionStatus === 'APPROVED' || i.recommendation === 'ACCEPT').length || 0}
                      </span>
                    </button>

                    <button
                      onClick={() => setExportFilter('all')}
                      className={`p-2.5 rounded-lg border text-xs font-semibold flex items-center justify-between transition-all ${
                        exportFilter === 'all'
                          ? 'bg-emerald-950/80 border-emerald-500 text-emerald-300 shadow-xs'
                          : 'bg-slate-900 border-slate-700 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      <span className="flex items-center gap-1.5">
                        <Table className="w-3.5 h-3.5 text-emerald-400" />
                        All Audited Records
                      </span>
                      <span className="bg-slate-800 text-slate-300 px-1.5 py-0.5 rounded text-[10px] font-mono font-bold">
                        {exportData?.auditedInvoices?.length || 0}
                      </span>
                    </button>
                  </div>
                </div>

                {/* Export Destination Mode */}
                <div className="grid grid-cols-2 gap-3 pt-2">
                  <button
                    onClick={() => setExportMode('new')}
                    className={`p-3 rounded-lg border text-left transition-all ${
                      exportMode === 'new'
                        ? 'bg-emerald-950/70 border-emerald-500 text-white'
                        : 'bg-slate-900 border-slate-700 text-slate-400 hover:text-slate-200'
                    }`}
                  >
                    <div className="flex items-center gap-2 font-bold text-xs">
                      <PlusCircle className="w-4 h-4 text-emerald-400" /> Create New Google Sheet
                    </div>
                    <p className="text-[10px] text-slate-400 mt-1">Generates a fresh standalone spreadsheet in your Drive.</p>
                  </button>

                  <button
                    onClick={() => setExportMode('existing')}
                    className={`p-3 rounded-lg border text-left transition-all ${
                      exportMode === 'existing'
                        ? 'bg-emerald-950/70 border-emerald-500 text-white'
                        : 'bg-slate-900 border-slate-700 text-slate-400 hover:text-slate-200'
                    }`}
                  >
                    <div className="flex items-center gap-2 font-bold text-xs">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-400" /> Append to Existing Sheet
                    </div>
                    <p className="text-[10px] text-slate-400 mt-1">Appends audit rows to a spreadsheet in your Drive.</p>
                  </button>
                </div>

                {exportMode === 'new' ? (
                  <div className="space-y-2 pt-2">
                    <label className="text-xs font-semibold text-slate-300">Spreadsheet Title</label>
                    <input
                      type="text"
                      value={newSheetTitle}
                      onChange={(e) => setNewSheetTitle(e.target.value)}
                      className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-xs text-white focus:outline-hidden focus:border-emerald-500"
                    />
                  </div>
                ) : (
                  <div className="space-y-2 pt-2">
                    <label className="text-xs font-semibold text-slate-300">Select Existing Spreadsheet</label>
                    <select
                      value={selectedSpreadsheet?.id || ''}
                      onChange={(e) => handleLoadSpreadsheet(e.target.value)}
                      className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-xs text-emerald-300 font-semibold focus:outline-hidden"
                    >
                      <option value="">-- Choose Spreadsheet --</option>
                      {driveFiles.map((f) => (
                        <option key={f.id} value={f.id}>
                          {f.name}
                        </option>
                      ))}
                    </select>
                  </div>
                )}

                {/* Audit Record Summary & Preview Table with Due Date */}
                {(() => {
                  const approvedCount = (exportData?.auditedInvoices || []).filter(i => i.actionStatus === 'APPROVED').length;
                  return (
                    <div className="space-y-3 pt-2">
                      {approvedCount === 0 ? (
                        <div className="bg-amber-950/60 border border-amber-600/70 p-3 rounded-lg text-xs text-amber-200 flex items-start gap-2.5">
                          <AlertCircle className="w-4 h-4 text-amber-400 shrink-0 mt-0.5" />
                          <div className="space-y-1">
                            <p className="font-bold text-amber-300">Approval Required Before Exporting</p>
                            <p className="text-[11px] text-amber-200/90 leading-relaxed">
                              You cannot export invoices to Google Sheets without approving them first. Please click <strong>"Approve & Log Payment"</strong> on invoice records in the audit dashboard to authorize them for export.
                            </p>
                          </div>
                        </div>
                      ) : (
                        <div className="bg-emerald-950/60 border border-emerald-600/70 p-3 rounded-lg text-xs text-emerald-200 flex items-center justify-between">
                          <div className="flex items-center gap-2">
                            <CheckCircle className="w-4 h-4 text-emerald-400 shrink-0" />
                            <span><strong>{approvedCount}</strong> Approved Invoice(s) Ready for Google Sheets Export</span>
                          </div>
                          <span className="bg-emerald-900/80 text-emerald-300 px-2 py-0.5 rounded text-[10px] font-mono font-bold">
                            NET 30 DATES INCLUDED
                          </span>
                        </div>
                      )}

                      {/* Table Column Preview Notice */}
                      <div className="bg-slate-950 border border-slate-800 p-2.5 rounded-lg text-[11px] text-slate-300 flex items-center justify-between">
                        <span className="flex items-center gap-2">
                          <ShieldCheck className="w-4 h-4 text-emerald-400 shrink-0" />
                          <span>Includes Payment Due Date (Net 30 days), PO Ref, and 3-Way Audit Status</span>
                        </span>
                        <span className="font-mono text-emerald-400 font-bold">{approvedCount} Record(s)</span>
                      </div>

                      {/* Preview Rows with Due Date */}
                      {approvedCount > 0 && (
                        <div className="border border-slate-700 rounded-lg overflow-hidden max-h-40 overflow-y-auto">
                          <table className="w-full text-[11px] text-left text-slate-300">
                            <thead className="bg-slate-900 text-slate-400 text-[10px] uppercase font-mono sticky top-0 border-b border-slate-700">
                              <tr>
                                <th className="p-2">Inv ID</th>
                                <th className="p-2">Supplier</th>
                                <th className="p-2">Inv Date</th>
                                <th className="p-2 text-emerald-400">Payment Due Date</th>
                                <th className="p-2 text-right">Amount</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-800 bg-slate-900/60">
                              {(exportData?.auditedInvoices || [])
                                .filter(i => i.actionStatus === 'APPROVED')
                                .slice(0, 5)
                                .map((invItem) => (
                                  <tr key={invItem.id || invItem.invoice?.id}>
                                    <td className="p-2 font-mono font-semibold text-emerald-300">{invItem.invoice?.id}</td>
                                    <td className="p-2 truncate max-w-[120px]">{invItem.invoice?.supplierName}</td>
                                    <td className="p-2 font-mono text-slate-400">{invItem.invoice?.invoiceDate}</td>
                                    <td className="p-2 font-mono font-bold text-emerald-400">
                                      {calculateDueDate(invItem.invoice?.invoiceDate, invItem.invoice?.dueDate)}
                                    </td>
                                    <td className="p-2 font-mono text-right text-slate-200">${invItem.invoice?.totalAmount?.toFixed(2)}</td>
                                  </tr>
                                ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  );
                })()}

                {/* Confirm dialog trigger */}
                <div className="pt-3">
                  {(() => {
                    const approvedCount = (exportData?.auditedInvoices || []).filter(i => i.actionStatus === 'APPROVED').length;
                    return (
                      <button
                        onClick={() => setShowConfirmExport(true)}
                        disabled={exporting || approvedCount === 0}
                        className={`w-full font-bold py-3 rounded-lg text-xs shadow-lg flex items-center justify-center gap-2 transition-all ${
                          approvedCount === 0
                            ? 'bg-slate-800 text-slate-500 cursor-not-allowed border border-slate-700'
                            : 'bg-emerald-600 hover:bg-emerald-500 text-white cursor-pointer'
                        }`}
                      >
                        <CloudUpload className="w-4 h-4" />
                        {approvedCount === 0 ? 'Approve an Invoice First to Unlock Export' : 'Prepare & Export Approved Invoices'}
                      </button>
                    );
                  })()}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* User Confirmation Modal for Destructive/Mutating Write Operations */}
        {showConfirmExport && (
          <div className="fixed inset-0 z-60 bg-black/80 backdrop-blur-xs flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-700 rounded-xl max-w-md w-full p-6 space-y-4 shadow-2xl text-slate-100">
              <div className="flex items-center gap-3 text-amber-400">
                <AlertCircle className="w-6 h-6 shrink-0" />
                <h3 className="text-base font-bold text-white">Confirm Google Sheets Write</h3>
              </div>
              <p className="text-xs text-slate-300 leading-relaxed">
                You are about to write <span className="font-bold text-emerald-400">{exportData?.auditedInvoices?.length} invoice audit records</span> into your Google Drive spreadsheet:
              </p>
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 text-xs text-slate-300 space-y-1">
                <p><span className="text-slate-500">Destination:</span> {exportMode === 'new' ? `New Sheet "${newSheetTitle}"` : `Existing Sheet "${selectedSpreadsheet?.title}"`}</p>
                <p><span className="text-slate-500">Action:</span> {exportMode === 'new' ? 'Create new file & insert rows' : 'Append rows'}</p>
              </div>
              <div className="flex gap-3 pt-2">
                <button
                  onClick={() => setShowConfirmExport(false)}
                  className="flex-1 bg-slate-800 hover:bg-slate-700 text-slate-300 py-2 rounded-lg text-xs font-semibold transition-colors"
                >
                  Cancel
                </button>
                <button
                  onClick={handleConfirmExport}
                  disabled={exporting}
                  className="flex-1 bg-emerald-600 hover:bg-emerald-500 text-white py-2 rounded-lg text-xs font-bold flex items-center justify-center gap-2 transition-colors"
                >
                  {exporting ? <Loader2 className="w-4 h-4 animate-spin" /> : <CheckCircle className="w-4 h-4" />}
                  Confirm Write
                </button>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
};
