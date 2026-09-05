import * as XLSX from 'xlsx';
import { useEffect, useState, useRef, useMemo } from "react";
import { db, auth, storage } from "../firebase";
import { collection, addDoc, query, orderBy, doc, updateDoc, onSnapshot, deleteDoc, where, getDocs, arrayUnion, writeBatch, setDoc, increment, runTransaction, serverTimestamp } from "firebase/firestore";
import { signOut } from "firebase/auth";
import { ref, uploadBytes, getDownloadURL, getMetadata } from "firebase/storage";
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import React from 'react';
/* ── SVG icon primitive ── */
const Icon = ({ d, d2, size = 16, stroke = "currentColor", fill = "none", sw = 1.8 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={fill} stroke={stroke} strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
    <path d={d} />{d2 && <path d={d2} />}
  </svg>
);
const Ic = {
  users:    (p) => <Icon {...p} d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" d2="M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" />,
  clock:    (p) => <Icon {...p} d="M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 6v6l4 2" />,
  bell:     (p) => <Icon {...p} d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9zM13.73 21a2 2 0 0 1-3.46 0" />,
  search:   (p) => <Icon {...p} d="M21 21l-4.35-4.35M17 11A6 6 0 1 1 5 11a6 6 0 0 1 12 0z" />,
  download: (p) => <Icon {...p} d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3" />,
  trash:    (p) => <Icon {...p} d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6" />,
  wrench:   (p) => <Icon {...p} d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z" />,
  logout:   (p) => <Icon {...p} d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9" />,
  moon:     (p) => <Icon {...p} d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />,
  sun:      (p) => <Icon {...p} d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42M12 5a7 7 0 1 0 0 14A7 7 0 0 0 12 5z" />,
  grid:     (p) => <Icon {...p} d="M3 3h7v7H3zM14 3h7v7h-7zM14 14h7v7h-7zM3 14h7v7H3z" />,
  history:  (p) => <Icon {...p} d="M3 3v5h5M3.05 13A9 9 0 1 0 6 5.3L3 8" />,
  wa:       (p) => <Icon {...p} d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z" />,
  chevron:  (p) => <Icon {...p} d="M9 18l6-6-6-6" />,
  plus:     (p) => <Icon {...p} d="M12 5v14M5 12h14" />,
  droplet:  (p) => <Icon {...p} d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z" />,
  filter:   (p) => <Icon {...p} d="M22 3H2l8 9.46V19l4 2v-8.54L22 3z" />,
  box:      (p) => <Icon {...p} d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16zM3.27 6.96 12 12l8.73-5.04M12 22.08V12" />,
  upload:   (p) => <Icon {...p} d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12" />,
  cash:     (p) => <Icon {...p} d="M2 7h20v10H2zM12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6zM6 7v.01M18 17v-.01" />,
  fileText: (p) => <Icon {...p} d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" d2="M14 2v6h6M9 13h6M9 17h6M9 9h1" />,
  quote:    (p) => <Icon {...p} d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" d2="M14 2v6h6M9 13h6M9 17h6M9 9h1" />,
  chart:    (p) => <Icon {...p} d="M3 3v18h18" d2="M18 17V9M13 17V5M8 17v-4" />,
  calendar: (p) => <Icon {...p} d="M8 2v4M16 2v4M3 10h18" d2="M4 4h16a1 1 0 0 1 1 1v15a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z" />,
  star:     (p) => <Icon {...p} d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01z" />,
  target:   (p) => <Icon {...p} d="M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20z" d2="M12 16a4 4 0 1 0 0-8 4 4 0 0 0 0 8z" />,
  route:    (p) => <Icon {...p} d="M6 3v12a3 3 0 0 0 3 3h9M6 3a2 2 0 1 1 0 .001M18 18a2 2 0 1 1 0 .001" />,
};

/* ════════════════════════════════════════════════════════════
   Warranty data parsing helpers (Ledger + Flat imports)
   ════════════════════════════════════════════════════════════ */

// Warranty covers this many free services per install (business rule
// confirmed with the shop owner — adjust here if the policy changes).
const WARRANTY_FREE_SERVICES = 3;

// Storage paths for the two source workbooks — shared by the periodic
// sync effect below and the backend Excel-writeback API.
const LEDGER_STORAGE_PATH = "company_data/Warranty_Ledger.xlsx";
const FLAT_STORAGE_PATH = "company_data/Warranty_Flat.xlsx";

/* ══════════════════════════════════════════════════════════
   NEW — Backend API base URL + authenticated fetch helper.

   All real-time writes into Excel-Warranty_Ledger.xlsx /
   Excel-Warranty_Flat.xlsx now go through a backend (Cloud Function /
   Cloud Run) instead of being read-modified-reuploaded directly from the
   browser. The backend serializes writes per file (so two admins saving
   at once can't clobber each other), verifies the caller's Firebase ID
   token, and keeps Firestore as the source of truth even if the Excel
   write itself fails.

   This project uses Vite, so env vars must be prefixed VITE_ and read
   via import.meta.env (process.env.* does not work in Vite client code).
   Configure this in your .env file at the project root:
     VITE_API_BASE=https://us-central1-<project>.cloudfunctions.net
   or, if you deploy to Cloud Run behind a custom domain:
     VITE_API_BASE=https://api.yourdomain.com
══════════════════════════════════════════════════════════ */
const API_BASE = import.meta.env.VITE_API_BASE || "";

// Wraps fetch with the current admin's Firebase ID token and consistent
// error handling. Never throws for a failed Excel sync specifically —
// callers decide whether to surface that as a soft warning, since the
// Firestore save (the thing the dashboard actually reads) already
// succeeded by the time this runs.
const callBackendApi = async (path, { method = "POST", body } = {}) => {
  if (!API_BASE) {
    console.warn(`[Excel Sync] REACT_APP_API_BASE is not set — skipping backend call to ${path}. Set it in your .env to enable real-time Excel writeback.`);
    return { ok: false, error: "API base URL not configured" };
  }
  try {
    const token = await auth.currentUser?.getIdToken();
    if (!token) return { ok: false, error: "Not signed in" };
    const res = await fetch(`${API_BASE}${path}`, {
      method,
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
      body: body ? JSON.stringify(body) : undefined,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) return { ok: false, error: data.error || `Request failed (${res.status})` };
    return { ok: true, ...data };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
};

// ── NEW: the three Excel-sync entry points the UI calls. Each is a thin
// wrapper around the backend endpoint responsible for one file. Kept as
// separate named functions (rather than one generic call) so each call
// site stays readable about which file it's touching. ──
const syncWarrantyClientToExcel = (client) =>
  callBackendApi("/api/warranty-clients", { body: client });

const syncServiceClientToExcel = (visit) =>
  callBackendApi("/api/service-clients", { body: visit });

// file: "ledger" | "flat" | "both" — lets the delete handler tell the
// backend which workbook(s) to remove the row from, based on which
// sheets the customer record actually came from (sourceSheets).
const syncCustomerDeleteToExcel = (customerId, phone, file) =>
  callBackendApi(`/api/customers/${encodeURIComponent(customerId)}`, {
    method: "DELETE",
    body: { phone, file },
  });

// Picks the first sheet that actually has a cell grid ('!ref'), skipping
// chart sheets. XLSX.js includes chart sheets in SheetNames but they have
// no '!ref' — sheet_to_json on one silently returns [] with no error.
const pickDataSheet = (wb) => {
  const named = wb.SheetNames.find(n => /^sheet\s*1$/i.test(n) && wb.Sheets[n]?.["!ref"]);
  if (named) return wb.Sheets[named];
  const withData = wb.SheetNames.find(n => wb.Sheets[n]?.["!ref"]);
  return withData ? wb.Sheets[withData] : wb.Sheets[wb.SheetNames[0]];
};

// Strips a phone string down to the last 10 digits. Handles cells that
// hold multiple numbers separated by commas/slashes by returning an array.
const extractPhones = (raw) => {
  if (!raw) return [];
  return String(raw)
    .split(/[,/]/)
    .map(s => s.replace(/\D/g, ""))
    .filter(s => s.length >= 10)
    .map(s => s.slice(-10));
};

// Normalizes a customer name for fuzzy matching across sheets: strips
// titles (Mr./Mrs./Dr.), parenthetical nicknames, punctuation, and case.
const normalizeName = (raw) => {
  if (!raw) return "";
  return String(raw)
    .toLowerCase()
    .replace(/\([^)]*\)/g, " ")
    .replace(/\b(mr|mrs|ms|miss|dr|smt|shri)\.?\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
};

const slugify = (raw) =>
  String(raw).trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 100);

// Parses a single date token in d/m/yy, d/m/yyyy, d.m.yyyy, or d-m-yyyy
// form (the formats found across both source sheets) into an ISO string.
const parseDateToken = (token) => {
  const m = String(token).trim().match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2,4})$/);
  if (!m) return null;
  let [, d, mo, y] = m;
  d = Number(d); mo = Number(mo); y = Number(y);
  if (y < 100) y += y <= 49 ? 2000 : 1900;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const iso = `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  const dt = new Date(iso);
  if (isNaN(dt)) return null;
  return iso;
};

// A cell may already be a JS Date (when read with cellDates:true). XLSX's
// date conversion can leave tiny rounding artifacts (e.g. a date landing at
// 23:59:50 instead of 00:00:00), which — if read with local Date methods —
// can shift the calendar date by a day. Rounding to the nearest UTC day
// first avoids that.
const roundToUTCDay = (dateObj) => new Date(Math.round(dateObj.getTime() / 86400000) * 86400000);

const normalizeDateValue = (val) => {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val)) {
    const r = roundToUTCDay(val);
    return `${r.getUTCFullYear()}-${String(r.getUTCMonth() + 1).padStart(2, "0")}-${String(r.getUTCDate()).padStart(2, "0")}`;
  }
  return parseDateToken(val);
};

// For display: Date objects get formatted as DD-MM-YYYY; anything else
// (free text like "warranty cus", or a plain date string) is shown as-is.
// This is what fixes the "Wed Sep 17 2025 23:59:50 GMT+0530..." bug — that
// was raw Date.toString() being dumped straight into the table before.
const formatDateForDisplay = (val) => {
  if (!val) return "";
  if (val instanceof Date && !isNaN(val)) {
    const r = roundToUTCDay(val);
    return `${String(r.getUTCDate()).padStart(2, "0")}-${String(r.getUTCMonth() + 1).padStart(2, "0")}-${r.getUTCFullYear()}`;
  }
  return String(val).trim();
};

// ── Best-effort parser for the Ledger's free-text "services" column ──
// Real examples: "9/12/19 full survice" · "SPUN CHANGING FINISH, spun
// chaning march, 6/10/2021 spun changing" · "free survice finish".
// Strategy: find every date-shaped token in the string; the text between
// one date and the next (or start/end of string) becomes that service's
// description. A chunk with no date anywhere still counts as one
// undated service entry rather than being dropped, since technicians
// often logged "full survice" style notes without a date.
const DATE_TOKEN_RE = /\d{1,2}[.\/\-]\d{1,2}[.\/\-]\d{2,4}/g;
const parseLedgerServiceText = (text) => {
  const str = String(text || "").trim();
  if (!str || str.length < 2) return [];
  const matches = [...str.matchAll(DATE_TOKEN_RE)];
  if (matches.length === 0) {
    return [{ dateISO: null, description: str.replace(/^[,.\s]+|[,.\s]+$/g, "") }];
  }
  const entries = [];
  for (let i = 0; i < matches.length; i++) {
    const cur = matches[i];
    const start = cur.index + cur[0].length;
    const end = i + 1 < matches.length ? matches[i + 1].index : str.length;
    const desc = str.slice(start, end).replace(/^[,.\s]+|[,.\s]+$/g, "");
    entries.push({ dateISO: parseDateToken(cur[0]), description: desc || "(no description)" });
  }
  return entries;
};

/* ════════════════════════ COMPONENT ════════════════════════ */
const AdminDashboard = () => {
  const [customers, setCustomers] = useState([]);
  const [staff, setStaff] = useState([
    { id: "S1", name: "Arun Kumar",   phone: "9884401234", status: "Offline", isBusy: false, currentCustomers: [] },
    { id: "S2", name: "Suresh Raina", phone: "9884405678", status: "Offline", isBusy: false, currentCustomers: [] },
  ]);
  const [searchTerm, setSearchTerm]     = useState("");
  const [filterStatus, setFilterStatus] = useState("All");
  const [activeTab, setActiveTab]       = useState("Directory");
  const [selectedCustomer, setSelectedCustomer] = useState(null);
  const [uploadStatus, setUploadStatus] = useState("");
  const [dark, setDark]                 = useState(false);
  const [historyTab, setHistoryTab]     = useState("Arun Kumar");
  const [history, setHistory]           = useState([]);
  // ── NEW: Store Sale records shown as a third "staff" tab in History ──
  const [storeSalesHistory, setStoreSalesHistory] = useState([]);
  // ── NEW: Analysis tab — raw data pulled from BOTH completedJobs (all
  // staff, unfiltered) and storeSales, combined client-side into
  // day/month/year revenue comparisons. ──
  const [analysisJobs, setAnalysisJobs] = useState([]);
  const [analysisSales, setAnalysisSales] = useState([]);
  const [analysisLoading, setAnalysisLoading] = useState(false);
  const [analysisGranularity, setAnalysisGranularity] = useState("month"); // "day" | "month" | "year"
  const [analysisMetric, setAnalysisMetric] = useState("revenue"); // "revenue" | "orders"
  const [sidebarOpen, setSidebarOpen]   = useState(true);
  const [lastLedgerUpdated, setLastLedgerUpdated] = useState(null);
  const [lastFlatUpdated, setLastFlatUpdated] = useState(null);
  const [syncStatus, setSyncStatus] = useState(""); // "syncing" | "done" | ""
  const [syncDiagnostics, setSyncDiagnostics] = useState(null);
  // shape: { ledgerFound, flatFound, ledgerRows, flatRows, mergedCount, error }
  const [historyPreviewId, setHistoryPreviewId] = useState(null);
  const [historyPreviewPos, setHistoryPreviewPos] = useState({ x: 0, y: 0 });

  /* ── Inventory Management state ── */
  const [inventory, setInventory] = useState([]);
  const [inventorySearch, setInventorySearch] = useState("");
  const [inventoryImportStatus, setInventoryImportStatus] = useState(""); // "importing" | "done:N" | "error" | ""
  const inventoryFileRef = useRef(null);
  // ── NEW: tracks which exact set of low-stock items the banner was last
  // dismissed for — re-shows automatically if the set changes (e.g. a
  // different/additional item drops low), rather than staying hidden
  // forever after one dismissal. ──
  const [lowStockDismissedKey, setLowStockDismissedKey] = useState("");

  const [qtyOverrides, setQtyOverrides] = useState({});
  const qtyTimers = useRef({});

  /* ── Direct Store Sale state ── */
  const storeItemIdCounter = useRef(0);
  const makeStoreRow = () => ({ id: `store-item-${storeItemIdCounter.current++}`, item: "", quantity: "", price: "" });
  const [storeCustomerName, setStoreCustomerName] = useState("");
  const [storeCustomerPhone, setStoreCustomerPhone] = useState("");
  const [storeItems, setStoreItems] = useState(() => [makeStoreRow()]);
  const [storeSubmitting, setStoreSubmitting] = useState(false);

  /* ══════════════════════════════════════════════════════════
     NEW — Quotation Maker state. Builds a professional PDF quotation
     for bulk/institutional clients (schools, apartments, companies),
     uploads it to Storage, saves a record to Firestore for history/
     resend, and opens WhatsApp pre-filled with a message + the PDF
     link, addressed to the phone number entered in the form.
  ══════════════════════════════════════════════════════════ */
  const quotItemIdCounter = useRef(0);
  const makeQuotItemRow = () => ({ id: `quot-item-${quotItemIdCounter.current++}`, description: "", quantity: "1", price: "" });
  const [quotClientName, setQuotClientName] = useState("");
  const [quotContactPerson, setQuotContactPerson] = useState("");
  const [quotPhone, setQuotPhone] = useState("");
  const [quotEmail, setQuotEmail] = useState("");
  const [quotAddress, setQuotAddress] = useState("");
  const [quotValidUntil, setQuotValidUntil] = useState(() => {
    const d = new Date(); d.setDate(d.getDate() + 15);
    return d.toISOString().split("T")[0];
  });
  const [quotTaxPercent, setQuotTaxPercent] = useState("18");
  const [quotTerms, setQuotTerms] = useState(
    "1. Prices are valid until the date mentioned above.\n" +
    "2. 50% advance required to confirm the order; balance on completion.\n" +
    "3. Installation and standard warranty included as per company policy.\n" +
    "4. Any additional plumbing/electrical work will be charged separately.\n" +
    "5. Delivery/installation timeline: 3-7 working days from confirmation."
  );
  const [quotItems, setQuotItems] = useState(() => [makeQuotItemRow()]);
  const [quotSubmitting, setQuotSubmitting] = useState(false);
  const [quotations, setQuotations] = useState([]);
  const [quotationsLoading, setQuotationsLoading] = useState(false);

  const [form, setForm] = useState({
    name: "", phone: "", address: "", nagar: "", place: "",
    machineModel: "", price: "", installDate: new Date().toISOString().split("T")[0],
    amcExpiry: "", paymentStatus: "Pending", remarks: "",
  });
  const [addFormMode, setAddFormMode] = useState("warranty"); // "warranty" | "service"
  const [serviceForm, setServiceForm] = useState({
    customerName: "", phone: "", address: "",
    serviceDate: new Date().toISOString().split("T")[0],
    serviceDescription: "", amount: "",
  });
  const [serviceFormStatus, setServiceFormStatus] = useState(""); // "" | "saving" | "error:<msg>"
  const [warrantyFormStatus, setWarrantyFormStatus] = useState(""); // "" | "saving" | "error:<msg>"
  // ── NEW: soft, non-blocking banner shown when Firestore saved fine but
  // the backend Excel writeback failed/was skipped — keeps the failure
  // visible instead of the old silent console.error-only behavior. ──
  const [excelSyncWarning, setExcelSyncWarning] = useState("");

  /* ══════════════════════════════════════════════════════════
     NEW — state for the five additions below: Stock Movement History
     (Inventory sub-view), AMC Renewals dashboard, Customer Feedback
     modal, Staff Performance tab, and a Schedule/job-queue tab.
  ══════════════════════════════════════════════════════════ */
  // Stock Movement History (a toggle within the existing Inventory tab)
  const [inventorySubView, setInventorySubView] = useState("stock"); // "stock" | "history"
  const [movementHistory, setMovementHistory] = useState([]);
  const [movementHistoryLoading, setMovementHistoryLoading] = useState(false);
  const [movementSearch, setMovementSearch] = useState("");

  // Customer Feedback modal (opened from the Directory table)
  const [feedbackTarget, setFeedbackTarget] = useState(null); // customer object or null
  const [feedbackRating, setFeedbackRating] = useState(0);
  const [feedbackNote, setFeedbackNote] = useState("");
  const [feedbackSaving, setFeedbackSaving] = useState(false);

  // Staff Performance tab
  const [performanceJobs, setPerformanceJobs] = useState([]);
  const [performanceLoading, setPerformanceLoading] = useState(false);

  // Schedule tab (today's assigned job queue, per technician)
  const [scheduleTasks, setScheduleTasks] = useState([]);
  const [scheduleLoading, setScheduleLoading] = useState(false);

  /* ══════════════════════════════════════════════════════════
     NEW — Duplicate Customer detection & merge. See detection/merge
     logic and UI further down. State here just tracks the modal and
     per-cluster merge-in-progress status.
  ══════════════════════════════════════════════════════════ */
  const [showDuplicatesModal, setShowDuplicatesModal] = useState(false);
  const [mergingClusterKey, setMergingClusterKey] = useState(null);
  const [dupBannerDismissedCount, setDupBannerDismissedCount] = useState(0);
  const [dupFilter, setDupFilter] = useState("all"); // "all" | "phone" | "name"

  /* ── Design tokens ── */
  const C = dark ? {
    bg:       "#070d1a",
    sidebar:  "#0c1526",
    card:     "#111d2e",
    cardHov:  "#162236",
    glass:    "rgba(17,29,46,0.85)",
    border:   "rgba(255,255,255,0.07)",
    borderHi: "rgba(99,179,237,0.35)",
    text:     "#e8f0fe",
    sub:      "#94a3b8",
    inp:      "#0a1525",
    accent:   "#3b82f6",
    accentHi: "#60a5fa",
    success:  "#10b981",
    danger:   "#ef4444",
    warn:     "#f59e0b",
  } : {
    bg:       "#f0f4ff",
    sidebar:  "#ffffff",
    card:     "#ffffff",
    cardHov:  "#f8faff",
    glass:    "rgba(255,255,255,0.92)",
    border:   "#e8edf5",
    borderHi: "#bfdbfe",
    text:     "#0f172a",
    sub:      "#64748b",
    inp:      "#f8fafc",
    accent:   "#3b82f6",
    accentHi: "#2563eb",
    success:  "#10b981",
    danger:   "#ef4444",
    warn:     "#f59e0b",
  };

  /* ── Firestore live listener ── */
  useEffect(() => {
    const todayDate = new Date().toISOString().split("T")[0];

    const qC = query(collection(db, "customers"), orderBy("installDate", "desc"));
    const u1 = onSnapshot(qC, s => setCustomers(s.docs.map(d => ({ id: d.id, ...d.data() }))));

    const qS = query(collection(db, "staffStatus"));
    const u2 = onSnapshot(qS, async s => {
      const db2 = s.docs.map(d => ({ id: d.id, ...d.data() }));

      for (const staffDoc of s.docs) {
        const { name, status } = staffDoc.data();
        if (status === "Online") {
          const attSnap = await getDocs(query(collection(db, "attendance"), where("name", "==", name), where("date", "==", todayDate)));
          if (attSnap.empty) {
            await updateDoc(doc(db, "staffStatus", staffDoc.id), { status: "Offline" });
          }
        }
      }

      setStaff(p => p.map(x => { const f = db2.find(y => y.name === x.name || y.id === x.id); return f ? { ...x, ...f } : x; }));
    });
    return () => { u1(); u2(); };
  }, []);

  const customersRef = useRef([]);
  const historyCloseTimerRef = useRef(null);
  const openHistoryPreview = (id, pos) => {
    if (historyCloseTimerRef.current) { clearTimeout(historyCloseTimerRef.current); historyCloseTimerRef.current = null; }
    if (pos) setHistoryPreviewPos(pos);
    setHistoryPreviewId(id);
  };
  const scheduleCloseHistoryPreview = () => {
    if (historyCloseTimerRef.current) clearTimeout(historyCloseTimerRef.current);
    historyCloseTimerRef.current = setTimeout(() => setHistoryPreviewId(null), 250);
  };
  useEffect(() => { customersRef.current = customers; }, [customers]);

  /* ══════════════════════════════════════════════════════════
     Poll Firebase Storage every 15 s for BOTH warranty sheets
     (Warranty_Ledger.xlsx + Warranty_Flat.xlsx). These are the
     *upstream* workbooks technicians/owners may edit directly — this
     read-only sync imports them into Firestore. It's separate from the
     real-time writeback path below, which pushes admin-added clients
     the other direction (Firestore → Excel-Warranty_Ledger.xlsx /
     Excel-Warranty_Flat.xlsx) via the backend API.
  ══════════════════════════════════════════════════════════ */
  useEffect(() => {
    const LEDGER_PATH = LEDGER_STORAGE_PATH;
    const FLAT_PATH = FLAT_STORAGE_PATH;

    const readWithDetectedHeader = (ws, headerMustInclude) => {
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
      let headerRowIdx = raw.findIndex(row =>
        row.some(cell => String(cell).toUpperCase().includes(headerMustInclude))
      );
      if (headerRowIdx === -1) headerRowIdx = 0;
      const headers = raw[headerRowIdx].map(h => String(h).trim());
      return raw.slice(headerRowIdx + 1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
        return obj;
      });
    };

    const parseLedgerWorkbook = (wb) => {
      const ws = pickDataSheet(wb);
      const rows = readWithDetectedHeader(ws, "NAME");
      const customersByPhone = new Map();   // phone -> record
      const customersByName = new Map();    // normalized name -> record
      const allRecords = [];

      for (const row of rows) {
        const name = String(row["NAME"] || "").trim();
        if (!name) continue;
        const phones = extractPhones(row["PHONE NUMBER"]);
        const serviceEntries = parseLedgerServiceText(row["services"]);
        const record = {
          name,
          phone: phones[0] || "",
          phoneAlt: phones[1] || "",
          address: String(row["ADDRESS"] || "").trim(),
          place: String(row["PLACE"] || "").trim(),
          machineModel: String(row["MODEL MECHINE"] || "").trim(),
          price: row["PRICE"] || "",
          installDate: formatDateForDisplay(row["DATE"]),
          installDateISO: normalizeDateValue(row["DATE"]),
          remarksFromLedger: String(row["FEEDBACK"] || "").trim(),
          ledgerServiceHistory: serviceEntries,
          warrantyServicesCompleted: serviceEntries.length,
          warrantyPending: Math.max(0, WARRANTY_FREE_SERVICES - serviceEntries.length),
        };
        allRecords.push(record);
        if (record.phone) customersByPhone.set(record.phone, record);
        const nName = normalizeName(name);
        if (nName) customersByName.set(nName, record);
      }
      return { allRecords, customersByPhone, customersByName };
    };

    const parseFlatWorkbook = (wb) => {
      const finalCleanName = wb.SheetNames.find(n => /final.*clean/i.test(n) && wb.Sheets[n]?.["!ref"]);
      const cleanName = wb.SheetNames.find(n => /clean/i.test(n) && wb.Sheets[n]?.["!ref"]);
      const ws = finalCleanName ? wb.Sheets[finalCleanName] : (cleanName ? wb.Sheets[cleanName] : pickDataSheet(wb));
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

      const pick = (row, keys) => {
        for (const k of keys) {
          if (row[k] !== undefined && row[k] !== "") return row[k];
        }
        return "";
      };

      const groups = new Map(); // key -> { name, phone, phoneAlt, address, place, services: [] }
      for (const row of rows) {
        const name = String(pick(row, ["Customer Name", "Name"])).trim();
        const phones = extractPhones(pick(row, ["Phone No 1", "Phone Number", "phone number", "Phone"]));
        const phone2 = extractPhones(pick(row, ["Phone No 2"]))[0] || "";
        if (!name && !phones.length) continue;

        const key = phones[0] || `name:${normalizeName(name)}`;
        if (!groups.has(key)) {
          groups.set(key, {
            name, phone: phones[0] || "", phoneAlt: phone2,
            address: String(pick(row, ["Address"])).trim(),
            place: String(pick(row, ["Place"])).trim(),
            services: [],
          });
        }
        const g = groups.get(key);
        const desc = String(pick(row, ["Service Description"])).trim();
        const dateISO = normalizeDateValue(pick(row, ["Service Date"]));
        if (desc || dateISO) {
          g.services.push({ dateISO, description: desc || "(no description)", amount: pick(row, ["Amount"]) });
        }
      }
      return groups;
    };

    const syncWarrantyDataToFirestore = async () => {
      try {
        const ledgerRef = ref(storage, LEDGER_PATH);
        const flatRef = ref(storage, FLAT_PATH);
        const [ledgerMeta, flatMeta] = await Promise.all([
          getMetadata(ledgerRef).catch(() => null),
          getMetadata(flatRef).catch(() => null),
        ]);
        const ledgerUpdated = ledgerMeta?.updated || null;
        const flatUpdated = flatMeta?.updated || null;

        if (ledgerUpdated === lastLedgerUpdated && flatUpdated === lastFlatUpdated) return;
        setLastLedgerUpdated(ledgerUpdated);
        setLastFlatUpdated(flatUpdated);
        setSyncStatus("syncing");

        if (!ledgerMeta) {
          console.warn(`[Warranty Sync] "${LEDGER_PATH}" was not found in Firebase Storage — check the exact filename/path/case. Ledger data will be skipped this sync.`);
        }
        if (!flatMeta) {
          console.warn(`[Warranty Sync] "${FLAT_PATH}" was not found in Firebase Storage — check the exact filename/path/case. Flat data will be skipped this sync.`);
        }

        let ledgerParsed = { allRecords: [], customersByPhone: new Map(), customersByName: new Map() };
        if (ledgerMeta) {
          const url = await getDownloadURL(ledgerRef);
          const buf = await (await fetch(url)).arrayBuffer();
          ledgerParsed = parseLedgerWorkbook(XLSX.read(buf, { type: "array", cellDates: true }));
          console.log(`[Warranty Sync] Ledger parsed: ${ledgerParsed.allRecords.length} customer rows.`);
        }

        let flatGroups = new Map();
        if (flatMeta) {
          const url = await getDownloadURL(flatRef);
          const buf = await (await fetch(url)).arrayBuffer();
          flatGroups = parseFlatWorkbook(XLSX.read(buf, { type: "array", cellDates: true }));
          console.log(`[Warranty Sync] Flat parsed: ${flatGroups.size} customer groups.`);
        }

        const merged = new Map(); // docKey -> merged record
        for (const rec of ledgerParsed.allRecords) {
          const docKey = rec.phone ? `phone_${rec.phone}` : `name_${slugify(rec.name)}`;
          merged.set(docKey, { ...rec, flatServiceHistory: [], sourceSheets: ["ledger"] });
        }

        for (const [key, g] of flatGroups) {
          let target = null;
          if (g.phone && ledgerParsed.customersByPhone.has(g.phone)) {
            target = ledgerParsed.customersByPhone.get(g.phone);
          } else {
            const nName = normalizeName(g.name);
            if (nName && ledgerParsed.customersByName.has(nName)) {
              target = ledgerParsed.customersByName.get(nName);
            }
          }

          if (target) {
            const docKey = target.phone ? `phone_${target.phone}` : `name_${slugify(target.name)}`;
            const entry = merged.get(docKey);
            entry.flatServiceHistory.push(...g.services);
            if (!entry.sourceSheets.includes("flat")) entry.sourceSheets.push("flat");
            if (!entry.phone && g.phone) entry.phone = g.phone;
          } else {
            const docKey = g.phone ? `phone_${g.phone}` : `name_${slugify(g.name)}`;
            if (!merged.has(docKey)) {
              merged.set(docKey, {
                name: g.name, phone: g.phone, phoneAlt: g.phoneAlt,
                address: g.address, place: g.place, machineModel: "", price: "",
                installDate: "", installDateISO: null, remarksFromLedger: "",
                ledgerServiceHistory: [], warrantyServicesCompleted: 0,
                warrantyPending: WARRANTY_FREE_SERVICES,
                flatServiceHistory: [], sourceSheets: ["flat"],
              });
            }
            merged.get(docKey).flatServiceHistory.push(...g.services);
          }
        }

        // ── Skip re-adding any customer that was deliberately deleted from
        // the dashboard. Without this, a customer who originated in the
        // upstream Warranty_Ledger.xlsx / Warranty_Flat.xlsx files would
        // reappear on the next 15s sync even after being deleted, since
        // this sync only knows how to add/update, never delete, and the
        // upstream files themselves aren't touched by the dashboard's
        // delete action. ──
        const excludedSnap = await getDocs(collection(db, "excludedCustomers"));
        const excludedIds = new Set(excludedSnap.docs.map(d => d.id));

        const existingIds = new Set(customersRef.current.map(c => c.id));
        const batch = writeBatch(db);
        for (const [docKey, rec] of merged) {
          if (excludedIds.has(docKey)) continue;
          const flatDates = rec.flatServiceHistory.map(s => s.dateISO).filter(Boolean).sort();
          const ledgerDates = rec.ledgerServiceHistory.map(s => s.dateISO).filter(Boolean).sort();
          const lastServiceDate = flatDates.at(-1) || ledgerDates.at(-1) || null;

          const allServiceHistory = [
            ...rec.ledgerServiceHistory.map(s => ({ ...s, source: "Warranty (Ledger)" })),
            ...rec.flatServiceHistory.map(s => ({ ...s, source: "Extra Service (Flat)" })),
          ].sort((a, b) => (a.dateISO || "").localeCompare(b.dateISO || ""));

          const data = {
            name: rec.name,
            phone: rec.phone,
            phoneAlt: rec.phoneAlt || "",
            address: rec.address,
            place: rec.place,
            machineModel: rec.machineModel || "General RO",
            price: rec.price,
            installDate: rec.installDate,
            installDateISO: rec.installDateISO,
            remarksFromLedger: rec.remarksFromLedger,
            ledgerServiceHistory: rec.ledgerServiceHistory,
            flatServiceHistory: rec.flatServiceHistory,
            allServiceHistory,
            warrantyServicesCompleted: rec.warrantyServicesCompleted,
            warrantyPending: rec.warrantyPending,
            lastServiceDate,
            sourceSheets: rec.sourceSheets,
            isMasterSynced: true,
          };
          if (!existingIds.has(docKey)) {
            data.paymentStatus = "N/A";
            data.remarks = "";
            data.lastNotified = "";
          }

          // ── NEW — permanent fix for the "customer count silently grew"
          // bug: this sync computes docKey fresh from the CURRENT sheet
          // contents every run. If a customer previously had no phone
          // number (so they were written under a "name_<slug>" doc ID)
          // and a phone number later appears for them in the sheet, this
          // run computes a brand-new "phone_<number>" docKey — which,
          // without this check, creates a SECOND Firestore doc instead of
          // migrating the first one, silently inflating Total Clients.
          // Detect that exact case and delete the stale name-keyed doc in
          // the same batch, carrying over any admin-entered fields
          // (payment status, remarks) so nothing is lost. ──
          if (docKey.startsWith("phone_")) {
            const staleNameKey = `name_${slugify(rec.name)}`;
            const staleDoc = customersRef.current.find(c => c.id === staleNameKey);
            if (staleDoc && staleDoc.id !== docKey) {
              if (staleDoc.paymentStatus && staleDoc.paymentStatus !== "N/A" && !data.paymentStatus) data.paymentStatus = staleDoc.paymentStatus;
              if (staleDoc.remarks) data.remarks = staleDoc.remarks;
              batch.delete(doc(db, "customers", staleNameKey));
            }
          }

          batch.set(doc(db, "customers", docKey), data, { merge: true });
        }
        await batch.commit();
        setSyncDiagnostics({
          ledgerFound: !!ledgerMeta, flatFound: !!flatMeta,
          ledgerRows: ledgerParsed.allRecords.length, flatRows: flatGroups.size,
          mergedCount: merged.size, error: null,
        });
        setSyncStatus("done");
        setTimeout(() => setSyncStatus(""), 3000);
      } catch (e) {
        if (e.code !== "storage/object-not-found") console.error("Warranty sync error:", e);
        setSyncDiagnostics(prev => ({ ...(prev || {}), error: e.message || String(e) }));
        setSyncStatus("");
      }
    };

    syncWarrantyDataToFirestore();
    const interval = setInterval(syncWarrantyDataToFirestore, 15000);
    return () => clearInterval(interval);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lastLedgerUpdated, lastFlatUpdated]);

  useEffect(() => {
    if (activeTab !== "History") return;
    // ── NEW: "Store Sales" is a separate collection (storeSales), not a
    // staff name, so it gets its own listener/query instead of filtering
    // completedJobs by employeeName. ──
    if (historyTab === "Store Sales") {
      const q = query(collection(db, "storeSales"), orderBy("timestamp", "desc"));
      const u = onSnapshot(q, s => {
        setStoreSalesHistory(s.docs.map(d => ({ id: d.id, ...d.data() })));
      });
      return () => u();
    }
    const q = query(collection(db, "completedJobs"), orderBy("timestamp", "desc"));
    const u = onSnapshot(q, s => {
      const all = s.docs.map(d => ({ id: d.id, ...d.data() }));
      setHistory(all.filter(j => j.employeeName?.toLowerCase().includes(historyTab.toLowerCase().split(" ")[0])));
    });
    return () => u();
  }, [activeTab, historyTab]);

  // ── NEW: Analysis tab data loader — pulls ALL completedJobs (every
  // staff member, not filtered like the staff tabs above) and ALL
  // storeSales, so revenue can be aggregated across the whole business
  // rather than one technician or one sale type at a time. ──
  useEffect(() => {
    if (activeTab !== "Analysis") return;
    setAnalysisLoading(true);
    let jobsReady = false, salesReady = false;
    const checkReady = () => { if (jobsReady && salesReady) setAnalysisLoading(false); };

    const qJobs = query(collection(db, "completedJobs"), orderBy("timestamp", "desc"));
    const u1 = onSnapshot(qJobs, s => {
      setAnalysisJobs(s.docs.map(d => d.data()));
      jobsReady = true; checkReady();
    }, () => { jobsReady = true; checkReady(); });

    const qSales = query(collection(db, "storeSales"), orderBy("timestamp", "desc"));
    const u2 = onSnapshot(qSales, s => {
      setAnalysisSales(s.docs.map(d => d.data()));
      salesReady = true; checkReady();
    }, () => { salesReady = true; checkReady(); });

    return () => { u1(); u2(); };
  }, [activeTab]);

  /* ══════════════════════════════════════════════════════════
     NEW — Analysis aggregation helpers. Combines completedJobs (staff
     service bills) and storeSales (walk-in sales) into one revenue
     timeline, then groups it by day/month/year for comparison. All
     computed client-side from the two live listeners above — no extra
     Firestore reads or writes involved.
  ══════════════════════════════════════════════════════════ */
  const analysisRecords = useMemo(() => {
    const fromJobs = analysisJobs
      .filter(j => j.date)
      .map(j => ({ dateISO: j.date, amount: Number(j.grandTotal) || 0, source: "service" }));
    const fromSales = analysisSales
      .filter(s => s.date)
      .map(s => ({ dateISO: s.date, amount: Number(s.grandTotal) || 0, source: "store" }));
    return [...fromJobs, ...fromSales];
  }, [analysisJobs, analysisSales]);

  const getAnalysisPeriodKey = (dateISO, granularity) => {
    if (!dateISO) return null;
    if (granularity === "day") return dateISO;
    if (granularity === "month") return dateISO.slice(0, 7);
    return dateISO.slice(0, 4);
  };

  const formatAnalysisPeriodLabel = (key, granularity) => {
    if (granularity === "day") {
      const d = new Date(key);
      return isNaN(d) ? key : d.toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" });
    }
    if (granularity === "month") {
      const [y, m] = key.split("-");
      const d = new Date(Number(y), Number(m) - 1, 1);
      return isNaN(d) ? key : d.toLocaleDateString("en-IN", { month: "short", year: "numeric" });
    }
    return key; // year is already human-readable
  };

  // Grouped totals per period, sorted most-recent-first — feeds the
  // breakdown table. { key, store, service, total, storeOrders,
  // serviceOrders, totalOrders }[]. Order count = number of individual
  // sale/job records in that period, distinct from the ₹ amount.
  const analysisAggregated = useMemo(() => {
    const map = new Map();
    for (const r of analysisRecords) {
      const key = getAnalysisPeriodKey(r.dateISO, analysisGranularity);
      if (!key) continue;
      if (!map.has(key)) map.set(key, { key, store: 0, service: 0, storeOrders: 0, serviceOrders: 0 });
      const entry = map.get(key);
      if (r.source === "store") { entry.store += r.amount; entry.storeOrders += 1; }
      else { entry.service += r.amount; entry.serviceOrders += 1; }
    }
    return Array.from(map.values())
      .map(e => ({ ...e, total: e.store + e.service, totalOrders: e.storeOrders + e.serviceOrders }))
      .sort((a, b) => b.key.localeCompare(a.key));
  }, [analysisRecords, analysisGranularity]);

  // Last 12 periods in chronological order — feeds the bar chart.
  const analysisChartData = useMemo(() => {
    return [...analysisAggregated].sort((a, b) => a.key.localeCompare(b.key)).slice(-12);
  }, [analysisAggregated]);

  // Today / this month / this year / all-time summary cards — both ₹
  // revenue and order counts.
  const analysisSummary = useMemo(() => {
    const todayISO = new Date().toISOString().split("T")[0];
    const monthKey = todayISO.slice(0, 7);
    const yearKey = todayISO.slice(0, 4);
    let todayTotal = 0, monthTotal = 0, yearTotal = 0, allTimeTotal = 0;
    let todayOrders = 0, monthOrders = 0, yearOrders = 0, allTimeOrders = 0;
    for (const r of analysisRecords) {
      allTimeTotal += r.amount; allTimeOrders += 1;
      if (r.dateISO === todayISO) { todayTotal += r.amount; todayOrders += 1; }
      if (r.dateISO?.slice(0, 7) === monthKey) { monthTotal += r.amount; monthOrders += 1; }
      if (r.dateISO?.slice(0, 4) === yearKey) { yearTotal += r.amount; yearOrders += 1; }
    }
    return {
      todayTotal, monthTotal, yearTotal, allTimeTotal,
      todayOrders, monthOrders, yearOrders, allTimeOrders,
      recordCount: analysisRecords.length,
    };
  }, [analysisRecords]);

  /* ══════════════════════════════════════════════════════════
     NEW — Inventory Stock Analysis. There's no dedicated stock-movement
     log in this app, so "units moved" is derived from the items array
     already recorded on each Direct Store Sale (analysisSales, loaded
     by the effect above) — no extra Firestore reads needed. Current
     stock levels come from the `inventory` collection listener that's
     already running elsewhere in this component. Note: this only
     captures units sold through Direct Store Sale; units used during a
     staff-completed service job aren't tracked here since that
     deduction happens outside this file.
  ══════════════════════════════════════════════════════════ */
  const inventoryMovementRecords = useMemo(() => {
    return analysisSales.flatMap(s =>
      (s.items || [])
        .map(it => ({ dateISO: s.date, itemName: String(it.item || "").trim(), qty: Number(it.quantity) || 0 }))
        .filter(r => r.dateISO && r.itemName && r.qty > 0)
    );
  }, [analysisSales]);

  // Total units sold per period, sorted most-recent-first.
  const inventoryAggregated = useMemo(() => {
    const map = new Map();
    for (const r of inventoryMovementRecords) {
      const key = getAnalysisPeriodKey(r.dateISO, analysisGranularity);
      if (!key) continue;
      map.set(key, (map.get(key) || 0) + r.qty);
    }
    return Array.from(map.entries())
      .map(([key, units]) => ({ key, units }))
      .sort((a, b) => b.key.localeCompare(a.key));
  }, [inventoryMovementRecords, analysisGranularity]);

  const inventoryChartData = useMemo(() => {
    return [...inventoryAggregated].sort((a, b) => a.key.localeCompare(b.key)).slice(-12);
  }, [inventoryAggregated]);

  // Best-selling items overall, matched against current stock so a low
  // remaining count next to a high sell-through rate stands out.
  const topSellingItems = useMemo(() => {
    const totals = new Map();
    for (const r of inventoryMovementRecords) {
      totals.set(r.itemName, (totals.get(r.itemName) || 0) + r.qty);
    }
    return Array.from(totals.entries())
      .map(([itemName, unitsSold]) => {
        const match = inventory.find(inv => (inv.itemName || "").toLowerCase() === itemName.toLowerCase());
        return { itemName, unitsSold, currentStock: match ? Number(match.quantity) || 0 : null };
      })
      .sort((a, b) => b.unitsSold - a.unitsSold)
      .slice(0, 8);
  }, [inventoryMovementRecords, inventory]);

  // Current stock snapshot — independent of the date range, always
  // reflects live inventory right now.
  const stockOverview = useMemo(() => {
    const totalItems = inventory.length;
    const totalUnits = inventory.reduce((sum, i) => sum + (Number(i.quantity) || 0), 0);
    const lowStock = inventory.filter(i => Number(i.quantity) > 0 && Number(i.quantity) <= 5).length;
    const outOfStock = inventory.filter(i => Number(i.quantity) <= 0).length;
    return { totalItems, totalUnits, lowStock, outOfStock };
  }, [inventory]);

  /* ══════════════════════════════════════════════════════════
     NEW — AMC Renewals dashboard data. Buckets every customer with an
     amcExpiry date by how soon it's due, using the `customers` list
     already loaded live at the top of this component — no extra
     Firestore reads needed.
  ══════════════════════════════════════════════════════════ */
  const renewalBuckets = useMemo(() => {
    const todayDate = new Date(); todayDate.setHours(0, 0, 0, 0);
    const buckets = { overdue: [], today: [], next7: [], next30: [] };
    for (const c of customers) {
      if (!c.amcExpiry) continue;
      const due = new Date(c.amcExpiry);
      if (isNaN(due)) continue;
      due.setHours(0, 0, 0, 0);
      const daysUntil = Math.round((due - todayDate) / 86400000);
      const entry = { ...c, daysUntil };
      if (daysUntil < 0) buckets.overdue.push(entry);
      else if (daysUntil === 0) buckets.today.push(entry);
      else if (daysUntil <= 7) buckets.next7.push(entry);
      else if (daysUntil <= 30) buckets.next30.push(entry);
    }
    for (const key of Object.keys(buckets)) buckets[key].sort((a, b) => a.daysUntil - b.daysUntil);
    return buckets;
  }, [customers]);

  const handleAmcRemind = (c) => {
    const cleanPhone = (c.phone || "").toString().replace(/\D/g, "");
    if (cleanPhone.length < 10) {
      alert(`${c.name || "This customer"} doesn't have a valid phone number on file.`);
      return;
    }
    const msg = `Hi ${c.name || "there"}, this is a reminder from AquaServe that your RO service/AMC renewal is due on ${c.amcExpiry}. Please let us know a convenient time to schedule it. Thank you!`;
    window.open(`https://wa.me/${cleanPhone}?text=${encodeURIComponent(msg)}`, "_blank");
  };

  /* ══════════════════════════════════════════════════════════
     NEW — Customer Feedback modal handlers. Stores a manually-recorded
     rating + note directly on the customer's Firestore doc, since
     collecting it automatically from the customer themselves would
     need a customer-facing channel this app doesn't have.
  ══════════════════════════════════════════════════════════ */
  const openFeedbackModal = (c) => {
    setFeedbackTarget(c);
    setFeedbackRating(c.feedbackRating || 0);
    setFeedbackNote(c.feedbackNote || "");
  };
  const closeFeedbackModal = () => { setFeedbackTarget(null); setFeedbackRating(0); setFeedbackNote(""); };
  const saveFeedback = async () => {
    if (!feedbackTarget) return;
    setFeedbackSaving(true);
    try {
      await updateDoc(doc(db, "customers", feedbackTarget.id), {
        feedbackRating, feedbackNote: feedbackNote.trim(),
        feedbackUpdatedAt: serverTimestamp(),
      });
      closeFeedbackModal();
    } catch (e) {
      console.error("Feedback save error:", e);
      alert("Failed to save feedback. Please try again.");
    } finally {
      setFeedbackSaving(false);
    }
  };

  /* ══════════════════════════════════════════════════════════
     NEW — Staff Performance aggregation. Groups completedJobs by
     employeeName: job count, total revenue, average per job.
  ══════════════════════════════════════════════════════════ */
  const staffPerformance = useMemo(() => {
    const map = new Map();
    for (const j of performanceJobs) {
      const name = j.employeeName || "Unknown";
      if (!map.has(name)) map.set(name, { name, jobCount: 0, revenue: 0 });
      const entry = map.get(name);
      entry.jobCount += 1;
      entry.revenue += Number(j.grandTotal) || 0;
    }
    return Array.from(map.values())
      .map(e => ({ ...e, avgPerJob: e.jobCount > 0 ? e.revenue / e.jobCount : 0 }))
      .sort((a, b) => b.revenue - a.revenue);
  }, [performanceJobs]);

  /* ══════════════════════════════════════════════════════════
     NEW — Schedule: groups the current assigned-job queue by
     technician, using the live `scheduleTasks` (tasks collection,
     status "Assigned") and `staff` list already in this component.
  ══════════════════════════════════════════════════════════ */
  const scheduleByStaff = useMemo(() => {
    const map = new Map();
    for (const s of staff) map.set(s.name, []);
    for (const t of scheduleTasks) {
      const name = t.employeeName || "Unassigned";
      if (!map.has(name)) map.set(name, []);
      map.get(name).push(t);
    }
    for (const [, list] of map) list.sort((a, b) => (a.timestamp || "").localeCompare(b.timestamp || ""));
    return Array.from(map.entries()).map(([name, tasks]) => ({ name, tasks }));
  }, [staff, scheduleTasks]);

  /* ══════════════════════════════════════════════════════════
     NEW — Low stock alerts. Derived live from the `inventory` listener
     (already running app-wide, not gated to the Inventory tab), so this
     is visible from anywhere in the dashboard rather than only when the
     Inventory tab happens to be open.
  ══════════════════════════════════════════════════════════ */
  const lowStockAlerts = useMemo(() => inventory.filter(i => Number(i.quantity) <= 5), [inventory]);
  const lowStockKey = useMemo(() => lowStockAlerts.map(i => i.id).sort().join(","), [lowStockAlerts]);
  const showLowStockBanner = lowStockAlerts.length > 0 && lowStockDismissedKey !== lowStockKey;

  /* ══════════════════════════════════════════════════════════
     NEW — Duplicate Customer detection. Groups every customer doc by
     normalized phone (last 10 digits) AND by normalized name (existing
     `normalizeName` helper), then unions any docs that share either —
     a simple union-find over the whole `customers` list, already live-
     loaded at the top of this component, so this costs zero extra
     Firestore reads. Any resulting group with more than one doc is a
     likely duplicate needing manual review (auto-merging is
     deliberately NOT automatic, since two different real people can
     legitimately share a phone or a similar name — e.g. family members
     — and a wrong auto-merge would destroy real data).
  ══════════════════════════════════════════════════════════ */
  const duplicateClusters = useMemo(() => {
    const byPhone = new Map();
    const byName = new Map();
    for (const c of customers) {
      const p = (c.phone || "").toString().replace(/\D/g, "").slice(-10);
      if (p.length === 10) {
        if (!byPhone.has(p)) byPhone.set(p, []);
        byPhone.get(p).push(c);
      }
      const n = normalizeName(c.name);
      if (n) {
        if (!byName.has(n)) byName.set(n, []);
        byName.get(n).push(c);
      }
    }
    const parent = new Map();
    customers.forEach(c => parent.set(c.id, c.id));
    const find = (id) => { let r = id; while (parent.get(r) !== r) r = parent.get(r); return r; };
    const union = (a, b) => { const ra = find(a), rb = find(b); if (ra !== rb) parent.set(ra, rb); };
    for (const group of byPhone.values()) if (group.length > 1) for (let i = 1; i < group.length; i++) union(group[i].id, group[0].id);
    for (const group of byName.values()) if (group.length > 1) for (let i = 1; i < group.length; i++) union(group[i].id, group[0].id);

    const clusters = new Map();
    for (const c of customers) {
      const root = find(c.id);
      if (!clusters.has(root)) clusters.set(root, []);
      clusters.get(root).push(c);
    }
    // ── NEW: tag each cluster by WHY its members got grouped —
    // "phone" if two or more members in the group actually share the
    // identical phone number (a strong, low-risk signal to merge), vs
    // "name" if they were only grouped because their names look similar
    // despite having different phone numbers (much weaker signal — could
    // easily be two different people, e.g. father & son, or a common
    // name in the area — needs a closer look before merging). ──
    return Array.from(clusters.entries())
      .filter(([, group]) => group.length > 1)
      .map(([key, group]) => {
        const phoneCounts = new Map();
        for (const m of group) {
          const p = (m.phone || "").toString().replace(/\D/g, "").slice(-10);
          if (p.length === 10) phoneCounts.set(p, (phoneCounts.get(p) || 0) + 1);
        }
        const hasSamePhone = Array.from(phoneCounts.values()).some(count => count > 1);
        return { key, members: group, matchType: hasSamePhone ? "phone" : "name" };
      });
  }, [customers]);

  const showDuplicatesBanner = duplicateClusters.length > 0 && dupBannerDismissedCount !== duplicateClusters.length;

  // ── NEW: counts per match type + the filtered list the modal actually
  // renders. "phone" clusters are the safe, high-confidence ones (an
  // identical phone number really is shared); "name" clusters are only
  // grouped by a similar-looking name with no shared phone at all, and
  // deserve more scrutiny before merging. ──
  const dupPhoneCount = useMemo(() => duplicateClusters.filter(c => c.matchType === "phone").length, [duplicateClusters]);
  const dupNameCount = useMemo(() => duplicateClusters.filter(c => c.matchType === "name").length, [duplicateClusters]);
  const filteredDuplicateClusters = useMemo(() => {
    if (dupFilter === "all") return duplicateClusters;
    return duplicateClusters.filter(c => c.matchType === dupFilter);
  }, [duplicateClusters, dupFilter]);

  // ── Merge a cluster of likely-duplicate customer docs into one.
  // Canonical doc = prefer a "phone_"-keyed doc (the current, correct ID
  // scheme) with the most service history; union all service history +
  // sourceSheets from every doc in the cluster into it; delete the rest,
  // and — critically — record each deleted ID in "excludedCustomers" so
  // the 15s upstream sync never recreates them, AND remove them from the
  // synced Excel files via the same backend the manual delete button
  // uses, so Excel and Firestore stay consistent. ──
  const handleMergeCluster = async (cluster) => {
    setMergingClusterKey(cluster.key);
    try {
      const members = [...cluster.members];
      const canonical = members
        .slice()
        .sort((a, b) => {
          const aPhone = a.id.startsWith("phone_") ? 1 : 0;
          const bPhone = b.id.startsWith("phone_") ? 1 : 0;
          if (aPhone !== bPhone) return bPhone - aPhone;
          const aHist = (a.allServiceHistory || []).length;
          const bHist = (b.allServiceHistory || []).length;
          return bHist - aHist;
        })[0];
      const others = members.filter(m => m.id !== canonical.id);

      const mergedLedgerHistory = [];
      const mergedFlatHistory = [];
      const mergedSourceSheets = new Set(canonical.sourceSheets || []);
      const seenHistoryKeys = new Set();
      const addHistory = (arr, target) => {
        for (const h of (arr || [])) {
          const key = `${h.dateISO || ""}|${h.description || ""}`;
          if (seenHistoryKeys.has(key)) continue;
          seenHistoryKeys.add(key);
          target.push(h);
        }
      };
      addHistory(canonical.ledgerServiceHistory, mergedLedgerHistory);
      addHistory(canonical.flatServiceHistory, mergedFlatHistory);
      for (const o of others) {
        addHistory(o.ledgerServiceHistory, mergedLedgerHistory);
        addHistory(o.flatServiceHistory, mergedFlatHistory);
        for (const s of (o.sourceSheets || [])) mergedSourceSheets.add(s);
      }
      const mergedAllHistory = [
        ...mergedLedgerHistory.map(h => ({ ...h, source: "Warranty (Ledger)" })),
        ...mergedFlatHistory.map(h => ({ ...h, source: "Extra Service (Flat)" })),
      ].sort((a, b) => (a.dateISO || "").localeCompare(b.dateISO || ""));

      const fallback = (field) => canonical[field] || others.find(o => o[field])?.[field] || "";

      await setDoc(doc(db, "customers", canonical.id), {
        ...canonical,
        address: fallback("address"),
        place: fallback("place"),
        machineModel: fallback("machineModel"),
        price: canonical.price || others.find(o => o.price)?.price || "",
        remarks: fallback("remarks"),
        remarksFromLedger: fallback("remarksFromLedger"),
        ledgerServiceHistory: mergedLedgerHistory,
        flatServiceHistory: mergedFlatHistory,
        allServiceHistory: mergedAllHistory,
        warrantyServicesCompleted: mergedLedgerHistory.length,
        warrantyPending: Math.max(0, WARRANTY_FREE_SERVICES - mergedLedgerHistory.length),
        sourceSheets: Array.from(mergedSourceSheets),
      }, { merge: true });

      const batch = writeBatch(db);
      for (const o of others) {
        batch.delete(doc(db, "customers", o.id));
        batch.set(doc(db, "excludedCustomers", o.id), {
          deletedAt: serverTimestamp(),
          deletedBy: auth?.currentUser?.uid || "unknown",
          reason: `Merged into ${canonical.id}`,
          name: o.name || "",
          phone: o.phone || "",
        });
      }
      await batch.commit();

      // Best-effort: also remove the merged-away rows from the synced
      // Excel files, same as the regular delete button does. A failure
      // here doesn't undo the merge — Firestore is already consistent.
      for (const o of others) {
        const sources = o.sourceSheets || [];
        const fileTarget = sources.includes("ledger") && sources.includes("flat")
          ? "both" : sources.includes("flat") ? "flat" : "ledger";
        syncCustomerDeleteToExcel(o.id, o.phone, fileTarget).catch(() => {});
      }
    } catch (e) {
      console.error("Merge duplicates error:", e);
      alert("Failed to merge these records. Please try again.");
    } finally {
      setMergingClusterKey(null);
    }
  };

  // ── NEW: live listener for saved quotations, active only while the
  // Quotation tab is open (same pattern as Job History above) ──
  useEffect(() => {
    if (activeTab !== "Quotation") return;
    setQuotationsLoading(true);
    const q = query(collection(db, "quotations"), orderBy("createdAt", "desc"));
    const u = onSnapshot(q, s => {
      setQuotations(s.docs.map(d => ({ id: d.id, ...d.data() })));
      setQuotationsLoading(false);
    }, () => setQuotationsLoading(false));
    return () => u();
  }, [activeTab]);

  // ── NEW: Stock Movement History — loads only when Inventory tab is
  // open AND the "History" sub-view is selected, so browsing the raw
  // log doesn't happen on every Inventory tab visit. ──
  useEffect(() => {
    if (activeTab !== "Inventory" || inventorySubView !== "history") return;
    setMovementHistoryLoading(true);
    const q = query(collection(db, "stockMovements"), orderBy("timestamp", "desc"));
    const u = onSnapshot(q, s => {
      setMovementHistory(s.docs.map(d => ({ id: d.id, ...d.data() })));
      setMovementHistoryLoading(false);
    }, () => setMovementHistoryLoading(false));
    return () => u();
  }, [activeTab, inventorySubView]);

  // ── NEW: Staff Performance — loads ALL completedJobs (unfiltered by
  // employee) only while the Performance tab is open. ──
  useEffect(() => {
    if (activeTab !== "Performance") return;
    setPerformanceLoading(true);
    const q = query(collection(db, "completedJobs"), orderBy("timestamp", "desc"));
    const u = onSnapshot(q, s => {
      setPerformanceJobs(s.docs.map(d => d.data()));
      setPerformanceLoading(false);
    }, () => setPerformanceLoading(false));
    return () => u();
  }, [activeTab]);

  // ── NEW: Schedule — the current job queue (assigned, not yet
  // completed) grouped by technician. Loads only while the Schedule tab
  // is open. ──
  useEffect(() => {
    if (activeTab !== "Schedule") return;
    setScheduleLoading(true);
    const q = query(collection(db, "tasks"), where("status", "==", "Assigned"));
    const u = onSnapshot(q, s => {
      setScheduleTasks(s.docs.map(d => ({ id: d.id, ...d.data() })));
      setScheduleLoading(false);
    }, () => setScheduleLoading(false));
    return () => u();
  }, [activeTab]);

  useEffect(() => {
    const qInv = query(collection(db, "inventory"), orderBy("itemName", "asc"));
    const uInv = onSnapshot(qInv, s => setInventory(s.docs.map(d => ({ id: d.id, ...d.data() }))));
    return () => uInv();
  }, []);

  useEffect(() => {
    return () => {
      Object.values(qtyTimers.current).forEach(t => t && clearTimeout(t));
    };
  }, []);

  /* ── Logic ── */
  const updateCloudExcel = async (list) => {
    try {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(list.map(c => ({
        Name: c.name, Phone: c.phone, Address: c.address, Nagar: c.nagar,
        Place: c.place, Model: c.machineModel, InstallationDate: c.installDate,
        AMCExpiry: c.amcExpiry, PaymentStatus: c.paymentStatus, Remarks: c.remarks || "",
      }))), "Sheet1");
      const blob = new Blob([XLSX.write(wb, { bookType: "xlsx", type: "array" })],
        { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      await uploadBytes(ref(storage, "company_data/Customer_Master.xlsx"), blob);
    } catch (e) { console.error(e); }
  };

  const sendAutoReminders = async () => {
    const today = new Date().toISOString().split("T")[0];
    const due = customers.filter(c => c.amcExpiry === today && c.lastNotified !== today);
    if (!due.length) { alert("All due customers have already been notified!"); return; }
    for (const [i, c] of due.entries()) {
      setTimeout(async () => {
        await updateDoc(doc(db, "customers", c.id), { lastNotified: today });
        window.open(`https://wa.me/${c.phone}?text=${encodeURIComponent(`Hi ${c.name}, your RO service is due today (${c.amcExpiry}). Our technician will contact you shortly.`)}`, "_blank");
      }, i * 2000);
    }
  };

  const handleWhatsAppRemind = (c) => {
    const cleanPhone = (c.phone || "").toString().replace(/\D/g, "");
    if (cleanPhone.length < 10) {
      alert(`${c.name || "This customer"} doesn't have a valid phone number on file, so WhatsApp can't be opened.`);
      return;
    }
    const msg = c.warrantyPending > 0
      ? `Hi ${c.name || "there"}, this is a reminder from AquaServe regarding your ${c.machineModel || "water purifier"}. You have ${c.warrantyPending} free warranty service${c.warrantyPending > 1 ? "s" : ""} remaining. Please let us know a convenient time to schedule it. Thank you!`
      : `Hi ${c.name || "there"}, this is a reminder from AquaServe. Your included warranty services have been fully used — feel free to reach out if you'd like to book a paid service visit. Thank you!`;
    window.open(`https://wa.me/${cleanPhone}?text=${encodeURIComponent(msg)}`, "_blank");
  };

  const processedList = [...customers]
    .filter(c => {
      const q = searchTerm.toLowerCase();
      const ok = c.name?.toLowerCase().includes(q) || c.phone?.includes(searchTerm) || c.nagar?.toLowerCase().includes(q);
      return ok && (filterStatus === "All" || filterStatus === "Alphabetical" ||
        (filterStatus === "PendingService" && c.warrantyPending > 0) ||
        (filterStatus === "Warranty" && c.sourceSheets?.includes("ledger")) ||
        (filterStatus === "NonWarranty" && c.sourceSheets && !c.sourceSheets.includes("ledger")));
    })
    .sort((a, b) => filterStatus === "Alphabetical" ? (a.name || "").localeCompare(b.name || "") : 0);

  /* ══════════════════════════════════════════════════════════
     Warranty Client form → Firestore "customers" ledger, PLUS a
     real-time write into Excel-Warranty_Ledger.xlsx via the backend API.
     Firestore is the source of truth and is saved first; the Excel sync
     is a best-effort follow-up call that never blocks or reverts the
     Firestore save if it fails — instead it surfaces a dismissible
     warning banner so a failure is visible, not silent.
  ══════════════════════════════════════════════════════════ */
  const handleAddWarrantyClient = async (e) => {
    e.preventDefault();
    setWarrantyFormStatus("");
    setExcelSyncWarning("");
    const cleanPhone = form.phone.replace(/\D/g, "");
    if (!form.name.trim()) { setWarrantyFormStatus("error:Client name is required."); return; }
    if (cleanPhone.length !== 10) { setWarrantyFormStatus("error:Phone number must be exactly 10 digits."); return; }
    if (form.price && isNaN(Number(form.price))) { setWarrantyFormStatus("error:Price must be a number."); return; }
    if (!form.installDate) { setWarrantyFormStatus("error:Installation date is required."); return; }

    setWarrantyFormStatus("saving");
    try {
      const inst = new Date(form.installDate); inst.setMonth(inst.getMonth() + 4);
      const amcExp = inst.toISOString().split("T")[0];
      const docId = `phone_${cleanPhone}`;
      const existing = customersRef.current.find(c => c.id === docId);

      const payload = {
        name: form.name.trim(),
        phone: cleanPhone,
        address: form.address.trim(),
        nagar: form.nagar.trim(),
        place: form.place.trim(),
        machineModel: form.machineModel.trim(),
        price: form.price ? Number(form.price) : "",
        installDate: form.installDate,
        installDateISO: form.installDate,
        amcExpiry: amcExp,
        remarksFromLedger: form.remarks.trim(),
        sourceSheets: existing ? Array.from(new Set([...(existing.sourceSheets || []), "ledger"])) : ["ledger"],
        isMasterSynced: false,
        createdAt: existing?.createdAt || serverTimestamp(),
        createdBy: auth?.currentUser?.uid || "unknown",
      };
      if (!existing) {
        payload.warrantyServicesCompleted = 0;
        payload.warrantyPending = WARRANTY_FREE_SERVICES;
        payload.ledgerServiceHistory = [];
        payload.flatServiceHistory = [];
        payload.allServiceHistory = [];
        payload.paymentStatus = form.paymentStatus;
        payload.remarks = form.remarks.trim();
        payload.lastNotified = "";
      }

      await setDoc(doc(db, "customers", docId), payload, { merge: true });

      // ── Real-time Excel writeback, via backend — replaces the old
      // browser-side download/edit/re-upload of the whole workbook.
      // Isolated from the Firestore save above: if this fails or the
      // backend is unreachable, the client is still saved and shows up
      // on the dashboard regardless. ──
      const excelResult = await syncWarrantyClientToExcel({
        customerId: docId,
        name: payload.name,
        phone: payload.phone,
        address: payload.address,
        place: payload.place,
        machineModel: payload.machineModel,
        price: payload.price,
        installDate: payload.installDate,
        remarks: payload.remarksFromLedger,
      });
      if (!excelResult.ok) {
        console.error("Excel writeback (Ledger) failed:", excelResult.error);
        setExcelSyncWarning(`Client saved, but the Excel-Warranty_Ledger.xlsx update failed: ${excelResult.error}. It will need a manual retry or resync.`);
      }

      setWarrantyFormStatus("done");
      setTimeout(() => setWarrantyFormStatus(""), 2500);
      setForm({ name: "", phone: "", address: "", nagar: "", place: "", machineModel: "", price: "", installDate: new Date().toISOString().split("T")[0], amcExpiry: "", paymentStatus: "Pending", remarks: "" });
    } catch (err) {
      console.error("Warranty client save error:", err);
      setWarrantyFormStatus(`error:${err.message || "Failed to save. Please try again."}`);
    }
  };

  /* ══════════════════════════════════════════════════════════
     Service Client form → Firestore, shaped like the
     Warranty_Flat.xlsx "Final_Cleaned" data, PLUS a real-time write into
     Excel-Warranty_Flat.xlsx via the backend API. Uses a Firestore
     transaction so two admins logging a service for the same customer at
     once can't silently overwrite each other's entry.
  ══════════════════════════════════════════════════════════ */
  const handleAddServiceClient = async (e) => {
    e.preventDefault();
    setServiceFormStatus("");
    setExcelSyncWarning("");
    const cleanPhone = serviceForm.phone.replace(/\D/g, "");
    if (!serviceForm.customerName.trim()) { setServiceFormStatus("error:Customer name is required."); return; }
    if (cleanPhone.length !== 10) { setServiceFormStatus("error:Phone number must be exactly 10 digits."); return; }
    if (!serviceForm.serviceDate) { setServiceFormStatus("error:Service date is required."); return; }
    if (!serviceForm.serviceDescription.trim()) { setServiceFormStatus("error:Service description is required."); return; }
    if (serviceForm.amount && isNaN(Number(serviceForm.amount))) { setServiceFormStatus("error:Amount must be a number."); return; }

    setServiceFormStatus("saving");
    try {
      const docId = `phone_${cleanPhone}`;
      const docRef = doc(db, "customers", docId);
      const newEntry = {
        dateISO: serviceForm.serviceDate,
        description: serviceForm.serviceDescription.trim(),
        amount: serviceForm.amount ? Number(serviceForm.amount) : "",
      };

      await runTransaction(db, async (tx) => {
        const snap = await tx.get(docRef);
        if (snap.exists()) {
          const data = snap.data();
          const flatServiceHistory = [...(data.flatServiceHistory || []), newEntry];
          const allServiceHistory = [
            ...(data.ledgerServiceHistory || []).map(s => ({ ...s, source: "Warranty (Ledger)" })),
            ...flatServiceHistory.map(s => ({ ...s, source: "Extra Service (Flat)" })),
          ].sort((a, b) => (a.dateISO || "").localeCompare(b.dateISO || ""));
          const lastServiceDate = flatServiceHistory.map(s => s.dateISO).filter(Boolean).sort().at(-1) || data.lastServiceDate || null;
          tx.set(docRef, {
            flatServiceHistory, allServiceHistory, lastServiceDate,
            sourceSheets: Array.from(new Set([...(data.sourceSheets || []), "flat"])),
            name: data.name || serviceForm.customerName.trim(),
            address: data.address || serviceForm.address.trim(),
          }, { merge: true });
        } else {
          tx.set(docRef, {
            name: serviceForm.customerName.trim(),
            phone: cleanPhone,
            address: serviceForm.address.trim(),
            place: "", nagar: "", machineModel: "", price: "",
            installDate: "", installDateISO: null, remarksFromLedger: "",
            ledgerServiceHistory: [], warrantyServicesCompleted: 0, warrantyPending: WARRANTY_FREE_SERVICES,
            flatServiceHistory: [newEntry],
            allServiceHistory: [{ ...newEntry, source: "Extra Service (Flat)" }],
            lastServiceDate: newEntry.dateISO,
            sourceSheets: ["flat"],
            isMasterSynced: false,
            createdAt: serverTimestamp(),
            createdBy: auth?.currentUser?.uid || "unknown",
          });
        }
      });

      // ── Real-time Excel writeback, via backend — isolated from the
      // transaction above; a failure here never blocks the service visit
      // from being logged for the dashboard. ──
      const excelResult = await syncServiceClientToExcel({
        customerId: docId,
        customerName: serviceForm.customerName.trim(),
        phone: cleanPhone,
        address: serviceForm.address.trim(),
        serviceDate: serviceForm.serviceDate,
        serviceDescription: serviceForm.serviceDescription.trim(),
        amount: newEntry.amount,
      });
      if (!excelResult.ok) {
        console.error("Excel writeback (Flat) failed:", excelResult.error);
        setExcelSyncWarning(`Service visit logged, but the Excel-Warranty_Flat.xlsx update failed: ${excelResult.error}. It will need a manual retry or resync.`);
      }

      setServiceFormStatus("done");
      setTimeout(() => setServiceFormStatus(""), 2500);
      setServiceForm({ customerName: "", phone: "", address: "", serviceDate: new Date().toISOString().split("T")[0], serviceDescription: "", amount: "" });
    } catch (err) {
      console.error("Service client save error:", err);
      setServiceFormStatus(`error:${err.message || "Failed to save. Please try again."}`);
    }
  };

  // ── Autofill Service form's name/address when a matching phone is found ──
  const handleServicePhoneBlur = () => {
    const cleanPhone = serviceForm.phone.replace(/\D/g, "");
    if (cleanPhone.length !== 10) return;
    const match = customersRef.current.find(c => c.phone === cleanPhone);
    if (match && !serviceForm.customerName) {
      setServiceForm(f => ({ ...f, customerName: match.name || f.customerName, address: match.address || f.address }));
    }
  };

  /* ══════════════════════════════════════════════════════════
     Delete a client — removes from Firestore, refreshes the legacy
     Customer_Master.xlsx export, AND now also removes the matching row
     from Excel-Warranty_Ledger.xlsx / Excel-Warranty_Flat.xlsx in real
     time via the backend, based on which sheet(s) the record came from.
  ══════════════════════════════════════════════════════════ */
  const handleDelete = async (id) => {
    if (!window.confirm("Remove this client from the database and Master Excel?")) return;
    setExcelSyncWarning("");
    const target = customersRef.current.find(c => c.id === id);

    await deleteDoc(doc(db, "customers", id));
    // ── Record this deletion so the 15s upstream re-import (which reads
    // Warranty_Ledger.xlsx / Warranty_Flat.xlsx) doesn't silently bring
    // this customer back if their row is still sitting in that source
    // file. This does NOT touch the original upstream file itself — it
    // only tells the sync to skip re-adding this specific customer ID. ──
    await setDoc(doc(db, "excludedCustomers", id), {
      deletedAt: serverTimestamp(),
      deletedBy: auth?.currentUser?.uid || "unknown",
      name: target?.name || "",
      phone: target?.phone || "",
    });
    const all = await getDocs(collection(db, "customers"));
    updateCloudExcel(all.docs.map(d => d.data()));

    // Decide which workbook(s) to remove this customer's row from, based
    // on where their record actually came from (a client added only via
    // the Service form, for example, was never in the Ledger file).
    const sources = target?.sourceSheets || [];
    const fileTarget = sources.includes("ledger") && sources.includes("flat")
      ? "both"
      : sources.includes("flat") ? "flat" : "ledger";

    const excelResult = await syncCustomerDeleteToExcel(id, target?.phone, fileTarget);
    if (!excelResult.ok) {
      console.error("Excel row delete failed:", excelResult.error);
      setExcelSyncWarning(`Client removed from the dashboard, but removing the row from Excel failed: ${excelResult.error}. It will need a manual retry or resync.`);
    }
  };

  const handleAssign = async (emp) => {
    if (!selectedCustomer) return;
    try {
      const snap = await getDocs(query(collection(db, "staffStatus"), where("name", "==", emp.name)));
      if (snap.empty) { alert("Staff not found"); return; }
      await updateDoc(doc(db, "staffStatus", snap.docs[0].id), {
        currentCustomers: arrayUnion({ name: selectedCustomer.name, address: selectedCustomer.address, id: selectedCustomer.id }),
        lastUpdated: new Date().toISOString(),
      });
      await addDoc(collection(db, "tasks"), { customerId: selectedCustomer.id, customerName: selectedCustomer.name, employeeName: emp.name, status: "Assigned", timestamp: new Date().toISOString() });
      window.open(`https://wa.me/${emp.phone}?text=${encodeURIComponent(`Hi ${emp.name}, new task: ${selectedCustomer.name}`)}`, "_blank");
      setSelectedCustomer(null);
      alert("Technician assigned successfully!");
    } catch (e) { console.error(e); alert("Failed to assign."); }
  };

  const exportToExcel = () => {
    if (!processedList.length) { alert("No data to export!"); return; }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(processedList.map(c => ({
      "Client Name": c.name || "", "Phone Number": c.phone || "", Address: c.address || "",
      Nagar: c.nagar || "", Place: c.place || "", "Machine Model": c.machineModel || "General RO",
      "Installation Date": c.installDate || "", "Service Due Date": c.amcExpiry || "",
      "Payment Status": c.paymentStatus || "", Remarks: c.remarks || "",
    }))), "Clients");
    XLSX.writeFile(wb, `Aquaserve_Pro_Clients_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  /* ── Inventory helpers ── */
  const slugifyItemName = (name) =>
    String(name).trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 150);

  const normalizeKey = (k) => String(k).toLowerCase().replace(/[^a-z0-9]/g, "");

  const getFieldValue = (row, candidates) => {
    const normalizedRow = {};
    for (const key of Object.keys(row)) {
      normalizedRow[normalizeKey(key)] = row[key];
    }
    for (const cand of candidates) {
      const val = normalizedRow[normalizeKey(cand)];
      if (val !== undefined && val !== null && String(val).trim() !== "") return val;
    }
    return "";
  };

  /* ══════════════════════════════════════════════════════════
     NEW — Stock movement log. Every quantity change now writes a record
     to the "stockMovements" collection instead of only overwriting the
     inventory doc's quantity field, so there's a real history of *why*
     an item's stock is what it is (import, manual edit, or a store
     sale) rather than just the current number.
  ══════════════════════════════════════════════════════════ */
  const buildStockMovementPayload = ({ itemId, itemName, change, reason, newQuantity }) => ({
    itemId, itemName, change, reason, newQuantity,
    date: new Date().toISOString().split("T")[0],
    timestamp: serverTimestamp(),
    by: auth?.currentUser?.uid || "unknown",
  });

  // For use inside an existing writeBatch (import, store sale) — queues
  // the log entry alongside the inventory write so both commit together.
  const logStockMovementInBatch = (batch, entry) => {
    batch.set(doc(collection(db, "stockMovements")), buildStockMovementPayload(entry));
  };

  // For a standalone quantity edit that isn't already part of a batch.
  const logStockMovementNow = async (entry) => {
    try {
      await addDoc(collection(db, "stockMovements"), buildStockMovementPayload(entry));
    } catch (e) {
      console.error("Stock movement log error:", e);
    }
  };

  const handleInventoryImport = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      setInventoryImportStatus("importing");
      const arrayBuffer = await file.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);

      const batch = writeBatch(db);
      let count = 0;
      for (const row of rows) {
        const name = String(getFieldValue(row, ["Item Name", "ItemName", "Item", "Name", "Product Name", "Product"])).trim();
        if (!name) continue;
        const sno = getFieldValue(row, ["S.No", "SNo", "Sno", "SNO", "S No", "Serial No", "Serial Number", "Sl No"]);
        const qtyRaw = getFieldValue(row, ["Quantity", "Qty", "Item Quantity", "ItemQuantity", "Stock", "Stock Quantity", "Available Quantity"]);
        const qty = Number(qtyRaw) || 0;
        const docId = slugifyItemName(name) || `item_${count}`;

        // ── NEW: log the delta caused by this import row against
        // whatever quantity that item currently has, so a re-import that
        // doesn't actually change anything doesn't clutter the log. ──
        const existingItem = inventory.find(i => i.id === docId);
        const prevQty = existingItem ? Number(existingItem.quantity) || 0 : 0;
        const delta = qty - prevQty;

        batch.set(doc(db, "inventory", docId), {
          itemName: name,
          quantity: qty,
          sno: String(sno),
          lastImported: new Date().toISOString(),
        }, { merge: true });

        if (delta !== 0) {
          logStockMovementInBatch(batch, { itemId: docId, itemName: name, change: delta, reason: "Excel Import", newQuantity: qty });
        }
        count++;
      }
      await batch.commit();
      setInventoryImportStatus(`done:${count}`);
      setTimeout(() => setInventoryImportStatus(""), 3500);
    } catch (err) {
      console.error("Inventory import error:", err);
      setInventoryImportStatus("error");
      setTimeout(() => setInventoryImportStatus(""), 3500);
    } finally {
      if (inventoryFileRef.current) inventoryFileRef.current.value = "";
    }
  };

  // ── UPDATED: now takes the full item object (not just its id) so it
  // can compute and log the delta against the previous quantity. ──
  const handleInventoryQtyUpdate = async (item, newQty) => {
    try {
      const prevQty = Number(item.quantity) || 0;
      await updateDoc(doc(db, "inventory", item.id), { quantity: newQty });
      const delta = newQty - prevQty;
      if (delta !== 0) {
        await logStockMovementNow({ itemId: item.id, itemName: item.itemName, change: delta, reason: "Manual edit", newQuantity: newQty });
      }
    } catch (e) { console.error("Inventory quantity update error:", e); }
  };

  const handleInventoryQtyInputChange = (item, rawValue) => {
    setQtyOverrides(prev => ({ ...prev, [item.id]: rawValue }));

    if (qtyTimers.current[item.id]) clearTimeout(qtyTimers.current[item.id]);
    qtyTimers.current[item.id] = setTimeout(() => {
      const val = Number(rawValue);
      if (!isNaN(val) && val !== item.quantity) {
        handleInventoryQtyUpdate(item, val);
      }
      setQtyOverrides(prev => {
        const next = { ...prev };
        delete next[item.id];
        return next;
      });
      qtyTimers.current[item.id] = null;
    }, 500);
  };

  const handleInventoryQtyBlur = (item, rawValue) => {
    if (qtyTimers.current[item.id]) {
      clearTimeout(qtyTimers.current[item.id]);
      qtyTimers.current[item.id] = null;
    }
    const val = Number(rawValue);
    if (!isNaN(val) && val !== item.quantity) {
      handleInventoryQtyUpdate(item, val);
    }
    setQtyOverrides(prev => {
      const next = { ...prev };
      delete next[item.id];
      return next;
    });
  };

  const handleInventoryDelete = async (id) => {
    if (!window.confirm("Remove this item from Inventory Management?")) return;
    try { await deleteDoc(doc(db, "inventory", id)); } catch (e) { console.error(e); }
  };

  const filteredInventory = inventory.filter(item =>
    (item.itemName || "").toLowerCase().includes(inventorySearch.toLowerCase())
  );

  const exportInventoryToExcel = () => {
    if (!filteredInventory.length) { alert("No inventory data to export!"); return; }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      filteredInventory.map(item => ({
        "S.No": item.sno || "",
        "Item Name": item.itemName || "",
        "Quantity": item.quantity ?? 0,
      }))
    ), "Inventory");
    XLSX.writeFile(wb, `Aquaserve_Pro_Inventory_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  /* ── Direct Store Sale helpers ── */
  const updateStoreItem = (id, field, value) =>
    setStoreItems(prev => prev.map(row => (row.id === id ? { ...row, [field]: value } : row)));
  const addStoreItemRow = () => setStoreItems(prev => [...prev, makeStoreRow()]);
  const removeStoreItemRow = (id) => setStoreItems(prev => prev.filter(row => row.id !== id));
  const storeRowTotal = (row) => (Number(row.quantity) || 0) * (Number(row.price) || 0);
  const storeItemsTotal = storeItems.reduce((sum, row) => sum + storeRowTotal(row), 0);

  const findInventoryMatch = (usedName) => {
    const nameLc = (usedName || "").trim().toLowerCase();
    if (!nameLc) return null;
    return inventory.find(inv => {
      const invName = (inv.itemName || "").trim().toLowerCase();
      return invName && (invName.includes(nameLc) || nameLc.includes(invName));
    });
  };

  const handleCompleteStoreSale = async () => {
    if (storeSubmitting) return;
    const validItems = storeItems.filter(row => row.item.trim() !== "");
    if (!storeCustomerName.trim() || !storeCustomerPhone.trim()) {
      alert("Please enter the customer's name and phone number.");
      return;
    }
    if (!validItems.length) {
      alert("Please add at least one item sold.");
      return;
    }

    setStoreSubmitting(true);
    try {
      const saleDate = new Date().toISOString().split("T")[0];
      const dateLabel = new Date().toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" });
      const grandTotal = storeItemsTotal;

      const docPdf = new jsPDF();
      docPdf.setFontSize(18);
      docPdf.text(`Store Sale Bill for ${storeCustomerName}`, 14, 20);
      docPdf.setFontSize(10);
      docPdf.text(`Date: ${dateLabel}`, 14, 30);
      docPdf.text(`Phone: ${storeCustomerPhone}`, 14, 35);

      autoTable(docPdf, {
        startY: 45,
        head: [["Item Name", "Qty", "Price", "Total"]],
        body: validItems.map(i => [i.item, i.quantity, i.price, storeRowTotal(i)]),
      });

      docPdf.text(`Grand Total: Rs. ${grandTotal}`, 14, docPdf.lastAutoTable.finalY + 10);

      const pdfBlob = docPdf.output("blob");
      const fileName = `billproff/storeSale_${storeCustomerName}_${Date.now()}.pdf`;
      const storageRef = ref(storage, fileName);
      await uploadBytes(storageRef, pdfBlob);
      const pdfUrl = await getDownloadURL(storageRef);

      await addDoc(collection(db, "storeSales"), {
        customerName: storeCustomerName,
        customerPhone: storeCustomerPhone,
        date: saleDate,
        timestamp: new Date(),
        items: validItems,
        grandTotal,
        pdfUrl,
        status: "Completed",
      });

      const batch = writeBatch(db);
      let hasUpdates = false;
      for (const used of validItems) {
        const qtyUsed = Number(used.quantity) || 0;
        if (qtyUsed <= 0) continue;
        const match = findInventoryMatch(used.item);
        if (match) {
          batch.update(doc(db, "inventory", match.id), { quantity: increment(-qtyUsed) });
          // ── NEW: log this deduction so the stock movement history (and
          // Inventory Stock Analysis) reflects it with a reason attached,
          // not just a bare quantity change. ──
          logStockMovementInBatch(batch, {
            itemId: match.id, itemName: match.itemName, change: -qtyUsed,
            reason: `Store Sale${storeCustomerName ? ` — ${storeCustomerName}` : ""}`,
            newQuantity: (Number(match.quantity) || 0) - qtyUsed,
          });
          hasUpdates = true;
        }
      }
      if (hasUpdates) await batch.commit();

      const cleanPhone = storeCustomerPhone.toString().replace(/\D/g, "");
      if (cleanPhone.length >= 10) {
        const msg = `Hi ${storeCustomerName}, thank you for your purchase! Please find your bill attached below:\n\n${pdfUrl}\n\nThank you for choosing Aquaserve Pro!`;
        window.open(`https://wa.me/${cleanPhone}?text=${encodeURIComponent(msg)}`, "_blank");
      } else {
        alert("Bill saved, but the phone number doesn't look valid, so WhatsApp couldn't be opened.");
      }

      setStoreCustomerName("");
      setStoreCustomerPhone("");
      setStoreItems([makeStoreRow()]);
    } catch (err) {
      console.error("Store sale error:", err);
      alert("Failed to complete the store sale bill. Please try again.");
    } finally {
      setStoreSubmitting(false);
    }
  };

  /* ══════════════════════════════════════════════════════════
     NEW — Quotation Maker helpers.
     Mirrors the Direct Store Sale PDF pattern (jsPDF + autoTable +
     upload to Storage), but shaped for a formal B2B/institutional
     quotation: client details, itemized services/products, tax,
     validity date, and terms. Saved to Firestore "quotations" for
     history/resend, then sent to the given phone number via WhatsApp.
  ══════════════════════════════════════════════════════════ */
  const updateQuotItem = (id, field, value) =>
    setQuotItems(prev => prev.map(row => (row.id === id ? { ...row, [field]: value } : row)));
  const addQuotItemRow = () => setQuotItems(prev => [...prev, makeQuotItemRow()]);
  const removeQuotItemRow = (id) => setQuotItems(prev => prev.filter(row => row.id !== id));
  const quotRowTotal = (row) => (Number(row.quantity) || 0) * (Number(row.price) || 0);
  const quotSubtotal = quotItems.reduce((sum, row) => sum + quotRowTotal(row), 0);
  const quotTaxAmount = Math.round(quotSubtotal * (Number(quotTaxPercent) || 0) / 100);
  const quotGrandTotal = quotSubtotal + quotTaxAmount;

  // Simple sequential-looking quotation number: AQ-YYYYMM-<short random>.
  // Not strictly collision-proof under heavy concurrent use, but more
  // than sufficient for a single-shop admin tool; each quotation's real
  // uniqueness comes from its Firestore document ID regardless.
  const generateQuotationNumber = () => {
    const now = new Date();
    const ym = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}`;
    const rand = Math.random().toString(36).slice(2, 6).toUpperCase();
    return `AQ-${ym}-${rand}`;
  };

  const handleCreateQuotation = async () => {
    if (quotSubmitting) return;
    const validItems = quotItems.filter(row => row.description.trim() !== "");
    if (!quotClientName.trim()) { alert("Please enter the client/company name."); return; }
    const cleanPhone = quotPhone.toString().replace(/\D/g, "");
    if (cleanPhone.length < 10) { alert("Please enter a valid phone number to send the quotation via WhatsApp."); return; }
    if (!validItems.length) { alert("Please add at least one item or service."); return; }

    setQuotSubmitting(true);
    try {
      const quotationNumber = generateQuotationNumber();
      const dateLabel = new Date().toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" });
      const validUntilLabel = quotValidUntil
        ? new Date(quotValidUntil).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })
        : "";

      // ── Build the quotation PDF ──
      const docPdf = new jsPDF();

      docPdf.setFontSize(20);
      docPdf.setTextColor(37, 99, 235);
      docPdf.text("AquaServe Pro", 14, 20);
      docPdf.setFontSize(10);
      docPdf.setTextColor(100);
      docPdf.text("Water Purifier Sales, Installation & Service", 14, 26);

      docPdf.setFontSize(16);
      docPdf.setTextColor(20);
      docPdf.text("QUOTATION", 196, 20, { align: "right" });
      docPdf.setFontSize(10);
      docPdf.setTextColor(90);
      docPdf.text(`No: ${quotationNumber}`, 196, 26, { align: "right" });
      docPdf.text(`Date: ${dateLabel}`, 196, 31, { align: "right" });
      if (validUntilLabel) docPdf.text(`Valid Until: ${validUntilLabel}`, 196, 36, { align: "right" });

      docPdf.setDrawColor(220);
      docPdf.line(14, 42, 196, 42);

      docPdf.setFontSize(11);
      docPdf.setTextColor(20);
      docPdf.text("Quotation For:", 14, 50);
      docPdf.setFontSize(10);
      docPdf.setTextColor(60);
      let y = 56;
      docPdf.text(quotClientName, 14, y); y += 5;
      if (quotContactPerson) { docPdf.text(`Attn: ${quotContactPerson}`, 14, y); y += 5; }
      if (quotAddress) {
        const addrLines = docPdf.splitTextToSize(quotAddress, 100);
        docPdf.text(addrLines, 14, y); y += addrLines.length * 5;
      }
      if (quotPhone) { docPdf.text(`Phone: ${quotPhone}`, 14, y); y += 5; }
      if (quotEmail) { docPdf.text(`Email: ${quotEmail}`, 14, y); y += 5; }

      const tableStartY = Math.max(y + 6, 78);
      autoTable(docPdf, {
        startY: tableStartY,
        head: [["#", "Description", "Qty", "Unit Price (₹)", "Amount (₹)"]],
        body: validItems.map((row, i) => [
          i + 1, row.description, row.quantity || "1",
          (Number(row.price) || 0).toLocaleString("en-IN"),
          quotRowTotal(row).toLocaleString("en-IN"),
        ]),
        headStyles: { fillColor: [37, 99, 235] },
        columnStyles: { 0: { cellWidth: 10 }, 2: { cellWidth: 18, halign: "center" }, 3: { cellWidth: 32, halign: "right" }, 4: { cellWidth: 32, halign: "right" } },
      });

      let afterTableY = docPdf.lastAutoTable.finalY + 8;
      docPdf.setFontSize(10);
      docPdf.setTextColor(60);
      docPdf.text(`Subtotal:`, 150, afterTableY);
      docPdf.text(`₹${quotSubtotal.toLocaleString("en-IN")}`, 196, afterTableY, { align: "right" });
      afterTableY += 6;
      docPdf.text(`Tax (${quotTaxPercent || 0}%):`, 150, afterTableY);
      docPdf.text(`₹${quotTaxAmount.toLocaleString("en-IN")}`, 196, afterTableY, { align: "right" });
      afterTableY += 7;
      docPdf.setDrawColor(220);
      docPdf.line(140, afterTableY - 4, 196, afterTableY - 4);
      docPdf.setFontSize(12);
      docPdf.setTextColor(20);
      docPdf.text(`Grand Total:`, 150, afterTableY);
      docPdf.text(`₹${quotGrandTotal.toLocaleString("en-IN")}`, 196, afterTableY, { align: "right" });

      if (quotTerms.trim()) {
        let termsY = afterTableY + 14;
        docPdf.setFontSize(11);
        docPdf.setTextColor(20);
        docPdf.text("Terms & Conditions:", 14, termsY);
        termsY += 6;
        docPdf.setFontSize(9);
        docPdf.setTextColor(80);
        const termsLines = docPdf.splitTextToSize(quotTerms, 182);
        docPdf.text(termsLines, 14, termsY);
      }

      docPdf.setFontSize(9);
      docPdf.setTextColor(140);
      docPdf.text("Thank you for considering AquaServe Pro.", 14, 285);

      const pdfBlob = docPdf.output("blob");
      const fileName = `quotations/${quotationNumber}_${Date.now()}.pdf`;
      const storageRef = ref(storage, fileName);
      await uploadBytes(storageRef, pdfBlob);
      const pdfUrl = await getDownloadURL(storageRef);

      // ── Save the quotation record for history/resend ──
      await addDoc(collection(db, "quotations"), {
        quotationNumber,
        clientName: quotClientName.trim(),
        contactPerson: quotContactPerson.trim(),
        phone: cleanPhone,
        email: quotEmail.trim(),
        address: quotAddress.trim(),
        items: validItems.map(({ id, ...rest }) => rest),
        subtotal: quotSubtotal,
        taxPercent: Number(quotTaxPercent) || 0,
        taxAmount: quotTaxAmount,
        grandTotal: quotGrandTotal,
        validUntil: quotValidUntil,
        terms: quotTerms,
        pdfUrl,
        createdAt: serverTimestamp(),
        createdBy: auth?.currentUser?.uid || "unknown",
        status: "Sent",
      });

      // ── Send via WhatsApp ──
      const msg = `Hi ${quotContactPerson || quotClientName}, please find the quotation from AquaServe Pro below.\n\nQuotation No: ${quotationNumber}\nTotal: ₹${quotGrandTotal.toLocaleString("en-IN")}\nValid Until: ${validUntilLabel || "N/A"}\n\n${pdfUrl}\n\nFeel free to reach out with any questions. Thank you!`;
      window.open(`https://wa.me/${cleanPhone}?text=${encodeURIComponent(msg)}`, "_blank");

      // ── Reset the form ──
      setQuotClientName(""); setQuotContactPerson(""); setQuotPhone(""); setQuotEmail(""); setQuotAddress("");
      setQuotItems([makeQuotItemRow()]);
    } catch (err) {
      console.error("Quotation creation error:", err);
      alert("Failed to generate/send the quotation. Please try again.");
    } finally {
      setQuotSubmitting(false);
    }
  };

  // ── Resend a previously generated quotation's existing PDF via
  // WhatsApp, without regenerating it — for a quick follow-up nudge. ──
  const handleResendQuotation = (q) => {
    const cleanPhone = (q.phone || "").toString().replace(/\D/g, "");
    if (cleanPhone.length < 10) { alert("This quotation doesn't have a valid phone number on file."); return; }
    const msg = `Hi ${q.contactPerson || q.clientName}, following up on our quotation from AquaServe Pro.\n\nQuotation No: ${q.quotationNumber}\nTotal: ₹${(q.grandTotal || 0).toLocaleString("en-IN")}\n\n${q.pdfUrl}\n\nLet us know if you have any questions. Thank you!`;
    window.open(`https://wa.me/${cleanPhone}?text=${encodeURIComponent(msg)}`, "_blank");
  };

  const handleDeleteQuotation = async (id) => {
    if (!window.confirm("Delete this quotation record? This won't delete the PDF file itself.")) return;
    try { await deleteDoc(doc(db, "quotations", id)); } catch (e) { console.error(e); }
  };

  const today = new Date().toISOString().split("T")[0];
  const dueCount = customers.filter(c => c.amcExpiry === today && c.lastNotified !== today).length;

  /* ── Shared styles ── */
  const btn = (bg, col, extra = {}) => ({
    display: "inline-flex", alignItems: "center", gap: 7, padding: "9px 18px",
    background: bg, color: col, border: "none", borderRadius: 10, fontSize: 12,
    fontWeight: 700, cursor: "pointer", letterSpacing: "0.03em", transition: "all .15s", ...extra,
  });
  const inp = {
    width: "100%", padding: "10px 14px", background: C.inp, border: `1.5px solid ${C.border}`,
    borderRadius: 10, color: C.text, fontSize: 13, outline: "none", boxSizing: "border-box", transition: "border .15s",
  };
  const card = (extra = {}) => ({
    background: C.card, border: `1px solid ${C.border}`, borderRadius: 20,
    boxShadow: dark ? "0 4px 24px rgba(0,0,0,.4)" : "0 2px 12px rgba(0,0,0,.06)", ...extra,
  });

  /* ── Payment badge ── */
  const PayBadge = ({ s }) => {
    const m = { Paid: ["#d1fae5","#065f46"], Pending: ["#fee2e2","#991b1b"], Partial: ["#ffedd5","#92400e"] }[s] || ["#f1f5f9","#475569"];
    return <span style={{ background: m[0], color: m[1], padding: "3px 11px", borderRadius: 999, fontSize: 10, fontWeight: 800, letterSpacing: "0.06em", textTransform: "uppercase", whiteSpace: "nowrap" }}>{s}</span>;
  };

  const navItems = [
    { key: "Directory",        label: "Directory",        icon: Ic.grid },
    { key: "Staff Management", label: "Staff",            icon: Ic.users },
    { key: "Inventory",        label: "Inventory",        icon: Ic.box },
    { key: "Store Sale",       label: "Store Sale",       icon: Ic.cash },
    { key: "Quotation",        label: "Quotation",        icon: Ic.fileText },
    { key: "Schedule",         label: "Schedule",         icon: Ic.route },
    { key: "Performance",      label: "Performance",      icon: Ic.target },
    { key: "History",          label: "History",          icon: Ic.history },
    { key: "Analysis",         label: "Analysis",         icon: Ic.chart },
  ];

  const tabTitles = {
    "Directory": "Client Directory",
    "Staff Management": "Staff Management",
    "Inventory": "Inventory Management",
    "Store Sale": "Direct Store Sale",
    "Quotation": "Quotation Maker",
    "Schedule": "Job Schedule",
    "Performance": "Staff Performance",
    "History": "Job History",
    "Analysis": "Revenue Analysis",
  };

  /* ════════════ RENDER ════════════ */
  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'Inter',system-ui,sans-serif", transition: "background .25s" }}>

      {/* ══ SIDEBAR ══ */}
      <aside style={{
        width: sidebarOpen ? 230 : 68, flexShrink: 0, background: C.sidebar,
        borderRight: `1px solid ${C.border}`, display: "flex", flexDirection: "column",
        transition: "width .22s ease", overflow: "hidden", position: "sticky", top: 0, height: "100vh",
        boxShadow: dark ? "2px 0 20px rgba(0,0,0,.3)" : "2px 0 12px rgba(0,0,0,.05)",
      }}>
        {/* Logo */}
        <div style={{ padding: sidebarOpen ? "24px 20px 20px" : "24px 14px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10, minHeight: 72 }}>
          <div style={{ width: 38, height: 38, borderRadius: 12, background: "linear-gradient(135deg,#3b82f6 0%,#6366f1 100%)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, boxShadow: "0 4px 12px rgba(99,102,241,.4)" }}>
            <Ic.droplet size={18} style={{ color: "#fff" }} />
          </div>
          {sidebarOpen && (
            <div>
              <p style={{ fontWeight: 800, fontSize: 15, letterSpacing: "-0.03em", lineHeight: 1.1 }}>AquaServe</p>
              <p style={{ fontSize: 10, color: C.accent, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>Pro Admin</p>
            </div>
          )}
        </div>

        {/* Nav */}
        <nav style={{ flex: 1, padding: "16px 10px", display: "flex", flexDirection: "column", gap: 4 }}>
          {navItems.map(({ key, label, icon: NavIcon }) => {
            const active = activeTab === key;
            return (
              <button key={key} onClick={() => setActiveTab(key)} style={{
                display: "flex", alignItems: "center", gap: 12, padding: sidebarOpen ? "11px 14px" : "11px 15px",
                borderRadius: 12, border: "none", cursor: "pointer", transition: "all .15s", textAlign: "left",
                background: active ? (dark ? "rgba(59,130,246,.18)" : "#eff6ff") : "transparent",
                color: active ? C.accent : C.sub,
                fontWeight: active ? 700 : 500, fontSize: 13,
                borderLeft: active ? `3px solid ${C.accent}` : "3px solid transparent",
                position: "relative",
              }}>
                <NavIcon size={17} style={{ color: active ? C.accent : C.sub, flexShrink: 0 }} />
                {sidebarOpen && label}
                {/* ── NEW: low-stock count badge, visible from any tab ── */}
                {key === "Inventory" && lowStockAlerts.length > 0 && (
                  <span style={{
                    position: "absolute", top: sidebarOpen ? 8 : 4, right: sidebarOpen ? 10 : 4,
                    minWidth: 16, height: 16, padding: "0 4px", borderRadius: 999,
                    background: C.danger, color: "#fff", fontSize: 9, fontWeight: 800,
                    display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1,
                  }}>
                    {lowStockAlerts.length}
                  </span>
                )}
                {/* ── NEW: duplicate-customer count badge on Directory ── */}
                {key === "Directory" && duplicateClusters.length > 0 && (
                  <span style={{
                    position: "absolute", top: sidebarOpen ? 8 : 4, right: sidebarOpen ? 10 : 4,
                    minWidth: 16, height: 16, padding: "0 4px", borderRadius: 999,
                    background: C.warn, color: "#fff", fontSize: 9, fontWeight: 800,
                    display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1,
                  }}>
                    {duplicateClusters.length}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        {/* Bottom */}
        <div style={{ padding: "16px 10px", borderTop: `1px solid ${C.border}`, display: "flex", flexDirection: "column", gap: 4 }}>
          <button onClick={() => setDark(!dark)} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 14px", borderRadius: 12, border: "none", cursor: "pointer", background: "transparent", color: C.sub, fontSize: 13, fontWeight: 500, transition: "background .15s" }}
            onMouseEnter={e => e.currentTarget.style.background = dark ? "rgba(255,255,255,.05)" : "#f1f5f9"}
            onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
            {dark ? <Ic.sun size={17} /> : <Ic.moon size={17} />}
            {sidebarOpen && (dark ? "Light Mode" : "Dark Mode")}
          </button>
          <button onClick={() => signOut(auth)} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 14px", borderRadius: 12, border: "none", cursor: "pointer", background: "transparent", color: C.sub, fontSize: 13, fontWeight: 500, transition: "all .15s" }}
            onMouseEnter={e => { e.currentTarget.style.background = "#fee2e2"; e.currentTarget.style.color = C.danger; }}
            onMouseLeave={e => { e.currentTarget.style.background = "transparent"; e.currentTarget.style.color = C.sub; }}>
            <Ic.logout size={17} />
            {sidebarOpen && "Sign Out"}
          </button>
          <button onClick={() => setSidebarOpen(!sidebarOpen)} style={{ display: "flex", alignItems: "center", justifyContent: sidebarOpen ? "flex-end" : "center", gap: 6, padding: "8px 14px", borderRadius: 10, border: "none", cursor: "pointer", background: "transparent", color: C.sub, fontSize: 11, fontWeight: 600 }}>
            <span style={{ display: "inline-block", transform: sidebarOpen ? "rotate(180deg)" : "rotate(0deg)", transition: "transform .2s" }}>
              <Ic.chevron size={14} />
            </span>
            {sidebarOpen && "Collapse"}
          </button>
        </div>
      </aside>

      {/* ══ MAIN AREA ══ */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>

        {/* Top bar */}
        <header style={{
          height: 60, background: C.glass, backdropFilter: "blur(16px)",
          borderBottom: `1px solid ${C.border}`, padding: "0 28px",
          display: "flex", alignItems: "center", justifyContent: "space-between",
          position: "sticky", top: 0, zIndex: 30,
        }}>
          <div>
            <p style={{ fontWeight: 800, fontSize: 16, letterSpacing: "-0.02em" }}>
              {tabTitles[activeTab] || activeTab}
            </p>
            <p style={{ fontSize: 11, color: C.sub, marginTop: 1 }}>
              {new Date().toLocaleDateString("en-IN", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}
            </p>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {syncStatus === "syncing" && (
              <span style={{ fontSize: 11, color: C.warn, fontWeight: 700, background: "#fef3c7", padding: "4px 12px", borderRadius: 999, border: "1px solid #fde68a" }}>
                ⟳ Syncing Warranty Data…
              </span>
            )}
            {syncStatus === "done" && (
              <span style={{ fontSize: 11, color: C.success, fontWeight: 700, background: "#d1fae5", padding: "4px 12px", borderRadius: 999, border: "1px solid #a7f3d0" }}>
                ✓ Warranty Data Synced
              </span>
            )}
            {syncDiagnostics && (
              <span
                title={`Ledger: ${syncDiagnostics.ledgerFound ? `${syncDiagnostics.ledgerRows} rows found` : "FILE NOT FOUND at company_data/Warranty_Ledger.xlsx"}\nFlat: ${syncDiagnostics.flatFound ? `${syncDiagnostics.flatRows} customer groups found` : "FILE NOT FOUND at company_data/Warranty_Flat.xlsx"}\nMerged customers written: ${syncDiagnostics.mergedCount}${syncDiagnostics.error ? `\nError: ${syncDiagnostics.error}` : ""}`}
                style={{
                  fontSize: 11, fontWeight: 700, padding: "4px 12px", borderRadius: 999, cursor: "help",
                  background: (!syncDiagnostics.ledgerFound || !syncDiagnostics.flatFound || syncDiagnostics.error) ? "#fee2e2" : (dark ? "rgba(59,130,246,.15)" : "#eff6ff"),
                  color: (!syncDiagnostics.ledgerFound || !syncDiagnostics.flatFound || syncDiagnostics.error) ? C.danger : C.accent,
                  border: `1px solid ${(!syncDiagnostics.ledgerFound || !syncDiagnostics.flatFound || syncDiagnostics.error) ? "#fecaca" : C.borderHi}`,
                }}
              >
                {(!syncDiagnostics.ledgerFound || !syncDiagnostics.flatFound) ? "⚠ " : ""}
                Ledger: {syncDiagnostics.ledgerFound ? syncDiagnostics.ledgerRows : "not found"} · Flat: {syncDiagnostics.flatFound ? syncDiagnostics.flatRows : "not found"}
              </span>
            )}
            {dueCount > 0 && (
              <button onClick={sendAutoReminders} style={{ ...btn("#fef2f2", C.danger, { border: `1.5px solid #fecaca`, fontSize: 11 }) }}>
                <Ic.bell size={13} />
                {dueCount} Due Today — Notify All
              </button>
            )}
            <div style={{ width: 36, height: 36, borderRadius: "50%", background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 14, cursor: "default" }}>A</div>
          </div>
        </header>

        {/* ── NEW: soft banner for Excel-sync failures (non-blocking, dismissible) ── */}
        {excelSyncWarning && (
          <div style={{
            margin: "14px 28px 0", padding: "10px 16px", borderRadius: 12,
            background: dark ? "rgba(245,158,11,.12)" : "#fffbeb",
            border: `1px solid ${dark ? "rgba(245,158,11,.35)" : "#fde68a"}`,
            color: C.warn, fontSize: 12.5, fontWeight: 600,
            display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12,
          }}>
            <span>⚠ {excelSyncWarning}</span>
            <button onClick={() => setExcelSyncWarning("")} style={{ background: "transparent", border: "none", color: C.warn, cursor: "pointer", fontSize: 16, lineHeight: 1, flexShrink: 0 }}>×</button>
          </div>
        )}

        {/* ── NEW: global low-stock banner — visible on every tab, not
             only when the Inventory tab happens to be open. Dismissing
             it hides it only until the specific set of low-stock items
             changes, so a new item dropping low re-surfaces it. ── */}
        {showLowStockBanner && (
          <div style={{
            margin: "14px 28px 0", padding: "10px 16px", borderRadius: 12,
            background: dark ? "rgba(239,68,68,.12)" : "#fef2f2",
            border: `1px solid ${dark ? "rgba(239,68,68,.35)" : "#fecaca"}`,
            color: C.danger, fontSize: 12.5, fontWeight: 600,
            display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap",
          }}>
            <span>
              ⚠ {lowStockAlerts.length} item{lowStockAlerts.length === 1 ? " is" : "s are"} low on stock:{" "}
              {lowStockAlerts.slice(0, 4).map(i => `${i.itemName} (${i.quantity} left)`).join(", ")}
              {lowStockAlerts.length > 4 ? `, +${lowStockAlerts.length - 4} more` : ""}
            </span>
            <button onClick={() => setLowStockDismissedKey(lowStockKey)} style={{ background: "transparent", border: "none", color: C.danger, cursor: "pointer", fontSize: 16, lineHeight: 1, flexShrink: 0 }}>×</button>
          </div>
        )}

        {/* ── NEW: global duplicate-customer banner — the permanent,
             always-visible fix for the "Total Clients silently grew"
             issue. Dismissing it only hides it until the cluster count
             changes, so a fresh duplicate re-surfaces it. ── */}
        {showDuplicatesBanner && (
          <div style={{
            margin: "14px 28px 0", padding: "10px 16px", borderRadius: 12,
            background: dark ? "rgba(245,158,11,.12)" : "#fffbeb",
            border: `1px solid ${dark ? "rgba(245,158,11,.35)" : "#fde68a"}`,
            color: C.warn, fontSize: 12.5, fontWeight: 600,
            display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap",
          }}>
            <span>
              ⚠ {duplicateClusters.length} possible duplicate customer record{duplicateClusters.length === 1 ? "" : "s"} found (same phone or very similar name).
            </span>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
              <button onClick={() => setShowDuplicatesModal(true)} style={{ ...btn(C.warn, "#fff", { fontSize: 11, padding: "6px 14px" }) }}>
                Review & Merge
              </button>
              <button onClick={() => setDupBannerDismissedCount(duplicateClusters.length)} style={{ background: "transparent", border: "none", color: C.warn, cursor: "pointer", fontSize: 16, lineHeight: 1 }}>×</button>
            </div>
          </div>
        )}

        <main style={{ flex: 1, padding: "24px 28px", overflowY: "auto" }}>

          {/* ── STAT CARDS ── */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 16, marginBottom: 28 }}>
            {[
              { label: "Total Clients",    val: customers.length,                                              grad: ["#3b82f6","#6366f1"], icon: Ic.users },
              { label: "Pending Services", val: customers.filter(c => c.warrantyPending > 0).length,               grad: ["#ef4444","#f97316"], icon: Ic.clock },
              { label: "Staff Members",    val: staff.length,                                                  grad: ["#10b981","#06b6d4"], icon: Ic.users },
              { label: "Due Today",        val: dueCount,                                                      grad: dueCount > 0 ? ["#f59e0b","#ef4444"] : ["#8b5cf6","#ec4899"], icon: Ic.bell },
            ].map((s, i) => (
              <div key={i} style={{ ...card(), padding: "20px 22px", position: "relative", overflow: "hidden", cursor: "default" }}>
                <div style={{ position: "absolute", top: -10, right: -10, width: 70, height: 70, borderRadius: "50%", background: `linear-gradient(135deg,${s.grad[0]}22,${s.grad[1]}22)` }} />
                <div style={{ width: 40, height: 40, borderRadius: 12, background: `linear-gradient(135deg,${s.grad[0]},${s.grad[1]})`, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 14, boxShadow: `0 4px 12px ${s.grad[0]}55` }}>
                  <s.icon size={18} style={{ color: "#fff" }} />
                </div>
                <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>{s.label}</p>
                <p style={{ fontSize: 34, fontWeight: 900, lineHeight: 1, background: `linear-gradient(135deg,${s.grad[0]},${s.grad[1]})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>{s.val}</p>
                {s.label === "Due Today" && dueCount > 0 && <span style={{ position: "absolute", top: 18, right: 18, width: 9, height: 9, borderRadius: "50%", background: C.danger, boxShadow: "0 0 0 3px rgba(239,68,68,.25)", animation: "pulseRing 1.4s ease infinite" }} />}
              </div>
            ))}
          </div>

          {/* ══ DIRECTORY ══ */}
          {activeTab === "Directory" && (
            <div style={{ display: "grid", gridTemplateColumns: "300px 1fr", gap: 20, alignItems: "start" }}>

              {/* Add Form */}
              <div style={{ ...card(), padding: 24, position: "sticky", top: 84 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20, paddingBottom: 16, borderBottom: `1px solid ${C.border}` }}>
                  <div style={{ width: 32, height: 32, borderRadius: 10, background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <Ic.plus size={15} style={{ color: "#fff" }} />
                  </div>
                  <p style={{ fontWeight: 800, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.05em" }}>Add Client</p>
                </div>

                {/* Warranty / Service tab switcher */}
                <div style={{ display: "flex", gap: 6, marginBottom: 16, padding: 4, background: dark ? "#0a1525" : "#f1f5f9", borderRadius: 12 }}>
                  {[{ k: "warranty", label: "Warranty Client" }, { k: "service", label: "Service Client" }].map(t => (
                    <button key={t.k} type="button" onClick={() => setAddFormMode(t.k)}
                      style={{
                        flex: 1, padding: "9px 6px", borderRadius: 9, border: "none", cursor: "pointer",
                        fontSize: 11.5, fontWeight: 700, letterSpacing: "0.02em",
                        background: addFormMode === t.k ? "linear-gradient(135deg,#3b82f6,#6366f1)" : "transparent",
                        color: addFormMode === t.k ? "#fff" : C.sub,
                        boxShadow: addFormMode === t.k ? "0 3px 10px rgba(99,102,241,.35)" : "none",
                        transition: "all .15s",
                      }}>
                      {t.label}
                    </button>
                  ))}
                </div>

                {/* Warranty Client form: new install → Ledger-shaped record */}
                {addFormMode === "warranty" && (
                <form onSubmit={handleAddWarrantyClient} style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {[{ ph: "Full name *", k: "name" }, { ph: "Phone number *", k: "phone" }, { ph: "Full address", k: "address" }, { ph: "Machine model", k: "machineModel" }].map(f => (
                    <input key={f.k} placeholder={f.ph} style={inp} value={form[f.k]} onChange={e => setForm({ ...form, [f.k]: e.target.value })}
                      onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  ))}
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <input placeholder="Nagar" style={inp} value={form.nagar} onChange={e => setForm({ ...form, nagar: e.target.value })} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                    <input placeholder="Place" style={inp} value={form.place} onChange={e => setForm({ ...form, place: e.target.value })} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  </div>
                  <input placeholder="Price (₹)" type="number" style={inp} value={form.price} onChange={e => setForm({ ...form, price: e.target.value })} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  <div>
                    <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 5 }}>Installation Date</p>
                    <input type="date" style={inp} value={form.installDate} onChange={e => setForm({ ...form, installDate: e.target.value })} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  </div>
                  <select style={{ ...inp, fontWeight: 600, cursor: "pointer" }} value={form.paymentStatus} onChange={e => setForm({ ...form, paymentStatus: e.target.value })}>
                    <option value="Pending">Pending</option>
                    <option value="Paid">Paid</option>
                    <option value="Partial">Partial</option>
                  </select>
                  {warrantyFormStatus === "error" || warrantyFormStatus.startsWith("error:") ? (
                    <p style={{ fontSize: 11.5, color: C.danger, fontWeight: 600, background: dark ? "rgba(239,68,68,.1)" : "#fef2f2", padding: "8px 10px", borderRadius: 8, border: `1px solid ${dark ? "rgba(239,68,68,.25)" : "#fecaca"}` }}>
                      {warrantyFormStatus.split("error:")[1] || "Something went wrong. Please try again."}
                    </p>
                  ) : null}
                  {warrantyFormStatus === "done" && (
                    <p style={{ fontSize: 11.5, color: C.success, fontWeight: 600, background: dark ? "rgba(16,185,129,.1)" : "#f0fdf4", padding: "8px 10px", borderRadius: 8, border: `1px solid ${dark ? "rgba(16,185,129,.25)" : "#bbf7d0"}` }}>
                      ✓ Client saved successfully.
                    </p>
                  )}
                  <button type="submit" disabled={warrantyFormStatus === "saving"} style={{ marginTop: 8, padding: "13px", background: "linear-gradient(135deg,#3b82f6,#6366f1)", color: "#fff", border: "none", borderRadius: 12, fontWeight: 800, fontSize: 13, cursor: warrantyFormStatus === "saving" ? "wait" : "pointer", letterSpacing: "0.03em", boxShadow: "0 6px 18px rgba(99,102,241,.4)", transition: "opacity .15s", opacity: warrantyFormStatus === "saving" ? 0.7 : 1 }}
                    onMouseEnter={e => e.currentTarget.style.opacity = ".88"} onMouseLeave={e => e.currentTarget.style.opacity = warrantyFormStatus === "saving" ? "0.7" : "1"}>
                    {warrantyFormStatus === "saving" ? "Saving…" : "Save Client"}
                  </button>
                </form>
                )}

                {/* Service Client form: logs one visit → Flat-shaped record */}
                {addFormMode === "service" && (
                <form onSubmit={handleAddServiceClient} style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  <input placeholder="Customer name *" style={inp} value={serviceForm.customerName} onChange={e => setServiceForm({ ...serviceForm, customerName: e.target.value })}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  <input placeholder="Phone number *" style={inp} value={serviceForm.phone} onChange={e => setServiceForm({ ...serviceForm, phone: e.target.value })}
                    onBlur={handleServicePhoneBlur} onFocus={e => e.target.style.borderColor = C.accent} />
                  <input placeholder="Address" style={inp} value={serviceForm.address} onChange={e => setServiceForm({ ...serviceForm, address: e.target.value })}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  <div>
                    <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 5 }}>Service Date</p>
                    <input type="date" style={inp} value={serviceForm.serviceDate} onChange={e => setServiceForm({ ...serviceForm, serviceDate: e.target.value })}
                      onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  </div>
                  <textarea placeholder="Service description *" rows={3} style={{ ...inp, resize: "vertical", fontFamily: "inherit" }} value={serviceForm.serviceDescription}
                    onChange={e => setServiceForm({ ...serviceForm, serviceDescription: e.target.value })}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  <input placeholder="Amount (₹)" type="number" style={inp} value={serviceForm.amount} onChange={e => setServiceForm({ ...serviceForm, amount: e.target.value })}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  {serviceFormStatus.startsWith("error:") ? (
                    <p style={{ fontSize: 11.5, color: C.danger, fontWeight: 600, background: dark ? "rgba(239,68,68,.1)" : "#fef2f2", padding: "8px 10px", borderRadius: 8, border: `1px solid ${dark ? "rgba(239,68,68,.25)" : "#fecaca"}` }}>
                      {serviceFormStatus.split("error:")[1]}
                    </p>
                  ) : null}
                  {serviceFormStatus === "done" && (
                    <p style={{ fontSize: 11.5, color: C.success, fontWeight: 600, background: dark ? "rgba(16,185,129,.1)" : "#f0fdf4", padding: "8px 10px", borderRadius: 8, border: `1px solid ${dark ? "rgba(16,185,129,.25)" : "#bbf7d0"}` }}>
                      ✓ Service visit logged successfully.
                    </p>
                  )}
                  <button type="submit" disabled={serviceFormStatus === "saving"} style={{ marginTop: 8, padding: "13px", background: "linear-gradient(135deg,#10b981,#3b82f6)", color: "#fff", border: "none", borderRadius: 12, fontWeight: 800, fontSize: 13, cursor: serviceFormStatus === "saving" ? "wait" : "pointer", letterSpacing: "0.03em", boxShadow: "0 6px 18px rgba(16,185,129,.35)", transition: "opacity .15s", opacity: serviceFormStatus === "saving" ? 0.7 : 1 }}
                    onMouseEnter={e => e.currentTarget.style.opacity = ".88"} onMouseLeave={e => e.currentTarget.style.opacity = serviceFormStatus === "saving" ? "0.7" : "1"}>
                    {serviceFormStatus === "saving" ? "Saving…" : "Log Service Visit"}
                  </button>
                </form>
                )}
              </div>

              {/* Table Panel */}
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>

                {/* Toolbar */}
                <div style={{ ...card({ borderRadius: 16 }), padding: "14px 18px", display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                  <div style={{ position: "relative", flex: "1 1 200px" }}>
                    <span style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: C.sub, pointerEvents: "none" }}>
                      <Ic.search size={14} />
                    </span>
                    <input placeholder="Search name, phone, nagar…" style={{ ...inp, paddingLeft: 36 }}
                      onChange={e => setSearchTerm(e.target.value)}
                      onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  </div>
                  <div style={{ position: "relative" }}>
                    <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: C.sub, pointerEvents: "none" }}>
                      <Ic.filter size={13} />
                    </span>
                    <select style={{ ...inp, paddingLeft: 30, width: "auto", cursor: "pointer" }} value={filterStatus} onChange={e => setFilterStatus(e.target.value)}>
                      <option value="All">All Clients</option>
                      <option value="PendingService">Pending Service</option>
                      <option value="Warranty">Warranty (Ledger)</option>
                      <option value="NonWarranty">Non-warranty (Flat only)</option>
                      <option value="Alphabetical">A – Z</option>
                    </select>
                  </div>
                  {duplicateClusters.length > 0 && (
                    <button onClick={() => setShowDuplicatesModal(true)} style={{ ...btn("#fef3c7", "#92400e", { border: `1.5px solid #fde68a`, flexShrink: 0 }) }}>
                      <Ic.users size={13} /> {duplicateClusters.length} Duplicate{duplicateClusters.length === 1 ? "" : "s"}
                    </button>
                  )}
                  <button onClick={exportToExcel} style={{ ...btn("#10b981", "#fff", { boxShadow: "0 4px 12px rgba(16,185,129,.35)", flexShrink: 0 }) }}>
                    <Ic.download size={13} /> Export Excel
                  </button>
                </div>

                <p style={{ fontSize: 12, color: C.sub, paddingLeft: 4 }}>
                  Showing <strong style={{ color: C.text }}>{processedList.length}</strong> of {customers.length} clients
                </p>

                {/* Table */}
                <div style={{ ...card(), overflow: "hidden" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ background: dark ? "#0c1526" : "#f8fafc", borderBottom: `1px solid ${C.border}` }}>
                        {["Client", "Machine", "Installed", "Last Service", "Warranty", ""].map((h, i) => (
                          <th key={i} style={{ padding: "13px 18px", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.09em", color: C.sub, textAlign: i >= 2 ? "center" : "left", ...(i === 5 ? { textAlign: "right" } : {}) }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {processedList.map(c => {
                        const pending = c.warrantyPending;
                        const hasWarrantyData = c.warrantyServicesCompleted !== undefined;
                        return (
                          <tr key={c.id} style={{ borderBottom: `1px solid ${C.border}`, transition: "background .12s" }}
                            onMouseEnter={e => e.currentTarget.style.background = C.cardHov}
                            onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                            <td style={{ padding: "15px 18px", position: "relative" }}>
                              <div style={{ display: "flex", alignItems: "center", gap: 10, cursor: "default" }}
                                onMouseEnter={e => {
                                  const r = e.currentTarget.getBoundingClientRect();
                                  openHistoryPreview(c.id, { x: r.left, y: r.bottom + 6 });
                                }}
                                onMouseLeave={scheduleCloseHistoryPreview}>
                                <div style={{ width: 34, height: 34, borderRadius: "50%", background: `linear-gradient(135deg,#3b82f6,#6366f1)`, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 13, flexShrink: 0 }}>
                                  {(c.name || "?")[0].toUpperCase()}
                                </div>
                                <div>
                                  <p style={{ fontWeight: 700, marginBottom: 1 }}>{c.name}</p>
                                  <p style={{ fontSize: 11, color: C.sub }}>{c.phone}</p>
                                </div>
                              </div>

                              {historyPreviewId === c.id && (c.allServiceHistory?.length > 0) && (
                                <div
                                  onMouseEnter={() => openHistoryPreview(c.id)}
                                  onMouseLeave={scheduleCloseHistoryPreview}
                                  style={{
                                    position: "fixed", left: historyPreviewPos.x, top: historyPreviewPos.y, zIndex: 200,
                                    width: 300, maxHeight: 320, overflowY: "auto", overscrollBehavior: "contain", background: C.card,
                                    border: `1px solid ${C.borderHi}`, borderRadius: 14,
                                    boxShadow: dark ? "0 12px 32px rgba(0,0,0,.5)" : "0 12px 32px rgba(0,0,0,.18)", padding: 14,
                                  }}>
                                  <p style={{ fontSize: 10, fontWeight: 800, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 10 }}>
                                    Full Service History
                                  </p>
                                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                                    {c.allServiceHistory.map((s, i) => (
                                      <div key={i} style={{ padding: "8px 10px", background: dark ? "#0a1525" : "#f8fafc", borderRadius: 10, border: `1px solid ${C.border}` }}>
                                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 3 }}>
                                          <span style={{ fontSize: 11, fontWeight: 700, color: C.text }}>{s.dateISO || "Date unknown"}</span>
                                          <span style={{ fontSize: 9, fontWeight: 700, color: C.accent, textTransform: "uppercase" }}>{s.source}</span>
                                        </div>
                                        <p style={{ fontSize: 12, color: C.sub }}>{s.description}{s.amount ? ` · ₹${s.amount}` : ""}</p>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              )}
                            </td>
                            <td style={{ padding: "15px 18px" }}>
                              <p style={{ fontSize: 12, color: C.sub, fontWeight: 500 }}>{c.machineModel || "—"}</p>
                            </td>
                            <td style={{ padding: "15px 18px", textAlign: "center" }}>
                              <span style={{ fontSize: 11, fontWeight: 600, background: dark ? "#1e293b" : "#f1f5f9", color: C.sub, padding: "4px 10px", borderRadius: 8 }}>{c.installDate || "—"}</span>
                            </td>
                            <td style={{ padding: "15px 18px", textAlign: "center" }}>
                              <span style={{ fontSize: 11, fontWeight: 600, background: dark ? "#1e293b" : "#f1f5f9", color: C.sub, padding: "4px 10px", borderRadius: 8 }}>{c.lastServiceDate || "No service yet"}</span>
                            </td>
                            <td style={{ padding: "15px 18px", textAlign: "center" }}>
                              {hasWarrantyData ? (
                                <span style={{ fontSize: 11, fontWeight: 700, background: pending > 0 ? (dark ? "rgba(245,158,11,.15)" : "#fef3c7") : "#d1fae5", color: pending > 0 ? C.warn : "#065f46", padding: "4px 10px", borderRadius: 8 }}>
                                  {pending > 0 ? `${pending} pending` : "Fully used"}
                                </span>
                              ) : (
                                <span style={{ fontSize: 11, color: C.sub }}>—</span>
                              )}
                            </td>
                            <td style={{ padding: "15px 18px" }}>
                              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                                <button onClick={() => setSelectedCustomer(c)} style={{ ...btn("linear-gradient(135deg,#3b82f6,#6366f1)", "#fff", { fontSize: 11, boxShadow: "0 3px 10px rgba(99,102,241,.35)" }) }}
                                  onMouseEnter={e => e.currentTarget.style.opacity = ".85"} onMouseLeave={e => e.currentTarget.style.opacity = "1"}>
                                  Assign
                                </button>
                                <button
                                  onClick={() => handleWhatsAppRemind(c)}
                                  title={c.phone ? `Remind ${c.name} via WhatsApp` : "No phone number on file"}
                                  style={{ ...btn("#25D366", "#fff", { fontSize: 11, boxShadow: "0 3px 10px rgba(37,211,102,.35)", display: "flex", alignItems: "center", gap: 5 }) }}
                                  onMouseEnter={e => e.currentTarget.style.opacity = ".85"} onMouseLeave={e => e.currentTarget.style.opacity = "1"}>
                                  <Ic.wa size={12} /> Remind
                                </button>
                                <button onClick={() => handleDelete(c.id)} style={{ ...btn("transparent", C.danger, { border: `1.5px solid #fecaca`, fontSize: 11 }) }}
                                  onMouseEnter={e => e.currentTarget.style.background = "#fee2e2"} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                                  <Ic.trash size={12} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                      {!processedList.length && (
                        <tr><td colSpan={6} style={{ padding: "56px 20px", textAlign: "center", color: C.sub }}>
                          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 12 }}>
                            <Ic.search size={32} style={{ color: C.border }} />
                            <p style={{ fontSize: 14, fontWeight: 600 }}>No clients match your search</p>
                            <p style={{ fontSize: 12 }}>Try adjusting the filter or search term</p>
                          </div>
                        </td></tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ══ STAFF MANAGEMENT ══ */}
          {activeTab === "Staff Management" && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(320px,1fr))", gap: 20 }}>
              {staff.map(m => (
                <div key={m.id} style={{ ...card(), padding: 24 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20, paddingBottom: 16, borderBottom: `1px solid ${C.border}` }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                      <div style={{ position: "relative" }}>
                        <div style={{ width: 50, height: 50, borderRadius: "50%", background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 18, boxShadow: "0 4px 14px rgba(99,102,241,.4)" }}>
                          {m.name[0]}
                        </div>
                        <span style={{ position: "absolute", bottom: 1, right: 1, width: 12, height: 12, borderRadius: "50%", background: m.status === "Online" ? C.success : C.sub, border: `2px solid ${C.card}` }} />
                      </div>
                      <div>
                        <p style={{ fontWeight: 800, fontSize: 15, marginBottom: 2 }}>{m.name}</p>
                        <p style={{ fontSize: 12, color: C.sub }}>{m.phone}</p>
                      </div>
                    </div>
                    <span style={{ padding: "4px 12px", borderRadius: 999, fontSize: 11, fontWeight: 700, background: m.status === "Online" ? "#d1fae5" : "#f1f5f9", color: m.status === "Online" ? "#065f46" : C.sub }}>
                      {m.status}
                    </span>
                  </div>
                  <div>
                    <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 10 }}>
                      {m.currentCustomers?.length > 0 ? `${m.currentCustomers.length} Active Assignment${m.currentCustomers.length > 1 ? "s" : ""}` : "No Assignments"}
                    </p>
                    {m.currentCustomers?.length > 0 ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        {m.currentCustomers.map((c, i) => (
                          <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", background: dark ? "rgba(59,130,246,.1)" : "#eff6ff", borderRadius: 12, border: `1px solid ${C.borderHi}` }}>
                            <span style={{ width: 22, height: 22, borderRadius: "50%", background: C.accent, color: "#fff", fontSize: 10, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{i + 1}</span>
                            <p style={{ fontSize: 13, fontWeight: 600, color: dark ? "#93c5fd" : "#1d4ed8" }}>{c.name}</p>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div style={{ background: dark ? "#0a1525" : "#f8fafc", borderRadius: 12, padding: "16px", textAlign: "center", color: C.sub, fontSize: 12, border: `1px dashed ${C.border}` }}>
                        Available for assignment
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* ══ INVENTORY MANAGEMENT ══ */}
          {activeTab === "Inventory" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>

              {/* Toolbar */}
              <div style={{ ...card({ borderRadius: 16 }), padding: "14px 18px", display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                <div style={{ position: "relative", flex: "1 1 200px" }}>
                  <span style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: C.sub, pointerEvents: "none" }}>
                    <Ic.search size={14} />
                  </span>
                  <input placeholder="Search item name…" style={{ ...inp, paddingLeft: 36 }}
                    value={inventorySearch}
                    onChange={e => setInventorySearch(e.target.value)}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                </div>

                {inventoryImportStatus === "importing" && (
                  <span style={{ fontSize: 11, color: C.warn, fontWeight: 700, background: "#fef3c7", padding: "4px 12px", borderRadius: 999, border: "1px solid #fde68a" }}>
                    ⟳ Importing…
                  </span>
                )}
                {inventoryImportStatus?.startsWith("done:") && (
                  <span style={{ fontSize: 11, color: C.success, fontWeight: 700, background: "#d1fae5", padding: "4px 12px", borderRadius: 999, border: "1px solid #a7f3d0" }}>
                    ✓ Imported {inventoryImportStatus.split(":")[1]} items
                  </span>
                )}
                {inventoryImportStatus === "error" && (
                  <span style={{ fontSize: 11, color: C.danger, fontWeight: 700, background: "#fee2e2", padding: "4px 12px", borderRadius: 999, border: "1px solid #fecaca" }}>
                    ✕ Import failed
                  </span>
                )}

                <input type="file" accept=".xlsx,.xls,.csv" ref={inventoryFileRef} onChange={handleInventoryImport} style={{ display: "none" }} />
                <button onClick={exportInventoryToExcel} style={{ ...btn("#3b82f6", "#fff", { boxShadow: "0 4px 12px rgba(59,130,246,.35)", flexShrink: 0 }) }}>
                  <Ic.download size={13} /> Export Excel
                </button>
                <button onClick={() => inventoryFileRef.current && inventoryFileRef.current.click()} style={{ ...btn("#10b981", "#fff", { boxShadow: "0 4px 12px rgba(16,185,129,.35)", flexShrink: 0 }) }}>
                  <Ic.upload size={13} /> Import Excel Stock
                </button>
              </div>

              {/* ── NEW: Current Stock / Movement History sub-view toggle ── */}
              <div style={{ display: "flex", gap: 6, padding: 4, background: dark ? "#0a1525" : "#f1f5f9", borderRadius: 12, width: "fit-content" }}>
                {[{ k: "stock", label: "Current Stock" }, { k: "history", label: "Movement History" }].map(v => (
                  <button key={v.k} onClick={() => setInventorySubView(v.k)} style={{
                    padding: "9px 18px", borderRadius: 9, border: "none", cursor: "pointer",
                    fontSize: 12, fontWeight: 700,
                    background: inventorySubView === v.k ? "linear-gradient(135deg,#3b82f6,#6366f1)" : "transparent",
                    color: inventorySubView === v.k ? "#fff" : C.sub,
                    boxShadow: inventorySubView === v.k ? "0 3px 10px rgba(99,102,241,.35)" : "none",
                    transition: "all .15s",
                  }}>
                    {v.label}
                  </button>
                ))}
              </div>

              {inventorySubView === "stock" ? (
                <>
              <p style={{ fontSize: 12, color: C.sub, paddingLeft: 4 }}>
                Showing <strong style={{ color: C.text }}>{filteredInventory.length}</strong> of {inventory.length} items · Excel columns expected: <strong style={{ color: C.text }}>S.No, Item Name, Quantity</strong>
              </p>

              {/* Table */}
              <div style={{ ...card(), overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: dark ? "#0c1526" : "#f8fafc", borderBottom: `1px solid ${C.border}` }}>
                      {["S.No", "Item Name", "Quantity", "Status", ""].map((h, i) => (
                        <th key={i} style={{ padding: "13px 18px", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.09em", color: C.sub, textAlign: i === 0 ? "left" : i === 4 ? "right" : "center" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredInventory.map(item => {
                      const low = Number(item.quantity) <= 5;
                      const displayQty = qtyOverrides[item.id] !== undefined ? qtyOverrides[item.id] : item.quantity;
                      return (
                        <tr key={item.id} style={{ borderBottom: `1px solid ${C.border}`, transition: "background .12s" }}
                          onMouseEnter={e => e.currentTarget.style.background = C.cardHov}
                          onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                          <td style={{ padding: "15px 18px" }}>
                            <span style={{ fontSize: 12, color: C.sub, fontWeight: 600 }}>{item.sno || "—"}</span>
                          </td>
                          <td style={{ padding: "15px 18px" }}>
                            <p style={{ fontWeight: 700 }}>{item.itemName}</p>
                          </td>
                          <td style={{ padding: "15px 18px", textAlign: "center" }}>
                            <input
                              type="number"
                              value={displayQty}
                              onChange={e => handleInventoryQtyInputChange(item, e.target.value)}
                              onBlur={e => handleInventoryQtyBlur(item, e.target.value)}
                              style={{ width: 80, padding: "6px 10px", textAlign: "center", background: C.inp, border: `1.5px solid ${C.border}`, borderRadius: 8, color: C.text, fontSize: 13, fontWeight: 700, outline: "none" }}
                            />
                          </td>
                          <td style={{ padding: "15px 18px", textAlign: "center" }}>
                            <span style={{ fontSize: 11, fontWeight: 700, background: low ? "#fee2e2" : "#d1fae5", color: low ? C.danger : "#065f46", padding: "4px 10px", borderRadius: 8 }}>
                              {low ? "Low Stock" : "In Stock"}
                            </span>
                          </td>
                          <td style={{ padding: "15px 18px" }}>
                            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                              <button onClick={() => handleInventoryDelete(item.id)} style={{ ...btn("transparent", C.danger, { border: `1.5px solid #fecaca`, fontSize: 11 }) }}
                                onMouseEnter={e => e.currentTarget.style.background = "#fee2e2"} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                                <Ic.trash size={12} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {!filteredInventory.length && (
                      <tr><td colSpan={5} style={{ padding: "56px 20px", textAlign: "center", color: C.sub }}>
                        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 12 }}>
                          <Ic.box size={32} style={{ color: C.border }} />
                          <p style={{ fontSize: 14, fontWeight: 600 }}>No inventory items yet</p>
                          <p style={{ fontSize: 12 }}>Import your Excel stock file to get started</p>
                        </div>
                      </td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              </>
              ) : (
                <>
                  {/* ── NEW: Movement History table — reads the stockMovements
                       collection populated by Excel imports, manual edits,
                       Store Sales, and Service Jobs. ── */}
                  <div style={{ ...card({ borderRadius: 16 }), padding: "14px 18px" }}>
                    <div style={{ position: "relative" }}>
                      <span style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: C.sub, pointerEvents: "none" }}>
                        <Ic.search size={14} />
                      </span>
                      <input placeholder="Search item name or reason…" style={{ ...inp, paddingLeft: 36 }}
                        value={movementSearch}
                        onChange={e => setMovementSearch(e.target.value)}
                        onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                    </div>
                  </div>

                  <div style={{ ...card(), overflow: "hidden" }}>
                    {movementHistoryLoading ? (
                      <div style={{ padding: "40px 20px", textAlign: "center", color: C.sub, fontSize: 13 }}>Loading movement history…</div>
                    ) : (() => {
                      const q = movementSearch.toLowerCase();
                      const filtered = movementHistory.filter(m =>
                        (m.itemName || "").toLowerCase().includes(q) || (m.reason || "").toLowerCase().includes(q)
                      );
                      if (!filtered.length) {
                        return (
                          <div style={{ padding: "56px 20px", textAlign: "center", color: C.sub }}>
                            <Ic.history size={32} style={{ color: C.border, marginBottom: 12 }} />
                            <p style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>No stock movements yet</p>
                            <p style={{ fontSize: 12 }}>Imports, manual edits, store sales, and service jobs will all show up here.</p>
                          </div>
                        );
                      }
                      return (
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                          <thead>
                            <tr style={{ background: dark ? "#0c1526" : "#f8fafc", borderBottom: `1px solid ${C.border}` }}>
                              {["Date", "Item", "Change", "New Qty", "Reason", "By"].map((h, i) => (
                                <th key={i} style={{ padding: "13px 18px", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.09em", color: C.sub, textAlign: i === 2 || i === 3 ? "right" : "left" }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {filtered.map(m => (
                              <tr key={m.id} style={{ borderBottom: `1px solid ${C.border}`, transition: "background .12s" }}
                                onMouseEnter={e => e.currentTarget.style.background = C.cardHov}
                                onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                                <td style={{ padding: "13px 18px", color: C.sub, whiteSpace: "nowrap" }}>{m.date || "—"}</td>
                                <td style={{ padding: "13px 18px", fontWeight: 700 }}>{m.itemName || "—"}</td>
                                <td style={{ padding: "13px 18px", textAlign: "right", fontWeight: 800, color: (m.change || 0) < 0 ? C.danger : C.success }}>
                                  {(m.change || 0) > 0 ? `+${m.change}` : m.change}
                                </td>
                                <td style={{ padding: "13px 18px", textAlign: "right", color: C.sub }}>{m.newQuantity ?? "—"}</td>
                                <td style={{ padding: "13px 18px", color: C.sub }}>{m.reason || "—"}</td>
                                <td style={{ padding: "13px 18px", color: C.sub, fontSize: 12 }}>{m.by || "—"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      );
                    })()}
                  </div>
                </>
              )}
            </div>
          )}

          {/* ══ DIRECT STORE SALE (its own dashboard page) ══ */}
          {activeTab === "Store Sale" && (
            <div style={{ ...card(), padding: 24 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20, paddingBottom: 16, borderBottom: `1px solid ${C.border}` }}>
                <div style={{ width: 32, height: 32, borderRadius: 10, background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <Ic.cash size={15} style={{ color: "#fff" }} />
                </div>
                <div>
                  <p style={{ fontWeight: 800, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.05em" }}>Direct Store Sale</p>
                  <p style={{ fontSize: 11, color: C.sub, marginTop: 2 }}>For walk-in customers buying items directly from the store</p>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
                <div>
                  <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 5 }}>Customer Name</p>
                  <input placeholder="Full name" style={inp} value={storeCustomerName} onChange={e => setStoreCustomerName(e.target.value)}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                </div>
                <div>
                  <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 5 }}>Customer Phone (for WhatsApp bill)</p>
                  <input placeholder="10-digit phone number" style={inp} value={storeCustomerPhone} onChange={e => setStoreCustomerPhone(e.target.value)}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                </div>
              </div>

              <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>Items Sold</p>
              <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 14 }}>
                {storeItems.map((row, index) => (
                  <div key={row.id} style={{ background: C.inp, border: `1px solid ${C.border}`, borderRadius: 14, padding: 14, display: "flex", flexDirection: "column", gap: 10 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 10, fontWeight: 800, color: C.sub, textTransform: "uppercase", letterSpacing: "0.06em" }}>Item #{index + 1}</span>
                      {storeItems.length > 1 && (
                        <button onClick={() => removeStoreItemRow(row.id)} style={{ background: "transparent", border: "none", color: C.danger, cursor: "pointer", display: "flex" }}>
                          <Ic.trash size={14} />
                        </button>
                      )}
                    </div>
                    <input list="admin-inventory-item-list" placeholder="Item name…" value={row.item}
                      onChange={e => updateStoreItem(row.id, "item", e.target.value)}
                      style={inp} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                      <input type="number" placeholder="Quantity" value={row.quantity}
                        onChange={e => updateStoreItem(row.id, "quantity", e.target.value)}
                        style={inp} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                      <input type="number" placeholder="Price" value={row.price}
                        onChange={e => updateStoreItem(row.id, "price", e.target.value)}
                        style={inp} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                    </div>
                  </div>
                ))}
                <datalist id="admin-inventory-item-list">
                  {inventory.map(inv => <option key={inv.id} value={inv.itemName} />)}
                </datalist>
                <button onClick={addStoreItemRow} style={{ padding: "12px", border: `2px dashed ${C.border}`, background: "transparent", color: C.accent, fontWeight: 800, fontSize: 12, borderRadius: 14, cursor: "pointer", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                  + Add Item
                </button>
              </div>

              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, paddingTop: 16, borderTop: `1px solid ${C.border}` }}>
                <div>
                  <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.06em" }}>Total Amount</p>
                  <p style={{ fontSize: 26, fontWeight: 900, color: C.accent }}>₹{storeItemsTotal.toLocaleString("en-IN")}</p>
                </div>
                <button onClick={handleCompleteStoreSale} disabled={storeSubmitting} style={{ ...btn("#10b981", "#fff", { padding: "13px 26px", fontSize: 13, boxShadow: "0 6px 18px rgba(16,185,129,.4)" }) }}>
                  <Ic.wa size={14} /> {storeSubmitting ? "Processing…" : "Complete Bill"}
                </button>
              </div>
            </div>
          )}

          {/* ══ QUOTATION MAKER ══ */}
          {activeTab === "Quotation" && (
            <div style={{ display: "grid", gridTemplateColumns: "380px 1fr", gap: 20, alignItems: "start" }}>

              {/* Form */}
              <div style={{ ...card(), padding: 24, position: "sticky", top: 84 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20, paddingBottom: 16, borderBottom: `1px solid ${C.border}` }}>
                  <div style={{ width: 32, height: 32, borderRadius: 10, background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <Ic.fileText size={15} style={{ color: "#fff" }} />
                  </div>
                  <div>
                    <p style={{ fontWeight: 800, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.05em" }}>New Quotation</p>
                    <p style={{ fontSize: 11, color: C.sub, marginTop: 2 }}>For schools, apartments & bulk/institutional clients</p>
                  </div>
                </div>

                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  <input placeholder="Client / company name *" style={inp} value={quotClientName} onChange={e => setQuotClientName(e.target.value)}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  <input placeholder="Contact person" style={inp} value={quotContactPerson} onChange={e => setQuotContactPerson(e.target.value)}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <input placeholder="Phone (for WhatsApp) *" style={inp} value={quotPhone} onChange={e => setQuotPhone(e.target.value)}
                      onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                    <input placeholder="Email (optional)" style={inp} value={quotEmail} onChange={e => setQuotEmail(e.target.value)}
                      onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  </div>
                  <textarea placeholder="Site address" rows={2} style={{ ...inp, resize: "vertical", fontFamily: "inherit" }} value={quotAddress}
                    onChange={e => setQuotAddress(e.target.value)}
                    onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <div>
                      <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 5 }}>Valid Until</p>
                      <input type="date" style={inp} value={quotValidUntil} onChange={e => setQuotValidUntil(e.target.value)}
                        onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                    </div>
                    <div>
                      <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 5 }}>Tax (%)</p>
                      <input type="number" placeholder="18" style={inp} value={quotTaxPercent} onChange={e => setQuotTaxPercent(e.target.value)}
                        onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                    </div>
                  </div>

                  <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginTop: 6 }}>Items / Services</p>
                  {quotItems.map((row, index) => (
                    <div key={row.id} style={{ background: C.inp, border: `1px solid ${C.border}`, borderRadius: 14, padding: 12, display: "flex", flexDirection: "column", gap: 8 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <span style={{ fontSize: 10, fontWeight: 800, color: C.sub, textTransform: "uppercase", letterSpacing: "0.06em" }}>Item #{index + 1}</span>
                        {quotItems.length > 1 && (
                          <button onClick={() => removeQuotItemRow(row.id)} style={{ background: "transparent", border: "none", color: C.danger, cursor: "pointer", display: "flex" }}>
                            <Ic.trash size={13} />
                          </button>
                        )}
                      </div>
                      <input placeholder="Description (e.g. RO Water Purifier - Emerald model)" value={row.description}
                        onChange={e => updateQuotItem(row.id, "description", e.target.value)}
                        style={inp} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                        <input type="number" placeholder="Qty" value={row.quantity}
                          onChange={e => updateQuotItem(row.id, "quantity", e.target.value)}
                          style={inp} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                        <input type="number" placeholder="Unit price (₹)" value={row.price}
                          onChange={e => updateQuotItem(row.id, "price", e.target.value)}
                          style={inp} onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                      </div>
                    </div>
                  ))}
                  <button onClick={addQuotItemRow} style={{ padding: "10px", border: `2px dashed ${C.border}`, background: "transparent", color: C.accent, fontWeight: 800, fontSize: 11.5, borderRadius: 12, cursor: "pointer", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                    + Add Item
                  </button>

                  <div>
                    <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 5, marginTop: 4 }}>Terms & Conditions</p>
                    <textarea rows={5} style={{ ...inp, resize: "vertical", fontFamily: "inherit", fontSize: 11.5, lineHeight: 1.5 }} value={quotTerms}
                      onChange={e => setQuotTerms(e.target.value)}
                      onFocus={e => e.target.style.borderColor = C.accent} onBlur={e => e.target.style.borderColor = C.border} />
                  </div>

                  <div style={{ background: dark ? "rgba(59,130,246,.08)" : "#eff6ff", border: `1px solid ${C.borderHi}`, borderRadius: 12, padding: "12px 14px", marginTop: 4 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.sub, marginBottom: 4 }}>
                      <span>Subtotal</span><span>₹{quotSubtotal.toLocaleString("en-IN")}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.sub, marginBottom: 8 }}>
                      <span>Tax ({quotTaxPercent || 0}%)</span><span>₹{quotTaxAmount.toLocaleString("en-IN")}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 800, color: C.text, paddingTop: 8, borderTop: `1px solid ${C.borderHi}` }}>
                      <span>Grand Total</span><span>₹{quotGrandTotal.toLocaleString("en-IN")}</span>
                    </div>
                  </div>

                  <button onClick={handleCreateQuotation} disabled={quotSubmitting} style={{ marginTop: 4, padding: "13px", background: "linear-gradient(135deg,#25D366,#128C7E)", color: "#fff", border: "none", borderRadius: 12, fontWeight: 800, fontSize: 13, cursor: quotSubmitting ? "wait" : "pointer", letterSpacing: "0.03em", boxShadow: "0 6px 18px rgba(37,211,102,.35)", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, opacity: quotSubmitting ? 0.7 : 1 }}>
                    <Ic.wa size={15} /> {quotSubmitting ? "Generating…" : "Generate & Send via WhatsApp"}
                  </button>
                </div>
              </div>

              {/* History */}
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                <p style={{ fontSize: 12, color: C.sub, paddingLeft: 4 }}>
                  <strong style={{ color: C.text }}>{quotations.length}</strong> quotation{quotations.length === 1 ? "" : "s"} sent
                </p>
                <div style={{ ...card(), overflow: "hidden" }}>
                  {quotationsLoading ? (
                    <div style={{ padding: "40px 20px", textAlign: "center", color: C.sub, fontSize: 13 }}>Loading quotations…</div>
                  ) : quotations.length === 0 ? (
                    <div style={{ padding: "56px 20px", textAlign: "center", color: C.sub }}>
                      <Ic.fileText size={32} style={{ color: C.border, marginBottom: 12 }} />
                      <p style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>No quotations yet</p>
                      <p style={{ fontSize: 12 }}>Fill out the form to generate and send your first one.</p>
                    </div>
                  ) : (
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: dark ? "#0c1526" : "#f8fafc", borderBottom: `1px solid ${C.border}` }}>
                          {["Quotation #", "Client", "Total", "Valid Until", ""].map((h, i) => (
                            <th key={i} style={{ padding: "13px 18px", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.09em", color: C.sub, textAlign: i >= 4 ? "right" : "left" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {quotations.map(q => (
                          <tr key={q.id} style={{ borderBottom: `1px solid ${C.border}`, transition: "background .12s" }}
                            onMouseEnter={e => e.currentTarget.style.background = C.cardHov}
                            onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                            <td style={{ padding: "15px 18px" }}>
                              <p style={{ fontWeight: 700, fontSize: 12.5 }}>{q.quotationNumber}</p>
                            </td>
                            <td style={{ padding: "15px 18px" }}>
                              <p style={{ fontWeight: 600, marginBottom: 1 }}>{q.clientName}</p>
                              <p style={{ fontSize: 11, color: C.sub }}>{q.phone}</p>
                            </td>
                            <td style={{ padding: "15px 18px" }}>
                              <p style={{ fontWeight: 700 }}>₹{(q.grandTotal || 0).toLocaleString("en-IN")}</p>
                            </td>
                            <td style={{ padding: "15px 18px" }}>
                              <span style={{ fontSize: 11, fontWeight: 600, background: dark ? "#1e293b" : "#f1f5f9", color: C.sub, padding: "4px 10px", borderRadius: 8 }}>{q.validUntil || "—"}</span>
                            </td>
                            <td style={{ padding: "15px 18px" }}>
                              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                                <a href={q.pdfUrl} target="_blank" rel="noreferrer" style={{ ...btn("#3b82f6", "#fff", { textDecoration: "none", fontSize: 11, boxShadow: "0 3px 10px rgba(59,130,246,.3)" }) }}>View PDF</a>
                                <button onClick={() => handleResendQuotation(q)} style={{ ...btn("#25D366", "#fff", { fontSize: 11, boxShadow: "0 3px 10px rgba(37,211,102,.35)", display: "flex", alignItems: "center", gap: 5 }) }}>
                                  <Ic.wa size={12} /> Resend
                                </button>
                                <button onClick={() => handleDeleteQuotation(q.id)} style={{ ...btn("transparent", C.danger, { border: `1.5px solid #fecaca`, fontSize: 11 }) }}
                                  onMouseEnter={e => e.currentTarget.style.background = "#fee2e2"} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                                  <Ic.trash size={12} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* ══ HISTORY ══ */}
          {activeTab === "History" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
                <div style={{ display: "flex", gap: 8 }}>
                  {["Arun Kumar", "Suresh Raina", "Store Sales"].map(name => {
                    const active = historyTab === name;
                    return (
                      <button key={name} onClick={() => setHistoryTab(name)} style={{ padding: "9px 22px", borderRadius: 12, border: active ? "none" : `1px solid ${C.border}`, cursor: "pointer", fontWeight: 700, fontSize: 13, background: active ? "linear-gradient(135deg,#3b82f6,#6366f1)" : C.card, color: active ? "#fff" : C.sub, boxShadow: active ? "0 4px 14px rgba(99,102,241,.35)" : "none", transition: "all .15s", display: "inline-flex", alignItems: "center", gap: 7 }}>
                        {name === "Store Sales" && <Ic.cash size={13} />}
                        {name}
                      </button>
                    );
                  })}
                </div>
                {/* ── Delete-all branches by tab: staff tabs clear
                     completedJobs for that employee; Store Sales clears
                     the storeSales collection instead. ── */}
                <button onClick={async () => {
                  if (!window.confirm(`Delete ALL ${historyTab === "Store Sales" ? "store sale" : `of ${historyTab}'s`} history?`)) return;
                  const targetCollection = historyTab === "Store Sales" ? "storeSales" : "completedJobs";
                  const snap = historyTab === "Store Sales"
                    ? await getDocs(collection(db, targetCollection))
                    : await getDocs(query(collection(db, targetCollection), where("employeeName", "==", historyTab)));
                  const b = writeBatch(db);
                  snap.docs.forEach(d => b.delete(d.ref));
                  await b.commit();
                }}
                  style={{ ...btn("#fee2e2", C.danger, { border: `1.5px solid #fecaca` }) }}>
                  <Ic.trash size={13} /> Delete All History
                </button>
              </div>

              {historyTab === "Store Sales" ? (
                storeSalesHistory.length === 0 ? (
                  <div style={{ ...card(), padding: "60px 20px", textAlign: "center", color: C.sub }}>
                    <Ic.cash size={36} style={{ color: C.border, marginBottom: 12 }} />
                    <p style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>No store sales yet</p>
                    <p style={{ fontSize: 12 }}>Completed Direct Store Sale bills will appear here.</p>
                  </div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {storeSalesHistory.map(sale => (
                      <div key={sale.id} style={{ ...card({ borderRadius: 16 }), padding: "18px 22px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 14, transition: "background .12s" }}
                        onMouseEnter={e => e.currentTarget.style.background = C.cardHov}
                        onMouseLeave={e => e.currentTarget.style.background = C.card}>
                        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                          <div style={{ width: 42, height: 42, borderRadius: 12, background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                            <Ic.cash size={18} style={{ color: "#fff" }} />
                          </div>
                          <div>
                            <p style={{ fontWeight: 700, fontSize: 14, marginBottom: 3 }}>{sale.customerName || "Walk-in customer"}</p>
                            <p style={{ fontSize: 12, color: C.sub }}>
                              {sale.date || "—"} &nbsp;·&nbsp; {sale.items?.length || 0} item{sale.items?.length === 1 ? "" : "s"} &nbsp;·&nbsp; Total: <strong style={{ color: C.text }}>₹{(sale.grandTotal || 0).toLocaleString("en-IN")}</strong>
                            </p>
                          </div>
                        </div>
                        <div style={{ display: "flex", gap: 8 }}>
                          {sale.pdfUrl && (
                            <a href={sale.pdfUrl} target="_blank" rel="noreferrer" style={{ ...btn("#10b981", "#fff", { textDecoration: "none", boxShadow: "0 3px 10px rgba(16,185,129,.3)" }) }}>View Bill</a>
                          )}
                          <button onClick={async () => { if (window.confirm("Delete this store sale record?")) await deleteDoc(doc(db, "storeSales", sale.id)); }}
                            style={{ ...btn("transparent", C.danger, { border: `1.5px solid #fecaca` }) }}
                            onMouseEnter={e => e.currentTarget.style.background = "#fee2e2"} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                            <Ic.trash size={12} />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )
              ) : (
                history.length === 0 ? (
                  <div style={{ ...card(), padding: "60px 20px", textAlign: "center", color: C.sub }}>
                    <Ic.history size={36} style={{ color: C.border, marginBottom: 12 }} />
                    <p style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>No records found</p>
                    <p style={{ fontSize: 12 }}>Completed jobs for {historyTab} will appear here.</p>
                  </div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {history.map(job => (
                      <div key={job.id} style={{ ...card({ borderRadius: 16 }), padding: "18px 22px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 14, transition: "background .12s" }}
                        onMouseEnter={e => e.currentTarget.style.background = C.cardHov}
                        onMouseLeave={e => e.currentTarget.style.background = C.card}>
                        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                          <div style={{ width: 42, height: 42, borderRadius: 12, background: "linear-gradient(135deg,#10b981,#06b6d4)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                            <Ic.history size={18} style={{ color: "#fff" }} />
                          </div>
                          <div>
                            <p style={{ fontWeight: 700, fontSize: 14, marginBottom: 3 }}>{job.customerName}</p>
                            <p style={{ fontSize: 12, color: C.sub }}>{job.date} &nbsp;·&nbsp; Total: <strong style={{ color: C.text }}>₹{job.grandTotal}</strong></p>
                          </div>
                        </div>
                        <div style={{ display: "flex", gap: 8 }}>
                          <a href={job.pdfUrl} target="_blank" rel="noreferrer" style={{ ...btn("#10b981", "#fff", { textDecoration: "none", boxShadow: "0 3px 10px rgba(16,185,129,.3)" }) }}>View Bill</a>
                          <button onClick={async () => { if (window.confirm("Delete this bill record?")) await deleteDoc(doc(db, "completedJobs", job.id)); }}
                            style={{ ...btn("transparent", C.danger, { border: `1.5px solid #fecaca` }) }}
                            onMouseEnter={e => e.currentTarget.style.background = "#fee2e2"} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                            <Ic.trash size={12} />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )
              )}
            </div>
          )}

          {/* ══ ANALYSIS ══ */}
          {activeTab === "Analysis" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
              {analysisLoading ? (
                <div style={{ ...card(), padding: "40px 20px", textAlign: "center", color: C.sub, fontSize: 13 }}>Loading revenue data…</div>
              ) : analysisRecords.length === 0 ? (
                <div style={{ ...card(), padding: "56px 20px", textAlign: "center", color: C.sub }}>
                  <Ic.chart size={32} style={{ color: C.border, marginBottom: 12 }} />
                  <p style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>No revenue data yet</p>
                  <p style={{ fontSize: 12 }}>Completed service jobs and store sales will show up here for analysis.</p>
                </div>
              ) : (
                <>
                  {/* Summary cards — revenue + order count together */}
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 14 }}>
                    {[
                      { label: "Today", rev: analysisSummary.todayTotal, orders: analysisSummary.todayOrders, grad: ["#3b82f6", "#6366f1"] },
                      { label: "This Month", rev: analysisSummary.monthTotal, orders: analysisSummary.monthOrders, grad: ["#10b981", "#06b6d4"] },
                      { label: "This Year", rev: analysisSummary.yearTotal, orders: analysisSummary.yearOrders, grad: ["#f59e0b", "#ef4444"] },
                      { label: "All Time", rev: analysisSummary.allTimeTotal, orders: analysisSummary.allTimeOrders, grad: ["#8b5cf6", "#ec4899"] },
                    ].map((s, i) => (
                      <div key={i} style={{ ...card(), padding: "16px 18px" }}>
                        <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>{s.label}</p>
                        <p style={{ fontSize: 22, fontWeight: 900, background: `linear-gradient(135deg,${s.grad[0]},${s.grad[1]})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", lineHeight: 1.2 }}>₹{s.rev.toLocaleString("en-IN")}</p>
                        <p style={{ fontSize: 11.5, color: C.sub, fontWeight: 600, marginTop: 3 }}>{s.orders} order{s.orders === 1 ? "" : "s"}</p>
                      </div>
                    ))}
                  </div>

                  {/* Chart card */}
                  <div style={{ ...card(), padding: 24 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20, flexWrap: "wrap", gap: 12 }}>
                      <p style={{ fontWeight: 800, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                        {analysisMetric === "revenue" ? "Revenue Trend" : "Order Volume Trend"}
                      </p>
                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                        {/* Revenue / Orders metric toggle */}
                        <div style={{ display: "flex", gap: 6, padding: 4, background: dark ? "#0a1525" : "#f1f5f9", borderRadius: 12 }}>
                          {[{ k: "revenue", label: "Revenue" }, { k: "orders", label: "Orders" }].map(m => (
                            <button key={m.k} onClick={() => setAnalysisMetric(m.k)} style={{
                              padding: "7px 14px", borderRadius: 9, border: "none", cursor: "pointer",
                              fontSize: 11.5, fontWeight: 700,
                              background: analysisMetric === m.k ? "linear-gradient(135deg,#10b981,#06b6d4)" : "transparent",
                              color: analysisMetric === m.k ? "#fff" : C.sub,
                              boxShadow: analysisMetric === m.k ? "0 3px 10px rgba(16,185,129,.35)" : "none",
                              transition: "all .15s",
                            }}>
                              {m.label}
                            </button>
                          ))}
                        </div>
                        {/* Day / Month / Year granularity toggle */}
                        <div style={{ display: "flex", gap: 6, padding: 4, background: dark ? "#0a1525" : "#f1f5f9", borderRadius: 12 }}>
                          {[{ k: "day", label: "Day" }, { k: "month", label: "Month" }, { k: "year", label: "Year" }].map(g => (
                            <button key={g.k} onClick={() => setAnalysisGranularity(g.k)} style={{
                              padding: "7px 16px", borderRadius: 9, border: "none", cursor: "pointer",
                              fontSize: 11.5, fontWeight: 700,
                              background: analysisGranularity === g.k ? "linear-gradient(135deg,#3b82f6,#6366f1)" : "transparent",
                              color: analysisGranularity === g.k ? "#fff" : C.sub,
                              boxShadow: analysisGranularity === g.k ? "0 3px 10px rgba(99,102,241,.35)" : "none",
                              transition: "all .15s",
                            }}>
                              {g.label}
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>

                    {(() => {
                      const CHART_H = 210;
                      const getVal = (d) => analysisMetric === "revenue" ? d.total : d.totalOrders;
                      const getStorePart = (d) => analysisMetric === "revenue" ? d.store : d.storeOrders;
                      const maxVal = Math.max(1, ...analysisChartData.map(getVal));
                      // Round the gridline ceiling up to a "nice" number so labels
                      // read cleanly (e.g. 1,200 -> 1,500 not an arbitrary max).
                      const niceMax = (() => {
                        const magnitude = Math.pow(10, Math.floor(Math.log10(maxVal || 1)));
                        const steps = [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10];
                        for (const s of steps) { if (maxVal <= s * magnitude) return s * magnitude; }
                        return maxVal;
                      })();
                      const gridLines = [1, 0.75, 0.5, 0.25, 0];
                      const fmtVal = (v) => analysisMetric === "revenue" ? `₹${Math.round(v).toLocaleString("en-IN")}` : `${Math.round(v)}`;

                      return (
                        <div style={{ position: "relative", paddingLeft: 54 }}>
                          {/* Gridlines + axis labels */}
                          <div style={{ position: "absolute", left: 0, top: 0, width: "100%", height: CHART_H }}>
                            {gridLines.map(f => (
                              <div key={f} style={{ position: "absolute", left: 54, right: 0, top: `${(1 - f) * CHART_H}px`, borderTop: `1px dashed ${C.border}` }}>
                                <span style={{ position: "absolute", left: -54, top: -7, width: 46, textAlign: "right", fontSize: 9.5, color: C.sub, fontWeight: 600 }}>
                                  {f === 0 ? "0" : fmtVal(niceMax * f)}
                                </span>
                              </div>
                            ))}
                          </div>

                          <div style={{ display: "flex", alignItems: "flex-end", gap: 14, height: CHART_H, position: "relative", zIndex: 1 }}>
                            {analysisChartData.map(d => {
                              const val = getVal(d);
                              const barHeightPx = val > 0 ? Math.max(3, (val / niceMax) * CHART_H) : 1;
                              const storePart = getStorePart(d);
                              const storeHeightPx = val > 0 ? (storePart / val) * barHeightPx : 0;
                              const servicedHeightPx = barHeightPx - storeHeightPx;
                              return (
                                <div key={d.key} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 8, minWidth: 0, height: "100%", justifyContent: "flex-end" }}>
                                  <span style={{ fontSize: 10.5, fontWeight: 800, color: C.text, whiteSpace: "nowrap" }}>
                                    {val > 0 ? fmtVal(val) : ""}
                                  </span>
                                  <div title={`${formatAnalysisPeriodLabel(d.key, analysisGranularity)}: ${fmtVal(val)}`}
                                    style={{
                                      width: "100%", maxWidth: 44, display: "flex", flexDirection: "column-reverse",
                                      height: barHeightPx, borderRadius: "8px 8px 3px 3px", overflow: "hidden",
                                      boxShadow: val > 0 ? "0 4px 10px rgba(59,130,246,.25)" : "none",
                                      transition: "transform .15s", cursor: "default",
                                    }}
                                    onMouseEnter={e => e.currentTarget.style.transform = "scaleY(1.02)"}
                                    onMouseLeave={e => e.currentTarget.style.transform = "scaleY(1)"}>
                                    <div style={{ height: storeHeightPx, background: "linear-gradient(180deg,#60a5fa,#3b82f6)" }} />
                                    <div style={{ height: servicedHeightPx, background: "linear-gradient(180deg,#34d399,#10b981)" }} />
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                          <div style={{ display: "flex", gap: 14, paddingLeft: 0, marginTop: 8 }}>
                            {analysisChartData.map(d => (
                              <span key={d.key} style={{ flex: 1, fontSize: 9.5, color: C.sub, fontWeight: 600, textAlign: "center", lineHeight: 1.3 }}>
                                {formatAnalysisPeriodLabel(d.key, analysisGranularity)}
                              </span>
                            ))}
                          </div>
                        </div>
                      );
                    })()}

                    <div style={{ display: "flex", gap: 18, marginTop: 18, paddingTop: 14, borderTop: `1px solid ${C.border}` }}>
                      <span style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 11.5, color: C.sub, fontWeight: 600 }}>
                        <span style={{ width: 10, height: 10, borderRadius: 3, background: "linear-gradient(135deg,#60a5fa,#3b82f6)", display: "inline-block" }} /> Store Sales
                      </span>
                      <span style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 11.5, color: C.sub, fontWeight: 600 }}>
                        <span style={{ width: 10, height: 10, borderRadius: 3, background: "linear-gradient(135deg,#34d399,#10b981)", display: "inline-block" }} /> Service Jobs
                      </span>
                    </div>
                  </div>

                  {/* Breakdown table — columns adapt to the selected metric */}
                  <div style={{ ...card(), overflow: "hidden" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: dark ? "#0c1526" : "#f8fafc", borderBottom: `1px solid ${C.border}` }}>
                          {[
                            analysisGranularity === "day" ? "Date" : analysisGranularity === "month" ? "Month" : "Year",
                            analysisMetric === "revenue" ? "Store Sales" : "Store Orders",
                            analysisMetric === "revenue" ? "Service Jobs" : "Service Orders",
                            analysisMetric === "revenue" ? "Total Revenue" : "Total Orders",
                          ].map((h, i) => (
                            <th key={i} style={{ padding: "13px 18px", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.09em", color: C.sub, textAlign: i === 0 ? "left" : "right" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {analysisAggregated.map(row => (
                          <tr key={row.key} style={{ borderBottom: `1px solid ${C.border}`, transition: "background .12s" }}
                            onMouseEnter={e => e.currentTarget.style.background = C.cardHov}
                            onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                            <td style={{ padding: "13px 18px", fontWeight: 700 }}>{formatAnalysisPeriodLabel(row.key, analysisGranularity)}</td>
                            {analysisMetric === "revenue" ? (
                              <>
                                <td style={{ padding: "13px 18px", textAlign: "right", color: C.sub }}>₹{row.store.toLocaleString("en-IN")}</td>
                                <td style={{ padding: "13px 18px", textAlign: "right", color: C.sub }}>₹{row.service.toLocaleString("en-IN")}</td>
                                <td style={{ padding: "13px 18px", textAlign: "right", fontWeight: 800 }}>₹{row.total.toLocaleString("en-IN")}</td>
                              </>
                            ) : (
                              <>
                                <td style={{ padding: "13px 18px", textAlign: "right", color: C.sub }}>{row.storeOrders}</td>
                                <td style={{ padding: "13px 18px", textAlign: "right", color: C.sub }}>{row.serviceOrders}</td>
                                <td style={{ padding: "13px 18px", textAlign: "right", fontWeight: 800 }}>{row.totalOrders}</td>
                              </>
                            )}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}

              {/* ── Inventory Stock Analysis — separate section, same page.
                   Derived from Store Sale line items (units sold per
                   day/month/year) plus the live current-stock snapshot. ── */}
              <div style={{ marginTop: 8, paddingTop: 8, borderTop: `1px solid ${C.border}` }}>
                <p style={{ fontWeight: 800, fontSize: 15, marginBottom: 3 }}>Inventory Stock Analysis</p>
                <p style={{ fontSize: 12, color: C.sub, marginBottom: 18 }}>Units sold and current stock, based on Direct Store Sale records.</p>
              </div>

              {/* Current stock overview cards */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(170px,1fr))", gap: 14 }}>
                {[
                  { label: "Total Items Tracked", val: stockOverview.totalItems, grad: ["#3b82f6", "#6366f1"], suffix: "" },
                  { label: "Total Units In Stock", val: stockOverview.totalUnits, grad: ["#10b981", "#06b6d4"], suffix: "" },
                  { label: "Low Stock Items", val: stockOverview.lowStock, grad: ["#f59e0b", "#ef4444"], suffix: "" },
                  { label: "Out Of Stock", val: stockOverview.outOfStock, grad: stockOverview.outOfStock > 0 ? ["#ef4444", "#b91c1c"] : ["#8b5cf6", "#ec4899"], suffix: "" },
                ].map((s, i) => (
                  <div key={i} style={{ ...card(), padding: "16px 18px" }}>
                    <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>{s.label}</p>
                    <p style={{ fontSize: 24, fontWeight: 900, background: `linear-gradient(135deg,${s.grad[0]},${s.grad[1]})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>{s.val}</p>
                  </div>
                ))}
              </div>

              {inventoryMovementRecords.length === 0 ? (
                <div style={{ ...card(), padding: "40px 20px", textAlign: "center", color: C.sub }}>
                  <Ic.box size={30} style={{ color: C.border, marginBottom: 10 }} />
                  <p style={{ fontSize: 13.5, fontWeight: 600, marginBottom: 4 }}>No stock movement recorded yet</p>
                  <p style={{ fontSize: 12 }}>Items sold through Direct Store Sale will show up here — matched against your current inventory levels.</p>
                </div>
              ) : (
                <>
                  {/* Units-sold chart — reuses the same Day/Month/Year granularity as the revenue chart above */}
                  <div style={{ ...card(), padding: 24 }}>
                    <p style={{ fontWeight: 800, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 20 }}>Units Sold Trend</p>
                    {(() => {
                      const CHART_H = 190;
                      const maxUnits = Math.max(1, ...inventoryChartData.map(d => d.units));
                      const niceMax = (() => {
                        const magnitude = Math.pow(10, Math.floor(Math.log10(maxUnits || 1)));
                        const steps = [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10];
                        for (const s of steps) { if (maxUnits <= s * magnitude) return s * magnitude; }
                        return maxUnits;
                      })();
                      const gridLines = [1, 0.75, 0.5, 0.25, 0];
                      return (
                        <div style={{ position: "relative", paddingLeft: 44 }}>
                          <div style={{ position: "absolute", left: 0, top: 0, width: "100%", height: CHART_H }}>
                            {gridLines.map(f => (
                              <div key={f} style={{ position: "absolute", left: 44, right: 0, top: `${(1 - f) * CHART_H}px`, borderTop: `1px dashed ${C.border}` }}>
                                <span style={{ position: "absolute", left: -44, top: -7, width: 36, textAlign: "right", fontSize: 9.5, color: C.sub, fontWeight: 600 }}>
                                  {f === 0 ? "0" : Math.round(niceMax * f)}
                                </span>
                              </div>
                            ))}
                          </div>
                          <div style={{ display: "flex", alignItems: "flex-end", gap: 14, height: CHART_H, position: "relative", zIndex: 1 }}>
                            {inventoryChartData.map(d => {
                              const barHeightPx = d.units > 0 ? Math.max(3, (d.units / niceMax) * CHART_H) : 1;
                              return (
                                <div key={d.key} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 8, minWidth: 0, height: "100%", justifyContent: "flex-end" }}>
                                  <span style={{ fontSize: 10.5, fontWeight: 800, color: C.text }}>{d.units > 0 ? d.units : ""}</span>
                                  <div title={`${formatAnalysisPeriodLabel(d.key, analysisGranularity)}: ${d.units} units`}
                                    style={{ width: "100%", maxWidth: 44, height: barHeightPx, borderRadius: "8px 8px 3px 3px", background: "linear-gradient(180deg,#c084fc,#8b5cf6)", boxShadow: d.units > 0 ? "0 4px 10px rgba(139,92,246,.25)" : "none" }} />
                                </div>
                              );
                            })}
                          </div>
                          <div style={{ display: "flex", gap: 14, marginTop: 8 }}>
                            {inventoryChartData.map(d => (
                              <span key={d.key} style={{ flex: 1, fontSize: 9.5, color: C.sub, fontWeight: 600, textAlign: "center", lineHeight: 1.3 }}>
                                {formatAnalysisPeriodLabel(d.key, analysisGranularity)}
                              </span>
                            ))}
                          </div>
                        </div>
                      );
                    })()}
                  </div>

                  {/* Top selling items vs current stock */}
                  <div style={{ ...card(), overflow: "hidden" }}>
                    <div style={{ padding: "16px 18px", borderBottom: `1px solid ${C.border}` }}>
                      <p style={{ fontWeight: 800, fontSize: 13, textTransform: "uppercase", letterSpacing: "0.05em" }}>Top Selling Items</p>
                    </div>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: dark ? "#0c1526" : "#f8fafc", borderBottom: `1px solid ${C.border}` }}>
                          {["Item", "Units Sold", "Current Stock", "Status"].map((h, i) => (
                            <th key={i} style={{ padding: "13px 18px", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.09em", color: C.sub, textAlign: i === 0 ? "left" : i === 3 ? "center" : "right" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {topSellingItems.map(item => {
                          const stockKnown = item.currentStock !== null;
                          const low = stockKnown && item.currentStock <= 5;
                          return (
                            <tr key={item.itemName} style={{ borderBottom: `1px solid ${C.border}`, transition: "background .12s" }}
                              onMouseEnter={e => e.currentTarget.style.background = C.cardHov}
                              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                              <td style={{ padding: "13px 18px", fontWeight: 700 }}>{item.itemName}</td>
                              <td style={{ padding: "13px 18px", textAlign: "right", color: C.sub }}>{item.unitsSold}</td>
                              <td style={{ padding: "13px 18px", textAlign: "right", color: C.sub }}>{stockKnown ? item.currentStock : "—"}</td>
                              <td style={{ padding: "13px 18px", textAlign: "center" }}>
                                {stockKnown ? (
                                  <span style={{ fontSize: 11, fontWeight: 700, background: low ? "#fee2e2" : "#d1fae5", color: low ? C.danger : "#065f46", padding: "4px 10px", borderRadius: 8 }}>
                                    {low ? "Low Stock" : "In Stock"}
                                  </span>
                                ) : (
                                  <span style={{ fontSize: 11, color: C.sub }}>Not in inventory</span>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          )}

          {/* ══ STAFF PERFORMANCE ══ */}
          {activeTab === "Performance" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
              {performanceLoading ? (
                <div style={{ ...card(), padding: "40px 20px", textAlign: "center", color: C.sub, fontSize: 13 }}>Loading staff performance…</div>
              ) : staffPerformance.length === 0 ? (
                <div style={{ ...card(), padding: "56px 20px", textAlign: "center", color: C.sub }}>
                  <Ic.target size={32} style={{ color: C.border, marginBottom: 12 }} />
                  <p style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>No completed jobs yet</p>
                  <p style={{ fontSize: 12 }}>Staff performance builds up as jobs are completed in the field app.</p>
                </div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 16 }}>
                  {staffPerformance.map((s, i) => (
                    <div key={s.name} style={{ ...card(), padding: 22 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 18, paddingBottom: 16, borderBottom: `1px solid ${C.border}` }}>
                        <div style={{ width: 42, height: 42, borderRadius: "50%", background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 15, flexShrink: 0 }}>
                          {(s.name || "?")[0]}
                        </div>
                        <div>
                          <p style={{ fontWeight: 800, fontSize: 15 }}>{s.name}</p>
                          {i === 0 && s.revenue > 0 && (
                            <span style={{ fontSize: 10, fontWeight: 700, color: "#f59e0b", textTransform: "uppercase", letterSpacing: "0.05em" }}>★ Top Performer</span>
                          )}
                        </div>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                        <div>
                          <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 4 }}>Jobs Completed</p>
                          <p style={{ fontSize: 22, fontWeight: 900, color: C.text }}>{s.jobCount}</p>
                        </div>
                        <div>
                          <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 4 }}>Revenue</p>
                          <p style={{ fontSize: 22, fontWeight: 900, background: "linear-gradient(135deg,#10b981,#06b6d4)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>₹{s.revenue.toLocaleString("en-IN")}</p>
                        </div>
                        <div>
                          <p style={{ fontSize: 10, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 4 }}>Avg / Job</p>
                          <p style={{ fontSize: 16, fontWeight: 800, color: C.text }}>₹{Math.round(s.avgPerJob).toLocaleString("en-IN")}</p>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* ══ SCHEDULE ══ */}
          {activeTab === "Schedule" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
              <div style={{ ...card({ borderRadius: 16 }), padding: "14px 18px", display: "flex", alignItems: "center", gap: 10 }}>
                <Ic.route size={15} style={{ color: C.sub, flexShrink: 0 }} />
                <p style={{ fontSize: 12, color: C.sub }}>
                  Shows each technician's currently-assigned jobs in the order they were assigned. There's no separate "scheduled date" in this app yet — this reflects assignment order, not a forward-looking calendar.
                </p>
              </div>

              {scheduleLoading ? (
                <div style={{ ...card(), padding: "40px 20px", textAlign: "center", color: C.sub, fontSize: 13 }}>Loading schedule…</div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(320px,1fr))", gap: 20 }}>
                  {scheduleByStaff.map(({ name, tasks }) => (
                    <div key={name} style={{ ...card(), padding: 22 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, paddingBottom: 14, borderBottom: `1px solid ${C.border}` }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <div style={{ width: 36, height: 36, borderRadius: "50%", background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 13, flexShrink: 0 }}>
                            {(name || "?")[0]}
                          </div>
                          <p style={{ fontWeight: 800, fontSize: 14 }}>{name}</p>
                        </div>
                        <span style={{ fontSize: 11, fontWeight: 700, color: C.sub, background: dark ? "#1e293b" : "#f1f5f9", padding: "3px 10px", borderRadius: 999 }}>
                          {tasks.length} job{tasks.length === 1 ? "" : "s"}
                        </span>
                      </div>
                      {tasks.length === 0 ? (
                        <p style={{ fontSize: 12, color: C.sub, textAlign: "center", padding: "16px 0" }}>No jobs currently assigned.</p>
                      ) : (
                        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                          {tasks.map((t, i) => {
                            const d = t.timestamp ? new Date(t.timestamp) : null;
                            const dateLabel = d && !isNaN(d) ? d.toLocaleDateString("en-IN", { day: "numeric", month: "short" }) : "—";
                            const timeLabel = d && !isNaN(d) ? d.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" }) : "";
                            return (
                              <div key={t.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: dark ? "#0a1525" : "#f8fafc", borderRadius: 12, border: `1px solid ${C.border}` }}>
                                <span style={{ width: 22, height: 22, borderRadius: "50%", background: C.accent, color: "#fff", fontSize: 10, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{i + 1}</span>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <p style={{ fontSize: 13, fontWeight: 700, color: C.text, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{t.customerName}</p>
                                </div>
                                <span style={{ fontSize: 10.5, color: C.sub, fontWeight: 600, whiteSpace: "nowrap" }}>{dateLabel}{timeLabel ? ` · ${timeLabel}` : ""}</span>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </main>
      </div>

      {/* ══ ASSIGN MODAL ══ */}
      {selectedCustomer && (
        <div onClick={() => setSelectedCustomer(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.6)", backdropFilter: "blur(6px)", zIndex: 60, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
          <div onClick={e => e.stopPropagation()} style={{ ...card(), width: "100%", maxWidth: 440, overflow: "hidden", animation: "slideUp .2s ease" }}>
            <div style={{ background: "linear-gradient(135deg,#3b82f6,#6366f1)", padding: "22px 24px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <p style={{ color: "rgba(255,255,255,.7)", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>Assign Technician</p>
                  <p style={{ color: "#fff", fontWeight: 800, fontSize: 16 }}>{selectedCustomer.name}</p>
                  <p style={{ color: "rgba(255,255,255,.7)", fontSize: 12, marginTop: 2 }}>{selectedCustomer.address}</p>
                </div>
                <button onClick={() => setSelectedCustomer(null)} style={{ background: "rgba(255,255,255,.2)", border: "none", borderRadius: 8, width: 30, height: 30, color: "#fff", fontSize: 18, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1 }}>×</button>
              </div>
            </div>
            <div style={{ padding: 20, display: "flex", flexDirection: "column", gap: 10 }}>
              <p style={{ fontSize: 11, color: C.sub, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>Choose Technician</p>
              {staff.map(emp => (
                <button key={emp.id} onClick={() => handleAssign(emp)} style={{ display: "flex", alignItems: "center", gap: 14, padding: "14px 16px", background: C.inp, border: `1.5px solid ${C.border}`, borderRadius: 14, cursor: "pointer", transition: "all .15s", width: "100%", textAlign: "left" }}
                  onMouseEnter={e => { e.currentTarget.style.borderColor = C.accent; e.currentTarget.style.background = dark ? "rgba(59,130,246,.1)" : "#eff6ff"; }}
                  onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.background = C.inp; }}>
                  <div style={{ width: 40, height: 40, borderRadius: "50%", background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 15, flexShrink: 0 }}>{emp.name[0]}</div>
                  <div style={{ flex: 1 }}>
                    <p style={{ fontWeight: 700, fontSize: 13, color: C.text, marginBottom: 2 }}>{emp.name}</p>
                    <p style={{ fontSize: 11, color: C.sub }}>{emp.phone}</p>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: emp.status === "Online" ? C.success : C.sub }} />
                    <span style={{ fontSize: 11, color: emp.status === "Online" ? C.success : C.sub, fontWeight: 600 }}>{emp.status}</span>
                    <Ic.wa size={14} style={{ color: "#25D366", marginLeft: 4 }} />
                  </div>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ══ DUPLICATE CUSTOMERS MODAL ══
           NEW — permanent-fix UI: shows every cluster of likely duplicate
           customer records (same phone or very similar name), side by
           side, so an admin can confirm and merge them without ever
           opening the Firebase console. Merging is always a manual,
           per-cluster confirmation — never automatic — since two
           different real people can legitimately share a phone or a
           similar name. */}
      {showDuplicatesModal && (
        <div onClick={() => setShowDuplicatesModal(false)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.6)", backdropFilter: "blur(6px)", zIndex: 70, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
          <div onClick={e => e.stopPropagation()} style={{ ...card(), width: "100%", maxWidth: 720, maxHeight: "85vh", overflow: "hidden", display: "flex", flexDirection: "column", animation: "slideUp .2s ease" }}>
            <div style={{ background: "linear-gradient(135deg,#f59e0b,#ef4444)", padding: "20px 24px", flexShrink: 0 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <p style={{ color: "rgba(255,255,255,.8)", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>Duplicate Customers</p>
                  <p style={{ color: "#fff", fontWeight: 800, fontSize: 17 }}>{duplicateClusters.length} possible duplicate{duplicateClusters.length === 1 ? "" : "s"} found</p>
                  <p style={{ color: "rgba(255,255,255,.85)", fontSize: 12, marginTop: 4 }}>Review each group below, then merge if they're really the same person.</p>
                </div>
                <button onClick={() => setShowDuplicatesModal(false)} style={{ background: "rgba(255,255,255,.2)", border: "none", borderRadius: 8, width: 30, height: 30, color: "#fff", fontSize: 18, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1, flexShrink: 0 }}>×</button>
              </div>
            </div>

            {/* ── NEW: filter by match type — "Same Phone" groups are a
                 much stronger signal (an identical phone number really is
                 shared) than "Similar Name" groups (only the name looks
                 alike; every phone number in the group is different). ── */}
            <div style={{ padding: "14px 20px 0", flexShrink: 0 }}>
              <div style={{ display: "flex", gap: 6, padding: 4, background: dark ? "#0a1525" : "#f1f5f9", borderRadius: 12, width: "fit-content" }}>
                {[
                  { k: "all", label: `All (${duplicateClusters.length})` },
                  { k: "phone", label: `Same Phone (${dupPhoneCount})` },
                  { k: "name", label: `Similar Name Only (${dupNameCount})` },
                ].map(f => (
                  <button key={f.k} onClick={() => setDupFilter(f.k)} style={{
                    padding: "8px 14px", borderRadius: 9, border: "none", cursor: "pointer",
                    fontSize: 11.5, fontWeight: 700,
                    background: dupFilter === f.k ? "linear-gradient(135deg,#3b82f6,#6366f1)" : "transparent",
                    color: dupFilter === f.k ? "#fff" : C.sub,
                    boxShadow: dupFilter === f.k ? "0 3px 10px rgba(99,102,241,.35)" : "none",
                    transition: "all .15s",
                  }}>
                    {f.label}
                  </button>
                ))}
              </div>
              {dupFilter === "name" && (
                <p style={{ fontSize: 11, color: C.warn, fontWeight: 600, marginTop: 8 }}>
                  ⚠ These share no phone number at all — only a similar-looking name. Double-check each one; it may well be two different people (e.g. father &amp; son, or a common local name).
                </p>
              )}
            </div>

            <div style={{ padding: 20, overflowY: "auto", display: "flex", flexDirection: "column", gap: 14 }}>
              {filteredDuplicateClusters.length === 0 ? (
                <p style={{ fontSize: 13, color: C.sub, textAlign: "center", padding: "30px 0" }}>
                  {duplicateClusters.length === 0 ? "No duplicates found. 🎉" : "No duplicates match this filter."}
                </p>
              ) : (
                filteredDuplicateClusters.map(cluster => (
                  <div key={cluster.key} style={{ border: `1px solid ${C.border}`, borderRadius: 14, padding: 16, background: dark ? "#0a1525" : "#f8fafc" }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                      <span style={{
                        fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.05em",
                        padding: "3px 10px", borderRadius: 999,
                        background: cluster.matchType === "phone" ? (dark ? "rgba(16,185,129,.15)" : "#d1fae5") : (dark ? "rgba(245,158,11,.15)" : "#fef3c7"),
                        color: cluster.matchType === "phone" ? "#065f46" : "#92400e",
                      }}>
                        {cluster.matchType === "phone" ? "Same Phone" : "Similar Name Only"}
                      </span>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 12 }}>
                      {cluster.members.map(m => (
                        <div key={m.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "10px 12px", background: C.card, borderRadius: 10, border: `1px solid ${C.border}` }}>
                          <div style={{ minWidth: 0 }}>
                            <p style={{ fontWeight: 700, fontSize: 13.5 }}>{m.name || "(no name)"}</p>
                            <p style={{ fontSize: 11.5, color: C.sub }}>
                              {m.phone || "no phone"} · {m.installDate || "no install date"} · doc id: <code style={{ fontSize: 10 }}>{m.id}</code>
                            </p>
                          </div>
                          <span style={{ fontSize: 10, fontWeight: 700, color: C.accent, textTransform: "uppercase", flexShrink: 0 }}>
                            {(m.allServiceHistory || []).length} service record{(m.allServiceHistory || []).length === 1 ? "" : "s"}
                          </span>
                        </div>
                      ))}
                    </div>
                    <div style={{ display: "flex", justifyContent: "flex-end" }}>
                      <button
                        onClick={() => handleMergeCluster(cluster)}
                        disabled={mergingClusterKey === cluster.key}
                        style={{ ...btn("linear-gradient(135deg,#3b82f6,#6366f1)", "#fff", { fontSize: 12, opacity: mergingClusterKey === cluster.key ? 0.7 : 1, cursor: mergingClusterKey === cluster.key ? "wait" : "pointer" }) }}>
                        {mergingClusterKey === cluster.key ? "Merging…" : "Merge These Into One"}
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { -webkit-font-smoothing: antialiased; }
        input::placeholder, textarea::placeholder { color: #94a3b8; }
        input:focus, select:focus { border-color: #3b82f6 !important; box-shadow: 0 0 0 3px rgba(59,130,246,.15) !important; outline: none; }
        ::-webkit-scrollbar { width: 5px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 99px; }
        @keyframes pulseRing { 0%,100%{box-shadow:0 0 0 0 rgba(239,68,68,.5)} 50%{box-shadow:0 0 0 6px rgba(239,68,68,0)} }
        @keyframes slideUp { from{opacity:0;transform:translateY(16px)} to{opacity:1;transform:translateY(0)} }
      `}</style>
    </div>
  );
};
export default AdminDashboard;