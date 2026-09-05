/**
 * AquaServe — Excel writeback backend (Firebase Cloud Functions v2)
 *
 * Exposes three authenticated HTTPS endpoints that the React admin app
 * calls after saving to Firestore:
 *
 *   POST   /api/warranty-clients        → append a row to
 *                                          Excel-Warranty_Ledger.xlsx
 *   POST   /api/service-clients         → append a row to
 *                                          Excel-Warranty_Flat.xlsx
 *   DELETE /api/customers/:id           → remove the matching row from
 *                                          one or both files
 *
 * Design:
 *  - Firestore stays the source of truth; the React app writes there
 *    directly and only calls this backend to keep the two Excel files
 *    in sync in real time.
 *  - Every write to a given workbook is serialized with a short-lived
 *    Firestore-based lock (see withFileLock) so two admins saving at
 *    the same moment can't clobber each other's row.
 *  - Every attempt (success or failure) is written to the `auditLog`
 *    collection for traceability.
 *  - Uses ExcelJS (not SheetJS) for the read-modify-write cycle because
 *    it round-trips existing charts/formatting far more reliably.
 *
 * Deploy:
 *   cd functions && npm install firebase-admin firebase-functions exceljs express cors
 *   firebase deploy --only functions
 *
 * Required env/config:
 *   None beyond the default Application Default Credentials Cloud
 *   Functions already has for your project's Firestore + Storage.
 */

const { onRequest } = require("firebase-functions/v2/https");
const { initializeApp } = require("firebase-admin/app");
const { getAuth } = require("firebase-admin/auth");
const { getFirestore, FieldValue } = require("firebase-admin/firestore");
const { getStorage } = require("firebase-admin/storage");
const ExcelJS = require("exceljs");
const express = require("express");
const cors = require("cors");

initializeApp();

const LEDGER_PATH = "company_data/Warranty_Ledger.xlsx";
const FLAT_PATH = "company_data/Warranty_Flat.xlsx";
const LEDGER_HEADERS = ["NAME", "DATE", "ADDRESS", "PHONE NUMBER", "MODEL MECHINE", "PRICE", "PLACE", "services", "FEEDBACK"];
const FLAT_HEADERS = ["Customer Name", "Phone No 1", "Address", "Service Date", "Service Description", "Amount"];

/* ══════════════════════════════════════════════════════════
   Per-file lock (Firestore-based mutex). Serializes concurrent writes
   to the same workbook so two admins saving within the same second
   queue instead of racing on a download → edit → upload cycle.
══════════════════════════════════════════════════════════ */
async function withFileLock(lockName, fn) {
  const db = getFirestore();
  const lockRef = db.doc(`locks/${lockName}`);
  const start = Date.now();

  while (true) {
    const acquired = await db.runTransaction(async (tx) => {
      const snap = await tx.get(lockRef);
      const now = Date.now();
      const held = snap.exists && snap.data().expiresAt > now;
      if (held) return false;
      tx.set(lockRef, { expiresAt: now + 30000, holder: process.env.K_REVISION || "local" });
      return true;
    });
    if (acquired) break;
    if (Date.now() - start > 15000) throw new Error("Timed out waiting for file lock: " + lockName);
    await new Promise((r) => setTimeout(r, 300));
  }

  try {
    return await fn();
  } finally {
    await db.doc(`locks/${lockName}`).delete().catch(() => {});
  }
}

/* ══════════════════════════════════════════════════════════
   Auth middleware — verifies the Firebase ID token sent by the React
   app and requires an admin custom claim. Least-privilege: this backend
   only accepts requests from signed-in admins, never trusts a client-
   supplied uid/role in the request body.
══════════════════════════════════════════════════════════ */
async function requireAdmin(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : null;
  if (!token) return res.status(401).json({ error: "Missing Authorization header" });
  try {
    const decoded = await getAuth().verifyIdToken(token);
    if (decoded.admin !== true) {
      // Adjust this check to however you set admin claims, e.g.:
      //   admin.auth().setCustomUserClaims(uid, { admin: true })
      return res.status(403).json({ error: "Admin privileges required" });
    }
    req.uid = decoded.uid;
    next();
  } catch (e) {
    return res.status(401).json({ error: "Invalid or expired token" });
  }
}

async function writeAudit(entry) {
  try {
    await getFirestore().collection("auditLog").add({
      ...entry,
      at: FieldValue.serverTimestamp(),
    });
  } catch (e) {
    console.error("Failed to write audit log entry:", e);
  }
}

/* ══════════════════════════════════════════════════════════
   Excel helpers (ExcelJS)
══════════════════════════════════════════════════════════ */
async function loadWorkbook(storagePath) {
  const bucket = getStorage().bucket();
  const file = bucket.file(storagePath);
  const [exists] = await file.exists();
  const wb = new ExcelJS.Workbook();
  if (exists) {
    const [buf] = await file.download();
    await wb.xlsx.load(buf);
  } else {
    // First write ever — create a fresh workbook with the expected header row.
    const ws = wb.addWorksheet("Sheet1");
    ws.addRow(storagePath === LEDGER_PATH ? LEDGER_HEADERS : FLAT_HEADERS);
  }
  return { wb, file };
}

async function saveWorkbook(wb, file) {
  const outBuf = await wb.xlsx.writeBuffer();
  await file.save(Buffer.from(outBuf), {
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function getDataSheet(wb) {
  return wb.worksheets.find((s) => /^sheet\s*1$/i.test(s.name)) || wb.worksheets[0];
}

function headerRowValues(ws) {
  const row = ws.getRow(1);
  const values = [];
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    values[colNumber - 1] = String(cell.value || "").trim();
  });
  return values;
}

function appendRow(ws, headers, rowObject) {
  ws.addRow(headers.map((h) => (rowObject[h] !== undefined ? rowObject[h] : "")));
}

// Removes the first row (excluding header) whose PHONE column matches.
// Returns true if a row was removed.
function removeRowByPhone(ws, phone, phoneColumnNames) {
  if (!phone) return false;
  const headers = headerRowValues(ws);
  const colIdx = headers.findIndex((h) => phoneColumnNames.some((n) => h.toUpperCase() === n.toUpperCase()));
  if (colIdx === -1) return false;

  let targetRow = null;
  for (let r = 2; r <= ws.rowCount; r++) {
    const cellVal = String(ws.getRow(r).getCell(colIdx + 1).value || "").replace(/\D/g, "");
    if (cellVal.slice(-10) === String(phone).slice(-10)) {
      targetRow = r;
      break;
    }
  }
  if (targetRow) {
    ws.spliceRows(targetRow, 1);
    return true;
  }
  return false;
}

/* ══════════════════════════════════════════════════════════
   Express app + routes
══════════════════════════════════════════════════════════ */
const app = express();
app.use(cors({ origin: true }));
app.use(express.json());
app.use(requireAdmin);

// POST /warranty-clients  (full URL: <function base>/api/warranty-clients,
// since the Cloud Function itself is named "api" and that name already
// supplies the leading /api segment — routes here must NOT repeat it.)
app.post("/warranty-clients", async (req, res) => {
  const { customerId, name, phone, address, place, machineModel, price, installDate, remarks } = req.body || {};
  if (!name || !phone) return res.status(400).json({ error: "name and phone are required" });

  try {
    await withFileLock("ledger", async () => {
      const { wb, file } = await loadWorkbook(LEDGER_PATH);
      const ws = getDataSheet(wb);
      appendRow(ws, LEDGER_HEADERS, {
        NAME: name,
        DATE: installDate || "",
        ADDRESS: address || "",
        "PHONE NUMBER": phone,
        "MODEL MECHINE": machineModel || "",
        PRICE: price ?? "",
        PLACE: place || "",
        services: "",
        FEEDBACK: remarks || "",
      });
      await saveWorkbook(wb, file);
    });

    await writeAudit({ type: "excel_append", file: "ledger", customerId, uid: req.uid, success: true });
    res.json({ customerId, excelSynced: true });
  } catch (e) {
    console.error("Ledger append failed:", e);
    await writeAudit({ type: "excel_append", file: "ledger", customerId, uid: req.uid, success: false, error: e.message });
    res.status(502).json({ customerId, excelSynced: false, error: e.message });
  }
});

// POST /service-clients
app.post("/service-clients", async (req, res) => {
  const { customerId, customerName, phone, address, serviceDate, serviceDescription, amount } = req.body || {};
  if (!customerName || !phone) return res.status(400).json({ error: "customerName and phone are required" });

  try {
    await withFileLock("flat", async () => {
      const { wb, file } = await loadWorkbook(FLAT_PATH);
      const ws = getDataSheet(wb);
      appendRow(ws, FLAT_HEADERS, {
        "Customer Name": customerName,
        "Phone No 1": phone,
        Address: address || "",
        "Service Date": serviceDate || "",
        "Service Description": serviceDescription || "",
        Amount: amount ?? "",
      });
      await saveWorkbook(wb, file);
    });

    await writeAudit({ type: "excel_append", file: "flat", customerId, uid: req.uid, success: true });
    res.json({ customerId, excelSynced: true });
  } catch (e) {
    console.error("Flat append failed:", e);
    await writeAudit({ type: "excel_append", file: "flat", customerId, uid: req.uid, success: false, error: e.message });
    res.status(502).json({ customerId, excelSynced: false, error: e.message });
  }
});

// DELETE /customers/:id
app.delete("/customers/:id", async (req, res) => {
  const { id } = req.params;
  const { phone, file: fileTarget } = req.body || {}; // fileTarget: "ledger" | "flat" | "both"
  if (!phone) return res.status(400).json({ error: "phone is required to locate the row to delete" });

  const targets = fileTarget === "both" ? ["ledger", "flat"] : [fileTarget || "ledger"];
  const results = {};

  for (const target of targets) {
    const path = target === "flat" ? FLAT_PATH : LEDGER_PATH;
    const phoneCols = target === "flat" ? ["Phone No 1", "Phone Number"] : ["PHONE NUMBER"];
    try {
      await withFileLock(target, async () => {
        const { wb, file } = await loadWorkbook(path);
        const ws = getDataSheet(wb);
        const removed = removeRowByPhone(ws, phone, phoneCols);
        if (removed) await saveWorkbook(wb, file);
        results[target] = { removed };
      });
      await writeAudit({ type: "excel_delete", file: target, customerId: id, uid: req.uid, success: true, removed: results[target]?.removed });
    } catch (e) {
      console.error(`${target} delete failed:`, e);
      results[target] = { removed: false, error: e.message };
      await writeAudit({ type: "excel_delete", file: target, customerId: id, uid: req.uid, success: false, error: e.message });
    }
  }

  const anyFailed = Object.values(results).some((r) => r.error);
  res.status(anyFailed ? 502 : 200).json({ deleted: true, excelSynced: !anyFailed, results });
});

exports.api = onRequest({ region: "us-central1", timeoutSeconds: 60, memory: "512MiB" }, app);
