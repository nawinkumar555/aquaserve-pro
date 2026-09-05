import React, { useEffect, useRef, useState } from 'react';
import { db, storage } from '../firebase';
import { collection, addDoc, query, where, getDocs, getDoc, updateDoc, deleteDoc, doc, increment, writeBatch, serverTimestamp, runTransaction } from 'firebase/firestore';
import { ref, uploadBytes, getDownloadURL } from 'firebase/storage';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

/* ── Inline icon set ── */
const Icon = ({ d, size = 16, sw = 2, className = '' }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor"
    strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" className={className} style={{ flexShrink: 0 }}>
    <path d={d} />
  </svg>
);
const Ic = {
  close:  (p) => <Icon {...p} d="M18 6 6 18M6 6l12 12" />,
  wrench: (p) => <Icon {...p} d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z" />,
  image:  (p) => <Icon {...p} d="M19 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2zM8.5 10a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3zM21 15l-5-5L5 21" />,
  plus:   (p) => <Icon {...p} d="M12 5v14M5 12h14" />,
  trash:  (p) => <Icon {...p} d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2m3 0-1 14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2L4 6h16z" />,
  cash:   (p) => <Icon {...p} d="M2 7h20v10H2zM12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6zM6 7v.01M18 17v-.01" />,
  upi:    (p) => <Icon {...p} d="M5 2h14a1 1 0 0 1 1 1v18a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V3a1 1 0 0 1 1-1zM12 18h.01" />,
  card:   (p) => <Icon {...p} d="M1 4h22v16H1zM1 10h22" />,
  bank:   (p) => <Icon {...p} d="M3 21h18M4 21V9l8-6 8 6v12M9 21V12h6v9" />,
  box:    (p) => <Icon {...p} d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16zM3.27 6.96 12 12l8.73-5.04M12 22.08V12" />,
  check:  (p) => <Icon {...p} d="M20 6 9 17l-5-5" />,
  loader: (p) => <Icon {...p} d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83" />,
  pin:    (p) => <Icon {...p} d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0zM12 10a1 1 0 1 0 0-2 1 1 0 0 0 0 2z" />,
  tool:   (p) => <Icon {...p} d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z" />,
  shield: (p) => <Icon {...p} d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />,
};

const PAYMENT_METHODS = [
  { value: 'Cash',          label: 'Cash',        icon: Ic.cash },
  { value: 'GPay',          label: 'GPay / UPI',  icon: Ic.upi },
  { value: 'Card',          label: 'Card',        icon: Ic.card },
  { value: 'Bank Transfer', label: 'Bank Transfer', icon: Ic.bank },
];

const INVENTORY_ITEMS = [
  "Membrane (75/80/100 GPD)", "Precarbon", "Sediment Carbon", "Post Carbon", "G.A.C. Carbon",
  "C.T.O. Carbon", "10 inch Spun 120gm", "10 inch Spun 170gm", "10 inch Spun 90gm",
  "Adopter (3 types)", "Float Dolphin", "Float Grand", "U.F. filter", "U.V. filter",
  "Minerals filter", "Pre filter Bowl Normal", "Pre filter Bowl Heavy", "Pump Normal (100GPD)",
  "Pump Heavy (100GPD)", "Pump Booster (100GPD)", "Membrane Housing", "Taf lon",
  "2 Side Elbow", "Pro Filter Elbow", "Mem Housing Elbow", "Pump Elbow", "Stemp Elbow",
  "Straight Elbow", "I - Conneter 3/4", "I - Conneter 3/8", "Pre-filter Elbow 3/8",
  "Body Connecter", "Fish", "Tape Dolphin", "Tape Grand", "Tape Luxury model",
  "F.R. 450", "F.R. 500", "S.V. 24", "S.V. 36 SLX", "Hero S.V. 24", "Hero S.V. 36",
  "Pump Plate", "On/Off Wall", "Float wire", "D - Stand", "Waterilly-Stand", "JADE-Stand",
  "Mars-Stand", "Emerald-Stand", "Aqua Touch-Stand", "V.S.-Stand", "Innovica-Stand",
  "HJC 80 Silvernet", "HJC 100 A.V.", "OToya 80 metrox", "OToya 100 Coconut",
  "Max 75 APL", "Max 80", "A.V. 75", "A.V. 100"
];

const SectionLabel = ({ title, subtitle }) => (
  <div className="mb-3">
    <p className="text-[11px] font-bold text-slate-400 uppercase tracking-wider leading-tight">{title}</p>
    {subtitle && <p className="text-[11px] text-slate-500 leading-tight mt-0.5">{subtitle}</p>}
  </div>
);

const JobCompletionModal = ({ isOpen, onClose, job, refreshJobs }) => {
  const idCounter = useRef(0);
  const makeRow = () => ({ id: `item-${idCounter.current++}`, item: '', quantity: '', price: '' });

  const [issue, setIssue] = useState('');
  const [items, setItems] = useState(() => [makeRow()]);
  const [paymentMethod, setPaymentMethod] = useState('Cash');
  const [serviceCost, setServiceCost] = useState('');
  const [preview, setPreview] = useState(null);
  const [photoFile, setPhotoFile] = useState(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const fileInputRef = useRef(null);

  // ── NEW: warranty status for this customer, looked up when the modal opens ──
  const [warrantyInfo, setWarrantyInfo] = useState(null); // { pending, completed } | null
  const [useWarranty, setUseWarranty] = useState(false);

  const date = new Date().toISOString().split('T')[0];
  const dateLabel = new Date().toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });

  useEffect(() => () => { if (preview) URL.revokeObjectURL(preview); }, [preview]);

  // ── NEW: fetch the customer's warranty status by their doc id (job.id is
  // the customer's Firestore doc id, set when the admin assigns the job) ──
  useEffect(() => {
    if (!isOpen || !job?.id) { setWarrantyInfo(null); setUseWarranty(false); return; }
    let cancelled = false;
    setUseWarranty(false);
    (async () => {
      try {
        const snap = await getDoc(doc(db, "customers", job.id));
        if (cancelled) return;
        if (snap.exists() && typeof snap.data().warrantyPending === 'number') {
          const data = snap.data();
          setWarrantyInfo({ pending: data.warrantyPending, completed: data.warrantyServicesCompleted || 0 });
        } else {
          setWarrantyInfo(null);
        }
      } catch (e) {
        console.error("Warranty lookup error:", e);
        setWarrantyInfo(null);
      }
    })();
    return () => { cancelled = true; };
  }, [isOpen, job?.id]);

  const updateItem = (id, field, value) => setItems(prev => prev.map(row => (row.id === id ? { ...row, [field]: value } : row)));
  const addItemRow = () => setItems(prev => [...prev, makeRow()]);
  const removeItemRow = (id) => setItems(prev => prev.filter(row => row.id !== id));

  const rowTotal = (row) => (Number(row.quantity) || 0) * (Number(row.price) || 0);
  const itemsTotal = items.reduce((sum, row) => sum + rowTotal(row), 0);
  const serviceCostNum = Number(serviceCost) || 0;
  const grandTotal = itemsTotal + serviceCostNum;
  const itemCount = items.filter(row => row.item.trim() !== '').length;
  const fmt = (n) => n.toLocaleString('en-IN');

  // ── NEW: read the selected photo file as a base64 data URL so it can be embedded in the PDF ──
  const fileToDataURL = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

  // ── NEW: get an image's natural pixel dimensions so it can be scaled into the PDF without distortion ──
  const getImageDimensions = (dataUrl) => new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = reject;
    img.src = dataUrl;
  });

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (preview) URL.revokeObjectURL(preview);
    setPhotoFile(file);
    setPreview(URL.createObjectURL(file));
  };
  const clearPhoto = () => {
    if (preview) URL.revokeObjectURL(preview);
    setPreview(null);
    setPhotoFile(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  /* ── Deduct used quantities from Inventory Management stock ──
     Best-effort case-insensitive partial match against the "inventory"
     collection maintained in the Admin Dashboard. If no matching
     inventory item is found, that item is silently skipped.

     UPDATED: each deduction now also writes a matching entry to the
     "stockMovements" collection — the same collection the Admin
     Dashboard writes to for Excel imports, manual quantity edits, and
     Direct Store Sales. Without this, a job's parts usage correctly
     reduced the stock count but left no trace of *why* it dropped, so
     it never showed up in the Admin Dashboard's Inventory Stock
     Analysis chart — only the Store Sale side of stock movement was
     ever visible there. */
  const deductInventory = async (usedItems) => {
    try {
      const invSnap = await getDocs(collection(db, "inventory"));
      const invDocs = invSnap.docs.map(d => ({ id: d.id, ...d.data() }));
      if (!invDocs.length) return;

      const batch = writeBatch(db);
      let hasUpdates = false;

      for (const used of usedItems) {
        const usedName = (used.item || "").trim().toLowerCase();
        const qtyUsed = Number(used.quantity) || 0;
        if (!usedName || qtyUsed <= 0) continue;

        const match = invDocs.find(inv => {
          const invName = (inv.itemName || "").trim().toLowerCase();
          return invName && (invName.includes(usedName) || usedName.includes(invName));
        });

        if (match) {
          batch.update(doc(db, "inventory", match.id), { quantity: increment(-qtyUsed) });

          // ── NEW: log this deduction to stockMovements ──
          batch.set(doc(collection(db, "stockMovements")), {
            itemId: match.id,
            itemName: match.itemName,
            change: -qtyUsed,
            reason: `Service Job — ${job?.employeeName || "Unknown"}${job?.name ? `, ${job.name}` : ""}`,
            newQuantity: (Number(match.quantity) || 0) - qtyUsed,
            date: new Date().toISOString().split("T")[0],
            timestamp: serverTimestamp(),
            by: job?.employeeName || "unknown",
          });

          hasUpdates = true;
        }
        // else: item not found in inventory — silently skip, per requirement
      }

      if (hasUpdates) await batch.commit();
    } catch (err) {
      console.error("Inventory deduction error:", err);
    }
  };

  // ── NEW: mark one free warranty service as used on the customer's
  // record. Runs as a transaction so the derived counts (completed/pending)
  // stay correct even if two visits get logged close together, and mirrors
  // the same ledgerServiceHistory / allServiceHistory shape the rest of the
  // app already uses, so the Directory's history view and warranty count
  // stay consistent with this write instead of drifting from it. ──
  const markWarrantyServiceUsed = async () => {
    try {
      await runTransaction(db, async (tx) => {
        const custRef = doc(db, "customers", job.id);
        const snap = await tx.get(custRef);
        if (!snap.exists()) return;
        const data = snap.data();
        const newEntry = { dateISO: date, description: issue.trim() || "Warranty service visit" };
        const newLedgerHistory = [...(data.ledgerServiceHistory || []), newEntry];
        const newAllHistory = [
          ...newLedgerHistory.map(s => ({ ...s, source: "Warranty (Ledger)" })),
          ...(data.flatServiceHistory || []).map(s => ({ ...s, source: "Extra Service (Flat)" })),
        ].sort((a, b) => (a.dateISO || "").localeCompare(b.dateISO || ""));
        const newCompleted = newLedgerHistory.length;
        tx.update(custRef, {
          ledgerServiceHistory: newLedgerHistory,
          allServiceHistory: newAllHistory,
          warrantyServicesCompleted: newCompleted,
          warrantyPending: Math.max(0, 3 - newCompleted),
          lastServiceDate: date,
        });
      });
    } catch (err) {
      console.error("Warranty update error:", err);
    }
  };

  const handleComplete = async () => {
    if (isSubmitting) return;
    setIsSubmitting(true);
    const waWindow = window.open('', '_blank'); // open now, while it's still directly tied to the click
    try {
      const validItems = items.filter(row => row.item.trim() !== '');

      let customerPhone = "";
      const q = query(collection(db, "customers"), where("name", "==", job.name));
      const querySnapshot = await getDocs(q);

      if (!querySnapshot.empty) {
        customerPhone = querySnapshot.docs[0].data().phone || "";
      }

      const docPdf = new jsPDF();
      docPdf.setFontSize(18);
      docPdf.text(`Bill for ${job.name}`, 14, 20);
      docPdf.setFontSize(10);
      docPdf.text(`Date: ${date}`, 14, 30);
      docPdf.text(`Issue: ${issue}`, 14, 35);

      autoTable(docPdf, {
        startY: 45,
        head: [['Item Name', 'Qty', 'Price', 'Total']],
        body: [
          ...validItems.map(i => [i.item, i.quantity, i.price, rowTotal(i)]),
          ['Service Cost', '', '', serviceCostNum],
        ],
      });

      docPdf.text(`Payment Method: ${paymentMethod}`, 14, docPdf.lastAutoTable.finalY + 10);
      docPdf.text(`Grand Total: Rs. ${grandTotal}`, 14, docPdf.lastAutoTable.finalY + 15);

      // ── NEW: embed the uploaded Work Proof photo into the PDF ──
      if (photoFile) {
        try {
          const photoDataUrl = await fileToDataURL(photoFile);
          const { width: natW, height: natH } = await getImageDimensions(photoDataUrl);

          const pageWidth = docPdf.internal.pageSize.getWidth();
          const pageHeight = docPdf.internal.pageSize.getHeight();
          const maxWidth = pageWidth - 28;   // 14mm margin either side
          const maxHeight = pageHeight - 60; // leave room for heading

          const scale = Math.min(maxWidth / natW, maxHeight / natH, 1);
          const imgWidth = natW * scale;
          const imgHeight = natH * scale;

          const imgFormat = photoFile.type && photoFile.type.includes('png') ? 'PNG' : 'JPEG';

          docPdf.addPage();
          docPdf.setFontSize(14);
          docPdf.text('Work Proof', 14, 20);
          docPdf.addImage(photoDataUrl, imgFormat, 14, 30, imgWidth, imgHeight);
        } catch (imgErr) {
          console.error('Failed to embed work-proof photo in PDF:', imgErr);
        }
      }

      const pdfBlob = docPdf.output('blob');
      const fileName = `billproff/${job.name}_${Date.now()}.pdf`;
      const storageRef = ref(storage, fileName);
      await uploadBytes(storageRef, pdfBlob);
      const pdfUrl = await getDownloadURL(storageRef);

      await addDoc(collection(db, "completedJobs"), {
        customerName: job.name,
        employeeName: job.employeeName || "Unknown",
        date: date,
        timestamp: new Date(),
        items: validItems,
        issue: issue,
        paymentMethod: paymentMethod,
        serviceCost: serviceCostNum,
        grandTotal: grandTotal,
        pdfUrl: pdfUrl,
        status: "Completed",
        usedWarrantyService: !!(useWarranty && warrantyInfo?.pending > 0), // NEW: record whether this visit used a free warranty service
      });

      // ── Automatically minus the used items from Inventory Management,
      // and (NEW) log each deduction to stockMovements ──
      await deductInventory(validItems);

      // ── NEW: if the technician marked this as a free warranty visit,
      // update the customer's warranty count ──
      if (useWarranty && warrantyInfo?.pending > 0 && job?.id) {
        await markWarrantyServiceUsed();
      }

      const staffQ = query(collection(db, "staffStatus"));
      const staffSnapshot = await getDocs(staffQ);

      for (const staffDoc of staffSnapshot.docs) {
        const staffData = staffDoc.data();
        if (staffData.currentCustomers) {
          const updatedJobs = staffData.currentCustomers.filter(c => c.name !== job.name);
          await updateDoc(staffDoc.ref, { currentCustomers: updatedJobs });
        }
      }

      const taskQ = query(collection(db, "tasks"), where("customerName", "==", job.name), where("status", "==", "Assigned"));
      const taskSnapshot = await getDocs(taskQ);
      for (const taskDoc of taskSnapshot.docs) {
        await deleteDoc(taskDoc.ref);
      }

      if (pdfUrl) {
        const msg = `Hi ${job.name}, your service is complete. Please find your bill attached below:\n\n${pdfUrl}\n\nThank you for choosing Aquaserve Pro!`;
        const cleanPhone = customerPhone.toString().replace(/\D/g, '');

        if (cleanPhone.length >= 10) {
          const waUrl = `https://wa.me/${cleanPhone}?text=${encodeURIComponent(msg)}`;
          if (waWindow) waWindow.location.href = waUrl;
          else window.open(waUrl, '_blank');
        } else {
          if (waWindow) waWindow.close();
          alert(`Bill saved, but could not find a valid phone number for ${job.name}.`);
        }
      }

      if (typeof refreshJobs === 'function') refreshJobs();
      onClose();
    } catch (error) {
      if (waWindow) waWindow.close();
      console.error("Error completing job:", error);
      alert("Failed to complete bill. Please try again.");
    } finally {
      setIsSubmitting(false);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[9999] bg-black/80 backdrop-blur-sm flex items-center justify-center box-border p-4 sm:p-8">
      <div
        className="bg-[#0b1120] w-full max-w-2xl lg:max-w-4xl rounded-2xl shadow-[0_25px_70px_-15px_rgba(0,0,0,0.8)] border border-white/10 flex flex-col text-white overflow-hidden"
        style={{ height: 'min(780px, calc(100dvh - 4rem))' }}
      >

        {/* ── Header ── */}
        <div className="relative flex-shrink-0 px-6 sm:px-10 py-6 sm:py-7 border-b border-white/10 bg-gradient-to-br from-[#1d2d50] via-[#1a1f3a] to-[#0b1120]">
          <button onClick={onClose} title="Close"
            className="absolute top-5 right-5 sm:top-7 sm:right-7 w-10 h-10 rounded-full bg-white/10 hover:bg-white/20 border border-white/10 flex items-center justify-center text-slate-400 hover:text-white transition">
            <Ic.close size={20} />
          </button>
          <p className="text-[11px] font-black text-blue-400 uppercase tracking-widest mb-1">Job Completion</p>
          <h2 className="text-2xl sm:text-3xl font-black tracking-tight mb-2 pr-12 break-words">{job?.name || 'Customer'}</h2>
          <div className="flex items-center gap-3 text-[13px] text-slate-400 font-semibold">
            {job?.phone && <span>{job.phone}</span>}
            {job?.phone && <span className="text-slate-600">•</span>}
            <span>{dateLabel}</span>
          </div>
        </div>

        {/* ── Scrollable body ── */}
        <div className="flex-1 min-h-0 overflow-y-auto px-6 sm:px-10 py-6">
          <div className="rounded-2xl border border-slate-700 bg-[#111827] overflow-hidden">

            {/* ── NEW: Warranty Status ── */}
            {warrantyInfo && (
              <div className="p-6 border-b border-slate-800 flex flex-col gap-3">
                <SectionLabel title="Warranty Status" />
                {warrantyInfo.pending > 0 ? (
                  <label className="flex items-center gap-3 bg-[#0b1120] border border-slate-700 rounded-xl p-4 cursor-pointer hover:border-blue-500 transition">
                    <input
                      type="checkbox"
                      checked={useWarranty}
                      onChange={(e) => setUseWarranty(e.target.checked)}
                      className="w-5 h-5 accent-blue-500 flex-shrink-0"
                    />
                    <Ic.shield size={18} className="text-blue-400 flex-shrink-0" />
                    <span className="text-sm font-semibold text-slate-200">
                      Use 1 free warranty service for this visit
                      <span className="block text-[11px] font-normal text-slate-500 mt-0.5">
                        {warrantyInfo.pending} free service{warrantyInfo.pending > 1 ? 's' : ''} remaining
                      </span>
                    </span>
                  </label>
                ) : (
                  <div className="flex items-center gap-3 bg-[#0b1120] border border-slate-700 rounded-xl p-4">
                    <Ic.shield size={18} className="text-slate-500 flex-shrink-0" />
                    <span className="text-sm text-slate-400">Warranty fully used — no free services remaining.</span>
                  </div>
                )}
              </div>
            )}

            {/* Issue */}
            <div className="p-6 border-b border-slate-800 flex flex-col gap-3">
              <SectionLabel title="Service Issue" subtitle="Briefly explain the problem solved" />
              <textarea rows={3} placeholder="e.g. Membrane changed, motor booster checked..."
                className="w-full bg-[#0b1120] p-4 rounded-xl border border-slate-700 text-sm font-medium placeholder:text-slate-600 focus:outline-none focus:border-blue-500 transition resize-none"
                value={issue} onChange={(e) => setIssue(e.target.value)} />
            </div>

            {/* Photo */}
            <div className="p-6 border-b border-slate-800 flex flex-col gap-3">
              <SectionLabel title="Work Proof" subtitle="Attach photo of the completed job" />
              {preview ? (
                <div className="relative rounded-xl overflow-hidden border border-slate-700 h-48">
                  <img src={preview} alt="Work" className="w-full h-full object-cover" />
                  <button onClick={clearPhoto} className="absolute top-3 right-3 p-2 rounded-full bg-black/60 hover:bg-black/80 text-white"><Ic.close size={16} /></button>
                </div>
              ) : (
                <button onClick={() => fileInputRef.current.click()}
                  className="w-full py-8 bg-[#0b1120] border-2 border-dashed border-slate-700 hover:border-blue-500 rounded-xl flex flex-col items-center gap-3 text-slate-500 hover:text-blue-400 transition">
                  <Ic.image size={32} />
                  <span className="text-sm font-bold uppercase">Upload Photo</span>
                </button>
              )}
              <input type="file" ref={fileInputRef} onChange={handleFileChange} accept="image/*" className="hidden" />
            </div>

            {/* Items */}
            <div className="p-6 border-b border-slate-800 flex flex-col gap-3">
              <SectionLabel title="Parts & Inventory" subtitle="List all items used for the service" />
              <div className="flex flex-col gap-4">
                {items.map((row, index) => (
                  <div key={row.id} className="bg-[#0b1120] p-4 rounded-xl border border-slate-700 flex flex-col gap-3">
                    <div className="flex justify-between items-center text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                      <span>Item #{index + 1}</span>
                      {items.length > 1 && <button onClick={() => removeItemRow(row.id)} className="hover:text-red-400"><Ic.trash size={16} /></button>}
                    </div>
                    <input list="inventory-list" placeholder="Item name..." value={row.item}
                      onChange={(e) => updateItem(row.id, 'item', e.target.value)}
                      className="w-full bg-[#111827] p-3.5 rounded-lg border border-slate-700 text-sm focus:outline-none focus:border-blue-500 transition" />
                    <div className="grid grid-cols-2 gap-3">
                      <input type="number" placeholder="Quantity" value={row.quantity}
                        onChange={(e) => updateItem(row.id, 'quantity', e.target.value)}
                        className="bg-[#111827] p-3.5 rounded-lg border border-slate-700 text-sm focus:outline-none focus:border-blue-500 transition" />
                      <input type="number" placeholder="Price" value={row.price}
                        onChange={(e) => updateItem(row.id, 'price', e.target.value)}
                        className="bg-[#111827] p-3.5 rounded-lg border border-slate-700 text-sm focus:outline-none focus:border-blue-500 transition" />
                    </div>
                  </div>
                ))}
                <button onClick={addItemRow} className="w-full py-3.5 border border-dashed border-slate-600 text-blue-400 font-bold text-sm rounded-xl uppercase tracking-wide hover:border-blue-500 hover:bg-blue-500/5 transition">
                  + Add Item
                </button>
              </div>
            </div>

            {/* Service Cost */}
            <div className="p-6 border-b border-slate-800 flex flex-col gap-3">
              <SectionLabel title="Service Cost" subtitle="Labor / visit charge (separate from parts)" />
              <div className="relative">
                <span className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500 font-bold text-sm">₹</span>
                <input type="number" placeholder="0" value={serviceCost}
                  onChange={(e) => setServiceCost(e.target.value)}
                  className="w-full bg-[#0b1120] pl-9 pr-4 py-4 rounded-xl border border-slate-700 text-sm font-bold placeholder:text-slate-600 focus:outline-none focus:border-blue-500 transition" />
              </div>
            </div>

            {/* Payment Method */}
            <div className="p-6 flex flex-col gap-3">
              <SectionLabel title="Payment Method" />
              <div className="grid grid-cols-2 gap-3">
                {PAYMENT_METHODS.map((m) => (
                  <button key={m.value} onClick={() => setPaymentMethod(m.value)}
                    className={`flex items-center gap-3 px-4 py-4 rounded-xl border font-bold text-sm transition ${paymentMethod === m.value ? 'bg-blue-600 border-blue-500 text-white' : 'bg-[#0b1120] border-slate-700 text-slate-400'}`}>
                    <m.icon size={18} /> {m.label}
                  </button>
                ))}
              </div>
            </div>

          </div>
        </div>

        {/* Footer */}
        <div className="flex-shrink-0 px-6 sm:px-10 py-6 border-t border-white/10 bg-[#0f172a] flex items-center justify-between gap-6">
          <div className="text-left">
            <p className="text-[10px] font-bold uppercase opacity-60">Total Amount</p>
            <p className="text-3xl font-black text-blue-400">₹{fmt(grandTotal)}</p>
            {serviceCostNum > 0 && (
              <p className="text-[10px] text-slate-500 font-semibold mt-0.5">
                Parts ₹{fmt(itemsTotal)} + Service ₹{fmt(serviceCostNum)}
              </p>
            )}
          </div>
          <button onClick={handleComplete} disabled={isSubmitting} className="flex-1 py-4 bg-green-500 hover:bg-green-400 text-slate-950 rounded-2xl font-black text-sm uppercase transition flex items-center justify-center gap-2">
            {isSubmitting ? <Ic.loader className="animate-spin" /> : <><Ic.check /> Submit Bill</>}
          </button>
        </div>
      </div>
      <datalist id="inventory-list">{INVENTORY_ITEMS.map(i => <option key={i} value={i} />)}</datalist>
    </div>
  );
};
export default JobCompletionModal;