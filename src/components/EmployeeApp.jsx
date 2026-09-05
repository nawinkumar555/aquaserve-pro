import { useEffect, useState } from "react";
import { db, auth } from "../firebase";
import { collection, query, where, onSnapshot, doc, updateDoc, getDocs, setDoc, getDoc } from "firebase/firestore";
import { signOut } from "firebase/auth";
import JobCompletionModal from '../components/JobCompletionModal';

/* ── SVG icons ── */
const Icon = ({ d, size = 16, sw = 1.8 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor"
    strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
    <path d={d} />
  </svg>
);
const Ic = {
  logout:  (p) => <Icon {...p} d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9" />,
  check:   (p) => <Icon {...p} d="M20 6L9 17l-5-5" />,
  wrench:  (p) => <Icon {...p} d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z" />,
  clock:   (p) => <Icon {...p} d="M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 6v6l4 2" />,
  user:    (p) => <Icon {...p} d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2M12 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8z" />,
  map:     (p) => <Icon {...p} d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0zM12 10a1 1 0 1 0 0-2 1 1 0 0 0 0 2z" />,
  droplet: (p) => <Icon {...p} d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z" />,
  briefcase:(p) => <Icon {...p} d="M20 7H4a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2zM16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2" />,
  chevron: (p) => <Icon {...p} d="M9 18l6-6-6-6" />,
  sun:     (p) => <Icon {...p} d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42M12 5a7 7 0 1 0 0 14A7 7 0 0 0 12 5z" />,
};

/* ════════════════════════ COMPONENT ════════════════════════ */
const EmployeeApp = ({ userDetails }) => {
  const [currentTask, setCurrentTask]     = useState([]);
  const [attendanceMarked, setAttendanceMarked] = useState(false);
  const [loading, setLoading]             = useState(true);
  const [isModalOpen, setIsModalOpen]     = useState(false);
  const [activeJob, setActiveJob]         = useState(null);
  const [time, setTime]                   = useState(new Date());

  const rawName  = userDetails?.name || "Technician";
  const staffName = rawName.toLowerCase().split(" ").map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(" ");
  const initials  = staffName.split(" ").map(w => w[0]).join("").slice(0, 2);
  const today     = new Date().toISOString().split("T")[0];

  /* live clock */
  useEffect(() => { const t = setInterval(() => setTime(new Date()), 1000); return () => clearInterval(t); }, []);

  useEffect(() => {
    if (!staffName) return;

    const initStaffStatus = async () => {
      // Check if employee has marked attendance TODAY
      const attSnap = await getDoc(doc(db, "attendance", `${staffName}_${today}`));
      const hasAttendanceToday = attSnap.exists();

      // Always sync staffStatus to reflect today's attendance state:
      // If no attendance today → force Offline, regardless of what's stored
      const staffSnap = await getDocs(query(collection(db, "staffStatus"), where("name", "==", staffName)));
      staffSnap.forEach(async (d) => {
        const currentStatus = d.data().status;
        if (!hasAttendanceToday && currentStatus === "Online") {
          // New day, not checked in yet → reset to Offline
          await updateDoc(doc(db, "staffStatus", d.id), { status: "Offline" });
        }
      });

      setAttendanceMarked(hasAttendanceToday);
    };

    initStaffStatus();

    // Live listener for assigned tasks
    const q = query(collection(db, "staffStatus"), where("name", "==", staffName));
    const unsub = onSnapshot(q, (snap) => {
      if (!snap.empty) setCurrentTask(snap.docs[0].data().currentCustomers || []);
      else setCurrentTask([]);
      setLoading(false);
    });
    return () => unsub();
  }, [staffName, today]);

  const handleAttendance = async () => {
    try {
      // Write attendance record for today
      await setDoc(doc(db, "attendance", `${staffName}_${today}`), {
        name: staffName, date: today,
        punchIn: new Date().toLocaleTimeString(), status: "Present",
      });
      // NOW set staffStatus to Online — only happens when Mark Present is tapped
      const snap = await getDocs(query(collection(db, "staffStatus"), where("name", "==", staffName)));
      snap.forEach(async d => await updateDoc(doc(db, "staffStatus", d.id), { status: "Online" }));
      setAttendanceMarked(true);
    } catch { alert("Error marking attendance."); }
  };

  const openCompleteModal = (job) => { setActiveJob(job); setIsModalOpen(true); };

  const dateStr = new Date().toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long" });
  const timeStr = time.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit", hour12: true });

  /* ── Loading screen ── */
  if (loading) return (
    <div style={{ minHeight: "100vh", background: "#070d1a", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ width: 48, height: 48, border: "3px solid rgba(59,130,246,.25)", borderTopColor: "#3b82f6", borderRadius: "50%", animation: "spin .8s linear infinite", margin: "0 auto 20px" }} />
        <p style={{ color: "#64748b", fontSize: 12, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase" }}>Syncing your dashboard…</p>
      </div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );

  /* ════════════ RENDER ════════════ */
  return (
    <div style={{ minHeight: "100vh", background: "#070d1a", color: "#e8f0fe", fontFamily: "'Inter',system-ui,sans-serif" }}>

      {/* ── TOP NAV ── */}
      <header style={{
        position: "sticky", top: 0, zIndex: 30,
        background: "rgba(7,13,26,.92)", backdropFilter: "blur(16px)",
        borderBottom: "1px solid rgba(255,255,255,.07)",
        padding: "0 24px", height: 64,
        display: "flex", alignItems: "center", justifyContent: "space-between",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 34, height: 34, borderRadius: 10, background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 4px 12px rgba(99,102,241,.5)" }}>
            <Ic.droplet size={16} style={{ color: "#fff" }} />
          </div>
          <div>
            <p style={{ fontWeight: 800, fontSize: 14, letterSpacing: "-0.02em", lineHeight: 1.1 }}>AquaServe Pro</p>
            <p style={{ fontSize: 10, color: "#3b82f6", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase" }}>Field Technician</p>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ textAlign: "right", display: "none" }} className="hide-mobile">
            <p style={{ fontSize: 13, fontWeight: 700, color: "#e8f0fe" }}>{staffName}</p>
            <p style={{ fontSize: 11, color: "#64748b" }}>Technician</p>
          </div>
          <div style={{ width: 38, height: 38, borderRadius: "50%", background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 14, boxShadow: "0 4px 12px rgba(99,102,241,.4)" }}>
            {initials}
          </div>
          <button onClick={() => signOut(auth)}
            style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: "1px solid rgba(255,255,255,.1)", borderRadius: 10, padding: "8px 14px", color: "#64748b", fontSize: 12, fontWeight: 600, cursor: "pointer", transition: "all .15s" }}
            onMouseEnter={e => { e.currentTarget.style.borderColor = "#ef4444"; e.currentTarget.style.color = "#ef4444"; }}
            onMouseLeave={e => { e.currentTarget.style.borderColor = "rgba(255,255,255,.1)"; e.currentTarget.style.color = "#64748b"; }}>
            <Ic.logout size={13} /> Logout
          </button>
        </div>
      </header>

      <main style={{ maxWidth: 680, margin: "0 auto", padding: "28px 20px", display: "flex", flexDirection: "column", gap: 20 }}>

        {/* ── HERO GREETING CARD ── */}
        <div style={{
          borderRadius: 24, overflow: "hidden", position: "relative",
          background: "linear-gradient(135deg,#1d2d50 0%,#1a1f3a 60%,#0f172a 100%)",
          border: "1px solid rgba(99,102,241,.25)",
          boxShadow: "0 20px 60px rgba(0,0,0,.5), inset 0 1px 0 rgba(255,255,255,.08)",
          padding: "28px 28px 24px",
        }}>
          {/* decorative blobs */}
          <div style={{ position: "absolute", top: -30, right: -30, width: 140, height: 140, borderRadius: "50%", background: "radial-gradient(circle,rgba(99,102,241,.25) 0%,transparent 70%)", pointerEvents: "none" }} />
          <div style={{ position: "absolute", bottom: -20, left: 60, width: 100, height: 100, borderRadius: "50%", background: "radial-gradient(circle,rgba(59,130,246,.2) 0%,transparent 70%)", pointerEvents: "none" }} />

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", position: "relative" }}>
            <div>
              <p style={{ fontSize: 12, color: "rgba(148,163,184,.8)", fontWeight: 600, marginBottom: 6, letterSpacing: "0.04em" }}>Good {time.getHours() < 12 ? "Morning" : time.getHours() < 17 ? "Afternoon" : "Evening"},</p>
              <h1 style={{ fontSize: 26, fontWeight: 900, letterSpacing: "-0.03em", lineHeight: 1.1, marginBottom: 4, background: "linear-gradient(135deg,#fff,#93c5fd)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>
                {staffName}
              </h1>
              <p style={{ fontSize: 13, color: "#64748b", marginTop: 6 }}>{dateStr}</p>
            </div>
            <div style={{ textAlign: "right" }}>
              <p style={{ fontSize: 28, fontWeight: 900, letterSpacing: "-0.03em", color: "#e8f0fe", lineHeight: 1 }}>{timeStr}</p>
              <div style={{ display: "inline-flex", alignItems: "center", gap: 5, marginTop: 6, background: attendanceMarked ? "rgba(16,185,129,.15)" : "rgba(239,68,68,.15)", padding: "4px 12px", borderRadius: 999, border: `1px solid ${attendanceMarked ? "rgba(16,185,129,.3)" : "rgba(239,68,68,.3)"}` }}>
                <span style={{ width: 6, height: 6, borderRadius: "50%", background: attendanceMarked ? "#10b981" : "#ef4444", boxShadow: attendanceMarked ? "0 0 0 3px rgba(16,185,129,.2)" : "0 0 0 3px rgba(239,68,68,.2)" }} />
                <span style={{ fontSize: 10, fontWeight: 700, color: attendanceMarked ? "#10b981" : "#ef4444", letterSpacing: "0.06em", textTransform: "uppercase" }}>
                  {attendanceMarked ? "Online" : "Not Checked In"}
                </span>
              </div>
            </div>
          </div>

          {/* stats row */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginTop: 24 }}>
            {[
              { label: "Assigned Jobs", val: currentTask.length, col: "#3b82f6" },
              { label: "Status", val: attendanceMarked ? "Present" : "Absent", col: attendanceMarked ? "#10b981" : "#ef4444" },
              { label: "Today", val: new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short" }), col: "#8b5cf6" },
            ].map((s, i) => (
              <div key={i} style={{ background: "rgba(255,255,255,.05)", borderRadius: 14, padding: "12px 14px", border: "1px solid rgba(255,255,255,.07)" }}>
                <p style={{ fontSize: 10, color: "#64748b", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 4 }}>{s.label}</p>
                <p style={{ fontSize: 18, fontWeight: 900, color: s.col }}>{s.val}</p>
              </div>
            ))}
          </div>
        </div>

        {/* ── ATTENDANCE CARD ── */}
        <div style={{
          background: "#111827", border: "1px solid rgba(255,255,255,.07)",
          borderRadius: 20, padding: 24,
          boxShadow: "0 4px 24px rgba(0,0,0,.4)",
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18, paddingBottom: 16, borderBottom: "1px solid rgba(255,255,255,.07)" }}>
            <div style={{ width: 36, height: 36, borderRadius: 10, background: "linear-gradient(135deg,#10b981,#06b6d4)", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Ic.clock size={16} style={{ color: "#fff" }} />
            </div>
            <div>
              <p style={{ fontWeight: 800, fontSize: 14 }}>Morning Check-In</p>
              <p style={{ fontSize: 11, color: "#64748b", marginTop: 1 }}>Mark your attendance to go online</p>
            </div>
          </div>

          {!attendanceMarked ? (
            <button onClick={handleAttendance}
              style={{ width: "100%", padding: "16px", background: "linear-gradient(135deg,#10b981,#06b6d4)", color: "#fff", border: "none", borderRadius: 14, fontWeight: 800, fontSize: 15, cursor: "pointer", letterSpacing: "0.02em", boxShadow: "0 8px 24px rgba(16,185,129,.4)", transition: "opacity .15s", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
              onMouseEnter={e => e.currentTarget.style.opacity = ".88"} onMouseLeave={e => e.currentTarget.style.opacity = "1"}>
              <Ic.check size={18} /> Mark Present
            </button>
          ) : (
            <div style={{ background: "rgba(16,185,129,.1)", border: "1px solid rgba(16,185,129,.3)", borderRadius: 14, padding: "16px", textAlign: "center", display: "flex", alignItems: "center", justifyContent: "center", gap: 10 }}>
              <div style={{ width: 32, height: 32, borderRadius: "50%", background: "rgba(16,185,129,.2)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Ic.check size={16} style={{ color: "#10b981" }} />
              </div>
              <div style={{ textAlign: "left" }}>
                <p style={{ fontWeight: 800, color: "#10b981", fontSize: 14 }}>Present Today</p>
                <p style={{ fontSize: 11, color: "rgba(16,185,129,.7)" }}>You're marked online & active</p>
              </div>
            </div>
          )}
        </div>

        {/* ── ASSIGNMENTS CARD ── */}
        <div style={{
          background: "#111827", border: "1px solid rgba(255,255,255,.07)",
          borderRadius: 20, padding: 24,
          boxShadow: "0 4px 24px rgba(0,0,0,.4)",
        }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18, paddingBottom: 16, borderBottom: "1px solid rgba(255,255,255,.07)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 36, height: 36, borderRadius: 10, background: "linear-gradient(135deg,#3b82f6,#6366f1)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Ic.briefcase size={16} style={{ color: "#fff" }} />
              </div>
              <div>
                <p style={{ fontWeight: 800, fontSize: 14 }}>Active Assignments</p>
                <p style={{ fontSize: 11, color: "#64748b", marginTop: 1 }}>Jobs assigned to you today</p>
              </div>
            </div>
            {currentTask.length > 0 && (
              <span style={{ background: "rgba(59,130,246,.15)", border: "1px solid rgba(59,130,246,.3)", color: "#60a5fa", padding: "4px 12px", borderRadius: 999, fontSize: 12, fontWeight: 800 }}>
                {currentTask.length} Job{currentTask.length > 1 ? "s" : ""}
              </span>
            )}
          </div>

          {currentTask.length > 0 ? (
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              {currentTask.map((job, i) => (
                <div key={job.id} style={{
                  background: "#0b1120", border: "1px solid rgba(59,130,246,.2)",
                  borderRadius: 16, padding: "18px 18px",
                  display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12,
                  transition: "border-color .15s, box-shadow .15s",
                }}
                  onMouseEnter={e => { e.currentTarget.style.borderColor = "rgba(99,102,241,.5)"; e.currentTarget.style.boxShadow = "0 4px 20px rgba(59,130,246,.15)"; }}
                  onMouseLeave={e => { e.currentTarget.style.borderColor = "rgba(59,130,246,.2)"; e.currentTarget.style.boxShadow = "none"; }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                    <div style={{ width: 42, height: 42, borderRadius: 12, background: "linear-gradient(135deg,#1e3a5f,#1e40af)", border: "1px solid rgba(59,130,246,.3)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                      <span style={{ fontWeight: 900, fontSize: 16, color: "#60a5fa" }}>{i + 1}</span>
                    </div>
                    <div>
                      <p style={{ fontSize: 10, color: "#3b82f6", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 3 }}>Job #{i + 1}</p>
                      <p style={{ fontWeight: 700, fontSize: 15, color: "#e8f0fe", marginBottom: 3 }}>{job.name}</p>
                      {job.address && (
                        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                          <Ic.map size={11} style={{ color: "#64748b" }} />
                          <p style={{ fontSize: 11, color: "#64748b" }}>{job.address}</p>
                        </div>
                      )}
                    </div>
                  </div>
                  <button onClick={() => openCompleteModal(job)}
                    style={{ flexShrink: 0, background: "linear-gradient(135deg,#3b82f6,#6366f1)", color: "#fff", border: "none", borderRadius: 12, padding: "10px 18px", fontSize: 12, fontWeight: 800, cursor: "pointer", letterSpacing: "0.04em", textTransform: "uppercase", boxShadow: "0 4px 14px rgba(99,102,241,.45)", transition: "opacity .15s", display: "flex", alignItems: "center", gap: 7 }}
                    onMouseEnter={e => e.currentTarget.style.opacity = ".85"} onMouseLeave={e => e.currentTarget.style.opacity = "1"}>
                    <Ic.check size={13} /> Done
                  </button>
                </div>
              ))}
            </div>
          ) : (
            <div style={{ padding: "40px 20px", textAlign: "center" }}>
              <div style={{ width: 56, height: 56, borderRadius: "50%", background: "rgba(99,102,241,.1)", border: "1px dashed rgba(99,102,241,.3)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
                <Ic.briefcase size={22} style={{ color: "rgba(99,102,241,.5)" }} />
              </div>
              <p style={{ fontWeight: 700, fontSize: 14, color: "#475569", marginBottom: 4 }}>No jobs assigned yet</p>
              <p style={{ fontSize: 12, color: "#334155" }}>Your admin will assign tasks to you shortly.</p>
            </div>
          )}
        </div>

      </main>

      <JobCompletionModal
        isOpen={isModalOpen}
        onClose={() => setIsModalOpen(false)}
        job={{ ...activeJob, employeeName: staffName }}
      />

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { -webkit-font-smoothing: antialiased; }
        @keyframes spin { to { transform: rotate(360deg); } }
        ::-webkit-scrollbar { width: 4px; }
        ::-webkit-scrollbar-thumb { background: #1e293b; border-radius: 99px; }
      `}</style>
    </div>
  );
};

export default EmployeeApp;