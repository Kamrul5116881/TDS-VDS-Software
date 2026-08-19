import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LayoutDashboard,
  BookOpen,
  ListFilter,
  FileBarChart,
  Plus,
  Search,
  Pencil,
  Trash2,
  X,
  Download,
  Printer,
  ChevronDown,
  ArrowUpDown,
  CircleDollarSign,
  Receipt,
  Landmark,
  Wallet,
  AlertCircle,
  Upload,
} from "lucide-react";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  PieChart,
  Pie,
  Cell,
} from "recharts";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------------------
   DESIGN TOKENS — "Ledger" system
   Paper ivory ground, deep ledger-green ink, brick-red rule (the classic
   double-entry cashbook margin line) as the one signature accent.
--------------------------------------------------------------------- */
const T = {
  paper: "#F6F2E9",
  paperDeep: "#EFE8D8",
  ink: "#1B2A29",
  inkSoft: "#4B5D5A",
  muted: "#8A8272",
  line: "#D9D0B8",
  card: "#FFFEFB",
  accent: "#0F6B4C",
  accentDeep: "#0B4F38",
  rule: "#B3472B",
  ruleSoft: "#E7CFC5",
  gold: "#B08D2B",
};

const FONT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
`;

const PAYMENT_TYPES = ["Advance", "Payment", "Final Payment"];
const CONDITIONS = [
  "Including",
  "Excluding",
  "TDS Including VDS Excluding",
  "N/A",
];

const emptyForm = {
  id: null,
  date: "",
  cheque: "",
  chequeDate: "",
  vendor: "",
  paymentType: "Payment",
  particular: "",
  vendorAddress: "",
  tin: "",
  bin: "",
  mainGL: "",
  subGL: "",
  sectionRef: "",
  invoiceAmount: "",
  condition: "Excluding",
  tdsRate: "",
  vdsRate: "",
  remarks: "",
};

/* ---------------------------------------------------------------------
   CALCULATION ENGINE — mirrors Work!R:AA exactly
--------------------------------------------------------------------- */
function calc(rec) {
  const N = parseFloat(rec.invoiceAmount) || 0;
  const P = parseFloat(rec.tdsRate) || 0; // fraction e.g. 0.075
  const Q = parseFloat(rec.vdsRate) || 0;
  const O = rec.condition;

  const R = O === "Including" ? null : 0; // placeholder, computed below w/ V
  let tdsIncluding = 0,
    tdsExcluding = 0,
    tdsInclVdsExcl = 0;
  let vdsIncluding = 0,
    vdsExcluding = 0,
    vdsInclVdsExcl = 0;

  // Excel evaluates V (VDS-Including) independent of R, and R depends on V —
  // both only fire in the "Including" branch, so compute together.
  if (O === "Including") {
    vdsIncluding = (N / (1 + Q)) * Q;
    tdsIncluding = (N - vdsIncluding) * P;
  }
  if (O === "Excluding") {
    tdsExcluding = (N / (1 - P)) * P;
    vdsExcluding = (N + tdsExcluding) * Q;
  }
  if (O === "TDS Including VDS Excluding") {
    tdsInclVdsExcl = N * P;
    vdsInclVdsExcl = N * Q;
  }

  const tdsTotal = tdsIncluding + tdsExcluding + tdsInclVdsExcl;
  const vdsTotal = vdsIncluding + vdsExcluding + vdsInclVdsExcl;

  let netPayment = 0;
  if (O === "Excluding") netPayment = N;
  else if (O === "TDS Including VDS Excluding") netPayment = N - tdsInclVdsExcl;
  else if (O === "Including") netPayment = N - tdsIncluding - vdsIncluding;
  else netPayment = N; // N/A — no withholding

  return {
    invoiceAmount: N,
    tdsAmount: round2(tdsTotal),
    vdsAmount: round2(vdsTotal),
    netPayment: round2(netPayment),
  };
}
function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}
function money(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("en-BD", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}
function fmtDate(d) {
  if (!d) return "—";
  const dt = new Date(d);
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}
function monthKey(d) {
  const dt = new Date(d);
  if (isNaN(dt)) return "Unknown";
  return dt.toLocaleDateString("en-US", { month: "short", year: "numeric" });
}

/* ---------------------------------------------------------------------
   EXCEL IMPORT — reads a workbook laid out like the real Work sheet
   (A Sl | B Date | C Cheque | D Cheque Date | E Vendor | F Payment Type |
    G Particular | H Vendor Address | I TIN | J BIN | K Main GL | L Sub-GL |
    M Section Ref | N Invoice Amount | O Condition | P TDS Rate | Q VDS Rate
    ... AD Remarks). Rows are detected by having a vendor name AND an
    invoice amount — header/blank rows are skipped automatically.
--------------------------------------------------------------------- */
function excelDateToStr(v) {
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getFullYear(),
      m = String(v.getMonth() + 1).padStart(2, "0"),
      d = String(v.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "string" && v.trim()) return v.trim();
  return "";
}
async function parseExcelFile(file, existingRecords) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames.includes("Work") ? "Work" : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  let sl = existingRecords.reduce((m, r) => Math.max(m, r.sl || 0), 0);
  const out = [];
  rows.forEach((row, i) => {
    const vendor = (row[4] || "").toString().trim();
    const invoice = parseFloat(row[13]);
    if (!vendor || !invoice) return; // skip header/blank/total rows
    sl += 1;
    out.push({
      id: `imp-${Date.now()}-${i}`,
      sl,
      date: excelDateToStr(row[1]),
      cheque: (row[2] || "").toString(),
      chequeDate: excelDateToStr(row[3]),
      vendor,
      paymentType: (row[5] || "Payment").toString().trim() || "Payment",
      particular: (row[6] || "").toString(),
      vendorAddress: (row[7] || "").toString(),
      tin: (row[8] || "").toString(),
      bin: (row[9] || "").toString(),
      mainGL: (row[10] || "").toString(),
      subGL: (row[11] || "").toString(),
      sectionRef: (row[12] || "").toString(),
      invoiceAmount: invoice,
      condition: (row[14] || "N/A").toString().trim() || "N/A",
      tdsRate: row[15] === "" || row[15] == null ? "" : parseFloat(row[15]),
      vdsRate: row[16] === "" || row[16] == null ? "" : parseFloat(row[16]),
      remarks: (row[29] || "").toString(),
    });
  });
  return out;
}

/* ---------------------------------------------------------------------
   SEED DATA — the 29 real rows read from the Work sheet
--------------------------------------------------------------------- */
const SEED = [
  [
    "2026-07-01",
    "2210003",
    "2026-07-01",
    "Rangs Industries Ltd",
    "Payment",
    "For the purpose of TV purchase for the admin block",
    "8768",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-01",
    "2210004",
    "2026-07-01",
    "AB Power Engineering Ltd.",
    "Payment",
    "",
    "470980",
    "Excluding",
    "",
    "",
  ],
  [
    "2026-07-02",
    "2201489",
    "2026-07-02",
    "GTCBL",
    "Payment",
    "Professional fees for the month of Jun'26",
    "687960",
    "Excluding",
    0.075,
    0.15,
  ],
  [
    "2026-07-02",
    "2201491",
    "2026-07-02",
    "M/S Hansa",
    "Payment",
    "Consultancy fee",
    "419230",
    "Excluding",
    0.075,
    0.15,
  ],
  [
    "2026-07-02",
    "2201492",
    "2026-07-02",
    "M/S Hansa",
    "Payment",
    "Consultancy fee",
    "68265",
    "Excluding",
    0.075,
    0.15,
  ],
  [
    "2026-07-02",
    "2201493",
    "2026-07-02",
    "Nazrul Islam",
    "Payment",
    "rent",
    "455625",
    "Excluding",
    0.1,
    0.15,
  ],
  [
    "2026-07-05",
    "2201494",
    "2026-07-05",
    "Unity Enterprise",
    "Payment",
    "Dining Table supply",
    "81000",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-05",
    "2201495",
    "2026-07-05",
    "Masud & Brothers",
    "Payment",
    "Paid mirazul islam for pureit supply",
    "32835",
    "Excluding",
    0.005,
    "",
  ],
  [
    "2026-07-06",
    "2201496",
    "2026-07-06",
    "Unity Enterprise",
    "Payment",
    "Dining Table supply",
    "27750",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-06",
    "2201500",
    "2026-07-08",
    "F. Rahman Construction",
    "Payment",
    "Supply of fixture and fittings",
    "70000",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-12",
    "2210007",
    "2026-07-12",
    "Bhai Bhai Fire Fighting Company",
    "Payment",
    "Supply of safety materials",
    "55650",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-12",
    "2210008",
    "2026-07-12",
    "Azad Ad",
    "Payment",
    "Supply of safety materials",
    "52280",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-13",
    "2210009",
    "2026-07-13",
    "Vai Vai Enterprise",
    "Payment",
    "Airline installation for air compressor",
    "12745",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-13",
    "2210010",
    "2026-07-13",
    "F. Rahman Construction",
    "Payment",
    "Civil work payment",
    "43000",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-19",
    "2210012",
    "2026-07-19",
    "Unity Enterprise",
    "Payment",
    "Supply of stationary Items, Hydraulic oil, PVC sign board etc",
    "397836",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-19",
    "2210013",
    "2026-07-19",
    "Binary Kraft",
    "Payment",
    "Admin Block – Roller Blinds and Glass sticker work – Completed",
    "106000",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-19",
    "2210015",
    "2026-07-19",
    "AB Power Engineering Ltd.",
    "Payment",
    "Remaining supply of electrical materials and labour charges",
    "291136",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-22",
    "2210017",
    "2026-07-22",
    "AB Power Engineering Ltd.",
    "Payment",
    "Electric materials supply",
    "91675",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-23",
    "2210018",
    "2026-07-23",
    "Fess Trade International",
    "Payment",
    "Floor cleaning machine purchase",
    "83200",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-23",
    "2210019",
    "2026-07-23",
    "M/S Hansa",
    "Payment",
    "Leather Purchase",
    "150652",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-23",
    "2210020",
    "2026-07-23",
    "M/S Leather Sewing and Accessories",
    "Payment",
    "machineries spare parts supplied",
    "16000",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-23",
    "2210021",
    "2026-07-23",
    "F. Rahman Construction",
    "Payment",
    "manpower hired during buyer visit",
    "10000",
    "Excluding",
    0.1,
    0.15,
  ],
  [
    "2026-07-23",
    "2210022",
    "2026-07-23",
    "M/S Ajad Ad",
    "Payment",
    "Safety sign board purchased.",
    "7960",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-23",
    "2210023",
    "2026-07-23",
    "K.A.R. Associates",
    "Payment",
    "Dehumidifier purchased for Hall-A materials stores",
    "143750",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-26",
    "",
    "",
    "Sikdar insurance PLC",
    "Payment",
    "Bill paid for marine insurance",
    "129528",
    "",
    "",
    "",
  ],
  [
    "2026-07-27",
    "2210025",
    "2026-07-27",
    "Transplace logistic Ltd.",
    "Payment",
    "Bill paid for Shipping agent",
    "488585",
    "Excluding",
    0.01,
    "",
  ],
  [
    "2026-07-28",
    "2210027",
    "",
    "Sentry Security Services Ltd",
    "Payment",
    "Security Service Bill for the month of June'26",
    "84096",
    "Excluding",
    0.02,
    "",
  ],
].map((r, i) => ({
  id: `seed-${i}`,
  sl: i + 1,
  date: r[0],
  cheque: r[1],
  chequeDate: r[2],
  vendor: r[3],
  paymentType: r[4],
  particular: r[5],
  vendorAddress: "",
  tin: "",
  bin: "",
  mainGL: "",
  subGL: "",
  sectionRef: "",
  invoiceAmount: r[6],
  condition: r[7] || "N/A",
  tdsRate: r[8] || "",
  vdsRate: r[9] || "",
  remarks: "",
}));

/* ---------------------------------------------------------------------
   STORAGE — writes to BOTH Claude's artifact storage (when available)
   AND the browser's localStorage on every save, and reads back from
   whichever one actually has data. This redundancy means a save
   survives even if one of the two mechanisms is unavailable or
   silently failing in a given environment (Claude preview vs. a local
   dev server vs. a browser with storage restrictions).
--------------------------------------------------------------------- */
const STORAGE_KEY = "ledger:transactions:v1";

function hasArtifactStorage() {
  return (
    typeof window !== "undefined" &&
    window.storage &&
    typeof window.storage.get === "function" &&
    typeof window.storage.set === "function"
  );
}
function hasLocalStorage() {
  try {
    const k = "__ledger_ls_probe__";
    window.localStorage.setItem(k, "1");
    window.localStorage.removeItem(k);
    return true;
  } catch (e) {
    return false;
  }
}

async function loadRecords() {
  let fromArtifact = null;
  if (hasArtifactStorage()) {
    try {
      const res = await window.storage.get(STORAGE_KEY);
      if (res && res.value) fromArtifact = JSON.parse(res.value);
    } catch (e) {
      /* key not present yet in artifact storage — not an error */
    }
  }
  let fromLocal = null;
  if (hasLocalStorage()) {
    try {
      const v = window.localStorage.getItem(STORAGE_KEY);
      if (v) fromLocal = JSON.parse(v);
    } catch (e) {
      console.error("localStorage read failed", e);
    }
  }
  if (fromArtifact !== null) return fromArtifact;
  if (fromLocal !== null) return fromLocal;
  return null;
}

async function saveRecords(records) {
  const payload = JSON.stringify(records);
  let savedSomewhere = false;
  if (hasArtifactStorage()) {
    try {
      await window.storage.set(STORAGE_KEY, payload, false);
      savedSomewhere = true;
    } catch (e) {
      console.error("artifact storage save failed:", e);
    }
  }
  if (hasLocalStorage()) {
    try {
      window.localStorage.setItem(STORAGE_KEY, payload);
      savedSomewhere = true;
    } catch (e) {
      console.error("localStorage save failed:", e);
    }
  }
  return savedSomewhere;
}

/* ---------------------------------------------------------------------
   SHARED UI BITS
--------------------------------------------------------------------- */
function Field({ label, children, span }) {
  return (
    <label
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 5,
        gridColumn: span ? `span ${span}` : undefined,
      }}
    >
      <span
        style={{
          fontFamily: "'IBM Plex Sans'",
          fontSize: 11.5,
          fontWeight: 600,
          letterSpacing: ".04em",
          textTransform: "uppercase",
          color: T.inkSoft,
        }}
      >
        {label}
      </span>
      {children}
    </label>
  );
}
const inputStyle = {
  fontFamily: "'IBM Plex Mono', monospace",
  fontSize: 14,
  color: T.ink,
  background: T.card,
  border: `1px solid ${T.line}`,
  borderRadius: 6,
  padding: "9px 11px",
  outline: "none",
  width: "100%",
  boxSizing: "border-box",
};
function TInput(props) {
  return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />;
}
function TSelect({ children, ...props }) {
  return (
    <div style={{ position: "relative" }}>
      <select
        {...props}
        style={{
          ...inputStyle,
          appearance: "none",
          paddingRight: 30,
          cursor: "pointer",
        }}
      >
        {children}
      </select>
      <ChevronDown
        size={14}
        color={T.muted}
        style={{
          position: "absolute",
          right: 10,
          top: 12,
          pointerEvents: "none",
        }}
      />
    </div>
  );
}
function StatCard({ icon: Icon, label, value, sub, accent }) {
  return (
    <div
      style={{
        background: T.card,
        border: `1px solid ${T.line}`,
        borderRadius: 10,
        padding: "18px 20px",
        display: "flex",
        flexDirection: "column",
        gap: 10,
        position: "relative",
        overflow: "hidden",
      }}
    >
      <div
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          bottom: 0,
          width: 3,
          background: accent || T.accent,
        }}
      />
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
          color: T.inkSoft,
        }}
      >
        <Icon size={15} strokeWidth={2} />
        <span
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 12,
            fontWeight: 600,
            letterSpacing: ".04em",
            textTransform: "uppercase",
          }}
        >
          {label}
        </span>
      </div>
      <div
        style={{
          fontFamily: "'Source Serif 4'",
          fontSize: 26,
          fontWeight: 600,
          color: T.ink,
        }}
      >
        {value}
      </div>
      {sub && (
        <div
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 12,
            color: T.muted,
          }}
        >
          {sub}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------
   MASTER DATA / ENTRY FORM
--------------------------------------------------------------------- */
function EntryForm({ records, initial, onSave, onCancel }) {
  const [form, setForm] = useState(initial || emptyForm);
  const isEdit = !!(initial && initial.id);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const nextSl = useMemo(() => {
    if (isEdit) return initial.sl;
    return records.reduce((m, r) => Math.max(m, r.sl || 0), 0) + 1;
  }, [records, isEdit, initial]);

  const preview = calc(form);
  const vendors = useMemo(
    () =>
      Array.from(new Set(records.map((r) => r.vendor).filter(Boolean))).sort(),
    [records],
  );
  const [vendorOpen, setVendorOpen] = useState(false);

  function submit(e) {
    e.preventDefault();
    if (!form.date || !form.vendor || !form.invoiceAmount) return;
    onSave({ ...form, sl: nextSl, id: isEdit ? form.id : `tx-${Date.now()}` });
  }

  return (
    <form
      onSubmit={submit}
      style={{
        background: T.card,
        border: `1px solid ${T.line}`,
        borderRadius: 12,
        padding: 26,
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "baseline",
          marginBottom: 20,
        }}
      >
        <div>
          <div
            style={{
              fontFamily: "'IBM Plex Sans'",
              fontSize: 11,
              fontWeight: 600,
              color: T.rule,
              letterSpacing: ".08em",
              textTransform: "uppercase",
            }}
          >
            Sl. No. {String(nextSl).padStart(4, "0")}
          </div>
          <h2
            style={{
              fontFamily: "'Source Serif 4'",
              fontSize: 22,
              margin: "2px 0 0",
              color: T.ink,
            }}
          >
            {isEdit ? "Edit entry" : "New ledger entry"}
          </h2>
        </div>
        {onCancel && (
          <button
            type="button"
            onClick={onCancel}
            style={{
              background: "none",
              border: "none",
              cursor: "pointer",
              color: T.muted,
            }}
          >
            <X size={20} />
          </button>
        )}
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(4, 1fr)",
          gap: 16,
        }}
      >
        <Field label="Date">
          <TInput
            type="date"
            value={form.date}
            onChange={set("date")}
            required
          />
        </Field>
        <Field label="Cheque No.">
          <TInput
            value={form.cheque}
            onChange={set("cheque")}
            placeholder="2210005"
          />
        </Field>
        <Field label="Cheque Date">
          <TInput
            type="date"
            value={form.chequeDate}
            onChange={set("chequeDate")}
          />
        </Field>
        <Field label="Payment Type">
          <TSelect value={form.paymentType} onChange={set("paymentType")}>
            {PAYMENT_TYPES.map((p) => (
              <option key={p} value={p}>
                {p}
              </option>
            ))}
          </TSelect>
        </Field>

        <Field label="Vendor Name" span={2}>
          <div style={{ position: "relative" }}>
            <TInput
              value={form.vendor}
              onChange={(e) => {
                set("vendor")(e);
                setVendorOpen(true);
              }}
              onFocus={() => setVendorOpen(true)}
              onBlur={() => setTimeout(() => setVendorOpen(false), 120)}
              placeholder="Search or type a vendor…"
              required
              autoComplete="off"
            />
            {vendorOpen && form.vendor && (
              <div
                style={{
                  position: "absolute",
                  top: "104%",
                  left: 0,
                  right: 0,
                  background: T.card,
                  border: `1px solid ${T.line}`,
                  borderRadius: 6,
                  maxHeight: 160,
                  overflowY: "auto",
                  zIndex: 5,
                  boxShadow: "0 6px 16px rgba(27,42,41,.12)",
                }}
              >
                {vendors
                  .filter(
                    (v) =>
                      v.toLowerCase().includes(form.vendor.toLowerCase()) &&
                      v !== form.vendor,
                  )
                  .slice(0, 6)
                  .map((v) => (
                    <div
                      key={v}
                      onMouseDown={() => setForm((f) => ({ ...f, vendor: v }))}
                      style={{
                        padding: "8px 12px",
                        fontFamily: "'IBM Plex Sans'",
                        fontSize: 13.5,
                        cursor: "pointer",
                        color: T.ink,
                      }}
                      onMouseEnter={(e) =>
                        (e.currentTarget.style.background = T.paperDeep)
                      }
                      onMouseLeave={(e) =>
                        (e.currentTarget.style.background = "transparent")
                      }
                    >
                      {v}
                    </div>
                  ))}
              </div>
            )}
          </div>
        </Field>
        <Field label="Particular" span={2}>
          <TInput
            value={form.particular}
            onChange={set("particular")}
            placeholder="Nature of the payment"
          />
        </Field>

        <Field label="Vendor Address" span={2}>
          <TInput value={form.vendorAddress} onChange={set("vendorAddress")} />
        </Field>
        <Field label="TIN">
          <TInput value={form.tin} onChange={set("tin")} />
        </Field>
        <Field label="BIN">
          <TInput value={form.bin} onChange={set("bin")} />
        </Field>

        <Field label="Main GL">
          <TInput value={form.mainGL} onChange={set("mainGL")} />
        </Field>
        <Field label="Sub-GL">
          <TInput value={form.subGL} onChange={set("subGL")} />
        </Field>
        <Field label="Section Ref">
          <TInput value={form.sectionRef} onChange={set("sectionRef")} />
        </Field>
        <Field label="Invoice Amount">
          <TInput
            type="number"
            step="0.01"
            value={form.invoiceAmount}
            onChange={set("invoiceAmount")}
            placeholder="500000"
            required
          />
        </Field>

        <Field label="TDS / VDS Condition">
          <TSelect value={form.condition} onChange={set("condition")}>
            {CONDITIONS.map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </TSelect>
        </Field>
        <Field label="TDS Rate">
          <TInput
            type="number"
            step="0.001"
            min="0"
            max="1"
            value={form.tdsRate}
            onChange={set("tdsRate")}
            placeholder="0.075 = 7.5%"
          />
        </Field>
        <Field label="VDS Rate">
          <TInput
            type="number"
            step="0.001"
            min="0"
            max="1"
            value={form.vdsRate}
            onChange={set("vdsRate")}
            placeholder="0.15 = 15%"
          />
        </Field>
        <Field label="Remarks">
          <TInput value={form.remarks} onChange={set("remarks")} />
        </Field>
      </div>

      {/* live calculation preview */}
      <div
        style={{
          marginTop: 22,
          borderTop: `1px dashed ${T.line}`,
          paddingTop: 18,
        }}
      >
        <div
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 11,
            fontWeight: 600,
            color: T.inkSoft,
            letterSpacing: ".05em",
            textTransform: "uppercase",
            marginBottom: 10,
          }}
        >
          Calculated automatically
        </div>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(4, 1fr)",
            gap: 14,
          }}
        >
          {[
            ["Invoice Amount", preview.invoiceAmount, T.ink],
            ["TDS Amount", preview.tdsAmount, T.rule],
            ["VDS Amount", preview.vdsAmount, T.rule],
            ["Net Payment", preview.netPayment, T.accent],
          ].map(([lbl, val, color]) => (
            <div
              key={lbl}
              style={{
                background: T.paperDeep,
                borderRadius: 8,
                padding: "12px 14px",
              }}
            >
              <div
                style={{
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  color: T.muted,
                }}
              >
                {lbl}
              </div>
              <div
                style={{
                  fontFamily: "'IBM Plex Mono'",
                  fontSize: 17,
                  fontWeight: 600,
                  color,
                }}
              >
                ৳ {money(val)}
              </div>
            </div>
          ))}
        </div>
      </div>

      <div
        style={{
          display: "flex",
          gap: 10,
          marginTop: 22,
          justifyContent: "flex-end",
        }}
      >
        {onCancel && (
          <button
            type="button"
            onClick={onCancel}
            style={{
              padding: "10px 18px",
              borderRadius: 7,
              border: `1px solid ${T.line}`,
              background: T.card,
              color: T.inkSoft,
              fontFamily: "'IBM Plex Sans'",
              fontSize: 13.5,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            Cancel
          </button>
        )}
        <button
          type="submit"
          style={{
            padding: "10px 22px",
            borderRadius: 7,
            border: "none",
            background: T.accent,
            color: "#fff",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13.5,
            fontWeight: 600,
            cursor: "pointer",
            display: "flex",
            alignItems: "center",
            gap: 8,
          }}
        >
          <Plus size={15} /> {isEdit ? "Save changes" : "Save entry"}
        </button>
      </div>
    </form>
  );
}

/* ---------------------------------------------------------------------
   TRANSACTIONS TABLE
--------------------------------------------------------------------- */
function Transactions({ records, onEdit, onDelete, onImport }) {
  const fileInputRef = React.useRef(null);
  const [importBusy, setImportBusy] = useState(false);
  const [q, setQ] = useState("");
  const [vendorF, setVendorF] = useState("");
  const [typeF, setTypeF] = useState("");
  const [condF, setCondF] = useState("");
  const [sort, setSort] = useState({ key: "sl", dir: "desc" });

  const vendors = useMemo(
    () =>
      Array.from(new Set(records.map((r) => r.vendor).filter(Boolean))).sort(),
    [records],
  );

  const filtered = useMemo(() => {
    let rows = records.map((r) => ({ ...r, ...calc(r) }));
    if (q) {
      const s = q.toLowerCase();
      rows = rows.filter((r) =>
        [r.vendor, r.particular, r.cheque, String(r.sl)].some((v) =>
          (v || "").toString().toLowerCase().includes(s),
        ),
      );
    }
    if (vendorF) rows = rows.filter((r) => r.vendor === vendorF);
    if (typeF) rows = rows.filter((r) => r.paymentType === typeF);
    if (condF) rows = rows.filter((r) => r.condition === condF);
    rows.sort((a, b) => {
      const { key, dir } = sort;
      let av = a[key],
        bv = b[key];
      if (key === "date") {
        av = new Date(av || 0).getTime();
        bv = new Date(bv || 0).getTime();
      }
      if (typeof av === "string") {
        av = (av || "").toLowerCase();
        bv = (bv || "").toLowerCase();
      }
      if (av < bv) return dir === "asc" ? -1 : 1;
      if (av > bv) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return rows;
  }, [records, q, vendorF, typeF, condF, sort]);

  const totals = useMemo(
    () =>
      filtered.reduce(
        (a, r) => ({
          invoice: a.invoice + r.invoiceAmount,
          tds: a.tds + r.tdsAmount,
          vds: a.vds + r.vdsAmount,
          net: a.net + r.netPayment,
        }),
        { invoice: 0, tds: 0, vds: 0, net: 0 },
      ),
    [filtered],
  );

  function toggleSort(key) {
    setSort((s) =>
      s.key === key
        ? { key, dir: s.dir === "asc" ? "desc" : "asc" }
        : { key, dir: "asc" },
    );
  }
  async function handleFileChosen(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = ""; // allow re-selecting the same file later
    if (!file) return;
    setImportBusy(true);
    try {
      const imported = await parseExcelFile(file, records);
      if (imported.length === 0) {
        window.alert(
          "No rows with both a vendor name and an invoice amount were found — check that the sheet matches the Work layout (Vendor in column E, Invoice Amount in column N).",
        );
        return;
      }
      const proceed = window.confirm(
        `Import ${imported.length} row(s) from "${file.name}"?\n\nThis will replace all data currently shown — the old data will not come back unless you import it again.`,
      );
      if (!proceed) return;
      onImport(imported);
    } catch (err) {
      window.alert(
        "Couldn't read that file — make sure it's a .xlsx or .xls export of your ledger.",
      );
      console.error(err);
    } finally {
      setImportBusy(false);
    }
  }
  function exportExcel() {
    const data = filtered.map((r) => ({
      Sl: r.sl,
      Date: r.date,
      Cheque: r.cheque,
      Vendor: r.vendor,
      "Payment Type": r.paymentType,
      Particular: r.particular,
      "Invoice Amount": r.invoiceAmount,
      Condition: r.condition,
      "TDS Rate": r.tdsRate,
      "VDS Rate": r.vdsRate,
      "TDS Amount": r.tdsAmount,
      "VDS Amount": r.vdsAmount,
      "Net Payment": r.netPayment,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Transactions");
    XLSX.writeFile(wb, "transactions.xlsx");
  }

  const Th = ({ label, k, align }) => (
    <th
      onClick={() => toggleSort(k)}
      style={{
        textAlign: align || "left",
        padding: "10px 12px",
        cursor: "pointer",
        userSelect: "none",
        fontFamily: "'IBM Plex Sans'",
        fontSize: 11,
        fontWeight: 600,
        letterSpacing: ".04em",
        textTransform: "uppercase",
        color: T.inkSoft,
        whiteSpace: "nowrap",
      }}
    >
      <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
        {label} {sort.key === k && <ArrowUpDown size={11} />}
      </span>
    </th>
  );

  return (
    <div>
      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: 10,
          marginBottom: 16,
          alignItems: "center",
        }}
      >
        <div style={{ position: "relative", flex: "1 1 220px" }}>
          <Search
            size={14}
            color={T.muted}
            style={{ position: "absolute", left: 11, top: 11 }}
          />
          <TInput
            placeholder="Search vendor, particular, cheque, Sl…"
            value={q}
            onChange={(e) => setQ(e.target.value)}
            style={{ paddingLeft: 32 }}
          />
        </div>
        <TSelect
          value={vendorF}
          onChange={(e) => setVendorF(e.target.value)}
          style={{ maxWidth: 190 }}
        >
          <option value="">All vendors</option>
          {vendors.map((v) => (
            <option key={v} value={v}>
              {v}
            </option>
          ))}
        </TSelect>
        <TSelect
          value={typeF}
          onChange={(e) => setTypeF(e.target.value)}
          style={{ maxWidth: 170 }}
        >
          <option value="">All payment types</option>
          {PAYMENT_TYPES.map((v) => (
            <option key={v} value={v}>
              {v}
            </option>
          ))}
        </TSelect>
        <TSelect
          value={condF}
          onChange={(e) => setCondF(e.target.value)}
          style={{ maxWidth: 220 }}
        >
          <option value="">All conditions</option>
          {CONDITIONS.map((v) => (
            <option key={v} value={v}>
              {v}
            </option>
          ))}
        </TSelect>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChosen}
          style={{ display: "none" }}
        />
        <button
          onClick={() => fileInputRef.current && fileInputRef.current.click()}
          disabled={importBusy}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: importBusy ? "wait" : "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Upload size={14} /> {importBusy ? "Reading…" : "Import Excel"}
        </button>
        <button
          onClick={exportExcel}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Download size={14} /> Export
        </button>
        <button
          onClick={() => window.print()}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Printer size={14} /> Print
        </button>
      </div>

      <div
        style={{
          background: T.card,
          border: `1px solid ${T.line}`,
          borderRadius: 12,
          overflow: "hidden",
        }}
      >
        <div style={{ overflowX: "auto" }}>
          <table
            style={{
              width: "100%",
              borderCollapse: "collapse",
              minWidth: 1100,
            }}
          >
            <thead>
              <tr style={{ borderBottom: `2px solid ${T.rule}` }}>
                <Th label="Sl" k="sl" />
                <Th label="Date" k="date" />
                <Th label="Cheque" k="cheque" />
                <Th label="Vendor" k="vendor" />
                <Th label="Type" k="paymentType" />
                <Th label="Particular" k="particular" />
                <Th label="Invoice" k="invoiceAmount" align="right" />
                <Th label="Condition" k="condition" />
                <Th label="TDS" k="tdsAmount" align="right" />
                <Th label="VDS" k="vdsAmount" align="right" />
                <Th label="Net Payment" k="netPayment" align="right" />
                <th style={{ padding: "10px 12px" }} />
              </tr>
            </thead>
            <tbody>
              {filtered.map((r, i) => (
                <tr
                  key={r.id}
                  style={{
                    borderBottom: `1px solid ${T.line}`,
                    background: i % 2 ? T.paper : T.card,
                  }}
                >
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 12.5,
                      color: T.muted,
                      borderLeft: `2px solid ${T.rule}`,
                    }}
                  >
                    {String(r.sl).padStart(3, "0")}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 13,
                    }}
                  >
                    {fmtDate(r.date)}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                    }}
                  >
                    {r.cheque || "—"}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 13,
                      fontWeight: 500,
                    }}
                  >
                    {r.vendor}
                  </td>
                  <td style={{ padding: "9px 12px" }}>
                    <span
                      style={{
                        fontFamily: "'IBM Plex Sans'",
                        fontSize: 11,
                        fontWeight: 600,
                        padding: "3px 8px",
                        borderRadius: 20,
                        background: T.paperDeep,
                        color: T.inkSoft,
                      }}
                    >
                      {r.paymentType}
                    </span>
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 12.5,
                      color: T.inkSoft,
                      maxWidth: 220,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                    title={r.particular}
                  >
                    {r.particular || "—"}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                    }}
                  >
                    {money(r.invoiceAmount)}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 12,
                    }}
                  >
                    {r.condition}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      color: T.rule,
                    }}
                  >
                    {money(r.tdsAmount)}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      color: T.rule,
                    }}
                  >
                    {money(r.vdsAmount)}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13.5,
                      textAlign: "right",
                      fontWeight: 600,
                      color: T.accentDeep,
                    }}
                  >
                    {money(r.netPayment)}
                  </td>
                  <td style={{ padding: "9px 12px" }}>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button
                        onClick={() => onEdit(r)}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          color: T.muted,
                        }}
                      >
                        <Pencil size={14} />
                      </button>
                      <button
                        onClick={() => onDelete(r.id)}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          color: T.rule,
                        }}
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr>
                  <td
                    colSpan={12}
                    style={{
                      padding: 40,
                      textAlign: "center",
                      fontFamily: "'IBM Plex Sans'",
                      color: T.muted,
                      fontSize: 13,
                    }}
                  >
                    {records.length === 0
                      ? "No entries yet. Add one from Master Data, or import your Excel file."
                      : "No entries match these filters."}
                  </td>
                </tr>
              )}
            </tbody>
            {filtered.length > 0 && (
              <tfoot>
                <tr
                  style={{
                    borderTop: `2px solid ${T.rule}`,
                    background: T.paperDeep,
                  }}
                >
                  <td
                    colSpan={6}
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 12,
                      fontWeight: 700,
                      color: T.ink,
                    }}
                  >
                    Total ({filtered.length} entries)
                  </td>
                  <td
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      fontWeight: 700,
                    }}
                  >
                    {money(totals.invoice)}
                  </td>
                  <td />
                  <td
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      fontWeight: 700,
                      color: T.rule,
                    }}
                  >
                    {money(totals.tds)}
                  </td>
                  <td
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      fontWeight: 700,
                      color: T.rule,
                    }}
                  >
                    {money(totals.vds)}
                  </td>
                  <td
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13.5,
                      textAlign: "right",
                      fontWeight: 700,
                      color: T.accentDeep,
                    }}
                  >
                    {money(totals.net)}
                  </td>
                  <td />
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   DASHBOARD
--------------------------------------------------------------------- */
function Dashboard({ records }) {
  const rows = useMemo(
    () => records.map((r) => ({ ...r, ...calc(r) })),
    [records],
  );
  const totals = rows.reduce(
    (a, r) => ({
      invoice: a.invoice + r.invoiceAmount,
      tds: a.tds + r.tdsAmount,
      vds: a.vds + r.vdsAmount,
      net: a.net + r.netPayment,
    }),
    { invoice: 0, tds: 0, vds: 0, net: 0 },
  );

  const byMonth = useMemo(() => {
    const m = {};
    rows.forEach((r) => {
      const k = monthKey(r.date);
      m[k] = (m[k] || 0) + r.netPayment;
    });
    return Object.entries(m).map(([month, net]) => ({ month, net }));
  }, [rows]);

  const byType = useMemo(() => {
    const m = {};
    rows.forEach((r) => {
      m[r.paymentType] = (m[r.paymentType] || 0) + r.netPayment;
    });
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [rows]);

  const byVendor = useMemo(() => {
    const m = {};
    rows.forEach((r) => {
      m[r.vendor] = (m[r.vendor] || 0) + r.netPayment;
    });
    return Object.entries(m)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6);
  }, [rows]);

  const pieColors = [T.accent, T.gold, T.rule];

  return (
    <div>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(4, 1fr)",
          gap: 14,
          marginBottom: 22,
        }}
      >
        <StatCard
          icon={Receipt}
          label="Total Invoice Amount"
          value={`৳ ${money(totals.invoice)}`}
          sub={`${records.length} transactions`}
        />
        <StatCard
          icon={Landmark}
          label="Total TDS"
          value={`৳ ${money(totals.tds)}`}
          accent={T.rule}
          sub="Withheld at source"
        />
        <StatCard
          icon={CircleDollarSign}
          label="Total VDS"
          value={`৳ ${money(totals.vds)}`}
          accent={T.rule}
          sub="Value Deducted at Source"
        />
        <StatCard
          icon={Wallet}
          label="Total Net Payment"
          value={`৳ ${money(totals.net)}`}
          accent={T.accent}
          sub="Disbursed to vendors"
        />
      </div>

      <div
        style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16 }}
      >
        <div
          style={{
            background: T.card,
            border: `1px solid ${T.line}`,
            borderRadius: 12,
            padding: 20,
          }}
        >
          <div
            style={{
              fontFamily: "'Source Serif 4'",
              fontSize: 16,
              fontWeight: 600,
              marginBottom: 12,
              color: T.ink,
            }}
          >
            Net payment by month
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={byMonth}>
              <CartesianGrid stroke={T.line} vertical={false} />
              <XAxis
                dataKey="month"
                tick={{
                  fontFamily: "IBM Plex Sans",
                  fontSize: 11,
                  fill: T.inkSoft,
                }}
                axisLine={{ stroke: T.line }}
                tickLine={false}
              />
              <YAxis
                tick={{
                  fontFamily: "IBM Plex Mono",
                  fontSize: 10,
                  fill: T.muted,
                }}
                axisLine={false}
                tickLine={false}
                width={70}
                tickFormatter={(v) => `৳${(v / 1000).toFixed(0)}k`}
              />
              <Tooltip
                formatter={(v) => `৳ ${money(v)}`}
                contentStyle={{
                  fontFamily: "IBM Plex Sans",
                  fontSize: 12,
                  border: `1px solid ${T.line}`,
                  borderRadius: 8,
                }}
              />
              <Bar dataKey="net" fill={T.accent} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div
          style={{
            background: T.card,
            border: `1px solid ${T.line}`,
            borderRadius: 12,
            padding: 20,
          }}
        >
          <div
            style={{
              fontFamily: "'Source Serif 4'",
              fontSize: 16,
              fontWeight: 600,
              marginBottom: 12,
              color: T.ink,
            }}
          >
            By payment type
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie
                data={byType}
                dataKey="value"
                nameKey="name"
                innerRadius={50}
                outerRadius={80}
                paddingAngle={2}
              >
                {byType.map((_, i) => (
                  <Cell key={i} fill={pieColors[i % pieColors.length]} />
                ))}
              </Pie>
              <Tooltip
                formatter={(v) => `৳ ${money(v)}`}
                contentStyle={{
                  fontFamily: "IBM Plex Sans",
                  fontSize: 12,
                  border: `1px solid ${T.line}`,
                  borderRadius: 8,
                }}
              />
            </PieChart>
          </ResponsiveContainer>
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 6,
              marginTop: 6,
            }}
          >
            {byType.map((t, i) => (
              <div
                key={t.name}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 12.5,
                  color: T.inkSoft,
                }}
              >
                <span
                  style={{
                    width: 9,
                    height: 9,
                    borderRadius: 9,
                    background: pieColors[i % pieColors.length],
                  }}
                />
                {t.name} — ৳ {money(t.value)}
              </div>
            ))}
          </div>
        </div>
      </div>

      <div
        style={{
          background: T.card,
          border: `1px solid ${T.line}`,
          borderRadius: 12,
          padding: 20,
          marginTop: 16,
        }}
      >
        <div
          style={{
            fontFamily: "'Source Serif 4'",
            fontSize: 16,
            fontWeight: 600,
            marginBottom: 14,
            color: T.ink,
          }}
        >
          Top vendors by net payment
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {byVendor.map(([name, val]) => {
            const pct = totals.net ? (val / totals.net) * 100 : 0;
            return (
              <div key={name}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    fontFamily: "'IBM Plex Sans'",
                    fontSize: 12.5,
                    marginBottom: 4,
                  }}
                >
                  <span style={{ color: T.ink, fontWeight: 500 }}>{name}</span>
                  <span
                    style={{ fontFamily: "'IBM Plex Mono'", color: T.inkSoft }}
                  >
                    ৳ {money(val)}
                  </span>
                </div>
                <div
                  style={{
                    height: 6,
                    borderRadius: 4,
                    background: T.paperDeep,
                    overflow: "hidden",
                  }}
                >
                  <div
                    style={{
                      height: "100%",
                      width: `${pct}%`,
                      background: T.accent,
                      borderRadius: 4,
                    }}
                  />
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   REPORTS
--------------------------------------------------------------------- */
function Reports({ records }) {
  const [groupBy, setGroupBy] = useState("vendor");
  const rows = useMemo(
    () => records.map((r) => ({ ...r, ...calc(r) })),
    [records],
  );

  const grouped = useMemo(() => {
    const m = {};
    rows.forEach((r) => {
      const key =
        groupBy === "vendor"
          ? r.vendor
          : groupBy === "paymentType"
            ? r.paymentType
            : groupBy === "month"
              ? monthKey(r.date)
              : r.condition;
      if (!m[key])
        m[key] = { key, count: 0, invoice: 0, tds: 0, vds: 0, net: 0 };
      m[key].count++;
      m[key].invoice += r.invoiceAmount;
      m[key].tds += r.tdsAmount;
      m[key].vds += r.vdsAmount;
      m[key].net += r.netPayment;
    });
    return Object.values(m).sort((a, b) => b.net - a.net);
  }, [rows, groupBy]);

  function exportExcel() {
    const ws = XLSX.utils.json_to_sheet(
      grouped.map((g) => ({
        Group: g.key,
        Entries: g.count,
        "Invoice Amount": round2(g.invoice),
        TDS: round2(g.tds),
        VDS: round2(g.vds),
        "Net Payment": round2(g.net),
      })),
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `report-by-${groupBy}.xlsx`);
  }

  const labels = {
    vendor: "Vendor",
    paymentType: "Payment Type",
    month: "Month",
    condition: "TDS/VDS Condition",
  };

  return (
    <div>
      <div
        style={{
          display: "flex",
          gap: 10,
          marginBottom: 16,
          alignItems: "center",
          flexWrap: "wrap",
        }}
      >
        <span
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            color: T.inkSoft,
            fontWeight: 600,
          }}
        >
          Group by
        </span>
        {Object.entries(labels).map(([k, l]) => (
          <button
            key={k}
            onClick={() => setGroupBy(k)}
            style={{
              padding: "7px 14px",
              borderRadius: 20,
              border: `1px solid ${groupBy === k ? T.accent : T.line}`,
              background: groupBy === k ? T.accent : T.card,
              color: groupBy === k ? "#fff" : T.inkSoft,
              fontFamily: "'IBM Plex Sans'",
              fontSize: 12.5,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            {l}
          </button>
        ))}
        <div style={{ flex: 1 }} />
        <button
          onClick={exportExcel}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Download size={14} /> Export Excel
        </button>
        <button
          onClick={() => window.print()}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Printer size={14} /> Print
        </button>
      </div>

      <div
        style={{
          background: T.card,
          border: `1px solid ${T.line}`,
          borderRadius: 12,
          overflow: "hidden",
        }}
      >
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ borderBottom: `2px solid ${T.rule}` }}>
              <th
                style={{
                  textAlign: "left",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                {labels[groupBy]}
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                Entries
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                Invoice
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                TDS
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                VDS
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                Net Payment
              </th>
            </tr>
          </thead>
          <tbody>
            {grouped.map((g, i) => (
              <tr
                key={g.key}
                style={{
                  borderBottom: `1px solid ${T.line}`,
                  background: i % 2 ? T.paper : T.card,
                }}
              >
                <td
                  style={{
                    padding: "10px 14px",
                    fontFamily: "'IBM Plex Sans'",
                    fontSize: 13,
                    fontWeight: 500,
                    borderLeft: `2px solid ${T.rule}`,
                  }}
                >
                  {g.key}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13,
                    color: T.muted,
                  }}
                >
                  {g.count}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13,
                  }}
                >
                  {money(g.invoice)}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13,
                    color: T.rule,
                  }}
                >
                  {money(g.tds)}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13,
                    color: T.rule,
                  }}
                >
                  {money(g.vds)}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13.5,
                    fontWeight: 600,
                    color: T.accentDeep,
                  }}
                >
                  {money(g.net)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   APP SHELL
--------------------------------------------------------------------- */
export default function App() {
  const [tab, setTab] = useState("dashboard");
  const [records, setRecords] = useState(null); // null = loading
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [isDefaultData, setIsDefaultData] = useState(true); // true = still showing sample data, never saved
  const [saveStatus, setSaveStatus] = useState("idle"); // idle | saving | saved | error
  const [lastSavedAt, setLastSavedAt] = useState(null);

  useEffect(() => {
    (async () => {
      const stored = await loadRecords();
      if (stored !== null) {
        // Saved data exists — even an empty list means "user cleared it",
        // and must NOT be overwritten by anything else.
        setRecords(stored);
        setIsDefaultData(false);
      } else {
        // Nothing saved yet — start empty. No sample data is shown by default.
        setRecords([]);
        setIsDefaultData(true);
      }
    })();
  }, []);

  const persist = useCallback(async (next) => {
    setRecords(next);
    setIsDefaultData(false);
    setSaveStatus("saving");
    const ok = await saveRecords(next);
    setSaveStatus(ok ? "saved" : "error");
    if (ok) setLastSavedAt(new Date());
  }, []);

  function handleSave(rec) {
    const withCalc = { ...rec };
    if (records.some((r) => r.id === rec.id)) {
      persist(records.map((r) => (r.id === rec.id ? withCalc : r)));
    } else {
      persist([...records, withCalc]);
    }
    setShowForm(false);
    setEditing(null);
    setTab("transactions");
  }
  function handleDelete(id) {
    if (window.confirm("Delete this ledger entry? This cannot be undone.")) {
      persist(records.filter((r) => r.id !== id));
    }
  }
  function handleImport(imported) {
    persist(imported);
  }
  const nav = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "entry", label: "Master Data", icon: BookOpen },
    { id: "transactions", label: "Transactions", icon: ListFilter },
    { id: "reports", label: "Reports", icon: FileBarChart },
  ];

  if (records === null) {
    return (
      <div
        style={{ padding: 40, fontFamily: "'IBM Plex Sans'", color: T.muted }}
      >
        Loading ledger…
      </div>
    );
  }

  return (
    <div
      style={{
        background: T.paper,
        minHeight: "100vh",
        width: "100vw",
        fontFamily: "'IBM Plex Sans', sans-serif",
      }}
    >
      <style>
        {FONT_CSS}
        {`
      html, body, #root {
      width: 100%;
      min-height: 100%;
      margin: 0;
      padding: 0;
      }
        * { box-sizing: border-box; }
        input:focus, select:focus { border-color: ${T.accent} !important; box-shadow: 0 0 0 3px ${T.ruleSoft}55; }
        @media print {
          .no-print { display: none !important; }
          body { background: white !important; }
        }
      `}
      </style>

      <div style={{ display: "flex" }}>
        {/* Sidebar */}
        <div
          className="no-print"
          style={{
            width: 220,
            minHeight: "100vh",
            background: T.ink,
            padding: "26px 16px",
            position: "sticky",
            top: 0,
            alignSelf: "flex-start",
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              marginBottom: 34,
              padding: "0 6px",
            }}
          >
            <div
              style={{
                width: 30,
                height: 30,
                borderRadius: 6,
                background: T.rule,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontFamily: "'Source Serif 4'",
                color: "#fff",
                fontWeight: 700,
                fontSize: 16,
              }}
            >
              H
            </div>
            <div>
              <div
                style={{
                  fontFamily: "'Source Serif 4'",
                  fontSize: 15.5,
                  color: "#fff",
                  fontWeight: 600,
                  lineHeight: 1.1,
                }}
              >
                Hastizam Ledger
              </div>
              <div
                style={{
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 10.5,
                  color: "#8FA39D",
                }}
              >
                TDS · VDS · Payments
              </div>
            </div>
          </div>
          <nav style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            {nav.map((n) => (
              <button
                key={n.id}
                onClick={() => {
                  setTab(n.id);
                  setShowForm(false);
                }}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 10,
                  padding: "10px 12px",
                  borderRadius: 7,
                  background:
                    tab === n.id ? "rgba(255,255,255,.08)" : "transparent",
                  border: "none",
                  color: tab === n.id ? "#fff" : "#9FB2AC",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 13.5,
                  fontWeight: tab === n.id ? 600 : 500,
                  cursor: "pointer",
                  textAlign: "left",
                  borderLeft:
                    tab === n.id
                      ? `2px solid ${T.rule}`
                      : "2px solid transparent",
                }}
              >
                <n.icon size={16} /> {n.label}
              </button>
            ))}
          </nav>
          <div
            style={{
              marginTop: 30,
              padding: "12px 12px",
              borderRadius: 8,
              background: "rgba(255,255,255,.05)",
              display: "flex",
              gap: 8,
            }}
          >
            <AlertCircle
              size={14}
              color={saveStatus === "error" ? "#E08A70" : "#8FA39D"}
              style={{ flexShrink: 0, marginTop: 1 }}
            />
            <div
              style={{
                fontFamily: "'IBM Plex Sans'",
                fontSize: 10.5,
                color: "#8FA39D",
                lineHeight: 1.5,
              }}
            >
              <div
                style={{
                  fontWeight: 600,
                  color: saveStatus === "error" ? "#E08A70" : "#C9D6D1",
                  marginBottom: 2,
                }}
              >
                {saveStatus === "saving" && "Saving…"}
                {saveStatus === "saved" &&
                  `Saved ${records ? records.length : 0} entries${lastSavedAt ? " · " + lastSavedAt.toLocaleTimeString() : ""}`}
                {saveStatus === "error" &&
                  "Save failed — open browser console (F12) for the error"}
                {saveStatus === "idle" &&
                  (isDefaultData
                    ? "No data yet — add an entry or import Excel"
                    : "Loaded your saved data")}
              </div>
              Prototype storage only — no live Excel or multi-user sync yet.
            </div>
          </div>
        </div>

        {/* Main */}
        <div
          style={{
            flex: 1,
            padding: "28px 34px",
            minWidth: 0,
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: 22,
            }}
          >
            <div>
              <div
                style={{
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11.5,
                  color: T.muted,
                  fontWeight: 600,
                  letterSpacing: ".05em",
                  textTransform: "uppercase",
                }}
              >
                Hastizam Limited
              </div>
              <h1
                style={{
                  fontFamily: "'Source Serif 4'",
                  fontSize: 27,
                  color: T.ink,
                  margin: "2px 0 0",
                }}
              >
                {nav.find((n) => n.id === tab)?.label}
              </h1>
            </div>
            {tab !== "entry" && (
              <button
                className="no-print"
                onClick={() => {
                  setEditing(null);
                  setTab("entry");
                }}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  padding: "10px 18px",
                  borderRadius: 8,
                  border: "none",
                  background: T.accent,
                  color: "#fff",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 13.5,
                  fontWeight: 600,
                  cursor: "pointer",
                }}
              >
                <Plus size={15} /> New Entry
              </button>
            )}
          </div>

          {tab === "dashboard" && <Dashboard records={records} />}
          {tab === "transactions" && (
            <Transactions
              records={records}
              onEdit={(r) => {
                setEditing(r);
                setTab("entry");
              }}
              onDelete={handleDelete}
              onImport={handleImport}
            />
          )}
          {tab === "reports" && <Reports records={records} />}
          {tab === "entry" && (
            <EntryForm
              records={records}
              initial={editing}
              onSave={handleSave}
              onCancel={() => {
                setEditing(null);
                setTab("transactions");
              }}
            />
          )}
        </div>
      </div>
    </div>
  );
}
