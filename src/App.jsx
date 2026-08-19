import React, { useState, useEffect, useLayoutEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Package, Users, CalendarCheck, Wallet, Briefcase,
  CreditCard, Plus, Trash2, AlertTriangle, X, Loader2, Pencil, Truck, Printer, Receipt, ShieldCheck, Boxes, Search, Building2, Workflow, ArrowRight, ArrowLeft, FileSpreadsheet, Lock, LogOut
} from "lucide-react";
import { supabase } from "./supabaseClient.js";
import { useAuth } from "./auth/AuthContext.jsx";

// ---------- storage helpers ----------
const KEYS = {
  inventory: "nova-inventory-items",
  attendance: "nova-attendance-records",
  employees: "nova-employees",
  finance: "nova-finance-entries",
  orders: "nova-orders-projects",
  payments: "nova-payments",
  deliveryDocs: "nova-delivery-challans",
  company: "nova-company-profile",
  payroll: "nova-payroll-slips",
  legalDocs: "nova-legal-documents",
  assets: "nova-assets",
  partyRegistrations: "nova-party-registrations",
  productionWorkflow: "nova-production-workflow",
  legalDocCategories: "nova-legal-doc-categories",
};

// Persistence: Supabase Postgres (table `app_storage`, one row per key) — real
// shared data across every device and every logged-in user, replacing the old
// per-browser localStorage. Function signatures are unchanged on purpose so
// every existing call site (loadList/saveList/loadObj/saveObj) needed zero edits.
async function loadObj(key, fallback) {
  try {
    const { data, error } = await supabase.from("app_storage").select("value").eq("key", key).maybeSingle();
    if (error || !data) return fallback;
    return data.value ?? fallback;
  } catch {
    return fallback;
  }
}
async function saveObj(key, value) {
  try {
    const { error } = await supabase.from("app_storage").upsert({ key, value, updated_at: new Date().toISOString() });
    if (error) console.error("save failed", key, error);
  } catch (e) {
    console.error("save failed", key, e);
  }
}

async function loadList(key) {
  try {
    const { data, error } = await supabase.from("app_storage").select("value").eq("key", key).maybeSingle();
    if (error || !data) return [];
    return Array.isArray(data.value) ? data.value : [];
  } catch {
    return [];
  }
}
async function saveList(key, value) {
  try {
    const { error } = await supabase.from("app_storage").upsert({ key, value, updated_at: new Date().toISOString() });
    if (error) console.error("save failed", key, error);
  } catch (e) {
    console.error("save failed", key, e);
  }
}

const uid = () => Math.random().toString(36).slice(2, 10);

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// Generic table export helpers — `columns` is [{ label, value(row) }, ...]
function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  // target=_blank + rel=noopener lets the browser open/download this outside a
  // sandboxed preview iframe (which blocks a same-frame forced download unless
  // the parent explicitly sets allow-downloads) — the sandbox still permits
  // popups to escape it, so this is what actually gets the file to save.
  a.target = "_blank";
  a.rel = "noopener noreferrer";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

function exportCSV(rows, columns, filename) {
  const escape = (val) => {
    const s = String(val ?? "");
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [
    columns.map((c) => escape(c.label)).join(","),
    ...rows.map((r) => columns.map((c) => escape(c.value(r))).join(",")),
  ];
  downloadBlob(lines.join("\n"), `${filename}.csv`, "text/csv;charset=utf-8;");
}

function exportExcel(rows, columns, filename, sheetName = "Sheet1") {
  const data = [columns.map((c) => c.label), ...rows.map((r) => columns.map((c) => c.value(r)))];
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  // XLSX.writeFile builds its own anchor internally (no way to add target=_blank
  // to it), which gets blocked the same way in a sandboxed preview iframe — build
  // the bytes ourselves and go through the shared downloadBlob() so the fix above
  // applies here too.
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  downloadBlob(wbout, `${filename}.xlsx`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

function formatAadhaar(raw) {
  const digits = raw.replace(/\D/g, "").slice(0, 12);
  return digits.replace(/(\d{4})(?=\d)/g, "$1-");
}
function formatPan(raw) {
  return raw.replace(/[^a-zA-Z0-9]/g, "").toUpperCase().slice(0, 10);
}
const PAN_REGEX = /^[A-Z]{5}[0-9]{4}[A-Z]$/;
// Enforces the PAN shape live, position by position, instead of just warning after
// the fact: first 5 characters must be letters, next 4 must be digits, last 1 must
// be a letter. A wrong-type keystroke for the current position is simply dropped.
function formatPanStrict(raw) {
  const cleaned = raw.toUpperCase().replace(/[^A-Z0-9]/g, "");
  let result = "";
  for (const ch of cleaned) {
    if (result.length < 5) {
      if (/[A-Z]/.test(ch)) result += ch;
    } else if (result.length < 9) {
      if (/[0-9]/.test(ch)) result += ch;
    } else if (result.length < 10) {
      if (/[A-Z]/.test(ch)) result += ch;
    }
    if (result.length === 10) break;
  }
  return result;
}

function formatPhone10(raw) {
  return raw.replace(/\D/g, "").slice(0, 10);
}

const fmt = (n) =>
  "₹" + (Number(n) || 0).toLocaleString("en-IN", { maximumFractionDigits: 0 });
const todayISO = () => new Date().toISOString().slice(0, 10);
const monthKey = (d) => (d ? d.slice(0, 7) : "");
const thisMonthKey = () => todayISO().slice(0, 7);

// ---------- generic UI bits ----------
function Card({ label, value, sub, tone = "default" }) {
  const toneMap = {
    default: "text-neutral-900",
    danger: "text-red-600",
    good: "text-emerald-600",
  };
  return (
    <div className="bg-white border border-neutral-200 rounded-xl p-4 shadow-sm">
      <div className="text-xs font-semibold uppercase tracking-wide text-neutral-500">
        {label}
      </div>
      <div className={`text-2xl font-bold mt-1 ${toneMap[tone]}`}>{value}</div>
      {sub && <div className="text-xs text-neutral-400 mt-1">{sub}</div>}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1 text-sm">
      <span className="text-neutral-600 font-medium">{label}</span>
      {children}
    </label>
  );
}

const inputCls =
  "border border-neutral-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500 bg-white";

// Dropdown that pulls names from Client / Vendor Registration records, with a manual
// fallback for a one-off party that hasn't been registered yet. `onPick` (optional)
// receives the full matched registration record so callers can auto-fill other fields
// (e.g. GSTIN, address). Give it a `key` that changes whenever the surrounding modal
// opens for a new/different row, so its "custom text" state resets correctly.
function PartyPicker({ label, value, onChange, records, onPick, placeholder }) {
  const names = records.map((r) => r.name);
  const [custom, setCustom] = useState(!!value && !names.includes(value));
  return (
    <Field label={label}>
      <select
        className={inputCls}
        value={custom ? "__other__" : value}
        onChange={(e) => {
          const v = e.target.value;
          if (v === "__other__") {
            setCustom(true);
            onChange("");
            onPick && onPick(null);
          } else if (v === "") {
            setCustom(false);
            onChange("");
            onPick && onPick(null);
          } else {
            setCustom(false);
            onChange(v);
            onPick && onPick(records.find((r) => r.name === v) || null);
          }
        }}
      >
        <option value="">— select from Client Registration —</option>
        {records.map((r) => (
          <option key={r.id} value={r.name}>{r.name}</option>
        ))}
        <option value="__other__">+ Other / not registered yet</option>
      </select>
      {custom && (
        <input
          className={`${inputCls} mt-1.5`}
          placeholder={placeholder || "Type name…"}
          value={value}
          onChange={(e) => onChange(e.target.value)}
        />
      )}
      {records.length === 0 && (
        <span className="text-[11px] text-neutral-400 mt-0.5">
          No one registered yet — add them in Client / Vendor Registration first, or type a name manually above.
        </span>
      )}
    </Field>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div className="fixed inset-0 bg-black/40 z-50 overflow-y-auto p-4">
      <div className="min-h-full flex items-start justify-center py-8">
        <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg">
          <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-200 sticky top-0 bg-white rounded-t-2xl z-10">
            <h3 className="font-bold text-lg text-neutral-900">{title}</h3>
            <button
              onClick={onClose}
              className="text-neutral-400 hover:text-neutral-700 transition"
            >
              <X size={20} />
            </button>
          </div>
          <div className="p-5">{children}</div>
        </div>
      </div>
    </div>
  );
}

function Table({ columns, rows, onDelete, onEdit, emptyMsg, hideActions }) {
  if (!rows.length) {
    return (
      <div className="text-center text-neutral-400 text-sm py-14 border border-dashed border-neutral-200 rounded-xl">
        {emptyMsg}
      </div>
    );
  }
  return (
    <div className="overflow-x-auto border border-neutral-200 rounded-xl">
      <table className="w-full text-sm">
        <thead>
          <tr className="bg-neutral-50 text-left text-neutral-500 text-xs uppercase tracking-wide">
            {columns.map((c) => (
              <th key={c.key} className="px-4 py-3 font-semibold">
                {c.label}
              </th>
            ))}
            {!hideActions && <th className="px-4 py-3 w-10"></th>}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr
              key={r.id ?? i}
              className={`border-t border-neutral-100 ${
                i % 2 ? "bg-white" : "bg-neutral-50/40"
              }`}
            >
              {columns.map((c) => (
                <td key={c.key} className="px-4 py-3 text-neutral-800">
                  {c.render ? c.render(r) : r[c.key]}
                </td>
              ))}
              {!hideActions && (
                <td className="px-4 py-3 text-right whitespace-nowrap">
                  {onEdit && (
                    <button
                      onClick={() => onEdit(r)}
                      className="text-neutral-300 hover:text-red-600 transition mr-2"
                      title="Edit"
                    >
                      <Pencil size={15} />
                    </button>
                  )}
                  <button
                    onClick={() => onDelete(r.id)}
                    className="text-neutral-300 hover:text-red-600 transition"
                    title="Remove"
                  >
                    <Trash2 size={15} />
                  </button>
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function AddButton({ onClick, text = "Add", disabled = false }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className="inline-flex items-center gap-1.5 bg-red-600 hover:bg-red-700 text-white text-sm font-semibold px-4 py-2 rounded-lg transition shadow-sm disabled:opacity-40 disabled:cursor-not-allowed disabled:hover:bg-red-600"
    >
      <Plus size={16} /> {text}
    </button>
  );
}

function Pill({ children, tone }) {
  const tones = {
    green: "bg-emerald-50 text-emerald-700 border-emerald-200",
    amber: "bg-amber-50 text-amber-700 border-amber-200",
    red: "bg-red-50 text-red-700 border-red-200",
    gray: "bg-neutral-100 text-neutral-600 border-neutral-200",
    blue: "bg-blue-50 text-blue-700 border-blue-200",
  };
  return (
    <span
      className={`inline-block text-xs font-semibold px-2 py-0.5 rounded-full border ${tones[tone] || tones.gray}`}
    >
      {children}
    </span>
  );
}

// Reusable attachment uploader — stores each record's files under their own
// storage key (nova-doc-<recordId>) so attachments never bloat the parent list.
const MAX_ATTACHMENT_BYTES = 1.5 * 1024 * 1024;

function AttachmentsField({ recordId, label = "Attachments" }) {
  const [files, setFiles] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    let cancelled = false;
    setLoaded(false);
    (async () => {
      const data = await loadObj(`nova-doc-${recordId}`, { files: [] });
      if (!cancelled) {
        setFiles(data.files || []);
        setLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, [recordId]);

  useEffect(() => {
    if (loaded) saveObj(`nova-doc-${recordId}`, { files });
  }, [files, loaded]);

  const handleUpload = async (e) => {
    setError("");
    const picked = Array.from(e.target.files || []);
    for (const file of picked) {
      if (file.size > MAX_ATTACHMENT_BYTES) {
        setError(`"${file.name}" is over 1.5 MB — try a smaller photo or a compressed PDF.`);
        continue;
      }
      try {
        const dataUrl = await fileToDataUrl(file);
        setFiles((prev) => [...prev, { id: uid(), name: file.name, size: file.size, dataUrl }]);
      } catch {
        setError(`Couldn't read "${file.name}" — try again.`);
      }
    }
    e.target.value = "";
  };

  const removeFile = (id) => setFiles(files.filter((f) => f.id !== id));

  return (
    <div className="border-t border-neutral-100 pt-3">
      <div className="flex justify-between items-center mb-2">
        <div className="text-sm font-semibold text-neutral-700">{label}</div>
        <label className="text-xs font-semibold text-red-600 hover:text-red-700 cursor-pointer inline-flex items-center gap-1">
          <Plus size={13} /> Attach file
          <input type="file" accept="image/*,application/pdf" multiple className="hidden" onChange={handleUpload} />
        </label>
      </div>
      {error && <p className="text-xs text-red-600 mb-2">{error}</p>}
      {!loaded ? (
        <p className="text-xs text-neutral-400">Loading attachments…</p>
      ) : files.length === 0 ? (
        <p className="text-xs text-neutral-400">No files attached yet — bills, POs, or invoices you attach here stay linked to this record.</p>
      ) : (
        <div className="space-y-1">
          {files.map((f) => (
            <div key={f.id} className="flex items-center justify-between text-xs bg-neutral-50 border border-neutral-200 rounded-lg px-3 py-2">
              <a href={f.dataUrl} download={f.name} target="_blank" rel="noopener noreferrer" className="text-neutral-700 hover:text-red-600 truncate">
                {f.name} <span className="text-neutral-400">({(f.size / 1024).toFixed(0)} KB)</span>
              </a>
              <button onClick={() => removeFile(f.id)} className="text-neutral-300 hover:text-red-600 ml-2 shrink-0">
                <Trash2 size={13} />
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// Every real dashboard tab (excludes "access", which is admin-only and injected
// separately) — shared between the sidebar nav and the Access Control panel so
// the two never drift out of sync.
const ALL_TABS = [
  { id: "overview", label: "Overview", icon: LayoutDashboard },
  { id: "inventory", label: "Inventory", icon: Package },
  { id: "assets", label: "Asset Management", icon: Boxes },
  { id: "attendance", label: "Attendance", icon: CalendarCheck },
  { id: "employees", label: "Employees & Salary", icon: Users },
  { id: "capex", label: "CapEx", icon: Wallet },
  { id: "opex", label: "OpEx", icon: Wallet },
  { id: "orders", label: "Orders / Projects", icon: Briefcase },
  { id: "workflow", label: "Production Workflow", icon: Workflow },
  { id: "clientPayments", label: "Client Payments", icon: CreditCard },
  { id: "vendorPayments", label: "Vendor Payments", icon: CreditCard },
  { id: "payroll", label: "Payroll", icon: Receipt },
  { id: "delivery", label: "Delivery Documents", icon: Truck },
  { id: "partyReg", label: "Client / Vendor Registration", icon: Building2 },
  { id: "legal", label: "Legal Documents", icon: ShieldCheck },
];

// Admin-only panel: manage who can log in to what. Only creates/edits rows in
// user_permissions — it can't create the underlying Supabase Auth user itself
// (that needs the Supabase dashboard, since the anon key has no admin rights).
function AccessControlTab({ currentEmail }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [form, setForm] = useState({ email: "", display_name: "", is_admin: false, allowed_tabs: [] });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const refresh = () => {
    setLoading(true);
    supabase
      .from("user_permissions")
      .select("*")
      .order("created_at", { ascending: true })
      .then(({ data, error }) => {
        if (error) setError(error.message);
        else setRows(data || []);
        setLoading(false);
      });
  };
  useEffect(refresh, []);

  const toggleTab = (tabId) => {
    setForm((f) => ({
      ...f,
      allowed_tabs: f.allowed_tabs.includes(tabId) ? f.allowed_tabs.filter((t) => t !== tabId) : [...f.allowed_tabs, tabId],
    }));
  };

  const addUser = async () => {
    if (!form.email.trim()) return;
    setSaving(true);
    setError("");
    const { error } = await supabase.from("user_permissions").insert({
      email: form.email.trim().toLowerCase(),
      display_name: form.display_name.trim(),
      is_admin: form.is_admin,
      allowed_tabs: form.allowed_tabs,
    });
    setSaving(false);
    if (error) {
      setError(error.message);
    } else {
      setForm({ email: "", display_name: "", is_admin: false, allowed_tabs: [] });
      refresh();
    }
  };

  const updateRow = async (row, patch) => {
    const { error } = await supabase.from("user_permissions").update(patch).eq("id", row.id);
    if (error) setError(error.message);
    else setRows((prev) => prev.map((r) => (r.id === row.id ? { ...r, ...patch } : r)));
  };

  const removeRow = async (row) => {
    if (row.email === currentEmail) return; // guard: can't remove yourself by accident
    const { error } = await supabase.from("user_permissions").delete().eq("id", row.id);
    if (error) setError(error.message);
    else setRows((prev) => prev.filter((r) => r.id !== row.id));
  };

  const rowToggleTab = (row, tabId) => {
    const next = row.allowed_tabs.includes(tabId)
      ? row.allowed_tabs.filter((t) => t !== tabId)
      : [...row.allowed_tabs, tabId];
    updateRow(row, { allowed_tabs: next });
  };

  return (
    <div className="space-y-4">
      <div>
        <h2 className="font-bold text-lg text-neutral-900">Access Control</h2>
        <p className="text-xs text-neutral-400 mt-0.5">
          Who can log in and which tabs they can see. Admins always see everything. To create the login itself, add the person in Supabase → Authentication → Users first — this panel only sets what they can access once they sign in.
        </p>
      </div>

      {error && <p className="text-xs text-red-600">{error}</p>}

      <div className="bg-white border border-neutral-200 rounded-xl p-4">
        <div className="text-sm font-semibold text-neutral-800 mb-3">Add a user</div>
        <div className="grid grid-cols-2 gap-3 mb-3">
          <label className="flex flex-col gap-1 text-sm">
            <span className="text-neutral-600 font-medium">Email (must match their Supabase Auth login)</span>
            <input
              type="email"
              className="border border-neutral-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500"
              value={form.email}
              onChange={(e) => setForm({ ...form, email: e.target.value })}
            />
          </label>
          <label className="flex flex-col gap-1 text-sm">
            <span className="text-neutral-600 font-medium">Display name</span>
            <input
              className="border border-neutral-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500"
              placeholder="e.g. Anbararsan"
              value={form.display_name}
              onChange={(e) => setForm({ ...form, display_name: e.target.value })}
            />
          </label>
        </div>
        <label className="flex items-center gap-2 text-sm mb-3">
          <input type="checkbox" className="w-4 h-4 accent-red-600" checked={form.is_admin} onChange={(e) => setForm({ ...form, is_admin: e.target.checked })} />
          <span className="text-neutral-700">Admin (sees every tab, can manage other users)</span>
        </label>
        {!form.is_admin && (
          <div className="mb-3">
            <div className="text-xs font-semibold text-neutral-500 mb-1.5">Allowed tabs</div>
            <div className="flex flex-wrap gap-1.5">
              {ALL_TABS.map((t) => (
                <button
                  key={t.id}
                  type="button"
                  onClick={() => toggleTab(t.id)}
                  className={`text-xs font-semibold px-2.5 py-1 rounded-full border transition ${
                    form.allowed_tabs.includes(t.id)
                      ? "bg-red-600 text-white border-red-600"
                      : "bg-white text-neutral-500 border-neutral-200 hover:border-red-300"
                  }`}
                >
                  {t.label}
                </button>
              ))}
            </div>
          </div>
        )}
        <button
          onClick={addUser}
          disabled={saving || !form.email.trim()}
          className="inline-flex items-center gap-1.5 bg-red-600 hover:bg-red-700 disabled:opacity-50 text-white text-sm font-semibold px-4 py-2 rounded-lg transition"
        >
          <Plus size={15} /> Add user
        </button>
      </div>

      <div className="bg-white border border-neutral-200 rounded-xl overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-neutral-50 text-neutral-500 text-xs uppercase tracking-wide">
            <tr>
              <th className="text-left px-4 py-2.5">User</th>
              <th className="text-left px-4 py-2.5">Admin</th>
              <th className="text-left px-4 py-2.5">Allowed tabs</th>
              <th className="px-4 py-2.5"></th>
            </tr>
          </thead>
          <tbody>
            {loading && (
              <tr><td colSpan={4} className="px-4 py-6 text-center text-neutral-400 text-sm">Loading…</td></tr>
            )}
            {!loading && rows.length === 0 && (
              <tr><td colSpan={4} className="px-4 py-6 text-center text-neutral-400 text-sm">No users added yet.</td></tr>
            )}
            {rows.map((row) => (
              <tr key={row.id} className="border-t border-neutral-100 align-top">
                <td className="px-4 py-3">
                  <div className="font-semibold text-neutral-800">{row.display_name || "—"}</div>
                  <div className="text-xs text-neutral-400">{row.email}</div>
                </td>
                <td className="px-4 py-3">
                  <input
                    type="checkbox"
                    className="w-4 h-4 accent-red-600 cursor-pointer"
                    checked={row.is_admin}
                    onChange={(e) => updateRow(row, { is_admin: e.target.checked })}
                  />
                </td>
                <td className="px-4 py-3">
                  {row.is_admin ? (
                    <span className="text-xs text-neutral-400">Everything (admin)</span>
                  ) : (
                    <div className="flex flex-wrap gap-1">
                      {ALL_TABS.map((t) => (
                        <button
                          key={t.id}
                          onClick={() => rowToggleTab(row, t.id)}
                          className={`text-[11px] font-semibold px-2 py-0.5 rounded-full border transition ${
                            row.allowed_tabs.includes(t.id)
                              ? "bg-red-50 text-red-700 border-red-200"
                              : "bg-neutral-50 text-neutral-400 border-neutral-200 hover:border-red-300"
                          }`}
                        >
                          {t.label}
                        </button>
                      ))}
                    </div>
                  )}
                </td>
                <td className="px-4 py-3 text-right">
                  <button
                    onClick={() => removeRow(row)}
                    disabled={row.email === currentEmail}
                    title={row.email === currentEmail ? "You can't remove yourself" : "Remove access"}
                    className="text-neutral-400 hover:text-red-600 disabled:opacity-30 disabled:cursor-not-allowed"
                  >
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- main app ----------
export default function NovaOps() {
  const { session, permissions, signOut } = useAuth();
  const isAdmin = !!permissions?.is_admin;
  const allowedTabIds = isAdmin ? null : (permissions?.allowed_tabs || []);

  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("overview");

  const [inventory, setInventory] = useState([]);
  const [attendance, setAttendance] = useState([]);
  const [employees, setEmployees] = useState([]);
  const [finance, setFinance] = useState([]);
  const [orders, setOrders] = useState([]);
  const [payments, setPayments] = useState([]);
  const [deliveryDocs, setDeliveryDocs] = useState([]);
  const [payrollSlips, setPayrollSlips] = useState([]);
  const [legalDocs, setLegalDocs] = useState([]);
  const [assets, setAssets] = useState([]);
  const [partyRegistrations, setPartyRegistrations] = useState([]);
  const [productionWorkflow, setProductionWorkflow] = useState([]);
  const [legalDocCategories, setLegalDocCategories] = useState([]);
  const [company, setCompany] = useState({ name: "NOVA", address: "", gstin: "" });
  const [printContent, setPrintContent] = useState(null);
  const [printTitle, setPrintTitle] = useState("nova-document");

  // useLayoutEffect (not useEffect + setTimeout) so this fires synchronously right
  // after React commits the hidden print node to the DOM — staying inside the same
  // user-gesture window as the button click. A setTimeout here, even a short one,
  // detaches the download from that gesture and gets it silently blocked, especially
  // inside a sandboxed preview iframe.
  useLayoutEffect(() => {
    if (printContent) {
      const node = document.getElementById("global-print-area");
      const html = node ? node.innerHTML : "";
      const fullHtml = `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>${printTitle}</title>
    <style>
      body { margin: 0; padding: 16px; font-family: Arial, sans-serif; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      table { border-collapse: collapse; }
      @media print { @page { margin: 12mm; } body { padding: 0; } }
    </style>
  </head>
  <body>${html}
    <script>
      window.onload = function () { setTimeout(function () { window.print(); }, 300); };
    </script>
  </body>
</html>`;
      const blob = new Blob([fullHtml], { type: "text/html" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${printTitle}.html`;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      setPrintContent(null);
    }
  }, [printContent]);

  useEffect(() => {
    (async () => {
      const [inv, att, emp, fin, ord, pay, dd, co, pr, ld, as, pReg, pw, ldc] = await Promise.all([
        loadList(KEYS.inventory),
        loadList(KEYS.attendance),
        loadList(KEYS.employees),
        loadList(KEYS.finance),
        loadList(KEYS.orders),
        loadList(KEYS.payments),
        loadList(KEYS.deliveryDocs),
        loadObj(KEYS.company, { name: "NOVA", address: "", gstin: "" }),
        loadList(KEYS.payroll),
        loadList(KEYS.legalDocs),
        loadList(KEYS.assets),
        loadList(KEYS.partyRegistrations),
        loadList(KEYS.productionWorkflow),
        loadList(KEYS.legalDocCategories),
      ]);
      setInventory(inv);
      setAttendance(att);
      setEmployees(emp);
      setFinance(fin);
      setOrders(ord);
      setPayments(pay);
      setDeliveryDocs(dd);
      setCompany(co);
      setPayrollSlips(pr);
      setLegalDocs(ld);
      setAssets(as);
      setPartyRegistrations(pReg);
      setProductionWorkflow(pw);
      setLegalDocCategories(ldc);
      setLoading(false);
    })();
  }, []);

  // persist on change (after initial load)
  useEffect(() => {
    if (!loading) saveList(KEYS.inventory, inventory);
  }, [inventory, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.attendance, attendance);
  }, [attendance, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.employees, employees);
  }, [employees, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.finance, finance);
  }, [finance, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.orders, orders);
  }, [orders, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.payments, payments);
  }, [payments, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.deliveryDocs, deliveryDocs);
  }, [deliveryDocs, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.payroll, payrollSlips);
  }, [payrollSlips, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.legalDocs, legalDocs);
  }, [legalDocs, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.assets, assets);
  }, [assets, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.partyRegistrations, partyRegistrations);
  }, [partyRegistrations, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.productionWorkflow, productionWorkflow);
  }, [productionWorkflow, loading]);
  useEffect(() => {
    if (!loading) saveList(KEYS.legalDocCategories, legalDocCategories);
  }, [legalDocCategories, loading]);
  useEffect(() => {
    if (!loading) saveObj(KEYS.company, company);
  }, [company, loading]);

  const tabs = ALL_TABS;

  // Admins see every tab plus the admin-only Access Control panel; everyone
  // else sees only whatever's in their allowed_tabs list.
  const visibleTabs = isAdmin
    ? [...tabs, { id: "access", label: "Access Control", icon: Lock }]
    : tabs.filter((t) => allowedTabIds.includes(t.id));

  // If the currently-selected tab isn't visible to this user (first load,
  // permissions changed, or they were on a tab they've since lost access to),
  // fall back to the first tab they're actually allowed to see.
  useEffect(() => {
    if (!loading && visibleTabs.length > 0 && !visibleTabs.some((t) => t.id === tab)) {
      setTab(visibleTabs[0].id);
    }
  }, [loading, permissions]);

  if (loading) {
    return (
      <div className="flex items-center justify-center h-96 text-neutral-400 gap-2">
        <Loader2 className="animate-spin" size={18} /> Loading dashboard…
      </div>
    );
  }

  return (
    <div className="bg-neutral-50 min-h-screen font-sans">
      <div className="max-w-7xl mx-auto px-4 py-6 flex gap-6 items-start">
        <aside className="w-64 shrink-0 sticky top-6">
          <div className="mb-5 px-1">
            <div className="text-xs font-bold tracking-widest text-red-600 uppercase">
              NOVA
            </div>
            <h1 className="text-xl font-bold text-neutral-900">
              Operations Dashboard
            </h1>
          </div>
          <div className="mb-3 px-1 flex items-center justify-between gap-2">
            <div className="min-w-0">
              <div className="text-xs font-semibold text-neutral-700 truncate">{session?.user?.email}</div>
              <div className="text-[11px] text-neutral-400">{isAdmin ? "Admin" : "Restricted access"}</div>
            </div>
            <button
              onClick={signOut}
              title="Sign out"
              className="shrink-0 text-neutral-400 hover:text-red-600 transition p-1.5 rounded-lg hover:bg-neutral-100"
            >
              <LogOut size={16} />
            </button>
          </div>
          <nav className="flex flex-col gap-1 bg-white border border-neutral-200 rounded-2xl p-2">
            {visibleTabs.map((t) => {
              const Icon = t.icon;
              const active = tab === t.id;
              return (
                <button
                  key={t.id}
                  onClick={() => setTab(t.id)}
                  className={`flex items-center gap-2 text-sm font-semibold px-3 py-2.5 rounded-xl transition text-left ${
                    active
                      ? "bg-red-600 text-white shadow-sm"
                      : "text-neutral-500 hover:text-neutral-900 hover:bg-neutral-100"
                  }`}
                >
                  <Icon size={16} /> {t.label}
                </button>
              );
            })}
          </nav>
        </aside>

        <main className="flex-1 min-w-0">
        {visibleTabs.length === 0 && (
          <div className="bg-white border border-neutral-200 rounded-2xl p-8 text-center">
            <p className="text-sm text-neutral-500">
              Your account doesn't have access to any tabs yet. Ask your admin to grant access in Access Control.
            </p>
          </div>
        )}
        {tab === "access" && isAdmin && <AccessControlTab currentEmail={session?.user?.email} />}
        {tab === "overview" && (
          <Overview
            inventory={inventory}
            attendance={attendance}
            employees={employees}
            finance={finance}
            orders={orders}
            payments={payments}
          />
        )}
        {tab === "inventory" && (
          <InventoryTab
            items={inventory}
            setItems={setInventory}
            entries={finance}
            setEntries={setFinance}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "assets" && (
          <AssetsTab
            assets={assets}
            setAssets={setAssets}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "attendance" && (
          <AttendanceTab
            records={attendance}
            setRecords={setAttendance}
            employees={employees}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "employees" && (
          <EmployeesTab
            employees={employees}
            setEmployees={setEmployees}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "capex" && (
          <FinanceTab
            entries={finance}
            setEntries={setFinance}
            type="CapEx"
            assets={assets}
            setAssets={setAssets}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "opex" && (
          <FinanceTab
            entries={finance}
            setEntries={setFinance}
            type="OpEx"
            assets={assets}
            setAssets={setAssets}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "orders" && (
          <OrdersTab
            orders={orders}
            setOrders={setOrders}
            company={company}
            partyRegistrations={partyRegistrations}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "workflow" && (
          <ProductionWorkflowTab
            orders={orders}
            setOrders={setOrders}
            employees={employees}
            workflow={productionWorkflow}
            setWorkflow={setProductionWorkflow}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "clientPayments" && (
          <PaymentsTab payments={payments} setPayments={setPayments} partyType="Client" partyRegistrations={partyRegistrations} orders={orders} />
        )}
        {tab === "vendorPayments" && (
          <PaymentsTab payments={payments} setPayments={setPayments} partyType="Vendor" partyRegistrations={partyRegistrations} orders={orders} />
        )}
        {tab === "payroll" && (
          <PayrollTab
            slips={payrollSlips}
            setSlips={setPayrollSlips}
            employees={employees}
            attendance={attendance}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "delivery" && (
          <DeliveryDocsTab
            docs={deliveryDocs}
            setDocs={setDeliveryDocs}
            company={company}
            setCompany={setCompany}
            orders={orders}
            partyRegistrations={partyRegistrations}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "partyReg" && (
          <PartyRegistrationTab
            registrations={partyRegistrations}
            setRegistrations={setPartyRegistrations}
            company={company}
            setPrintContent={setPrintContent}
            setPrintTitle={setPrintTitle}
          />
        )}
        {tab === "legal" && (
          <LegalDocsTab
            docs={legalDocs}
            setDocs={setLegalDocs}
            customCategories={legalDocCategories}
            setCustomCategories={setLegalDocCategories}
          />
        )}
        </main>
      </div>

      <style>{`#global-print-area { display: none; }`}</style>
      <div id="global-print-area">{printContent}</div>
    </div>
  );
}

// ---------- OVERVIEW ----------
function Overview({ inventory, attendance, employees, finance, orders, payments }) {
  const stockValue = inventory.reduce(
    (s, i) => s + Number(i.quantity) * Number(i.unitCost || 0),
    0
  );
  const lowStock = inventory.filter(
    (i) => Number(i.quantity) <= Number(i.reorderLevel || 0)
  );
  const today = todayISO();
  const presentToday = attendance.filter(
    (a) => a.date === today && a.status === "Present"
  ).length;
  const mKey = thisMonthKey();
  const capexTotal = finance
    .filter((f) => f.type === "CapEx")
    .reduce((s, f) => s + Number(f.amount), 0);
  const opexMonth = finance
    .filter((f) => f.type === "OpEx" && monthKey(f.date) === mKey)
    .reduce((s, f) => s + Number(f.amount), 0);
  const activeOrders = orders.filter(
    (o) => o.status !== "Delivered" && o.status !== "Cancelled"
  ).length;
  const pendingPayable = payments
  const pendingReceivable = payments
    .filter((p) => p.type === "Receivable" && clientPaymentStatus(p) !== "Paid")
    .reduce((s, p) => s + (Number(p.amount || 0) - paymentReceived(p)), 0);

  // Invoice closing date reminders: unpaid/partially-paid client invoices due within
  // 5 days, plus anything already past its closing date, so nothing quietly slips by.
  const dueSoonInvoices = payments
    .filter((p) => p.type === "Receivable" && p.dueDate)
    .map((p) => ({
      ...p,
      liveStatus: clientPaymentStatus(p),
      balance: Number(p.amount || 0) - paymentReceived(p),
      daysLeft: Math.ceil((new Date(p.dueDate) - new Date(today)) / (1000 * 60 * 60 * 24)),
    }))
    .filter((p) => p.liveStatus !== "Paid" && p.daysLeft <= 5)
    .sort((a, b) => a.daysLeft - b.daysLeft);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <Card label="Stock value" value={fmt(stockValue)} sub={`${inventory.length} SKUs`} />
        <Card
          label="Low stock alerts"
          value={lowStock.length}
          tone={lowStock.length ? "danger" : "good"}
          sub={lowStock.length ? "Needs reorder" : "All healthy"}
        />
        <Card label="Employees" value={employees.length} sub={`${presentToday} present today`} />
        <Card label="Active orders" value={activeOrders} sub={`${orders.length} total`} />
        <Card label="Total CapEx" value={fmt(capexTotal)} sub="Machinery, equipment, land etc." />
        <Card label="OpEx this month" value={fmt(opexMonth)} />
        <Card label="Payable pending" value={fmt(pendingPayable)} tone={pendingPayable ? "danger" : "good"} />
        <Card label="Receivable pending" value={fmt(pendingReceivable)} tone="good" sub="Balance outstanding" />
        <Card
          label="Invoices due soon"
          value={dueSoonInvoices.length}
          tone={dueSoonInvoices.length ? "danger" : "good"}
          sub={dueSoonInvoices.length ? "Within 5 days or overdue" : "None due soon"}
        />
      </div>

      {dueSoonInvoices.length > 0 && (
        <div className="bg-white border border-amber-200 rounded-xl p-4">
          <div className="flex items-center gap-2 text-amber-600 font-semibold text-sm mb-2">
            <AlertTriangle size={16} /> Invoices due within 5 days
          </div>
          <ul className="text-sm text-neutral-700 space-y-1">
            {dueSoonInvoices.slice(0, 6).map((p) => (
              <li key={p.id} className="flex justify-between">
                <span>{p.party}{p.orderRef ? ` — ${p.orderRef}` : ""} {p.liveStatus === "Partially Paid" && <span className="text-blue-600">(partial)</span>}</span>
                <span className={p.daysLeft < 0 ? "text-red-600 font-semibold" : "text-neutral-400"}>
                  {p.daysLeft < 0 ? `Overdue by ${Math.abs(p.daysLeft)}d` : p.daysLeft === 0 ? "Due today" : `Due in ${p.daysLeft}d`} · {fmt(p.balance)} due
                </span>
              </li>
            ))}
          </ul>
        </div>
      )}

      {lowStock.length > 0 && (
        <div className="bg-white border border-red-200 rounded-xl p-4">
          <div className="flex items-center gap-2 text-red-600 font-semibold text-sm mb-2">
            <AlertTriangle size={16} /> Items at or below reorder level
          </div>
          <ul className="text-sm text-neutral-700 space-y-1">
            {lowStock.slice(0, 6).map((i) => (
              <li key={i.id} className="flex justify-between">
                <span>{i.name}</span>
                <span className="text-neutral-400">
                  {i.quantity} {i.unit} left (reorder at {i.reorderLevel})
                </span>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}

// ---------- INVENTORY ----------
const BLANK_ITEM = { name: "", category: "", quantity: "", unit: "pcs", reorderLevel: "", unitCost: "" };

function InventoryTab({ items, setItems, entries, setEntries, company, setPrintContent, setPrintTitle }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [form, setForm] = useState(BLANK_ITEM);
  const [query, setQuery] = useState("");

  const openAdd = () => { setForm(BLANK_ITEM); setEditingId(null); setActiveId(uid()); setOpen(true); };
  const openEdit = (row) => { setForm({ ...BLANK_ITEM, ...row }); setEditingId(row.id); setActiveId(row.id); setOpen(true); };

  const save = () => {
    if (!form.name) return;
    if (editingId) {
      setItems(items.map((i) => (i.id === editingId ? { ...i, ...form, id: editingId } : i)));
      // If this item was already sent to OpEx, keep that OpEx entry's cost/qty/name
      // in sync — otherwise correcting quantity or unit cost here leaves the OpEx
      // entry frozen at the old, un-updated total.
      setEntries((prev) =>
        prev.map((e) =>
          e.inventoryItemId === editingId
            ? {
                ...e,
                category: form.category,
                description: form.name,
                unitPrice: form.unitCost,
                units: form.quantity,
                unit: form.unit,
                amount: Number(form.unitCost || 0) * Number(form.quantity || 0),
              }
            : e
        )
      );
    } else {
      setItems([...items, { id: activeId, ...form }]);
    }
    setForm(BLANK_ITEM);
    setEditingId(null);
    setOpen(false);
  };

  const sendToOpex = (item) => {
    // Idempotency guard: check both the item's own flag AND whether an OpEx entry
    // already exists for it, so a stale click or race never creates a duplicate.
    if (item.sentToOpex || entries.some((e) => e.inventoryItemId === item.id)) return;
    const newEntry = {
      id: uid(),
      type: "OpEx",
      category: item.category,
      description: item.name,
      unitPrice: item.unitCost,
      units: item.quantity,
      unit: item.unit,
      amount: Number(item.unitCost || 0) * Number(item.quantity || 0),
      date: todayISO(),
      allocations: [],
      fromInventory: true,
      inventoryItemId: item.id,
    };
    setEntries([...entries, newEntry]);
    setItems(items.map((i) => (i.id === item.id ? { ...i, sentToOpex: true } : i)));
  };

  const q = query.trim().toLowerCase();
  const filteredItems = q
    ? items.filter(
        (i) =>
          (i.name || "").toLowerCase().includes(q) ||
          (i.category || "").toLowerCase().includes(q)
      )
    : items;

  const totalInventoryValue = items.reduce((s, i) => s + Number(i.quantity || 0) * Number(i.unitCost || 0), 0);
  const lowStockCount = items.filter((i) => Number(i.quantity) <= Number(i.reorderLevel || 0)).length;

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center flex-wrap gap-3">
        <h2 className="font-bold text-lg text-neutral-900">Inventory / Stock</h2>
        <div className="flex items-center gap-2">
          <div className="relative">
            <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              className={`${inputCls} pl-8 w-56`}
              placeholder="Search items or category…"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <button
            onClick={() => {
              setPrintTitle("Inventory-Register");
              setPrintContent(<InventoryPrintLayout items={filteredItems} company={company} />);
            }}
            className="inline-flex items-center gap-1.5 bg-white border border-neutral-300 hover:border-red-400 text-sm font-semibold px-4 py-2 rounded-lg transition"
          >
            <Printer size={16} /> Download PDF
          </button>
          <AddButton onClick={openAdd} text="Add item" />
        </div>
      </div>
      <div className="grid grid-cols-2 gap-4">
        <Card label="Inventory Value" value={fmt(totalInventoryValue)} sub={`${items.length} SKUs tracked`} />
        <Card
          label="Low Stock Alerts"
          value={lowStockCount}
          tone={lowStockCount ? "danger" : "good"}
          sub={lowStockCount ? "Items at or below reorder level" : "All items healthy"}
        />
      </div>
      <p className="text-xs text-neutral-400 -mt-2">
        Tick "Send to OpEx" on any item to log it as an operating expense — its cost stays linked, so quantity or unit-cost corrections here update the OpEx entry too.
      </p>
      <Table
        emptyMsg={q ? `No items match "${query}".` : "No inventory items yet. Add your first SKU to start tracking stock."}
        columns={[
          { key: "name", label: "Item" },
          { key: "category", label: "Category" },
          {
            key: "quantity",
            label: "Qty",
            render: (r) => (
              <span
                className={
                  Number(r.quantity) <= Number(r.reorderLevel || 0)
                    ? "text-red-600 font-semibold"
                    : ""
                }
              >
                {r.quantity} {r.unit}
              </span>
            ),
          },
          { key: "reorderLevel", label: "Reorder at" },
          { key: "unitCost", label: "Unit cost", render: (r) => fmt(r.unitCost) },
          {
            key: "value",
            label: "Value",
            render: (r) => fmt(Number(r.quantity) * Number(r.unitCost || 0)),
          },
          {
            key: "sendToOpex",
            label: "Send to OpEx",
            render: (r) =>
              r.sentToOpex ? (
                <span className="text-xs text-emerald-600 font-semibold whitespace-nowrap">✓ In OpEx</span>
              ) : (
                <input
                  type="checkbox"
                  checked={false}
                  onChange={(e) => { e.target.disabled = true; sendToOpex(r); }}
                  className="w-4 h-4 accent-red-600 cursor-pointer"
                  title="Log this item's cost as an OpEx entry"
                />
              ),
          },
        ]}
        rows={filteredItems}
        onDelete={(id) => {
          setItems(items.filter((i) => i.id !== id));
          // Don't leave a linked OpEx entry pointing at a deleted inventory item —
          // unlink it (keep the OpEx entry itself, just mark it Manual) instead of
          // a dangling reference.
          setEntries((prev) =>
            prev.map((e) => (e.inventoryItemId === id ? { ...e, fromInventory: false, inventoryItemId: null } : e))
          );
        }}
        onEdit={openEdit}
      />
      {open && (
        <Modal title={editingId ? "Edit inventory item" : "Add inventory item"} onClose={() => setOpen(false)}>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Item name">
              <input className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </Field>
            <Field label="Category">
              <input className={inputCls} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} />
            </Field>
            <Field label="Quantity">
              <input type="number" className={inputCls} value={form.quantity} onChange={(e) => setForm({ ...form, quantity: e.target.value })} />
            </Field>
            <Field label="Unit">
              <input className={inputCls} value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} />
            </Field>
            <Field label="Reorder level">
              <input type="number" className={inputCls} value={form.reorderLevel} onChange={(e) => setForm({ ...form, reorderLevel: e.target.value })} />
            </Field>
            <Field label="Unit cost (₹)">
              <input type="number" className={inputCls} value={form.unitCost} onChange={(e) => setForm({ ...form, unitCost: e.target.value })} />
            </Field>
          </div>
          {activeId && <AttachmentsField recordId={activeId} label="Purchase bills" />}
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update item" : "Save item"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- ASSET MANAGEMENT ----------
const ASSET_CATEGORIES = [
  "Computers", "Electrical Equipments & Fittings", "Furniture & Fixtures",
  "Office Equipment", "Safety Equipment", "Plant & Machinery",
];
const BLANK_ASSET = { name: "", category: ASSET_CATEGORIES[0], assetId: "", value: "", purchaseDate: "", expiryDate: "" };

const ASSET_EXPORT_COLUMNS = [
  { label: "Asset ID", value: (a) => a.assetId || "" },
  { label: "Description", value: (a) => a.name || "" },
  { label: "Category", value: (a) => a.category || "" },
  { label: "Purchase Date", value: (a) => a.purchaseDate || "" },
  { label: "Expiry Date", value: (a) => a.expiryDate || "" },
  { label: "Value", value: (a) => Number(a.value || 0) },
  { label: "Residual (5%)", value: (a) => Number(a.value || 0) * 0.05 },
  { label: "Source", value: (a) => (a.fromCapex ? "CAPEX" : "Manual") },
];

function AssetsTab({ assets, setAssets, company, setPrintContent, setPrintTitle }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [form, setForm] = useState(BLANK_ASSET);
  const [idError, setIdError] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("All");

  const openAdd = () => { setForm(BLANK_ASSET); setEditingId(null); setActiveId(uid()); setIdError(""); setOpen(true); };
  const openEdit = (row) => { setForm({ ...BLANK_ASSET, ...row }); setEditingId(row.id); setActiveId(row.id); setIdError(""); setOpen(true); };

  const save = () => {
    if (!form.name) return;
    const trimmedId = (form.assetId || "").trim();
    if (trimmedId) {
      const clash = assets.some(
        (a) => a.id !== editingId && (a.assetId || "").trim().toLowerCase() === trimmedId.toLowerCase()
      );
      if (clash) {
        setIdError(`Asset ID "${trimmedId}" is already assigned to another asset — use a different one.`);
        return;
      }
    }
    setIdError("");
    if (editingId) {
      setAssets(assets.map((a) => (a.id === editingId ? { ...a, ...form, id: editingId } : a)));
    } else {
      setAssets([...assets, { id: activeId, ...form }]);
    }
    setForm(BLANK_ASSET);
    setEditingId(null);
    setOpen(false);
  };

  const totalValue = assets.reduce((s, a) => s + Number(a.value || 0), 0);
  const RESIDUAL_RATE = 0.05;
  const residualValue = totalValue * RESIDUAL_RATE;
  const netValue = totalValue - residualValue;

  const filteredAssets = categoryFilter === "All" ? assets : assets.filter((a) => a.category === categoryFilter);

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center flex-wrap gap-3">
        <h2 className="font-bold text-lg text-neutral-900">Asset Management</h2>
        <div className="flex gap-2 items-center flex-wrap">
          <select
            className={inputCls}
            value={categoryFilter}
            onChange={(e) => setCategoryFilter(e.target.value)}
          >
            <option value="All">All Categories</option>
            {ASSET_CATEGORIES.map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
          <button
            onClick={() => {
              setPrintTitle(categoryFilter === "All" ? "Asset-Register" : `Asset-Register-${categoryFilter}`);
              setPrintContent(<AssetPrintLayout assets={filteredAssets} company={company} totalValue={totalValue} residualValue={residualValue} netValue={netValue} residualRate={RESIDUAL_RATE} />);
            }}
            className="inline-flex items-center gap-1.5 bg-white border border-neutral-300 hover:border-red-400 text-sm font-semibold px-4 py-2 rounded-lg transition"
          >
            <Printer size={16} /> Download PDF
          </button>
          <button
            onClick={() =>
              exportExcel(
                filteredAssets,
                ASSET_EXPORT_COLUMNS,
                categoryFilter === "All" ? "Asset-Register" : `Asset-Register-${categoryFilter}`,
                "Assets"
              )
            }
            className="inline-flex items-center gap-1.5 bg-white border border-neutral-300 hover:border-red-400 text-sm font-semibold px-4 py-2 rounded-lg transition"
          >
            <FileSpreadsheet size={16} /> Excel
          </button>
          <AddButton onClick={openAdd} text="Add asset" />
        </div>
      </div>
      <div className="grid grid-cols-2 gap-4">
        <Card label="Total Asset Value" value={fmt(totalValue)} sub="Sum of all recorded assets" />
        <Card label="Asset Net Value (After Residual)" value={fmt(netValue)} sub="Total value less 5% residual, for audit purposes" />
      </div>
      <Table
        emptyMsg={categoryFilter === "All" ? "No assets added yet." : `No assets in "${categoryFilter}" yet.`}
        columns={[
          { key: "assetId", label: "Asset ID", render: (r) => r.assetId || <span className="text-neutral-300">Not set</span> },
          { key: "name", label: "Asset Description" },
          { key: "category", label: "Category" },
          { key: "purchaseDate", label: "Purchase Date" },
          {
            key: "expiryDate",
            label: "Expiry Date",
            render: (r) => {
              const info = expiryTone(r.expiryDate);
              return r.expiryDate ? (
                <span>{r.expiryDate} {info && <Pill tone={info.tone}>{info.text}</Pill>}</span>
              ) : (
                <span className="text-neutral-300">—</span>
              );
            },
          },
          { key: "value", label: "Asset Value", render: (r) => fmt(r.value) },
          { key: "residual", label: "Residual (5%)", render: (r) => fmt(Number(r.value || 0) * RESIDUAL_RATE) },
          {
            key: "source",
            label: "Source",
            render: (r) => r.fromCapex ? <Pill tone="blue">CAPEX</Pill> : <Pill tone="gray">Manual</Pill>,
          },
        ]}
        rows={filteredAssets}
        onDelete={(id) => setAssets(assets.filter((a) => a.id !== id))}
        onEdit={openEdit}
      />
      {open && (
        <Modal title={editingId ? "Edit asset" : "Add asset"} onClose={() => setOpen(false)}>
          <div className="grid grid-cols-2 gap-3">
            <div className="col-span-2">
              <Field label="Asset List (name)">
                <input className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="e.g. Dell Latitude Laptop" />
              </Field>
            </div>
            <Field label="Category">
              <select className={inputCls} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>
                {ASSET_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="Asset ID">
              <input
                className={inputCls}
                value={form.assetId}
                onChange={(e) => { setForm({ ...form, assetId: e.target.value }); setIdError(""); }}
                placeholder="e.g. NOVA-AST-001"
              />
              {idError && <span className="text-xs text-red-600">{idError}</span>}
            </Field>
            <Field label="Asset Value (₹)">
              <input type="number" className={inputCls} value={form.value} onChange={(e) => setForm({ ...form, value: e.target.value })} />
            </Field>
            <Field label="Purchase Date">
              <input type="date" className={inputCls} value={form.purchaseDate} onChange={(e) => setForm({ ...form, purchaseDate: e.target.value })} />
            </Field>
            <Field label="Expiry Date (warranty/AMC, if any)">
              <input type="date" className={inputCls} value={form.expiryDate} onChange={(e) => setForm({ ...form, expiryDate: e.target.value })} />
            </Field>
          </div>
          {activeId && <AttachmentsField recordId={activeId} label="Asset bills / purchase proof" />}
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update asset" : "Save asset"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- ATTENDANCE ----------
// Two separate permission slot systems, both shown in the same dropdown (grouped):
// 9 one-hour slots covering the full work day, and 18 half-hour slots covering the
// same day at finer granularity. Shown only in the dropdown below — not spelled
// out in the HR policy text.
const PERMISSION_SLOTS_HOURLY = [
  "09:00 AM - 10:00 AM", "10:00 AM - 11:00 AM", "11:00 AM - 12:00 PM", "12:00 PM - 01:00 PM",
  "01:00 PM - 02:00 PM", "02:00 PM - 03:00 PM", "03:00 PM - 04:00 PM", "04:00 PM - 05:00 PM",
  "05:00 PM - 06:00 PM",
]; // 9 slots
const PERMISSION_SLOTS_HALF_HOURLY = [
  "09:00 AM - 09:30 AM", "09:30 AM - 10:00 AM", "10:00 AM - 10:30 AM", "10:30 AM - 11:00 AM",
  "11:00 AM - 11:30 AM", "11:30 AM - 12:00 PM", "12:00 PM - 12:30 PM", "12:30 PM - 01:00 PM",
  "01:00 PM - 01:30 PM", "01:30 PM - 02:00 PM", "02:00 PM - 02:30 PM", "02:30 PM - 03:00 PM",
  "03:00 PM - 03:30 PM", "03:30 PM - 04:00 PM", "04:00 PM - 04:30 PM", "04:30 PM - 05:00 PM",
  "05:00 PM - 05:30 PM", "05:30 PM - 06:00 PM",
]; // 18 slots
const PERMISSION_SLOTS = [...PERMISSION_SLOTS_HOURLY, ...PERMISSION_SLOTS_HALF_HOURLY];
// Renders both groups as <optgroup>s inside a <select>, so either granularity can
// be picked from the one dropdown.
function PermissionSlotOptions() {
  return (
    <>
      <optgroup label="1-hour slots">
        {PERMISSION_SLOTS_HOURLY.map((slot) => (
          <option key={slot} value={slot}>{slot}</option>
        ))}
      </optgroup>
      <optgroup label="30-min slots">
        {PERMISSION_SLOTS_HALF_HOURLY.map((slot) => (
          <option key={slot} value={slot}>{slot}</option>
        ))}
      </optgroup>
    </>
  );
}
// HR policy: 12 CL + 12 ML per calendar year (no carryover to next year),
// 4 Permissions per calendar month (no carryover to next month).
const LEAVE_LIMITS = { CL: 12, ML: 12, Permission: 4 };
const SHORT_CODE = { Present: "P", "Half-day": "HD", CL: "CL", "Half CL": "HCL", ML: "ML", Permission: "PM" };
function isSunday(dateStr) {
  return new Date(`${dateStr}T00:00:00`).getDay() === 0;
}
function daysInMonth(monthStr) {
  const [y, m] = monthStr.split("-").map(Number);
  return new Date(y, m, 0).getDate();
}
// Resolves the display code for a given employee/date: real record > Sunday holiday (H) > blank.
function codeForDate(records, employeeName, dateStr) {
  const rec = records.find((r) => r.employeeName === employeeName && r.date === dateStr);
  if (rec) return SHORT_CODE[rec.status] || rec.status;
  if (isSunday(dateStr)) return "H";
  return "-";
}
const CODE_TONE_BG = {
  P: "bg-emerald-50 text-emerald-700",
  HD: "bg-amber-50 text-amber-700",
  CL: "bg-blue-50 text-blue-700",
  HCL: "bg-sky-50 text-sky-600",
  ML: "bg-neutral-200 text-neutral-700",
  PM: "bg-amber-50 text-amber-700",
  H: "bg-neutral-100 text-neutral-400",
  "-": "text-neutral-300",
};

// Shows "2" instead of "2.0", but "2.5" when there's a genuine half-CL in the mix.
function formatLeaveCount(n) {
  return Number.isInteger(n) ? String(n) : n.toFixed(1);
}

function AttendanceTab({ records, setRecords, employees, company, setPrintContent, setPrintTitle }) {
  const [date, setDate] = useState(todayISO());
  const [draft, setDraft] = useState({});
  const [draftSlot, setDraftSlot] = useState({});
  const [editRow, setEditRow] = useState(null);

  const [historyMode, setHistoryMode] = useState("day");
  const [historyDay, setHistoryDay] = useState(todayISO());
  const [rangeFrom, setRangeFrom] = useState(todayISO());
  const [rangeTo, setRangeTo] = useState(todayISO());
  const [historyMonth, setHistoryMonth] = useState(thisMonthKey());

  const dayRecords = records.filter((r) => r.date === date);
  const statusFor = (name) => {
    const found = dayRecords.find((r) => r.employeeName === name);
    return draft[name] ?? found?.status ?? "";
  };
  const slotFor = (name) => {
    const found = dayRecords.find((r) => r.employeeName === name);
    return draftSlot[name] ?? found?.permissionSlot ?? PERMISSION_SLOTS[0];
  };

  const mark = (name, status) => setDraft({ ...draft, [name]: status });
  const markSlot = (name, slot) => setDraftSlot({ ...draftSlot, [name]: slot });

  const save = () => {
    if (date > todayISO()) return; // guard: never persist attendance for a future date
    const others = records.filter((r) => r.date !== date);
    const newDay = employees
      .map((e) => {
        const status = draft[e.name] ?? dayRecords.find((r) => r.employeeName === e.name)?.status ?? "Not marked";
        return {
          id: uid(),
          date,
          employeeName: e.name,
          status,
          permissionSlot: status === "Permission" ? slotFor(e.name) : undefined,
        };
      })
      .filter((r) => r.status !== "Not marked");
    setRecords([...others, ...newDay]);
    setDraft({});
    setDraftSlot({});
  };

  const statusOptions = ["Present", "Half-day", "CL", "Half CL", "ML", "Permission"];
  const toneOf = (s) =>
    s === "Present" ? "green" : s === "Half-day" ? "amber" : s === "CL" ? "blue" : s === "Half CL" ? "blue" : s === "ML" ? "gray" : s === "Permission" ? "amber" : "gray";

  // HR usage: CL/ML tracked per calendar year, Permission tracked per calendar month.
  const usageYear = date.slice(0, 4);
  const usageMonth = monthKey(date);
  const monthlyUsage = employees.map((e) => {
    const yearRows = records.filter((r) => r.employeeName === e.name && r.date.slice(0, 4) === usageYear);
    const monthRows = records.filter((r) => r.employeeName === e.name && monthKey(r.date) === usageMonth);
    const fullCL = yearRows.filter((r) => r.status === "CL").length;
    const halfCL = yearRows.filter((r) => r.status === "Half CL").length;
    return {
      name: e.name,
      cl: fullCL + halfCL * 0.5, // 2 × Half CL counts as 1 CL against the 12/year quota
      halfCL,
      ml: yearRows.filter((r) => r.status === "ML").length,
      permission: monthRows.filter((r) => r.status === "Permission").length,
    };
  });

  const filteredHistory = records
    .filter((r) => {
      if (historyMode === "day") return r.date === historyDay;
      if (historyMode === "range") return r.date >= rangeFrom && r.date <= rangeTo;
      if (historyMode === "month") return monthKey(r.date) === historyMonth;
      return true;
    })
    .sort((a, b) => (a.date < b.date ? 1 : -1));

  const periodLabel =
    historyMode === "day" ? historyDay :
    historyMode === "range" ? `${rangeFrom} to ${rangeTo}` :
    historyMonth;

  const downloadHistory = () => {
    setPrintTitle(`Attendance-${periodLabel}`);
    if (historyMode === "month") {
      setPrintContent(
        <AttendanceGridPrintLayout employees={employees} records={records} monthStr={historyMonth} company={company} />
      );
    } else {
      setPrintContent(
        <AttendancePrintLayout records={filteredHistory} employees={employees} company={company} periodLabel={periodLabel} />
      );
    }
  };

  const isFutureDate = date > todayISO();

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center flex-wrap gap-3">
        <h2 className="font-bold text-lg text-neutral-900">Attendance</h2>
        <div className="flex items-center gap-2">
          <input
            type="date"
            className={inputCls}
            value={date}
            max={todayISO()}
            onChange={(e) => setDate(e.target.value)}
          />
          <AddButton onClick={save} text="Save day" disabled={isFutureDate} />
        </div>
      </div>

      {isFutureDate && (
        <div className="bg-red-50 border border-red-200 rounded-xl px-4 py-2 text-xs text-red-700 flex items-center gap-2">
          <AlertTriangle size={14} className="shrink-0" />
          <span><strong>{date}</strong> is a future date. Attendance can only be recorded for today or an earlier date — this entry will not be saved.</span>
        </div>
      )}

      {isSunday(date) && (
        <div className="bg-neutral-100 border border-neutral-200 rounded-xl px-4 py-2 text-xs text-neutral-600">
          <strong>Sunday — Weekly Holiday (H).</strong> Marked automatically in history/reports — no action needed unless someone actually worked, in which case mark them below.
        </div>
      )}

      {employees.length === 0 ? (
        <div className="text-center text-neutral-400 text-sm py-14 border border-dashed border-neutral-200 rounded-xl">
          Add employees first (Employees & Salary tab) to start marking attendance.
        </div>
      ) : (
        <div className="border border-neutral-200 rounded-xl divide-y divide-neutral-100 bg-white">
          {employees.map((e) => (
            <div key={e.id} className="flex items-center justify-between px-4 py-3 flex-wrap gap-2">
              <div>
                <div className="font-medium text-neutral-800 text-sm">{e.name}</div>
                <div className="text-xs text-neutral-400">
                  {e.role}{e.department ? ` · ${e.department}` : ""}{e.employeeId ? ` · ID: ${e.employeeId}` : ""}
                </div>
              </div>
              <div className="flex items-center gap-1.5 flex-wrap">
                {statusOptions.map((s) => (
                  <button
                    key={s}
                    onClick={() => mark(e.name, s)}
                    className={`text-xs font-semibold px-2.5 py-1 rounded-full border transition ${
                      statusFor(e.name) === s
                        ? {
                            Present: "bg-emerald-600 text-white border-emerald-600",
                            "Half-day": "bg-amber-500 text-white border-amber-500",
                            CL: "bg-blue-600 text-white border-blue-600",
                            "Half CL": "bg-sky-400 text-white border-sky-400",
                            ML: "bg-neutral-600 text-white border-neutral-600",
                            Permission: "bg-amber-500 text-white border-amber-500",
                          }[s]
                        : "text-neutral-500 border-neutral-200 hover:bg-neutral-50"
                    }`}
                  >
                    {s}
                  </button>
                ))}
                {statusFor(e.name) === "Permission" && (
                  <select
                    className="border border-neutral-300 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-red-500"
                    value={slotFor(e.name)}
                    onChange={(ev) => markSlot(e.name, ev.target.value)}
                  >
                    <PermissionSlotOptions />
                  </select>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      <div className="bg-white border border-neutral-200 rounded-xl p-4">
        <div className="text-sm font-semibold text-neutral-800 mb-2">
          Leave usage — CL/ML for {usageYear}, Permission for {usageMonth}
        </div>
        <ul className="text-xs text-neutral-400 mb-3 space-y-0.5 list-disc pl-4">
          <li>12 Casual Leave (CL) per year — no carryover to next year.</li>
          <li>2 × Half CL = 1 CL, counted against the same 12/year quota.</li>
          <li>12 Medical Leave (ML) per year — no carryover to next year.</li>
          <li>4 Permissions per month — no carryover to next month.</li>
          <li>Sundays count as a weekly holiday (H) automatically.</li>
        </ul>
        <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
          {monthlyUsage.map((u) => {
            const over = u.cl > LEAVE_LIMITS.CL || u.ml > LEAVE_LIMITS.ML || u.permission > LEAVE_LIMITS.Permission;
            return (
              <div key={u.name} className={`text-xs rounded-lg border px-3 py-2 ${over ? "border-red-200 bg-red-50" : "border-neutral-200 bg-neutral-50"}`}>
                <div className="font-semibold text-neutral-800">{u.name}</div>
                <div className={u.cl > LEAVE_LIMITS.CL ? "text-red-600 font-semibold" : "text-neutral-500"}>
                  CL (year): {formatLeaveCount(u.cl)}/{LEAVE_LIMITS.CL}{u.halfCL > 0 ? ` (incl. ${u.halfCL} Half CL)` : ""}
                </div>
                <div className={u.ml > LEAVE_LIMITS.ML ? "text-red-600 font-semibold" : "text-neutral-500"}>ML (year): {u.ml}/{LEAVE_LIMITS.ML}</div>
                <div className={u.permission > LEAVE_LIMITS.Permission ? "text-red-600 font-semibold" : "text-neutral-500"}>Permission (month): {u.permission}/{LEAVE_LIMITS.Permission}</div>
              </div>
            );
          })}
        </div>
      </div>

      <div>
        <div className="flex justify-between items-center flex-wrap gap-3 mt-6 mb-2">
          <h3 className="font-semibold text-sm text-neutral-600">History</h3>
          <div className="flex items-center gap-2 flex-wrap">
            <select className={inputCls} value={historyMode} onChange={(e) => setHistoryMode(e.target.value)}>
              <option value="day">Single day</option>
              <option value="range">Date range</option>
              <option value="month">Month</option>
            </select>
            {historyMode === "day" && (
              <input type="date" className={inputCls} value={historyDay} onChange={(e) => setHistoryDay(e.target.value)} />
            )}
            {historyMode === "range" && (
              <>
                <input type="date" className={inputCls} value={rangeFrom} onChange={(e) => setRangeFrom(e.target.value)} />
                <span className="text-xs text-neutral-400">to</span>
                <input type="date" className={inputCls} value={rangeTo} onChange={(e) => setRangeTo(e.target.value)} />
              </>
            )}
            {historyMode === "month" && (
              <input type="month" className={inputCls} value={historyMonth} onChange={(e) => setHistoryMonth(e.target.value)} />
            )}
            <button
              onClick={downloadHistory}
              className="inline-flex items-center gap-1.5 bg-white border border-neutral-300 hover:border-red-400 text-sm font-semibold px-3 py-2 rounded-lg transition"
            >
              <Printer size={15} /> Download PDF
            </button>
          </div>
        </div>
        {historyMode === "month" ? (
          <MonthlyAttendanceGrid employees={employees} records={records} monthStr={historyMonth} />
        ) : (
          <Table
            emptyMsg="No attendance records for this period."
            columns={[
              { key: "date", label: "Date" },
              { key: "employeeName", label: "Employee" },
              { key: "status", label: "Status", render: (r) => (
                <Pill tone={toneOf(r.status)}>
                  {r.status}{r.status === "Permission" && r.permissionSlot ? ` (${r.permissionSlot})` : ""}
                </Pill>
              ) },
            ]}
            rows={filteredHistory}
            onDelete={(id) => setRecords(records.filter((r) => r.id !== id))}
            onEdit={(row) => setEditRow({ ...row })}
          />
        )}
      </div>

      {editRow && (
        <Modal title="Edit attendance record" onClose={() => setEditRow(null)}>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Date">
              <input type="date" className={inputCls} value={editRow.date} onChange={(e) => setEditRow({ ...editRow, date: e.target.value })} />
            </Field>
            <Field label="Status">
              <select className={inputCls} value={editRow.status} onChange={(e) => setEditRow({ ...editRow, status: e.target.value })}>
                {statusOptions.map((s) => (
                  <option key={s}>{s}</option>
                ))}
              </select>
            </Field>
            {editRow.status === "Permission" && (
              <Field label="Permission slot">
                <select className={inputCls} value={editRow.permissionSlot || PERMISSION_SLOTS[0]} onChange={(e) => setEditRow({ ...editRow, permissionSlot: e.target.value })}>
                  <PermissionSlotOptions />
                </select>
              </Field>
            )}
          </div>
          <div className="mt-4 flex justify-end">
            <AddButton
              onClick={() => {
                setRecords(records.map((r) => (r.id === editRow.id ? editRow : r)));
                setEditRow(null);
              }}
              text="Update record"
            />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- EMPLOYEES ----------
const BLANK_EMPLOYEE = {
  name: "", employeeId: "", role: "", department: "", contact: "", workStatus: "Active", joiningDate: "", salary: "",
  bankName: "", accountNumber: "", ifsc: "",
  pfNumber: "", pfContribution: "", esiNumber: "", esiContribution: "",
  pan: "", aadhaar: "",
};

function EmployeesTab({ employees, setEmployees, company, setPrintContent, setPrintTitle }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(BLANK_EMPLOYEE);
  const [idError, setIdError] = useState("");

  const openAdd = () => { setForm(BLANK_EMPLOYEE); setEditingId(null); setIdError(""); setOpen(true); };
  const openEdit = (row) => { setForm({ ...BLANK_EMPLOYEE, ...row }); setEditingId(row.id); setIdError(""); setOpen(true); };

  const save = () => {
    if (!form.name) return;
    const trimmedId = form.employeeId.trim();
    if (trimmedId) {
      const clash = employees.some(
        (e) => e.id !== editingId && e.employeeId.trim().toLowerCase() === trimmedId.toLowerCase()
      );
      if (clash) {
        setIdError(`Employee ID "${trimmedId}" is already assigned to another employee — use a different one.`);
        return;
      }
    }
    setIdError("");
    if (editingId) {
      setEmployees(employees.map((e) => (e.id === editingId ? { ...e, ...form, id: editingId } : e)));
    } else {
      setEmployees([...employees, { id: uid(), ...form }]);
    }
    setForm(BLANK_EMPLOYEE);
    setEditingId(null);
    setOpen(false);
  };

  const totalSalary = employees.reduce((s, e) => s + Number(e.salary || 0), 0);
  const totalPF = employees.reduce((s, e) => s + Number(e.pfContribution || 0), 0);
  const totalESI = employees.reduce((s, e) => s + Number(e.esiContribution || 0), 0);
  const sortedEmployees = [...employees].sort((a, b) =>
    (a.employeeId || "").localeCompare(b.employeeId || "", undefined, { numeric: true, sensitivity: "base" })
  );

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="font-bold text-lg text-neutral-900">Employees & Salary</h2>
          <p className="text-xs text-neutral-400 mt-0.5">
            Monthly payroll: {fmt(totalSalary)} · PF: {fmt(totalPF)} · ESI: {fmt(totalESI)}
          </p>
        </div>
        <div className="flex gap-2">
          <button
            onClick={() => { setPrintTitle("Employee-Register"); setPrintContent(<EmployeeListPrintLayout employees={sortedEmployees} company={company} />); }}
            className="inline-flex items-center gap-1.5 bg-white border border-neutral-300 hover:border-red-400 text-sm font-semibold px-4 py-2 rounded-lg transition"
          >
            <Printer size={16} /> Download PDF
          </button>
          <AddButton onClick={openAdd} text="Add employee" />
        </div>
      </div>
      <Table
        emptyMsg="No employees added yet."
        columns={[
          { key: "employeeId", label: "Emp ID" },
          { key: "name", label: "Name" },
          { key: "role", label: "Role" },
          { key: "department", label: "Department" },
          { key: "workStatus", label: "Status", render: (r) => <Pill tone={r.workStatus === "Active" ? "green" : r.workStatus === "Resigned" ? "red" : "amber"}>{r.workStatus}</Pill> },
          { key: "joiningDate", label: "Joined" },
          { key: "salary", label: "Monthly salary", render: (r) => fmt(r.salary) },
          { key: "pfNumber", label: "PF no." },
          { key: "esiNumber", label: "ESI no." },
        ]}
        rows={sortedEmployees}
        onDelete={(id) => setEmployees(employees.filter((e) => e.id !== id))}
        onEdit={openEdit}
      />
      {open && (
        <Modal title={editingId ? "Edit employee" : "Add employee"} onClose={() => setOpen(false)}>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Full name">
              <input className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </Field>
            <Field label="Employee ID">
              <input
                className={inputCls}
                value={form.employeeId}
                onChange={(e) => { setForm({ ...form, employeeId: e.target.value }); setIdError(""); }}
              />
              {idError && <span className="text-xs text-red-600">{idError}</span>}
            </Field>
            <Field label="Role / designation">
              <input className={inputCls} value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })} />
            </Field>
            <Field label="Department">
              <input className={inputCls} value={form.department} onChange={(e) => setForm({ ...form, department: e.target.value })} />
            </Field>
            <Field label="Work status">
              <select className={inputCls} value={form.workStatus} onChange={(e) => setForm({ ...form, workStatus: e.target.value })}>
                <option>Active</option>
                <option>On Leave</option>
                <option>Resigned</option>
              </select>
            </Field>
            <Field label="Contact">
              <input
                className={inputCls}
                value={form.contact}
                maxLength={10}
                placeholder="10-digit mobile number"
                onChange={(e) => setForm({ ...form, contact: formatPhone10(e.target.value) })}
              />
              {form.contact && form.contact.length !== 10 && (
                <span className="text-[11px] text-red-600">Contact number must be exactly 10 digits.</span>
              )}
            </Field>
            <Field label="Joining date">
              <input type="date" className={inputCls} value={form.joiningDate} onChange={(e) => setForm({ ...form, joiningDate: e.target.value })} />
            </Field>
            <Field label="Monthly salary (₹)">
              <input type="number" className={inputCls} value={form.salary} onChange={(e) => setForm({ ...form, salary: e.target.value })} />
            </Field>

            <div className="col-span-2 border-t border-neutral-100 pt-3 text-sm font-semibold text-neutral-700">Bank details</div>
            <Field label="Bank name">
              <input className={inputCls} value={form.bankName} onChange={(e) => setForm({ ...form, bankName: e.target.value })} />
            </Field>
            <Field label="Account number">
              <input className={inputCls} value={form.accountNumber} onChange={(e) => setForm({ ...form, accountNumber: e.target.value })} />
            </Field>
            <Field label="IFSC code">
              <input className={inputCls} value={form.ifsc} onChange={(e) => setForm({ ...form, ifsc: e.target.value })} />
            </Field>

            <div className="col-span-2 border-t border-neutral-100 pt-3 text-sm font-semibold text-neutral-700">Statutory details</div>
            <Field label="PF number (UAN)">
              <input className={inputCls} value={form.pfNumber} onChange={(e) => setForm({ ...form, pfNumber: e.target.value })} />
            </Field>
            <Field label="PF contribution (₹/month)">
              <input type="number" className={inputCls} value={form.pfContribution} onChange={(e) => setForm({ ...form, pfContribution: e.target.value })} />
            </Field>
            <Field label="ESI number">
              <input className={inputCls} value={form.esiNumber} onChange={(e) => setForm({ ...form, esiNumber: e.target.value })} />
            </Field>
            <Field label="ESI contribution (₹/month)">
              <input type="number" className={inputCls} value={form.esiContribution} onChange={(e) => setForm({ ...form, esiContribution: e.target.value })} />
            </Field>
            <Field label="PAN number">
              <input
                className={inputCls}
                value={form.pan}
                maxLength={10}
                placeholder="ABCDE1234F"
                onChange={(e) => setForm({ ...form, pan: formatPanStrict(e.target.value) })}
              />
              <span className="text-[11px] text-neutral-400">Format: 5 letters + 4 digits + 1 letter (e.g. ABCDE1234F).</span>
              {form.pan && form.pan.length < 10 && (
                <span className="text-xs text-amber-600">PAN must be exactly 10 characters — {10 - form.pan.length} more needed.</span>
              )}
            </Field>
            <Field label="Aadhaar number">
              <input
                className={inputCls}
                value={form.aadhaar}
                maxLength={14}
                placeholder="1234-5678-9012"
                onChange={(e) => setForm({ ...form, aadhaar: formatAadhaar(e.target.value) })}
              />
            </Field>
          </div>
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update employee" : "Save employee"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

function PartyRegistrationPrintLayout({ reg, company }) {
  const labelCell = { border: "1px solid #ddd", padding: "6px", fontSize: "12px", background: "#f8f8f8", fontWeight: "bold", width: "28%" };
  const cell = { border: "1px solid #ddd", padding: "6px", fontSize: "12px" };
  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "800px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>
      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 4px" }}>
        {reg.partyType === "Vendor" ? "VENDOR" : "CLIENT"} PROFILE
      </div>
      <div style={{ textAlign: "center", fontSize: "12px", color: "#555", marginBottom: "16px" }}>
        Company on record with {company.name}
      </div>

      <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "14px" }}>
        <tbody>
          <tr><td style={labelCell}>{reg.partyType} Name</td><td style={cell} colSpan={3}>{reg.name}</td></tr>
          <tr><td style={labelCell}>Registered Office Address</td><td style={cell} colSpan={3}>{reg.address || "—"}</td></tr>
          <tr><td style={labelCell}>GSTIN No.</td><td style={cell}>{reg.gstin || "—"}</td><td style={labelCell}>MSME Status</td><td style={cell}>{reg.msmeStatus}</td></tr>
        </tbody>
      </table>

      <div style={{ fontSize: "13px", fontWeight: "bold", margin: "12px 0 6px" }}>Bank Details</div>
      <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "14px" }}>
        <tbody>
          <tr><td style={labelCell}>Beneficiary Name</td><td style={cell} colSpan={3}>{reg.beneficiaryName || "—"}</td></tr>
          <tr><td style={labelCell}>Bank Name</td><td style={cell}>{reg.bankName || "—"}</td><td style={labelCell}>Account Number</td><td style={cell}>{reg.accountNumber || "—"}</td></tr>
          <tr><td style={labelCell}>IFSC Code</td><td style={cell} colSpan={3}>{reg.ifsc || "—"}</td></tr>
        </tbody>
      </table>

      <div style={{ fontSize: "13px", fontWeight: "bold", margin: "12px 0 6px" }}>Authorized Contacts</div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Authority</th>
            <th style={cell}>Name</th>
            <th style={cell}>Contact No.</th>
            <th style={cell}>Email ID</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style={cell}>Primary</td>
            <td style={cell}>{reg.authority1Name || "—"}</td>
            <td style={cell}>{reg.authority1Contact || "—"}</td>
            <td style={cell}>{reg.authority1Email || "—"}</td>
          </tr>
          {(reg.authority2Name || reg.authority2Contact || reg.authority2Email) && (
            <tr>
              <td style={cell}>Secondary</td>
              <td style={cell}>{reg.authority2Name || "—"}</td>
              <td style={cell}>{reg.authority2Contact || "—"}</td>
              <td style={cell}>{reg.authority2Email || "—"}</td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

function OrdersPrintLayout({ orders, company }) {
  const cell = { border: "1px solid #ddd", padding: "6px", fontSize: "11px" };
  const totalValue = orders.reduce((s, o) => s + Number(o.value || 0), 0);
  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "1000px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>
      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 16px" }}>
        ORDERS / PROJECTS REGISTER
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Order / Project</th>
            <th style={cell}>Client</th>
            <th style={cell}>Status</th>
            <th style={cell}>Start</th>
            <th style={cell}>Deadline</th>
            <th style={{ ...cell, textAlign: "right" }}>Value</th>
          </tr>
        </thead>
        <tbody>
          {orders.map((o, idx) => (
            <tr key={o.id || idx}>
              <td style={cell}>{o.name}</td>
              <td style={cell}>{o.client}</td>
              <td style={cell}>{o.status}</td>
              <td style={cell}>{o.startDate}</td>
              <td style={cell}>{o.deadline}</td>
              <td style={{ ...cell, textAlign: "right" }}>{fmt(o.value)}</td>
            </tr>
          ))}
          <tr style={{ fontWeight: "bold" }}>
            <td style={cell} colSpan={5}>Total Value</td>
            <td style={{ ...cell, textAlign: "right" }}>{fmt(totalValue)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function InventoryPrintLayout({ items, company }) {
  const cell = { border: "1px solid #ddd", padding: "6px", fontSize: "11px" };
  const totalValue = items.reduce((s, i) => s + Number(i.quantity || 0) * Number(i.unitCost || 0), 0);
  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "1000px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>
      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 16px" }}>
        INVENTORY REGISTER
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Item</th>
            <th style={cell}>Category</th>
            <th style={{ ...cell, textAlign: "right" }}>Qty</th>
            <th style={{ ...cell, textAlign: "right" }}>Reorder At</th>
            <th style={{ ...cell, textAlign: "right" }}>Unit Cost</th>
            <th style={{ ...cell, textAlign: "right" }}>Value</th>
          </tr>
        </thead>
        <tbody>
          {items.map((i, idx) => (
            <tr key={i.id || idx}>
              <td style={cell}>{i.name}</td>
              <td style={cell}>{i.category}</td>
              <td style={{ ...cell, textAlign: "right" }}>{i.quantity} {i.unit}</td>
              <td style={{ ...cell, textAlign: "right" }}>{i.reorderLevel}</td>
              <td style={{ ...cell, textAlign: "right" }}>{fmt(i.unitCost)}</td>
              <td style={{ ...cell, textAlign: "right" }}>{fmt(Number(i.quantity || 0) * Number(i.unitCost || 0))}</td>
            </tr>
          ))}
          <tr style={{ fontWeight: "bold" }}>
            <td style={cell} colSpan={5}>Total Stock Value</td>
            <td style={{ ...cell, textAlign: "right" }}>{fmt(totalValue)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function AssetPrintLayout({ assets, company, totalValue, residualValue, netValue, residualRate }) {
  const cell = { border: "1px solid #ddd", padding: "6px", fontSize: "11px" };
  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "1000px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>
      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 4px" }}>
        ASSET REGISTER
      </div>
      <div style={{ textAlign: "center", fontSize: "11px", color: "#555", marginBottom: "16px" }}>
        Total Asset Value: {fmt(totalValue)} &nbsp;·&nbsp; Residual ({Math.round(residualRate * 100)}%): {fmt(residualValue)} &nbsp;·&nbsp; Net Value: {fmt(netValue)}
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Asset ID</th>
            <th style={cell}>Description</th>
            <th style={cell}>Category</th>
            <th style={cell}>Purchase Date</th>
            <th style={cell}>Expiry Date</th>
            <th style={{ ...cell, textAlign: "right" }}>Value</th>
            <th style={{ ...cell, textAlign: "right" }}>Residual</th>
          </tr>
        </thead>
        <tbody>
          {assets.map((a, idx) => (
            <tr key={a.id || idx}>
              <td style={cell}>{a.assetId || "—"}</td>
              <td style={cell}>{a.name}</td>
              <td style={cell}>{a.category}</td>
              <td style={cell}>{a.purchaseDate}</td>
              <td style={cell}>{a.expiryDate}</td>
              <td style={{ ...cell, textAlign: "right" }}>{fmt(a.value)}</td>
              <td style={{ ...cell, textAlign: "right" }}>{fmt(Number(a.value || 0) * residualRate)}</td>
            </tr>
          ))}
          <tr style={{ fontWeight: "bold" }}>
            <td style={cell} colSpan={5}>Total</td>
            <td style={{ ...cell, textAlign: "right" }}>{fmt(totalValue)}</td>
            <td style={{ ...cell, textAlign: "right" }}>{fmt(residualValue)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function EmployeeListPrintLayout({ employees, company }) {
  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "1000px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>
      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 16px" }}>
        EMPLOYEE REGISTER
      </div>
      <table style={{ width: "100%", fontSize: "11px", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            {["Emp ID", "Name", "Role", "Department", "Status", "Joined", "Salary (₹)", "PF No.", "ESI No.", "Contact"].map((h) => (
              <th key={h} style={{ border: "1px solid #ddd", padding: "5px", textAlign: "left" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {employees.map((e) => (
            <tr key={e.id}>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.employeeId}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.name}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.role}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.department}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.workStatus}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.joiningDate}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px", textAlign: "right" }}>{fmt(e.salary)}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.pfNumber}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.esiNumber}</td>
              <td style={{ border: "1px solid #ddd", padding: "5px" }}>{e.contact}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function FinanceListPrintLayout({ entries, type, total, company, isCapex }) {
  const cell = { border: "1px solid #ddd", padding: "6px", fontSize: "11px" };
  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "900px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>
      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 16px" }}>
        {type.toUpperCase()} REGISTER
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Date</th>
            <th style={cell}>Category</th>
            <th style={cell}>Description</th>
            {isCapex && <th style={{ ...cell, textAlign: "right" }}>Unit Price</th>}
            {isCapex && <th style={{ ...cell, textAlign: "right" }}>No. of Units</th>}
            <th style={{ ...cell, textAlign: "right" }}>{isCapex ? "Total Price" : "Amount"}</th>
          </tr>
        </thead>
        <tbody>
          {entries.map((e, i) => (
            <tr key={e.id || i}>
              <td style={cell}>{e.date}</td>
              <td style={cell}>{e.category}</td>
              <td style={cell}>{e.description}</td>
              {isCapex && <td style={{ ...cell, textAlign: "right" }}>{fmt(e.unitPrice)}</td>}
              {isCapex && <td style={{ ...cell, textAlign: "right" }}>{e.units}</td>}
              <td style={{ ...cell, textAlign: "right" }}>{fmt(e.amount)}</td>
            </tr>
          ))}
          <tr style={{ fontWeight: "bold" }}>
            <td style={cell} colSpan={isCapex ? 5 : 3}>Total {type}</td>
            <td style={{ ...cell, textAlign: "right" }}>{fmt(total)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function MonthlyAttendanceGrid({ employees, records, monthStr }) {
  const numDays = daysInMonth(monthStr);
  const days = Array.from({ length: numDays }, (_, i) => i + 1);
  const dateFor = (d) => `${monthStr}-${String(d).padStart(2, "0")}`;

  if (employees.length === 0) {
    return (
      <div className="text-center text-neutral-400 text-sm py-14 border border-dashed border-neutral-200 rounded-xl">
        Add employees first to see the monthly grid.
      </div>
    );
  }

  return (
    <div className="border border-neutral-200 rounded-xl bg-white overflow-hidden">
      <div className="px-3 py-2 text-xs text-neutral-400 border-b border-neutral-100">
        P=Present · HD=Half-day · CL=Casual Leave · HCL=Half CL · ML=Medical Leave · PM=Permission · H=Sunday Holiday · -=No record
      </div>
      <div className="overflow-x-auto">
        <table className="text-xs w-full">
          <thead>
            <tr className="bg-neutral-50 text-neutral-500">
              <th className="px-2 py-2 text-left whitespace-nowrap">Emp ID</th>
              <th className="px-2 py-2 text-left whitespace-nowrap">Name</th>
              {days.map((d) => <th key={d} className="px-1 py-2 text-center">{d}</th>)}
              <th className="px-2 py-2 text-center">CL</th>
              <th className="px-2 py-2 text-center">ML</th>
              <th className="px-2 py-2 text-center">PM</th>
            </tr>
          </thead>
          <tbody>
            {employees.map((e) => {
              const codes = days.map((d) => codeForDate(records, e.name, dateFor(d)));
              const cl = codes.filter((c) => c === "CL").length;
              const ml = codes.filter((c) => c === "ML").length;
              const pm = codes.filter((c) => c === "PM").length;
              return (
                <tr key={e.id} className="border-t border-neutral-100">
                  <td className="px-2 py-1.5 whitespace-nowrap text-neutral-600">{e.employeeId}</td>
                  <td className="px-2 py-1.5 whitespace-nowrap font-medium text-neutral-800">{e.name}</td>
                  {codes.map((code, i) => (
                    <td key={i} className={`px-1 py-1.5 text-center ${CODE_TONE_BG[code] || "text-neutral-500"}`}>{code}</td>
                  ))}
                  <td className="px-2 py-1.5 text-center font-semibold text-neutral-800">{cl}</td>
                  <td className="px-2 py-1.5 text-center font-semibold text-neutral-800">{ml}</td>
                  <td className="px-2 py-1.5 text-center font-semibold text-neutral-800">{pm}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function AttendanceGridPrintLayout({ employees, records, monthStr, company }) {
  const numDays = daysInMonth(monthStr);
  const days = Array.from({ length: numDays }, (_, i) => i + 1);
  const dateFor = (d) => `${monthStr}-${String(d).padStart(2, "0")}`;
  const cell = { border: "1px solid #ddd", padding: "3px", fontSize: "9px", textAlign: "center" };
  const nameCell = { border: "1px solid #ddd", padding: "3px", fontSize: "9px", textAlign: "left", whiteSpace: "nowrap" };

  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "18px" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "8px", marginBottom: "12px" }}>
        <div style={{ fontSize: "18px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "11px", color: "#555" }}>{company.address}</div>
      </div>
      <div style={{ textAlign: "center", fontSize: "14px", fontWeight: "bold", letterSpacing: "1px", margin: "8px 0 4px" }}>
        MONTHLY ATTENDANCE — {monthStr}
      </div>
      <div style={{ textAlign: "center", fontSize: "10px", color: "#555", marginBottom: "10px" }}>
        P=Present · HD=Half-day · CL=Casual Leave · HCL=Half CL · ML=Medical Leave · PM=Permission · H=Sunday Holiday · -=No record
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={nameCell}>Emp ID</th>
            <th style={nameCell}>Name</th>
            {days.map((d) => <th key={d} style={cell}>{d}</th>)}
            <th style={cell}>CL</th>
            <th style={cell}>ML</th>
            <th style={cell}>PM</th>
          </tr>
        </thead>
        <tbody>
          {employees.map((e) => {
            const codes = days.map((d) => codeForDate(records, e.name, dateFor(d)));
            const cl = codes.filter((c) => c === "CL").length;
            const ml = codes.filter((c) => c === "ML").length;
            const pm = codes.filter((c) => c === "PM").length;
            return (
              <tr key={e.id}>
                <td style={nameCell}>{e.employeeId}</td>
                <td style={nameCell}>{e.name}</td>
                {codes.map((code, i) => <td key={i} style={cell}>{code}</td>)}
                <td style={{ ...cell, fontWeight: "bold" }}>{cl}</td>
                <td style={{ ...cell, fontWeight: "bold" }}>{ml}</td>
                <td style={{ ...cell, fontWeight: "bold" }}>{pm}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function AttendancePrintLayout({ records, employees, company, periodLabel }) {
  const summary = employees.map((e) => {
    const rows = records.filter((r) => r.employeeName === e.name);
    return {
      name: e.name,
      present: rows.filter((r) => r.status === "Present").length,
      halfDay: rows.filter((r) => r.status === "Half-day").length,
      cl: rows.filter((r) => r.status === "CL").length,
      ml: rows.filter((r) => r.status === "ML").length,
      permission: rows.filter((r) => r.status === "Permission").length,
    };
  });
  const cell = { border: "1px solid #ddd", padding: "5px", fontSize: "11px" };

  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "1000px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>
      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 4px" }}>
        ATTENDANCE REPORT
      </div>
      <div style={{ textAlign: "center", fontSize: "12px", marginBottom: "16px" }}>
        Period: <strong>{periodLabel}</strong>
      </div>

      <div style={{ fontSize: "13px", fontWeight: "bold", margin: "12px 0 6px" }}>Summary</div>
      <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "16px" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Employee</th>
            <th style={{ ...cell, textAlign: "right" }}>Present</th>
            <th style={{ ...cell, textAlign: "right" }}>Half-day</th>
            <th style={{ ...cell, textAlign: "right" }}>CL</th>
            <th style={{ ...cell, textAlign: "right" }}>ML</th>
            <th style={{ ...cell, textAlign: "right" }}>Permission</th>
          </tr>
        </thead>
        <tbody>
          {summary.map((s) => (
            <tr key={s.name}>
              <td style={cell}>{s.name}</td>
              <td style={{ ...cell, textAlign: "right" }}>{s.present}</td>
              <td style={{ ...cell, textAlign: "right" }}>{s.halfDay}</td>
              <td style={{ ...cell, textAlign: "right" }}>{s.cl}</td>
              <td style={{ ...cell, textAlign: "right" }}>{s.ml}</td>
              <td style={{ ...cell, textAlign: "right" }}>{s.permission}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <div style={{ fontSize: "13px", fontWeight: "bold", margin: "12px 0 6px" }}>Daily Records</div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Date</th>
            <th style={cell}>Employee</th>
            <th style={cell}>Status</th>
            <th style={cell}>Permission Slot</th>
          </tr>
        </thead>
        <tbody>
          {records.map((r, i) => (
            <tr key={r.id || i}>
              <td style={cell}>{r.date}</td>
              <td style={cell}>{r.employeeName}</td>
              <td style={cell}>{r.status}</td>
              <td style={cell}>{r.status === "Permission" ? r.permissionSlot || "" : ""}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ---------- FINANCE (CapEx / OpEx) ----------
function FinanceTab({ entries, setEntries, type, assets, setAssets, company, setPrintContent, setPrintTitle }) {
  const isCapex = type === "CapEx";
  const blank = { type, category: "", description: "", unitPrice: "", units: "1", amount: "", date: todayISO() };
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [form, setForm] = useState(blank);
  const [query, setQuery] = useState("");

  const filtered = entries.filter((e) => e.type === type);
  const q = query.trim().toLowerCase();
  const searchedEntries = q ? filtered.filter((e) => (e.category || "").toLowerCase().includes(q)) : filtered;

  const openAdd = () => { setForm(blank); setEditingId(null); setActiveId(uid()); setOpen(true); };
  const openEdit = (row) => { setForm({ ...blank, ...row }); setEditingId(row.id); setActiveId(row.id); setOpen(true); };

  const computedTotal = isCapex ? Number(form.unitPrice || 0) * Number(form.units || 0) : Number(form.amount || 0);

  const save = () => {
    const finalAmount = isCapex ? computedTotal : Number(form.amount || 0);
    if (!form.description || !finalAmount) return;
    const payload = { ...form, amount: finalAmount, type };
    if (editingId) {
      setEntries(entries.map((e) => (e.id === editingId ? { ...e, ...payload, id: editingId } : e)));
      // If this CapEx row was already sent to Asset Management, keep that asset's
      // value/name/date in sync — otherwise editing unit price or no. of units here
      // (e.g. correcting units from 1 to 2) leaves the Asset record frozen at the
      // old, un-updated total.
      if (isCapex) {
        setAssets((prev) =>
          prev.map((a) =>
            a.capexEntryId === editingId
              ? { ...a, name: payload.description, value: finalAmount, purchaseDate: payload.date }
              : a
          )
        );
      }
    } else {
      setEntries([...entries, { id: activeId, ...payload }]);
    }
    setForm(blank);
    setEditingId(null);
    setOpen(false);
  };

  const sendToAssets = (entry) => {
    // Idempotency guard: check both the entry's own flag AND whether an asset
    // already exists for this CapEx entry, so a stale click or race never
    // creates a second linked asset for the same CapEx row.
    if (entry.sentToAssets || assets.some((a) => a.capexEntryId === entry.id)) return;
    const matchedCategory =
      ASSET_CATEGORIES.find((c) => c.toLowerCase() === (entry.category || "").toLowerCase()) || "Plant & Machinery";
    const newAsset = {
      id: uid(),
      name: entry.description,
      category: matchedCategory,
      assetId: "",
      value: entry.amount,
      purchaseDate: entry.date,
      expiryDate: "",
      fromCapex: true,
      capexEntryId: entry.id,
    };
    setAssets([...assets, newAsset]);
    setEntries(entries.map((e) => (e.id === entry.id ? { ...e, sentToAssets: true } : e)));
  };

  const total = filtered.reduce((s, e) => s + Number(e.amount), 0);
  const sub = type === "CapEx" ? "Machinery, equipment, land etc." : "Rent, utilities, materials etc.";

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center flex-wrap gap-3">
        <h2 className="font-bold text-lg text-neutral-900">{type}</h2>
        <div className="flex gap-2 items-center flex-wrap">
          <div className="relative">
            <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              className={`${inputCls} pl-8 w-48`}
              placeholder="Search category…"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <button
            onClick={() => {
              setPrintTitle(`${type}-Register`);
              setPrintContent(
                <FinanceListPrintLayout entries={[...searchedEntries].sort((a, b) => (a.date < b.date ? 1 : -1))} type={type} total={total} company={company} isCapex={isCapex} />
              );
            }}
            className="inline-flex items-center gap-1.5 bg-white border border-neutral-300 hover:border-red-400 text-sm font-semibold px-4 py-2 rounded-lg transition"
          >
            <Printer size={16} /> Download PDF
          </button>
          <AddButton onClick={openAdd} text="Add entry" />
        </div>
      </div>
      <Card label={`Total ${type}`} value={fmt(total)} sub={sub} />
      {isCapex && (
        <p className="text-xs text-neutral-400 -mt-2">
          Tick "Add to Assets" on any row to send its details straight into Asset Management (Asset ID and expiry stay editable there).
        </p>
      )}
      <Table
        emptyMsg={q ? `No ${type} entries match "${query}".` : `No ${type} entries yet.`}
        columns={
          isCapex
            ? [
                { key: "date", label: "Date" },
                { key: "category", label: "Category" },
                { key: "description", label: "Description" },
                { key: "unitPrice", label: "Unit Price", render: (r) => fmt(r.unitPrice) },
                { key: "units", label: "No. of Units" },
                { key: "amount", label: "Total Price", render: (r) => fmt(r.amount) },
                {
                  key: "sendToAssets",
                  label: "Add to Assets",
                  render: (r) =>
                    r.sentToAssets ? (
                      <span className="text-xs text-emerald-600 font-semibold whitespace-nowrap">✓ In Assets</span>
                    ) : (
                      <input
                        type="checkbox"
                        checked={false}
                        onChange={(e) => { e.target.disabled = true; sendToAssets(r); }}
                        className="w-4 h-4 accent-red-600 cursor-pointer"
                        title="Send this entry to Asset Management"
                      />
                    ),
                },
              ]
            : [
                { key: "date", label: "Date" },
                { key: "category", label: "Category" },
                { key: "description", label: "Description" },
                { key: "amount", label: "Amount", render: (r) => fmt(r.amount) },
              ]
        }
        rows={[...searchedEntries].sort((a, b) => (a.date < b.date ? 1 : -1))}
        onDelete={(id) => {
          setEntries(entries.filter((e) => e.id !== id));
          // Don't leave a linked Asset pointing at a deleted CapEx row — unlink it
          // (keep the asset itself, just mark it Manual) instead of a dangling reference.
          if (isCapex) {
            setAssets((prev) =>
              prev.map((a) => (a.capexEntryId === id ? { ...a, fromCapex: false, capexEntryId: null } : a))
            );
          }
        }}
        onEdit={openEdit}
      />
      {open && (
        <Modal title={editingId ? `Edit ${type} entry` : `Add ${type} entry`} onClose={() => setOpen(false)}>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Category">
              <input className={inputCls} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} placeholder={type === "CapEx" ? "e.g. Machinery, Equipment" : "e.g. Rent, Raw material"} />
            </Field>
            <Field label="Date">
              <input type="date" className={inputCls} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
            </Field>
            {isCapex ? (
              <>
                <Field label="Unit Price (₹)">
                  <input type="number" className={inputCls} value={form.unitPrice} onChange={(e) => setForm({ ...form, unitPrice: e.target.value })} />
                </Field>
                <Field label="No. of Units">
                  <input type="number" className={inputCls} value={form.units} onChange={(e) => setForm({ ...form, units: e.target.value })} />
                </Field>
                <div className="col-span-2 text-right text-sm font-semibold text-neutral-800">
                  Total Price: {fmt(computedTotal)}
                </div>
              </>
            ) : (
              <Field label="Amount (₹)">
                <input type="number" className={inputCls} value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} />
              </Field>
            )}
            <div className="col-span-2">
              <Field label="Description">
                <input className={inputCls} value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} />
              </Field>
            </div>
          </div>
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update entry" : "Save entry"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- ORDERS / PROJECTS ----------
function OrdersTab({ orders, setOrders, company, partyRegistrations, setPrintContent, setPrintTitle }) {
  const blank = { name: "", client: "", status: "In progress", startDate: todayISO(), deadline: "", value: "" };
  const clientRecords = partyRegistrations.filter((r) => r.partyType === "Client");
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [form, setForm] = useState(blank);
  const [query, setQuery] = useState("");
  const [idError, setIdError] = useState("");

  const openAdd = () => { setForm(blank); setEditingId(null); setActiveId(uid()); setIdError(""); setOpen(true); };
  const openEdit = (row) => { setForm({ ...blank, ...row }); setEditingId(row.id); setActiveId(row.id); setIdError(""); setOpen(true); };

  const save = () => {
    if (!form.name) return;
    const trimmedName = form.name.trim();
    const clash = orders.some(
      (o) => o.id !== editingId && (o.name || "").trim().toLowerCase() === trimmedName.toLowerCase()
    );
    if (clash) {
      setIdError(`Order/Project ID "${trimmedName}" is already assigned to another project — use a different one.`);
      return;
    }
    setIdError("");
    if (editingId) {
      setOrders(orders.map((o) => (o.id === editingId ? { ...o, ...form, id: editingId } : o)));
    } else {
      setOrders([...orders, { id: activeId, ...form }]);
    }
    setForm(blank);
    setEditingId(null);
    setOpen(false);
  };

  const statusTone = { "Not started": "gray", "In progress": "blue", Delivered: "green", Delayed: "red", Cancelled: "gray" };
  const STATUS_ORDER = { "Not started": 0, "In progress": 1, Delayed: 2, Delivered: 3, Cancelled: 4 };

  const q = query.trim().toLowerCase();
  const filteredOrders = q
    ? orders.filter(
        (o) =>
          (o.name || "").toLowerCase().includes(q) ||
          (o.client || "").toLowerCase().includes(q)
      )
    : orders;
  const sortedOrders = [...filteredOrders].sort((a, b) => {
    const so = (STATUS_ORDER[a.status] ?? 5) - (STATUS_ORDER[b.status] ?? 5);
    if (so !== 0) return so;
    return (a.deadline || "").localeCompare(b.deadline || "");
  });

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center flex-wrap gap-3">
        <h2 className="font-bold text-lg text-neutral-900">Monthly Orders / Projects</h2>
        <div className="flex items-center gap-2 flex-wrap">
          <div className="relative">
            <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              className={`${inputCls} pl-8 w-56`}
              placeholder="Search project or client…"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <button
            onClick={() => {
              setPrintTitle("Orders-Projects-Register");
              setPrintContent(<OrdersPrintLayout orders={sortedOrders} company={company} />);
            }}
            className="inline-flex items-center gap-1.5 bg-white border border-neutral-300 hover:border-red-400 text-sm font-semibold px-4 py-2 rounded-lg transition"
          >
            <Printer size={16} /> Download PDF
          </button>
          <AddButton onClick={openAdd} text="Add order" />
        </div>
      </div>
      <Table
        emptyMsg={q ? `No orders/projects match "${query}".` : "No orders or projects tracked yet."}
        columns={[
          { key: "name", label: "Order / Project ID" },
          { key: "client", label: "Client" },
          { key: "status", label: "Status", render: (r) => <Pill tone={statusTone[r.status] || "gray"}>{r.status}</Pill> },
          { key: "startDate", label: "Start" },
          { key: "deadline", label: "Deadline" },
          { key: "value", label: "Value", render: (r) => fmt(r.value) },
        ]}
        rows={sortedOrders}
        onDelete={(id) => setOrders(orders.filter((o) => o.id !== id))}
        onEdit={openEdit}
      />
      {open && (
        <Modal title={editingId ? "Edit order / project" : "Add order / project"} onClose={() => setOpen(false)}>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Order / Project ID (must be unique)">
              <input
                className={inputCls}
                value={form.name}
                onChange={(e) => { setForm({ ...form, name: e.target.value }); setIdError(""); }}
                placeholder="e.g. NOVA-ORD-001"
              />
              {idError && <span className="text-xs text-red-600">{idError}</span>}
            </Field>
            <PartyPicker
              key={editingId || activeId}
              label="Client"
              value={form.client}
              onChange={(v) => setForm({ ...form, client: v })}
              records={clientRecords}
            />
            <Field label="Status">
              <select className={inputCls} value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>
                <option>Not started</option>
                <option>In progress</option>
                <option>Delayed</option>
                <option>Delivered</option>
                <option>Cancelled</option>
              </select>
            </Field>
            <Field label="Order value (₹)">
              <input type="number" className={inputCls} value={form.value} onChange={(e) => setForm({ ...form, value: e.target.value })} />
            </Field>
            <Field label="Start date">
              <input type="date" className={inputCls} value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} />
            </Field>
            <Field label="Deadline">
              <input type="date" className={inputCls} value={form.deadline} onChange={(e) => setForm({ ...form, deadline: e.target.value })} />
            </Field>
          </div>
          {activeId && <AttachmentsField recordId={activeId} label="Purchase order / related files" />}
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update order" : "Save order"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- VENDOR & CLIENT PAYMENTS ----------
// Indian financial year: April 1 – March 31
function currentFYRange() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth(); // 0 = Jan, 3 = Apr
  const startYear = m >= 3 ? y : y - 1;
  return {
    start: `${startYear}-04-01`,
    end: `${startYear + 1}-03-31`,
    label: `FY ${startYear}-${String(startYear + 1).slice(-2)} (Apr ${startYear} - Mar ${startYear + 1})`,
  };
}

// Reads a payment's installments, with a safe fallback for legacy records saved
// before installment tracking existed: a manually-marked "Paid" row is treated as
// one full installment already received, so old data keeps working with no migration.
function paymentInstallments(p) {
  if (Array.isArray(p.installments) && p.installments.length > 0) return p.installments;
  if (p.status === "Paid") {
    return [{ id: "legacy", amount: p.amount, status: "Received" }];
  }
  return [];
}
function paymentReceived(p) {
  return paymentInstallments(p)
    .filter((i) => i.status === "Received")
    .reduce((s, i) => s + Number(i.amount || 0), 0);
}
// Live status for a client invoice, derived from its installments + closing date —
// computed on every render (not stored) so it never goes stale, e.g. it flips to
// "Overdue" automatically once the closing date passes, even with no edits.
function clientPaymentStatus(p) {
  const total = Number(p.amount || 0);
  const received = paymentReceived(p);
  const isPastDue = p.dueDate && p.dueDate < todayISO();
  if (total > 0 && received >= total) return "Paid";
  if (received > 0) return isPastDue ? "Overdue" : "Partially Paid";
  return isPastDue ? "Overdue" : "Pending";
}

function PaymentsTab({ payments, setPayments, partyType, partyRegistrations, orders }) {
  const defaultType = partyType === "Vendor" ? "Payable" : "Receivable";
  const blank = { party: "", partyType, type: defaultType, amount: "", date: todayISO(), status: "Pending", reference: "", orderRef: "", dueDate: "", installments: [] };
  const partyRecords = partyRegistrations.filter((r) => r.partyType === partyType);
  const isClient = partyType === "Client";
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [form, setForm] = useState(blank);
  const [query, setQuery] = useState("");

  // Status shown/aggregated for a row: live-computed from installments for Client
  // Payments, or the plain manual field for Vendor Payments (unchanged behaviour).
  const displayStatus = (p) => (isClient ? clientPaymentStatus(p) : p.status);

  const addInstallment = () => {
    setForm((f) => ({
      ...f,
      installments: [...(f.installments || []), { id: uid(), percent: "", amount: "", date: todayISO(), status: "Pending", reference: "" }],
    }));
  };
  const removeInstallment = (idx) => {
    setForm((f) => ({ ...f, installments: (f.installments || []).filter((_, i) => i !== idx) }));
  };
  const updateInstallment = (idx, field, value) => {
    setForm((f) => {
      const installments = [...(f.installments || [])];
      const inst = { ...installments[idx], [field]: value };
      const total = Number(f.amount || 0);
      // Keep % and ₹ in sync with each other and with the invoice total, whichever
      // side the user is actually typing into.
      if (field === "percent" && total) {
        inst.amount = String(Math.round((Number(value || 0) / 100) * total));
      } else if (field === "amount" && total) {
        inst.percent = ((Number(value || 0) / total) * 100).toFixed(1);
      }
      installments[idx] = inst;
      return { ...f, installments };
    });
  };

  const formInstallments = form.installments || [];
  const formReceived = formInstallments.filter((i) => i.status === "Received").reduce((s, i) => s + Number(i.amount || 0), 0);
  const formAllocated = formInstallments.reduce((s, i) => s + Number(i.amount || 0), 0);
  const formBalance = Number(form.amount || 0) - formReceived;
  const formProgressPct = form.amount ? Math.min(100, Math.round((formReceived / Number(form.amount)) * 100)) : 0;
  const formOverAllocated = Number(form.amount || 0) > 0 && formAllocated > Number(form.amount) + 0.5;
  const formLiveStatus = clientPaymentStatus({ ...form, installments: formInstallments });

  // Projects belonging to the party currently selected in the form (falls back
  // to all projects if no party chosen yet, so the dropdown isn't empty pre-selection).
  const relevantOrders = orders.filter(
    (o) => !form.party || (o.client || "").trim().toLowerCase() === form.party.trim().toLowerCase()
  );

  const filtered = payments.filter((p) => p.partyType === partyType);
  const q = query.trim().toLowerCase();
  const searchedPayments = q ? filtered.filter((p) => (p.party || "").toLowerCase().includes(q)) : filtered;

  const openAdd = () => { setForm(blank); setEditingId(null); setActiveId(uid()); setOpen(true); };
  const openEdit = (row) => { setForm({ ...blank, ...row }); setEditingId(row.id); setActiveId(row.id); setOpen(true); };

  const save = () => {
    if (!form.party || !form.amount) return;
    if (editingId) {
      setPayments(payments.map((p) => (p.id === editingId ? { ...p, ...form, partyType, id: editingId } : p)));
    } else {
      setPayments([...payments, { id: activeId, ...form, partyType }]);
    }
    setForm(blank);
    setEditingId(null);
    setOpen(false);
  };

  const statusTone = { Paid: "green", Pending: "amber", Overdue: "red", "Partially Paid": "blue" };
  const label = partyType === "Vendor" ? "Vendor Payments" : "Client Payments";

  const fy = currentFYRange();
  const fyPayments = filtered.filter((p) => p.date >= fy.start && p.date <= fy.end);
  const fyTotal = fyPayments.reduce((s, p) => s + Number(p.amount || 0), 0);
  const fyOverdue = isClient
    ? fyPayments.filter((p) => displayStatus(p) === "Overdue").reduce((s, p) => s + (Number(p.amount || 0) - paymentReceived(p)), 0)
    : fyPayments.filter((p) => p.status === "Overdue").reduce((s, p) => s + Number(p.amount), 0);
  const fyPending = isClient
    ? fyPayments.filter((p) => displayStatus(p) === "Pending").reduce((s, p) => s + Number(p.amount || 0), 0)
    : fyPayments.filter((p) => p.status === "Pending").reduce((s, p) => s + Number(p.amount), 0);
  const fyPartial = isClient
    ? fyPayments.filter((p) => displayStatus(p) === "Partially Paid").reduce((s, p) => s + (Number(p.amount || 0) - paymentReceived(p)), 0)
    : 0;
  const fyPaid = isClient
    ? fyPayments.filter((p) => displayStatus(p) === "Paid").reduce((s, p) => s + paymentReceived(p), 0)
    : fyPayments.filter((p) => p.status === "Paid").reduce((s, p) => s + Number(p.amount), 0);

  const STATUS_ORDER = { Overdue: 0, "Partially Paid": 1, Pending: 2, Paid: 3 };
  const sortedPayments = [...searchedPayments].sort((a, b) => {
    const so = (STATUS_ORDER[displayStatus(a)] ?? 4) - (STATUS_ORDER[displayStatus(b)] ?? 4);
    if (so !== 0) return so;
    return a.date < b.date ? 1 : -1;
  });

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center flex-wrap gap-3">
        <h2 className="font-bold text-lg text-neutral-900">{label}</h2>
        <div className="flex gap-2 items-center flex-wrap">
          <div className="relative">
            <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              className={`${inputCls} pl-8 w-48`}
              placeholder={`Search ${partyType.toLowerCase()} name…`}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <AddButton onClick={openAdd} text="Add payment" />
        </div>
      </div>
      <div className={`grid grid-cols-2 gap-4 ${isClient ? "md:grid-cols-5" : "md:grid-cols-4"}`}>
        <Card label={fy.label} value={fmt(fyTotal)} sub={`Total ${partyType.toLowerCase()} this FY`} />
        <Card label="Overdue" value={fmt(fyOverdue)} tone={fyOverdue ? "danger" : "good"} sub={`${fyPayments.filter((p) => displayStatus(p) === "Overdue").length} record(s)`} />
        {isClient && (
          <Card label="Partially Paid" value={fmt(fyPartial)} tone={fyPartial ? "default" : "good"} sub={`${fyPayments.filter((p) => displayStatus(p) === "Partially Paid").length} record(s) · balance due`} />
        )}
        <Card label="Pending" value={fmt(fyPending)} sub={`${fyPayments.filter((p) => displayStatus(p) === "Pending").length} record(s)`} />
        <Card label="Paid" value={fmt(fyPaid)} tone="good" sub={`${fyPayments.filter((p) => displayStatus(p) === "Paid").length} record(s)${isClient ? " · amount received" : ""}`} />
      </div>
      <Table
        emptyMsg={q ? `No ${partyType.toLowerCase()} records match "${query}".` : `No ${partyType.toLowerCase()} payment records yet.`}
        columns={[
          { key: "party", label: "Party" },
          ...(isClient ? [{ key: "orderRef", label: "Project", render: (r) => r.orderRef || "—" }] : []),
          { key: "type", label: "Payable / Receivable" },
          { key: "amount", label: isClient ? "Invoice Amount" : "Amount", render: (r) => fmt(r.amount) },
          ...(isClient
            ? [
                { key: "received", label: "Received", render: (r) => fmt(paymentReceived(r)) },
                { key: "balance", label: "Balance", render: (r) => fmt(Number(r.amount || 0) - paymentReceived(r)) },
              ]
            : []),
          { key: "date", label: "Date" },
          ...(isClient ? [{ key: "dueDate", label: "Invoice due", render: (r) => r.dueDate || "—" }] : []),
          { key: "status", label: "Status", render: (r) => <Pill tone={statusTone[displayStatus(r)]}>{displayStatus(r)}</Pill> },
          { key: "reference", label: "Reference" },
        ]}
        rows={sortedPayments}
        onDelete={(id) => setPayments(payments.filter((p) => p.id !== id))}
        onEdit={openEdit}
      />
      {open && (
        <Modal title={editingId ? "Edit payment record" : "Add payment record"} onClose={() => setOpen(false)}>
          <div className="grid grid-cols-2 gap-3">
            <PartyPicker
              key={editingId || activeId}
              label={`${partyType} name`}
              value={form.party}
              onChange={(v) => setForm({ ...form, party: v })}
              records={partyRecords}
            />
            <Field label="Payable / Receivable">
              <select className={inputCls} value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                <option>Payable</option>
                <option>Receivable</option>
              </select>
            </Field>
            {isClient && (
              <Field label="Project / Order">
                <select
                  className={inputCls}
                  value={form.orderRef}
                  onChange={(e) => {
                    const val = e.target.value;
                    const matched = orders.find((o) => o.name === val);
                    setForm((f) => ({
                      ...f,
                      orderRef: val,
                      amount: matched ? matched.value : f.amount,
                    }));
                  }}
                >
                  <option value="">— none / manual amount —</option>
                  {relevantOrders.map((o) => (
                    <option key={o.id} value={o.name}>{o.name} · {fmt(o.value)}</option>
                  ))}
                </select>
                {relevantOrders.length === 0 && (
                  <span className="text-[11px] text-neutral-400">
                    No projects found for this client in Orders/Projects yet.
                  </span>
                )}
              </Field>
            )}
            <Field label={isClient ? "Invoice Amount (₹)" : "Amount (₹)"}>
              <input type="number" className={inputCls} value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} />
              {form.orderRef && (
                <span className="text-[11px] text-neutral-400">Auto-filled from {form.orderRef} — edit if the actual invoice differs.</span>
              )}
            </Field>
            <Field label="Date">
              <input type="date" className={inputCls} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
            </Field>
            {isClient && (
              <Field label="Invoice closing date">
                <input
                  type="date"
                  className={inputCls}
                  value={form.dueDate}
                  onChange={(e) => setForm({ ...form, dueDate: e.target.value })}
                />
              </Field>
            )}
            {!isClient && (
              <Field label="Status">
                <select className={inputCls} value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>
                  <option>Pending</option>
                  <option>Paid</option>
                  <option>Overdue</option>
                </select>
              </Field>
            )}
            <div className="col-span-2">
              <Field label="Reference / invoice no.">
                <input className={inputCls} value={form.reference} onChange={(e) => setForm({ ...form, reference: e.target.value })} />
              </Field>
            </div>
          </div>

          {isClient && (
            <div className="border-t border-neutral-100 pt-3 mt-3">
              <div className="flex items-center justify-between mb-1">
                <div className="text-sm font-semibold text-neutral-700">Payment installments / tranches</div>
                <button onClick={addInstallment} className="text-xs font-semibold text-red-600 hover:text-red-700 inline-flex items-center gap-1">
                  <Plus size={13} /> Add installment
                </button>
              </div>
              <p className="text-xs text-neutral-400 mb-2">
                Split this invoice however that client's terms work — e.g. 20% advance + 50% mid + 30% on delivery. Enter each tranche as a % or a fixed ₹ amount; the other side fills in automatically.
              </p>
              <div className="space-y-2">
                {formInstallments.map((inst, idx) => (
                  <div key={inst.id} className="grid grid-cols-12 gap-1.5 items-center bg-neutral-50 border border-neutral-200 rounded-lg p-2">
                    <input
                      className={`${inputCls} col-span-2 px-2`}
                      type="number"
                      placeholder="%"
                      value={inst.percent}
                      onChange={(e) => updateInstallment(idx, "percent", e.target.value)}
                    />
                    <input
                      className={`${inputCls} col-span-3 px-2`}
                      type="number"
                      placeholder="₹ amount"
                      value={inst.amount}
                      onChange={(e) => updateInstallment(idx, "amount", e.target.value)}
                    />
                    <input
                      className={`${inputCls} col-span-2 px-2`}
                      type="date"
                      value={inst.date}
                      onChange={(e) => updateInstallment(idx, "date", e.target.value)}
                    />
                    <select
                      className={`${inputCls} col-span-2 px-2`}
                      value={inst.status}
                      onChange={(e) => updateInstallment(idx, "status", e.target.value)}
                    >
                      <option>Pending</option>
                      <option>Received</option>
                    </select>
                    <input
                      className={`${inputCls} col-span-2 px-2`}
                      placeholder="Ref / UTR"
                      value={inst.reference}
                      onChange={(e) => updateInstallment(idx, "reference", e.target.value)}
                    />
                    <button onClick={() => removeInstallment(idx)} className="col-span-1 text-neutral-400 hover:text-red-600 flex justify-center">
                      <Trash2 size={14} />
                    </button>
                  </div>
                ))}
                {formInstallments.length === 0 && (
                  <div className="text-center text-[11px] text-neutral-400 py-3 border border-dashed border-neutral-200 rounded-lg">
                    No installments added yet — this invoice will show as Pending until you add one.
                  </div>
                )}
              </div>

              <div className="bg-white border border-neutral-200 rounded-lg p-3 mt-3">
                <div className="flex justify-between text-sm">
                  <span className="text-neutral-500">Received: <strong className="text-emerald-600">{fmt(formReceived)}</strong></span>
                  <span className="text-neutral-500">Balance: <strong className={formBalance > 0 ? "text-red-600" : "text-emerald-600"}>{fmt(formBalance)}</strong></span>
                </div>
                <div className="w-full bg-neutral-200 rounded-full h-2 mt-2 overflow-hidden">
                  <div className="bg-emerald-500 h-2" style={{ width: `${formProgressPct}%` }} />
                </div>
                <div className="text-xs text-neutral-500 mt-1.5 flex items-center justify-between flex-wrap gap-1">
                  <span>Status: <strong className="text-neutral-800">{formLiveStatus}</strong></span>
                  {formOverAllocated && (
                    <span className="text-red-600 font-semibold">⚠ Installments add up to more than the invoice amount</span>
                  )}
                </div>
              </div>
            </div>
          )}

          {activeId && <AttachmentsField recordId={activeId} label="Invoice / proof of payment" />}
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update payment" : "Save payment"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- DELIVERY DOCUMENTS (Challan + E-Way Bill reference) ----------
const REASONS = ["Supply", "Job Work", "Sales Return", "Exhibition / Fair", "Own Use", "Others"];
const TRANSPORT_MODES = ["Road", "Rail", "Air", "Ship"];

const blankItemRow = () => ({ id: uid(), description: "", hsn: "", qty: "", unit: "pcs", rate: "" });

const blankChallan = () => ({
  challanNo: "",
  date: todayISO(),
  consigneeName: "",
  consigneeAddress: "",
  consigneeGstin: "",
  orderRef: "",
  reason: "Supply",
  transportMode: "Road",
  vehicleNo: "",
  transporterName: "",
  distanceKm: "",
  ewayBillNo: "",
  ewayBillDate: "",
  notes: "",
  items: [blankItemRow()],
});

function DeliveryDocsTab({ docs, setDocs, company, setCompany, orders, partyRegistrations, setPrintContent, setPrintTitle }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(blankChallan());
  const clientRecords = partyRegistrations.filter((r) => r.partyType === "Client");

  const [activeId, setActiveId] = useState(null);
  const openAdd = () => {
    setForm({ ...blankChallan(), challanNo: `DC-${Date.now().toString().slice(-6)}` });
    setEditingId(null);
    setActiveId(uid());
    setOpen(true);
  };
  const openEdit = (row) => {
    setForm({ ...blankChallan(), ...row, items: row.items?.length ? row.items : [blankItemRow()] });
    setEditingId(row.id);
    setActiveId(row.id);
    setOpen(true);
  };

  const updateItem = (idx, field, value) => {
    const items = [...form.items];
    items[idx] = { ...items[idx], [field]: value };
    setForm({ ...form, items });
  };
  const addItemRow = () => setForm({ ...form, items: [...form.items, blankItemRow()] });
  const removeItemRow = (idx) => setForm({ ...form, items: form.items.filter((_, i) => i !== idx) });

  const save = () => {
    if (!form.consigneeName || !form.items.some((i) => i.description)) return;
    if (editingId) {
      setDocs(docs.map((d) => (d.id === editingId ? { ...d, ...form, id: editingId } : d)));
    } else {
      setDocs([...docs, { id: uid(), ...form }]);
    }
    setForm(blankChallan());
    setEditingId(null);
    setOpen(false);
  };

  const total = (doc) =>
    (doc.items || []).reduce((s, i) => s + Number(i.qty || 0) * Number(i.rate || 0), 0);

  return (
    <div className="space-y-6">
      <div>
        <h2 className="font-bold text-lg text-neutral-900">Delivery Documents</h2>
        <p className="text-xs text-neutral-400 mt-0.5">
          Generate a delivery challan for goods leaving for a client site — print it or save as PDF.
        </p>
      </div>

      <div className="bg-white border border-neutral-200 rounded-xl p-4 space-y-3">
        <div className="text-sm font-semibold text-neutral-800">Company details (shown on printed documents)</div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Company name">
            <input className={inputCls} value={company.name} onChange={(e) => setCompany({ ...company, name: e.target.value })} />
          </Field>
          <Field label="GSTIN">
            <input className={inputCls} value={company.gstin} onChange={(e) => setCompany({ ...company, gstin: e.target.value })} />
          </Field>
          <div className="col-span-2">
            <Field label="Address">
              <input className={inputCls} value={company.address} onChange={(e) => setCompany({ ...company, address: e.target.value })} />
            </Field>
          </div>
        </div>
      </div>

      <div className="flex justify-between items-center">
        <h3 className="font-semibold text-neutral-800 text-sm">Challans</h3>
        <AddButton onClick={openAdd} text="New challan" />
      </div>

      <Table
        emptyMsg="No delivery challans yet."
        columns={[
          { key: "challanNo", label: "Challan no." },
          { key: "date", label: "Date" },
          { key: "consigneeName", label: "Client" },
          { key: "reason", label: "Reason" },
          { key: "value", label: "Value", render: (r) => fmt(total(r)) },
          {
            key: "print",
            label: "",
            render: (r) => (
              <button
                onClick={() => { setPrintTitle(`Delivery-Challan-${r.challanNo || r.id}`); setPrintContent(<ChallanPrintLayout doc={r} company={company} />); }}
                className="inline-flex items-center gap-1 text-xs font-semibold text-red-600 hover:text-red-700"
              >
                <Printer size={13} /> Print / PDF
              </button>
            ),
          },
        ]}
        rows={[...docs].sort((a, b) => (a.date < b.date ? 1 : -1))}
        onDelete={(id) => setDocs(docs.filter((d) => d.id !== id))}
        onEdit={openEdit}
      />

      {open && (
        <Modal title={editingId ? "Edit delivery challan" : "New delivery challan"} onClose={() => setOpen(false)}>
          <div className="space-y-4">
            <div className="grid grid-cols-2 gap-3">
              <Field label="Challan no.">
                <input className={inputCls} value={form.challanNo} onChange={(e) => setForm({ ...form, challanNo: e.target.value })} />
              </Field>
              <Field label="Date">
                <input type="date" className={inputCls} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
              </Field>
              <PartyPicker
                key={editingId || activeId}
                label="Client / consignee name"
                value={form.consigneeName}
                onChange={(v) => setForm({ ...form, consigneeName: v })}
                records={clientRecords}
                onPick={(rec) => {
                  if (rec) {
                    setForm((f) => ({
                      ...f,
                      consigneeName: rec.name,
                      consigneeGstin: rec.gstin || f.consigneeGstin,
                      consigneeAddress: rec.address || f.consigneeAddress,
                    }));
                  }
                }}
              />
              <Field label="Client GSTIN (if registered)">
                <input className={inputCls} value={form.consigneeGstin} onChange={(e) => setForm({ ...form, consigneeGstin: e.target.value })} />
              </Field>
              <div className="col-span-2">
                <Field label="Delivery address">
                  <input className={inputCls} value={form.consigneeAddress} onChange={(e) => setForm({ ...form, consigneeAddress: e.target.value })} />
                </Field>
              </div>
              <Field label="Link to order/project (optional)">
                <select className={inputCls} value={form.orderRef} onChange={(e) => setForm({ ...form, orderRef: e.target.value })}>
                  <option value="">— none —</option>
                  {orders.map((o) => (
                    <option key={o.id} value={o.name}>{o.name}</option>
                  ))}
                </select>
              </Field>
              <Field label="Reason for delivery">
                <select className={inputCls} value={form.reason} onChange={(e) => setForm({ ...form, reason: e.target.value })}>
                  {REASONS.map((r) => <option key={r}>{r}</option>)}
                </select>
              </Field>
            </div>

            <div className="border-t border-neutral-100 pt-3">
              <div className="text-sm font-semibold text-neutral-700 mb-2">Transport details</div>
              <div className="grid grid-cols-2 gap-3">
                <Field label="Mode">
                  <select className={inputCls} value={form.transportMode} onChange={(e) => setForm({ ...form, transportMode: e.target.value })}>
                    {TRANSPORT_MODES.map((m) => <option key={m}>{m}</option>)}
                  </select>
                </Field>
                <Field label="Vehicle no.">
                  <input className={inputCls} value={form.vehicleNo} onChange={(e) => setForm({ ...form, vehicleNo: e.target.value })} />
                </Field>
                <Field label="Transporter name">
                  <input className={inputCls} value={form.transporterName} onChange={(e) => setForm({ ...form, transporterName: e.target.value })} />
                </Field>
                <Field label="Approx. distance (km)">
                  <input type="number" className={inputCls} value={form.distanceKm} onChange={(e) => setForm({ ...form, distanceKm: e.target.value })} />
                </Field>
              </div>
            </div>

            <div className="border-t border-neutral-100 pt-3">
              <div className="text-sm font-semibold text-neutral-700 mb-1">E-Way bill reference</div>
              <p className="text-xs text-neutral-400 mb-2">
                Generate the actual e-way bill on the GST portal first, then paste its number here so it prints on the challan.
              </p>
              <div className="grid grid-cols-2 gap-3">
                <Field label="E-Way bill no.">
                  <input className={inputCls} value={form.ewayBillNo} onChange={(e) => setForm({ ...form, ewayBillNo: e.target.value })} placeholder="e.g. 1234 5678 9012" />
                </Field>
                <Field label="Generated on">
                  <input type="date" className={inputCls} value={form.ewayBillDate} onChange={(e) => setForm({ ...form, ewayBillDate: e.target.value })} />
                </Field>
              </div>
            </div>

            <div className="border-t border-neutral-100 pt-3">
              <div className="flex justify-between items-center mb-2">
                <div className="text-sm font-semibold text-neutral-700">Items</div>
                <button onClick={addItemRow} className="text-xs font-semibold text-red-600 hover:text-red-700 inline-flex items-center gap-1">
                  <Plus size={13} /> Add item
                </button>
              </div>
              <div className="space-y-2">
                {form.items.map((it, idx) => (
                  <div key={it.id} className="grid grid-cols-12 gap-2 items-center">
                    <input
                      className={`${inputCls} col-span-4`}
                      placeholder="Description"
                      value={it.description}
                      onChange={(e) => updateItem(idx, "description", e.target.value)}
                    />
                    <input
                      className={`${inputCls} col-span-2`}
                      placeholder="HSN"
                      value={it.hsn}
                      onChange={(e) => updateItem(idx, "hsn", e.target.value)}
                    />
                    <input
                      type="number"
                      className={`${inputCls} col-span-2`}
                      placeholder="Qty"
                      value={it.qty}
                      onChange={(e) => updateItem(idx, "qty", e.target.value)}
                    />
                    <input
                      className={`${inputCls} col-span-1`}
                      placeholder="Unit"
                      value={it.unit}
                      onChange={(e) => updateItem(idx, "unit", e.target.value)}
                    />
                    <input
                      type="number"
                      className={`${inputCls} col-span-2`}
                      placeholder="Rate ₹"
                      value={it.rate}
                      onChange={(e) => updateItem(idx, "rate", e.target.value)}
                    />
                    <button onClick={() => removeItemRow(idx)} className="col-span-1 text-neutral-300 hover:text-red-600">
                      <Trash2 size={15} />
                    </button>
                  </div>
                ))}
              </div>
              <div className="text-right text-sm font-semibold text-neutral-800 mt-2">
                Total: {fmt(form.items.reduce((s, i) => s + Number(i.qty || 0) * Number(i.rate || 0), 0))}
              </div>
            </div>

            <Field label="Notes (optional)">
              <input className={inputCls} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} />
            </Field>
          </div>
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update challan" : "Save challan"} />
          </div>
        </Modal>
      )}

    </div>
  );
}

function ChallanPrintLayout({ doc, company }) {
  const total = (doc.items || []).reduce((s, i) => s + Number(i.qty || 0) * Number(i.rate || 0), 0);
  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "800px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
        {company.gstin && <div style={{ fontSize: "12px", color: "#555" }}>GSTIN: {company.gstin}</div>}
      </div>

      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 16px" }}>
        DELIVERY CHALLAN
      </div>

      <table style={{ width: "100%", fontSize: "12px", marginBottom: "14px" }}>
        <tbody>
          <tr>
            <td style={{ verticalAlign: "top", width: "50%" }}>
              <strong>Challan No:</strong> {doc.challanNo}<br />
              <strong>Date:</strong> {doc.date}<br />
              <strong>Reason:</strong> {doc.reason}
              {doc.orderRef && <><br /><strong>Order/Project:</strong> {doc.orderRef}</>}
            </td>
            <td style={{ verticalAlign: "top", width: "50%" }}>
              <strong>Deliver to:</strong><br />
              {doc.consigneeName}<br />
              {doc.consigneeAddress}<br />
              {doc.consigneeGstin && <>GSTIN: {doc.consigneeGstin}</>}
            </td>
          </tr>
        </tbody>
      </table>

      <table style={{ width: "100%", fontSize: "12px", marginBottom: "14px", border: "1px solid #ddd", borderCollapse: "collapse" }}>
        <tbody>
          <tr>
            <td style={{ border: "1px solid #ddd", padding: "6px" }}>
              <strong>Transport mode:</strong> {doc.transportMode}<br />
              <strong>Vehicle no.:</strong> {doc.vehicleNo || "—"}<br />
              <strong>Transporter:</strong> {doc.transporterName || "—"}<br />
              <strong>Distance:</strong> {doc.distanceKm ? `${doc.distanceKm} km` : "—"}
            </td>
            <td style={{ border: "1px solid #ddd", padding: "6px" }}>
              <strong>E-Way Bill No.:</strong> {doc.ewayBillNo || "Not applicable / generated separately"}<br />
              {doc.ewayBillDate && <><strong>Generated on:</strong> {doc.ewayBillDate}</>}
            </td>
          </tr>
        </tbody>
      </table>

      <table style={{ width: "100%", fontSize: "12px", borderCollapse: "collapse", marginBottom: "14px" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={{ border: "1px solid #ddd", padding: "6px", textAlign: "left" }}>#</th>
            <th style={{ border: "1px solid #ddd", padding: "6px", textAlign: "left" }}>Description</th>
            <th style={{ border: "1px solid #ddd", padding: "6px", textAlign: "left" }}>HSN</th>
            <th style={{ border: "1px solid #ddd", padding: "6px", textAlign: "right" }}>Qty</th>
            <th style={{ border: "1px solid #ddd", padding: "6px", textAlign: "right" }}>Rate</th>
            <th style={{ border: "1px solid #ddd", padding: "6px", textAlign: "right" }}>Amount</th>
          </tr>
        </thead>
        <tbody>
          {(doc.items || []).filter((i) => i.description).map((it, idx) => (
            <tr key={it.id || idx}>
              <td style={{ border: "1px solid #ddd", padding: "6px" }}>{idx + 1}</td>
              <td style={{ border: "1px solid #ddd", padding: "6px" }}>{it.description}</td>
              <td style={{ border: "1px solid #ddd", padding: "6px" }}>{it.hsn}</td>
              <td style={{ border: "1px solid #ddd", padding: "6px", textAlign: "right" }}>{it.qty} {it.unit}</td>
              <td style={{ border: "1px solid #ddd", padding: "6px", textAlign: "right" }}>{fmt(it.rate)}</td>
              <td style={{ border: "1px solid #ddd", padding: "6px", textAlign: "right" }}>{fmt(Number(it.qty || 0) * Number(it.rate || 0))}</td>
            </tr>
          ))}
          <tr>
            <td colSpan="5" style={{ border: "1px solid #ddd", padding: "6px", textAlign: "right", fontWeight: "bold" }}>Total</td>
            <td style={{ border: "1px solid #ddd", padding: "6px", textAlign: "right", fontWeight: "bold" }}>{fmt(total)}</td>
          </tr>
        </tbody>
      </table>

      {doc.notes && <div style={{ fontSize: "12px", marginBottom: "14px" }}><strong>Notes:</strong> {doc.notes}</div>}

      <table style={{ width: "100%", fontSize: "12px", marginTop: "40px" }}>
        <tbody>
          <tr>
            <td style={{ width: "50%" }}>____________________<br />Dispatched by</td>
            <td style={{ width: "50%" }}>____________________<br />Received by (client signature & stamp)</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ---------- PAYROLL ----------
const EARNINGS_FIELDS = [
  { key: "basic", label: "Basic Salary" },
  { key: "hra", label: "HRA" },
  { key: "da", label: "DA" },
  { key: "conveyance", label: "Conveyance" },
  { key: "medical", label: "Medical" },
  { key: "specialAllowance", label: "Special Allowance" },
  { key: "travelAllowance", label: "Travel Allowance" },
  { key: "foodAllowance", label: "Food Allowance" },
  { key: "bonus", label: "Bonus" },
  { key: "incentive", label: "Incentive" },
  { key: "overtime", label: "Overtime" },
  { key: "arrears", label: "Arrears" },
  { key: "otherEarnings", label: "Other Earnings" },
];

const DEDUCTION_FIELDS = [
  { key: "pf", label: "PF" },
  { key: "esi", label: "ESI" },
  { key: "professionalTax", label: "Professional Tax" },
  { key: "tds", label: "TDS" },
  { key: "loanDeduction", label: "Loan Deduction" },
  { key: "advanceRecovery", label: "Advance Recovery" },
  { key: "leaveDeduction", label: "Leave Deduction" },
  { key: "insurance", label: "Insurance" },
  { key: "otherDeduction", label: "Other Deduction" },
];

const blankSlip = () => {
  const e = {}, d = {};
  EARNINGS_FIELDS.forEach((f) => (e[f.key] = ""));
  DEDUCTION_FIELDS.forEach((f) => (d[f.key] = ""));
  return {
    employeeRowId: "",
    month: thisMonthKey(),
    earnings: e,
    deductions: d,
    paymentMode: "Bank Transfer",
    transactionId: "",
    salaryStatus: "Pending",
    paymentDate: "",
  };
};

function attendanceSummary(attendance, employeeName, month) {
  const year = month.slice(0, 4);
  const rows = attendance.filter((a) => a.employeeName === employeeName && monthKey(a.date) === month);
  const yearRows = attendance.filter((a) => a.employeeName === employeeName && a.date.slice(0, 4) === year);
  return {
    present: rows.filter((r) => r.status === "Present").length,
    halfDay: rows.filter((r) => r.status === "Half-day").length,
    cl: rows.filter((r) => r.status === "CL").length,
    ml: rows.filter((r) => r.status === "ML").length,
    permission: rows.filter((r) => r.status === "Permission").length,
    clYear: yearRows.filter((r) => r.status === "CL").length,
    mlYear: yearRows.filter((r) => r.status === "ML").length,
  };
}

function PayrollTab({ slips, setSlips, employees, attendance, company, setPrintContent, setPrintTitle }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(blankSlip());
  const [overviewMonth, setOverviewMonth] = useState(thisMonthKey());

  const selectedEmployee = employees.find((e) => e.id === form.employeeRowId);
  const att = selectedEmployee ? attendanceSummary(attendance, selectedEmployee.name, form.month) : null;

  const openAdd = () => { setForm(blankSlip()); setEditingId(null); setOpen(true); };
  const openEdit = (row) => { setForm({ ...blankSlip(), ...row }); setEditingId(row.id); setOpen(true); };

  const generateForEmployee = (emp, month) => {
    const base = blankSlip();
    setForm({
      ...base,
      month,
      employeeRowId: emp.id,
      earnings: { ...base.earnings, basic: emp.salary || "" },
      deductions: { ...base.deductions, pf: emp.pfContribution || "", esi: emp.esiContribution || "" },
    });
    setEditingId(null);
    setOpen(true);
  };

  const onEmployeeChange = (id) => {
    const emp = employees.find((e) => e.id === id);
    setForm({
      ...form,
      employeeRowId: id,
      earnings: {
        ...form.earnings,
        basic: emp ? emp.salary : form.earnings.basic,
      },
      deductions: {
        ...form.deductions,
        pf: emp ? emp.pfContribution : form.deductions.pf,
        esi: emp ? emp.esiContribution : form.deductions.esi,
      },
    });
  };

  const suggestLeaveDeduction = () => {
    if (!selectedEmployee || !att) return;
    const perDay = Number(selectedEmployee.salary || 0) / 30;
    const excessDays = Math.max(0, att.clYear - LEAVE_LIMITS.CL) + Math.max(0, att.mlYear - LEAVE_LIMITS.ML);
    const suggested = Math.round(perDay * excessDays);
    setForm({ ...form, deductions: { ...form.deductions, leaveDeduction: suggested } });
  };

  const gross = EARNINGS_FIELDS.reduce((s, f) => s + Number(form.earnings[f.key] || 0), 0);
  const totalDeductions = DEDUCTION_FIELDS.reduce((s, f) => s + Number(form.deductions[f.key] || 0), 0);
  const net = gross - totalDeductions;

  const save = () => {
    if (!form.employeeRowId || !form.month) return;
    if (editingId) {
      setSlips(slips.map((s) => (s.id === editingId ? { ...s, ...form, id: editingId } : s)));
    } else {
      setSlips([...slips, { id: uid(), ...form }]);
    }
    setForm(blankSlip());
    setEditingId(null);
    setOpen(false);
  };

  const nameFor = (rowId) => employees.find((e) => e.id === rowId)?.name || "—";
  const totalOf = (s) => {
    const g = EARNINGS_FIELDS.reduce((sum, f) => sum + Number(s.earnings?.[f.key] || 0), 0);
    const d = DEDUCTION_FIELDS.reduce((sum, f) => sum + Number(s.deductions?.[f.key] || 0), 0);
    return { g, d, net: g - d };
  };
  const statusTone = { Paid: "green", Pending: "amber", Hold: "red" };

  // Merged view: every current employee, for the selected month, combining
  // live Employee master data + live Attendance with any saved slip.
  const overviewRows = employees.map((emp) => {
    const savedSlip = slips.find((s) => s.employeeRowId === emp.id && s.month === overviewMonth);
    const att = attendanceSummary(attendance, emp.name, overviewMonth);
    if (savedSlip) {
      const t = totalOf(savedSlip);
      return { emp, att, slip: savedSlip, gross: t.g, deductions: t.d, net: t.net, status: savedSlip.salaryStatus, generated: true };
    }
    const gross = Number(emp.salary || 0);
    const deductions = Number(emp.pfContribution || 0) + Number(emp.esiContribution || 0);
    return { emp, att, slip: null, gross, deductions, net: gross - deductions, status: "Not generated", generated: false };
  });

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center flex-wrap gap-3">
        <div>
          <h2 className="font-bold text-lg text-neutral-900">Payroll</h2>
          <p className="text-xs text-neutral-400 mt-0.5">
            Every current employee, merged live with Attendance — pick a month to see who's generated and who isn't.
          </p>
        </div>
        <AddButton onClick={openAdd} text="New salary slip" />
      </div>

      <div className="flex items-center gap-2">
        <span className="text-xs font-semibold text-neutral-500">Month:</span>
        <input type="month" className={inputCls} value={overviewMonth} onChange={(e) => setOverviewMonth(e.target.value)} />
      </div>

      <Table
        emptyMsg="Add employees first (Employees & Salary tab) to see them here."
        columns={[
          { key: "empId", label: "Emp ID", render: (r) => r.emp.employeeId || "—" },
          { key: "employee", label: "Employee", render: (r) => r.emp.name },
          { key: "gross", label: "Gross", render: (r) => fmt(r.gross) },
          { key: "deductions", label: "Deductions", render: (r) => fmt(r.deductions) },
          { key: "net", label: "Net payable", render: (r) => fmt(r.net) },
          {
            key: "status",
            label: "Status",
            render: (r) => r.generated ? <Pill tone={statusTone[r.status] || "gray"}>{r.status}</Pill> : <Pill tone="gray">Not generated</Pill>,
          },
          {
            key: "action",
            label: "",
            render: (r) =>
              r.generated ? (
                <button
                  onClick={() => { setPrintTitle(`Salary-Slip-${r.emp.name}-${overviewMonth}`); setPrintContent(<SalarySlipPrintLayout slip={r.slip} employee={r.emp} company={company} />); }}
                  className="inline-flex items-center gap-1 text-xs font-semibold text-red-600 hover:text-red-700"
                >
                  <Printer size={13} /> Print / PDF
                </button>
              ) : (
                <button
                  onClick={() => generateForEmployee(r.emp, overviewMonth)}
                  className="inline-flex items-center gap-1 text-xs font-semibold text-red-600 hover:text-red-700"
                >
                  <Plus size={13} /> Generate slip
                </button>
              ),
          },
        ]}
        rows={overviewRows}
        hideActions
      />

      <h3 className="font-semibold text-sm text-neutral-600 mt-6">All generated slips (every month)</h3>
      <Table
        emptyMsg="No salary slips generated yet."
        columns={[
          { key: "employee", label: "Employee", render: (r) => nameFor(r.employeeRowId) },
          { key: "month", label: "Month" },
          { key: "gross", label: "Gross", render: (r) => fmt(totalOf(r).g) },
          { key: "deductions", label: "Deductions", render: (r) => fmt(totalOf(r).d) },
          { key: "net", label: "Net payable", render: (r) => fmt(totalOf(r).net) },
          { key: "salaryStatus", label: "Status", render: (r) => <Pill tone={statusTone[r.salaryStatus] || "gray"}>{r.salaryStatus}</Pill> },
          {
            key: "print",
            label: "",
            render: (r) => {
              const emp = employees.find((e) => e.id === r.employeeRowId);
              return (
                <button
                  onClick={() => { setPrintTitle(`Salary-Slip-${nameFor(r.employeeRowId)}-${r.month}`); setPrintContent(<SalarySlipPrintLayout slip={r} employee={emp} company={company} />); }}
                  className="inline-flex items-center gap-1 text-xs font-semibold text-red-600 hover:text-red-700"
                >
                  <Printer size={13} /> Print / PDF
                </button>
              );
            },
          },
        ]}
        rows={[...slips].sort((a, b) => (a.month < b.month ? 1 : -1))}
        onDelete={(id) => setSlips(slips.filter((s) => s.id !== id))}
        onEdit={openEdit}
      />

      {open && (
        <Modal title={editingId ? "Edit salary slip" : "New salary slip"} onClose={() => setOpen(false)}>
          <div className="space-y-4">
            <div className="grid grid-cols-2 gap-3">
              <Field label="Employee">
                <select className={inputCls} value={form.employeeRowId} onChange={(e) => onEmployeeChange(e.target.value)}>
                  <option value="">— select —</option>
                  {employees.map((e) => (
                    <option key={e.id} value={e.id}>{e.name}</option>
                  ))}
                </select>
              </Field>
              <Field label="Payroll month">
                <input type="month" className={inputCls} value={form.month} onChange={(e) => setForm({ ...form, month: e.target.value })} />
              </Field>
            </div>

            {selectedEmployee && (
              <div className="bg-neutral-50 border border-neutral-200 rounded-lg p-3 text-xs text-neutral-600 grid grid-cols-2 gap-x-4 gap-y-1">
                <div><strong>Emp ID:</strong> {selectedEmployee.employeeId || "—"}</div>
                <div><strong>Department:</strong> {selectedEmployee.department || "—"}</div>
                <div><strong>Designation:</strong> {selectedEmployee.role || "—"}</div>
                <div><strong>Work status:</strong> {selectedEmployee.workStatus || "—"}</div>
                <div><strong>DOJ:</strong> {selectedEmployee.joiningDate || "—"}</div>
                <div><strong>Bank:</strong> {selectedEmployee.bankName || "—"} {selectedEmployee.accountNumber ? `(${selectedEmployee.accountNumber})` : ""}</div>
                <div><strong>IFSC:</strong> {selectedEmployee.ifsc || "—"}</div>
                <div><strong>UAN (PF):</strong> {selectedEmployee.pfNumber || "—"}</div>
                <div><strong>ESI No.:</strong> {selectedEmployee.esiNumber || "—"}</div>
                <div><strong>PAN:</strong> {selectedEmployee.pan || "—"}</div>
                {att && (
                  <div className="col-span-2 pt-1 border-t border-neutral-200 mt-1">
                    <strong>Attendance this month:</strong> Present {att.present} · Half-day {att.halfDay} · CL (YTD) {att.clYear}/{LEAVE_LIMITS.CL} · ML (YTD) {att.mlYear}/{LEAVE_LIMITS.ML} · Permission (month) {att.permission}/{LEAVE_LIMITS.Permission}{" "}
                    <button onClick={suggestLeaveDeduction} className="ml-2 text-red-600 font-semibold hover:text-red-700">
                      Suggest leave deduction
                    </button>
                  </div>
                )}
              </div>
            )}

            <div className="border-t border-neutral-100 pt-3">
              <div className="text-sm font-semibold text-neutral-700 mb-2">Earnings</div>
              <div className="grid grid-cols-3 gap-3">
                {EARNINGS_FIELDS.map((f) => (
                  <Field key={f.key} label={f.label}>
                    <input
                      type="number"
                      className={inputCls}
                      value={form.earnings[f.key]}
                      onChange={(e) => setForm({ ...form, earnings: { ...form.earnings, [f.key]: e.target.value } })}
                    />
                  </Field>
                ))}
              </div>
            </div>

            <div className="border-t border-neutral-100 pt-3">
              <div className="text-sm font-semibold text-neutral-700 mb-2">Deductions</div>
              <div className="grid grid-cols-3 gap-3">
                {DEDUCTION_FIELDS.map((f) => (
                  <Field key={f.key} label={f.label}>
                    <input
                      type="number"
                      className={inputCls}
                      value={form.deductions[f.key]}
                      onChange={(e) => setForm({ ...form, deductions: { ...form.deductions, [f.key]: e.target.value } })}
                    />
                  </Field>
                ))}
              </div>
            </div>

            <div className="flex justify-end gap-6 text-sm font-semibold border-t border-neutral-100 pt-3">
              <div>Gross: <span className="text-neutral-900">{fmt(gross)}</span></div>
              <div>Deductions: <span className="text-neutral-900">{fmt(totalDeductions)}</span></div>
              <div>Net payable: <span className="text-red-600">{fmt(net)}</span></div>
            </div>

            <div className="border-t border-neutral-100 pt-3">
              <div className="text-sm font-semibold text-neutral-700 mb-2">Payment information</div>
              <div className="grid grid-cols-2 gap-3">
                <Field label="Payment mode">
                  <select className={inputCls} value={form.paymentMode} onChange={(e) => setForm({ ...form, paymentMode: e.target.value })}>
                    <option>Bank Transfer</option>
                    <option>Cash</option>
                    <option>UPI</option>
                    <option>Cheque</option>
                  </select>
                </Field>
                <Field label="Transaction ID / UTR">
                  <input className={inputCls} value={form.transactionId} onChange={(e) => setForm({ ...form, transactionId: e.target.value })} />
                </Field>
                <Field label="Salary status">
                  <select className={inputCls} value={form.salaryStatus} onChange={(e) => setForm({ ...form, salaryStatus: e.target.value })}>
                    <option>Pending</option>
                    <option>Paid</option>
                    <option>Hold</option>
                  </select>
                </Field>
                <Field label="Payment date">
                  <input type="date" className={inputCls} value={form.paymentDate} onChange={(e) => setForm({ ...form, paymentDate: e.target.value })} />
                </Field>
              </div>
            </div>
          </div>
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update slip" : "Save slip"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

function SalarySlipPrintLayout({ slip, employee, company }) {
  const gross = EARNINGS_FIELDS.reduce((s, f) => s + Number(slip.earnings?.[f.key] || 0), 0);
  const totalDeductions = DEDUCTION_FIELDS.reduce((s, f) => s + Number(slip.deductions?.[f.key] || 0), 0);
  const net = gross - totalDeductions;
  const [year, mo] = (slip.month || "").split("-");
  const monthName = mo ? new Date(Number(year), Number(mo) - 1, 1).toLocaleString("en-IN", { month: "long" }) : "";
  const cell = { border: "1px solid #ddd", padding: "6px", fontSize: "12px" };
  const labelCell = { ...cell, background: "#f8f8f8", fontWeight: "bold", width: "18%" };

  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "800px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>

      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "6px 0" }}>
        SALARY SLIP
      </div>
      <div style={{ textAlign: "center", fontSize: "12px", marginBottom: "14px" }}>
        Payroll Month: <strong>{monthName}</strong> &nbsp; Payroll Year: <strong>{year}</strong>
      </div>

      <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "12px" }}>
        <tbody>
          <tr><td style={labelCell}>Employee Name</td><td style={cell}>{employee?.name}</td><td style={labelCell}>Employee ID</td><td style={cell}>{employee?.employeeId}</td></tr>
          <tr><td style={labelCell}>Department</td><td style={cell}>{employee?.department}</td><td style={labelCell}>Designation</td><td style={cell}>{employee?.role}</td></tr>
          <tr><td style={labelCell}>Work Status</td><td style={cell}>{employee?.workStatus}</td><td style={labelCell}>Date of Joining</td><td style={cell}>{employee?.joiningDate}</td></tr>
          <tr><td style={labelCell}>Bank Name</td><td style={cell}>{employee?.bankName}</td><td style={labelCell}>Account Number</td><td style={cell}>{employee?.accountNumber}</td></tr>
          <tr><td style={labelCell}>IFSC Code</td><td style={cell}>{employee?.ifsc}</td><td style={labelCell}>UAN (PF)</td><td style={cell}>{employee?.pfNumber}</td></tr>
          <tr><td style={labelCell}>ESI Number</td><td style={cell}>{employee?.esiNumber}</td><td style={labelCell}>PAN Number</td><td style={cell}>{employee?.pan}</td></tr>
          <tr><td style={labelCell}>Aadhaar Number</td><td style={cell}>{employee?.aadhaar}</td><td style={labelCell}>Payment Date</td><td style={cell}>{slip.paymentDate}</td></tr>
        </tbody>
      </table>

      <div style={{ fontSize: "13px", fontWeight: "bold", margin: "12px 0 6px" }}>EARNINGS &amp; DEDUCTIONS</div>
      <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "12px" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Earnings</th><th style={{ ...cell, textAlign: "right" }}>Amount (₹)</th>
            <th style={cell}>Deductions</th><th style={{ ...cell, textAlign: "right" }}>Amount (₹)</th>
          </tr>
        </thead>
        <tbody>
          {Array.from({ length: Math.max(EARNINGS_FIELDS.length, DEDUCTION_FIELDS.length) }).map((_, i) => {
            const e = EARNINGS_FIELDS[i];
            const d = DEDUCTION_FIELDS[i];
            return (
              <tr key={i}>
                <td style={cell}>{e ? e.label : ""}</td>
                <td style={{ ...cell, textAlign: "right" }}>{e ? fmt(slip.earnings[e.key] || 0) : ""}</td>
                <td style={cell}>{d ? d.label : ""}</td>
                <td style={{ ...cell, textAlign: "right" }}>{d ? fmt(slip.deductions[d.key] || 0) : ""}</td>
              </tr>
            );
          })}
          <tr style={{ fontWeight: "bold" }}>
            <td style={cell}>Gross Earnings</td>
            <td style={{ ...cell, textAlign: "right" }}>{fmt(gross)}</td>
            <td style={cell}>Total Deductions</td>
            <td style={{ ...cell, textAlign: "right" }}>{fmt(totalDeductions)}</td>
          </tr>
          <tr style={{ fontWeight: "bold", background: "#fef2f2" }}>
            <td style={cell} colSpan={3}>Net Salary Payable</td>
            <td style={{ ...cell, textAlign: "right" }}>{fmt(net)}</td>
          </tr>
        </tbody>
      </table>

      <div style={{ fontSize: "13px", fontWeight: "bold", margin: "12px 0 6px" }}>Payment Information</div>
      <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "16px" }}>
        <tbody>
          <tr><td style={labelCell}>Payment Mode</td><td style={cell}>{slip.paymentMode}</td><td style={labelCell}>Transaction ID / UTR</td><td style={cell}>{slip.transactionId}</td></tr>
          <tr><td style={labelCell}>Salary Status</td><td style={cell} colSpan={3}>{slip.salaryStatus}</td></tr>
        </tbody>
      </table>

      <div style={{ fontSize: "10px", color: "#777", marginBottom: "36px" }}>
        This is a computer-generated salary slip.
      </div>

      <table style={{ width: "100%", fontSize: "12px" }}>
        <tbody>
          <tr>
            <td style={{ width: "33%" }}>____________________<br />Employee Signature</td>
            <td style={{ width: "33%" }}>____________________<br />HR Manager</td>
            <td style={{ width: "33%" }}>____________________<br />Authorized Signatory</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ---------- LEGAL DOCUMENTS ----------
const LEGAL_CATEGORIES = [
  "Incorporation Certificate", "GST Certificate", "PAN", "Lease Agreement",
  "MOA / AOA", "Trade License", "Trademark / IP", "Insurance Policy",
  "Bank Documents", "Other",
];

const blankLegalDoc = () => ({
  title: "",
  category: "Incorporation Certificate",
  issueDate: "",
  expiryDate: "",
  notes: "",
});

function expiryTone(expiryDate) {
  if (!expiryDate) return null;
  const days = (new Date(expiryDate) - new Date(todayISO())) / (1000 * 60 * 60 * 24);
  if (days < 0) return { tone: "red", text: "Expired" };
  if (days <= 30) return { tone: "amber", text: `Expires in ${Math.ceil(days)}d` };
  return { tone: "green", text: "Valid" };
}

function LegalDocsTab({ docs, setDocs, customCategories, setCustomCategories }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [form, setForm] = useState(blankLegalDoc());
  const [showNewCategory, setShowNewCategory] = useState(false);
  const [newCategoryText, setNewCategoryText] = useState("");

  const openAdd = () => { setForm(blankLegalDoc()); setEditingId(null); setActiveId(uid()); setShowNewCategory(false); setOpen(true); };
  const openEdit = (row) => { setForm({ ...blankLegalDoc(), ...row }); setEditingId(row.id); setActiveId(row.id); setShowNewCategory(false); setOpen(true); };

  const save = () => {
    if (!form.title) return;
    if (editingId) {
      setDocs(docs.map((d) => (d.id === editingId ? { ...d, ...form, id: editingId } : d)));
    } else {
      setDocs([...docs, { id: activeId, ...form }]);
    }
    setForm(blankLegalDoc());
    setEditingId(null);
    setOpen(false);
  };

  // Built-in categories stay fixed; your own categories slot in before "Other",
  // which always stays available as a catch-all fallback.
  const allCategories = [...LEGAL_CATEGORIES.filter((c) => c !== "Other"), ...customCategories, "Other"];

  const addCustomCategory = () => {
    const name = newCategoryText.trim();
    if (!name) return;
    const exists = allCategories.some((c) => c.toLowerCase() === name.toLowerCase());
    if (!exists) setCustomCategories([...customCategories, name]);
    setForm({ ...form, category: name });
    setNewCategoryText("");
    setShowNewCategory(false);
  };

  const removeCustomCategory = (name) => {
    setCustomCategories(customCategories.filter((c) => c !== name));
  };

  return (
    <div className="space-y-4">
      <div>
        <h2 className="font-bold text-lg text-neutral-900">Legal Documents</h2>
        <p className="text-xs text-neutral-400 mt-0.5">
          Company-wide records — incorporation certificate, GST certificate, lease agreement, and other documents not tied to a specific transaction.
        </p>
      </div>

      {customCategories.length > 0 && (
        <div className="bg-white border border-neutral-200 rounded-xl p-3">
          <div className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1.5">Your custom categories</div>
          <div className="flex flex-wrap gap-1.5">
            {customCategories.map((c) => (
              <span key={c} className="inline-flex items-center gap-1 text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200 rounded-full pl-2.5 pr-1.5 py-0.5">
                {c}
                <button onClick={() => removeCustomCategory(c)} className="text-blue-400 hover:text-red-600" title="Remove this category">
                  <X size={12} />
                </button>
              </span>
            ))}
          </div>
        </div>
      )}

      <div className="flex justify-end">
        <AddButton onClick={openAdd} text="Add document" />
      </div>

      <Table
        emptyMsg="No legal documents added yet."
        columns={[
          { key: "title", label: "Document" },
          { key: "category", label: "Category" },
          { key: "issueDate", label: "Issued" },
          {
            key: "expiryDate",
            label: "Expiry",
            render: (r) => {
              const info = expiryTone(r.expiryDate);
              return r.expiryDate ? (
                <span>
                  {r.expiryDate} {info && <Pill tone={info.tone}>{info.text}</Pill>}
                </span>
              ) : (
                <span className="text-neutral-400">No expiry</span>
              );
            },
          },
          { key: "notes", label: "Notes" },
        ]}
        rows={[...docs].sort((a, b) => (a.title > b.title ? 1 : -1))}
        onDelete={(id) => setDocs(docs.filter((d) => d.id !== id))}
        onEdit={openEdit}
      />

      {open && (
        <Modal title={editingId ? "Edit legal document" : "Add legal document"} onClose={() => setOpen(false)}>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Document title">
              <input className={inputCls} value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} placeholder="e.g. Certificate of Incorporation" />
            </Field>
            <Field label="Category">
              <select
                className={inputCls}
                value={showNewCategory ? "__new__" : form.category}
                onChange={(e) => {
                  if (e.target.value === "__new__") {
                    setShowNewCategory(true);
                  } else {
                    setShowNewCategory(false);
                    setForm({ ...form, category: e.target.value });
                  }
                }}
              >
                {allCategories.map((c) => <option key={c}>{c}</option>)}
                <option value="__new__">+ Add new category…</option>
              </select>
              {showNewCategory && (
                <div className="flex gap-2 mt-1.5">
                  <input
                    className={inputCls}
                    placeholder="e.g. Fire Safety Certificate"
                    value={newCategoryText}
                    onChange={(e) => setNewCategoryText(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && (e.preventDefault(), addCustomCategory())}
                  />
                  <button
                    onClick={addCustomCategory}
                    className="shrink-0 bg-red-600 hover:bg-red-700 text-white text-sm font-semibold px-3 rounded-lg transition"
                  >
                    Add
                  </button>
                </div>
              )}
              {!showNewCategory && (
                <span className="text-[11px] text-neutral-400">Don't see the right one? Pick "+ Add new category…" — it'll be saved for next time.</span>
              )}
            </Field>
            <Field label="Issue date">
              <input type="date" className={inputCls} value={form.issueDate} onChange={(e) => setForm({ ...form, issueDate: e.target.value })} />
            </Field>
            <Field label="Expiry date (if applicable)">
              <input type="date" className={inputCls} value={form.expiryDate} onChange={(e) => setForm({ ...form, expiryDate: e.target.value })} />
            </Field>
            <div className="col-span-2">
              <Field label="Notes">
                <input className={inputCls} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} />
              </Field>
            </div>
          </div>
          {activeId && <AttachmentsField recordId={activeId} label="Document file(s)" />}
          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update document" : "Save document"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- CLIENT / VENDOR REGISTRATION ----------
function formatGSTIN(raw) {
  return raw.replace(/[^a-zA-Z0-9]/g, "").toUpperCase().slice(0, 15);
}
const GSTIN_REGEX = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/;

const blankPartyReg = () => ({
  partyType: "Client",
  name: "",
  address: "",
  gstin: "",
  msmeStatus: "Non-MSME",
  beneficiaryName: "",
  bankName: "",
  accountNumber: "",
  ifsc: "",
  authority1Name: "",
  authority1Contact: "",
  authority1Email: "",
  authority2Name: "",
  authority2Contact: "",
  authority2Email: "",
});

function PartyRegistrationTab({ registrations, setRegistrations, company, setPrintContent, setPrintTitle }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [form, setForm] = useState(blankPartyReg());
  const [query, setQuery] = useState("");
  const [viewType, setViewType] = useState("Client");

  const openAdd = () => { setForm({ ...blankPartyReg(), partyType: viewType }); setEditingId(null); setActiveId(uid()); setOpen(true); };
  const openEdit = (row) => { setForm({ ...blankPartyReg(), ...row }); setEditingId(row.id); setActiveId(row.id); setOpen(true); };

  const save = () => {
    if (!form.name) return;
    if (editingId) {
      setRegistrations(registrations.map((r) => (r.id === editingId ? { ...r, ...form, id: editingId } : r)));
    } else {
      setRegistrations([...registrations, { id: activeId, ...form }]);
    }
    setForm(blankPartyReg());
    setEditingId(null);
    setOpen(false);
  };

  const nameLabel = form.partyType === "Vendor" ? "Vendor Name" : "Client Name";

  const clientCount = registrations.filter((r) => r.partyType === "Client").length;
  const vendorCount = registrations.filter((r) => r.partyType === "Vendor").length;

  // Everything below is scoped to whichever type is selected in the dropdown —
  // switching it changes the whole view, not just a filter on top of a mixed list.
  const viewRegistrations = registrations.filter((r) => r.partyType === viewType);
  const viewMsmeCount = viewRegistrations.filter((r) => r.msmeStatus === "MSME").length;
  const viewNonMsmeCount = viewRegistrations.length - viewMsmeCount;

  const q = query.trim().toLowerCase();
  const filteredRegistrations = q
    ? viewRegistrations.filter((r) => (r.name || "").toLowerCase().includes(q))
    : viewRegistrations;

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center flex-wrap gap-3">
        <div>
          <h2 className="font-bold text-lg text-neutral-900">Client / Vendor Registration</h2>
          <p className="text-xs text-neutral-400 mt-0.5">Registered office, GSTIN, MSME status, bank details, and authority contacts for each client or vendor.</p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <select
            className={`${inputCls} font-semibold`}
            value={viewType}
            onChange={(e) => setViewType(e.target.value)}
          >
            <option value="Client">Client</option>
            <option value="Vendor">Vendor</option>
          </select>
          <div className="relative">
            <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              className={`${inputCls} pl-8 w-56`}
              placeholder={`Search ${viewType.toLowerCase()}…`}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <AddButton onClick={openAdd} text={`Add ${viewType.toLowerCase()}`} />
        </div>
      </div>

      <div className="grid grid-cols-2 gap-4">
        <Card label={`${viewType}s Registered`} value={viewRegistrations.length} sub={`Out of ${registrations.length} total — ${clientCount} Clients · ${vendorCount} Vendors`} />
        <Card label={`${viewType} MSME`} value={viewMsmeCount} tone={viewMsmeCount ? "good" : "default"} sub={`${viewNonMsmeCount} Non-MSME`} />
      </div>

      <Table
        emptyMsg={
          q
            ? `No ${viewType.toLowerCase()} matches "${query}".`
            : `No ${viewType.toLowerCase()} registrations added yet.`
        }
        columns={[
          { key: "name", label: "Name" },
          { key: "gstin", label: "GSTIN" },
          { key: "msmeStatus", label: "MSME Status", render: (r) => <Pill tone={r.msmeStatus === "MSME" ? "green" : "gray"}>{r.msmeStatus}</Pill> },
          { key: "authority1Name", label: "Primary Contact", render: (r) => r.authority1Name ? `${r.authority1Name} (${r.authority1Contact || "—"})` : "—" },
          { key: "authority2Name", label: "Secondary Contact", render: (r) => r.authority2Name ? `${r.authority2Name} (${r.authority2Contact || "—"})` : "—" },
          {
            key: "download",
            label: "",
            render: (r) => (
              <button
                onClick={() => {
                  setPrintTitle(`${r.name}-Company-Details`);
                  setPrintContent(<PartyRegistrationPrintLayout reg={r} company={company} />);
                }}
                className="inline-flex items-center gap-1 text-xs font-semibold text-red-600 hover:text-red-700 whitespace-nowrap"
              >
                <Printer size={13} /> Download File
              </button>
            ),
          },
        ]}
        rows={filteredRegistrations}
        onDelete={(id) => setRegistrations(registrations.filter((r) => r.id !== id))}
        onEdit={openEdit}
      />

      {open && (
        <Modal title={editingId ? "Edit registration" : "Add client / vendor registration"} onClose={() => setOpen(false)}>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Party type">
              <select className={inputCls} value={form.partyType} onChange={(e) => setForm({ ...form, partyType: e.target.value })}>
                <option>Client</option>
                <option>Vendor</option>
              </select>
            </Field>
            <Field label={nameLabel}>
              <input className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </Field>

            <div className="col-span-2">
              <Field label="Registered Office Address">
                <input className={inputCls} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} />
              </Field>
            </div>

            <Field label="GSTIN No.">
              <input
                className={inputCls}
                value={form.gstin}
                maxLength={15}
                placeholder="22AAAAA0000A1Z5"
                onChange={(e) => setForm({ ...form, gstin: formatGSTIN(e.target.value) })}
              />
              {form.gstin.length === 15 && !GSTIN_REGEX.test(form.gstin) && (
                <span className="text-xs text-amber-600">Doesn't match the standard GSTIN pattern — double-check it.</span>
              )}
            </Field>
            <Field label="MSME Status">
              <select className={inputCls} value={form.msmeStatus} onChange={(e) => setForm({ ...form, msmeStatus: e.target.value })}>
                <option>MSME</option>
                <option>Non-MSME</option>
              </select>
            </Field>

            <div className="col-span-2 border-t border-neutral-100 pt-3 text-sm font-semibold text-neutral-700">Bank details</div>
            <Field label="Beneficiary name">
              <input className={inputCls} value={form.beneficiaryName} onChange={(e) => setForm({ ...form, beneficiaryName: e.target.value })} />
            </Field>
            <Field label="Bank name">
              <input className={inputCls} value={form.bankName} onChange={(e) => setForm({ ...form, bankName: e.target.value })} />
            </Field>
            <Field label="Account number">
              <input className={inputCls} value={form.accountNumber} onChange={(e) => setForm({ ...form, accountNumber: e.target.value })} />
            </Field>
            <Field label="IFSC code">
              <input className={inputCls} value={form.ifsc} onChange={(e) => setForm({ ...form, ifsc: e.target.value })} />
            </Field>

            <div className="col-span-2 border-t border-neutral-100 pt-3 text-sm font-semibold text-neutral-700">Authority 1</div>
            <Field label="Authority name">
              <input className={inputCls} value={form.authority1Name} onChange={(e) => setForm({ ...form, authority1Name: e.target.value })} />
            </Field>
            <Field label="Contact number">
              <input
                className={inputCls}
                value={form.authority1Contact}
                maxLength={10}
                placeholder="10-digit mobile number"
                onChange={(e) => setForm({ ...form, authority1Contact: formatPhone10(e.target.value) })}
              />
            </Field>
            <div className="col-span-2">
              <Field label="Email ID">
                <input type="email" className={inputCls} value={form.authority1Email} onChange={(e) => setForm({ ...form, authority1Email: e.target.value })} />
              </Field>
            </div>

            <div className="col-span-2 border-t border-neutral-100 pt-3 text-sm font-semibold text-neutral-700">Authority 2 (optional)</div>
            <Field label="Authority name">
              <input className={inputCls} value={form.authority2Name} onChange={(e) => setForm({ ...form, authority2Name: e.target.value })} />
            </Field>
            <Field label="Contact number">
              <input
                className={inputCls}
                value={form.authority2Contact}
                maxLength={10}
                placeholder="10-digit mobile number"
                onChange={(e) => setForm({ ...form, authority2Contact: formatPhone10(e.target.value) })}
              />
            </Field>
            <div className="col-span-2">
              <Field label="Email ID">
                <input type="email" className={inputCls} value={form.authority2Email} onChange={(e) => setForm({ ...form, authority2Email: e.target.value })} />
              </Field>
            </div>
          </div>

          <div className="border-t border-neutral-100 pt-3">
            <div className="text-sm font-semibold text-neutral-700 mb-1">Document soft copies</div>
            <p className="text-xs text-neutral-400 mb-2">Upload scanned/photographed copies (image or PDF, under ~1.5 MB each).</p>
            {activeId && (
              <div className="space-y-3">
                <AttachmentsField recordId={`${activeId}-gstin`} label="GSTIN Certificate" />
                <AttachmentsField recordId={`${activeId}-pan`} label="PAN Card" />
                <AttachmentsField recordId={`${activeId}-msme`} label="MSME Certificate" />
                <AttachmentsField recordId={`${activeId}-incorp`} label="Certificate of Incorporation" />
                <AttachmentsField recordId={`${activeId}-cheque`} label="Cancelled Bank Cheque" />
              </div>
            )}
          </div>

          <div className="mt-4 flex justify-end">
            <AddButton onClick={save} text={editingId ? "Update registration" : "Save registration"} />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- PRODUCTION WORKFLOW (Kanban) ----------
const PRODUCTION_STAGES = [
  { key: "marketing", label: "Marketing", desc: "Requirement / remarks from client" },
  { key: "cad", label: "CAD Design", desc: "Design work + approval remarks" },
  { key: "machining", label: "Machining / 3D Printing / Preparation", desc: "Production based on CAD approval" },
  { key: "assembly", label: "Assembly & Painting", desc: "With 3D Printing remarks" },
  { key: "inspection", label: "Final Inspection", desc: "QC against original requirement" },
  { key: "completed", label: "Completed & Packaging / Dispatched", desc: "Remarks from QC team" },
];

// Mild/pastel color per stage so cards visually shift as they move through the Kanban board.
const STAGE_COLORS = {
  marketing:  { col: "bg-sky-50",     header: "text-sky-700",     badge: "bg-sky-100 text-sky-600",         card: "bg-sky-50/70 border-sky-200 hover:border-sky-400" },
  cad:        { col: "bg-violet-50",  header: "text-violet-700",  badge: "bg-violet-100 text-violet-600",   card: "bg-violet-50/70 border-violet-200 hover:border-violet-400" },
  machining:  { col: "bg-amber-50",   header: "text-amber-700",   badge: "bg-amber-100 text-amber-600",     card: "bg-amber-50/70 border-amber-200 hover:border-amber-400" },
  assembly:   { col: "bg-orange-50",  header: "text-orange-700",  badge: "bg-orange-100 text-orange-600",   card: "bg-orange-50/70 border-orange-200 hover:border-orange-400" },
  inspection: { col: "bg-cyan-50",    header: "text-cyan-700",    badge: "bg-cyan-100 text-cyan-600",       card: "bg-cyan-50/70 border-cyan-200 hover:border-cyan-400" },
  completed:  { col: "bg-emerald-50", header: "text-emerald-700", badge: "bg-emerald-100 text-emerald-600", card: "bg-emerald-50/70 border-emerald-200 hover:border-emerald-400" },
};

const blankStageData = () => ({ remarks: "", responsibleId: "", status: "Not started", date: "", pmApproved: false, pmApprovedBy: "", pmApprovalDate: "" });

function blankWorkflowRecord(orderId) {
  const stages = {};
  PRODUCTION_STAGES.forEach((s, i) => (stages[s.key] = { ...blankStageData(), status: i === 0 ? "In progress" : "Not started" }));
  return { id: orderId, orderId, currentStageIndex: 0, stages };
}

function ProjectCompletionPrintLayout({ order, rec, employees, company }) {
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
  const labelCell = { border: "1px solid #ddd", padding: "6px", fontSize: "12px", background: "#f8f8f8", fontWeight: "bold", width: "22%" };
  const cell = { border: "1px solid #ddd", padding: "6px", fontSize: "12px" };
  const dispatchDate = rec.stages.completed.date;

  return (
    <div style={{ fontFamily: "Arial, sans-serif", color: "#111", padding: "24px", maxWidth: "850px", margin: "0 auto" }}>
      <div style={{ borderBottom: "2px solid #dc2626", paddingBottom: "10px", marginBottom: "16px" }}>
        <div style={{ fontSize: "20px", fontWeight: "bold" }}>{company.name}</div>
        <div style={{ fontSize: "12px", color: "#555" }}>{company.address}</div>
      </div>

      <div style={{ textAlign: "center", fontSize: "16px", fontWeight: "bold", letterSpacing: "1px", margin: "10px 0 4px" }}>
        PROJECT COMPLETION REPORT
      </div>
      <div style={{ textAlign: "center", fontSize: "12px", color: "#555", marginBottom: "16px" }}>
        Completed &amp; Dispatched on {dispatchDate}
      </div>

      <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "16px" }}>
        <tbody>
          <tr><td style={labelCell}>Project / Order ID</td><td style={cell} colSpan={3}>{order.name}</td></tr>
          <tr><td style={labelCell}>Client</td><td style={cell} colSpan={3}>{order.client}</td></tr>
          <tr><td style={labelCell}>Start Date</td><td style={cell}>{order.startDate}</td><td style={labelCell}>Deadline</td><td style={cell}>{order.deadline}</td></tr>
        </tbody>
      </table>

      <div style={{ fontSize: "13px", fontWeight: "bold", margin: "12px 0 6px" }}>Production Stage History</div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f3f3f3" }}>
            <th style={cell}>Stage</th>
            <th style={cell}>Responsible</th>
            <th style={cell}>Remarks</th>
            <th style={cell}>PM Approved By</th>
            <th style={cell}>Date Completed</th>
          </tr>
        </thead>
        <tbody>
          {PRODUCTION_STAGES.map((s) => {
            const st = rec.stages[s.key];
            return (
              <tr key={s.key}>
                <td style={cell}>{s.label}</td>
                <td style={cell}>{st.responsibleId ? empName(st.responsibleId) : "—"}</td>
                <td style={cell}>{st.remarks || "—"}</td>
                <td style={cell}>{st.pmApprovedBy ? `${empName(st.pmApprovedBy)} (${st.pmApprovalDate})` : "—"}</td>
                <td style={cell}>{st.date || "—"}</td>
              </tr>
            );
          })}
        </tbody>
      </table>

      <div style={{ fontSize: "10px", color: "#777", margin: "20px 0 36px" }}>
        This is a computer-generated project completion report.
      </div>

      <table style={{ width: "100%", fontSize: "12px" }}>
        <tbody>
          <tr>
            <td style={{ width: "33%" }}>____________________<br />Prepared by</td>
            <td style={{ width: "33%" }}>____________________<br />QC Approved by</td>
            <td style={{ width: "33%" }}>____________________<br />Authorized Signatory</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function ProductionWorkflowTab({ orders, setOrders, employees, workflow, setWorkflow, company, setPrintContent, setPrintTitle }) {
  const [detailOrderId, setDetailOrderId] = useState(null);

  // The link: any order without a workflow record yet gets one automatically,
  // starting at the Marketing stage — this is what makes new Orders/Projects
  // entries show up here without any manual setup.
  useEffect(() => {
    const missing = orders.filter((o) => !workflow.some((w) => w.orderId === o.id));
    if (missing.length > 0) {
      setWorkflow([...workflow, ...missing.map((o) => blankWorkflowRecord(o.id))]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [orders.length]);

  const recordFor = (orderId) => workflow.find((w) => w.orderId === orderId) || blankWorkflowRecord(orderId);
  const orderFor = (orderId) => orders.find((o) => o.id === orderId);
  const empName = (id) => employees.find((e) => e.id === id)?.name || "";

  const updateRecord = (orderId, updater) => {
    setWorkflow((prev) => {
      const exists = prev.some((w) => w.orderId === orderId);
      const base = exists ? prev.find((w) => w.orderId === orderId) : blankWorkflowRecord(orderId);
      const updated = updater({ ...base, stages: { ...base.stages } });
      return exists ? prev.map((w) => (w.orderId === orderId ? updated : w)) : [...prev, updated];
    });
  };

  const updateStageField = (orderId, stageKey, field, value) => {
    updateRecord(orderId, (r) => ({
      ...r,
      stages: { ...r.stages, [stageKey]: { ...r.stages[stageKey], [field]: value } },
    }));
  };

  const approveStage = (orderId, stageKey, approverId) => {
    updateRecord(orderId, (r) => ({
      ...r,
      stages: {
        ...r.stages,
        [stageKey]: { ...r.stages[stageKey], pmApproved: true, pmApprovedBy: approverId, pmApprovalDate: todayISO() },
      },
    }));
  };

  const revokeApproval = (orderId, stageKey) => {
    updateRecord(orderId, (r) => ({
      ...r,
      stages: { ...r.stages, [stageKey]: { ...r.stages[stageKey], pmApproved: false, pmApprovalDate: "" } },
    }));
  };

  const moveNext = (orderId) => {
    const rec = recordFor(orderId);
    const stageKey = PRODUCTION_STAGES[rec.currentStageIndex].key;
    if (!rec.stages[stageKey].pmApproved) return; // gated: needs Program Manager approval first
    const newIndex = Math.min(PRODUCTION_STAGES.length - 1, rec.currentStageIndex + 1);
    const newStageKey = PRODUCTION_STAGES[newIndex].key;
    updateRecord(orderId, (r) => ({
      ...r,
      currentStageIndex: newIndex,
      stages: {
        ...r.stages,
        [stageKey]: { ...r.stages[stageKey], status: "Done", date: r.stages[stageKey].date || todayISO() },
        [newStageKey]: { ...r.stages[newStageKey], status: r.stages[newStageKey].status === "Done" ? "Done" : "In progress" },
      },
    }));
  };

  const moveBack = (orderId) => {
    const rec = recordFor(orderId);
    const newIndex = Math.max(0, rec.currentStageIndex - 1);
    updateRecord(orderId, (r) => ({ ...r, currentStageIndex: newIndex }));
  };

  const markCompleted = (orderId) => {
    const rec = recordFor(orderId);
    if (!rec.stages.completed.pmApproved) return; // gated: needs Program Manager approval first
    updateRecord(orderId, (r) => ({
      ...r,
      stages: { ...r.stages, completed: { ...r.stages.completed, status: "Done", date: r.stages.completed.date || todayISO() } },
    }));
    setOrders((prev) => prev.map((o) => (o.id === orderId ? { ...o, status: "Delivered" } : o)));
  };

  const detailOrder = detailOrderId ? orderFor(detailOrderId) : null;
  const detailRec = detailOrderId ? recordFor(detailOrderId) : null;

  return (
    <div className="space-y-4">
      <div>
        <h2 className="font-bold text-lg text-neutral-900">Production Workflow</h2>
        <p className="text-xs text-neutral-400 mt-0.5">
          Every order from Orders/Projects flows through here automatically — click a card to move it forward and log remarks.
        </p>
      </div>

      {orders.length === 0 ? (
        <div className="text-center text-neutral-400 text-sm py-14 border border-dashed border-neutral-200 rounded-xl">
          Add an order in Orders/Projects first — it'll appear here in Marketing automatically.
        </div>
      ) : (
        <div className="flex gap-3 overflow-x-auto pb-2">
          {PRODUCTION_STAGES.map((stage, colIndex) => {
            const colOrders = orders.filter((o) => recordFor(o.id).currentStageIndex === colIndex);
            const sc = STAGE_COLORS[stage.key];
            return (
              <div key={stage.key} className={`w-64 shrink-0 ${sc.col} rounded-xl p-2 transition-colors`}>
                <div className="px-2 py-1.5 mb-2">
                  <div className={`text-xs font-bold uppercase tracking-wide ${sc.header}`}>{stage.label}</div>
                  <div className="text-[11px] text-neutral-400">
                    {stage.desc} · <span className={`inline-block px-1.5 py-0.5 rounded-full font-semibold ${sc.badge}`}>{colOrders.length}</span>
                  </div>
                </div>
                <div className="space-y-2">
                  {colOrders.map((o) => {
                    const rec = recordFor(o.id);
                    const stageData = rec.stages[stage.key];
                    const isCompletedCard = stage.key === "completed" && stageData?.status === "Done";
                    return (
                      <div
                        key={o.id}
                        onClick={() => setDetailOrderId(o.id)}
                        className={`w-full text-left border rounded-lg p-3 shadow-sm transition cursor-pointer ${sc.card}`}
                      >
                        <div className="font-semibold text-sm text-neutral-900">{o.name}</div>
                        <div className="text-xs text-neutral-500 mb-1">{o.client}</div>
                        {o.deadline && <div className="text-[11px] text-neutral-400">Deadline: {o.deadline}</div>}
                        {stageData?.responsibleId && (
                          <div className="text-[11px] text-neutral-500 mt-1">👤 {empName(stageData.responsibleId)}</div>
                        )}
                        {stageData?.remarks && (
                          <div className="text-[11px] text-neutral-500 mt-1 line-clamp-2">"{stageData.remarks}"</div>
                        )}
                        {isCompletedCard && (
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              setPrintTitle(`Project-Completion-${o.name}`);
                              setPrintContent(<ProjectCompletionPrintLayout order={o} rec={rec} employees={employees} company={company} />);
                            }}
                            className="mt-2 inline-flex items-center gap-1 text-[11px] font-semibold text-red-600 hover:text-red-700"
                          >
                            <Printer size={12} /> Download PDF
                          </button>
                        )}
                      </div>
                    );
                  })}
                  {colOrders.length === 0 && (
                    <div className="text-center text-[11px] text-neutral-300 py-4">No projects here</div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {detailOrder && detailRec && (
        <Modal title={`${detailOrder.name} — Production Workflow`} onClose={() => setDetailOrderId(null)}>
          <div className="flex items-start justify-between gap-3 mb-4">
            <div className="text-xs text-neutral-500">
              Client: <strong className="text-neutral-700">{detailOrder.client}</strong>
              {detailOrder.deadline && <> · Deadline: <strong className="text-neutral-700">{detailOrder.deadline}</strong></>}
            </div>
            {detailRec.stages.completed.status === "Done" && (
              <button
                onClick={() => {
                  setPrintTitle(`Project-Completion-${detailOrder.name}`);
                  setPrintContent(<ProjectCompletionPrintLayout order={detailOrder} rec={detailRec} employees={employees} company={company} />);
                }}
                className="inline-flex items-center gap-1.5 bg-white border border-neutral-300 hover:border-red-400 text-xs font-semibold px-3 py-2 rounded-lg transition shrink-0"
              >
                <Printer size={14} /> Download PDF
              </button>
            )}
          </div>

          {/* Stepper */}
          <div className="flex items-center mb-5 overflow-x-auto pb-1">
            {PRODUCTION_STAGES.map((s, i) => {
              const st = detailRec.stages[s.key];
              const isCurrent = i === detailRec.currentStageIndex;
              const isDone = st.status === "Done";
              return (
                <div key={s.key} className="flex items-center shrink-0">
                  <div
                    className={`flex flex-col items-center px-2 ${isCurrent ? "" : ""}`}
                    title={s.label}
                  >
                    <div
                      className={`w-6 h-6 rounded-full flex items-center justify-center text-[10px] font-bold ${
                        isDone
                          ? "bg-emerald-600 text-white"
                          : isCurrent
                          ? "bg-red-600 text-white"
                          : "bg-neutral-200 text-neutral-500"
                      }`}
                    >
                      {isDone ? "✓" : i + 1}
                    </div>
                    <div className={`text-[10px] mt-1 w-16 text-center leading-tight ${isCurrent ? "text-red-600 font-semibold" : "text-neutral-400"}`}>
                      {s.label}
                    </div>
                  </div>
                  {i < PRODUCTION_STAGES.length - 1 && <div className="w-6 h-0.5 bg-neutral-200 mx-0.5" />}
                </div>
              );
            })}
          </div>

          {/* Current stage editor */}
          {(() => {
            const stage = PRODUCTION_STAGES[detailRec.currentStageIndex];
            const st = detailRec.stages[stage.key];
            const isLast = detailRec.currentStageIndex === PRODUCTION_STAGES.length - 1;
            return (
              <div className="border border-neutral-200 rounded-xl p-4 space-y-3">
                <div className="text-sm font-semibold text-neutral-800">Current stage: {stage.label}</div>
                <p className="text-xs text-neutral-400 -mt-2">{stage.desc}</p>
                <Field label="Responsible">
                  <select
                    className={inputCls}
                    value={st.responsibleId}
                    onChange={(e) => updateStageField(detailOrder.id, stage.key, "responsibleId", e.target.value)}
                  >
                    <option value="">— none —</option>
                    {employees.map((e) => (
                      <option key={e.id} value={e.id}>{e.name}</option>
                    ))}
                  </select>
                </Field>
                <Field label="Remarks">
                  <input
                    className={inputCls}
                    value={st.remarks}
                    onChange={(e) => updateStageField(detailOrder.id, stage.key, "remarks", e.target.value)}
                    placeholder="Requirement notes, approval remarks, QC notes…"
                  />
                </Field>
                <AttachmentsField recordId={`${detailOrder.id}-${stage.key}`} label="Attachments (CAD file, inspection sheet, etc.)" />

                <div className="border-t border-neutral-100 pt-3">
                  <div className="text-sm font-semibold text-neutral-700 mb-2">Program Manager Approval</div>
                  {st.pmApproved ? (
                    <div className="flex items-center justify-between bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2">
                      <div className="text-xs text-emerald-700">
                        ✓ Approved by <strong>{empName(st.pmApprovedBy) || "—"}</strong> on {st.pmApprovalDate}
                      </div>
                      <button
                        onClick={() => revokeApproval(detailOrder.id, stage.key)}
                        className="text-xs font-semibold text-neutral-400 hover:text-red-600"
                      >
                        Revoke
                      </button>
                    </div>
                  ) : (
                    <div className="space-y-2">
                      <p className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
                        This stage can't move forward until a Program Manager approves it here.
                      </p>
                      <div className="flex gap-2">
                        <select id={`pm-select-${detailOrder.id}`} key={stage.key} className={`${inputCls} flex-1`} defaultValue="">
                          <option value="">— select Program Manager —</option>
                          {employees.map((e) => (
                            <option key={e.id} value={e.id}>{e.name}</option>
                          ))}
                        </select>
                        <button
                          onClick={() => {
                            const sel = document.getElementById(`pm-select-${detailOrder.id}`);
                            if (sel && sel.value) approveStage(detailOrder.id, stage.key, sel.value);
                          }}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3 py-2 rounded-lg transition whitespace-nowrap"
                        >
                          Approve stage
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <div className="flex justify-between pt-2">
                  <button
                    onClick={() => moveBack(detailOrder.id)}
                    disabled={detailRec.currentStageIndex === 0}
                    className="inline-flex items-center gap-1 text-xs font-semibold text-neutral-500 hover:text-neutral-800 disabled:opacity-30 disabled:cursor-not-allowed"
                  >
                    <ArrowLeft size={14} /> Move back
                  </button>
                  {isLast ? (
                    <button
                      onClick={() => markCompleted(detailOrder.id)}
                      disabled={st.status === "Done" || !st.pmApproved}
                      title={!st.pmApproved ? "Needs Program Manager approval first" : ""}
                      className="inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold px-4 py-2 rounded-lg transition disabled:opacity-40 disabled:cursor-not-allowed"
                    >
                      ✓ {st.status === "Done" ? "Completed" : "Mark completed & dispatch"}
                    </button>
                  ) : (
                    <button
                      onClick={() => moveNext(detailOrder.id)}
                      disabled={!st.pmApproved}
                      title={!st.pmApproved ? "Needs Program Manager approval first" : ""}
                      className="inline-flex items-center gap-1.5 bg-red-600 hover:bg-red-700 text-white text-sm font-semibold px-4 py-2 rounded-lg transition disabled:opacity-40 disabled:cursor-not-allowed"
                    >
                      Move to next stage <ArrowRight size={14} />
                    </button>
                  )}
                </div>
              </div>
            );
          })()}

          {/* History of completed stages */}
          <div className="mt-4">
            <div className="text-xs font-semibold text-neutral-500 uppercase tracking-wide mb-2">Stage history</div>
            <div className="space-y-2">
              {PRODUCTION_STAGES.map((s, i) => {
                const st = detailRec.stages[s.key];
                if (st.status !== "Done") return null;
                return (
                  <div key={s.key} className="text-xs bg-neutral-50 border border-neutral-200 rounded-lg px-3 py-2">
                    <div className="font-semibold text-neutral-700">{s.label} <span className="text-neutral-400 font-normal">· {st.date}</span></div>
                    {st.responsibleId && <div className="text-neutral-500">Responsible: {empName(st.responsibleId)}</div>}
                    {st.pmApprovedBy && <div className="text-neutral-500">Approved by: {empName(st.pmApprovedBy)} (PM) on {st.pmApprovalDate}</div>}
                    {st.remarks && <div className="text-neutral-500">"{st.remarks}"</div>}
                  </div>
                );
              })}
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}
