import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, onValue, get, set as dbSet } from "firebase/database";
import { getAuth, signInWithEmailAndPassword, onAuthStateChanged, signOut } from "firebase/auth";
import { firebaseConfig, CLOUD_SYNC_ENABLED, CLOUD_CLIENTS_BASE } from "./firebaseConfig.js";
import DashboardProduccion from "./DashboardProduccion.jsx";
import DashboardCultivos from "./DashboardCultivos.jsx";
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, Legend, LabelList,
} from "recharts";
import {
  Upload, RefreshCw, Search, X, Sprout, MapPin, DollarSign, Ruler,
  ChevronLeft, ChevronRight, ArrowUpDown, Trash2, FileSpreadsheet, AlertCircle, Eye,
  ChevronDown, ChevronUp, ClipboardCheck, Download,
  Cloud, CloudOff, LogOut, Plus, ArrowLeft,
} from "lucide-react";

// ---------------------------------------------------------------------------
// Config: columnas esperadas en la planilla (nombres exactos de encabezado)
// ---------------------------------------------------------------------------
const COLS = [
  "Admin", "Campaña", "Campo", "Lote", "Sup.Lote", "Cultivo", "Variedad",
  "Sup.Cultivo", "Fecha", "OTA", "Origen", "Cant.Orig.", "U.Serv.",
  "Tipo Det.", "Tipo item", "Concepto", "Cant.Afect", "Cantidad", "Unid.",
  "U$S/U", "U$S/Total",
];
// Fuera del entorno de Claude no existe window.storage: lo reemplazamos por un
// equivalente simple basado en localStorage, con la misma forma (get/set/delete/list).
if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    async get(key, shared) {
      const raw = localStorage.getItem(key);
      if (raw === null) return null;
      return { key, value: raw, shared: !!shared };
    },
    async set(key, value, shared) {
      localStorage.setItem(key, value);
      return { key, value, shared: !!shared };
    },
    async delete(key, shared) {
      localStorage.removeItem(key);
      return { key, deleted: true, shared: !!shared };
    },
    async list(prefix, shared) {
      const keys = Object.keys(localStorage).filter((k) => !prefix || k.startsWith(prefix));
      return { keys, prefix, shared: !!shared };
    },
  };
}

// El caché local es por cliente, para no mezclar datos de distintos clientes
// en el mismo navegador (ej. si el admin revisa varios seguidos).
const STORAGE_KEY_PREFIX = "produccion-agricola-dataset-v1";
const PAGE_SIZE = 30;
// Columnas que pueden venir sin precio cargado (quedan en null). Firebase Realtime
// Database borra los valores "null" dentro de un array al guardarlos (los trata
// como una instrucción de "eliminar"), lo que descuadra la fila entera al leerla
// de vuelta. Para evitarlo, estos nulls se codifican como "" antes de guardar y
// se decodifican de nuevo a null al reconstruir las filas.
const NULLABLE_COLS = new Set(["U$S/U", "U$S/Total"]);
const encodeCell = (col, val) => (val === null && NULLABLE_COLS.has(col) ? "" : val);
const decodeCell = (col, val) => (val === "" && NULLABLE_COLS.has(col) ? null : val);
const rowsFromPayload = (payload) => payload.rows.map((arr) => {
  const o = {};
  payload.columns.forEach((c, i) => { o[c] = decodeCell(c, arr[i]); });
  return o;
});

// Conexión a Firebase (solo si se completó firebaseConfig.js con datos reales)
let cloudDb = null;
let cloudAuth = null;
if (CLOUD_SYNC_ENABLED) {
  try {
    const fbApp = initializeApp(firebaseConfig);
    cloudDb = getDatabase(fbApp);
    cloudAuth = getAuth(fbApp);
  } catch (e) {
    console.error("No se pudo inicializar Firebase:", e);
  }
}

const CULTIVO_COLORS = { SOJA: "#5B7C4B", MAIZ: "#C68F41", POROTO: "#35606B" };
const FALLBACK_COLORS = ["#5B7C4B", "#C68F41", "#35606B", "#8A5A3B", "#7C8A4B", "#A1462F"];
const cultivoColor = (name, i) => CULTIVO_COLORS[String(name).toUpperCase()] || FALLBACK_COLORS[i % FALLBACK_COLORS.length];

// Estilos compartidos por toda la app (dashboard, login y panel de admin)
const AGRI_STYLES = `
  .agri-root {
    --paper: #F6F1E2;
    --paper-raised: #FCFAF2;
    --ink: #2B2118;
    --ink-soft: #6B5E4F;
    --line: #DCD2B8;
    --green: #4B6B3A;
    --gold: #B8842E;
    --teal: #2F5B66;
    --rust: #A1462F;
    background: var(--paper);
    color: var(--ink);
    font-family: 'Inter', system-ui, sans-serif;
    min-height: 100vh;
    padding: 20px 16px 60px;
  }
  .agri-root * { box-sizing: border-box; }
  .agri-serif { font-family: 'Inter', system-ui, sans-serif; font-weight: 600; letter-spacing: -0.005em; }
  .agri-shell { max-width: 1180px; margin: 0 auto; }
  .agri-masthead {
    display: flex; align-items: flex-end; justify-content: space-between; gap: 16px;
    border-bottom: 2px solid var(--ink); padding-bottom: 14px; margin-bottom: 22px; flex-wrap: wrap;
  }
  .agri-title { font-size: 24px; font-weight: 700; letter-spacing: -0.01em; line-height: 1.1; font-family: 'Inter', system-ui, sans-serif; }
  .agri-sub { color: var(--ink-soft); font-size: 13px; margin-top: 4px; }
  .agri-card {
    background: var(--paper-raised); border: 1px solid var(--line); border-radius: 6px;
  }
  .agri-btn {
    display: inline-flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 600;
    padding: 8px 14px; border-radius: 5px; border: 1px solid var(--ink); background: var(--ink); color: var(--paper);
    cursor: pointer; transition: opacity .15s ease;
  }
  .agri-btn:hover { opacity: .85; }
  .agri-btn-outline {
    background: transparent; color: var(--ink); border: 1px solid var(--line);
  }
  .agri-btn-outline:hover { border-color: var(--ink); opacity: 1; }
  .agri-input, .agri-select {
    background: var(--paper-raised); border: 1px solid var(--line); border-radius: 5px;
    padding: 8px 10px; font-size: 13px; color: var(--ink); font-family: 'Inter', sans-serif;
  }
  .agri-input:focus, .agri-select:focus { outline: 2px solid var(--teal); outline-offset: 1px; border-color: var(--teal); }
  .agri-chip {
    font-size: 13px; font-weight: 600; padding: 6px 13px; border-radius: 999px; border: 1px solid var(--line);
    background: var(--paper-raised); cursor: pointer; color: var(--ink-soft);
  }
  .agri-chip[data-active="true"] { background: var(--green); border-color: var(--green); color: #fff; }
  .agri-kpi-label { font-size: 12px; color: var(--ink-soft); font-weight: 500; }
  .agri-kpi-value { font-family: 'Inter', system-ui, sans-serif; font-size: 20px; font-weight: 700; margin-top: 2px; font-variant-numeric: tabular-nums; }
  .agri-th {
    text-align: left; font-size: 11px; text-transform: none; color: var(--ink-soft); font-weight: 600;
    padding: 8px 10px; border-bottom: 1px solid var(--ink); cursor: pointer; white-space: nowrap; user-select: none;
  }
  .agri-td { padding: 8px 10px; font-size: 13px; border-bottom: 1px solid var(--line); font-variant-numeric: tabular-nums; }
  .agri-tr:hover { background: rgba(75,107,58,0.06); }
  .agri-dropzone {
    border: 2px dashed var(--line); border-radius: 10px; background: var(--paper-raised);
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    padding: 56px 24px; text-align: center; gap: 12px; transition: border-color .15s ease, background .15s ease;
  }
  .agri-dropzone[data-drag="true"] { border-color: var(--green); background: rgba(75,107,58,0.06); }
  @media (max-width: 640px) {
    .agri-title { font-size: 19px; }
    .agri-kpi-value { font-size: 16px; }
    .agri-detail-grid { grid-template-columns: 1fr !important; }
  }
`;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const str = (v) => (v === null || v === undefined ? "" : String(v).trim());
const num = (v) => {
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const n = parseFloat(String(v ?? "").replace(",", "."));
  return isNaN(n) ? 0 : n;
};
const numOrNull = (v) => {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return isNaN(v) ? null : v;
  const n = parseFloat(String(v).replace(",", "."));
  return isNaN(n) ? null : n;
};
// Las fechas que entrega SheetJS (cellDates:true) están ancladas en UTC medianoche.
// Usamos los getters UTC para evitar que el huso horario local corra el día (p.ej. UTC-3).
const toISODate = (v) => {
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getUTCFullYear(), m = String(v.getUTCMonth() + 1).padStart(2, "0"), d = String(v.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "string" && v) {
    const d = new Date(v);
    if (!isNaN(d)) {
      const y = d.getUTCFullYear(), m = String(d.getUTCMonth() + 1).padStart(2, "0"), dd = String(d.getUTCDate()).padStart(2, "0");
      return `${y}-${m}-${dd}`;
    }
  }
  return "";
};

const fmtUSD = (n) => new Intl.NumberFormat("es-AR", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(n || 0);
const fmtUSD2 = (n) => new Intl.NumberFormat("es-AR", { style: "currency", currency: "USD", maximumFractionDigits: 2 }).format(n || 0);
const fmtNum = (n, d = 0) => new Intl.NumberFormat("es-AR", { maximumFractionDigits: d, minimumFractionDigits: d }).format(n || 0);
const fmtDate = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
const monthLabel = (ym) => {
  const [y, m] = ym.split("-");
  const names = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  return `${names[parseInt(m, 10) - 1]} ${y.slice(2)}`;
};

// Renderers de etiquetas dentro de las barras: se ocultan solos si no entra el
// texto, para que nunca sobresalgan del área de la barra.
const VerticalBarLabel = (props) => {
  const { x, y, width, height, value } = props;
  if (!value || height < 16) return null;
  return (
    <text x={x + width / 2} y={y + height / 2} textAnchor="middle" dominantBaseline="middle" fill="#fff" fontSize={11} fontWeight={700}>
      {fmtUSD2(value)}
    </text>
  );
};

const HorizontalBarLabel = (props) => {
  const { x, y, width, height, value } = props;
  if (!value) return null;
  const text = fmtUSD2(value);
  if (width < text.length * 6.4 + 10) return null;
  return (
    <text x={x + width - 6} y={y + height / 2} textAnchor="end" dominantBaseline="middle" fill="#fff" fontSize={10} fontWeight={700}>
      {text}
    </text>
  );
};

const LoteCostoHaTooltip = ({ active, payload }) => {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  return (
    <div style={{ background: "#FCFAF2", border: "1px solid #DCD2B8", borderRadius: 6, padding: "8px 10px", fontSize: 12 }}>
      <div style={{ fontWeight: 600, marginBottom: 3 }}>{d.campo} · Lote {d.lote} · {d.cultivo}</div>
      {d.variedad && <div style={{ color: "var(--ink-soft)" }}>{d.variedad}</div>}
      <div style={{ color: "var(--ink-soft)" }}>{fmtNum(d.ha, 1)} ha</div>
      <div style={{ fontWeight: 600, marginTop: 2 }}>{fmtUSD2(d.value)}/ha</div>
    </div>
  );
};

function normalizeRow(r) {
  return {
    "Admin": str(r["Admin"]),
    "Campaña": str(r["Campaña"]),
    "Campo": str(r["Campo"]),
    "Lote": str(r["Lote"]),
    "Sup.Lote": num(r["Sup.Lote"]),
    "Cultivo": str(r["Cultivo"]),
    "Variedad": str(r["Variedad"]),
    "Sup.Cultivo": num(r["Sup.Cultivo"]),
    "Fecha": toISODate(r["Fecha"]),
    "OTA": str(r["OTA"]),
    "Origen": str(r["Origen"]),
    "Cant.Orig.": num(r["Cant.Orig."]),
    "U.Serv.": str(r["U.Serv."]),
    "Tipo Det.": str(r["Tipo Det."]),
    "Tipo item": str(r["Tipo item"]),
    "Concepto": str(r["Concepto"]),
    "Cant.Afect": num(r["Cant.Afect"]),
    "Cantidad": num(r["Cantidad"]),
    "Unid.": str(r["Unid."]),
    // El Excel de origen renombró estas dos columnas; aceptamos ambos nombres
    // por si en algún momento se sube un archivo con la nomenclatura anterior.
    "U$S/U": numOrNull(r["Precio/U"] !== undefined ? r["Precio/U"] : r["U$S/U"]),
    "U$S/Total": numOrNull(r["Total"] !== undefined ? r["Total"] : r["U$S/Total"]),
  };
}

// ---------------------------------------------------------------------------
// Componente principal
// ---------------------------------------------------------------------------
function ClienteDashboard({ slug, isAdmin }) {
  const baseUrl = typeof window !== "undefined" ? `${window.location.origin}${window.location.pathname}` : "";

  const DASHBOARDS = [
    { key: "costos",     label: "Costos Agrícolas" },
    { key: "produccion", label: "Producción" },
    { key: "cultivos",   label: "Cultivos" },
  ];

  const [dashKey, setDashKey] = useState(() => {
    if (typeof window === "undefined") return "costos";
    return new URLSearchParams(window.location.search).get("dashboard") || "costos";
  });

  useEffect(() => {
    const onPop = () => {
      const k = new URLSearchParams(window.location.search).get("dashboard") || "costos";
      setDashKey(k);
    };
    window.addEventListener("popstate", onPop);
    return () => window.removeEventListener("popstate", onPop);
  }, []);

  const navTo = (key) => {
    const u = new URLSearchParams(window.location.search);
    u.set("dashboard", key);
    window.history.pushState({}, "", `?${u.toString()}`);
    window.dispatchEvent(new Event("popstate"));
  };

  const [clienteNombre, setClienteNombre] = useState("");
  useEffect(() => {
    if (!CLOUD_SYNC_ENABLED || !cloudDb) return;
    const unsub = onValue(ref(cloudDb, `${CLOUD_CLIENTS_BASE}/index/${slug}/nombre`), (snap) => {
      setClienteNombre(snap.val() || "");
    });
    return unsub;
  }, [slug]);

  const masterhead = (
    <div className="agri-masthead">
      <div>
        {isAdmin && (
          <a href={baseUrl} style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 12, color: "var(--ink-soft)", textDecoration: "none", marginBottom: 6 }}>
            <ArrowLeft size={12} /> Panel de clientes
          </a>
        )}
        <div className="agri-title agri-serif">Costos Agrícolas - {clienteNombre || slug}</div>
        <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
          {DASHBOARDS.map((d) => (
            <button key={d.key} className="agri-chip" data-active={dashKey === d.key} onClick={() => navTo(d.key)}>
              {d.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );

  return (
    <div className="agri-shell">
      {masterhead}
      {dashKey === "produccion"
        ? <DashboardProduccion slug={slug} isAdmin={isAdmin} cloudDb={cloudDb} />
        : dashKey === "cultivos"
        ? <DashboardCultivos slug={slug} isAdmin={isAdmin} cloudDb={cloudDb} />
        : <DashboardCostos slug={slug} isAdmin={isAdmin} />
      }
    </div>
  );
}

// ---------------------------------------------------------------------------
// Dashboard de costos (extraído de ClienteDashboard para cumplir reglas de hooks)
// ---------------------------------------------------------------------------
function DashboardCostos({ slug, isAdmin }) {
  const storageKey = `${STORAGE_KEY_PREFIX}-${slug}`;
  const cloudPath = `${CLOUD_CLIENTS_BASE}/${slug}/dataset`;  const [rows, setRows] = useState([]);
  const [meta, setMeta] = useState(null); // {fileName, updatedAt, rowCount}
  const [bootLoading, setBootLoading] = useState(true);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [syncStatus, setSyncStatus] = useState(CLOUD_SYNC_ENABLED ? "connecting" : "local");
  const fileInputRef = useRef(null);

  // Filtros
  const [cultivo, setCultivo] = useState("Todos");
  const [campo, setCampo] = useState("Todos");
  const [lote, setLote] = useState("Todos");
  const [administracion, setAdministracion] = useState("Todos");
  const [tipoDet, setTipoDet] = useState("Todos");
  const [tipoItem, setTipoItem] = useState("Todos");
  const [desde, setDesde] = useState("");
  const [hasta, setHasta] = useState("");
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [sort, setSort] = useState({ key: "Fecha", dir: "desc" });
  const [showDetalle, setShowDetalle] = useState(false);

  // Cargar último dataset guardado: primero del caché local (instantáneo),
  // y si hay Firebase configurado, nos suscribimos a la nube (tiempo real).
  useEffect(() => {
    let cloudUnsub = null;

    (async () => {
      try {
        const res = await window.storage.get(storageKey, false);
        if (res && res.value) {
          const payload = JSON.parse(res.value);
          const objRows = rowsFromPayload(payload);
          setRows(objRows);
          setMeta({ fileName: payload.fileName, updatedAt: payload.updatedAt, rowCount: objRows.length });
        }
      } catch (e) {
        // no hay datos guardados todavía en este dispositivo, es normal
      } finally {
        if (!CLOUD_SYNC_ENABLED) setBootLoading(false);
      }

      if (CLOUD_SYNC_ENABLED && cloudDb) {
        const dbRef = ref(cloudDb, cloudPath);
        cloudUnsub = onValue(
          dbRef,
          (snapshot) => {
            const payload = snapshot.val();
            if (payload && payload.rows && payload.columns) {
              const objRows = rowsFromPayload(payload);
              setRows(objRows);
              setMeta({ fileName: payload.fileName, updatedAt: payload.updatedAt, rowCount: objRows.length });
            }
            setSyncStatus("cloud");
            setBootLoading(false);
          },
          (err) => {
            console.error("Error de sincronización con la nube:", err);
            setSyncStatus("error");
            setBootLoading(false);
          }
        );
      }
    })();

    return () => { if (cloudUnsub) cloudUnsub(); };
  }, []);

  const persist = useCallback(async (normalizedRows, fileName) => {
    const columns = COLS;
    const dataRows = normalizedRows.map((r) => columns.map((c) => encodeCell(c, r[c])));
    const payload = { columns, rows: dataRows, fileName, updatedAt: new Date().toISOString() };

    try {
      await window.storage.set(storageKey, JSON.stringify(payload), false);
    } catch (e) {
      console.error("No se pudo guardar en caché local:", e);
    }

    if (CLOUD_SYNC_ENABLED && cloudDb) {
      try {
        await dbSet(ref(cloudDb, cloudPath), payload);
        setSyncStatus("cloud");
      } catch (e) {
        console.error("No se pudo sincronizar con la nube:", e);
        setSyncStatus("error");
      }
    }
  }, []);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setParsing(true);
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      let sheetName = wb.SheetNames[0];
      let best = 0;
      wb.SheetNames.forEach((n) => {
        const ws = wb.Sheets[n];
        const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
        const rc = range ? range.e.r - range.s.r : 0;
        if (rc > best) { best = rc; sheetName = n; }
      });
      const ws = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
      if (!json.length) throw new Error("La hoja no tiene filas de datos.");
      const normalized = json.map(normalizeRow).filter((r) => r["Campo"] || r["Cultivo"] || r["Concepto"]);
      setRows(normalized);
      const newMeta = { fileName: file.name, updatedAt: new Date().toISOString(), rowCount: normalized.length };
      setMeta(newMeta);
      setPage(1);
      await persist(normalized, file.name);

      // Si los filtros activos dejan el dashboard vacío con los datos nuevos, los limpiamos
      const q = search.trim().toLowerCase();
      const wouldMatch = normalized.some((r) => {
        if (cultivo !== "Todos" && r["Cultivo"] !== cultivo) return false;
        if (campo !== "Todos" && r["Campo"] !== campo) return false;
        if (administracion !== "Todos" && r["Admin"] !== administracion) return false;
        if (lote !== "Todos") {
          const [lc, ll] = lote.split("::");
          if (r["Campo"] !== lc || r["Lote"] !== ll) return false;
        }
        if (tipoDet !== "Todos" && r["Tipo Det."] !== tipoDet) return false;
        if (tipoItem !== "Todos" && r["Tipo item"] !== tipoItem) return false;
        if (desde && r["Fecha"] && r["Fecha"] < desde) return false;
        if (hasta && r["Fecha"] && r["Fecha"] > hasta) return false;
        if (q) {
          const hay = `${r["Concepto"]} ${r["Origen"]} ${r["Lote"]} ${r["OTA"]} ${r["Variedad"]} ${r["Tipo item"]}`.toLowerCase();
          if (!hay.includes(q)) return false;
        }
        return true;
      });
      if (!wouldMatch) {
        setCultivo("Todos"); setCampo("Todos"); setLote("Todos"); setAdministracion("Todos"); setTipoDet("Todos");
        setTipoItem("Todos"); setDesde(""); setHasta(""); setSearch("");
      }
    } catch (e) {
      setError("No pude leer el archivo. Verificá que sea un Excel exportado del mismo formato (" + e.message + ")");
    } finally {
      setParsing(false);
    }
  }, [persist, cultivo, campo, lote, tipoDet, tipoItem, desde, hasta, search]);

  const onDrop = (e) => {
    e.preventDefault(); setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) handleFile(f);
  };

  const clearDataset = async () => {
    setRows([]); setMeta(null);
    try { await window.storage.delete(storageKey, false); } catch (e) {}
    if (CLOUD_SYNC_ENABLED && cloudDb) {
      try { await dbSet(ref(cloudDb, cloudPath), null); } catch (e) {}
    }
  };

  // Listas para filtros
  const cultivos = useMemo(() => Array.from(new Set(rows.map((r) => r["Cultivo"]).filter(Boolean))).sort(), [rows]);
  const administraciones = useMemo(() => Array.from(new Set(rows.map((r) => r["Admin"]).filter(Boolean))).sort(), [rows]);
  // Los campos dependen de la administración elegida, para no listar campos que no le pertenecen
  const campos = useMemo(() => {
    const base = administracion === "Todos" ? rows : rows.filter((r) => r["Admin"] === administracion);
    return Array.from(new Set(base.map((r) => r["Campo"]).filter(Boolean))).sort();
  }, [rows, administracion]);
  // Si el campo elegido deja de pertenecer a la administración elegida, lo reseteamos
  useEffect(() => {
    if (campo !== "Todos" && !campos.includes(campo)) setCampo("Todos");
  }, [campos]); // eslint-disable-line react-hooks/exhaustive-deps
  const tipoItems = useMemo(() => Array.from(new Set(rows.map((r) => r["Tipo item"]).filter(Boolean))).sort(), [rows]);
  // Los lotes dependen de la administración, el campo y el cultivo elegidos
  const loteOptions = useMemo(() => {
    const base = rows.filter((r) =>
      (administracion === "Todos" || r["Admin"] === administracion) &&
      (campo === "Todos" || r["Campo"] === campo) &&
      (cultivo === "Todos" || r["Cultivo"] === cultivo)
    );
    const map = new Map();
    base.forEach((r) => {
      if (!r["Lote"]) return;
      const value = `${r["Campo"]}::${r["Lote"]}`;
      if (!map.has(value)) map.set(value, campo === "Todos" ? `${r["Campo"]} · Lote ${r["Lote"]}` : `Lote ${r["Lote"]}`);
    });
    return Array.from(map.entries()).map(([value, label]) => ({ value, label })).sort((a, b) => a.label.localeCompare(b.label, "es", { numeric: true }));
  }, [rows, administracion, campo, cultivo]);
  // Si el lote elegido deja de tener sentido con el nuevo campo/cultivo/administración, lo reseteamos
  useEffect(() => {
    if (lote !== "Todos" && !loteOptions.some((o) => o.value === lote)) setLote("Todos");
  }, [loteOptions]); // eslint-disable-line react-hooks/exhaustive-deps

  // Filtrado
  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return rows.filter((r) => {
      if (cultivo !== "Todos" && r["Cultivo"] !== cultivo) return false;
      if (campo !== "Todos" && r["Campo"] !== campo) return false;
      if (administracion !== "Todos" && r["Admin"] !== administracion) return false;
      if (lote !== "Todos") {
        const [lc, ll] = lote.split("::");
        if (r["Campo"] !== lc || r["Lote"] !== ll) return false;
      }
      if (tipoDet !== "Todos" && r["Tipo Det."] !== tipoDet) return false;
      if (tipoItem !== "Todos" && r["Tipo item"] !== tipoItem) return false;
      if (desde && r["Fecha"] && r["Fecha"] < desde) return false;
      if (hasta && r["Fecha"] && r["Fecha"] > hasta) return false;
      if (q) {
        const hay = `${r["Concepto"]} ${r["Origen"]} ${r["Lote"]} ${r["OTA"]} ${r["Variedad"]} ${r["Tipo item"]}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [rows, cultivo, campo, administracion, lote, tipoDet, tipoItem, desde, hasta, search]);

  useEffect(() => { setPage(1); }, [cultivo, campo, administracion, lote, tipoDet, tipoItem, desde, hasta, search]);

  // KPIs
  const kpis = useMemo(() => {
    let gasto = 0, sinPrecio = 0;
    const supMap = new Map();
    const lotesSet = new Set(), camposSet = new Set(), cultivosSet = new Set();
    filtered.forEach((r) => {
      if (r["U$S/Total"] === null) sinPrecio++; else gasto += r["U$S/Total"];
      const key = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}`;
      if (!supMap.has(key)) supMap.set(key, r["Sup.Cultivo"] || 0);
      if (r["Lote"]) lotesSet.add(`${r["Campo"]}|${r["Lote"]}`);
      if (r["Campo"]) camposSet.add(r["Campo"]);
      if (r["Cultivo"]) cultivosSet.add(r["Cultivo"]);
    });
    const superficie = Array.from(supMap.values()).reduce((a, b) => a + b, 0);
    return {
      gasto, sinPrecio, superficie,
      lotes: lotesSet.size, campos: camposSet.size, cultivos: cultivosSet.size,
      costoPorHa: superficie > 0 ? gasto / superficie : 0,
    };
  }, [filtered]);

  // Datos para gráficos: todos normalizados por hectárea y desglosados por cultivo
  // (deduplicando superficie por Campo+Lote+Cultivo para no sumar la misma ha varias veces)
  const porCultivo = useMemo(() => {
    const m = new Map();
    const seenSup = new Set();
    filtered.forEach((r) => {
      if (!r["Cultivo"]) return;
      if (!m.has(r["Cultivo"])) m.set(r["Cultivo"], { gasto: 0, ha: 0, insumos: 0, servicios: 0 });
      const e = m.get(r["Cultivo"]);
      const v = r["U$S/Total"] || 0;
      e.gasto += v;
      if (r["Tipo Det."] === "INSUMOS") e.insumos += v; else if (r["Tipo Det."] === "SERVICIOS") e.servicios += v;
      const supKey = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}`;
      if (!seenSup.has(supKey)) { seenSup.add(supKey); e.ha += r["Sup.Cultivo"] || 0; }
    });
    return Array.from(m.entries()).map(([name, e]) => ({
      name, value: e.gasto, ha: e.ha,
      costoHa: e.ha > 0 ? e.gasto / e.ha : 0,
      insumosHa: e.ha > 0 ? e.insumos / e.ha : 0,
      serviciosHa: e.ha > 0 ? e.servicios / e.ha : 0,
    })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  const costoHaPorCultivo = useMemo(() => [...porCultivo].filter((e) => e.ha > 0).sort((a, b) => b.costoHa - a.costoHa), [porCultivo]);
  const cultivosConHa = useMemo(() => costoHaPorCultivo.map((e) => e.name), [costoHaPorCultivo]);
  const haPorCultivoMap = useMemo(() => new Map(porCultivo.map((e) => [e.name, e.ha])), [porCultivo]);

  // Principales rubros de gasto por hectárea, agrupados por cultivo
  const rubrosPorCultivo = useMemo(() => {
    const m = new Map(); // tipoItem -> { cultivo: gasto }
    filtered.forEach((r) => {
      if (!r["Tipo item"] || !r["Cultivo"]) return;
      if (!m.has(r["Tipo item"])) m.set(r["Tipo item"], {});
      const bucket = m.get(r["Tipo item"]);
      bucket[r["Cultivo"]] = (bucket[r["Cultivo"]] || 0) + (r["U$S/Total"] || 0);
    });
    const rows = Array.from(m.entries()).map(([name, bucket]) => {
      const row = { name };
      let rank = 0;
      cultivosConHa.forEach((c) => {
        const ha = haPorCultivoMap.get(c) || 0;
        row[c] = ha > 0 ? (bucket[c] || 0) / ha : 0;
        rank += row[c];
      });
      row._rank = rank;
      return row;
    }).sort((a, b) => b._rank - a._rank);
    return rows.slice(0, 8);
  }, [filtered, cultivosConHa, haPorCultivoMap]);

  // Control de datos: detecta filas con problemas que podrían ensuciar el dashboard
  // (se calcula sobre TODO el dataset, no solo lo filtrado, para dar una foto completa
  // de lo que conviene corregir en el Excel de origen).
  const [showDataQuality, setShowDataQuality] = useState(false);
  // Precios y consumo por ítem (Rubro + Concepto), respetando los filtros activos
  // Agrupa en dos secciones: Insumos primero, Servicios después.
  const [itemSort, setItemSort] = useState({ key: "costo", dir: "desc" });
  const toggleItemSort = (key) => setItemSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "desc" }));
  const TIPODET_ORDER_ITEMS = { "INSUMOS": 0, "SERVICIOS": 1 };
  const itemsSummary = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const concepto = r["Concepto"] || "(sin concepto)";
      const key = `${r["Tipo item"] || "(sin rubro)"}|${concepto}`;
      if (!m.has(key)) m.set(key, { rubro: r["Tipo item"] || "(sin rubro)", concepto, tipoDet: r["Tipo Det."] || "", unid: r["Unid."], cant: 0, costo: 0, sumUnit: 0, countUnit: 0, movimientos: 0 });
      const e = m.get(key);
      e.movimientos += 1;
      e.cant += r["Cantidad"] || 0;
      e.costo += r["U$S/Total"] || 0;
      if (r["U$S/U"] !== null) { e.sumUnit += r["U$S/U"]; e.countUnit += 1; }
    });
    return Array.from(m.values()).map((e) => ({
      rubro: e.rubro, concepto: e.concepto, tipoDet: e.tipoDet, unid: e.unid, movimientos: e.movimientos,
      cant: e.cant, costo: e.costo, precioProm: e.countUnit > 0 ? e.sumUnit / e.countUnit : null,
    }));
  }, [filtered]);
  const itemsSorted = useMemo(() => {
    const arr = [...itemsSummary];
    const { key, dir } = itemSort;
    arr.sort((a, b) => {
      // Siempre Insumos antes que Servicios, sin importar la columna elegida
      const da = TIPODET_ORDER_ITEMS[a.tipoDet] ?? 2;
      const db = TIPODET_ORDER_ITEMS[b.tipoDet] ?? 2;
      if (da !== db) return da - db;
      // Dentro de cada grupo, ordenar por la columna elegida
      let va = a[key], vb = b[key];
      if (va === null) va = dir === "asc" ? Infinity : -Infinity;
      if (vb === null) vb = dir === "asc" ? Infinity : -Infinity;
      if (typeof va === "string") { va = va.toLowerCase(); vb = String(vb).toLowerCase(); }
      if (va < vb) return dir === "asc" ? -1 : 1;
      if (va > vb) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return arr;
  }, [itemsSummary, itemSort]);

  // Exporta a Excel los registros filtrados, con las mismas 21 columnas del archivo original

  const dataQuality = useMemo(() => {
    const totals = { sinPrecio: 0, precioCero: 0, sinSuperficie: 0, sinCultivo: 0, sinCampo: 0, sinConcepto: 0, filasConProblemas: 0 };
    const groups = new Map();
    rows.forEach((r) => {
      const problems = [];
      if (r["U$S/Total"] === null) { problems.push("Sin precio"); totals.sinPrecio++; }
      else if (r["U$S/Total"] === 0 && r["Cantidad"] > 0) { problems.push("Precio en $0"); totals.precioCero++; }
      if (!r["Sup.Cultivo"]) { problems.push("Sin superficie"); totals.sinSuperficie++; }
      if (!r["Cultivo"]) { problems.push("Sin cultivo"); totals.sinCultivo++; }
      if (!r["Campo"]) { problems.push("Sin campo"); totals.sinCampo++; }
      if (!r["Concepto"]) { problems.push("Sin concepto"); totals.sinConcepto++; }
      if (!problems.length) return;
      totals.filasConProblemas++;
      const key = `${r["Tipo item"] || "(sin rubro)"}|${r["Concepto"] || "(sin concepto)"}`;
      if (!groups.has(key)) groups.set(key, { rubro: r["Tipo item"] || "(sin rubro)", concepto: r["Concepto"] || "(sin concepto)", unid: r["Unid."], count: 0, cant: 0, problemsSet: new Set() });
      const g = groups.get(key);
      g.count += 1;
      g.cant += r["Cantidad"] || 0;
      problems.forEach((p) => g.problemsSet.add(p));
    });
    const list = Array.from(groups.values()).map((g) => ({ ...g, problems: Array.from(g.problemsSet) })).sort((a, b) => b.count - a.count);
    return { totals, list };
  }, [rows]);

  // Comparativo por lotes (agrupa Campo + Lote + Cultivo dentro de lo ya filtrado)
  const loteAgg = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const key = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}`;
      if (!m.has(key)) {
        m.set(key, { campo: r["Campo"], lote: r["Lote"], cultivo: r["Cultivo"], variedad: r["Variedad"], ha: r["Sup.Cultivo"] || 0, insumos: 0, servicios: 0, total: 0 });
      }
      const e = m.get(key);
      const v = r["U$S/Total"] || 0;
      if (r["Tipo Det."] === "INSUMOS") e.insumos += v; else if (r["Tipo Det."] === "SERVICIOS") e.servicios += v;
      e.total += v;
    });
    return Array.from(m.values()).map((e) => ({ ...e, costoHa: e.ha > 0 ? e.total / e.ha : 0 }));
  }, [filtered]);

  const [loteSort, setLoteSort] = useState({ key: "total", dir: "desc" });
  const toggleLoteSort = (key) => setLoteSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "desc" }));
  const loteSorted = useMemo(() => {
    const arr = [...loteAgg];
    const { key, dir } = loteSort;
    arr.sort((a, b) => {
      let va = a[key], vb = b[key];
      if (typeof va === "string") { va = va.toLowerCase(); vb = String(vb).toLowerCase(); }
      if (va < vb) return dir === "asc" ? -1 : 1;
      if (va > vb) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return arr;
  }, [loteAgg, loteSort]);
  const maxCostoHa = Math.max(1, ...loteAgg.map((e) => e.costoHa));
  const topLotesPorGasto = useMemo(() => [...loteAgg].filter((e) => e.ha > 0).sort((a, b) => b.costoHa - a.costoHa).slice(0, 12).map((e) => ({
    name: `${e.campo} · L${e.lote} · ${e.cultivo}`,
    value: e.costoHa, campo: e.campo, lote: e.lote, cultivo: e.cultivo, variedad: e.variedad, ha: e.ha,
  })), [loteAgg]);

  // Detalle de un lote seleccionado
  const [selectedLoteKey, setSelectedLoteKey] = useState(null);
  const selectLote = (key) => setSelectedLoteKey((cur) => (cur === key ? null : key));
  const selectedLoteInfo = useMemo(() => loteAgg.find((e) => `${e.campo}|${e.lote}|${e.cultivo}|${e.variedad}` === selectedLoteKey) || null, [loteAgg, selectedLoteKey]);
  const selectedLoteRows = useMemo(() => {
    if (!selectedLoteKey) return [];
    return filtered.filter((r) => `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}` === selectedLoteKey)
      .slice().sort((a, b) => {
        const fa = a["Fecha"] || "", fb = b["Fecha"] || "";
        if (fa !== fb) return fa < fb ? -1 : 1;
        const ra = a["Tipo item"] || "", rb = b["Tipo item"] || "";
        if (ra !== rb) return ra.localeCompare(rb, "es");
        const ca = a["Concepto"] || "", cb = b["Concepto"] || "";
        return ca.localeCompare(cb, "es");
      });
  }, [filtered, selectedLoteKey]);
  const TIPODET_ORDER = { "INSUMOS": 0, "SERVICIOS": 1 };
  const selectedLoteByItem = useMemo(() => {
    const m = new Map();
    selectedLoteRows.forEach((r) => {
      if (!r["Tipo item"]) return;
      const key = r["Tipo item"];
      if (!m.has(key)) m.set(key, { value: 0, tipoDet: r["Tipo Det."] || "" });
      m.get(key).value += r["U$S/Total"] || 0;
    });
    return Array.from(m.entries())
      .map(([name, v]) => ({ name, value: v.value, tipoDet: v.tipoDet }))
      .sort((a, b) => (TIPODET_ORDER[a.tipoDet] ?? 2) - (TIPODET_ORDER[b.tipoDet] ?? 2) || b.value - a.value);
  }, [selectedLoteRows]);

  // Resumen agrupado por Tipo Det. (Insumos primero, Servicios después) > Rubro > Concepto
  const [loteDetailTab, setLoteDetailTab] = useState("resumen");
  useEffect(() => { setLoteDetailTab("resumen"); }, [selectedLoteKey]);
  const selectedLoteSummary = useMemo(() => {
    const detMap = new Map();
    selectedLoteRows.forEach((r) => {
      const tipoDet = r["Tipo Det."] || "(sin tipo)";
      if (!detMap.has(tipoDet)) detMap.set(tipoDet, { items: 0, cant: 0, costo: 0, rubros: new Map() });
      const dg = detMap.get(tipoDet);
      dg.items += 1; dg.cant += r["Cantidad"] || 0; dg.costo += r["U$S/Total"] || 0;
      const rubro = r["Tipo item"] || "(sin rubro)";
      if (!dg.rubros.has(rubro)) dg.rubros.set(rubro, { items: 0, cant: 0, costo: 0, conceptos: new Map() });
      const rg = dg.rubros.get(rubro);
      rg.items += 1; rg.cant += r["Cantidad"] || 0; rg.costo += r["U$S/Total"] || 0;
      const concepto = r["Concepto"] || "(sin concepto)";
      if (!rg.conceptos.has(concepto)) rg.conceptos.set(concepto, { items: 0, cant: 0, costo: 0, unid: r["Unid."], sumUnit: 0, countUnit: 0 });
      const cg = rg.conceptos.get(concepto);
      cg.items += 1; cg.cant += r["Cantidad"] || 0; cg.costo += r["U$S/Total"] || 0;
      if (r["U$S/U"] !== null) { cg.sumUnit += r["U$S/U"]; cg.countUnit += 1; }
    });
    const ha = selectedLoteInfo ? selectedLoteInfo.ha : 0;
    return Array.from(detMap.entries())
      .map(([tipoDet, dg]) => ({
        tipoDet, items: dg.items, cant: dg.cant, costo: dg.costo, costoHa: ha > 0 ? dg.costo / ha : null,
        rubros: Array.from(dg.rubros.entries()).map(([rubro, rg]) => ({
          rubro, items: rg.items, cant: rg.cant, costo: rg.costo, costoHa: ha > 0 ? rg.costo / ha : null,
          conceptos: Array.from(rg.conceptos.entries()).map(([concepto, cg]) => ({
            concepto, unid: cg.unid, items: cg.items, cant: cg.cant, costo: cg.costo,
            unitPrice: cg.countUnit > 0 ? cg.sumUnit / cg.countUnit : null,
            dosisHa: ha > 0 ? cg.cant / ha : null,
            costoHa: ha > 0 ? cg.costo / ha : null,
          })).sort((a, b) => b.costo - a.costo),
        })).sort((a, b) => b.costo - a.costo),
      }))
      .sort((a, b) => (TIPODET_ORDER[a.tipoDet] ?? 2) - (TIPODET_ORDER[b.tipoDet] ?? 2));
  }, [selectedLoteRows, selectedLoteInfo]);
  // Si cambian los filtros generales y el lote seleccionado deja de existir, lo deseleccionamos
  useEffect(() => {
    if (selectedLoteKey && !loteAgg.some((e) => `${e.campo}|${e.lote}|${e.cultivo}|${e.variedad}` === selectedLoteKey)) setSelectedLoteKey(null);
  }, [loteAgg, selectedLoteKey]);

  // Orden y paginación de tabla
  const sorted = useMemo(() => {
    const arr = [...filtered];
    const { key, dir } = sort;
    arr.sort((a, b) => {
      let va = a[key], vb = b[key];
      if (va === null || va === undefined) va = dir === "asc" ? Infinity : -Infinity;
      if (vb === null || vb === undefined) vb = dir === "asc" ? Infinity : -Infinity;
      if (typeof va === "string") { va = va.toLowerCase(); vb = String(vb).toLowerCase(); }
      if (va < vb) return dir === "asc" ? -1 : 1;
      if (va > vb) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return arr;
  }, [filtered, sort]);

  // Exporta a Excel los registros filtrados, con las mismas 21 columnas del archivo original
  // Columnas y orden para la exportación a Excel: distinto del orden interno (COLS).
  // Excluye "Cant.Orig." y "U.Serv.", y mueve/renombra "Cant.Afect" -> "Sup.Origen"
  // justo después de "Origen".
  const EXPORT_COLUMNS = [
    ["Admin", "Admin"], ["Campaña", "Campaña"], ["Campo", "Campo"], ["Lote", "Lote"], ["Sup.Lote", "Sup.Lote"],
    ["Cultivo", "Cultivo"], ["Variedad", "Variedad"], ["Sup.Cultivo", "Sup.Cultivo"], ["Fecha", "Fecha"], ["OTA", "OTA"],
    ["Origen", "Origen"], ["Cant.Afect", "Sup.Origen"], ["Tipo Det.", "Tipo Det."], ["Tipo item", "Tipo item"],
    ["Concepto", "Concepto"], ["Cantidad", "Cantidad"], ["Unid.", "Unid."], ["U$S/U", "U$S/U"], ["U$S/Total", "U$S/Total"],
  ];
  const exportToExcel = useCallback(() => {
    const data = sorted.map((r) => {
      const o = {};
      EXPORT_COLUMNS.forEach(([srcKey, header]) => {
        let v = r[srcKey];
        if (srcKey === "Fecha") v = fmtDate(v);
        if (v === null) v = "";
        o[header] = v;
      });
      return o;
    });
    const ws = XLSX.utils.json_to_sheet(data, { header: EXPORT_COLUMNS.map(([, h]) => h) });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Movimientos");
    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `movimientos_filtrados_${fecha}.xlsx`);
  }, [sorted]);

  const totalPages = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
  const pageRows = sorted.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  const toggleSort = (key) => {
    setSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "asc" }));
  };

  const resetFiltros = () => { setCultivo("Todos"); setCampo("Todos"); setLote("Todos"); setAdministracion("Todos"); setTipoDet("Todos"); setTipoItem("Todos"); setDesde(""); setHasta(""); setSearch(""); };
  const hayFiltrosActivos = cultivo !== "Todos" || campo !== "Todos" || lote !== "Todos" || administracion !== "Todos" || tipoDet !== "Todos" || tipoItem !== "Todos" || desde || hasta || search;

  // -------------------------------------------------------------------------
  return (
    <>
      {/* Sub-header de costos: datos y controles de carga */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
        <div className="agri-sub">
          {meta ? (
            <>Datos de <strong>{meta.fileName}</strong> · {fmtNum(meta.rowCount)} registros · actualizado {new Date(meta.updatedAt).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" })}</>
          ) : isAdmin ? "Cargá tu planilla para empezar a consultar los datos" : "Todavía no hay datos cargados para este cliente."}
        </div>
        {CLOUD_SYNC_ENABLED && (
          <div style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 600, padding: "3px 9px", borderRadius: 999, background: syncStatus === "cloud" ? "rgba(75,107,58,0.12)" : syncStatus === "error" ? "rgba(161,70,47,0.12)" : "rgba(107,94,79,0.12)", color: syncStatus === "cloud" ? "var(--green)" : syncStatus === "error" ? "var(--rust)" : "var(--ink-soft)" }}>
            {syncStatus === "cloud" ? <><Cloud size={12} /> Sincronizado</> : syncStatus === "error" ? <><CloudOff size={12} /> Sin conexión</> : <><Cloud size={12} /> Conectando…</>}
          </div>
        )}
        {isAdmin && (
          <div style={{ display: "flex", gap: 8 }}>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
              onChange={(e) => handleFile(e.target.files?.[0])} />
            {meta && (
              <>
                <button className="agri-btn" onClick={() => fileInputRef.current?.click()}>
                  <RefreshCw size={14} /> Actualizar datos
                </button>
                <button className="agri-btn agri-btn-outline" onClick={clearDataset}>
                  <Trash2 size={14} />
                </button>
              </>
            )}
            <button className="agri-btn agri-btn-outline" onClick={() => signOut(cloudAuth)} title="Cerrar sesión">
              <LogOut size={14} />
            </button>
          </div>
        )}
      </div>

      {bootLoading ? (
        <div style={{ padding: 60, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>
        ) : !meta || rows.length === 0 ? (
          isAdmin ? (
          <div
            className="agri-dropzone" data-drag={dragOver}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
          >
            <FileSpreadsheet size={40} color="var(--ink-soft)" />
            <div className="agri-serif" style={{ fontSize: 16, fontWeight: 700 }}>
              {parsing ? "Leyendo el archivo…" : "Arrastrá tu Excel acá"}
            </div>
            <div style={{ fontSize: 13, color: "var(--ink-soft)", maxWidth: 360 }}>
              Formato .xlsx con las columnas de campo, lote, cultivo, insumos y costos. Los datos se guardan en este dispositivo hasta que subas una versión nueva.
            </div>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
              onChange={(e) => handleFile(e.target.files?.[0])} />
            <button className="agri-btn" disabled={parsing} onClick={() => fileInputRef.current?.click()}>
              <Upload size={14} /> {parsing ? "Procesando…" : "Elegir archivo"}
            </button>
            {error && (
              <div style={{ color: "var(--rust)", fontSize: 13, display: "flex", alignItems: "center", gap: 6, marginTop: 6 }}>
                <AlertCircle size={14} /> {error}
              </div>
            )}
          </div>
          ) : (
            <div className="agri-card" style={{ padding: 56, textAlign: "center" }}>
              <FileSpreadsheet size={36} color="var(--ink-soft)" style={{ marginBottom: 10 }} />
              <div className="agri-serif" style={{ fontSize: 16, fontWeight: 700, marginBottom: 6 }}>Todavía no hay datos cargados</div>
              <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Contactá a tu administrador para que suba la planilla de esta campaña.</div>
            </div>
          )
        ) : (
          <>
            {/* Filtros */}
            <div className="agri-card" style={{ padding: 14, marginBottom: 18 }}>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", marginBottom: 12 }}>
                <button className="agri-chip" data-active={cultivo === "Todos"} onClick={() => setCultivo("Todos")}>Todos los cultivos</button>
                {cultivos.map((c) => (
                  <button key={c} className="agri-chip" data-active={cultivo === c} onClick={() => setCultivo(c)}>{c}</button>
                ))}
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
                <div style={{ position: "relative", flex: "1 1 220px" }}>
                  <Search size={14} style={{ position: "absolute", left: 9, top: 10, color: "var(--ink-soft)" }} />
                  <input className="agri-input" style={{ width: "100%", paddingLeft: 30 }}
                    placeholder="Buscar concepto, lote, labor…" value={search} onChange={(e) => setSearch(e.target.value)} />
                </div>
                <select className="agri-select" value={administracion} onChange={(e) => setAdministracion(e.target.value)}>
                  <option value="Todos">Todas las administraciones</option>
                  {administraciones.map((a) => <option key={a} value={a}>{a}</option>)}
                </select>
                <select className="agri-select" value={campo} onChange={(e) => setCampo(e.target.value)}>
                  <option value="Todos">Todos los campos</option>
                  {campos.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
                <select className="agri-select" value={lote} onChange={(e) => setLote(e.target.value)}>
                  <option value="Todos">Todos los lotes</option>
                  {loteOptions.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                </select>
                <select className="agri-select" value={tipoDet} onChange={(e) => setTipoDet(e.target.value)}>
                  <option value="Todos">Insumos y servicios</option>
                  <option value="INSUMOS">Solo insumos</option>
                  <option value="SERVICIOS">Solo servicios</option>
                </select>
                <select className="agri-select" value={tipoItem} onChange={(e) => setTipoItem(e.target.value)}>
                  <option value="Todos">Todos los rubros</option>
                  {tipoItems.map((t) => <option key={t} value={t}>{t}</option>)}
                </select>
                <input className="agri-input" type="date" value={desde} onChange={(e) => setDesde(e.target.value)} />
                <span style={{ color: "var(--ink-soft)", fontSize: 13 }}>a</span>
                <input className="agri-input" type="date" value={hasta} onChange={(e) => setHasta(e.target.value)} />
                {hayFiltrosActivos && (
                  <button className="agri-btn agri-btn-outline" onClick={resetFiltros}><X size={13} /> Limpiar</button>
                )}
              </div>
            </div>

            {/* KPIs */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 20 }}>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label"><DollarSign size={12} style={{ display: "inline", marginBottom: -1 }} /> Gasto total</div>
                <div className="agri-kpi-value">{fmtUSD(kpis.gasto)}</div>
                {kpis.sinPrecio > 0 && <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>{kpis.sinPrecio} sin precio cargado</div>}
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label"><Ruler size={12} style={{ display: "inline", marginBottom: -1 }} /> Superficie</div>
                <div className="agri-kpi-value">{fmtNum(kpis.superficie)} ha</div>
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Costo por hectárea</div>
                <div className="agri-kpi-value">{fmtUSD2(kpis.costoPorHa)}</div>
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label"><Sprout size={12} style={{ display: "inline", marginBottom: -1 }} /> Cultivos</div>
                <div className="agri-kpi-value">{kpis.cultivos}</div>
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label"><MapPin size={12} style={{ display: "inline", marginBottom: -1 }} /> Campos · Lotes</div>
                <div className="agri-kpi-value">{kpis.campos} · {kpis.lotes}</div>
              </div>
            </div>

            {/* Gráficos — todos normalizados por hectárea y desglosados por cultivo */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 14, marginBottom: 20 }}>
              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Costo por hectárea · por cultivo</div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={costoHaPorCultivo} margin={{ left: -10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                    <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v) => fmtUSD2(v) + "/ha"} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Bar dataKey="costoHa" radius={[3, 3, 0, 0]}>
                      {costoHaPorCultivo.map((e, i) => <Cell key={i} fill={cultivoColor(e.name, i)} />)}
                      <LabelList dataKey="costoHa" content={<VerticalBarLabel />} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Composición del costo/ha · por cultivo</div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={costoHaPorCultivo} margin={{ left: -10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                    <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v, name) => [fmtUSD2(v) + "/ha", name === "insumosHa" ? "Insumos" : "Servicios"]} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Legend wrapperStyle={{ fontSize: 11 }} formatter={(v) => (v === "insumosHa" ? "Insumos" : "Servicios")} />
                    <Bar dataKey="insumosHa" stackId="costo" fill="#4B6B3A" radius={[0, 0, 0, 0]}>
                      <LabelList dataKey="insumosHa" content={<VerticalBarLabel />} />
                    </Bar>
                    <Bar dataKey="serviciosHa" stackId="costo" fill="#B8842E" radius={[3, 3, 0, 0]}>
                      <LabelList dataKey="serviciosHa" content={<VerticalBarLabel />} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Principales rubros por hectárea · por cultivo</div>
                <ResponsiveContainer width="100%" height={Math.max(220, rubrosPorCultivo.length * 30)}>
                  <BarChart data={rubrosPorCultivo} layout="vertical" margin={{ left: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 10, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v) => fmtUSD2(v) + "/ha"} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    {cultivosConHa.map((c, i) => (
                      <Bar key={c} dataKey={c} name={c} fill={cultivoColor(c, i)} radius={[0, 3, 3, 0]}>
                        <LabelList dataKey={c} content={<HorizontalBarLabel />} />
                      </Bar>
                    ))}
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Comparativo por lotes */}
            <div className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Lotes con mayor gasto por hectárea</div>
              <ResponsiveContainer width="100%" height={Math.max(180, topLotesPorGasto.length * 26)}>
                <BarChart data={topLotesPorGasto} layout="vertical" margin={{ left: 10, right: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" width={230} tick={{ fontSize: 11, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                  <Tooltip content={<LoteCostoHaTooltip />} />
                  <Bar dataKey="value" fill="#2F5B66" radius={[0, 3, 3, 0]}>
                    <LabelList dataKey="value" content={<HorizontalBarLabel />} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="agri-card" style={{ padding: 0, overflow: "hidden", marginBottom: 20 }}>
              <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Comparativo por lote</div>
                <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(loteSorted.length)} lotes · tocá una fila para ver el detalle</div>
              </div>
              <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 780 }}>
                  <thead>
                    <tr>
                      {[["campo", "Campo"], ["lote", "Lote"], ["cultivo", "Cultivo"], ["variedad", "Variedad"], ["ha", "Ha"], ["insumos", "Insumos"], ["servicios", "Servicios"], ["total", "Total"], ["costoHa", "USD/ha"]].map(([key, label]) => (
                        <th key={key} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)" }} onClick={() => toggleLoteSort(key)}>
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 3 }}>
                            {label} {loteSort.key === key && <ArrowUpDown size={11} />}
                          </span>
                        </th>
                      ))}
                      <th className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)" }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {loteSorted.map((e, i) => {
                      const key = `${e.campo}|${e.lote}|${e.cultivo}|${e.variedad}`;
                      const isSelected = key === selectedLoteKey;
                      return (
                      <tr key={i} className="agri-tr" onClick={() => selectLote(key)} style={{ cursor: "pointer", background: isSelected ? "rgba(184,132,46,0.14)" : undefined }}>
                        <td className="agri-td">{e.campo}</td>
                        <td className="agri-td">{e.lote}</td>
                        <td className="agri-td">{e.cultivo}</td>
                        <td className="agri-td" style={{ color: "var(--ink-soft)" }}>{e.variedad}</td>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(e.ha, 1)}</td>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtUSD(e.insumos)}</td>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtUSD(e.servicios)}</td>
                        <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{fmtUSD(e.total)}</td>
                        <td className="agri-td">
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <div style={{ flex: "0 0 58px", textAlign: "right", fontWeight: 600 }}>{fmtUSD2(e.costoHa)}</div>
                            <div style={{ flex: 1, minWidth: 40, height: 6, background: "var(--line)", borderRadius: 3, overflow: "hidden" }}>
                              <div style={{ width: `${Math.min(100, (e.costoHa / maxCostoHa) * 100)}%`, height: "100%", background: "var(--gold)" }} />
                            </div>
                          </div>
                        </td>
                        <td className="agri-td" style={{ textAlign: "center" }}>
                          <Eye size={15} color={isSelected ? "var(--gold)" : "var(--ink-soft)"} />
                        </td>
                      </tr>
                      );
                    })}
                    {loteSorted.length === 0 && (
                      <tr><td className="agri-td" colSpan={10} style={{ textAlign: "center", padding: 30, color: "var(--ink-soft)" }}>Ningún lote coincide con los filtros aplicados.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Detalle del lote seleccionado */}
            {selectedLoteInfo && (
              <div className="agri-card" style={{ padding: 16, marginBottom: 20, borderColor: "var(--gold)" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
                  <div>
                    <div className="agri-serif" style={{ fontSize: 15, fontWeight: 700 }}>
                      {selectedLoteInfo.campo} · Lote {selectedLoteInfo.lote} · {selectedLoteInfo.cultivo}
                    </div>
                    <div style={{ fontSize: 13, color: "var(--ink-soft)", marginTop: 2 }}>
                      {selectedLoteInfo.variedad ? `${selectedLoteInfo.variedad} · ` : ""}{fmtNum(selectedLoteInfo.ha, 1)} ha · {fmtNum(selectedLoteRows.length)} movimientos
                    </div>
                  </div>
                  <button className="agri-btn agri-btn-outline" onClick={() => setSelectedLoteKey(null)}><X size={13} /> Cerrar</button>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 10, marginBottom: 16 }}>
                  <div>
                    <div className="agri-kpi-label">Costo total</div>
                    <div className="agri-kpi-value" style={{ fontSize: 15 }}>{fmtUSD(selectedLoteInfo.total)}</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Insumos</div>
                    <div className="agri-kpi-value" style={{ fontSize: 15 }}>{fmtUSD(selectedLoteInfo.insumos)}</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Servicios</div>
                    <div className="agri-kpi-value" style={{ fontSize: 15 }}>{fmtUSD(selectedLoteInfo.servicios)}</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Costo por hectárea</div>
                    <div className="agri-kpi-value" style={{ fontSize: 15 }}>{fmtUSD2(selectedLoteInfo.costoHa)}</div>
                  </div>
                </div>

                <div className="agri-detail-grid" style={{ display: "grid", gridTemplateColumns: "minmax(190px, 0.8fr) minmax(320px, 1.8fr)", gap: 16 }}>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Composición del costo por rubro</div>
                    <ResponsiveContainer width="100%" height={Math.max(180, selectedLoteByItem.length * 26 + 30)}>
                      <BarChart data={selectedLoteByItem} layout="vertical" margin={{ left: 10 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                        <XAxis type="number" tick={{ fontSize: 10, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                        <YAxis type="category" dataKey="name" width={100} tick={{ fontSize: 10, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                        <Tooltip formatter={(v) => fmtUSD2(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                        <Legend
                          wrapperStyle={{ fontSize: 11 }}
                          payload={[{ value: "Insumos", type: "square", color: "#4B6B3A" }, { value: "Servicios", type: "square", color: "#B8842E" }]}
                        />
                        <Bar dataKey="value" radius={[0, 3, 3, 0]}>
                          {selectedLoteByItem.map((e, i) => (
                            <Cell key={i} fill={e.tipoDet === "INSUMOS" ? "#4B6B3A" : e.tipoDet === "SERVICIOS" ? "#B8842E" : FALLBACK_COLORS[i % FALLBACK_COLORS.length]} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
                      <button
                        className="agri-chip" data-active={loteDetailTab === "resumen"}
                        style={{ padding: "3px 10px", fontSize: 12 }}
                        onClick={() => setLoteDetailTab("resumen")}
                      >Resumen por rubro</button>
                      <button
                        className="agri-chip" data-active={loteDetailTab === "movimientos"}
                        style={{ padding: "3px 10px", fontSize: 12 }}
                        onClick={() => setLoteDetailTab("movimientos")}
                      >Movimientos</button>
                    </div>
                    {loteDetailTab === "movimientos" ? (
                    <div style={{ overflowX: "auto", maxHeight: 320, overflowY: "auto", border: "1px solid var(--line)", borderRadius: 6 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 560 }}>
                        <thead>
                          <tr>
                            {["Fecha / Labor", "Rubro", "Concepto", "Cant.", "Dosis", "USD"].map((h) => (
                              <th key={h} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)", fontSize: 9, padding: "5px 7px" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {selectedLoteRows.map((r, i) => {
                            const prev = selectedLoteRows[i - 1];
                            const sameFechaLabor = prev && prev["Fecha"] === r["Fecha"] && prev["Origen"] === r["Origen"];
                            const sameRubro = sameFechaLabor && prev["Tipo item"] === r["Tipo item"];
                            const tdSm = { padding: "4px 7px", fontSize: 11 };
                            const tdLabor = { padding: "4px 7px", fontSize: 9, color: "var(--ink-soft)", fontWeight: 400 };
                            const ha = selectedLoteInfo.ha;
                            const dosis = ha > 0 ? r["Cantidad"] / ha : null;
                            return (
                              <tr key={i} className="agri-tr" style={{ borderTop: !sameFechaLabor && i > 0 ? "1px solid var(--ink)" : undefined }}>
                                <td className="agri-td" style={{ ...tdSm, color: sameFechaLabor ? "var(--line)" : "var(--ink)" }}>
                                  {sameFechaLabor ? "″" : (
                                    <>
                                      <div>{fmtDate(r["Fecha"])}</div>
                                      <div style={{ fontSize: 9, color: "var(--ink-soft)", fontWeight: 400 }}>{r["Origen"]}</div>
                                    </>
                                  )}
                                </td>
                                <td className="agri-td" style={{ ...tdLabor, color: sameRubro ? "var(--line)" : "var(--ink-soft)" }}>{sameRubro ? "″" : r["Tipo item"]}</td>
                                <td className="agri-td" style={tdLabor}>{r["Concepto"]}</td>
                                <td className="agri-td" style={{ ...tdSm, textAlign: "right" }}>{fmtNum(r["Cantidad"], 1)} {r["Unid."]}</td>
                                <td className="agri-td" style={{ ...tdSm, textAlign: "right", color: "var(--ink-soft)" }}>{dosis === null ? "—" : `${fmtNum(dosis, 2)} ${r["Unid."]}/ha`}</td>
                                <td className="agri-td" style={{ ...tdSm, textAlign: "right", fontWeight: 600 }}>{r["U$S/Total"] === null ? "—" : fmtUSD2(r["U$S/Total"])}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                    ) : (
                    <div style={{ overflowX: "auto", maxHeight: 320, overflowY: "auto", border: "1px solid var(--line)", borderRadius: 6 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 620 }}>
                        <thead>
                          <tr>
                            {["Rubro / Concepto", "$/U", "Dosis/ha", "Items", "Cant.", "Costo", "Costo/ha"].map((h) => (
                              <th key={h} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)", fontSize: 9, padding: "5px 7px", textAlign: h === "Rubro / Concepto" ? "left" : "right" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {selectedLoteSummary.map((dg, di) => (
                            <React.Fragment key={di}>
                              <tr className="agri-tr" style={{ background: "rgba(75,107,58,0.16)" }}>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.02em" }}>{dg.tipoDet}</td>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11 }}></td>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11 }}></td>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{dg.items}</td>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{fmtNum(dg.cant, 1)}</td>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{fmtUSD2(dg.costo)}</td>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{dg.costoHa === null ? "—" : fmtUSD2(dg.costoHa)}</td>
                              </tr>
                              {dg.rubros.map((rg, ri) => (
                                <React.Fragment key={ri}>
                                  <tr className="agri-tr" style={{ background: "rgba(75,107,58,0.07)" }}>
                                    <td className="agri-td" style={{ padding: "4px 7px 4px 14px", fontSize: 11, fontWeight: 700 }}>{rg.rubro}</td>
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11 }}></td>
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11 }}></td>
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{rg.items}</td>
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{fmtNum(rg.cant, 1)}</td>
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{fmtUSD2(rg.costo)}</td>
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{rg.costoHa === null ? "—" : fmtUSD2(rg.costoHa)}</td>
                                  </tr>
                                  {rg.conceptos.map((cg, ci) => (
                                    <tr key={ci} className="agri-tr">
                                      <td className="agri-td" style={{ padding: "4px 7px 4px 26px", fontSize: 10, color: "var(--ink-soft)" }}>{cg.concepto}</td>
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{cg.unitPrice === null ? "—" : fmtUSD2(cg.unitPrice)}</td>
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{cg.dosisHa === null ? "—" : `${fmtNum(cg.dosisHa, 2)} ${cg.unid}/ha`}</td>
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{cg.items}</td>
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{fmtNum(cg.cant, 1)} {cg.unid}</td>
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{fmtUSD2(cg.costo)}</td>
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{cg.costoHa === null ? "—" : fmtUSD2(cg.costoHa)}</td>
                                    </tr>
                                  ))}
                                </React.Fragment>
                              ))}
                            </React.Fragment>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    )}
                  </div>
                </div>
              </div>
            )}

            {/* Precios y consumo por ítem */}
            <div className="agri-card" style={{ padding: 0, overflow: "hidden", marginBottom: 20 }}>
              <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Precios y consumo por ítem</div>
                <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(itemsSorted.length)} ítems · según filtros de arriba</div>
              </div>
              <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700 }}>
                  <thead>
                    <tr>
                      {[["rubro", "Rubro"], ["concepto", "Concepto"], ["precioProm", "Precio prom."], ["cant", "Cantidad total"], ["costo", "Costo total"], ["movimientos", "Movs."]].map(([key, label]) => (
                        <th key={key} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)" }} onClick={() => toggleItemSort(key)}>
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 3 }}>
                            {label} {itemSort.key === key && <ArrowUpDown size={11} />}
                          </span>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {itemsSorted.map((e, i) => {
                      const prev = itemsSorted[i - 1];
                      const esNuevoGrupo = !prev || prev.tipoDet !== e.tipoDet;
                      return (
                        <React.Fragment key={i}>
                          {esNuevoGrupo && e.tipoDet && (
                            <tr>
                              <td colSpan={6} style={{ padding: "8px 10px 4px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", background: "rgba(75,107,58,0.10)", color: "var(--green)", borderTop: i > 0 ? "2px solid var(--ink)" : "none", borderBottom: "1px solid var(--line)" }}>
                                {e.tipoDet}
                              </td>
                            </tr>
                          )}
                          <tr className="agri-tr">
                            <td className="agri-td">{e.rubro}</td>
                            <td className="agri-td">{e.concepto}</td>
                            <td className="agri-td" style={{ textAlign: "right" }}>{e.precioProm === null ? "—" : fmtUSD2(e.precioProm)}</td>
                            <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(e.cant, 1)} {e.unid}</td>
                            <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{fmtUSD2(e.costo)}</td>
                            <td className="agri-td" style={{ textAlign: "right", color: "var(--ink-soft)" }}>{e.movimientos}</td>
                          </tr>
                        </React.Fragment>
                      );
                    })}
                    {itemsSorted.length === 0 && (
                      <tr><td className="agri-td" colSpan={6} style={{ textAlign: "center", padding: 30, color: "var(--ink-soft)" }}>Ningún ítem coincide con los filtros aplicados.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Tabla */}
            <div className="agri-card" style={{ padding: 0, overflow: "hidden" }}>
              <div style={{ padding: "12px 16px", borderBottom: showDetalle ? "1px solid var(--line)" : "none", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Detalle de movimientos</div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(sorted.length)} registros filtrados</span>
                  <button className="agri-btn agri-btn-outline" onClick={exportToExcel}><Download size={13} /> Exportar a Excel</button>
                  <button className="agri-btn agri-btn-outline" onClick={() => setShowDetalle((v) => !v)} title={showDetalle ? "Colapsar" : "Expandir"}>
                    {showDetalle ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                  </button>
                </div>
              </div>
              {showDetalle && (
              <>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 780 }}>
                  <thead>
                    <tr>
                      {[["Fecha", "Fecha"], ["Campo", "Campo"], ["Lote", "Lote"], ["Cultivo", "Cultivo"], ["Origen", "Labor"], ["Tipo item", "Rubro"], ["Concepto", "Concepto"], ["Cantidad", "Cant."], ["Unid.", "Un."], ["U$S/Total", "USD"]].map(([key, label]) => (
                        <th key={key} className="agri-th" onClick={() => toggleSort(key)}>
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 3 }}>
                            {label} {sort.key === key && <ArrowUpDown size={11} />}
                          </span>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {pageRows.map((r, i) => (
                      <tr key={i} className="agri-tr">
                        <td className="agri-td">{fmtDate(r["Fecha"])}</td>
                        <td className="agri-td">{r["Campo"]}</td>
                        <td className="agri-td">{r["Lote"]}</td>
                        <td className="agri-td">{r["Cultivo"]}</td>
                        <td className="agri-td">{r["Origen"]}</td>
                        <td className="agri-td">{r["Tipo item"]}</td>
                        <td className="agri-td">{r["Concepto"]}</td>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(r["Cantidad"], 2)}</td>
                        <td className="agri-td">{r["Unid."]}</td>
                        <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{r["U$S/Total"] === null ? "—" : fmtUSD2(r["U$S/Total"])}</td>
                      </tr>
                    ))}
                    {pageRows.length === 0 && (
                      <tr><td className="agri-td" colSpan={10} style={{ textAlign: "center", padding: 30, color: "var(--ink-soft)" }}>Ningún registro coincide con los filtros aplicados.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 16px", borderTop: "1px solid var(--line)" }}>
                <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>Página {page} de {totalPages}</span>
                <div style={{ display: "flex", gap: 6 }}>
                  <button className="agri-btn agri-btn-outline" disabled={page <= 1} onClick={() => setPage((p) => p - 1)} style={{ opacity: page <= 1 ? 0.4 : 1 }}><ChevronLeft size={14} /></button>
                  <button className="agri-btn agri-btn-outline" disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)} style={{ opacity: page >= totalPages ? 0.4 : 1 }}><ChevronRight size={14} /></button>
                </div>
              </div>
              </>
              )}
            </div>

            {/* Control de datos: colapsado por defecto, no invasivo */}
            <div className="agri-card" style={{ padding: 16, marginTop: 16 }}>
              <button
                onClick={() => setShowDataQuality((v) => !v)}
                style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%", background: "none", border: "none", cursor: "pointer", padding: 0, font: "inherit", color: "inherit" }}
              >
                <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <ClipboardCheck size={16} color={dataQuality.list.length > 0 ? "var(--rust)" : "var(--green)"} />
                  <span className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Control de datos</span>
                  {dataQuality.list.length > 0 ? (
                    <span style={{ fontSize: 11, fontWeight: 700, color: "#fff", background: "var(--rust)", borderRadius: 999, padding: "2px 9px" }}>
                      {fmtNum(dataQuality.list.length)} ítems con problemas
                    </span>
                  ) : (
                    <span style={{ fontSize: 11, fontWeight: 700, color: "#fff", background: "var(--green)", borderRadius: 999, padding: "2px 9px" }}>Sin observaciones</span>
                  )}
                </span>
                {showDataQuality ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
              </button>

              {showDataQuality && (
                <div style={{ marginTop: 14 }}>
                  <div style={{ fontSize: 12, color: "var(--ink-soft)", marginBottom: 12, maxWidth: 720 }}>
                    Estos casos no rompen el dashboard (se tratan como $0 o se excluyen de los promedios), pero conviene corregirlos en el Excel de origen para que los números reflejen la realidad. Se calculan sobre todos los datos cargados, sin importar los filtros de arriba.
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
                    {[
                      ["Sin precio", dataQuality.totals.sinPrecio],
                      ["Precio en $0", dataQuality.totals.precioCero],
                      ["Sin superficie", dataQuality.totals.sinSuperficie],
                      ["Sin cultivo", dataQuality.totals.sinCultivo],
                      ["Sin campo", dataQuality.totals.sinCampo],
                      ["Sin concepto", dataQuality.totals.sinConcepto],
                    ].filter(([, n]) => n > 0).map(([label, n]) => (
                      <div key={label} style={{ fontSize: 12, padding: "5px 11px", borderRadius: 999, border: "1px solid var(--line)", background: "var(--paper)" }}>
                        <strong>{fmtNum(n)}</strong> {label.toLowerCase()}
                      </div>
                    ))}
                  </div>
                  {dataQuality.list.length === 0 ? (
                    <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>No se encontraron problemas en los datos cargados. 🎉</div>
                  ) : (
                    <div style={{ overflowX: "auto", maxHeight: 320, overflowY: "auto", border: "1px solid var(--line)", borderRadius: 6 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 640 }}>
                        <thead>
                          <tr>
                            {["Rubro", "Concepto", "Problema", "Filas", "Cantidad total"].map((h) => (
                              <th key={h} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {dataQuality.list.map((g, i) => (
                            <tr key={i} className="agri-tr">
                              <td className="agri-td">{g.rubro}</td>
                              <td className="agri-td">{g.concepto}</td>
                              <td className="agri-td" style={{ color: "var(--rust)" }}>{g.problems.join(", ")}</td>
                              <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(g.count)}</td>
                              <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(g.cant, 1)} {g.unid}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              )}
            </div>
          </>
        )}
    </>
  );
}

// ---------------------------------------------------------------------------
// Login del administrador
// ---------------------------------------------------------------------------
function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      await signInWithEmailAndPassword(cloudAuth, email, password);
    } catch (err) {
      setError("Usuario o contraseña incorrectos.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="agri-shell" style={{ maxWidth: 380, paddingTop: 70 }}>
      <div className="agri-card" style={{ padding: 28 }}>
        <div className="agri-serif" style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>Panel de administrador</div>
        <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 20 }}>Ingresá con tu cuenta para gestionar los clientes y sus datos.</div>
        <form onSubmit={handleSubmit} style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <input className="agri-input" type="email" placeholder="Email" value={email} autoComplete="username"
            onChange={(e) => setEmail(e.target.value)} required />
          <input className="agri-input" type="password" placeholder="Contraseña" value={password} autoComplete="current-password"
            onChange={(e) => setPassword(e.target.value)} required />
          {error && <div style={{ color: "var(--rust)", fontSize: 13 }}>{error}</div>}
          <button className="agri-btn" type="submit" disabled={loading} style={{ justifyContent: "center", marginTop: 4 }}>
            {loading ? "Ingresando…" : "Ingresar"}
          </button>
        </form>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Panel de administrador: alta/baja de clientes y acceso a sus dashboards
// ---------------------------------------------------------------------------
const slugify = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "")
  .replace(/[^a-z0-9]+/g, "-").replace(/(^-+|-+$)/g, "");

function AdminPanel() {
  const [clients, setClients] = useState(null); // null = cargando
  const [nombre, setNombre] = useState("");
  const [adding, setAdding] = useState(false);
  const [addError, setAddError] = useState("");
  const [copiedSlug, setCopiedSlug] = useState("");
  const [editingSlug, setEditingSlug] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [editError, setEditError] = useState("");
  const [savingEdit, setSavingEdit] = useState(false);
  const [pendingAdd, setPendingAdd] = useState(null); // { slug, nombre } mientras confirmamos que se creó
  const [pendingEdit, setPendingEdit] = useState(null); // { slug, nombre } mientras confirmamos que se guardó

  const baseUrl = `${window.location.origin}${window.location.pathname}`;
  const linkFor = (slug) => `${baseUrl}?cliente=${slug}`;

  useEffect(() => {
    const dbRef = ref(cloudDb, "clientes/index");
    const unsub = onValue(dbRef, (snap) => {
      const val = snap.val() || {};
      const list = Object.entries(val).map(([slug, v]) => ({ slug, nombre: (v && v.nombre) || slug, creadoEn: v && v.creadoEn }));
      list.sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
      setClients(list);
    });
    return unsub;
  }, []);

  // Confirmación real vía suscripción en vivo (única fuente de verdad).
  // La promesa de dbSet no es confiable en este entorno — la ignoramos y
  // miramos directamente si el dato aparece en la lista.
  useEffect(() => {
    if (!pendingAdd || !clients) return;
    if (clients.some((c) => c.slug === pendingAdd.slug && c.nombre === pendingAdd.nombre)) {
      // Éxito confirmado
      setNombre(""); setAddError(""); setAdding(false); setPendingAdd(null);
      return;
    }
    // Si ya pasaron 8 segundos y no llegó, recién entonces avisamos error real
    const t = setTimeout(() => {
      setPendingAdd((cur) => {
        if (cur && cur.slug === pendingAdd.slug) {
          setAddError("No se pudo crear el cliente. Revisá tu conexión e intentá de nuevo.");
          setAdding(false);
          return null;
        }
        return cur;
      });
    }, 8000);
    return () => clearTimeout(t);
  }, [clients, pendingAdd?.slug]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (!pendingEdit || !clients) return;
    // Para migración de slug, chequeamos el nuevo; para edición de nombre, chequeamos el nombre
    const confirmado = clients.some((c) => c.slug === pendingEdit.slug && c.nombre === pendingEdit.nombre);
    const viejoDesaparecio = pendingEdit.oldSlug !== pendingEdit.slug
      ? !clients.some((c) => c.slug === pendingEdit.oldSlug)
      : true;
    if (confirmado && viejoDesaparecio) {
      setEditingSlug(null); setEditValue(""); setEditError(""); setSavingEdit(false); setPendingEdit(null);
      return;
    }
    const t = setTimeout(() => {
      setPendingEdit((cur) => {
        if (cur) {
          setEditError("No se pudo guardar. Intentá de nuevo.");
          setSavingEdit(false);
          return null;
        }
        return cur;
      });
    }, 8000);
    return () => clearTimeout(t);
  }, [clients, pendingEdit?.slug]); // eslint-disable-line react-hooks/exhaustive-deps

  const nombreDuplicado = (nombreNuevo, slugAEexcluir) => {
    const norm = nombreNuevo.trim().toLowerCase();
    return (clients || []).some((c) => c.slug !== slugAEexcluir && c.nombre.trim().toLowerCase() === norm);
  };

  const addClient = async (e) => {
    e.preventDefault();
    setAddError("");
    const trimmed = nombre.trim();
    const base = slugify(trimmed);
    if (!base) { setAddError("Ingresá un nombre válido."); return; }
    if (nombreDuplicado(trimmed, null)) { setAddError("Ya existe un cliente con ese nombre."); return; }
    const existingSlugs = new Set((clients || []).map((c) => c.slug));
    let slug = base, n = 2;
    while (existingSlugs.has(slug)) { slug = `${base}-${n}`; n++; }
    setAdding(true);
    setPendingAdd({ slug, nombre: trimmed });
    // Disparamos la escritura sin await — la suscripción onValue va a confirmar
    // el resultado real. Si en 8 segundos no llegó, recién ahí mostramos error.
    dbSet(ref(cloudDb, `clientes/index/${slug}`), { nombre: trimmed, creadoEn: new Date().toISOString() });
  };

  const startEdit = (c) => { setEditingSlug(c.slug); setEditValue(c.nombre); setEditError(""); };
  const cancelEdit = () => { setEditingSlug(null); setEditValue(""); setEditError(""); };

  const saveEdit = async (slug) => {
    const trimmed = editValue.trim();
    if (!trimmed) { setEditError("El nombre no puede quedar vacío."); return; }
    if (nombreDuplicado(trimmed, slug)) { setEditError("Ya existe un cliente con ese nombre."); return; }
    const newBase = slugify(trimmed);
    if (!newBase) { setEditError("Ingresá un nombre válido."); return; }
    const otherSlugs = new Set((clients || []).map((c) => c.slug).filter((s) => s !== slug));
    let newSlug = newBase, n = 2;
    while (otherSlugs.has(newSlug)) { newSlug = `${newBase}-${n}`; n++; }

    if (newSlug !== slug) {
      const ok = window.confirm(
        "Ese nombre genera un link nuevo para este cliente. El link actual va a dejar de funcionar " +
        "(sus datos se migran al nuevo link). ¿Continuar?"
      );
      if (!ok) return;
    }

    setSavingEdit(true);
    setPendingEdit({ oldSlug: slug, slug: newSlug, nombre: trimmed });

    if (newSlug !== slug) {
      // Migración: leer datos del viejo, escribir en el nuevo, borrar el viejo
      const oldClient = (clients || []).find((c) => c.slug === slug);
      const snap = await get(ref(cloudDb, `clientes/${slug}/dataset`));
      if (snap.exists()) {
        dbSet(ref(cloudDb, `clientes/${newSlug}/dataset`), snap.val());
      }
      dbSet(ref(cloudDb, `clientes/index/${newSlug}`), {
        nombre: trimmed,
        creadoEn: (oldClient && oldClient.creadoEn) || new Date().toISOString(),
      });
      dbSet(ref(cloudDb, `clientes/index/${slug}`), null);
      dbSet(ref(cloudDb, `clientes/${slug}`), null);
    } else {
      dbSet(ref(cloudDb, `clientes/index/${slug}/nombre`), trimmed);
    }
    // La suscripción onValue confirma el resultado real
  };

  const deleteClient = async (slug, nombreCliente) => {
    if (!window.confirm(`¿Eliminar a "${nombreCliente}" y todos sus datos? Esta acción no se puede deshacer.`)) return;
    try {
      await dbSet(ref(cloudDb, `clientes/index/${slug}`), null);
      await dbSet(ref(cloudDb, `clientes/${slug}`), null);
    } catch (err) { /* silencioso: el listado se refresca solo si falla parcialmente */ }
  };

  const copyLink = (slug) => {
    const link = linkFor(slug);
    if (navigator.clipboard) navigator.clipboard.writeText(link).catch(() => {});
    setCopiedSlug(slug);
    setTimeout(() => setCopiedSlug(""), 1500);
  };

  return (
    <div className="agri-shell">
      <div className="agri-masthead">
        <div>
          <div className="agri-title agri-serif">Panel de administrador</div>
          <div className="agri-sub">Gestioná tus clientes y el acceso a sus datos.</div>
        </div>
        <button className="agri-btn agri-btn-outline" onClick={() => signOut(cloudAuth)}>
          <LogOut size={14} /> Cerrar sesión
        </button>
      </div>

      <div className="agri-card" style={{ padding: 16, marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 10 }}>Agregar cliente</div>
        <form onSubmit={addClient} style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <input className="agri-input" style={{ flex: "1 1 220px" }} placeholder="Nombre del cliente (ej. Cliente1)"
            value={nombre} onChange={(e) => setNombre(e.target.value)} />
          <button className="agri-btn" disabled={adding} type="submit"><Plus size={14} /> Agregar</button>
        </form>
        {addError && <div style={{ color: "var(--rust)", fontSize: 13, marginTop: 6 }}>{addError}</div>}
      </div>

      <div className="agri-card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)" }}>
          <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Clientes ({clients ? clients.length : 0})</div>
        </div>
        {clients === null ? (
          <div style={{ padding: 30, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>
        ) : clients.length === 0 ? (
          <div style={{ padding: 30, textAlign: "center", color: "var(--ink-soft)" }}>Todavía no agregaste ningún cliente.</div>
        ) : (
          clients.map((c) => (
            <div key={c.slug} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "12px 16px", borderBottom: "1px solid var(--line)", flexWrap: "wrap" }}>
              {editingSlug === c.slug ? (
                <div style={{ flex: "1 1 220px" }}>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    <input className="agri-input" style={{ flex: "1 1 180px" }} value={editValue}
                      onChange={(e) => setEditValue(e.target.value)} autoFocus
                      onKeyDown={(e) => { if (e.key === "Enter") saveEdit(c.slug); if (e.key === "Escape") cancelEdit(); }} />
                    <button className="agri-btn" disabled={savingEdit} onClick={() => saveEdit(c.slug)}>{savingEdit ? "Guardando…" : "Guardar"}</button>
                    <button className="agri-btn agri-btn-outline" onClick={cancelEdit}>Cancelar</button>
                  </div>
                  {editError && <div style={{ color: "var(--rust)", fontSize: 12, marginTop: 4 }}>{editError}</div>}
                </div>
              ) : (
                <div>
                  <div style={{ fontWeight: 600, fontSize: 14 }}>{c.nombre}</div>
                  <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{linkFor(c.slug)}</div>
                </div>
              )}
              {editingSlug !== c.slug && (
                <div style={{ display: "flex", gap: 6 }}>
                  <button className="agri-btn agri-btn-outline" onClick={() => copyLink(c.slug)}>
                    {copiedSlug === c.slug ? "¡Copiado!" : "Copiar link"}
                  </button>
                  <a className="agri-btn agri-btn-outline" href={linkFor(c.slug)} target="_blank" rel="noreferrer">Abrir</a>
                  <button className="agri-btn agri-btn-outline" onClick={() => startEdit(c)}>Editar</button>
                  <button className="agri-btn agri-btn-outline" style={{ color: "var(--rust)" }} onClick={() => deleteClient(c.slug, c.nombre)}>
                    <Trash2 size={13} />
                  </button>
                </div>
              )}
            </div>
          ))
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Componente raíz: decide qué mostrar según la URL (?cliente=slug) y la sesión
// ---------------------------------------------------------------------------
export default function App() {
  const [authUser, setAuthUser] = useState(undefined); // undefined = resolviendo, null = sin sesión
  const clienteSlug = useMemo(() => {
    if (typeof window === "undefined") return null;
    return new URLSearchParams(window.location.search).get("cliente");
  }, []);

  // Fuente tipográfica (una sola vez, para toda la app)
  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap";
    document.head.appendChild(link);
    return () => { try { document.head.removeChild(link); } catch (e) {} };
  }, []);

  // Sesión de administrador (persiste sola entre recargas gracias a Firebase Auth)
  useEffect(() => {
    if (!CLOUD_SYNC_ENABLED || !cloudAuth) { setAuthUser(null); return; }
    const unsub = onAuthStateChanged(cloudAuth, (u) => setAuthUser(u));
    return unsub;
  }, []);

  let content;
  if (!CLOUD_SYNC_ENABLED) {
    content = (
      <div className="agri-shell" style={{ paddingTop: 60, textAlign: "center", color: "var(--ink-soft)" }}>
        El modo multi-cliente necesita Firebase configurado en <code>firebaseConfig.js</code>.
      </div>
    );
  } else if (authUser === undefined) {
    content = <div style={{ padding: 60, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>;
  } else if (!clienteSlug) {
    content = authUser ? <AdminPanel /> : <LoginScreen />;
  } else {
    content = <ClienteDashboard slug={clienteSlug} isAdmin={!!authUser} />;
  }

  return (
    <div className="agri-root">
      <style>{AGRI_STYLES}</style>
      {content}
    </div>
  );
}
