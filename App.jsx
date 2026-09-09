import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, onValue, set as dbSet } from "firebase/database";
import { firebaseConfig, CLOUD_SYNC_ENABLED, CLOUD_PATH } from "./firebaseConfig.js";
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell,
} from "recharts";
import {
  Upload, RefreshCw, Search, X, Sprout, MapPin, DollarSign, Ruler,
  ChevronLeft, ChevronRight, ArrowUpDown, Trash2, FileSpreadsheet, AlertCircle, Eye,
  Cloud, CloudOff,
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

const STORAGE_KEY = "produccion-agricola-dataset-v1";
const PAGE_SIZE = 30;

// Conexión a Firebase (solo si se completó firebaseConfig.js con datos reales)
let cloudDb = null;
if (CLOUD_SYNC_ENABLED) {
  try {
    const fbApp = initializeApp(firebaseConfig);
    cloudDb = getDatabase(fbApp);
  } catch (e) {
    console.error("No se pudo inicializar Firebase:", e);
  }
}

const CULTIVO_COLORS = { SOJA: "#5B7C4B", MAIZ: "#C68F41", POROTO: "#35606B" };
const FALLBACK_COLORS = ["#5B7C4B", "#C68F41", "#35606B", "#8A5A3B", "#7C8A4B", "#A1462F"];
const cultivoColor = (name, i) => CULTIVO_COLORS[String(name).toUpperCase()] || FALLBACK_COLORS[i % FALLBACK_COLORS.length];

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
    "U$S/U": numOrNull(r["U$S/U"]),
    "U$S/Total": numOrNull(r["U$S/Total"]),
  };
}

// ---------------------------------------------------------------------------
// Componente principal
// ---------------------------------------------------------------------------
export default function App() {
  const [rows, setRows] = useState([]);
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
  const [tipoDet, setTipoDet] = useState("Todos");
  const [tipoItem, setTipoItem] = useState("Todos");
  const [desde, setDesde] = useState("");
  const [hasta, setHasta] = useState("");
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [sort, setSort] = useState({ key: "Fecha", dir: "desc" });

  // Cargar fuente tipográfica
  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap";
    document.head.appendChild(link);
    return () => { try { document.head.removeChild(link); } catch (e) {} };
  }, []);

  // Cargar último dataset guardado: primero del caché local (instantáneo),
  // y si hay Firebase configurado, nos suscribimos a la nube (tiempo real).
  useEffect(() => {
    let cloudUnsub = null;

    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && res.value) {
          const payload = JSON.parse(res.value);
          const objRows = payload.rows.map((arr) => {
            const o = {};
            payload.columns.forEach((c, i) => { o[c] = arr[i]; });
            return o;
          });
          setRows(objRows);
          setMeta({ fileName: payload.fileName, updatedAt: payload.updatedAt, rowCount: objRows.length });
        }
      } catch (e) {
        // no hay datos guardados todavía en este dispositivo, es normal
      } finally {
        if (!CLOUD_SYNC_ENABLED) setBootLoading(false);
      }

      if (CLOUD_SYNC_ENABLED && cloudDb) {
        const dbRef = ref(cloudDb, CLOUD_PATH);
        cloudUnsub = onValue(
          dbRef,
          (snapshot) => {
            const payload = snapshot.val();
            if (payload && payload.rows && payload.columns) {
              const objRows = payload.rows.map((arr) => {
                const o = {};
                payload.columns.forEach((c, i) => { o[c] = arr[i]; });
                return o;
              });
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
    const dataRows = normalizedRows.map((r) => columns.map((c) => r[c]));
    const payload = { columns, rows: dataRows, fileName, updatedAt: new Date().toISOString() };

    try {
      await window.storage.set(STORAGE_KEY, JSON.stringify(payload), false);
    } catch (e) {
      console.error("No se pudo guardar en caché local:", e);
    }

    if (CLOUD_SYNC_ENABLED && cloudDb) {
      try {
        await dbSet(ref(cloudDb, CLOUD_PATH), payload);
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
    } catch (e) {
      setError("No pude leer el archivo. Verificá que sea un Excel exportado del mismo formato (" + e.message + ")");
    } finally {
      setParsing(false);
    }
  }, [persist]);

  const onDrop = (e) => {
    e.preventDefault(); setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) handleFile(f);
  };

  const clearDataset = async () => {
    setRows([]); setMeta(null);
    try { await window.storage.delete(STORAGE_KEY, false); } catch (e) {}
    if (CLOUD_SYNC_ENABLED && cloudDb) {
      try { await dbSet(ref(cloudDb, CLOUD_PATH), null); } catch (e) {}
    }
  };

  // Listas para filtros
  const cultivos = useMemo(() => Array.from(new Set(rows.map((r) => r["Cultivo"]).filter(Boolean))).sort(), [rows]);
  const campos = useMemo(() => Array.from(new Set(rows.map((r) => r["Campo"]).filter(Boolean))).sort(), [rows]);
  const tipoItems = useMemo(() => Array.from(new Set(rows.map((r) => r["Tipo item"]).filter(Boolean))).sort(), [rows]);
  // Los lotes dependen del campo y cultivo elegidos, para no listar lotes que no aplican
  const loteOptions = useMemo(() => {
    const base = rows.filter((r) => (campo === "Todos" || r["Campo"] === campo) && (cultivo === "Todos" || r["Cultivo"] === cultivo));
    const map = new Map();
    base.forEach((r) => {
      if (!r["Lote"]) return;
      const value = `${r["Campo"]}::${r["Lote"]}`;
      if (!map.has(value)) map.set(value, campo === "Todos" ? `${r["Campo"]} · Lote ${r["Lote"]}` : `Lote ${r["Lote"]}`);
    });
    return Array.from(map.entries()).map(([value, label]) => ({ value, label })).sort((a, b) => a.label.localeCompare(b.label, "es", { numeric: true }));
  }, [rows, campo, cultivo]);
  // Si el lote elegido deja de tener sentido con el nuevo campo/cultivo, lo reseteamos
  useEffect(() => {
    if (lote !== "Todos" && !loteOptions.some((o) => o.value === lote)) setLote("Todos");
  }, [loteOptions]); // eslint-disable-line react-hooks/exhaustive-deps

  // Filtrado
  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return rows.filter((r) => {
      if (cultivo !== "Todos" && r["Cultivo"] !== cultivo) return false;
      if (campo !== "Todos" && r["Campo"] !== campo) return false;
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
  }, [rows, cultivo, campo, lote, tipoDet, tipoItem, desde, hasta, search]);

  useEffect(() => { setPage(1); }, [cultivo, campo, lote, tipoDet, tipoItem, desde, hasta, search]);

  // KPIs
  const kpis = useMemo(() => {
    let gasto = 0, sinPrecio = 0;
    const supMap = new Map();
    const lotesSet = new Set(), camposSet = new Set(), cultivosSet = new Set();
    filtered.forEach((r) => {
      if (r["U$S/Total"] === null) sinPrecio++; else gasto += r["U$S/Total"];
      const key = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}`;
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

  // Datos para gráficos (incluyen hectáreas y costo/ha, deduplicando superficie
  // por Campo+Lote+Cultivo para no sumar la misma hectárea varias veces)
  const porCultivo = useMemo(() => {
    const m = new Map();
    const seenSup = new Set();
    filtered.forEach((r) => {
      if (!r["Cultivo"]) return;
      if (!m.has(r["Cultivo"])) m.set(r["Cultivo"], { gasto: 0, ha: 0 });
      const e = m.get(r["Cultivo"]);
      e.gasto += r["U$S/Total"] || 0;
      const supKey = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}`;
      if (!seenSup.has(supKey)) { seenSup.add(supKey); e.ha += r["Sup.Cultivo"] || 0; }
    });
    return Array.from(m.entries()).map(([name, e]) => ({ name, value: e.gasto, ha: e.ha, costoHa: e.ha > 0 ? e.gasto / e.ha : 0 })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  const porCampo = useMemo(() => {
    const m = new Map();
    const seenSup = new Set();
    filtered.forEach((r) => {
      if (!r["Campo"]) return;
      if (!m.has(r["Campo"])) m.set(r["Campo"], { gasto: 0, ha: 0 });
      const e = m.get(r["Campo"]);
      e.gasto += r["U$S/Total"] || 0;
      const supKey = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}`;
      if (!seenSup.has(supKey)) { seenSup.add(supKey); e.ha += r["Sup.Cultivo"] || 0; }
    });
    return Array.from(m.entries()).map(([name, e]) => ({ name, value: e.gasto, ha: e.ha, costoHa: e.ha > 0 ? e.gasto / e.ha : 0 })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  const costoHaPorCultivo = useMemo(() => [...porCultivo].filter((e) => e.ha > 0).sort((a, b) => b.costoHa - a.costoHa), [porCultivo]);
  const costoHaPorCampo = useMemo(() => [...porCampo].filter((e) => e.ha > 0).sort((a, b) => b.costoHa - a.costoHa), [porCampo]);

  // Tooltip que muestra gasto total y costo/ha juntos
  const GastoTooltip = ({ active, payload, label }) => {
    if (!active || !payload || !payload.length) return null;
    const d = payload[0].payload;
    return (
      <div style={{ background: "#FCFAF2", border: "1px solid #DCD2B8", borderRadius: 6, padding: "8px 10px", fontSize: 12 }}>
        <div style={{ fontWeight: 600, marginBottom: 3 }}>{label ?? d.name}</div>
        <div>Gasto: {fmtUSD(d.value)}</div>
        {d.ha > 0 && <div style={{ color: "var(--ink-soft)" }}>{fmtNum(d.ha, 1)} ha · {fmtUSD2(d.costoHa)}/ha</div>}
      </div>
    );
  };

  const evolucionMensual = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      if (!r["Fecha"]) return;
      const ym = r["Fecha"].slice(0, 7);
      m.set(ym, (m.get(ym) || 0) + (r["U$S/Total"] || 0));
    });
    return Array.from(m.entries()).sort(([a], [b]) => a.localeCompare(b)).map(([ym, value]) => ({ ym, label: monthLabel(ym), value }));
  }, [filtered]);

  const porTipoItem = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => { if (!r["Tipo item"]) return; m.set(r["Tipo item"], (m.get(r["Tipo item"]) || 0) + (r["U$S/Total"] || 0)); });
    const arr = Array.from(m.entries()).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    const top = arr.slice(0, 7);
    const restoVal = arr.slice(7).reduce((a, b) => a + b.value, 0);
    if (restoVal > 0) top.push({ name: "Otros", value: restoVal });
    return top;
  }, [filtered]);

  // Comparativo por lotes (agrupa Campo + Lote + Cultivo dentro de lo ya filtrado)
  const loteAgg = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const key = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}`;
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
  const topLotesPorGasto = useMemo(() => [...loteAgg].sort((a, b) => b.total - a.total).slice(0, 12).map((e) => ({ name: `${e.campo} · L${e.lote}`, value: e.total })), [loteAgg]);

  // Detalle de un lote seleccionado
  const [selectedLoteKey, setSelectedLoteKey] = useState(null);
  const selectLote = (key) => setSelectedLoteKey((cur) => (cur === key ? null : key));
  const selectedLoteInfo = useMemo(() => loteAgg.find((e) => `${e.campo}|${e.lote}|${e.cultivo}` === selectedLoteKey) || null, [loteAgg, selectedLoteKey]);
  const selectedLoteRows = useMemo(() => {
    if (!selectedLoteKey) return [];
    return filtered.filter((r) => `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}` === selectedLoteKey)
      .slice().sort((a, b) => {
        const fa = a["Fecha"] || "", fb = b["Fecha"] || "";
        if (fa !== fb) return fa < fb ? -1 : 1;
        const ra = a["Tipo item"] || "", rb = b["Tipo item"] || "";
        if (ra !== rb) return ra.localeCompare(rb, "es");
        const ca = a["Concepto"] || "", cb = b["Concepto"] || "";
        return ca.localeCompare(cb, "es");
      });
  }, [filtered, selectedLoteKey]);
  const selectedLoteByItem = useMemo(() => {
    const m = new Map();
    selectedLoteRows.forEach((r) => { if (!r["Tipo item"]) return; m.set(r["Tipo item"], (m.get(r["Tipo item"]) || 0) + (r["U$S/Total"] || 0)); });
    return Array.from(m.entries()).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [selectedLoteRows]);
  // Si cambian los filtros generales y el lote seleccionado deja de existir, lo deseleccionamos
  useEffect(() => {
    if (selectedLoteKey && !loteAgg.some((e) => `${e.campo}|${e.lote}|${e.cultivo}` === selectedLoteKey)) setSelectedLoteKey(null);
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

  const totalPages = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
  const pageRows = sorted.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  const toggleSort = (key) => {
    setSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "asc" }));
  };

  const resetFiltros = () => { setCultivo("Todos"); setCampo("Todos"); setLote("Todos"); setTipoDet("Todos"); setTipoItem("Todos"); setDesde(""); setHasta(""); setSearch(""); };
  const hayFiltrosActivos = cultivo !== "Todos" || campo !== "Todos" || lote !== "Todos" || tipoDet !== "Todos" || tipoItem !== "Todos" || desde || hasta || search;

  // -------------------------------------------------------------------------
  return (
    <div className="agri-root">
      <style>{`
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
        .agri-serif { font-family: 'Fraunces', Georgia, serif; }
        .agri-shell { max-width: 1180px; margin: 0 auto; }
        .agri-masthead {
          display: flex; align-items: flex-end; justify-content: space-between; gap: 16px;
          border-bottom: 2px solid var(--ink); padding-bottom: 14px; margin-bottom: 22px; flex-wrap: wrap;
        }
        .agri-title { font-size: 28px; font-weight: 600; letter-spacing: -0.01em; line-height: 1.1; }
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
        .agri-kpi-value { font-family: 'Fraunces', Georgia, serif; font-size: 26px; font-weight: 600; margin-top: 2px; font-variant-numeric: tabular-nums; }
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
          .agri-title { font-size: 22px; }
          .agri-kpi-value { font-size: 20px; }
          .agri-detail-grid { grid-template-columns: 1fr !important; }
        }
      `}</style>

      <div className="agri-shell">
        {/* Masthead */}
        <div className="agri-masthead">
          <div>
            <div className="agri-title agri-serif">Campaña C25/26 · Producción</div>
            <div className="agri-sub">
              {meta ? (
                <>Datos de <strong>{meta.fileName}</strong> · {fmtNum(meta.rowCount)} registros · actualizado {new Date(meta.updatedAt).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" })}</>
              ) : "Cargá tu planilla para empezar a consultar los datos"}
            </div>
            {CLOUD_SYNC_ENABLED && (
              <div style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 600, marginTop: 6, padding: "3px 9px", borderRadius: 999, background: syncStatus === "cloud" ? "rgba(75,107,58,0.12)" : syncStatus === "error" ? "rgba(161,70,47,0.12)" : "rgba(107,94,79,0.12)", color: syncStatus === "cloud" ? "var(--green)" : syncStatus === "error" ? "var(--rust)" : "var(--ink-soft)" }}>
                {syncStatus === "cloud" ? <><Cloud size={12} /> Sincronizado en todos tus dispositivos</> : syncStatus === "error" ? <><CloudOff size={12} /> Sin conexión a la nube · usando datos locales</> : <><Cloud size={12} /> Conectando…</>}
              </div>
            )}
          </div>
          {meta && (
            <div style={{ display: "flex", gap: 8 }}>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
                onChange={(e) => handleFile(e.target.files?.[0])} />
              <button className="agri-btn" onClick={() => fileInputRef.current?.click()}>
                <RefreshCw size={14} /> Actualizar datos
              </button>
              <button className="agri-btn agri-btn-outline" onClick={clearDataset}>
                <Trash2 size={14} />
              </button>
            </div>
          )}
        </div>

        {bootLoading ? (
          <div style={{ padding: 60, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>
        ) : !meta || rows.length === 0 ? (
          <div
            className="agri-dropzone" data-drag={dragOver}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
          >
            <FileSpreadsheet size={40} color="var(--ink-soft)" />
            <div className="agri-serif" style={{ fontSize: 19, fontWeight: 600 }}>
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

            {/* Gráficos */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 14, marginBottom: 20 }}>
              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Gasto por cultivo</div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={porCultivo} margin={{ left: -10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                    <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v / 1000) + "k"} axisLine={false} tickLine={false} />
                    <Tooltip content={<GastoTooltip />} />
                    <Bar dataKey="value" radius={[3, 3, 0, 0]}>
                      {porCultivo.map((e, i) => <Cell key={i} fill={cultivoColor(e.name, i)} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Gasto por campo</div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={porCampo} layout="vertical" margin={{ left: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v / 1000) + "k"} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="name" width={100} tick={{ fontSize: 11, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                    <Tooltip content={<GastoTooltip />} />
                    <Bar dataKey="value" fill="#B8842E" radius={[0, 3, 3, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

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
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Costo por hectárea · por campo</div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={costoHaPorCampo} layout="vertical" margin={{ left: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="name" width={100} tick={{ fontSize: 11, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v) => fmtUSD2(v) + "/ha"} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Bar dataKey="costoHa" fill="#2F5B66" radius={[0, 3, 3, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Evolución mensual del gasto</div>
                <ResponsiveContainer width="100%" height={220}>
                  <AreaChart data={evolucionMensual} margin={{ left: -10 }}>
                    <defs>
                      <linearGradient id="gradGasto" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor="#2F5B66" stopOpacity={0.35} />
                        <stop offset="95%" stopColor="#2F5B66" stopOpacity={0.03} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v / 1000) + "k"} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v) => fmtUSD(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Area type="monotone" dataKey="value" stroke="#2F5B66" strokeWidth={2} fill="url(#gradGasto)" />
                  </AreaChart>
                </ResponsiveContainer>
              </div>

              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Principales rubros de gasto</div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={porTipoItem} layout="vertical" margin={{ left: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v / 1000) + "k"} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 10, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v) => fmtUSD(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Bar dataKey="value" fill="#4B6B3A" radius={[0, 3, 3, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Comparativo por lotes */}
            <div className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Lotes con mayor gasto</div>
              <ResponsiveContainer width="100%" height={Math.max(180, topLotesPorGasto.length * 26)}>
                <BarChart data={topLotesPorGasto} layout="vertical" margin={{ left: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v / 1000) + "k"} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 11, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                  <Tooltip formatter={(v) => fmtUSD(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                  <Bar dataKey="value" fill="#2F5B66" radius={[0, 3, 3, 0]} />
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
                      const key = `${e.campo}|${e.lote}|${e.cultivo}`;
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
                    <div className="agri-serif" style={{ fontSize: 18, fontWeight: 600 }}>
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
                    <div className="agri-kpi-value" style={{ fontSize: 20 }}>{fmtUSD(selectedLoteInfo.total)}</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Insumos</div>
                    <div className="agri-kpi-value" style={{ fontSize: 20 }}>{fmtUSD(selectedLoteInfo.insumos)}</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Servicios</div>
                    <div className="agri-kpi-value" style={{ fontSize: 20 }}>{fmtUSD(selectedLoteInfo.servicios)}</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Costo por hectárea</div>
                    <div className="agri-kpi-value" style={{ fontSize: 20 }}>{fmtUSD2(selectedLoteInfo.costoHa)}</div>
                  </div>
                </div>

                <div className="agri-detail-grid" style={{ display: "grid", gridTemplateColumns: "minmax(260px, 1fr) minmax(280px, 1.3fr)", gap: 16 }}>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Composición del costo por rubro</div>
                    <ResponsiveContainer width="100%" height={Math.max(160, selectedLoteByItem.length * 26)}>
                      <BarChart data={selectedLoteByItem} layout="vertical" margin={{ left: 10 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                        <XAxis type="number" tick={{ fontSize: 10, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                        <YAxis type="category" dataKey="name" width={125} tick={{ fontSize: 10, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                        <Tooltip formatter={(v) => fmtUSD2(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                        <Bar dataKey="value" radius={[0, 3, 3, 0]}>
                          {selectedLoteByItem.map((_, i) => <Cell key={i} fill={FALLBACK_COLORS[i % FALLBACK_COLORS.length]} />)}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Movimientos del lote</div>
                    <div style={{ overflowX: "auto", maxHeight: 320, overflowY: "auto", border: "1px solid var(--line)", borderRadius: 6 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 480 }}>
                        <thead>
                          <tr>
                            {["Fecha", "Rubro", "Concepto", "Cant.", "USD"].map((h) => (
                              <th key={h} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {selectedLoteRows.map((r, i) => (
                            <tr key={i} className="agri-tr">
                              <td className="agri-td">{fmtDate(r["Fecha"])}</td>
                              <td className="agri-td">{r["Tipo item"]}</td>
                              <td className="agri-td">{r["Concepto"]}</td>
                              <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(r["Cantidad"], 1)} {r["Unid."]}</td>
                              <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{r["U$S/Total"] === null ? "—" : fmtUSD2(r["U$S/Total"])}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>
            )}
            {/* Tabla */}
            <div className="agri-card" style={{ padding: 0, overflow: "hidden" }}>
              <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Detalle de movimientos</div>
                <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(sorted.length)} registros filtrados</div>
              </div>
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
            </div>
          </>
        )}
      </div>
    </div>
  );
}
