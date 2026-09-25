import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, Legend, LabelList,
} from "recharts";
import {
  Upload, RefreshCw, FileSpreadsheet, AlertCircle, Trash2,
  ChevronDown, ChevronUp, Download, Search, X, Cloud, CloudOff, LogOut,
} from "lucide-react";
import { ref as dbRef, onValue, set as dbSet } from "firebase/database";
import { PdfMenu, generarPDF, filaGrupo, filaTotal } from "./pdfExport.jsx";

// ---------------------------------------------------------------------------
// Constantes
// ---------------------------------------------------------------------------
const COLS_CULT = [
  "Campaña","Admin","Campo","Lote","Cultivo","Variedad","Siembra",
  "Has.Semb.","Has.Act.","Trilla","Has.Tri.","% Trilla",
  "Neto (O)","Seco (O)","Desc (O)","Neto (D)","Seco (D)","Desc (D)",
  "Rinde","Rinde (Con.)",
];
const STORAGE_KEY_CULT = "cultivos-dataset-v1";

const CULTIVO_COLORS = { SOJA: "#5B7C4B", MAIZ: "#C68F41", POROTO: "#35606B", GARBANZO: "#8A5A3B" };
const FALLBACK_COLORS = ["#5B7C4B","#C68F41","#35606B","#8A5A3B","#7C8A4B","#A1462F"];
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
const toISODate = (v) => {
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getUTCFullYear(), m = String(v.getUTCMonth()+1).padStart(2,"0"), d = String(v.getUTCDate()).padStart(2,"0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "string" && v) {
    const d = new Date(v);
    if (!isNaN(d)) return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;
  }
  return "";
};
const fmtNum = (n, d=0) => new Intl.NumberFormat("es-AR",{maximumFractionDigits:d,minimumFractionDigits:d}).format(n||0);
const fmtDate = (iso) => {
  if (!iso) return "—";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
const fmtDateShort = (iso) => {
  if (!iso) return "";
  const [, m, d] = iso.split("-");
  return `${d}/${m}`;
};
// Orden natural de lotes: "2" < "10", "8A" < "8B", "N2" < "N11"
const cmpLote = (a, b) => String(a).localeCompare(String(b), "es", { numeric: true, sensitivity: "base" });

function normalizeRowCult(r) {
  return {
    "Campaña":      str(r["Campaña"]),
    "Admin":        str(r["Admin"]),
    "Campo":        str(r["Campo"]),
    "Lote":         str(r["Lote"]),
    "Cultivo":      str(r["Cultivo"]).toUpperCase(),
    "Variedad":     str(r["Variedad"]),
    "Siembra":      toISODate(r["Siembra"]),
    "Has.Semb.":    num(r["Has.Semb."]),
    "Has.Act.":     num(r["Has.Act."]),
    "Trilla":       toISODate(r["Trilla"]),
    "Has.Tri.":     num(r["Has.Tri."]),
    "% Trilla":     num(r["% Trilla"]),
    "Neto (O)":     num(r["Neto (O)"]),
    "Seco (O)":     num(r["Seco (O)"]),
    "Desc (O)":     num(r["Desc (O)"]),
    "Neto (D)":     num(r["Neto (D)"]),
    "Seco (D)":     num(r["Seco (D)"]),
    "Desc (D)":     num(r["Desc (D)"]),
    "Rinde":        num(r["Rinde"]),
    "Rinde (Con.)": num(r["Rinde (Con.)"]),
  };
}

const rowsFromPayload = (payload) => payload.rows.map((arr) => {
  const o = {};
  payload.columns.forEach((c, i) => { o[c] = arr[i] ?? (typeof arr[i] === "number" ? 0 : ""); });
  return o;
});

// Suma de un conjunto de filas → métricas. Los rindes se recalculan con sumas,
// nunca promediando rindes individuales:
//   Rinde        = Σ Neto (O) / Σ Has.Tri.   (kg por ha trillada)
//   Rinde (Con.) = Σ Desc (D) / Σ Has.Act.   (misma base que la planilla de origen)
function sumar(rows) {
  const t = { hasSemb: 0, hasAct: 0, hasTri: 0, netoO: 0, secoO: 0, descO: 0, netoD: 0, secoD: 0, descD: 0 };
  rows.forEach((r) => {
    t.hasSemb += r["Has.Semb."]; t.hasAct += r["Has.Act."]; t.hasTri += r["Has.Tri."];
    t.netoO += r["Neto (O)"]; t.secoO += r["Seco (O)"]; t.descO += r["Desc (O)"];
    t.netoD += r["Neto (D)"]; t.secoD += r["Seco (D)"]; t.descD += r["Desc (D)"];
  });
  t.pctTri   = t.hasAct > 0 ? (t.hasTri / t.hasAct) * 100 : 0;
  t.rinde    = t.hasTri > 0 ? t.netoO / t.hasTri : 0;
  t.rindeCon = t.hasAct > 0 ? t.descD / t.hasAct : 0;
  return t;
}

// ---------------------------------------------------------------------------
// Tooltips y etiquetas
// ---------------------------------------------------------------------------
const TooltipBox = ({ children }) => (
  <div style={{ background: "#FCFAF2", border: "1px solid #DCD2B8", borderRadius: 6, padding: "8px 10px", fontSize: 12 }}>{children}</div>
);

const RindeTooltip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>
      {payload.filter((p) => p.value).map((p, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ width: 8, height: 8, borderRadius: 2, background: p.color || p.payload?.color }} />
          <span>{p.name}: {fmtNum(p.value)} kg/ha</span>
        </div>
      ))}
    </TooltipBox>
  );
};

const VariedadTooltip = ({ active, payload }) => {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{d.name} · {d.cultivo}</div>
      <div>Rinde: {fmtNum(d.rinde)} kg/ha</div>
      <div style={{ color: "#6B5E4F" }}>{fmtNum(d.hasTri, 1)} ha trilladas · {fmtNum(d.netoO / 1000, 1)} tn</div>
    </TooltipBox>
  );
};

const AvanceTooltip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{fmtDate(payload[0].payload.fecha)}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ width: 10, height: 2, background: p.color }} />
          <span>{p.name}: {fmtNum(p.value, 1)} % · {fmtNum(p.payload[`${p.dataKey}__ha`], 0)} ha</span>
        </div>
      ))}
    </TooltipBox>
  );
};

const AvanceCampoTooltip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>
      <div>Trilladas: {fmtNum(d.trilladas, 1)} ha ({fmtNum(d.pct, 1)} %)</div>
      <div>Pendientes: {fmtNum(d.pendientes, 1)} ha</div>
    </TooltipBox>
  );
};

const HorizontalBarLabel = ({ x, y, width, height, value }) => {
  if (!value || width < 44) return null;
  return (
    <text x={x + width - 6} y={y + height / 2} textAnchor="end" dominantBaseline="middle" fill="#fff" fontSize={10} fontWeight={700}>
      {fmtNum(value)}
    </text>
  );
};

const legendText = (v) => <span style={{ color: "#2B2118" }}>{v}</span>;

// Barra de progreso de % trilla (usada en tabla)
const BarraTrilla = ({ pct }) => {
  const p = Math.max(0, Math.min(100, pct));
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, justifyContent: "flex-end" }}>
      <div style={{ width: 54, height: 6, background: "var(--line)", borderRadius: 3, overflow: "hidden" }}>
        <div style={{ width: `${p}%`, height: "100%", background: p >= 99.5 ? "var(--green)" : "var(--gold)" }} />
      </div>
      <span style={{ minWidth: 38, textAlign: "right" }}>{fmtNum(pct, 0)} %</span>
    </div>
  );
};

// ---------------------------------------------------------------------------
// Componente principal
// ---------------------------------------------------------------------------
export default function DashboardCultivos({ slug, isAdmin, cloudDb, onLogout, clienteNombre }) {
  const pdfRootRef = useRef(null);
  const storageKey = `${STORAGE_KEY_CULT}-${slug}`;
  const cloudPath = `clientes/${slug}/cultivos`;

  const [rows, setRows] = useState([]);
  const [meta, setMeta] = useState(null);
  const [bootLoading, setBootLoading] = useState(true);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState("");
  const [aviso, setAviso] = useState("");
  const [syncStatus, setSyncStatus] = useState(cloudDb ? "connecting" : "local");
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef(null);

  // Filtros
  const [buscar, setBuscar]     = useState("");
  const [campania, setCampania] = useState("");
  const [cultivo, setCultivo]   = useState("Todos");
  const [admin, setAdmin]       = useState("Todos");
  const [campo, setCampo]       = useState("Todos");
  const [variedad, setVariedad] = useState("Todos");

  // Tabla por campo
  const [camposAbiertos, setCamposAbiertos] = useState({});
  const [selectedUnidad, setSelectedUnidad] = useState(null);
  const toggleCampo = (c) => setCamposAbiertos((prev) => ({ ...prev, [c]: !prev[c] }));

  // Cargar caché local y suscribirse a Firebase
  useEffect(() => {
    let unsub = null;
    try {
      const raw = localStorage.getItem(storageKey);
      if (raw) {
        const payload = JSON.parse(raw);
        const r = rowsFromPayload(payload);
        setRows(r);
        setMeta({ fileName: payload.fileName, updatedAt: payload.updatedAt, rowCount: r.length });
      }
    } catch (e) {}
    if (cloudDb) {
      unsub = onValue(dbRef(cloudDb, cloudPath), (snap) => {
        const payload = snap.val();
        if (payload && payload.rows && payload.columns) {
          const r = rowsFromPayload(payload);
          setRows(r);
          setMeta({ fileName: payload.fileName, updatedAt: payload.updatedAt, rowCount: r.length });
        }
        setSyncStatus("cloud");
        setBootLoading(false);
      }, () => { setSyncStatus("error"); setBootLoading(false); });
    } else {
      setBootLoading(false);
    }
    return () => { if (unsub) unsub(); };
  }, [slug]);

  const persist = useCallback(async (all, fileName) => {
    const payload = {
      columns: COLS_CULT,
      rows: all.map((r) => COLS_CULT.map((c) => r[c] ?? "")),
      fileName, updatedAt: new Date().toISOString(),
    };
    try { localStorage.setItem(storageKey, JSON.stringify(payload)); } catch (e) {}
    if (cloudDb) {
      try { await dbSet(dbRef(cloudDb, cloudPath), payload); setSyncStatus("cloud"); }
      catch (e) { setSyncStatus("error"); throw e; }
    }
  }, [slug, cloudDb]);

  // Al subir un Excel se reemplazan sólo las campañas que contiene; las demás se conservan.
  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setParsing(true); setError(""); setAviso("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
      if (!json.length) throw new Error("La hoja no tiene datos.");
      const faltan = ["Campaña","Campo","Lote","Cultivo","Variedad","Has.Act.","Has.Tri.","Neto (O)","Desc (D)"].filter((c) => !(c in json[0]));
      if (faltan.length) throw new Error(`Faltan columnas: ${faltan.join(", ")}`);
      const nuevas = json.map(normalizeRowCult).filter((r) => r["Campo"] && r["Cultivo"]);
      const campNuevas = new Set(nuevas.map((r) => r["Campaña"]));
      const conservadas = rows.filter((r) => !campNuevas.has(r["Campaña"]));
      const all = [...conservadas, ...nuevas];
      setRows(all);
      setMeta({ fileName: file.name, updatedAt: new Date().toISOString(), rowCount: all.length });
      const lista = Array.from(campNuevas).sort();
      setCampania(lista[lista.length - 1] || "");
      setAviso(`Actualizada${lista.length > 1 ? "s" : ""}: ${lista.join(", ")}${conservadas.length ? " · se conservaron las demás campañas" : ""}`);
      try {
        await persist(all, file.name);
      } catch (e) {
        setAviso("");
        setError(`El archivo se leyó, pero no se pudo guardar en la nube (${e.code || e.message}). Sólo lo ves en este navegador.`);
      }
    } catch (e) {
      setError("No pude leer el archivo: " + e.message);
    } finally {
      setParsing(false);
    }
  }, [persist, rows]);

  const clearDataset = async () => {
    if (!window.confirm("¿Borrar todos los datos de Cultivos de este cliente?")) return;
    setRows([]); setMeta(null);
    try { localStorage.removeItem(storageKey); } catch (e) {}
    if (cloudDb) { try { await dbSet(dbRef(cloudDb, cloudPath), null); } catch (e) {} }
  };

  // ---------- Listas para filtros ----------
  const campanias = useMemo(() => Array.from(new Set(rows.map((r) => r["Campaña"]).filter(Boolean))).sort(), [rows]);
  // Por defecto: la campaña más reciente
  useEffect(() => {
    if (!campanias.length) return;
    if (campania === "" || (campania !== "Todas" && !campanias.includes(campania))) setCampania(campanias[campanias.length - 1]);
  }, [campanias]);

  const baseCamp = useMemo(() => (campania && campania !== "Todas" ? rows.filter((r) => r["Campaña"] === campania) : rows), [rows, campania]);
  const cultivos = useMemo(() => Array.from(new Set(baseCamp.map((r) => r["Cultivo"]).filter(Boolean))).sort(), [baseCamp]);
  const admins   = useMemo(() => Array.from(new Set(baseCamp.map((r) => r["Admin"]).filter(Boolean))).sort(), [baseCamp]);
  const campos   = useMemo(() => {
    const b = admin === "Todos" ? baseCamp : baseCamp.filter((r) => r["Admin"] === admin);
    return Array.from(new Set(b.map((r) => r["Campo"]).filter(Boolean))).sort();
  }, [baseCamp, admin]);
  const variedades = useMemo(() => {
    const b = cultivo === "Todos" ? baseCamp : baseCamp.filter((r) => r["Cultivo"] === cultivo);
    return Array.from(new Set(b.map((r) => r["Variedad"]).filter(Boolean))).sort();
  }, [baseCamp, cultivo]);
  useEffect(() => { if (campo !== "Todos" && !campos.includes(campo)) setCampo("Todos"); }, [campos]);
  useEffect(() => { if (variedad !== "Todos" && !variedades.includes(variedad)) setVariedad("Todos"); }, [variedades]);

  // Color fijo por cultivo según su posición en la lista completa (no cambia al filtrar)
  const colorDe = useMemo(() => {
    const todos = Array.from(new Set(rows.map((r) => r["Cultivo"]))).sort();
    const m = {};
    todos.forEach((c, i) => { m[c] = cultivoColor(c, i); });
    return (c) => m[c] || "#6B5E4F";
  }, [rows]);

  // ---------- Filtrado (fila a fila) ----------
  const buscarQ = buscar.trim().toLowerCase();
  const filtered = useMemo(() => baseCamp.filter((r) => {
    if (cultivo  !== "Todos" && r["Cultivo"]  !== cultivo)  return false;
    if (admin    !== "Todos" && r["Admin"]    !== admin)    return false;
    if (campo    !== "Todos" && r["Campo"]    !== campo)    return false;
    if (variedad !== "Todos" && r["Variedad"] !== variedad) return false;
    if (buscarQ) {
      const hay = ["Campo","Lote","Cultivo","Variedad","Admin"].some((k) => str(r[k]).toLowerCase().includes(buscarQ));
      if (!hay) return false;
    }
    return true;
  }), [baseCamp, cultivo, admin, campo, variedad, buscarQ]);

  // ---------- Unidades de análisis: Campo + Lote + Cultivo + Variedad ----------
  const unidades = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const key = `${r["Campaña"]}||${r["Campo"]}||${r["Lote"]}||${r["Cultivo"]}||${r["Variedad"]}`;
      if (!m.has(key)) m.set(key, { key, campania: r["Campaña"], admin: r["Admin"], campo: r["Campo"], lote: r["Lote"], cultivo: r["Cultivo"], variedad: r["Variedad"], registros: [] });
      m.get(key).registros.push(r);
    });
    return Array.from(m.values()).map((u) => {
      const t = sumar(u.registros);
      const siembras = u.registros.map((r) => r["Siembra"]).filter(Boolean).sort();
      const trillas  = u.registros.map((r) => r["Trilla"]).filter(Boolean).sort();
      return { ...u, ...t, siembra: siembras[0] || "", trilla: trillas[trillas.length - 1] || "" };
    });
  }, [filtered]);

  // ---------- KPIs ----------
  const tot = useMemo(() => sumar(filtered), [filtered]);
  const mermaPct = tot.netoO > 0 ? ((tot.netoO - tot.descD) / tot.netoO) * 100 : 0;
  const cultivosEnFiltro = useMemo(() => Array.from(new Set(filtered.map((r) => r["Cultivo"]))).sort(), [filtered]);
  const unSoloCultivo = cultivosEnFiltro.length === 1;

  // Rinde por cultivo (para KPI cuando hay varios cultivos, ya que mezclar rindes no tiene sentido)
  const rindePorCultivo = useMemo(() => cultivosEnFiltro.map((c) => ({ cultivo: c, ...sumar(filtered.filter((r) => r["Cultivo"] === c)) })), [filtered, cultivosEnFiltro]);

  // ---------- Avance de cosecha acumulado (% de has activas) por cultivo ----------
  const avance = useMemo(() => {
    const totalAct = {};
    cultivosEnFiltro.forEach((c) => { totalAct[c] = 0; });
    filtered.forEach((r) => { totalAct[r["Cultivo"]] += r["Has.Act."]; });
    const porFecha = new Map();
    filtered.forEach((r) => {
      if (!r["Trilla"] || !(r["Has.Tri."] > 0)) return;
      if (!porFecha.has(r["Trilla"])) porFecha.set(r["Trilla"], {});
      const b = porFecha.get(r["Trilla"]);
      b[r["Cultivo"]] = (b[r["Cultivo"]] || 0) + r["Has.Tri."];
    });
    const fechas = Array.from(porFecha.keys()).sort();
    const acum = {};
    cultivosEnFiltro.forEach((c) => { acum[c] = 0; });
    return fechas.map((f) => {
      const pt = { fecha: f, t: Date.parse(f + "T00:00:00Z") };
      cultivosEnFiltro.forEach((c) => {
        acum[c] += porFecha.get(f)[c] || 0;
        pt[c] = totalAct[c] > 0 ? (acum[c] / totalAct[c]) * 100 : 0;
        pt[`${c}__ha`] = acum[c];
      });
      return pt;
    });
  }, [filtered, cultivosEnFiltro]);

  // ---------- Avance por campo (has trilladas vs pendientes) ----------
  const avancePorCampo = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      if (!m.has(r["Campo"])) m.set(r["Campo"], []);
      m.get(r["Campo"]).push(r);
    });
    return Array.from(m.entries()).map(([name, rs]) => {
      const t = sumar(rs);
      return { name, trilladas: t.hasTri, pendientes: Math.max(0, t.hasAct - t.hasTri), pct: t.pctTri, hasAct: t.hasAct };
    }).sort((a, b) => b.hasAct - a.hasAct);
  }, [filtered]);

  // ---------- Rinde por campo, una barra por cultivo ----------
  const rindePorCampo = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      if (!m.has(r["Campo"])) m.set(r["Campo"], {});
      const b = m.get(r["Campo"]);
      if (!b[r["Cultivo"]]) b[r["Cultivo"]] = [];
      b[r["Cultivo"]].push(r);
    });
    return Array.from(m.entries()).map(([name, b]) => {
      const o = { name };
      Object.entries(b).forEach(([c, rs]) => { const t = sumar(rs); if (t.hasTri > 0) o[c] = Math.round(t.rinde); });
      return o;
    }).sort((a, b) => a.name.localeCompare(b.name));
  }, [filtered]);

  // ---------- Rinde por variedad ----------
  const rindePorVariedad = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const k = `${r["Cultivo"]}||${r["Variedad"]}`;
      if (!m.has(k)) m.set(k, { name: r["Variedad"], cultivo: r["Cultivo"], rs: [] });
      m.get(k).rs.push(r);
    });
    return Array.from(m.values()).map((v) => {
      const t = sumar(v.rs);
      return { name: v.name, cultivo: v.cultivo, rinde: Math.round(t.rinde), hasTri: t.hasTri, netoO: t.netoO, color: colorDe(v.cultivo) };
    }).filter((v) => v.hasTri > 0)
      .sort((a, b) => (a.cultivo === b.cultivo ? b.rinde - a.rinde : a.cultivo.localeCompare(b.cultivo)));
  }, [filtered, colorDe]);

  // ---------- Tabla agrupada por campo ----------
  const tablaCampos = useMemo(() => {
    const m = new Map();
    unidades.forEach((u) => {
      if (!m.has(u.campo)) m.set(u.campo, []);
      m.get(u.campo).push(u);
    });
    return Array.from(m.entries()).map(([c, us]) => {
      const regs = us.flatMap((u) => u.registros);
      const cults = Array.from(new Set(us.map((u) => u.cultivo)));
      return { campo: c, unidades: us.sort((a, b) => cmpLote(a.lote, b.lote) || a.cultivo.localeCompare(b.cultivo) || a.variedad.localeCompare(b.variedad)), tot: sumar(regs), unCultivo: cults.length === 1 };
    }).sort((a, b) => a.campo.localeCompare(b.campo));
  }, [unidades]);

  const exportToExcel = useCallback(() => {
    const data = unidades.map((u) => ({
      "Campaña": u.campania, "Admin": u.admin, "Campo": u.campo, "Lote": u.lote, "Cultivo": u.cultivo, "Variedad": u.variedad,
      "Siembra": fmtDate(u.siembra), "Has.Semb.": +u.hasSemb.toFixed(2), "Has.Act.": +u.hasAct.toFixed(2),
      "Trilla": u.trilla ? fmtDate(u.trilla) : "", "Has.Tri.": +u.hasTri.toFixed(2), "% Trilla": +u.pctTri.toFixed(2),
      "Neto (O)": +u.netoO.toFixed(2), "Seco (O)": +u.secoO.toFixed(2), "Desc (O)": +u.descO.toFixed(2),
      "Neto (D)": +u.netoD.toFixed(2), "Seco (D)": +u.secoD.toFixed(2), "Desc (D)": +u.descD.toFixed(2),
      "Rinde": +u.rinde.toFixed(2), "Rinde (Con.)": +u.rindeCon.toFixed(2),
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Cultivos");
    XLSX.writeFile(wb, `cultivos_filtrado_${new Date().toISOString().slice(0,10)}.xlsx`);
  }, [unidades]);

  // ---------- PDF ----------
  const exportarPDF = async (modo) => {
    const filtros = [];
    if (campania) filtros.push(["Campaña", campania === "Todas" ? "Todas" : campania]);
    if (cultivo !== "Todos") filtros.push(["Cultivo", cultivo]);
    if (admin !== "Todos") filtros.push(["Admin", admin]);
    if (campo !== "Todos") filtros.push(["Campo", campo]);
    if (variedad !== "Todos") filtros.push(["Variedad", variedad]);
    if (buscar.trim()) filtros.push(["Búsqueda", `"${buscar.trim()}"`]);
    const head = ["Lote", "Cultivo", "Variedad", "Siembra", "Has semb.", "Has act.", "Trilla", "Has tri.", "% Trilla", "Neto O (tn)", "Desc D (tn)", "Rinde", "Rinde con."];
    const align = [null, null, null, null, "right", "right", null, "right", "right", "right", "right", "right", "right"];
    const fila = (u) => [
      u.lote + (u.hasSemb - u.hasAct > 0.005 ? " (R)" : ""), u.cultivo, u.variedad, fmtDate(u.siembra),
      fmtNum(u.hasSemb, 2), fmtNum(u.hasAct, 2), u.trilla ? fmtDate(u.trilla) : "Pendiente", fmtNum(u.hasTri, 2),
      `${fmtNum(u.pctTri, 1)} %`, fmtNum(u.netoO / 1000, 1), fmtNum(u.descD / 1000, 1),
      u.hasTri > 0 ? fmtNum(u.rinde) : "-", u.hasAct > 0 && u.descD > 0 ? fmtNum(u.rindeCon) : "-",
    ];
    const body = [];
    tablaCampos.forEach(({ campo: c, unidades: us, tot: t, unCultivo }) => {
      body.push(filaGrupo(`${c} · ${us.length} unidades`, head.length, [
        "", fmtNum(t.hasAct, 1), "", fmtNum(t.hasTri, 1), `${fmtNum(t.pctTri, 1)} %`, fmtNum(t.netoO / 1000, 1), fmtNum(t.descD / 1000, 1),
        unCultivo && t.hasTri > 0 ? fmtNum(t.rinde) : "-", unCultivo && t.hasAct > 0 ? fmtNum(t.rindeCon) : "-",
      ]));
      us.forEach((u) => body.push(fila(u)));
    });
    body.push(filaTotal(["Total", "", "", "", fmtNum(tot.hasSemb, 1), fmtNum(tot.hasAct, 1), "", fmtNum(tot.hasTri, 1), `${fmtNum(tot.pctTri, 1)} %`,
      fmtNum(tot.netoO / 1000, 1), fmtNum(tot.descD / 1000, 1), unSoloCultivo ? fmtNum(tot.rinde) : "-", unSoloCultivo ? fmtNum(tot.rindeCon) : "-"]));
    await generarPDF({
      modo, titulo: "Cultivos", cliente: clienteNombre || slug, filtros, root: pdfRootRef.current, archivo: "Cultivos",
      tablas: [{
        titulo: "Cosecha por lote",
        nota: "(R) lote con resiembra. Rinde = Neto (O) / has trilladas · Rinde con. = Desc (D) / has activas. Pesos en toneladas.",
        head, align, body,
      }],
    });
  };

  const resetFiltros = () => { setBuscar(""); setCultivo("Todos"); setAdmin("Todos"); setCampo("Todos"); setVariedad("Todos"); };
  const hayFiltros = buscar !== "" || cultivo !== "Todos" || admin !== "Todos" || campo !== "Todos" || variedad !== "Todos";

  // -------------------------------------------------------------------------
  // Panel de detalle de una unidad (Lote · Cultivo · Variedad)
  // -------------------------------------------------------------------------
  const PanelDetalle = ({ u }) => {
    const resembrado = u.hasSemb - u.hasAct;
    const humO = u.netoO - u.secoO, otrosO = u.secoO - u.descO;
    const humD = u.netoD - u.secoD, otrosD = u.secoD - u.descD;
    const difTransito = u.netoD - u.netoO;
    const pct = (a, b) => (b > 0 ? `${fmtNum((a / b) * 100, 2)} %` : "—");
    const Kpi = ({ label, val, sub, color }) => (
      <div style={{ background: "var(--paper)", borderRadius: 6, padding: "8px 12px", border: "1px solid var(--line)" }}>
        <div className="agri-kpi-label">{label}</div>
        <div className="agri-kpi-value" style={{ fontSize: 17, color }}>{val}</div>
        {sub && <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>{sub}</div>}
      </div>
    );
    return (
      <tr>
        <td colSpan={9} style={{ padding: 0 }}>
          <div style={{ margin: "0 0 4px 32px", padding: 16, background: "var(--paper-raised)", borderLeft: "3px solid var(--green)", borderRadius: "0 0 8px 8px" }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 8, marginBottom: 14 }}>
              <Kpi label="Siembra" val={fmtDate(u.siembra)}
                sub={resembrado > 0.005 ? `Sembradas ${fmtNum(u.hasSemb, 2)} ha · resembradas ${fmtNum(resembrado, 2)} ha` : `${fmtNum(u.hasSemb, 2)} ha sembradas`} />
              <Kpi label="Has activas" val={`${fmtNum(u.hasAct, 2)} ha`} />
              <Kpi label="Trilla" val={u.trilla ? fmtDate(u.trilla) : "Pendiente"}
                sub={`${fmtNum(u.hasTri, 2)} ha trilladas · ${fmtNum(u.pctTri, 1)} %`} color={u.hasTri > 0 ? undefined : "var(--gold)"} />
              <Kpi label="Rinde" val={`${fmtNum(u.rinde)} kg/ha`} sub="Neto (O) / Has trilladas" color="var(--green)" />
              <Kpi label="Rinde consignado" val={`${fmtNum(u.rindeCon)} kg/ha`} sub="Desc (D) / Has activas" />
            </div>

            <div style={{ fontWeight: 600, fontSize: 12, marginBottom: 6 }}>Pesos: de origen a destino (kg)</div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 520 }}>
                <thead>
                  <tr>
                    <th className="agri-th" style={{ textAlign: "left" }}></th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Origen</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Destino</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Dif. D − O</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    ["Neto", u.netoO, u.netoD],
                    ["Seco (sin humedad)", u.secoO, u.secoD],
                    ["Descontado (consignado)", u.descO, u.descD],
                  ].map(([l, o, d]) => (
                    <tr key={l} className="agri-tr">
                      <td className="agri-td">{l}</td>
                      <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(o)}</td>
                      <td className="agri-td" style={{ textAlign: "right", fontWeight: l.startsWith("Descontado") ? 700 : undefined }}>{fmtNum(d)}</td>
                      <td className="agri-td" style={{ textAlign: "right", color: d - o < 0 ? "var(--rust)" : "var(--green)" }}>{d - o >= 0 ? "+" : ""}{fmtNum(d - o)}</td>
                    </tr>
                  ))}
                  <tr className="agri-tr">
                    <td className="agri-td" style={{ color: "var(--ink-soft)", fontSize: 12 }}>Merma por humedad</td>
                    <td className="agri-td" style={{ textAlign: "right", color: "var(--ink-soft)", fontSize: 12 }}>{fmtNum(humO)} · {pct(humO, u.netoO)}</td>
                    <td className="agri-td" style={{ textAlign: "right", color: "var(--ink-soft)", fontSize: 12 }}>{fmtNum(humD)} · {pct(humD, u.netoD)}</td>
                    <td className="agri-td" />
                  </tr>
                  <tr className="agri-tr">
                    <td className="agri-td" style={{ color: "var(--ink-soft)", fontSize: 12 }}>Otros descuentos</td>
                    <td className="agri-td" style={{ textAlign: "right", color: "var(--ink-soft)", fontSize: 12 }}>{fmtNum(otrosO)} · {pct(otrosO, u.netoO)}</td>
                    <td className="agri-td" style={{ textAlign: "right", color: "var(--ink-soft)", fontSize: 12 }}>{fmtNum(otrosD)} · {pct(otrosD, u.netoD)}</td>
                    <td className="agri-td" />
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 8 }}>
              Diferencia de pesaje en destino: {difTransito >= 0 ? "+" : ""}{fmtNum(difTransito)} kg ({pct(difTransito, u.netoO)} sobre Neto O).
              {u.registros.length > 1 && ` Esta unidad agrupa ${u.registros.length} registros de la planilla.`}
            </div>
          </div>
        </td>
      </tr>
    );
  };

  // Sub-header común a todos los dashboards: datos, estado de sincronización y controles de carga
  const subHeader = (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
      <div className="agri-sub">
        {meta ? (
          <>Datos de <strong>{meta.fileName}</strong> · {fmtNum(meta.rowCount)} registros · actualizado {new Date(meta.updatedAt).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" })}</>
        ) : isAdmin ? "Cargá tu planilla para empezar a consultar los datos" : "Todavía no hay datos cargados para este cliente."}
        {aviso && <span style={{ color: "var(--green)", marginLeft: 8 }}>· {aviso}</span>}
        {error && meta && <span style={{ color: "var(--rust)", marginLeft: 8 }}>· {error}</span>}
      </div>
      {cloudDb && (
        <div style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 600, padding: "3px 9px", borderRadius: 999, background: syncStatus === "cloud" ? "rgba(75,107,58,0.12)" : syncStatus === "error" ? "rgba(161,70,47,0.12)" : "rgba(107,94,79,0.12)", color: syncStatus === "cloud" ? "var(--green)" : syncStatus === "error" ? "var(--rust)" : "var(--ink-soft)" }}>
          {syncStatus === "cloud" ? <><Cloud size={12} /> Sincronizado</> : syncStatus === "error" ? <><CloudOff size={12} /> Sin conexión</> : <><Cloud size={12} /> Conectando…</>}
        </div>
      )}
      <div style={{ display: "flex", gap: 8 }}>
        {meta && rows.length > 0 && <PdfMenu onExport={exportarPDF} />}
        {isAdmin && (<>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
            onChange={(e) => { handleFile(e.target.files?.[0]); e.target.value = ""; }} />
          {meta && (
            <>
              <button className="agri-btn" disabled={parsing} onClick={() => fileInputRef.current?.click()}>
                <RefreshCw size={14} /> {parsing ? "Procesando…" : "Actualizar datos"}
              </button>
              <button className="agri-btn agri-btn-outline" onClick={clearDataset}>
                <Trash2 size={14} />
              </button>
            </>
          )}
          {onLogout && (
            <button className="agri-btn agri-btn-outline" onClick={onLogout} title="Cerrar sesión">
              <LogOut size={14} />
            </button>
          )}
        </>)}
      </div>
    </div>
  );

  // -------------------------------------------------------------------------
  return (
    <div className="agri-shell" ref={pdfRootRef}>
      {!bootLoading && subHeader}
      {bootLoading ? (
        <div style={{ padding: 60, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>
      ) : !meta || rows.length === 0 ? (
        isAdmin ? (
          <div
            className="agri-dropzone" data-drag={dragOver}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files?.[0]); }}
          >
            <FileSpreadsheet size={40} color="var(--ink-soft)" />
            <div style={{ fontSize: 16, fontWeight: 700 }}>{parsing ? "Leyendo…" : "Arrastrá el Excel de Cultivos"}</div>
            <div style={{ fontSize: 13, color: "var(--ink-soft)", maxWidth: 380 }}>
              Planilla con Campaña, Campo, Lote, Cultivo, Variedad, Has.Act., Has.Tri., pesos y rindes.
            </div>
            <button className="agri-btn" disabled={parsing} onClick={() => fileInputRef.current?.click()}>
              <Upload size={14} /> {parsing ? "Procesando…" : "Elegir archivo"}
            </button>
            {error && <div style={{ color: "var(--rust)", fontSize: 13 }}><AlertCircle size={13} style={{ display: "inline", marginRight: 4 }} />{error}</div>}
          </div>
        ) : (
          <div className="agri-card" style={{ padding: 56, textAlign: "center" }}>
            <FileSpreadsheet size={36} color="var(--ink-soft)" style={{ marginBottom: 10 }} />
            <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 6 }}>Todavía no hay datos de cultivos</div>
            <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Contactá a tu administrador para que suba la planilla.</div>
          </div>
        )
      ) : (
        <>
          {/* Filtros */}
          <div className="agri-card" style={{ padding: 14, marginBottom: 18 }}>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
              <button className="agri-chip" data-active={cultivo === "Todos"} onClick={() => setCultivo("Todos")}>Todos los cultivos</button>
              {cultivos.map((c) => <button key={c} className="agri-chip" data-active={cultivo === c} onClick={() => setCultivo(c)}>{c}</button>)}
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
              <div style={{ position: "relative", flex: "1 1 220px" }}>
                <Search size={14} style={{ position: "absolute", left: 9, top: 10, color: "var(--ink-soft)" }} />
                <input className="agri-input" style={{ width: "100%", paddingLeft: 30 }}
                  placeholder="Buscar campo, lote, variedad…" value={buscar} onChange={(e) => setBuscar(e.target.value)} />
              </div>
              <select className="agri-select" value={campania} onChange={(e) => setCampania(e.target.value)}>
                {campanias.map((c) => <option key={c} value={c}>{c}</option>)}
                {campanias.length > 1 && <option value="Todas">Todas las campañas</option>}
              </select>
              <select className="agri-select" value={admin} onChange={(e) => setAdmin(e.target.value)}>
                <option value="Todos">Todas las admins</option>
                {admins.map((a) => <option key={a} value={a}>{a}</option>)}
              </select>
              <select className="agri-select" value={campo} onChange={(e) => setCampo(e.target.value)}>
                <option value="Todos">Todos los campos</option>
                {campos.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
              <select className="agri-select" value={variedad} onChange={(e) => setVariedad(e.target.value)}>
                <option value="Todos">Todas las variedades</option>
                {variedades.map((v) => <option key={v} value={v}>{v}</option>)}
              </select>
              {hayFiltros && <button className="agri-btn agri-btn-outline" onClick={resetFiltros}><X size={13} /> Limpiar</button>}
            </div>
          </div>

          {filtered.length === 0 ? (
            <div className="agri-card" style={{ padding: 40, textAlign: "center", color: "var(--ink-soft)" }}>Ningún registro coincide con los filtros.</div>
          ) : (
          <>
          {/* KPIs */}
          <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 12, marginBottom: 20 }}>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Has activas</div>
              <div className="agri-kpi-value">{fmtNum(tot.hasAct, 0)}</div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>{unidades.length} lotes · cultivo · variedad</div>
            </div>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Avance de cosecha</div>
              <div className="agri-kpi-value">{fmtNum(tot.pctTri, 1)} %</div>
              <div style={{ height: 5, background: "var(--line)", borderRadius: 3, marginTop: 6, overflow: "hidden" }}>
                <div style={{ width: `${Math.min(100, tot.pctTri)}%`, height: "100%", background: "var(--green)" }} />
              </div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 4 }}>{fmtNum(tot.hasTri, 0)} ha trilladas · {fmtNum(Math.max(0, tot.hasAct - tot.hasTri), 0)} pendientes</div>
            </div>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Producción (Neto O)</div>
              <div className="agri-kpi-value">{fmtNum(tot.netoO / 1000, 1)} tn</div>
            </div>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Consignado (Desc D)</div>
              <div className="agri-kpi-value">{fmtNum(tot.descD / 1000, 1)} tn</div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>{fmtNum(mermaPct, 2)} % menos que Neto O</div>
            </div>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Rinde (kg/ha trillada)</div>
              {unSoloCultivo ? (
                <>
                  <div className="agri-kpi-value">{fmtNum(tot.rinde)}</div>
                  <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>Consignado: {fmtNum(tot.rindeCon)} kg/ha activa</div>
                </>
              ) : (
                <div style={{ marginTop: 4 }}>
                  {rindePorCultivo.map((c) => (
                    <div key={c.cultivo} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, lineHeight: 1.6 }}>
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                        <span style={{ width: 8, height: 8, borderRadius: 2, background: colorDe(c.cultivo) }} />{c.cultivo}
                      </span>
                      <span style={{ fontWeight: 700 }}>{c.hasTri > 0 ? fmtNum(c.rinde) : "—"}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>

          {/* Gráficos — fila 1: avance */}
          <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(340px,1fr))", gap: 14, marginBottom: 14 }}>
            <div className="agri-card" style={{ padding: 16 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 2 }}>Avance de cosecha</div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginBottom: 8 }}>% acumulado de has activas trilladas, por fecha de trilla</div>
              {avance.length === 0 ? (
                <div style={{ height: 230, display: "flex", alignItems: "center", justifyContent: "center", color: "var(--ink-soft)", fontSize: 13 }}>Todavía no hay trilla registrada.</div>
              ) : (
                <ResponsiveContainer width="100%" height={230}>
                  <LineChart data={avance} margin={{ left: -10, right: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                    <XAxis dataKey="t" type="number" scale="time" domain={["dataMin", "dataMax"]} tickFormatter={(t) => fmtDateShort(new Date(t).toISOString().slice(0, 10))} tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} minTickGap={24} />
                    <YAxis domain={[0, 100]} ticks={[0, 25, 50, 75, 100]} tickFormatter={(v) => `${v}%`} tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={false} tickLine={false} />
                    <Tooltip content={<AvanceTooltip />} cursor={{ stroke: "#6B5E4F", strokeDasharray: "3 3" }} />
                    <Legend wrapperStyle={{ fontSize: 11 }} formatter={legendText} />
                    {cultivosEnFiltro.map((c) => (
                      <Line key={c} type="stepAfter" dataKey={c} name={c} stroke={colorDe(c)} strokeWidth={2} dot={false} activeDot={{ r: 4 }} />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              )}
            </div>

            <div className="agri-card" style={{ padding: 16 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 2 }}>Avance por campo (ha)</div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginBottom: 8 }}>Trilladas y pendientes sobre has activas</div>
              <ResponsiveContainer width="100%" height={Math.max(230, avancePorCampo.length * 30 + 40)}>
                <BarChart data={avancePorCampo} layout="vertical" margin={{ left: 0, right: 10 }} barCategoryGap={6}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" width={60} tick={{ fontSize: 11, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                  <Tooltip content={<AvanceCampoTooltip />} cursor={{ fill: "rgba(107,94,79,0.08)" }} />
                  <Legend wrapperStyle={{ fontSize: 11 }} formatter={legendText} />
                  <Bar dataKey="trilladas" name="Trilladas" stackId="a" fill="#4B6B3A" stroke="#FCFAF2" strokeWidth={1}>
                    <LabelList dataKey="trilladas" content={<HorizontalBarLabel />} />
                  </Bar>
                  <Bar dataKey="pendientes" name="Pendientes" stackId="a" fill="#DCD2B8" stroke="#FCFAF2" strokeWidth={1} radius={[0,3,3,0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Gráficos — fila 2: rindes */}
          <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(340px,1fr))", gap: 14, marginBottom: 14 }}>
            <div className="agri-card" style={{ padding: 16 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 2 }}>Rinde por campo (kg/ha)</div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginBottom: 8 }}>Neto (O) / has trilladas, por cultivo</div>
              <ResponsiveContainer width="100%" height={250}>
                <BarChart data={rindePorCampo} margin={{ left: -6, right: 6 }} barGap={2}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                  <XAxis dataKey="name" tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                  <Tooltip content={<RindeTooltip />} cursor={{ fill: "rgba(107,94,79,0.08)" }} />
                  <Legend wrapperStyle={{ fontSize: 11 }} formatter={legendText} />
                  {cultivosEnFiltro.map((c) => (
                    <Bar key={c} dataKey={c} name={c} fill={colorDe(c)} radius={[3,3,0,0]} maxBarSize={28} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="agri-card" style={{ padding: 16 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 2 }}>Rinde por variedad (kg/ha)</div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginBottom: 8, display: "flex", gap: 10, flexWrap: "wrap" }}>
                <span>Neto (O) / has trilladas</span>
                {cultivosEnFiltro.map((c) => (
                  <span key={c} style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                    <span style={{ width: 8, height: 8, borderRadius: 2, background: colorDe(c) }} />{c}
                  </span>
                ))}
              </div>
              <ResponsiveContainer width="100%" height={Math.max(250, rindePorVariedad.length * 22 + 20)}>
                <BarChart data={rindePorVariedad} layout="vertical" margin={{ left: 0, right: 10 }} barCategoryGap={3}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" width={160} tick={{ fontSize: 10.5, fill: "#2B2118" }} axisLine={false} tickLine={false} interval={0} />
                  <Tooltip content={<VariedadTooltip />} cursor={{ fill: "rgba(107,94,79,0.08)" }} />
                  <Bar dataKey="rinde" name="Rinde" radius={[0,3,3,0]}>
                    {rindePorVariedad.map((v, i) => <Cell key={i} fill={v.color} />)}
                    <LabelList dataKey="rinde" content={<HorizontalBarLabel />} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Tabla por campo → lote · cultivo · variedad */}
          <div data-pdf="pantalla" className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Cosecha por lote</div>
              <div data-pdf-ignore style={{ display: "flex", gap: 8 }}>
                <button className="agri-btn agri-btn-outline" onClick={() => {
                  const todosAbiertos = tablaCampos.every((c) => camposAbiertos[c.campo]);
                  const o = {}; tablaCampos.forEach((c) => { o[c.campo] = !todosAbiertos; });
                  setCamposAbiertos(o); setSelectedUnidad(null);
                }}>
                  {tablaCampos.every((c) => camposAbiertos[c.campo]) ? "Colapsar todo" : "Expandir todo"}
                </button>
                <button className="agri-btn agri-btn-outline" onClick={exportToExcel}><Download size={13} /> Exportar</button>
              </div>
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
                <thead>
                  <tr>
                    <th className="agri-th" style={{ textAlign: "left" }}>Campo / Lote</th>
                    <th className="agri-th" style={{ textAlign: "left" }}>Cultivo</th>
                    <th className="agri-th" style={{ textAlign: "left" }}>Variedad</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Has act.</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>% Trilla</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Neto O (tn)</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Rinde</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Rinde con.</th>
                    <th className="agri-th" style={{ width: 30 }} />
                  </tr>
                </thead>
                <tbody>
                  {tablaCampos.map(({ campo: c, unidades: us, tot: t, unCultivo }) => {
                    const abierto = !!camposAbiertos[c];
                    return (
                      <React.Fragment key={c}>
                        <tr style={{ cursor: "pointer", background: "#F0EBE0" }} onClick={() => { toggleCampo(c); setSelectedUnidad(null); }}>
                          <td style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13 }} colSpan={3}>
                            <span style={{ color: "var(--ink-soft)", fontSize: 12, marginRight: 6 }}>{abierto ? "▾" : "▸"}</span>
                            {c} <span style={{ fontWeight: 400, color: "var(--ink-soft)", fontSize: 12 }}>· {us.length} unidades</span>
                          </td>
                          <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{fmtNum(t.hasAct, 1)}</td>
                          <td style={{ padding: "10px 12px", fontSize: 13, fontWeight: 700 }}><BarraTrilla pct={t.pctTri} /></td>
                          <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{fmtNum(t.netoO / 1000, 1)}</td>
                          <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }} title={unCultivo ? "" : "Varios cultivos: ver por lote"}>{unCultivo && t.hasTri > 0 ? fmtNum(t.rinde) : "—"}</td>
                          <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{unCultivo && t.hasAct > 0 ? fmtNum(t.rindeCon) : "—"}</td>
                          <td />
                        </tr>
                        {abierto && us.map((u) => {
                          const sel = selectedUnidad === u.key;
                          return (
                            <React.Fragment key={u.key}>
                              <tr className="agri-tr" style={{ cursor: "pointer" }} onClick={() => setSelectedUnidad(sel ? null : u.key)}>
                                <td className="agri-td" style={{ paddingLeft: 32, fontWeight: sel ? 700 : 500, color: sel ? "var(--green)" : undefined }}>
                                  {u.lote}
                                  {u.hasSemb - u.hasAct > 0.005 && <span title="Con resiembra" style={{ marginLeft: 6, fontSize: 10, color: "var(--gold)", border: "1px solid var(--gold)", borderRadius: 3, padding: "0 3px" }}>R</span>}
                                </td>
                                <td className="agri-td">
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                                    <span style={{ width: 8, height: 8, borderRadius: 2, background: colorDe(u.cultivo) }} />{u.cultivo}
                                  </span>
                                </td>
                                <td className="agri-td" style={{ fontSize: 12 }}>{u.variedad}</td>
                                <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(u.hasAct, 2)}</td>
                                <td className="agri-td"><BarraTrilla pct={u.pctTri} /></td>
                                <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(u.netoO / 1000, 1)}</td>
                                <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{u.hasTri > 0 ? fmtNum(u.rinde) : "—"}</td>
                                <td className="agri-td" style={{ textAlign: "right" }}>{u.hasAct > 0 && u.descD > 0 ? fmtNum(u.rindeCon) : "—"}</td>
                                <td className="agri-td" style={{ textAlign: "center", color: "var(--ink-soft)" }}>{sel ? <ChevronUp size={14} /> : <ChevronDown size={14} />}</td>
                              </tr>
                              {sel && <PanelDetalle u={u} />}
                            </React.Fragment>
                          );
                        })}
                        <tr><td colSpan={9} style={{ height: 4 }} /></tr>
                      </React.Fragment>
                    );
                  })}
                  <tr style={{ borderTop: "2px solid var(--line)" }}>
                    <td style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13 }} colSpan={3}>Total</td>
                    <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{fmtNum(tot.hasAct, 1)}</td>
                    <td style={{ padding: "10px 12px", fontSize: 13, fontWeight: 700 }}><BarraTrilla pct={tot.pctTri} /></td>
                    <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{fmtNum(tot.netoO / 1000, 1)}</td>
                    <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{unSoloCultivo ? fmtNum(tot.rinde) : "—"}</td>
                    <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{unSoloCultivo ? fmtNum(tot.rindeCon) : "—"}</td>
                    <td />
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 8 }}>
              <span style={{ color: "var(--gold)", border: "1px solid var(--gold)", borderRadius: 3, padding: "0 3px", marginRight: 4 }}>R</span>
              lote con resiembra (has activas menores a las sembradas). Rinde = Neto (O) / has trilladas · Rinde con. = Desc (D) / has activas.
            </div>
          </div>
          </>
          )}
        </>
      )}
    </div>
  );
}
