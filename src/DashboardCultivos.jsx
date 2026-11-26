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
import { leerFilasArchivo, leerEncabezado, leerCultivos, mapaColoresCultivos, colorDeCultivo, ACCEPT_ARCHIVOS } from "./lectorArchivo.js";

// ---------------------------------------------------------------------------
// Cultivos: datos generados por la pantalla "Dashboards > Cultivos" de Gorrión.
// Reemplaza por completo al dashboard de Cultivos anterior (17/11/2026): mismo nodo, otro formato.
// Una fila por siembra (o por habilitación del lote, si no tiene cultivo) de UNA campaña.
// Llega por tres vías y siempre se guarda en clientes/<slug>/cultivos:
//   - publicación directa desde Gorrión (PUT del JSON),
//   - "Actualizar datos" con el JSON que genera Gorrión,
//   - "Actualizar datos" con el Excel que exporta Gorrión.
// Reglas de cálculo acordadas con el usuario (17/11/2026):
//   - Sup Lote y Sup Plan se repiten en cada fila del lote: se cuentan UNA vez por Campo + Lote.
//   - Fila sin Cultivo = habilitación del lote: Siembra y Sup Semb. son fecha y superficie habilitadas.
//   - 0 y vacío significan lo mismo.
//   - Rinde = Neto (O) / Sup Trilla., sólo con superficie trillada; si no hay, el rinde queda vacío.
// ---------------------------------------------------------------------------
export const NODO_CULTIVOS = "cultivos";

const COLS = [
  "Admin", "Campo", "Lote", "Sup Lote", "Estado", "Cultivo Plan", "Variedad Plan", "Sup Plan",
  "Cultivo", "Variedad", "Siembra", "Sup Semb.", "Sup Superp.", "Sup Activa",
  "Trilla", "Sup Trilla.", "Neto (O)", "Neto (D)", "Seco (D)", "Desc (D)",
];
const COLS_NUM = new Set(["Sup Lote", "Sup Plan", "Sup Semb.", "Sup Superp.", "Sup Activa", "Sup Trilla.", "Neto (O)", "Neto (D)", "Seco (D)", "Desc (D)"]);
const COLS_FECHA = new Set(["Siembra", "Trilla"]);
const COLS_MAYUS = new Set(["Cultivo", "Cultivo Plan"]);
const REQUERIDAS = ["Campo", "Lote", "Sup Lote", "Cultivo", "Siembra", "Sup Semb.", "Sup Activa", "Sup Trilla.", "Neto (O)"];
const STORAGE_KEY = "cultivos-gorrion-v1";   // distinta de la del dashboard anterior: no lee su caché
const SIN_CULTIVO = "__SIN__";

const CULTIVO_COLORS = { SOJA: "#5B7C4B", MAIZ: "#C68F41", POROTO: "#35606B", GARBANZO: "#8A5A3B", TRIGO: "#B8842E", GIRASOL: "#A1462F" };
// Colores para cultivos sin color propio: distintos de los fijos de arriba, así dos cultivos no comparten color
const FALLBACK_COLORS = ["#7C8A4B", "#6E5A8A", "#3F7F6E", "#B26A8E", "#8C7A3A", "#4F6F8F"];
// Colores de los cultivos: primero los que manda Gorrión en la sección "cultivos" del JSON
// (administrador de cultivos de cada Gorrión); si no vino, la paleta fija y después la de respaldo.
let coloresGorrion = new Map();
const cultivoColor = (name, i) => colorDeCultivo(coloresGorrion, name) || CULTIVO_COLORS[String(name).toUpperCase()] || FALLBACK_COLORS[i % FALLBACK_COLORS.length];
const GRIS_HABILITADO = "#B9AE94";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const str = (v) => (v === null || v === undefined ? "" : String(v).trim());
// Números: acepta 1234.5, "1234,5", "1.234,5", "1,234.5" y textos con unidad ("45.0 Has")
const num = (v) => {
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  let s = String(v ?? "").replace(/[^0-9,.\-]/g, "");
  if (!s) return 0;
  const coma = s.lastIndexOf(","), punto = s.lastIndexOf(".");
  if (coma >= 0 && punto >= 0) s = coma > punto ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  else if (coma >= 0) s = s.replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};
// Fechas: ISO o Date de Excel. La fecha cero de Delphi (1899-12-30) y cualquier año < 1901 = vacía.
const toISODate = (v) => {
  let d = null;
  const dmy = typeof v === "string" ? v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/) : null;   // dd/mm/aaaa
  if (v instanceof Date) d = v;
  else if (dmy) d = new Date(Date.UTC(+dmy[3], +dmy[2] - 1, +dmy[1]));
  else if (typeof v === "string" && v.trim()) d = new Date(v.trim().length === 10 ? v.trim() + "T00:00:00Z" : v.trim());
  else if (typeof v === "number" && v > 0) d = new Date(Math.round((v - 25569) * 86400000)); // serial de Excel
  if (!d || isNaN(d) || d.getUTCFullYear() < 1901) return "";
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
};
const fmtNum = (n, d = 0) => new Intl.NumberFormat("es-AR", { maximumFractionDigits: d, minimumFractionDigits: d }).format(n || 0);
const fmtDate = (iso) => {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
const fmtDateShort = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y.slice(2)}`;
};
const pct = (a, b) => (b > 0 ? (a / b) * 100 : 0);
const cmpLote = (a, b) => String(a).localeCompare(String(b), "es", { numeric: true, sensitivity: "base" });

// Encabezados tolerantes: "Sup Trilla." = "Sup.Trilla" = "sup trilla"
const normKey = (s) => String(s).toLowerCase().replace(/[\s.]/g, "");
function mapaColumnas(filas) {
  const presentes = new Map();
  filas.slice(0, 50).forEach((f) => Object.keys(f).forEach((k) => { if (!presentes.has(normKey(k))) presentes.set(normKey(k), k); }));
  const mapa = {};
  COLS.forEach((c) => { mapa[c] = presentes.get(normKey(c)) || null; });
  return mapa;
}
function normalizar(filas) {
  const mapa = mapaColumnas(filas);
  return filas.map((r) => {
    const o = {};
    COLS.forEach((c) => {
      const v = mapa[c] ? r[mapa[c]] : null;
      if (COLS_NUM.has(c)) o[c] = num(v);
      else if (COLS_FECHA.has(c)) o[c] = toISODate(v);
      else o[c] = COLS_MAYUS.has(c) ? str(v).toUpperCase() : str(v);
    });
    return o;
  }).filter((r) => r["Campo"] && r["Lote"]);
}
function validar(filas) {
  if (!filas.length) throw new Error("el archivo no tiene filas de datos");
  const mapa = mapaColumnas(filas);
  const faltan = REQUERIDAS.filter((c) => !mapa[c]);
  if (faltan.length) throw new Error(`faltan columnas: ${faltan.join(", ")}`);
}
const filasDesdePayload = (payload) => normalizar(payload.rows.map((arr) => {
  const o = {};
  payload.columns.forEach((c, i) => { o[c] = arr[i]; });
  return o;
}));

// Totales de cosecha de un conjunto de filas con cultivo.
// El rinde se calcula sólo con las filas que tienen superficie trillada, para que los kilos
// cargados sin superficie trillada no inflen el resultado.
function sumarCosecha(rows) {
  const t = { activa: 0, haTri: 0, netoO: 0, netoD: 0, secoD: 0, descD: 0, netoOConTrilla: 0, kgSinTrilla: 0 };
  rows.forEach((r) => {
    t.activa += r["Sup Activa"]; t.haTri += r["Sup Trilla."];
    t.netoO += r["Neto (O)"]; t.netoD += r["Neto (D)"]; t.secoD += r["Seco (D)"]; t.descD += r["Desc (D)"];
    if (r["Sup Trilla."] > 0) t.netoOConTrilla += r["Neto (O)"];
    else if (r["Neto (O)"] > 0) t.kgSinTrilla += r["Neto (O)"];
  });
  t.pctTri = pct(t.haTri, t.activa);
  t.rinde = t.haTri > 0 ? t.netoOConTrilla / t.haTri : null;
  t.rindeCon = t.activa > 0 && t.descD > 0 ? t.descD / t.activa : null;
  return t;
}

// ---------------------------------------------------------------------------
// Tooltips y etiquetas
// ---------------------------------------------------------------------------
const TooltipBox = ({ children }) => (
  <div style={{ background: "#FCFAF2", border: "1px solid #DCD2B8", borderRadius: 6, padding: "8px 10px", fontSize: 12 }}>{children}</div>
);
const Punto = ({ color, linea }) => <span style={{ width: linea ? 10 : 8, height: linea ? 2 : 8, borderRadius: linea ? 0 : 2, background: color, display: "inline-block" }} />;

const AvanceTooltip = ({ active, payload, unidad }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{fmtDate(payload[0].payload.fecha)}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <Punto color={p.color} linea />
          <span>{p.name}: {unidad === "%" ? `${fmtNum(p.value, 1)} % · ${fmtNum(p.payload[`${p.dataKey}__ha`], 1)} ha` : `${fmtNum(p.value, 1)} ha`}</span>
        </div>
      ))}
    </TooltipBox>
  );
};
const BarrasHaTooltip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>
      {payload.filter((p) => p.value).map((p, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}><Punto color={p.color || p.fill} /><span>{p.name}: {fmtNum(p.value, 1)} ha</span></div>
      ))}
    </TooltipBox>
  );
};
const VariedadHaTooltip = ({ active, payload }) => {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{d.name}{d.cultivo ? ` · ${d.cultivo}` : ""}</div>
      <div>{fmtNum(d.ha, 1)} ha activas</div>
    </TooltipBox>
  );
};
const RindeTooltip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>
      {payload.filter((p) => p.value).map((p, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}><Punto color={p.color || p.payload?.color} /><span>{p.name}: {fmtNum(p.value)} kg/ha</span></div>
      ))}
    </TooltipBox>
  );
};
const VariedadRindeTooltip = ({ active, payload }) => {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  return (
    <TooltipBox>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{d.name} · {d.cultivo}</div>
      <div>Rinde: {fmtNum(d.rinde)} kg/ha</div>
      <div style={{ color: "#6B5E4F" }}>{fmtNum(d.haTri, 1)} ha trilladas · {fmtNum(d.netoO / 1000, 1)} tn</div>
    </TooltipBox>
  );
};
const HorizontalBarLabel = ({ x, y, width, height, value }) => {
  if (!value || width < 44) return null;
  return (
    <text x={x + width - 6} y={y + height / 2} textAnchor="end" dominantBaseline="middle" fill="#fff" fontSize={10} fontWeight={700}>{fmtNum(value)}</text>
  );
};
const legendText = (v) => <span style={{ color: "#2B2118" }}>{v}</span>;

const Barra = ({ valor }) => {
  const p = Math.max(0, Math.min(100, valor));
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, justifyContent: "flex-end" }}>
      <div style={{ width: 54, height: 6, background: "var(--line)", borderRadius: 3, overflow: "hidden" }}>
        <div style={{ width: `${p}%`, height: "100%", background: p >= 99.5 ? "var(--green)" : "var(--gold)" }} />
      </div>
      <span style={{ minWidth: 38, textAlign: "right" }}>{fmtNum(valor, 0)} %</span>
    </div>
  );
};
const Kpi = ({ label, value, sub, children }) => (
  <div className="agri-card" style={{ padding: 14 }}>
    <div className="agri-kpi-label">{label}</div>
    {value !== undefined && <div className="agri-kpi-value">{value}</div>}
    {children}
    {sub && <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 4 }}>{sub}</div>}
  </div>
);
const Titulo = ({ children, sub }) => (
  <>
    <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 2 }}>{children}</div>
    <div style={{ fontSize: 11, color: "var(--ink-soft)", marginBottom: 8 }}>{sub}</div>
  </>
);
const SinDatos = ({ alto = 230, children }) => (
  <div style={{ height: alto, display: "flex", alignItems: "center", justifyContent: "center", color: "var(--ink-soft)", fontSize: 13, textAlign: "center", padding: 12 }}>{children}</div>
);

// ---------------------------------------------------------------------------
// Componente principal
// ---------------------------------------------------------------------------
export default function DashboardCultivos({ slug, isAdmin, cloudDb, onLogout, clienteNombre }) {
  const pdfRootRef = useRef(null);
  const storageKey = `${STORAGE_KEY}-${slug}`;
  const cloudPath = `clientes/${slug}/${NODO_CULTIVOS}`;

  const [rows, setRows] = useState([]);
  const [meta, setMeta] = useState(null);
  const [bootLoading, setBootLoading] = useState(true);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState("");
  const [syncStatus, setSyncStatus] = useState(cloudDb ? "connecting" : "local");
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef(null);

  const [vista, setVista] = useState("siembra");
  const [buscar, setBuscar] = useState("");
  const [cultivo, setCultivo] = useState("Todos");
  const [admin, setAdmin] = useState("Todos");
  const [campo, setCampo] = useState("Todos");
  const [variedad, setVariedad] = useState("Todos");
  const [estado, setEstado] = useState("Todos");

  const [camposAbiertos, setCamposAbiertos] = useState({});
  const [seleccion, setSeleccion] = useState(null);
  const toggleCampo = (c) => setCamposAbiertos((p) => ({ ...p, [c]: !p[c] }));
  useEffect(() => { setSeleccion(null); }, [vista]);

  // ---------- Carga: caché local y luego Firebase en vivo ----------
  useEffect(() => {
    let unsub = null;
    const aplicar = (payload) => {
      const r = filasDesdePayload(payload);
      const cultivosGorrion = leerCultivos(payload.cultivos);
      coloresGorrion = mapaColoresCultivos(cultivosGorrion);
      setRows(r);
      setMeta({ fileName: payload.fileName || "", updatedAt: payload.updatedAt, rowCount: r.length, encabezado: leerEncabezado(payload.encabezado), cultivos: cultivosGorrion });
    };
    try {
      const raw = localStorage.getItem(storageKey);
      if (raw) aplicar(JSON.parse(raw));
    } catch (e) {}
    if (cloudDb) {
      unsub = onValue(dbRef(cloudDb, cloudPath), (snap) => {
        const payload = snap.val();
        if (payload && Array.isArray(payload.rows) && Array.isArray(payload.columns)) aplicar(payload);
        else { setRows([]); setMeta(null); }
        setSyncStatus("cloud");
        setBootLoading(false);
      }, () => { setSyncStatus("error"); setBootLoading(false); });
    } else {
      setBootLoading(false);
    }
    return () => { if (unsub) unsub(); };
  }, [slug]); // eslint-disable-line react-hooks/exhaustive-deps

  const persist = useCallback(async (all, fileName, encabezado, cultivos) => {
    const payload = {
      columns: COLS,
      rows: all.map((r) => COLS.map((c) => r[c] ?? "")),
      fileName, updatedAt: new Date().toISOString(),
      ...(encabezado ? { encabezado } : {}),
      ...(cultivos ? { cultivos } : {}),
    };
    try { localStorage.setItem(storageKey, JSON.stringify(payload)); } catch (e) {}
    if (cloudDb) {
      try { await dbSet(dbRef(cloudDb, cloudPath), payload); setSyncStatus("cloud"); }
      catch (e) { setSyncStatus("error"); throw e; }
    }
  }, [slug, cloudDb]); // eslint-disable-line react-hooks/exhaustive-deps

  // Al subir un archivo se reemplaza TODO (una sola campaña por perfil)
  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setParsing(true); setError("");
    try {
      const { filas, fileName, encabezado, cultivos } = await leerFilasArchivo(file, { hoja: "mayor" });
      validar(filas);
      const all = normalizar(filas);
      coloresGorrion = mapaColoresCultivos(cultivos);   // un Excel no trae colores: vuelve a la paleta
      setRows(all);
      setMeta({ fileName, updatedAt: new Date().toISOString(), rowCount: all.length, encabezado, cultivos });
      setCultivo("Todos"); setAdmin("Todos"); setCampo("Todos"); setVariedad("Todos"); setEstado("Todos"); setBuscar("");
      try { await persist(all, fileName, encabezado, cultivos); }
      catch (e) { setError(`El archivo se leyó, pero no se pudo guardar en la nube (${e.code || e.message}). Sólo lo ves en este navegador.`); }
    } catch (e) {
      setError("No pude leer el archivo: " + e.message);
    } finally {
      setParsing(false);
    }
  }, [persist]);

  const clearDataset = async () => {
    if (!window.confirm("¿Borrar todos los datos de Cultivos de este perfil?")) return;
    setRows([]); setMeta(null);
    try { localStorage.removeItem(storageKey); } catch (e) {}
    if (cloudDb) { try { await dbSet(dbRef(cloudDb, cloudPath), null); } catch (e) {} }
  };

  // ---------- Listas para filtros ----------
  const cultivos = useMemo(() => Array.from(new Set(rows.map((r) => r["Cultivo"]).filter(Boolean))).sort(), [rows]);
  const haySinCultivo = useMemo(() => rows.some((r) => !r["Cultivo"]), [rows]);
  const admins = useMemo(() => Array.from(new Set(rows.map((r) => r["Admin"]).filter(Boolean))).sort(), [rows]);
  const estados = useMemo(() => Array.from(new Set(rows.map((r) => r["Estado"]).filter(Boolean))).sort(), [rows]);
  const campos = useMemo(() => {
    const b = admin === "Todos" ? rows : rows.filter((r) => r["Admin"] === admin);
    return Array.from(new Set(b.map((r) => r["Campo"]).filter(Boolean))).sort();
  }, [rows, admin]);
  const variedades = useMemo(() => {
    const b = cultivo === "Todos" ? rows : rows.filter((r) => r["Cultivo"] === cultivo);
    return Array.from(new Set(b.map((r) => r["Variedad"]).filter(Boolean))).sort();
  }, [rows, cultivo]);
  useEffect(() => { if (campo !== "Todos" && !campos.includes(campo)) setCampo("Todos"); }, [campos]); // eslint-disable-line
  useEffect(() => { if (variedad !== "Todos" && !variedades.includes(variedad)) setVariedad("Todos"); }, [variedades]); // eslint-disable-line

  // Color fijo por cultivo (no cambia al filtrar)
  const colorDe = useMemo(() => {
    const m = {};
    cultivos.forEach((c, i) => { m[c] = cultivoColor(c, i); });
    return (c) => (c ? m[c] || "#6B5E4F" : GRIS_HABILITADO);
  }, [cultivos]);

  // ---------- Filtrado fila a fila ----------
  const buscarQ = buscar.trim().toLowerCase();
  const filtered = useMemo(() => rows.filter((r) => {
    if (cultivo === SIN_CULTIVO ? r["Cultivo"] : cultivo !== "Todos" && r["Cultivo"] !== cultivo) return false;
    if (admin !== "Todos" && r["Admin"] !== admin) return false;
    if (campo !== "Todos" && r["Campo"] !== campo) return false;
    if (variedad !== "Todos" && r["Variedad"] !== variedad) return false;
    if (estado !== "Todos" && r["Estado"] !== estado) return false;
    if (buscarQ && !["Campo", "Lote", "Cultivo", "Variedad", "Admin", "Estado"].some((k) => str(r[k]).toLowerCase().includes(buscarQ))) return false;
    return true;
  }), [rows, cultivo, admin, campo, variedad, estado, buscarQ]);

  const conCultivo = useMemo(() => filtered.filter((r) => r["Cultivo"]), [filtered]);
  const cultivosEnFiltro = useMemo(() => Array.from(new Set(conCultivo.map((r) => r["Cultivo"]))).sort(), [conCultivo]);

  // ---------- Lotes (Campo + Lote): superficie del lote y planificada, una sola vez ----------
  const lotes = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const key = `${r["Campo"]}||${r["Lote"]}`;
      if (!m.has(key)) m.set(key, { key, campo: r["Campo"], lote: r["Lote"], admins: new Set(), estados: new Set(), supLote: 0, supPlan: 0, cultivoPlan: "", variedadPlan: "", filas: [] });
      const l = m.get(key);
      l.filas.push(r);
      if (r["Admin"]) l.admins.add(r["Admin"]);
      if (r["Estado"]) l.estados.add(r["Estado"]);
      if (!l.supLote && r["Sup Lote"]) l.supLote = r["Sup Lote"];
      if (!l.supPlan && r["Sup Plan"]) l.supPlan = r["Sup Plan"];
      if (!l.cultivoPlan && r["Cultivo Plan"]) { l.cultivoPlan = r["Cultivo Plan"]; l.variedadPlan = r["Variedad Plan"]; }
    });
    return Array.from(m.values()).map((l) => {
      const sembrada = l.filas.filter((r) => r["Cultivo"]);
      const habil = l.filas.filter((r) => !r["Cultivo"]);
      const activa = sembrada.reduce((s, r) => s + r["Sup Activa"], 0);
      const habilitada = habil.reduce((s, r) => s + r["Sup Activa"], 0);
      return {
        ...l, admins: Array.from(l.admins), estados: Array.from(l.estados),
        activa, habilitada,
        semb: sembrada.reduce((s, r) => s + r["Sup Semb."], 0),
        superp: sembrada.reduce((s, r) => s + r["Sup Superp."], 0),
        libre: Math.max(0, l.supLote - activa - habilitada),
        pctSemb: pct(activa, l.supLote),
        cultivos: Array.from(new Set(sembrada.map((r) => r["Cultivo"]))),
      };
    });
  }, [filtered]);

  // ---------- Siembra ----------
  const siembra = useMemo(() => {
    const t = { supLote: 0, supPlan: 0, activa: 0, semb: 0, superp: 0, habilitada: 0, libre: 0 };
    lotes.forEach((l) => {
      t.supLote += l.supLote; t.supPlan += l.supPlan; t.activa += l.activa; t.semb += l.semb;
      t.superp += l.superp; t.habilitada += l.habilitada; t.libre += l.libre;
    });
    t.pctSemb = pct(t.activa, t.supLote);
    t.lotes = lotes.length;
    t.campos = new Set(lotes.map((l) => l.campo)).size;
    t.lotesSinCultivo = lotes.filter((l) => !l.cultivos.length).length;
    t.porCultivo = cultivosEnFiltro.map((c) => ({ cultivo: c, ha: conCultivo.filter((r) => r["Cultivo"] === c).reduce((s, r) => s + r["Sup Activa"], 0) }));
    return t;
  }, [lotes, conCultivo, cultivosEnFiltro]);

  // Avance de siembra: ha sembradas acumuladas por fecha de siembra, por cultivo
  const avanceSiembra = useMemo(() => {
    const porFecha = new Map();
    conCultivo.forEach((r) => {
      if (!r["Siembra"] || !(r["Sup Semb."] > 0)) return;
      if (!porFecha.has(r["Siembra"])) porFecha.set(r["Siembra"], {});
      const b = porFecha.get(r["Siembra"]);
      b[r["Cultivo"]] = (b[r["Cultivo"]] || 0) + r["Sup Semb."];
    });
    const acum = {};
    cultivosEnFiltro.forEach((c) => { acum[c] = 0; });
    return Array.from(porFecha.keys()).sort().map((f) => {
      const pt = { fecha: f, t: Date.parse(f + "T00:00:00Z") };
      cultivosEnFiltro.forEach((c) => { acum[c] += porFecha.get(f)[c] || 0; pt[c] = acum[c]; });
      return pt;
    });
  }, [conCultivo, cultivosEnFiltro]);

  // Superficie activa por variedad (y habilitada sin cultivo)
  const haPorVariedad = useMemo(() => {
    const m = new Map();
    conCultivo.forEach((r) => {
      const k = `${r["Cultivo"]}||${r["Variedad"]}`;
      if (!m.has(k)) m.set(k, { name: r["Variedad"] || "(sin variedad)", cultivo: r["Cultivo"], ha: 0 });
      m.get(k).ha += r["Sup Activa"];
    });
    const lista = Array.from(m.values()).filter((v) => v.ha > 0)
      .sort((a, b) => (a.cultivo === b.cultivo ? b.ha - a.ha : a.cultivo.localeCompare(b.cultivo)))
      .map((v) => ({ ...v, color: colorDe(v.cultivo) }));
    if (siembra.habilitada > 0) lista.push({ name: "Habilitado sin cultivo", cultivo: "", ha: siembra.habilitada, color: GRIS_HABILITADO });
    return lista;
  }, [conCultivo, siembra.habilitada, colorDe]);

  // Ocupación por campo: sembrada activa, habilitada sin cultivo y libre
  const ocupacionPorCampo = useMemo(() => {
    const m = new Map();
    lotes.forEach((l) => {
      if (!m.has(l.campo)) m.set(l.campo, { name: l.campo, sembrada: 0, habilitada: 0, libre: 0, supLote: 0 });
      const c = m.get(l.campo);
      c.sembrada += l.activa; c.habilitada += l.habilitada; c.libre += l.libre; c.supLote += l.supLote;
    });
    return Array.from(m.values()).sort((a, b) => b.supLote - a.supLote);
  }, [lotes]);

  // Tabla de siembra: Campo > Lote
  const tablaSiembra = useMemo(() => {
    const m = new Map();
    lotes.forEach((l) => { if (!m.has(l.campo)) m.set(l.campo, []); m.get(l.campo).push(l); });
    return Array.from(m.entries()).map(([c, ls]) => {
      const t = ls.reduce((a, l) => ({ supLote: a.supLote + l.supLote, supPlan: a.supPlan + l.supPlan, activa: a.activa + l.activa, habilitada: a.habilitada + l.habilitada }), { supLote: 0, supPlan: 0, activa: 0, habilitada: 0 });
      return { campo: c, lotes: ls.sort((a, b) => cmpLote(a.lote, b.lote)), tot: { ...t, pctSemb: pct(t.activa, t.supLote) } };
    }).sort((a, b) => a.campo.localeCompare(b.campo));
  }, [lotes]);

  // ---------- Cosecha (sólo filas con cultivo) ----------
  const cosecha = useMemo(() => sumarCosecha(conCultivo), [conCultivo]);
  const unSoloCultivo = cultivosEnFiltro.length === 1;
  const rindePorCultivo = useMemo(() => cultivosEnFiltro.map((c) => ({ cultivo: c, ...sumarCosecha(conCultivo.filter((r) => r["Cultivo"] === c)) })), [conCultivo, cultivosEnFiltro]);
  const mermaPct = cosecha.netoO > 0 ? ((cosecha.netoO - cosecha.descD) / cosecha.netoO) * 100 : 0;

  const unidades = useMemo(() => {
    const m = new Map();
    conCultivo.forEach((r) => {
      const key = `${r["Campo"]}||${r["Lote"]}||${r["Cultivo"]}||${r["Variedad"]}`;
      if (!m.has(key)) m.set(key, { key, admin: r["Admin"], campo: r["Campo"], lote: r["Lote"], cultivo: r["Cultivo"], variedad: r["Variedad"], registros: [] });
      m.get(key).registros.push(r);
    });
    return Array.from(m.values()).map((u) => {
      const siembras = u.registros.map((r) => r["Siembra"]).filter(Boolean).sort();
      const trillas = u.registros.map((r) => r["Trilla"]).filter(Boolean).sort();
      return {
        ...u, ...sumarCosecha(u.registros),
        semb: u.registros.reduce((s, r) => s + r["Sup Semb."], 0), superp: u.registros.reduce((s, r) => s + r["Sup Superp."], 0),
        fechaSiembra: siembras[0] || "", fechaTrilla: trillas[trillas.length - 1] || "",
      };
    });
  }, [conCultivo]);

  const avanceCosecha = useMemo(() => {
    const totalAct = {};
    cultivosEnFiltro.forEach((c) => { totalAct[c] = 0; });
    conCultivo.forEach((r) => { totalAct[r["Cultivo"]] += r["Sup Activa"]; });
    const porFecha = new Map();
    conCultivo.forEach((r) => {
      if (!r["Trilla"] || !(r["Sup Trilla."] > 0)) return;
      if (!porFecha.has(r["Trilla"])) porFecha.set(r["Trilla"], {});
      const b = porFecha.get(r["Trilla"]);
      b[r["Cultivo"]] = (b[r["Cultivo"]] || 0) + r["Sup Trilla."];
    });
    const acum = {};
    cultivosEnFiltro.forEach((c) => { acum[c] = 0; });
    return Array.from(porFecha.keys()).sort().map((f) => {
      const pt = { fecha: f, t: Date.parse(f + "T00:00:00Z") };
      cultivosEnFiltro.forEach((c) => {
        acum[c] += porFecha.get(f)[c] || 0;
        pt[c] = pct(acum[c], totalAct[c]);
        pt[`${c}__ha`] = acum[c];
      });
      return pt;
    });
  }, [conCultivo, cultivosEnFiltro]);
  const trillaSinFecha = useMemo(() => conCultivo.filter((r) => r["Sup Trilla."] > 0 && !r["Trilla"]).reduce((s, r) => s + r["Sup Trilla."], 0), [conCultivo]);

  const avancePorCampo = useMemo(() => {
    const m = new Map();
    conCultivo.forEach((r) => { if (!m.has(r["Campo"])) m.set(r["Campo"], []); m.get(r["Campo"]).push(r); });
    return Array.from(m.entries()).map(([name, rs]) => {
      const t = sumarCosecha(rs);
      return { name, trilladas: t.haTri, pendientes: Math.max(0, t.activa - t.haTri), pct: t.pctTri, activa: t.activa };
    }).sort((a, b) => b.activa - a.activa);
  }, [conCultivo]);

  const rindePorCampo = useMemo(() => {
    const m = new Map();
    conCultivo.forEach((r) => {
      if (!m.has(r["Campo"])) m.set(r["Campo"], {});
      const b = m.get(r["Campo"]);
      (b[r["Cultivo"]] = b[r["Cultivo"]] || []).push(r);
    });
    return Array.from(m.entries()).map(([name, b]) => {
      const o = { name };
      Object.entries(b).forEach(([c, rs]) => { const t = sumarCosecha(rs); if (t.rinde !== null) o[c] = Math.round(t.rinde); });
      return o;
    }).filter((o) => Object.keys(o).length > 1).sort((a, b) => a.name.localeCompare(b.name));
  }, [conCultivo]);

  const rindePorVariedad = useMemo(() => {
    const m = new Map();
    conCultivo.forEach((r) => {
      const k = `${r["Cultivo"]}||${r["Variedad"]}`;
      if (!m.has(k)) m.set(k, { name: r["Variedad"] || "(sin variedad)", cultivo: r["Cultivo"], rs: [] });
      m.get(k).rs.push(r);
    });
    return Array.from(m.values()).map((v) => {
      const t = sumarCosecha(v.rs);
      return { name: v.name, cultivo: v.cultivo, rinde: t.rinde === null ? null : Math.round(t.rinde), haTri: t.haTri, netoO: t.netoOConTrilla, color: colorDe(v.cultivo) };
    }).filter((v) => v.rinde !== null)
      .sort((a, b) => (a.cultivo === b.cultivo ? b.rinde - a.rinde : a.cultivo.localeCompare(b.cultivo)));
  }, [conCultivo, colorDe]);

  const tablaCosecha = useMemo(() => {
    const m = new Map();
    unidades.forEach((u) => { if (!m.has(u.campo)) m.set(u.campo, []); m.get(u.campo).push(u); });
    return Array.from(m.entries()).map(([c, us]) => ({
      campo: c,
      unidades: us.sort((a, b) => cmpLote(a.lote, b.lote) || a.cultivo.localeCompare(b.cultivo) || a.variedad.localeCompare(b.variedad)),
      tot: sumarCosecha(us.flatMap((u) => u.registros)),
      unCultivo: new Set(us.map((u) => u.cultivo)).size === 1,
    })).sort((a, b) => a.campo.localeCompare(b.campo));
  }, [unidades]);

  // ---------- Exportar a Excel: las filas filtradas, con las columnas de Gorrión ----------
  const exportToExcel = useCallback(() => {
    const data = filtered.map((r) => {
      const o = {};
      COLS.forEach((c) => { o[c] = COLS_FECHA.has(c) ? (r[c] ? fmtDate(r[c]) : "") : r[c]; });
      return o;
    });
    const ws = XLSX.utils.json_to_sheet(data, { header: COLS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Cultivos");
    XLSX.writeFile(wb, `cultivos_filtrado_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, [filtered]);

  // ---------- PDF ----------
  const textoRinde = (v) => (v === null ? "-" : fmtNum(v));
  const exportarPDF = async (modo) => {
    const filtros = [];
    if (cultivo !== "Todos") filtros.push(["Cultivo", cultivo === SIN_CULTIVO ? "Sin cultivo (habilitado)" : cultivo]);
    if (admin !== "Todos") filtros.push(["Admin", admin]);
    if (campo !== "Todos") filtros.push(["Campo", campo]);
    if (variedad !== "Todos") filtros.push(["Variedad", variedad]);
    if (estado !== "Todos") filtros.push(["Estado", estado]);
    if (buscar.trim()) filtros.push(["Búsqueda", `"${buscar.trim()}"`]);

    const headS = ["Lote", "Estado", "Sup. lote", "Planificada", "Cultivo plan", "Sembrada activa", "Habilitada s/cult.", "% sembrado", "Cultivos"];
    const bodyS = [];
    tablaSiembra.forEach(({ campo: c, lotes: ls, tot: t }) => {
      bodyS.push(filaGrupo(`${c} · ${ls.length} lotes`, headS.length, [
        fmtNum(t.supLote, 1), fmtNum(t.supPlan, 1), "", fmtNum(t.activa, 1), fmtNum(t.habilitada, 1), `${fmtNum(t.pctSemb, 1)} %`, "",
      ]));
      ls.forEach((l) => bodyS.push([
        l.lote, l.estados.join(", "), fmtNum(l.supLote, 1), fmtNum(l.supPlan, 1), l.cultivoPlan ? `${l.cultivoPlan}${l.variedadPlan ? " " + l.variedadPlan : ""}` : "-",
        fmtNum(l.activa, 1), fmtNum(l.habilitada, 1), `${fmtNum(l.pctSemb, 1)} %`, l.cultivos.join(", ") || "-",
      ]));
    });
    bodyS.push(filaTotal(["Total", "", fmtNum(siembra.supLote, 1), fmtNum(siembra.supPlan, 1), "", fmtNum(siembra.activa, 1), fmtNum(siembra.habilitada, 1), `${fmtNum(siembra.pctSemb, 1)} %`, ""]));

    const headC = ["Lote", "Cultivo", "Variedad", "Siembra", "Has activas", "Trilla", "Has trill.", "% Trilla", "Neto O (tn)", "Desc D (tn)", "Rinde", "Rinde con."];
    const bodyC = [];
    tablaCosecha.forEach(({ campo: c, unidades: us, tot: t, unCultivo }) => {
      bodyC.push(filaGrupo(`${c} · ${us.length} unidades`, headC.length, [
        fmtNum(t.activa, 1), "", fmtNum(t.haTri, 1), `${fmtNum(t.pctTri, 1)} %`, fmtNum(t.netoO / 1000, 1), fmtNum(t.descD / 1000, 1),
        unCultivo ? textoRinde(t.rinde) : "-", unCultivo ? textoRinde(t.rindeCon) : "-",
      ]));
      us.forEach((u) => bodyC.push([
        u.lote + (u.superp > 0.005 ? " (R)" : ""), u.cultivo, u.variedad, fmtDate(u.fechaSiembra), fmtNum(u.activa, 2),
        u.fechaTrilla ? fmtDate(u.fechaTrilla) : u.haTri > 0 ? "Sin fecha" : "Pendiente", fmtNum(u.haTri, 2), `${fmtNum(u.pctTri, 1)} %`,
        fmtNum(u.netoO / 1000, 1), fmtNum(u.descD / 1000, 1), textoRinde(u.rinde), textoRinde(u.rindeCon),
      ]));
    });
    bodyC.push(filaTotal(["Total", "", "", "", fmtNum(cosecha.activa, 1), "", fmtNum(cosecha.haTri, 1), `${fmtNum(cosecha.pctTri, 1)} %`,
      fmtNum(cosecha.netoO / 1000, 1), fmtNum(cosecha.descD / 1000, 1), unSoloCultivo ? textoRinde(cosecha.rinde) : "-", unSoloCultivo ? textoRinde(cosecha.rindeCon) : "-"]));

    const tSiembra = { titulo: "Siembra por lote", nota: "Hectáreas. Superficie del lote y planificada contadas una vez por lote. Habilitada s/cult. = habilitada para sembrar, todavía sin cultivo.", head: headS, align: [null, null, "right", "right", null, "right", "right", "right", null], body: bodyS };
    const tCosecha = { titulo: "Cosecha por lote", nota: "(R) con superficie superpuesta (resiembra). Rinde = Neto (O) / has trilladas, sólo donde hay superficie trillada · Rinde con. = Desc (D) / has activas. Pesos en toneladas.", head: headC, align: [null, null, null, null, "right", null, "right", "right", "right", "right", "right", "right"], body: bodyC };
    await generarPDF({
      modo, titulo: `Cultivos · ${vista === "siembra" ? "Siembra" : "Cosecha"}`, cliente: clienteNombre || slug, filtros,
      root: pdfRootRef.current, archivo: "Cultivos", encabezado: meta?.encabezado,
      tablas: vista === "siembra" ? [tSiembra, tCosecha] : [tCosecha, tSiembra],
    });
  };

  const resetFiltros = () => { setBuscar(""); setCultivo("Todos"); setAdmin("Todos"); setCampo("Todos"); setVariedad("Todos"); setEstado("Todos"); };
  const hayFiltros = buscar !== "" || cultivo !== "Todos" || admin !== "Todos" || campo !== "Todos" || variedad !== "Todos" || estado !== "Todos";

  // -------------------------------------------------------------------------
  // Paneles de detalle
  // -------------------------------------------------------------------------
  const KpiChico = ({ label, val, sub, color }) => (
    <div style={{ background: "var(--paper)", borderRadius: 6, padding: "8px 12px", border: "1px solid var(--line)" }}>
      <div className="agri-kpi-label">{label}</div>
      <div className="agri-kpi-value" style={{ fontSize: 17, color }}>{val}</div>
      {sub && <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>{sub}</div>}
    </div>
  );

  const PanelLote = ({ l }) => (
    <tr>
      <td colSpan={9} style={{ padding: 0 }}>
        <div style={{ margin: "0 0 4px 32px", padding: 16, background: "var(--paper-raised)", borderLeft: "3px solid var(--green)", borderRadius: "0 0 8px 8px" }}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 8, marginBottom: 14 }}>
            <KpiChico label="Superficie del lote" val={`${fmtNum(l.supLote, 1)} ha`} sub={l.admins.join(", ")} />
            <KpiChico label="Planificada" val={`${fmtNum(l.supPlan, 1)} ha`} sub={l.cultivoPlan ? `${l.cultivoPlan}${l.variedadPlan ? " · " + l.variedadPlan : ""}` : "Sin cultivo planificado"} />
            <KpiChico label="Sembrada activa" val={`${fmtNum(l.activa, 1)} ha`} sub={l.superp > 0.005 ? `Sembradas ${fmtNum(l.semb, 1)} ha · superpuestas ${fmtNum(l.superp, 1)} ha` : `${fmtNum(l.pctSemb, 1)} % del lote`} color="var(--green)" />
            <KpiChico label="Habilitada sin cultivo" val={`${fmtNum(l.habilitada, 1)} ha`} sub={`Libre: ${fmtNum(l.libre, 1)} ha`} />
          </div>
          <div style={{ fontWeight: 600, fontSize: 12, marginBottom: 6 }}>Siembras y habilitaciones del lote</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 560 }}>
              <thead>
                <tr>{["Fecha", "Cultivo", "Variedad", "Estado", "Sembrada", "Superpuesta", "Activa"].map((h, i) => (
                  <th key={h} className="agri-th" style={{ textAlign: i >= 4 ? "right" : "left", cursor: "default" }}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {l.filas.slice().sort((a, b) => (a["Siembra"] || "").localeCompare(b["Siembra"] || "")).map((r, i) => (
                  <tr key={i} className="agri-tr">
                    <td className="agri-td" style={{ whiteSpace: "nowrap" }}>{fmtDate(r["Siembra"])}</td>
                    <td className="agri-td">
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 5, color: r["Cultivo"] ? undefined : "var(--ink-soft)" }}>
                        <Punto color={colorDe(r["Cultivo"])} />{r["Cultivo"] || "Habilitado"}
                      </span>
                    </td>
                    <td className="agri-td" style={{ fontSize: 12 }}>{r["Variedad"] || "—"}</td>
                    <td className="agri-td" style={{ fontSize: 12, color: "var(--ink-soft)" }}>{r["Estado"]}</td>
                    <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(r["Sup Semb."], 2)}</td>
                    <td className="agri-td" style={{ textAlign: "right", color: r["Sup Superp."] > 0 ? "var(--gold)" : "var(--ink-soft)" }}>{r["Sup Superp."] > 0 ? fmtNum(r["Sup Superp."], 2) : "—"}</td>
                    <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{fmtNum(r["Sup Activa"], 2)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 8 }}>Sin cultivo = habilitación del lote: la fecha y la superficie son las de habilitación.</div>
        </div>
      </td>
    </tr>
  );

  const PanelUnidad = ({ u }) => {
    const humD = u.netoD - u.secoD, otrosD = u.secoD - u.descD;
    const difTransito = u.netoD - u.netoO;
    const p = (a, b) => (b > 0 ? `${fmtNum((a / b) * 100, 2)} %` : "—");
    return (
      <tr>
        <td colSpan={9} style={{ padding: 0 }}>
          <div style={{ margin: "0 0 4px 32px", padding: 16, background: "var(--paper-raised)", borderLeft: "3px solid var(--green)", borderRadius: "0 0 8px 8px" }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 8, marginBottom: 14 }}>
              <KpiChico label="Siembra" val={fmtDate(u.fechaSiembra)} sub={u.superp > 0.005 ? `Sembradas ${fmtNum(u.semb, 2)} ha · superpuestas ${fmtNum(u.superp, 2)} ha` : `${fmtNum(u.semb, 2)} ha sembradas`} />
              <KpiChico label="Has activas" val={`${fmtNum(u.activa, 2)} ha`} />
              <KpiChico label="Trilla" val={u.fechaTrilla ? fmtDate(u.fechaTrilla) : u.haTri > 0 ? "Sin fecha" : "Pendiente"}
                sub={`${fmtNum(u.haTri, 2)} ha trilladas · ${fmtNum(u.pctTri, 1)} %`} color={u.haTri > 0 ? undefined : "var(--gold)"} />
              <KpiChico label="Rinde" val={u.rinde === null ? "—" : `${fmtNum(u.rinde)} kg/ha`} sub={u.rinde === null ? "Sin superficie trillada" : "Neto (O) / has trilladas"} color="var(--green)" />
              <KpiChico label="Rinde consignado" val={u.rindeCon === null ? "—" : `${fmtNum(u.rindeCon)} kg/ha`} sub="Desc (D) / has activas" />
            </div>
            <div style={{ fontWeight: 600, fontSize: 12, marginBottom: 6 }}>Pesos (kg)</div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 460 }}>
                <tbody>
                  {[
                    ["Neto en origen", u.netoO, null],
                    ["Neto en destino", u.netoD, `${difTransito >= 0 ? "+" : ""}${fmtNum(difTransito)} kg (${p(difTransito, u.netoO)}) contra origen`],
                    ["Seco en destino", u.secoD, `Merma por humedad ${fmtNum(humD)} kg · ${p(humD, u.netoD)}`],
                    ["Descontado en destino (consignado)", u.descD, `Otros descuentos ${fmtNum(otrosD)} kg · ${p(otrosD, u.netoD)}`],
                  ].map(([l, v, nota]) => (
                    <tr key={l} className="agri-tr">
                      <td className="agri-td">{l}</td>
                      <td className="agri-td" style={{ textAlign: "right", fontWeight: l.startsWith("Descontado") ? 700 : 500 }}>{fmtNum(v)}</td>
                      <td className="agri-td" style={{ fontSize: 12, color: "var(--ink-soft)" }}>{nota || ""}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {(u.kgSinTrilla > 0 || u.registros.length > 1) && (
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 8 }}>
                {u.kgSinTrilla > 0 && <span style={{ color: "var(--rust)" }}>{fmtNum(u.kgSinTrilla)} kg cargados sin superficie trillada (no entran en el rinde). </span>}
                {u.registros.length > 1 && `Agrupa ${u.registros.length} siembras.`}
              </div>
            )}
          </div>
        </td>
      </tr>
    );
  };

  // -------------------------------------------------------------------------
  // Sub-header común: datos, sincronización y carga
  // -------------------------------------------------------------------------
  const subHeader = (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
      <div className="agri-sub">
        {meta ? (
          <>Datos de <strong>{meta.fileName}</strong> · {fmtNum(meta.rowCount)} registros{meta.updatedAt ? <> · actualizado {new Date(meta.updatedAt).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" })}</> : null}</>
        ) : isAdmin ? "Cargá el JSON o el Excel que genera Gorrión, o publicalo desde Gorrión" : "Todavía no hay datos cargados para este perfil."}
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
          <input ref={fileInputRef} type="file" accept={ACCEPT_ARCHIVOS} style={{ display: "none" }}
            onChange={(e) => { handleFile(e.target.files?.[0]); e.target.value = ""; }} />
          {meta && (
            <>
              <button className="agri-btn" disabled={parsing} onClick={() => fileInputRef.current?.click()}>
                <RefreshCw size={14} /> {parsing ? "Procesando…" : "Actualizar datos"}
              </button>
              <button className="agri-btn agri-btn-outline" onClick={clearDataset} title="Borrar los datos"><Trash2 size={14} /></button>
            </>
          )}
          {onLogout && <button className="agri-btn agri-btn-outline" onClick={onLogout} title="Cerrar sesión"><LogOut size={14} /></button>}
        </>)}
      </div>
    </div>
  );

  // -------------------------------------------------------------------------
  // Vista Siembra
  // -------------------------------------------------------------------------
  const vistaSiembra = (
    <>
      <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 12, marginBottom: 20 }}>
        <Kpi label="Superficie de los lotes" value={`${fmtNum(siembra.supLote, 0)} ha`} sub={`${siembra.campos} ${siembra.campos === 1 ? "campo" : "campos"} · ${siembra.lotes} ${siembra.lotes === 1 ? "lote" : "lotes"}`} />
        <Kpi label="Planificada" value={`${fmtNum(siembra.supPlan, 0)} ha`} sub={`${fmtNum(pct(siembra.supPlan, siembra.supLote), 1)} % de la superficie`} />
        <Kpi label="Sembrada activa" value={`${fmtNum(siembra.activa, 0)} ha`}
          sub={siembra.superp > 0.005 ? `${fmtNum(siembra.semb, 0)} ha sembradas · ${fmtNum(siembra.superp, 0)} superpuestas` : `${fmtNum(siembra.semb, 0)} ha sembradas`} />
        <Kpi label="Avance de siembra" value={`${fmtNum(siembra.pctSemb, 1)} %`} sub={`${fmtNum(siembra.habilitada, 0)} ha habilitadas sin cultivo · ${fmtNum(siembra.libre, 0)} libres`}>
          <div style={{ height: 5, background: "var(--line)", borderRadius: 3, marginTop: 6, overflow: "hidden", display: "flex" }}>
            <div style={{ width: `${Math.min(100, siembra.pctSemb)}%`, height: "100%", background: "var(--green)" }} />
            <div style={{ width: `${Math.min(100, pct(siembra.habilitada, siembra.supLote))}%`, height: "100%", background: GRIS_HABILITADO }} />
          </div>
        </Kpi>
        <Kpi label="Sembrada por cultivo">
          <div style={{ marginTop: 4 }}>
            {siembra.porCultivo.length === 0 ? <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Sin siembras</div> : siembra.porCultivo.map((c) => (
              <div key={c.cultivo} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, lineHeight: 1.6 }}>
                <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}><Punto color={colorDe(c.cultivo)} />{c.cultivo}</span>
                <span style={{ fontWeight: 700 }}>{fmtNum(c.ha, 0)} ha</span>
              </div>
            ))}
          </div>
        </Kpi>
      </div>

      <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(340px,1fr))", gap: 14, marginBottom: 14 }}>
        <div className="agri-card" style={{ padding: 16 }}>
          <Titulo sub="Hectáreas sembradas acumuladas, por fecha de siembra">Avance de siembra</Titulo>
          {avanceSiembra.length === 0 ? <SinDatos>Todavía no hay siembras registradas.</SinDatos> : (
            <ResponsiveContainer width="100%" height={230}>
              <LineChart data={avanceSiembra} margin={{ left: -4, right: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                <XAxis dataKey="t" type="number" scale="time" domain={["dataMin", "dataMax"]} tickFormatter={(t) => fmtDateShort(new Date(t).toISOString().slice(0, 10))} tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} minTickGap={28} />
                <YAxis tickFormatter={(v) => fmtNum(v)} tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={false} tickLine={false} />
                <Tooltip content={<AvanceTooltip unidad="ha" />} cursor={{ stroke: "#6B5E4F", strokeDasharray: "3 3" }} />
                <Legend wrapperStyle={{ fontSize: 11 }} formatter={legendText} />
                {cultivosEnFiltro.map((c) => <Line key={c} type="stepAfter" dataKey={c} name={c} stroke={colorDe(c)} strokeWidth={2} dot={false} activeDot={{ r: 4 }} />)}
              </LineChart>
            </ResponsiveContainer>
          )}
        </div>
        <div className="agri-card" style={{ padding: 16 }}>
          <Titulo sub="Sembrada activa, habilitada sin cultivo y libre, sobre la superficie de los lotes">Ocupación por campo (ha)</Titulo>
          <ResponsiveContainer width="100%" height={Math.max(170, ocupacionPorCampo.length * 36 + 70)}>
            <BarChart data={ocupacionPorCampo} layout="vertical" margin={{ left: 0, right: 10 }} barCategoryGap={6} maxBarSize={34}>
              <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
              <YAxis type="category" dataKey="name" width={120} tick={{ fontSize: 11, fill: "#2B2118" }} axisLine={false} tickLine={false} />
              <Tooltip content={<BarrasHaTooltip />} cursor={{ fill: "rgba(107,94,79,0.08)" }} />
              <Legend wrapperStyle={{ fontSize: 11 }} formatter={legendText} />
              <Bar dataKey="sembrada" name="Sembrada activa" stackId="a" fill="#4B6B3A" stroke="#FCFAF2" strokeWidth={1}><LabelList dataKey="sembrada" content={<HorizontalBarLabel />} /></Bar>
              <Bar dataKey="habilitada" name="Habilitada sin cultivo" stackId="a" fill={GRIS_HABILITADO} stroke="#FCFAF2" strokeWidth={1} />
              <Bar dataKey="libre" name="Libre" stackId="a" fill="#EDE6D3" stroke="#FCFAF2" strokeWidth={1} radius={[0, 3, 3, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div data-pdf="resumen" className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
        <Titulo sub="Hectáreas activas por variedad, agrupadas por cultivo">Superficie por variedad</Titulo>
        {haPorVariedad.length === 0 ? <SinDatos alto={120}>Sin superficie sembrada.</SinDatos> : (
          <ResponsiveContainer width="100%" height={Math.max(160, haPorVariedad.length * 24 + 20)}>
            <BarChart data={haPorVariedad} layout="vertical" margin={{ left: 0, right: 10 }} barCategoryGap={3} maxBarSize={26}>
              <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
              <YAxis type="category" dataKey="name" width={170} tick={{ fontSize: 10.5, fill: "#2B2118" }} axisLine={false} tickLine={false} interval={0} />
              <Tooltip content={<VariedadHaTooltip />} cursor={{ fill: "rgba(107,94,79,0.08)" }} />
              <Bar dataKey="ha" name="Ha activas" radius={[0, 3, 3, 0]}>
                {haPorVariedad.map((v, i) => <Cell key={i} fill={v.color} />)}
                <LabelList dataKey="ha" content={<HorizontalBarLabel />} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>

      <div data-pdf="pantalla" className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
          <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Siembra por lote</div>
          <div data-pdf-ignore style={{ display: "flex", gap: 8 }}>
            <BotonExpandir grupos={tablaSiembra.map((g) => g.campo)} abiertos={camposAbiertos} setAbiertos={setCamposAbiertos} onCambio={() => setSeleccion(null)} />
            <button className="agri-btn agri-btn-outline" onClick={exportToExcel}><Download size={13} /> Exportar</button>
          </div>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 820 }}>
            <thead>
              <tr>
                {[["Campo / Lote", "left"], ["Estado", "left"], ["Sup. lote", "right"], ["Planificada", "right"], ["Sembrada activa", "right"], ["Habilitada s/cult.", "right"], ["% sembrado", "right"], ["Cultivos", "left"], ["", "center"]].map(([h, a], i) => (
                  <th key={i} className="agri-th" style={{ textAlign: a, cursor: "default", width: h ? undefined : 30 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {tablaSiembra.map(({ campo: c, lotes: ls, tot: t }) => {
                const abierto = !!camposAbiertos[c];
                const celdaG = { padding: "10px 12px", fontWeight: 700, fontSize: 13, textAlign: "right" };
                return (
                  <React.Fragment key={c}>
                    <tr style={{ cursor: "pointer", background: "#F0EBE0" }} onClick={() => { toggleCampo(c); setSeleccion(null); }}>
                      <td style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13 }} colSpan={2}>
                        <span style={{ color: "var(--ink-soft)", fontSize: 12, marginRight: 6 }}>{abierto ? "▾" : "▸"}</span>
                        {c} <span style={{ fontWeight: 400, color: "var(--ink-soft)", fontSize: 12 }}>· {ls.length} lotes</span>
                      </td>
                      <td style={celdaG}>{fmtNum(t.supLote, 1)}</td>
                      <td style={celdaG}>{fmtNum(t.supPlan, 1)}</td>
                      <td style={celdaG}>{fmtNum(t.activa, 1)}</td>
                      <td style={celdaG}>{fmtNum(t.habilitada, 1)}</td>
                      <td style={{ ...celdaG, textAlign: undefined }}><Barra valor={t.pctSemb} /></td>
                      <td /><td />
                    </tr>
                    {abierto && ls.map((l) => {
                      const sel = seleccion === l.key;
                      return (
                        <React.Fragment key={l.key}>
                          <tr className="agri-tr" style={{ cursor: "pointer" }} onClick={() => setSeleccion(sel ? null : l.key)}>
                            <td className="agri-td" style={{ paddingLeft: 32, fontWeight: sel ? 700 : 500, color: sel ? "var(--green)" : undefined }}>
                              {l.lote}
                              {l.superp > 0.005 && <span title="Con superficie superpuesta (resiembra)" style={{ marginLeft: 6, fontSize: 10, color: "var(--gold)", border: "1px solid var(--gold)", borderRadius: 3, padding: "0 3px" }}>R</span>}
                            </td>
                            <td className="agri-td" style={{ fontSize: 12, color: "var(--ink-soft)" }}>{l.estados.join(", ")}</td>
                            <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(l.supLote, 1)}</td>
                            <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(l.supPlan, 1)}</td>
                            <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{fmtNum(l.activa, 1)}</td>
                            <td className="agri-td" style={{ textAlign: "right", color: l.habilitada > 0 ? undefined : "var(--ink-soft)" }}>{l.habilitada > 0 ? fmtNum(l.habilitada, 1) : "—"}</td>
                            <td className="agri-td"><Barra valor={l.pctSemb} /></td>
                            <td className="agri-td" style={{ fontSize: 12 }}>
                              {l.cultivos.length ? l.cultivos.map((cu) => (
                                <span key={cu} style={{ display: "inline-flex", alignItems: "center", gap: 4, marginRight: 8 }}><Punto color={colorDe(cu)} />{cu}</span>
                              )) : <span style={{ color: "var(--ink-soft)" }}>Sin cultivo</span>}
                            </td>
                            <td className="agri-td" style={{ textAlign: "center", color: "var(--ink-soft)" }}>{sel ? <ChevronUp size={14} /> : <ChevronDown size={14} />}</td>
                          </tr>
                          {sel && <PanelLote l={l} />}
                        </React.Fragment>
                      );
                    })}
                    <tr><td colSpan={9} style={{ height: 4 }} /></tr>
                  </React.Fragment>
                );
              })}
              <tr style={{ borderTop: "2px solid var(--line)" }}>
                {["Total", "", fmtNum(siembra.supLote, 1), fmtNum(siembra.supPlan, 1), fmtNum(siembra.activa, 1), fmtNum(siembra.habilitada, 1)].map((v, i) => (
                  <td key={i} style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13, textAlign: i >= 2 ? "right" : "left" }}>{v}</td>
                ))}
                <td style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13 }}><Barra valor={siembra.pctSemb} /></td>
                <td /><td />
              </tr>
            </tbody>
          </table>
        </div>
        <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 8 }}>
          Superficie del lote y planificada: una vez por lote. Habilitada s/cult. = habilitada para sembrar, todavía sin cultivo.
          <span style={{ color: "var(--gold)", border: "1px solid var(--gold)", borderRadius: 3, padding: "0 3px", margin: "0 4px" }}>R</span>
          lote con superficie superpuesta (resiembra).
        </div>
      </div>
    </>
  );

  // -------------------------------------------------------------------------
  // Vista Cosecha
  // -------------------------------------------------------------------------
  const vistaCosecha = conCultivo.length === 0 ? (
    <div className="agri-card" style={{ padding: 40, textAlign: "center", color: "var(--ink-soft)" }}>Con estos filtros no hay lotes sembrados.</div>
  ) : (
    <>
      <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 12, marginBottom: 20 }}>
        <Kpi label="Has activas" value={fmtNum(cosecha.activa, 0)} sub={`${unidades.length} lotes · cultivo · variedad`} />
        <Kpi label="Avance de cosecha" value={`${fmtNum(cosecha.pctTri, 1)} %`} sub={`${fmtNum(cosecha.haTri, 0)} ha trilladas · ${fmtNum(Math.max(0, cosecha.activa - cosecha.haTri), 0)} pendientes`}>
          <div style={{ height: 5, background: "var(--line)", borderRadius: 3, marginTop: 6, overflow: "hidden" }}>
            <div style={{ width: `${Math.min(100, cosecha.pctTri)}%`, height: "100%", background: "var(--green)" }} />
          </div>
        </Kpi>
        <Kpi label="Producción (Neto O)" value={`${fmtNum(cosecha.netoO / 1000, 1)} tn`}
          sub={cosecha.kgSinTrilla > 0 ? <span style={{ color: "var(--rust)" }}>{fmtNum(cosecha.kgSinTrilla / 1000, 1)} tn sin superficie trillada</span> : undefined} />
        <Kpi label="Consignado (Desc D)" value={`${fmtNum(cosecha.descD / 1000, 1)} tn`} sub={cosecha.netoO > 0 ? `${fmtNum(mermaPct, 2)} % menos que Neto O` : undefined} />
        <Kpi label="Rinde (kg/ha trillada)">
          {unSoloCultivo ? (
            <>
              <div className="agri-kpi-value">{cosecha.rinde === null ? "—" : fmtNum(cosecha.rinde)}</div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>{cosecha.rinde === null ? "Sin superficie trillada" : `Consignado: ${cosecha.rindeCon === null ? "—" : fmtNum(cosecha.rindeCon)} kg/ha activa`}</div>
            </>
          ) : (
            <div style={{ marginTop: 4 }}>
              {rindePorCultivo.map((c) => (
                <div key={c.cultivo} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, lineHeight: 1.6 }}>
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}><Punto color={colorDe(c.cultivo)} />{c.cultivo}</span>
                  <span style={{ fontWeight: 700 }}>{c.rinde === null ? "—" : fmtNum(c.rinde)}</span>
                </div>
              ))}
            </div>
          )}
        </Kpi>
      </div>

      <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(340px,1fr))", gap: 14, marginBottom: 14 }}>
        <div className="agri-card" style={{ padding: 16 }}>
          <Titulo sub={`% acumulado de has activas trilladas, por fecha de trilla${trillaSinFecha > 0 ? ` · ${fmtNum(trillaSinFecha, 1)} ha trilladas sin fecha no se grafican` : ""}`}>Avance de cosecha</Titulo>
          {avanceCosecha.length === 0 ? <SinDatos>Todavía no hay trilla con fecha registrada.</SinDatos> : (
            <ResponsiveContainer width="100%" height={230}>
              <LineChart data={avanceCosecha} margin={{ left: -10, right: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                <XAxis dataKey="t" type="number" scale="time" domain={["dataMin", "dataMax"]} tickFormatter={(t) => fmtDateShort(new Date(t).toISOString().slice(0, 10))} tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} minTickGap={28} />
                <YAxis domain={[0, 100]} ticks={[0, 25, 50, 75, 100]} tickFormatter={(v) => `${v}%`} tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={false} tickLine={false} />
                <Tooltip content={<AvanceTooltip unidad="%" />} cursor={{ stroke: "#6B5E4F", strokeDasharray: "3 3" }} />
                <Legend wrapperStyle={{ fontSize: 11 }} formatter={legendText} />
                {cultivosEnFiltro.map((c) => <Line key={c} type="stepAfter" dataKey={c} name={c} stroke={colorDe(c)} strokeWidth={2} dot={false} activeDot={{ r: 4 }} />)}
              </LineChart>
            </ResponsiveContainer>
          )}
        </div>
        <div className="agri-card" style={{ padding: 16 }}>
          <Titulo sub="Trilladas y pendientes sobre has activas">Avance por campo (ha)</Titulo>
          <ResponsiveContainer width="100%" height={Math.max(170, avancePorCampo.length * 36 + 70)}>
            <BarChart data={avancePorCampo} layout="vertical" margin={{ left: 0, right: 10 }} barCategoryGap={6} maxBarSize={34}>
              <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
              <YAxis type="category" dataKey="name" width={120} tick={{ fontSize: 11, fill: "#2B2118" }} axisLine={false} tickLine={false} />
              <Tooltip content={<BarrasHaTooltip />} cursor={{ fill: "rgba(107,94,79,0.08)" }} />
              <Legend wrapperStyle={{ fontSize: 11 }} formatter={legendText} />
              <Bar dataKey="trilladas" name="Trilladas" stackId="a" fill="#4B6B3A" stroke="#FCFAF2" strokeWidth={1}><LabelList dataKey="trilladas" content={<HorizontalBarLabel />} /></Bar>
              <Bar dataKey="pendientes" name="Pendientes" stackId="a" fill="#DCD2B8" stroke="#FCFAF2" strokeWidth={1} radius={[0, 3, 3, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(340px,1fr))", gap: 14, marginBottom: 14 }}>
        <div className="agri-card" style={{ padding: 16 }}>
          <Titulo sub="Neto (O) / has trilladas, por cultivo">Rinde por campo (kg/ha)</Titulo>
          {rindePorCampo.length === 0 ? <SinDatos alto={250}>Sin superficie trillada: el rinde queda vacío.</SinDatos> : (
            <ResponsiveContainer width="100%" height={250}>
              <BarChart data={rindePorCampo} margin={{ left: -6, right: 6 }} barGap={2}>
                <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                <Tooltip content={<RindeTooltip />} cursor={{ fill: "rgba(107,94,79,0.08)" }} />
                <Legend wrapperStyle={{ fontSize: 11 }} formatter={legendText} />
                {cultivosEnFiltro.map((c) => <Bar key={c} dataKey={c} name={c} fill={colorDe(c)} radius={[3, 3, 0, 0]} maxBarSize={28} />)}
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
        <div className="agri-card" style={{ padding: 16 }}>
          <Titulo sub="Neto (O) / has trilladas">Rinde por variedad (kg/ha)</Titulo>
          {rindePorVariedad.length === 0 ? <SinDatos alto={250}>Sin superficie trillada: el rinde queda vacío.</SinDatos> : (
            <ResponsiveContainer width="100%" height={Math.max(250, rindePorVariedad.length * 22 + 20)}>
              <BarChart data={rindePorVariedad} layout="vertical" margin={{ left: 0, right: 10 }} barCategoryGap={3} maxBarSize={26}>
                <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                <YAxis type="category" dataKey="name" width={160} tick={{ fontSize: 10.5, fill: "#2B2118" }} axisLine={false} tickLine={false} interval={0} />
                <Tooltip content={<VariedadRindeTooltip />} cursor={{ fill: "rgba(107,94,79,0.08)" }} />
                <Bar dataKey="rinde" name="Rinde" radius={[0, 3, 3, 0]}>
                  {rindePorVariedad.map((v, i) => <Cell key={i} fill={v.color} />)}
                  <LabelList dataKey="rinde" content={<HorizontalBarLabel />} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>

      <div data-pdf="pantalla" className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
          <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Cosecha por lote</div>
          <div data-pdf-ignore style={{ display: "flex", gap: 8 }}>
            <BotonExpandir grupos={tablaCosecha.map((g) => g.campo)} abiertos={camposAbiertos} setAbiertos={setCamposAbiertos} onCambio={() => setSeleccion(null)} />
            <button className="agri-btn agri-btn-outline" onClick={exportToExcel}><Download size={13} /> Exportar</button>
          </div>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
            <thead>
              <tr>
                {[["Campo / Lote", "left"], ["Cultivo", "left"], ["Variedad", "left"], ["Has act.", "right"], ["% Trilla", "right"], ["Neto O (tn)", "right"], ["Rinde", "right"], ["Rinde con.", "right"], ["", "center"]].map(([h, a], i) => (
                  <th key={i} className="agri-th" style={{ textAlign: a, cursor: "default", width: h ? undefined : 30 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {tablaCosecha.map(({ campo: c, unidades: us, tot: t, unCultivo }) => {
                const abierto = !!camposAbiertos[c];
                const celdaG = { padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 };
                return (
                  <React.Fragment key={c}>
                    <tr style={{ cursor: "pointer", background: "#F0EBE0" }} onClick={() => { toggleCampo(c); setSeleccion(null); }}>
                      <td style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13 }} colSpan={3}>
                        <span style={{ color: "var(--ink-soft)", fontSize: 12, marginRight: 6 }}>{abierto ? "▾" : "▸"}</span>
                        {c} <span style={{ fontWeight: 400, color: "var(--ink-soft)", fontSize: 12 }}>· {us.length} unidades</span>
                      </td>
                      <td style={celdaG}>{fmtNum(t.activa, 1)}</td>
                      <td style={{ ...celdaG, textAlign: undefined }}><Barra valor={t.pctTri} /></td>
                      <td style={celdaG}>{fmtNum(t.netoO / 1000, 1)}</td>
                      <td style={celdaG} title={unCultivo ? "" : "Varios cultivos: ver por lote"}>{unCultivo && t.rinde !== null ? fmtNum(t.rinde) : "—"}</td>
                      <td style={celdaG}>{unCultivo && t.rindeCon !== null ? fmtNum(t.rindeCon) : "—"}</td>
                      <td />
                    </tr>
                    {abierto && us.map((u) => {
                      const sel = seleccion === u.key;
                      return (
                        <React.Fragment key={u.key}>
                          <tr className="agri-tr" style={{ cursor: "pointer" }} onClick={() => setSeleccion(sel ? null : u.key)}>
                            <td className="agri-td" style={{ paddingLeft: 32, fontWeight: sel ? 700 : 500, color: sel ? "var(--green)" : undefined }}>
                              {u.lote}
                              {u.superp > 0.005 && <span title="Con superficie superpuesta (resiembra)" style={{ marginLeft: 6, fontSize: 10, color: "var(--gold)", border: "1px solid var(--gold)", borderRadius: 3, padding: "0 3px" }}>R</span>}
                            </td>
                            <td className="agri-td"><span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}><Punto color={colorDe(u.cultivo)} />{u.cultivo}</span></td>
                            <td className="agri-td" style={{ fontSize: 12 }}>{u.variedad}</td>
                            <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(u.activa, 2)}</td>
                            <td className="agri-td"><Barra valor={u.pctTri} /></td>
                            <td className="agri-td" style={{ textAlign: "right" }}>
                              {fmtNum(u.netoO / 1000, 1)}
                              {u.kgSinTrilla > 0 && <span title="Kilos cargados sin superficie trillada" style={{ color: "var(--rust)", marginLeft: 4 }}>*</span>}
                            </td>
                            <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{u.rinde === null ? "—" : fmtNum(u.rinde)}</td>
                            <td className="agri-td" style={{ textAlign: "right" }}>{u.rindeCon === null ? "—" : fmtNum(u.rindeCon)}</td>
                            <td className="agri-td" style={{ textAlign: "center", color: "var(--ink-soft)" }}>{sel ? <ChevronUp size={14} /> : <ChevronDown size={14} />}</td>
                          </tr>
                          {sel && <PanelUnidad u={u} />}
                        </React.Fragment>
                      );
                    })}
                    <tr><td colSpan={9} style={{ height: 4 }} /></tr>
                  </React.Fragment>
                );
              })}
              <tr style={{ borderTop: "2px solid var(--line)" }}>
                <td style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13 }} colSpan={3}>Total</td>
                <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{fmtNum(cosecha.activa, 1)}</td>
                <td style={{ padding: "10px 12px", fontSize: 13, fontWeight: 700 }}><Barra valor={cosecha.pctTri} /></td>
                <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{fmtNum(cosecha.netoO / 1000, 1)}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{unSoloCultivo && cosecha.rinde !== null ? fmtNum(cosecha.rinde) : "—"}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{unSoloCultivo && cosecha.rindeCon !== null ? fmtNum(cosecha.rindeCon) : "—"}</td>
                <td />
              </tr>
            </tbody>
          </table>
        </div>
        <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 8 }}>
          <span style={{ color: "var(--gold)", border: "1px solid var(--gold)", borderRadius: 3, padding: "0 3px", marginRight: 4 }}>R</span>
          con superficie superpuesta (resiembra). Rinde = Neto (O) / has trilladas, sólo donde hay superficie trillada (si no hay, queda vacío) · Rinde con. = Desc (D) / has activas.
          {cosecha.kgSinTrilla > 0 && <span style={{ color: "var(--rust)" }}> * kilos cargados sin superficie trillada.</span>}
        </div>
      </div>
    </>
  );

  // -------------------------------------------------------------------------
  return (
    <div className="agri-shell" ref={pdfRootRef}>
      {!bootLoading && subHeader}
      {bootLoading ? (
        <div style={{ padding: 60, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>
      ) : !meta || rows.length === 0 ? (
        isAdmin ? (
          <div className="agri-dropzone" data-drag={dragOver}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files?.[0]); }}>
            <FileSpreadsheet size={40} color="var(--ink-soft)" />
            <div style={{ fontSize: 16, fontWeight: 700 }}>{parsing ? "Leyendo…" : "Arrastrá el JSON o el Excel de Cultivos que genera Gorrión"}</div>
            <div style={{ fontSize: 13, color: "var(--ink-soft)", maxWidth: 420 }}>
              Una campaña por perfil: Campo, Lote, Sup Lote, Estado, Cultivo, Variedad, Siembra, superficies, Trilla y pesos.
              También se puede publicar directo desde Gorrión.
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
            <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Contactá a tu administrador para que publique los datos.</div>
          </div>
        )
      ) : (
        <>
          {/* Filtros y vista */}
          <div className="agri-card" style={{ padding: 14, marginBottom: 18 }}>
            <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 10 }}>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                <button className="agri-chip" data-active={cultivo === "Todos"} onClick={() => setCultivo("Todos")}>Todos los cultivos</button>
                {cultivos.map((c) => <button key={c} className="agri-chip" data-active={cultivo === c} onClick={() => setCultivo(c)}>{c}</button>)}
                {haySinCultivo && <button className="agri-chip" data-active={cultivo === SIN_CULTIVO} onClick={() => setCultivo(SIN_CULTIVO)}>Sin cultivo</button>}
              </div>
              <div style={{ display: "inline-flex", border: "1px solid var(--ink)", borderRadius: 6, overflow: "hidden" }}>
                {[["siembra", "Siembra"], ["cosecha", "Cosecha"]].map(([k, l]) => (
                  <button key={k} onClick={() => setVista(k)}
                    style={{ border: "none", padding: "6px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer", font: "inherit", background: vista === k ? "var(--ink)" : "transparent", color: vista === k ? "var(--paper)" : "var(--ink)" }}>
                    {l}
                  </button>
                ))}
              </div>
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
              <div style={{ position: "relative", flex: "1 1 220px" }}>
                <Search size={14} style={{ position: "absolute", left: 9, top: 10, color: "var(--ink-soft)" }} />
                <input className="agri-input" style={{ width: "100%", paddingLeft: 30 }} placeholder="Buscar campo, lote, variedad…" value={buscar} onChange={(e) => setBuscar(e.target.value)} />
              </div>
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
              {estados.length > 1 && (
                <select className="agri-select" value={estado} onChange={(e) => setEstado(e.target.value)}>
                  <option value="Todos">Todos los estados</option>
                  {estados.map((v) => <option key={v} value={v}>{v}</option>)}
                </select>
              )}
              {hayFiltros && <button className="agri-btn agri-btn-outline" onClick={resetFiltros}><X size={13} /> Limpiar</button>}
            </div>
          </div>

          {filtered.length === 0 ? (
            <div className="agri-card" style={{ padding: 40, textAlign: "center", color: "var(--ink-soft)" }}>Ningún registro coincide con los filtros.</div>
          ) : vista === "siembra" ? vistaSiembra : vistaCosecha}
        </>
      )}
    </div>
  );
}

// Botón "Expandir todo / Colapsar todo" de las tablas por campo
function BotonExpandir({ grupos, abiertos, setAbiertos, onCambio }) {
  const todos = grupos.length > 0 && grupos.every((g) => abiertos[g]);
  return (
    <button className="agri-btn agri-btn-outline" onClick={() => {
      const o = { ...abiertos };
      grupos.forEach((g) => { o[g] = !todos; });
      setAbiertos(o); onCambio?.();
    }}>{todos ? "Colapsar todo" : "Expandir todo"}</button>
  );
}
