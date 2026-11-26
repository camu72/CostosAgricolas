import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, Legend, LabelList,
} from "recharts";
import {
  Upload, RefreshCw, FileSpreadsheet, AlertCircle, Trash2,
  ChevronDown, ChevronUp, Download, ArrowUpDown, Cloud, CloudOff, LogOut,
} from "lucide-react";
import { ref as dbRef, onValue, set as dbSet } from "firebase/database";
import { PdfMenu, generarPDF, filaGrupo, filaTotal, MODOS_PDF } from "./pdfExport.jsx";
import { leerFilasArchivo, validarColumnas, leerEncabezado, leerCultivos, mapaColoresCultivos, colorDeCultivo, ACCEPT_ARCHIVOS } from "./lectorArchivo.js";

// PDF de Producción: Resumen, Pantalla y Totales por depósito (con detalle opcional)
const MODOS_PDF_PROD = [
  ...MODOS_PDF.filter((m) => m.key !== "completo"),
  { key: "depositos", label: "Totales por depósito", desc: "Resumen + stock de cada depósito, con las categorías desplegadas",
    opcion: { key: "detalle", label: "Incluir el detalle de cada depósito" } },
];

// ---------------------------------------------------------------------------
// Producción: datos generados por la pantalla "Dashboards > Producción" de Gorrión.
// Sigue el modelo de Cultivos (25/11/2026). Primera versión: sin compatibilidad con el formato anterior.
// Cada movimiento de grano llega en DOS filas, una por lado (ODT ↔ ODT Contrap.), con kilos opuestos.
// Llega por tres vías y siempre se guarda en clientes/<slug>/produccion:
//   - publicación directa desde Gorrión (PUT del JSON; NomSeccionFirebase = "produccion"),
//   - "Actualizar datos" con el JSON que genera Gorrión,
//   - "Actualizar datos" con el Excel que exporta Gorrión.
// Reglas acordadas con el usuario (25/11/2026):
//   - una sola campaña por perfil: sin selector de campaña;
//   - selector de peso Neto O. / Neto D.; con Neto D., si una fila no lo tiene (vacío o 0) se usa su Neto O.;
//   - cosecha = filas cuyo TipoDep es Cultivo con kilos negativos (el lote entrega grano); se muestra por Campo → Lote;
//   - stock de cada depósito = suma de los kilos de sus filas; sin rinde en este dashboard;
//   - el Articulo puede cambiar entre los dos lados de un movimiento (lo maneja Gorrión).
// ---------------------------------------------------------------------------
export const NODO_PRODUCCION = "produccion";
const COLS_PROD = [
  "Campaña","Cultivo","Fecha","Comp.","Suc","Num","Dominio","Acoplado",
  "Conductor","Transporte","Nº CP","Campo","Lote","Articulo","Titular",
  "Admin","Socio","TipoDep","ODT","TipoDepContrap.","ODT Contrap.",
  "Neto O.","Seco O.","Desc O.","Neto D.","Seco D.","Desc D.",
  "CampoOrigen","SocioContraparte",
];
const REQUERIDAS = ["Cultivo", "Fecha", "Campo", "TipoDep", "ODT", "TipoDepContrap.", "ODT Contrap.", "Neto O."];
const STORAGE_KEY_PROD = "produccion-gorrion-v1";   // distinta de la del dashboard anterior: no lee su caché
const PAGE_SIZE = 30;

const CULTIVO_COLORS = { SOJA: "#5B7C4B", MAIZ: "#C68F41", POROTO: "#35606B", GARBANZO: "#8A5A3B", TRIGO: "#B8842E", GIRASOL: "#A1462F" };
// Colores para cultivos sin color propio: distintos de los fijos de arriba, así dos cultivos no comparten color
const FALLBACK_COLORS = ["#7C8A4B", "#6E5A8A", "#3F7F6E", "#B26A8E", "#8C7A3A", "#4F6F8F"];
// Colores de los cultivos: primero los que manda Gorrión en la sección "cultivos" del JSON
// (administrador de cultivos de cada Gorrión); si no vino, la paleta fija y después la de respaldo.
let coloresGorrion = new Map();
const cultivoColor = (name, i) => colorDeCultivo(coloresGorrion, name) || CULTIVO_COLORS[String(name).toUpperCase()] || FALLBACK_COLORS[i % FALLBACK_COLORS.length];

// Encabezados: se comparan sin espacios, puntos ni acentos, sin distinguir mayúsculas; ° y º valen lo mismo (N° CP / Nº CP)
const normKey = (k) => String(k || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[°º]/g, "o").replace(/[\s.]/g, "").toLowerCase();
const COLS_POR_CLAVE = new Map(COLS_PROD.map((c) => [normKey(c), c]));
const alinearColumnas = (o) => {
  const r = {};
  Object.keys(o).forEach((k) => { r[COLS_POR_CLAVE.get(normKey(k)) || k] = o[k]; });
  return r;
};
const rowsFromPayload = (payload) => payload.rows.map((arr) => {
  const o = {};
  payload.columns.forEach((c, i) => { o[c] = arr[i]; });
  return normalizeRowProd(alinearColumnas(o));
}).filter((r) => r["Cultivo"] || r["Campo"]);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const str = (v) => (v === null || v === undefined ? "" : String(v).trim());
// Números: acepta 1234.5, "1234,5", "1.234,5" y "1,234.5"
const parseNum = (v) => {
  if (typeof v === "number") return isNaN(v) ? NaN : v;
  let t = String(v ?? "").replace(/[^0-9,.\-]/g, "");
  if (!t) return NaN;
  const coma = t.lastIndexOf(","), punto = t.lastIndexOf(".");
  if (coma >= 0 && punto >= 0) t = coma > punto ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "");
  else if (coma >= 0) t = t.replace(",", ".");
  return parseFloat(t);
};
const num = (v) => { const n = parseNum(v); return isNaN(n) ? 0 : n; };
const numOrNull = (v) => {
  if (v === null || v === undefined || String(v).trim() === "") return null;
  const n = parseNum(v);
  return isNaN(n) ? null : n;
};
// Fechas: yyyy-mm-dd (JSON de Gorrión), Date de SheetJS (UTC) o dd/mm/aaaa (Excel exportado).
// La fecha cero de Delphi (1899-12-30) y los años < 1901 se toman como vacías.
const toISODate = (v) => {
  let y, m, d;
  if (v instanceof Date && !isNaN(v)) { y = v.getUTCFullYear(); m = v.getUTCMonth() + 1; d = v.getUTCDate(); }
  else if (typeof v === "string" && v.trim()) {
    const t = v.trim();
    let mm = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (mm) { y = +mm[1]; m = +mm[2]; d = +mm[3]; }
    else if ((mm = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) { y = +mm[3]; m = +mm[2]; d = +mm[1]; }
    else { const dt = new Date(t); if (isNaN(dt)) return ""; y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate(); }
  } else return "";
  if (!(y >= 1901)) return "";
  return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
};
// Kilos de una fila según el peso elegido. Con Neto D., si la fila no lo tiene (vacío o 0) se usa Neto O.
const kgDe = (r, peso) => (peso === "D" && r["Neto D."] ? r["Neto D."] : (r["Neto O."] || 0));
const fmtNum = (n, d=0) => new Intl.NumberFormat("es-AR",{maximumFractionDigits:d,minimumFractionDigits:d}).format(n||0);
const fmtTn = (n) => `${fmtNum(n,1)} tn`;
const fmtDate = (iso) => {
  if (!iso) return "";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
const monthLabel = (ym) => {
  const [y,m] = ym.split("-");
  const names=["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];
  return `${names[parseInt(m,10)-1]} ${y.slice(2)}`;
};

// Clasifica si un TipoDep es un "cultivo" (campo) o un depósito físico
const esCultivo = (tipo) => {
  const t = str(tipo).toLowerCase();
  return t === "cultivo" || t === "cultivos";   // Gorrión manda "Cultivo" en TipoDep y "Cultivos" en TipoDepContrap.
};

// Tipo de movimiento inferido
const tipoMovimiento = (row) => {
  if (esCultivo(row["TipoDep"]) || esCultivo(row["TipoDepContrap."])) return "Cosecha";
  return "Movimiento de grano";
};

function normalizeRowProd(r) {
  return {
    "Campaña":          str(r["Campaña"]),
    "Cultivo":          str(r["Cultivo"]),
    "Fecha":            toISODate(r["Fecha"]),
    "Comp.":            str(r["Comp."]),
    "Suc":              num(r["Suc"]),
    "Num":              num(r["Num"]),
    "Dominio":          str(r["Dominio"]),
    "Acoplado":         str(r["Acoplado"]),
    "Conductor":        str(r["Conductor"]),
    "Transporte":       str(r["Transporte"]),
    "Nº CP":            str(r["Nº CP"]),
    "Campo":            str(r["Campo"]),
    "Lote":             str(r["Lote"]),
    "Articulo":         str(r["Articulo"]),
    "Titular":          str(r["Titular"]),
    "Admin":            str(r["Admin"]),
    "Socio":            str(r["Socio"]),
    "TipoDep":          str(r["TipoDep"]),
    "ODT":              str(r["ODT"]),
    "TipoDepContrap.":  str(r["TipoDepContrap."]),
    "ODT Contrap.":     str(r["ODT Contrap."]),
    "Neto O.":          numOrNull(r["Neto O."]),
    "Seco O.":          numOrNull(r["Seco O."]),
    "Desc O.":          numOrNull(r["Desc O."]),
    "Neto D.":          numOrNull(r["Neto D."]),
    "Seco D.":          numOrNull(r["Seco D."]),
    "Desc D.":          numOrNull(r["Desc D."]),
    "CampoOrigen":      str(r["CampoOrigen"]),
    "SocioContraparte": str(r["SocioContraparte"]),
  };
}

// ---------------------------------------------------------------------------
// Tooltip personalizado
// ---------------------------------------------------------------------------
const ChartTooltip = ({ active, payload, label, unit = "tn" }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <div style={{ background: "#FCFAF2", border: "1px solid #DCD2B8", borderRadius: 6, padding: "8px 10px", fontSize: 12 }}>
      {label && <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>}
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || "#2B2118" }}>
          {p.name}: {fmtNum(p.value, 1)} {unit}
        </div>
      ))}
    </div>
  );
};

const HorizontalBarLabel = ({ x, y, width, height, value }) => {
  if (!value || width < 55) return null;
  return (
    <text x={x + width - 6} y={y + height / 2} textAnchor="end" dominantBaseline="middle" fill="#fff" fontSize={10} fontWeight={700}>
      {fmtNum(value, 1)}
    </text>
  );
};

const VerticalBarLabel = ({ x, y, width, height, value }) => {
  if (!value || height < 16) return null;
  return (
    <text x={x + width / 2} y={y + height / 2} textAnchor="middle" dominantBaseline="middle" fill="#fff" fontSize={10} fontWeight={700}>
      {fmtNum(value, 1)}
    </text>
  );
};

// ---------------------------------------------------------------------------
// Componente principal
// ---------------------------------------------------------------------------
export default function DashboardProduccion({ slug, isAdmin, cloudDb, onLogout, clienteNombre }) {
  const pdfRootRef = useRef(null);
  const storageKey = `${STORAGE_KEY_PROD}-${slug}`;
  const cloudPath = `clientes/${slug}/${NODO_PRODUCCION}`;

  const [rows, setRows] = useState([]);
  const [meta, setMeta] = useState(null);
  const [bootLoading, setBootLoading] = useState(true);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState("");
  const [aviso, setAviso] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [syncStatus, setSyncStatus] = useState(cloudDb ? "connecting" : "local");
  const fileInputRef = useRef(null);
  const [showDetalle, setShowDetalle] = useState(false);
  const [selectedDeposito, setSelectedDeposito] = useState(null);
  const [gruposAbiertos, setGruposAbiertos] = useState({ planta: true, silo: false, otros: false });
  const toggleGrupo = (g) => setGruposAbiertos((prev) => ({ ...prev, [g]: !prev[g] }));

  // Filtros
  const [buscar, setBuscar]   = useState("");
  const [peso, setPeso]       = useState("O");   // "O" = Neto O., "D" = Neto D. (con Neto O. si falta)
  const [camposAbiertos, setCamposAbiertos] = useState({});
  const [cultivo, setCultivo] = useState("Todos");
  const [admin, setAdmin]     = useState("Todos");
  const [campo, setCampo]     = useState("Todos");
  const [tipoMov, setTipoMov] = useState("Todos");
  const [sort, setSort] = useState({ key: "Fecha", dir: "desc" });
  const [page, setPage] = useState(1);

  // ---------- Carga: caché local y luego Firebase en vivo (igual que Cultivos) ----------
  useEffect(() => {
    let unsub = null;
    const aplicar = (payload) => {
      const r = rowsFromPayload(payload);
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

  const persist = useCallback(async (normalized, fileName, encabezado, cultivos) => {
    const payload = {
      columns: COLS_PROD,
      rows: normalized.map((r) => COLS_PROD.map((c) => r[c] ?? "")),
      fileName, updatedAt: new Date().toISOString(),
      ...(encabezado ? { encabezado } : {}),
      ...(cultivos ? { cultivos } : {}),
    };
    try { localStorage.setItem(storageKey, JSON.stringify(payload)); } catch (e) {}
    if (cloudDb) {
      try { await dbSet(dbRef(cloudDb, cloudPath), payload); setSyncStatus("cloud"); }
      catch (e) { setSyncStatus("error"); throw e; }
    }
  }, [slug, cloudDb]);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setParsing(true); setError(""); setAviso("");
    try {
      // Excel (primera hoja) o JSON con el mismo formato
      const { filas: json, fileName, encabezado, cultivos } = await leerFilasArchivo(file);
      const alineadas = json.map(alinearColumnas);
      validarColumnas(alineadas, REQUERIDAS);
      const normalized = alineadas.map(normalizeRowProd).filter((r) => r["Cultivo"] || r["Campo"]);
      coloresGorrion = mapaColoresCultivos(cultivos);   // un Excel no trae colores: vuelve a la paleta
      setRows(normalized);
      setMeta({ fileName, updatedAt: new Date().toISOString(), rowCount: normalized.length, encabezado, cultivos });
      setPage(1); setCultivo("Todos"); setAdmin("Todos"); setCampo("Todos"); setTipoMov("Todos"); setBuscar("");
      try {
        await persist(normalized, fileName, encabezado, cultivos);
      } catch (e) {
        setError(`El archivo se leyó, pero no se pudo guardar en la nube (${e.code || e.message}). Sólo lo ves en este navegador.`);
      }
    } catch (e) {
      setError("No pude leer el archivo: " + e.message);
    } finally {
      setParsing(false);
    }
  }, [persist]);

  const clearDataset = async () => {
    setRows([]); setMeta(null);
    try { localStorage.removeItem(storageKey); } catch (e) {}
    if (cloudDb) { try { await dbSet(dbRef(cloudDb, cloudPath), null); } catch (e) {} }
  };

  // Listas para filtros
  // Kilos según el peso elegido (campo calculado "kg"; no se guarda)
  const rowsKg = useMemo(() => rows.map((r) => ({ ...r, kg: kgDe(r, peso) })), [rows, peso]);
  const sinNetoD = useMemo(() => (peso === "D" ? rows.filter((r) => !r["Neto D."]).length : 0), [rows, peso]);
  const cultivos    = useMemo(() => Array.from(new Set(rows.map((r) => r["Cultivo"]).filter(Boolean))).sort(), [rows]);
  const admins      = useMemo(() => Array.from(new Set(rows.map((r) => r["Admin"]).filter(Boolean))).sort(), [rows]);
  const campos      = useMemo(() => {
    const base = admin === "Todos" ? rows : rows.filter((r) => r["Admin"] === admin);
    return Array.from(new Set(base.map((r) => r["Campo"]).filter((c) => c && !c.toUpperCase().includes("DEPOSITO")))).sort();
  }, [rows, admin]);
  useEffect(() => { if (campo !== "Todos" && !campos.includes(campo)) setCampo("Todos"); }, [campos]);

  // Filtrado
  const buscarQ = buscar.trim().toLowerCase();
  const filtered = useMemo(() => rowsKg.filter((r) => {
    if (cultivo !== "Todos" && r["Cultivo"] !== cultivo) return false;
    if (admin   !== "Todos" && r["Admin"]   !== admin)   return false;
    if (campo   !== "Todos" && r["Campo"]   !== campo && r["CampoOrigen"] !== campo) return false;
    if (tipoMov !== "Todos" && tipoMovimiento(r) !== tipoMov) return false;
    if (buscarQ) {
      const hay = ["Campo","CampoOrigen","Lote","ODT","ODT Contrap.","Dominio","Transporte","Socio","SocioContraparte","Conductor"]
        .some((k) => str(r[k]).toLowerCase().includes(buscarQ));
      if (!hay) return false;
    }
    return true;
  }), [rowsKg, cultivo, admin, campo, tipoMov, buscarQ]);

  useEffect(() => { setPage(1); }, [cultivo, admin, campo, tipoMov, buscarQ]);

  // ---------- KPIs ----------
  const kpis = useMemo(() => {
    // Cosecha: filas donde el TipoDep es un cultivo y kilos < 0 (el lote entrega grano)
    const cosechaRows = filtered.filter((r) => esCultivo(r["TipoDep"]) && r.kg < 0);
    const tnCosechadas = cosechaRows.reduce((s, r) => s + Math.abs(r.kg), 0) / 1000;

    // Depósitos: filas cuyo TipoDep NO es cultivo → acumulamos los kilos del peso elegido
    const deposRows = filtered.filter((r) => !esCultivo(r["TipoDep"]));
    let stockPlanta = 0, stockSilo = 0, stockPermanente = 0, haySilo = false;
    deposRows.forEach((r) => {
      const v = r.kg / 1000;
      const t = str(r["TipoDep"]).toLowerCase();
      if (t === "planta de acopio") stockPlanta += v;
      else if (t === "temporal") { stockSilo += v; haySilo = true; }
      else if (t === "permanente") stockPermanente += v;
    });

    const viajeRows = filtered.filter((r) => r["Dominio"] && r.kg > 0);
    return { tnCosechadas, stockPlanta, stockSilo, haySilo, stockPermanente, viajes: viajeRows.length };
  }, [filtered]);

  // ---------- Gráficos ----------
  const cosechaPorCultivo = useMemo(() => {
    const m = new Map();
    filtered.filter((r) => esCultivo(r["TipoDep"]) && r.kg < 0)
      .forEach((r) => {
        const c = r["Cultivo"] || "(sin cultivo)";
        m.set(c, (m.get(c) || 0) + Math.abs(r.kg) / 1000);
      });
    return Array.from(m.entries()).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  const cosechaPorCampo = useMemo(() => {
    const m = new Map();
    filtered.filter((r) => esCultivo(r["TipoDep"]) && r.kg < 0)
      .forEach((r) => {
        const c = r["Campo"] || r["CampoOrigen"] || "(sin campo)";
        if (c.toUpperCase().includes("DEPOSITO")) return;
        m.set(c, (m.get(c) || 0) + Math.abs(r.kg) / 1000);
      });
    return Array.from(m.entries()).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  const evolucionMensual = useMemo(() => {
    const m = new Map();
    const cultList = cultivos;
    filtered.filter((r) => esCultivo(r["TipoDep"]) && r.kg < 0 && r["Fecha"])
      .forEach((r) => {
        const ym = r["Fecha"].slice(0, 7);
        if (!m.has(ym)) m.set(ym, {});
        const bucket = m.get(ym);
        const c = r["Cultivo"] || "(sin cultivo)";
        bucket[c] = (bucket[c] || 0) + Math.abs(r.kg) / 1000;
      });
    if (!m.size) return [];
    // Meses completos entre el primero y el último con cosecha: los meses sin cosecha valen 0,
    // así la línea baja a cero en vez de unir meses que no son consecutivos.
    const meses = Array.from(m.keys()).sort();
    const sumar = (ym) => { const [y, mo] = ym.split("-").map(Number); return mo === 12 ? `${y + 1}-01` : `${y}-${String(mo + 1).padStart(2, "0")}`; };
    const out = [];
    for (let ym = meses[0]; ym <= meses[meses.length - 1]; ym = sumar(ym)) {
      const bucket = m.get(ym) || {};
      const fila = { ym, label: monthLabel(ym) };
      cultList.forEach((c) => { fila[c] = bucket[c] || 0; });
      Object.keys(bucket).forEach((c) => { if (!(c in fila)) fila[c] = bucket[c]; });
      out.push(fila);
    }
    return out;
  }, [filtered, cultivos]);

  const stockPorDeposito = useMemo(() => {
    const m = new Map();
    filtered.filter((r) => !esCultivo(r["TipoDep"])).forEach((r) => {
      const key = `${r["TipoDep"]}||${r["ODT"]}`;
      if (!m.has(key)) m.set(key, { tipo: r["TipoDep"], nombre: r["ODT"], stock: 0 });
      m.get(key).stock += r.kg / 1000;
    });
    return Array.from(m.values()).filter((e) => Math.abs(e.stock) > 0.01).sort((a, b) => b.stock - a.stock);
  }, [filtered]);

  // Cosecha por Campo → Lote (filas de cosecha: TipoDep = Cultivo y kilos < 0)
  const cosechaPorLote = useMemo(() => {
    const campos = new Map();
    filtered.filter((r) => esCultivo(r["TipoDep"]) && r.kg < 0).forEach((r) => {
      const c = r["Campo"] || r["CampoOrigen"] || "(sin campo)";
      if (!campos.has(c)) campos.set(c, { campo: c, tn: 0, viajes: 0, lotes: new Map() });
      const g = campos.get(c);
      const lk = `${r["Lote"]}||${r["Cultivo"]}`;
      if (!g.lotes.has(lk)) g.lotes.set(lk, { lote: r["Lote"] || "(sin lote)", cultivo: r["Cultivo"], tn: 0, viajes: 0, primera: "", ultima: "" });
      const l = g.lotes.get(lk);
      const tn = Math.abs(r.kg) / 1000;
      g.tn += tn; g.viajes += 1; l.tn += tn; l.viajes += 1;
      if (r["Fecha"]) {
        if (!l.primera || r["Fecha"] < l.primera) l.primera = r["Fecha"];
        if (!l.ultima || r["Fecha"] > l.ultima) l.ultima = r["Fecha"];
      }
    });
    return Array.from(campos.values())
      .map((g) => ({ ...g, lotes: Array.from(g.lotes.values()).sort((a, b) => a.lote.localeCompare(b.lote, "es", { numeric: true }) || a.cultivo.localeCompare(b.cultivo)) }))
      .sort((a, b) => b.tn - a.tn);
  }, [filtered]);
  const toggleCampoLote = (c) => setCamposAbiertos((p) => ({ ...p, [c]: !p[c] }));

  // ---------- Tabla detalle ----------
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
  const toggleSort = (key) => setSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "asc" }));

  const exportToExcel = useCallback(() => {
    const data = sorted.map((r) => {
      const o = {};
      COLS_PROD.forEach((c) => { o[c] = r[c] === null ? "" : c === "Fecha" ? fmtDate(r[c]) : r[c]; });
      return o;
    });
    const ws = XLSX.utils.json_to_sheet(data, { header: COLS_PROD });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Produccion");
    XLSX.writeFile(wb, `produccion_filtrada_${new Date().toISOString().slice(0,10)}.xlsx`);
  }, [sorted]);

  // ---------- PDF ----------
  const exportarPDF = async (modo, opciones = {}) => {
    const filtros = [];
    filtros.push(["Peso", peso === "D" ? "Neto D. (Neto O. si falta)" : "Neto O."]);
    if (cultivo !== "Todos") filtros.push(["Cultivo", cultivo]);
    if (admin !== "Todos") filtros.push(["Admin", admin]);
    if (campo !== "Todos") filtros.push(["Campo", campo]);
    if (tipoMov !== "Todos") filtros.push(["Tipo", tipoMov === "Cosecha" ? "Solo cosecha" : "Solo movimientos"]);
    if (buscar.trim()) filtros.push(["Búsqueda", `"${buscar.trim()}"`]);

    // Depósitos agrupados igual que en pantalla, todos desplegados
    const grupos = [
      { label: "Plantas de acopio", match: (t) => t === "planta de acopio" },
      { label: "Silo bolsa",        match: (t) => t === "temporal" },
      { label: "Otros depósitos",   match: (t) => t !== "planta de acopio" && t !== "temporal" },
    ];
    const movsDe = (nombre) => filtered.filter((r) => r["ODT"] === nombre).sort((x, y) => {
      const fa = x["Fecha"] || "", fb = y["Fecha"] || "";
      return fa > fb ? -1 : fa < fb ? 1 : 0;
    });
    const depositos = [];   // [{ grupo, dep, movs, ent, sal }] en orden de la tabla
    const stockBody = [];
    let stockTotal = 0, entTotal = 0, salTotal = 0;
    grupos.forEach(({ label, match }) => {
      const items = stockPorDeposito.filter((e) => match(String(e.tipo).toLowerCase()));
      if (!items.length) return;
      const filasGrupo = items.map((e) => {
        const movs = movsDe(e.nombre);
        const ent = movs.filter((r) => r.kg > 0).reduce((acc, r) => acc + r.kg, 0) / 1000;
        const sal = movs.filter((r) => r.kg < 0).reduce((acc, r) => acc + Math.abs(r.kg), 0) / 1000;
        depositos.push({ grupo: label, dep: e, movs, ent, sal });
        return { e, ent, sal };
      });
      const tg = items.reduce((acc, e) => acc + e.stock, 0);
      const eg = filasGrupo.reduce((acc, f) => acc + f.ent, 0), sg = filasGrupo.reduce((acc, f) => acc + f.sal, 0);
      stockTotal += tg; entTotal += eg; salTotal += sg;
      stockBody.push(filaGrupo(label, 4, [fmtNum(eg, 1), fmtNum(sg, 1), fmtTn(tg)]));
      filasGrupo.forEach(({ e, ent, sal }) => stockBody.push([
        { content: e.nombre, styles: { cellPadding: { top: 1.4, bottom: 1.4, left: 6, right: 1.4 } } },
        fmtNum(ent, 1), fmtNum(sal, 1), { content: fmtTn(e.stock), styles: { fontStyle: "bold" } },
      ]));
    });
    stockBody.push(filaTotal(["Total", fmtNum(entTotal, 1), fmtNum(salTotal, 1), fmtTn(stockTotal)]));

    // Detalle de cada depósito (opcional): KPIs + todos los movimientos, sin scroll
    const detalle = {
      render: ({ doc, autoTable, pdfTxt, C, PAGE, TOP, BOTTOM, CONTENT_W, getY, setY, nuevaPagina }) => {
        const VERDE = [75, 107, 58], OXIDO = [161, 70, 47];
        depositos.forEach(({ grupo, dep, movs, ent, sal }, idx) => {
          if (idx === 0) {
            if (BOTTOM - getY() < 60) nuevaPagina();
            doc.setFont("helvetica", "bold"); doc.setFontSize(13); doc.setTextColor(...C.ink);
            doc.text(pdfTxt(`Detalle de cada depósito (${depositos.length})`), PAGE.m, getY() + 4);
            setY(getY() + 10);
          } else if (BOTTOM - getY() < 55) {
            nuevaPagina();
          } else {
            doc.setDrawColor(...C.line); doc.setLineWidth(0.3);
            doc.line(PAGE.m, getY() - 3, PAGE.w - PAGE.m, getY() - 3);
          }
          let y = getY();
          const pagina = doc.getNumberOfPages();
          // Encabezado del depósito
          doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setTextColor(...VERDE);
          doc.text(pdfTxt(dep.nombre), PAGE.m, y + 4);
          doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(...C.soft);
          doc.text(pdfTxt(`${grupo} · ${movs.length} movimientos`), PAGE.m, y + 8.5);
          y += 11;
          // Indicadores
          const kpis = [["Stock actual", fmtTn(dep.stock), VERDE], ["Total ingresado", `${fmtNum(ent, 1)} tn`, C.ink], ["Total egresado", `${fmtNum(sal, 1)} tn`, OXIDO]];
          const kw = (CONTENT_W - 8) / 3;
          kpis.forEach(([l, v, col], i) => {
            const x = PAGE.m + i * (kw + 4);
            doc.setDrawColor(...C.line); doc.setLineWidth(0.2); doc.setFillColor(...C.raised);
            doc.roundedRect(x, y, kw, 11, 1.2, 1.2, "FD");
            doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(...C.soft);
            doc.text(pdfTxt(l), x + 3, y + 4);
            doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setTextColor(...col);
            doc.text(pdfTxt(v), x + 3, y + 9);
          });
          y += 14;
          // Movimientos
          autoTable(doc, {
            startY: y,
            head: [["Fecha", "Cultivo", "Campo / Origen", "Contraparte", "Kg (neto)"]],
            body: movs.map((r) => {
              const kg = r.kg;
              return [
                fmtDate(r["Fecha"]), r["Cultivo"], r["CampoOrigen"] || r["Campo"] || "-", r["ODT Contrap."] || "-",
                { content: `${kg > 0 ? "+" : ""}${fmtNum(kg / 1000, 1)} tn`, styles: { fontStyle: "bold", textColor: kg > 0 ? VERDE : OXIDO } },
              ].map((c) => (typeof c === "object" ? { ...c, content: pdfTxt(c.content) } : pdfTxt(c)));
            }),
            margin: { left: PAGE.m, right: PAGE.m, top: TOP + 11, bottom: PAGE.h - BOTTOM },
            didDrawPage: () => {
              if (doc.internal.getCurrentPageInfo().pageNumber === pagina) return;
              doc.setFont("helvetica", "bold"); doc.setFontSize(10); doc.setTextColor(...VERDE);
              doc.text(pdfTxt(dep.nombre), PAGE.m, TOP + 3);
              doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(...C.soft);
              doc.text(pdfTxt(`${grupo} · movimientos (continuación)`), PAGE.m, TOP + 7);
            },
            styles: { font: "helvetica", fontSize: 8, cellPadding: 1.3, textColor: C.ink, lineColor: C.line, lineWidth: 0.1 },
            headStyles: { fillColor: C.paper, textColor: C.soft, fontStyle: "bold", fontSize: 7.5, lineWidth: 0 },
            bodyStyles: { fillColor: [255, 255, 255] },
            columnStyles: { 0: { cellWidth: 24 }, 1: { cellWidth: 24 }, 2: { cellWidth: 60 }, 4: { halign: "right", cellWidth: 28 } },
            showHead: "everyPage",
          });
          setY(doc.lastAutoTable.finalY + 10);
        });
      },
    };

    await generarPDF({
      modo, titulo: "Producción", cliente: clienteNombre || slug, filtros, root: pdfRootRef.current, archivo: "Produccion", encabezado: meta?.encabezado,
      modoLabel: modo === "depositos" ? `Totales por depósito${opciones.detalle ? " con detalle" : ""}` : undefined,
      tablas: [
        ...(cosechaPorLote.length ? [{
          titulo: "Cosecha por campo y lote", nota: "Toneladas entregadas por cada lote. Viajes = filas de cosecha.",
          head: ["Campo / Lote", "Cultivo", "Período", "Viajes", "Cosechado"], align: [null, null, null, "right", "right"],
          body: [
            ...cosechaPorLote.flatMap((g) => [
              filaGrupo(g.campo, 5, [fmtNum(g.viajes), fmtTn(g.tn)]),
              ...g.lotes.map((l) => [
                { content: l.lote, styles: { cellPadding: { top: 1.4, bottom: 1.4, left: 6, right: 1.4 } } },
                l.cultivo, l.primera ? `${fmtDate(l.primera)} – ${fmtDate(l.ultima)}` : "-", fmtNum(l.viajes),
                { content: fmtTn(l.tn), styles: { fontStyle: "bold" } },
              ]),
            ]),
            filaTotal(["Total", "", "", fmtNum(cosechaPorLote.reduce((a, g) => a + g.viajes, 0)), fmtTn(cosechaPorLote.reduce((a, g) => a + g.tn, 0))]),
          ],
        }] : []),
        { titulo: "Stock por depósito", nota: `Toneladas. Ingresado y egresado según ${peso === "D" ? "Neto D. (Neto O. si falta)" : "Neto O."} de los movimientos filtrados.`,
          head: ["Depósito", "Ingresado (tn)", "Egresado (tn)", "Stock"], align: [null, "right", "right", "right"], body: stockBody },
        ...(opciones.detalle ? [detalle] : []),
      ],
    });
  };

  const resetFiltros = () => { setBuscar(""); setCultivo("Todos"); setAdmin("Todos"); setCampo("Todos"); setTipoMov("Todos"); };
  const hayFiltros = buscar !== "" || cultivo !== "Todos" || admin !== "Todos" || campo !== "Todos" || tipoMov !== "Todos";

  // Sub-header común a todos los dashboards: datos, estado de sincronización y controles de carga
  const subHeader = (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
      <div className="agri-sub">
        {meta ? (
          <>Datos de <strong>{meta.fileName}</strong> · {fmtNum(meta.rowCount)} registros · actualizado {new Date(meta.updatedAt).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" })}</>
        ) : isAdmin ? "Cargá tu planilla para empezar a consultar los datos" : "Todavía no hay datos cargados para este cliente."}
        {aviso && <span style={{ color: "var(--green)", marginLeft: 8 }}>· {aviso}</span>}
        {meta && sinNetoD > 0 && <span style={{ color: "var(--ink-soft)", marginLeft: 8 }}>· {fmtNum(sinNetoD)} {sinNetoD === 1 ? "fila sin Neto D. usa" : "filas sin Neto D. usan"} Neto O. (*)</span>}
        {error && meta && <span style={{ color: "var(--rust)", marginLeft: 8 }}>· {error}</span>}
      </div>
      {cloudDb && (
        <div style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 600, padding: "3px 9px", borderRadius: 999, background: syncStatus === "cloud" ? "rgba(75,107,58,0.12)" : syncStatus === "error" ? "rgba(161,70,47,0.12)" : "rgba(107,94,79,0.12)", color: syncStatus === "cloud" ? "var(--green)" : syncStatus === "error" ? "var(--rust)" : "var(--ink-soft)" }}>
          {syncStatus === "cloud" ? <><Cloud size={12} /> Sincronizado</> : syncStatus === "error" ? <><CloudOff size={12} /> Sin conexión</> : <><Cloud size={12} /> Conectando…</>}
        </div>
      )}
      <div style={{ display: "flex", gap: 8 }}>
        {meta && rows.length > 0 && <PdfMenu onExport={exportarPDF} modos={MODOS_PDF_PROD} />}
        {isAdmin && (<>
          <input ref={fileInputRef} type="file" accept={ACCEPT_ARCHIVOS} style={{ display: "none" }}
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
            <div style={{ fontSize: 16, fontWeight: 700 }}>{parsing ? "Leyendo…" : "Arrastrá el Excel o JSON de Producción"}</div>
            <div style={{ fontSize: 13, color: "var(--ink-soft)", maxWidth: 380 }}>
              JSON o Excel que genera la pantalla de Producción de Gorrión.
            </div>
            <button className="agri-btn" disabled={parsing} onClick={() => fileInputRef.current?.click()}>
              <Upload size={14} /> {parsing ? "Procesando…" : "Elegir archivo"}
            </button>
            {error && <div style={{ color: "var(--rust)", fontSize: 13 }}><AlertCircle size={13} style={{ display: "inline", marginRight: 4 }} />{error}</div>}
          </div>
        ) : (
          <div className="agri-card" style={{ padding: 56, textAlign: "center" }}>
            <FileSpreadsheet size={36} color="var(--ink-soft)" style={{ marginBottom: 10 }} />
            <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 6 }}>Todavía no hay datos de producción</div>
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
              <div style={{ position: "relative", display: "flex", alignItems: "center" }}>
                <svg style={{ position: "absolute", left: 9, color: "var(--ink-soft)", pointerEvents: "none" }} width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                <input
                  className="agri-select"
                  style={{ paddingLeft: 30, minWidth: 200 }}
                  placeholder="Buscar campo, lote, ODT…"
                  value={buscar}
                  onChange={(e) => setBuscar(e.target.value)}
                />
              </div>
              <div className="agri-chip-group" title="Peso con el que se calculan cosecha y stock" style={{ display: "inline-flex", gap: 4, alignItems: "center" }}>
                <span style={{ fontSize: 12, color: "var(--ink-soft)", marginRight: 2 }}>Peso:</span>
                <button className="agri-chip" data-active={peso === "O"} onClick={() => setPeso("O")}>Neto O.</button>
                <button className="agri-chip" data-active={peso === "D"} onClick={() => setPeso("D")}>Neto D.</button>
              </div>
              <select className="agri-select" value={admin}    onChange={(e) => setAdmin(e.target.value)}>
                <option value="Todos">Todas las admins</option>
                {admins.map((a) => <option key={a} value={a}>{a}</option>)}
              </select>
              <select className="agri-select" value={campo}    onChange={(e) => setCampo(e.target.value)}>
                <option value="Todos">Todos los campos</option>
                {campos.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
              <select className="agri-select" value={tipoMov}  onChange={(e) => setTipoMov(e.target.value)}>
                <option value="Todos">Cosecha y movimientos</option>
                <option value="Cosecha">Solo cosecha</option>
                <option value="Movimiento de grano">Solo movimientos</option>
              </select>
              {hayFiltros && <button className="agri-btn agri-btn-outline" onClick={resetFiltros}>✕ Limpiar</button>}
            </div>
          </div>

          {/* KPIs */}
          <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 12, marginBottom: 20 }}>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Tn cosechadas</div>
              <div className="agri-kpi-value">{fmtNum(kpis.tnCosechadas, 1)}</div>
            </div>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Stock en planta</div>
              <div className="agri-kpi-value">{fmtNum(kpis.stockPlanta, 1)} tn</div>
            </div>
            {/* Sólo si hay silos bolsa (TipoDep = Temporal) en los datos filtrados; con silos y stock neto 0, se muestra */}
            {kpis.haySilo && (
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Stock en silo bolsa</div>
                <div className="agri-kpi-value">{fmtNum(kpis.stockSilo, 1)} tn</div>
              </div>
            )}
            {kpis.stockPermanente !== 0 && (
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Stock en otros dep.</div>
                <div className="agri-kpi-value">{fmtNum(kpis.stockPermanente, 1)} tn</div>
              </div>
            )}
          </div>

          {/* Gráficos — fila 1: cosecha */}
          <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(320px,1fr))", gap: 14, marginBottom: 14 }}>
            <div className="agri-card" style={{ padding: 16 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Cosecha por cultivo (tn)</div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={cosechaPorCultivo} margin={{ left: -10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                  <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                  <Tooltip content={<ChartTooltip unit="tn" />} />
                  <Bar dataKey="value" name="Cosechado" radius={[3,3,0,0]} maxBarSize={56}>
                    {cosechaPorCultivo.map((e,i) => <Cell key={i} fill={cultivoColor(e.name,i)} />)}
                    <LabelList dataKey="value" content={<VerticalBarLabel />} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="agri-card" style={{ padding: 16 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Cosecha por campo (tn)</div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={cosechaPorCampo} layout="vertical" margin={{ left: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" width={120} tick={{ fontSize: 11, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                  <Tooltip content={<ChartTooltip unit="tn" />} />
                  <Bar dataKey="value" name="Cosechado" fill="#B8842E" radius={[0,3,3,0]} maxBarSize={32}>
                    <LabelList dataKey="value" content={<HorizontalBarLabel />} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="agri-card" style={{ padding: 16 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Evolución mensual de cosecha (tn)</div>
              <ResponsiveContainer width="100%" height={220}>
                <LineChart data={evolucionMensual} margin={{ left: -10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                  <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                  <Tooltip content={<ChartTooltip unit="tn" />} />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  {cultivos.map((c, i) => (
                    <Line key={c} type="linear" dataKey={c} name={c} stroke={cultivoColor(c, i)} strokeWidth={2} dot={evolucionMensual.length <= 18 ? { r: 2.5 } : false} />
                  ))}
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Cosecha por campo y lote */}
          {cosechaPorLote.length > 0 && (
            <div data-pdf="pantalla" className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 12, flexWrap: "wrap", gap: 6 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Cosecha por campo y lote</div>
                <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>tocá un campo para ver sus lotes</span>
              </div>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th className="agri-th">Campo / Lote</th>
                    <th className="agri-th">Cultivo</th>
                    <th className="agri-th">Período de cosecha</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Viajes</th>
                    <th className="agri-th" style={{ textAlign: "right" }}>Cosechado</th>
                  </tr>
                </thead>
                <tbody>
                  {cosechaPorLote.map((g) => {
                    const abierto = !!camposAbiertos[g.campo];
                    return (
                      <React.Fragment key={g.campo}>
                        <tr style={{ cursor: "pointer", background: "var(--surface-alt, #F0EBE0)" }} onClick={() => toggleCampoLote(g.campo)}>
                          <td style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13 }} colSpan={3}>
                            <span style={{ color: "var(--ink-soft)", fontSize: 12, marginRight: 6 }}>{abierto ? "▾" : "▸"}</span>
                            {g.campo} <span style={{ fontWeight: 400, color: "var(--ink-soft)", fontSize: 12 }}>· {g.lotes.length} {g.lotes.length === 1 ? "lote" : "lotes"}</span>
                          </td>
                          <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13 }}>{fmtNum(g.viajes)}</td>
                          <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13, whiteSpace: "nowrap" }}>{fmtTn(g.tn)}</td>
                        </tr>
                        {abierto && g.lotes.map((l) => (
                          <tr key={`${l.lote}|${l.cultivo}`} className="agri-tr">
                            <td className="agri-td" style={{ paddingLeft: 32 }}>{l.lote}</td>
                            <td className="agri-td">{l.cultivo}</td>
                            <td className="agri-td" style={{ fontSize: 12, color: "var(--ink-soft)", whiteSpace: "nowrap" }}>
                              {l.primera ? (l.primera === l.ultima ? fmtDate(l.primera) : `${fmtDate(l.primera)} – ${fmtDate(l.ultima)}`) : "—"}
                            </td>
                            <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(l.viajes)}</td>
                            <td className="agri-td" style={{ textAlign: "right", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtTn(l.tn)}</td>
                          </tr>
                        ))}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* Stock por depósito — acordeón unificado */}
          {stockPorDeposito.length > 0 && (() => {
            const grupos = [
              { key: "planta", label: "Plantas de acopio", tipoMatch: (t) => t === "planta de acopio" },
              { key: "silo",   label: "Silo bolsa",        tipoMatch: (t) => t === "temporal" },
              { key: "otros",  label: "Otros depósitos",   tipoMatch: (t) => t !== "planta de acopio" && t !== "temporal" },
            ];

            const PanelDetalle = ({ dep }) => {
              const allMovs = filtered.filter((r) => r["ODT"] === dep.nombre).sort((a, b) => {
                const fa = a["Fecha"] || "", fb = b["Fecha"] || "";
                return fa > fb ? -1 : fa < fb ? 1 : 0;
              });
              const totalEntradas = allMovs.filter((r) => r.kg > 0).reduce((s, r) => s + r.kg, 0) / 1000;
              const totalSalidas  = allMovs.filter((r) => r.kg < 0).reduce((s, r) => s + Math.abs(r.kg), 0) / 1000;
              return (
                <tr>
                  <td colSpan={3} style={{ padding: 0, background: "var(--bg)" }}>
                    <div style={{ margin: "0 0 2px 32px", padding: 16, background: "var(--surface)", borderLeft: "3px solid var(--green)", borderRadius: "0 0 8px 8px" }}>
                      <div style={{ fontSize: 12, color: "var(--ink-soft)", marginBottom: 12 }}>{allMovs.length} movimientos</div>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(130px,1fr))", gap: 8, marginBottom: 14 }}>
                        {[
                          { label: "Stock actual",    val: fmtTn(dep.stock),              color: "var(--green)" },
                          { label: "Total ingresado", val: `${fmtNum(totalEntradas,1)} tn`, color: undefined },
                          { label: "Total egresado",  val: `${fmtNum(totalSalidas,1)} tn`,  color: "var(--rust)" },
                        ].map(({ label, val, color }) => (
                          <div key={label} style={{ background: "var(--bg)", borderRadius: 6, padding: "8px 12px", border: "1px solid var(--line)" }}>
                            <div className="agri-kpi-label">{label}</div>
                            <div className="agri-kpi-value" style={{ fontSize: 18, color }}>{val}</div>
                          </div>
                        ))}
                      </div>
                      <div style={{ fontWeight: 600, fontSize: 12, marginBottom: 6, color: "var(--ink)" }}>Movimientos</div>
                      <div style={{ overflowX: "auto", maxHeight: 300, overflowY: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                          <thead>
                            <tr>
                              {[["Fecha","Fecha"],["Cultivo","Cultivo"],["Campo","Campo / Origen"],["ODT Contrap.","Contraparte"],["kg","Kg (neto)"]].map(([k,l]) => (
                                <th key={k} className="agri-th">{l}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {allMovs.map((r, i) => {
                              const kg = r.kg;
                              const esEntrada = kg > 0;
                              return (
                                <tr key={i} className="agri-tr">
                                  <td className="agri-td" style={{ whiteSpace: "nowrap" }}>{fmtDate(r["Fecha"])}</td>
                                  <td className="agri-td">{r["Cultivo"]}</td>
                                  <td className="agri-td" style={{ fontSize: 11 }}>{r["CampoOrigen"] || r["Campo"] || "—"}</td>
                                  <td className="agri-td" style={{ fontSize: 11, color: "var(--ink-soft)" }}>{r["ODT Contrap."] || "—"}</td>
                                  <td className="agri-td" style={{ textAlign: "right", fontWeight: 600, color: esEntrada ? "var(--green)" : "var(--rust)" }}>
                                    {esEntrada ? "+" : ""}{fmtNum(kg / 1000, 1)} tn
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </td>
                </tr>
              );
            };

            return (
              <div data-pdf="pantalla" className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 12 }}>Stock por depósito</div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <tbody>
                    {grupos.map(({ key, label, tipoMatch }) => {
                      const items = stockPorDeposito.filter((e) => tipoMatch(e.tipo.toLowerCase()));
                      if (items.length === 0) return null;
                      const totalGrupo = items.reduce((s, e) => s + e.stock, 0);
                      const abierto = gruposAbiertos[key];
                      return (
                        <React.Fragment key={key}>
                          {/* Fila encabezado de grupo */}
                          <tr
                            style={{ cursor: "pointer", background: "var(--surface-alt, #F0EBE0)" }}
                            onClick={() => { toggleGrupo(key); setSelectedDeposito(null); }}
                          >
                            <td style={{ padding: "10px 12px", fontWeight: 700, fontSize: 13, color: "var(--ink)", display: "flex", alignItems: "center", gap: 6 }}>
                              <span style={{ color: "var(--ink-soft)", fontSize: 12 }}>{abierto ? "▾" : "▸"}</span>
                              {label}
                            </td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: 13, color: "var(--ink)", whiteSpace: "nowrap" }}>
                              {fmtNum(totalGrupo, 1)} tn
                            </td>
                            <td style={{ width: 30 }} />
                          </tr>
                          {/* Filas hijos */}
                          {abierto && items.map((e, i) => (
                            <React.Fragment key={i}>
                              <tr
                                className="agri-tr"
                                style={{ cursor: "pointer" }}
                                onClick={() => setSelectedDeposito(selectedDeposito?.nombre === e.nombre ? null : e)}
                              >
                                <td className="agri-td" style={{ paddingLeft: 32, color: selectedDeposito?.nombre === e.nombre ? "var(--green)" : undefined, fontWeight: selectedDeposito?.nombre === e.nombre ? 600 : undefined }}>
                                  {e.nombre}
                                </td>
                                <td className="agri-td" style={{ textAlign: "right", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtTn(e.stock)}</td>
                                <td className="agri-td" style={{ textAlign: "center", color: "var(--ink-soft)", width: 30 }}>
                                  {selectedDeposito?.nombre === e.nombre ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                                </td>
                              </tr>
                              {selectedDeposito?.nombre === e.nombre && <PanelDetalle dep={e} />}
                            </React.Fragment>
                          ))}
                          {/* Separador entre grupos */}
                          <tr><td colSpan={3} style={{ height: 6 }} /></tr>
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            );
          })()}

          {/* Tabla detalle colapsable */}
          <div data-pdf="pantalla" className="agri-card" style={{ padding: 0, overflow: "hidden" }}>
            <div style={{ padding: "12px 16px", borderBottom: showDetalle ? "1px solid var(--line)" : "none", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Detalle de movimientos</div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(sorted.length)} registros</span>
                <span data-pdf-ignore style={{ display: "contents" }}>
                <button className="agri-btn agri-btn-outline" onClick={exportToExcel}><Download size={13} /> Exportar</button>
                <button className="agri-btn agri-btn-outline" onClick={() => setShowDetalle((v) => !v)}>
                  {showDetalle ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                </button>
                </span>
              </div>
            </div>
            {showDetalle && (
              <>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
                    <thead>
                      <tr>
                        {[["Fecha","Fecha"],["Cultivo","Cultivo"],["Campo","Campo"],["Lote","Lote"],["TipoDep","Tipo dep."],["ODT","ODT"],["TipoDepContrap.","Tipo contrap."],["ODT Contrap.","ODT contrap."],["kg", peso === "D" ? "Neto D. (kg)" : "Neto O. (kg)"],["Dominio","Dominio"],["Transporte","Transporte"]].map(([key, label]) => (
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
                          <td className="agri-td">{r["Cultivo"]}</td>
                          <td className="agri-td">{r["Campo"]}</td>
                          <td className="agri-td">{r["Lote"]}</td>
                          <td className="agri-td" style={{ fontSize: 11, color: "var(--ink-soft)" }}>{r["TipoDep"]}</td>
                          <td className="agri-td" style={{ fontSize: 11 }}>{r["ODT"]}</td>
                          <td className="agri-td" style={{ fontSize: 11, color: "var(--ink-soft)" }}>{r["TipoDepContrap."]}</td>
                          <td className="agri-td" style={{ fontSize: 11 }}>{r["ODT Contrap."]}</td>
                          <td className="agri-td" style={{ textAlign: "right", fontWeight: 600, color: r.kg < 0 ? "var(--rust)" : "var(--green)" }}
                            title={peso === "D" && !r["Neto D."] ? "Sin Neto D.: se usa Neto O." : undefined}>
                            {fmtNum(r.kg, 0)}{peso === "D" && !r["Neto D."] ? " *" : ""}
                          </td>
                          <td className="agri-td" style={{ fontSize: 11 }}>{r["Dominio"]}</td>
                          <td className="agri-td" style={{ fontSize: 11 }}>{r["Transporte"]}</td>
                        </tr>
                      ))}
                      {pageRows.length === 0 && (
                        <tr><td className="agri-td" colSpan={11} style={{ textAlign: "center", padding: 30, color: "var(--ink-soft)" }}>Ningún registro coincide con los filtros.</td></tr>
                      )}
                    </tbody>
                  </table>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 16px", borderTop: "1px solid var(--line)" }}>
                  <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>Página {page} de {totalPages}</span>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button className="agri-btn agri-btn-outline" disabled={page <= 1} onClick={() => setPage((p) => p - 1)} style={{ opacity: page <= 1 ? 0.4 : 1 }}>‹</button>
                    <button className="agri-btn agri-btn-outline" disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)} style={{ opacity: page >= totalPages ? 0.4 : 1 }}>›</button>
                  </div>
                </div>
              </>
            )}
          </div>
        </>
      )}
    </div>
  );
}
