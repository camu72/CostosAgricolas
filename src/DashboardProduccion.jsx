import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, Legend, LabelList,
} from "recharts";
import {
  Upload, RefreshCw, FileSpreadsheet, AlertCircle, Trash2,
  ChevronDown, ChevronUp, Download, ArrowUpDown,
} from "lucide-react";
import { ref as dbRef, onValue, set as dbSet } from "firebase/database";

// ---------------------------------------------------------------------------
// Constantes
// ---------------------------------------------------------------------------
const COLS_PROD = [
  "Periodo","Cultivo","Fecha","Comp.","Suc","Num","Dominio","Acoplado",
  "Conductor","Transporte","N° CP","Campo","Lote","Articulo","Titular",
  "Admin","Socio","TipoDep","ODT","TipoDepContrap.","ODT Contrap.",
  "Neto O.","Seco O.","Desc O.","Neto D.","Seco D.","Desc D.",
  "CampoOrigen","SocioContraparte",
];
const STORAGE_KEY_PROD = "produccion-dataset-v1";
const PAGE_SIZE = 30;

const CULTIVO_COLORS = { SOJA: "#5B7C4B", MAIZ: "#C68F41", POROTO: "#35606B", GARBANZO: "#8A5A3B" };
const FALLBACK_COLORS = ["#5B7C4B","#C68F41","#35606B","#8A5A3B","#7C8A4B","#A1462F"];
const cultivoColor = (name, i) => CULTIVO_COLORS[String(name).toUpperCase()] || FALLBACK_COLORS[i % FALLBACK_COLORS.length];

const NULLABLE_COLS_PROD = new Set(["Neto O.","Seco O.","Desc O.","Neto D.","Seco D.","Desc D.","N° CP"]);
const encodeCell = (col, val) => (val === null && NULLABLE_COLS_PROD.has(col) ? "" : val);
const decodeCell = (col, val) => (val === "" && NULLABLE_COLS_PROD.has(col) ? null : val);
const rowsFromPayload = (payload) => payload.rows.map((arr) => {
  const o = {};
  payload.columns.forEach((c, i) => { o[c] = decodeCell(c, arr[i]); });
  return o;
});

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
  return t === "cultivo" || t === "cultivos";
};

// Tipo de movimiento inferido
const tipoMovimiento = (row) => {
  if (esCultivo(row["TipoDep"]) || esCultivo(row["TipoDepContrap."])) return "Cosecha";
  return "Movimiento de grano";
};

function normalizeRowProd(r) {
  return {
    "Periodo":          str(r["Periodo"]),
    "Cultivo":          str(r["Cultivo"]),
    "Fecha":            toISODate(r["Fecha"]),
    "Comp.":            str(r["Comp."]),
    "Suc":              num(r["Suc"]),
    "Num":              num(r["Num"]),
    "Dominio":          str(r["Dominio"]),
    "Acoplado":         str(r["Acoplado"]),
    "Conductor":        str(r["Conductor"]),
    "Transporte":       str(r["Transporte"]),
    "N° CP":            numOrNull(r["N° CP"]),
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
export default function DashboardProduccion({ slug, isAdmin, cloudDb }) {
  const storageKey = `${STORAGE_KEY_PROD}-${slug}`;
  const cloudPath = `clientes/${slug}/produccion`;

  const [rows, setRows] = useState([]);
  const [meta, setMeta] = useState(null);
  const [bootLoading, setBootLoading] = useState(true);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [syncStatus, setSyncStatus] = useState(cloudDb ? "connecting" : "local");
  const fileInputRef = useRef(null);
  const [showDetalle, setShowDetalle] = useState(false);
  const [selectedDeposito, setSelectedDeposito] = useState(null);
  const [gruposAbiertos, setGruposAbiertos] = useState({ planta: true, silo: true, otros: true });
  const toggleGrupo = (g) => setGruposAbiertos((prev) => ({ ...prev, [g]: !prev[g] }));

  // Filtros
  const [buscar, setBuscar]   = useState("");
  const [periodo, setPeriodo] = useState("Todos");
  const [cultivo, setCultivo] = useState("Todos");
  const [admin, setAdmin]     = useState("Todos");
  const [socio, setSocio]     = useState("Todos");
  const [campo, setCampo]     = useState("Todos");
  const [tipoMov, setTipoMov] = useState("Todos");
  const [sort, setSort] = useState({ key: "Fecha", dir: "desc" });
  const [page, setPage] = useState(1);

  // Cargar desde caché local y luego suscribirse a Firebase
  useEffect(() => {
    let unsub = null;
    (async () => {
      try {
        const raw = localStorage.getItem(storageKey);
        if (raw) {
          const payload = JSON.parse(raw);
          setRows(rowsFromPayload(payload));
          setMeta({ fileName: payload.fileName, updatedAt: payload.updatedAt, rowCount: rowsFromPayload(payload).length });
        }
      } catch (e) {}

      if (cloudDb) {
        unsub = onValue(dbRef(cloudDb, cloudPath), (snap) => {
          const payload = snap.val();
          if (payload && payload.rows && payload.columns) {
            const objRows = rowsFromPayload(payload);
            setRows(objRows);
            setMeta({ fileName: payload.fileName, updatedAt: payload.updatedAt, rowCount: objRows.length });
          }
          setSyncStatus("cloud");
          setBootLoading(false);
        }, () => { setSyncStatus("error"); setBootLoading(false); });
      } else {
        setBootLoading(false);
      }
    })();
    return () => { if (unsub) unsub(); };
  }, [slug]);

  const persist = useCallback(async (normalized, fileName) => {
    const payload = {
      columns: COLS_PROD,
      rows: normalized.map((r) => COLS_PROD.map((c) => encodeCell(c, r[c]))),
      fileName, updatedAt: new Date().toISOString(),
    };
    try { localStorage.setItem(storageKey, JSON.stringify(payload)); } catch (e) {}
    if (cloudDb) {
      try { await dbSet(dbRef(cloudDb, cloudPath), payload); setSyncStatus("cloud"); }
      catch (e) { setSyncStatus("error"); }
    }
  }, [slug, cloudDb]);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setParsing(true); setError("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
      if (!json.length) throw new Error("La hoja no tiene datos.");
      const normalized = json.map(normalizeRowProd).filter((r) => r["Cultivo"] || r["Campo"]);
      setRows(normalized);
      setMeta({ fileName: file.name, updatedAt: new Date().toISOString(), rowCount: normalized.length });
      setPage(1);
      await persist(normalized, file.name);
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
  const periodos    = useMemo(() => Array.from(new Set(rows.map((r) => r["Periodo"]).filter(Boolean))).sort(), [rows]);
  const cultivos    = useMemo(() => Array.from(new Set(rows.map((r) => r["Cultivo"]).filter(Boolean))).sort(), [rows]);
  const admins      = useMemo(() => Array.from(new Set(rows.map((r) => r["Admin"]).filter(Boolean))).sort(), [rows]);
  const socios      = useMemo(() => Array.from(new Set(rows.map((r) => r["Socio"]).filter(Boolean))).sort(), [rows]);
  const campos      = useMemo(() => {
    const base = admin === "Todos" ? rows : rows.filter((r) => r["Admin"] === admin);
    return Array.from(new Set(base.map((r) => r["Campo"]).filter((c) => c && !c.toUpperCase().includes("DEPOSITO")))).sort();
  }, [rows, admin]);
  useEffect(() => { if (campo !== "Todos" && !campos.includes(campo)) setCampo("Todos"); }, [campos]);

  // Filtrado
  const buscarQ = buscar.trim().toLowerCase();
  const filtered = useMemo(() => rows.filter((r) => {
    if (periodo !== "Todos" && r["Periodo"] !== periodo) return false;
    if (cultivo !== "Todos" && r["Cultivo"] !== cultivo) return false;
    if (admin   !== "Todos" && r["Admin"]   !== admin)   return false;
    if (socio   !== "Todos" && r["Socio"]   !== socio)   return false;
    if (campo   !== "Todos" && r["Campo"]   !== campo && r["CampoOrigen"] !== campo) return false;
    if (tipoMov !== "Todos" && tipoMovimiento(r) !== tipoMov) return false;
    if (buscarQ) {
      const hay = ["Campo","CampoOrigen","Lote","ODT","ODT Contrap.","Dominio","Transporte","Socio","SocioContraparte","Conductor"]
        .some((k) => str(r[k]).toLowerCase().includes(buscarQ));
      if (!hay) return false;
    }
    return true;
  }), [rows, periodo, cultivo, admin, socio, campo, tipoMov, buscarQ]);

  useEffect(() => { setPage(1); }, [periodo, cultivo, admin, socio, campo, tipoMov, buscarQ]);

  // ---------- KPIs ----------
  const kpis = useMemo(() => {
    // Cosecha: filas donde el TipoDep es un cultivo y Neto O. < 0 (el campo pierde grano)
    const cosechaRows = filtered.filter((r) => esCultivo(r["TipoDep"]) && (r["Neto O."] || 0) < 0);
    const tnCosechadas = cosechaRows.reduce((s, r) => s + Math.abs(r["Neto O."] || 0), 0) / 1000;

    // Depósitos: filas cuyo TipoDep NO es cultivo → acumulamos Neto O.
    const deposRows = filtered.filter((r) => !esCultivo(r["TipoDep"]));
    let stockPlanta = 0, stockSilo = 0, stockPermanente = 0;
    deposRows.forEach((r) => {
      const v = (r["Neto O."] || 0) / 1000;
      const t = str(r["TipoDep"]).toLowerCase();
      if (t === "planta de acopio") stockPlanta += v;
      else if (t === "temporal") stockSilo += v;
      else if (t === "permanente") stockPermanente += v;
    });

    const viajeRows = filtered.filter((r) => r["Dominio"] && (r["Neto O."] || 0) > 0);
    return { tnCosechadas, stockPlanta, stockSilo, stockPermanente, viajes: viajeRows.length };
  }, [filtered]);

  // ---------- Gráficos ----------
  const cosechaPorCultivo = useMemo(() => {
    const m = new Map();
    filtered.filter((r) => esCultivo(r["TipoDep"]) && (r["Neto O."] || 0) < 0)
      .forEach((r) => {
        const c = r["Cultivo"] || "(sin cultivo)";
        m.set(c, (m.get(c) || 0) + Math.abs(r["Neto O."] || 0) / 1000);
      });
    return Array.from(m.entries()).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  const cosechaPorCampo = useMemo(() => {
    const m = new Map();
    filtered.filter((r) => esCultivo(r["TipoDep"]) && (r["Neto O."] || 0) < 0)
      .forEach((r) => {
        const c = r["Campo"] || r["CampoOrigen"] || "(sin campo)";
        if (c.toUpperCase().includes("DEPOSITO")) return;
        m.set(c, (m.get(c) || 0) + Math.abs(r["Neto O."] || 0) / 1000);
      });
    return Array.from(m.entries()).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  const evolucionMensual = useMemo(() => {
    const m = new Map();
    const cultList = cultivos;
    filtered.filter((r) => esCultivo(r["TipoDep"]) && (r["Neto O."] || 0) < 0 && r["Fecha"])
      .forEach((r) => {
        const ym = r["Fecha"].slice(0, 7);
        if (!m.has(ym)) m.set(ym, {});
        const bucket = m.get(ym);
        const c = r["Cultivo"] || "(sin cultivo)";
        bucket[c] = (bucket[c] || 0) + Math.abs(r["Neto O."] || 0) / 1000;
      });
    return Array.from(m.entries()).sort(([a], [b]) => a.localeCompare(b)).map(([ym, bucket]) => ({
      ym, label: monthLabel(ym), ...bucket,
    }));
  }, [filtered, cultivos]);

  const stockPorDeposito = useMemo(() => {
    const m = new Map();
    filtered.filter((r) => !esCultivo(r["TipoDep"])).forEach((r) => {
      const key = `${r["TipoDep"]}||${r["ODT"]}`;
      if (!m.has(key)) m.set(key, { tipo: r["TipoDep"], nombre: r["ODT"], stock: 0 });
      m.get(key).stock += (r["Neto O."] || 0) / 1000;
    });
    return Array.from(m.values()).filter((e) => Math.abs(e.stock) > 0.01).sort((a, b) => b.stock - a.stock);
  }, [filtered]);

  // stockPlantaChart / stockSiloChart ya no se usan — reemplazados por el acordeón unificado

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

  const resetFiltros = () => { setBuscar(""); setPeriodo("Todos"); setCultivo("Todos"); setAdmin("Todos"); setSocio("Todos"); setCampo("Todos"); setTipoMov("Todos"); };
  const hayFiltros = buscar !== "" || periodo !== "Todos" || cultivo !== "Todos" || admin !== "Todos" || socio !== "Todos" || campo !== "Todos" || tipoMov !== "Todos";

  // -------------------------------------------------------------------------
  return (
    <div className="agri-shell">
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
            <div style={{ fontSize: 16, fontWeight: 700 }}>{parsing ? "Leyendo…" : "Arrastrá el Excel de Producción"}</div>
            <div style={{ fontSize: 13, color: "var(--ink-soft)", maxWidth: 380 }}>
              Planilla de movimientos con columnas ODT, TipoDep, Neto O., etc.
            </div>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
              onChange={(e) => handleFile(e.target.files?.[0])} />
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
          {/* Sub-header de este dashboard */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
              {meta.rowCount.toLocaleString("es-AR")} movimientos · actualizado {new Date(meta.updatedAt).toLocaleString("es-AR",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"})}
            </div>
            {isAdmin && (
              <div style={{ display: "flex", gap: 8 }}>
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files?.[0])} />
                <button className="agri-btn" onClick={() => fileInputRef.current?.click()}><RefreshCw size={13} /> Actualizar datos</button>
                <button className="agri-btn agri-btn-outline" onClick={clearDataset}><Trash2 size={13} /></button>
              </div>
            )}
          </div>

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
              <select className="agri-select" value={periodo}  onChange={(e) => setPeriodo(e.target.value)}>
                <option value="Todos">Todos los períodos</option>
                {periodos.map((p) => <option key={p} value={p}>{p}</option>)}
              </select>
              <select className="agri-select" value={admin}    onChange={(e) => setAdmin(e.target.value)}>
                <option value="Todos">Todas las admins</option>
                {admins.map((a) => <option key={a} value={a}>{a}</option>)}
              </select>
              <select className="agri-select" value={socio}    onChange={(e) => setSocio(e.target.value)}>
                <option value="Todos">Todos los socios</option>
                {socios.map((s) => <option key={s} value={s}>{s}</option>)}
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
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 12, marginBottom: 20 }}>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Tn cosechadas</div>
              <div className="agri-kpi-value">{fmtNum(kpis.tnCosechadas, 1)}</div>
            </div>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Stock en planta</div>
              <div className="agri-kpi-value">{fmtNum(kpis.stockPlanta, 1)} tn</div>
            </div>
            <div className="agri-card" style={{ padding: 14 }}>
              <div className="agri-kpi-label">Stock en silo bolsa</div>
              <div className="agri-kpi-value">{fmtNum(kpis.stockSilo, 1)} tn</div>
            </div>
            {kpis.stockPermanente !== 0 && (
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Stock en otros dep.</div>
                <div className="agri-kpi-value">{fmtNum(kpis.stockPermanente, 1)} tn</div>
              </div>
            )}
          </div>

          {/* Gráficos — fila 1: cosecha */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(320px,1fr))", gap: 14, marginBottom: 14 }}>
            <div className="agri-card" style={{ padding: 16 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Cosecha por cultivo (tn)</div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={cosechaPorCultivo} margin={{ left: -10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                  <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                  <Tooltip content={<ChartTooltip unit="tn" />} />
                  <Bar dataKey="value" name="Cosechado" radius={[3,3,0,0]}>
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
                  <Bar dataKey="value" name="Cosechado" fill="#B8842E" radius={[0,3,3,0]}>
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
                    <Line key={c} type="monotone" dataKey={c} name={c} stroke={cultivoColor(c, i)} strokeWidth={2} dot={false} />
                  ))}
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>

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
              const totalEntradas = allMovs.filter((r) => (r["Neto O."] || 0) > 0).reduce((s, r) => s + (r["Neto O."] || 0), 0) / 1000;
              const totalSalidas  = allMovs.filter((r) => (r["Neto O."] || 0) < 0).reduce((s, r) => s + Math.abs(r["Neto O."] || 0), 0) / 1000;
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
                              {[["Fecha","Fecha"],["Cultivo","Cultivo"],["Campo","Campo / Origen"],["ODT Contrap.","Contraparte"],["Neto O.","Kg (neto)"]].map(([k,l]) => (
                                <th key={k} className="agri-th">{l}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {allMovs.map((r, i) => {
                              const kg = r["Neto O."] || 0;
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
              <div className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
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
          <div className="agri-card" style={{ padding: 0, overflow: "hidden" }}>
            <div style={{ padding: "12px 16px", borderBottom: showDetalle ? "1px solid var(--line)" : "none", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
              <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Detalle de movimientos</div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(sorted.length)} registros</span>
                <button className="agri-btn agri-btn-outline" onClick={exportToExcel}><Download size={13} /> Exportar</button>
                <button className="agri-btn agri-btn-outline" onClick={() => setShowDetalle((v) => !v)}>
                  {showDetalle ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                </button>
              </div>
            </div>
            {showDetalle && (
              <>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
                    <thead>
                      <tr>
                        {[["Fecha","Fecha"],["Cultivo","Cultivo"],["Campo","Campo"],["Lote","Lote"],["TipoDep","Tipo dep."],["ODT","ODT"],["TipoDepContrap.","Tipo contrap."],["ODT Contrap.","ODT contrap."],["Neto O.","Neto O. (kg)"],["Dominio","Dominio"],["Transporte","Transporte"]].map(([key, label]) => (
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
                          <td className="agri-td" style={{ textAlign: "right", fontWeight: 600, color: (r["Neto O."] || 0) < 0 ? "var(--rust)" : "var(--green)" }}>
                            {r["Neto O."] === null ? "—" : fmtNum(r["Neto O."], 0)}
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
