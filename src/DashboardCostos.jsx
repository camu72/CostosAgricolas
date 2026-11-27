import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { ref, onValue, set as dbSet } from "firebase/database";
import { CLOUD_CLIENTS_BASE } from "./firebaseConfig.js";
import { PdfMenu, generarPDF, filaGrupo, filaTotal, MODOS_PDF } from "./pdfExport.jsx";
import { leerFilasArchivo, validarColumnas, leerEncabezado, leerCultivos, mapaColoresCultivos, colorDeCultivo, ACCEPT_ARCHIVOS } from "./lectorArchivo.js";
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
// Costos: datos generados por la pantalla "Dashboards > Costos" de Gorrión.
// Sigue el modelo de Cultivos (20/11/2026). Reemplaza al dashboard de Costos anterior: mismo nodo
// (antes clientes/<slug>/dataset; ahora clientes/<slug>/costos, como publica Gorrión), sin compatibilidad con los datos viejos.
// Una fila por ítem de una OTA aplicado a una línea de cultivo, de UNA campaña (sin selector).
// Llega por tres vías y siempre se guarda en clientes/<slug>/costos:
//   - publicación directa desde Gorrión (PUT del JSON),
//   - "Actualizar datos" con el JSON que genera Gorrión,
//   - "Actualizar datos" con el Excel que exporta Gorrión.
// Reglas de cálculo acordadas con el usuario (20/11/2026):
//   - Cultivo = "Habilitación": costo de habilitar el lote. Se muestra como categoría aparte,
//     separada de los cultivos (no cuenta como cultivo ni se suma a otro).
//   - Superficie: Sup.Cultivo una vez por Campo + Lote + Cultivo + Variedad (como antes).
//   - Insumos y servicios llegan con precio; una fila sin Total se marca "sin precio".
//   - Hectáreas de la labor: Cant.Afect (las de esa línea); Cant.Orig. = las de toda la OTA.
// ---------------------------------------------------------------------------
export const NODO_COSTOS = "costos";
const COLS = [
  "Admin", "Campaña", "Campo", "Lote", "Sup.Lote", "Cultivo", "Variedad",
  "Sup.Cultivo", "Fecha", "OTA", "Origen", "Cant.Orig.", "U.Serv.",
  "Tipo Det.", "Tipo item", "Concepto", "Cant.Afect", "Cantidad", "Unid.",
  "Precio/U", "Total",
];

const REQUERIDAS = ["Campo", "Lote", "Cultivo", "Sup.Cultivo", "Fecha", "Tipo Det.", "Tipo item", "Concepto", "Cantidad", "Total"];
const STORAGE_KEY_PREFIX = "costos-gorrion-v1";   // distinta de la del dashboard anterior: no lee su caché
const PAGE_SIZE = 30;
// Precio/U y Total pueden venir vacíos (""): quedan en null y se cuentan como "sin precio".
// Filas guardadas o publicadas por Gorrión ({columns, rows}) → objetos normalizados.
// Los encabezados se comparan sin espacios ni puntos y sin distinguir mayúsculas.
const normKey = (k) => String(k || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s.]/g, "").toLowerCase();
const COLS_POR_CLAVE = new Map(COLS.map((c) => [normKey(c), c]));
const alinearColumnas = (o) => {
  const r = {};
  Object.keys(o).forEach((k) => { r[COLS_POR_CLAVE.get(normKey(k)) || k] = o[k]; });
  return r;
};
const rowsFromPayload = (payload) => payload.rows.map((arr) => {
  const o = {};
  payload.columns.forEach((c, i) => { o[c] = arr[i]; });
  return normalizeRow(alinearColumnas(o));
}).filter((r) => r["Campo"] || r["Cultivo"] || r["Concepto"]);
// Habilitación del lote: va como categoría aparte, nunca como un cultivo más.
const esHabilitacion = (c) => normKey(c) === "habilitacion";
const GRIS_HABILITACION = "#B9AE94";

const CULTIVO_COLORS = { SOJA: "#5B7C4B", MAIZ: "#C68F41", POROTO: "#35606B", GARBANZO: "#8A5A3B", TRIGO: "#B8842E", GIRASOL: "#A1462F" };
// Colores para cultivos sin color propio: distintos de los fijos de arriba, así dos cultivos no comparten color
const FALLBACK_COLORS = ["#7C8A4B", "#6E5A8A", "#3F7F6E", "#B26A8E", "#8C7A3A", "#4F6F8F"];
// Colores de los cultivos: primero los que manda Gorrión en la sección "cultivos" del JSON
// (administrador de cultivos de cada Gorrión); si no vino, la paleta fija y después la de respaldo.
let coloresGorrion = new Map();
const cultivoColor = (name, i) => colorDeCultivo(coloresGorrion, name) || (esHabilitacion(name) ? GRIS_HABILITACION : (CULTIVO_COLORS[String(name).toUpperCase()] || FALLBACK_COLORS[i % FALLBACK_COLORS.length]));
// Orden de cultivos: alfabético, con Habilitación siempre al final
const ordenCultivos = (a, b) => (esHabilitacion(a) - esHabilitacion(b)) || String(a).localeCompare(String(b), "es");

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const str = (v) => (v === null || v === undefined ? "" : String(v).trim());
// Números: acepta 1234.5, "1234,5", "1.234,5" y "1,234.5"
const parseNum = (v) => {
  if (typeof v === "number") return isNaN(v) ? NaN : v;
  let s = String(v ?? "").replace(/[^0-9,.\-]/g, "");
  if (!s) return NaN;
  const coma = s.lastIndexOf(","), punto = s.lastIndexOf(".");
  if (coma >= 0 && punto >= 0) s = coma > punto ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  else if (coma >= 0) s = s.replace(",", ".");
  return parseFloat(s);
};
const num = (v) => { const n = parseNum(v); return isNaN(n) ? 0 : n; };
const numOrNull = (v) => {
  if (v === null || v === undefined || String(v).trim() === "") return null;
  const n = parseNum(v);
  return isNaN(n) ? null : n;
};
// Las fechas que entrega SheetJS (cellDates:true) están ancladas en UTC medianoche.
// Usamos los getters UTC para evitar que el huso horario local corra el día (p.ej. UTC-3).
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
// Etiquetas dentro de las barras. "formato" permite usarlas para hectáreas (vista Labores).
const VerticalBarLabel = (props) => {
  const { x, y, width, height, value, formato = fmtUSD2 } = props;
  if (!value || height < 16) return null;
  const text = formato(value);
  if (width < text.length * 6.2 + 4) return null;
  return (
    <text x={x + width / 2} y={y + height / 2} textAnchor="middle" dominantBaseline="middle" fill="#fff" fontSize={11} fontWeight={700}>
      {text}
    </text>
  );
};
const fmtHa = (v) => `${fmtNum(v, v < 10 ? 1 : 0)} ha`;

const HorizontalBarLabel = (props) => {
  const { x, y, width, height, value, formato = fmtUSD2 } = props;
  if (!value) return null;
  const text = formato(value);
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
    "Precio/U": numOrNull(r["Precio/U"]),
    "Total": numOrNull(r["Total"]),
  };
}


// ---------------------------------------------------------------------------
// Resumen de un lote: Tipo Det. (Insumos, Servicios) > Rubro > Concepto.
// Lo usan el detalle en pantalla y el PDF "Completo por lote", así los números coinciden.
// ---------------------------------------------------------------------------
const TIPODET_ORDER_LOTE = { "INSUMOS": 0, "SERVICIOS": 1 };
function construirResumenLote(rows, ha) {
  const detMap = new Map();
  rows.forEach((r) => {
    const tipoDet = r["Tipo Det."] || "(sin tipo)";
    if (!detMap.has(tipoDet)) detMap.set(tipoDet, { items: 0, cant: 0, costo: 0, rubros: new Map() });
    const dg = detMap.get(tipoDet);
    dg.items += 1; dg.cant += r["Cantidad"] || 0; dg.costo += r["Total"] || 0;
    const rubro = r["Tipo item"] || "(sin rubro)";
    if (!dg.rubros.has(rubro)) dg.rubros.set(rubro, { items: 0, cant: 0, costo: 0, conceptos: new Map() });
    const rg = dg.rubros.get(rubro);
    rg.items += 1; rg.cant += r["Cantidad"] || 0; rg.costo += r["Total"] || 0;
    const concepto = r["Concepto"] || "(sin concepto)";
    if (!rg.conceptos.has(concepto)) rg.conceptos.set(concepto, { items: 0, cant: 0, costo: 0, unid: r["Unid."], sumUnit: 0, countUnit: 0 });
    const cg = rg.conceptos.get(concepto);
    cg.items += 1; cg.cant += r["Cantidad"] || 0; cg.costo += r["Total"] || 0;
    if (r["Precio/U"] !== null) { cg.sumUnit += r["Precio/U"]; cg.countUnit += 1; }
  });
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
    .sort((a, b) => (TIPODET_ORDER_LOTE[a.tipoDet] ?? 2) - (TIPODET_ORDER_LOTE[b.tipoDet] ?? 2));
}
// OTA abreviada: "OTA N° 020-001008 C25/26 - SA" -> "OTA 020-001008" (tipo + número)
const otaCorta = (ota) => {
  const m = String(ota || "").match(/^(\S+)\s+N[°º]?\s*(\d+-\d+)/);
  return m ? `${m[1]} ${m[2]}` : String(ota || "");
};

// Hectáreas de la labor: columna COL_HAS_LABOR ("Cant.Afect" = has afectadas a esta línea
// de cultivo; alternativa "Cant.Orig." = has de toda la orden de labor), en la unidad del
// servicio ("U.Serv.": Has → ha, Hs → hs). Si falta, se usan las ha del cultivo.
const COL_HAS_LABOR = "Cant.Afect";
const unidadLabor = (r) => {
  const u = String(r["U.Serv."] || "").trim().toLowerCase();
  return !u || u.startsWith("ha") ? "ha" : u;
};
const dosisLabor = (r, haCultivo) => {
  const hl = r[COL_HAS_LABOR];
  const base = hl > 0 ? hl : haCultivo;
  return base > 0 ? { valor: (r["Cantidad"] || 0) / base, unidad: hl > 0 ? unidadLabor(r) : "ha" } : null;
};
// Hectáreas de cada grupo fecha + labor + OTA (índice de la primera fila → { has, unidad }).
// Si el lote tiene varias líneas de cultivo (p. ej. una resiembra con la misma variedad),
// se suman las hectáreas de cada línea (identificada por su Sup.Cultivo).
function hasPorGrupoMovs(movs) {
  const res = {};
  let ini = 0;
  for (let i = 1; i <= movs.length; i++) {
    const a = movs[i - 1], b = movs[i];
    const corte = !b || b["Fecha"] !== a["Fecha"] || b["Origen"] !== a["Origen"] || b["OTA"] !== a["OTA"];
    if (!corte) continue;
    const porLinea = new Map();
    movs.slice(ini, i).forEach((r) => { if (r[COL_HAS_LABOR] > 0 && !porLinea.has(r["Sup.Cultivo"])) porLinea.set(r["Sup.Cultivo"], r[COL_HAS_LABOR]); });
    const has = Array.from(porLinea.values()).reduce((x, y) => x + y, 0);
    res[ini] = has > 0 ? { has, unidad: unidadLabor(movs[ini]) } : null;
    ini = i;
  }
  return res;
}

// Movimientos de un lote en el orden de la pantalla: fecha, OTA, rubro, concepto.
// Cada grupo es fecha + labor + OTA, y queda siempre en un bloque continuo.
function ordenarMovimientosLote(rows) {
  return rows.slice().sort((a, b) => {
    const fa = a["Fecha"] || "", fb = b["Fecha"] || "";
    if (fa !== fb) return fa < fb ? -1 : 1;
    const oa = a["OTA"] || "", ob = b["OTA"] || "";
    if (oa !== ob) return oa.localeCompare(ob, "es", { numeric: true });
    const ra = a["Tipo item"] || "", rb = b["Tipo item"] || "";
    if (ra !== rb) return ra.localeCompare(rb, "es");
    const ca = a["Concepto"] || "", cb = b["Concepto"] || "";
    return ca.localeCompare(cb, "es");
  });
}
// Costo por rubro de un lote (para el gráfico de composición)
function costoPorRubroLote(rows) {
  const m = new Map();
  rows.forEach((r) => {
    if (!r["Tipo item"]) return;
    const key = r["Tipo item"];
    if (!m.has(key)) m.set(key, { value: 0, tipoDet: r["Tipo Det."] || "" });
    m.get(key).value += r["Total"] || 0;
  });
  return Array.from(m.entries())
    .map(([name, v]) => ({ name, value: v.value, tipoDet: v.tipoDet }))
    .sort((a, b) => (TIPODET_ORDER_LOTE[a.tipoDet] ?? 2) - (TIPODET_ORDER_LOTE[b.tipoDet] ?? 2) || b.value - a.value);
}

// ---------------------------------------------------------------------------
// Vista "Labores" (27/11/2026): seguimiento de labores sin precios.
// Reglas acordadas con el usuario:
//   - la labor es "Origen"; cada OTA es una labor; toda labor (propia o contratada) tiene su fila de servicio;
//   - hectáreas trabajadas = Cant.Afect de la fila de SERVICIO, una vez por OTA + línea de cultivo
//     (si una OTA tiene dos servicios en la misma línea, las ha no se duplican);
//   - contratista = Concepto de la fila de servicio;
//   - servicios en otra unidad (U.Serv. = Hs, etc.) no suman hectáreas: se informan aparte.
// ---------------------------------------------------------------------------
const esServicio = (r) => r["Tipo Det."] === "SERVICIOS";
const esInsumo = (r) => r["Tipo Det."] === "INSUMOS";
const claveLinea = (r) => `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}|${r["Sup.Cultivo"]}`;

// Trabajo (ha u otra unidad) de cada OTA + línea, tomado una sola vez de su fila de servicio
function trabajosLabores(rows) {
  const m = new Map();
  rows.forEach((r) => {
    if (!esServicio(r) || !(r[COL_HAS_LABOR] > 0)) return;
    const k = `${r["OTA"]}|${claveLinea(r)}`;
    if (m.has(k)) return;
    m.set(k, { ota: r["OTA"], labor: r["Origen"] || "(sin labor)", cultivo: r["Cultivo"], fecha: r["Fecha"], contratista: r["Concepto"] || "(sin contratista)",
      lote: `${r["Campo"]}|${r["Lote"]}`, cant: r[COL_HAS_LABOR], unidad: unidadLabor(r) });
  });
  return Array.from(m.values());
}
function resumenLabores(rows) {
  const otas = new Set(), insumos = new Set(), contratistas = new Set();
  let ultima = null;
  rows.forEach((r) => {
    if (r["OTA"]) otas.add(r["OTA"]);
    if (esInsumo(r) && r["Concepto"]) insumos.add(r["Concepto"]);
    if (esServicio(r) && r["Concepto"]) contratistas.add(r["Concepto"]);
    if (r["Fecha"] && (!ultima || r["Fecha"] > ultima.fecha)) ultima = { fecha: r["Fecha"], labor: r["Origen"] };
  });
  let ha = 0; const otras = {};
  trabajosLabores(rows).forEach((t) => { if (t.unidad === "ha") ha += t.cant; else otras[t.unidad] = (otras[t.unidad] || 0) + t.cant; });
  return { labores: otas.size, ha, otras, insumos: insumos.size, contratistas: contratistas.size, ultima };
}
const textoOtras = (otras) => Object.entries(otras || {}).map(([u, v]) => `${fmtNum(v, 1)} ${u}`).join(" · ");

// Hectáreas trabajadas por labor, desglosadas por cultivo (para el gráfico apilado)
function haPorLabor(rows, cultivosOrden) {
  const m = new Map();
  trabajosLabores(rows).forEach((t) => {
    if (t.unidad !== "ha") return;
    if (!m.has(t.labor)) m.set(t.labor, { name: t.labor, _total: 0, _otas: new Set() });
    const e = m.get(t.labor);
    e[t.cultivo] = (e[t.cultivo] || 0) + t.cant; e._total += t.cant; e._otas.add(t.ota);
  });
  return Array.from(m.values()).map((e) => {
    const o = { name: e.name, _total: e._total, labores: e._otas.size };
    (cultivosOrden || []).forEach((c) => { o[c] = e[c] || 0; });
    return o;
  }).sort((a, b) => b._total - a._total);
}
// Hectáreas trabajadas por mes
function haPorMes(rows) {
  const m = new Map();
  trabajosLabores(rows).forEach((t) => {
    if (t.unidad !== "ha" || !t.fecha) return;
    const ym = t.fecha.slice(0, 7);
    m.set(ym, (m.get(ym) || 0) + t.cant);
  });
  const meses = Array.from(m.keys()).sort();
  if (!meses.length) return [];
  const sig = (ym) => { const [y, mo] = ym.split("-").map(Number); return mo === 12 ? `${y + 1}-01` : `${y}-${String(mo + 1).padStart(2, "0")}`; };
  const out = [];
  for (let ym = meses[0]; ym <= meses[meses.length - 1]; ym = sig(ym)) out.push({ ym, name: monthLabel(ym), ha: m.get(ym) || 0 });
  return out;
}
// Resumen por contratista
function resumenContratistas(rows) {
  const m = new Map();
  trabajosLabores(rows).forEach((t) => {
    if (!m.has(t.contratista)) m.set(t.contratista, { contratista: t.contratista, otas: new Set(), lotes: new Set(), labores: new Set(), ha: 0, otras: {}, ultima: "" });
    const e = m.get(t.contratista);
    e.otas.add(t.ota); e.lotes.add(t.lote); e.labores.add(t.labor);
    if (t.unidad === "ha") e.ha += t.cant; else e.otras[t.unidad] = (e.otras[t.unidad] || 0) + t.cant;
    if (t.fecha && t.fecha > e.ultima) e.ultima = t.fecha;
  });
  return Array.from(m.values()).map((e) => ({
    contratista: e.contratista, labores: e.otas.size, lotes: e.lotes.size, tiposLabor: Array.from(e.labores).sort().join(", "),
    ha: e.ha, otras: e.otras, ultima: e.ultima,
  })).sort((a, b) => b.ha - a.ha);
}
// Insumos aplicados: producto + unidad (nunca se suman unidades distintas)
function insumosAplicados(rows) {
  const m = new Map();
  rows.forEach((r) => {
    if (!esInsumo(r)) return;
    const concepto = r["Concepto"] || "(sin concepto)";
    const k = `${concepto}|${r["Unid."]}`;
    if (!m.has(k)) m.set(k, { concepto, rubro: r["Tipo item"] || "(sin rubro)", unid: r["Unid."] || "", cant: 0, ha: 0, min: null, max: null, movimientos: 0, lotes: new Set() });
    const e = m.get(k);
    e.cant += r["Cantidad"] || 0; e.movimientos += 1; e.lotes.add(`${r["Campo"]}|${r["Lote"]}`);
    if (r[COL_HAS_LABOR] > 0 && unidadLabor(r) === "ha") {
      e.ha += r[COL_HAS_LABOR];
      const d = (r["Cantidad"] || 0) / r[COL_HAS_LABOR];
      if (e.min === null || d < e.min) e.min = d;
      if (e.max === null || d > e.max) e.max = d;
    }
  });
  return Array.from(m.values()).map((e) => ({ ...e, lotes: e.lotes.size, dosis: e.ha > 0 ? e.cant / e.ha : null }))
    .sort((a, b) => a.rubro.localeCompare(b.rubro, "es") || a.concepto.localeCompare(b.concepto, "es"));
}
// Dosis con hasta 3 decimales significativos (0,0035 Lts/ha no queda en 0,00)
const fmtDosis = (v) => (v === null || v === undefined ? "—" : fmtNum(v, v !== 0 && Math.abs(v) < 0.1 ? 4 : 2));


// ---------------------------------------------------------------------------
// PDF "Informe completo" con detalle por lote: para cada lote, encabezado + indicadores +
// composición por rubro (barras) + resumen por rubro completo (sin scroll).
// ---------------------------------------------------------------------------
const MODOS_PDF_COSTOS = [
  ...MODOS_PDF.filter((m) => m.key !== "completo"),
  { key: "completo", label: "Informe completo", desc: "Resumen, comparativo por lote y precios por ítem",
    opciones: [
      { key: "detalleLotes", label: "Incluir detalle por lote" },
      { key: "tipoDetalle", tipo: "radio", requiere: "detalleLotes", def: "resumen",
        valores: [{ value: "resumen", label: "Resumen por rubro" }, { value: "movimientos", label: "Movimientos" }] },
      { key: "hojaNueva", label: "Cada lote en hoja nueva", requiere: "detalleLotes" },
    ] },
];
// Vista Labores: el informe completo es el cuaderno de campo (movimientos por lote, sin precios)
const MODOS_PDF_LABORES = [
  ...MODOS_PDF.filter((m) => m.key !== "completo"),
  { key: "completo", label: "Cuaderno de campo", desc: "Comparativo por lote, insumos aplicados, contratistas y labores de cada lote",
    opciones: [
      { key: "detalleLotes", label: "Incluir detalle por lote" },
      { key: "tipoDetalle", tipo: "radio", requiere: "detalleLotes", def: "movimientos",
        valores: [{ value: "movimientos", label: "Labores (cuaderno)" }, { value: "resumen", label: "Resumen por rubro" }] },
      { key: "hojaNueva", label: "Cada lote en hoja nueva", requiere: "detalleLotes" },
    ] },
];
const COLOR_INSUMOS = [75, 107, 58], COLOR_SERVICIOS = [184, 132, 46];

async function dibujarResumenLotes(api, lotes, hojaNueva, tipoDetalle = "resumen", sinPrecios = false) {
  const esMovs = tipoDetalle === "movimientos";
  const { doc, autoTable, pdfTxt, C, PAGE, TOP, BOTTOM, CONTENT_W } = api;
  const CHART_W = 78, GAP = 8, TABLE_X = PAGE.m + CHART_W + GAP;

  lotes.forEach(({ info, rows }, idx) => {
    const resumen = construirResumenLote(rows, info.ha);
    // Sin precios: el gráfico muestra hectáreas trabajadas por labor en lugar del costo por rubro
    const porRubro = sinPrecios
      ? haPorLabor(rows, []).map((e) => ({ name: e.name, value: e._total, tipoDet: "SERVICIOS" }))
      : costoPorRubroLote(rows);
    const altoChart = 8 + porRubro.length * 5.2 + 10;

    // ---- Salto de página ----
    if (idx === 0) {
      if (hojaNueva || BOTTOM - api.getY() < 60) api.nuevaPagina();
      doc.setFont("helvetica", "bold"); doc.setFontSize(13); doc.setTextColor(...C.ink);
      doc.text(pdfTxt(`${sinPrecios && esMovs ? "Cuaderno de campo por lote" : `Detalle por lote: ${esMovs ? "movimientos" : "resumen por rubro"}`} (${lotes.length})`), PAGE.m, api.getY() + 4);
      api.setY(api.getY() + 9);
    } else if (hojaNueva) {
      api.nuevaPagina();
    } else if (BOTTOM - api.getY() < Math.min(95, 38 + altoChart)) {
      api.nuevaPagina();
    } else {
      doc.setDrawColor(...C.ink); doc.setLineWidth(0.3);
      doc.line(PAGE.m, api.getY() - 3, PAGE.w - PAGE.m, api.getY() - 3);
    }
    let y = api.getY();

    // ---- Encabezado del lote ----
    doc.setFont("helvetica", "bold"); doc.setFontSize(12); doc.setTextColor(...C.ink);
    doc.text(pdfTxt(`${info.campo} · Lote ${info.lote} · ${info.cultivo}`), PAGE.m, y + 4);
    doc.setFont("helvetica", "normal"); doc.setFontSize(8.5); doc.setTextColor(...C.soft);
    doc.text(pdfTxt(`${info.variedad ? info.variedad + " · " : ""}${fmtNum(info.ha, 1)} ha · ${fmtNum(rows.length)} movimientos`), PAGE.m, y + 8.5);
    y += 12;

    // ---- Indicadores ----
    const rl = sinPrecios ? resumenLabores(rows) : null;
    const kpis = sinPrecios ? [
      ["Labores", fmtNum(rl.labores)], ["Hectáreas trabajadas", `${fmtNum(rl.ha, 1)} ha${textoOtras(rl.otras) ? " · " + textoOtras(rl.otras) : ""}`],
      ["Insumos aplicados", fmtNum(rl.insumos)], ["Última labor", rl.ultima ? `${fmtDate(rl.ultima.fecha)} · ${rl.ultima.labor}` : "-"],
    ] : [
      ["Costo total", fmtUSD(info.total)], ["Insumos", fmtUSD(info.insumos)],
      ["Servicios", fmtUSD(info.servicios)], ["Costo por hectárea", fmtUSD2(info.costoHa)],
    ];
    const kw = CONTENT_W / kpis.length;
    kpis.forEach(([l, v], i) => {
      doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(...C.soft);
      doc.text(pdfTxt(l), PAGE.m + i * kw, y + 3);
      doc.setFont("helvetica", "bold"); doc.setFontSize(sinPrecios && i === 3 ? 9 : 11); doc.setTextColor(...C.ink);
      doc.text(doc.splitTextToSize(pdfTxt(v), kw - 3)[0], PAGE.m + i * kw, y + 8);
    });
    y += 13;
    const yBloque = y;
    const paginaBloque = doc.getNumberOfPages();

    // ---- Composición del costo por rubro (barras) ----
    doc.setFont("helvetica", "bold"); doc.setFontSize(9); doc.setTextColor(...C.ink);
    doc.text(sinPrecios ? "Hectáreas trabajadas por labor" : "Composición del costo por rubro", PAGE.m, y + 3);
    let cy = y + 7;
    const LABEL_W = 30, BAR_W = CHART_W - LABEL_W - 16;
    const max = Math.max(1, ...porRubro.map((r) => r.value));
    porRubro.forEach((r) => {
      doc.setFont("helvetica", "normal"); doc.setFontSize(6.5); doc.setTextColor(...C.ink);
      const lab = doc.splitTextToSize(pdfTxt(r.name), LABEL_W - 1);
      doc.text(lab.slice(0, 2), PAGE.m + LABEL_W - 1, cy + (lab.length > 1 ? 1.3 : 2.6), { align: "right" });
      const w = Math.max(0.6, (r.value / max) * BAR_W);
      doc.setFillColor(...(r.tipoDet === "SERVICIOS" ? COLOR_SERVICIOS : COLOR_INSUMOS));
      doc.rect(PAGE.m + LABEL_W, cy, w, 3.6, "F");
      doc.setFontSize(6.3); doc.setTextColor(...C.soft);
      doc.text(pdfTxt(sinPrecios ? `${fmtNum(r.value, 1)} ha` : fmtUSD(r.value)), PAGE.m + LABEL_W + w + 1, cy + 2.6);
      cy += 5.2;
    });
    // leyenda
    cy += 2;
    (sinPrecios ? [] : [["Insumos", COLOR_INSUMOS], ["Servicios", COLOR_SERVICIOS]]).forEach(([l, col], i) => {
      const lx = PAGE.m + LABEL_W + i * 22;
      doc.setFillColor(...col); doc.rect(lx, cy, 2.6, 2.6, "F");
      doc.setFontSize(7); doc.setTextColor(...C.ink); doc.text(l, lx + 3.6, cy + 2.2);
    });
    const finChart = cy + 6;

    // ---- Tabla del detalle: Resumen por rubro o Movimientos (completa, sin scroll) ----
    const limpiar = (c) => (c && typeof c === "object" ? { ...c, content: pdfTxt(c.content) } : pdfTxt(c));
    const continuacion = () => {
      if (doc.internal.getCurrentPageInfo().pageNumber === paginaBloque) return;
      doc.setFont("helvetica", "bold"); doc.setFontSize(10); doc.setTextColor(...C.ink);
      doc.text(pdfTxt(`${info.campo} · Lote ${info.lote} · ${info.cultivo}`), PAGE.m, TOP + 3);
      doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(...C.soft);
      doc.text(pdfTxt(`${info.variedad ? info.variedad + " · " : ""}${esMovs ? "movimientos" : "resumen por rubro"} (continuación)`), PAGE.m, TOP + 7);
    };
    const base = {
      startY: yBloque,
      margin: { left: TABLE_X, right: PAGE.m, top: TOP + 11, bottom: PAGE.h - BOTTOM },
      didDrawPage: continuacion,
      styles: { font: "helvetica", fontSize: 7.5, cellPadding: 1.2, textColor: C.ink, lineColor: C.line, lineWidth: 0.1 },
      headStyles: { fillColor: C.paper, textColor: C.soft, fontStyle: "bold", fontSize: 7, lineWidth: 0 },
      bodyStyles: { fillColor: [255, 255, 255] },
      showHead: "everyPage",
      rowPageBreak: "avoid",   // una fila nunca se parte entre dos hojas
    };
    const num = (v, d) => fmtNum(v, d);
    const tot = { fontStyle: "bold", fillColor: C.raised, textColor: C.ink, lineWidth: { top: 0.4 }, lineColor: C.ink };

    if (esMovs) {
      // Igual que la pestaña "Movimientos": orden fecha > rubro > concepto; la fecha/labor
      // y el rubro no se repiten si son iguales a la fila anterior; línea al cambiar de fecha/labor.
      const movs = ordenarMovimientosLote(rows);
      const hasGrupo = hasPorGrupoMovs(movs);
      const meta = [];
      const body = movs.map((r, i) => {
        const prev = movs[i - 1];
        const mismaFL = !!prev && prev["Fecha"] === r["Fecha"] && prev["Origen"] === r["Origen"] && prev["OTA"] === r["OTA"];
        const mismoRubro = mismaFL && prev["Tipo item"] === r["Tipo item"];
        const fecha = fmtDate(r["Fecha"]), labor = r["Origen"] || "", ota = otaCorta(r["OTA"]);
        const hg = hasGrupo[i];
        const laborTxt = labor + (hg ? ` · ${fmtNum(hg.has, 2)} ${hg.unidad}` : "");
        meta.push({ inicio: !mismaFL && i > 0, fecha: mismaFL ? "" : fecha, labor: mismaFL ? "" : laborTxt, ota: mismaFL ? "" : ota });
        const dosis = dosisLabor(r, info.ha);
        return [
          mismaFL ? "" : fecha,
          { content: mismoRubro ? "" : (r["Tipo item"] || ""), styles: { textColor: C.soft, fontSize: 6.8 } },
          { content: r["Concepto"] || "", styles: { textColor: C.soft, fontSize: 6.8 } },
          `${num(r["Cantidad"], 1)} ${r["Unid."] || ""}`,
          { content: dosis === null ? "-" : `${num(dosis.valor, 4)} ${r["Unid."] || ""}/${dosis.unidad}`, styles: { textColor: C.soft } },
          { content: r["Total"] === null ? "-" : fmtUSD2(r["Total"]), styles: { fontStyle: "bold" } },
        ];
      });
      body.push([
        { content: "Total", styles: tot }, { content: `${fmtNum(movs.length)} movimientos`, styles: tot }, { content: "", styles: tot },
        { content: "", styles: tot }, { content: "", styles: tot }, { content: fmtUSD2(info.total), styles: { ...tot, halign: "right" } },
      ]);
      const sinUSD = (fila) => (sinPrecios ? fila.slice(0, 5) : fila);
      autoTable(doc, {
        ...base,
        head: [sinUSD(["Fecha / Labor", "Rubro", "Concepto", "Cant.", "Dosis", "USD"])],
        body: body.map((r) => sinUSD(r).map(limpiar)),
        columnStyles: { 0: { cellWidth: 46 }, 1: { cellWidth: 26 }, 3: { halign: "right", cellWidth: 24 }, 4: { halign: "right", cellWidth: 26 }, 5: { halign: "right", cellWidth: 24 } },
        // Alto justo para fecha + labor + OTA (se dibujan a mano, con tamaños distintos)
        didParseCell: (d) => {
          if (d.section !== "body" || d.column.index !== 0) return;
          const m = meta[d.row.index];
          if (!m || !m.fecha) return;
          const w = 46 - 2.4;
          doc.setFont("helvetica", "normal"); doc.setFontSize(6.2);
          const nl = doc.splitTextToSize(pdfTxt(m.labor), w).length;
          d.cell.styles.minCellHeight = 4.4 + nl * 2.6;
        },
        // Fecha en tinta, labor más chica y gris, OTA en dorado, dentro de la misma celda
        willDrawCell: (d) => {
          if (d.section === "body" && d.column.index === 0 && meta[d.row.index]) d.cell.text = [];
        },
        didDrawCell: (d) => {
          if (d.section !== "body") return;
          const m = meta[d.row.index];
          if (!m) return;
          if (m.inicio) {
            doc.setDrawColor(...C.ink); doc.setLineWidth(0.3);
            doc.line(d.cell.x, d.cell.y, d.cell.x + d.cell.width, d.cell.y);
          }
          if (d.column.index === 0 && m.fecha) {
            const px = d.cell.x + 1.2;
            doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(...C.ink);
            doc.text(pdfTxt(m.fecha), px, d.cell.y + 3.6);
            if (m.ota) {   // OTA abreviada en la misma línea que la fecha
              const fx = px + doc.getTextWidth(pdfTxt(m.fecha)) + 2;
              doc.setFontSize(6.2); doc.setTextColor(184, 132, 46);
              doc.text(pdfTxt(m.ota), fx, d.cell.y + 3.6);
            }
            doc.setFontSize(6.2); doc.setTextColor(...C.soft);
            doc.text(doc.splitTextToSize(pdfTxt(m.labor), d.cell.width - 2.4), px, d.cell.y + 6.6);
          }
        },
      });
    } else {
      const VERDE_1 = [219, 226, 214], VERDE_2 = [237, 240, 233];
      const body = [];
      resumen.forEach((dg) => {
        const st = { fontStyle: "bold", fillColor: VERDE_1, textColor: C.ink };
        body.push([
          { content: dg.tipoDet, styles: st }, { content: "", styles: st }, { content: "", styles: st },
          { content: String(dg.items), styles: { ...st, halign: "right" } }, { content: num(dg.cant, 1), styles: { ...st, halign: "right" } },
          { content: fmtUSD2(dg.costo), styles: { ...st, halign: "right" } }, { content: dg.costoHa === null ? "-" : fmtUSD2(dg.costoHa), styles: { ...st, halign: "right" } },
        ]);
        dg.rubros.forEach((rg) => {
          const sr = { fontStyle: "bold", fillColor: VERDE_2, textColor: C.ink };
          body.push([
            { content: rg.rubro, styles: { ...sr, cellPadding: { top: 1.2, bottom: 1.2, left: 4, right: 1.4 } } }, { content: "", styles: sr }, { content: "", styles: sr },
            { content: String(rg.items), styles: { ...sr, halign: "right" } }, { content: num(rg.cant, 1), styles: { ...sr, halign: "right" } },
            { content: fmtUSD2(rg.costo), styles: { ...sr, halign: "right" } }, { content: rg.costoHa === null ? "-" : fmtUSD2(rg.costoHa), styles: { ...sr, halign: "right" } },
          ]);
          rg.conceptos.forEach((cg) => {
            body.push([
              { content: cg.concepto, styles: { textColor: C.soft, cellPadding: { top: 1.2, bottom: 1.2, left: 7, right: 1.4 } } },
              cg.unitPrice === null ? "-" : fmtUSD2(cg.unitPrice),
              cg.dosisHa === null ? "-" : `${fmtDosis(cg.dosisHa)} ${cg.unid || ""}/ha`,
              String(cg.items), `${num(cg.cant, 1)} ${cg.unid || ""}`, fmtUSD2(cg.costo), cg.costoHa === null ? "-" : fmtUSD2(cg.costoHa),
            ]);
          });
        });
      });
      body.push([
        { content: "Total", styles: tot }, { content: "", styles: tot }, { content: "", styles: tot },
        { content: String(rows.length), styles: { ...tot, halign: "right" } }, { content: "", styles: tot },
        { content: fmtUSD2(info.total), styles: { ...tot, halign: "right" } }, { content: fmtUSD2(info.costoHa), styles: { ...tot, halign: "right" } },
      ]);
      // Sin precios: se quitan $/U, Costo y Costo/ha
      const cols = sinPrecios ? [0, 2, 3, 4] : [0, 1, 2, 3, 4, 5, 6];
      autoTable(doc, {
        ...base,
        head: [cols.map((i) => ["Rubro / Concepto", "$/U", "Dosis/ha", "Items", "Cant.", "Costo", "Costo/ha"][i])],
        body: body.map((r) => cols.map((i) => r[i]).map(limpiar)),
        columnStyles: Object.fromEntries(cols.map((_, j) => [j, j === 0 ? {} : { halign: "right" }])),
      });
    }
    const finTabla = doc.lastAutoTable.finalY;
    const mismaPagina = doc.getNumberOfPages() === paginaBloque;
    api.setY((mismaPagina ? Math.max(finChart, finTabla) : finTabla) + 10);
  });
}

// ---------------------------------------------------------------------------
// Dashboard de costos (extraído de ClienteDashboard para cumplir reglas de hooks)
// ---------------------------------------------------------------------------
export default function DashboardCostos({ slug, isAdmin, cloudDb, onLogout, clienteNombre }) {
  const pdfRootRef = useRef(null);
  const storageKey = `${STORAGE_KEY_PREFIX}-${slug}`;
  const cloudPath = `${CLOUD_CLIENTS_BASE}/${slug}/${NODO_COSTOS}`;
  const [rows, setRows] = useState([]);
  const [meta, setMeta] = useState(null); // {fileName, updatedAt, rowCount}
  const [bootLoading, setBootLoading] = useState(true);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [syncStatus, setSyncStatus] = useState(cloudDb ? "connecting" : "local");
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
  // Vista: "labores" (qué se hizo: labores, cantidades y dosis, sin precios) o "costos" (cuánto costó).
  // Al cargar datos se elige sola: Labores si la mayoría de los ítems no tiene precio.
  const [vista, setVista] = useState("costos");
  const L = vista === "labores";
  useEffect(() => {
    if (!rows.length) return;
    const sinPrecio = rows.filter((r) => r["Total"] === null).length;
    setVista(sinPrecio > rows.length / 2 ? "labores" : "costos");
  }, [rows]);

  // ---------- Carga: caché local y luego Firebase en vivo (igual que Cultivos) ----------
  useEffect(() => {
    let unsub = null;
    const aplicar = (payload) => {
      const objRows = rowsFromPayload(payload);
      const cultivosGorrion = leerCultivos(payload.cultivos);
      coloresGorrion = mapaColoresCultivos(cultivosGorrion);
      setRows(objRows);
      setMeta({ fileName: payload.fileName || "", updatedAt: payload.updatedAt, rowCount: objRows.length, encabezado: leerEncabezado(payload.encabezado), cultivos: cultivosGorrion });
    };
    try {
      const raw = localStorage.getItem(storageKey);
      if (raw) aplicar(JSON.parse(raw));
    } catch (e) {}
    if (cloudDb) {
      unsub = onValue(ref(cloudDb, cloudPath), (snap) => {
        const payload = snap.val();
        if (payload && Array.isArray(payload.rows) && Array.isArray(payload.columns)) aplicar(payload);
        else { setRows([]); setMeta(null); }
        setSyncStatus("cloud");
        setBootLoading(false);
      }, (err) => { console.error("Error de sincronización con la nube:", err); setSyncStatus("error"); setBootLoading(false); });
    } else {
      setBootLoading(false);
    }
    return () => { if (unsub) unsub(); };
  }, [slug]); // eslint-disable-line react-hooks/exhaustive-deps

  // Se guarda con las columnas de Gorrión: el mismo formato que publica Gorrión
  const persist = useCallback(async (normalizedRows, fileName, encabezado, cultivos) => {
    const payload = {
      columns: COLS,
      rows: normalizedRows.map((r) => COLS.map((c) => r[c] ?? "")),
      fileName, updatedAt: new Date().toISOString(),
      ...(encabezado ? { encabezado } : {}),
      ...(cultivos ? { cultivos } : {}),
    };
    try { localStorage.setItem(storageKey, JSON.stringify(payload)); } catch (e) {}
    if (cloudDb) {
      try { await dbSet(ref(cloudDb, cloudPath), payload); setSyncStatus("cloud"); }
      catch (e) { setSyncStatus("error"); throw e; }
    }
  }, [slug, cloudDb]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setParsing(true);
    setError("");
    try {
      // Excel (hoja con más filas) o JSON con el mismo formato
      const { filas: json, fileName, encabezado, cultivos } = await leerFilasArchivo(file, { hoja: "mayor" });
      const alineadas = json.map(alinearColumnas);
      validarColumnas(alineadas, REQUERIDAS);
      const normalized = alineadas.map(normalizeRow).filter((r) => r["Campo"] || r["Cultivo"] || r["Concepto"]);
      coloresGorrion = mapaColoresCultivos(cultivos);   // un Excel no trae colores: vuelve a la paleta
      setRows(normalized);
      const newMeta = { fileName, updatedAt: new Date().toISOString(), rowCount: normalized.length, encabezado, cultivos };
      setMeta(newMeta);
      setPage(1);
      try { await persist(normalized, fileName, encabezado, cultivos); }
      catch (e) { setError(`El archivo se leyó, pero no se pudo guardar en la nube (${e.code || e.message}). Sólo lo ves en este navegador.`); }

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
    try { localStorage.removeItem(storageKey); } catch (e) {}
    if (cloudDb) {
      try { await dbSet(ref(cloudDb, cloudPath), null); } catch (e) {}
    }
  };


  // Listas para filtros
  const cultivos = useMemo(() => Array.from(new Set(rows.map((r) => r["Cultivo"]).filter(Boolean))).sort(ordenCultivos), [rows]);
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
    let habGasto = 0; const habSup = new Map();
    filtered.forEach((r) => {
      if (r["Total"] === null) sinPrecio++; else gasto += r["Total"];
      const key = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}`;
      if (!supMap.has(key)) supMap.set(key, r["Sup.Cultivo"] || 0);
      if (r["Lote"]) lotesSet.add(`${r["Campo"]}|${r["Lote"]}`);
      if (r["Campo"]) camposSet.add(r["Campo"]);
      if (r["Cultivo"] && !esHabilitacion(r["Cultivo"])) cultivosSet.add(r["Cultivo"]);
      if (esHabilitacion(r["Cultivo"])) {
        habGasto += r["Total"] || 0;
        const hk = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}`;
        if (!habSup.has(hk)) habSup.set(hk, r["Sup.Cultivo"] || 0);
      }
    });
    const superficie = Array.from(supMap.values()).reduce((a, b) => a + b, 0);
    return {
      gasto, sinPrecio, superficie,
      lotes: lotesSet.size, campos: camposSet.size, cultivos: cultivosSet.size,
      habilitacion: habSup.size ? { gasto: habGasto, ha: Array.from(habSup.values()).reduce((a, b) => a + b, 0) } : null,
      costoPorHa: superficie > 0 ? gasto / superficie : 0,
    };
  }, [filtered]);

  // Vista Labores
  const labKpis = useMemo(() => resumenLabores(filtered), [filtered]);
  const labPorTipo = useMemo(() => haPorLabor(filtered, cultivos).slice(0, 10), [filtered, cultivos]);
  const labCultivos = useMemo(() => cultivos.filter((c) => labPorTipo.some((e) => e[c] > 0)), [cultivos, labPorTipo]);
  const labMes = useMemo(() => haPorMes(filtered), [filtered]);
  const contratistas = useMemo(() => resumenContratistas(filtered), [filtered]);
  const insumosApl = useMemo(() => insumosAplicados(filtered), [filtered]);
  const pctConPrecio = filtered.length ? Math.round(((filtered.length - kpis.sinPrecio) / filtered.length) * 100) : 100;

  // Datos para gráficos: todos normalizados por hectárea y desglosados por cultivo
  // (deduplicando superficie por Campo+Lote+Cultivo para no sumar la misma ha varias veces)
  const porCultivo = useMemo(() => {
    const m = new Map();
    const seenSup = new Set();
    filtered.forEach((r) => {
      if (!r["Cultivo"]) return;
      if (!m.has(r["Cultivo"])) m.set(r["Cultivo"], { gasto: 0, ha: 0, insumos: 0, servicios: 0 });
      const e = m.get(r["Cultivo"]);
      const v = r["Total"] || 0;
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

  // Habilitación va siempre al final y en gris: es una categoría aparte, no un cultivo
  const costoHaPorCultivo = useMemo(() => [...porCultivo].filter((e) => e.ha > 0).sort((a, b) => (esHabilitacion(a.name) - esHabilitacion(b.name)) || b.costoHa - a.costoHa), [porCultivo]);
  const cultivosConHa = useMemo(() => costoHaPorCultivo.map((e) => e.name), [costoHaPorCultivo]);
  const haPorCultivoMap = useMemo(() => new Map(porCultivo.map((e) => [e.name, e.ha])), [porCultivo]);

  // Principales rubros de gasto por hectárea, agrupados por cultivo
  const rubrosPorCultivo = useMemo(() => {
    const m = new Map(); // tipoItem -> { cultivo: gasto }
    filtered.forEach((r) => {
      if (!r["Tipo item"] || !r["Cultivo"]) return;
      if (!m.has(r["Tipo item"])) m.set(r["Tipo item"], {});
      const bucket = m.get(r["Tipo item"]);
      bucket[r["Cultivo"]] = (bucket[r["Cultivo"]] || 0) + (r["Total"] || 0);
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
  // Por defecto: dentro de Insumos / Servicios, por rubro y luego por concepto (A-Z)
  const [itemSort, setItemSort] = useState({ key: "rubro", dir: "asc" });
  const toggleItemSort = (key) => setItemSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: key === "rubro" || key === "concepto" ? "asc" : "desc" }));
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
      e.costo += r["Total"] || 0;
      if (r["Precio/U"] !== null) { e.sumUnit += r["Precio/U"]; e.countUnit += 1; }
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
      // Desempate: rubro y después concepto, alfabético
      return a.rubro.localeCompare(b.rubro, "es", { numeric: true }) || a.concepto.localeCompare(b.concepto, "es", { numeric: true });
    });
    return arr;
  }, [itemsSummary, itemSort]);

  // Exporta a Excel los registros filtrados, con las mismas 21 columnas del archivo original

  const dataQuality = useMemo(() => {
    const totals = { sinPrecio: 0, precioCero: 0, sinSuperficie: 0, sinCultivo: 0, sinCampo: 0, sinConcepto: 0, filasConProblemas: 0 };
    const groups = new Map();
    rows.forEach((r) => {
      const problems = [];
      // En la vista Labores no se controlan los precios
      if (!L && r["Total"] === null) { problems.push("Sin precio"); totals.sinPrecio++; }
      else if (!L && r["Total"] === 0 && r["Cantidad"] > 0) { problems.push("Precio en $0"); totals.precioCero++; }
      if (L && esServicio(r) && !(r[COL_HAS_LABOR] > 0)) { problems.push("Servicio sin hectáreas"); totals.servSinHa = (totals.servSinHa || 0) + 1; }
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
  }, [rows, L]);

  // Comparativo por lotes (agrupa Campo + Lote + Cultivo dentro de lo ya filtrado)
  const loteAgg = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const key = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}`;
      if (!m.has(key)) {
        m.set(key, { campo: r["Campo"], lote: r["Lote"], cultivo: r["Cultivo"], variedad: r["Variedad"], ha: r["Sup.Cultivo"] || 0, insumos: 0, servicios: 0, total: 0, sinPrecio: 0, _rows: [] });
      }
      const e = m.get(key);
      e._rows.push(r);
      if (r["Total"] === null) e.sinPrecio += 1;
      const v = r["Total"] || 0;
      if (r["Tipo Det."] === "INSUMOS") e.insumos += v; else if (r["Tipo Det."] === "SERVICIOS") e.servicios += v;
      e.total += v;
    });
    return Array.from(m.values()).map(({ _rows, ...e }) => {
      const rl = resumenLabores(_rows);
      return { ...e, costoHa: e.ha > 0 ? e.total / e.ha : 0,
        labores: rl.labores, haTrab: rl.ha, insumosN: rl.insumos, ultima: rl.ultima ? rl.ultima.fecha : "", ultimaLabor: rl.ultima ? rl.ultima.labor : "" };
    });
  }, [filtered]);

  const [loteSort, setLoteSort] = useState({ key: "total", dir: "desc" });
  useEffect(() => { setLoteSort(L ? { key: "ultima", dir: "desc" } : { key: "total", dir: "desc" }); }, [L]);
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
    name: `${e.campo} · ${e.lote} · ${e.cultivo}`,
    value: e.costoHa, campo: e.campo, lote: e.lote, cultivo: e.cultivo, variedad: e.variedad, ha: e.ha,
  })), [loteAgg]);

  // Detalle de un lote seleccionado
  const [selectedLoteKey, setSelectedLoteKey] = useState(null);
  const selectLote = (key) => setSelectedLoteKey((cur) => (cur === key ? null : key));
  const selectedLoteInfo = useMemo(() => loteAgg.find((e) => `${e.campo}|${e.lote}|${e.cultivo}|${e.variedad}` === selectedLoteKey) || null, [loteAgg, selectedLoteKey]);
  const selectedLoteRows = useMemo(() => {
    if (!selectedLoteKey) return [];
    return ordenarMovimientosLote(filtered.filter((r) => `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}` === selectedLoteKey));
  }, [filtered, selectedLoteKey]);
  const selectedLoteByItem = useMemo(() => costoPorRubroLote(selectedLoteRows), [selectedLoteRows]);
  const selLab = useMemo(() => resumenLabores(selectedLoteRows), [selectedLoteRows]);
  const selLabChart = useMemo(() => haPorLabor(selectedLoteRows, []).map((e) => ({ name: e.name, value: e._total })), [selectedLoteRows]);

  // Resumen agrupado por Tipo Det. (Insumos primero, Servicios después) > Rubro > Concepto
  const [loteDetailTab, setLoteDetailTab] = useState("resumen");
  // La pestaña elegida (Resumen por rubro / Movimientos) se mantiene al cambiar de lote
  const selectedLoteSummary = useMemo(() => construirResumenLote(selectedLoteRows, selectedLoteInfo ? selectedLoteInfo.ha : 0), [selectedLoteRows, selectedLoteInfo]);
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
    ["Concepto", "Concepto"], ["Cantidad", "Cantidad"], ["Unid.", "Unid."], ["Precio/U", "Precio/U"], ["Total", "Total"],
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

  // ---------- PDF ----------
  const exportarPDF = async (modo, opciones = {}) => {
    const filtros = [];
    if (cultivo !== "Todos") filtros.push(["Cultivo", cultivo]);
    if (administracion !== "Todos") filtros.push(["Admin", administracion]);
    if (campo !== "Todos") filtros.push(["Campo", campo]);
    if (lote !== "Todos") filtros.push(["Lote", loteOptions.find((o) => o.value === lote)?.label || lote]);
    if (tipoDet !== "Todos") filtros.push(["Tipo", tipoDet === "INSUMOS" ? "Solo insumos" : "Solo servicios"]);
    if (tipoItem !== "Todos") filtros.push(["Rubro", tipoItem]);
    if (desde || hasta) filtros.push(["Fechas", `${desde ? fmtDate(desde) : "..."} a ${hasta ? fmtDate(hasta) : "..."}`]);
    if (search.trim()) filtros.push(["Búsqueda", `"${search.trim()}"`]);

    const itemsBody = [];
    itemsSorted.forEach((e, i) => {
      if (e.tipoDet && (i === 0 || itemsSorted[i - 1].tipoDet !== e.tipoDet)) itemsBody.push(filaGrupo(e.tipoDet, 6));
      itemsBody.push([e.rubro, e.concepto, e.precioProm === null ? "-" : fmtUSD2(e.precioProm), `${fmtNum(e.cant, 1)} ${e.unid || ""}`, fmtUSD2(e.costo), String(e.movimientos)]);
    });
    if (L) return exportarPDFLabores(modo, opciones, filtros);
    const tot = loteSorted.reduce((a, e) => ({ ha: a.ha + e.ha, insumos: a.insumos + e.insumos, servicios: a.servicios + e.servicios, total: a.total + e.total }), { ha: 0, insumos: 0, servicios: 0, total: 0 });
    await generarPDF({
      modo, modoLabel: modo === "completo" ? `Informe completo${opciones.detalleLotes ? ` con detalle por lote: ${opciones.tipoDetalle === "movimientos" ? "movimientos" : "resumen por rubro"}${opciones.hojaNueva ? " (un lote por hoja)" : ""}` : ""}` : undefined,
      titulo: "Costos", cliente: clienteNombre || slug, filtros, root: pdfRootRef.current, archivo: "Costos", encabezado: meta?.encabezado,
      tablas: [
        { titulo: "Comparativo por lote", nota: `${fmtNum(loteSorted.length)} lotes · orden actual de la tabla.`,
          head: ["Campo", "Lote", "Cultivo", "Variedad", "Ha", "Insumos", "Servicios", "Total", "USD/ha"],
          align: [null, null, null, null, "right", "right", "right", "right", "right"],
          body: [
            ...loteSorted.map((e) => [e.campo, e.lote, e.cultivo, e.variedad, fmtNum(e.ha, 1), fmtUSD(e.insumos), fmtUSD(e.servicios), fmtUSD(e.total), fmtUSD2(e.costoHa)]),
            filaTotal(["Total", "", "", "", fmtNum(tot.ha, 1), fmtUSD(tot.insumos), fmtUSD(tot.servicios), fmtUSD(tot.total), tot.ha > 0 ? fmtUSD2(tot.total / tot.ha) : "-"]),
          ] },
        { titulo: "Precios y consumo por ítem", nota: `${fmtNum(itemsSorted.length)} ítems.`,
          head: ["Rubro", "Concepto", "Precio prom.", "Cantidad total", "Costo total", "Movs."],
          align: [null, null, "right", "right", "right", "right"], body: itemsBody },
        ...(opciones.detalleLotes ? [{
          render: (api) => {
            const porLote = new Map();
            filtered.forEach((r) => {
              const k = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}`;
              if (!porLote.has(k)) porLote.set(k, []);
              porLote.get(k).push(r);
            });
            const lotes = loteSorted.map((e) => ({ info: e, rows: porLote.get(`${e.campo}|${e.lote}|${e.cultivo}|${e.variedad}`) || [] }));
            return dibujarResumenLotes(api, lotes, !!opciones.hojaNueva, opciones.tipoDetalle || "resumen");
          },
        }] : []),
      ],
    });
  };

  // PDF de la vista Labores: sin precios; el informe completo es el cuaderno de campo
  const exportarPDFLabores = async (modo, opciones, filtros) => {
    const totL = loteSorted.reduce((a, e) => ({ ha: a.ha + e.ha, labores: a.labores + e.labores, haTrab: a.haTrab + e.haTrab }), { ha: 0, labores: 0, haTrab: 0 });
    const insBody = [];
    insumosApl.forEach((e, i) => {
      if (i === 0 || insumosApl[i - 1].rubro !== e.rubro) insBody.push(filaGrupo(e.rubro, 6));
      insBody.push([e.concepto, `${fmtNum(e.cant, 1)} ${e.unid}`, fmtNum(e.ha, 1), e.dosis === null ? "-" : `${fmtDosis(e.dosis)} ${e.unid}/ha`,
        e.min === null ? "-" : `${fmtDosis(e.min)} - ${fmtDosis(e.max)}`, String(e.lotes)]);
    });
    await generarPDF({
      modo, modoLabel: modo === "completo" ? `Cuaderno de campo${opciones.detalleLotes ? ` con detalle por lote: ${(opciones.tipoDetalle || "movimientos") === "movimientos" ? "labores" : "resumen por rubro"}${opciones.hojaNueva ? " (un lote por hoja)" : ""}` : ""}` : undefined,
      titulo: "Labores", cliente: clienteNombre || slug, filtros, root: pdfRootRef.current, archivo: "Labores", encabezado: meta?.encabezado,
      tablas: [
        { titulo: "Comparativo por lote", nota: `${fmtNum(loteSorted.length)} lotes · orden actual de la tabla.`,
          head: ["Campo", "Lote", "Cultivo", "Variedad", "Ha", "Labores", "Ha trabajadas", "Insumos", "Última labor"],
          align: [null, null, null, null, "right", "right", "right", "right", null],
          body: [
            ...loteSorted.map((e) => [e.campo, e.lote, e.cultivo, e.variedad, fmtNum(e.ha, 1), fmtNum(e.labores), fmtNum(e.haTrab, 1), fmtNum(e.insumosN), e.ultima ? `${fmtDate(e.ultima)} ${e.ultimaLabor}` : "-"]),
            filaTotal(["Total", "", "", "", fmtNum(totL.ha, 1), fmtNum(labKpis.labores), fmtNum(totL.haTrab, 1), fmtNum(labKpis.insumos), ""]),
          ] },
        { titulo: "Insumos aplicados", nota: `${fmtNum(insumosApl.length)} productos · dosis = cantidad / hectáreas aplicadas.`,
          head: ["Producto", "Cantidad total", "Ha aplicadas", "Dosis prom./ha", "Dosis mín. - máx.", "Lotes"],
          align: [null, "right", "right", "right", "right", "right"], body: insBody },
        { titulo: "Resumen por contratista", nota: `${fmtNum(contratistas.length)} contratistas · incluye labores propias.`,
          head: ["Contratista", "Labores", "Hectáreas", "Lotes", "Tipos de labor", "Última"],
          align: [null, "right", "right", "right", null, null],
          body: contratistas.map((e) => [e.contratista, fmtNum(e.labores), `${fmtNum(e.ha, 1)} ha${textoOtras(e.otras) ? " · " + textoOtras(e.otras) : ""}`, fmtNum(e.lotes), e.tiposLabor, e.ultima ? fmtDate(e.ultima) : "-"]) },
        ...(opciones.detalleLotes ? [{
          render: (api) => {
            const porLote = new Map();
            filtered.forEach((r) => {
              const k = `${r["Campo"]}|${r["Lote"]}|${r["Cultivo"]}|${r["Variedad"]}`;
              if (!porLote.has(k)) porLote.set(k, []);
              porLote.get(k).push(r);
            });
            const lotes = loteSorted.map((e) => ({ info: e, rows: porLote.get(`${e.campo}|${e.lote}|${e.cultivo}|${e.variedad}`) || [] }));
            return dibujarResumenLotes(api, lotes, !!opciones.hojaNueva, opciones.tipoDetalle || "movimientos", true);
          },
        }] : []),
      ],
    });
  };

  const resetFiltros = () => { setCultivo("Todos"); setCampo("Todos"); setLote("Todos"); setAdministracion("Todos"); setTipoDet("Todos"); setTipoItem("Todos"); setDesde(""); setHasta(""); setSearch(""); };
  const hayFiltrosActivos = cultivo !== "Todos" || campo !== "Todos" || lote !== "Todos" || administracion !== "Todos" || tipoDet !== "Todos" || tipoItem !== "Todos" || desde || hasta || search;

  // -------------------------------------------------------------------------
  return (
    <div ref={pdfRootRef}>
      {/* Sub-header de costos: datos y controles de carga */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
        <div className="agri-sub">
          {meta ? (
            <>Datos de <strong>{meta.fileName}</strong> · {fmtNum(meta.rowCount)} registros · actualizado {new Date(meta.updatedAt).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" })}</>
          ) : isAdmin ? "Cargá tu planilla para empezar a consultar los datos" : "Todavía no hay datos cargados para este cliente."}
        </div>
        {cloudDb && (
          <div style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 600, padding: "3px 9px", borderRadius: 999, background: syncStatus === "cloud" ? "rgba(75,107,58,0.12)" : syncStatus === "error" ? "rgba(161,70,47,0.12)" : "rgba(107,94,79,0.12)", color: syncStatus === "cloud" ? "var(--green)" : syncStatus === "error" ? "var(--rust)" : "var(--ink-soft)" }}>
            {syncStatus === "cloud" ? <><Cloud size={12} /> Sincronizado</> : syncStatus === "error" ? <><CloudOff size={12} /> Sin conexión</> : <><Cloud size={12} /> Conectando…</>}
          </div>
        )}
        <div style={{ display: "flex", gap: 8 }}>
          {meta && rows.length > 0 && <PdfMenu key={vista} onExport={exportarPDF} modos={L ? MODOS_PDF_LABORES : MODOS_PDF_COSTOS} />}
          {isAdmin && (<>
            <input ref={fileInputRef} type="file" accept={ACCEPT_ARCHIVOS} style={{ display: "none" }}
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
            {onLogout && (
              <button className="agri-btn agri-btn-outline" onClick={onLogout} title="Cerrar sesión">
                <LogOut size={14} />
              </button>
            )}
          </>)}
        </div>
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
              {parsing ? "Leyendo el archivo…" : "Arrastrá tu Excel o JSON acá"}
            </div>
            <div style={{ fontSize: 13, color: "var(--ink-soft)", maxWidth: 360 }}>
              Excel (.xlsx) o JSON con las columnas de campo, lote, cultivo, insumos y costos.
            </div>
            <input ref={fileInputRef} type="file" accept={ACCEPT_ARCHIVOS} style={{ display: "none" }}
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
              <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 12 }}>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
                  <button className="agri-chip" data-active={cultivo === "Todos"} onClick={() => setCultivo("Todos")}>Todos los cultivos</button>
                  {cultivos.map((c) => (
                    <button key={c} className="agri-chip" data-active={cultivo === c} onClick={() => setCultivo(c)}>{c}</button>
                  ))}
                </div>
                {/* Vista: Labores | Costos (mismo selector que Siembra | Cosecha en Cultivos) */}
                <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 4 }}>
                  <div style={{ display: "inline-flex", border: "1px solid var(--ink)", borderRadius: 6, overflow: "hidden" }}>
                    {[["labores", "Labores"], ["costos", "Costos"]].map(([k, l]) => (
                      <button key={k} onClick={() => setVista(k)}
                        style={{ border: "none", padding: "6px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer", font: "inherit", background: vista === k ? "var(--ink)" : "transparent", color: vista === k ? "var(--paper)" : "var(--ink)" }}>
                        {l}
                      </button>
                    ))}
                  </div>
                  {!L && pctConPrecio < 100 && (
                    <span style={{ fontSize: 11, color: pctConPrecio < 80 ? "var(--rust)" : "var(--ink-soft)" }}>{pctConPrecio} % de los ítems con precio cargado</span>
                  )}
                </div>
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
            <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 20 }}>
              {L ? (<>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Labores realizadas</div>
                <div className="agri-kpi-value">{fmtNum(labKpis.labores)}</div>
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label"><Ruler size={12} style={{ display: "inline", marginBottom: -1 }} /> Hectáreas trabajadas</div>
                <div className="agri-kpi-value">{fmtNum(labKpis.ha)} ha</div>
                {textoOtras(labKpis.otras) && <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>además {textoOtras(labKpis.otras)}</div>}
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Superficie</div>
                <div className="agri-kpi-value">{fmtNum(kpis.superficie)} ha</div>
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Insumos aplicados</div>
                <div className="agri-kpi-value">{fmtNum(labKpis.insumos)}</div>
                <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>productos distintos</div>
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Contratistas</div>
                <div className="agri-kpi-value">{fmtNum(labKpis.contratistas)}</div>
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label">Última labor</div>
                <div className="agri-kpi-value" style={{ fontSize: 16 }}>{labKpis.ultima ? fmtDate(labKpis.ultima.fecha) : "—"}</div>
                {labKpis.ultima && <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>{labKpis.ultima.labor}</div>}
              </div>
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label"><MapPin size={12} style={{ display: "inline", marginBottom: -1 }} /> Campos · Lotes</div>
                <div className="agri-kpi-value">{kpis.campos} · {kpis.lotes}</div>
              </div>
              </>) : (<>
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
              {kpis.habilitacion && (
                <div className="agri-card" style={{ padding: 14 }}>
                  <div className="agri-kpi-label">Habilitación</div>
                  <div className="agri-kpi-value">{fmtUSD(kpis.habilitacion.gasto)}</div>
                  <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>
                    {fmtNum(kpis.habilitacion.ha)} ha{kpis.habilitacion.ha > 0 ? ` · ${fmtUSD2(kpis.habilitacion.gasto / kpis.habilitacion.ha)}/ha` : ""}
                  </div>
                </div>
              )}
              <div className="agri-card" style={{ padding: 14 }}>
                <div className="agri-kpi-label"><MapPin size={12} style={{ display: "inline", marginBottom: -1 }} /> Campos · Lotes</div>
                <div className="agri-kpi-value">{kpis.campos} · {kpis.lotes}</div>
              </div>
              </>)}
            </div>

            {L ? (<>
            {/* Gráficos de labores: hectáreas trabajadas */}
            <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 14, marginBottom: 20 }}>
              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Hectáreas trabajadas por labor</div>
                <ResponsiveContainer width="100%" height={Math.max(220, labPorTipo.length * 30 + 40)}>
                  <BarChart data={labPorTipo} layout="vertical" margin={{ left: 10, right: 30 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 10, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v, n) => [`${fmtNum(v, 1)} ha`, n]} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    {labCultivos.map((c, i) => (
                      <Bar key={c} dataKey={c} name={c} stackId="ha" fill={cultivoColor(c, i)} maxBarSize={26} />
                    ))}
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Hectáreas trabajadas por mes</div>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={labMes} margin={{ left: -10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" vertical={false} />
                    <XAxis dataKey="name" tick={{ fontSize: 11, fill: "#6B5E4F" }} axisLine={{ stroke: "#DCD2B8" }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v) => `${fmtNum(v, 1)} ha`} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Bar dataKey="ha" name="Hectáreas" fill="#4B6B3A" radius={[3, 3, 0, 0]} maxBarSize={48}>
                      <LabelList dataKey="ha" content={<VerticalBarLabel formato={fmtHa} />} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="agri-card" style={{ padding: 16 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Hectáreas por contratista</div>
                <ResponsiveContainer width="100%" height={Math.max(220, Math.min(10, contratistas.length) * 30 + 30)}>
                  <BarChart data={contratistas.slice(0, 10)} layout="vertical" margin={{ left: 10, right: 30 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                    <YAxis type="category" dataKey="contratista" width={140} tick={{ fontSize: 10, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                    <Tooltip formatter={(v) => `${fmtNum(v, 1)} ha`} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                    <Bar dataKey="ha" name="Hectáreas" fill="#B8842E" radius={[0, 3, 3, 0]} maxBarSize={26}>
                      <LabelList dataKey="ha" content={<HorizontalBarLabel formato={fmtHa} />} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            </>) : (<>
            {/* Gráficos — todos normalizados por hectárea y desglosados por cultivo */}
            <div data-pdf="resumen" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 14, marginBottom: 20 }}>
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

            </>)}

            {!L && (<>
            {/* Comparativo por lotes */}
            <div data-pdf="resumen" className="agri-card" style={{ padding: 16, marginBottom: 14 }}>
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

            </>)}

            <div data-pdf="pantalla" className="agri-card" style={{ padding: 0, overflow: "hidden", marginBottom: 20 }}>
              <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Comparativo por lote</div>
                <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(loteSorted.length)} lotes · tocá una fila para ver el detalle</div>
              </div>
              <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 780 }}>
                  <thead>
                    <tr>
                      {(L
                        ? [["campo", "Campo"], ["lote", "Lote"], ["cultivo", "Cultivo"], ["variedad", "Variedad"], ["ha", "Ha"], ["labores", "Labores"], ["haTrab", "Ha trabajadas"], ["insumosN", "Insumos"], ["ultima", "Última labor"]]
                        : [["campo", "Campo"], ["lote", "Lote"], ["cultivo", "Cultivo"], ["variedad", "Variedad"], ["ha", "Ha"], ["insumos", "Insumos"], ["servicios", "Servicios"], ["total", "Total"], ["costoHa", "USD/ha"]]
                      ).map(([key, label]) => (
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
                        {L ? (<>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(e.labores)}</td>
                        <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{fmtNum(e.haTrab, 1)}</td>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(e.insumosN)}</td>
                        <td className="agri-td" style={{ whiteSpace: "nowrap" }}>
                          {e.ultima ? fmtDate(e.ultima) : "—"}
                          {e.ultimaLabor && <span style={{ fontSize: 11, color: "var(--ink-soft)", marginLeft: 6 }}>{e.ultimaLabor}</span>}
                        </td>
                        </>) : (<>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtUSD(e.insumos)}</td>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtUSD(e.servicios)}</td>
                        <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{fmtUSD(e.total)}</td>
                        <td className="agri-td">
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <div style={{ flex: "0 0 58px", textAlign: "right", fontWeight: 600 }}>{fmtUSD2(e.costoHa)}</div>
                            <div style={{ flex: 1, minWidth: 40, height: 6, background: "var(--line)", borderRadius: 3, overflow: "hidden" }}>
                              <div style={{ width: `${Math.min(100, (e.costoHa / maxCostoHa) * 100)}%`, height: "100%", background: "var(--gold)" }} />
                            </div>
                            {e.sinPrecio > 0 && <span title={`${e.sinPrecio} ítems sin precio: el costo está incompleto`} style={{ color: "var(--rust)", fontWeight: 700 }}>*</span>}
                          </div>
                        </td>
                        </>)}
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
              <div data-pdf="pantalla" className="agri-card" style={{ padding: 16, marginBottom: 20, borderColor: "var(--gold)" }}>
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
                  {L ? (<>
                  <div>
                    <div className="agri-kpi-label">Labores</div>
                    <div className="agri-kpi-value" style={{ fontSize: 15 }}>{fmtNum(selLab.labores)}</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Hectáreas trabajadas</div>
                    <div className="agri-kpi-value" style={{ fontSize: 15 }}>{fmtNum(selLab.ha, 1)} ha</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Insumos aplicados</div>
                    <div className="agri-kpi-value" style={{ fontSize: 15 }}>{fmtNum(selLab.insumos)}</div>
                  </div>
                  <div>
                    <div className="agri-kpi-label">Última labor</div>
                    <div className="agri-kpi-value" style={{ fontSize: 15 }}>{selLab.ultima ? fmtDate(selLab.ultima.fecha) : "—"}</div>
                    {selLab.ultima && <div style={{ fontSize: 11, color: "var(--ink-soft)" }}>{selLab.ultima.labor}</div>}
                  </div>
                  </>) : (<>
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
                  </>)}
                </div>

                <div className="agri-detail-grid" style={{ display: "grid", gridTemplateColumns: "minmax(190px, 0.8fr) minmax(320px, 1.8fr)", gap: 16 }}>
                  <div>
                    {L ? (<>
                    <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Hectáreas trabajadas por labor</div>
                    <ResponsiveContainer width="100%" height={Math.max(180, selLabChart.length * 26 + 30)}>
                      <BarChart data={selLabChart} layout="vertical" margin={{ left: 10, right: 34 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#DCD2B8" horizontal={false} />
                        <XAxis type="number" tick={{ fontSize: 10, fill: "#6B5E4F" }} tickFormatter={(v) => fmtNum(v)} axisLine={false} tickLine={false} />
                        <YAxis type="category" dataKey="name" width={100} tick={{ fontSize: 10, fill: "#2B2118" }} axisLine={false} tickLine={false} />
                        <Tooltip formatter={(v) => `${fmtNum(v, 1)} ha`} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #DCD2B8" }} />
                        <Bar dataKey="value" name="Hectáreas" fill="#B8842E" radius={[0, 3, 3, 0]} maxBarSize={22}>
                          <LabelList dataKey="value" content={<HorizontalBarLabel formato={fmtHa} />} />
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                    </>) : (<>
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
                    </>)}
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
                            {["Fecha / Labor", "Rubro", "Concepto", "Cant.", "Dosis", ...(L ? [] : ["USD"])].map((h) => (
                              <th key={h} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)", fontSize: 9, padding: "5px 7px", width: h === "Fecha / Labor" ? 210 : undefined }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {(() => { const hasGrupo = hasPorGrupoMovs(selectedLoteRows); return selectedLoteRows.map((r, i) => {
                            const prev = selectedLoteRows[i - 1];
                            const sameFechaLabor = prev && prev["Fecha"] === r["Fecha"] && prev["Origen"] === r["Origen"] && prev["OTA"] === r["OTA"];
                            const sameRubro = sameFechaLabor && prev["Tipo item"] === r["Tipo item"];
                            const tdSm = { padding: "4px 7px", fontSize: 11 };
                            const tdLabor = { padding: "4px 7px", fontSize: 9, color: "var(--ink-soft)", fontWeight: 400 };
                            const dosis = dosisLabor(r, selectedLoteInfo.ha);
                            const hg = hasGrupo[i];
                            return (
                              <tr key={i} className="agri-tr" style={{ borderTop: !sameFechaLabor && i > 0 ? "1px solid var(--ink)" : undefined }}>
                                <td className="agri-td" style={{ ...tdSm, color: sameFechaLabor ? "var(--line)" : "var(--ink)" }}>
                                  {sameFechaLabor ? "″" : (
                                    <>
                                      <div style={{ display: "flex", alignItems: "baseline", gap: 6, whiteSpace: "nowrap" }}>
                                        <span>{fmtDate(r["Fecha"])}</span>
                                        {r["OTA"] && <span style={{ fontSize: 9, color: "var(--gold)", fontWeight: 400 }} title={r["OTA"]}>{otaCorta(r["OTA"])}</span>}
                                      </div>
                                      <div style={{ display: "flex", gap: 4, fontSize: 9, color: "var(--ink-soft)", fontWeight: 400, whiteSpace: "nowrap", maxWidth: 200 }}>
                                        <span style={{ overflow: "hidden", textOverflow: "ellipsis", minWidth: 0 }} title={r["Origen"]}>{r["Origen"]}</span>
                                        {hg && <span style={{ flexShrink: 0, color: "var(--ink)", fontWeight: 600 }}>· {fmtNum(hg.has, 2)} {hg.unidad}</span>}
                                      </div>
                                    </>
                                  )}
                                </td>
                                <td className="agri-td" style={{ ...tdLabor, color: sameRubro ? "var(--line)" : "var(--ink-soft)" }}>{sameRubro ? "″" : r["Tipo item"]}</td>
                                <td className="agri-td" style={tdLabor}>{r["Concepto"]}</td>
                                <td className="agri-td" style={{ ...tdSm, textAlign: "right", whiteSpace: "nowrap" }}>{fmtNum(r["Cantidad"], 1)} {r["Unid."]}</td>
                                <td className="agri-td" style={{ ...tdSm, textAlign: "right", color: "var(--ink-soft)", whiteSpace: "nowrap" }}>{dosis === null ? "—" : `${fmtDosis(dosis.valor)} ${r["Unid."]}/${dosis.unidad}`}</td>
                                {!L && <td className="agri-td" style={{ ...tdSm, textAlign: "right", fontWeight: 600 }}>{r["Total"] === null ? "—" : fmtUSD2(r["Total"])}</td>}
                              </tr>
                            );
                          }); })()}
                        </tbody>
                      </table>
                    </div>
                    ) : (
                    <div style={{ overflowX: "auto", maxHeight: 320, overflowY: "auto", border: "1px solid var(--line)", borderRadius: 6 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 620 }}>
                        <thead>
                          <tr>
                            {(L ? ["Rubro / Concepto", "Dosis/ha", "Items", "Cant."] : ["Rubro / Concepto", "$/U", "Dosis/ha", "Items", "Cant.", "Costo", "Costo/ha"]).map((h) => (
                              <th key={h} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)", fontSize: 9, padding: "5px 7px", textAlign: h === "Rubro / Concepto" ? "left" : "right" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {selectedLoteSummary.map((dg, di) => (
                            <React.Fragment key={di}>
                              <tr className="agri-tr" style={{ background: "rgba(75,107,58,0.16)" }}>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.02em" }}>{dg.tipoDet}</td>
                                {!L && <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11 }}></td>}
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11 }}></td>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{dg.items}</td>
                                <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{fmtNum(dg.cant, 1)}</td>
                                {!L && <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{fmtUSD2(dg.costo)}</td>}
                                {!L && <td className="agri-td" style={{ padding: "5px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{dg.costoHa === null ? "—" : fmtUSD2(dg.costoHa)}</td>}
                              </tr>
                              {dg.rubros.map((rg, ri) => (
                                <React.Fragment key={ri}>
                                  <tr className="agri-tr" style={{ background: "rgba(75,107,58,0.07)" }}>
                                    <td className="agri-td" style={{ padding: "4px 7px 4px 14px", fontSize: 11, fontWeight: 700 }}>{rg.rubro}</td>
                                    {!L && <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11 }}></td>}
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11 }}></td>
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{rg.items}</td>
                                    <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{fmtNum(rg.cant, 1)}</td>
                                    {!L && <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{fmtUSD2(rg.costo)}</td>}
                                    {!L && <td className="agri-td" style={{ padding: "4px 7px", fontSize: 11, textAlign: "right", fontWeight: 700 }}>{rg.costoHa === null ? "—" : fmtUSD2(rg.costoHa)}</td>}
                                  </tr>
                                  {rg.conceptos.map((cg, ci) => (
                                    <tr key={ci} className="agri-tr">
                                      <td className="agri-td" style={{ padding: "4px 7px 4px 26px", fontSize: 10, color: "var(--ink-soft)" }}>{cg.concepto}</td>
                                      {!L && <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{cg.unitPrice === null ? "—" : fmtUSD2(cg.unitPrice)}</td>}
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{cg.dosisHa === null ? "—" : `${fmtDosis(cg.dosisHa)} ${cg.unid}/ha`}</td>
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{cg.items}</td>
                                      <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{fmtNum(cg.cant, 1)} {cg.unid}</td>
                                      {!L && <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{fmtUSD2(cg.costo)}</td>}
                                      {!L && <td className="agri-td" style={{ padding: "4px 7px", fontSize: 10, textAlign: "right" }}>{cg.costoHa === null ? "—" : fmtUSD2(cg.costoHa)}</td>}
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

            {L ? (<>
            {/* Insumos aplicados (sin precios) */}
            <div data-pdf="pantalla" className="agri-card" style={{ padding: 0, overflow: "hidden", marginBottom: 20 }}>
              <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Insumos aplicados</div>
                <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(insumosApl.length)} productos · dosis = cantidad / hectáreas aplicadas</div>
              </div>
              <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
                  <thead>
                    <tr>
                      {["Rubro", "Producto", "Cantidad total", "Ha aplicadas", "Dosis prom./ha", "Dosis mín. – máx.", "Lotes"].map((h, k) => (
                        <th key={h} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)", textAlign: k >= 2 ? "right" : "left" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {insumosApl.map((e, i) => {
                      const prev = insumosApl[i - 1];
                      const nuevoRubro = !prev || prev.rubro !== e.rubro;
                      const dispar = e.min !== null && e.max !== null && e.min > 0 && e.max / e.min >= 3;
                      return (
                        <tr key={i} className="agri-tr" style={{ borderTop: nuevoRubro && i > 0 ? "1px solid var(--ink)" : undefined }}>
                          <td className="agri-td" style={{ fontWeight: nuevoRubro ? 600 : 400, color: nuevoRubro ? "var(--ink)" : "var(--line)" }}>{nuevoRubro ? e.rubro : "″"}</td>
                          <td className="agri-td">{e.concepto}</td>
                          <td className="agri-td" style={{ textAlign: "right", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtNum(e.cant, 1)} {e.unid}</td>
                          <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(e.ha, 1)}</td>
                          <td className="agri-td" style={{ textAlign: "right", whiteSpace: "nowrap" }}>{e.dosis === null ? "—" : `${fmtDosis(e.dosis)} ${e.unid}/ha`}</td>
                          <td className="agri-td" style={{ textAlign: "right", whiteSpace: "nowrap", color: dispar ? "var(--rust)" : "var(--ink-soft)" }}
                            title={dispar ? "La dosis máxima es 3 veces o más la mínima: revisar la carga" : undefined}>
                            {e.min === null ? "—" : `${fmtDosis(e.min)} – ${fmtDosis(e.max)}`}
                          </td>
                          <td className="agri-td" style={{ textAlign: "right", color: "var(--ink-soft)" }}>{e.lotes}</td>
                        </tr>
                      );
                    })}
                    {insumosApl.length === 0 && (
                      <tr><td className="agri-td" colSpan={7} style={{ textAlign: "center", padding: 30, color: "var(--ink-soft)" }}>No hay insumos aplicados con los filtros de arriba.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Contratistas */}
            <div data-pdf="pantalla" className="agri-card" style={{ padding: 0, overflow: "hidden", marginBottom: 20 }}>
              <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Resumen por contratista</div>
                <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(contratistas.length)} contratistas · incluye labores propias</div>
              </div>
              <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
                  <thead>
                    <tr>
                      {["Contratista", "Labores", "Hectáreas", "Lotes", "Tipos de labor", "Última"].map((h, k) => (
                        <th key={h} className="agri-th" style={{ position: "sticky", top: 0, background: "var(--paper-raised)", textAlign: k >= 1 && k <= 3 ? "right" : "left" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {contratistas.map((e, i) => (
                      <tr key={i} className="agri-tr">
                        <td className="agri-td" style={{ fontWeight: 600 }}>{e.contratista}</td>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(e.labores)}</td>
                        <td className="agri-td" style={{ textAlign: "right", fontWeight: 600, whiteSpace: "nowrap" }}>
                          {fmtNum(e.ha, 1)} ha{textoOtras(e.otras) && <span style={{ fontWeight: 400, color: "var(--ink-soft)" }}> · {textoOtras(e.otras)}</span>}
                        </td>
                        <td className="agri-td" style={{ textAlign: "right" }}>{fmtNum(e.lotes)}</td>
                        <td className="agri-td" style={{ fontSize: 12, color: "var(--ink-soft)" }}>{e.tiposLabor}</td>
                        <td className="agri-td" style={{ whiteSpace: "nowrap" }}>{e.ultima ? fmtDate(e.ultima) : "—"}</td>
                      </tr>
                    ))}
                    {contratistas.length === 0 && (
                      <tr><td className="agri-td" colSpan={6} style={{ textAlign: "center", padding: 30, color: "var(--ink-soft)" }}>No hay servicios con los filtros de arriba.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            </>) : (<>
            {/* Precios y consumo por ítem */}
            <div data-pdf="pantalla" className="agri-card" style={{ padding: 0, overflow: "hidden", marginBottom: 20 }}>
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

            </>)}

            {/* Tabla */}
            <div data-pdf="pantalla" className="agri-card" style={{ padding: 0, overflow: "hidden" }}>
              <div style={{ padding: "12px 16px", borderBottom: showDetalle ? "1px solid var(--line)" : "none", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Detalle de movimientos</div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>{fmtNum(sorted.length)} registros filtrados</span>
                  <span data-pdf-ignore style={{ display: "contents" }}>
                  <button className="agri-btn agri-btn-outline" onClick={exportToExcel}><Download size={13} /> Exportar a Excel</button>
                  <button className="agri-btn agri-btn-outline" onClick={() => setShowDetalle((v) => !v)} title={showDetalle ? "Colapsar" : "Expandir"}>
                    {showDetalle ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                  </button>
                  </span>
                </div>
              </div>
              {showDetalle && (
              <>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 780 }}>
                  <thead>
                    <tr>
                      {[["Fecha", "Fecha"], ["Campo", "Campo"], ["Lote", "Lote"], ["Cultivo", "Cultivo"], ["Origen", "Labor"], ["Tipo item", "Rubro"], ["Concepto", "Concepto"], ["Cantidad", "Cant."], ["Unid.", "Un."], ...(L ? [] : [["Total", "USD"]])].map(([key, label]) => (
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
                        {!L && <td className="agri-td" style={{ textAlign: "right", fontWeight: 600 }}>{r["Total"] === null ? "—" : fmtUSD2(r["Total"])}</td>}
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
                      ["Servicio sin hectáreas", dataQuality.totals.servSinHa || 0],
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
    </div>
  );
}
