import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, onValue, get, set as dbSet, query, orderByChild, equalTo } from "firebase/database";
import { getAuth, signInWithEmailAndPassword, onAuthStateChanged, signOut } from "firebase/auth";
import { firebaseConfig, CLOUD_SYNC_ENABLED, CLOUD_CLIENTS_BASE } from "./firebaseConfig.js";
import DashboardProduccion from "./DashboardProduccion.jsx";
import DashboardCultivos from "./DashboardCultivos.jsx";
import DashboardCostos from "./DashboardCostos.jsx";
import { PdfMenu, generarPDF, filaGrupo, filaTotal, MODOS_PDF } from "./pdfExport.jsx";
import { leerFilasArchivo, validarColumnas, leerEncabezado, ACCEPT_ARCHIVOS } from "./lectorArchivo.js";
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
// Componente principal
// ---------------------------------------------------------------------------
function ClienteDashboard({ slug, user, rol }) {
  // Puede actualizar datos el administrador o el dueño del perfil (las reglas de Firebase lo exigen igual)
  const [esDueno, setEsDueno] = useState(false);
  useEffect(() => {
    setEsDueno(false);
    if (!user || rol !== "publicador" || !cloudDb) return;
    get(ref(cloudDb, `${CLOUD_CLIENTS_BASE}/index/${slug}/owner`))
      .then((s) => setEsDueno(s.val() === user.uid)).catch(() => setEsDueno(false));
  }, [slug, user, rol]);
  const isAdmin = rol === "admin" || esDueno;

  const DASHBOARDS = [
    { key: "cultivos",   label: "Cultivos" },
    { key: "costos",     label: "Costos Agrícolas" },
    { key: "produccion", label: "Producción" },
  ];

  const [dashKey, setDashKey] = useState(() => {
    if (typeof window === "undefined") return "cultivos";
    return new URLSearchParams(window.location.search).get("dashboard") || "cultivos";
  });

  useEffect(() => {
    const onPop = () => {
      const k = new URLSearchParams(window.location.search).get("dashboard") || "cultivos";
      setDashKey(k);
    };
    window.addEventListener("popstate", onPop);
    return () => window.removeEventListener("popstate", onPop);
  }, []);

  const navTo = (key) => {
    const u = new URLSearchParams(window.location.search);
    u.set("dashboard", key);
    window.history.replaceState({}, "", `?${u.toString()}`);
    window.dispatchEvent(new Event("popstate"));
  };

  // Encabezado opcional que manda Gorrión con cada dashboard (p.ej. la campaña): se lee sólo
  // clientes/<slug>/<nodo>/encabezado del dashboard activo, sin bajar los datos.
  const NODO_DASH = { costos: "costos", produccion: "produccion", cultivos: "cultivos" };
  const [encabezado, setEncabezado] = useState(null);
  useEffect(() => {
    setEncabezado(null);
    if (!CLOUD_SYNC_ENABLED || !cloudDb) return;
    const nodo = NODO_DASH[dashKey] || "costos";
    const unsub = onValue(ref(cloudDb, `${CLOUD_CLIENTS_BASE}/${slug}/${nodo}/encabezado`),
      (snap) => setEncabezado(leerEncabezado(snap.val())), () => setEncabezado(null));
    return unsub;
  }, [slug, dashKey]); // eslint-disable-line react-hooks/exhaustive-deps

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
        {user && (rol === "admin" || rol === "publicador") && (
          <a href={urlBase()} onClick={(e) => { e.preventDefault(); irA({}); }}
            style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 12, color: "var(--ink-soft)", textDecoration: "none", marginBottom: 6 }}>
            <ArrowLeft size={12} /> {rol === "admin" ? "Volver al panel" : "Volver a mis perfiles"}
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
      {encabezado && (
        <div style={{ textAlign: "right", maxWidth: 560, marginLeft: "auto" }}>
          {encabezado.titulo && <div className="agri-serif" style={{ fontSize: 20, fontWeight: 600, lineHeight: 1.2 }}>{encabezado.titulo}</div>}
          {encabezado.detalle && <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 4, whiteSpace: "pre-line" }}>{encabezado.detalle}</div>}
        </div>
      )}
    </div>
  );

  return (
    <div className="agri-shell">
      {masterhead}
      {dashKey === "produccion"
        ? <DashboardProduccion slug={slug} isAdmin={isAdmin} cloudDb={cloudDb} clienteNombre={clienteNombre} onLogout={cloudAuth ? () => signOut(cloudAuth) : null} />
        : dashKey === "cultivos"
        ? <DashboardCultivos slug={slug} isAdmin={isAdmin} cloudDb={cloudDb} clienteNombre={clienteNombre} onLogout={cloudAuth ? () => signOut(cloudAuth) : null} />
        : <DashboardCostos slug={slug} isAdmin={isAdmin} cloudDb={cloudDb} clienteNombre={clienteNombre} onLogout={cloudAuth ? () => signOut(cloudAuth) : null} />
      }
    </div>
  );
}

// ---------------------------------------------------------------------------
// Login (administrador o publicador)
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
        <div className="agri-serif" style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>Ingresar</div>
        <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 20 }}>Ingresá con tu cuenta para ver tus perfiles y sus dashboards.</div>
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
// Portal según el rol (etapa 2, 26/11/2026)
//   - Publicador: "Mis perfiles" (clientes/index filtrado por owner = su UID).
//   - Administrador: "Mis perfiles", "Publicadores" (con los perfiles de cada uno en ?publicador=<UID>)
//     y "Sin publicador" (perfiles cuyo owner no es un publicador registrado ni el administrador).
// Los perfiles se crean, renombran y eliminan desde Gorrión. En la App sólo se ven; el administrador
// además puede reasignar el dueño, eliminar un perfil y dar de baja a un publicador (borra sus perfiles).
// ---------------------------------------------------------------------------
const DASH_PORTAL = [
  { key: "cultivos",   label: "Cultivos",   nodo: "cultivos" },
  { key: "costos",     label: "Costos",     nodo: "costos" },
  { key: "produccion", label: "Producción", nodo: "produccion" },
];
const NODOS_DATOS = DASH_PORTAL.map((d) => d.nodo);

const urlBase = () => (typeof window !== "undefined" ? `${window.location.origin}${window.location.pathname}` : "");
const irA = (params) => {
  const u = new URLSearchParams(params);
  const qs = u.toString();
  window.history.pushState({}, "", qs ? `?${qs}` : urlBase());
  window.dispatchEvent(new Event("popstate"));
};
// Fecha y hora por separado: la tarjeta muestra la fecha en una línea y la hora completa en la siguiente
const fmtFechaHora = (iso) => {
  const d = iso ? new Date(iso) : null;
  if (!d || isNaN(d)) return null;
  return {
    fecha: d.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric" }),
    hora: d.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" }),
  };
};
const listaPerfiles = (val) => Object.entries(val || {})
  .filter(([, v]) => v && typeof v === "object")
  .map(([slug, v]) => ({ slug, nombre: v.nombre || slug, owner: v.owner || "", creadoEn: v.creadoEn || "", actualizado: v.actualizado || {} }))
  .sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));

// Rol del usuario: lee admins/<uid> y publicadores/<uid> (cada uno puede leer el suyo)
function useRol(user) {
  const [rol, setRol] = useState(undefined); // undefined = resolviendo
  useEffect(() => {
    if (!user) { setRol(null); return; }
    let vivo = true;
    setRol(undefined);
    (async () => {
      const leer = async (p) => { try { return (await get(ref(cloudDb, p))).val(); } catch (e) { return null; } };
      const esAdmin = await leer(`admins/${user.uid}`);
      const pub = esAdmin ? null : await leer(`publicadores/${user.uid}`);
      if (!vivo) return;
      setRol(esAdmin ? "admin" : pub ? "publicador" : "ninguno");
    })();
    return () => { vivo = false; };
  }, [user]);
  return rol;
}

// Link público de un perfil: usa config/appUrl si está cargado; si no, la dirección actual
function useAppUrl() {
  const [appUrl, setAppUrl] = useState("");
  useEffect(() => {
    get(ref(cloudDb, "config/appUrl")).then((s) => setAppUrl(String(s.val() || "").trim())).catch(() => {});
  }, []);
  return appUrl || urlBase();
}
const linkPerfil = (appUrl, slug) => `${appUrl}${appUrl.includes("?") ? "&" : "?"}cliente=${slug}`;

// Borra un perfil completo: cada nodo de datos por separado (una rama pesada no falla de una vez),
// después lo que quede en clientes/<slug> y por último la ficha del índice.
async function eliminarPerfilNube(slug) {
  for (const nodo of NODOS_DATOS) await dbSet(ref(cloudDb, `${CLOUD_CLIENTS_BASE}/${slug}/${nodo}`), null);
  await dbSet(ref(cloudDb, `${CLOUD_CLIENTS_BASE}/${slug}`), null);
  await dbSet(ref(cloudDb, `${CLOUD_CLIENTS_BASE}/index/${slug}`), null);
}

function TarjetaPerfil({ perfil, appUrl, esAdmin, publicadores, adminUid, onError }) {
  const [copiado, setCopiado] = useState(false);
  const [borrando, setBorrando] = useState(false);
  const link = linkPerfil(appUrl, perfil.slug);

  const copiar = () => {
    if (navigator.clipboard) navigator.clipboard.writeText(link).catch(() => {});
    setCopiado(true);
    setTimeout(() => setCopiado(false), 1500);
  };
  const reasignar = async (owner) => {
    if (owner === perfil.owner) return;
    try { await dbSet(ref(cloudDb, `${CLOUD_CLIENTS_BASE}/index/${perfil.slug}/owner`), owner); }
    catch (e) { onError(`No se pudo cambiar el dueño de "${perfil.nombre}" (${e.code || e.message}).`); }
  };
  const eliminar = async () => {
    if (!window.confirm(`¿Eliminar el perfil "${perfil.nombre}" con todos sus dashboards? El link deja de funcionar y no se puede deshacer.`)) return;
    setBorrando(true);
    try { await eliminarPerfilNube(perfil.slug); }
    catch (e) { onError(`No se pudo eliminar "${perfil.nombre}" (${e.code || e.message}).`); setBorrando(false); }
  };

  const opcionesDueno = esAdmin ? [
    { uid: adminUid, nombre: "Yo (administrador)" },
    ...Object.entries(publicadores || {}).filter(([uid]) => uid !== adminUid).map(([uid, p]) => ({ uid, nombre: (p && p.nombre) || uid })),
  ] : [];
  const duenoConocido = opcionesDueno.some((o) => o.uid === perfil.owner);

  return (
    <div className="agri-card" style={{ padding: 16, display: "flex", flexDirection: "column", gap: 12, opacity: borrando ? 0.5 : 1 }}>
      <div>
        <div className="agri-serif" style={{ fontWeight: 700, fontSize: 18, lineHeight: 1.2, marginBottom: 2 }}>{perfil.nombre}</div>
        <div style={{ fontSize: 12, color: "var(--ink-soft)", wordBreak: "break-all" }}>{link}</div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 8 }}>
        {DASH_PORTAL.map((d) => {
          const f = perfil.actualizado[d.nodo];
          return (
            <button key={d.key} className="agri-btn agri-btn-outline" disabled={!f}
              title={f ? `Abrir ${d.label}` : "Todavía no se publicó este dashboard"}
              onClick={() => irA({ cliente: perfil.slug, dashboard: d.key })}
              style={{ flexDirection: "column", alignItems: "flex-start", gap: 2, padding: "8px 10px", opacity: f ? 1 : 0.55, cursor: f ? "pointer" : "default" }}>
              <span>{d.label}</span>
              {f && fmtFechaHora(f) ? (
                <span style={{ fontSize: 11, fontWeight: 400, color: "var(--ink-soft)", lineHeight: 1.35, textAlign: "left" }}>
                  <span style={{ display: "block", whiteSpace: "nowrap" }}>{fmtFechaHora(f).fecha}</span>
                  <span style={{ display: "block", whiteSpace: "nowrap" }}>{fmtFechaHora(f).hora}</span>
                </span>
              ) : (
                <span style={{ fontSize: 11, fontWeight: 400, color: "var(--ink-soft)" }}>sin publicar</span>
              )}
            </button>
          );
        })}
      </div>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
        <button className="agri-btn agri-btn-outline" onClick={copiar}>{copiado ? "¡Copiado!" : "Copiar link"}</button>
        {esAdmin && (
          <>
            <select className="agri-select" value={duenoConocido ? perfil.owner : ""} onChange={(e) => reasignar(e.target.value)}
              title="Dueño del perfil (el publicador que lo publica desde su Gorrión)" style={{ flex: "1 1 150px", minWidth: 0 }}>
              {!duenoConocido && <option value="">Sin publicador</option>}
              {opcionesDueno.map((o) => <option key={o.uid} value={o.uid}>Dueño: {o.nombre}</option>)}
            </select>
            <button className="agri-btn agri-btn-outline" style={{ color: "var(--rust)" }} disabled={borrando} onClick={eliminar} title="Eliminar perfil">
              <Trash2 size={13} />
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function GrillaPerfiles({ perfiles, vacio, ...resto }) {
  if (perfiles === null) return <div style={{ padding: 30, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>;
  if (!perfiles.length) return <div className="agri-card" style={{ padding: 30, textAlign: "center", color: "var(--ink-soft)" }}>{vacio}</div>;
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(330px, 1fr))", gap: 14 }}>
      {perfiles.map((p) => <TarjetaPerfil key={p.slug} perfil={p} {...resto} />)}
    </div>
  );
}

function SeccionPublicadores({ publicadores, perfiles, adminUid, onError }) {
  const [uid, setUid] = useState("");
  const [nombre, setNombre] = useState("");
  const [email, setEmail] = useState("");
  const [guardando, setGuardando] = useState(false);
  const [baja, setBaja] = useState(null);       // { uid, nombre, confirmacion }
  const [progreso, setProgreso] = useState("");

  const lista = Object.entries(publicadores || {})
    .map(([u, p]) => ({ uid: u, nombre: (p && p.nombre) || u, email: (p && p.email) || "", perfiles: (perfiles || []).filter((x) => x.owner === u) }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));

  const registrar = async (e) => {
    e.preventDefault();
    const u = uid.trim(), n = nombre.trim(), m = email.trim();
    if (!u || !n) { onError("Completá UID y nombre."); return; }
    if (/[.#$[\]/]/.test(u)) { onError("El UID no puede tener . # $ [ ] /"); return; }
    if (publicadores && publicadores[u]) { onError("Ese UID ya está registrado como publicador."); return; }
    setGuardando(true);
    try { await dbSet(ref(cloudDb, `publicadores/${u}`), { nombre: n, email: m }); setUid(""); setNombre(""); setEmail(""); }
    catch (err) { onError(`No se pudo registrar (${err.code || err.message}).`); }
    finally { setGuardando(false); }
  };

  const darDeBaja = async () => {
    const pub = lista.find((p) => p.uid === baja.uid);
    if (!pub) return;
    try {
      for (let i = 0; i < pub.perfiles.length; i++) {
        setProgreso(`Eliminando perfil ${i + 1} de ${pub.perfiles.length}: ${pub.perfiles[i].nombre}…`);
        await eliminarPerfilNube(pub.perfiles[i].slug);
      }
      setProgreso("Eliminando el publicador…");
      await dbSet(ref(cloudDb, `publicadores/${pub.uid}`), null);
      setBaja(null);
    } catch (err) {
      onError(`La baja quedó incompleta (${err.code || err.message}). Volvé a intentarlo: sigue desde donde quedó.`);
    } finally { setProgreso(""); }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div className="agri-card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)" }}>
          <div className="agri-serif" style={{ fontSize: 15, fontWeight: 600 }}>Publicadores ({lista.length})</div>
        </div>
        {!lista.length && <div style={{ padding: 30, textAlign: "center", color: "var(--ink-soft)" }}>Todavía no hay publicadores registrados.</div>}
        {lista.map((p) => (
          <div key={p.uid} style={{ padding: "12px 16px", borderBottom: "1px solid var(--line)" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontWeight: 600, fontSize: 14 }}>{p.nombre}{p.uid === adminUid ? " (vos)" : ""}</div>
                <div style={{ fontSize: 12, color: "var(--ink-soft)", wordBreak: "break-all" }}>{p.email || "sin email"} · UID {p.uid}</div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button className="agri-btn agri-btn-outline" onClick={() => irA({ publicador: p.uid })}>
                  Ver perfiles ({p.perfiles.length})
                </button>
                {p.uid !== adminUid && (
                  <button className="agri-btn agri-btn-outline" style={{ color: "var(--rust)" }} onClick={() => setBaja({ uid: p.uid, nombre: p.nombre, confirmacion: "" })}>
                    Dar de baja
                  </button>
                )}
              </div>
            </div>
            {baja && baja.uid === p.uid && (
              <div style={{ marginTop: 12, padding: 12, border: "1px solid var(--rust)", borderRadius: 6, background: "rgba(161,70,47,0.05)" }}>
                <div style={{ fontSize: 13, marginBottom: 8 }}>
                  Se van a eliminar <b>{p.perfiles.length} {p.perfiles.length === 1 ? "perfil" : "perfiles"}</b> con todos sus dashboards
                  {p.perfiles.length ? ` (${p.perfiles.map((x) => x.nombre).join(", ")})` : ""} y el registro del publicador. Los links dejan de funcionar y no se puede deshacer.
                  <br />Después inhabilitá su usuario en Firebase → Authentication.
                </div>
                <div style={{ fontSize: 13, marginBottom: 6 }}>Para confirmar, escribí <b>{p.nombre}</b>:</div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  <input className="agri-input" style={{ flex: "1 1 200px" }} value={baja.confirmacion} disabled={!!progreso}
                    onChange={(e) => setBaja({ ...baja, confirmacion: e.target.value })} autoFocus />
                  <button className="agri-btn" style={{ background: "var(--rust)", borderColor: "var(--rust)" }}
                    disabled={baja.confirmacion.trim() !== p.nombre.trim() || !!progreso} onClick={darDeBaja}>
                    Dar de baja
                  </button>
                  <button className="agri-btn agri-btn-outline" disabled={!!progreso} onClick={() => setBaja(null)}>Cancelar</button>
                </div>
                {progreso && <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 6 }}>{progreso}</div>}
              </div>
            )}
          </div>
        ))}
      </div>

      <div className="agri-card" style={{ padding: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4 }}>Registrar publicador</div>
        <div style={{ fontSize: 12, color: "var(--ink-soft)", marginBottom: 10 }}>
          Primero creá el usuario en Firebase → Authentication → Agregar usuario, y copiá su UID acá.
        </div>
        <form onSubmit={registrar} style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <input className="agri-input" style={{ flex: "2 1 220px" }} placeholder="UID" value={uid} onChange={(e) => setUid(e.target.value)} />
          <input className="agri-input" style={{ flex: "1 1 160px" }} placeholder="Nombre del cliente" value={nombre} onChange={(e) => setNombre(e.target.value)} />
          <input className="agri-input" style={{ flex: "1 1 180px" }} type="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} />
          <button className="agri-btn" type="submit" disabled={guardando}><Plus size={14} /> Registrar</button>
        </form>
      </div>
    </div>
  );
}

function Portal({ user, rol }) {
  const esAdmin = rol === "admin";
  const appUrl = useAppUrl();
  const [perfiles, setPerfiles] = useState(null);
  const [publicadores, setPublicadores] = useState(esAdmin ? null : {});
  const [error, setError] = useState("");
  const [params, setParams] = useState(() => new URLSearchParams(window.location.search));

  useEffect(() => {
    const onPop = () => setParams(new URLSearchParams(window.location.search));
    window.addEventListener("popstate", onPop);
    return () => window.removeEventListener("popstate", onPop);
  }, []);

  // Perfiles: el administrador lee todo el índice; el publicador sólo los suyos (consulta por owner)
  useEffect(() => {
    const idx = ref(cloudDb, `${CLOUD_CLIENTS_BASE}/index`);
    const q = esAdmin ? idx : query(idx, orderByChild("owner"), equalTo(user.uid));
    return onValue(q, (s) => setPerfiles(listaPerfiles(s.val())),
      (e) => { setPerfiles([]); setError(`No se pudieron leer los perfiles (${e.code || e.message}).`); });
  }, [esAdmin, user.uid]);

  useEffect(() => {
    if (!esAdmin) return;
    return onValue(ref(cloudDb, "publicadores"), (s) => setPublicadores(s.val() || {}),
      (e) => { setPublicadores({}); setError(`No se pudieron leer los publicadores (${e.code || e.message}).`); });
  }, [esAdmin]);

  const vistaPub = params.get("publicador");
  const vista = !esAdmin ? "mios" : vistaPub ? "publicador" : (params.get("vista") || "mios");
  const mios = perfiles && perfiles.filter((p) => p.owner === user.uid);
  const sinPublicador = perfiles && publicadores
    ? perfiles.filter((p) => p.owner !== user.uid && !publicadores[p.owner]) : null;
  const delPublicador = perfiles && vistaPub ? perfiles.filter((p) => p.owner === vistaPub) : null;
  const pubActual = vistaPub && publicadores ? publicadores[vistaPub] : null;

  const props = { appUrl, esAdmin, publicadores, adminUid: user.uid, onError: setError };
  const tabs = [
    { key: "mios", label: `Mis perfiles${mios ? ` (${mios.length})` : ""}` },
    { key: "publicadores", label: `Publicadores${publicadores ? ` (${Object.keys(publicadores).length})` : ""}` },
    ...(sinPublicador && sinPublicador.length ? [{ key: "sin", label: `Sin publicador (${sinPublicador.length})` }] : []),
  ];

  return (
    <div className="agri-shell">
      <div className="agri-masthead">
        <div>
          <div className="agri-title agri-serif">{esAdmin ? "Panel de administrador" : "Mis perfiles"}</div>
          <div className="agri-sub">{user.email}{esAdmin ? " · administrador" : " · publicador"}</div>
          {esAdmin && (
            <div style={{ display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" }}>
              {tabs.map((t) => (
                <button key={t.key} className="agri-chip"
                  data-active={vista === t.key || (vista === "publicador" && t.key === "publicadores")}
                  onClick={() => irA(t.key === "mios" ? {} : { vista: t.key })}>
                  {t.label}
                </button>
              ))}
            </div>
          )}
        </div>
        <button className="agri-btn agri-btn-outline" onClick={() => signOut(cloudAuth)}>
          <LogOut size={14} /> Cerrar sesión
        </button>
      </div>

      {error && (
        <div className="agri-card" style={{ padding: "10px 14px", marginBottom: 14, borderColor: "var(--rust)", color: "var(--rust)", fontSize: 13, display: "flex", justifyContent: "space-between", gap: 10 }}>
          <span>{error}</span>
          <button className="agri-btn agri-btn-outline" style={{ padding: "2px 8px" }} onClick={() => setError("")}><X size={12} /></button>
        </div>
      )}

      {vista === "mios" && (
        <GrillaPerfiles perfiles={mios} vacio="Todavía no publicaste ningún perfil. Los perfiles se crean y publican desde Gorrión." {...props} />
      )}
      {vista === "publicadores" && (
        publicadores === null || perfiles === null
          ? <div style={{ padding: 30, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>
          : <SeccionPublicadores publicadores={publicadores} perfiles={perfiles} adminUid={user.uid} onError={setError} />
      )}
      {vista === "publicador" && (
        <>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
            <button className="agri-btn agri-btn-outline" onClick={() => irA({ vista: "publicadores" })}><ArrowLeft size={13} /> Publicadores</button>
            <div className="agri-serif" style={{ fontSize: 16 }}>
              {pubActual ? pubActual.nombre : vistaPub === user.uid ? "Mis perfiles" : "Publicador no registrado"}
            </div>
            {pubActual && pubActual.email && <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>{pubActual.email}</span>}
          </div>
          <GrillaPerfiles perfiles={delPublicador} vacio="Este publicador todavía no publicó ningún perfil." {...props} />
        </>
      )}
      {vista === "sin" && (
        <>
          <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 12 }}>
            Perfiles cuyo dueño no es un publicador registrado. Asignales un dueño o eliminalos.
          </div>
          <GrillaPerfiles perfiles={sinPublicador} vacio="No hay perfiles sin publicador." {...props} />
        </>
      )}
    </div>
  );
}

function SinRol() {
  return (
    <div className="agri-shell" style={{ maxWidth: 460, paddingTop: 70 }}>
      <div className="agri-card" style={{ padding: 28 }}>
        <div className="agri-serif" style={{ fontSize: 18, fontWeight: 700, marginBottom: 6 }}>Usuario sin acceso</div>
        <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 18 }}>
          Tu usuario no está registrado como publicador. Pedile al administrador que te dé de alta.
        </div>
        <button className="agri-btn agri-btn-outline" onClick={() => signOut(cloudAuth)}><LogOut size={14} /> Cerrar sesión</button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Componente raíz: decide qué mostrar según la URL (?cliente=slug) y la sesión
// ---------------------------------------------------------------------------
export default function App() {
  const [authUser, setAuthUser] = useState(undefined); // undefined = resolviendo, null = sin sesión
  const rol = useRol(authUser);
  const [clienteSlug, setClienteSlug] = useState(() =>
    typeof window === "undefined" ? null : new URLSearchParams(window.location.search).get("cliente"));

  // La navegación interna (portal ↔ dashboard) usa pushState + popstate, sin recargar
  useEffect(() => {
    const onPop = () => setClienteSlug(new URLSearchParams(window.location.search).get("cliente"));
    window.addEventListener("popstate", onPop);
    return () => window.removeEventListener("popstate", onPop);
  }, []);

  // Fuente tipográfica (una sola vez, para toda la app)
  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap";
    document.head.appendChild(link);
    return () => { try { document.head.removeChild(link); } catch (e) {} };
  }, []);

  // Sesión (persiste sola entre recargas gracias a Firebase Auth)
  useEffect(() => {
    if (!CLOUD_SYNC_ENABLED || !cloudAuth) { setAuthUser(null); return; }
    const unsub = onAuthStateChanged(cloudAuth, (u) => setAuthUser(u));
    return unsub;
  }, []);

  let content;
  if (!CLOUD_SYNC_ENABLED) {
    content = (
      <div className="agri-shell" style={{ paddingTop: 60, textAlign: "center", color: "var(--ink-soft)" }}>
        La App necesita Firebase configurado en <code>firebaseConfig.js</code>.
      </div>
    );
  } else if (authUser === undefined || (authUser && rol === undefined)) {
    content = <div style={{ padding: 60, textAlign: "center", color: "var(--ink-soft)" }}>Cargando…</div>;
  } else if (clienteSlug) {
    // Link de un perfil: público. Con sesión iniciada se muestra "Volver" al portal.
    content = <ClienteDashboard slug={clienteSlug} user={authUser} rol={rol} />;
  } else if (!authUser) {
    content = <LoginScreen />;
  } else if (rol === "admin" || rol === "publicador") {
    content = <Portal user={authUser} rol={rol} />;
  } else {
    content = <SinRol />;
  }

  return (
    <div className="agri-root">
      <style>{AGRI_STYLES}</style>
      {content}
    </div>
  );
}
