import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, onValue, get, set as dbSet } from "firebase/database";
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
function ClienteDashboard({ slug, isAdmin }) {
  const baseUrl = typeof window !== "undefined" ? `${window.location.origin}${window.location.pathname}` : "";

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
    window.history.pushState({}, "", `?${u.toString()}`);
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
      const snap = await get(ref(cloudDb, `clientes/${slug}/costos`));
      if (snap.exists()) {
        dbSet(ref(cloudDb, `clientes/${newSlug}/costos`), snap.val());
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
