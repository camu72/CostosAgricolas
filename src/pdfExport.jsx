import React, { useState, useRef, useEffect } from "react";
import { FileDown, ChevronDown } from "lucide-react";

// ---------------------------------------------------------------------------
// Exportación a PDF, común a todos los dashboards.
//
// Cada dashboard marca sus bloques con el atributo data-pdf:
//   data-pdf="resumen"   → KPIs y gráficos (entran en Resumen, Pantalla y Completo)
//   data-pdf="pantalla"  → tablas tal como se ven (sólo en Pantalla)
// Los elementos con data-pdf-ignore (botones, etc.) no se dibujan.
// En modo Completo, las tablas se generan desde los datos (no desde la pantalla),
// así salen enteras aunque estén colapsadas.
// ---------------------------------------------------------------------------

export const MODOS_PDF = [
  { key: "resumen",  label: "Resumen",                  desc: "Filtros, indicadores y gráficos" },
  { key: "pantalla", label: "Lo que se ve en pantalla", desc: "Todo como está ahora, abierto o colapsado" },
  { key: "completo", label: "Completo",                 desc: "Resumen + todas las tablas de detalle" },
];

// Colores de la app (RGB) para jsPDF
const C = {
  ink: [43, 33, 24], soft: [107, 94, 79], line: [220, 210, 184],
  green: [75, 107, 58], paper: [246, 241, 226], raised: [252, 250, 242], grupo: [240, 235, 224],
};

const PAGE = { w: 297, h: 210, m: 10 };            // A4 horizontal, mm
const CONTENT_W = PAGE.w - PAGE.m * 2;
const TOP = 12, BOTTOM = PAGE.h - 12;              // zona útil (deja lugar a encabezado/pie)

// jsPDF usa WinAnsi: reemplazamos los caracteres que no puede dibujar
const pdfTxt = (v) => String(v ?? "")
  .replace(/[−–]/g, "-").replace(/[≥]/g, ">=").replace(/[≤]/g, "<=")
  .replace(/[▸▾]/g, "");

// Fila de grupo (ocupa todas las columnas) para usar dentro de body de una tabla
export const filaGrupo = (texto, nCols, extra = []) => {
  const celdas = [{ content: texto, colSpan: nCols - extra.length, styles: { fontStyle: "bold", fillColor: C.grupo, textColor: C.ink } }];
  extra.forEach((e) => celdas.push({ content: e, styles: { fontStyle: "bold", fillColor: C.grupo, textColor: C.ink, halign: "right" } }));
  return celdas;
};
export const filaTotal = (celdas) => celdas.map((c) => ({ content: c, styles: { fontStyle: "bold", fillColor: C.raised } }));

/**
 * @param {object}   o
 * @param {string}   o.modo          "resumen" | "pantalla" | "completo"
 * @param {string}   o.titulo        p.ej. "Producción"
 * @param {string}   o.cliente
 * @param {Array}    o.filtros       [[etiqueta, valor], ...] sólo los activos
 * @param {Element}  o.root          contenedor del dashboard (donde están los data-pdf)
 * @param {Array}    o.tablas        para modo completo: [{ titulo, nota?, head:[...], body:[[...]], align:[...] }]
 * @param {string}   o.archivo       nombre base del archivo
 */
export async function generarPDF({ modo, titulo, cliente, filtros, root, tablas = [], archivo }) {
  const [{ jsPDF }, { autoTable }, { default: html2canvas }] = await Promise.all([
    import("jspdf"), import("jspdf-autotable"), import("html2canvas"),
  ]);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setFont("helvetica", "normal");

  // ---------- Encabezado de la primera página ----------
  const ahora = new Date();
  const fechaTxt = ahora.toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
  const modoTxt = MODOS_PDF.find((m) => m.key === modo)?.label || "";
  doc.setTextColor(...C.ink); doc.setFont("helvetica", "bold"); doc.setFontSize(16);
  doc.text(pdfTxt(`${titulo} - ${cliente}`), PAGE.m, 16);
  doc.setFont("helvetica", "normal"); doc.setFontSize(9); doc.setTextColor(...C.soft);
  doc.text(pdfTxt(`Generado el ${fechaTxt} · ${modoTxt}`), PAGE.m, 21.5);
  const filtrosTxt = filtros.length ? "Filtros: " + filtros.map(([k, v]) => `${k}: ${v}`).join("  ·  ") : "Filtros: ninguno (todos los datos)";
  const lineas = doc.splitTextToSize(pdfTxt(filtrosTxt), CONTENT_W);
  doc.setTextColor(...C.ink);
  doc.text(lineas, PAGE.m, 27);
  let y = 27 + lineas.length * 4 + 1;
  doc.setDrawColor(...C.ink); doc.setLineWidth(0.4); doc.line(PAGE.m, y, PAGE.w - PAGE.m, y);
  y += 4;

  const nuevaPagina = () => { doc.addPage(); y = TOP; };

  // ---------- Bloques capturados de la pantalla ----------
  const selector = modo === "pantalla" ? "[data-pdf]" : '[data-pdf="resumen"]';
  const bloques = root ? Array.from(root.querySelectorAll(selector)) : [];
  for (const el of bloques) {
    if (!el.offsetParent && el.getClientRects().length === 0) continue; // oculto
    const canvas = await html2canvas(el, {
      scale: 2, backgroundColor: "#FFFFFF", logging: false, useCORS: true,
      ignoreElements: (n) => n.hasAttribute && n.hasAttribute("data-pdf-ignore"),
    });
    if (!canvas.width || !canvas.height) continue;
    const pxPorMm = canvas.width / CONTENT_W;
    const altoTotal = canvas.height / pxPorMm;
    const util = BOTTOM - TOP;

    if (altoTotal <= BOTTOM - y) {
      doc.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", PAGE.m, y, CONTENT_W, altoTotal);
      y += altoTotal + 4;
      continue;
    }
    if (altoTotal <= util) {           // entra en una página nueva sin cortarse
      nuevaPagina();
      doc.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", PAGE.m, y, CONTENT_W, altoTotal);
      y += altoTotal + 4;
      continue;
    }
    // Bloque más alto que una página: se corta en tramos, siempre entre filas de tabla
    const rect = el.getBoundingClientRect();
    const escala = canvas.height / rect.height;
    const cortes = Array.from(el.querySelectorAll("tr"))
      .map((tr) => Math.round((tr.getBoundingClientRect().bottom - rect.top) * escala))
      .filter((c) => c > 0 && c < canvas.height)
      .sort((a, b) => a - b);
    let offsetPx = 0;
    if (BOTTOM - y < 40) nuevaPagina();
    while (offsetPx < canvas.height) {
      const disponibleMm = BOTTOM - y;
      const maxPx = Math.floor(disponibleMm * pxPorMm);
      let tramoPx = Math.min(canvas.height - offsetPx, maxPx);
      if (offsetPx + tramoPx < canvas.height) {
        const limite = offsetPx + maxPx;
        const corte = cortes.filter((c) => c > offsetPx + maxPx * 0.3 && c <= limite).pop();
        if (corte) tramoPx = corte - offsetPx;
      }
      const c2 = document.createElement("canvas");
      c2.width = canvas.width; c2.height = tramoPx;
      c2.getContext("2d").drawImage(canvas, 0, offsetPx, canvas.width, tramoPx, 0, 0, canvas.width, tramoPx);
      doc.addImage(c2.toDataURL("image/jpeg", 0.92), "JPEG", PAGE.m, y, CONTENT_W, tramoPx / pxPorMm);
      offsetPx += tramoPx;
      y += tramoPx / pxPorMm + 4;
      if (offsetPx < canvas.height) nuevaPagina();
    }
  }

  // ---------- Tablas completas (modo completo) ----------
  if (modo === "completo") {
    for (const t of tablas) {
      if (BOTTOM - y < 30) nuevaPagina();
      doc.setFont("helvetica", "bold"); doc.setFontSize(12); doc.setTextColor(...C.ink);
      doc.text(pdfTxt(t.titulo), PAGE.m, y + 4);
      y += 6;
      if (t.nota) {
        doc.setFont("helvetica", "normal"); doc.setFontSize(8.5); doc.setTextColor(...C.soft);
        const nl = doc.splitTextToSize(pdfTxt(t.nota), CONTENT_W);
        doc.text(nl, PAGE.m, y + 3);
        y += nl.length * 3.6 + 1;
      }
      const columnStyles = {};
      (t.align || []).forEach((a, i) => { if (a) columnStyles[i] = { halign: a }; });
      const limpiar = (cell) => (cell && typeof cell === "object" ? { ...cell, content: pdfTxt(cell.content) } : pdfTxt(cell));
      autoTable(doc, {
        startY: y + 1,
        head: [t.head.map(pdfTxt)],
        body: t.body.map((r) => r.map(limpiar)),
        margin: { left: PAGE.m, right: PAGE.m, top: TOP, bottom: PAGE.h - BOTTOM },
        styles: { font: "helvetica", fontSize: 8, cellPadding: 1.4, textColor: C.ink, lineColor: C.line, lineWidth: 0.1 },
        headStyles: { fillColor: C.paper, textColor: C.soft, fontStyle: "bold", lineWidth: 0, fontSize: 7.5 },
        alternateRowStyles: { fillColor: [255, 255, 255] },
        bodyStyles: { fillColor: [255, 255, 255] },
        columnStyles,
        showHead: "everyPage",
      });
      y = doc.lastAutoTable.finalY + 8;
    }
  }

  // ---------- Pie de página en todas las hojas ----------
  const n = doc.getNumberOfPages();
  for (let i = 1; i <= n; i++) {
    doc.setPage(i);
    doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(...C.soft);
    doc.setDrawColor(...C.line); doc.setLineWidth(0.2);
    doc.line(PAGE.m, PAGE.h - 8.5, PAGE.w - PAGE.m, PAGE.h - 8.5);
    doc.text(pdfTxt(`${titulo} - ${cliente}`), PAGE.m, PAGE.h - 5);
    doc.text(`Página ${i} de ${n}`, PAGE.w - PAGE.m, PAGE.h - 5, { align: "right" });
  }

  const fecha = ahora.toISOString().slice(0, 10);
  const slugify = (s) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Za-z0-9]+/g, "_").replace(/^_|_$/g, "");
  doc.save(`${slugify(archivo || titulo)}_${slugify(cliente)}_${modo}_${fecha}.pdf`);
}

// ---------------------------------------------------------------------------
// Botón "PDF" con menú de 3 opciones
// ---------------------------------------------------------------------------
export function PdfMenu({ onExport }) {
  const [abierto, setAbierto] = useState(false);
  const [generando, setGenerando] = useState(null);
  const [err, setErr] = useState("");
  const ref = useRef(null);

  useEffect(() => {
    if (!abierto) return;
    const cerrar = (e) => { if (ref.current && !ref.current.contains(e.target)) setAbierto(false); };
    document.addEventListener("mousedown", cerrar);
    return () => document.removeEventListener("mousedown", cerrar);
  }, [abierto]);

  const elegir = async (modo) => {
    setAbierto(false); setGenerando(modo); setErr("");
    try { await onExport(modo); }
    catch (e) { console.error(e); setErr("No se pudo generar el PDF"); }
    finally { setGenerando(null); }
  };

  return (
    <div ref={ref} style={{ position: "relative" }} data-pdf-ignore>
      <button className="agri-btn agri-btn-outline" disabled={!!generando} onClick={() => setAbierto((v) => !v)} title={err || "Descargar PDF"}
        style={err ? { borderColor: "var(--rust)", color: "var(--rust)" } : undefined}>
        <FileDown size={14} /> {generando ? "Generando…" : "PDF"} <ChevronDown size={12} />
      </button>
      {abierto && (
        <div style={{ position: "absolute", right: 0, top: "calc(100% + 4px)", zIndex: 50, width: 250, background: "var(--paper-raised)", border: "1px solid var(--line)", borderRadius: 8, boxShadow: "0 6px 18px rgba(43,33,24,0.14)", padding: 4 }}>
          {MODOS_PDF.map((m) => (
            <button key={m.key} onClick={() => elegir(m.key)}
              style={{ display: "block", width: "100%", textAlign: "left", background: "none", border: "none", borderRadius: 6, padding: "8px 10px", cursor: "pointer", font: "inherit", color: "var(--ink)" }}
              onMouseEnter={(e) => { e.currentTarget.style.background = "rgba(107,94,79,0.08)"; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = "none"; }}>
              <div style={{ fontSize: 13, fontWeight: 600 }}>{m.label}</div>
              <div style={{ fontSize: 11, color: "var(--ink-soft)" }}>{m.desc}</div>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
