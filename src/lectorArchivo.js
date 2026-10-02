import * as XLSX from "xlsx";

// ---------------------------------------------------------------------------
// Lectura común de archivos para los dashboards: Excel (.xlsx/.xls/.xlsm) o JSON.
// Devuelve siempre filas como objetos { "Columna": valor }, igual que
// XLSX.utils.sheet_to_json, para que cada dashboard aplique su misma normalización.
//
// Formatos JSON aceptados:
//   1) { "columns": ["A","B",...], "rows": [[...], [...]], "fileName": "...", "updatedAt": "..." }
//      (el mismo formato que la App guarda en Firebase; fileName/updatedAt opcionales)
//   2) [ { "A": ..., "B": ... }, ... ]   (lista de objetos)
// En el formato 1, "" en una celda se toma como vacío (null).
// ---------------------------------------------------------------------------

export const ACCEPT_ARCHIVOS = ".xlsx,.xls,.xlsm,.json";

const esJSON = (file) => /\.json$/i.test(file.name) || file.type === "application/json";

/**
 * @param {File} file
 * @param {object} [opts]
 * @param {"primera"|"mayor"} [opts.hoja="primera"]  En Excel: primera hoja o la de más filas
 * @returns {Promise<{ filas: object[], fileName: string, origen: "excel"|"json" }>}
 */
export async function leerFilasArchivo(file, { hoja = "primera" } = {}) {
  if (esJSON(file)) return leerJSON(file);

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  let sheetName = wb.SheetNames[0];
  if (hoja === "mayor") {
    let best = -1;
    wb.SheetNames.forEach((n) => {
      const ref = wb.Sheets[n]["!ref"];
      const r = ref ? XLSX.utils.decode_range(ref) : null;
      const rc = r ? r.e.r - r.s.r : 0;
      if (rc > best) { best = rc; sheetName = n; }
    });
  }
  const filas = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null, raw: true });
  return { filas, fileName: file.name, origen: "excel" };
}

async function leerJSON(file) {
  let texto = await file.text();
  if (texto.charCodeAt(0) === 0xfeff) texto = texto.slice(1);   // BOM de UTF-8
  let data;
  try { data = JSON.parse(texto); }
  catch (e) { throw new Error(`el JSON no es válido (${e.message})`); }

  // Formato 2: lista de objetos
  if (Array.isArray(data)) {
    if (data.some((o) => !o || typeof o !== "object" || Array.isArray(o)))
      throw new Error("la lista debe contener objetos { columna: valor }");
    return { filas: data, fileName: file.name, origen: "json" };
  }

  // Formato 1: { columns, rows }
  if (!data || typeof data !== "object" || !Array.isArray(data.columns) || !Array.isArray(data.rows))
    throw new Error('se esperaba { "columns": [...], "rows": [[...]] } o una lista de objetos');
  const cols = data.columns.map((c) => String(c));
  const dup = cols.find((c, i) => cols.indexOf(c) !== i);
  if (dup) throw new Error(`la columna "${dup}" está repetida`);
  const filas = data.rows.map((arr, i) => {
    if (!Array.isArray(arr)) throw new Error(`la fila ${i + 1} no es una lista de valores`);
    if (arr.length !== cols.length)
      throw new Error(`la fila ${i + 1} tiene ${arr.length} valores y hay ${cols.length} columnas`);
    const o = {};
    cols.forEach((c, j) => { const v = arr[j]; o[c] = v === "" || v === undefined ? null : v; });
    return o;
  });
  return { filas, fileName: typeof data.fileName === "string" && data.fileName ? data.fileName : file.name, origen: "json" };
}

/**
 * Verifica que estén las columnas obligatorias. Cada elemento puede ser un nombre
 * o una lista de nombres alternativos (alcanza con que esté uno).
 */
export function validarColumnas(filas, requeridas) {
  if (!filas.length) throw new Error("el archivo no tiene filas de datos");
  const presentes = new Set();
  filas.slice(0, 50).forEach((f) => Object.keys(f).forEach((k) => presentes.add(k)));
  const faltan = requeridas
    .filter((req) => (Array.isArray(req) ? !req.some((r) => presentes.has(r)) : !presentes.has(req)))
    .map((req) => (Array.isArray(req) ? req.join(" o ") : req));
  if (faltan.length) throw new Error(`faltan columnas: ${faltan.join(", ")}`);
}
