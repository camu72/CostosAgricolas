// Pegá acá los valores que te da Firebase al registrar tu app web
// (Configuración del proyecto → tus apps → ícono </> → firebaseConfig).
// Si dejás los valores de ejemplo tal cual, el dashboard sigue funcionando
// como antes: guarda los datos solo en el navegador de cada dispositivo.
export const firebaseConfig = {
  apiKey: "TU_API_KEY",
  authDomain: "TU_PROYECTO.firebaseapp.com",
  databaseURL: "https://TU_PROYECTO-default-rtdb.firebaseio.com",
  projectId: "TU_PROYECTO",
};

// Se activa automáticamente cuando completás los datos de arriba.
export const CLOUD_SYNC_ENABLED = !firebaseConfig.apiKey.startsWith("TU_");

// Prefijo dentro de la base de datos donde vive cada cliente. Cada uno se
// guarda en `clientes/{slug}/dataset`; el slug se toma de la URL (?cliente=...).
export const CLOUD_CLIENTS_BASE = "clientes";
