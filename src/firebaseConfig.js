// Pegá acá los valores que te da Firebase al registrar tu app web
// (Configuración del proyecto → tus apps → ícono </> → firebaseConfig).
// Si dejás los valores de ejemplo tal cual, el dashboard sigue funcionando
// como antes: guarda los datos solo en el navegador de cada dispositivo.
export const firebaseConfig = {
  apiKey: "AIzaSyADTGt-Av99zkXMS3suMsk3z6IG2_LQ5DA",
  authDomain: "costosagricolas-c3179.firebaseapp.com",
  databaseURL: "https://costosagricolas-c3179-default-rtdb.firebaseio.com",
  projectId: "costosagricolas-c3179",
};

// Se activa automáticamente cuando completás los datos de arriba.
export const CLOUD_SYNC_ENABLED = !firebaseConfig.apiKey.startsWith("TU_");

// Ruta dentro de la base de datos donde se guarda el dataset. No hace falta
// tocarla, pero tiene que coincidir con las reglas de seguridad de Firebase.
export const CLOUD_PATH = "produccionAgricola/dataset";
