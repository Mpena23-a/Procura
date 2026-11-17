/* =========================================================================
   CONFIGURACIÓN DEL DASHBOARD
   Para actualizar los datos basta con reemplazar el Excel en la carpeta
   /data del repositorio manteniendo el MISMO nombre de archivo.
   ========================================================================= */
window.DASHBOARD_CONFIG = {
  // Ruta del Excel dentro del repositorio
  dataFile: 'data/Seguimiento eq Mantenimiento.xlsx',

  // Moneda para los montos de la hoja de anticipos
  currency: 'USD',

  // Días desde la emisión de la orden a partir de los cuales se marca como "antigua"
  agingAlertDays: 45,

  // Contraseña del botón "Cargar Excel" guardada como huella SHA-256 (no en texto plano).
  // Para cambiarla, calcula el SHA-256 de la nueva contraseña y pégalo aquí.
  uploadPasswordHash: '5f4ca1d191439bfc60e03bb3cff02c1a6b482ea27b8efac1e7485903811b35ad',

  // Cómo identificar cada hoja (por nombre). Si no coincide, se usa el orden de las hojas visibles.
  sheets: {
    mtto: /mtto|manten/i,
    oper: /oper/i,
    mat:  /material/i,
    rq:   /^rq|anticip|monto/i
  }
};
