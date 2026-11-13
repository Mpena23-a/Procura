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

  // Cómo identificar cada hoja (por nombre). Si no coincide, se usa el orden de las hojas visibles.
  sheets: {
    mtto: /mtto|manten/i,
    oper: /oper/i,
    mat:  /material/i,
    rq:   /^rq|anticip|monto/i
  }
};
