// ==========================================
// CÓDIGO GOOGLE APPS SCRIPT - NEGOCIO PISCINAS
// ==========================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var accion = data.accion;
    
    if (accion === "venta") {
      registrarVenta(data);
    } else if (accion === "compra") {
      registrarCompra(data);
    } else if (accion === "gasto") {
      registrarGasto(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({success: true}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// REGISTRAR VENTA
// ==========================================
function registrarVenta(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetVentas = ss.getSheetByName("Registro_Ventas");
  var sheetProductos = ss.getSheetByName("Productos");
  
  // Buscar el producto en la hoja Productos
  var datosProductos = sheetProductos.getDataRange().getValues();
  var productoEncontrado = null;
  var filaProducto = -1;
  
  for (var i = 1; i < datosProductos.length; i++) {
    if (datosProductos[i][0] === data.producto) {
      productoEncontrado = datosProductos[i];
      filaProducto = i + 1;
      break;
    }
  }
  
  if (!productoEncontrado) {
    throw new Error("Producto no encontrado");
  }
  
  var precioVenta = productoEncontrado[1];  // Columna B
  var precioCosto = productoEncontrado[2];  // Columna C
  var stockActual = productoEncontrado[3];  // Columna D
  
  // Si hay precio personalizado, usarlo
  if (data.precioPersonalizado && data.precioPersonalizado > 0) {
    precioVenta = data.precioPersonalizado;
  }
  
  // Calcular totales
  var cantidad = parseFloat(data.cantidad);
  var totalVenta = precioVenta * cantidad;
  var costoTotal = precioCosto * cantidad;
  var ganancia = totalVenta - costoTotal;
  
  // Agregar venta al registro
  sheetVentas.appendRow([
    new Date(),
    data.producto,
    cantidad,
    data.cliente || "Sin especificar",
    precioVenta,
    totalVenta,
    costoTotal,
    ganancia
  ]);
  
  // Actualizar stock (restar)
  var nuevoStock = stockActual - cantidad;
  sheetProductos.getRange(filaProducto, 4).setValue(nuevoStock);
}

// ==========================================
// REGISTRAR COMPRA DE STOCK
// ==========================================
function registrarCompra(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCompras = ss.getSheetByName("Compras_Stock");
  var sheetProductos = ss.getSheetByName("Productos");
  
  // Buscar el producto
  var datosProductos = sheetProductos.getDataRange().getValues();
  var filaProducto = -1;
  var stockActual = 0;
  
  for (var i = 1; i < datosProductos.length; i++) {
    if (datosProductos[i][0] === data.producto) {
      stockActual = datosProductos[i][3];
      filaProducto = i + 1;
      break;
    }
  }
  
  if (filaProducto === -1) {
    throw new Error("Producto no encontrado");
  }
  
  // Registrar la compra
  sheetCompras.appendRow([
    new Date(),
    data.producto,
    parseFloat(data.cantidad),
    parseFloat(data.precioTotal) || 0,
    data.notas || ""
  ]);
  
  // Actualizar stock (sumar)
  var nuevoStock = stockActual + parseFloat(data.cantidad);
  sheetProductos.getRange(filaProducto, 4).setValue(nuevoStock);
}

// ==========================================
// REGISTRAR GASTO
// ==========================================
function registrarGasto(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGastos = ss.getSheetByName("Gastos");
  
  sheetGastos.appendRow([
    new Date(),
    data.concepto,
    parseFloat(data.monto),
    data.notas || ""
  ]);
}
