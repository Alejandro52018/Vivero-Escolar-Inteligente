// ==================== VIVERO ESCOLAR INTELIGENTE - CÓDIGO COMPLETO CORREGIDO ====================
// Proyecto SENATIC - Grupo 3147177
// I.E. Rafael Uribe Uribe - Buesaco, Nariño
// Equipo: Nicolas Chavez, Eliana Jojoa, David Guerrero, Angie Jojoa, Andres Ordoñez

// ==================== CONFIGURACIÓN ====================
const ID_HOJA = '1_OkfO1BDLyOMtgJyragA-o4TGHiKcpMA8KqxyiARSgM';

function obtenerHoja() {
  return SpreadsheetApp.openById(ID_HOJA);
}

function obtenerHojaEspecifica(nombre) {
  const ss = obtenerHoja();
  return ss.getSheetByName(nombre);
}

// ==================== FUNCIÓN PRINCIPAL ====================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Vivero Escolar Inteligente')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== SISTEMA DE LOGIN ====================
function verificarLogin(correo, contrasena) {
  try {
    const sheet = obtenerHojaEspecifica('usuarios');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const correoBD = String(fila[2]).trim().toLowerCase();
      const contrasenaBD = String(fila[3]).trim();
      const correoIngresado = String(correo).trim().toLowerCase();
      const contrasenaIngresada = String(contrasena).trim();
      if (correoBD === correoIngresado && contrasenaBD === contrasenaIngresada) {
        return { exito: true, mensaje: '¡Bienvenido al Vivero Escolar!', usuario: { id: fila[0], nombre: fila[1], correo: fila[2], rol: fila[4] } };
      }
    }
    return { exito: false, mensaje: 'Correo o contraseña incorrectos' };
  } catch (error) {
    Logger.log('Error en verificarLogin: ' + error.toString());
    return { exito: false, mensaje: 'Error del servidor: ' + error.toString() };
  }
}

function registrarUsuario(nombre, correo, contrasena) {
  try {
    const sheet = obtenerHojaEspecifica('usuarios');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toLowerCase() === String(correo).trim().toLowerCase()) {
        return { exito: false, mensaje: 'Este correo ya está registrado' };
      }
    }
    const nuevoId = data.length;
    sheet.appendRow([nuevoId, nombre, correo, "'" + contrasena, 'estudiante', new Date()]);
    return { exito: true, mensaje: 'Usuario registrado exitosamente', usuario: { id: nuevoId, nombre: nombre, correo: correo, rol: 'estudiante' } };
  } catch (error) {
    Logger.log('Error en registrarUsuario: ' + error.toString());
    return { exito: false, mensaje: 'Error: ' + error.toString() };
  }
}

// ==================== CRUD DE SEMILLAS ====================
function crearSemilla(nombreComun, nombreCientifico, tipo, temporada) {
  try {
    const sheet = obtenerHojaEspecifica('semillas');
    const data = sheet.getDataRange().getValues();
    const nuevoId = data.length;
    sheet.appendRow([nuevoId, nombreComun, nombreCientifico, tipo, temporada]);
    return { exito: true, mensaje: 'Semilla registrada', semilla: { id: nuevoId, nombreComun, nombreCientifico, tipo, temporada } };
  } catch (error) {
    return { exito: false, mensaje: 'Error: ' + error.toString() };
  }
}

function listarSemillas() {
  try {
    const sheet = obtenerHojaEspecifica('semillas');
    const data = sheet.getDataRange().getValues();
    const semillas = [];
    for (let i = 1; i < data.length; i++) {
      semillas.push({ id: data[i][0], nombreComun: data[i][1], nombreCientifico: data[i][2], tipo: data[i][3], temporada: data[i][4] });
    }
    return semillas;
  } catch (error) {
    Logger.log('Error en listarSemillas: ' + error.toString());
    return [];
  }
}

function eliminarSemilla(id) {
  try {
    const sheet = obtenerHojaEspecifica('semillas');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { exito: true, mensaje: 'Semilla eliminada' };
      }
    }
    return { exito: false, mensaje: 'Semilla no encontrada' };
  } catch (error) {
    return { exito: false, mensaje: 'Error: ' + error.toString() };
  }
}

// ==================== CRUD DE SIEMBRAS (CORREGIDO) ====================
function crearSiembra(semillaId, fecha, cantidad, ubicacion, responsable) {
  try {
    const sheet = obtenerHojaEspecifica('siembras');
    const data = sheet.getDataRange().getValues();
    const nuevoId = data.length;
    sheet.appendRow([nuevoId, parseInt(semillaId), fecha, parseInt(cantidad), ubicacion, responsable]);
    
    // Crear recordatorios en la pestaña calendario si existe
    const calendarioSheet = obtenerHojaEspecifica('calendario');
    if (calendarioSheet) {
      const calendarioData = calendarioSheet.getDataRange().getValues();
      const hoy = new Date(fecha);
      for (let d = 0; d < 30; d += 2) {
        const fechaRiego = new Date(hoy);
        fechaRiego.setDate(hoy.getDate() + d);
        const idCalendario = calendarioData.length;
        calendarioSheet.appendRow([idCalendario, nuevoId, fechaRiego.toISOString().split('T')[0], false]);
      }
    }
    return { exito: true, mensaje: 'Siembra registrada y calendario creado', siembra: { id: nuevoId, semillaId, fecha, cantidad, ubicacion, responsable } };
  } catch (error) {
    return { exito: false, mensaje: 'Error: ' + error.toString() };
  }
}

function listarSiembras() {
  try {
    const sheet = obtenerHojaEspecifica('siembras');
    const data = sheet.getDataRange().getValues();
    const siembras = [];
    const semillasSheet = obtenerHojaEspecifica('semillas');
    const semillasData = semillasSheet ? semillasSheet.getDataRange().getValues() : [];
    
    const semillasMap = {};
    for (let i = 1; i < semillasData.length; i++) {
      semillasMap[semillasData[i][0]] = semillasData[i][1];
    }
    
    for (let i = 1; i < data.length; i++) {
      // Formateo seguro de fecha para evitar errores al enviar datos al frontend
      let fechaVal = data[i][2];
      if (fechaVal instanceof Date) {
        fechaVal = fechaVal.toLocaleDateString();
      } else {
        fechaVal = String(fechaVal || '');
      }

      siembras.push({ 
        id: data[i][0], 
        semillaId: data[i][1], 
        nombreSemilla: semillasMap[data[i][1]] || 'Desconocida', 
        fecha: fechaVal, 
        cantidad: data[i][3], 
        ubicacion: data[i][4], 
        responsable: data[i][5] 
      });
    }
    return siembras;
  } catch (error) {
    Logger.log('Error en listarSiembras: ' + error.toString());
    return [];
  }
}

function eliminarSiembra(id) {
  try {
    const sheet = obtenerHojaEspecifica('siembras');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { exito: true, mensaje: 'Siembra eliminada' };
      }
    }
    return { exito: false, mensaje: 'Siembra no encontrada' };
  } catch (error) {
    return { exito: false, mensaje: 'Error: ' + error.toString() };
  }
}

// ==================== CRUD DE RIEGOS (CORREGIDO) ====================
function crearRiego(siembraId, fecha, cantidadLitros, responsable, observaciones) {
  try {
    const sheet = obtenerHojaEspecifica('riegos');
    const data = sheet.getDataRange().getValues();
    const nuevoId = data.length;
    sheet.appendRow([nuevoId, parseInt(siembraId), fecha, parseFloat(cantidadLitros), responsable, observaciones]);
    
    const calendarioSheet = obtenerHojaEspecifica('calendario');
    if (calendarioSheet) {
      const calendarioData = calendarioSheet.getDataRange().getValues();
      for (let i = 1; i < calendarioData.length; i++) {
        if (calendarioData[i][1] == siembraId && String(calendarioData[i][2]).includes(fecha)) {
          calendarioSheet.getRange(i + 1, 4).setValue(true);
          break;
        }
      }
    }
    return { exito: true, mensaje: 'Riego registrado', riego: { id: nuevoId, siembraId, fecha, cantidadLitros, responsable, observaciones } };
  } catch (error) {
    return { exito: false, mensaje: 'Error: ' + error.toString() };
  }
}

function listarRiegos() {
  try {
    const sheet = obtenerHojaEspecifica('riegos');
    const data = sheet.getDataRange().getValues();
    const riegos = [];
    const siembrasSheet = obtenerHojaEspecifica('siembras');
    const siembrasData = siembrasSheet ? siembrasSheet.getDataRange().getValues() : [];
    
    const siembrasMap = {};
    for (let i = 1; i < siembrasData.length; i++) {
      siembrasMap[siembrasData[i][0]] = siembrasData[i][4];
    }
    
    for (let i = 1; i < data.length; i++) {
      let fechaVal = data[i][2];
      if (fechaVal instanceof Date) {
        fechaVal = fechaVal.toLocaleDateString();
      } else {
        fechaVal = String(fechaVal || '');
      }

      riegos.push({ 
        id: data[i][0], 
        siembraId: data[i][1], 
        ubicacion: siembrasMap[data[i][1]] || 'Desconocida', 
        fecha: fechaVal, 
        cantidadLitros: data[i][3], 
        responsable: data[i][4], 
        observaciones: data[i][5] 
      });
    }
    return riegos;
  } catch (error) {
    Logger.log('Error en listarRiegos: ' + error.toString());
    return [];
  }
}

// ==================== REPORTE DEL VIVERO ====================
function obtenerReporte() {
  try {
    const siembrasSheet = obtenerHojaEspecifica('siembras');
    const siembrasData = siembrasSheet ? siembrasSheet.getDataRange().getValues() : [];
    const riegosSheet = obtenerHojaEspecifica('riegos');
    const riegosData = riegosSheet ? riegosSheet.getDataRange().getValues() : [];
    const semillasSheet = obtenerHojaEspecifica('semillas');
    const semillasData = semillasSheet ? semillasSheet.getDataRange().getValues() : [];
    const calendarioSheet = obtenerHojaEspecifica('calendario');
    const calendarioData = calendarioSheet ? calendarioSheet.getDataRange().getValues() : [];
    
    let totalRiegos = 0, totalLitros = 0, riegosPendientes = 0;
    for (let i = 1; i < riegosData.length; i++) {
      totalRiegos++;
      totalLitros += parseFloat(riegosData[i][3]) || 0;
    }
    for (let i = 1; i < calendarioData.length; i++) {
      if (!calendarioData[i][3]) riegosPendientes++;
    }
    return {
      totalSemillas: Math.max(0, semillasData.length - 1),
      totalSiembras: Math.max(0, siembrasData.length - 1),
      totalRiegos: totalRiegos,
      totalLitros: Math.round(totalLitros * 100) / 100,
      riegosPendientes: riegosPendientes,
      mensaje: '🌱 El vivero tiene ' + Math.max(0, siembrasData.length - 1) + ' siembras activas y ' + riegosPendientes + ' riegos pendientes'
    };
  } catch (error) {
    return { totalSemillas: 0, totalSiembras: 0, totalRiegos: 0, totalLitros: 0, riegosPendientes: 0, mensaje: 'Error: ' + error.toString() };
  }
}

// ==================== FUNCIÓN DE PRUEBA ====================
function probarConexion() {
  try {
    const ss = obtenerHoja();
    Logger.log('✅ Conexión exitosa con la hoja: ' + ss.getName());
    const hojas = ss.getSheets();
    Logger.log('📋 Hojas encontradas: ' + hojas.map(h => h.getName()).join(', '));
    return { exito: true, nombre: ss.getName(), hojas: hojas.map(h => h.getName()) };
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return { exito: false, mensaje: error.toString() };
  }
}
