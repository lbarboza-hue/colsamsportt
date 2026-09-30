// Pega aquí el ID de tu Google Sheet
const SHEET_ID = '12G2oNViW5vtl1qBmSE38PWH7lwNWETjxgEnhPPVzMc4';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Colsam-Sport | Torneo Intercursos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Validar y crear las pestañas en Google Sheets si no existen.
 */
function verificarEstructuraHojas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const hojasRequeridas = [
    { nombre: 'Usuarios', cols: ['Documento', 'Nombre Completo', 'Correo Institucional', 'Clave', 'Rol'] },
    { nombre: 'Salones', cols: ['ID', 'Nombre'] },
    { nombre: 'Jugadores', cols: ['ID', 'Nombre', 'Salon'] },
    { nombre: 'Partidos', cols: ['ID', 'Deporte', 'Equipo1', 'Equipo2', 'Fecha', 'Estado', 'Goles1', 'Goles2'] },
    { nombre: 'Goleadores', cols: ['Jugador', 'Salon', 'Deporte', 'Goles'] }
  ];

  hojasRequeridas.forEach(h => {
    let sheet = ss.getSheetByName(h.nombre);
    if (!sheet) {
      sheet = ss.insertSheet(h.nombre);
      sheet.appendRow(h.cols);
      
      // Si es usuarios, crear el admin por defecto
      if (h.nombre === 'Usuarios') {
        sheet.appendRow(['123', 'Admin Colsam', 'admin@colsam.edu.co', 'admin123', 'admin']);
      }
    }
  });
}

/**
 * Autenticar Usuario
 */
function validarLogin(email, pass) {
  verificarEstructuraHojas();
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Usuarios');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase() === email.toLowerCase() && data[i][3].toString() === pass) {
      return {
        success: true,
        user: { nombre: data[i][1], rol: data[i][4] }
      };
    }
  }
  return { success: false };
}

/**
 * Obtener todos los datos al cargar la app
 */
function obtenerDatosGlobales() {
  verificarEstructuraHojas();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Función helper para convertir tabla en JSON
  const getJson = (nombre) => {
    const data = ss.getSheetByName(nombre).getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const result = [];
    for(let i=1; i<data.length; i++) {
      let obj = {};
      headers.forEach((h, j) => obj[h] = data[i][j]);
      result.push(obj);
    }
    return result;
  };

  return {
    salones: getJson('Salones'),
    jugadores: getJson('Jugadores'),
    partidos: getJson('Partidos'),
    goleadores: getJson('Goleadores')
  };
}

/**
 * Crear Salón
 */
function crearSalon(nombre) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Salones');
  sheet.appendRow([Date.now(), nombre]);
  return { salones: obtenerDatosGlobales().salones };
}

/**
 * Inscribir Jugador
 */
function crearJugador(nombre, salon) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Jugadores');
  sheet.appendRow([Date.now(), nombre, salon]);
  return { jugadores: obtenerDatosGlobales().jugadores };
}

/**
 * Crear Editor (Admin only)
 */
function crearEditor(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Usuarios');
  sheet.appendRow([data.doc, data.nom, data.email, data.pass, 'editor']);
  return true;
}

/**
 * Guardar Partido y Asignar Goles a Jugadores
 */
function guardarPartidoYGoles(partido) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sPartidos = ss.getSheetByName('Partidos');
  const dPartidos = sPartidos.getDataRange().getValues();
  
  let rowIndex = -1;
  for (let i = 1; i < dPartidos.length; i++) {
    if (dPartidos[i][0].toString() === partido.id.toString()) {
      rowIndex = i + 1;
      break;
    }
  }

  const fila = [partido.id, partido.deporte, partido.equipo1, partido.equipo2, partido.fecha, partido.estado, partido.goles1, partido.goles2];
  
  if (rowIndex !== -1) {
    sPartidos.getRange(rowIndex, 1, 1, fila.length).setValues([fila]);
  } else {
    sPartidos.appendRow(fila);
  }

  // Acumular goles en Goleadores
  if (partido.goleadoresInfo && partido.goleadoresInfo.length > 0) {
    const sGol = ss.getSheetByName('Goleadores');
    
    partido.goleadoresInfo.forEach(g => {
      const dGol = sGol.getDataRange().getValues();
      let golRow = -1;
      
      for(let j=1; j<dGol.length; j++) {
        if(dGol[j][0] === g.jugador && dGol[j][2] === partido.deporte) {
          golRow = j + 1;
          break;
        }
      }
      
      if(golRow !== -1) {
        let currentGoals = parseInt(sGol.getRange(golRow, 4).getValue()) || 0;
        sGol.getRange(golRow, 4).setValue(currentGoals + 1);
      } else {
        sGol.appendRow([g.jugador, g.salon, partido.deporte, 1]);
      }
    });
  }

  return {
    partidos: obtenerDatosGlobales().partidos,
    goleadores: obtenerDatosGlobales().goleadores
  };
}
