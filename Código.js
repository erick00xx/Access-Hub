// El verdadero lenguaje de este archivo es .gs , pero para fines de edicion se ha nombrado como .js 
// para facilitar la lectura. Este archivo contiene el código del backend desarrollado en Google Apps Script, 
// que se encarga de manejar las operaciones CRUD sobre los proyectos, gestionar la subida de imágenes a Google Drive 
// y enviar notificaciones por email.


// ============================================
// CONFIGURACIÓN
// ============================================

const SPREADSHEET_ID = '11Ujs6F-0sfpY3tUTlId5J28avdux6buSgsQAVgprrTg';
const SHEET_NAME = 'Proyectos';
const DRIVE_FOLDER_ID = '10OlacDpa7rBZ4zyX4AImW8jKpaJA8I40';

// Columnas en el orden del Google Sheet
const COLUMNS = {
  ID: 0,
  FECHA_CREACION: 1,
  MARCA: 2,
  RESPONSABLE: 3,
  AREA: 4,
  PARTICIPANTES: 5,
  REQUERIMIENTO: 6,
  COMENTARIOS: 7,
  NOMBRE_PROYECTO: 8,
  ACCESOS: 9,
  PLATAFORMAS: 10,
  IMAGENES: 11,
  ESTADO: 12,
  FECHA_ACTUALIZACION: 13,
  COMENTARIOS_ESTADO: 14
};

// ============================================
// ENDPOINTS PRINCIPALES
// ============================================

/**
 * Maneja las solicitudes GET
 */
function doGet(e) {
  const action = e.parameter.action;
  let response;
  
  try {
    switch(action) {
      case 'getProyectos':
        response = getProyectos();
        break;
      case 'getProyecto':
        response = getProyecto(e.parameter.id);
        break;
      case 'getResponsables':
        response = getResponsables();
        break;
      case 'getPassword':
        response = getPassword();
        break;
      default:
        response = { success: false, message: 'Acción no válida' };
    }
  } catch(error) {
    response = { success: false, message: error.toString() };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Maneja las solicitudes POST
 */
function doPost(e) {
  let response;
  
  try {
    // Usamos el contenido parseado de la petición
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch(action) {
      case 'createProyecto':
        response = createProyecto(data.proyecto);
        break;
      case 'updateProyecto':
        response = updateProyecto(data.proyecto);
        break;
      case 'deleteProyecto':
        response = deleteProyecto(data.id);
        break;
      case 'uploadImage':
        response = uploadImage(data.fileName, data.base64Data, data.mimeType);
        break;
      default:
        response = { success: false, message: 'Acción no válida' };
    }
  } catch(error) {
    response = { success: false, message: error.toString() };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// OPERACIONES CRUD
// ============================================

/**
 * Obtiene todos los proyectos
 */
function getProyectos() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Saltar la fila de encabezados
  const proyectos = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[COLUMNS.ID]) { // Solo incluir filas con ID
      proyectos.push(rowToProyecto(row));
    }
  }
  
  return { success: true, data: proyectos };
}

/**
 * Obtiene un proyecto por ID
 */
function getProyecto(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.ID].toString() === id.toString()) {
      return { success: true, data: rowToProyecto(data[i]) };
    }
  }
  
  return { success: false, message: 'Proyecto no encontrado' };
}

/**
 * Crea un nuevo proyecto
 */
function createProyecto(proyecto) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  
  // Generar ID si no tiene
  if (!proyecto.id) {
    proyecto.id = generateId();
  }
  
  // Establecer fecha de creación
  if (!proyecto.fechaCreacion) {
    proyecto.fechaCreacion = formatDate(new Date());
  }
  
  // Convertir proyecto a fila
  const row = proyectoToRow(proyecto);
  
  // Agregar al final de la hoja
  sheet.appendRow(row);
  
  return { success: true, data: proyecto, message: 'Proyecto creado exitosamente' };
}

/**
 * Actualiza un proyecto existente
 */
function updateProyecto(proyecto) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Buscar la fila del proyecto
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.ID].toString() === proyecto.id.toString()) {
      rowIndex = i + 1; // +1 porque las filas en Sheets empiezan en 1
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: 'Proyecto no encontrado' };
  }
  
  // Actualizar fecha de actualización
  proyecto.fechaActualizacion = formatDate(new Date());
  
  // Convertir proyecto a fila
  const row = proyectoToRow(proyecto);
  
  // Actualizar la fila
  const range = sheet.getRange(rowIndex, 1, 1, row.length);
  range.setValues([row]);
  
  return { success: true, data: proyecto, message: 'Proyecto actualizado exitosamente' };
}

/**
 * Elimina un proyecto
 */
function deleteProyecto(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Buscar la fila del proyecto
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.ID].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Proyecto eliminado exitosamente' };
    }
  }
  
  return { success: false, message: 'Proyecto no encontrado' };
}

// ============================================
// GESTIÓN DE ARCHIVOS (GOOGLE DRIVE)
// ============================================

/**
 * Sube un archivo a Google Drive
 * Nombre final: yyyyMMdd_HHmm_XXXXX_nombreOriginal.ext
 * Donde XXXXX es un identificador aleatorio único de 5 caracteres
 */
function uploadImage(fileName, base64Data, mimeType) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    
    // Generar nombre con fecha-hora Perú + ID único
    const now      = new Date();
    const datePart = Utilities.formatDate(now, 'America/Lima', 'yyyyMMdd_HHmm');
    const uniqueId = Math.random().toString(36).substring(2, 7).toUpperCase();
    const safeOrig = fileName.replace(/[^a-zA-Z0-9._\-]/g, '_');
    const finalName = `${datePart}_${uniqueId}_${safeOrig}`;
    
    // Decodificar base64 y crear archivo
    const decoded = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(decoded, mimeType, finalName);
    const file    = folder.createFile(blob);
    
    // Permisos públicos de lectura
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileId     = file.getId();
    const isImage    = mimeType.startsWith('image/');
    
    // thumbnail?id= funciona para <img> sin problemas de CORS/CSP
    const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`;
    const viewUrl      = `https://drive.google.com/file/d/${fileId}/view`;
    const previewUrl   = `https://drive.google.com/file/d/${fileId}/preview`;
    
    return {
      success: true,
      data: {
        fileId        : fileId,
        fileName      : finalName,
        originalName  : fileName,
        url           : isImage ? thumbnailUrl : viewUrl,   // valor principal guardado en sheet
        thumbnailUrl  : thumbnailUrl,
        viewUrl       : viewUrl,
        previewUrl    : previewUrl,
        isImage       : isImage
      }
    };
  } catch(error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Lista los archivos de la carpeta de imágenes
 */
function listImages() {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFiles();
    const images = [];
    
    while (files.hasNext()) {
      const file = files.next();
      images.push({
        id: file.getId(),
        name: file.getName(),
        url: `https://drive.google.com/uc?export=view&id=${file.getId()}`,
        webViewLink: file.getUrl(),
        mimeType: file.getMimeType(),
        dateCreated: file.getDateCreated()
      });
    }
    
    return { success: true, data: images };
  } catch(error) {
    return { success: false, message: error.toString() };
  }
}

// ============================================
// RESPONSABLES Y CONTRASEÑA
// ============================================

/**
 * Lee la hoja "Responsables" y devuelve los nombres (columna A, sin cabecera).
 */
function getResponsables() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Responsables');
  if (!sheet) return { success: false, message: 'Hoja Responsables no encontrada' };

  const data   = sheet.getDataRange().getValues();
  const nombres = [];
  // Saltamos fila 0 (encabezados: Nombre | Cargo | Correo)
  for (let i = 1; i < data.length; i++) {
    const nombre = (data[i][0] || '').toString().trim();
    if (nombre) nombres.push(nombre);
  }
  return { success: true, data: nombres };
}

/**
 * Lee la contraseña de administrador desde la hoja "Password", celda A1.
 */
function getPassword() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Password');
  if (!sheet) return { success: false, message: 'Hoja Password no encontrada' };

  const pass = (sheet.getRange('A1').getValue() || '').toString().trim();
  return { success: true, data: pass };
}

// ============================================
// FUNCIONES AUXILIARES
// ============================================

/**
 * Convierte una fila del sheet a objeto proyecto
 */
function rowToProyecto(row) {
  return {
    id: row[COLUMNS.ID]?.toString() || '',
    fechaCreacion: formatDateValue(row[COLUMNS.FECHA_CREACION]),
    marca: row[COLUMNS.MARCA] || '',
    responsable: row[COLUMNS.RESPONSABLE] || '',
    area: row[COLUMNS.AREA] || '',
    participantes: row[COLUMNS.PARTICIPANTES] || '',
    requerimiento: row[COLUMNS.REQUERIMIENTO] || '',
    comentarios: row[COLUMNS.COMENTARIOS] || '',
    nombreProyecto: row[COLUMNS.NOMBRE_PROYECTO] || '',
    accesos: parseMultipleItems(row[COLUMNS.ACCESOS]),
    plataformas: parseMultipleItems(row[COLUMNS.PLATAFORMAS]),
    imagenes: parseMultipleItems(row[COLUMNS.IMAGENES]),
    estado: row[COLUMNS.ESTADO] || 'En proceso',
    fechaActualizacion: formatDateValue(row[COLUMNS.FECHA_ACTUALIZACION]),
    comentariosEstado: row[COLUMNS.COMENTARIOS_ESTADO] || ''
  };
}

/**
 * Convierte un objeto proyecto a fila del sheet
 */
function proyectoToRow(proyecto) {
  return [
    proyecto.id || '',
    proyecto.fechaCreacion || '',
    proyecto.marca || '',
    proyecto.responsable || '',
    proyecto.area || '',
    proyecto.participantes || '',
    proyecto.requerimiento || '',
    proyecto.comentarios || '',
    proyecto.nombreProyecto || '',
    serializeMultipleItems(proyecto.accesos),
    serializeMultipleItems(proyecto.plataformas),
    serializeMultipleItems(proyecto.imagenes),
    proyecto.estado || 'En proceso',
    proyecto.fechaActualizacion || '',
    proyecto.comentariosEstado || ''
  ];
}

/**
 * Parsea items múltiples (formato: "titulo: valor || titulo2: valor2")
 */
function parseMultipleItems(value) {
  if (!value) return [];
  
  const items = [];
  const parts = value.toString().split('||');
  
  parts.forEach(part => {
    const trimmed = part.trim();
    if (trimmed) {
      // Nuevo separador: '>>>' (evita conflictos con URLs que usan ':')
      const sepIdx = trimmed.indexOf('>>>');
      if (sepIdx > -1) {
        items.push({
          titulo: trimmed.substring(0, sepIdx).trim(),
          valor: trimmed.substring(sepIdx + 3).trim()
        });
      } else {
        // Compatibilidad hacia atrás: separador legado ':'
        const colonIndex = trimmed.indexOf(':');
        if (colonIndex > -1) {
          items.push({
            titulo: trimmed.substring(0, colonIndex).trim(),
            valor: trimmed.substring(colonIndex + 1).trim()
          });
        } else {
          items.push({
            titulo: '',
            valor: trimmed
          });
        }
      }
    }
  });
  
  return items;
}

/**
 * Serializa items múltiples a string
 */
function serializeMultipleItems(items) {
  if (!items || !Array.isArray(items) || items.length === 0) return '';
  
  return items
    .filter(item => item.titulo || item.valor)
    .map(item => `${item.titulo}>>>${item.valor}`)
    .join(' || ');
}

/**
 * Genera un ID único
 */
function generateId() {
  return Math.floor(100000000 + Math.random() * 900000000).toString();
}

/**
 * Formatea una fecha con hora en zona horaria de Perú (America/Lima)
 * Formato de salida: d/M/yyyy, HH:mm  Ej: 10/3/2026, 18:30
 */
function formatDate(date) {
  return Utilities.formatDate(date, 'America/Lima', 'd/M/yyyy, HH:mm');
}

/**
 * Convierte un valor de celda (Date o string) a formato de fecha legible
 */
function formatDateValue(value) {
  if (!value) return '';
  if (value instanceof Date) return formatDate(value);
  return value.toString();
}

// ============================================
// FUNCIONES DE EMAIL (OPCIONAL)
// ============================================

/**
 * Envía una notificación por email cuando se crea un proyecto
 */
function sendNotificationEmail(proyecto, recipients) {
  const subject = `Nuevo Proyecto: ${proyecto.nombreProyecto}`;
  
  const htmlBody = `
    <h2>Nuevo Proyecto Registrado</h2>
    <p><strong>Nombre:</strong> ${proyecto.nombreProyecto}</p>
    <p><strong>Marca:</strong> ${proyecto.marca}</p>
    <p><strong>Responsable:</strong> ${proyecto.responsable}</p>
    <p><strong>Área:</strong> ${proyecto.area}</p>
    <p><strong>Fecha:</strong> ${proyecto.fechaCreacion}</p>
    <hr>
    <h3>Requerimiento:</h3>
    <div>${proyecto.requerimiento}</div>
  `;
  
  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: htmlBody
    });
    return { success: true, message: 'Email enviado' };
  } catch(error) {
    return { success: false, message: error.toString() };
  }
}

// ============================================
// TESTING (Ejecutar desde el editor)
// ============================================

/**
 * Función de prueba para verificar la conexión
 */
function testConnection() {
  const result = getProyectos();
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * Función de prueba para crear un proyecto
 */
function testCreateProyecto() {
  const proyecto = {
    marca: 'Test',
    responsable: 'Erick',
    area: 'OSE',
    participantes: 'Test User',
    requerimiento: '<p>Este es un requerimiento de prueba</p>',
    comentarios: '<p>Comentarios de prueba</p>',
    nombreProyecto: 'Proyecto de Prueba',
    accesos: [{ titulo: 'Admin', valor: 'admin:123' }],
    plataformas: [{ titulo: 'Github', valor: 'https://github.com' }],
    imagenes: [],
    estado: 'En proceso',
    comentariosEstado: ''
  };
  
  const result = createProyecto(proyecto);
  Logger.log(JSON.stringify(result, null, 2));
}
