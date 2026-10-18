const SPREADSHEET_ID = "10zgHs7Tt6CA4YuJWe9Lfb6OyWPKJ9N8G9YlIrWzzjHo";
const DRIVE_FOLDER_NAME = "Sistema_Asistencia_UNEFA_Listas";
const MAX_FILE_BYTES = 7 * 1024 * 1024;
const MAX_ATTENDANCE_RECORDS = 500;

function doPost(e) {
  let request;
  let authenticated = false;
  try {
    request = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : "{}");
    const expectedKey = PropertiesService.getScriptProperties().getProperty("API_KEY");
    if (!expectedKey || !request.apiKey || request.apiKey !== expectedKey) {
      throw new Error("No autorizado.");
    }
    authenticated = true;
    if (!/^[a-f0-9]{32}$/.test(request.requestId || "")) {
      throw new Error("El identificador de solicitud no es válido.");
    }
    if (request.spreadsheetId !== SPREADSHEET_ID) {
      throw new Error("La hoja de cálculo indicada no es válida.");
    }

    let result;
    if (request.action === "ping") {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      DriveApp.getRootFolder().getId();
      result = { spreadsheetName: spreadsheet.getName(), driveAccess: true };
    } else if (request.action === "recordAttendance") {
      result = recordAttendance_(request.records);
    } else if (request.action === "saveListFile") {
      result = saveListFile_(request);
    } else {
      throw new Error("Acción no reconocida.");
    }
    CacheService.getScriptCache().put(
      "SYNC_" + request.requestId,
      JSON.stringify({ ok: true, result: result }),
      600
    );
    return jsonResponse_({ ok: true, result: result });
  } catch (error) {
    console.error(error);
    if (authenticated && request && /^[a-f0-9]{32}$/.test(request.requestId || "")) {
      CacheService.getScriptCache().put(
        "SYNC_" + request.requestId,
        JSON.stringify({ ok: false, error: error.message || String(error) }),
        600
      );
    }
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function doGet(e) {
  const requestId = e && e.parameter ? e.parameter.requestId : "";
  const callback = e && e.parameter ? e.parameter.callback : "";
  if (!/^[a-f0-9]{32}$/.test(requestId || "") || callback !== "__unefaGoogle_" + requestId) {
    return jsonResponse_({ ok: false, error: "Consulta de estado inválida." });
  }

  const cache = CacheService.getScriptCache();
  const savedResult = cache.get("SYNC_" + requestId);
  const response = savedResult
    ? JSON.parse(savedResult)
    : { pending: true };
  if (savedResult) cache.remove("SYNC_" + requestId);
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(response) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function recordAttendance_(records) {
  if (!Array.isArray(records) || records.length === 0 || records.length > MAX_ATTENDANCE_RECORDS) {
    throw new Error("La solicitud debe incluir entre 1 y 500 registros de asistencia.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getOrCreateSheet_("Asistencias", [
      "Fecha", "Hora de registro", "Sección", "Materia", "Cédula",
      "Alumno", "Género", "Método", "Archivo origen", "Clave única"
    ]);
    const rows = sheet.getDataRange().getValues();
    const existing = new Set();
    rows.slice(1).forEach(row => {
      existing.add(attendanceKey_(row[0], row[2], row[4]));
    });

    const newRows = [];
    const now = new Date();
    records.forEach(record => {
      validateAttendanceRecord_(record);
      const key = attendanceKey_(record.fecha, record.seccion, record.cedula);
      if (existing.has(key)) return;
      existing.add(key);
      newRows.push([
        record.fecha,
        Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss"),
        cleanText_(record.seccion, 80),
        cleanText_(record.materia, 160),
        cleanText_(record.cedula, 8),
        cleanText_(record.nombre, 160),
        cleanText_(record.genero, 10),
        cleanText_(record.metodo, 60),
        cleanText_(record.archivoOrigen || "", 200),
        key
      ]);
    });

    if (newRows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    return { inserted: newRows.length, duplicates: records.length - newRows.length };
  } finally {
    lock.releaseLock();
  }
}

function saveListFile_(request) {
  const fileName = cleanText_(request.fileName, 200);
  const mimeType = cleanText_(request.mimeType, 120);
  if (!fileName || !mimeType || typeof request.base64 !== "string") {
    throw new Error("Faltan el nombre, el tipo o el contenido del archivo.");
  }
  if (!/^(application\/pdf|image\/(jpeg|png|webp|heic|heif))$/i.test(mimeType)) {
    throw new Error("Solo se permiten archivos PDF o imágenes.");
  }
  if (request.base64.length > Math.ceil(MAX_FILE_BYTES * 4 / 3) + 8) {
    throw new Error("El archivo supera el límite de 7 MB.");
  }
  const date = cleanText_(request.date, 10);
  const section = cleanText_(request.section, 80);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !section) {
    throw new Error("Faltan una fecha válida o la sección de la lista.");
  }

  const bytes = Utilities.base64Decode(request.base64);
  if (!bytes.length || bytes.length > MAX_FILE_BYTES) {
    throw new Error("El archivo está vacío o supera el límite de 7 MB.");
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const hash = sha256_(bytes);
    const properties = PropertiesService.getScriptProperties();
    const hashProperty = "LIST_FILE_" + hash + "_" + date.replace(/\D/g, "") + "_" + section.replace(/[^a-z0-9]/gi, "");
    const previousFileId = properties.getProperty(hashProperty);
    if (previousFileId) {
      try {
        const previousFile = DriveApp.getFileById(previousFileId);
        return { fileId: previousFileId, fileName: previousFile.getName(), duplicate: true };
      } catch (error) {
        properties.deleteProperty(hashProperty);
      }
    }

    const folder = getOrCreateDriveFolder_();
    const safeName = cleanText_(fileName.replace(/[\\/:*?"<>|]/g, "_"), 200);
    const blob = Utilities.newBlob(bytes, mimeType, safeName);
    const file = folder.createFile(blob);
    file.setDescription([
      "Sistema de asistencia UNEFA",
      "Fecha: " + date,
      "Sección: " + section,
      "Origen: " + cleanText_(request.source, 80)
    ].join("\n"));

    const sheet = getOrCreateSheet_("Listas", [
      "Fecha de carga", "Fecha de lista", "Sección", "Materia",
      "Nombre de archivo", "Tipo", "Origen", "Enlace de Drive", "ID de Drive"
    ]);
    sheet.appendRow([
      new Date(),
      date,
      section,
      cleanText_(request.subject, 160),
      safeName,
      mimeType,
      cleanText_(request.source, 80),
      file.getUrl(),
      file.getId()
    ]);
    properties.setProperty(hashProperty, file.getId());
    return { fileId: file.getId(), fileName: safeName, duplicate: false };
  } finally {
    lock.releaseLock();
  }
}

function validateAttendanceRecord_(record) {
  if (!record || typeof record !== "object") {
    throw new Error("Un registro de asistencia tiene formato inválido.");
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(record.fecha || "")) {
    throw new Error("La fecha de asistencia no es válida.");
  }
  if (!/^\d{7,8}$/.test(String(record.cedula || ""))) {
    throw new Error("Una cédula de asistencia no es válida.");
  }
  if (!record.seccion || !record.nombre) {
    throw new Error("Faltan la sección o el nombre del alumno.");
  }
}

function attendanceKey_(date, section, cedula) {
  const normalizedDate = date instanceof Date
    ? Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd")
    : String(date).slice(0, 10);
  return [normalizedDate, String(section).trim(), String(cedula).trim()].join("|");
}

function getOrCreateSheet_(name, headers) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (headers.some((header, index) => currentHeaders[index] !== header)) {
      throw new Error("La pestaña “" + name + "” ya existe con encabezados distintos. No se modificó para evitar sobrescribir datos.");
    }
  }
  return sheet;
}

function getOrCreateDriveFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty("DRIVE_FOLDER_ID");
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (error) {
      properties.deleteProperty("DRIVE_FOLDER_ID");
    }
  }
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  properties.setProperty("DRIVE_FOLDER_ID", folder.getId());
  return folder;
}

function cleanText_(value, maxLength) {
  const text = String(value || "").trim().slice(0, maxLength);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function sha256_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    .map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, "0"))
    .join("");
}

function jsonResponse_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
