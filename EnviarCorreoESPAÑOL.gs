/***************************************************************
 * ENVIAR CORREOS DE PROPUESTAS Y/O FICHAS TÉCNICAS
 * + Caso A: Solo Fichas Técnicas (enviarFichasCliente) -> Busca en CARPETA_FICHAS_ID
 * + Caso B: Solo Propuesta PDF (EnviarPropuestaFinalPDFPorCorreo) -> Busca en CARPETA_PROPUESTAS_ID
 * + Caso C: Propuesta + Fichas Técnicas (enviarPropuestaYFichasCliente) -> Busca en AMBAS carpetas
 * + Inserta/actualiza tabla de log desde columna D
 * + Lee la firma REAL de Gmail (SendAs.signature) y la añade limpia al correo
 ***************************************************************/

// IDs de las Carpetas Padre
const ID_RAIZ_PROPUESTAS = "13aW3gRhAHuVlF2Wqc4R1KLqe3ZucQSPX";
const ID_RAIZ_FICHAS      = "1YG1LMk8D0zYWib7YqnH9q3XudZZY4_cY";

// =============================================================================
// CASO A: ENVIAR SOLO FICHAS TÉCNICAS
// =============================================================================
function enviarFichasCliente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const datosModal = pedirDatosEnvio_(ui);
  if (!datosModal) return;

  const { nombreCliente, email, folderName } = datosModal;
  let status = "ENVIADO";
  let pdfName = "";

  try {
    const folderFichas = getSubfolderInParent_(ID_RAIZ_FICHAS, folderName);
    if (!folderFichas) throw new Error('No encontré la carpeta "' + folderName + '" dentro de la raíz de Fichas.');

    const itFiles = folderFichas.getFilesByType(MimeType.PDF);
    const attachments = [];
    const pdfNames = [];

    while (itFiles.hasNext()) {
      const f = itFiles.next();
      attachments.push(f.getBlob().setName(f.getName()));
      pdfNames.push(f.getName());
    }

    if (attachments.length === 0) throw new Error('No encontré ningún PDF dentro de la carpeta de Fichas: "' + folderName + '".');
    pdfName = pdfNames.join(", ");

    const plainBodyBase =
      "Hola " + nombreCliente + ",\n\n" +
      "Te comparto en este correo las fichas técnicas de las opciones de naves industriales que consideramos se ajustan a tu requerimiento.\n\n" +
      "Si buscas una superficie mayor o menor, otra ubicación, o deseas confirmar algún aspecto técnico, con gusto ajustamos la selección y te enviamos nuevas opciones.\n\n" +
      "Si deseas agendar recorridos para conocer las propiedades, con gusto lo coordinamos con 1 a 2 días hábiles de anticipación.\n\n" +
      "Quedo a tu disposición.\n";

    const htmlBodyBase =
      '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' +
        "<p>Hola " + escapeHtml_(nombreCliente) + ",</p>" +
        "<p>Te comparto en este correo las fichas técnicas de las opciones de naves industriales que consideramos se ajustan a tu requerimiento.</p>" +
        "<p>Si buscas una superficie mayor o menor, otra ubicación, o deseas confirmar algún aspecto técnico, con gusto ajustamos la selección y te enviamos nuevas opciones.</p>" +
        "<p>Si deseas agendar recorridos para conocer las propiedades, con gusto lo coordinamos con 1 a 2 días hábiles de anticipación.</p>" +
        "<p>Quedo a tu disposición.</p>" +
      "</div>";

    const subject = "Fichas técnicas de naves industriales - Industrial Estate Mexico";

    enviarCorreoProcesado_({
      email,
      subject,
      plainBodyBase,
      htmlBodyBase,
      attachments
    });

    ss.toast("Correo enviado a " + email + " (" + attachments.length + " fichas adjuntas)", "Enviar Fichas", 5);

  } catch (e) {
    status = "ERROR: " + String(e && e.message ? e.message : e).slice(0, 180);
    ss.toast("No se pudo enviar: " + status, "Enviar Fichas", 6);
    throw e;

  } finally {
    appendEnvioLog_(sheet, {
      nombre: nombreCliente,
      correo: email,
      estatus: status,
      fecha: new Date(),
      archivo: pdfName,
      carpeta: folderName
    });
  }
}

// =============================================================================
// CASO B: ENVIAR SOLO RESUMEN / PROPUESTA PDF
// =============================================================================
function EnviarPropuestaFinalPDFPorCorreo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const datosModal = pedirDatosEnvio_(ui);
  if (!datosModal) return;

  const { nombreCliente, email, folderName } = datosModal;
  let status = "ENVIADO";
  let pdfName = "";

  try {
    const folderPropuestas = getSubfolderInParent_(ID_RAIZ_PROPUESTAS, folderName);
    if (!folderPropuestas) throw new Error('No encontré la carpeta "' + folderName + '" dentro de la raíz de Propuestas.');

    const pdfFile = findBestPdfInFolder_(folderPropuestas);
    if (!pdfFile) throw new Error('No encontré ningún PDF dentro de la carpeta de Propuestas: "' + folderName + '".');
    pdfName = pdfFile.getName();

    const plainBodyBase =
      "Hola " + nombreCliente + ",\n\n" +
      "Te comparto en este correo el PDF con las propuestas de naves industriales que consideramos se ajustan a tu requerimiento.\n\n" +
      "Si buscas una superficie mayor o menor, otra ubicación, o deseas confirmar algún aspecto técnico, con gusto ajustamos la propuesta y te enviamos nuevas opciones.\n\n" +
      "Si deseas revisar las fichas técnicas de alguna de las partidas incluidas en la propuesta, indícanos cuáles son de tu interés y con gusto te las enviaremos.\n\n" +
      "Si deseas agendar recorridos para conocer las propiedades, con gusto lo coordinamos con 1 a 2 días hábiles de anticipación.\n\n" +
      "Quedo a tu disposición.\n";

    const htmlBodyBase =
      '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' +
        "<p>Hola " + escapeHtml_(nombreCliente) + ",</p>" +
        "<p>Te comparto en este correo el PDF con las propuestas de naves industriales que consideramos se ajustan a tu requerimiento.</p>" +
        "<p>Si buscas una superficie mayor o menor, otra ubicación, o deseas confirmar algún aspecto técnico, con gusto ajustamos la propuesta y te enviamos nuevas opciones.</p>" +
        "<p>Si deseas revisar las fichas técnicas de alguna de las partidas incluidas en la propuesta, indícanos cuáles son de tu interés y con gusto te las enviaremos.</p>" +
        "<p>Si deseas agendar recorridos para conocer las propiedades, con gusto lo coordinamos con 1 a 2 días hábiles de anticipación.</p>" +
        "<p>Quedo a tu disposición.</p>" +
      "</div>";

    const subject = "Propuesta de naves industriales - Industrial Estate Mexico";

    enviarCorreoProcesado_({
      email,
      subject,
      plainBodyBase,
      htmlBodyBase,
      attachments: [pdfFile.getBlob().setName(pdfFile.getName())]
    });

    ss.toast("Correo enviado a " + email + " (adjunto: " + pdfFile.getName() + ")", "Enviar Propuesta", 5);

  } catch (e) {
    status = "ERROR: " + String(e && e.message ? e.message : e).slice(0, 180);
    ss.toast("No se pudo enviar: " + status, "Enviar Propuesta", 6);
    throw e;

  } finally {
    appendEnvioLog_(sheet, {
      nombre: nombreCliente,
      correo: email,
      estatus: status,
      fecha: new Date(),
      archivo: pdfName,
      carpeta: folderName
    });
  }
}

// =============================================================================
// CASO C: ENVIAR PROPUESTA (RESUMEN) + FICHAS TÉCNICAS (COMBINADAS)
// =============================================================================
function enviarPropuestaYFichasCliente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const datosModal = pedirDatosEnvio_(ui);
  if (!datosModal) return;

  const { nombreCliente, email, folderName } = datosModal;
  let status = "ENVIADO";
  let pdfName = "";

  try {
    const attachments = [];
    const pdfNames = [];

    // 1. Obtener la propuesta en la raíz de Propuestas
    const folderPropuestas = getSubfolderInParent_(ID_RAIZ_PROPUESTAS, folderName);
    if (folderPropuestas) {
      const pdfProp = findBestPdfInFolder_(folderPropuestas);
      if (pdfProp) {
        attachments.push(pdfProp.getBlob().setName(pdfProp.getName()));
        pdfNames.push(pdfProp.getName());
      }
    }

    // 2. Obtener las fichas en la raíz de Fichas
    const folderFichas = getSubfolderInParent_(ID_RAIZ_FICHAS, folderName);
    if (folderFichas) {
      const itFiles = folderFichas.getFilesByType(MimeType.PDF);
      while (itFiles.hasNext()) {
        const f = itFiles.next();
        attachments.push(f.getBlob().setName(f.getName()));
        pdfNames.push(f.getName());
      }
    }

    if (attachments.length === 0) {
      throw new Error('No encontré ningún PDF dentro de las carpetas de Propuestas o Fichas para: "' + folderName + '".');
    }
    pdfName = pdfNames.join(", ");

    const plainBodyBase =
      "Hola " + nombreCliente + ",\n\n" +
      "Te comparto en este correo la propuesta comercial, junto con las fichas técnicas correspondientes a las opciones de naves industriales que consideramos se ajustan a tu requerimiento.\n\n" +
      "Si buscas una superficie mayor o menor, otra ubicación, o deseas confirmar algún aspecto técnico, con gusto ajustamos la selección y te enviamos nuevas opciones.\n\n" +
      "Si deseas agendar recorridos para conocer las propiedades, con gusto lo coordinamos con 1 a 2 días hábiles de anticipación.\n\n" +
      "Saludos.\n";

    const htmlBodyBase =
      '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' +
        "<p>Hola " + escapeHtml_(nombreCliente) + ",</p>" +
        "<p>Te comparto en este correo la propuesta comercial, junto con las fichas técnicas correspondientes a las opciones de naves industriales que consideramos se ajustan a tu requerimiento.</p>" +
        "<p>Si buscas una superficie mayor o menor, otra ubicación, o deseas confirmar algún aspecto técnico, con gusto ajustamos la selección y te enviamos nuevas opciones.</p>" +
        "<p>Si deseas agendar recorridos para conocer las propiedades, con gusto lo coordinamos con 1 a 2 días hábiles de anticipación.</p>" +
        "<p>Saludos.</p>" +
      "</div>";

    const subject = "Propuesta y fichas técnicas de naves industriales - Industrial Estate Mexico";

    enviarCorreoProcesado_({
      email,
      subject,
      plainBodyBase,
      htmlBodyBase,
      attachments
    });

    ss.toast("Correo enviado a " + email + " (" + attachments.length + " adjuntos)", "Enviar Propuesta y Fichas", 5);

  } catch (e) {
    status = "ERROR: " + String(e && e.message ? e.message : e).slice(0, 180);
    ss.toast("No se pudo enviar: " + status, "Enviar Propuesta y Fichas", 6);
    throw e;

  } finally {
    appendEnvioLog_(sheet, {
      nombre: nombreCliente,
      correo: email,
      estatus: status,
      fecha: new Date(),
      archivo: pdfName,
      carpeta: folderName
    });
  }
}

// =============================================================================
// FUNCIONES AUXILIARES DE PROCESAMIENTO Y ENVÍO DE EMAIL
// =============================================================================

function pedirDatosEnvio_(ui) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const defaultFolder = sheet.getName();

  const rNombre = ui.prompt(
    "Nombre del cliente",
    'Escribe el nombre para el saludo (saldrá como "Hola XXX").\nEj: Juan Pérez',
    ui.ButtonSet.OK_CANCEL
  );
  if (rNombre.getSelectedButton() !== ui.Button.OK) return null;

  const nombreCliente = String(rNombre.getResponseText() || "").trim();
  if (!nombreCliente) throw new Error("El nombre del cliente es obligatorio.");

  const rEmail = ui.prompt(
    "Correo para envío",
    "Escribe el correo del destinatario.\nEj: cliente@empresa.com",
    ui.ButtonSet.OK_CANCEL
  );
  if (rEmail.getSelectedButton() !== ui.Button.OK) return null;

  const email = String(rEmail.getResponseText() || "").trim();
  if (!isValidEmail_(email)) throw new Error("El correo no parece válido: " + email);

  const rFolder = ui.prompt(
    "Carpeta del cliente (Drive)",
    'Escribe el nombre exacto de la carpeta.\nSi lo dejas en blanco se usará: "' + defaultFolder + '"',
    ui.ButtonSet.OK_CANCEL
  );
  if (rFolder.getSelectedButton() !== ui.Button.OK) return null;

  const folderName = String(rFolder.getResponseText() || "").trim() || defaultFolder;

  return { nombreCliente, email, folderName };
}

function enviarCorreoProcesado_(opts) {
  const execEmail = getExecutingUserEmail_();
  const sendAsEmail = pickSendAsEmail_(execEmail);
  const signatureHtml = getGmailSignatureHtml_(sendAsEmail);

  const finalHtmlBody = appendSignatureHtml_(opts.htmlBodyBase, signatureHtml);

  let signaturePlain = htmlToPlain_(sanitizeSignatureHtml_(signatureHtml));
  signaturePlain = signaturePlain
    .split("\n")
    .map(l => l.trim())
    .filter(l => {
      const x = l.toLowerCase();
      return x !== "anexo" && x !== "(anexo)";
    })
    .join("\n")
    .trim();

  const finalPlainBody = opts.plainBodyBase + (signaturePlain ? ("\n\n" + signaturePlain) : "");

  MailApp.sendEmail({
    to: opts.email,
    cc: "sebastien@industrialestatemexico.com",
    subject: opts.subject,
    body: finalPlainBody,
    htmlBody: finalHtmlBody,
    attachments: opts.attachments
  });
}

function sanitizeSignatureHtml_(signatureHtml) {
  let sig = String(signatureHtml || "").trim();
  if (!sig) return "";

  sig = sig
    .replace(/<o:p>\s*<\/o:p>/gi, "")
    .replace(/<o:p>.*?<\/o:p>/gi, "")
    .replace(/class="Mso[a-zA-Z0-9]+"/g, "")
    .replace(/style="[^"]*mso-[^"]*"/gi, 'style=""')
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "");

  sig = sig.replace(/<hr\b[^>]*>/gi, "");
  sig = sig.replace(/<a\b[^>]*>\s*<img\b[^>]*>\s*<\/a>/gi, "");
  sig = sig.replace(/<img\b[^>]*>/gi, "");

  return sig.trim();
}

function appendSignatureHtml_(htmlBody, signatureHtml) {
  const sig = sanitizeSignatureHtml_(signatureHtml);
  if (!sig) return htmlBody;

  const wrapper = '<div style="margin-top:14px;">' + sig + "</div>";
  return htmlBody + wrapper;
}

// =============================================================================
// LOG EN HOJA (TABLA DESDE COLUMNA D)
// =============================================================================
function appendEnvioLog_(sheet, rowObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const colStart = 4; // Columna D
  const meta = getOrCreateEnvioLogTable_(sheet, colStart);

  const dataStartRow = meta.headerRow + 1;
  const lastRow = Math.max(sheet.getLastRow(), dataStartRow);
  const vals = sheet.getRange(dataStartRow, colStart, lastRow - dataStartRow + 1, 1).getValues().flat();

  let writeRow = dataStartRow;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i] || "").trim() !== "") {
      writeRow = dataStartRow + i + 1;
      break;
    }
  }
  if (vals.length && String(vals[0] || "").trim() === "" && vals.every(v => String(v || "").trim() === "")) {
    writeRow = dataStartRow;
  }

  const tz = ss.getSpreadsheetTimeZone();
  const fechaStr = Utilities.formatDate(rowObj.fecha, tz, "dd/MM/yy HH:mm");

  sheet.getRange(writeRow, colStart, 1, 6).setValues([[
    rowObj.nombre || "",
    rowObj.correo || "",
    rowObj.estatus || "",
    fechaStr,
    rowObj.archivo || "",
    rowObj.carpeta || ""
  ]]);

  sheet.getRange(writeRow, colStart + 2).setWrap(true);
}

function getOrCreateEnvioLogTable_(sheet, colStart) {
  const TITLE = "Historial de envíos";
  const headers = ["Nombre", "Correo", "Estatus del envío", "Fecha/Hora", "Archivo", "Carpeta"];

  const found = sheet.createTextFinder(TITLE).matchEntireCell(true).findNext();
  if (found) {
    return { titleRow: found.getRow(), headerRow: found.getRow() + 1, colStart: found.getColumn() };
  }

  const startRow = sheet.getLastRow() + 3;

  sheet.getRange(startRow, colStart).setValue(TITLE).setFontWeight("bold");
  sheet.getRange(startRow + 1, colStart, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold");

  try {
    sheet.setColumnWidth(colStart + 0, 180);
    sheet.setColumnWidth(colStart + 1, 220);
    sheet.setColumnWidth(colStart + 2, 320);
    sheet.setColumnWidth(colStart + 3, 140);
    sheet.setColumnWidth(colStart + 4, 220);
    sheet.setColumnWidth(colStart + 5, 200);
  } catch (e) {}

  return { titleRow: startRow, headerRow: startRow + 1, colStart: colStart };
}

// =============================================================================
// DRIVE HELPERS (BÚSQUEDA INTELIGENTE QUE IGNORA LA FECHA AL FINAL)
// =============================================================================
function getSubfolderInParent_(parentId, folderName, autoCrear = false) {
  try {
    const parentFolder = DriveApp.getFolderById(parentId);
    const subfolders = parentFolder.getFolders();
    const targetNameClean = folderName.trim().toLowerCase();

    let coincidenciaPorInicio = null;

    while (subfolders.hasNext()) {
      const sub = subfolders.next();
      const subNameClean = sub.getName().trim().toLowerCase();

      // 1. Coincidencia Exacta (ej: "prueba01.08")
      if (subNameClean === targetNameClean) {
        return sub;
      }

      // 2. Coincidencia que COMIENZA con el nombre de la pestaña (ej: "prueba01.08 - 02-08-26")
      if (subNameClean.startsWith(targetNameClean)) {
        coincidenciaPorInicio = sub; // Guarda la carpeta más reciente que coincida
      }
    }

    if (coincidenciaPorInicio) {
      return coincidenciaPorInicio;
    }

    // 3. Si no existe ninguna y autoCrear es true, la crea
    if (autoCrear) {
      return parentFolder.createFolder(folderName.trim());
    }

    return null;
  } catch (e) {
    return null;
  }
}

function findBestPdfInFolder_(folder) {
  const it = folder.getFilesByType(MimeType.PDF);

  let bestPropuesta = null;
  let bestPropuestaUpdated = 0;

  let bestAny = null;
  let bestAnyUpdated = 0;

  while (it.hasNext()) {
    const f = it.next();
    const n = (f.getName() || "").toLowerCase();
    const updated = f.getLastUpdated().getTime();

    if (updated > bestAnyUpdated) {
      bestAnyUpdated = updated;
      bestAny = f;
    }

    if (n.includes("propuesta")) {
      if (updated > bestPropuestaUpdated) {
        bestPropuestaUpdated = updated;
        bestPropuesta = f;
      }
    }
  }

  return bestPropuesta || bestAny;
}

// =============================================================================
// FIRMA REAL DESDE GMAIL (Gmail Advanced Service)
// =============================================================================
function getExecutingUserEmail_() {
  let email = "";
  try { email = Session.getEffectiveUser().getEmail(); } catch (e) {}
  if (!email) {
    try { email = Session.getActiveUser().getEmail(); } catch (e2) {}
  }
  return String(email || "").trim().toLowerCase();
}

function pickSendAsEmail_(userEmail) {
  const CONTACTO = "contacto@industrialestatemexico.com";
  const SEBASTIEN_GMAIL = "sebastien.derieux@gmail.com";
  const SEBASTIEN_WORK = "sebastien@industrialestatemexico.com";

  if (userEmail === CONTACTO) return CONTACTO;
  if (userEmail === SEBASTIEN_WORK) return SEBASTIEN_WORK;
  if (userEmail === SEBASTIEN_GMAIL) return SEBASTIEN_GMAIL;

  return userEmail;
}

function getGmailSignatureHtml_(sendAsEmail) {
  const target = String(sendAsEmail || "").trim().toLowerCase();

  try {
    const list = Gmail.Users.Settings.SendAs.list("me");
    const arr = (list && list.sendAs) ? list.sendAs : [];

    const exact = arr.find(x => String(x.sendAsEmail || "").toLowerCase() === target);
    if (exact && String(exact.signature || "").trim()) return String(exact.signature);

    const def = arr.find(x => x.isDefault && String(x.signature || "").trim())
             || arr.find(x => x.isPrimary && String(x.signature || "").trim())
             || arr.find(x => String(x.signature || "").trim());

    return def ? String(def.signature || "") : "";
  } catch (e) {
    try {
      const sendAs = Gmail.Users.Settings.SendAs.get("me", sendAsEmail);
      return String(sendAs.signature || "");
    } catch (e2) {
      return "";
    }
  }
}

function htmlToPlain_(html) {
  return String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(div|p|tr|table|span)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .trim();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}