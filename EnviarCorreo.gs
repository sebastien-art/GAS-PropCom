/***************************************************************
 * ENVIAR PDF PROPUESTA FINAL POR CORREO (DESDE CARPETA DEL CLIENTE)
 * + Inserta/actualiza tabla de log desde columna D
 * + Lee la firma REAL de Gmail (SendAs.signature) y la añade al correo
 *
 * ✅ FIX solicitado:
 * - Limpia la firma para evitar la línea "(anexo)" antes de la firma.
 * - Quita <img> embebidos y elimina "(anexo)" si viene en la firma.
 *
 * Popups:
 *  1) Nombre del cliente (para "Estimado XXX")
 *  2) Correo destino
 *  3) Nombre de la carpeta del cliente (Drive)
 ***************************************************************/

function EnviarPropuestaFinalPDFPorCorreo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // 1) Nombre del cliente
  const rNombre = ui.prompt(
    "Nombre del cliente",
    'Escribe el nombre para el saludo (saldrá como "Estimado XXX").\nEj: Juan Pérez',
    ui.ButtonSet.OK_CANCEL
  );
  if (rNombre.getSelectedButton() !== ui.Button.OK) return;

  const nombreCliente = String(rNombre.getResponseText() || "").trim();
  if (!nombreCliente) throw new Error("El nombre del cliente es obligatorio.");

  // 2) Correo destino
  const rEmail = ui.prompt(
    "Correo para envío",
    "Escribe el correo del destinatario.\nEj: cliente@empresa.com",
    ui.ButtonSet.OK_CANCEL
  );
  if (rEmail.getSelectedButton() !== ui.Button.OK) return;

  const email = String(rEmail.getResponseText() || "").trim();
  if (!isValidEmail_(email)) throw new Error("El correo no parece válido: " + email);

  // 3) Carpeta del cliente
  const rFolder = ui.prompt(
    "Carpeta del cliente (Drive)",
    "Escribe el nombre exacto de la carpeta donde está el PDF.",
    ui.ButtonSet.OK_CANCEL
  );
  if (rFolder.getSelectedButton() !== ui.Button.OK) return;

  const folderName = String(rFolder.getResponseText() || "").trim() || nombreCliente;

  let status = "ENVIADO";
  let pdfName = "";

  try {
    // === Buscar carpeta
    const folder = getFirstFolderByName_(folderName);
    if (!folder) throw new Error('No encontré una carpeta en Drive con el nombre: "' + folderName + '".');

    // === Buscar PDF (prioriza "propuesta"; si no, el más reciente)
    const pdfFile = findBestPdfInFolder_(folder);
    if (!pdfFile) throw new Error('No encontré ningún PDF dentro de la carpeta: "' + folderName + '".');
    pdfName = pdfFile.getName();

    // === Cuerpo EXACTO pedido
    const plainBodyBase =
      "Estimado " + nombreCliente + ",\n" +
      "Te comparto en este correo el PDF con las propuestas de naves industriales conforme a tu requerimiento.\n" +
      "Si buscas más o menos superficie, otra zona, o necesitas confirmar algún punto técnico, nos dices y ajustamos el documento con otras propuestas. Y si quieres ver fichas técnicas a detalle o agendar recorridos, con gusto lo coordinamos.\n" +
      "Quedo a tu disposición.\n";

    const htmlBodyBase =
      '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' +
        "<p>Estimado " + escapeHtml_(nombreCliente) + ",</p>" +
        "<p>Te comparto en este correo el PDF con las propuestas de naves industriales conforme a tu requerimiento.</p>" +
        "<p>Si buscas más o menos superficie, otra zona, o necesitas confirmar algún punto técnico, nos dices y ajustamos el documento con otras propuestas. Y si quieres ver fichas técnicas a detalle o agendar recorridos, con gusto lo coordinamos.</p>" +
        "<p>Quedo a tu disposición.</p>" +
      "</div>";

    const subject = "Propuestas de naves industriales (PDF)";

    // === Firma real según quién ejecuta
    const execEmail = getExecutingUserEmail_();
    const sendAsEmail = pickSendAsEmail_(execEmail);
    const signatureHtml = getGmailSignatureHtml_(sendAsEmail);

    // ✅ FIX: limpiar firma para evitar "(anexo)"
    const finalHtmlBody = appendSignatureHtml_(htmlBodyBase, signatureHtml);

    // ✅ FIX: limpiar firma también en texto plano
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

    const finalPlainBody = plainBodyBase + (signaturePlain ? ("\n\n" + signaturePlain) : "");

    // Logs de diagnóstico (opcional; ver en View > Logs)
    Logger.log("Executing user: " + execEmail);
    Logger.log("Send-as email: " + sendAsEmail);
    Logger.log("Sig len raw: " + String(signatureHtml || "").length);
    Logger.log("Sig len plain: " + signaturePlain.length);
    Logger.log("Final html len: " + finalHtmlBody.length);

    // === Envío
    MailApp.sendEmail({
  to: email,
  cc: "sebastien@industrialestatemexico.com",   // <-- AQUÍ
  subject: subject,
  body: finalPlainBody,
  htmlBody: finalHtmlBody,
  attachments: [pdfFile.getBlob().setName(pdfFile.getName())]
});

    ss.toast("Correo enviado a " + email + " (adjunto: " + pdfFile.getName() + ")", "Enviar PDF", 5);

  } catch (e) {
    status = "ERROR: " + String(e && e.message ? e.message : e).slice(0, 180);
    ss.toast("No se pudo enviar: " + status, "Enviar PDF", 6);
    throw e;

  } finally {
    // === Log en tabla desde columna D
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

/***************************************************************
 * ✅ FIX: Sanitizar firma para evitar "(anexo)"
 * - Quita imágenes embebidas (<img>) que suelen disparar "(anexo)"
 * - Quita "(anexo)" si viene dentro de la firma
 ***************************************************************/
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

  // ✅ Quita separadores horizontales dentro de la firma
  sig = sig.replace(/<hr\b[^>]*>/gi, "");

  // Quitar imágenes embebidas
  sig = sig.replace(/<a\b[^>]*>\s*<img\b[^>]*>\s*<\/a>/gi, "");
  sig = sig.replace(/<img\b[^>]*>/gi, "");

  return sig.trim();
}

function appendSignatureHtml_(htmlBody, signatureHtml) {
  const sig = sanitizeSignatureHtml_(signatureHtml);
  if (!sig) return htmlBody;

  // ✅ Sin border-top (sin separador)
  const wrapper = '<div style="margin-top:14px;">' + sig + "</div>";
  return htmlBody + wrapper;
}
  

/***************************************************************
 * LOG EN HOJA (TABLA DESDE COLUMNA D)
 ***************************************************************/
function appendEnvioLog_(sheet, rowObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const colStart = 4; // ✅ Columna D
  const meta = getOrCreateEnvioLogTable_(sheet, colStart);

  const dataStartRow = meta.headerRow + 1;

  // Encuentra siguiente fila libre (columna "Nombre")
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

  // Anchos (opcional)
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

/***************************************************************
 * DRIVE HELPERS (PDF dentro de carpeta)
 ***************************************************************/
function getFirstFolderByName_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
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

/***************************************************************
 * FIRMA REAL DESDE GMAIL (Gmail Advanced Service)
 ***************************************************************/
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

/***************************************************************
 * STRING HELPERS
 ***************************************************************/
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
