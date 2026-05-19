/***************************************************************
 * SEND FINAL PROPOSAL PDF BY EMAIL (FROM CLIENT FOLDER)
 * + Inserts/updates log table from column D
 * + Reads REAL Gmail signature (SendAs.signature) and adds to email
 ***************************************************************/

function SendFinalProposalPDFByEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const rNombre = ui.prompt(
    "Client Name / 客户名称",
    'Enter name for greeting (will appear as "Dear XXX / 尊敬的XXX").\nEx: John Smith',
    ui.ButtonSet.OK_CANCEL
  );
  if (rNombre.getSelectedButton() !== ui.Button.OK) return;

  const nombreCliente = String(rNombre.getResponseText() || "").trim();
  if (!nombreCliente) throw new Error("Client name is required / 客户名称为必填项。");

  const rEmail = ui.prompt(
    "Recipient Email / 收件人邮箱",
    "Enter recipient email address.\nEx: client@company.com",
    ui.ButtonSet.OK_CANCEL
  );
  if (rEmail.getSelectedButton() !== ui.Button.OK) return;

  const email = String(rEmail.getResponseText() || "").trim();
  if (!isValidEmail_ENG_(email)) throw new Error("Email does not appear valid / 邮箱格式无效: " + email);

  const rFolder = ui.prompt(
    "Client Folder (Drive) / 客户文件夹",
    "Enter the exact name of the folder where the PDF is stored.",
    ui.ButtonSet.OK_CANCEL
  );
  if (rFolder.getSelectedButton() !== ui.Button.OK) return;

  const folderName = String(rFolder.getResponseText() || "").trim() || nombreCliente;

  let status = "SENT / 已发送";
  let pdfName = "";

  try {
    const folder = getFirstFolderByName_ENG_(folderName);
    if (!folder) throw new Error('Could not find a Drive folder named: "' + folderName + '" / 未找到名为"' + folderName + '"的文件夹。');

    const pdfFile = findBestPdfInFolder_ENG_(folder);
    if (!pdfFile) throw new Error('No PDF found inside folder: "' + folderName + '" / 文件夹中未找到PDF。');
    pdfName = pdfFile.getName();

    const plainBodyBase =
      "Dear " + nombreCliente + " / 尊敬的" + nombreCliente + ",\n\n" +
      "Please find attached the PDF with industrial warehouse proposals according to your requirements.\n" +
      "请查收附件中根据您需求整理的工业厂房/仓库方案PDF文件。\n\n" +
      "If you are looking for more or less space, a different location, or need to confirm any technical details, please let us know and we will adjust the document with other options. If you would like to review detailed technical sheets or schedule a site visit, we are happy to coordinate.\n" +
      "如需调整面积、区域或确认技术细节，请告知我们，我们将为您提供更多方案。如需查看详细技术资料或安排实地参观，我们随时可以协调安排。\n\n" +
      "We remain at your disposal.\n" +
      "我们随时恭候您的回复。\n";

    const htmlBodyBase =
      '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' +
        "<p>Dear " + escapeHtml_ENG_(nombreCliente) + " / 尊敬的" + escapeHtml_ENG_(nombreCliente) + ",</p>" +
        "<p>Please find attached the PDF with industrial warehouse proposals according to your requirements.<br>" +
        "请查收附件中根据您需求整理的工业厂房/仓库方案PDF文件。</p>" +
        "<p>If you are looking for more or less space, a different location, or need to confirm any technical details, please let us know and we will adjust the document with other options. If you would like to review detailed technical sheets or schedule a site visit, we are happy to coordinate.<br>" +
        "如需调整面积、区域或确认技术细节，请告知我们，我们将为您提供更多方案。如需查看详细技术资料或安排实地参观，我们随时可以协调安排。</p>" +
        "<p>We remain at your disposal.<br>" +
        "我们随时恭候您的回复。</p>" +
      "</div>";

    const subject = "Industrial Warehouse Proposals / 工业厂房方案 - Industrial Estate";

    const execEmail = getExecutingUserEmail_ENG_();
    const sendAsEmail = pickSendAsEmail_ENG_(execEmail);
    const signatureHtml = getGmailSignatureHtml_ENG_(sendAsEmail);

    const finalHtmlBody = appendSignatureHtml_ENG_(htmlBodyBase, signatureHtml);

    let signaturePlain = htmlToPlain_ENG_(sanitizeSignatureHtml_ENG_(signatureHtml));
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

    Logger.log("Executing user: " + execEmail);
    Logger.log("Send-as email: " + sendAsEmail);
    Logger.log("Sig len raw: " + String(signatureHtml || "").length);
    Logger.log("Sig len plain: " + signaturePlain.length);
    Logger.log("Final html len: " + finalHtmlBody.length);

    MailApp.sendEmail({
      to: email,
      cc: "sebastien@industrialestatemexico.com",
      subject: subject,
      body: finalPlainBody,
      htmlBody: finalHtmlBody,
      attachments: [pdfFile.getBlob().setName(pdfFile.getName())]
    });

    ss.toast("Email sent to " + email + " (attachment: " + pdfFile.getName() + ") / 邮件已发送", "Send PDF", 5);

  } catch (e) {
    status = "ERROR: " + String(e && e.message ? e.message : e).slice(0, 180);
    ss.toast("Could not send / 发送失败: " + status, "Send PDF", 6);
    throw e;

  } finally {
    appendEnvioLog_ENG_(sheet, {
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
 * SANITIZE SIGNATURE
 ***************************************************************/
function sanitizeSignatureHtml_ENG_(signatureHtml) {
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

function appendSignatureHtml_ENG_(htmlBody, signatureHtml) {
  const sig = sanitizeSignatureHtml_ENG_(signatureHtml);
  if (!sig) return htmlBody;

  const wrapper = '<div style="margin-top:14px;">' + sig + "</div>";
  return htmlBody + wrapper;
}

/***************************************************************
 * LOG TABLE FROM COLUMN D
 ***************************************************************/
function appendEnvioLog_ENG_(sheet, rowObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const colStart = 4;
  const meta = getOrCreateEnvioLogTable_ENG_(sheet, colStart);

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

function getOrCreateEnvioLogTable_ENG_(sheet, colStart) {
  const TITLE = "Email Send History / 发送记录";
  const headers = ["Client / 客户", "Email / 邮箱", "Status / 状态", "Date/Time / 日期时间", "File / 文件", "Folder / 文件夹"];

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

/***************************************************************
 * DRIVE HELPERS
 ***************************************************************/
function getFirstFolderByName_ENG_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function findBestPdfInFolder_ENG_(folder) {
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

    if (n.includes("proposal")) {
      if (updated > bestPropuestaUpdated) {
        bestPropuestaUpdated = updated;
        bestPropuesta = f;
      }
    }
  }

  return bestPropuesta || bestAny;
}

/***************************************************************
 * GMAIL SIGNATURE
 ***************************************************************/
function getExecutingUserEmail_ENG_() {
  let email = "";
  try { email = Session.getEffectiveUser().getEmail(); } catch (e) {}
  if (!email) {
    try { email = Session.getActiveUser().getEmail(); } catch (e2) {}
  }
  return String(email || "").trim().toLowerCase();
}

function pickSendAsEmail_ENG_(userEmail) {
  const CONTACTO = "contacto@industrialestatemexico.com";
  const SEBASTIEN_GMAIL = "sebastien.derieux@gmail.com";
  const SEBASTIEN_WORK = "sebastien@industrialestatemexico.com";

  if (userEmail === CONTACTO) return CONTACTO;
  if (userEmail === SEBASTIEN_WORK) return SEBASTIEN_WORK;
  if (userEmail === SEBASTIEN_GMAIL) return SEBASTIEN_GMAIL;

  return userEmail;
}

function getGmailSignatureHtml_ENG_(sendAsEmail) {
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

function htmlToPlain_ENG_(html) {
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
function isValidEmail_ENG_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function escapeHtml_ENG_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}