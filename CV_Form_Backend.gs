/**
 * CV Intake Backend
 * Deploy: Extensions > Apps Script > paste this > Deploy > Web App
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the /exec URL into SCRIPT_URL in intake.html.
 *
 * First-time setup: run authTest() once manually from the Apps Script
 * editor (select it in the function dropdown, click Run) so Google
 * asks you to authorize Docs/Drive/Sheets access. Without this the
 * PDF generation step will silently fail on first real submission.
 */

// ====== CONFIG ======
const ROOT_FOLDER_ID = '1_QvOlHYVLohFbNH4AyXK72XhuLqt5pGS'; // your Drive folder
const SPREADSHEET_ID = '1UjJISUjAcdJWaLRGWc9gH2QMQ8nEoXFZkX0Il0C60UE'; // your new sheet
const CALLMEBOT_PHONE = '971544564191';
const CALLMEBOT_APIKEY = '2457694';
// =====================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const clientName = (data.fullName || 'Unknown Client').trim();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const safeName = clientName.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s+/g, '_');

    // 1. Create/find per-client dated subfolder in your Drive folder
    const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const folderName = `${safeName}_${today}`;
    const clientFolder = getOrCreateFolder(rootFolder, folderName);

    // 2. Save each uploaded certificate/document, renamed to Name_Date_OriginalFilename
    const savedFiles = [];
    if (data.files && data.files.length) {
      data.files.forEach(f => {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(f.data),
          f.mimeType,
          `${safeName}_${today}_${f.name}`
        );
        const file = clientFolder.createFile(blob);
        savedFiles.push(file.getUrl());
      });
    }

    // 3. Generate a full PDF summary of everything the client submitted
    const roleLabel = formatRole(data);
    const summaryPdfUrl = generateSummaryPDF(data, clientFolder, safeName, today, roleLabel);

    // 4. Log the submission to your spreadsheet
    const sheet = getSheet();
    sheet.appendRow([
      new Date(),
      clientName,
      data.email || '',
      data.phone || '',
      data.location || '',
      formatVisa(data),
      data.noticePeriod || '',
      data.selectedService || '',
      data.selectedPrice || '',
      roleLabel,
      data.summary || '',
      formatJobs(data),
      formatEducation(data),
      data.skills || '',
      clientFolder.getUrl(),
      summaryPdfUrl,
      savedFiles.join(', ')
    ]);

    // 5. WhatsApp notification
    const serviceLine = data.selectedService ? `Service: ${data.selectedService} (${data.selectedPrice || '?'} AED)\n` : '';
    const msg = `New CV intake: ${clientName}\nPhone: ${data.phone || 'n/a'}\n${serviceLine}Visa: ${formatVisa(data)}\nRole: ${roleLabel}\nSummary PDF: ${summaryPdfUrl}\nFiles: ${savedFiles.length}\nFolder: ${clientFolder.getUrl()}`;
    sendWhatsApp(msg);

    return jsonResponse({ status: 'success', folder: clientFolder.getUrl(), summaryPdf: summaryPdfUrl });

  } catch (err) {
    try { sendWhatsApp(`CV intake ERROR: ${err.message}`); } catch (e2) {}
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function getOrCreateFolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Name', 'Email', 'Phone', 'Location', 'Visa Status', 'Notice Period',
      'Selected Service', 'Service Price (AED)', 'Target Role', 'Summary', 'Jobs', 'Education',
      'Skills', 'Drive Folder', 'Summary PDF', 'File Links'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatVisa(data) {
  if (data.visaStatus === 'Other' && data.visaOther) return data.visaOther;
  return data.visaStatus || '';
}

function formatRole(data) {
  const parts = [];
  if (data.industry === 'Other' || data.sector === 'Other' || data.specificRole === 'Other') {
    if (data.roleOther) parts.push(data.roleOther);
  }
  if (data.industry && data.industry !== 'Other') parts.push(data.industry);
  if (data.sector && data.sector !== 'Other') parts.push(data.sector);
  if (data.specificRole && data.specificRole !== 'Other') parts.push(data.specificRole);
  return parts.join(' > ');
}

function formatJobs(data) {
  const jobs = [];
  for (let i = 1; i <= 6; i++) {
    const title = data[`jobTitle_${i}`];
    if (!title) continue;
    jobs.push(`${title} @ ${data[`company_${i}`] || ''} (${data[`startDate_${i}`] || ''} - ${data[`endDate_${i}`] || 'Present'}): ${data[`description_${i}`] || ''}`);
  }
  return jobs.join(' | ');
}

function formatEducation(data) {
  const edu = [];
  let i = 1;
  while (data[`degree_${i}`] !== undefined) {
    if (data[`degree_${i}`]) {
      edu.push(`${data[`degree_${i}`]} - ${data[`school_${i}`] || ''} (${data[`gradYear_${i}`] || ''})`);
    }
    i++;
  }
  return edu.join(' | ');
}

/**
 * Builds a formatted Google Doc with every field the client submitted,
 * exports it as a PDF into the client's Drive folder, then deletes the
 * temporary Doc so only the PDF remains.
 */
function generateSummaryPDF(data, clientFolder, safeName, today, roleLabel) {
  const clientName = (data.fullName || 'Unknown Client').trim();
  const doc = DocumentApp.create(`TEMP_${safeName}_${today}_Summary`);
  const body = doc.getBody();

  body.appendParagraph('CV Intake Summary').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(`Generated: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a")}`)
    .setForegroundColor('#666666');

  addSectionHeading(body, 'Personal Details');
  addLine(body, 'Full Name', clientName);
  addLine(body, 'Email', data.email);
  addLine(body, 'Phone / WhatsApp', data.phone);
  addLine(body, 'Location', data.location);
  addLine(body, 'LinkedIn', data.linkedin);

  addSectionHeading(body, 'Visa & Availability');
  addLine(body, 'Visa Status', formatVisa(data));
  addLine(body, 'Notice Period', data.noticePeriod);

  if (data.selectedService) {
    addSectionHeading(body, 'Selected Service');
    addLine(body, 'Service', data.selectedService);
    addLine(body, 'Price', data.selectedPrice ? `${data.selectedPrice} AED` : '');
  }

  addSectionHeading(body, 'Target Role');
  addLine(body, 'Category', roleLabel);
  addLine(body, 'Professional Summary', data.summary);

  addSectionHeading(body, 'Work Experience');
  let hasJob = false;
  for (let i = 1; i <= 6; i++) {
    const title = data[`jobTitle_${i}`];
    if (!title) continue;
    hasJob = true;
    body.appendParagraph(`${title} — ${data[`company_${i}`] || ''}`).setBold(true).setSpacingBefore(8);
    body.appendParagraph(`${data[`startDate_${i}`] || ''} to ${data[`endDate_${i}`] || 'Present'}`).setForegroundColor('#666666').setItalic(true);
    if (data[`description_${i}`]) body.appendParagraph(data[`description_${i}`]);
  }
  if (!hasJob) body.appendParagraph('No work experience provided.').setForegroundColor('#888888');

  addSectionHeading(body, 'Education & Certifications');
  let hasEdu = false;
  let j = 1;
  while (data[`degree_${j}`] !== undefined) {
    if (data[`degree_${j}`]) {
      hasEdu = true;
      body.appendParagraph(`${data[`degree_${j}`]} — ${data[`school_${j}`] || ''} (${data[`gradYear_${j}`] || ''})`);
    }
    j++;
  }
  if (!hasEdu) body.appendParagraph('No education entries provided.').setForegroundColor('#888888');

  addSectionHeading(body, 'Skills');
  body.appendParagraph(data.skills || 'None listed.');

  addSectionHeading(body, 'Attached Documents');
  const fileCount = data.files ? data.files.length : 0;
  body.appendParagraph(fileCount ? `${fileCount} file(s) uploaded — see client Drive folder.` : 'No files uploaded.');

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  pdfBlob.setName(`${safeName}_${today}_Summary.pdf`);
  const pdfFile = clientFolder.createFile(pdfBlob);

  // Clean up the temporary Google Doc, keep only the PDF
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return pdfFile.getUrl();
}

function addSectionHeading(body, text) {
  body.appendParagraph(text).setHeading(DocumentApp.ParagraphHeading.HEADING2).setSpacingBefore(14);
}

function addLine(body, label, value) {
  if (!value) return;
  const p = body.appendParagraph('');
  p.appendText(`${label}: `).setBold(true);
  p.appendText(String(value));
}

function sendWhatsApp(text) {
  const url = `https://api.callmebot.com/whatsapp.php?phone=${CALLMEBOT_PHONE}&text=${encodeURIComponent(text)}&apikey=${CALLMEBOT_APIKEY}`;
  UrlFetchApp.fetch(url, { muteHttpExceptions: true });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this ONCE manually from the Apps Script editor before your first
 * real submission. It forces Google to prompt you for Docs/Drive/Sheets
 * authorization — without it, PDF generation will fail silently.
 */
function authTest() {
  const doc = DocumentApp.create('auth_test_delete_me');
  doc.getBody().appendParagraph('test');
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0].getName();
  Logger.log('Authorization OK.');
}
