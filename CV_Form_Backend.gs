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

    // 3. Generate a full data summary PDF (for your records) + a formatted CV draft PDF (client-ready)
    const roleLabel = formatRole(data);
    const summaryPdfUrl = generateSummaryPDF(data, clientFolder, safeName, today, roleLabel);
    const cvDraftUrl = generateCVDraft(data, clientFolder, safeName, today, roleLabel);
    const wantsLetter = data.wantsCoverLetter === 'Yes';
    const coverLetterUrl = wantsLetter ? generateCoverLetter(data, clientFolder, safeName, today, roleLabel) : '';

    // 4. Log the submission to your spreadsheet
    const sheet = getSheet();
    const rushLabel = data.isRush === 'Yes' ? `Yes (+${data.rushFee || '?'} AED)` : 'No';
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
      rushLabel,
      roleLabel,
      data.summary || '',
      formatJobs(data),
      formatEducation(data),
      data.skills || '',
      wantsLetter ? 'Yes' : 'No',
      clientFolder.getUrl(),
      summaryPdfUrl,
      cvDraftUrl,
      coverLetterUrl,
      savedFiles.join(', ')
    ]);

    // 5. WhatsApp notification (to you)
    const serviceLine = data.selectedService ? `Service: ${data.selectedService} (${data.selectedPrice || '?'} AED)\n` : '';
    const rushLine = data.isRush === 'Yes' ? `⚡ URGENT REQUEST (+${data.rushFee || '?'} AED)\n` : '';
    const letterLine = wantsLetter ? `Cover Letter: ${coverLetterUrl}\n` : '';
    const msg = `New CV intake: ${clientName}\nPhone: ${data.phone || 'n/a'}\n${serviceLine}${rushLine}Visa: ${formatVisa(data)}\nRole: ${roleLabel}\nCV Draft: ${cvDraftUrl}\n${letterLine}Summary PDF: ${summaryPdfUrl}\nFiles: ${savedFiles.length}\nFolder: ${clientFolder.getUrl()}`;
    sendWhatsApp(msg);

    // 6. Email confirmation (to the client)
    if (data.email) {
      try { sendClientConfirmationEmail(data, clientName, wantsLetter); } catch (emailErr) {
        sendWhatsApp(`Email confirmation FAILED for ${clientName}: ${emailErr.message}`);
      }
    }

    return jsonResponse({ status: 'success', folder: clientFolder.getUrl(), summaryPdf: summaryPdfUrl, cvDraft: cvDraftUrl, coverLetter: coverLetterUrl });

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
      'Selected Service', 'Service Price (AED)', 'Rush Request', 'Target Role', 'Summary', 'Jobs', 'Education',
      'Skills', 'Wants Cover Letter', 'Drive Folder', 'Summary PDF', 'CV Draft PDF', 'Cover Letter PDF', 'File Links'
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

/**
 * ============================================================
 * CV TEMPLATE SYSTEM
 * Six selectable layouts, built from shared content helpers so
 * each template only needs to define arrangement, not re-parse data.
 * ============================================================
 */

function extractJobs(data) {
  const jobs = [];
  for (let i = 1; i <= 6; i++) {
    const title = data[`jobTitle_${i}`];
    if (!title) continue;
    jobs.push({
      title: title,
      company: data[`company_${i}`] || '',
      start: data[`startDate_${i}`] || '',
      end: data[`endDate_${i}`] || 'Present',
      bullets: (data[`description_${i}`] || '').split(/\r?\n/).map(l => l.trim()).filter(Boolean)
    });
  }
  return jobs;
}

function extractEducation(data) {
  const edu = [];
  let j = 1;
  while (data[`degree_${j}`] !== undefined) {
    if (data[`degree_${j}`]) {
      edu.push({ degree: data[`degree_${j}`], school: data[`school_${j}`] || '', year: data[`gradYear_${j}`] || '' });
    }
    j++;
  }
  return edu;
}

function extractSkillsList(data) {
  return (data.skills || '').split(',').map(s => s.trim()).filter(Boolean);
}

/** Strips the stray empty first paragraph Google Docs auto-creates in a new Body. */
function stripLeadingBlankParagraph(body) {
  try {
    const first = body.getChild(0);
    if (first.getType() === DocumentApp.ElementType.PARAGRAPH && first.asParagraph().getText() === '') {
      body.removeChild(first);
    }
  } catch (e) { /* ignore */ }
}

/** Writes a section heading into any container (Body or TableCell). */
function writeHeading(container, text, opts) {
  opts = opts || {};
  const p = container.appendParagraph(text);
  p.setBold(true)
   .setFontSize(opts.size || 11.5)
   .setForegroundColor(opts.color || '#12212e')
   .setSpacingBefore(opts.spaceBefore != null ? opts.spaceBefore : 10)
   .setSpacingAfter(opts.spaceAfter != null ? opts.spaceAfter : 4);
  if (opts.align) p.setAlignment(opts.align);
  return p;
}

/** Writes work experience into any container. */
function writeJobsInto(container, jobs, opts) {
  opts = opts || {};
  const textColor = opts.textColor || '#12212e';
  const mutedColor = opts.mutedColor || '#666666';
  jobs.forEach(job => {
    const titleLine = container.appendParagraph('');
    titleLine.appendText(job.title).setBold(true).setFontSize(opts.titleSize || 11).setForegroundColor(textColor);
    if (job.company) titleLine.appendText('  —  ' + job.company).setFontSize(opts.titleSize || 11).setBold(false).setForegroundColor(textColor);
    titleLine.setSpacingBefore(8).setSpacingAfter(1);

    const dateLine = container.appendParagraph(`${job.start} – ${job.end}`);
    dateLine.setItalic(true).setFontSize(9.5).setForegroundColor(mutedColor).setSpacingAfter(3);

    job.bullets.forEach(line => {
      const bullet = container.appendParagraph(line.replace(/^[-•]\s*/, ''));
      bullet.setFontSize(opts.bodySize || 10.5).setSpacingAfter(2).setForegroundColor(textColor);
      bullet.setGlyphType(DocumentApp.GlyphType.BULLET);
    });
  });
}

/** Writes education entries into any container. */
function writeEducationInto(container, edu, opts) {
  opts = opts || {};
  const textColor = opts.textColor || '#12212e';
  edu.forEach(item => {
    const line = container.appendParagraph('');
    line.appendText(item.degree).setBold(true).setFontSize(opts.size || 10.5).setForegroundColor(textColor);
    const rest = [item.school, item.year].filter(Boolean).join(', ');
    if (rest) line.appendText(' — ' + rest).setBold(false).setFontSize(opts.size || 10.5).setForegroundColor(textColor);
    line.setSpacingAfter(4);
  });
}

/** Writes skills — either inline comma text or one-per-line (for narrow sidebars). */
function writeSkillsInto(container, skillsList, opts) {
  opts = opts || {};
  const textColor = opts.textColor || '#12212e';
  if (opts.mode === 'list') {
    skillsList.forEach(skill => {
      container.appendParagraph(skill).setFontSize(opts.size || 9.5).setForegroundColor(textColor).setSpacingAfter(3);
    });
  } else {
    container.appendParagraph(skillsList.join('  •  ')).setFontSize(opts.size || 10.5).setForegroundColor(textColor).setSpacingAfter(10);
  }
}

function writeReferencesInto(container, opts) {
  opts = opts || {};
  container.appendParagraph('Available upon request.')
    .setFontSize(10.5).setItalic(true).setForegroundColor(opts.mutedColor || '#666666');
}

/**
 * Dispatcher — picks the right layout based on data.cvTemplate,
 * exports as PDF into the client's Drive folder.
 */
function generateCVDraft(data, clientFolder, safeName, today, roleLabel) {
  const clientName = (data.fullName || 'Unknown Client').trim();
  const template = data.cvTemplate || 'ats-classic';
  const jobs = extractJobs(data);
  const edu = extractEducation(data);
  const skillsList = extractSkillsList(data);

  const doc = DocumentApp.create(`TEMP_${safeName}_${today}_CV_Draft`);
  const body = doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(55).setMarginRight(55);

  const ctx = { data, clientName, roleLabel, jobs, edu, skillsList };

  switch (template) {
    case 'ats-minimal':        buildAtsMinimal(body, ctx); break;
    case 'navy-sidebar':       body.setMarginLeft(0).setMarginRight(0).setMarginTop(0); buildNavySidebar(body, ctx); break;
    case 'bold-header':        body.setMarginTop(0); buildBoldHeader(body, ctx); break;
    case 'minimalist-elegant': buildMinimalistElegant(body, ctx); break;
    case 'two-column':         buildTwoColumn(body, ctx); break;
    case 'ats-classic':
    default:                   buildAtsClassic(body, ctx); break;
  }

  stripLeadingBlankParagraph(body);
  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  pdfBlob.setName(`${safeName}_${today}_CV_Draft_${template}.pdf`);
  const pdfFile = clientFolder.createFile(pdfBlob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return pdfFile.getUrl();
}

/** 1. ATS Classic — centered header, single column, light brand blue accents. */
function buildAtsClassic(body, ctx) {
  const { data, clientName, roleLabel, jobs, edu, skillsList } = ctx;

  const namePara = body.appendParagraph(clientName.toUpperCase());
  namePara.setFontSize(22).setBold(true).setForegroundColor('#12212e').setSpacingAfter(2);
  namePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const roleTitle = data.roleOther || (roleLabel ? roleLabel.split(' > ').pop() : '') || data.selectedService;
  if (roleTitle) {
    body.appendParagraph(roleTitle).setFontSize(13).setBold(true).setForegroundColor('#3f7dfa')
      .setSpacingAfter(6).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  const contactParts = [data.phone, data.email, data.location, data.linkedin].filter(Boolean);
  if (contactParts.length) {
    body.appendParagraph(contactParts.join('   |   ')).setFontSize(9.5).setForegroundColor('#555555')
      .setSpacingAfter(10).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
  body.appendHorizontalRule();

  if (data.summary) {
    writeHeading(body, 'PROFESSIONAL SUMMARY');
    body.appendParagraph(data.summary).setFontSize(10.5).setSpacingAfter(10);
  }
  if (skillsList.length) {
    writeHeading(body, 'KEY SKILLS');
    writeSkillsInto(body, skillsList, { mode: 'inline' });
  }
  if (jobs.length) {
    writeHeading(body, 'WORK EXPERIENCE');
    writeJobsInto(body, jobs);
  }
  if (edu.length) {
    writeHeading(body, 'EDUCATION & CERTIFICATIONS');
    writeEducationInto(body, edu);
  }
  writeHeading(body, 'REFERENCES');
  writeReferencesInto(body);
  appendVisaFootnote(body, data);
}

/** 2. ATS Minimal — left-aligned, monochrome, zero graphical risk. */
function buildAtsMinimal(body, ctx) {
  const { data, clientName, roleLabel, jobs, edu, skillsList } = ctx;

  body.appendParagraph(clientName.toUpperCase()).setFontSize(18).setBold(true).setForegroundColor('#000000').setSpacingAfter(2);

  const roleTitle = data.roleOther || (roleLabel ? roleLabel.split(' > ').pop() : '') || data.selectedService;
  if (roleTitle) body.appendParagraph(roleTitle).setFontSize(11.5).setForegroundColor('#333333').setSpacingAfter(4);

  const contactParts = [data.phone, data.email, data.location, data.linkedin].filter(Boolean);
  if (contactParts.length) body.appendParagraph(contactParts.join('  |  ')).setFontSize(9).setForegroundColor('#444444').setSpacingAfter(12);

  if (data.summary) {
    writeHeading(body, 'SUMMARY', { color: '#000000', size: 10.5 });
    body.appendParagraph(data.summary).setFontSize(10).setForegroundColor('#111111').setSpacingAfter(10);
  }
  if (skillsList.length) {
    writeHeading(body, 'SKILLS', { color: '#000000', size: 10.5 });
    writeSkillsInto(body, skillsList, { mode: 'inline', textColor: '#111111' });
  }
  if (jobs.length) {
    writeHeading(body, 'EXPERIENCE', { color: '#000000', size: 10.5 });
    writeJobsInto(body, jobs, { textColor: '#111111', mutedColor: '#555555' });
  }
  if (edu.length) {
    writeHeading(body, 'EDUCATION', { color: '#000000', size: 10.5 });
    writeEducationInto(body, edu, { textColor: '#111111' });
  }
  writeHeading(body, 'REFERENCES', { color: '#000000', size: 10.5 });
  writeReferencesInto(body, { mutedColor: '#555555' });
  appendVisaFootnote(body, data, { plain: true });
}

/** 3. Navy Sidebar — dark left column (contact + skills), white right column (everything else). */
function buildNavySidebar(body, ctx) {
  const { data, clientName, roleLabel, jobs, edu, skillsList } = ctx;

  const table = body.appendTable([['', '']]);
  table.setBorderWidth(0);
  const leftCell = table.getCell(0, 0);
  const rightCell = table.getCell(0, 1);
  leftCell.setBackgroundColor('#12212e');
  leftCell.setWidth(165);
  rightCell.setWidth(330);
  leftCell.setPaddingTop(20).setPaddingLeft(16).setPaddingRight(16).setPaddingBottom(20);
  rightCell.setPaddingTop(20).setPaddingLeft(20).setPaddingRight(16).setPaddingBottom(20);

  // Left column
  const namePara = leftCell.getChild(0).asParagraph();
  namePara.setText(clientName).setBold(true).setFontSize(15).setForegroundColor('#ffffff').setSpacingAfter(10);

  writeHeading(leftCell, 'CONTACT', { color: '#8bb4ff', size: 10 });
  [data.phone, data.email, data.location, data.linkedin].filter(Boolean).forEach(line => {
    leftCell.appendParagraph(line).setFontSize(9).setForegroundColor('#e8f0ff').setSpacingAfter(3);
  });

  if (skillsList.length) {
    writeHeading(leftCell, 'SKILLS', { color: '#8bb4ff', size: 10 });
    writeSkillsInto(leftCell, skillsList, { mode: 'list', textColor: '#e8f0ff', size: 9 });
  }

  writeHeading(leftCell, 'REFERENCES', { color: '#8bb4ff', size: 10 });
  leftCell.appendParagraph('Available upon request.').setFontSize(8.5).setItalic(true).setForegroundColor('#b9c9e8');

  // Right column
  const roleTitle = data.roleOther || (roleLabel ? roleLabel.split(' > ').pop() : '') || data.selectedService;
  const rp = rightCell.getChild(0).asParagraph();
  rp.setText(roleTitle || 'Professional').setBold(true).setFontSize(13).setForegroundColor('#3f7dfa').setSpacingAfter(10);

  if (data.summary) {
    writeHeading(rightCell, 'PROFESSIONAL SUMMARY', { size: 11 });
    rightCell.appendParagraph(data.summary).setFontSize(10).setSpacingAfter(10);
  }
  if (jobs.length) {
    writeHeading(rightCell, 'WORK EXPERIENCE', { size: 11 });
    writeJobsInto(rightCell, jobs, { bodySize: 9.7 });
  }
  if (edu.length) {
    writeHeading(rightCell, 'EDUCATION & CERTIFICATIONS', { size: 11 });
    writeEducationInto(rightCell, edu, { size: 9.7 });
  }
  const visa = formatVisa(data);
  if (visa || data.noticePeriod) {
    const availPara = rightCell.appendParagraph('');
    availPara.appendText('Visa: ').setBold(true).setFontSize(8.5);
    availPara.appendText(visa || 'N/A').setFontSize(8.5);
    if (data.noticePeriod) {
      availPara.appendText('  |  Availability: ').setBold(true).setFontSize(8.5);
      availPara.appendText(data.noticePeriod).setFontSize(8.5);
    }
    availPara.setForegroundColor('#666666').setSpacingBefore(10);
  }
}

/** 4. Bold Header — full-width gradient-style color band, single column below. */
function buildBoldHeader(body, ctx) {
  const { data, clientName, roleLabel, jobs, edu, skillsList } = ctx;

  const table = body.appendTable([['']]);
  table.setBorderWidth(0);
  const cell = table.getCell(0, 0);
  cell.setBackgroundColor('#3f7dfa');
  cell.setPaddingTop(28).setPaddingBottom(28).setPaddingLeft(40).setPaddingRight(40);
  const namePara = cell.getChild(0).asParagraph();
  namePara.setText(clientName.toUpperCase()).setBold(true).setFontSize(22).setForegroundColor('#ffffff').setSpacingAfter(4);

  const roleTitle = data.roleOther || (roleLabel ? roleLabel.split(' > ').pop() : '') || data.selectedService;
  if (roleTitle) cell.appendParagraph(roleTitle).setFontSize(12).setForegroundColor('#e8f0ff').setSpacingAfter(4);

  const contactParts = [data.phone, data.email, data.location, data.linkedin].filter(Boolean);
  if (contactParts.length) cell.appendParagraph(contactParts.join('   |   ')).setFontSize(9).setForegroundColor('#dce9ff');

  if (data.summary) {
    writeHeading(body, 'PROFESSIONAL SUMMARY', { spaceBefore: 20 });
    body.appendParagraph(data.summary).setFontSize(10.5).setSpacingAfter(10);
  }
  if (skillsList.length) {
    writeHeading(body, 'KEY SKILLS');
    writeSkillsInto(body, skillsList, { mode: 'inline' });
  }
  if (jobs.length) {
    writeHeading(body, 'WORK EXPERIENCE');
    writeJobsInto(body, jobs);
  }
  if (edu.length) {
    writeHeading(body, 'EDUCATION & CERTIFICATIONS');
    writeEducationInto(body, edu);
  }
  writeHeading(body, 'REFERENCES');
  writeReferencesInto(body);
  appendVisaFootnote(body, data);
}

/** 5. Minimalist Elegant — airy whitespace, thin dividers, refined serif-style headers. */
function buildMinimalistElegant(body, ctx) {
  const { data, clientName, roleLabel, jobs, edu, skillsList } = ctx;

  const namePara = body.appendParagraph(clientName);
  namePara.setFontSize(24).setBold(false).setForegroundColor('#12212e').setSpacingAfter(2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const roleTitle = data.roleOther || (roleLabel ? roleLabel.split(' > ').pop() : '') || data.selectedService;
  if (roleTitle) {
    body.appendParagraph(roleTitle.toUpperCase()).setFontSize(9.5).setForegroundColor('#a8752f')
      .setSpacingAfter(8).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
  const contactParts = [data.phone, data.email, data.location, data.linkedin].filter(Boolean);
  if (contactParts.length) {
    body.appendParagraph(contactParts.join('    ·    ')).setFontSize(9).setForegroundColor('#8a8f9c')
      .setSpacingAfter(18).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  function thinDivider() {
    body.appendParagraph('—').setFontSize(9).setForegroundColor('#c9ced8')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(14);
  }

  if (data.summary) {
    writeHeading(body, 'Profile', { size: 10.5, color: '#12212e', align: DocumentApp.HorizontalAlignment.CENTER });
    body.appendParagraph(data.summary).setFontSize(10.5).setSpacingAfter(14).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    thinDivider();
  }
  if (jobs.length) {
    writeHeading(body, 'Experience', { size: 10.5, align: DocumentApp.HorizontalAlignment.CENTER, spaceAfter: 8 });
    writeJobsInto(body, jobs);
    thinDivider();
  }
  if (edu.length) {
    writeHeading(body, 'Education', { size: 10.5, align: DocumentApp.HorizontalAlignment.CENTER, spaceAfter: 8 });
    writeEducationInto(body, edu);
    thinDivider();
  }
  if (skillsList.length) {
    writeHeading(body, 'Skills', { size: 10.5, align: DocumentApp.HorizontalAlignment.CENTER, spaceAfter: 8 });
    writeSkillsInto(body, skillsList, { mode: 'inline' });
    thinDivider();
  }
  writeHeading(body, 'References', { size: 10.5, align: DocumentApp.HorizontalAlignment.CENTER, spaceAfter: 4 });
  writeReferencesInto(body);
  appendVisaFootnote(body, data);
}

/** 6. Two-Column Compact — table-based split body for a denser magazine feel. */
function buildTwoColumn(body, ctx) {
  const { data, clientName, roleLabel, jobs, edu, skillsList } = ctx;

  const namePara = body.appendParagraph(clientName.toUpperCase());
  namePara.setFontSize(20).setBold(true).setForegroundColor('#12212e').setSpacingAfter(2);
  const roleTitle = data.roleOther || (roleLabel ? roleLabel.split(' > ').pop() : '') || data.selectedService;
  if (roleTitle) body.appendParagraph(roleTitle).setFontSize(12).setBold(true).setForegroundColor('#3f7dfa').setSpacingAfter(4);
  const contactParts = [data.phone, data.email, data.location, data.linkedin].filter(Boolean);
  if (contactParts.length) body.appendParagraph(contactParts.join('   |   ')).setFontSize(9).setForegroundColor('#555555').setSpacingAfter(10);
  body.appendHorizontalRule();

  const table = body.appendTable([['', '']]);
  table.setBorderWidth(0);
  const leftCell = table.getCell(0, 0);
  const rightCell = table.getCell(0, 1);
  leftCell.setWidth(190);
  rightCell.setWidth(285);
  leftCell.setPaddingRight(14);
  rightCell.setPaddingLeft(14);

  const leftFirst = leftCell.getChild(0).asParagraph();
  leftFirst.removeFromParent();
  if (data.summary) {
    writeHeading(leftCell, 'SUMMARY', { size: 10.5 });
    leftCell.appendParagraph(data.summary).setFontSize(9.5).setSpacingAfter(10);
  }
  if (skillsList.length) {
    writeHeading(leftCell, 'SKILLS', { size: 10.5 });
    writeSkillsInto(leftCell, skillsList, { mode: 'list', size: 9.5 });
  }
  if (edu.length) {
    writeHeading(leftCell, 'EDUCATION', { size: 10.5 });
    writeEducationInto(leftCell, edu, { size: 9.5 });
  }
  writeHeading(leftCell, 'REFERENCES', { size: 10.5 });
  writeReferencesInto(leftCell);

  const rightFirst = rightCell.getChild(0).asParagraph();
  rightFirst.removeFromParent();
  if (jobs.length) {
    writeHeading(rightCell, 'WORK EXPERIENCE', { size: 10.5 });
    writeJobsInto(rightCell, jobs, { bodySize: 9.7 });
  }
  appendVisaFootnote(rightCell, data, { compact: true });
}

function appendVisaFootnote(container, data, opts) {
  opts = opts || {};
  const visa = formatVisa(data);
  if (!visa && !data.noticePeriod) return;
  if (!opts.compact && container.appendHorizontalRule) {
    try { container.appendHorizontalRule(); } catch (e) { /* table cells can't take rules */ }
  }
  const availPara = container.appendParagraph('');
  availPara.appendText('Visa Status: ').setBold(true).setFontSize(9.5);
  availPara.appendText(visa || 'N/A').setFontSize(9.5);
  if (data.noticePeriod) {
    availPara.appendText('    |    Availability: ').setBold(true).setFontSize(9.5);
    availPara.appendText(data.noticePeriod).setFontSize(9.5);
  }
  availPara.setForegroundColor(opts.plain ? '#444444' : '#666666').setSpacingBefore(8);
}

/**
 * ============================================================
 * APPLICATION / COVER LETTER
 * Standard single-column business-letter format (the correct
 * convention regardless of CV layout), with accent color and
 * tone matched to whichever CV template the client picked.
 * ============================================================
 */
function letterThemeFor(template) {
  const themes = {
    'ats-classic':        { accent: '#3f7dfa', useColor: true,  band: false },
    'ats-minimal':        { accent: '#000000', useColor: false, band: false },
    'navy-sidebar':       { accent: '#12212e', useColor: true,  band: false },
    'bold-header':        { accent: '#3f7dfa', useColor: true,  band: true  },
    'minimalist-elegant': { accent: '#a8752f', useColor: true,  band: false },
    'two-column':         { accent: '#3f7dfa', useColor: true,  band: false }
  };
  return themes[template] || themes['ats-classic'];
}

function naturalJoin(arr) {
  if (!arr.length) return '';
  if (arr.length === 1) return arr[0];
  if (arr.length === 2) return arr.join(' and ');
  return arr.slice(0, -1).join(', ') + ', and ' + arr[arr.length - 1];
}

function generateCoverLetter(data, clientFolder, safeName, today, roleLabel) {
  const clientName = (data.fullName || 'Unknown Client').trim();
  const template = data.cvTemplate || 'ats-classic';
  const theme = letterThemeFor(template);
  const jobs = extractJobs(data);
  const skillsList = extractSkillsList(data);
  const roleTitle = data.roleOther || (roleLabel ? roleLabel.split(' > ').pop() : '') || data.selectedService || 'this role';

  const doc = DocumentApp.create(`TEMP_${safeName}_${today}_Cover_Letter`);
  const body = doc.getBody();
  body.setMarginTop(theme.band ? 0 : 55).setMarginBottom(55).setMarginLeft(65).setMarginRight(65);

  // ---- Header block ----
  if (theme.band) {
    const table = body.appendTable([['']]);
    table.setBorderWidth(0);
    const cell = table.getCell(0, 0);
    cell.setBackgroundColor(theme.accent);
    cell.setPaddingTop(24).setPaddingBottom(24).setPaddingLeft(45).setPaddingRight(45);
    const nameP = cell.getChild(0).asParagraph();
    nameP.setText(clientName).setBold(true).setFontSize(18).setForegroundColor('#ffffff').setSpacingAfter(3);
    const contactParts = [data.phone, data.email, data.location].filter(Boolean);
    if (contactParts.length) cell.appendParagraph(contactParts.join('   |   ')).setFontSize(9).setForegroundColor('#e8f0ff');
    body.appendParagraph('').setSpacingAfter(16);
  } else {
    const nameP = body.appendParagraph(clientName);
    nameP.setBold(true).setFontSize(16).setForegroundColor(theme.useColor ? theme.accent : '#12212e').setSpacingAfter(2);
    const contactParts = [data.phone, data.email, data.location].filter(Boolean);
    if (contactParts.length) body.appendParagraph(contactParts.join('   |   ')).setFontSize(9.5).setForegroundColor('#555555').setSpacingAfter(16);
  }

  // ---- Date ----
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  body.appendParagraph(dateStr).setFontSize(10).setForegroundColor('#555555').setSpacingAfter(16);

  // ---- Salutation ----
  body.appendParagraph('Dear Hiring Manager,').setFontSize(10.5).setSpacingAfter(10);

  // ---- Opening paragraph ----
  const mostRecent = jobs[0];
  const topSkills = skillsList.slice(0, 4);
  let opening = `I am writing to express my interest in the ${roleTitle} position. `;
  if (mostRecent) {
    opening += `With hands-on experience as ${aOrAn(mostRecent.title)} at ${mostRecent.company || 'my current organization'}, `;
  }
  if (topSkills.length) {
    opening += `I bring strong capability in ${naturalJoin(topSkills)}${mostRecent ? ', among other areas,' : ''} that I believe would add real value to your team.`;
  } else {
    opening += `I believe my background makes me a strong fit for this opportunity.`;
  }
  body.appendParagraph(opening).setFontSize(10.5).setSpacingAfter(10);

  // ---- Body paragraph ----
  let bodyPara = '';
  if (data.summary) {
    bodyPara += data.summary + ' ';
  }
  if (mostRecent && mostRecent.bullets.length) {
    bodyPara += `In my most recent role, I ${lowerFirst(mostRecent.bullets[0])}`;
    if (!bodyPara.trim().endsWith('.')) bodyPara += '.';
  }
  if (bodyPara.trim()) {
    body.appendParagraph(bodyPara.trim()).setFontSize(10.5).setSpacingAfter(10);
  }

  // ---- Closing paragraph ----
  const closing = `I would welcome the opportunity to discuss how my experience aligns with your needs. Thank you for considering my application — I look forward to the possibility of contributing to your organization.`;
  body.appendParagraph(closing).setFontSize(10.5).setSpacingAfter(18);

  // ---- Sign-off ----
  body.appendParagraph('Sincerely,').setFontSize(10.5).setSpacingAfter(2);
  body.appendParagraph(clientName).setFontSize(10.5).setBold(true).setForegroundColor(theme.useColor ? theme.accent : '#12212e');

  stripLeadingBlankParagraph(body);
  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  pdfBlob.setName(`${safeName}_${today}_Cover_Letter.pdf`);
  const pdfFile = clientFolder.createFile(pdfBlob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return pdfFile.getUrl();
}

function aOrAn(word) {
  if (!word) return 'a professional';
  const vowels = ['a', 'e', 'i', 'o', 'u'];
  return (vowels.includes(word.trim()[0].toLowerCase()) ? 'an ' : 'a ') + word;
}

function lowerFirst(str) {
  if (!str) return '';
  return str.charAt(0).toLowerCase() + str.slice(1);
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

/**
 * Sends an automatic confirmation email to the client. This is separate
 * from the WhatsApp ping (which goes to you) — this one lands in the
 * client's own inbox immediately after they submit.
 */
function sendClientConfirmationEmail(data, clientName, wantsLetter) {
  const firstName = clientName.split(' ')[0];
  const serviceLine = data.selectedService
    ? `<p><strong>Service selected:</strong> ${data.selectedService}${data.selectedPrice ? ` — ${data.selectedPrice} AED` : ''}</p>`
    : '';
  const rushLine = data.isRush === 'Yes'
    ? `<p><strong>⚡ Urgent request</strong> — includes a +${data.rushFee || ''} AED express fee.</p>`
    : '';
  const letterLine = wantsLetter
    ? `<p>We'll also prepare a matching application/cover letter alongside your CV.</p>`
    : '';

  const subject = 'We received your CV details — KINETIX';
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#12212e;max-width:480px;margin:0 auto;">
      <h2 style="color:#3f7dfa;">Thank you, ${firstName}!</h2>
      <p>We're working on it, and we'll get back to you with feedback as soon as your CV is ready.</p>
      ${serviceLine}
      ${rushLine}
      ${letterLine}
      <p>Standard turnaround is 24–48 hours after payment and details are received. If you haven't sent payment yet, please do so to start work — full payment is required upfront.</p>
      <p>Questions in the meantime? Message us directly on WhatsApp:<br>
      <a href="https://wa.me/971544564191" style="color:#3f7dfa;">+971 54 456 4191</a></p>
      <p style="margin-top:24px;color:#888;font-size:12px;">— KINETIX · Perpetual Momentum. Infinite Impact.</p>
    </div>`;

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    htmlBody: htmlBody
  });
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
