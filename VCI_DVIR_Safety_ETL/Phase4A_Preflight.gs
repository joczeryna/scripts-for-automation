/**
 * VCI DVIR + Safety Deviations consolidated automation
 * Phase 4A: read-only preflight validation.
 *
 * This file does not write destination data, archive source files, or send email.
 */

const PHASE4_INTAKE_FOLDER_ID = '1-pkJhlzx80h2bGsAxPhu_WXkul2UiwQa';
const PHASE4_TIME_ZONE = 'America/Los_Angeles';

function validateAutomationInputs() {
  const results = [];
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFolderById(PHASE4_INTAKE_FOLDER_ID);
    const runDate = Utilities.formatDate(new Date(), PHASE4_TIME_ZONE, 'MM-dd-yyyy');

    const dvirFile = phase4FindUniqueFile_(folder, ['DVIR_Details', 'DVIR_Details.xlsx']);
    const safetyFile = phase4FindUniqueFile_(folder, [
      'safety-deviations-' + runDate,
      'safety-deviations-' + runDate + '.xlsx'
    ]);

    const dvirSource = phase4OpenGoogleSheet_(dvirFile, 'DVIR source');
    const safetySource = phase4OpenGoogleSheet_(safetyFile, 'Safety Deviations source');

    const dvirSheet = phase4RequireSheet_(dvirSource, 'DVIR_Details');
    phase4ValidateDvir_(dvirSheet);
    results.push('DVIR source: PASS (' + dvirSheet.getLastRow() + ' rows, ' +
      dvirSheet.getLastColumn() + ' columns)');

    const safetySheet = phase4RequireSheet_(safetySource, 'Export');
    phase4ValidateSafety_(safetySheet);
    results.push('Safety Deviations source: PASS (' + safetySheet.getLastRow() +
      ' rows, ' + safetySheet.getLastColumn() + ' columns)');

    const issues = phase4RequireSheet_(active, 'Issues');
    const gateChecks = phase4RequireSheet_(active, 'Gate Checks');
    const jso = phase4RequireSheet_(active, 'JSO');
    const closed = phase4RequireSheet_(active, 'Closed');
    const equipment = phase4RequireSheet_(active, 'Equipment');
    if (equipment.getMaxColumns() < 9) {
      throw new Error('Equipment must contain at least columns A through I.');
    }
    results.push('Operational tabs: PASS (Issues GID ' + issues.getSheetId() +
      ', Gate Checks GID ' + gateChecks.getSheetId() + ', JSO GID ' +
      jso.getSheetId() + ', Closed GID ' + closed.getSheetId() + ')');

    const recipientIndex = phase4RequireSheet_(active, 'Automation Recipient Index');
    phase4RequireHeaders_(recipientIndex, 1, [
      'Source Type', 'Audit ID', 'Worker No', 'Assigned Driver',
      'Match Source', 'Last Updated', 'Match Status', 'Notes'
    ]);
    const protections = recipientIndex.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (protections.length === 0) {
      throw new Error('Automation Recipient Index exists but is not protected.');
    }
    if (!protections.some(function(protection) { return protection.canEdit(); })) {
      throw new Error('The executing Google account cannot edit Automation Recipient Index.');
    }
    results.push('Automation Recipient Index: PASS (protected and editable)');

    const directory = phase4RequireSheet_(active, 'Company Directory');
    phase4RequireHeaderSet_(directory, 1, [
      'Name', 'Worker No', 'Employee AD', 'Email Address',
      'Worker Status', 'Supervisor'
    ]);
    results.push('Company Directory: PASS (GID ' + directory.getSheetId() + ')');

    const message = 'Phase 4A preflight passed.\n\n' + results.join('\n');
    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);
    return results;
  } catch (error) {
    const message = 'Phase 4A preflight failed:\n\n' + error.message;
    Logger.log(message + '\n' + (error.stack || ''));
    SpreadsheetApp.getUi().alert(message);
    throw error;
  }
}

function phase4FindUniqueFile_(folder, acceptedNames) {
  const accepted = {};
  acceptedNames.forEach(function(name) { accepted[name.toLowerCase()] = true; });
  const matches = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (accepted[file.getName().toLowerCase()]) matches.push(file);
  }
  if (matches.length === 0) {
    throw new Error('Missing intake file. Expected one of: ' + acceptedNames.join(', '));
  }
  if (matches.length > 1) {
    throw new Error('Duplicate intake files found for ' + acceptedNames[0] +
      '. Remove duplicates before processing. Matching file IDs: ' +
      matches.map(function(file) { return file.getId(); }).join(', '));
  }
  return matches[0];
}

function phase4OpenGoogleSheet_(file, label) {
  if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    throw new Error(label + ' is not a Google Sheet after Drive upload: ' +
      file.getName() + ' (' + file.getMimeType() + ').');
  }
  return SpreadsheetApp.openById(file.getId());
}

function phase4RequireSheet_(spreadsheet, name) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error('Required worksheet not found: ' + name);
  return sheet;
}

function phase4ValidateDvir_(sheet) {
  if (sheet.getLastRow() < 3 || sheet.getLastColumn() < 28) {
    throw new Error('DVIR_Details contains no usable report rows.');
  }
  const width = Math.max(29, sheet.getLastColumn());
  const values = sheet.getRange(3, 1, sheet.getLastRow() - 2, width).getDisplayValues();
  const usable = values.some(function(row) {
    return phase4Normalize_(row[0]) === 'vci' && String(row[4] || '').trim() !== '';
  });
  if (!usable) throw new Error('DVIR_Details has no VCI row with a populated Audit ID.');
}

function phase4ValidateSafety_(sheet) {
  if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 15) {
    throw new Error('Safety Deviations Export contains no usable report rows.');
  }
  phase4RequireHeaderSet_(sheet, 1, [
    'Subsidiary', 'Worker Nbr', 'Auditid', 'Audit Name',
    'Deviation Question', 'Deviation Status'
  ]);
}

function phase4RequireHeaders_(sheet, rowNumber, expected) {
  const actual = sheet.getRange(rowNumber, 1, 1, expected.length).getDisplayValues()[0];
  for (let index = 0; index < expected.length; index++) {
    if (phase4Normalize_(actual[index]) !== phase4Normalize_(expected[index])) {
      throw new Error(sheet.getName() + ' header mismatch in column ' + (index + 1) +
        '. Expected "' + expected[index] + '"; found "' + actual[index] + '".');
    }
  }
}

function phase4RequireHeaderSet_(sheet, rowNumber, expected) {
  const actual = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(phase4Normalize_);
  const missing = expected.filter(function(header) {
    return actual.indexOf(phase4Normalize_(header)) === -1;
  });
  if (missing.length > 0) {
    throw new Error(sheet.getName() + ' is missing required headers: ' + missing.join(', '));
  }
}

function phase4Normalize_(value) {
  return String(value == null ? '' : value).trim().toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}
