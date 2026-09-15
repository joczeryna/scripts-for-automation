/**
 * Phase 4G: establish the historical email baseline.
 *
 * One-time controlled action:
 * - validates the workbook before writing;
 * - creates a complete spreadsheet backup;
 * - stamps today's date into Closed column P only where Audit ID exists and
 *   Emailed Date is blank;
 * - preserves every existing Emailed Date;
 * - verifies the result;
 * - sends no emails and archives no source files.
 *
 * Requires validateWorkbookIntegrity() from Phase 4D/4E.
 */

const PHASE4G_VALIDATION_ERROR_EMAIL = 'jo.deleon@vcicom.com';

function establishHistoricalEmailBaseline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preflight = validateWorkbookIntegrity('EMAIL_BASELINE_PREFLIGHT');
  const validation = preflight.validation || preflight;

  if (validation.status === 'BLOCKED' || (validation.blockers || []).length) {
    throw new Error(
      'Email baseline was not written because workbook validation is BLOCKED. ' +
      'No email was sent. Review the execution log.'
    );
  }

  const closed = ss.getSheetByName('Closed');
  if (!closed) throw new Error('Required worksheet "Closed" was not found.');
  if (closed.getLastColumn() < 16) {
    throw new Error('Closed must contain Emailed Date in column P.');
  }

  const lastRow = closed.getLastRow();
  if (lastRow < 2) {
    throw new Error('Closed contains no data rows; no baseline was written.');
  }

  const backup = ss.copy(
    ss.getName() + ' - Phase4G Email Baseline Backup - ' +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmmss')
  );

  const rowCount = lastRow - 1;
  const auditIds = closed.getRange(2, 2, rowCount, 1).getDisplayValues();
  const emailedRange = closed.getRange(2, 16, rowCount, 1);
  const emailedValues = emailedRange.getValues();
  const baselineDate = phase4gTodayAtNoon_(ss.getSpreadsheetTimeZone());

  let stamped = 0;
  let preserved = 0;
  let ignoredBlankAuditRows = 0;

  for (let index = 0; index < rowCount; index++) {
    const auditId = String(auditIds[index][0] || '').trim();
    const emailedDate = emailedValues[index][0];
    if (!auditId) {
      ignoredBlankAuditRows++;
      continue;
    }
    if (emailedDate !== '' && emailedDate !== null) {
      preserved++;
      continue;
    }
    emailedValues[index][0] = baselineDate;
    stamped++;
  }

  if (stamped) {
    emailedRange.setValues(emailedValues);
    emailedRange.setNumberFormat('MM/dd/yyyy');
    SpreadsheetApp.flush();
  }

  const verification = closed.getRange(2, 2, rowCount, 15).getValues();
  const remaining = [];
  verification.forEach(function(row, index) {
    const auditId = String(row[0] || '').trim();
    const emailedDate = row[14];
    if (auditId && (emailedDate === '' || emailedDate === null)) {
      if (remaining.length < 20) remaining.push({ row: index + 2, auditId: auditId });
    }
  });
  if (remaining.length) {
    throw new Error(
      'Baseline verification failed; eligible Closed rows still have blank Emailed Date: ' +
      JSON.stringify(remaining)
    );
  }

  const result = {
    mode: 'ONE-TIME EMAIL BASELINE - NO EMAILS OR ARCHIVING',
    baselineDate: Utilities.formatDate(
      baselineDate, ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy'
    ),
    validationStatus: validation.status,
    backupId: backup.getId(),
    backupName: backup.getName(),
    closedRowsReviewed: rowCount,
    placeholderDatesWritten: stamped,
    existingEmailedDatesPreserved: preserved,
    blankAuditIdRowsIgnored: ignoredBlankAuditRows,
    eligibleBlankEmailedDatesRemaining: 0,
    emailsSent: false,
    sourcesArchived: false,
    futurePolicy: [
      'Only closures detected after this baseline may become email candidates.',
      'Operational email To is Assigned Driver only.',
      'Operational email CC is the Assigned Driver supervisor only.',
      'No hard-coded operational recipients are permitted.',
      'Validation and processing error notifications go only to ' +
        PHASE4G_VALIDATION_ERROR_EMAIL + '.',
      'Emailed Date is written for future records only after successful delivery.'
    ]
  };

  Logger.log('PHASE4G_BASELINE_RESULT=' + JSON.stringify(result));
  SpreadsheetApp.getUi().alert([
    'Historical email baseline completed.',
    '',
    'Placeholder date: ' + result.baselineDate,
    'Dates written: ' + stamped,
    'Existing emailed dates preserved: ' + preserved,
    'Remaining eligible blanks: 0',
    '',
    'No emails were sent. No source files were archived.',
    'Backup: ' + result.backupName
  ].join('\n'));
  return result;
}

function previewHistoricalEmailBaseline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const closed = ss.getSheetByName('Closed');
  if (!closed || closed.getLastRow() < 2) {
    throw new Error('Closed contains no data rows.');
  }
  const rows = closed.getRange(2, 2, closed.getLastRow() - 1, 15).getValues();
  let wouldStamp = 0;
  let wouldPreserve = 0;
  rows.forEach(function(row) {
    const auditId = String(row[0] || '').trim();
    if (!auditId) return;
    if (row[14] === '' || row[14] === null) wouldStamp++;
    else wouldPreserve++;
  });
  const result = {
    mode: 'READ ONLY',
    placeholderDate: Utilities.formatDate(
      phase4gTodayAtNoon_(ss.getSpreadsheetTimeZone()),
      ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy'
    ),
    wouldWritePlaceholderDates: wouldStamp,
    wouldPreserveExistingDates: wouldPreserve,
    emailsSent: false
  };
  Logger.log('PHASE4G_BASELINE_PREVIEW=' + JSON.stringify(result));
  return result;
}

function phase4gTodayAtNoon_(timeZone) {
  const parts = Utilities.formatDate(new Date(), timeZone, 'yyyy,MM,dd')
    .split(',').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0, 0);
}
