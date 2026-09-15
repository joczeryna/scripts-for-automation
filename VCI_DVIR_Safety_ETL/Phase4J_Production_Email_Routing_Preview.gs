/**
 * Phase 4J: read-only production email-routing preview.
 *
 * Requires Phase 4A, Phase 4B, and Phase 4D in the Apps Script project.
 * Reads only Closed rows with a blank Emailed Date. Sends no email and writes nothing.
 */

const PHASE4J_VALIDATION_EMAIL = 'jo.deleon@vcicom.com';

function previewPhase4JProductionEmailRouting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = validateWorkbookIntegrity('PHASE4J_PRE_EMAIL_PREVIEW');
  const validation = base.validation || base;
  const blockers = (validation.blockers || []).slice();
  const warnings = (validation.warnings || []).slice();

  const closed = phase4RequireSheet_(ss, 'Closed');
  if (closed.getLastColumn() < 16) {
    phase4dFinding_(blockers, 'Closed', 1, '',
      'Closed must contain at least 16 columns, including Emailed Date in column P.');
  }

  const index = phase4jRecipientIndex_(ss, blockers);
  const directory = phase4jDirectory_(ss, blockers);
  const rows = closed.getLastRow() > 1
    ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];

  const ready = [];
  const unresolved = [];
  let alreadyEmailed = 0;
  rows.forEach(function(row, offset) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (String(row[15] || '').trim()) {
      alreadyEmailed++;
      return;
    }
    const sourceType = phase4jSourceType_(row, index);
    if (!sourceType) {
      unresolved.push(phase4jFailure_(offset + 2, auditId, '',
        'Source Type could not be resolved as Issues or Gate Checks.'));
      return;
    }
    const mapping = index[sourceType + '|' + auditId];
    if (!mapping) {
      unresolved.push(phase4jFailure_(offset + 2, auditId, sourceType,
        'Recipient index mapping is missing.'));
      return;
    }
    if (['Matched', 'Name Fallback'].indexOf(mapping.status) === -1) {
      unresolved.push(phase4jFailure_(offset + 2, auditId, sourceType,
        'Recipient index status is ' + (mapping.status || 'blank') + '.'));
      return;
    }
    const employees = directory.byWorker[phase4Normalize_(mapping.workerNo)] || [];
    if (employees.length !== 1) {
      unresolved.push(phase4jFailure_(offset + 2, auditId, sourceType,
        employees.length ? 'Worker No is ambiguous in Company Directory.' :
          'Worker No was not found in Company Directory.'));
      return;
    }
    const employee = employees[0];
    if (!phase4dIsEmail_(employee.email)) {
      unresolved.push(phase4jFailure_(offset + 2, auditId, sourceType,
        'Assigned Driver email is blank or invalid.'));
      return;
    }
    if (!phase4dIsEmail_(employee.supervisor)) {
      unresolved.push(phase4jFailure_(offset + 2, auditId, sourceType,
        'Supervisor email is blank or invalid.'));
      return;
    }
    ready.push({ closedRow: offset + 2, auditId: auditId, sourceType: sourceType,
      to: employee.email.toLowerCase(), cc: employee.supervisor.toLowerCase() });
  });

  if (unresolved.length) {
    phase4dFinding_(warnings, 'Closed', '', '', unresolved.length +
      ' pending operational emails are unresolved and would be suppressed; ' +
      'their validation-error report would go only to ' + PHASE4J_VALIDATION_EMAIL + '.');
  }
  const status = blockers.length ? 'BLOCKED' : (warnings.length ? 'WARNING' : 'PASS');
  const report = {
    mode: 'READ ONLY - NO WRITES, EMAILS OR ARCHIVING',
    validation: { status: status, blockers: blockers, warnings: warnings },
    queue: {
      closedRowsReviewed: rows.filter(function(row) { return String(row[1] || '').trim(); }).length,
      alreadyEmailedRows: alreadyEmailed,
      pendingRows: ready.length + unresolved.length,
      validOperationalEmails: ready.length,
      unresolvedOperationalEmails: unresolved.length,
      uniqueDrivers: phase4jUniqueCount_(ready, 'to'),
      uniqueSupervisors: phase4jUniqueCount_(ready, 'cc'),
      validSample: ready.slice(0, 10),
      unresolvedSample: unresolved.slice(0, 10)
    },
    routing: {
      valid: 'To Assigned Driver; CC Assigned Driver Supervisor',
      unresolved: 'Suppress operational email; send validation-error report only to ' +
        PHASE4J_VALIDATION_EMAIL,
      validationErrorRecipient: PHASE4J_VALIDATION_EMAIL,
      hardcodedOperationalRecipients: 0
    },
    safeguards: [
      'Only Closed rows with a blank Emailed Date enter the queue.',
      'A valid operational record requires exactly one directory employee, a valid driver email, and a valid supervisor email.',
      'Unresolved records are never partially sent to a driver or supervisor.',
      'Jo receives validation/error reporting only, never valid operational emails.',
      'Emailed Date will be written only after successful operational delivery in a later production phase.'
    ],
    emailsSent: false,
    writesPerformed: false,
    sourcesArchived: false
  };

  Logger.log('PHASE4J_SUMMARY=' + JSON.stringify(report));
  SpreadsheetApp.getUi().alert([
    'Phase 4J completed — READ ONLY.', '',
    'Validation: ' + status,
    'Pending rows: ' + report.queue.pendingRows,
    'Valid driver/supervisor routes: ' + ready.length,
    'Unresolved routes for Jo report: ' + unresolved.length, '',
    'No emails were sent. Copy PHASE4J_SUMMARY from Executions → Logs.'
  ].join('\n'));
  return report;
}

function phase4jRecipientIndex_(ss, blockers) {
  const sheet = ss.getSheetByName('Automation Recipient Index');
  const result = {};
  if (!sheet || sheet.getLastRow() < 2) return result;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getDisplayValues();
  rows.forEach(function(row, offset) {
    const sourceType = String(row[0] || '').trim();
    const auditId = String(row[1] || '').trim();
    if (!sourceType || !auditId) return;
    const key = sourceType + '|' + auditId;
    if (result[key]) {
      phase4dFinding_(blockers, 'Automation Recipient Index', offset + 2, auditId,
        'Duplicate Source Type + Audit ID recipient mapping.');
      return;
    }
    result[key] = { workerNo: String(row[2] || '').trim(),
      assignedDriver: String(row[3] || '').trim(), status: String(row[6] || '').trim() };
    if (!result.__byAudit) result.__byAudit = {};
    if (!result.__byAudit[auditId]) result.__byAudit[auditId] = [];
    result.__byAudit[auditId].push({ sourceType: sourceType, key: key,
      assignedDriver: String(row[3] || '').trim() });
  });
  return result;
}

function phase4jDirectory_(ss, blockers) {
  const sheet = ss.getSheetByName('Company Directory');
  const data = sheet.getDataRange().getDisplayValues();
  const header = phase4bHeaderMap_(data[0]);
  const byWorker = {};
  data.slice(1).forEach(function(row) {
    const workerNo = phase4Normalize_(row[header.workerno]);
    if (!workerNo) return;
    if (!byWorker[workerNo]) byWorker[workerNo] = [];
    byWorker[workerNo].push({ name: String(row[header.name] || '').trim(),
      email: String(row[header.emailaddress] || '').trim(),
      supervisor: String(row[header.supervisor] || '').trim() });
  });
  return { byWorker: byWorker };
}

function phase4jSourceType_(row, index) {
  const auditId = String(row[1] || '').trim();
  const candidates = (index.__byAudit && index.__byAudit[auditId]) || [];
  if (candidates.length === 1) return candidates[0].sourceType;
  const driver = String(row[5] || '').trim();
  const matches = candidates.filter(function(candidate) {
    return phase4hNamesOverlap_(driver, candidate.assignedDriver);
  });
  return matches.length === 1 ? matches[0].sourceType : '';
}

function phase4jFailure_(closedRow, auditId, sourceType, reason) {
  return { closedRow: closedRow, auditId: auditId, sourceType: sourceType,
    reason: reason, errorRecipient: PHASE4J_VALIDATION_EMAIL };
}

function phase4jUniqueCount_(rows, field) {
  const values = {};
  rows.forEach(function(row) { values[row[field]] = true; });
  return Object.keys(values).length;
}
