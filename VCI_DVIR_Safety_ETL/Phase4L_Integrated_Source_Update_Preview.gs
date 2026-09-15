/**
 * Phase 4L: integrated, read-only source-update preview.
 *
 * Requires the installed Phase 4A, 4B, 4D, 4E, 4H, 4I, and 4J helpers.
 * No workbook writes, email, backup, or archiving occurs.
 */

function previewPhase4LIntegratedSourceUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = validateWorkbookIntegrity('PHASE4L_INTEGRATED_PREVIEW');
  const validation = base.validation || base;
  const blockers = (validation.blockers || []).slice();
  const warnings = (validation.warnings || []).slice();

  const sources = phase4lSourceFiles_();
  const context = phase4eBuildContext_(ss);
  const targets = {
    Issues: phase4lTargetPreview_(ss.getSheetByName('Issues'),
      context.plans.Issues, context.equipment),
    GateChecks: phase4lTargetPreview_(ss.getSheetByName('Gate Checks'),
      context.plans['Gate Checks'], context.equipment),
    JSO: phase4lTargetPreview_(ss.getSheetByName('JSO'),
      context.plans.JSO, context.equipment)
  };
  const closed = phase4lClosedPreview_(ss, context.plans);
  const email = phase4lEmailPreview_(ss);

  if (email.unresolvedRows) {
    phase4dFinding_(warnings, 'Closed', '', '', email.unresolvedRows +
      ' pending operational emails would be suppressed and reported only to jo.deleon@vcicom.com.');
  }
  const status = blockers.length ? 'BLOCKED' : (warnings.length ? 'WARNING' : 'PASS');
  const archiveEligibleAfterProductionRun = !blockers.length;
  const report = {
    mode: 'READ ONLY - NO WRITES, EMAILS, BACKUPS OR ARCHIVING',
    validation: { status: status, blockers: blockers, warnings: warnings },
    sourceFiles: sources,
    proposedWorksheetChanges: targets,
    proposedClosedSynchronization: closed,
    currentEmailQueue: email,
    archiveReadiness: {
      dvirEligibleAfterSuccessfulProductionRun: archiveEligibleAfterProductionRun,
      safetyDeviationsEligibleAfterSuccessfulProductionRun: archiveEligibleAfterProductionRun,
      archivedNow: false,
      rule: 'Archive only after source updates, Closed synchronization, email handling, and post-run validation all complete successfully.'
    },
    productionEmailPolicy: {
      valid: 'To Assigned Driver; CC Assigned Driver Supervisor',
      unresolved: 'Suppress operational email; notify only jo.deleon@vcicom.com',
      senderDisplayName: 'DVIR Report Automation',
      emailedDateRule: 'Write only after successful operational delivery.'
    }
  };
  Logger.log('PHASE4L_SUMMARY=' + JSON.stringify(report));
  SpreadsheetApp.getUi().alert([
    'Phase 4L completed — READ ONLY.', '',
    'Validation: ' + status,
    'Issues inserts/updates: ' + targets.Issues.inserts + '/' + targets.Issues.updates,
    'Gate Checks inserts/updates: ' + targets.GateChecks.inserts + '/' + targets.GateChecks.updates,
    'JSO inserts/updates: ' + targets.JSO.inserts + '/' + targets.JSO.updates,
    'Closed inserts/updates: ' + closed.inserts + '/' + closed.updates,
    'Pending valid/unresolved emails: ' + email.validRows + '/' + email.unresolvedRows, '',
    'No data was changed. Copy PHASE4L_SUMMARY from Executions → Logs.'
  ].join('\n'));
  return report;
}

function phase4lSourceFiles_() {
  const folder = DriveApp.getFolderById(PHASE4_INTAKE_FOLDER_ID);
  const runDate = Utilities.formatDate(new Date(), PHASE4_TIME_ZONE, 'MM-dd-yyyy');
  const dvir = phase4FindUniqueFile_(folder, ['DVIR_Details', 'DVIR_Details.xlsx']);
  const safety = phase4FindUniqueFile_(folder, [
    'safety-deviations-' + runDate, 'safety-deviations-' + runDate + '.xlsx'
  ]);
  return {
    dvir: { name: dvir.getName(), id: dvir.getId(), found: true },
    safetyDeviations: { name: safety.getName(), id: safety.getId(), found: true }
  };
}

function phase4lTargetPreview_(sheet, plans, equipment) {
  const name = sheet.getName();
  const isJso = name === 'JSO';
  const issueIndex = isJso ? 5 : 6;
  const statusIndex = isJso ? 8 : 9;
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues() : [];
  const groups = {};
  rows.forEach(function(row, offset) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!groups[auditId]) groups[auditId] = [];
    groups[auditId].push({ rowNumber: offset + 2, row: row });
  });
  const result = { sourceAuditIds: Object.keys(plans).length, inserts: 0, updates: 0,
    issueChanges: 0, statusChanges: 0, unchanged: 0, sample: [] };
  Object.keys(plans).forEach(function(auditId) {
    const plan = plans[auditId];
    const matches = phase4eRelevantMatches_(groups[auditId] || [], plan, name);
    if (!matches.length) {
      result.inserts++;
      if (result.sample.length < 8) result.sample.push({ action: 'INSERT', auditId: auditId });
      return;
    }
    let changedForAudit = false;
    matches.forEach(function(match) {
      const issueChanged = phase4dIssueSignature_(match.row[issueIndex]) !==
        phase4dIssueSignature_(phase4bIssueText_(plan));
      const plannedStatus = phase4eStatus_(plan, match.row[statusIndex], equipment, isJso);
      const statusChanged = String(match.row[statusIndex] || '').trim() !== plannedStatus;
      if (issueChanged) result.issueChanges++;
      if (statusChanged) result.statusChanges++;
      if (issueChanged || statusChanged) changedForAudit = true;
    });
    if (changedForAudit) {
      result.updates++;
      if (result.sample.length < 8) result.sample.push({ action: 'UPDATE', auditId: auditId });
    } else result.unchanged++;
  });
  return result;
}

function phase4lClosedPreview_(ss, plans) {
  const sourceRows = phase4iCurrentClosedRows_(ss, plans);
  const closed = ss.getSheetByName('Closed');
  const existing = closed.getLastRow() > 1
    ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];
  const byAudit = {};
  existing.forEach(function(row, offset) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!byAudit[auditId]) byAudit[auditId] = [];
    byAudit[auditId].push({ row: offset + 2, values: row, used: false });
  });
  const result = { eligibleSourceRows: sourceRows.length, inserts: 0, updates: 0,
    unchanged: 0, emailedDatesPreserved: 0, sample: [] };
  sourceRows.forEach(function(source) {
    const candidates = (byAudit[source.auditId] || []).filter(function(x) { return !x.used; });
    let best = null;
    let score = -1;
    candidates.forEach(function(candidate) {
      const candidateScore = phase4hSimilarity_(source.values, candidate.values);
      if (candidateScore > score) { score = candidateScore; best = candidate; }
    });
    if (!best) {
      result.inserts++;
      if (result.sample.length < 8) result.sample.push({ action: 'INSERT', auditId: source.auditId });
      return;
    }
    best.used = true;
    if (String(best.values[15] || '').trim()) result.emailedDatesPreserved++;
    if (phase4lClosedEquivalent_(source.values, best.values)) result.unchanged++;
    else {
      result.updates++;
      if (result.sample.length < 8) result.sample.push({ action: 'UPDATE', auditId: source.auditId });
    }
  });
  return result;
}

function phase4lEmailPreview_(ss) {
  const closed = ss.getSheetByName('Closed');
  const rows = closed.getLastRow() > 1
    ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];
  const index = phase4jRecipientIndex_(ss, []);
  const directory = phase4jDirectory_(ss, []);
  const result = { pendingRows: 0, validRows: 0, unresolvedRows: 0,
    alreadyEmailedRows: 0, validSample: [], unresolvedSample: [] };
  rows.forEach(function(row, offset) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (String(row[15] || '').trim()) { result.alreadyEmailedRows++; return; }
    result.pendingRows++;
    const sourceType = phase4jSourceType_(row, index);
    const mapping = sourceType ? index[sourceType + '|' + auditId] : null;
    const employees = mapping ? directory.byWorker[phase4Normalize_(mapping.workerNo)] || [] : [];
    const employee = employees.length === 1 ? employees[0] : null;
    if (mapping && ['Matched', 'Name Fallback'].indexOf(mapping.status) !== -1 &&
        employee && phase4dIsEmail_(employee.email) && phase4dIsEmail_(employee.supervisor)) {
      result.validRows++;
      if (result.validSample.length < 8) result.validSample.push({ closedRow: offset + 2,
        auditId: auditId, sourceType: sourceType, to: employee.email.toLowerCase(),
        cc: employee.supervisor.toLowerCase() });
    } else {
      result.unresolvedRows++;
      if (result.unresolvedSample.length < 8) result.unresolvedSample.push({
        closedRow: offset + 2, auditId: auditId,
        reason: !sourceType ? 'Source Type unresolved' : !mapping ? 'Mapping missing' :
          'Driver or supervisor recipient unresolved' });
    }
  });
  return result;
}

function phase4lClosedEquivalent_(source, closed) {
  for (let column = 0; column < 15; column++) {
    if (column === 1) continue;
    if ([0, 8, 11, 12].indexOf(column) !== -1) {
      if (phase4lDateKey_(source[column]) !== phase4lDateKey_(closed[column])) return false;
    } else if (column === 3) {
      if (phase4dVehicleKey_(source[column]) !== phase4dVehicleKey_(closed[column])) return false;
    } else if (column === 5) {
      if (!phase4hNamesOverlap_(source[column], closed[column])) return false;
    } else if (column === 6) {
      if (phase4dIssueSignature_(source[column]) !== phase4dIssueSignature_(closed[column])) return false;
    } else if (phase4hCell_(source[column]) !== phase4hCell_(closed[column])) return false;
  }
  return true;
}

function phase4lDateKey_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return match[3] + '-' + ('0' + match[1]).slice(-2) + '-' + ('0' + match[2]).slice(-2);
  match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) return match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
  return phase4hCell_(text);
}
