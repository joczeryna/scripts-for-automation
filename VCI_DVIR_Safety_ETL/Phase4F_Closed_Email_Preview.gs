/**
 * Phase 4F: read-only Closed synchronization and email-readiness preview.
 * Requires Phase 4A, Phase 4B, and Phase 4D in the same Apps Script project.
 */

function previewClosedAndEmailReadiness() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseValidation = validateWorkbookIntegrity('PRE_EMAIL_PREVIEW');
  const validation = {
    status: baseValidation.validation.status,
    blockers: baseValidation.validation.blockers.slice(),
    warnings: baseValidation.validation.warnings.slice()
  };
  const closedPlan = phase4fClosedPlan_(ss, validation);
  const emailReadiness = phase4fEmailReadiness_(ss, closedPlan, validation);
  if (validation.blockers.length) validation.status = 'BLOCKED';
  else if (validation.warnings.length) validation.status = 'WARNING';
  else validation.status = 'PASS';

  const report = {
    mode: 'READ ONLY - NO CLOSED WRITES OR EMAILS',
    validation: validation,
    closedSynchronization: closedPlan.summary,
    emailReadiness: emailReadiness,
    policies: [
      'Only Issues and Gate Checks rows with Status Closed are eligible for Closed.',
      'JSO and Out of Service rows are excluded.',
      'Repeated Audit IDs are matched one-to-one using Vehicle, Assigned Driver, issue meaning, and row similarity.',
      'Existing Emailed Dates are preserved on matched Closed rows.',
      'Unresolved individualized recipients are skipped and logged; the full safety distribution remains independently eligible.',
      'No email may run when validation status is BLOCKED.'
    ]
  };
  Logger.log('PHASE4F_SUMMARY=' + JSON.stringify(report));
  SpreadsheetApp.getUi().alert([
    'Phase 4F completed — NO DATA WAS WRITTEN AND NO EMAIL WAS SENT.',
    '',
    'Validation: ' + validation.status,
    'Closed inserts: ' + closedPlan.summary.inserts,
    'Closed updates: ' + closedPlan.summary.updates,
    'Closed unchanged: ' + closedPlan.summary.unchanged,
    'Emailed Dates preserved: ' + closedPlan.summary.emailedDatesPreserved,
    '',
    'Pending full-report rows: ' + emailReadiness.pendingFullReportRows,
    'Individualized rows ready: ' + emailReadiness.individualizedRowsReady,
    'Individualized rows skipped: ' + emailReadiness.individualizedRowsSkipped,
    '',
    'Copy PHASE4F_SUMMARY from Executions → Logs.'
  ].join('\n'));
  return report;
}

function phase4fClosedPlan_(ss, validation) {
  const closedSheet = ss.getSheetByName('Closed');
  if (!closedSheet || closedSheet.getLastColumn() < 16) {
    phase4dFinding_(validation.blockers, 'Closed', 1, '',
      'Closed must exist and contain at least 16 columns.');
    return phase4fEmptyClosedPlan_();
  }

  const sourceRows = [];
  ['Issues', 'Gate Checks'].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1,
      sheet.getLastColumn()).getDisplayValues();
    rows.forEach(function(row, index) {
      const auditId = String(row[1] || '').trim();
      const status = String(row[9] || '').trim();
      if (!auditId || status !== 'Closed') return;
      sourceRows.push({ sourceType: name, sourceRow: index + 2,
        auditId: auditId, values: phase4fPad_(row, 15) });
    });
  });

  const closedRows = closedSheet.getLastRow() > 1
    ? closedSheet.getRange(2, 1, closedSheet.getLastRow() - 1, 16).getDisplayValues()
    : [];
  const closedByAudit = {};
  const exactClosedSignatures = {};
  closedRows.forEach(function(row, index) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!closedByAudit[auditId]) closedByAudit[auditId] = [];
    const item = { closedRow: index + 2, values: row, used: false };
    closedByAudit[auditId].push(item);
    const signature = row.slice(0, 15).map(phase4fCell_).join('|');
    exactClosedSignatures[signature] = (exactClosedSignatures[signature] || 0) + 1;
  });

  let exactDuplicateClosedRows = 0;
  Object.keys(exactClosedSignatures).forEach(function(key) {
    if (exactClosedSignatures[key] > 1) {
      exactDuplicateClosedRows += exactClosedSignatures[key] - 1;
    }
  });
  if (exactDuplicateClosedRows) {
    phase4dFinding_(validation.warnings, 'Closed', '', '',
      exactDuplicateClosedRows + ' exact duplicate Closed rows require review.');
  }

  const actions = [];
  const summary = { eligibleSourceRows: sourceRows.length, inserts: 0, updates: 0,
    unchanged: 0, emailedDatesPreserved: 0, exactDuplicateClosedRows: exactDuplicateClosedRows,
    unmatchedExistingClosedRowsPreserved: 0, sample: [] };

  sourceRows.forEach(function(source) {
    const candidates = (closedByAudit[source.auditId] || []).filter(function(item) {
      return !item.used && phase4fCoreCompatible_(source.values, item.values);
    });
    let match = null;
    let bestScore = -1;
    candidates.forEach(function(candidate) {
      const score = phase4fSimilarity_(source.values, candidate.values);
      if (score > bestScore) { bestScore = score; match = candidate; }
    });
    if (!match) {
      summary.inserts++;
      actions.push({ action: 'INSERT', source: source, emailedDate: '' });
      return;
    }
    match.used = true;
    const emailedDate = String(match.values[15] || '').trim();
    if (emailedDate) summary.emailedDatesPreserved++;
    const equal = phase4fRowsEquivalent_(source.values, match.values);
    if (equal) summary.unchanged++;
    else summary.updates++;
    actions.push({ action: equal ? 'UNCHANGED' : 'UPDATE', source: source,
      closedRow: match.closedRow, emailedDate: emailedDate });
    if (!equal && summary.sample.length < 8) {
      summary.sample.push({ auditId: source.auditId, sourceType: source.sourceType,
        sourceRow: source.sourceRow, closedRow: match.closedRow,
        differingColumns: phase4fDifferingColumns_(closedSheet, source.values, match.values) });
    }
  });
  Object.keys(closedByAudit).forEach(function(auditId) {
    closedByAudit[auditId].forEach(function(item) {
      if (!item.used) summary.unmatchedExistingClosedRowsPreserved++;
    });
  });
  return { summary: summary, actions: actions };
}

function phase4fEmailReadiness_(ss, closedPlan, validation) {
  const indexSheet = ss.getSheetByName('Automation Recipient Index');
  const indexRows = indexSheet.getLastRow() > 1
    ? indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 8).getDisplayValues() : [];
  const index = {};
  indexRows.forEach(function(row) {
    const key = String(row[0] || '').trim() + '|' + String(row[1] || '').trim();
    if (key !== '|') index[key] = row;
  });
  const directory = ss.getSheetByName('Company Directory').getDataRange().getDisplayValues();
  const header = phase4bHeaderMap_(directory[0]);
  const byWorker = {};
  directory.slice(1).forEach(function(row) {
    const workerNo = phase4Normalize_(row[header.workerno]);
    if (!workerNo) return;
    if (!byWorker[workerNo]) byWorker[workerNo] = [];
    byWorker[workerNo].push({ email: String(row[header.emailaddress] || '').trim(),
      supervisor: String(row[header.supervisor] || '').trim() });
  });

  const result = { pendingFullReportRows: 0, alreadyEmailedRows: 0,
    individualizedRowsReady: 0, individualizedRowsSkipped: 0,
    uniqueEmployeeEmails: {}, uniqueSupervisorEmails: {}, skippedSample: [] };
  closedPlan.actions.forEach(function(action) {
    if (action.emailedDate) {
      result.alreadyEmailedRows++;
      return;
    }
    result.pendingFullReportRows++;
    const key = action.source.sourceType + '|' + action.source.auditId;
    const mapping = index[key];
    if (!mapping || ['Matched', 'Name Fallback'].indexOf(String(mapping[6])) === -1) {
      result.individualizedRowsSkipped++;
      if (result.skippedSample.length < 10) result.skippedSample.push({
        auditId: action.source.auditId, sourceType: action.source.sourceType,
        reason: mapping ? String(mapping[6]) : 'Recipient index missing' });
      return;
    }
    const workers = byWorker[phase4Normalize_(mapping[2])] || [];
    if (workers.length !== 1 || !phase4dIsEmail_(workers[0].email) ||
        !phase4dIsEmail_(workers[0].supervisor)) {
      result.individualizedRowsSkipped++;
      if (result.skippedSample.length < 10) result.skippedSample.push({
        auditId: action.source.auditId, sourceType: action.source.sourceType,
        reason: 'Current directory email or supervisor email is unresolved' });
      return;
    }
    result.individualizedRowsReady++;
    result.uniqueEmployeeEmails[workers[0].email.toLowerCase()] = true;
    result.uniqueSupervisorEmails[workers[0].supervisor.toLowerCase()] = true;
  });
  result.uniqueEmployeeEmailCount = Object.keys(result.uniqueEmployeeEmails).length;
  result.uniqueSupervisorEmailCount = Object.keys(result.uniqueSupervisorEmails).length;
  delete result.uniqueEmployeeEmails;
  delete result.uniqueSupervisorEmails;
  if (result.individualizedRowsSkipped) phase4dFinding_(validation.warnings,
    'Automation Recipient Index', '', '', result.individualizedRowsSkipped +
    ' pending Closed rows would be skipped from individualized email.');
  return result;
}

function phase4fCoreCompatible_(source, closed) {
  return phase4dVehicleKey_(source[3]) === phase4dVehicleKey_(closed[3]) &&
    phase4fNamesOverlap_(source[5], closed[5]) &&
    phase4dIssueSignature_(source[6]) === phase4dIssueSignature_(closed[6]);
}

function phase4fNamesOverlap_(left, right) {
  if (!String(left || '').trim() && !String(right || '').trim()) return true;
  const rightKeys = {};
  phase4dNameKeys_(right).forEach(function(key) { rightKeys[key] = true; });
  return phase4dNameKeys_(left).some(function(key) { return rightKeys[key]; });
}

function phase4fSimilarity_(source, closed) {
  let score = 0;
  for (let index = 0; index < 15; index++) {
    if (index === 1 || index === 9) continue;
    if (index === 6) {
      if (phase4dIssueSignature_(source[index]) === phase4dIssueSignature_(closed[index])) score += 5;
    } else if (phase4fCell_(source[index]) === phase4fCell_(closed[index])) {
      score += [3, 5].indexOf(index) !== -1 ? 4 : 1;
    }
  }
  return score;
}

function phase4fRowsEquivalent_(source, closed) {
  for (let index = 0; index < 15; index++) {
    if (index === 1) continue;
    if (index === 6) {
      if (phase4dIssueSignature_(source[index]) !== phase4dIssueSignature_(closed[index])) return false;
    } else if (phase4fCell_(source[index]) !== phase4fCell_(closed[index])) return false;
  }
  return true;
}

function phase4fDifferingColumns_(sheet, source, closed) {
  const result = [];
  for (let index = 0; index < 15; index++) {
    const same = index === 6
      ? phase4dIssueSignature_(source[index]) === phase4dIssueSignature_(closed[index])
      : phase4fCell_(source[index]) === phase4fCell_(closed[index]);
    if (!same) result.push(String(sheet.getRange(1, index + 1).getDisplayValue() ||
      'Column ' + (index + 1)));
  }
  return result;
}

function phase4fPad_(row, length) {
  const copy = row.slice(0, length);
  while (copy.length < length) copy.push('');
  return copy;
}

function phase4fCell_(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
}

function phase4fEmptyClosedPlan_() {
  return { summary: { eligibleSourceRows: 0, inserts: 0, updates: 0,
    unchanged: 0, emailedDatesPreserved: 0, exactDuplicateClosedRows: 0,
    unmatchedExistingClosedRowsPreserved: 0, sample: [] }, actions: [] };
}
