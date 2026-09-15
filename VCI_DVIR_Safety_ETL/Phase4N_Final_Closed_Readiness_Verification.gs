/**
 * Phase 4N: uniquely named final Closed-readiness verification.
 * Read only. Requires Phase 4A, 4B, 4D, 4E, 4H, and 4I helpers.
 */

function verifyPhase4NFinalClosedReadiness() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = validateWorkbookIntegrity('PHASE4N_FINAL_CLOSED_READINESS');
  const validation = base.validation || base;
  if ((validation.blockers || []).length || validation.status === 'BLOCKED') {
    throw new Error('Phase 4N stopped: workbook validation is BLOCKED.');
  }
  const plans = phase4eBuildContext_(ss).plans;
  const sources = phase4iCurrentClosedRows_(ss, plans);
  const closed = phase4RequireSheet_(ss, 'Closed');
  const existing = closed.getLastRow() > 1
    ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];
  const byAudit = {};
  existing.forEach(function(row, offset) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!byAudit[auditId]) byAudit[auditId] = [];
    byAudit[auditId].push({ closedRow: offset + 2, values: row, used: false });
  });
  const result = { eligibleSourceRows: sources.length, inserts: 0, updates: 0,
    unchanged: 0, emailedDatesPreserved: 0, differences: [] };
  sources.forEach(function(source) {
    const candidates = (byAudit[source.auditId] || []).filter(function(x) { return !x.used; });
    let best = null;
    let bestScore = -1;
    candidates.forEach(function(candidate) {
      const score = phase4hSimilarity_(source.values, candidate.values);
      if (score > bestScore) { best = candidate; bestScore = score; }
    });
    if (!best) { result.inserts++; return; }
    best.used = true;
    if (String(best.values[15] || '').trim()) result.emailedDatesPreserved++;
    const fields = phase4nDifferentFields_(source.values, best.values);
    if (!fields.length) result.unchanged++;
    else {
      result.updates++;
      result.differences.push({ auditId: source.auditId, sourceType: source.sourceType,
        sourceRow: source.sourceRow, closedRow: best.closedRow, fields: fields });
    }
  });
  const report = {
    mode: 'READ ONLY - UNIQUE DATE-AWARE CLOSED VERIFICATION',
    validationStatus: validation.status,
    closed: result,
    readyForControlledProductionWorkflow:
      result.inserts === 0 && result.updates === 0 && result.unchanged === result.eligibleSourceRows,
    datePolicy: 'Equivalent calendar dates remain unchanged when one display value includes a midnight suffix.',
    writesPerformed: false,
    emailsSent: false,
    sourcesArchived: false
  };
  Logger.log('PHASE4N_RESULT=' + JSON.stringify(report));
  SpreadsheetApp.getUi().alert([
    'Phase 4N completed — READ ONLY.', '',
    'Closed inserts: ' + result.inserts,
    'Closed updates: ' + result.updates,
    'Closed unchanged: ' + result.unchanged,
    'Emailed Dates preserved: ' + result.emailedDatesPreserved,
    'Production-ready: ' + report.readyForControlledProductionWorkflow, '',
    'Copy PHASE4N_RESULT from Executions → Logs.'
  ].join('\n'));
  return report;
}

function phase4nDifferentFields_(source, closed) {
  const fields = [];
  for (let column = 0; column < 15; column++) {
    if (column === 1) continue;
    let equal;
    if ([0, 8, 11, 12].indexOf(column) !== -1) {
      equal = phase4nDateKey_(source[column]) === phase4nDateKey_(closed[column]);
    } else if (column === 3) {
      equal = phase4dVehicleKey_(source[column]) === phase4dVehicleKey_(closed[column]);
    } else if (column === 5) {
      equal = phase4hNamesOverlap_(source[column], closed[column]);
    } else if (column === 6) {
      equal = phase4dIssueSignature_(source[column]) === phase4dIssueSignature_(closed[column]);
    } else {
      equal = phase4hCell_(source[column]) === phase4hCell_(closed[column]);
    }
    if (!equal) fields.push(column + 1);
  }
  return fields;
}

function phase4nDateKey_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return match[3] + '-' + ('0' + match[1]).slice(-2) + '-' + ('0' + match[2]).slice(-2);
  match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) return match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
  return phase4hCell_(text);
}
