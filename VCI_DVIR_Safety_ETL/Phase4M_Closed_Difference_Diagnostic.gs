/**
 * Phase 4M: read-only diagnostic for proposed Closed updates.
 * Requires installed Phase 4A, 4B, 4D, 4E, 4H, and 4I helpers.
 * No writes, emails, backups, or archiving occur.
 */

function diagnosePhase4MClosedDifferences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = validateWorkbookIntegrity('PHASE4M_CLOSED_DIAGNOSTIC');
  const validation = base.validation || base;
  if ((validation.blockers || []).length || validation.status === 'BLOCKED') {
    throw new Error('Phase 4M stopped: workbook validation is BLOCKED.');
  }

  const context = phase4eBuildContext_(ss);
  const sourceRows = phase4iCurrentClosedRows_(ss, context.plans);
  const closed = phase4RequireSheet_(ss, 'Closed');
  const headers = closed.getRange(1, 1, 1, 16).getDisplayValues()[0];
  const existing = closed.getLastRow() > 1
    ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];
  const byAudit = {};
  existing.forEach(function(row, offset) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!byAudit[auditId]) byAudit[auditId] = [];
    byAudit[auditId].push({ closedRow: offset + 2, values: row, used: false });
  });

  const differences = [];
  let matched = 0;
  let unchanged = 0;
  let missing = 0;
  sourceRows.forEach(function(source) {
    const candidates = (byAudit[source.auditId] || []).filter(function(x) { return !x.used; });
    let best = null;
    let bestScore = -1;
    candidates.forEach(function(candidate) {
      const score = phase4hSimilarity_(source.values, candidate.values);
      if (score > bestScore) { bestScore = score; best = candidate; }
    });
    if (!best) { missing++; return; }
    best.used = true;
    matched++;
    if (phase4hEquivalent_(source.values, best.values)) { unchanged++; return; }
    const fields = [];
    for (let column = 0; column < 15; column++) {
      if (column === 1) continue;
      if (phase4mEquivalentCell_(column, source.values[column], best.values[column])) continue;
      fields.push({
        column: column + 1,
        header: String(headers[column] || ('Column ' + (column + 1))),
        closedValue: String(best.values[column] || ''),
        sourceValue: String(source.values[column] || ''),
        classification: phase4mDifferenceClass_(column, source.values[column], best.values[column])
      });
    }
    differences.push({ sourceType: source.sourceType, auditId: source.auditId,
      sourceRow: source.sourceRow, closedRow: best.closedRow, matchScore: bestScore,
      emailedDatePreserved: String(best.values[15] || ''), fields: fields });
  });

  const classificationCounts = {};
  differences.forEach(function(item) {
    item.fields.forEach(function(field) {
      classificationCounts[field.classification] =
        (classificationCounts[field.classification] || 0) + 1;
    });
  });
  const report = {
    mode: 'READ ONLY - CLOSED DIFFERENCE DIAGNOSTIC',
    validationStatus: validation.status,
    eligibleCurrentClosedRows: sourceRows.length,
    matchedRows: matched,
    unchangedRows: unchanged,
    proposedUpdateRows: differences.length,
    missingClosedRows: missing,
    classificationCounts: classificationCounts,
    differences: differences,
    policy: [
      'Emailed Date is never compared as updateable business data and will be preserved.',
      'Vehicle leading-zero, driver-name ordering, and issue formatting/aliases are treated as equivalent.',
      'No Closed update will be enabled until every remaining difference is classified.'
    ],
    writesPerformed: false,
    emailsSent: false,
    sourcesArchived: false
  };
  Logger.log('PHASE4M_DIAGNOSTIC=' + JSON.stringify(report));
  SpreadsheetApp.getUi().alert([
    'Phase 4M diagnostic completed — READ ONLY.', '',
    'Current Closed rows: ' + sourceRows.length,
    'Unchanged: ' + unchanged,
    'Proposed updates requiring review: ' + differences.length,
    'Missing Closed rows: ' + missing, '',
    'Copy PHASE4M_DIAGNOSTIC from Executions → Logs.'
  ].join('\n'));
  return report;
}

function phase4mEquivalentCell_(column, source, closed) {
  if (column === 3) return phase4dVehicleKey_(source) === phase4dVehicleKey_(closed);
  if (column === 5) return phase4hNamesOverlap_(source, closed);
  if (column === 6) return phase4dIssueSignature_(source) === phase4dIssueSignature_(closed);
  return phase4hCell_(source) === phase4hCell_(closed);
}

function phase4mDifferenceClass_(column, source, closed) {
  if (column === 0 || column === 8 || column === 11 || column === 12) return 'DATE_OR_TIME';
  if (column === 2) return 'DIVISION';
  if (column === 4) return 'EQUIPMENT_LOOKUP';
  if (column === 7 || column === 10 || column === 13 || column === 14) return 'OPERATIONAL_NOTE_OR_REPAIR_FIELD';
  if (column === 9) return 'STATUS';
  return 'OTHER';
}
