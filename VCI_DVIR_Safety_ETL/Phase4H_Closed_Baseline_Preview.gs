/**
 * Phase 4H: corrected, read-only Closed baseline preview.
 *
 * Requires Phase 4A, 4B and 4D/4E helpers in the Apps Script project.
 * No writes, emails or archiving occur.
 */

const PHASE4H_ERROR_EMAIL = 'jo.deleon@vcicom.com';

function previewCorrectedClosedBaseline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = validateWorkbookIntegrity('PHASE4H_CLOSED_BASELINE_PREVIEW');
  const validation = base.validation || base;
  if ((validation.blockers || []).length || validation.status === 'BLOCKED') {
    throw new Error('Phase 4H stopped: workbook validation is BLOCKED. No changes were made.');
  }

  const currentPlans = phase4hCurrentPlans_();
  const currentAuditKeys = {};
  ['Issues', 'Gate Checks'].forEach(function(sourceType) {
    Object.keys(currentPlans[sourceType]).forEach(function(auditId) {
      currentAuditKeys[sourceType + '|' + auditId] = true;
    });
  });

  const sourceRows = [];
  ['Issues', 'Gate Checks'].forEach(function(sourceType) {
    const sheet = phase4RequireSheet_(ss, sourceType);
    if (sheet.getLastRow() < 2) return;
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1,
      Math.max(16, sheet.getLastColumn())).getDisplayValues();
    rows.forEach(function(row, index) {
      const auditId = String(row[1] || '').trim();
      if (!auditId || !currentAuditKeys[sourceType + '|' + auditId]) return;
      if (String(row[9] || '').trim() !== 'Closed') return;
      sourceRows.push({
        sourceType: sourceType,
        auditId: auditId,
        sourceRow: index + 2,
        values: phase4hPad_(row, 15)
      });
    });
  });

  const closed = phase4RequireSheet_(ss, 'Closed');
  const closedRows = closed.getLastRow() > 1
    ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];
  const closedByAudit = {};
  closedRows.forEach(function(row, index) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!closedByAudit[auditId]) closedByAudit[auditId] = [];
    closedByAudit[auditId].push({ row: index + 2, values: row, used: false });
  });

  const actions = [];
  sourceRows.forEach(function(source) {
    const candidates = (closedByAudit[source.auditId] || []).filter(function(candidate) {
      return !candidate.used;
    });
    let best = null;
    let bestScore = -1;
    candidates.forEach(function(candidate) {
      const score = phase4hSimilarity_(source.values, candidate.values);
      if (score > bestScore) {
        best = candidate;
        bestScore = score;
      }
    });
    if (!best) {
      actions.push({ action: 'BASELINE_INSERT', source: source });
      return;
    }
    best.used = true;
    actions.push({
      action: phase4hEquivalent_(source.values, best.values) ? 'UNCHANGED' : 'UPDATE',
      source: source,
      closedRow: best.row,
      existingEmailedDate: String(best.values[15] || '').trim(),
      matchScore: bestScore
    });
  });

  const summary = {
    mode: 'READ ONLY - NO WRITES, EMAILS OR ARCHIVING',
    validationStatus: validation.status,
    currentSourceAuditIds: Object.keys(currentAuditKeys).length,
    currentClosedSourceRows: sourceRows.length,
    matchedExistingClosedRows: actions.filter(function(x) { return x.action !== 'BASELINE_INSERT'; }).length,
    unchanged: actions.filter(function(x) { return x.action === 'UNCHANGED'; }).length,
    proposedUpdates: actions.filter(function(x) { return x.action === 'UPDATE'; }).length,
    proposedBaselineInserts: actions.filter(function(x) { return x.action === 'BASELINE_INSERT'; }).length,
    proposedEmailCandidates: 0,
    historicalClosedRowsPreserved: closedRows.length,
    historicalBlankEmailedDates: phase4hBlankEmailedDates_(closedRows),
    baselineDateForAnyInsert: Utilities.formatDate(
      new Date(), ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy'
    ),
    validationErrorRecipient: PHASE4H_ERROR_EMAIL,
    sampleBaselineInserts: actions.filter(function(x) {
      return x.action === 'BASELINE_INSERT';
    }).slice(0, 10).map(function(x) {
      return { sourceType: x.source.sourceType, auditId: x.source.auditId,
        sourceRow: x.source.sourceRow };
    }),
    policies: [
      'Only current Power BI source Audit IDs are considered.',
      'Only Issues and Gate Checks rows whose final Status is Closed are considered.',
      'JSO and Out of Service rows are excluded.',
      'Rows match within Audit ID first, then one-to-one by similarity.',
      'Different vehicle, driver, division, or issue formatting does not force an insert.',
      'Any missing Closed rows found before activation are baseline inserts, dated today, and never emailed.',
      'Existing Closed rows and Emailed Dates are preserved.',
      'After activation, only newly detected Closed inserts may become email candidates.',
      'Operational emails go only to Assigned Driver with Supervisor CC.',
      'Validation/error notifications go only to ' + PHASE4H_ERROR_EMAIL + '.'
    ]
  };

  Logger.log('PHASE4H_SUMMARY=' + JSON.stringify(summary));
  SpreadsheetApp.getUi().alert([
    'Phase 4H completed — READ ONLY.',
    '',
    'Current Closed source rows: ' + summary.currentClosedSourceRows,
    'Matched existing Closed rows: ' + summary.matchedExistingClosedRows,
    'Proposed baseline inserts: ' + summary.proposedBaselineInserts,
    'Historical blank Emailed Dates: ' + summary.historicalBlankEmailedDates,
    'Proposed email candidates: 0',
    '',
    'Copy PHASE4H_SUMMARY from Executions → Logs.'
  ].join('\n'));
  return summary;
}

function phase4hCurrentPlans_() {
  const folder = DriveApp.getFolderById(PHASE4_INTAKE_FOLDER_ID);
  const runDate = Utilities.formatDate(new Date(), PHASE4_TIME_ZONE, 'MM-dd-yyyy');
  const dvirFile = phase4FindUniqueFile_(folder, ['DVIR_Details', 'DVIR_Details.xlsx']);
  const safetyFile = phase4FindUniqueFile_(folder, [
    'safety-deviations-' + runDate,
    'safety-deviations-' + runDate + '.xlsx'
  ]);
  const dvir = phase4RequireSheet_(phase4OpenGoogleSheet_(dvirFile, 'DVIR source'), 'DVIR_Details');
  const safety = phase4RequireSheet_(phase4OpenGoogleSheet_(safetyFile, 'Safety source'), 'Export');
  const plans = { Issues: {}, 'Gate Checks': {}, JSO: {} };
  phase4bReadDvir_(dvir, plans);
  phase4bReadSafety_(safety, plans);
  return plans;
}

function phase4hSimilarity_(source, closed) {
  let score = 0;
  for (let index = 0; index < 15; index++) {
    if (index === 1) continue;
    if (index === 3) {
      if (phase4dVehicleKey_(source[index]) === phase4dVehicleKey_(closed[index])) score += 8;
    } else if (index === 5) {
      if (phase4hNamesOverlap_(source[index], closed[index])) score += 8;
    } else if (index === 6) {
      if (phase4dIssueSignature_(source[index]) === phase4dIssueSignature_(closed[index])) score += 10;
    } else if (phase4hCell_(source[index]) === phase4hCell_(closed[index])) {
      score += 1;
    }
  }
  return score;
}

function phase4hEquivalent_(source, closed) {
  for (let index = 0; index < 15; index++) {
    if (index === 1 || index === 15) continue;
    if (index === 3) {
      if (phase4dVehicleKey_(source[index]) !== phase4dVehicleKey_(closed[index])) return false;
    } else if (index === 5) {
      if (!phase4hNamesOverlap_(source[index], closed[index])) return false;
    } else if (index === 6) {
      if (phase4dIssueSignature_(source[index]) !== phase4dIssueSignature_(closed[index])) return false;
    } else if (phase4hCell_(source[index]) !== phase4hCell_(closed[index])) return false;
  }
  return true;
}

function phase4hNamesOverlap_(left, right) {
  if (!String(left || '').trim() && !String(right || '').trim()) return true;
  const rightKeys = {};
  phase4dNameKeys_(right).forEach(function(key) { rightKeys[key] = true; });
  return phase4dNameKeys_(left).some(function(key) { return rightKeys[key]; });
}

function phase4hCell_(value) {
  return String(value == null ? '' : value).trim().replace(/\s+/g, ' ').toLowerCase();
}

function phase4hPad_(row, width) {
  const result = row.slice(0, width);
  while (result.length < width) result.push('');
  return result;
}

function phase4hBlankEmailedDates_(rows) {
  return rows.filter(function(row) {
    return String(row[1] || '').trim() && !String(row[15] || '').trim();
  }).length;
}
