/**
 * Phase 4I: controlled Closed baseline write.
 *
 * Requires Phase 4H plus Phase 4A/4B/4D helpers.
 * Creates a full backup, appends only current Closed source rows that have no
 * one-to-one Audit ID match in Closed, and stamps the append with today's
 * placeholder Emailed Date. Sends no emails and archives nothing.
 */

function runPhase4IClosedBaselineWrite() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) throw new Error('Another automation run is already active.');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const base = validateWorkbookIntegrity('PHASE4I_CLOSED_BASELINE_PREFLIGHT');
    const validation = base.validation || base;
    if ((validation.blockers || []).length || validation.status === 'BLOCKED') {
      throw new Error('Phase 4I stopped: workbook validation is BLOCKED. Nothing was written.');
    }

    const closed = phase4RequireSheet_(ss, 'Closed');
    if (closed.getLastColumn() < 16) {
      throw new Error('Closed must contain at least 16 columns, with Emailed Date in P.');
    }

    const currentPlans = phase4hCurrentPlans_();
    const sourceRows = phase4iCurrentClosedRows_(ss, currentPlans);
    const existingRows = closed.getLastRow() > 1
      ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];
    const existingByAudit = {};
    existingRows.forEach(function(row, index) {
      const auditId = String(row[1] || '').trim();
      if (!auditId) return;
      if (!existingByAudit[auditId]) existingByAudit[auditId] = [];
      existingByAudit[auditId].push({ row: index + 2, values: row, used: false });
    });

    const missing = [];
    let matched = 0;
    sourceRows.forEach(function(source) {
      const candidates = (existingByAudit[source.auditId] || []).filter(function(item) {
        return !item.used;
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
      if (best) {
        best.used = true;
        matched++;
      } else {
        missing.push(source);
      }
    });

    const backup = ss.copy(
      ss.getName() + ' - Phase4I Closed Baseline Backup - ' +
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmmss')
    );
    const baselineDate = phase4iTodayAtNoon_(ss.getSpreadsheetTimeZone());

    if (missing.length) {
      const startRow = closed.getLastRow() + 1;
      closed.insertRowsAfter(closed.getLastRow(), missing.length);
      const output = missing.map(function(source) {
        return phase4hPad_(source.values, 15).concat([baselineDate]);
      });
      closed.getRange(startRow, 1, output.length, 16).setValues(output);
      if (startRow > 2) {
        closed.getRange(startRow - 1, 1, 1, 16).copyTo(
          closed.getRange(startRow, 1, output.length, 16),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
      }
      closed.getRange(startRow, 16, output.length, 1).setNumberFormat('MM/dd/yyyy');
      SpreadsheetApp.flush();
    }

    const verifyRows = closed.getLastRow() > 1
      ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];
    const availableCounts = {};
    verifyRows.forEach(function(row) {
      const auditId = String(row[1] || '').trim();
      if (auditId) availableCounts[auditId] = (availableCounts[auditId] || 0) + 1;
    });
    const requiredCounts = {};
    sourceRows.forEach(function(source) {
      requiredCounts[source.auditId] = (requiredCounts[source.auditId] || 0) + 1;
    });
    const shortages = [];
    Object.keys(requiredCounts).forEach(function(auditId) {
      if ((availableCounts[auditId] || 0) < requiredCounts[auditId]) {
        shortages.push({ auditId: auditId, required: requiredCounts[auditId],
          available: availableCounts[auditId] || 0 });
      }
    });
    const appendedStart = verifyRows.length - missing.length;
    const blankAppendedDates = missing.length
      ? verifyRows.slice(appendedStart).filter(function(row) {
          return !String(row[15] || '').trim();
        }).length : 0;
    if (shortages.length || blankAppendedDates) {
      throw new Error('Phase 4I post-write verification failed: ' + JSON.stringify({
        shortages: shortages.slice(0, 20), blankAppendedDates: blankAppendedDates
      }));
    }

    const post = validateWorkbookIntegrity('PHASE4I_CLOSED_BASELINE_POSTWRITE');
    const postValidation = post.validation || post;
    if ((postValidation.blockers || []).length || postValidation.status === 'BLOCKED') {
      throw new Error(
        'Phase 4I wrote the baseline but post-write validation is BLOCKED. ' +
        'Restore backup ' + backup.getName() + ' before continuing.'
      );
    }

    const result = {
      mode: 'CONTROLLED CLOSED BASELINE WRITE - NO EMAILS OR ARCHIVING',
      backupId: backup.getId(),
      backupName: backup.getName(),
      baselineDate: Utilities.formatDate(
        baselineDate, ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy'
      ),
      currentClosedSourceRows: sourceRows.length,
      matchedExistingRows: matched,
      baselineRowsAppended: missing.length,
      placeholderDatesWritten: missing.length,
      proposedEmailCandidates: 0,
      verificationShortages: 0,
      blankAppendedEmailedDates: 0,
      postValidationStatus: postValidation.status,
      validationErrorRecipient: PHASE4H_ERROR_EMAIL,
      emailsSent: false,
      sourcesArchived: false
    };
    Logger.log('PHASE4I_RESULT=' + JSON.stringify(result));
    SpreadsheetApp.getUi().alert([
      'Phase 4I Closed baseline completed.',
      '',
      'Rows appended: ' + missing.length,
      'Placeholder dates written: ' + missing.length,
      'Email candidates: 0',
      'Post-validation: ' + postValidation.status,
      '',
      'No emails were sent. No files were archived.',
      'Backup: ' + backup.getName()
    ].join('\n'));
    return result;
  } finally {
    lock.releaseLock();
  }
}

function phase4iCurrentClosedRows_(ss, plans) {
  const currentKeys = {};
  ['Issues', 'Gate Checks'].forEach(function(sourceType) {
    Object.keys(plans[sourceType]).forEach(function(auditId) {
      currentKeys[sourceType + '|' + auditId] = true;
    });
  });
  const rows = [];
  ['Issues', 'Gate Checks'].forEach(function(sourceType) {
    const sheet = phase4RequireSheet_(ss, sourceType);
    if (sheet.getLastRow() < 2) return;
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1,
      Math.max(16, sheet.getLastColumn())).getDisplayValues();
    values.forEach(function(row, index) {
      const auditId = String(row[1] || '').trim();
      if (!auditId || !currentKeys[sourceType + '|' + auditId]) return;
      if (String(row[9] || '').trim() !== 'Closed') return;
      rows.push({ sourceType: sourceType, sourceRow: index + 2,
        auditId: auditId, values: phase4hPad_(row, 15) });
    });
  });
  return rows;
}

function phase4iTodayAtNoon_(timeZone) {
  const parts = Utilities.formatDate(new Date(), timeZone, 'yyyy,MM,dd')
    .split(',').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0, 0);
}
