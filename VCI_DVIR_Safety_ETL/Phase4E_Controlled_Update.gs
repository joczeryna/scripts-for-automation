/**
 * Phase 4E controlled-write checkpoint.
 * Requires Phase 4A, Phase 4B, and Phase 4D files in the same Apps Script project.
 *
 * Test safeguards:
 * - Creates a full spreadsheet backup before writes.
 * - Runs pre- and post-update validation.
 * - Does not update Closed, archive sources, or send email.
 */

const PHASE4E_TEST_MODE = true;

function runPhase4EControlledUpdate() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Another automation run is already active.');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const precheck = validateWorkbookIntegrity('UPDATE_PRECHECK');
    if (precheck.validation.status === 'BLOCKED') {
      throw new Error('Pre-update validation blocked the run: ' +
        JSON.stringify(precheck.validation.blockers));
    }

    const backup = phase4eCreateBackup_(ss);
    const context = phase4eBuildContext_(ss);
    const result = {
      backupId: backup.getId(),
      backupName: backup.getName(),
      Issues: phase4eApplyTarget_(ss.getSheetByName('Issues'), context.plans.Issues,
        context.equipment, context.locationDivisions),
      GateChecks: phase4eApplyTarget_(ss.getSheetByName('Gate Checks'),
        context.plans['Gate Checks'], context.equipment, context.locationDivisions),
      JSO: phase4eApplyTarget_(ss.getSheetByName('JSO'), context.plans.JSO,
        context.equipment, context.locationDivisions),
      legacyStatusNormalization: phase4eNormalizeLegacyStatuses_(ss),
      recipientIndex: phase4eRefreshRecipientIndex_(ss, context.plans),
      testMode: PHASE4E_TEST_MODE,
      closedUpdated: false,
      sourcesArchived: false,
      emailsSent: false
    };

    SpreadsheetApp.flush();
    const postcheck = validateWorkbookIntegrity('POST_UPDATE');
    result.postValidation = postcheck.validation;
    if (postcheck.validation.status === 'BLOCKED') {
      throw new Error('Post-update validation failed. Restore backup ' + backup.getName() +
        '. Findings: ' + JSON.stringify(postcheck.validation.blockers));
    }

    Logger.log('PHASE4E_RESULT=' + JSON.stringify(result));
    SpreadsheetApp.getUi().alert([
      'Phase 4E controlled update completed.',
      '',
      'Backup: ' + backup.getName(),
      'Issues — inserts: ' + result.Issues.inserts + ', updates: ' + result.Issues.updates,
      'Gate Checks — inserts: ' + result.GateChecks.inserts + ', updates: ' + result.GateChecks.updates,
      'JSO — inserts: ' + result.JSO.inserts + ', updates: ' + result.JSO.updates,
      'Workbook-wide Open cleared: ' + result.legacyStatusNormalization.openCleared,
      'Workbook-wide Pending Parts converted: ' +
        result.legacyStatusNormalization.pendingPartsConverted,
      'Recipient Index rows: ' + result.recipientIndex.rows,
      'Post-validation: ' + postcheck.validation.status,
      '',
      'TEST MODE: Closed was not updated; sources were not archived; no emails were sent.'
    ].join('\n'));
    return result;
  } finally {
    lock.releaseLock();
  }
}

function phase4eNormalizeLegacyStatuses_(ss) {
  const result = { openCleared: 0, pendingPartsConverted: 0 };
  ['Issues', 'Gate Checks'].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (sheet.getLastRow() < 2) return;
    const range = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1);
    const values = range.getDisplayValues();
    let changed = false;
    values.forEach(function(row) {
      const status = String(row[0] || '').trim();
      if (status === 'Open') {
        row[0] = '';
        result.openCleared++;
        changed = true;
      } else if (status === 'Pending Parts') {
        row[0] = 'Out of Service';
        result.pendingPartsConverted++;
        changed = true;
      }
    });
    if (changed) range.setValues(values);
  });
  return result;
}

function phase4eBuildContext_(ss) {
  const folder = DriveApp.getFolderById(PHASE4_INTAKE_FOLDER_ID);
  const runDate = Utilities.formatDate(new Date(), PHASE4_TIME_ZONE, 'MM-dd-yyyy');
  const dvirFile = phase4FindUniqueFile_(folder, ['DVIR_Details', 'DVIR_Details.xlsx']);
  const safetyFile = phase4FindUniqueFile_(folder, [
    'safety-deviations-' + runDate, 'safety-deviations-' + runDate + '.xlsx'
  ]);
  const dvir = phase4RequireSheet_(phase4OpenGoogleSheet_(dvirFile, 'DVIR source'), 'DVIR_Details');
  const safety = phase4RequireSheet_(phase4OpenGoogleSheet_(safetyFile, 'Safety source'), 'Export');
  const plans = { Issues: {}, 'Gate Checks': {}, JSO: {} };
  phase4bReadDvir_(dvir, plans);
  phase4bReadSafety_(safety, plans);
  phase4eEnrichSafetyPlans_(safety, plans);
  return {
    plans: plans,
    equipment: phase4eEquipmentMap_(ss.getSheetByName('Equipment')),
    locationDivisions: phase4eLocationDivisionMap_(ss.getSheetByName('Issues'), safety)
  };
}

function phase4eEnrichSafetyPlans_(sheet, plans) {
  const display = sheet.getDataRange().getDisplayValues();
  const raw = sheet.getDataRange().getValues();
  const header = phase4bHeaderMap_(display[0]);

  display.slice(1).forEach(function(row, index) {
    const auditId = String(row[header.auditid] || '').trim();
    const target = phase4bTarget_(row[header.auditname]);

    if (!auditId || !target || !plans[target][auditId]) return;

    const plan = plans[target][auditId];
    const sourceAuditDate = raw[index + 1][header.auditdate];

    // Do not replace a valid DVIR timestamp with a blank or date-only value.
    plan.auditDate = phase4eChooseBestAuditDate_(
      plan.auditDate,
      sourceAuditDate
    );

    if (!plan.location) {
      plan.location = String(row[header.location] || '').trim();
    }

    if (!plan.supervisor) {
      plan.supervisor = String(row[header.supervisor] || '').trim();
    }

    if (!plan.auditLink) {
      plan.auditLink = String(row[header.auditlink] || '').trim();
    }
  });
}

function phase4eApplyTarget_(sheet, plans, equipment, locationDivisions) {
  const name = sheet.getName();
  const isJso = name === 'JSO';
  const issueIndex = isJso ? 5 : 6;
  const statusIndex = isJso ? 8 : 9;
  const width = sheet.getLastColumn();
  const rowCount = Math.max(0, sheet.getLastRow() - 1);
  const values = rowCount ? sheet.getRange(2, 1, rowCount, width).getDisplayValues() : [];
  const groups = {};
  values.forEach(function(row, index) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!groups[auditId]) groups[auditId] = [];
    groups[auditId].push({ rowNumber: index + 2, row: row });
  });

  const result = { inserts: 0, updates: 0, statusChanges: 0,
    issueChanges: 0, openCleared: 0, pendingPartsConverted: 0 };
  const newPlans = [];
  Object.keys(plans).forEach(function(auditId) {
    const plan = plans[auditId];
    const matches = phase4eRelevantMatches_(groups[auditId] || [], plan, name);
    if (!matches.length) {
      newPlans.push(plan);
      return;
    }
    matches.forEach(function(match) {
      const currentIssue = String(match.row[issueIndex] || '').trim();
      const plannedIssue = phase4bIssueText_(plan);
      const currentStatus = String(match.row[statusIndex] || '').trim();
      const plannedStatus = phase4eStatus_(plan, currentStatus, equipment, isJso);
      let changed = false;
      if (phase4dIssueSignature_(currentIssue) !== phase4dIssueSignature_(plannedIssue)) {
        sheet.getRange(match.rowNumber, issueIndex + 1).setValue(plannedIssue);
        result.issueChanges++;
        changed = true;
      }
      if (currentStatus !== plannedStatus) {
        sheet.getRange(match.rowNumber, statusIndex + 1).setValue(plannedStatus);
        result.statusChanges++;
        if (currentStatus === 'Open' && plannedStatus === '') result.openCleared++;
        if (currentStatus === 'Pending Parts' && plannedStatus === 'Out of Service') {
          result.pendingPartsConverted++;
        }
        changed = true;
      }
      if (changed) result.updates++;
    });
  });

  let nextInsertRow = newPlans.length
    ? phase4eNextBlankRow_(sheet)
    : 0;

  newPlans.forEach(function(plan) {
    if (nextInsertRow > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 100);
    }

    const rowNumber = nextInsertRow++;
    const status = phase4eStatus_(plan, '', equipment, isJso);

    const templateRow = phase4eApplyRowTemplate_(
      sheet,
      rowNumber
    );

    if (isJso) {
      const row = [[
        plan.auditDate || '',
        phase4eAuditFormula_(plan),
        locationDivisions[phase4Normalize_(plan.location)] || '-1',
        (plan.assignedNames || []).join('\n'),
        plan.supervisor || '',
        phase4bIssueText_(plan),
        '',
        '',
        status
      ]];

      sheet
        .getRange(rowNumber, 1, 1, 9)
        .setValues(row);

      sheet
        .getRange(rowNumber, 8)
        .setFormula(
          '=WORKDAY(A' + rowNumber +
          ',5,Holidays!$A$2:$A$8)'
        );
    } else {
      /*
       * Write only the non-formula columns.
       *
       * Columns C, E, and I are copied from the preceding
       * template row and are never overwritten with values.
       */

      sheet
        .getRange(rowNumber, 1)
        .setValue(plan.auditDate || '');

      sheet
        .getRange(rowNumber, 2)
        .setFormula(phase4eAuditFormula_(plan));

      sheet
        .getRange(rowNumber, 4)
        .setValue(plan.vehicle || '');

      sheet
        .getRange(rowNumber, 6)
        .setValue((plan.assignedNames || []).join('\n'));

      sheet
        .getRange(rowNumber, 7)
        .setValue(phase4bIssueText_(plan));

      sheet
        .getRange(rowNumber, 8)
        .clearContent();

      sheet
        .getRange(rowNumber, 10)
        .setValue(status);

      if (!templateRow) {
        throw new Error(
          'No formula template row was found for ' +
          name +
          ' row ' +
          rowNumber +
          '.'
        );
      }

      // Copy formulas and automatically adjust their row references.
      [3, 5, 9].forEach(function(column) {
        sheet
          .getRange(templateRow, column)
          .copyTo(
            sheet.getRange(rowNumber, column),
            SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
            false
          );
      });
    }

    sheet
      .getRange(rowNumber, 1)
      .setNumberFormat('M/d/yyyy H:mm:ss');

    result.inserts++;
  });
  return result;
}

function phase4eRelevantMatches_(matches, plan, sheetName) {
  if (matches.length < 2) return matches;
  const vehicle = phase4dVehicleKey_(plan.vehicle);
  const names = {};
  (plan.assignedNames || []).forEach(function(value) {
    phase4dNameKeys_(value).forEach(function(key) { names[key] = true; });
  });
  let composite = matches.filter(function(match) {
    const vehicleMatch = sheetName === 'JSO' || !vehicle ||
      phase4dVehicleKey_(match.row[3]) === vehicle;
    const driverKeys = phase4dNameKeys_(match.row[5]);
    const driverMatch = !Object.keys(names).length ||
      driverKeys.some(function(key) { return names[key]; });
    return vehicleMatch && driverMatch;
  });
  if (composite.length) return composite;
  composite = matches.filter(function(match) {
    return sheetName !== 'JSO' && vehicle &&
      phase4dVehicleKey_(match.row[3]) === vehicle;
  });
  if (composite.length) return composite;
  throw new Error(sheetName + ' Audit ID ' + plan.auditId +
    ' has repeated rows but none match the source Vehicle and Assigned Driver.');
}

function phase4eStatus_(plan, currentStatus, equipment, isJso) {
  if (!isJso) {
    if (currentStatus === 'Pending Parts' || currentStatus === 'Out of Service') {
      return 'Out of Service';
    }
    if (equipment[phase4dVehicleKey_(plan.vehicle)] === 'OUT OF SERVICE') {
      return 'Out of Service';
    }
  }
  if ((plan.openQuestions || []).length === 0 &&
      (plan.resolvedQuestions || []).length > 0) return 'Closed';
  return '';
}

function phase4eRefreshRecipientIndex_(ss, plans) {
  const sheet = ss.getSheetByName('Automation Recipient Index');
  const directory = ss.getSheetByName('Company Directory').getDataRange().getDisplayValues();
  const header = phase4bHeaderMap_(directory[0]);
  const byWorker = {};
  const byName = {};
  directory.slice(1).forEach(function(row) {
    const record = { name: String(row[header.name] || '').trim(),
      workerNo: String(row[header.workerno] || '').trim(),
      email: String(row[header.emailaddress] || '').trim(),
      supervisor: String(row[header.supervisor] || '').trim() };
    if (record.workerNo) phase4dMapAdd_(byWorker, phase4Normalize_(record.workerNo), record);
    phase4dNameKeys_(record.name).forEach(function(key) { phase4dMapAdd_(byName, key, record); });
  });
  const now = new Date();
  const output = [];
  ['Issues', 'Gate Checks'].forEach(function(target) {
    Object.keys(plans[target]).sort().forEach(function(auditId) {
      const plan = plans[target][auditId];
      let matches = [];
      (plan.workerNos || []).filter(function(value) { return value && value !== '-1'; })
        .forEach(function(value) {
          (byWorker[phase4Normalize_(value)] || []).forEach(function(record) {
            if (matches.indexOf(record) === -1) matches.push(record);
          });
        });
      let source = 'Worker No';
      if (!matches.length) {
        source = 'Name Fallback';
        (plan.assignedNames || []).forEach(function(value) {
          phase4dNameKeys_(value).forEach(function(key) {
            (byName[key] || []).forEach(function(record) {
              if (matches.indexOf(record) === -1) matches.push(record);
            });
          });
        });
      }
      let status = 'Not Found';
      let notes = 'Employee could not be resolved; individualized email will be skipped.';
      let workerNo = (plan.workerNos || []).filter(function(value) { return value !== '-1'; })[0] || '';
      if (matches.length === 1) {
        workerNo = matches[0].workerNo;
        status = source === 'Worker No' ? 'Matched' : 'Name Fallback';
        notes = phase4dIsEmail_(matches[0].supervisor)
          ? 'Employee and direct supervisor email resolved.'
          : 'Employee resolved; supervisor email requires review.';
      } else if (matches.length > 1) {
        status = 'Ambiguous';
        notes = 'Multiple directory employees matched; individualized email will be skipped.';
      }
      output.push([target, auditId, workerNo, (plan.assignedNames || []).join('\n'),
        source, now, status, notes]);
    });
  });
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).clearContent();
  }
  if (output.length) sheet.getRange(2, 1, output.length, 8).setValues(output);
  return { rows: output.length,
    matched: output.filter(function(row) { return row[6] === 'Matched'; }).length,
    nameFallback: output.filter(function(row) { return row[6] === 'Name Fallback'; }).length,
    unresolved: output.filter(function(row) {
      return row[6] === 'Not Found' || row[6] === 'Ambiguous';
    }).length };
}

function phase4eEquipmentMap_(sheet) {
  const map = {};
  if (sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 5, sheet.getLastRow() - 1, 5).getDisplayValues().forEach(function(row) {
    const key = phase4dVehicleKey_(row[0]);
    if (key) map[key] = String(row[4] || '').trim().toUpperCase();
  });
  return map;
}

function phase4eLocationDivisionMap_(issuesSheet, safetySheet) {
  const issueRows = issuesSheet.getLastRow() > 1
    ? issuesSheet.getRange(2, 1, issuesSheet.getLastRow() - 1,
      issuesSheet.getLastColumn()).getDisplayValues() : [];
  const divisionByAudit = {};
  issueRows.forEach(function(row) {
    const auditId = String(row[1] || '').trim();
    const division = String(row[2] || '').trim();
    if (auditId && division && division !== '-1') divisionByAudit[auditId] = division;
  });
  const values = safetySheet.getDataRange().getDisplayValues();
  const header = phase4bHeaderMap_(values[0]);
  const counts = {};
  values.slice(1).forEach(function(row) {
    const location = phase4Normalize_(row[header.location]);
    const division = divisionByAudit[String(row[header.auditid] || '').trim()];
    if (!location || !division) return;
    if (!counts[location]) counts[location] = {};
    counts[location][division] = (counts[location][division] || 0) + 1;
  });
  const map = {};
  Object.keys(counts).forEach(function(location) {
    map[location] = Object.keys(counts[location]).sort(function(a, b) {
      return counts[location][b] - counts[location][a];
    })[0];
  });
  return map;
}

function phase4eNextBlankRow_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    const ids = sheet.getRange(2, 2, maxRows - 1, 1).getDisplayValues();
    for (let index = 0; index < ids.length; index++) {
      if (!String(ids[index][0] || '').trim()) return index + 2;
    }
  }
  sheet.insertRowsAfter(maxRows, 100);
  return maxRows + 1;
}

function phase4eAuditFormula_(plan) {
  const link = plan.auditLink ||
    'https://fieldaudit.certusview.com/Audit/Details/' + plan.auditId;
  return '=HYPERLINK("' + String(link).replace(/"/g, '""') + '","' +
    String(plan.auditId).replace(/"/g, '""') + '")';
}

function phase4eCreateBackup_(ss) {
  const source = DriveApp.getFileById(ss.getId());
  const parents = source.getParents();
  const stamp = Utilities.formatDate(new Date(), PHASE4_TIME_ZONE, 'yyyy-MM-dd HHmmss');
  const name = ss.getName() + ' - Phase4E Backup - ' + stamp;
  return parents.hasNext() ? source.makeCopy(name, parents.next()) : source.makeCopy(name);
}

function phase4eChooseBestAuditDate_(existingValue, candidateValue) {
  if (!phase4eHasAuditDate_(candidateValue)) return existingValue;
  if (!phase4eHasAuditDate_(existingValue)) return candidateValue;

  const existingHasTime = phase4eAuditDateHasTime_(existingValue);
  const candidateHasTime = phase4eAuditDateHasTime_(candidateValue);

  if (candidateHasTime && !existingHasTime) return candidateValue;

  return existingValue;
}

function phase4eHasAuditDate_(value) {
  if (value instanceof Date) return !isNaN(value.getTime());
  return String(value == null ? '' : value).trim() !== '';
}

function phase4eAuditDateHasTime_(value) {
  const date = value instanceof Date ? value : new Date(value);

  if (isNaN(date.getTime())) {
    return /\d{1,2}:\d{2}/.test(String(value || ''));
  }

  return (
    date.getHours() !== 0 ||
    date.getMinutes() !== 0 ||
    date.getSeconds() !== 0 ||
    date.getMilliseconds() !== 0
  );
}

function phase4eApplyRowTemplate_(sheet, rowNumber) {
  const width = sheet.getLastColumn();
  let templateRow = 0;

  // Find the closest populated row above the new row.
  for (let row = rowNumber - 1; row >= 2; row--) {
    if (String(sheet.getRange(row, 2).getDisplayValue()).trim()) {
      templateRow = row;
      break;
    }
  }

  // If necessary, find a populated row below it.
  if (!templateRow) {
    for (let row = rowNumber + 1; row <= sheet.getLastRow(); row++) {
      if (String(sheet.getRange(row, 2).getDisplayValue()).trim()) {
        templateRow = row;
        break;
      }
    }
  }

  if (!templateRow) return 0;

  const template = sheet.getRange(templateRow, 1, 1, width);
  const destination = sheet.getRange(rowNumber, 1, 1, width);

  template.copyTo(
    destination,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );

  template.copyTo(
    destination,
    SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
    false
  );
  return templateRow;
}
