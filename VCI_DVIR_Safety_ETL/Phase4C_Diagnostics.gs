/**
 * Phase 4C: read-only multi-column duplicate and mapping diagnostics.
 * Requires Phase4A_Preflight.gs and Phase4B_Dry_Run.gs in the same project.
 */

function diagnoseConsolidatedMappings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  const report = {
    mode: 'READ ONLY - NO WRITES OR ARCHIVING',
    repeatedAuditAnalysis: {},
    issueDifferenceAnalysis: {},
    supervisorLookupAnalysis: phase4cSupervisorAnalysis_(
      phase4RequireSheet_(ss, 'Company Directory'), plans
    ),
    equipmentLookupAnalysis: phase4cEquipmentAnalysis_(
      phase4RequireSheet_(ss, 'Equipment'), plans
    )
  };

  ['Issues', 'Gate Checks', 'JSO'].forEach(function(name) {
    const sheet = phase4RequireSheet_(ss, name);
    report.repeatedAuditAnalysis[name] = phase4cRepeatedAudits_(sheet);
    report.issueDifferenceAnalysis[name] = phase4cIssueDifferences_(sheet, plans[name]);
  });

  const compactReport = JSON.parse(JSON.stringify(report));
  Object.keys(compactReport.repeatedAuditAnalysis).forEach(function(name) {
    delete compactReport.repeatedAuditAnalysis[name].repeatedGroupsSample;
  });
  Object.keys(compactReport.issueDifferenceAnalysis).forEach(function(name) {
    delete compactReport.issueDifferenceAnalysis[name].substantiveSample;
  });
  delete compactReport.supervisorLookupAnalysis.notFoundSample;
  delete compactReport.supervisorLookupAnalysis.ambiguousSample;
  delete compactReport.equipmentLookupAnalysis.unmatchedVehicleSample;
  Logger.log('PHASE4C_SUMMARY=' + JSON.stringify(compactReport));
  Logger.log('PHASE4C_REPEATED_SAMPLES=' + JSON.stringify(
    report.repeatedAuditAnalysis.Issues.repeatedGroupsSample
  ));
  Logger.log('PHASE4C_ISSUE_SAMPLES=' + JSON.stringify({
    Issues: report.issueDifferenceAnalysis.Issues.substantiveSample,
    GateChecks: report.issueDifferenceAnalysis['Gate Checks'].substantiveSample,
    JSO: report.issueDifferenceAnalysis.JSO.substantiveSample
  }));
  Logger.log('PHASE4C_SUPERVISOR_SAMPLES=' + JSON.stringify({
    notFound: report.supervisorLookupAnalysis.notFoundSample,
    ambiguous: report.supervisorLookupAnalysis.ambiguousSample
  }));
  Logger.log('PHASE4C_EQUIPMENT_SAMPLES=' + JSON.stringify(
    report.equipmentLookupAnalysis.unmatchedVehicleSample
  ));
  const issues = report.repeatedAuditAnalysis.Issues;
  const differences = report.issueDifferenceAnalysis.Issues;
  const supervisors = report.supervisorLookupAnalysis;
  const equipment = report.equipmentLookupAnalysis;
  SpreadsheetApp.getUi().alert([
    'Phase 4C diagnostics completed — NO DATA WAS WRITTEN.',
    '',
    'Issues repeated Audit IDs: ' + issues.repeatedAuditIds,
    'Exact duplicate rows: ' + issues.exactDuplicateRows,
    'Same Audit ID with different issues: ' + issues.auditIdsWithDifferentIssues,
    'Same Audit ID with different drivers/vehicles: ' +
      issues.auditIdsWithDifferentDriversOrVehicles,
    '',
    'Issues comparisons: ' + differences.exactTextMatch + ' exact, ' +
      differences.formatOrOrderOnly + ' format/order only, ' +
      differences.substantiveDifference + ' substantive.',
    '',
    'Supervisor lookup: ' + supervisors.exactMatch + ' exact, ' +
      supervisors.flexibleMatch + ' flexible, ' + supervisors.ambiguous +
      ' ambiguous, ' + supervisors.notFound + ' not found.',
    '',
    'Equipment: ' + equipment.plannedVehiclesMatched + ' planned vehicles matched; ' +
      equipment.outOfServiceMatches + ' are Out of Service.',
    '',
    'Open Executions → Logs for row numbers and samples.'
  ].join('\n'));
  return report;
}

function phase4cRepeatedAudits_(sheet) {
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
    : [];
  const groups = {};
  values.forEach(function(row, index) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!groups[auditId]) groups[auditId] = [];
    groups[auditId].push({ rowNumber: index + 2, values: row });
  });

  const result = {
    repeatedAuditIds: 0,
    rowsWithinRepeatedAuditIds: 0,
    exactDuplicateRows: 0,
    auditIdsContainingExactDuplicates: 0,
    auditIdsWithDifferentIssues: 0,
    auditIdsWithDifferentDriversOrVehicles: 0,
    differingColumnFrequency: {},
    repeatedGroupsSample: []
  };

  Object.keys(groups).forEach(function(auditId) {
    const rows = groups[auditId];
    if (rows.length < 2) return;
    result.repeatedAuditIds++;
    result.rowsWithinRepeatedAuditIds += rows.length;

    const fullSignatures = {};
    const issues = {};
    const drivers = {};
    const vehicles = {};
    const statuses = {};
    rows.forEach(function(item) {
      const row = item.values;
      const full = row.map(phase4cCell_).join('|');
      fullSignatures[full] = (fullSignatures[full] || 0) + 1;
      issues[phase4cIssueSignature_(row[phase4cIssueIndex_(sheet.getName())])] = true;
      drivers[phase4cCell_(row[5])] = true;
      vehicles[phase4cCell_(row[3])] = true;
      statuses[phase4cCell_(row[phase4cStatusIndex_(sheet.getName())])] = true;
    });
    let groupExactDuplicates = 0;
    Object.keys(fullSignatures).forEach(function(signature) {
      if (fullSignatures[signature] > 1) groupExactDuplicates += fullSignatures[signature] - 1;
    });
    if (groupExactDuplicates > 0) {
      result.auditIdsContainingExactDuplicates++;
      result.exactDuplicateRows += groupExactDuplicates;
    }
    const differentIssues = Object.keys(issues).length > 1;
    const differentIdentity = Object.keys(drivers).length > 1 || Object.keys(vehicles).length > 1;
    if (differentIssues) result.auditIdsWithDifferentIssues++;
    if (differentIdentity) result.auditIdsWithDifferentDriversOrVehicles++;

    const differingColumns = [];
    const width = Math.max.apply(null, rows.map(function(item) { return item.values.length; }));
    for (let columnIndex = 0; columnIndex < width; columnIndex++) {
      const distinct = phase4cDistinct_(rows, columnIndex);
      if (distinct.length > 1) {
        const header = String(sheet.getRange(1, columnIndex + 1).getDisplayValue() ||
          'Column ' + (columnIndex + 1)).trim();
        differingColumns.push(header);
        result.differingColumnFrequency[header] =
          (result.differingColumnFrequency[header] || 0) + 1;
      }
    }

    if (result.repeatedGroupsSample.length < 5) {
      result.repeatedGroupsSample.push({
        auditId: auditId,
        rowNumbers: rows.map(function(item) { return item.rowNumber; }),
        exactDuplicateRowsBeyondFirst: groupExactDuplicates,
        differingColumns: differingColumns,
        dates: phase4cDistinct_(rows, 0),
        divisions: phase4cDistinct_(rows, 2),
        vehicles: phase4cDistinct_(rows, 3),
        assignedDrivers: phase4cDistinct_(rows, 5),
        issues: phase4cDistinct_(rows, phase4cIssueIndex_(sheet.getName())),
        comments: phase4cDistinct_(rows, 7),
        statuses: phase4cDistinct_(rows, phase4cStatusIndex_(sheet.getName()))
      });
    }
  });
  return result;
}

function phase4cIssueDifferences_(sheet, plans) {
  const issueIndex = phase4cIssueIndex_(sheet.getName());
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
    : [];
  const existing = {};
  values.forEach(function(row) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!existing[auditId]) existing[auditId] = [];
    existing[auditId].push(String(row[issueIndex] || ''));
  });
  const result = {
    comparedExistingAuditIds: 0,
    exactTextMatch: 0,
    formatOrOrderOnly: 0,
    substantiveDifference: 0,
    missingDestinationAuditId: 0,
    substantiveSample: []
  };
  Object.keys(plans).forEach(function(auditId) {
    if (!existing[auditId]) {
      result.missingDestinationAuditId++;
      return;
    }
    result.comparedExistingAuditIds++;
    const plannedText = phase4bIssueText_(plans[auditId]);
    const currentText = existing[auditId].join('\n');
    if (currentText.trim() === plannedText.trim()) {
      result.exactTextMatch++;
    } else if (phase4cIssueSignature_(currentText) === phase4cIssueSignature_(plannedText)) {
      result.formatOrOrderOnly++;
    } else {
      result.substantiveDifference++;
      if (result.substantiveSample.length < 5) {
        result.substantiveSample.push({
          auditId: auditId,
          currentIssues: currentText,
          plannedIssues: plannedText
        });
      }
    }
  });
  return result;
}

function phase4cSupervisorAnalysis_(sheet, plans) {
  const data = sheet.getDataRange().getDisplayValues();
  const header = phase4bHeaderMap_(data[0]);
  const byWorker = {};
  const byExactName = {};
  const byFlexibleName = {};
  data.slice(1).forEach(function(row) {
    const record = {
      name: String(row[header.name] || '').trim(),
      workerNo: String(row[header.workerno] || '').trim(),
      email: String(row[header.emailaddress] || '').trim(),
      supervisor: String(row[header.supervisor] || '').trim()
    };
    if (record.workerNo) phase4cAddMap_(byWorker, phase4Normalize_(record.workerNo), record);
    if (record.name) {
      phase4cAddMap_(byExactName, phase4Normalize_(record.name), record);
      phase4cNameKeys_(record.name).forEach(function(key) {
        phase4cAddMap_(byFlexibleName, key, record);
      });
    }
  });
  const result = {
    employeesResolved: 0, exactMatch: 0, flexibleMatch: 0,
    ambiguous: 0, notFound: 0, supervisorMissingEmail: 0,
    notFoundSample: [], ambiguousSample: []
  };
  ['Issues', 'Gate Checks'].forEach(function(target) {
    Object.keys(plans[target]).forEach(function(auditId) {
      const plan = plans[target][auditId];
      let employees = [];
      (plan.workerNos || []).filter(function(value) { return value && value !== '-1'; })
        .forEach(function(value) {
          (byWorker[phase4Normalize_(value)] || []).forEach(function(record) {
            if (employees.indexOf(record) === -1) employees.push(record);
          });
        });
      if (employees.length !== 1) return;
      result.employeesResolved++;
      const employee = employees[0];
      const exact = byExactName[phase4Normalize_(employee.supervisor)] || [];
      if (exact.length === 1) {
        result.exactMatch++;
        if (!exact[0].email) result.supervisorMissingEmail++;
        return;
      }
      let flexible = [];
      phase4cNameKeys_(employee.supervisor).forEach(function(key) {
        (byFlexibleName[key] || []).forEach(function(record) {
          if (flexible.indexOf(record) === -1) flexible.push(record);
        });
      });
      if (flexible.length === 1) {
        result.flexibleMatch++;
        if (!flexible[0].email) result.supervisorMissingEmail++;
      } else if (flexible.length > 1) {
        result.ambiguous++;
        if (result.ambiguousSample.length < 5) {
          result.ambiguousSample.push({ auditId: auditId, employee: employee.name,
            supervisorValue: employee.supervisor,
            candidates: flexible.map(function(record) { return record.name; }) });
        }
      } else {
        result.notFound++;
        if (result.notFoundSample.length < 8) {
          result.notFoundSample.push({ auditId: auditId, employee: employee.name,
            supervisorValue: employee.supervisor,
            generatedKeys: phase4cNameKeys_(employee.supervisor) });
        }
      }
    });
  });
  return result;
}

function phase4cEquipmentAnalysis_(sheet, plans) {
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 5, sheet.getLastRow() - 1, 5).getDisplayValues()
    : [];
  const equipment = {};
  const statuses = {};
  rows.forEach(function(row) {
    const vehicle = phase4Normalize_(row[0]);
    const status = String(row[4] || '').trim() || '<blank>';
    statuses[status] = (statuses[status] || 0) + 1;
    if (vehicle) equipment[vehicle] = status;
  });
  const result = {
    equipmentRows: rows.length,
    statusDistribution: statuses,
    uniquePlannedVehicles: 0,
    plannedVehiclesMatched: 0,
    plannedVehiclesNotFound: 0,
    outOfServiceMatches: 0,
    unmatchedVehicleSample: []
  };
  const plannedVehicles = {};
  ['Issues', 'Gate Checks'].forEach(function(target) {
    Object.keys(plans[target]).forEach(function(auditId) {
      const vehicle = String(plans[target][auditId].vehicle || '').trim();
      if (vehicle) plannedVehicles[phase4Normalize_(vehicle)] = vehicle;
    });
  });
  Object.keys(plannedVehicles).forEach(function(key) {
    result.uniquePlannedVehicles++;
    if (equipment[key] == null) {
      result.plannedVehiclesNotFound++;
      if (result.unmatchedVehicleSample.length < 8) {
        result.unmatchedVehicleSample.push(plannedVehicles[key]);
      }
    } else {
      result.plannedVehiclesMatched++;
      if (String(equipment[key]).trim().toUpperCase() === 'OUT OF SERVICE') {
        result.outOfServiceMatches++;
      }
    }
  });
  return result;
}

function phase4cNameKeys_(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return [];
  const keys = {};
  keys[phase4Normalize_(raw)] = true;
  const comma = raw.split(',');
  if (comma.length >= 2) {
    keys[phase4Normalize_(comma.slice(1).join(' ') + ' ' + comma[0])] = true;
  }
  const words = raw.replace(/[^a-z0-9 ]+/g, ' ').split(/\s+/).filter(Boolean)
    .filter(function(word) { return ['jr', 'sr', 'ii', 'iii', 'iv'].indexOf(word) === -1; });
  if (words.length >= 2) {
    keys[words[0] + words[words.length - 1]] = true;
    keys[words[words.length - 1] + words[0]] = true;
  }
  return Object.keys(keys);
}

function phase4cAddMap_(map, key, value) {
  if (!key) return;
  if (!map[key]) map[key] = [];
  if (map[key].indexOf(value) === -1) map[key].push(value);
}

function phase4cIssueSignature_(value) {
  const tokens = String(value || '')
    .replace(/RESOLVED\s*:/gi, '\n')
    .split(/\r?\n/)
    .map(function(item) { return phase4Normalize_(item); })
    .filter(Boolean);
  const unique = {};
  tokens.forEach(function(token) { unique[token] = true; });
  return Object.keys(unique).sort().join('|');
}

function phase4cDistinct_(rows, columnIndex) {
  const values = {};
  rows.forEach(function(item) {
    const value = String(item.values[columnIndex] || '').trim();
    values[value || '<blank>'] = true;
  });
  return Object.keys(values);
}

function phase4cCell_(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
}

function phase4cIssueIndex_(sheetName) {
  return sheetName === 'JSO' ? 5 : 6;
}

function phase4cStatusIndex_(sheetName) {
  return sheetName === 'JSO' ? 8 : 9;
}
