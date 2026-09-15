/**
 * Phase 4D: reconciled dry run plus reusable workbook-integrity validator.
 * READ ONLY in this checkpoint.
 * Requires the Phase 4A and Phase 4B script files in the same project.
 */

function previewPhase4D() {
  const report = validateWorkbookIntegrity('PREVIEW');
  Logger.log('PHASE4D_SUMMARY=' + JSON.stringify(report));
  SpreadsheetApp.getUi().alert([
    'Phase 4D completed — NO DATA WAS WRITTEN.',
    '',
    'Validation: ' + report.validation.status,
    'Blockers: ' + report.validation.blockers.length,
    'Warnings: ' + report.validation.warnings.length,
    '',
    'Semantic issue comparison: ' + report.issueComparison.exact + ' exact, ' +
      report.issueComparison.equivalent + ' equivalent, ' +
      report.issueComparison.substantive + ' substantive.',
    'Repeated Audit-ID groups safely matched: ' + report.compositeMatching.matchedGroups,
    'Repeated groups requiring review: ' + report.compositeMatching.blockedGroups,
    '',
    'Recipients: ' + report.recipients.employeeMatched + ' employee matched; ' +
      report.recipients.directSupervisorEmail + ' direct supervisor emails; ' +
      report.recipients.unresolved + ' unresolved.',
    '',
    'Equipment: ' + report.equipment.matched + ' matched, ' +
      report.equipment.notFound + ' not found, ' +
      report.equipment.outOfService + ' Out of Service.',
    '',
    'Copy PHASE4D_SUMMARY from Executions → Logs.'
  ].join('\n'));
  return report;
}

/**
 * Reusable validation gate. Later phases will call this after Update Report and
 * again immediately before email. Phase 4D only calls PREVIEW and performs no writes.
 */
function validateWorkbookIntegrity(mode) {
  mode = mode || 'PREVIEW';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const validation = { status: 'PASS', blockers: [], warnings: [], passedChecks: [] };
  let plans = { Issues: {}, 'Gate Checks': {}, JSO: {} };

  try {
    const folder = DriveApp.getFolderById(PHASE4_INTAKE_FOLDER_ID);
    const runDate = Utilities.formatDate(new Date(), PHASE4_TIME_ZONE, 'MM-dd-yyyy');
    const dvirFile = phase4FindUniqueFile_(folder, ['DVIR_Details', 'DVIR_Details.xlsx']);
    const safetyFile = phase4FindUniqueFile_(folder, [
      'safety-deviations-' + runDate,
      'safety-deviations-' + runDate + '.xlsx'
    ]);
    const dvir = phase4RequireSheet_(phase4OpenGoogleSheet_(dvirFile, 'DVIR source'), 'DVIR_Details');
    const safety = phase4RequireSheet_(phase4OpenGoogleSheet_(safetyFile, 'Safety source'), 'Export');
    phase4ValidateDvir_(dvir);
    phase4ValidateSafety_(safety);
    phase4bReadDvir_(dvir, plans);
    phase4bReadSafety_(safety, plans);
    validation.passedChecks.push('Both source workbooks and worksheets are valid.');
  } catch (error) {
    phase4dFinding_(validation.blockers, 'SOURCES', '', '', error.message);
  }

  const requiredSheets = ['Issues', 'Gate Checks', 'JSO', 'Closed', 'Equipment',
    'Company Directory', 'Automation Recipient Index'];
  requiredSheets.forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      phase4dFinding_(validation.blockers, name, '', '', 'Required worksheet is missing.');
    }
  });
  if (validation.blockers.length) {
    validation.status = 'BLOCKED';
    return phase4dEmptyReport_(mode, validation);
  }

  phase4dValidateHeadersAndProtection_(ss, validation);
  const issueComparison = phase4dIssueComparison_(ss, plans);
  const compositeMatching = phase4dCompositeMatching_(ss, plans, validation);
  const recipients = phase4dRecipients_(ss.getSheetByName('Company Directory'), plans, validation);
  const equipment = phase4dEquipment_(ss.getSheetByName('Equipment'), plans, validation);
  phase4dValidateStatuses_(ss, validation);

  if (validation.blockers.length) validation.status = 'BLOCKED';
  else if (validation.warnings.length) validation.status = 'WARNING';

  return {
    mode: mode,
    readOnly: true,
    validation: validation,
    issueComparison: issueComparison,
    compositeMatching: compositeMatching,
    recipients: recipients,
    equipment: equipment,
    policies: [
      'Repeated Audit IDs are preserved; no row is deleted or collapsed.',
      'Repeated groups match by Audit ID + normalized Vehicle + Assigned Driver.',
      'Formatting, ordering, and approved issue aliases do not trigger updates.',
      'Purely numeric vehicle IDs ignore leading zeros for Equipment matching.',
      'A Supervisor value formatted as an email is used directly.',
      'The validator will run after Update Report and again before email in a later phase.'
    ]
  };
}

function phase4dValidateHeadersAndProtection_(ss, validation) {
  try {
    phase4RequireHeaders_(ss.getSheetByName('Automation Recipient Index'), 1, [
      'Source Type', 'Audit ID', 'Worker No', 'Assigned Driver',
      'Match Source', 'Last Updated', 'Match Status', 'Notes'
    ]);
    const protections = ss.getSheetByName('Automation Recipient Index')
      .getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (!protections.length || !protections.some(function(item) { return item.canEdit(); })) {
      throw new Error('Protected sheet is not editable by the executing account.');
    }
    validation.passedChecks.push('Automation Recipient Index headers and protection are valid.');
  } catch (error) {
    phase4dFinding_(validation.blockers, 'Automation Recipient Index', 1, '', error.message);
  }
  try {
    phase4RequireHeaderSet_(ss.getSheetByName('Company Directory'), 1, [
      'Name', 'Worker No', 'Employee AD', 'Email Address', 'Worker Status', 'Supervisor'
    ]);
    validation.passedChecks.push('Company Directory headers are valid.');
  } catch (error) {
    phase4dFinding_(validation.blockers, 'Company Directory', 1, '', error.message);
  }
}

function phase4dIssueComparison_(ss, plans) {
  const result = { exact: 0, equivalent: 0, substantive: 0, missing: 0, sample: [] };
  ['Issues', 'Gate Checks', 'JSO'].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    const issueIndex = name === 'JSO' ? 5 : 6;
    const rows = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
      : [];
    const existing = {};
    rows.forEach(function(row) {
      const auditId = String(row[1] || '').trim();
      if (!auditId) return;
      if (!existing[auditId]) existing[auditId] = [];
      existing[auditId].push(String(row[issueIndex] || ''));
    });
    Object.keys(plans[name]).forEach(function(auditId) {
      if (!existing[auditId]) {
        result.missing++;
        return;
      }
      const current = existing[auditId].join('\n');
      const planned = phase4bIssueText_(plans[name][auditId]);
      if (current.trim() === planned.trim()) result.exact++;
      else if (phase4dIssueSignature_(current) === phase4dIssueSignature_(planned)) {
        result.equivalent++;
      } else {
        result.substantive++;
        if (result.sample.length < 8) {
          result.sample.push({ sheet: name, auditId: auditId,
            currentIssues: current, plannedIssues: planned });
        }
      }
    });
  });
  return result;
}

function phase4dCompositeMatching_(ss, plans, validation) {
  const result = { repeatedGroups: 0, matchedGroups: 0, blockedGroups: 0, sample: [] };
  ['Issues', 'Gate Checks', 'JSO'].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    const rows = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
      : [];
    const grouped = {};
    rows.forEach(function(row, index) {
      const auditId = String(row[1] || '').trim();
      if (!auditId) return;
      if (!grouped[auditId]) grouped[auditId] = [];
      grouped[auditId].push({ rowNumber: index + 2, vehicle: row[3], driver: row[5] });
    });
    Object.keys(grouped).forEach(function(auditId) {
      const group = grouped[auditId];
      if (group.length < 2 || !plans[name][auditId]) return;
      result.repeatedGroups++;
      const plan = plans[name][auditId];
      const planVehicle = phase4dVehicleKey_(plan.vehicle);
      const planNameKeys = {};
      (plan.assignedNames || []).forEach(function(value) {
        phase4dNameKeys_(value).forEach(function(key) { planNameKeys[key] = true; });
      });
      let matches = group.filter(function(item) {
        const vehicleMatch = !planVehicle || phase4dVehicleKey_(item.vehicle) === planVehicle;
        const driverKeys = phase4dNameKeys_(item.driver);
        const driverMatch = !Object.keys(planNameKeys).length ||
          driverKeys.some(function(key) { return planNameKeys[key]; });
        return vehicleMatch && driverMatch;
      });
      if (!matches.length) {
        matches = group.filter(function(item) {
          return planVehicle && phase4dVehicleKey_(item.vehicle) === planVehicle;
        });
      }
      if (matches.length) {
        result.matchedGroups++;
      } else {
        result.blockedGroups++;
        phase4dFinding_(validation.blockers, name, '', auditId,
          'Repeated Audit ID could not be reconciled by Vehicle and Assigned Driver.');
        if (result.sample.length < 8) result.sample.push({ sheet: name,
          auditId: auditId, sourceVehicle: plan.vehicle,
          sourceDrivers: plan.assignedNames, destinationRows: group });
      }
    });
  });
  return result;
}

function phase4dRecipients_(sheet, plans, validation) {
  const data = sheet.getDataRange().getDisplayValues();
  const header = phase4bHeaderMap_(data[0]);
  const byWorker = {};
  const byName = {};
  data.slice(1).forEach(function(row) {
    const record = { name: String(row[header.name] || '').trim(),
      workerNo: String(row[header.workerno] || '').trim(),
      email: String(row[header.emailaddress] || '').trim(),
      supervisor: String(row[header.supervisor] || '').trim() };
    if (record.workerNo) phase4dMapAdd_(byWorker, phase4Normalize_(record.workerNo), record);
    phase4dNameKeys_(record.name).forEach(function(key) { phase4dMapAdd_(byName, key, record); });
  });
  const result = { total: 0, employeeMatched: 0, workerNoMatched: 0,
    nameFallback: 0, directSupervisorEmail: 0, supervisorNameMatched: 0,
    invalidEmployeeEmail: 0, unresolved: 0, sample: [] };
  ['Issues', 'Gate Checks'].forEach(function(name) {
    Object.keys(plans[name]).forEach(function(auditId) {
      result.total++;
      const plan = plans[name][auditId];
      let employees = [];
      (plan.workerNos || []).filter(function(value) { return value && value !== '-1'; })
        .forEach(function(value) {
          (byWorker[phase4Normalize_(value)] || []).forEach(function(record) {
            if (employees.indexOf(record) === -1) employees.push(record);
          });
        });
      let matchSource = 'Worker No';
      if (!employees.length) {
        matchSource = 'Name Fallback';
        (plan.assignedNames || []).forEach(function(value) {
          phase4dNameKeys_(value).forEach(function(key) {
            (byName[key] || []).forEach(function(record) {
              if (employees.indexOf(record) === -1) employees.push(record);
            });
          });
        });
      }
      if (employees.length !== 1) {
        result.unresolved++;
        if (result.sample.length < 10) result.sample.push({ auditId: auditId,
          reason: employees.length ? 'Ambiguous employee' : 'Employee not found',
          workerNos: plan.workerNos, assignedNames: plan.assignedNames });
        return;
      }
      const employee = employees[0];
      result.employeeMatched++;
      if (matchSource === 'Worker No') result.workerNoMatched++;
      else result.nameFallback++;
      if (!phase4dIsEmail_(employee.email)) result.invalidEmployeeEmail++;
      if (phase4dIsEmail_(employee.supervisor)) {
        result.directSupervisorEmail++;
      } else {
        let supervisors = [];
        phase4dNameKeys_(employee.supervisor).forEach(function(key) {
          (byName[key] || []).forEach(function(record) {
            if (supervisors.indexOf(record) === -1) supervisors.push(record);
          });
        });
        if (supervisors.length === 1 && phase4dIsEmail_(supervisors[0].email)) {
          result.supervisorNameMatched++;
        } else {
          result.unresolved++;
          if (result.sample.length < 10) result.sample.push({ auditId: auditId,
            reason: 'Supervisor email unresolved', employee: employee.name,
            supervisorValue: employee.supervisor });
        }
      }
    });
  });
  if (result.unresolved) {
    phase4dFinding_(validation.warnings, 'Automation Recipient Index', '', '',
      result.unresolved + ' recipient mappings require review before individualized email.');
  }
  if (result.invalidEmployeeEmail) {
    phase4dFinding_(validation.warnings, 'Company Directory', '', '',
      result.invalidEmployeeEmail + ' matched employees have blank or invalid email addresses.');
  }
  return result;
}

function phase4dEquipment_(sheet, plans, validation) {
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 5, sheet.getLastRow() - 1, 5).getDisplayValues()
    : [];
  const equipment = {};
  rows.forEach(function(row) {
    const key = phase4dVehicleKey_(row[0]);
    if (key) equipment[key] = String(row[4] || '').trim();
  });
  const vehicles = {};
  ['Issues', 'Gate Checks'].forEach(function(name) {
    Object.keys(plans[name]).forEach(function(auditId) {
      const raw = String(plans[name][auditId].vehicle || '').trim();
      if (raw) vehicles[phase4dVehicleKey_(raw)] = raw;
    });
  });
  const result = { uniquePlanned: Object.keys(vehicles).length, matched: 0,
    notFound: 0, outOfService: 0, unmatched: [] };
  Object.keys(vehicles).forEach(function(key) {
    if (equipment[key] == null) {
      result.notFound++;
      if (result.unmatched.length < 10) result.unmatched.push(vehicles[key]);
    } else {
      result.matched++;
      if (String(equipment[key]).trim().toUpperCase() === 'OUT OF SERVICE') {
        result.outOfService++;
      }
    }
  });
  if (result.notFound) phase4dFinding_(validation.warnings, 'Equipment', '', '',
    result.notFound + ' planned vehicle IDs were not found: ' + result.unmatched.join(', '));
  return result;
}

function phase4dValidateStatuses_(ss, validation) {
  ['Issues', 'Gate Checks', 'JSO'].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    const column = name === 'JSO' ? 9 : 10;
    if (sheet.getLastRow() < 2) return;
    const statuses = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getDisplayValues();
    statuses.forEach(function(row, index) {
      const value = String(row[0] || '').trim();
      const allowed = name === 'JSO' ? ['', 'Closed'] : ['', 'Closed', 'Out of Service'];
      if (allowed.indexOf(value) === -1) {
        phase4dFinding_(validation.warnings, name, index + 2, '',
          'Unexpected Status value: ' + value);
      }
    });
  });
}

function phase4dIssueSignature_(value) {
  const aliases = {
    parkingbrake: 'parkingbrakes', parkingbrakes: 'parkingbrakes',
    lighteningdevices: 'lightingdevices', lightingdevices: 'lightingdevices',
    wheelsandrims: 'wheelsandrims', wheelandrims: 'wheelsandrims',
    vehiclesecuredlocked: 'vehiclesecuredlocked',
    propdevices: 'propdevices', properpinwhipclamps: 'properpinwhipclamps'
  };
  const unique = {};
  String(value || '').replace(/RESOLVED\s*:/gi, '\n').split(/\r?\n/)
    .map(phase4Normalize_).filter(Boolean).forEach(function(token) {
      unique[aliases[token] || token] = true;
    });
  return Object.keys(unique).sort().join('|');
}

function phase4dVehicleKey_(value) {
  const normalized = phase4Normalize_(value);
  if (/^\d+$/.test(normalized)) return normalized.replace(/^0+(?=\d)/, '');
  return normalized;
}

function phase4dNameKeys_(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return [];
  const keys = {};
  keys[phase4Normalize_(raw)] = true;
  const comma = raw.split(',');
  if (comma.length >= 2) keys[phase4Normalize_(comma.slice(1).join(' ') + ' ' + comma[0])] = true;
  const words = raw.replace(/[^a-z0-9 ]+/g, ' ').split(/\s+/).filter(Boolean)
    .filter(function(word) { return ['jr', 'sr', 'ii', 'iii', 'iv'].indexOf(word) === -1; });
  if (words.length >= 2) {
    keys[words[0] + words[words.length - 1]] = true;
    keys[words[words.length - 1] + words[0]] = true;
  }
  return Object.keys(keys);
}

function phase4dMapAdd_(map, key, value) {
  if (!key) return;
  if (!map[key]) map[key] = [];
  if (map[key].indexOf(value) === -1) map[key].push(value);
}

function phase4dIsEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function phase4dFinding_(array, sheet, row, auditId, reason) {
  array.push({ sheet: sheet, row: row || '', auditId: auditId || '', reason: reason });
}

function phase4dEmptyReport_(mode, validation) {
  return { mode: mode, readOnly: true, validation: validation,
    issueComparison: { exact: 0, equivalent: 0, substantive: 0, missing: 0, sample: [] },
    compositeMatching: { repeatedGroups: 0, matchedGroups: 0, blockedGroups: 0, sample: [] },
    recipients: { total: 0, employeeMatched: 0, directSupervisorEmail: 0, unresolved: 0 },
    equipment: { uniquePlanned: 0, matched: 0, notFound: 0, outOfService: 0, unmatched: [] },
    policies: [] };
}
