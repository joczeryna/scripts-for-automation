/**
 * Phase 4B: read-only consolidated-processing dry run.
 * Requires Phase4A_Preflight.gs to remain in the same Apps Script project.
 */

function previewConsolidatedUpdate() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) throw new Error('Another automation run is already active.');
  try {
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
    phase4ValidateDvir_(dvir);
    phase4ValidateSafety_(safety);

    const plans = { Issues: {}, 'Gate Checks': {}, JSO: {} };
    const sourceCounts = phase4bReadDvir_(dvir, plans);
    const safetyCounts = phase4bReadSafety_(safety, plans);
    const outOfService = phase4bOutOfService_(phase4RequireSheet_(ss, 'Equipment'));

    const comparisons = {};
    ['Issues', 'Gate Checks', 'JSO'].forEach(function(name) {
      comparisons[name] = phase4bCompareTarget_(
        phase4RequireSheet_(ss, name), plans[name], outOfService
      );
    });

    const recipients = phase4bPreviewRecipients_(
      phase4RequireSheet_(ss, 'Company Directory'), plans
    );

    const report = {
      mode: 'READ ONLY - NO WRITES OR ARCHIVING',
      source: {
        dvirRowsRead: sourceCounts.rowsRead,
        dvirIssueAudits: sourceCounts.issueAudits,
        safetyRowsRead: safetyCounts.rowsRead,
        safetyAuditIds: safetyCounts.auditIds,
        skippedSafetyAuditTypes: safetyCounts.skippedTypes
      },
      plannedDestinationChanges: comparisons,
      recipientResolution: recipients,
      assumptions: [
        'Destination Audit IDs absent from the current export remain unchanged.',
        'Status precedence for Issues/Gate Checks is Out of Service, then Closed, then blank.',
        'JSO is excluded from Out of Service, Closed, recipient, and email workflows.',
        'No ambiguous employee or supervisor name is guessed.'
      ]
    };

    Logger.log(JSON.stringify(report, null, 2));
    const summary = [
      'Phase 4B dry run completed — NO DATA WAS WRITTEN.',
      '',
      phase4bSummaryLine_('Issues', comparisons.Issues),
      phase4bSummaryLine_('Gate Checks', comparisons['Gate Checks']),
      phase4bSummaryLine_('JSO', comparisons.JSO),
      '',
      'Recipient mappings: ' + recipients.matched + ' Worker No matched, ' +
        recipients.nameFallback + ' name fallback, ' + recipients.ambiguous +
        ' ambiguous, ' + recipients.notFound + ' not found.',
      '',
      'Open Executions → Logs for the complete JSON report.'
    ].join('\n');
    SpreadsheetApp.getUi().alert(summary);
    return report;
  } finally {
    lock.releaseLock();
  }
}

function phase4bReadDvir_(sheet, plans) {
  const rows = sheet.getRange(3, 1, sheet.getLastRow() - 2, Math.max(29, sheet.getLastColumn()))
    .getDisplayValues();
  let issueAudits = 0;
  rows.forEach(function(row) {
    if (phase4Normalize_(row[0]) !== 'vci') return;
    const auditId = String(row[4] || '').trim();
    if (!auditId) return;
    const issueNames = [
      'Service Brakes', 'Parking Brake', 'Steering Mechanism', 'Lighting Devices',
      'Tires', 'Horn', 'Windshield Wipers', 'Rear Vision Mirrors',
      'Coupling Devices', 'Wheels And Rims', 'Emergency Equipment',
      'Windshield Condition', 'Oil Level', 'Trailer'
    ];
    const issues = [];
    for (let index = 12; index <= 25; index++) {
      const answer = String(row[index] || '').trim().toUpperCase();
      if (answer === 'DEFECT/DEFICIENCY' || answer === 'NEEDS ATTENTION') {
        issues.push(issueNames[index - 12]);
      }
    }
    const deviations = Number(row[26] || 0);
    const openDeviations = Number(row[27] || 0);
    const comments = String(row[28] || '').trim();
    if (deviations > 0 && openDeviations > 0 && issues.length === 0) issues.push('SEE FIELD AUDIT');
    if (issues.length === 0 && !(comments && deviations === 0)) return;

    plans.Issues[auditId] = phase4bMergePlan_(plans.Issues[auditId], {
      auditId: auditId,
      sourceType: 'Issues',
      auditType: 'DVIR',
      vehicle: String(row[10] || '').trim(),
      assignedNames: [String(row[6] || '').trim()],
      workerNos: [String(row[7] || '').trim()],
      openQuestions: issues,
      resolvedQuestions: [],
      plannedStatus: '',
      comments: comments
    });
    issueAudits++;
  });
  return { rowsRead: rows.length, issueAudits: issueAudits };
}

function phase4bReadSafety_(sheet, plans) {
  const values = sheet.getDataRange().getDisplayValues();
  const header = phase4bHeaderMap_(values[0]);
  const required = ['worker', 'workernbr', 'assetnumber', 'auditid', 'auditname',
    'deviationquestion', 'deviationstatus'];
  required.forEach(function(name) {
    if (header[name] == null) throw new Error('Safety Export missing required field: ' + name);
  });
  const grouped = {};
  const skipped = {};
  values.slice(1).forEach(function(row) {
    const auditId = String(row[header.auditid] || '').trim();
    if (!auditId) return;
    const auditType = String(row[header.auditname] || '').trim();
    const target = phase4bTarget_(auditType);
    if (!target) {
      skipped[auditType || '<blank>'] = (skipped[auditType || '<blank>'] || 0) + 1;
      return;
    }
    const key = target + '|' + auditId;
    const plan = grouped[key] || {
      auditId: auditId,
      sourceType: target,
      auditType: auditType,
      vehicle: String(row[header.assetnumber] || '').trim(),
      assignedNames: [], workerNos: [], openQuestions: [], resolvedQuestions: [], comments: ''
    };
    phase4bPushUnique_(plan.assignedNames, String(row[header.worker] || '').trim());
    phase4bPushUnique_(plan.workerNos, String(row[header.workernbr] || '').trim());
    const question = phase4bQuestion_(String(row[header.deviationquestion] || '').trim());
    const status = String(row[header.deviationstatus] || '').trim().toUpperCase();
    if (status === 'ACTIVE') phase4bPushUnique_(plan.openQuestions, question);
    else phase4bPushUnique_(plan.resolvedQuestions, question);
    grouped[key] = plan;
  });
  Object.keys(grouped).forEach(function(key) {
    const plan = grouped[key];
    plan.plannedStatus = plan.openQuestions.length === 0 && plan.resolvedQuestions.length > 0
      ? 'Closed' : '';
    plans[plan.sourceType][plan.auditId] = phase4bMergePlan_(
      plans[plan.sourceType][plan.auditId], plan
    );
  });
  return {
    rowsRead: Math.max(0, values.length - 1),
    auditIds: Object.keys(grouped).length,
    skippedTypes: skipped
  };
}

function phase4bMergePlan_(existing, incoming) {
  if (!existing) return incoming;
  ['assignedNames', 'workerNos', 'openQuestions', 'resolvedQuestions'].forEach(function(field) {
    (incoming[field] || []).forEach(function(value) { phase4bPushUnique_(existing[field], value); });
  });
  if (incoming.vehicle) existing.vehicle = incoming.vehicle;
  if (incoming.auditType) existing.auditType = incoming.auditType;
  if (incoming.comments) existing.comments = incoming.comments;
  existing.plannedStatus = incoming.plannedStatus;
  return existing;
}

function phase4bCompareTarget_(sheet, plans, outOfService) {
  const statusIndex = sheet.getName() === 'JSO' ? 8 : 9;
  const issueIndex = sheet.getName() === 'JSO' ? 5 : 6;
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
    : [];
  const existing = {};
  let duplicateAuditIds = 0;
  values.forEach(function(row) {
    const id = String(row[1] || '').trim();
    if (!id) return;
    if (existing[id]) duplicateAuditIds++;
    else existing[id] = row;
  });
  const result = {
    plannedRecords: Object.keys(plans).length,
    inserts: 0, issueUpdates: 0, statusUpdates: 0,
    unchanged: 0, duplicateDestinationAuditIds: duplicateAuditIds,
    outOfService: 0, closed: 0, openOrMixed: 0
  };
  Object.keys(plans).forEach(function(id) {
    const plan = plans[id];
    if (sheet.getName() !== 'JSO' && outOfService[phase4Normalize_(plan.vehicle)]) {
      plan.plannedStatus = 'Out of Service';
    }
    if (plan.plannedStatus === 'Out of Service') result.outOfService++;
    else if (plan.plannedStatus === 'Closed') result.closed++;
    else result.openOrMixed++;
    const row = existing[id];
    if (!row) {
      result.inserts++;
      return;
    }
    const plannedIssues = phase4bIssueText_(plan);
    let changed = false;
    if (String(row[issueIndex] || '').trim() !== plannedIssues) {
      result.issueUpdates++;
      changed = true;
    }
    if (String(row[statusIndex] || '').trim() !== plan.plannedStatus) {
      result.statusUpdates++;
      changed = true;
    }
    if (!changed) result.unchanged++;
  });
  return result;
}

function phase4bPreviewRecipients_(directorySheet, plans) {
  const data = directorySheet.getDataRange().getDisplayValues();
  const header = phase4bHeaderMap_(data[0]);
  ['name', 'workerno', 'emailaddress', 'supervisor'].forEach(function(name) {
    if (header[name] == null) throw new Error('Company Directory missing field: ' + name);
  });
  const byWorker = {};
  const byName = {};
  data.slice(1).forEach(function(row) {
    const record = {
      name: String(row[header.name] || '').trim(),
      workerNo: String(row[header.workerno] || '').trim(),
      email: String(row[header.emailaddress] || '').trim(),
      supervisor: String(row[header.supervisor] || '').trim()
    };
    if (record.workerNo) {
      const workerKey = phase4Normalize_(record.workerNo);
      if (!byWorker[workerKey]) byWorker[workerKey] = [];
      byWorker[workerKey].push(record);
    }
    if (record.name) {
      const nameKey = phase4Normalize_(record.name);
      if (!byName[nameKey]) byName[nameKey] = [];
      byName[nameKey].push(record);
    }
  });
  const result = {
    total: 0, matched: 0, nameFallback: 0, ambiguous: 0, notFound: 0,
    missingEmployeeEmail: 0, missingSupervisorEmail: 0
  };
  ['Issues', 'Gate Checks'].forEach(function(target) {
    Object.keys(plans[target]).forEach(function(id) {
      result.total++;
      const plan = plans[target][id];
      const workerNos = (plan.workerNos || []).filter(function(value) {
        return value && value !== '-1';
      });
      let candidates = [];
      workerNos.forEach(function(value) {
        (byWorker[phase4Normalize_(value)] || []).forEach(function(row) {
          if (candidates.indexOf(row) === -1) candidates.push(row);
        });
      });
      let employee = null;
      if (candidates.length === 1) {
        employee = candidates[0];
        result.matched++;
      } else if (candidates.length > 1) {
        result.ambiguous++;
        return;
      } else {
        const nameMatches = [];
        (plan.assignedNames || []).forEach(function(value) {
          (byName[phase4Normalize_(value)] || []).forEach(function(row) {
            if (nameMatches.indexOf(row) === -1) nameMatches.push(row);
          });
        });
        if (nameMatches.length === 1) {
          employee = nameMatches[0];
          result.nameFallback++;
        } else if (nameMatches.length > 1) {
          result.ambiguous++;
          return;
        } else {
          result.notFound++;
          return;
        }
      }
      if (!employee.email) result.missingEmployeeEmail++;
      const supervisors = byName[phase4Normalize_(employee.supervisor)] || [];
      if (supervisors.length !== 1 || !supervisors[0].email) result.missingSupervisorEmail++;
    });
  });
  return result;
}

function phase4bOutOfService_(sheet) {
  const result = {};
  if (sheet.getLastRow() < 2) return result;
  const rows = sheet.getRange(2, 5, sheet.getLastRow() - 1, 5).getDisplayValues();
  rows.forEach(function(row) {
    const vehicle = phase4Normalize_(row[0]);
    const status = String(row[4] || '').trim().toUpperCase();
    if (vehicle && status === 'OUT OF SERVICE') result[vehicle] = true;
  });
  return result;
}

function phase4bTarget_(auditType) {
  const value = String(auditType || '').toUpperCase();
  if (value.indexOf('JSO') !== -1) return 'JSO';
  if (value.indexOf('GATE') !== -1) return 'Gate Checks';
  if (value.indexOf('DVIR') !== -1) return 'Issues';
  return '';
}

function phase4bIssueText_(plan) {
  const open = (plan.openQuestions || []).filter(Boolean);
  const resolved = (plan.resolvedQuestions || []).filter(Boolean);
  if (!resolved.length) return open.join('\n');
  if (!open.length) return 'RESOLVED:\n' + resolved.join('\n');
  return open.join('\n') + '\nRESOLVED:\n' + resolved.join('\n');
}

function phase4bQuestion_(value) {
  const replacements = {
    'LIGHTENINGDEVICES': 'Lighting Devices',
    'VEHICLESECUREDLOCKED': 'Vehicle Secured/Locked',
    'PROPDEVICES': 'Prop. Devices',
    'PROPERPINWHIPCLAMPS': 'Proper Pin/Whip/Clamps'
  };
  return replacements[phase4Normalize_(value).toUpperCase()] || value;
}

function phase4bHeaderMap_(headers) {
  const result = {};
  headers.forEach(function(value, index) { result[phase4Normalize_(value)] = index; });
  return result;
}

function phase4bPushUnique_(array, value) {
  const text = String(value || '').trim();
  if (text && array.indexOf(text) === -1) array.push(text);
}

function phase4bSummaryLine_(name, result) {
  return name + ': ' + result.inserts + ' inserts, ' + result.issueUpdates +
    ' issue updates, ' + result.statusUpdates + ' status updates, ' +
    result.unchanged + ' unchanged.';
}
