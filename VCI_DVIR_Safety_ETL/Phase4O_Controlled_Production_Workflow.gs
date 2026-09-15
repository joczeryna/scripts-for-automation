/**
 * Phase 4O: controlled production workflow for DVIR + Safety Deviations.
 *
 * Requires installed Phase 4A, 4B, 4D, 4E, 4H, 4I, 4J, and 4N helpers.
 */

const PHASE4O_ARCHIVE_FOLDER_ID = '1Wt2P3yZNscQqqfL4A0e1Kwn7T5vhcCou';
const PHASE4O_ERROR_RECIPIENT = 'jo.deleon@vcicom.com';
const PHASE4O_SENDER_NAME = 'DVIR Report Automation';

function runPhase4OControlledProductionWorkflow() {
  const lock = LockService.getDocumentLock();
  let operationalFiltersRemoved = false;
  let ss = null;

  if (!lock.tryLock(10000)) {
    throw new Error(
      'Another automation run is already active.'
    );
  }

  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();

    const sources = phase4oResolveSources_();

    const archiveFolder = DriveApp.getFolderById(
      PHASE4O_ARCHIVE_FOLDER_ID
    );

    const confirmation = SpreadsheetApp.getUi().alert(
      'Run controlled production workflow?',
      'DVIR: ' + sources.dvir.getName() + '\n' +
      'Safety: ' + sources.safety.getName() + '\n\n' +
      'A backup will be created. Valid new closures may be emailed. ' +
      'Both source files will be archived only after successful ' +
      'validation and email handling.',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (
      confirmation !==
      SpreadsheetApp.getUi().Button.YES
    ) {
      Logger.log(
        'PHASE4O_RESULT=' +
        JSON.stringify({
          cancelledByUser: true
        })
      );

      return {
        cancelledByUser: true
      };
    }

    const pre = validateWorkbookIntegrity(
      'PHASE4O_PRODUCTION_PRECHECK'
    );

    if (pre.validation.status === 'BLOCKED') {
      phase4oSendBlockingNotice_(
        'Pre-update validation',
        pre.validation.blockers
      );

      throw new Error(
        'Production run blocked before writes.'
      );
    }

    /*
     * Remove operational filters before any rows are
     * inserted, formatted, or assigned formulas.
     */
    phase4oRemoveOperationalFilters_(ss);
    operationalFiltersRemoved = true;

    const backup = phase4oCreateProductionBackup_(ss);
    const context = phase4eBuildContext_(ss);

    const result = {
      mode: 'CONTROLLED PRODUCTION WORKFLOW',

      backupId: backup.getId(),
      backupName: backup.getName(),

      Issues: phase4eApplyTarget_(
        ss.getSheetByName('Issues'),
        context.plans.Issues,
        context.equipment,
        context.locationDivisions
      ),

      GateChecks: phase4eApplyTarget_(
        ss.getSheetByName('Gate Checks'),
        context.plans['Gate Checks'],
        context.equipment,
        context.locationDivisions
      ),

      JSO: phase4eApplyTarget_(
        ss.getSheetByName('JSO'),
        context.plans.JSO,
        context.equipment,
        context.locationDivisions
      ),

      legacyStatusNormalization:
        phase4eNormalizeLegacyStatuses_(ss)
    };

    result.recipientIndex =
      phase4eRefreshRecipientIndex_(
        ss,
        context.plans
      );

    result.Closed =
      phase4oAppendNewClosed_(
        ss,
        context.plans
      );

    SpreadsheetApp.flush();

    const preEmail = validateWorkbookIntegrity(
      'PHASE4O_PRE_EMAIL_GATE'
    );

    result.preEmailValidationStatus =
      preEmail.validation.status;

    if (preEmail.validation.status === 'BLOCKED') {
      phase4oSendBlockingNotice_(
        'Pre-email validation',
        preEmail.validation.blockers
      );

      throw new Error(
        'Production run blocked after workbook updates. ' +
        'Restore backup ' +
        backup.getName() +
        ' before retrying.'
      );
    }

    result.email =
      phase4oProcessEmailQueue_(ss);

    if (
      result.email.deliveryFailures.length ||
      result.email.errorNoticeFailed
    ) {
      throw new Error(
        'Email handling was incomplete. ' +
        'Source files were not archived. ' +
        JSON.stringify(
          result.email.deliveryFailures.slice(0, 10)
        )
      );
    }

    /*
     * Restore operational filters after all data,
     * formulas, formatting, and email handling finish.
     */
    phase4oApplyOperationalFilters_(ss);
    operationalFiltersRemoved = false;

    const finalCheck = validateWorkbookIntegrity(
      'PHASE4O_FINAL_POSTRUN'
    );

    result.finalValidationStatus =
      finalCheck.validation.status;

    if (
      finalCheck.validation.status === 'BLOCKED'
    ) {
      phase4oSendBlockingNotice_(
        'Final post-run validation',
        finalCheck.validation.blockers
      );

      throw new Error(
        'Final validation is BLOCKED. ' +
        'Source files were not archived.'
      );
    }

    result.archived =
      phase4oArchiveSources_(
        sources,
        archiveFolder
      );

    result.senderDisplayName =
      PHASE4O_SENDER_NAME;

    result.validationErrorRecipient =
      PHASE4O_ERROR_RECIPIENT;

    Logger.log(
      'PHASE4O_RESULT=' +
      JSON.stringify(result)
    );

    SpreadsheetApp.getUi().alert(
      [
        'Phase 4O production workflow completed.',
        '',
        'Issues inserts/updates: ' +
          result.Issues.inserts +
          '/' +
          result.Issues.updates,

        'Gate Checks inserts/updates: ' +
          result.GateChecks.inserts +
          '/' +
          result.GateChecks.updates,

        'JSO inserts/updates: ' +
          result.JSO.inserts +
          '/' +
          result.JSO.updates,

        'New Closed rows: ' +
          result.Closed.inserted,

        'Operational emails sent: ' +
          result.email.sent,

        'Unresolved records reported to Jo: ' +
          result.email.unresolved.length,

        'Sources archived: ' +
          result.archived.length,

        'Final validation: ' +
          result.finalValidationStatus,

        '',
        'Backup: ' +
          result.backupName,

        'Copy PHASE4O_RESULT from Executions → Logs.'
      ].join('\n')
    );

    return result;
  } finally {
    /*
     * If the workflow fails after removing filters,
     * restore them before releasing the lock.
     */
    try {
      if (
        operationalFiltersRemoved &&
        ss
      ) {
        phase4oApplyOperationalFilters_(ss);
      }
    } catch (filterError) {
      Logger.log(
        'PHASE4O_FILTER_RESTORE_ERROR=' +
        filterError.message
      );
    } finally {
      lock.releaseLock();
    }
  }
}

function phase4oRemoveOperationalFilters_(ss) {
  [
    'Issues',
    'Gate Checks',
    'JSO'
  ].forEach(function(sheetName) {
    const sheet = phase4RequireSheet_(
      ss,
      sheetName
    );

    const filter = sheet.getFilter();

    if (filter) {
      filter.remove();
    }
  });

  SpreadsheetApp.flush();
}

function phase4oApplyOperationalFilters_(ss) {
  const configurations = [
    {
      sheetName: 'Issues',
      statusColumn: 10
    },
    {
      sheetName: 'Gate Checks',
      statusColumn: 10
    },
    {
      sheetName: 'JSO',
      statusColumn: 9
    }
  ];

  configurations.forEach(function(config) {
    const sheet = phase4RequireSheet_(
      ss,
      config.sheetName
    );

    const existingFilter = sheet.getFilter();

    if (existingFilter) {
      existingFilter.remove();
    }

    const lastRow = Math.max(
      sheet.getLastRow(),
      2
    );

    const lastColumn =
      sheet.getLastColumn();

    const filter = sheet
      .getRange(
        1,
        1,
        lastRow,
        lastColumn
      )
      .createFilter();

    const statusValues = sheet
      .getRange(
        2,
        config.statusColumn,
        lastRow - 1,
        1
      )
      .getDisplayValues();

    const hiddenValues = [];

    statusValues.forEach(function(row) {
      const status = String(
        row[0] || ''
      ).trim();

      /*
       * Keep only blank and Out of Service visible.
       *
       * For JSO, Out of Service does not occur,
       * so this leaves only blank-status rows visible.
       */
      if (
        status &&
        status !== 'Out of Service' &&
        hiddenValues.indexOf(status) === -1
      ) {
        hiddenValues.push(status);
      }
    });

    if (hiddenValues.length) {
      const criteria = SpreadsheetApp
        .newFilterCriteria()
        .setHiddenValues(hiddenValues)
        .build();

      filter.setColumnFilterCriteria(
        config.statusColumn,
        criteria
      );
    }
  });

  SpreadsheetApp.flush();
}

function phase4oResolveSources_() {
  const folder = DriveApp.getFolderById(PHASE4_INTAKE_FOLDER_ID);
  const runDate = Utilities.formatDate(new Date(), PHASE4_TIME_ZONE, 'MM-dd-yyyy');
  return {
    intake: folder,
    dvir: phase4FindUniqueFile_(folder, ['DVIR_Details', 'DVIR_Details.xlsx']),
    safety: phase4FindUniqueFile_(folder, [
      'safety-deviations-' + runDate, 'safety-deviations-' + runDate + '.xlsx'
    ])
  };
}

function phase4oCreateProductionBackup_(ss) {
  return ss.copy(
    ss.getName() + ' - Production Backup - ' +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmmss')
  );
}

function phase4oAppendNewClosed_(ss, plans) {
  const sourceRows = phase4iCurrentClosedRows_(ss, plans);
  const closed = phase4RequireSheet_(ss, 'Closed');
  const existing = closed.getLastRow() > 1
    ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues() : [];
  const byAudit = {};
  existing.forEach(function(row) {
    const auditId = String(row[1] || '').trim();
    if (!auditId) return;
    if (!byAudit[auditId]) byAudit[auditId] = [];
    byAudit[auditId].push({ values: row, used: false });
  });
  const missing = [];
  let matched = 0;
  sourceRows.forEach(function(source) {
    const candidates = (byAudit[source.auditId] || []).filter(function(x) { return !x.used; });
    let best = null;
    let bestScore = -1;
    candidates.forEach(function(candidate) {
      const score = phase4hSimilarity_(source.values, candidate.values);
      if (score > bestScore) { best = candidate; bestScore = score; }
    });
    if (best) { best.used = true; matched++; }
    else missing.push(source);
  });
  if (missing.length) {
    const start = closed.getLastRow() + 1;
    if (closed.getMaxRows() < start + missing.length - 1) {
      closed.insertRowsAfter(closed.getMaxRows(), start + missing.length - 1 - closed.getMaxRows());
    }
    closed.getRange(start, 1, missing.length, 16).setValues(missing.map(function(source) {
      return phase4hPad_(source.values, 15).concat(['']);
    }));
    if (start > 2) {
      closed.getRange(start - 1, 1, 1, 16).copyTo(
        closed.getRange(start, 1, missing.length, 16),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false
      );
    }
    closed.getRange(start, 16, missing.length, 1).setNumberFormat('MM/dd/yyyy');
  }
  return { eligible: sourceRows.length, matched: matched, inserted: missing.length,
    existingRowsPreserved: existing.length };
}

function phase4oProcessEmailQueue_(ss) {
  const closed = phase4RequireSheet_(ss, 'Closed');
  const headers = closed.getRange(1, 1, 1, 16).getDisplayValues()[0];
  const rows = closed.getLastRow() > 1
    ? closed.getRange(2, 1, closed.getLastRow() - 1, 16).getDisplayValues()
    : [];

  const index = phase4jRecipientIndex_(ss, []);
  const directory = phase4jDirectory_(ss, []);

  const result = {
    pending: 0,
    sent: 0,
    unresolved: [],
    deliveryFailures: [],
    errorNoticeSent: false,
    errorNoticeFailed: false,
    sentSample: []
  };

  rows.forEach(function(row, offset) {
    const auditId = String(row[1] || '').trim();

    if (!auditId || String(row[15] || '').trim()) return;

    result.pending++;

    const sourceType = phase4jSourceType_(row, index);
    const mapping = sourceType
      ? index[sourceType + '|' + auditId]
      : null;

    const employees = mapping
      ? directory.byWorker[phase4Normalize_(mapping.workerNo)] || []
      : [];

    const employee = employees.length === 1 ? employees[0] : null;
    let reason = '';

    if (!sourceType) {
      reason = 'Source Type unresolved';
    } else if (!mapping) {
      reason = 'Recipient mapping missing';
    } else if (['Matched', 'Name Fallback'].indexOf(mapping.status) === -1) {
      reason = 'Recipient mapping status is ' +
        (mapping.status || 'blank');
    } else if (!employee) {
      reason = employees.length
        ? 'Employee mapping ambiguous'
        : 'Employee not found';
    } else if (!phase4dIsEmail_(employee.email)) {
      reason = 'Driver email invalid';
    } else if (!phase4dIsEmail_(employee.supervisor)) {
      reason = 'Supervisor email invalid';
    }

    if (reason) {
      result.unresolved.push({
        closedRow: offset + 2,
        auditId: auditId,
        sourceType: sourceType,
        reason: reason
      });
      return;
    }

    try {
      const subject = '[DVIR CLOSED] Audit ' + auditId +
        (String(row[3] || '').trim()
          ? ' — Vehicle ' + String(row[3]).trim()
          : '');

      const ccRecipients = phase4oBuildOperationalCc_(
        employee.email,
        employee.supervisor
      );

      MailApp.sendEmail({
        to: employee.email,
        cc: ccRecipients,
        subject: subject,
        body: phase4oOperationalText_(headers, row),
        htmlBody: phase4oOperationalHtml_(headers, row),
        name: PHASE4O_SENDER_NAME
      });

      closed
        .getRange(offset + 2, 16)
        .setValue(new Date())
        .setNumberFormat('MM/dd/yyyy');

      result.sent++;

      if (result.sentSample.length < 10) {
        result.sentSample.push({
          auditId: auditId,
          to: employee.email.toLowerCase(),
          cc: ccRecipients
        });
      }
    } catch (error) {
      result.deliveryFailures.push({
        closedRow: offset + 2,
        auditId: auditId,
        reason: error.message
      });
    }
  });

  if (result.unresolved.length || result.deliveryFailures.length) {
    try {
      MailApp.sendEmail({
        to: PHASE4O_ERROR_RECIPIENT,
        subject: '[ACTION REQUIRED] DVIR Report Automation — ' +
          (result.unresolved.length + result.deliveryFailures.length) +
          ' Email Exception(s)',
        body: phase4oExceptionText_(result),
        htmlBody: phase4oExceptionHtml_(result),
        name: PHASE4O_SENDER_NAME
      });

      result.errorNoticeSent = true;
    } catch (error) {
      result.errorNoticeFailed = true;
      result.deliveryFailures.push({
        auditId: '',
        reason: 'Jo notification failed: ' + error.message
      });
    }
  }

  SpreadsheetApp.flush();
  return result;
}

function phase4oBuildOperationalCc_(driverEmail, supervisorEmail) {
  const driver = String(driverEmail || '').trim().toLowerCase();
  const candidates = [
    supervisorEmail,
    PHASE4O_ERROR_RECIPIENT
  ];

  const unique = [];

  candidates.forEach(function(email) {
    const normalized = String(email || '').trim().toLowerCase();

    if (
      normalized &&
      normalized !== driver &&
      unique.indexOf(normalized) === -1
    ) {
      unique.push(normalized);
    }
  });

  return unique.join(',');
}

function phase4oArchiveSources_(sources, archiveFolder) {
  const runDate = Utilities.formatDate(new Date(), PHASE4_TIME_ZONE, 'MM-dd-yyyy');
  const moved = [];
  const originals = [];
  try {
    [sources.dvir, sources.safety].forEach(function(file) {
      const originalName = file.getName();
      originals.push({ file: file, name: originalName });
      const archiveName = originalName.indexOf(runDate) === -1
        ? originalName + ' - ' + runDate : originalName;
      file.setName(archiveName).moveTo(archiveFolder);
      moved.push({ id: file.getId(), originalName: originalName, archiveName: archiveName });
    });
    return moved;
  } catch (error) {
    originals.forEach(function(item) {
      try { item.file.setName(item.name).moveTo(sources.intake); } catch (ignored) {}
    });
    throw new Error('Source archiving failed and rollback was attempted: ' + error.message);
  }
}

function phase4oOperationalText_(headers, row) {
  const lines = ['A reported DVIR or safety deviation has been resolved and closed.', ''];
  for (let i = 0; i < 15; i++) {
    if (String(row[i] || '').trim()) lines.push((headers[i] || ('Field ' + (i + 1))) + ': ' + row[i]);
  }
  lines.push('', 'This message was sent automatically by DVIR Report Automation.');
  return lines.join('\n');
}

function phase4oOperationalHtml_(headers, row) {
  let table = '';
  for (let i = 0; i < 15; i++) {
    if (!String(row[i] || '').trim()) continue;
    table += '<tr><td style="padding:6px;border:1px solid #bbb"><strong>' +
      phase4oEscape_(headers[i] || ('Field ' + (i + 1))) + '</strong></td><td style="padding:6px;border:1px solid #bbb">' +
      phase4oEscape_(row[i]) + '</td></tr>';
  }
  return '<div style="font-family:Arial,sans-serif;color:#222"><h2>Closed DVIR / Safety Record</h2>' +
    '<p>A reported DVIR or safety deviation has been resolved and closed.</p>' +
    '<table style="border-collapse:collapse">' + table + '</table>' +
    '<p>This message was sent automatically by DVIR Report Automation.</p></div>';
}

function phase4oExceptionText_(result) {
  const lines = ['DVIR Report Automation found email exceptions.', '',
    'Unresolved recipients: ' + result.unresolved.length,
    'Delivery failures: ' + result.deliveryFailures.length, ''];
  result.unresolved.concat(result.deliveryFailures).slice(0, 50).forEach(function(item) {
    lines.push('Audit ' + (item.auditId || '(unknown)') +
      (item.closedRow ? ', Closed row ' + item.closedRow : '') + ': ' + item.reason);
  });
  lines.push('', 'Affected operational emails were suppressed or left unsent. Their Emailed Dates remain blank.');
  return lines.join('\n');
}

function phase4oExceptionHtml_(result) {
  return '<div style="font-family:Arial,sans-serif;color:#222"><h2>Action Required: Email Exceptions</h2><pre>' +
    phase4oEscape_(phase4oExceptionText_(result)) + '</pre></div>';
}

function phase4oSendBlockingNotice_(stage, blockers) {
  MailApp.sendEmail({
    to: PHASE4O_ERROR_RECIPIENT,
    subject: '[BLOCKED] DVIR Report Automation — ' + stage,
    body: 'The production workflow stopped at ' + stage + '.\n\n' +
      JSON.stringify(blockers, null, 2),
    name: PHASE4O_SENDER_NAME
  });
}

function phase4oEscape_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/\r?\n/g, '<br>');
}

function sendPhase4OJoOnlyClosedDateTest() {
  const TEST_DATE = '2026-09-15';
  const TEST_RECIPIENT = 'jo.deleon@vcicom.com';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = ss.getSpreadsheetTimeZone();
  const closed = phase4RequireSheet_(ss, 'Closed');
  const lastRow = closed.getLastRow();

  if (lastRow < 2) {
    const emptyResult = {
      mode: 'JO-ONLY CLOSED-DATE EMAIL TEST',
      testDate: TEST_DATE,
      recipient: TEST_RECIPIENT,
      matchingRecords: 0,
      messagesSent: 0,
      deliveryFailures: [],
      workbookWrites: 0,
      emailedDatesWritten: 0,
      driversContacted: 0,
      supervisorsContacted: 0,
      sourcesArchived: false
    };

    Logger.log(
      'PHASE4O_JO_ONLY_DATE_TEST=' +
      JSON.stringify(emptyResult)
    );

    return emptyResult;
  }

  const headers = closed
    .getRange(1, 1, 1, 16)
    .getDisplayValues()[0];

  const range = closed.getRange(
    2,
    1,
    lastRow - 1,
    16
  );

  const rawRows = range.getValues();
  const displayRows = range.getDisplayValues();
  const selected = [];

  rawRows.forEach(function(rawRow, index) {
    const displayRow = displayRows[index];
    const auditId =
      String(displayRow[1] || '').trim();

    if (!auditId) return;

    const emailedDate = phase4oTestDateKey_(
      rawRow[15],
      displayRow[15],
      timeZone
    );

    if (emailedDate === TEST_DATE) {
      selected.push({
        rowNumber: index + 2,
        auditId: auditId,
        row: displayRow
      });
    }
  });

  const result = {
    mode: 'JO-ONLY CLOSED-DATE EMAIL TEST',
    testDate: TEST_DATE,
    recipient: TEST_RECIPIENT,
    matchingRecords: selected.length,
    messagesSent: 0,
    deliveryFailures: [],
    workbookWrites: 0,
    emailedDatesWritten: 0,
    driversContacted: 0,
    supervisorsContacted: 0,
    sourcesArchived: false,
    sentSample: []
  };

  if (!selected.length) {
    Logger.log(
      'PHASE4O_JO_ONLY_DATE_TEST=' +
      JSON.stringify(result)
    );

    return result;
  }

  selected.forEach(function(item) {
    try {
      const vehicle =
        String(item.row[3] || '').trim();

      const subject =
        '[TEST — JO ONLY — NO ACTION REQUIRED] ' +
        'DVIR Closed Audit ' +
        item.auditId +
        (
          vehicle
            ? ' — Vehicle ' + vehicle
            : ''
        );

      const textNotice =
        'TEST — JO ONLY — NO ACTION REQUIRED\n\n' +
        'Production closure batch: 09/15/2026\n' +
        'The assigned driver and supervisor were not contacted.\n' +
        'No workbook values or Emailed Dates were changed.\n\n';

      const htmlNotice =
        '<div style="' +
        'font-family:Arial,sans-serif;' +
        'background:#fff3cd;' +
        'border:1px solid #ffe69c;' +
        'padding:12px;' +
        'margin-bottom:16px;">' +

        '<strong>' +
        'TEST — JO ONLY — NO ACTION REQUIRED' +
        '</strong><br>' +

        'Production closure batch: 09/15/2026<br>' +
        'The assigned driver and supervisor were not contacted.<br>' +
        'No workbook values or Emailed Dates were changed.' +

        '</div>';

      MailApp.sendEmail({
        to: TEST_RECIPIENT,
        subject: subject,

        body:
          textNotice +
          phase4oOperationalText_(
            headers,
            item.row
          ),

        htmlBody:
          htmlNotice +
          phase4oOperationalHtml_(
            headers,
            item.row
          ),

        name: PHASE4O_SENDER_NAME
      });

      result.messagesSent++;

      if (result.sentSample.length < 10) {
        result.sentSample.push({
          auditId: item.auditId,
          closedRow: item.rowNumber,
          to: TEST_RECIPIENT
        });
      }
    } catch (error) {
      result.deliveryFailures.push({
        auditId: item.auditId,
        closedRow: item.rowNumber,
        reason: error.message
      });
    }
  });

  Logger.log(
    'PHASE4O_JO_ONLY_DATE_TEST=' +
    JSON.stringify(result)
  );

  return result;
}

function phase4oTestDateKey_(
  rawValue,
  displayValue,
  timeZone
) {
  if (
    rawValue instanceof Date &&
    !isNaN(rawValue.getTime())
  ) {
    return Utilities.formatDate(
      rawValue,
      timeZone,
      'yyyy-MM-dd'
    );
  }

  const text =
    String(displayValue || '').trim();

  if (!text) return '';

  const match = text.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})/
  );

  if (!match) return '';

  return (
    match[3] +
    '-' +
    ('0' + match[1]).slice(-2) +
    '-' +
    ('0' + match[2]).slice(-2)
  );
}