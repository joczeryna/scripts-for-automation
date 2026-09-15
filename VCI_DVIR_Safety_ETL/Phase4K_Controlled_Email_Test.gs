/**
 * Phase 4K: controlled live-email test.
 *
 * Sends exactly two synthetic test emails to jo.deleon@vcicom.com:
 *  1) operational Closed-record format/routing simulation;
 *  2) validation-error notification format.
 *
 * No spreadsheet writes, Emailed Date changes, or archiving occur.
 */

const PHASE4K_TEST_RECIPIENT = 'jo.deleon@vcicom.com';
const PHASE4K_OPERATIONAL_SUBJECT =
  '[TEST — NO ACTION REQUIRED] DVIR Closed Record Email';
const PHASE4K_VALIDATION_SUBJECT =
  '[TEST — NO ACTION REQUIRED] DVIR Validation Error Notification';

function previewPhase4KControlledEmailTest() {
  const base = validateWorkbookIntegrity('PHASE4K_CONTROLLED_EMAIL_TEST_PREVIEW');
  const validation = base.validation || base;
  const report = phase4kReport_(validation, false, []);
  Logger.log('PHASE4K_PREVIEW=' + JSON.stringify(report));
  SpreadsheetApp.getUi().alert([
    'Phase 4K preview completed — NO EMAILS SENT.', '',
    'Recipient: ' + PHASE4K_TEST_RECIPIENT,
    'Messages prepared: 2',
    'Validation: ' + validation.status, '',
    'Run sendPhase4KControlledEmailTest only when ready.'
  ].join('\n'));
  return report;
}

function sendPhase4KControlledEmailTest() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) throw new Error('Another automation run is active.');
  try {
    const base = validateWorkbookIntegrity('PHASE4K_CONTROLLED_EMAIL_TEST_SEND');
    const validation = base.validation || base;
    if ((validation.blockers || []).length || validation.status === 'BLOCKED') {
      throw new Error('Phase 4K stopped: workbook validation is BLOCKED. No test email was sent.');
    }
    if (MailApp.getRemainingDailyQuota() < 2) {
      throw new Error('Phase 4K stopped: fewer than two email recipients remain in today\'s quota.');
    }

    const confirmation = SpreadsheetApp.getUi().alert(
      'Send two controlled test emails?',
      'Both synthetic messages will go only to ' + PHASE4K_TEST_RECIPIENT +
        '. No workbook data will be changed.',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (confirmation !== SpreadsheetApp.getUi().Button.YES) {
      const cancelled = phase4kReport_(validation, false, []);
      cancelled.cancelledByUser = true;
      Logger.log('PHASE4K_RESULT=' + JSON.stringify(cancelled));
      return cancelled;
    }

    const sent = [];
    MailApp.sendEmail({
      to: PHASE4K_TEST_RECIPIENT,
      subject: PHASE4K_OPERATIONAL_SUBJECT,
      body: phase4kOperationalText_(),
      htmlBody: phase4kOperationalHtml_(),
      name: 'DVIR Report Automation'
    });
    sent.push({ type: 'OPERATIONAL ROUTING TEST', to: PHASE4K_TEST_RECIPIENT,
      cc: '', subject: PHASE4K_OPERATIONAL_SUBJECT });

    MailApp.sendEmail({
      to: PHASE4K_TEST_RECIPIENT,
      subject: PHASE4K_VALIDATION_SUBJECT,
      body: phase4kValidationText_(validation),
      htmlBody: phase4kValidationHtml_(validation),
      name: 'DVIR Report Automation'
    });
    sent.push({ type: 'VALIDATION NOTIFICATION TEST', to: PHASE4K_TEST_RECIPIENT,
      cc: '', subject: PHASE4K_VALIDATION_SUBJECT });

    const report = phase4kReport_(validation, true, sent);
    Logger.log('PHASE4K_RESULT=' + JSON.stringify(report));
    SpreadsheetApp.getUi().alert([
      'Phase 4K controlled test completed.', '',
      'Test emails sent: 2',
      'Recipient: ' + PHASE4K_TEST_RECIPIENT,
      'Workbook writes: 0',
      'Files archived: 0', '',
      'Confirm both messages arrived, then send the PHASE4K_RESULT log.'
    ].join('\n'));
    return report;
  } finally {
    lock.releaseLock();
  }
}

function phase4kReport_(validation, emailsSent, messages) {
  return {
    mode: emailsSent ? 'CONTROLLED LIVE EMAIL TEST' : 'READ-ONLY EMAIL TEST PREVIEW',
    recipient: PHASE4K_TEST_RECIPIENT,
    validationStatus: validation.status,
    validationBlockerCount: (validation.blockers || []).length,
    validationWarningCount: (validation.warnings || []).length,
    messagesPrepared: 2,
    messagesSent: emailsSent ? messages.length : 0,
    messages: messages,
    sameAddressRoutingNote: 'Jo represents both driver and supervisor for this synthetic test; the address is used once in To and is not duplicated in CC.',
    emailedDatesWritten: 0,
    workbookWrites: 0,
    sourcesArchived: false,
    emailsSent: emailsSent
  };
}

function phase4kOperationalText_() {
  return [
    'TEST — NO ACTION REQUIRED', '',
    'This is a synthetic test of the future DVIR closed-record email.',
    'No employee, vehicle, Audit ID, or workbook record is associated with this message.', '',
    'Production routing after approval:',
    'To: Assigned Driver',
    'CC: Assigned Driver Supervisor', '',
    'For this controlled test, jo.deleon@vcicom.com represents both roles and appears only once in To.', '',
    'Test Audit ID: TEST-PHASE4K',
    'Test Vehicle: TEST-VEHICLE',
    'Test Status: Closed', '',
    'No action is required.'
  ].join('\n');
}

function phase4kOperationalHtml_() {
  return '<div style="font-family:Arial,sans-serif;color:#222">' +
    '<h2 style="color:#b42318">TEST — NO ACTION REQUIRED</h2>' +
    '<p>This is a <strong>synthetic test</strong> of the future DVIR closed-record email.</p>' +
    '<p>No employee, vehicle, Audit ID, or workbook record is associated with this message.</p>' +
    '<h3>Production routing after approval</h3><ul>' +
    '<li><strong>To:</strong> Assigned Driver</li>' +
    '<li><strong>CC:</strong> Assigned Driver Supervisor</li></ul>' +
    '<p>For this controlled test, <strong>jo.deleon@vcicom.com</strong> represents both roles and appears only once in To.</p>' +
    '<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse">' +
    '<tr><td><strong>Test Audit ID</strong></td><td>TEST-PHASE4K</td></tr>' +
    '<tr><td><strong>Test Vehicle</strong></td><td>TEST-VEHICLE</td></tr>' +
    '<tr><td><strong>Test Status</strong></td><td>Closed</td></tr></table>' +
    '<p><strong>No action is required.</strong></p></div>';
}

function phase4kValidationText_(validation) {
  return [
    'TEST — NO ACTION REQUIRED', '',
    'This is a synthetic test of the validation-error notification sent only to Jo.',
    'It does not indicate a production failure.', '',
    'Validation status observed during test: ' + validation.status,
    'Blockers observed: ' + (validation.blockers || []).length,
    'Warnings observed: ' + (validation.warnings || []).length, '',
    'In production, unresolved operational records will be suppressed and summarized in this notification.',
    'Valid operational records will not be copied to Jo.', '',
    'No action is required.'
  ].join('\n');
}

function phase4kValidationHtml_(validation) {
  return '<div style="font-family:Arial,sans-serif;color:#222">' +
    '<h2 style="color:#b42318">TEST — NO ACTION REQUIRED</h2>' +
    '<p>This is a <strong>synthetic test</strong> of the validation-error notification sent only to Jo. ' +
    'It does not indicate a production failure.</p>' +
    '<ul><li><strong>Validation status:</strong> ' + validation.status + '</li>' +
    '<li><strong>Blockers:</strong> ' + (validation.blockers || []).length + '</li>' +
    '<li><strong>Warnings:</strong> ' + (validation.warnings || []).length + '</li></ul>' +
    '<p>In production, unresolved operational records will be suppressed and summarized here. ' +
    'Valid operational records will not be copied to Jo.</p>' +
    '<p><strong>No action is required.</strong></p></div>';
}
