/**
 * Phase 4P: final user-facing production entry point.
 *
 * Keep the validated support files installed. This file exposes only the
 * production Update Report command and never calls a preview/test entry point.
 */

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('DVIR Report Automation')
    .addItem('Update Report', 'updateReport')
    .addSeparator()
    .addItem('Validate Workbook Only', 'validatePhase4PWorkbookOnly')
    .addToUi();
}

function installPhase4PMenu() {
  onOpen();
  SpreadsheetApp.getUi().alert([
    'DVIR Report Automation menu installed.', '',
    'Use Update Report only after both new source files are present in intake:',
    '• DVIR_Details',
    '• safety-deviations-MM-DD-YYYY'
  ].join('\n'));
}

function updateReport() {
  return runPhase4OControlledProductionWorkflow();
}

function validatePhase4PWorkbookOnly() {
  const report = validateWorkbookIntegrity('PHASE4P_MANUAL_VALIDATION');
  const validation = report.validation || report;
  Logger.log('PHASE4P_VALIDATION=' + JSON.stringify(validation));
  SpreadsheetApp.getUi().alert([
    'Workbook validation: ' + validation.status,
    'Blockers: ' + (validation.blockers || []).length,
    'Warnings: ' + (validation.warnings || []).length, '',
    'No workbook data was changed and no email was sent.'
  ].join('\n'));
  return validation;
}
