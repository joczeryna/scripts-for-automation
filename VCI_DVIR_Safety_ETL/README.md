# VCI Apps Script Phase 4A — Read-Only Preflight

This checkpoint validates the consolidated automation inputs without changing
production data, archiving files, or sending email.

## Install and run

1. Open the DVIR Report Google Sheet.
2. Open **Extensions → Apps Script**.
3. Add a new script file named `Phase4A_Preflight`.
4. Paste the contents of `Phase4A_Preflight.gs` and save.
5. Select `validateAutomationInputs` from the function dropdown.
6. Click **Run** and approve only the requested Drive/Sheets permissions.
7. Send the alert text or execution log back for review.

Do not replace the existing production script at this stage.
