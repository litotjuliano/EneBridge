# Changelog

## [1.0.5] - 2026-09-16
- Fix duplicate-document check silently letting a re-export through when EMAS has the data files open (now blocks and asks you to close EMAS and retry)

## [1.0.4] - 2026-09-16
- Fix Stock Received not saving Qty/Price/Amount values
- Detect and warn when the wrong Excel file is used on a tab (Invoice vs Stock Received)
- Fix a crash on Confirm & Export caused by the duplicate-document check
- Fix duplicate-document check incorrectly blocking re-export of a deleted document
- Revert DBF export to back up and recreate each run instead of appending, fixing duplicated line items in EMAS

## [1.0.3] - 2026-09-14
- Stock Received Module

## [1.0.2] - 2026-09-08
- icmast and ictran exist validations, excel driver validations

## [1.0.1] - 2026-08-28
- Adding DBF Viewer for verifications

## [1.0.0] - 2026-08-17
- Initial release.
