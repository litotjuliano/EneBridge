# Tolerate the Source Excel File Being Open in Excel

## Context

`ExcelReaderService.OpenWorkbook` (`src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs`) opens
the source workbook with `new XLWorkbook(path)`, letting ClosedXML open the file using .NET's
default file-sharing mode. When the source Excel file is already open in Microsoft Excel — a
realistic scenario for a client who opens the invoice file to double-check numbers before running
the tool — that default sharing mode conflicts with the share flags Excel's own open handle
requires, and the read fails with a raw `IOException` ("...being used by another process").

This is not hypothetical: it was reproduced twice in the same working session while testing an
unrelated change, once via an automated UI test and once by the client attempting to run the app
against a file they had open in Excel.

Today this failure is already caught and surfaces without crashing the app — `OpenWorkbook` wraps
it in `ExcelReadException`, and `MainViewModel.TryOpenWorkbook` displays the exception message on
both stage panels and logs it via `FileLogger`. So the "nothing should crash or fail silently"
principle already holds. What's missing is that the raw `.NET` exception text isn't actionable for
a non-technical user, and — more importantly — the failure is avoidable in the common case.

## Goal

In the common case (Excel has the file open normally, not in some exclusive/protected mode), the
app should be able to read the file successfully with no error at all. In the rarer case where the
file is still genuinely inaccessible, show a clear, actionable message instead of the raw
exception text.

Scope: the source Excel file only. The DBF output side (`icmaste.dbf`/`ictrane.dbf` being open in
another program) already fails gracefully via the same "catch and report, don't crash" pattern in
`DbfExportService.Export` and is explicitly out of scope for this change.

## Design

### Open the file with a more permissive share mode

Windows file-sharing rules mean a new file open only succeeds if, for every existing open handle
on the file: (a) the new request's access is a subset of what that handle's share flags allow, and
(b) that handle's access is a subset of what the new request's share flags allow. Excel typically
opens a workbook it has open for editing with `Access=ReadWrite, Share=Read`. `ClosedXML`'s
path-based constructor uses .NET's default read sharing (`FileShare.Read`), which fails check (b)
above — Excel's `ReadWrite` access is not a subset of a `Read`-only share grant from us.

Fix: instead of `new XLWorkbook(path)`, open our own stream explicitly —
`new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)` — and pass that stream
to `XLWorkbook`'s stream-accepting constructor. Requesting `FileShare.ReadWrite` on our end
satisfies check (b) against Excel's typical `ReadWrite` access, while our own `FileAccess.Read`
request is already a subset of Excel's `Share=Read` grant, satisfying check (a). In the common
case, this should let the read succeed with the file still open in Excel — no error, no user
action needed.

### Fallback: a clear message for the cases that still fail

If the file is still inaccessible after this change — Excel opened it in an exclusive/protected
mode, or a different program holds a stricter lock — `OpenWorkbook` should detect specifically a
sharing-violation `IOException` and raise `ExcelReadException` with a clear, actionable message,
e.g.: *"The Excel file is currently open in another program. Please close it and click Run
again."* Detection should check the exception's `HResult` against `ERROR_SHARING_VIOLATION`
(`0x80070020`) rather than matching message text, since the latter is locale-fragile. Any other
kind of I/O failure (missing file, permissions, corruption) is unaffected and keeps today's
existing message path (`"Failed to open Excel file '{path}': {ex.Message}"`).

No retry loop, no polling, no UI changes: this is strictly a one-shot open attempt, same as today
— just one that succeeds more often, and fails more clearly when it doesn't.

### Testing

- The existing `OpenWorkbook_FileIsLockedByAnotherProcess_ThrowsExcelReadExceptionWithInnerException`
  test simulates a lock via `FileShare.None` — a fully exclusive lock, stronger than what this fix
  addresses. It should still correctly throw after this change, so it stays valid, but its expected
  message assertion needs updating to the new friendlier text.
- A new test should simulate Excel's typical open mode
  (`new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read)`) held open by the
  test while `OpenWorkbook` is called against the same path, and assert that the call now succeeds
  (returns a usable `XLWorkbook`) — this is the core behavior this change adds, and the prior test
  suite has no case that exercises it.

## Out of scope

- DBF output file locking (`DbfExportService`) — already handled gracefully, not addressed here.
- Auto-retry / polling while waiting for the user to close Excel.
- Any UI changes — the existing error-display path in `MainViewModel` is reused as-is.
