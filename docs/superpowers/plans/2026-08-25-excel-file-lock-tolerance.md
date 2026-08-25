# Excel File Lock Tolerance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `ExcelReaderService.OpenWorkbook` succeed when the source Excel file is already open in Microsoft Excel (the common real-world case), and show a clear, actionable message instead of a raw `.NET` exception in the rarer case where the file is still genuinely inaccessible.

**Architecture:** A single-method change in `ExcelReaderService.OpenWorkbook`: open the file via our own `FileStream` requesting `FileShare.ReadWrite` (compatible with how Excel holds a file open for editing), copy its bytes into a `MemoryStream`, and construct `XLWorkbook` from that — fully decoupling the returned workbook from the original file handle. A sharing-violation `IOException` (detected by `HResult`, not by matching message text) gets a friendly message; every other failure keeps today's existing message path unchanged.

**Tech Stack:** .NET 8, ClosedXML (`XLWorkbook`), xUnit.

**Spec:** `docs/superpowers/specs/2026-08-25-excel-file-lock-tolerance-design.md`

Both of the following facts were independently verified on this machine before writing this plan (not assumed from general knowledge):
- A sharing-violation `IOException` (opening a file with `FileShare.None` held by another handle) carries `HResult = 0x80070020` (`unchecked((int)0x80070020)` = `-2147024864`) with message "The process cannot access the file '...' because it is being used by another process."
- Opening a file with `FileAccess.Read, FileShare.ReadWrite` succeeds even while another handle holds it open with `FileAccess.ReadWrite, FileShare.Read` — exactly how Excel holds a workbook open for editing. Today's code (`new XLWorkbook(path)`, which uses .NET's default `FileShare.Read`) does NOT succeed against that same scenario — confirmed live in this session, twice, against the real app.

---

### Task 1: Make `OpenWorkbook` tolerate Excel's typical file lock, with a friendly fallback message

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs:15-30`
- Test: `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs`

- [ ] **Step 1: Update the existing exclusive-lock test to expect the new friendly message**

In `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs`, find `OpenWorkbook_FileIsLockedByAnotherProcess_ThrowsExcelReadExceptionWithInnerException` (currently lines 159-179):

```csharp
    [Fact]
    public void OpenWorkbook_FileIsLockedByAnotherProcess_ThrowsExcelReadExceptionWithInnerException()
    {
        var service = new ExcelReaderService();
        var lockedPath = Path.Combine(Path.GetTempPath(), "eneBridge-locked-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(FixturePath, lockedPath);

        var lockStream = new FileStream(lockedPath, FileMode.Open, FileAccess.Read, FileShare.None);
        try
        {
            var ex = Assert.Throws<ExcelReadException>(() => service.OpenWorkbook(lockedPath));

            Assert.NotNull(ex.InnerException);
            Assert.StartsWith("Failed to open Excel file", ex.Message);
        }
        finally
        {
            lockStream.Dispose();
            try { File.Delete(lockedPath); } catch (IOException) { }
        }
    }
```

Change the two assertion lines from:

```csharp
            Assert.NotNull(ex.InnerException);
            Assert.StartsWith("Failed to open Excel file", ex.Message);
```

to:

```csharp
            Assert.NotNull(ex.InnerException);
            Assert.Equal(
                "The Excel file is currently open in another program. Please close it and click Run again.",
                ex.Message);
```

This test opens its simulated lock with `FileShare.None` — fully exclusive, stronger than the fix addresses — so it should still correctly throw after this change; only the expected message text changes.

- [ ] **Step 2: Add a new test for the case the fix actually addresses — Excel's typical open mode**

Add this test to `ExcelReaderServiceTests.cs`, directly after the test from Step 1:

```csharp
    [Fact]
    public void OpenWorkbook_FileIsOpenInExcel_SucceedsAnyway()
    {
        var service = new ExcelReaderService();
        var sharedPath = Path.Combine(Path.GetTempPath(), "eneBridge-excelopen-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(FixturePath, sharedPath);

        // Simulates how Excel holds a file open for editing: ReadWrite access, Read share.
        var excelLikeStream = new FileStream(sharedPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        try
        {
            using var workbook = service.OpenWorkbook(sharedPath);

            Assert.NotNull(workbook);
            Assert.NotEmpty(workbook.Worksheets);
        }
        finally
        {
            excelLikeStream.Dispose();
            try { File.Delete(sharedPath); } catch (IOException) { }
        }
    }
```

- [ ] **Step 3: Run both tests to verify they fail**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~OpenWorkbook_FileIsLockedByAnotherProcess_ThrowsExcelReadExceptionWithInnerException|FullyQualifiedName~OpenWorkbook_FileIsOpenInExcel_SucceedsAnyway"`

Expected: FAIL on both.
- `OpenWorkbook_FileIsLockedByAnotherProcess_ThrowsExcelReadExceptionWithInnerException` fails because the actual message still starts with "Failed to open Excel file" (the old generic text), not the new friendly message.
- `OpenWorkbook_FileIsOpenInExcel_SucceedsAnyway` fails because `OpenWorkbook` currently throws `ExcelReadException` against this scenario instead of succeeding.

- [ ] **Step 4: Implement the fix**

In `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs`, change `OpenWorkbook` (currently lines 15-30) from:

```csharp
    public XLWorkbook OpenWorkbook(string path)
    {
        if (!File.Exists(path))
        {
            throw new ExcelReadException($"Excel file not found: {path}");
        }

        try
        {
            return new XLWorkbook(path);
        }
        catch (Exception ex) when (ex is not ExcelReadException)
        {
            throw new ExcelReadException($"Failed to open Excel file '{path}': {ex.Message}", ex);
        }
    }
```

to:

```csharp
    public XLWorkbook OpenWorkbook(string path)
    {
        if (!File.Exists(path))
        {
            throw new ExcelReadException($"Excel file not found: {path}");
        }

        try
        {
            using var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var memoryStream = new MemoryStream();
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;
            return new XLWorkbook(memoryStream);
        }
        catch (IOException ex) when (ex.HResult == unchecked((int)0x80070020))
        {
            throw new ExcelReadException(
                "The Excel file is currently open in another program. Please close it and click Run again.",
                ex);
        }
        catch (Exception ex) when (ex is not ExcelReadException)
        {
            throw new ExcelReadException($"Failed to open Excel file '{path}': {ex.Message}", ex);
        }
    }
```

Notes for the implementer:
- Requesting `FileShare.ReadWrite` on our own `FileStream` open is what makes this compatible with Excel's typical `Access=ReadWrite, Share=Read` hold on the file — this was verified empirically before writing this plan, not assumed.
- Reading fully into a `MemoryStream` (rather than passing the `FileStream` directly to `XLWorkbook`) decouples the returned workbook from the original file handle's lifetime — the `using` on `fileStream` releases the OS handle immediately after the copy, before `XLWorkbook` even parses the content. Do not wrap `memoryStream` in a `using` — `XLWorkbook` (or its caller, via `using var workbook = ...` at call sites) owns disposing it from here on; disposing it ourselves risks an `ObjectDisposedException` if `XLWorkbook` retains a reference to it internally.
- The `catch (IOException ex) when (ex.HResult == unchecked((int)0x80070020))` clause must come *before* the existing generic `catch (Exception ex) when (ex is not ExcelReadException)` clause (C# evaluates catch clauses in order) — the code above already has them in the correct order.
- `0x80070020` is `ERROR_SHARING_VIOLATION`. This check is on the exception's `HResult`, not its message text, so it isn't affected by OS locale.

- [ ] **Step 5: Run both tests to verify they pass**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~OpenWorkbook_FileIsLockedByAnotherProcess_ThrowsExcelReadExceptionWithInnerException|FullyQualifiedName~OpenWorkbook_FileIsOpenInExcel_SucceedsAnyway"`

Expected: PASS on both.

- [ ] **Step 6: Run the full test suite to check for regressions**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj"`

Expected: PASS for every test, including `OpenWorkbook_FileDoesNotExist_ThrowsExcelReadException` and `OpenWorkbook_FileIsNotValidExcel_ThrowsExcelReadExceptionWithInnerException`, which are unaffected by this change (the first is short-circuited before the `try` block; the second still fails to parse invalid content and falls into the unchanged generic `catch` clause, producing the same "Failed to open Excel file" message as before).

- [ ] **Step 7: Commit**

```bash
git add src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs
git commit -m "Tolerate the source Excel file being open in Excel when reading it"
```

---

## Self-Review Notes

- **Spec coverage:** The spec's "Open the file with a more permissive share mode" section → Step 4's `FileShare.ReadWrite` + `MemoryStream` implementation. The "Fallback: a clear message" section → the `HResult`-based catch clause in the same step, using the exact message text and `HResult` check the spec specifies. The "Testing" section's two required test changes → Steps 1 and 2, using the exact simulated-lock modes (`FileShare.None` for the existing test, `FileAccess.ReadWrite, FileShare.Read` for the new one) the spec calls for. The spec's explicit "no retry loop, no UI changes" constraint is satisfied — this plan touches only `OpenWorkbook` and its tests, nothing in `MainViewModel` or the UI layer.
- **Placeholder scan:** No TBD/TODO; every step shows complete before/after code.
- **Type consistency:** `ExcelReadException`'s constructor signature (`string message, Exception? innerException = null`) was verified against `src/eneBridge.Wpf.Core/Exceptions/ExcelReadException.cs` before use — the plan's calls match it exactly. `XLWorkbook`'s `Stream`-accepting constructor was verified to exist via reflection on the built `ClosedXML.dll` before this plan assumed it.
