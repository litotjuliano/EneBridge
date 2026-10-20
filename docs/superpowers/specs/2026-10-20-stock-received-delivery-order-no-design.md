# Stock Received: Delivery Order No -> icmaste HFIELD1

Date: 2026-10-20. Reference workbook: `sample-19-10-2026.xlsx` (client-provided).

## Goal

The Stock Received template gained a 12th column, "Delivery Order No" (column L, 0-based index 11).
Its value must be written to `HFIELD1` in the master table.

## Scope

- Stock Received workflow only. Invoice is untouched.
- Target is `icmaste.dbf`, the staging file eneBridge writes; EMAS imports it into `icmast.dbf`.
  eneBridge never writes `icmast.dbf` directly.
- `HFIELD1` already exists in `IcmasteSchema` (varchar 40). The 151-column layout does not change.
- `ictrane` is unchanged.

## Decisions

- The column is **optional**. `StockReceivedExcelCol.MinColumnCount` stays at 11, so existing
  11-column workbooks still validate and read.
- A missing column or blank cell leaves `HFIELD1` blank. It never skips a row or adds a skip reason.
- Values longer than 40 characters are truncated via `TextTruncation.Truncate`.

## Changes

1. `Models/StockReceivedExcelCol.cs`: add `DeliveryOrderNo = 11` (doc comment noting it is optional).
2. `Models/IcmasteSchema.cs`: add `const string HField1 = "HFIELD1"` for mapping code to reference.
3. `Services/ExcelReaderService.cs`: `ReadStockReceivedRowFields` reads the column (blank if absent);
   `ReadStockReceivedIcmaste` sets `HFIELD1`.
4. Stock Received preview grid: check whether it uses a fixed column list and include `HFIELD1` if so.
5. `CLAUDE.md`: short note alongside the Item Description 2 paragraph.

## Tests

- Value present, blank cell, and an 11-column workbook (HFIELD1 blank, row still read).
- Truncation at 40 characters.
- Real DBF round-trip confirming HFIELD1 survives export.
