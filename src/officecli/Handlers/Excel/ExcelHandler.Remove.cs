// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public string? Remove(string path)
    {
        path = NormalizeExcelPath(path);
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        if (segments.Length == 1)
        {
            // Remove entire sheet
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var sheets = GetWorkbook().GetFirstChild<Sheets>();
            var sheet = sheets?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetName}");

            var sheetCount = sheets!.Elements<Sheet>().Count();
            if (sheetCount <= 1)
                throw new InvalidOperationException($"Cannot remove the last sheet. A workbook must contain at least one sheet.");

            var relId = sheet.Id?.Value;
            sheet.Remove();
            if (relId != null)
                workbookPart.DeletePart(workbookPart.GetPartById(relId));

            // Clean up named ranges referencing the deleted sheet
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                var toRemove = definedNames.Elements<DefinedName>()
                    .Where(dn => dn.Text?.Contains(sheetName + "!", StringComparison.OrdinalIgnoreCase) == true)
                    .ToList();
                foreach (var dn in toRemove) dn.Remove();
                if (!definedNames.HasChildren) definedNames.Remove();
            }

            // Fix ActiveTab to prevent workbook corruption when deleting the last tab
            var remainingCount = sheets!.Elements<Sheet>().Count();
            var bookViews = workbook.GetFirstChild<BookViews>();
            if (bookViews != null)
            {
                foreach (var bv in bookViews.Elements<WorkbookView>())
                {
                    if (bv.ActiveTab?.Value >= (uint)remainingCount)
                        bv.ActiveTab = (uint)Math.Max(0, remainingCount - 1);
                }
            }

            workbook.Save();
            return null;
        }

        var cellRef = segments[1];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // row[N] — true shift delete
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = int.Parse(rowMatch.Groups[1].Value);
            sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIdx)
                ?.Remove();
            var affected = CollectFormulaCellsAffectedByRowDelete(worksheet, rowIdx);
            ShiftRowsUp(worksheet, rowIdx);
            SaveWorksheet(worksheet);
            return FormatFormulaWarning(affected);
        }

        // col[X] — true shift delete
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Za-z]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colName = colMatch.Groups[1].Value.ToUpperInvariant();
            var deletedColIdx = ColumnNameToIndex(colName);
            var affected = CollectFormulaCellsAffectedByColDelete(worksheet, deletedColIdx);
            ShiftColumnsLeft(worksheet, colName);
            SaveWorksheet(worksheet);
            return FormatFormulaWarning(affected);
        }

        // Single cell
        var cell = FindCell(sheetData, cellRef)
            ?? throw new ArgumentException($"Cell {cellRef} not found");
        cell.Remove();
        SaveWorksheet(worksheet);
        return null;
    }

    // ==================== Row shift ====================

    private void ShiftRowsUp(WorksheetPart worksheet, int deletedRow)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();

        // 1. Shift all rows after the deleted row: update RowIndex + all CellReferences
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>().ToList())
            {
                var rowIdx = (int)(row.RowIndex?.Value ?? 0);
                if (rowIdx <= deletedRow) continue;

                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value != null)
                    {
                        var (col, _) = ParseCellReference(cell.CellReference.Value);
                        cell.CellReference = $"{col}{rowIdx - 1}";
                    }
                }
                row.RowIndex = (uint)(rowIdx - 1);
            }
        }

        // 2. Merge cells
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells != null)
        {
            foreach (var mc in mergeCells.Elements<MergeCell>().ToList())
            {
                var shifted = ShiftRowInRef(mc.Reference?.Value, deletedRow);
                if (shifted == null) mc.Remove();
                else mc.Reference = shifted;
            }
            if (!mergeCells.HasChildren) mergeCells.Remove();
        }

        // 3. Conditional formatting sqref
        foreach (var cf in ws.Elements<ConditionalFormatting>().ToList())
        {
            if (cf.SequenceOfReferences?.HasValue != true) continue;
            var newRefs = cf.SequenceOfReferences.Items
                .Select(r => ShiftRowInRef(r.Value, deletedRow))
                .OfType<string>().ToList();
            if (newRefs.Count == 0) cf.Remove();
            else cf.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
        }

        // 4. Data validations sqref
        var dvs = ws.GetFirstChild<DataValidations>();
        if (dvs != null)
        {
            foreach (var dv in dvs.Elements<DataValidation>().ToList())
            {
                if (dv.SequenceOfReferences?.HasValue != true) continue;
                var newRefs = dv.SequenceOfReferences.Items
                    .Select(r => ShiftRowInRef(r.Value, deletedRow))
                    .OfType<string>().ToList();
                if (newRefs.Count == 0) dv.Remove();
                else dv.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
            }
            if (!dvs.HasChildren) dvs.Remove();
        }

        // 5. AutoFilter
        var af = ws.GetFirstChild<AutoFilter>();
        if (af?.Reference?.Value != null)
        {
            var shifted = ShiftRowInRef(af.Reference.Value, deletedRow);
            if (shifted != null) af.Reference = shifted;
            else af.Remove();
        }

        // 6. Named ranges (workbook-level)
        ShiftNamedRangeRows(worksheet, deletedRow);
    }

    // ==================== Column shift ====================

    private void ShiftColumnsLeft(WorksheetPart worksheet, string deletedColName)
    {
        var ws = GetSheet(worksheet);
        var deletedColIdx = ColumnNameToIndex(deletedColName);
        var sheetData = ws.GetFirstChild<SheetData>();

        // 1. Remove cells in deleted column, shift remaining cell references left
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>().ToList())
                {
                    if (cell.CellReference?.Value == null) continue;
                    var (col, rowIdx) = ParseCellReference(cell.CellReference.Value);
                    var colIdx = ColumnNameToIndex(col);

                    if (colIdx == deletedColIdx)
                        cell.Remove();
                    else if (colIdx > deletedColIdx)
                        cell.CellReference = $"{IndexToColumnName(colIdx - 1)}{rowIdx}";
                }
            }
        }

        // 2. Column width/style definitions
        var columns = ws.GetFirstChild<Columns>();
        if (columns != null)
        {
            foreach (var col in columns.Elements<Column>().ToList())
            {
                var min = (int)(col.Min?.Value ?? 0);
                var max = (int)(col.Max?.Value ?? 0);

                if (min >= deletedColIdx && max <= deletedColIdx)
                {
                    col.Remove();
                }
                else if (min > deletedColIdx)
                {
                    col.Min = (uint)(min - 1);
                    col.Max = (uint)(max - 1);
                }
                else if (max >= deletedColIdx)
                {
                    col.Max = (uint)(max - 1);
                }
            }
            if (!columns.HasChildren) columns.Remove();
        }

        // 3. Merge cells
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells != null)
        {
            foreach (var mc in mergeCells.Elements<MergeCell>().ToList())
            {
                var shifted = ShiftColInRef(mc.Reference?.Value, deletedColIdx);
                if (shifted == null) mc.Remove();
                else mc.Reference = shifted;
            }
            if (!mergeCells.HasChildren) mergeCells.Remove();
        }

        // 4. Conditional formatting sqref
        foreach (var cf in ws.Elements<ConditionalFormatting>().ToList())
        {
            if (cf.SequenceOfReferences?.HasValue != true) continue;
            var newRefs = cf.SequenceOfReferences.Items
                .Select(r => ShiftColInRef(r.Value, deletedColIdx))
                .OfType<string>().ToList();
            if (newRefs.Count == 0) cf.Remove();
            else cf.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
        }

        // 5. Data validations sqref
        var dvs = ws.GetFirstChild<DataValidations>();
        if (dvs != null)
        {
            foreach (var dv in dvs.Elements<DataValidation>().ToList())
            {
                if (dv.SequenceOfReferences?.HasValue != true) continue;
                var newRefs = dv.SequenceOfReferences.Items
                    .Select(r => ShiftColInRef(r.Value, deletedColIdx))
                    .OfType<string>().ToList();
                if (newRefs.Count == 0) dv.Remove();
                else dv.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
            }
            if (!dvs.HasChildren) dvs.Remove();
        }

        // 6. AutoFilter
        var af = ws.GetFirstChild<AutoFilter>();
        if (af?.Reference?.Value != null)
        {
            var shifted = ShiftColInRef(af.Reference.Value, deletedColIdx);
            if (shifted != null) af.Reference = shifted;
            else af.Remove();
        }

        // 7. Named ranges
        ShiftNamedRangeCols(worksheet, deletedColIdx);
    }

    // ==================== Shift helpers ====================

    /// <summary>
    /// Shift row numbers in a cell/range reference after a row deletion.
    /// Returns null if the reference sits exactly on the deleted row (should be removed).
    /// For ranges: if either endpoint is on the deleted row the range is removed;
    /// endpoints after the deleted row are decremented by 1.
    /// </summary>
    private static string? ShiftRowInRef(string? refStr, int deletedRow)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var (col, row) = ParseCellReference(part);
                if (row == deletedRow) return null;
                shifted.Add(row > deletedRow ? $"{col}{row - 1}" : part);
            }
            catch { shifted.Add(part); }
        }
        return string.Join(":", shifted);
    }

    /// <summary>
    /// Shift column letters in a cell/range reference after a column deletion.
    /// Returns null if the reference sits exactly on the deleted column.
    /// </summary>
    private static string? ShiftColInRef(string? refStr, int deletedColIdx)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var (col, row) = ParseCellReference(part);
                var colIdx = ColumnNameToIndex(col);
                if (colIdx == deletedColIdx) return null;
                shifted.Add(colIdx > deletedColIdx ? $"{IndexToColumnName(colIdx - 1)}{row}" : part);
            }
            catch { shifted.Add(part); }
        }
        return string.Join(":", shifted);
    }

    /// <summary>
    /// Update workbook-level named ranges after a row deletion.
    /// Handles both relative (A1) and absolute ($A$1) references.
    /// Note: formula expressions inside named ranges are not updated.
    /// </summary>
    private void ShiftNamedRangeRows(WorksheetPart worksheet, int deletedRow)
    {
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name;
        if (string.IsNullOrEmpty(sheetName)) return;

        var definedNames = GetWorkbook().GetFirstChild<DefinedNames>();
        if (definedNames == null) return;

        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            if (dn.Text == null) continue;
            dn.Text = ShiftRowNumbersInText(dn.Text, sheetName, deletedRow);
        }
        GetWorkbook().Save();
    }

    /// <summary>
    /// Update workbook-level named ranges after a column deletion.
    /// </summary>
    private void ShiftNamedRangeCols(WorksheetPart worksheet, int deletedColIdx)
    {
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name;
        if (string.IsNullOrEmpty(sheetName)) return;

        var definedNames = GetWorkbook().GetFirstChild<DefinedNames>();
        if (definedNames == null) return;

        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            if (dn.Text == null) continue;
            dn.Text = ShiftColLettersInText(dn.Text, sheetName, deletedColIdx);
        }
        GetWorkbook().Save();
    }

    // ==================== Formula impact detection ====================

    /// <summary>
    /// Find all cells with formulas that reference rows >= deletedRow (will shift or become #REF!).
    /// Returns list of cell references like ["C3", "D5"].
    /// </summary>
    private List<string> CollectFormulaCellsAffectedByRowDelete(WorksheetPart worksheet, int deletedRow)
    {
        var affected = new List<string>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return affected;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var formula = cell.CellFormula?.Text;
                if (string.IsNullOrEmpty(formula)) continue;
                // Formula references a row >= deletedRow
                if (Regex.IsMatch(formula, $@"\$?[A-Z]+\$?({deletedRow}|[{deletedRow + 1}-9]\d*|\d{{2,}})", RegexOptions.IgnoreCase)
                    || Regex.IsMatch(formula, $@"\$?[A-Z]+\$?[{deletedRow}-9]\d*", RegexOptions.IgnoreCase))
                {
                    // Simpler: any row number >= deletedRow in the formula
                    if (FormulaReferencesRowOrAbove(formula, deletedRow))
                        affected.Add(cell.CellReference?.Value ?? "?");
                }
            }
        }
        return affected;
    }

    private static bool FormulaReferencesRowOrAbove(string formula, int deletedRow)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?[A-Z]+\$?(\d+)", RegexOptions.IgnoreCase))
        {
            if (int.TryParse(m.Groups[1].Value, out var row) && row >= deletedRow)
                return true;
        }
        return false;
    }

    /// <summary>
    /// Find all cells with formulas that reference columns >= deletedColIdx.
    /// </summary>
    private List<string> CollectFormulaCellsAffectedByColDelete(WorksheetPart worksheet, int deletedColIdx)
    {
        var affected = new List<string>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return affected;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var formula = cell.CellFormula?.Text;
                if (string.IsNullOrEmpty(formula)) continue;
                if (FormulaReferencesColOrAbove(formula, deletedColIdx))
                    affected.Add(cell.CellReference?.Value ?? "?");
            }
        }
        return affected;
    }

    private static bool FormulaReferencesColOrAbove(string formula, int deletedColIdx)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?([A-Z]+)\$?\d+", RegexOptions.IgnoreCase))
        {
            var colIdx = ColumnNameToIndex(m.Groups[1].Value.ToUpperInvariant());
            if (colIdx >= deletedColIdx) return true;
        }
        return false;
    }

    private static string? FormatFormulaWarning(List<string> affected)
    {
        if (affected.Count == 0) return null;
        var cells = string.Join(", ", affected);
        return $"Warning: {affected.Count} formula cell(s) may reference shifted cells — verify: {cells}";
    }

    /// <summary>
    /// In a formula/reference string like "Sheet1!$A$3:$B$5", decrement row numbers > deletedRow.
    /// Only touches references that belong to the given sheet.
    /// </summary>
    private static string ShiftRowNumbersInText(string text, string sheetName, int deletedRow)
    {
        // Match: optional sheet prefix (Sheet1! or 'Sheet 1'!), optional $, column letters, optional $, row number
        return Regex.Replace(text,
            $@"(?<={Regex.Escape(sheetName)}!\$?[A-Z]+\$?)(\d+)",
            m =>
            {
                var row = int.Parse(m.Value);
                return row > deletedRow ? (row - 1).ToString() : m.Value;
            },
            RegexOptions.IgnoreCase);
    }

    /// <summary>
    /// In a formula/reference string like "Sheet1!$B$1:$D$5", shift column letters > deletedColIdx left by one.
    /// Only touches references that belong to the given sheet.
    /// </summary>
    private static string ShiftColLettersInText(string text, string sheetName, int deletedColIdx)
    {
        return Regex.Replace(text,
            $@"(?<={Regex.Escape(sheetName)}!)\$?([A-Z]+)\$?(\d+)",
            m =>
            {
                var col = m.Groups[1].Value.ToUpperInvariant();
                var row = m.Groups[2].Value;
                var colIdx = ColumnNameToIndex(col);
                if (colIdx <= deletedColIdx) return m.Value;
                var dollar1 = m.Value.StartsWith("$") ? "$" : "";
                var dollar2 = m.Value.Contains("$" + col + "$") ? "$" : "";
                return $"{dollar1}{IndexToColumnName(colIdx - 1)}{dollar2}{row}";
            },
            RegexOptions.IgnoreCase);
    }
}
