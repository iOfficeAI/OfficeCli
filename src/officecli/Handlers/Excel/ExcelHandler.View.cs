// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int sheetIdx = 0;
        int emitted = 0;
        bool truncated = false;

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            if (truncated) break;
            sb.AppendLine($"=== Sheet: {sheetName} ===");
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            int totalRows = sheetData.Elements<Row>().Count();
            int lineNum = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                lineNum++;
                if (startLine.HasValue && lineNum < startLine.Value) continue;
                if (endLine.HasValue && lineNum > endLine.Value) break;

                if (maxLines.HasValue && emitted >= maxLines.Value)
                {
                    sb.AppendLine($"... (showed {emitted} rows, {totalRows} total in sheet, use --start/--end to view more)");
                    truncated = true;
                    break;
                }

                var cellElements = row.Elements<Cell>();
                if (cols != null)
                    cellElements = cellElements.Where(c => cols.Contains(ParseCellReference(c.CellReference?.Value ?? "A1").Column));
                var cells = cellElements.Select(c => GetCellDisplayValue(c)).ToArray();
                var rowRef = row.RowIndex?.Value ?? (uint)lineNum;
                sb.AppendLine($"[/{sheetName}/row[{rowRef}]] {string.Join("\t", cells)}");
                emitted++;
            }

            sheetIdx++;
            if (sheetIdx < sheets.Count) sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int emitted = 0;
        bool truncated = false;

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            if (truncated) break;
            sb.AppendLine($"=== Sheet: {sheetName} ===");
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            int totalRows = sheetData.Elements<Row>().Count();
            int lineNum = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                lineNum++;
                if (startLine.HasValue && lineNum < startLine.Value) continue;
                if (endLine.HasValue && lineNum > endLine.Value) break;

                if (maxLines.HasValue && emitted >= maxLines.Value)
                {
                    sb.AppendLine($"... (showed {emitted} rows, {totalRows} total in sheet, use --start/--end to view more)");
                    truncated = true;
                    break;
                }

                var cellElements = row.Elements<Cell>();
                if (cols != null)
                    cellElements = cellElements.Where(c => cols.Contains(ParseCellReference(c.CellReference?.Value ?? "A1").Column));

                foreach (var cell in cellElements)
                {
                    var cellRef = cell.CellReference?.Value ?? "?";
                    var value = GetCellDisplayValue(cell);
                    var formula = cell.CellFormula?.Text;
                    var type = cell.DataType?.Value.ToString() ?? "Number";

                    var annotation = formula != null ? $"={formula}" : type;
                    var warn = "";

                    if (string.IsNullOrEmpty(value) && formula == null)
                        warn = " \u26a0 empty";
                    else if (formula != null && (value == "#REF!" || value == "#VALUE!" || value == "#NAME?"))
                        warn = " \u26a0 formula error";

                    sb.AppendLine($"  {cellRef}: [{value}] \u2190 {annotation}{warn}");
                }
                emitted++;
            }
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return "(empty workbook)";

        var sheets = workbook.GetFirstChild<Sheets>();
        if (sheets == null) return "(no sheets)";

        sb.AppendLine($"File: {Path.GetFileName(_filePath)}");

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var sheetId = sheet.Id?.Value;
            if (sheetId == null) continue;

            var worksheetPart = (WorksheetPart)_doc.WorkbookPart!.GetPartById(sheetId);
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();

            int rowCount = sheetData?.Elements<Row>().Count() ?? 0;
            int colCount = sheetData?.Elements<Row>().FirstOrDefault()?.Elements<Cell>().Count() ?? 0;

            int formulaCount = 0;
            if (sheetData != null)
            {
                formulaCount = sheetData.Descendants<CellFormula>().Count();
            }

            var formulaInfo = formulaCount > 0 ? $", {formulaCount} formula(s)" : "";
            sb.AppendLine($"\u251c\u2500\u2500 \"{name}\" ({rowCount} rows \u00d7 {colCount} cols{formulaInfo})");
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsStats()
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int totalCells = 0;
        int emptyCells = 0;
        int formulaCells = 0;
        int errorCells = 0;
        var typeCounts = new Dictionary<string, int>();

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    totalCells++;
                    var value = GetCellDisplayValue(cell);
                    if (string.IsNullOrEmpty(value)) emptyCells++;
                    if (cell.CellFormula != null) formulaCells++;
                    if (value is "#REF!" or "#VALUE!" or "#NAME?" or "#DIV/0!") errorCells++;

                    var type = cell.DataType?.Value.ToString() ?? "Number";
                    typeCounts[type] = typeCounts.GetValueOrDefault(type) + 1;
                }
            }
        }

        sb.AppendLine($"Sheets: {sheets.Count}");
        sb.AppendLine($"Total Cells: {totalCells}");
        sb.AppendLine($"Empty Cells: {emptyCells}");
        sb.AppendLine($"Formula Cells: {formulaCells}");
        sb.AppendLine($"Error Cells: {errorCells}");
        sb.AppendLine();
        sb.AppendLine("Data Type Distribution:");
        foreach (var (type, count) in typeCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {type}: {count}");

        return sb.ToString().TrimEnd();
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueNum = 0;

        var sheets = GetWorksheets();
        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value ?? "?";
                    var value = GetCellDisplayValue(cell);

                    if (cell.CellFormula != null && value is "#REF!" or "#VALUE!" or "#NAME?" or "#DIV/0!")
                    {
                        issues.Add(new DocumentIssue
                        {
                            Id = $"F{++issueNum}",
                            Type = IssueType.Content,
                            Severity = IssueSeverity.Error,
                            Path = $"{sheetName}!{cellRef}",
                            Message = $"Formula error: {value}",
                            Context = $"={cell.CellFormula.Text}"
                        });
                    }

                    if (limit.HasValue && issues.Count >= limit.Value) break;
                }
                if (limit.HasValue && issues.Count >= limit.Value) break;
            }
        }

        return issues;
    }
}
