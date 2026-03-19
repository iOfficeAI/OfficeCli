// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Selector ====================

    private record CellSelector(string? Sheet, string? Column, string? ValueEquals, string? ValueNotEquals,
        string? ValueContains, bool? HasFormula, bool? IsEmpty, string? TypeEquals);

    private CellSelector ParseCellSelector(string selector)
    {
        string? sheet = null;
        string? column = null;
        string? valueEquals = null;
        string? valueNotEquals = null;
        string? valueContains = null;
        bool? hasFormula = null;
        bool? isEmpty = null;
        string? typeEquals = null;

        // Check for sheet prefix: Sheet1!cell[...]
        var exclIdx = selector.IndexOf('!');
        if (exclIdx > 0)
        {
            sheet = selector[..exclIdx];
            selector = selector[(exclIdx + 1)..];
        }

        // Parse element and attributes: cell[attr=value]
        var match = Regex.Match(selector, @"^(\w+)?(.*)$");
        var element = match.Groups[1].Value;

        // Column filter: e.g., "B" or "cell" in column context
        if (element.Length <= 3 && Regex.IsMatch(element, @"^[A-Z]+$", RegexOptions.IgnoreCase))
        {
            column = element.ToUpperInvariant();
        }

        // Parse attributes
        foreach (Match attrMatch in Regex.Matches(selector, @"\[(\w+)(!?=)([^\]]*)\]"))
        {
            var key = attrMatch.Groups[1].Value.ToLowerInvariant();
            var op = attrMatch.Groups[2].Value;
            var val = attrMatch.Groups[3].Value.Trim('\'', '"');

            switch (key)
            {
                case "value" when op == "=": valueEquals = val; break;
                case "value" when op == "!=": valueNotEquals = val; break;
                case "type": typeEquals = val; break;
                case "formula": hasFormula = val.ToLowerInvariant() != "false"; break;
                case "empty": isEmpty = val.ToLowerInvariant() != "false"; break;
            }
        }

        // :contains() pseudo-selector
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) valueContains = containsMatch.Groups[1].Value;

        // Shorthand: "cell:text" → treat as :contains(text)
        if (valueContains == null)
        {
            var shorthandMatch = Regex.Match(selector, @"^(?:\w+)?:(?!contains|empty|has)(.+)$");
            if (shorthandMatch.Success) valueContains = shorthandMatch.Groups[1].Value;
        }

        // :empty pseudo-selector
        if (selector.Contains(":empty")) isEmpty = true;

        // :has(formula) pseudo-selector
        if (selector.Contains(":has(formula)")) hasFormula = true;

        return new CellSelector(sheet, column, valueEquals, valueNotEquals, valueContains, hasFormula, isEmpty, typeEquals);
    }

    private bool MatchesCellSelector(Cell cell, string sheetName, CellSelector selector)
    {
        // Column filter
        if (selector.Column != null)
        {
            var cellRef = cell.CellReference?.Value ?? "";
            var (colName, _) = ParseCellReference(cellRef);
            if (!colName.Equals(selector.Column, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        var value = GetCellDisplayValue(cell);

        // Value filters
        if (selector.ValueEquals != null && !value.Equals(selector.ValueEquals, StringComparison.OrdinalIgnoreCase))
            return false;
        if (selector.ValueNotEquals != null && value.Equals(selector.ValueNotEquals, StringComparison.OrdinalIgnoreCase))
            return false;
        if (selector.ValueContains != null && !value.Contains(selector.ValueContains, StringComparison.OrdinalIgnoreCase))
            return false;

        // Formula filter
        if (selector.HasFormula == true && cell.CellFormula == null)
            return false;
        if (selector.HasFormula == false && cell.CellFormula != null)
            return false;

        // Empty filter
        if (selector.IsEmpty == true && !string.IsNullOrEmpty(value))
            return false;
        if (selector.IsEmpty == false && string.IsNullOrEmpty(value))
            return false;

        // Type filter
        if (selector.TypeEquals != null)
        {
            var type = cell.DataType?.Value.ToString() ?? "Number";
            if (!type.Equals(selector.TypeEquals, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    // ==================== Cell Reference Utils ====================

    private static (string Column, int Row) ParseCellReference(string cellRef)
    {
        var match = Regex.Match(cellRef, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success) return ("A", 1);
        return (match.Groups[1].Value.ToUpperInvariant(), int.Parse(match.Groups[2].Value));
    }

    private static int ColumnNameToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpperInvariant())
        {
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }

    private static string IndexToColumnName(int index)
    {
        var result = "";
        while (index > 0)
        {
            index--;
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }
        return result;
    }

    private static DocumentFormat.OpenXml.Packaging.ChartPart GetChartPart(WorksheetPart worksheetPart, int index)
    {
        var drawingsPart = worksheetPart.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/charts");
        var chartParts = drawingsPart.ChartParts.ToList();
        if (index < 1 || index > chartParts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{chartParts.Count})");
        return chartParts[index - 1];
    }

    private DocumentFormat.OpenXml.Packaging.ChartPart GetGlobalChartPart(int index)
    {
        var allCharts = new List<DocumentFormat.OpenXml.Packaging.ChartPart>();
        foreach (var (_, worksheetPart) in GetWorksheets())
        {
            if (worksheetPart.DrawingsPart != null)
                allCharts.AddRange(worksheetPart.DrawingsPart.ChartParts);
        }
        if (allCharts.Count == 0)
            throw new ArgumentException("No charts found in workbook");
        if (index < 1 || index > allCharts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{allCharts.Count})");
        return allCharts[index - 1];
    }
}
