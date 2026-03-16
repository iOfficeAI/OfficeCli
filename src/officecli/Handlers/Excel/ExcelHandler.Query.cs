// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (path == "/" || path == "")
        {
            var node = new DocumentNode { Path = "/", Type = "workbook" };
            foreach (var (name, part) in GetWorksheets())
            {
                var sheetNode = new DocumentNode { Path = $"/{name}", Type = "sheet", Preview = name };
                var sheetData = GetSheet(part).GetFirstChild<SheetData>();
                sheetNode.ChildCount = sheetData?.Elements<Row>().Count() ?? 0;

                if (depth > 0 && sheetData != null)
                {
                    sheetNode.Children = GetSheetChildNodes(name, sheetData, depth);
                }

                node.Children.Add(sheetNode);
            }
            node.ChildCount = node.Children.Count;
            return node;
        }

        // Parse path: /SheetName or /SheetName/A1 or /SheetName/A1:D10
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetNameFromPath = segments[0];
        var worksheet = FindWorksheet(sheetNameFromPath);
        if (worksheet == null)
            throw new ArgumentException($"Sheet not found: {sheetNameFromPath}");

        var data = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (data == null)
            return new DocumentNode { Path = path, Type = "sheet", Preview = "(empty)" };

        if (segments.Length == 1)
        {
            // Return sheet overview
            var sheetNode = new DocumentNode
            {
                Path = path,
                Type = "sheet",
                Preview = sheetNameFromPath,
                ChildCount = data.Elements<Row>().Count()
            };
            if (depth > 0)
            {
                sheetNode.Children = GetSheetChildNodes(sheetNameFromPath, data, depth);
            }
            return sheetNode;
        }

        // Cell reference: A1 or range A1:D10
        var cellRef = segments[1];

        // Check if it's a cell reference or a generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = System.Text.RegularExpressions.Regex.IsMatch(firstPart, @"^[A-Z]+\d+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        if (!isCellRef)
        {
            // Generic XML fallback: navigate worksheet XML tree
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"Element not found: {cellRef}" };
            return GenericXmlQuery.ElementToNode(target, path, depth);
        }

        if (cellRef.Contains(':'))
        {
            // Range
            return GetCellRange(sheetNameFromPath, data, cellRef, depth);
        }
        else
        {
            // Single cell
            var cell = FindCell(data, cellRef);
            if (cell == null)
                return new DocumentNode { Path = path, Type = "cell", Text = "(empty)", Preview = cellRef };
            return CellToNode(sheetNameFromPath, cell, worksheet);
        }
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();

        // Check if element type is known (Scheme A) or should fall back to generic XML (Scheme B)
        var elementMatch = Regex.Match(selector.Split('!').Last(), @"^([\w:]+)");
        var elementName = elementMatch.Success ? elementMatch.Groups[1].Value : "";
        bool isKnownType = string.IsNullOrEmpty(elementName)
            || elementName is "cell" or "row" or "sheet"
            || (elementName.Length <= 3 && Regex.IsMatch(elementName, @"^[A-Z]+$", RegexOptions.IgnoreCase));
        if (!isKnownType)
        {
            // Scheme B: generic XML fallback
            var genericParsed = GenericXmlQuery.ParseSelector(selector);
            foreach (var (_, worksheetPart) in GetWorksheets())
            {
                results.AddRange(GenericXmlQuery.Query(
                    GetSheet(worksheetPart), genericParsed.element, genericParsed.attrs, genericParsed.containsText));
            }
            return results;
        }

        var parsed = ParseCellSelector(selector);

        foreach (var (sheetName, worksheetPart) in GetWorksheets())
        {
            // If selector specifies a sheet, skip non-matching sheets
            if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                continue;

            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (MatchesCellSelector(cell, sheetName, parsed))
                    {
                        results.Add(CellToNode(sheetName, cell));
                    }
                }
            }
        }

        return results;
    }
}
