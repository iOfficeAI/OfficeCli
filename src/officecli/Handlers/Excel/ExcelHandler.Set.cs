// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;


namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        // Parse path: /SheetName/A1
        var segments = path.TrimStart('/').Split('/', 2);
        if (segments.Length < 2)
            throw new ArgumentException($"Path must include sheet and cell reference: /SheetName/A1");

        var sheetName = segments[0];
        var cellRef = segments[1];

        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null)
            throw new ArgumentException($"Sheet not found: {sheetName}");

        // Check if path is a cell reference or generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = System.Text.RegularExpressions.Regex.IsMatch(firstPart, @"^[A-Z]+\d+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!isCellRef)
        {
            // Generic XML fallback: navigate to element and set attributes
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                throw new ArgumentException($"Element not found: {cellRef}");
            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            GetSheet(worksheet).Save();
            return unsup;
        }

        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            GetSheet(worksheet).Append(sheetData);
        }

        var cell = FindOrCreateCell(sheetData, cellRef);

        // Separate content props from style props
        var styleProps = new Dictionary<string, string>();
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            if (ExcelStyleManager.IsStyleKey(key))
            {
                styleProps[key] = value;
                continue;
            }

            switch (key.ToLowerInvariant())
            {
                case "value":
                    cell.CellValue = new CellValue(value);
                    // Auto-detect type
                    if (double.TryParse(value, out _))
                        cell.DataType = null; // Number is default
                    else
                    {
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                    break;
                case "formula":
                    cell.CellFormula = new CellFormula(value);
                    cell.CellValue = null;
                    break;
                case "type":
                    cell.DataType = value.ToLowerInvariant() switch
                    {
                        "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                        _ => cell.DataType
                    };
                    break;
                case "clear":
                    cell.CellValue = null;
                    cell.CellFormula = null;
                    break;
                case "link":
                {
                    var ws = GetSheet(worksheet);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        hyperlinksEl?.Elements<Hyperlink>()
                            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                            .ToList().ForEach(h => h.Remove());
                    }
                    else
                    {
                        var hlRel = worksheet.AddHyperlinkRelationship(new Uri(value), isExternal: true);
                        if (hyperlinksEl == null)
                        {
                            hyperlinksEl = new Hyperlinks();
                            ws.AppendChild(hyperlinksEl);
                        }
                        hyperlinksEl.Elements<Hyperlink>()
                            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                            .ToList().ForEach(h => h.Remove());
                        hyperlinksEl.AppendChild(new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id });
                    }
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                        unsupported.Add(key);
                    break;
            }
        }

        // Apply style properties if any
        if (styleProps.Count > 0)
        {
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var styleManager = new ExcelStyleManager(workbookPart);
            cell.StyleIndex = styleManager.ApplyStyle(cell, styleProps);
        }

        GetSheet(worksheet).Save();
        return unsupported;
    }
}
