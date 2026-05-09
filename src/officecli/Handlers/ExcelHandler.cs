// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler : IDocumentHandler
{
    private readonly SpreadsheetDocument _doc;
    private readonly string _filePath;
    private readonly HashSet<string> _initialSheetNames;
    private readonly HashSet<WorksheetPart> _dirtyWorksheets = new();
    private bool _dirtyStylesheet;
    private bool _disposed;
    // Row index cache: SheetData → sorted map of rowIndex → Row.
    // Turns the O(n) linear scan in FindOrCreateCell into O(1) lookup + O(log n) insert.
    // Invalidated by InvalidateRowIndex() whenever rows are structurally modified (shift, remove).
    private Dictionary<SheetData, SortedList<uint, Row>>? _rowIndex;
    public int LastFindMatchCount { get; internal set; }

    public ExcelHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        try
        {
            _doc = SpreadsheetDocument.Open(filePath, editable);
            // Force early validation: access WorkbookPart to catch corrupt packages now
            _ = _doc.WorkbookPart?.Workbook;
            // Capture initial sheet names to detect duplicate additions
            _initialSheetNames = new HashSet<string>(
                GetWorkbook().GetFirstChild<Sheets>()?.Elements<Sheet>()
                    .Select(s => s.Name?.Value ?? "")
                    .Where(n => !string.IsNullOrEmpty(n)) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);
        }
        catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException ex)
        {
            throw new InvalidOperationException(
                $"Cannot open {Path.GetFileName(filePath)}: {ex.Message}", ex);
        }
    }

    // ==================== Raw Layer ====================

    // CONSISTENCY(zip-uri-lookup): any partPath ending in `.xml` is treated
    // as a literal zip-internal URI and resolved via `RawXmlHelper.FindPartByZipUri`.
    // This replaces the old hand-curated alias table (workbook/styles/...) which
    // could never cover everything — sheet/slide-scoped parts, footnotes,
    // custom XML, etc. were all unreachable. Semantic short names (`/workbook`,
    // `/Sheet1`, `/chart[N]`) continue to route through the switch below.

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return "(empty)";

        // Zip-URI form: any path ending in .xml or .rels is resolved literally
        // against the package. No alias table needed.
        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            // CONSISTENCY(zip-uri-row-filter): if the resolved part is a
            // worksheet AND the caller asked for row/column filtering,
            // route through the same filter as the semantic /SheetName
            // path. Without this, --start/--end/--cols would be silently
            // ignored on zip-URI worksheet reads.
            if ((startRow.HasValue || endRow.HasValue || cols != null)
                && RawXmlHelper.FindPartByZipUri(_doc, partPath) is WorksheetPart wsp)
            {
                return RawSheetWithFilter(wsp, startRow, endRow, cols);
            }

            var xml = RawXmlHelper.TryReadByZipUri(_doc, _filePath, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/workbook, /Sheet1, /chart[N]) for stable identification.");
            return xml;
        }

        if (partPath == "/" || partPath == "/workbook")
            return workbookPart.Workbook?.OuterXml ?? "(empty)";

        if (partPath == "/styles")
        {
            // Raw is read-only; do not create the part if missing (would fail
            // when the package is opened read-only).
            var stylesPart = workbookPart.WorkbookStylesPart;
            return stylesPart?.Stylesheet?.OuterXml ?? "(no styles)";
        }

        if (partPath == "/sharedstrings")
        {
            var sst = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            return sst?.SharedStringTable?.OuterXml ?? "(no shared strings)";
        }

        if (partPath == "/theme")
        {
            var themePart = workbookPart.ThemePart;
            return themePart?.Theme?.OuterXml ?? "(no theme)";
        }

        // Drawing part: /SheetName/drawing
        var drawingMatch = Regex.Match(partPath, @"^/(.+)/drawing$");
        if (drawingMatch.Success)
        {
            var drawSheetName = drawingMatch.Groups[1].Value;
            var drawWs = FindWorksheet(drawSheetName)
                ?? throw SheetNotFoundException(drawSheetName);
            var dp = drawWs.DrawingsPart
                ?? throw new ArgumentException($"Sheet '{drawSheetName}' has no drawings");
            return dp.WorksheetDrawing!.OuterXml;
        }

        // Chart part: /SheetName/chart[N] or /chart[N]
        var chartMatch = Regex.Match(partPath, @"^/(.+)/chart\[(\d+)\]$");
        if (chartMatch.Success)
        {
            var chartSheetName = chartMatch.Groups[1].Value;
            var chartIdx = int.Parse(chartMatch.Groups[2].Value);
            var chartWs = FindWorksheet(chartSheetName)
                ?? throw SheetNotFoundException(chartSheetName);
            var chartPart = GetChartPart(chartWs, chartIdx);
            return chartPart.ChartSpace!.OuterXml;
        }

        // Global chart: /chart[N] — searches all sheets
        var globalChartMatch = Regex.Match(partPath, @"^/chart\[(\d+)\]$");
        if (globalChartMatch.Success)
        {
            var chartIdx = int.Parse(globalChartMatch.Groups[1].Value);
            var chartPart = GetGlobalChartPart(chartIdx);
            return chartPart.ChartSpace!.OuterXml;
        }

        // Try as sheet name
        var sheetName = partPath.TrimStart('/');
        var worksheet = FindWorksheet(sheetName);
        if (worksheet != null)
        {
            if (startRow.HasValue || endRow.HasValue || cols != null)
                return RawSheetWithFilter(worksheet, startRow, endRow, cols);
            return GetSheet(worksheet).OuterXml;
        }

        // /SheetName/<relId> fallback — resolve a worksheet relationship by id
        // (covers OLE embed parts, image parts, etc. that have no named path).
        // Open XML SDK generates relIds like "rId12" or "Rff3244f593f8481a";
        // accept both forms (any non-slash token starting with R/r).
        var relIdMatch = Regex.Match(partPath, @"^/([^/]+)/([Rr][A-Za-z0-9]+)$");
        if (relIdMatch.Success)
        {
            var relSheetName = relIdMatch.Groups[1].Value;
            var relId = relIdMatch.Groups[2].Value;
            var relWs = FindWorksheet(relSheetName)
                ?? throw SheetNotFoundException(relSheetName);
            try
            {
                var part = relWs.GetPartById(relId);
                if (part != null)
                {
                    var ct = part.ContentType ?? "";
                    bool isText = ct.Contains("xml", StringComparison.OrdinalIgnoreCase)
                        || ct.StartsWith("text/", StringComparison.OrdinalIgnoreCase);
                    using var partStream = part.GetStream();
                    if (isText)
                    {
                        using var reader = new StreamReader(partStream);
                        return reader.ReadToEnd();
                    }
                    long size = 0;
                    try { size = partStream.Length; } catch { /* non-seekable */ }
                    return $"(binary part: {ct}, {size} bytes)";
                }
            }
            catch (KeyNotFoundException)
            {
                // fall through to the unknown-part error
            }
            catch (ArgumentException)
            {
                // fall through to the unknown-part error
            }
        }

        throw new ArgumentException($"Unknown part: {partPath}. Available: /workbook, /styles, /sharedstrings, /theme, /<SheetName>, /<SheetName>/drawing, /<SheetName>/chart[N], /chart[N], /<SheetName>/<relId>");
    }

    private static string RawSheetWithFilter(WorksheetPart worksheetPart, int? startRow, int? endRow, HashSet<string>? cols)
    {
        var worksheet = GetSheet(worksheetPart);
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData == null)
            return worksheet.OuterXml;

        var cloned = (Worksheet)worksheet.CloneNode(true);
        var clonedSheetData = cloned.GetFirstChild<SheetData>()!;
        clonedSheetData.RemoveAllChildren();

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowNum = (int)row.RowIndex!.Value;
            if (startRow.HasValue && rowNum < startRow.Value) continue;
            if (endRow.HasValue && rowNum > endRow.Value) break;

            if (cols != null)
            {
                var filteredRow = (Row)row.CloneNode(false);
                filteredRow.RowIndex = row.RowIndex;
                foreach (var cell in row.Elements<Cell>())
                {
                    var colName = ParseCellReference(cell.CellReference?.Value ?? "A1").Column;
                    if (cols.Contains(colName))
                        filteredRow.AppendChild(cell.CloneNode(true));
                }
                clonedSheetData.AppendChild(filteredRow);
            }
            else
            {
                clonedSheetData.AppendChild(row.CloneNode(true));
            }
        }

        return cloned.OuterXml;
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("No workbook part");

        // Zip-URI form: resolve via package part tree, mutate the part's XML
        // stream directly (no SDK typed root needed — handles arbitrary XML
        // parts like footnotes, customXml, untyped sheet1.xml, etc.).
        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"because it ends in .xml, but no matching part exists in the package. " +
                    $"Use semantic paths (/workbook, /Sheet1, /chart[N]) for stable identification.");
            RawXmlHelper.Execute(part, xpath, action, xml);
            return;
        }

        OpenXmlPartRootElement rootElement;
        if (partPath is "/" or "/workbook")
        {
            rootElement = workbookPart.Workbook
                ?? throw new InvalidOperationException("No workbook");
        }
        else if (partPath == "/styles")
        {
            var styleManager = new ExcelStyleManager(workbookPart);
            rootElement = styleManager.EnsureStylesPart().Stylesheet!;
        }
        else if (partPath == "/sharedstrings")
        {
            var sst = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                ?? throw new InvalidOperationException("No shared strings");
            rootElement = sst.SharedStringTable!;
        }
        else if (partPath == "/theme")
        {
            rootElement = workbookPart.ThemePart?.Theme
                ?? throw new ArgumentException("No theme part");
        }
        else
        {
            // Drawing part: /SheetName/drawing
            var drawingMatch = Regex.Match(partPath, @"^/(.+)/drawing$");
            if (drawingMatch.Success)
            {
                var drawSheetName = drawingMatch.Groups[1].Value;
                var drawWs = FindWorksheet(drawSheetName)
                    ?? throw SheetNotFoundException(drawSheetName);
                var dp = drawWs.DrawingsPart
                    ?? throw new ArgumentException($"Sheet '{drawSheetName}' has no drawings");
                rootElement = dp.WorksheetDrawing!;
            }
            else
            {
            // Chart part: /SheetName/chart[N] or /chart[N]
            var chartMatch = Regex.Match(partPath, @"^/(.+)/chart\[(\d+)\]$");
            if (chartMatch.Success)
            {
                var chartSheetName = chartMatch.Groups[1].Value;
                var chartIdx = int.Parse(chartMatch.Groups[2].Value);
                var chartWs = FindWorksheet(chartSheetName)
                    ?? throw SheetNotFoundException(chartSheetName);
                var chartPart = GetChartPart(chartWs, chartIdx);
                rootElement = chartPart.ChartSpace!;
            }
            else
            {
                var globalChartMatch = Regex.Match(partPath, @"^/chart\[(\d+)\]$");
                if (globalChartMatch.Success)
                {
                    var chartIdx = int.Parse(globalChartMatch.Groups[1].Value);
                    var chartPart = GetGlobalChartPart(chartIdx);
                    rootElement = chartPart.ChartSpace!;
                }
                else
                {
                    // Try as sheet name
                    var sheetName = partPath.TrimStart('/');
                    var worksheet = FindWorksheet(sheetName)
                        ?? throw new ArgumentException($"Unknown part: {partPath}. Available: /workbook, /styles, /sharedstrings, /theme, /<SheetName>, /<SheetName>/chart[N], /chart[N]");
                    rootElement = GetSheet(worksheet);
                }
            }
            }
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        // BUG-R5-01: silent — CLI wrappers print their own structured message.
        _ = affected;
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try { FlushDirtyParts(); }
        finally { _doc.Dispose(); }
    }

}
