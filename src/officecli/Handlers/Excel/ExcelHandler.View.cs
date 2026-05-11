// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;

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
            var evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
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
                var cells = cellElements.Select(c => GetCellDisplayValue(c, evaluator)).ToArray();
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
                    var type = GetCellTypeName(cell);

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
            var worksheet = GetSheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>();

            int rowCount = sheetData?.Elements<Row>().Count() ?? 0;
            int colCount = GetSheetColumnCount(worksheet, sheetData);

            int formulaCount = 0;
            if (sheetData != null)
            {
                formulaCount = sheetData.Descendants<CellFormula>().Count();
            }

            var formulaInfo = formulaCount > 0 ? $", {formulaCount} formula(s)" : "";

            // Pivot tables are stored as pivotTableDefinition XML; their rendered cells
            // are NOT materialized into sheetData (Excel/Calc re-render from pivotCacheRecords
            // at display time). Without this hint, a pivot-only sheet looks like "0 rows × 0 cols"
            // and users think it's empty. Surface the pivot count explicitly — same strategy POI
            // takes via XSSFSheet.getPivotTables(). See also: query pivottable.
            int pivotCount = worksheetPart.PivotTableParts.Count();
            var pivotInfo = pivotCount > 0 ? $", {pivotCount} pivot table(s)" : "";

            int oleCount = CountSheetOleObjects(worksheetPart);
            var oleInfo = oleCount > 0 ? $", {oleCount} ole object(s)" : "";

            sb.AppendLine($"\u251c\u2500\u2500 \"{name}\" ({rowCount} rows \u00d7 {colCount} cols{formulaInfo}{pivotInfo}{oleInfo})");
        }

        return sb.ToString().TrimEnd();
    }

    // CONSISTENCY(ole-stats): per-sheet OLE counter shared by outline and
    // outlineJson. Same dedup rule as ViewAsStats — referenced oleObject
    // elements count once, orphan embedded/package parts add extras.
    private int CountSheetOleObjects(WorksheetPart worksheetPart)
    {
        int count = 0;
        var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var oleEl in GetSheet(worksheetPart).Descendants<OleObject>())
        {
            count++;
            if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                referenced.Add(rid);
        }
        count += worksheetPart.EmbeddedObjectParts.Count(p => !referenced.Contains(worksheetPart.GetIdOfPart(p)));
        count += worksheetPart.EmbeddedPackageParts.Count(p => !referenced.Contains(worksheetPart.GetIdOfPart(p)));
        return count;
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

                    var type = GetCellTypeName(cell);
                    typeCounts[type] = typeCounts.GetValueOrDefault(type) + 1;
                }
            }
        }

        // OLE object count across all sheets. Same dedup rule as
        // CollectOleNodesForSheet: referenced parts count as one entry
        // (via their oleObject element), orphan parts add extras.
        int oleCount = 0;
        foreach (var (_, worksheetPart) in sheets)
        {
            var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var oleEl in GetSheet(worksheetPart).Descendants<OleObject>())
            {
                oleCount++;
                if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                    referenced.Add(rid);
            }
            oleCount += worksheetPart.EmbeddedObjectParts.Count(p => !referenced.Contains(worksheetPart.GetIdOfPart(p)));
            oleCount += worksheetPart.EmbeddedPackageParts.Count(p => !referenced.Contains(worksheetPart.GetIdOfPart(p)));
        }

        sb.AppendLine($"Sheets: {sheets.Count}");
        sb.AppendLine($"Total Cells: {totalCells}");
        sb.AppendLine($"Empty Cells: {emptyCells}");
        sb.AppendLine($"Formula Cells: {formulaCells}");
        sb.AppendLine($"Error Cells: {errorCells}");
        if (oleCount > 0) sb.AppendLine($"OLE Objects: {oleCount}");
        sb.AppendLine();
        sb.AppendLine("Data Type Distribution:");
        foreach (var (type, count) in typeCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {type}: {count}");

        return sb.ToString().TrimEnd();
    }

    public JsonNode ViewAsStatsJson()
    {
        var sheets = GetWorksheets();
        int totalCells = 0, emptyCells = 0, formulaCells = 0, errorCells = 0;
        var typeCounts = new Dictionary<string, int>();

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
                foreach (var cell in row.Elements<Cell>())
                {
                    totalCells++;
                    var value = GetCellDisplayValue(cell);
                    if (string.IsNullOrEmpty(value)) emptyCells++;
                    if (cell.CellFormula != null) formulaCells++;
                    if (value is "#REF!" or "#VALUE!" or "#NAME?" or "#DIV/0!") errorCells++;
                    var type = GetCellTypeName(cell);
                    typeCounts[type] = typeCounts.GetValueOrDefault(type) + 1;
                }
        }

        int oleCountJson = 0;
        foreach (var (_, worksheetPart) in sheets)
        {
            var refSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var oleEl in GetSheet(worksheetPart).Descendants<OleObject>())
            {
                oleCountJson++;
                if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                    refSet.Add(rid);
            }
            oleCountJson += worksheetPart.EmbeddedObjectParts.Count(p => !refSet.Contains(worksheetPart.GetIdOfPart(p)));
            oleCountJson += worksheetPart.EmbeddedPackageParts.Count(p => !refSet.Contains(worksheetPart.GetIdOfPart(p)));
        }

        var result = new JsonObject
        {
            ["sheets"] = sheets.Count,
            ["totalCells"] = totalCells,
            ["emptyCells"] = emptyCells,
            ["formulaCells"] = formulaCells,
            ["errorCells"] = errorCells,
            ["oleObjects"] = oleCountJson,
        };

        var types = new JsonObject();
        foreach (var (type, count) in typeCounts.OrderByDescending(kv => kv.Value))
            types[type] = count;
        result["dataTypeDistribution"] = types;

        return result;
    }

    public JsonNode ViewAsOutlineJson()
    {
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return new JsonObject();

        var sheetsEl = workbook.GetFirstChild<Sheets>();
        if (sheetsEl == null) return new JsonObject { ["fileName"] = Path.GetFileName(_filePath), ["sheets"] = new JsonArray() };

        var sheetsArray = new JsonArray();
        foreach (var sheet in sheetsEl.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var sheetId = sheet.Id?.Value;
            if (sheetId == null) continue;

            var worksheetPart = (WorksheetPart)_doc.WorkbookPart!.GetPartById(sheetId);
            var worksheet = GetSheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>();
            int rowCount = sheetData?.Elements<Row>().Count() ?? 0;
            int colCount = GetSheetColumnCount(worksheet, sheetData);
            int formulaCount = sheetData?.Descendants<CellFormula>().Count() ?? 0;

            int oleCount = CountSheetOleObjects(worksheetPart);
            var sheetObj = new JsonObject
            {
                ["name"] = name,
                ["rows"] = rowCount,
                ["cols"] = colCount,
                ["formulas"] = formulaCount,
                ["oleObjects"] = oleCount
            };
            sheetsArray.Add((JsonNode)sheetObj);
        }

        return new JsonObject
        {
            ["fileName"] = Path.GetFileName(_filePath),
            ["sheets"] = sheetsArray
        };
    }

    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sheetsArray = new JsonArray();
        var worksheets = GetWorksheets();
        int emitted = 0;
        bool truncated = false;

        foreach (var (sheetName, worksheetPart) in worksheets)
        {
            if (truncated) break;
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            var rowsArray = new JsonArray();
            int lineNum = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                lineNum++;
                if (startLine.HasValue && lineNum < startLine.Value) continue;
                if (endLine.HasValue && lineNum > endLine.Value) break;
                if (maxLines.HasValue && emitted >= maxLines.Value) { truncated = true; break; }

                var cellElements = row.Elements<Cell>();
                if (cols != null)
                    cellElements = cellElements.Where(c => cols.Contains(ParseCellReference(c.CellReference?.Value ?? "A1").Column));

                var cellsObj = new JsonObject();
                foreach (var cell in cellElements)
                {
                    var cellRef = cell.CellReference?.Value ?? "?";
                    cellsObj[cellRef] = GetCellDisplayValue(cell);
                }

                var rowRef = row.RowIndex?.Value ?? (uint)lineNum;
                rowsArray.Add((JsonNode)new JsonObject
                {
                    ["row"] = (int)rowRef,
                    ["cells"] = cellsObj
                });
                emitted++;
            }

            sheetsArray.Add((JsonNode)new JsonObject
            {
                ["name"] = sheetName,
                ["rows"] = rowsArray
            });
        }

        return new JsonObject { ["sheets"] = sheetsArray };
    }

    private static int GetSheetColumnCount(Worksheet worksheet, SheetData? sheetData)
    {
        // Try SheetDimension first (e.g., <dimension ref="A1:F20"/>)
        var dimRef = worksheet.GetFirstChild<SheetDimension>()?.Reference?.Value;
        if (!string.IsNullOrEmpty(dimRef))
        {
            var parts = dimRef.Split(':');
            if (parts.Length == 2)
            {
                var endRef = parts[1];
                var col = new string(endRef.TakeWhile(char.IsLetter).ToArray());
                if (!string.IsNullOrEmpty(col))
                    return ColumnNameToIndex(col);
            }
            // Single-cell dimension like "A1" means 1 column
            if (parts.Length == 1)
            {
                var col = new string(parts[0].TakeWhile(char.IsLetter).ToArray());
                if (!string.IsNullOrEmpty(col))
                    return ColumnNameToIndex(col);
            }
        }

        // Fallback: scan all rows for max cell count
        if (sheetData == null) return 0;
        int maxCols = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            var count = row.Elements<Cell>().Count();
            if (count > maxCols) maxCols = count;
        }
        return maxCols;
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueNum = 0;
        // Reset the per-invocation worksheet cache so a long-lived handler
        // sees sheet add/rename/delete between successive calls.
        _viewAsIssuesWorksheetCache = null;
        _viewAsIssuesSheetNameCache = null;

        var sheets = GetWorksheets();
        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            var evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);

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
                    else if (cell.CellFormula?.Text is { } fText
                        && string.IsNullOrEmpty(cell.CellValue?.Text)
                        && (issueType == null || issueType == "formula_not_evaluated"))
                    {
                        // Route through the shared EvaluateForReport so view
                        // text / view issues / Format["evaluated"] all agree on
                        // what counts as "evaluator gave up".
                        var report = evaluator.EvaluateForReport(fText);
                        if (report.Status == Core.EvalReportStatus.NotEvaluated)
                        {
                            issues.Add(new DocumentIssue
                            {
                                Id = $"U{++issueNum}",
                                Type = IssueType.Content,
                                Subtype = "formula_not_evaluated",
                                Severity = IssueSeverity.Warning,
                                Path = $"{sheetName}!{cellRef}",
                                Message = "Formula written but not evaluated (no cachedValue, evaluator unsupported)",
                                Context = $"={fText}"
                            });
                        }
                    }

                    if (limit.HasValue && issues.Count >= limit.Value) break;
                }
                if (limit.HasValue && issues.Count >= limit.Value) break;
            }
        }

        // Defined names whose body references a sheet that no longer exists.
        // Excel persists the stale ref (or writes #REF!) and silently returns
        // 0 in any formula using the name — see ResolveSheetCellResult. The
        // B3 fix in d535587a catches literal `<definedName>#REF!</definedName>`
        // bodies; this scanner catches the still-formula-shaped form like
        // `<definedName>Sheet99!A1:B3</definedName>` where Sheet99 was deleted
        // before the name was cleaned up.
        var workbook = _doc.WorkbookPart?.Workbook;
        var definedNames = workbook?.DefinedNames?.Elements<DefinedName>();
        if (definedNames != null)
        {
            foreach (var dn in definedNames)
            {
                if (limit.HasValue && issues.Count >= limit.Value) break;
                var body = dn.Text?.Trim();
                var name = dn.Name?.Value;
                if (string.IsNullOrEmpty(body) || string.IsNullOrEmpty(name)) continue;
                // Body that is an error literal (#REF!) is already handled
                // by the evaluator's TT.Error path (B3 fix) — that branch
                // propagates the error to formulas. Surface it as an issue
                // too so it's discoverable.
                if (body.StartsWith('#') && body.EndsWith('!'))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"D{++issueNum}",
                        Type = IssueType.Content,
                        Subtype = "definedname_broken",
                        Severity = IssueSeverity.Error,
                        Path = $"/namedrange[{name}]",
                        Message = $"Defined name '{name}' has error body {body}",
                        Context = body,
                        Suggestion = "Rebind to a valid range or remove the name."
                    });
                    continue;
                }
                if (!ChartRefSheetExists(body, out var missingSheet)) continue;
                issues.Add(new DocumentIssue
                {
                    Id = $"D{++issueNum}",
                    Type = IssueType.Content,
                    Subtype = "definedname_target_missing",
                    Severity = IssueSeverity.Error,
                    Path = $"/namedrange[{name}]",
                    Message = $"Defined name '{name}' references missing sheet '{missingSheet}'",
                    Context = body,
                    Suggestion = "Restore the sheet or rebind the name to an existing range."
                });
            }
        }

        // Chart series references pointing at sheets / cells that no longer
        // exist — same observability family as formula_not_evaluated.
        // Excel won't refuse to load a chart whose series formula references
        // a deleted sheet; it just renders the last cached values and the
        // ref becomes a silent landmine for the next refresh. Detect by
        // scanning every chart's c:f formulas and matching the sheet prefix
        // against the live workbook.
        foreach (var (slug, formula) in EnumerateChartRefFormulas())
        {
            if (limit.HasValue && issues.Count >= limit.Value) break;
            if (!ChartRefSheetExists(formula, out var missingSheet)) continue;
            issues.Add(new DocumentIssue
            {
                Id = $"R{++issueNum}",
                Type = IssueType.Content,
                Subtype = "chart_series_ref_missing_sheet",
                Severity = IssueSeverity.Error,
                Path = slug,
                Message = $"Chart series references missing sheet '{missingSheet}'",
                Context = formula,
                Suggestion = "Restore the sheet, or rebuild the chart against an existing range."
            });
        }

        // Chart numCache vs live cell values — stale-cache detection.
        // Off by default because it walks every chart × every series × every
        // point and evaluates every referenced range; opt in via
        // `--type chart_cache_stale`. When a cell's value changed but the
        // chart hasn't been refreshed by Excel, the chart silently shows
        // old values. The check requires explicit opt-in because some
        // numCache deltas are legitimate (rounding, formatting), and we
        // don't want noise on every casual `view issues`.
        if (issueType == "chart_cache_stale")
        {
            foreach (var (slug, numRef) in EnumerateChartNumberRefs())
            {
                if (limit.HasValue && issues.Count >= limit.Value) break;
                var formula = numRef.Formula?.Text;
                if (string.IsNullOrWhiteSpace(formula)) continue;
                if (ChartRefSheetExists(formula, out _)) continue; // skip; #2 already reports
                var cached = numRef.NumberingCache?.Elements<C.NumericPoint>()
                    .Select(p => p.NumericValue?.Text ?? "").ToList();
                if (cached == null || cached.Count == 0) continue;
                // First try the cheap range-only resolver (preserves cell-format
                // text exactly). If the formula is wrapped in functions like
                // SUM/AVERAGE/INDEX/OFFSET it returns null — fall back to the
                // FormulaEvaluator, which produces a single scalar that we
                // compare against the cached scalar (collapses N points to 1
                // for aggregate functions; that's the right answer).
                var live = ResolveChartFormulaValues(formula);
                if (live == null) live = TryEvaluateChartFormulaScalar(formula);
                if (live == null) continue;
                // Compare cached vs live string-wise — both come from cell text;
                // numeric formatting normalisation would mask real edits.
                if (cached.Count != live.Count
                    || cached.Zip(live, (c, l) => c != l).Any(b => b))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Subtype = "chart_cache_stale",
                        Severity = IssueSeverity.Warning,
                        Path = slug,
                        Message = "Chart numCache out of sync with source cells",
                        Context = $"f=\"{formula}\" cached=[{string.Join(",", cached)}] live=[{string.Join(",", live)}]",
                        Suggestion = "Open in Excel to refresh, or call validate after data changes."
                    });
                }
            }
        }

        // CONSISTENCY(text-overflow-check): merged in from former `check` command.
        // Emits wrapText-cells whose visible row-height budget can't fit the wrapped text.
        foreach (var (path, msg) in CheckAllCellOverflow())
        {
            if (limit.HasValue && issues.Count >= limit.Value) break;
            issues.Add(new DocumentIssue
            {
                Id = $"O{++issueNum}",
                Type = IssueType.Format,
                Severity = IssueSeverity.Warning,
                Path = path,
                Message = msg
            });
        }

        return issues;
    }

    // ==================== Chart reference observability ====================

    // Cached worksheet list — ViewAsIssues + EnumerateChartRefFormulas +
    // EnumerateChartNumberRefs + ChartRefSheetExists + ResolveChartFormulaValues
    // all call GetWorksheets() repeatedly during a single scan. The underlying
    // GetPartById walk isn't free on workbooks with many sheets; computing it
    // once per scan keeps complexity O(sheets + charts) instead of O(sheets ×
    // charts). The cache is scoped to one ViewAsIssues invocation; subsequent
    // calls rebuild it (sheet rename/add via Set invalidates).
    private List<(string Name, WorksheetPart Part)>? _viewAsIssuesWorksheetCache;
    private HashSet<string>? _viewAsIssuesSheetNameCache;
    private List<(string Name, WorksheetPart Part)> CachedWorksheets()
        => _viewAsIssuesWorksheetCache ??= GetWorksheets();
    private HashSet<string> CachedSheetNames()
        => _viewAsIssuesSheetNameCache ??= new HashSet<string>(
            CachedWorksheets().Select(w => w.Name), StringComparer.OrdinalIgnoreCase);

    /// <summary>Yield every &lt;c:numRef&gt; with its slug, for chart-cache-stale
    /// cross-checks against live cell values.</summary>
    private IEnumerable<(string Slug, C.NumberReference NumRef)> EnumerateChartNumberRefs()
    {
        foreach (var (sheetName, wsPart) in CachedWorksheets())
        {
            if (wsPart.DrawingsPart is not { } dp) continue;
            int idx = 0;
            foreach (var cp in dp.ChartParts)
            {
                idx++;
                if (cp.ChartSpace is null) continue;
                foreach (var nr in cp.ChartSpace.Descendants<C.NumberReference>())
                    yield return ($"/{sheetName}/chart[{idx}]", nr);
            }
        }
    }

    /// <summary>Fall back for chart formulas that wrap a range in functions
    /// (SUM/AVERAGE/INDEX/OFFSET/…). Pipe through FormulaEvaluator and return
    /// the scalar as a single-element list so the cached/live comparator can
    /// run uniformly. Returns null when evaluator can't handle the formula
    /// (then the scanner silently skips — accepted: agent can opt-in to
    /// formula_not_evaluated for the underlying issue).</summary>
    private List<string>? TryEvaluateChartFormulaScalar(string formula)
    {
        // Use the first worksheet's evaluator — chart c:f formulas can reference
        // any sheet by name. The evaluator follows Sheet!Ref prefixes itself.
        var first = CachedWorksheets().FirstOrDefault();
        if (first.Part == null) return null;
        var sheetData = GetSheet(first.Part).GetFirstChild<SheetData>();
        if (sheetData == null) return null;
        var ev = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        var report = ev.EvaluateForReport(formula);
        if (report.Status != Core.EvalReportStatus.Evaluated) return null;
        return new List<string> { report.Result!.ToCellValueText() };
    }

    /// <summary>Resolve a chart c:f formula like "Sheet1!$A$1:$A$3" against
    /// current cell values; returns the cell text in row-major order, or null
    /// if the sheet is missing (caller already reports that via #2).</summary>
    private List<string>? ResolveChartFormulaValues(string formula)
    {
        var bang = formula.IndexOf('!');
        if (bang <= 0) return null;
        var sheetPart = formula[..bang].Trim();
        if (sheetPart.StartsWith('\'') && sheetPart.EndsWith('\''))
            sheetPart = sheetPart[1..^1].Replace("''", "'");
        var rangePart = formula[(bang + 1)..].Replace("$", "");
        var wsPart = CachedWorksheets()
            .FirstOrDefault(w => string.Equals(w.Name, sheetPart, StringComparison.OrdinalIgnoreCase))
            .Part;
        if (wsPart == null) return null;
        var sheetData = GetSheet(wsPart).GetFirstChild<SheetData>();
        if (sheetData == null) return null;

        // Cell or range A1 / A1:B3 — fall back to null on anything else
        // (named ranges, table refs); not in scope for this scanner.
        var parts = rangePart.Split(':');
        var first = parts[0];
        var last = parts.Length > 1 ? parts[1] : parts[0];
        if (!System.Text.RegularExpressions.Regex.IsMatch(first, "^[A-Z]+\\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return null;
        if (!System.Text.RegularExpressions.Regex.IsMatch(last, "^[A-Z]+\\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return null;

        var (col1Str, r1) = ParseCellReference(first.ToUpperInvariant());
        var (col2Str, r2) = ParseCellReference(last.ToUpperInvariant());
        int c1 = ColumnNameToIndex(col1Str), c2 = ColumnNameToIndex(col2Str);
        var cellIndex = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheetData.Elements<Row>())
            foreach (var cell in row.Elements<Cell>())
                if (cell.CellReference?.Value is { } cr) cellIndex[cr] = cell;

        var values = new List<string>();
        for (int r = r1; r <= r2; r++)
            for (int c = c1; c <= c2; c++)
            {
                var addr = $"{IndexToColumnName(c)}{r}";
                if (!cellIndex.TryGetValue(addr, out var cell)) { values.Add(""); continue; }
                values.Add(cell.CellValue?.Text ?? "");
            }
        return values;
    }

    /// <summary>
    /// Yield every &lt;c:f&gt; formula text across all chart parts (standard and
    /// extended) attached to any worksheet. The slug identifies the chart for
    /// the issue Path — sheet name plus chart index, matching what ExcelHandler
    /// already emits for chart-level Set/Get paths.
    /// </summary>
    private IEnumerable<(string Slug, string Formula)> EnumerateChartRefFormulas()
    {
        foreach (var (sheetName, wsPart) in CachedWorksheets())
        {
            if (wsPart.DrawingsPart is not { } dp) continue;
            int idx = 0;
            foreach (var cp in dp.ChartParts)
            {
                idx++;
                if (cp.ChartSpace is null) continue;
                foreach (var f in cp.ChartSpace.Descendants<C.Formula>())
                {
                    var t = f.Text;
                    if (!string.IsNullOrWhiteSpace(t))
                        yield return ($"/{sheetName}/chart[{idx}]", t);
                }
            }
            foreach (var ep in dp.ExtendedChartParts)
            {
                idx++;
                if (ep.ChartSpace is null) continue;
                // Extended charts use the same c:f shape via cx:formula-equivalent
                // descendants — defensive descendant scan picks them up.
                foreach (var e in ep.ChartSpace.Descendants())
                {
                    if (e.LocalName == "f" && !string.IsNullOrWhiteSpace(e.InnerText))
                        yield return ($"/{sheetName}/chart[{idx}]", e.InnerText);
                }
            }
        }
    }

    /// <summary>
    /// Parse a chart c:f formula like "Sheet1!$A$1:$B$5" or "'Quoted Sheet'!A1"
    /// and return true if the sheet prefix names a sheet that no longer exists
    /// in the workbook. Out-parameter receives the missing sheet name. Returns
    /// false (no issue) when the formula has no sheet prefix or the sheet
    /// resolves cleanly. Range / cell validity itself is intentionally not
    /// checked here — that's the chart-cache-stale gap (#5).
    /// </summary>
    private bool ChartRefSheetExists(string formula, out string missingSheet)
    {
        missingSheet = "";
        var bang = formula.IndexOf('!');
        if (bang <= 0) return false;
        var sheetPart = formula[..bang].Trim();
        // Quoted sheet names: 'Sheet Name'; '' escapes a literal apostrophe.
        if (sheetPart.StartsWith('\'') && sheetPart.EndsWith('\''))
            sheetPart = sheetPart[1..^1].Replace("''", "'");
        // Multi-sheet 3D refs (Sheet1:Sheet3) — split at colon. If either end
        // missing, report the first one to keep messaging concrete.
        var colon = sheetPart.IndexOf(':');
        var first = colon >= 0 ? sheetPart[..colon] : sheetPart;
        var second = colon >= 0 ? sheetPart[(colon + 1)..] : null;
        var liveSheets = CachedSheetNames();
        if (!liveSheets.Contains(first)) { missingSheet = first; return true; }
        if (second != null && !liveSheets.Contains(second)) { missingSheet = second; return true; }
        return false;
    }
}
