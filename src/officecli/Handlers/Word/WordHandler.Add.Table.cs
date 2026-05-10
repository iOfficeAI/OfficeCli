// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddTable(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var table = new Table();
        // BUG-R7-03: Previously this always seeded all 6 borders (top/bottom/
        // left/right/insideH/insideV) and then applied user props on top —
        // which corrupted three-line tables (top+bottom only) on round-trip
        // because dump emits only the user-set sides. When the caller passes
        // ANY border.* prop, treat it as an explicit specification: start
        // with an empty TableBorders and apply only the requested sides.
        // Otherwise (no border props at all), keep the historical default
        // grid look so bare `add table` still produces a visible table.
        var hasExplicitBorders = properties.Keys.Any(k =>
            k.StartsWith("border", StringComparison.OrdinalIgnoreCase));
        TableProperties tblProps;
        if (hasExplicitBorders)
        {
            tblProps = new TableProperties(new TableBorders());
            table.AppendChild(tblProps);
            foreach (var (bk, bv) in properties)
            {
                if (bk.StartsWith("border", StringComparison.OrdinalIgnoreCase))
                    ApplyTableBorders(tblProps, bk, bv);
            }
        }
        else
        {
            tblProps = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                )
            );
            table.AppendChild(tblProps);
        }

        // Parse data if provided: "H1,H2;R1C1,R1C2;R2C1,R2C2" or CSV file/URL/data-URI
        string[][]? tableData = null;
        if (properties.TryGetValue("data", out var dataStr))
        {
            if (OfficeCli.Core.FileSource.IsResolvable(dataStr))
                tableData = OfficeCli.Core.FileSource.ResolveLines(dataStr)
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(l => l.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
            else
                tableData = dataStr.Split(';')
                    .Select(r => r.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
        }

        int rows, cols;
        if (tableData != null)
        {
            rows = tableData.Length;
            cols = tableData.Max(r => r.Length);
        }
        else
        {
            rows = 1;
            if (properties.TryGetValue("rows", out var rowsStr))
            {
                if (!int.TryParse(rowsStr, out rows))
                    throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Expected a positive integer.");
                if (rows <= 0)
                    throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Must be a positive integer (> 0).");
            }
            cols = 1;
            if (properties.TryGetValue("cols", out var colsStr))
            {
                cols = ParseHelpers.SafeParseInt(colsStr, "cols");
                if (cols <= 0)
                    throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Must be a positive integer (> 0).");
            }
        }

        // Parse per-column widths: colWidths="3000,2000,5000"
        int[]? colWidthArr = null;
        if (properties.TryGetValue("colwidths", out var cwStr) || properties.TryGetValue("colWidths", out cwStr))
        {
            var parts = cwStr.Split(',');
            colWidthArr = new int[parts.Length];
            for (int ci = 0; ci < parts.Length; ci++)
            {
                if (!int.TryParse(parts[ci].Trim(), out colWidthArr[ci]))
                    throw new ArgumentException($"Invalid 'colwidths' value: '{parts[ci].Trim()}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
                // BUG-R1-01: reject negative or zero up front (Set already
                // does this; Add did not). Invalid OOXML otherwise.
                if (colWidthArr[ci] <= 0)
                    throw new ArgumentException($"Invalid 'colwidths' value: '{parts[ci].Trim()}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
            }
        }

        // Add table grid
        // BUG-R1-P0-4: when colWidths is not specified, default per-column
        // width should be computed from the section's usable body width
        // (page width − left/right margins) divided by `cols`. The previous
        // hard-coded 2400-twips default overflowed the page once cols > 3
        // on default A4 / Letter section properties.
        long defaultColTwips = 2400;
        if (colWidthArr == null)
        {
            var sectPr = _doc.MainDocumentPart?.Document?.Body?
                .Descendants<SectionProperties>().LastOrDefault();
            var pgSz = sectPr?.GetFirstChild<PageSize>();
            var pgMar = sectPr?.GetFirstChild<PageMargin>();
            long pageW = pgSz?.Width?.Value ?? 12240u;
            long mL = pgMar?.Left?.Value ?? 1440u;
            long mR = pgMar?.Right?.Value ?? 1440u;
            long usable = Math.Max(1, pageW - mL - mR);
            defaultColTwips = Math.Max(1, usable / Math.Max(1, cols));
        }

        var tblGrid = new TableGrid();
        for (int gc = 0; gc < cols; gc++)
        {
            // BUG-R1-01: reject negative or zero gridCol widths up front
            // (Set already does this; Add did not). Invalid OOXML otherwise.
            if (colWidthArr != null && gc < colWidthArr.Length)
            {
                if (colWidthArr[gc] <= 0)
                    throw new ArgumentException($"Invalid 'colwidths' value: '{colWidthArr[gc]}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
            }
            var w = colWidthArr != null && gc < colWidthArr.Length
                ? colWidthArr[gc].ToString()
                : defaultColTwips.ToString();
            tblGrid.AppendChild(new GridColumn { Width = w });
        }
        table.AppendChild(tblGrid);

        // Apply table-level properties from Add parameters
        foreach (var (tk, tv) in properties)
        {
            var tkl = tk.ToLowerInvariant();
            if (tkl is "rows" or "cols" or "colwidths" || tkl.StartsWith("border")) continue;
            switch (tkl)
            {
                case "align" or "alignment":
                    tblProps.TableJustification = new TableJustification
                    {
                        Val = tv.ToLowerInvariant() switch
                        {
                            "center" => TableRowAlignmentValues.Center,
                            "right" => TableRowAlignmentValues.Right,
                            "left" => TableRowAlignmentValues.Left,
                            _ => throw new ArgumentException($"Invalid table alignment value: '{tv}'. Valid values: left, center, right.")
                        }
                    };
                    break;
                case "width":
                    // BUG-DUMP19-03: accept "auto" so dump round-trip preserves
                    // <w:tblW w:type="auto"/>. Without this, SafeParseUint("auto")
                    // throws and the prop is silently dropped/normalized.
                    if (string.Equals(tv, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        tblProps.TableWidth = new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto };
                    }
                    else if (tv.EndsWith('%'))
                    {
                        var pct = ParseHelpers.SafeParseInt(tv.TrimEnd('%'), "width") * 50;
                        tblProps.TableWidth = new TableWidth { Width = pct.ToString(), Type = TableWidthUnitValues.Pct };
                    }
                    else
                    {
                        tblProps.TableWidth = new TableWidth { Width = ParseHelpers.SafeParseUint(tv, "width").ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "indent":
                    tblProps.TableIndentation = new TableIndentation { Width = ParseHelpers.SafeParseInt(tv, "indent"), Type = TableWidthUnitValues.Dxa };
                    break;
                case "cellspacing":
                    tblProps.TableCellSpacing = new TableCellSpacing { Width = ParseHelpers.SafeParseUint(tv, "cellspacing").ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                case "layout":
                    tblProps.TableLayout = new TableLayout
                    {
                        Type = tv.ToLowerInvariant() == "fixed" ? TableLayoutValues.Fixed : TableLayoutValues.Autofit
                    };
                    break;
                case "padding":
                    var cm = tblProps.TableCellMarginDefault ?? tblProps.AppendChild(new TableCellMarginDefault());
                    var paddingVal = ParseHelpers.SafeParseInt(tv, "padding");
                    cm.TopMargin = new TopMargin { Width = tv, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    cm.BottomMargin = new BottomMargin { Width = tv, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    break;
                // BUG-DUMP13-04: per-side default cell margins. BatchEmitter
                // passes asymmetric padding.* keys through unfolded when sides
                // differ; without these cases AddTable warned UNSUPPORTED and
                // the values became zero on round-trip. Mirrors the per-cell
                // tcMar handling in Set.Element.cs.
                case "padding.top":
                    {
                        var cmt = tblProps.TableCellMarginDefault ?? tblProps.AppendChild(new TableCellMarginDefault());
                        cmt.TopMargin = new TopMargin { Width = tv, Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "padding.bottom":
                    {
                        var cmb = tblProps.TableCellMarginDefault ?? tblProps.AppendChild(new TableCellMarginDefault());
                        cmb.BottomMargin = new BottomMargin { Width = tv, Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "padding.left":
                    {
                        var cml = tblProps.TableCellMarginDefault ?? tblProps.AppendChild(new TableCellMarginDefault());
                        var lv = ParseHelpers.SafeParseInt(tv, "padding.left");
                        cml.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(lv, short.MaxValue), Type = TableWidthValues.Dxa };
                    }
                    break;
                case "padding.right":
                    {
                        var cmr = tblProps.TableCellMarginDefault ?? tblProps.AppendChild(new TableCellMarginDefault());
                        var rv = ParseHelpers.SafeParseInt(tv, "padding.right");
                        cmr.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(rv, short.MaxValue), Type = TableWidthValues.Dxa };
                    }
                    break;
                case "style":
                    tblProps.TableStyle = new TableStyle { Val = tv };
                    // Add TableLook so built-in styles apply banding correctly
                    tblProps.RemoveAllChildren<TableLook>();
                    tblProps.AppendChild(new TableLook { Val = "04A0" });
                    break;
                case "shd" or "shading":
                    {
                        // BUG-DUMP21-01: w:tblPr/w:shd table-level shading
                        // round-trip. Mirrors paragraph/cell `shading` parsing
                        // — accepts FILL, VAL;FILL, or VAL;FILL;COLOR.
                        var shdParts = tv.Split(';');
                        var tShd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            tShd.Val = ShadingPatternValues.Clear;
                            tShd.Fill = SanitizeHex(shdParts[0]);
                        }
                        else if (shdParts.Length >= 2)
                        {
                            var pat = shdParts[0].TrimStart('#');
                            if (pat.Length >= 6 && pat.All(char.IsAsciiHexDigit))
                            {
                                tShd.Val = ShadingPatternValues.Clear;
                                tShd.Fill = SanitizeHex(shdParts[0]);
                            }
                            else
                            {
                                tShd.Val = new ShadingPatternValues(shdParts[0]);
                                tShd.Fill = SanitizeHex(shdParts[1]);
                                if (shdParts.Length >= 3)
                                    tShd.Color = SanitizeHex(shdParts[2]);
                            }
                        }
                        tblProps.Shading = tShd;
                    }
                    break;
                case "direction" or "dir" or "bidi":
                    // Table-level bidi: emit <w:bidiVisual/> on tblPr in schema
                    // order. Mirrors paragraph/cell direction=rtl vocabulary.
                    // CONSISTENCY(rtl-cascade).
                    tblProps.RemoveAllChildren<BiDiVisual>();
                    if (ParseDirectionRtl(tv))
                        InsertTblPrChildInOrder(tblProps, new BiDiVisual());
                    break;
            }
        }

        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
            {
                var cellText = tableData != null && r < tableData.Length && c < tableData[r].Length
                    ? tableData[r][c] : (properties.TryGetValue($"r{r + 1}c{c + 1}", out var rc) ? rc : "");
                // CONSISTENCY(table-cell-defaults): do not stamp explicit
                // spaceAfter=0 / lineSpacing=240 Auto on freshly-created cell
                // paragraphs — let them inherit from style/docDefaults like
                // regular body paragraphs. Otherwise dump→batch round-trip
                // grows 67 extra `set spaceAfter=0pt lineSpacing=1x` commands
                // per cell (BUG-R3-3).
                var cellPara = new Paragraph();
                AssignParaId(cellPara);
                if (!string.IsNullOrEmpty(cellText))
                    cellPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
                var cell = new TableCell(cellPara);
                // BUG-R6-06 / BUG-R6-01: do NOT stamp an explicit
                // <w:tcW> on every cell when the user supplied colWidths
                // — w:tblGrid/w:gridCol already encodes the column
                // widths, and per-cell tcW makes dump→batch→dump
                // non-idempotent (each round-trip emits N×M extra
                // `set width=…` commands). Cells without a tcW inherit
                // the column width from tblGrid as the schema intends.
                row.AppendChild(cell);
            }
            table.AppendChild(row);
        }

        // Dotted-key fallback for tblPr-level attrs not modeled by the
        // hand-rolled blocks above (single-attr forms like tblpPr.* or
        // future schema additions). CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            // border.{top,bottom,left,right,insideH,insideV,all} were already
            // applied at the top of AddTable via ApplyTableBorders. Skip them
            // here so they don't get mis-flagged UNSUPPORTED by the generic
            // TypedAttributeFallback (which doesn't model border.*).
            // CONSISTENCY(add-set-symmetry).
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase)) continue;
            // BUG-DUMP14-04: padding.{top,bottom,left,right} are handled by
            // the main switch above (round-13 added tblCellMar emit). Skip
            // them here so they aren't double-tagged as UNSUPPORTED by the
            // generic TypedAttributeFallback. Mirrors border.* skip.
            if (key.StartsWith("padding.", StringComparison.OrdinalIgnoreCase)) continue;
            if (Core.TypedAttributeFallback.TrySet(tblProps, key, value)) continue;
            LastAddUnsupportedProps.Add(key);
        }

        if (index.HasValue)
            InsertAtPosition(parent, table, index);
        else
            AppendToParent(parent, table);
        var tbls = parent.Elements<Table>().ToList();
        var idx = tbls.FindIndex(t => ReferenceEquals(t, table));
        return $"{parentPath}/tbl[{(idx >= 0 ? idx + 1 : tbls.Count)}]";
    }

    private string AddRow(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not Table targetTable)
            throw new ArgumentException("Rows can only be added to a table: /body/tbl[N]");

        var grid = targetTable.GetFirstChild<TableGrid>()
            ?? targetTable.PrependChild(new TableGrid());
        var existingGridCols = grid.Elements<GridColumn>().ToList();
        var existingCols = existingGridCols.Count > 0 ? existingGridCols.Count : 1;
        int newCols = existingCols;
        if (properties.TryGetValue("cols", out var colsVal))
        {
            newCols = ParseHelpers.SafeParseInt(colsVal, "cols");
            // BUG-R1-P0-3a: cols=0 silently produces an empty <w:tr> with no
            // cells; per OOXML spec a row must contain at least one cell.
            if (newCols <= 0)
                throw new ArgumentException($"Invalid 'cols' value: '{colsVal}'. Must be a positive integer (> 0); a row with 0 cells is invalid OOXML.");
        }

        // BUG-R1-P0-3b: cols > existing tblGrid count must expand tblGrid
        // to keep tcW / gridCol in agreement. Otherwise the extra cells
        // have no column-width definition and Word misaligns them.
        if (existingGridCols.Count > 0 && newCols > existingGridCols.Count)
        {
            // Width: average of existing cols, falling back to 2400.
            long avg = (long)existingGridCols.Average(gc =>
                long.TryParse(gc.Width?.Value, out var w) ? w : 2400L);
            for (int extra = existingGridCols.Count; extra < newCols; extra++)
                grid.AppendChild(new GridColumn { Width = avg.ToString() });
        }

        var newRow = new TableRow();
        TableRowProperties? newRowProps = null;
        if (properties.TryGetValue("height", out var rowHeight))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeight), HeightType = HeightRuleValues.AtLeast });
        }
        if (properties.TryGetValue("height.exact", out var rowHeightExact))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.GetFirstChild<TableRowHeight>()?.Remove();
            newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeightExact), HeightType = HeightRuleValues.Exact });
        }
        if (properties.TryGetValue("header", out var headerVal) && IsTruthy(headerVal))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.AppendChild(new TableHeader());
        }

        for (int c = 0; c < newCols; c++)
        {
            var cellText = properties.TryGetValue($"c{c + 1}", out var ct) ? ct : "";
            var cellPara = new Paragraph();
            AssignParaId(cellPara);
            if (!string.IsNullOrEmpty(cellText))
                cellPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
            newRow.AppendChild(new TableCell(cellPara));
        }

        // Dotted-key fallback for trPr-level attrs (trHeight.*, etc.) not
        // modeled by hand-rolled blocks. Lazy-create trPr if any dotted
        // attr binds. CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            var trPrTarget = newRowProps ?? new TableRowProperties();
            if (Core.TypedAttributeFallback.TrySet(trPrTarget, key, value))
            {
                if (newRowProps == null)
                {
                    newRow.PrependChild(trPrTarget);
                    newRowProps = trPrTarget;
                }
                continue;
            }
            LastAddUnsupportedProps.Add(key);
        }

        if (index.HasValue)
        {
            var existingRows = targetTable.Elements<TableRow>().ToList();
            if (index.Value < existingRows.Count)
                targetTable.InsertBefore(newRow, existingRows[index.Value]);
            else
                targetTable.AppendChild(newRow);
        }
        else
        {
            targetTable.AppendChild(newRow);
        }

        var rowIdx = targetTable.Elements<TableRow>().ToList().IndexOf(newRow) + 1;
        return $"{parentPath}/tr[{rowIdx}]";
    }

    /// <summary>
    /// Insert a new virtual column into a Word table. OOXML has no <w:col>
    /// element, so this synthesizes one by inserting a <w:gridCol> in
    /// <w:tblGrid> and a fresh <w:tc> at the same positional index in every
    /// existing <w:tr>. Rejects when any affected row carries gridSpan or
    /// vMerge in that column slot — those merge directives reference column
    /// positions and would silently break.
    /// </summary>
    private string AddTableColumn(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not Table targetTable)
            throw new ArgumentException("Columns can only be added to a table: /body/tbl[N]");

        var grid = targetTable.GetFirstChild<TableGrid>()
            ?? targetTable.PrependChild(new TableGrid());
        var existingGridCols = grid.Elements<GridColumn>().ToList();
        var insertIdx = index.HasValue && index.Value >= 0 && index.Value < existingGridCols.Count
            ? index.Value
            : existingGridCols.Count; // append by default

        // Reject if any row at insertIdx straddles the boundary via merge.
        foreach (var row in targetTable.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            // Check the cell currently occupying slot `insertIdx` (the one
            // that will be pushed right). gridSpan or vMerge here means
            // re-indexing the column slot would split a merged region.
            if (insertIdx < cells.Count && CellHasMerge(cells[insertIdx]))
                throw new ArgumentException(
                    $"Cannot insert column at index {insertIdx} of {parentPath}: " +
                    $"a row contains a merged cell straddling that boundary (gridSpan/vMerge). " +
                    "Unmerge first or pick a different position.");
        }

        // Width: explicit, or average of existing cols, or default 2400 twips
        long defaultWidthTwips = 2400;
        long newWidth = properties.TryGetValue("width", out var wVal)
            ? ParseTwips(wVal)
            : (existingGridCols.Count > 0
                ? (long)existingGridCols.Average(gc => long.TryParse(gc.Width?.Value, out var w) ? w : defaultWidthTwips)
                : defaultWidthTwips);

        var newGridCol = new GridColumn { Width = newWidth.ToString() };
        if (insertIdx < existingGridCols.Count)
            grid.InsertBefore(newGridCol, existingGridCols[insertIdx]);
        else
            grid.AppendChild(newGridCol);

        var cellText = properties.GetValueOrDefault("text", "");
        foreach (var row in targetTable.Elements<TableRow>())
        {
            var newPara = new Paragraph();
            AssignParaId(newPara);
            if (!string.IsNullOrEmpty(cellText))
                newPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
            var newCell = new TableCell(newPara);

            var cells = row.Elements<TableCell>().ToList();
            if (insertIdx < cells.Count)
                row.InsertBefore(newCell, cells[insertIdx]);
            else
                row.AppendChild(newCell);
        }

        var newColIdx = grid.Elements<GridColumn>().ToList().IndexOf(newGridCol) + 1;
        return $"{parentPath}/col[{newColIdx}]";
    }

    /// <summary>
    /// True if the cell carries gridSpan > 1 (horizontal merge) or any
    /// vMerge directive (vertical merge — restart or continue).
    /// </summary>
    private static bool CellHasMerge(TableCell cell)
    {
        var tcPr = cell.GetFirstChild<TableCellProperties>();
        if (tcPr == null) return false;
        var span = tcPr.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
        if (span > 1) return true;
        if (tcPr.GetFirstChild<VerticalMerge>() != null) return true;
        return false;
    }

    private string AddCell(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not TableRow targetRow)
            throw new ArgumentException("Cells can only be added to a table row: /body/tbl[N]/tr[M]");

        // BUG-R1-P0-2: AddCell on an existing row must keep tblGrid in sync.
        // Without this, the new cell has no matching <w:gridCol> and the
        // last "virtual column" collapses in Word. We synchronize lazily:
        // if the row's total grid-column occupancy after appending exceeds
        // the existing tblGrid, append matching gridCol entries averaging
        // the existing widths. Mirrors AddTableColumn's width logic.
        Table? cellParentTable = targetRow.Parent as Table;
        TableGrid? cellGrid = cellParentTable?.GetFirstChild<TableGrid>();

        var cellParagraph = new Paragraph();
        AssignParaId(cellParagraph);
        if (properties.TryGetValue("text", out var cellTxt))
            cellParagraph.AppendChild(new Run(new Text(cellTxt) { Space = SpaceProcessingModeValues.Preserve }));

        // Reading direction (Arabic / Hebrew). Mirrors AddParagraph: 'rtl'
        // writes <w:bidi/> on the cell paragraph's pPr and stamps <w:rtl/>
        // on the paragraph mark + any text run that was just appended.
        // CONSISTENCY(rtl-cascade).
        if (properties.TryGetValue("direction", out var cellDirRaw)
            || properties.TryGetValue("dir", out cellDirRaw)
            || properties.TryGetValue("bidi", out cellDirRaw))
        {
            bool cellRtl = ParseDirectionRtl(cellDirRaw);
            var cellPProps = cellParagraph.ParagraphProperties ?? cellParagraph.PrependChild(new ParagraphProperties());
            if (cellRtl) cellPProps.BiDi = new BiDi();
            var cellMarkRPr = cellPProps.ParagraphMarkRunProperties ?? cellPProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(cellMarkRPr, "direction", cellRtl ? "rtl" : "ltr");
            foreach (var existingRun in cellParagraph.Descendants<Run>())
                ApplyRunFormatting(EnsureRunProperties(existingRun), "direction", cellRtl ? "rtl" : "ltr");
        }

        var newCell = new TableCell(cellParagraph);

        if (properties.TryGetValue("width", out var cellWidth))
        {
            // BUG-DUMP6-04: accept "N%" alongside bare twips so dump→batch
            // round-trips pct cell widths. OOXML stores pct as fifths-of-percent.
            TableCellWidth tcw;
            if (cellWidth.EndsWith('%') &&
                double.TryParse(cellWidth.AsSpan(0, cellWidth.Length - 1),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var pctCw))
            {
                tcw = new TableCellWidth
                {
                    Width = ((int)Math.Round(pctCw * 50)).ToString(),
                    Type = TableWidthUnitValues.Pct
                };
            }
            else
            {
                tcw = new TableCellWidth { Width = cellWidth, Type = TableWidthUnitValues.Dxa };
            }
            newCell.PrependChild(new TableCellProperties(tcw));
        }

        // Dotted-key fallback for tcPr-level attrs (shd.fill, etc.) not
        // modeled by hand-rolled blocks. Lazy-create tcPr if any dotted
        // attr binds. CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            var tcPr = newCell.GetFirstChild<TableCellProperties>();
            var lazyTcPr = tcPr ?? new TableCellProperties();
            // CONSISTENCY(add-set-symmetry): route border.{top,bottom,left,
            // right,all,tl2br,tr2bl} through the same ApplyCellBorders helper
            // Set uses, instead of falling through to TypedAttributeFallback
            // which doesn't model border.* and would mis-flag UNSUPPORTED.
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase)
                || key.Equals("border", StringComparison.OrdinalIgnoreCase))
            {
                ApplyCellBorders(lazyTcPr, key, value);
                if (tcPr == null) newCell.PrependChild(lazyTcPr);
                continue;
            }
            if (Core.TypedAttributeFallback.TrySet(lazyTcPr, key, value))
            {
                if (tcPr == null) newCell.PrependChild(lazyTcPr);
                continue;
            }
            LastAddUnsupportedProps.Add(key);
        }

        if (index.HasValue)
        {
            var cells = targetRow.Elements<TableCell>().ToList();
            if (index.Value < cells.Count)
                targetRow.InsertBefore(newCell, cells[index.Value]);
            else
                targetRow.AppendChild(newCell);
        }
        else
        {
            targetRow.AppendChild(newCell);
        }

        // BUG-R1-P0-2: expand tblGrid if this row's grid-column occupancy
        // (sum of gridSpan) now exceeds existing gridCol count.
        if (cellGrid != null)
        {
            var existingGridCount = cellGrid.Elements<GridColumn>().Count();
            var rowSpan = targetRow.Elements<TableCell>().Sum(tc =>
                tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            if (rowSpan > existingGridCount)
            {
                long avgWidth;
                var existingWidths = cellGrid.Elements<GridColumn>().ToList();
                avgWidth = existingWidths.Count > 0
                    ? (long)existingWidths.Average(gc => long.TryParse(gc.Width?.Value, out var w) ? w : 2400L)
                    : 2400L;
                for (int i = existingGridCount; i < rowSpan; i++)
                    cellGrid.AppendChild(new GridColumn { Width = avgWidth.ToString() });
            }
        }

        var cellIdx = targetRow.Elements<TableCell>().ToList().IndexOf(newCell) + 1;
        return $"{parentPath}/tc[{cellIdx}]";
    }
}
