// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        // Document-level properties
        if (path == "/" || path == "")
        {
            SetDocumentProperties(properties);
            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        var parts = ParsePath(path);
        var element = NavigateToElement(parts);
        if (element == null)
            throw new ArgumentException($"Path not found: {path}");

        if (element is Run run)
        {
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "text":
                        var textEl = run.GetFirstChild<Text>();
                        if (textEl != null) textEl.Text = value;
                        break;
                    case "bold":
                        EnsureRunProperties(run).Bold = bool.Parse(value) ? new Bold() : null;
                        break;
                    case "italic":
                        EnsureRunProperties(run).Italic = bool.Parse(value) ? new Italic() : null;
                        break;
                    case "caps":
                        EnsureRunProperties(run).Caps = bool.Parse(value) ? new Caps() : null;
                        break;
                    case "smallcaps":
                        EnsureRunProperties(run).SmallCaps = bool.Parse(value) ? new SmallCaps() : null;
                        break;
                    case "dstrike":
                        EnsureRunProperties(run).DoubleStrike = bool.Parse(value) ? new DoubleStrike() : null;
                        break;
                    case "vanish":
                        EnsureRunProperties(run).Vanish = bool.Parse(value) ? new Vanish() : null;
                        break;
                    case "outline":
                        EnsureRunProperties(run).Outline = bool.Parse(value) ? new Outline() : null;
                        break;
                    case "shadow":
                        EnsureRunProperties(run).Shadow = bool.Parse(value) ? new Shadow() : null;
                        break;
                    case "emboss":
                        EnsureRunProperties(run).Emboss = bool.Parse(value) ? new Emboss() : null;
                        break;
                    case "imprint":
                        EnsureRunProperties(run).Imprint = bool.Parse(value) ? new Imprint() : null;
                        break;
                    case "noproof":
                        EnsureRunProperties(run).NoProof = bool.Parse(value) ? new NoProof() : null;
                        break;
                    case "rtl":
                        EnsureRunProperties(run).RightToLeftText = bool.Parse(value) ? new RightToLeftText() : null;
                        break;
                    case "font":
                        var rPrFont = EnsureRunProperties(run);
                        var existingFonts = rPrFont.RunFonts;
                        if (existingFonts != null)
                        {
                            existingFonts.Ascii = value;
                            existingFonts.HighAnsi = value;
                            existingFonts.EastAsia = value;
                        }
                        else
                        {
                            rPrFont.RunFonts = new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value };
                        }
                        break;
                    case "size":
                        EnsureRunProperties(run).FontSize = new FontSize
                        {
                            Val = (int.Parse(value) * 2).ToString() // half-points
                        };
                        break;
                    case "highlight":
                        EnsureRunProperties(run).Highlight = new Highlight
                        {
                            Val = new HighlightColorValues(value)
                        };
                        break;
                    case "color":
                        EnsureRunProperties(run).Color = new Color { Val = value.ToUpperInvariant() };
                        break;
                    case "underline":
                        EnsureRunProperties(run).Underline = new Underline
                        {
                            Val = new UnderlineValues(value)
                        };
                        break;
                    case "strike":
                        EnsureRunProperties(run).Strike = bool.Parse(value) ? new Strike() : null;
                        break;
                    case "shading":
                    case "shd":
                        // shd has w:val, w:fill, w:color — value format: "fill" or "val;fill" or "val;fill;color"
                        var shdParts = value.Split(';');
                        var shd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            shd.Val = ShadingPatternValues.Clear;
                            shd.Fill = shdParts[0];
                        }
                        else if (shdParts.Length >= 2)
                        {
                            shd.Val = new ShadingPatternValues(shdParts[0]);
                            shd.Fill = shdParts[1];
                            if (shdParts.Length >= 3) shd.Color = shdParts[2];
                        }
                        EnsureRunProperties(run).Shading = shd;
                        break;
                    case "alt":
                        var drawingAlt = run.GetFirstChild<Drawing>();
                        if (drawingAlt != null)
                        {
                            var docPropsAlt = drawingAlt.Descendants<DW.DocProperties>().FirstOrDefault();
                            if (docPropsAlt != null) docPropsAlt.Description = value;
                        }
                        else unsupported.Add(key);
                        break;
                    case "width":
                        var drawingW = run.GetFirstChild<Drawing>();
                        if (drawingW != null)
                        {
                            var extentW = drawingW.Descendants<DW.Extent>().FirstOrDefault();
                            if (extentW != null) extentW.Cx = ParseEmu(value);
                            var extentsW = drawingW.Descendants<A.Extents>().FirstOrDefault();
                            if (extentsW != null) extentsW.Cx = ParseEmu(value);
                        }
                        else unsupported.Add(key);
                        break;
                    case "height":
                        var drawingH = run.GetFirstChild<Drawing>();
                        if (drawingH != null)
                        {
                            var extentH = drawingH.Descendants<DW.Extent>().FirstOrDefault();
                            if (extentH != null) extentH.Cy = ParseEmu(value);
                            var extentsH = drawingH.Descendants<A.Extents>().FirstOrDefault();
                            if (extentsH != null) extentsH.Cy = ParseEmu(value);
                        }
                        else unsupported.Add(key);
                        break;
                    case "link":
                    {
                        var mainPart3 = _doc.MainDocumentPart!;
                        if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            // Remove hyperlink wrapper if present
                            if (run.Parent is Hyperlink existingHlNone)
                            {
                                foreach (var childRun in existingHlNone.Elements<Run>().ToList())
                                    existingHlNone.InsertBeforeSelf(childRun);
                                existingHlNone.Remove();
                            }
                        }
                        else
                        {
                            var newRelId = mainPart3.AddHyperlinkRelationship(new Uri(value), isExternal: true).Id;
                            if (run.Parent is Hyperlink existingHl)
                            {
                                existingHl.Id = newRelId;
                            }
                            else
                            {
                                var newHl = new Hyperlink { Id = newRelId };
                                run.InsertBeforeSelf(newHl);
                                run.Remove();
                                newHl.AppendChild(run);
                            }
                        }
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(EnsureRunProperties(run), key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is Paragraph para)
        {
            var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "style":
                        pProps.ParagraphStyleId = new ParagraphStyleId { Val = value };
                        break;
                    case "alignment":
                        pProps.Justification = new Justification
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => JustificationValues.Center,
                                "right" => JustificationValues.Right,
                                "justify" => JustificationValues.Both,
                                _ => JustificationValues.Left
                            }
                        };
                        break;
                    case "firstlineindent":
                        var indent = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                        indent.FirstLine = (int.Parse(value) * 480).ToString(); // chars to twips (~480 per char)
                        break;
                    case "shading":
                    case "shd":
                        var shdPartsP = value.Split(';');
                        var shdP = new Shading();
                        if (shdPartsP.Length == 1)
                        {
                            shdP.Val = ShadingPatternValues.Clear;
                            shdP.Fill = shdPartsP[0];
                        }
                        else if (shdPartsP.Length >= 2)
                        {
                            shdP.Val = new ShadingPatternValues(shdPartsP[0]);
                            shdP.Fill = shdPartsP[1];
                            if (shdPartsP.Length >= 3) shdP.Color = shdPartsP[2];
                        }
                        pProps.Shading = shdP;
                        break;
                    case "spacebefore":
                        var spacingBefore = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        spacingBefore.Before = value;
                        break;
                    case "spaceafter":
                        var spacingAfter = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        spacingAfter.After = value;
                        break;
                    case "linespacing":
                        var spacingLine = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        spacingLine.Line = value;
                        spacingLine.LineRule = LineSpacingRuleValues.Auto;
                        break;
                    case "numid":
                        var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                        numPr.NumberingId = new NumberingId { Val = int.Parse(value) };
                        break;
                    case "numlevel" or "ilvl":
                        var numPr2 = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                        numPr2.NumberingLevelReference = new NumberingLevelReference { Val = int.Parse(value) };
                        break;
                    case "liststyle":
                        ApplyListStyle(para, value);
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(pProps, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }

        else if (element is TableCell cell)
        {
            var tcPr = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "text":
                        var firstPara = cell.Elements<Paragraph>().FirstOrDefault();
                        if (firstPara == null)
                        {
                            firstPara = new Paragraph();
                            cell.AppendChild(firstPara);
                        }
                        // Remove existing runs
                        foreach (var r in firstPara.Elements<Run>().ToList()) r.Remove();
                        firstPara.AppendChild(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                        break;
                    case "font":
                    case "size":
                    case "bold":
                    case "italic":
                    case "color":
                        // Apply to all runs in all paragraphs in the cell
                        foreach (var cellPara in cell.Elements<Paragraph>())
                        {
                            foreach (var cellRun in cellPara.Elements<Run>())
                            {
                                var rPr = EnsureRunProperties(cellRun);
                                switch (key.ToLowerInvariant())
                                {
                                    case "font":
                                        rPr.RunFonts = new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value };
                                        break;
                                    case "size":
                                        rPr.FontSize = new FontSize { Val = (int.Parse(value) * 2).ToString() };
                                        break;
                                    case "bold":
                                        rPr.Bold = bool.Parse(value) ? new Bold() : null;
                                        break;
                                    case "italic":
                                        rPr.Italic = bool.Parse(value) ? new Italic() : null;
                                        break;
                                    case "color":
                                        rPr.Color = new Color { Val = value.ToUpperInvariant() };
                                        break;
                                }
                            }
                        }
                        break;
                    case "shd" or "shading":
                        var shdParts = value.Split(';');
                        var shd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            shd.Val = ShadingPatternValues.Clear;
                            shd.Fill = shdParts[0];
                        }
                        else if (shdParts.Length >= 2)
                        {
                            shd.Val = new ShadingPatternValues(shdParts[0]);
                            shd.Fill = shdParts[1];
                            if (shdParts.Length >= 3) shd.Color = shdParts[2];
                        }
                        tcPr.Shading = shd;
                        break;
                    case "alignment":
                        var cellFirstPara = cell.Elements<Paragraph>().FirstOrDefault();
                        if (cellFirstPara != null)
                        {
                            var cpProps = cellFirstPara.ParagraphProperties ?? cellFirstPara.PrependChild(new ParagraphProperties());
                            cpProps.Justification = new Justification
                            {
                                Val = value.ToLowerInvariant() switch
                                {
                                    "center" => JustificationValues.Center,
                                    "right" => JustificationValues.Right,
                                    "justify" => JustificationValues.Both,
                                    _ => JustificationValues.Left
                                }
                            };
                        }
                        break;
                    case "valign":
                        tcPr.TableCellVerticalAlignment = new TableCellVerticalAlignment
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => TableVerticalAlignmentValues.Center,
                                "bottom" => TableVerticalAlignmentValues.Bottom,
                                _ => TableVerticalAlignmentValues.Top
                            }
                        };
                        break;
                    case "width":
                        tcPr.TableCellWidth = new TableCellWidth { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    case "vmerge":
                        tcPr.VerticalMerge = new VerticalMerge
                        {
                            Val = value.ToLowerInvariant() == "restart" ? MergedCellValues.Restart : MergedCellValues.Continue
                        };
                        break;
                    case "gridspan":
                        var newSpan = int.Parse(value);
                        tcPr.GridSpan = new GridSpan { Val = newSpan };
                        // Ensure the row has the correct number of tc elements.
                        // Calculate total grid columns occupied by all cells in this row,
                        // then remove/add cells so it matches the table grid.
                        if (element.Parent is TableRow parentRow)
                        {
                            var table = parentRow.Parent as Table;
                            var gridCols = table?.GetFirstChild<TableGrid>()
                                ?.Elements<GridColumn>().Count() ?? 0;
                            if (gridCols > 0)
                            {
                                // Calculate total columns occupied by current cells
                                var totalSpan = parentRow.Elements<TableCell>().Sum(tc =>
                                    tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                                // Remove excess cells after the current cell
                                while (totalSpan > gridCols)
                                {
                                    var nextCell = ((TableCell)element).NextSibling<TableCell>();
                                    if (nextCell == null) break;
                                    totalSpan -= nextCell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                                    nextCell.Remove();
                                }
                            }
                        }
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(tcPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is TableRow row)
        {
            var trPr = row.TableRowProperties ?? row.PrependChild(new TableRowProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "height":
                        trPr.AppendChild(new TableRowHeight { Val = uint.Parse(value), HeightType = HeightRuleValues.AtLeast });
                        break;
                    case "header":
                        if (bool.Parse(value))
                            trPr.AppendChild(new TableHeader());
                        else
                            trPr.GetFirstChild<TableHeader>()?.Remove();
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(trPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is Table tbl)
        {
            var tblPr = tbl.GetFirstChild<TableProperties>() ?? tbl.PrependChild(new TableProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "alignment":
                        tblPr.TableJustification = new TableJustification
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => TableRowAlignmentValues.Center,
                                "right" => TableRowAlignmentValues.Right,
                                _ => TableRowAlignmentValues.Left
                            }
                        };
                        break;
                    case "width":
                        tblPr.TableWidth = new TableWidth { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(tblPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }

        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }
}
