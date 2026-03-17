// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        OpenXmlElement parent;
        if (parentPath is "/" or "" or "/body")
        {
            parent = body;
        }
        else if (parentPath == "/styles")
        {
            // Ensure styles part exists for style operations
            var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
                ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles ??= new Styles();
            parent = stylesPart.Styles;
        }
        else
        {
            var parts = ParsePath(parentPath);
            parent = NavigateToElement(parts)
                ?? throw new ArgumentException($"Path not found: {parentPath}");
        }

        OpenXmlElement newElement;
        string resultPath;

        switch (type.ToLowerInvariant())
        {
            case "paragraph" or "p":
                var para = new Paragraph();
                var pProps = new ParagraphProperties();

                if (properties.TryGetValue("style", out var style))
                    pProps.ParagraphStyleId = new ParagraphStyleId { Val = style };
                if (properties.TryGetValue("alignment", out var alignment))
                    pProps.Justification = new Justification
                    {
                        Val = alignment.ToLowerInvariant() switch
                        {
                            "center" => JustificationValues.Center,
                            "right" => JustificationValues.Right,
                            "justify" => JustificationValues.Both,
                            _ => JustificationValues.Left
                        }
                    };
                if (properties.TryGetValue("firstlineindent", out var indent))
                {
                    var indentVal = (long)(double.Parse(indent, System.Globalization.CultureInfo.InvariantCulture) * 480);
                    if (indentVal < 0 || indentVal > int.MaxValue)
                        throw new OverflowException($"First line indent value out of range: {indent}");
                    pProps.Indentation = new Indentation
                    {
                        FirstLine = indentVal.ToString()
                    };
                }
                if (properties.TryGetValue("spacebefore", out var sb4))
                {
                    var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                    spacing.Before = sb4;
                }
                if (properties.TryGetValue("spaceafter", out var sa4))
                {
                    var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                    spacing.After = sa4;
                }
                if (properties.TryGetValue("linespacing", out var ls4))
                {
                    var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                    spacing.Line = ls4;
                    spacing.LineRule = LineSpacingRuleValues.Auto;
                }
                if (properties.TryGetValue("numid", out var numId))
                {
                    var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                    numPr.NumberingId = new NumberingId { Val = int.Parse(numId) };
                    if (properties.TryGetValue("numlevel", out var numLevel))
                        numPr.NumberingLevelReference = new NumberingLevelReference { Val = int.Parse(numLevel) };
                }
                if (properties.TryGetValue("shd", out var pShdVal) || properties.TryGetValue("shading", out pShdVal))
                {
                    var shdParts = pShdVal.Split(';');
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
                    pProps.Shading = shd;
                }
                if (properties.TryGetValue("leftindent", out var addLI) || properties.TryGetValue("indentleft", out addLI))
                {
                    var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                    ind.Left = addLI;
                }
                if (properties.TryGetValue("rightindent", out var addRI) || properties.TryGetValue("indentright", out addRI))
                {
                    var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                    ind.Right = addRI;
                }
                if (properties.TryGetValue("hangingindent", out var addHI) || properties.TryGetValue("hanging", out addHI))
                {
                    var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                    ind.Hanging = addHI;
                }
                if (properties.TryGetValue("keepnext", out var addKN) && IsTruthy(addKN))
                    pProps.KeepNext = new KeepNext();
                if ((properties.TryGetValue("keeplines", out var addKL) || properties.TryGetValue("keeptogether", out addKL)) && IsTruthy(addKL))
                    pProps.KeepLines = new KeepLines();
                if (properties.TryGetValue("pagebreakbefore", out var addPBB) && IsTruthy(addPBB))
                    pProps.PageBreakBefore = new PageBreakBefore();
                if (properties.TryGetValue("widowcontrol", out var addWC) && IsTruthy(addWC))
                    pProps.WidowControl = new WidowControl();
                if (properties.TryGetValue("liststyle", out var listStyle))
                {
                    para.AppendChild(pProps);
                    int? startVal = properties.TryGetValue("start", out var sv) ? int.Parse(sv) : null;
                    ApplyListStyle(para, listStyle, startVal);
                    // pProps already appended, skip the append below
                    goto paragraphPropsApplied;
                }

                para.AppendChild(pProps);
                paragraphPropsApplied:

                if (properties.TryGetValue("text", out var text))
                {
                    var run = new Run();
                    var rProps = new RunProperties();
                    if (properties.TryGetValue("font", out var font))
                    {
                        rProps.AppendChild(new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font });
                    }
                    if (properties.TryGetValue("size", out var size))
                    {
                        rProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(size) * 2)).ToString() });
                    }
                    if (properties.TryGetValue("bold", out var bold) && IsTruthy(bold))
                        rProps.Bold = new Bold();
                    if (properties.TryGetValue("italic", out var pItalic) && IsTruthy(pItalic))
                        rProps.Italic = new Italic();
                    if (properties.TryGetValue("color", out var pColor))
                        rProps.Color = new Color { Val = pColor.ToUpperInvariant() };
                    if (properties.TryGetValue("underline", out var pUnderline))
                        rProps.Underline = new Underline { Val = new UnderlineValues(pUnderline) };
                    if (properties.TryGetValue("strike", out var pStrike) && IsTruthy(pStrike))
                        rProps.Strike = new Strike();
                    if (properties.TryGetValue("highlight", out var pHighlight))
                        rProps.Highlight = new Highlight { Val = new HighlightColorValues(pHighlight) };
                    if (properties.TryGetValue("caps", out var pCaps) && IsTruthy(pCaps))
                        rProps.Caps = new Caps();
                    if (properties.TryGetValue("smallcaps", out var pSmallCaps) && IsTruthy(pSmallCaps))
                        rProps.SmallCaps = new SmallCaps();
                    if (properties.TryGetValue("superscript", out var pSup) && IsTruthy(pSup))
                        rProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Superscript };
                    if (properties.TryGetValue("subscript", out var pSub) && IsTruthy(pSub))
                        rProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };
                    if (properties.TryGetValue("shd", out var pShd) || properties.TryGetValue("shading", out pShd))
                    {
                        var shdParts = pShd.Split(';');
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
                        rProps.Shading = shd;
                    }

                    run.AppendChild(rProps);
                    run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                    para.AppendChild(run);
                }

                newElement = para;
                var paraCount = parent.Elements<Paragraph>().Count();
                if (index.HasValue && index.Value < paraCount)
                {
                    var refElement = parent.Elements<Paragraph>().ElementAt(index.Value);
                    parent.InsertBefore(para, refElement);
                    resultPath = $"{parentPath}/p[{index.Value + 1}]";
                }
                else
                {
                    parent.AppendChild(para);
                    resultPath = $"{parentPath}/p[{paraCount + 1}]";
                }
                break;

            case "equation" or "formula" or "math":
                if (!properties.TryGetValue("formula", out var formula))
                    throw new ArgumentException("'formula' property is required for equation type");

                var mode = properties.GetValueOrDefault("mode", "display");

                if (mode == "inline" && parent is Paragraph inlinePara)
                {
                    // Insert inline math into existing paragraph
                    var mathElement = FormulaParser.Parse(formula);
                    if (mathElement is M.OfficeMath oMathInline)
                        inlinePara.AppendChild(oMathInline);
                    else
                        inlinePara.AppendChild(new M.OfficeMath(mathElement.CloneNode(true)));
                    var mathCount = inlinePara.Elements<M.OfficeMath>().Count();
                    resultPath = $"{parentPath}/oMath[{mathCount}]";
                    newElement = inlinePara;
                }
                else
                {
                    // Display mode: create m:oMathPara
                    var mathContent = FormulaParser.Parse(formula);
                    M.OfficeMath oMath;
                    if (mathContent is M.OfficeMath directMath)
                        oMath = directMath;
                    else
                        oMath = new M.OfficeMath(mathContent.CloneNode(true));

                    var mathPara = new M.Paragraph(oMath);

                    if (parent is Body || parent is SdtBlock)
                    {
                        // Wrap m:oMathPara in w:p for schema validity
                        var wrapPara = new Paragraph(mathPara);
                        var mathParaCount = parent.Descendants<M.Paragraph>().Count();
                        if (index.HasValue)
                        {
                            var children = parent.ChildElements.ToList();
                            if (index.Value < children.Count)
                                parent.InsertBefore(wrapPara, children[index.Value]);
                            else
                                parent.AppendChild(wrapPara);
                        }
                        else
                        {
                            parent.AppendChild(wrapPara);
                        }
                        resultPath = $"{parentPath}/oMathPara[{mathParaCount + 1}]";
                    }
                    else
                    {
                        parent.AppendChild(mathPara);
                        resultPath = $"{parentPath}/oMathPara[1]";
                    }
                    newElement = mathPara;
                }

                _doc.MainDocumentPart?.Document?.Save();
                return resultPath;

            case "run" or "r":
                if (parent is not Paragraph targetPara)
                    throw new ArgumentException("Runs can only be added to paragraphs");

                var newRun = new Run();
                var newRProps = new RunProperties();
                if (properties.TryGetValue("font", out var rFont))
                    newRProps.AppendChild(new RunFonts { Ascii = rFont, HighAnsi = rFont, EastAsia = rFont });
                if (properties.TryGetValue("size", out var rSize))
                    newRProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(rSize) * 2)).ToString() });
                if (properties.TryGetValue("bold", out var rBold) && IsTruthy(rBold))
                    newRProps.Bold = new Bold();
                if (properties.TryGetValue("italic", out var rItalic) && IsTruthy(rItalic))
                    newRProps.Italic = new Italic();
                if (properties.TryGetValue("color", out var rColor))
                    newRProps.Color = new Color { Val = rColor.ToUpperInvariant() };
                if (properties.TryGetValue("underline", out var rUnderline))
                    newRProps.Underline = new Underline { Val = new UnderlineValues(rUnderline) };
                if (properties.TryGetValue("strike", out var rStrike) && IsTruthy(rStrike))
                    newRProps.Strike = new Strike();
                if (properties.TryGetValue("highlight", out var rHighlight))
                    newRProps.Highlight = new Highlight { Val = new HighlightColorValues(rHighlight) };
                if (properties.TryGetValue("caps", out var rCaps) && IsTruthy(rCaps))
                    newRProps.Caps = new Caps();
                if (properties.TryGetValue("smallcaps", out var rSmallCaps) && IsTruthy(rSmallCaps))
                    newRProps.SmallCaps = new SmallCaps();
                if (properties.TryGetValue("superscript", out var rSup) && IsTruthy(rSup))
                    newRProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Superscript };
                if (properties.TryGetValue("subscript", out var rSub) && IsTruthy(rSub))
                    newRProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };
                if (properties.TryGetValue("shd", out var rShd) || properties.TryGetValue("shading", out rShd))
                {
                    var shdParts = rShd.Split(';');
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
                    newRProps.Shading = shd;
                }

                newRun.AppendChild(newRProps);
                var runText = properties.GetValueOrDefault("text", "");
                newRun.AppendChild(new Text(runText) { Space = SpaceProcessingModeValues.Preserve });

                var runCount = targetPara.Elements<Run>().Count();
                if (index.HasValue && index.Value < runCount)
                {
                    var refRun = targetPara.Elements<Run>().ElementAt(index.Value);
                    targetPara.InsertBefore(newRun, refRun);
                    resultPath = $"{parentPath}/r[{index.Value + 1}]";
                }
                else
                {
                    targetPara.AppendChild(newRun);
                    resultPath = $"{parentPath}/r[{runCount + 1}]";
                }

                newElement = newRun;
                break;

            case "table" or "tbl":
                var table = new Table();
                var tblProps = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new StartBorder { Val = BorderValues.Single, Size = 4 },
                        new EndBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                    )
                );
                table.AppendChild(tblProps);

                int rows = properties.TryGetValue("rows", out var rowsStr) ? int.Parse(rowsStr) : 1;
                int cols = properties.TryGetValue("cols", out var colsStr) ? int.Parse(colsStr) : 1;

                // Add table grid
                var tblGrid = new TableGrid();
                for (int gc = 0; gc < cols; gc++)
                    tblGrid.AppendChild(new GridColumn { Width = "2400" });
                table.AppendChild(tblGrid);

                for (int r = 0; r < rows; r++)
                {
                    var row = new TableRow();
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = new TableCell(new Paragraph());
                        row.AppendChild(cell);
                    }
                    table.AppendChild(row);
                }

                parent.AppendChild(table);
                var tblCount = parent.Elements<Table>().Count();
                resultPath = $"{parentPath}/tbl[{tblCount}]";
                newElement = table;
                break;

            case "row" or "tr":
            {
                if (parent is not Table targetTable)
                    throw new ArgumentException("Rows can only be added to a table: /body/tbl[N]");

                var existingCols = targetTable.Elements<TableGrid>().FirstOrDefault()
                    ?.Elements<GridColumn>().Count() ?? 1;
                int newCols = properties.TryGetValue("cols", out var colsVal) ? int.Parse(colsVal) : existingCols;

                var newRow = new TableRow();
                if (properties.TryGetValue("height", out var rowHeight))
                    newRow.AppendChild(new TableRowProperties(
                        new TableRowHeight { Val = uint.Parse(rowHeight) }));

                for (int c = 0; c < newCols; c++)
                {
                    var cellText = properties.TryGetValue($"c{c + 1}", out var ct) ? ct : "";
                    var cellPara = new Paragraph();
                    if (!string.IsNullOrEmpty(cellText))
                        cellPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
                    newRow.AppendChild(new TableCell(cellPara));
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
                resultPath = $"{parentPath}/tr[{rowIdx}]";
                newElement = newRow;
                break;
            }

            case "cell" or "tc":
            {
                if (parent is not TableRow targetRow)
                    throw new ArgumentException("Cells can only be added to a table row: /body/tbl[N]/tr[M]");

                var cellParagraph = new Paragraph();
                if (properties.TryGetValue("text", out var cellTxt))
                    cellParagraph.AppendChild(new Run(new Text(cellTxt) { Space = SpaceProcessingModeValues.Preserve }));

                var newCell = new TableCell(cellParagraph);

                if (properties.TryGetValue("width", out var cellWidth))
                    newCell.PrependChild(new TableCellProperties(new TableCellWidth { Width = cellWidth, Type = TableWidthUnitValues.Dxa }));

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

                var cellIdx = targetRow.Elements<TableCell>().ToList().IndexOf(newCell) + 1;
                resultPath = $"{parentPath}/tc[{cellIdx}]";
                newElement = newCell;
                break;
            }

            case "picture" or "image" or "img":
                if (!properties.TryGetValue("path", out var imgPath) && !properties.TryGetValue("src", out imgPath))
                    throw new ArgumentException("'path' (or 'src') property is required for picture type");
                if (!File.Exists(imgPath))
                    throw new FileNotFoundException($"Image file not found: {imgPath}");

                var imgExtension = Path.GetExtension(imgPath).ToLowerInvariant();
                var imgPartType = imgExtension switch
                {
                    ".png" => ImagePartType.Png,
                    ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                    ".gif" => ImagePartType.Gif,
                    ".bmp" => ImagePartType.Bmp,
                    ".tif" or ".tiff" => ImagePartType.Tiff,
                    ".emf" => ImagePartType.Emf,
                    ".wmf" => ImagePartType.Wmf,
                    ".svg" => ImagePartType.Svg,
                    _ => throw new ArgumentException($"Unsupported image format: {imgExtension}")
                };

                var mainPart = _doc.MainDocumentPart!;
                var imagePart = mainPart.AddImagePart(imgPartType);
                using (var stream = File.OpenRead(imgPath))
                    imagePart.FeedData(stream);
                var relId = mainPart.GetIdOfPart(imagePart);

                // Determine dimensions (default: 6 inches wide, auto height)
                long cxEmu = 5486400; // 6 inches in EMUs (914400 * 6)
                long cyEmu = 3657600; // 4 inches default
                if (properties.TryGetValue("width", out var widthStr))
                    cxEmu = ParseEmu(widthStr);
                if (properties.TryGetValue("height", out var heightStr))
                    cyEmu = ParseEmu(heightStr);

                var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

                Run imgRun;
                if (properties.TryGetValue("anchor", out var anchorVal) && IsTruthy(anchorVal))
                {
                    var wrapType = properties.GetValueOrDefault("wrap", "none");
                    long hPos = properties.TryGetValue("hposition", out var hPosStr) ? ParseEmu(hPosStr) : 0;
                    long vPos = properties.TryGetValue("vposition", out var vPosStr) ? ParseEmu(vPosStr) : 0;
                    var hRel = properties.TryGetValue("hrelative", out var hRelStr)
                        ? ParseHorizontalRelative(hRelStr)
                        : DW.HorizontalRelativePositionValues.Margin;
                    var vRel = properties.TryGetValue("vrelative", out var vRelStr)
                        ? ParseVerticalRelative(vRelStr)
                        : DW.VerticalRelativePositionValues.Margin;
                    var behind = properties.TryGetValue("behindtext", out var behindStr) && IsTruthy(behindStr);
                    imgRun = CreateAnchorImageRun(relId, cxEmu, cyEmu, altText, wrapType, hPos, vPos, hRel, vRel, behind);
                }
                else
                {
                    imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText);
                }

                Paragraph imgPara;
                if (parent is Paragraph existingPara)
                {
                    existingPara.AppendChild(imgRun);
                    imgPara = existingPara;
                    var imgRunCount = existingPara.Elements<Run>().Count();
                    resultPath = $"{parentPath}/r[{imgRunCount}]";
                }
                else
                {
                    imgPara = new Paragraph(imgRun);
                    var imgParaCount = parent.Elements<Paragraph>().Count();
                    if (index.HasValue && index.Value < imgParaCount)
                    {
                        var refPara = parent.Elements<Paragraph>().ElementAt(index.Value);
                        parent.InsertBefore(imgPara, refPara);
                        resultPath = $"{parentPath}/p[{index.Value + 1}]";
                    }
                    else
                    {
                        parent.AppendChild(imgPara);
                        resultPath = $"{parentPath}/p[{imgParaCount + 1}]";
                    }
                }
                newElement = imgPara;
                break;

            case "comment":
            {
                if (!properties.TryGetValue("text", out var commentText))
                    throw new ArgumentException("'text' property is required for comment type");

                var commentRun = parent as Run;
                var commentPara = commentRun?.Parent as Paragraph ?? parent as Paragraph
                    ?? throw new ArgumentException("Comments must be added to a paragraph or run: /body/p[N] or /body/p[N]/r[M]");

                var author = properties.GetValueOrDefault("author", "officecli");
                var initials = properties.GetValueOrDefault("initials", author[..1]);
                var commentsPart = _doc.MainDocumentPart!.WordprocessingCommentsPart
                    ?? _doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.Comments ??= new Comments();

                var commentId = (commentsPart.Comments.Elements<Comment>()
                    .Select(c => int.TryParse(c.Id?.Value, out var id) ? id : 0)
                    .DefaultIfEmpty(0).Max() + 1).ToString();

                commentsPart.Comments.AppendChild(new Comment(
                    new Paragraph(new Run(new Text(commentText) { Space = SpaceProcessingModeValues.Preserve })))
                {
                    Id = commentId, Author = author, Initials = initials,
                    Date = properties.TryGetValue("date", out var ds) ? DateTime.Parse(ds) : DateTime.UtcNow
                });
                commentsPart.Comments.Save();

                var rangeStart = new CommentRangeStart { Id = commentId };
                var rangeEnd = new CommentRangeEnd { Id = commentId };
                var refRun = new Run(new CommentReference { Id = commentId });

                if (commentRun != null)
                {
                    commentRun.InsertBeforeSelf(rangeStart);
                    commentRun.InsertAfterSelf(rangeEnd);
                    rangeEnd.InsertAfterSelf(refRun);
                }
                else
                {
                    var after = commentPara.ParagraphProperties as OpenXmlElement;
                    if (after != null) after.InsertAfterSelf(rangeStart);
                    else commentPara.InsertAt(rangeStart, 0);
                    commentPara.AppendChild(rangeEnd);
                    commentPara.AppendChild(refRun);
                }

                newElement = rangeStart;
                resultPath = $"{parentPath}/comment[{commentId}]";
                break;
            }

            case "bookmark":
            {
                var bkName = properties.GetValueOrDefault("name", "");
                if (string.IsNullOrEmpty(bkName))
                    throw new ArgumentException("'name' property is required for bookmark");

                var existingIds = body.Descendants<BookmarkStart>()
                    .Select(b => int.TryParse(b.Id?.Value, out var id) ? id : 0);
                var bkId = (existingIds.Any() ? existingIds.Max() + 1 : 1).ToString();

                var bookmarkStart = new BookmarkStart { Id = bkId, Name = bkName };
                var bookmarkEnd = new BookmarkEnd { Id = bkId };

                if (properties.TryGetValue("text", out var bkText))
                {
                    parent.AppendChild(bookmarkStart);
                    parent.AppendChild(new Run(new Text(bkText) { Space = SpaceProcessingModeValues.Preserve }));
                    parent.AppendChild(bookmarkEnd);
                }
                else
                {
                    parent.AppendChild(bookmarkStart);
                    parent.AppendChild(bookmarkEnd);
                }

                newElement = bookmarkStart;
                resultPath = $"/bookmark[{bkName}]";
                break;
            }

            case "hyperlink" or "link":
            {
                if (!properties.TryGetValue("url", out var hlUrl) && !properties.TryGetValue("href", out hlUrl))
                    throw new ArgumentException("'url' property is required for hyperlink type");

                if (parent is not Paragraph hlPara)
                    throw new ArgumentException("Hyperlinks can only be added to paragraphs: /body/p[N]");

                var mainDocPart = _doc.MainDocumentPart!;
                var hlRelId = mainDocPart.AddHyperlinkRelationship(new Uri(hlUrl), isExternal: true).Id;

                var hlRProps = new RunProperties();
                hlRProps.Color = new Color { Val = "0563C1" };
                hlRProps.Underline = new Underline { Val = UnderlineValues.Single };
                if (properties.TryGetValue("font", out var hlFont))
                    hlRProps.RunFonts = new RunFonts { Ascii = hlFont, HighAnsi = hlFont };
                if (properties.TryGetValue("size", out var hlSize))
                    hlRProps.FontSize = new FontSize { Val = ((int)(ParseFontSize(hlSize) * 2)).ToString() };

                var hlRun = new Run(hlRProps);
                var hlText = properties.GetValueOrDefault("text", hlUrl);
                hlRun.AppendChild(new Text(hlText) { Space = SpaceProcessingModeValues.Preserve });

                var hyperlink = new Hyperlink(hlRun) { Id = hlRelId };
                if (index.HasValue)
                    hlPara.InsertAt(hyperlink, index.Value);
                else
                    hlPara.AppendChild(hyperlink);

                var hlCount = hlPara.Elements<Hyperlink>().Count();
                resultPath = $"{parentPath}/hyperlink[{hlCount}]";
                newElement = hyperlink;
                break;
            }

            case "section" or "sectionbreak":
            {
                // Section break: adds SectionProperties to the last paragraph before the break point
                var breakType = properties.GetValueOrDefault("type", "nextPage").ToLowerInvariant();
                var sectType = breakType switch
                {
                    "nextpage" or "next" => SectionMarkValues.NextPage,
                    "continuous" => SectionMarkValues.Continuous,
                    "evenpage" or "even" => SectionMarkValues.EvenPage,
                    "oddpage" or "odd" => SectionMarkValues.OddPage,
                    _ => SectionMarkValues.NextPage
                };

                // Create a paragraph with section properties to mark the break
                var sectPara = new Paragraph();
                var sectPProps = new ParagraphProperties();
                var sectPr = new SectionProperties();
                sectPr.AppendChild(new SectionType { Val = sectType });

                // Copy page size/margins from document section, or use A4 defaults
                var bodySectPr = body.GetFirstChild<SectionProperties>();
                var srcPageSize = bodySectPr?.GetFirstChild<PageSize>();
                sectPr.AppendChild(new PageSize
                {
                    Width = srcPageSize?.Width ?? 11906,   // A4 width
                    Height = srcPageSize?.Height ?? 16838,  // A4 height
                    Orient = srcPageSize?.Orient
                });
                var srcMargin = bodySectPr?.GetFirstChild<PageMargin>();
                sectPr.AppendChild(new PageMargin
                {
                    Top = srcMargin?.Top ?? 1440,
                    Bottom = srcMargin?.Bottom ?? 1440,
                    Left = srcMargin?.Left ?? 1800,
                    Right = srcMargin?.Right ?? 1800
                });

                // Allow per-section overrides
                if (properties.TryGetValue("pagewidth", out var sw) || properties.TryGetValue("width", out sw))
                {
                    if (!uint.TryParse(sw, out var swVal))
                        throw new FormatException($"Invalid page width: {sw}");
                    (sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize())).Width = swVal;
                }
                if (properties.TryGetValue("pageheight", out var sh) || properties.TryGetValue("height", out sh))
                {
                    if (!uint.TryParse(sh, out var shVal))
                        throw new FormatException($"Invalid page height: {sh}");
                    (sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize())).Height = shVal;
                }
                if (properties.TryGetValue("orientation", out var orient))
                {
                    var ps = sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize());
                    ps.Orient = orient.ToLowerInvariant() == "landscape"
                        ? PageOrientationValues.Landscape
                        : PageOrientationValues.Portrait;
                    // Swap width/height for landscape if needed
                    if (ps.Orient == PageOrientationValues.Landscape && ps.Width < ps.Height)
                        (ps.Width!.Value, ps.Height!.Value) = (ps.Height.Value, ps.Width.Value);
                }

                sectPProps.AppendChild(sectPr);
                sectPara.AppendChild(sectPProps);
                parent.AppendChild(sectPara);

                // Count section properties in document
                var secCount = body.Elements<Paragraph>()
                    .Count(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null);
                resultPath = $"/section[{secCount}]";
                newElement = sectPara;
                break;
            }

            case "footnote":
            {
                if (!properties.TryGetValue("text", out var fnText))
                    throw new ArgumentException("'text' property is required for footnote type");

                if (parent is not Paragraph fnPara)
                    throw new ArgumentException("Footnotes must be added to a paragraph: /body/p[N]");

                var mainPart2 = _doc.MainDocumentPart!;
                var fnPart = mainPart2.FootnotesPart ?? mainPart2.AddNewPart<FootnotesPart>();
                fnPart.Footnotes ??= new Footnotes(
                    new Footnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                    new Footnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
                );

                var fnId = (fnPart.Footnotes.Elements<Footnote>()
                    .Where(f => f.Id?.Value > 0)
                    .Select(f => f.Id!.Value)
                    .DefaultIfEmpty(0).Max() + 1);

                var footnote = new Footnote { Id = fnId };
                var fnContentPara = new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
                    new Run(
                        new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                        new FootnoteReferenceMark()),
                    new Run(new Text(" " + fnText) { Space = SpaceProcessingModeValues.Preserve })
                );
                footnote.AppendChild(fnContentPara);
                fnPart.Footnotes.AppendChild(footnote);
                fnPart.Footnotes.Save();

                // Insert reference in document body
                var fnRefRun = new Run(
                    new RunProperties(new RunStyle { Val = "FootnoteReference" }),
                    new FootnoteReference { Id = fnId }
                );
                fnPara.AppendChild(fnRefRun);

                resultPath = $"/footnote[{fnId}]";
                newElement = fnRefRun;
                break;
            }

            case "endnote":
            {
                if (!properties.TryGetValue("text", out var enText))
                    throw new ArgumentException("'text' property is required for endnote type");

                if (parent is not Paragraph enPara)
                    throw new ArgumentException("Endnotes must be added to a paragraph: /body/p[N]");

                var mainPart3 = _doc.MainDocumentPart!;
                var enPart = mainPart3.EndnotesPart ?? mainPart3.AddNewPart<EndnotesPart>();
                enPart.Endnotes ??= new Endnotes(
                    new Endnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                    new Endnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
                );

                var enId = (enPart.Endnotes.Elements<Endnote>()
                    .Where(e => e.Id?.Value > 0)
                    .Select(e => e.Id!.Value)
                    .DefaultIfEmpty(0).Max() + 1);

                var endnote = new Endnote { Id = enId };
                var enContentPara = new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "EndnoteText" }),
                    new Run(
                        new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                        new EndnoteReferenceMark()),
                    new Run(new Text(" " + enText) { Space = SpaceProcessingModeValues.Preserve })
                );
                endnote.AppendChild(enContentPara);
                enPart.Endnotes.AppendChild(endnote);
                enPart.Endnotes.Save();

                // Insert reference in document body
                var enRefRun = new Run(
                    new RunProperties(new RunStyle { Val = "EndnoteReference" }),
                    new EndnoteReference { Id = enId }
                );
                enPara.AppendChild(enRefRun);

                resultPath = $"/endnote[{enId}]";
                newElement = enRefRun;
                break;
            }

            case "toc" or "tableofcontents":
            {
                // Table of Contents field code
                var levels = properties.GetValueOrDefault("levels", "1-3");
                var tocTitle = properties.GetValueOrDefault("title", "");
                var hyperlinks = !properties.TryGetValue("hyperlinks", out var hlVal) || IsTruthy(hlVal);
                var pageNumbers = !properties.TryGetValue("pagenumbers", out var pnVal) || IsTruthy(pnVal);

                // Build field code instruction
                var instrBuilder = new StringBuilder($" TOC \\o \"{levels}\"");
                if (hyperlinks) instrBuilder.Append(" \\h");
                if (!pageNumbers) instrBuilder.Append(" \\z");
                instrBuilder.Append(" \\u ");

                var tocPara = new Paragraph();

                // Optional title
                if (!string.IsNullOrEmpty(tocTitle))
                {
                    var titlePara = new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                        new Run(new Text(tocTitle))
                    );
                    parent.AppendChild(titlePara);
                }

                // Field begin
                tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
                // Field code
                tocPara.AppendChild(new Run(new FieldCode(instrBuilder.ToString()) { Space = SpaceProcessingModeValues.Preserve }));
                // Field separate
                tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                // Placeholder text
                tocPara.AppendChild(new Run(new Text("Update field to see table of contents") { Space = SpaceProcessingModeValues.Preserve }));
                // Field end
                tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

                parent.AppendChild(tocPara);

                // Add UpdateFieldsOnOpen setting
                var settingsPart2 = _doc.MainDocumentPart!.DocumentSettingsPart
                    ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                settingsPart2.Settings ??= new Settings();
                if (settingsPart2.Settings.GetFirstChild<UpdateFieldsOnOpen>() == null)
                    settingsPart2.Settings.AppendChild(new UpdateFieldsOnOpen { Val = true });
                settingsPart2.Settings.Save();

                // Count TOC fields in document to determine index
                var tocCount = body.Elements<Paragraph>()
                    .Count(p => p.Descendants<FieldCode>().Any(fc =>
                        fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)));
                resultPath = $"/toc[{tocCount}]";
                newElement = tocPara;
                break;
            }

            case "style":
            {
                // Create a new style in the styles part
                var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
                    ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles ??= new Styles();

                var styleId = properties.GetValueOrDefault("id", properties.GetValueOrDefault("name", "CustomStyle"));
                var styleName = properties.GetValueOrDefault("name", styleId);
                var styleType = properties.GetValueOrDefault("type", "paragraph").ToLowerInvariant() switch
                {
                    "character" or "char" => StyleValues.Character,
                    "table" => StyleValues.Table,
                    "numbering" => StyleValues.Numbering,
                    _ => StyleValues.Paragraph
                };

                var newStyle = new Style
                {
                    Type = styleType,
                    StyleId = styleId,
                    CustomStyle = true
                };
                newStyle.AppendChild(new StyleName { Val = styleName });

                if (properties.TryGetValue("basedon", out var basedOn) && !string.IsNullOrEmpty(basedOn))
                    newStyle.AppendChild(new BasedOn { Val = basedOn });
                if (properties.TryGetValue("next", out var nextStyle))
                    newStyle.AppendChild(new NextParagraphStyle { Val = nextStyle });

                // Style paragraph properties
                var stylePPr = new StyleParagraphProperties();
                bool hasPPr = false;
                if (properties.TryGetValue("alignment", out var sAlign))
                {
                    stylePPr.Justification = new Justification
                    {
                        Val = sAlign.ToLowerInvariant() switch
                        {
                            "center" => JustificationValues.Center,
                            "right" => JustificationValues.Right,
                            "justify" => JustificationValues.Both,
                            _ => JustificationValues.Left
                        }
                    };
                    hasPPr = true;
                }
                if (properties.TryGetValue("spacebefore", out var sSBefore))
                {
                    var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
                    sp.Before = sSBefore;
                    hasPPr = true;
                }
                if (properties.TryGetValue("spaceafter", out var sSAfter))
                {
                    var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
                    sp.After = sSAfter;
                    hasPPr = true;
                }
                if (hasPPr) newStyle.AppendChild(stylePPr);

                // Style run properties
                var styleRPr = new StyleRunProperties();
                bool hasRPr = false;
                if (properties.TryGetValue("font", out var sFont))
                {
                    styleRPr.RunFonts = new RunFonts { Ascii = sFont, HighAnsi = sFont, EastAsia = sFont };
                    hasRPr = true;
                }
                if (properties.TryGetValue("size", out var sSize))
                {
                    styleRPr.FontSize = new FontSize { Val = ((int)(ParseFontSize(sSize) * 2)).ToString() };
                    hasRPr = true;
                }
                if (properties.TryGetValue("bold", out var sBold) && IsTruthy(sBold))
                {
                    styleRPr.Bold = new Bold();
                    hasRPr = true;
                }
                if (properties.TryGetValue("italic", out var sItalic) && IsTruthy(sItalic))
                {
                    styleRPr.Italic = new Italic();
                    hasRPr = true;
                }
                if (properties.TryGetValue("color", out var sColor))
                {
                    styleRPr.Color = new Color { Val = sColor.ToUpperInvariant() };
                    hasRPr = true;
                }
                if (hasRPr) newStyle.AppendChild(styleRPr);

                stylesPart.Styles.AppendChild(newStyle);
                stylesPart.Styles.Save();

                resultPath = $"/styles/{styleId}";
                newElement = newStyle;
                break;
            }

            case "header":
            {
                var mainPartH = _doc.MainDocumentPart!;
                var headerPart = mainPartH.AddNewPart<HeaderPart>();

                var hPara = new Paragraph();
                var hPProps = new ParagraphProperties();

                if (properties.TryGetValue("alignment", out var hAlign))
                    hPProps.Justification = new Justification
                    {
                        Val = hAlign.ToLowerInvariant() switch
                        {
                            "center" => JustificationValues.Center,
                            "right" => JustificationValues.Right,
                            "justify" => JustificationValues.Both,
                            _ => JustificationValues.Left
                        }
                    };
                hPara.AppendChild(hPProps);

                if (properties.TryGetValue("text", out var hText))
                {
                    var hRun = new Run();
                    var hRProps = new RunProperties();
                    if (properties.TryGetValue("font", out var hFont))
                        hRProps.AppendChild(new RunFonts { Ascii = hFont, HighAnsi = hFont, EastAsia = hFont });
                    if (properties.TryGetValue("size", out var hSize))
                        hRProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(hSize) * 2)).ToString() });
                    if (properties.TryGetValue("bold", out var hBold) && IsTruthy(hBold))
                        hRProps.Bold = new Bold();
                    if (properties.TryGetValue("italic", out var hItalic) && IsTruthy(hItalic))
                        hRProps.Italic = new Italic();
                    if (properties.TryGetValue("color", out var hColor))
                        hRProps.Color = new Color { Val = hColor.ToUpperInvariant() };
                    hRun.AppendChild(hRProps);
                    hRun.AppendChild(new Text(hText) { Space = SpaceProcessingModeValues.Preserve });
                    hPara.AppendChild(hRun);
                }

                headerPart.Header = new Header(hPara);
                headerPart.Header.Save();

                var hBody = mainPartH.Document!.Body!;
                var hSectPr = hBody.Elements<SectionProperties>().LastOrDefault()
                    ?? hBody.AppendChild(new SectionProperties());

                var headerType = HeaderFooterValues.Default;
                if (properties.TryGetValue("type", out var hTypeStr))
                {
                    headerType = hTypeStr.ToLowerInvariant() switch
                    {
                        "first" => HeaderFooterValues.First,
                        "even" => HeaderFooterValues.Even,
                        _ => HeaderFooterValues.Default
                    };
                }

                var headerRef = new HeaderReference
                {
                    Id = mainPartH.GetIdOfPart(headerPart),
                    Type = headerType
                };
                hSectPr.PrependChild(headerRef);

                if (headerType == HeaderFooterValues.First)
                {
                    var settingsPart = mainPartH.DocumentSettingsPart
                        ?? mainPartH.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<TitlePage>() == null)
                        settingsPart.Settings.AppendChild(new TitlePage());
                    settingsPart.Settings.Save();
                }

                mainPartH.Document.Save();
                var hIdx = mainPartH.HeaderParts.ToList().IndexOf(headerPart);
                return $"/header[{hIdx + 1}]";
            }

            case "footer":
            {
                var mainPartF = _doc.MainDocumentPart!;
                var footerPart = mainPartF.AddNewPart<FooterPart>();

                var fPara = new Paragraph();
                var fPProps = new ParagraphProperties();

                if (properties.TryGetValue("alignment", out var fAlign))
                    fPProps.Justification = new Justification
                    {
                        Val = fAlign.ToLowerInvariant() switch
                        {
                            "center" => JustificationValues.Center,
                            "right" => JustificationValues.Right,
                            "justify" => JustificationValues.Both,
                            _ => JustificationValues.Left
                        }
                    };
                fPara.AppendChild(fPProps);

                if (properties.TryGetValue("text", out var fText))
                {
                    var fRun = new Run();
                    var fRProps = new RunProperties();
                    if (properties.TryGetValue("font", out var fFont))
                        fRProps.AppendChild(new RunFonts { Ascii = fFont, HighAnsi = fFont, EastAsia = fFont });
                    if (properties.TryGetValue("size", out var fSize))
                        fRProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(fSize) * 2)).ToString() });
                    if (properties.TryGetValue("bold", out var fBold) && IsTruthy(fBold))
                        fRProps.Bold = new Bold();
                    if (properties.TryGetValue("italic", out var fItalic) && IsTruthy(fItalic))
                        fRProps.Italic = new Italic();
                    if (properties.TryGetValue("color", out var fColor))
                        fRProps.Color = new Color { Val = fColor.ToUpperInvariant() };
                    fRun.AppendChild(fRProps);
                    fRun.AppendChild(new Text(fText) { Space = SpaceProcessingModeValues.Preserve });
                    fPara.AppendChild(fRun);
                }

                footerPart.Footer = new Footer(fPara);
                footerPart.Footer.Save();

                var fBody = mainPartF.Document!.Body!;
                var fSectPr = fBody.Elements<SectionProperties>().LastOrDefault()
                    ?? fBody.AppendChild(new SectionProperties());

                var footerType = HeaderFooterValues.Default;
                if (properties.TryGetValue("type", out var fTypeStr))
                {
                    footerType = fTypeStr.ToLowerInvariant() switch
                    {
                        "first" => HeaderFooterValues.First,
                        "even" => HeaderFooterValues.Even,
                        _ => HeaderFooterValues.Default
                    };
                }

                var footerRef = new FooterReference
                {
                    Id = mainPartF.GetIdOfPart(footerPart),
                    Type = footerType
                };
                fSectPr.PrependChild(footerRef);

                if (footerType == HeaderFooterValues.First)
                {
                    var settingsPart = mainPartF.DocumentSettingsPart
                        ?? mainPartF.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<TitlePage>() == null)
                        settingsPart.Settings.AppendChild(new TitlePage());
                    settingsPart.Settings.Save();
                }

                mainPartF.Document.Save();
                var fIdx = mainPartF.FooterParts.ToList().IndexOf(footerPart);
                return $"/footer[{fIdx + 1}]";
            }

            default:
            {
                // Generic fallback: create typed element via SDK schema validation
                var created = GenericXmlQuery.TryCreateTypedElement(parent, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                newElement = created;
                var siblings = parent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                resultPath = $"{parentPath}/{created.LocalName}[{createdIdx}]";
                break;
            }
        }

        _doc.MainDocumentPart?.Document?.Save();
        return resultPath;
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var mainPart = _doc.MainDocumentPart!;

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                var chartPart = mainPart.AddNewPart<ChartPart>();
                var relId = mainPart.GetIdOfPart(chartPart);
                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new C.ChartSpace(
                    new C.Chart(new C.PlotArea(new C.Layout()))
                );
                chartPart.ChartSpace.Save();
                var chartIdx = mainPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/chart[{chartIdx + 1}]");

            case "header":
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                var hRelId = mainPart.GetIdOfPart(headerPart);
                headerPart.Header = new Header(new Paragraph());
                headerPart.Header.Save();
                var hIdx = mainPart.HeaderParts.ToList().IndexOf(headerPart);
                return (hRelId, $"/header[{hIdx + 1}]");

            case "footer":
                var footerPart = mainPart.AddNewPart<FooterPart>();
                var fRelId = mainPart.GetIdOfPart(footerPart);
                footerPart.Footer = new Footer(new Paragraph());
                footerPart.Footer.Save();
                var fIdx = mainPart.FooterParts.ToList().IndexOf(footerPart);
                return (fRelId, $"/footer[{fIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, header, footer");
        }
    }

    public void Remove(string path)
    {
        var parts = ParsePath(path);
        var element = NavigateToElement(parts)
            ?? throw new ArgumentException($"Path not found: {path}");

        element.Remove();
        _doc.MainDocumentPart?.Document?.Save();
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        // Determine target parent
        string effectiveParentPath;
        OpenXmlElement targetParent;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within current parent
            targetParent = element.Parent
                ?? throw new InvalidOperationException("Element has no parent");
            // Compute parent path by removing last segment
            var lastSlash = sourcePath.LastIndexOf('/');
            effectiveParentPath = lastSlash > 0 ? sourcePath[..lastSlash] : "/body";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            if (targetParentPath is "/" or "" or "/body")
                targetParent = _doc.MainDocumentPart!.Document!.Body!;
            else
            {
                var tgtParts = ParsePath(targetParentPath);
                targetParent = NavigateToElement(tgtParts)
                    ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
            }
        }

        element.Remove();

        // Insert at the specified position among same-type siblings
        if (index.HasValue)
        {
            var sameTypeSiblings = targetParent.ChildElements
                .Where(e => e.LocalName == element.LocalName).ToList();
            var insertIdx = index.Value - 1; // Convert 1-based to 0-based
            if (insertIdx >= 0 && insertIdx < sameTypeSiblings.Count)
                sameTypeSiblings[insertIdx].InsertBeforeSelf(element);
            else
                targetParent.AppendChild(element);
        }
        else
        {
            targetParent.AppendChild(element);
        }

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == element.LocalName).ToList();
        var newIdx = siblings.IndexOf(element) + 1;
        return $"{effectiveParentPath}/{element.LocalName}[{newIdx}]";
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        var clone = element.CloneNode(true);

        OpenXmlElement targetParent;
        if (targetParentPath is "/" or "" or "/body")
            targetParent = _doc.MainDocumentPart!.Document!.Body!;
        else
        {
            var tgtParts = ParsePath(targetParentPath);
            targetParent = NavigateToElement(tgtParts)
                ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
        }

        InsertAtPosition(targetParent, clone, index);

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == clone.LocalName).ToList();
        var newIdx = siblings.IndexOf(clone) + 1;
        return $"{targetParentPath}/{clone.LocalName}[{newIdx}]";
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                parent.AppendChild(element);
        }
        else
        {
            parent.AppendChild(element);
        }
    }

    private void SetDocumentProperties(Dictionary<string, string> properties)
    {
        var doc = _doc.MainDocumentPart?.Document
            ?? throw new InvalidOperationException("Document not found");

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pagebackground" or "background":
                    doc.DocumentBackground = new DocumentBackground { Color = value };
                    // Enable background display in settings
                    var settingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<DisplayBackgroundShape>() == null)
                        settingsPart.Settings.AppendChild(new DisplayBackgroundShape());
                    settingsPart.Settings.Save();
                    break;

                case "defaultfont":
                    var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart;
                    if (stylesPart?.Styles != null)
                    {
                        var defaultRunProps = stylesPart.Styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                        if (defaultRunProps != null)
                        {
                            var fonts = defaultRunProps.GetFirstChild<RunFonts>()
                                ?? defaultRunProps.AppendChild(new RunFonts());
                            fonts.Ascii = value;
                            fonts.HighAnsi = value;
                            fonts.EastAsia = value;
                            stylesPart.Styles.Save();
                        }
                    }
                    break;

                case "pagewidth":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Width = uint.Parse(value);
                    break;
                case "pageheight":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Height = uint.Parse(value);
                    break;
                case "margintop":
                    EnsurePageMargin().Top = int.Parse(value);
                    break;
                case "marginbottom":
                    EnsurePageMargin().Bottom = int.Parse(value);
                    break;
                case "marginleft":
                    EnsurePageMargin().Left = uint.Parse(value);
                    break;
                case "marginright":
                    EnsurePageMargin().Right = uint.Parse(value);
                    break;

                // Core document properties
                case "title":
                    _doc.PackageProperties.Title = value;
                    break;
                case "author" or "creator":
                    _doc.PackageProperties.Creator = value;
                    break;
                case "subject":
                    _doc.PackageProperties.Subject = value;
                    break;
                case "keywords":
                    _doc.PackageProperties.Keywords = value;
                    break;
                case "description":
                    _doc.PackageProperties.Description = value;
                    break;
                case "category":
                    _doc.PackageProperties.Category = value;
                    break;
                case "lastmodifiedby":
                    _doc.PackageProperties.LastModifiedBy = value;
                    break;
                case "revision":
                    _doc.PackageProperties.Revision = value;
                    break;
            }
        }
    }

    private SectionProperties EnsureSectionProperties()
    {
        var body = _doc.MainDocumentPart!.Document!.Body!;
        var sectPr = body.GetFirstChild<SectionProperties>();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
        if (sectPr.GetFirstChild<PageSize>() == null)
            sectPr.AppendChild(new PageSize { Width = 11906, Height = 16838 }); // A4 default
        return sectPr;
    }

    private PageMargin EnsurePageMargin()
    {
        var sectPr = EnsureSectionProperties();
        var margin = sectPr.GetFirstChild<PageMargin>();
        if (margin == null)
        {
            margin = new PageMargin { Top = 1440, Bottom = 1440, Left = 1800, Right = 1800 };
            sectPr.AppendChild(margin);
        }
        return margin;
    }
}
