// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
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
                    pProps.Indentation = new Indentation
                    {
                        FirstLine = (int.Parse(indent) * 480).ToString()
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
                if (properties.TryGetValue("liststyle", out var listStyle))
                {
                    para.AppendChild(pProps);
                    ApplyListStyle(para, listStyle);
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
                        rProps.AppendChild(new FontSize { Val = (int.Parse(size) * 2).ToString() });
                    }
                    if (properties.TryGetValue("bold", out var bold) && bool.Parse(bold))
                        rProps.Bold = new Bold();
                    if (properties.TryGetValue("italic", out var pItalic) && bool.Parse(pItalic))
                        rProps.Italic = new Italic();
                    if (properties.TryGetValue("color", out var pColor))
                        rProps.Color = new Color { Val = pColor.ToUpperInvariant() };
                    if (properties.TryGetValue("underline", out var pUnderline))
                        rProps.Underline = new Underline { Val = new UnderlineValues(pUnderline) };
                    if (properties.TryGetValue("strike", out var pStrike) && bool.Parse(pStrike))
                        rProps.Strike = new Strike();
                    if (properties.TryGetValue("highlight", out var pHighlight))
                        rProps.Highlight = new Highlight { Val = new HighlightColorValues(pHighlight) };
                    if (properties.TryGetValue("caps", out var pCaps) && bool.Parse(pCaps))
                        rProps.Caps = new Caps();
                    if (properties.TryGetValue("smallcaps", out var pSmallCaps) && bool.Parse(pSmallCaps))
                        rProps.SmallCaps = new SmallCaps();
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
                    newRProps.AppendChild(new FontSize { Val = (int.Parse(rSize) * 2).ToString() });
                if (properties.TryGetValue("bold", out var rBold) && bool.Parse(rBold))
                    newRProps.Bold = new Bold();
                if (properties.TryGetValue("italic", out var rItalic) && bool.Parse(rItalic))
                    newRProps.Italic = new Italic();
                if (properties.TryGetValue("color", out var rColor))
                    newRProps.Color = new Color { Val = rColor.ToUpperInvariant() };
                if (properties.TryGetValue("underline", out var rUnderline))
                    newRProps.Underline = new Underline { Val = new UnderlineValues(rUnderline) };
                if (properties.TryGetValue("strike", out var rStrike) && bool.Parse(rStrike))
                    newRProps.Strike = new Strike();
                if (properties.TryGetValue("highlight", out var rHighlight))
                    newRProps.Highlight = new Highlight { Val = new HighlightColorValues(rHighlight) };
                if (properties.TryGetValue("caps", out var rCaps) && bool.Parse(rCaps))
                    newRProps.Caps = new Caps();
                if (properties.TryGetValue("smallcaps", out var rSmallCaps) && bool.Parse(rSmallCaps))
                    newRProps.SmallCaps = new SmallCaps();
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

            case "picture" or "image" or "img":
                if (!properties.TryGetValue("path", out var imgPath))
                    throw new ArgumentException("'path' property is required for picture type");
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

                var imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText);

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
                    hlRProps.FontSize = new FontSize { Val = (int.Parse(hlSize) * 2).ToString() };

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
        InsertAtPosition(targetParent, element, index);

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
