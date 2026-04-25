// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

// Per-path-pattern Set helpers extracted from the WordHandler.Set() entry
// method. Each helper owns one path-pattern's full handling. Mechanically
// extracted, no behavior change.
public partial class WordHandler
{
    private List<string> SetWatermarkPath(Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        // Find watermark VML shape in headers and modify properties
        foreach (var hp in _doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
        {
            if (hp.Header == null) continue;
            var picts = hp.Header.Descendants<Picture>().ToList();
            foreach (var pict in picts)
            {
                if (!pict.InnerXml.Contains("WaterMark", StringComparison.OrdinalIgnoreCase)) continue;

                // Rebuild VML with updated properties — parse existing values as defaults
                var xml = pict.InnerXml;
                foreach (var (key, value) in properties)
                {
                    switch (key.ToLowerInvariant())
                    {
                        case "text":
                            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                @"string=""[^""]*""", $@"string=""{System.Security.SecurityElement.Escape(value)}""");
                            break;
                        case "color":
                            var clr = "#" + SanitizeHex(value);
                            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                @"fillcolor=""[^""]*""", $@"fillcolor=""{clr}""");
                            break;
                        case "font":
                            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                @"font-family:&quot;[^&]*&quot;", $@"font-family:&quot;{System.Security.SecurityElement.Escape(value)}&quot;");
                            break;
                        case "opacity":
                            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                @"opacity=""[^""]*""", $@"opacity=""{value}""");
                            break;
                        case "rotation":
                            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                @"rotation:\d+", $@"rotation:{value}");
                            break;
                        default:
                            unsupported.Add(key);
                            break;
                    }
                }
                pict.InnerXml = xml;
            }
            hp.Header.Save();
        }
        return unsupported;
    }

    private List<string> SetChartAxisPath(System.Text.RegularExpressions.Match chartAxisSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var caChartIdx = int.Parse(chartAxisSetMatch.Groups[1].Value);
        var caRole = chartAxisSetMatch.Groups[2].Value;
        var caAllCharts = GetAllWordCharts();
        if (caAllCharts.Count == 0)
            throw new ArgumentException("No charts in this document");
        if (caChartIdx < 1 || caChartIdx > caAllCharts.Count)
            throw new ArgumentException($"Chart {caChartIdx} not found (total: {caAllCharts.Count})");
        var caChartInfo = caAllCharts[caChartIdx - 1];
        if (caChartInfo.IsExtended || caChartInfo.StandardPart == null)
            throw new ArgumentException($"Axis Set not supported on extended charts.");
        unsupported.AddRange(Core.ChartHelper.SetAxisProperties(
            caChartInfo.StandardPart, caRole, properties));
        return unsupported;
    }

    private List<string> SetChartPath(System.Text.RegularExpressions.Match chartMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var chartIdx = int.Parse(chartMatch.Groups[1].Value);
        var allCharts = GetAllWordCharts();
        if (allCharts.Count == 0)
            throw new ArgumentException("No charts in this document");
        if (chartIdx < 1 || chartIdx > allCharts.Count)
            throw new ArgumentException($"Chart {chartIdx} not found (total: {allCharts.Count})");

        var chartInfo = allCharts[chartIdx - 1];

        // If series sub-path, prefix all properties with series{N}. for ChartSetter
        var chartProps = properties;
        var isSeriesPath = chartMatch.Groups[2].Success;
        if (isSeriesPath)
        {
            var seriesIdx = int.Parse(chartMatch.Groups[2].Value);
            chartProps = new Dictionary<string, string>();
            foreach (var (key, value) in properties)
                chartProps[$"series{seriesIdx}.{key}"] = value;
        }

        // Chart-level position/size Set — mutate the hosting wp:inline's
        // wp:extent. Word inline charts have no positional x/y (they
        // flow in text), so only width/height are meaningful here.
        //
        // CONSISTENCY(chart-position-set): same vocabulary as Excel and
        // PPTX. x/y are silently dropped (flagged as unsupported) since
        // inline mode has no absolute position.
        if (!isSeriesPath && chartInfo.Inline != null)
        {
            ApplyWordChartPositionSet(chartInfo.Inline, chartProps, unsupported);
            // Drop ALL position keys (x/y/width/height) from chartProps
            // after handling — unsupported ones were already reported by
            // ApplyWordChartPositionSet. Forwarding them to ChartHelper
            // would double-report them.
            foreach (var k in new[] { "x", "y", "width", "height" })
            {
                var matched = chartProps.Keys
                    .FirstOrDefault(key => key.Equals(k, StringComparison.OrdinalIgnoreCase));
                if (matched != null) chartProps.Remove(matched);
            }
        }

        if (chartInfo.IsExtended)
        {
            // cx:chart — delegates to ChartExBuilder.SetChartProperties.
            // Same shared implementation as Excel/PPTX: title/axis/gridline
            // styling, series fill, histogram binning, etc.
            unsupported.AddRange(Core.ChartExBuilder.SetChartProperties(
                chartInfo.ExtendedPart!, chartProps));
        }
        else
        {
            unsupported.AddRange(Core.ChartHelper.SetChartProperties(chartInfo.StandardPart!, chartProps));
        }
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private List<string> SetFieldPath(System.Text.RegularExpressions.Match fieldSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var fieldIdx = int.Parse(fieldSetMatch.Groups[1].Value);
        var allFields = FindFields();
        if (fieldIdx < 1 || fieldIdx > allFields.Count)
            throw new ArgumentException($"Field {fieldIdx} not found (total: {allFields.Count})");

        var field = allFields[fieldIdx - 1];

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "instruction" or "instr":
                    field.InstrCode.Text = value.StartsWith(" ") ? value : $" {value} ";
                    // Auto-mark dirty when instruction changes
                    var beginCharI = field.BeginRun.GetFirstChild<FieldChar>();
                    if (beginCharI != null) beginCharI.Dirty = true;
                    break;
                case "text" or "result":
                    // Replace result text (between separate and end)
                    if (field.ResultRuns.Count > 0)
                    {
                        // Set text on first result run, clear the rest
                        var firstResultText = field.ResultRuns[0].GetFirstChild<Text>();
                        if (firstResultText != null)
                            firstResultText.Text = value;
                        else
                            field.ResultRuns[0].AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                        for (int ri = 1; ri < field.ResultRuns.Count; ri++)
                        {
                            var t = field.ResultRuns[ri].GetFirstChild<Text>();
                            if (t != null) t.Text = "";
                        }
                    }
                    break;
                case "dirty":
                    var beginCharD = field.BeginRun.GetFirstChild<FieldChar>();
                    if (beginCharD != null) beginCharD.Dirty = IsTruthy(value);
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private List<string> SetTocPath(System.Text.RegularExpressions.Match tocMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var tocIdx = int.Parse(tocMatch.Groups[1].Value);
        var tocParas = FindTocParagraphs();
        if (tocIdx < 1 || tocIdx > tocParas.Count)
            throw new ArgumentException($"TOC {tocIdx} not found (total: {tocParas.Count})");

        var tocPara = tocParas[tocIdx - 1];

        // Rebuild the field code from properties
        var instrRun = tocPara.Descendants<Run>()
            .FirstOrDefault(r => r.GetFirstChild<FieldCode>() != null);
        if (instrRun == null)
            throw new InvalidOperationException("TOC field code not found");

        var fieldCode = instrRun.GetFirstChild<FieldCode>()!;
        var instr = fieldCode.Text ?? "";

        // Update levels
        if (properties.TryGetValue("levels", out var newLevels))
        {
            var levelsRx = System.Text.RegularExpressions.Regex.Match(instr, @"\\o\s+""[^""]+""");
            instr = levelsRx.Success
                ? instr.Replace(levelsRx.Value, $"\\o \"{newLevels}\"")
                : instr.TrimEnd() + $" \\o \"{newLevels}\" ";
        }

        // Update hyperlinks switch
        if (properties.TryGetValue("hyperlinks", out var hlSwitch))
        {
            if (IsTruthy(hlSwitch) && !instr.Contains("\\h"))
                instr = instr.TrimEnd() + " \\h ";
            else if (!IsTruthy(hlSwitch))
                instr = instr.Replace("\\h", "").Replace("  ", " ");
        }

        // Update page numbers switch (\\z = hide page numbers)
        if (properties.TryGetValue("pagenumbers", out var pnSwitch))
        {
            if (!IsTruthy(pnSwitch) && !instr.Contains("\\z"))
                instr = instr.TrimEnd() + " \\z ";
            else if (IsTruthy(pnSwitch))
                instr = instr.Replace("\\z", "").Replace("  ", " ");
        }

        fieldCode.Text = instr;

        // Mark field as dirty so Word updates it on open
        var beginRun = tocPara.Descendants<Run>()
            .FirstOrDefault(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Begin);
        if (beginRun != null)
        {
            var fldChar = beginRun.GetFirstChild<FieldChar>()!;
            fldChar.Dirty = true;
        }

        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private List<string> SetFootnotePath(System.Text.RegularExpressions.Match fnSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var fnId = int.Parse(fnSetMatch.Groups[1].Value);
        var fn = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
            .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId);
        if (fn == null)
        {
            // Try ordinal lookup (1-based index among user footnotes)
            var userFns = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
                .Elements<Footnote>().Where(f => f.Id?.Value > 0).ToList();
            if (userFns != null && fnId >= 1 && fnId <= userFns.Count)
                fn = userFns[fnId - 1];
            else
                throw new ArgumentException($"Footnote {fnId} not found");
        }

        if (properties.TryGetValue("text", out var fnText))
        {
            // Find the content paragraph (skip the reference mark run)
            var contentRuns = fn.Descendants<Run>()
                .Where(r => r.GetFirstChild<FootnoteReferenceMark>() == null).ToList();
            if (contentRuns.Count > 0)
            {
                // Update first content run; keep space as separate element
                var textEl = contentRuns[0].GetFirstChild<Text>();
                if (textEl != null)
                {
                    textEl.Text = fnText;
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                }
                else
                    contentRuns[0].AppendChild(new Text(fnText) { Space = SpaceProcessingModeValues.Preserve });
                // Remove extra runs so text is not duplicated
                for (int i = 1; i < contentRuns.Count; i++)
                    contentRuns[i].Remove();
            }
        }
        // Report any keys besides "text" as unsupported
        foreach (var k in properties.Keys)
        {
            if (!k.Equals("text", StringComparison.OrdinalIgnoreCase))
                unsupported.Add(k);
        }
        _doc.MainDocumentPart?.FootnotesPart?.Footnotes?.Save();
        return unsupported;
    }

    private List<string> SetEndnotePath(System.Text.RegularExpressions.Match enSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var enId = int.Parse(enSetMatch.Groups[1].Value);
        var en = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
            .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId);
        if (en == null)
        {
            // Try ordinal lookup (1-based index among user endnotes)
            var userEns = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
                .Elements<Endnote>().Where(e => e.Id?.Value > 0).ToList();
            if (userEns != null && enId >= 1 && enId <= userEns.Count)
                en = userEns[enId - 1];
            else
                throw new ArgumentException($"Endnote {enId} not found");
        }

        if (properties.TryGetValue("text", out var enText))
        {
            var contentRuns = en.Descendants<Run>()
                .Where(r => r.GetFirstChild<EndnoteReferenceMark>() == null).ToList();
            if (contentRuns.Count > 0)
            {
                var textEl = contentRuns[0].GetFirstChild<Text>();
                if (textEl != null)
                {
                    textEl.Text = enText;
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                }
                else
                    contentRuns[0].AppendChild(new Text(enText) { Space = SpaceProcessingModeValues.Preserve });
                // Remove extra runs so text is not duplicated
                for (int i = 1; i < contentRuns.Count; i++)
                    contentRuns[i].Remove();
            }
        }
        // Report any keys besides "text" as unsupported
        foreach (var k in properties.Keys)
        {
            if (!k.Equals("text", StringComparison.OrdinalIgnoreCase))
                unsupported.Add(k);
        }
        _doc.MainDocumentPart?.EndnotesPart?.Endnotes?.Save();
        return unsupported;
    }

    private List<string> SetSectionPath(System.Text.RegularExpressions.Match secSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var secIdxStr = secSetMatch.Groups[1].Success ? secSetMatch.Groups[1].Value
            : (secSetMatch.Groups[2].Success ? secSetMatch.Groups[2].Value : "1");
        var secIdx = int.Parse(secIdxStr);
        var sectionProps = FindSectionProperties();

        // If no section properties exist and requesting section 1, create one
        if (sectionProps.Count == 0 && secIdx == 1)
        {
            var sBody = _doc.MainDocumentPart?.Document?.Body;
            if (sBody != null)
            {
                var newSectPr = new SectionProperties();
                sBody.AppendChild(newSectPr);
                sectionProps = FindSectionProperties();
            }
        }

        if (secIdx < 1 || secIdx > sectionProps.Count)
            throw new ArgumentException($"Section {secIdx} not found (total: {sectionProps.Count})");

        var sectPr = sectionProps[secIdx - 1];
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "type":
                    var st = sectPr.GetFirstChild<SectionType>() ?? sectPr.PrependChild(new SectionType());
                    st.Val = value.ToLowerInvariant() switch
                    {
                        "nextpage" or "next" => SectionMarkValues.NextPage,
                        "continuous" => SectionMarkValues.Continuous,
                        "evenpage" or "even" => SectionMarkValues.EvenPage,
                        "oddpage" or "odd" => SectionMarkValues.OddPage,
                        _ => throw new ArgumentException($"Invalid section break type: '{value}'. Valid values: nextPage, continuous, evenPage, oddPage.")
                    };
                    break;
                case "pagewidth" or "pageWidth":
                    EnsureSectPrPageSize(sectPr).Width = ParseTwips(value);
                    break;
                case "pageheight" or "pageHeight":
                    EnsureSectPrPageSize(sectPr).Height = ParseTwips(value);
                    break;
                case "orientation":
                {
                    var ps = EnsureSectPrPageSize(sectPr);
                    var isLandscape = value.ToLowerInvariant() == "landscape";
                    ps.Orient = isLandscape
                        ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
                    // Default to A4 if no dimensions set
                    var w = ps.Width?.Value ?? WordPageDefaults.A4WidthTwips;
                    var h = ps.Height?.Value ?? WordPageDefaults.A4HeightTwips;
                    // Swap width/height if orientation changes and dimensions are misaligned
                    if ((isLandscape && w < h) || (!isLandscape && w > h))
                    {
                        ps.Width = h;
                        ps.Height = w;
                    }
                    break;
                }
                case "margintop":
                    EnsureSectPrPageMargin(sectPr).Top = (int)ParseTwips(value);
                    break;
                case "marginbottom":
                    EnsureSectPrPageMargin(sectPr).Bottom = (int)ParseTwips(value);
                    break;
                case "marginleft":
                    EnsureSectPrPageMargin(sectPr).Left = ParseTwips(value);
                    break;
                case "marginright":
                    EnsureSectPrPageMargin(sectPr).Right = ParseTwips(value);
                    break;
                case "columns" or "cols" or "col":
                {
                    // Equal-width columns: "3" or "3,720" (count,space in twips)
                    var eqCols = EnsureColumns(sectPr);
                    var colParts = value.Split(',');
                    if (!short.TryParse(colParts[0], out var colCount))
                        throw new ArgumentException($"Invalid 'columns' value: '{value}'. Expected an integer or integer,space (e.g. '3' or '3,720').");
                    eqCols.ColumnCount = (Int16Value)colCount;
                    eqCols.EqualWidth = true;
                    if (colParts.Length > 1)
                        eqCols.Space = colParts[1];
                    else
                        eqCols.Space ??= "720"; // default ~1.27cm
                    // Remove any individual column definitions for equal width
                    eqCols.RemoveAllChildren<Column>();
                    break;
                }
                case "columnspace" or "columns.space":
                {
                    // Standalone column-spacing update — preserves existing
                    // column count/widths. Pairs with the canonical 'columnSpace'
                    // key returned by Get/Query (WordHandler.Query.cs:491).
                    var spaceCols = EnsureColumns(sectPr);
                    spaceCols.Space = ParseTwips(value).ToString();
                    break;
                }
                case "colwidths":
                {
                    // Custom column widths: "3000,720,2000,720,3000"
                    // Alternating: width,space,width,space,...,width
                    var cwCols = EnsureColumns(sectPr);
                    cwCols.EqualWidth = false;
                    cwCols.RemoveAllChildren<Column>();
                    var vals = value.Split(',');
                    int colCount = 0;
                    for (int ci = 0; ci < vals.Length; ci += 2)
                    {
                        var col = new Column { Width = vals[ci] };
                        if (ci + 1 < vals.Length)
                            col.Space = vals[ci + 1];
                        cwCols.AppendChild(col);
                        colCount++;
                    }
                    cwCols.ColumnCount = (Int16Value)(short)colCount;
                    break;
                }
                case "separator" or "sep":
                {
                    var sepCols = EnsureColumns(sectPr);
                    sepCols.Separator = IsTruthy(value);
                    break;
                }
                case "linenumbers" or "linenumbering":
                {
                    var lower = value.ToLowerInvariant();
                    if (lower == "none" || lower == "off" || lower == "false")
                    {
                        sectPr.RemoveAllChildren<LineNumberType>();
                    }
                    else
                    {
                        var lnNum = sectPr.GetFirstChild<LineNumberType>();
                        if (lnNum == null)
                        {
                            lnNum = new LineNumberType();
                            sectPr.AppendChild(lnNum);
                        }
                        // If value is a number, set CountBy to that number
                        if (int.TryParse(lower, out var countBy))
                        {
                            lnNum.CountBy = (short)countBy;
                            lnNum.Restart = LineNumberRestartValues.Continuous;
                        }
                        else
                        {
                            lnNum.CountBy = 1;
                            lnNum.Restart = lower switch
                            {
                                "continuous" => LineNumberRestartValues.Continuous,
                                "restartpage" or "page" => LineNumberRestartValues.NewPage,
                                "restartsection" or "section" => LineNumberRestartValues.NewSection,
                                _ => LineNumberRestartValues.Continuous
                            };
                        }
                    }
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private List<string> SetStylePath(System.Text.RegularExpressions.Match styleSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var styleId = styleSetMatch.Groups[1].Value;
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart
            ?? _doc.MainDocumentPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        if (stylesPart.Styles == null) stylesPart.Styles = new Styles();
        var styles = stylesPart.Styles;
        var style = styles.Elements<Style>().FirstOrDefault(s =>
            s.StyleId?.Value == styleId || s.StyleName?.Val?.Value == styleId);
        if (style == null)
        {
            var isBuiltIn = styleId is "Normal" or "Heading1" or "Heading2" or "Heading3" or "Heading4"
                or "Heading5" or "Heading6" or "Heading7" or "Heading8" or "Heading9"
                or "Title" or "Subtitle" or "Quote" or "IntenseQuote" or "ListParagraph"
                or "NoSpacing" or "TOCHeading";
            style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
            if (!isBuiltIn) style.CustomStyle = true;
            style.AppendChild(new StyleName { Val = styleId });
            styles.AppendChild(style);
        }

        foreach (var (key, value) in properties)
        {
            // CONSISTENCY(run-prop-helper): rPr-style props (font/size/bold/
            // italic/color/highlight/underline/strike/caps/smallcaps/...)
            // delegate to ApplyRunFormatting which works on
            // StyleRunProperties via its OpenXmlCompositeElement base. This
            // also extends Style's previously narrow rPr surface (was 7
            // props) to cover the full ~23-prop ApplyRunFormatting set,
            // matching what Word actually accepts in style/rPr.
            if (ApplyRunFormatting(
                    style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties()),
                    key, value))
                continue;

            switch (key.ToLowerInvariant())
            {
                case "name":
                    var sn = style.StyleName ?? style.AppendChild(new StyleName());
                    sn.Val = value;
                    break;
                case "basedon":
                    var bo = style.BasedOn ?? style.AppendChild(new BasedOn());
                    bo.Val = value;
                    break;
                case "next":
                    var ns = style.NextParagraphStyle ?? style.AppendChild(new NextParagraphStyle());
                    ns.Val = value;
                    break;
                case "alignment":
                    var pPr = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    pPr.Justification = new Justification { Val = ParseJustification(value) };
                    break;
                case "spacebefore" or "spaceBefore":
                    var pPr2 = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    var sp2 = pPr2.SpacingBetweenLines ?? (pPr2.SpacingBetweenLines = new SpacingBetweenLines());
                    sp2.Before = SpacingConverter.ParseWordSpacing(value).ToString();
                    break;
                case "spaceafter" or "spaceAfter":
                    var pPr3 = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    var sp3 = pPr3.SpacingBetweenLines ?? (pPr3.SpacingBetweenLines = new SpacingBetweenLines());
                    sp3.After = SpacingConverter.ParseWordSpacing(value).ToString();
                    break;
                case "linespacing" or "lineSpacing":
                {
                    var pPr4 = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    var sp4 = pPr4.SpacingBetweenLines ?? (pPr4.SpacingBetweenLines = new SpacingBetweenLines());
                    var (twips, isMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
                    sp4.Line = twips.ToString();
                    sp4.LineRule = isMultiplier
                        ? new DocumentFormat.OpenXml.EnumValue<LineSpacingRuleValues>(LineSpacingRuleValues.Auto)
                        : new DocumentFormat.OpenXml.EnumValue<LineSpacingRuleValues>(LineSpacingRuleValues.Exact);
                    break;
                }
                case "contextualspacing" or "contextualSpacing":
                {
                    var pPrCs = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    if (IsTruthy(value))
                        pPrCs.ContextualSpacing ??= new ContextualSpacing();
                    else
                        pPrCs.ContextualSpacing = null;
                    break;
                }
                case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
                case "border.all" or "border" or "border.top" or "border.bottom" or "border.left" or "border.right":
                {
                    var pPrB = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    ApplyStyleParagraphBorders(pPrB, key, value);
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }
        styles.Save();
        return unsupported;
    }

    private List<string> SetWordOlePath(System.Text.RegularExpressions.Match wordOleSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var wOleIdx = int.Parse(wordOleSetMatch.Groups["idx"].Value);
        var wOleParent = wordOleSetMatch.Groups["parent"].Success && wordOleSetMatch.Groups["parent"].Value.Length > 0
            ? wordOleSetMatch.Groups["parent"].Value
            : "/body";
        var allOles = Query("ole")
            .Where(n => n.Path.StartsWith(wOleParent + "/", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (wOleIdx < 1 || wOleIdx > allOles.Count)
            throw new ArgumentException(
                $"OLE object {wOleIdx} not found at {wOleParent} (available: {allOles.Count}).");
        return Set(allOles[wOleIdx - 1].Path, properties);
    }

}
