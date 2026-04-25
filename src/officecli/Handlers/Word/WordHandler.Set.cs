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

public partial class WordHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        // Batch Set: if path looks like a selector (not starting with /), Query → Set each
        if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
        {
            var targets = Query(path);
            if (targets.Count == 0)
                throw new ArgumentException($"No elements matched selector: {path}");
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        // Unified find: if 'find' key is present (at any path level), route to ProcessFind
        if (properties.TryGetValue("find", out var findText))
        {
            var replace = properties.TryGetValue("replace", out var r) ? r : null;
            // Separate run-level format properties from paragraph-level properties
            var formatProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var paraProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (key, value) in properties)
            {
                var k = key.ToLowerInvariant();
                if (k is "find" or "replace" or "scope" or "regex") continue;
                // Paragraph-level properties go to paraProps
                if (k is "style" or "alignment" or "align" or "firstlineindent" or "leftindent" or "indentleft"
                    or "indent" or "rightindent" or "indentright" or "hangingindent" or "spacebefore"
                    or "spaceafter" or "linespacing" or "keepnext" or "keeplines" or "pagebreakbefore"
                    or "widowcontrol" or "liststyle" or "start" or "text" or "formula"
                    or "contextualspacing")
                    paraProps[key] = value;
                else
                    formatProps[key] = value;
            }

            if (replace == null && formatProps.Count == 0 && paraProps.Count == 0)
                throw new ArgumentException("'find' requires either 'replace' and/or format properties (e.g. bold, highlight, color).");

            // CONSISTENCY(find-regex): canonical site for the `regex=true` → `r"..."`
            // raw-string normalization. `mark` and the other handlers' Set paths all
            // copy this pattern verbatim. To change the find/regex protocol,
            // grep "CONSISTENCY(find-regex)" and update every site project-wide;
            // do not diverge in a single handler.
            if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findText.StartsWith("r\"") && !findText.StartsWith("r'"))
                findText = $"r\"{findText}\"";

            var effectivePath = (path is "" or "/") ? "/body" : path;
            var matchCount = ProcessFind(effectivePath, findText, replace, formatProps.Count > 0 ? formatProps : new Dictionary<string, string>());
            LastFindMatchCount = matchCount;

            // Apply paragraph-level properties to the matched paragraphs
            if (paraProps.Count > 0)
            {
                var paragraphs = ResolveParagraphsForFind(effectivePath);
                foreach (var para in paragraphs)
                {
                    var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                    foreach (var (key, value) in paraProps)
                        ApplyParagraphLevelProperty(pProps, key, value);
                }
            }

            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Document-level properties
        if (path == "/" || path == "" || path.Equals("/body", StringComparison.OrdinalIgnoreCase))
        {
            SetDocumentProperties(properties, unsupported);
            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Handle /settings path — route to SetDocumentProperties which calls TrySetDocSetting
        if (path.Equals("/settings", StringComparison.OrdinalIgnoreCase))
        {
            SetDocumentProperties(properties, unsupported);
            EnsureSettings().Save();
            return unsupported;
        }

        // Handle /watermark path
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase))
        {
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

        // FormField paths: /formfield[N] or /formfield[name]
        // Routed BEFORE ParsePath because the generic predicate validator
        // only accepts positive-integer / last() / [@attr=v] predicates and
        // would reject the documented /formfield[name] form.
        var ffSetMatchEarly = System.Text.RegularExpressions.Regex.Match(path, @"^/formfield\[(\w+)\]$");
        if (ffSetMatchEarly.Success)
        {
            var allFormFields = FindFormFields();
            var indexOrName = ffSetMatchEarly.Groups[1].Value;
            (FieldInfo Field, FormFieldData FfData) target;
            if (int.TryParse(indexOrName, out var ffIdx))
            {
                if (ffIdx < 1 || ffIdx > allFormFields.Count)
                    throw new ArgumentException($"FormField {ffIdx} not found (total: {allFormFields.Count})");
                target = allFormFields[ffIdx - 1];
            }
            else
            {
                target = allFormFields.FirstOrDefault(ff =>
                    ff.FfData.GetFirstChild<FormFieldName>()?.Val?.Value == indexOrName);
                if (target.Field == null)
                    throw new ArgumentException($"FormField '{indexOrName}' not found");
            }
            return SetFormField(target, properties);
        }

        // Handle header/footer paths
        var hfParts = ParsePath(path);
        if (hfParts.Count >= 1)
        {
            var firstName = hfParts[0].Name.ToLowerInvariant();
            if ((firstName == "header" || firstName == "footer") && hfParts.Count == 1)
            {
                SetHeaderFooter(firstName, (hfParts[0].Index ?? 1) - 1, properties, unsupported);
                return unsupported;
            }
        }

        // Chart axis-by-role sub-path: /chart[N]/axis[@role=ROLE].
        var chartAxisSetMatch = System.Text.RegularExpressions.Regex.Match(path,
            @"^/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success)
        {
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

        // Chart paths: /chart[N] or /chart[N]/series[K]
        var chartMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartMatch.Success)
        {
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

        // Field paths: /field[N]
        var fieldSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/field\[(\d+)\]$");
        if (fieldSetMatch.Success)
        {
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

        // TOC paths: /toc[N]
        var tocMatch = System.Text.RegularExpressions.Regex.Match(path, @"/toc\[(\d+)\]$");
        if (tocMatch.Success)
        {
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

        // Footnote paths: /footnote[N] or .../footnote[N]
        var fnSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"/footnote\[(\d+)\]$");
        if (fnSetMatch.Success)
        {
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

        // Endnote paths: /endnote[N] or .../endnote[N]
        var enSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"/endnote\[(\d+)\]$");
        if (enSetMatch.Success)
        {
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

        // Section paths: /section[N] or /body/sectPr[N] (canonical form returned by Get/Query)
        var secSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^(?:/section\[(\d+)\]|/body/sectPr(?:\[(\d+)\])?)$");
        if (secSetMatch.Success)
        {
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

        // Style paths: /styles/StyleId
        var styleSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/styles/(.+)$");
        if (styleSetMatch.Success)
        {
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

        // CONSISTENCY(ole-shorthand-set): mirror the /body/ole[N] shorthand
        // already supported in Get (WordHandler.Query.cs) and Remove
        // (WordHandler.Mutations.cs). Without this intercept, Set falls through
        // to NavigateToElement which hits "No ole found at /body" because OLE
        // lives inside a run, not as a direct child of the body.
        var wordOleSetMatch = System.Text.RegularExpressions.Regex.Match(
            path,
            @"^(?<parent>/body|/header\[\d+\]|/footer\[\d+\])?/(?:ole|object|embed)\[(?<idx>\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (wordOleSetMatch.Success)
        {
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

        var parts = ParsePath(path);
        var element = NavigateToElement(parts, out var ctx);
        if (element == null)
            throw new ArgumentException($"Path not found: {path}" + (ctx != null ? $". {ctx}" : ""));

        // Clone element for rollback on failure (atomic: no partial modifications)
        var elementBackup = element.CloneNode(true);
        try
        {
        return SetElement(element, properties);
        }
        catch
        {
            // Rollback: restore element to pre-modification state
            element.Parent?.ReplaceChild(elementBackup, element);
            throw;
        }
    }

    private List<string> SetElement(OpenXmlElement element, Dictionary<string, string> properties)
    {
        if (element is BookmarkStart bk) return SetElementBookmark(bk, properties);
        if (element is SdtBlock || element is SdtRun) return SetElementSdt(element, properties);
        if (element is Run run) return SetElementRun(run, properties);
        if (element is Hyperlink hl) return SetElementHyperlink(hl, properties);
        if (element is M.Paragraph mPara) return SetElementMPara(mPara, properties);
        if (element is Paragraph para) return SetElementParagraph(para, properties);
        if (element is TableCell cell) return SetElementTableCell(cell, properties);
        if (element is TableRow row) return SetElementTableRow(row, properties);
        if (element is Table tbl) return SetElementTable(tbl, properties);
        return new List<string>();
    }

    private void SetHeaderFooter(string kind, int index, Dictionary<string, string> properties, List<string> unsupported)
    {
        var mainPart = _doc.MainDocumentPart!;
        OpenXmlCompositeElement? container;
        OpenXmlPart partRef;

        if (kind == "header")
        {
            var part = mainPart.HeaderParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Header not found: /header[{index + 1}]");
            container = part.Header;
            partRef = part;
        }
        else
        {
            var part = mainPart.FooterParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Footer not found: /footer[{index + 1}]");
            container = part.Footer;
            partRef = part;
        }

        if (container == null)
            throw new ArgumentException($"{kind} content not found at index {index + 1}");

        var firstPara = container.Elements<Paragraph>().FirstOrDefault();
        if (firstPara == null)
        {
            firstPara = new Paragraph();
            container.AppendChild(firstPara);
        }
        var pProps = firstPara.ParagraphProperties ?? firstPara.PrependChild(new ParagraphProperties());

        foreach (var (key, value) in properties)
        {
            var k = key.ToLowerInvariant();
            if (ApplyParagraphLevelProperty(pProps, key, value))
            {
                // handled by paragraph-level helper
            }
            else switch (k)
            {
                case "text":
                {
                    RunProperties? existingRProps = null;
                    var existingRun = firstPara.Elements<Run>().FirstOrDefault();
                    if (existingRun?.RunProperties != null)
                        existingRProps = (RunProperties)existingRun.RunProperties.CloneNode(true);
                    foreach (var r in firstPara.Elements<Run>().ToList()) r.Remove();
                    var newRun = new Run();
                    if (existingRProps != null)
                        newRun.AppendChild(existingRProps);
                    newRun.AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                    firstPara.AppendChild(newRun);
                    break;
                }
                case "size" or "font" or "bold" or "italic" or "color" or "highlight" or "underline" or "strike":
                    // Apply run-level formatting to all runs in the container
                    foreach (var run in container.Descendants<Run>())
                        ApplyRunFormatting(EnsureRunProperties(run), key, value);
                    // Also update paragraph mark run properties so new runs inherit formatting
                    var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    ApplyRunFormatting(markRPr, key, value);
                    break;
                case "type":
                {
                    // Mutate the HeaderReference/FooterReference Type attribute
                    // pointing at this part. Read side (WordHandler.Query.cs:660-666,
                    // 717-723) only inspects body-level SectionProperties, so the
                    // write side stays scoped to the same set for round-trip parity.
                    var newType = value.ToLowerInvariant() switch
                    {
                        "first" => HeaderFooterValues.First,
                        "even" => HeaderFooterValues.Even,
                        "default" => HeaderFooterValues.Default,
                        _ => throw new ArgumentException(
                            $"Invalid {kind} type: '{value}'. Valid values: default, first, even.")
                    };
                    var partRid = mainPart.GetIdOfPart(partRef);
                    var body = mainPart.Document?.Body
                        ?? throw new InvalidOperationException("Document body not found");
                    bool found = false;
                    foreach (var sp in body.Elements<SectionProperties>())
                    {
                        if (kind == "header")
                        {
                            var ownRef = sp.Elements<HeaderReference>().FirstOrDefault(r => r.Id?.Value == partRid);
                            if (ownRef == null) continue;
                            if (ownRef.Type?.Value == newType) { found = true; continue; }
                            if (sp.Elements<HeaderReference>().Any(r => r != ownRef && r.Type?.Value == newType))
                                throw new ArgumentException(
                                    $"Header of type '{value}' already exists in this section.");
                            ownRef.Type = newType;
                            found = true;
                        }
                        else
                        {
                            var ownRef = sp.Elements<FooterReference>().FirstOrDefault(r => r.Id?.Value == partRid);
                            if (ownRef == null) continue;
                            if (ownRef.Type?.Value == newType) { found = true; continue; }
                            if (sp.Elements<FooterReference>().Any(r => r != ownRef && r.Type?.Value == newType))
                                throw new ArgumentException(
                                    $"Footer of type '{value}' already exists in this section.");
                            ownRef.Type = newType;
                            found = true;
                        }
                        // Mirrors AddHeader: Title-page header requires <w:titlePg/> on the section.
                        if (newType == HeaderFooterValues.First && sp.GetFirstChild<TitlePage>() == null)
                            sp.AddChild(new TitlePage(), throwOnError: false);
                    }
                    if (!found) unsupported.Add(key);
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        if (kind == "header")
            mainPart.HeaderParts.ElementAt(index).Header?.Save();
        else
            mainPart.FooterParts.ElementAt(index).Footer?.Save();
    }

    // Border style format: "style" or "style;size" or "style;size;color" or "style;size;color;space"
    // Styles: none, single, thick, double, dotted, dashed, dotDash, dotDotDash, triple,
    //         thinThickSmallGap, thickThinSmallGap, thinThickThinSmallGap,
    //         thinThickMediumGap, thickThinMediumGap, thinThickThinMediumGap,
    //         thinThickLargeGap, thickThinLargeGap, thinThickThinLargeGap, wave, doubleWave, threeDEmboss, threeDEngrave
    /// <summary>Insert StyleParagraphProperties before StyleRunProperties to maintain OOXML schema order.</summary>
    private static StyleParagraphProperties EnsureStyleParagraphProperties(Style style)
    {
        var pPr = new StyleParagraphProperties();
        var rPr = style.StyleRunProperties;
        if (rPr != null)
            style.InsertBefore(pPr, rPr);
        else
            style.AppendChild(pPr);
        return pPr;
    }

    private static BorderValues ParseBorderStyle(string style) => style.ToLowerInvariant() switch
    {
        "none" => BorderValues.None,
        "nil" => BorderValues.Nil,
        "single" or "thin" => BorderValues.Single,
        "thick" or "medium" => BorderValues.Thick,
        "double" => BorderValues.Double,
        "dotted" => BorderValues.Dotted,
        "dashed" => BorderValues.Dashed,
        "dotdash" => BorderValues.DotDash,
        "dotdotdash" => BorderValues.DotDotDash,
        "triple" => BorderValues.Triple,
        "thinthicksmallgap" => BorderValues.ThinThickSmallGap,
        "thickthinsmallgap" => BorderValues.ThickThinSmallGap,
        "thinthickthinsmallgap" => BorderValues.ThinThickThinSmallGap,
        "thinthickmediumgap" => BorderValues.ThinThickMediumGap,
        "thickthinmediumgap" => BorderValues.ThickThinMediumGap,
        "thinthickthinmediumgap" => BorderValues.ThinThickThinMediumGap,
        "thinthicklargegap" => BorderValues.ThinThickLargeGap,
        "thickthinlargegap" => BorderValues.ThickThinLargeGap,
        "thinthickthinlargegap" => BorderValues.ThinThickThinLargeGap,
        "wave" => BorderValues.Wave,
        "doublewave" => BorderValues.DoubleWave,
        "threedembed" or "3demboss" => BorderValues.ThreeDEmboss,
        "threedengrave" or "3dengrave" => BorderValues.ThreeDEngrave,
        _ => throw new ArgumentException($"Invalid border style: '{style}'. Valid values: single, thick, double, dotted, dashed, none, triple, wave, etc.")
    };

    private static (BorderValues style, uint size, string? color, uint space) ParseBorderValue(string value)
    {
        var parts = value.Split(';');
        var style = ParseBorderStyle(parts[0]);
        uint size;
        if (parts.Length > 1)
        {
            if (!uint.TryParse(parts[1], out size))
                throw new ArgumentException($"Invalid border size '{parts[1]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
        }
        else
            size = style == BorderValues.Nil ? 0u : style == BorderValues.Thick ? 12u : 4u;
        string? color = parts.Length > 2 ? SanitizeHex(parts[2]) : null;
        uint space = 0u;
        if (parts.Length > 3 && !uint.TryParse(parts[3], out space))
            throw new ArgumentException($"Invalid border space '{parts[3]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
        return (style, size, color, space);
    }

    private static T MakeBorder<T>(BorderValues style, uint size, string? color, uint space) where T : BorderType, new()
    {
        var b = new T { Val = style, Size = size, Space = space };
        if (color != null) b.Color = color;
        return b;
    }

    /// <summary>
    /// Apply a paragraph-level property. Returns true if handled, false if not recognized.
    /// Handles: style, alignment, indent, spacing, keepNext, keepLines, pageBreakBefore, widowControl, shading, pbdr.
    /// </summary>
    private static bool ApplyParagraphLevelProperty(ParagraphProperties pProps, string key, string? value)
    {
        if (value is null) return false;
        switch (key.ToLowerInvariant())
        {
            case "style":
                pProps.ParagraphStyleId = new ParagraphStyleId { Val = value };
                return true;
            case "alignment" or "align":
                pProps.Justification = new Justification { Val = ParseJustification(value) };
                return true;
            case "firstlineindent":
                var indent = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                // Lenient input: accept "2cm", "0.5in", "18pt", or bare twips.
                indent.FirstLine = SpacingConverter.ParseWordSpacing(value).ToString();
                indent.Hanging = null;
                return true;
            case "leftindent" or "indentleft" or "indent":
                var indentL = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentL.Left = ParseHelpers.SafeParseUint(value, "leftindent").ToString();
                return true;
            case "rightindent" or "indentright":
                var indentR = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentR.Right = ParseHelpers.SafeParseUint(value, "rightindent").ToString();
                return true;
            case "hangingindent" or "hanging":
                var indentH = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentH.Hanging = ParseHelpers.SafeParseUint(value, "hangingindent").ToString();
                indentH.FirstLine = null;
                return true;
            case "keepnext" or "keepwithnext":
                if (IsTruthy(value)) pProps.KeepNext ??= new KeepNext();
                else pProps.KeepNext = null;
                return true;
            case "keeplines" or "keeptogether":
                if (IsTruthy(value)) pProps.KeepLines ??= new KeepLines();
                else pProps.KeepLines = null;
                return true;
            case "pagebreakbefore":
                if (IsTruthy(value)) pProps.PageBreakBefore ??= new PageBreakBefore();
                else pProps.PageBreakBefore = null;
                return true;
            case "widowcontrol" or "widoworphan":
                if (IsTruthy(value)) pProps.WidowControl ??= new WidowControl();
                else pProps.WidowControl = new WidowControl { Val = false };
                return true;
            case "contextualspacing" or "contextualSpacing":
                if (IsTruthy(value)) pProps.ContextualSpacing ??= new ContextualSpacing();
                else pProps.ContextualSpacing = null;
                return true;
            case "shading" or "shd":
                pProps.Shading = ParseShadingValue(value);
                return true;
            case "spacebefore":
                var spacingBefore = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingBefore.Before = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "spaceafter":
                var spacingAfter = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingAfter.After = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "linespacing":
                var spacingLine = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                var (lsTwips, lsIsMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
                spacingLine.Line = lsTwips.ToString();
                spacingLine.LineRule = lsIsMultiplier ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact;
                return true;
            case "numId" or "numid":
                var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                numPr.NumberingId = new NumberingId { Val = ParseHelpers.SafeParseInt(value, "numId") };
                return true;
            case "numLevel" or "numlevel" or "ilvl":
                var numPr2 = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                numPr2.NumberingLevelReference = new NumberingLevelReference { Val = ParseHelpers.SafeParseInt(value, "numLevel") };
                return true;
            case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
            case "border.all" or "border" or "border.top" or "border.bottom" or "border.left" or "border.right":
                ApplyParagraphBorders(pProps, key, value);
                return true;
            default:
                return false;
        }
    }


    private static void ApplyParagraphBorders(ParagraphProperties pProps, string key, string value)
    {
        var borders = pProps.ParagraphBorders;
        if (borders == null)
        {
            borders = new ParagraphBorders();
            pProps.ParagraphBorders = borders; // typed setter maintains CT_PPr schema order
        }
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "pbdr.all" or "pbdr" or "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.top" or "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "pbdr.bottom" or "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "pbdr.left" or "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "pbdr.right" or "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "pbdr.between":
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.bar":
                borders.BarBorder = MakeBorder<BarBorder>(style, size, color, space);
                break;
        }
    }

    private static void ApplyStyleParagraphBorders(StyleParagraphProperties spPr, string key, string value)
    {
        var borders = spPr.GetFirstChild<ParagraphBorders>();
        if (borders == null)
        {
            borders = new ParagraphBorders();
            // StyleParagraphProperties is also OneSequence — use SetElement pattern
            // ParagraphBorders element order index is after Indentation and before Shading
            var afterRef = (OpenXmlElement?)spPr.GetFirstChild<Indentation>()
                ?? (OpenXmlElement?)spPr.GetFirstChild<SpacingBetweenLines>()
                ?? (OpenXmlElement?)spPr.GetFirstChild<Justification>();
            if (afterRef != null)
                spPr.InsertAfter(borders, afterRef);
            else
                spPr.PrependChild(borders);
        }
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "pbdr.all" or "pbdr" or "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.top" or "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "pbdr.bottom" or "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "pbdr.left" or "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "pbdr.right" or "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "pbdr.between":
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.bar":
                borders.BarBorder = MakeBorder<BarBorder>(style, size, color, space);
                break;
        }
    }

    private static void ApplyTableBorders(TableProperties tblPr, string key, string value)
    {
        var borders = tblPr.TableBorders ?? tblPr.AppendChild(new TableBorders());
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.InsideHorizontalBorder = MakeBorder<InsideHorizontalBorder>(style, size, color, space);
                borders.InsideVerticalBorder = MakeBorder<InsideVerticalBorder>(style, size, color, space);
                break;
            case "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "border.insideh" or "border.horizontal":
                borders.InsideHorizontalBorder = MakeBorder<InsideHorizontalBorder>(style, size, color, space);
                break;
            case "border.insidev" or "border.vertical":
                borders.InsideVerticalBorder = MakeBorder<InsideVerticalBorder>(style, size, color, space);
                break;
        }
    }

    private static void ApplyCellBorders(TableCellProperties tcPr, string key, string value)
    {
        var borders = tcPr.TableCellBorders ?? tcPr.AppendChild(new TableCellBorders());
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "border.tl2br":
                borders.TopLeftToBottomRightCellBorder = MakeBorder<TopLeftToBottomRightCellBorder>(style, size, color, space);
                break;
            case "border.tr2bl":
                borders.TopRightToBottomLeftCellBorder = MakeBorder<TopRightToBottomLeftCellBorder>(style, size, color, space);
                break;
        }
    }

    /// <summary>
    /// Apply gradient fill to a Word table cell using mc:AlternativeContent with w14:gradFill.
    /// Fallback is a solid shading with the start color.
    /// </summary>
    private static void ApplyCellGradient(TableCellProperties tcPr, string startColor, string endColor, int angleDeg)
    {
        // Sanitize colors: strip 8-char RRGGBBAA to 6-char RGB (w14:srgbClr requires 6 chars)
        var (startRgb, _) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(startColor);
        var (endRgb, _) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(endColor);

        // Remove existing shading/gradient
        RemoveCellGradient(tcPr);
        tcPr.Shading?.Remove();

        // Set fallback solid fill
        tcPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = startRgb };

        // Build w14:gradFill XML via raw OpenXml
        var w14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        // Convert angle to OOXML 60000ths of a degree
        var angleOoxml = angleDeg * 60000;

        var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);
        acElement.InnerXml = $@"<mc:Choice xmlns:mc=""{mcNs}"" xmlns:w14=""{w14Ns}"" Requires=""w14"">
    <w14:gradFill>
      <w14:gsLst>
        <w14:gs w14:pos=""0"">
          <w14:srgbClr w14:val=""{startRgb}""/>
        </w14:gs>
        <w14:gs w14:pos=""100000"">
          <w14:srgbClr w14:val=""{endRgb}""/>
        </w14:gs>
      </w14:gsLst>
      <w14:lin w14:ang=""{angleOoxml}"" w14:scaled=""1""/>
    </w14:gradFill>
  </mc:Choice>";

        tcPr.AppendChild(acElement);
    }

    /// <summary>
    /// Remove any existing gradient mc:AlternateContent from table cell properties.
    /// </summary>
    private static void RemoveCellGradient(TableCellProperties tcPr)
    {
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        var existing = tcPr.ChildElements
            .Where(e => e.LocalName == "AlternateContent" && e.NamespaceUri == mcNs)
            .ToList();
        foreach (var e in existing) e.Remove();
    }

    /// <summary>
    /// Parse twips from a string with optional unit suffix: "1.5cm", "0.5in", "36pt", or raw twips.
    /// 1 inch = 1440 twips, 1 cm = 567 twips, 1 pt = 20 twips.
    /// </summary>
    private static TablePositionProperties EnsureTablePositionProperties(TableProperties tblPr)
    {
        var tpp = tblPr.GetFirstChild<TablePositionProperties>();
        if (tpp == null)
        {
            tpp = new TablePositionProperties
            {
                VerticalAnchor = VerticalAnchorValues.Page,
                HorizontalAnchor = HorizontalAnchorValues.Page
            };
            // CT_TblPr schema order: tblStyle → tblpPr → tblOverlap → ...
            var tblStyle = tblPr.GetFirstChild<TableStyle>();
            if (tblStyle != null)
                tblStyle.InsertAfterSelf(tpp);
            else
                tblPr.PrependChild(tpp);
        }
        return tpp;
    }

    internal static uint ParseTwips(string value)
    {
        value = value.Trim();
        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (cm)");
            return (uint)Math.Round(num * 1440.0 / 2.54);
        }
        if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (in)");
            return (uint)Math.Round(num * 1440);
        }
        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (pt)");
            return (uint)Math.Round(num * 20);
        }
        return ParseHelpers.SafeParseUint(value, "twips");
    }
}
