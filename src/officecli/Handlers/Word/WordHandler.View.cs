// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Semantic Layer ====================

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        int lineNum = 0;
        int pIdx = 0, tblIdx = 0, eqIdx = 0;
        int emitted = 0;
        var bodyElements = GetBodyElements(body).ToList();
        int totalElements = bodyElements.Count;

        foreach (var element in bodyElements)
        {
            lineNum++;
            string path;

            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                eqIdx++;
                path = $"/body/oMathPara[{eqIdx}]";
            }
            else if (element is Paragraph)
            {
                pIdx++;
                path = $"/body/p[{pIdx}]";
            }
            else if (element is Table)
            {
                tblIdx++;
                path = $"/body/tbl[{tblIdx}]";
            }
            else if (IsStructuralElement(element))
            {
                path = $"/body/{element.LocalName}";
            }
            else
            {
                // Skip non-content elements
                continue;
            }

            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {emitted} rows, {totalElements} total, use --start/--end to view more)");
                break;
            }

            if (element is Paragraph para)
            {
                // Check if paragraph contains display equation (oMathPara)
                var oMathParaChild = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathParaChild != null)
                {
                    var mathText = FormulaParser.ToReadableText(oMathParaChild);
                    sb.AppendLine($"[{path}] [Equation] {mathText}");
                }
                else
                {
                    // Check for inline math
                    var mathElements = FindMathElements(para);
                    if (mathElements.Count > 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
                    {
                        var mathText = string.Concat(mathElements.Select(FormulaParser.ToReadableText));
                        sb.AppendLine($"[{path}] [Equation] {mathText}");
                    }
                    else if (mathElements.Count > 0)
                    {
                        var text = GetParagraphTextWithMath(para);
                        var listPrefix = GetListPrefix(para);
                        sb.AppendLine($"[{path}] {listPrefix}{text}");
                    }
                    else
                    {
                        var text = GetParagraphText(para);
                        var listPrefix = GetListPrefix(para);
                        sb.AppendLine($"[{path}] {listPrefix}{text}");
                    }
                }
            }
            else if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                var mathText = FormulaParser.ToReadableText(element);
                sb.AppendLine($"[{path}] [Equation] {mathText}");
            }
            else if (element is Table table)
            {
                sb.AppendLine($"[{path}] [Table: {table.Elements<TableRow>().Count()} rows]");
            }
            else if (IsStructuralElement(element))
            {
                sb.AppendLine($"[{path}] [{element.LocalName}]");
            }
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        int lineNum = 0;
        int pIdx = 0, tblIdx = 0, eqIdx = 0;
        int emitted = 0;
        var bodyElements = GetBodyElements(body).ToList();
        int totalElements = bodyElements.Count;

        foreach (var element in bodyElements)
        {
            lineNum++;
            string path;

            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                eqIdx++;
                path = $"/body/oMathPara[{eqIdx}]";
            }
            else if (element is Paragraph)
            {
                pIdx++;
                path = $"/body/p[{pIdx}]";
            }
            else if (element is Table)
            {
                tblIdx++;
                path = $"/body/tbl[{tblIdx}]";
            }
            else
            {
                path = $"/body/?[{lineNum}]";
            }

            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {emitted} rows, {totalElements} total, use --start/--end to view more)");
                break;
            }

            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                var latex = FormulaParser.ToLatex(element);
                sb.AppendLine($"[{path}] [Equation: \"{latex}\"] ← display");
            }
            else if (element is Paragraph para)
            {
                // Check if paragraph contains display equation (oMathPara)
                var oMathParaChild = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathParaChild != null)
                {
                    var latex = FormulaParser.ToLatex(oMathParaChild);
                    sb.AppendLine($"[{path}] [Equation: \"{latex}\"] ← display");
                    emitted++;
                    continue;
                }

                var styleName = GetStyleName(para);
                var runs = GetAllRuns(para);

                // Check for inline math
                var inlineMath = FindMathElements(para);
                if (inlineMath.Count > 0 && runs.Count == 0)
                {
                    var latex = string.Concat(inlineMath.Select(FormulaParser.ToLatex));
                    sb.AppendLine($"[{path}] [Equation: \"{latex}\"] ← {styleName} | inline");
                    emitted++;
                    continue;
                }

                if (runs.Count == 0 && inlineMath.Count == 0)
                {
                    sb.AppendLine($"[{path}] [] <- {styleName} | empty paragraph");
                    emitted++;
                    continue;
                }

                var listPrefix = GetListPrefix(para);

                foreach (var run in runs)
                {
                    // Check if run contains an image
                    var drawing = run.GetFirstChild<Drawing>();
                    if (drawing != null)
                    {
                        var imgInfo = GetDrawingInfo(drawing);
                        sb.AppendLine($"[{path}] {listPrefix}[Image: {imgInfo}] ← {styleName}");
                        continue;
                    }

                    var text = GetRunText(run);
                    var fmt = GetRunFormatDescription(run, para);
                    var warn = "";

                    sb.AppendLine($"[{path}] {listPrefix}「{text}」 ← {styleName} | {fmt}{warn}");
                }

                // Show inline math elements
                foreach (var math in inlineMath)
                {
                    var latex = FormulaParser.ToLatex(math);
                    sb.AppendLine($"[{path}] {listPrefix}[Equation: \"{latex}\"] ← {styleName} | inline");
                }
            }
            else if (element is Table table)
            {
                var rows = table.Elements<TableRow>().Count();
                var colCount = table.Elements<TableRow>().FirstOrDefault()
                    ?.Elements<TableCell>().Count() ?? 0;
                sb.AppendLine($"[{path}] [Table: {rows}×{colCount}]");
            }
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        // Document info
        var paragraphs = GetBodyElements(body).OfType<Paragraph>().ToList();
        var tables = GetBodyElements(body).OfType<Table>().ToList();
        var imageCount = body.Descendants<Drawing>().Count();
        var equationCount = body.Descendants().Count(e => e.LocalName == "oMathPara" || e is M.Paragraph);
        var statsLine = $"File: {Path.GetFileName(_filePath)} | {paragraphs.Count} paragraphs | {tables.Count} tables | {imageCount} images";
        if (equationCount > 0) statsLine += $" | {equationCount} equations";
        sb.AppendLine(statsLine);

        // Watermark
        var watermark = FindWatermark();
        if (watermark != null)
            sb.AppendLine($"Watermark: \"{watermark}\"");

        // Headers
        var headers = GetHeaderTexts();
        foreach (var h in headers)
            sb.AppendLine($"Header: \"{h}\"");

        // Footers
        var footers = GetFooterTexts();
        foreach (var f in footers)
            sb.AppendLine($"Footer: \"{f}\"");

        sb.AppendLine();

        // Heading structure
        int lineNum = 0;
        foreach (var para in paragraphs)
        {
            lineNum++;
            var styleName = GetStyleName(para);
            var text = GetParagraphText(para);

            if (styleName.Contains("Heading") || styleName.Contains("标题")
                || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase)
                || styleName == "Title" || styleName == "Subtitle")
            {
                var level = GetHeadingLevel(styleName);
                var indent = level <= 1 ? "" : new string(' ', (level - 1) * 2);
                var prefix = level == 0 ? "■" : "├──";
                sb.AppendLine($"{indent}{prefix} [{lineNum}] \"{text}\" ({styleName})");
            }
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsStats()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        var paragraphs = GetBodyElements(body).OfType<Paragraph>().ToList();

        // Style counts
        var styleCounts = new Dictionary<string, int>();
        var fontCounts = new Dictionary<string, int>();
        var sizeCounts = new Dictionary<string, int>();
        int emptyParagraphs = 0;
        int doubleSpaces = 0;
        int totalChars = 0;

        foreach (var para in paragraphs)
        {
            var style = GetStyleName(para);
            styleCounts[style] = styleCounts.GetValueOrDefault(style) + 1;

            var runs = GetAllRuns(para);
            if (runs.Count == 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
            {
                emptyParagraphs++;
                continue;
            }

            foreach (var run in runs)
            {
                var text = GetRunText(run);
                totalChars += text.Length;

                if (text.Contains("  "))
                    doubleSpaces++;

                var resolved = ResolveEffectiveRunProperties(run, para);
                var font = GetFontFromProperties(resolved) ?? "(default)";
                fontCounts[font] = fontCounts.GetValueOrDefault(font) + 1;

                var size = GetSizeFromProperties(resolved) ?? "(default)";
                sizeCounts[size] = sizeCounts.GetValueOrDefault(size) + 1;
            }
        }

        sb.AppendLine($"Paragraphs: {paragraphs.Count} | Total Characters: {totalChars}");
        sb.AppendLine();

        sb.AppendLine("Style Distribution:");
        foreach (var (style, count) in styleCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {style}: {count}");

        sb.AppendLine();
        sb.AppendLine("Font Usage:");
        foreach (var (font, count) in fontCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {font}: {count}");

        sb.AppendLine();
        sb.AppendLine("Font Size Usage:");
        foreach (var (size, count) in sizeCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {size}: {count}");

        sb.AppendLine();
        sb.AppendLine($"Empty Paragraphs: {emptyParagraphs}");
        sb.AppendLine($"Consecutive Spaces: {doubleSpaces}");

        return sb.ToString().TrimEnd();
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return issues;

        int issueNum = 0;
        int lineNum = -1;

        foreach (var para in GetBodyElements(body).OfType<Paragraph>())
        {
            lineNum++;
            var styleName = GetStyleName(para);
            var runs = GetAllRuns(para);

            // Empty paragraph
            if (runs.Count == 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"S{++issueNum}",
                    Type = IssueType.Structure,
                    Severity = IssueSeverity.Warning,
                    Path = $"/body/p[{lineNum + 1}]",
                    Message = "Empty paragraph"
                });
            }

            // Paragraph format checks
            var pProps = para.ParagraphProperties;
            if (pProps != null && IsNormalStyle(styleName))
            {
                var indent = pProps.Indentation;
                if (indent?.FirstLine == null || indent.FirstLine.Value == "0")
                {
                    // Only flag if there's actual text
                    if (runs.Any(r => !string.IsNullOrWhiteSpace(GetRunText(r))))
                    {
                        issues.Add(new DocumentIssue
                        {
                            Id = $"F{++issueNum}",
                            Type = IssueType.Format,
                            Severity = IssueSeverity.Warning,
                            Path = $"/body/p[{lineNum + 1}]",
                            Message = "Body paragraph missing first-line indent",
                            Suggestion = "Set first-line indent to 2 characters"
                        });
                    }
                }
            }

            int runIdx = 0;
            foreach (var run in runs)
            {
                var text = GetRunText(run);

                // Double spaces
                if (text.Contains("  "))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Warning,
                        Path = $"/body/p[{lineNum + 1}]/r[{runIdx + 1}]",
                        Message = "Consecutive spaces",
                        Context = text,
                        Suggestion = "Merge into a single space"
                    });
                }

                // Duplicate punctuation
                if (System.Text.RegularExpressions.Regex.IsMatch(text, @"[，。！？、；：]{2,}"))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Warning,
                        Path = $"/body/p[{lineNum + 1}]/r[{runIdx + 1}]",
                        Message = "Duplicate punctuation",
                        Context = text
                    });
                }

                // Mixed Chinese/English punctuation
                if (HasMixedPunctuation(text))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Info,
                        Path = $"/body/p[{lineNum + 1}]/r[{runIdx + 1}]",
                        Message = "Mixed CJK/Latin punctuation",
                        Context = text
                    });
                }

                runIdx++;
            }

            if (limit.HasValue && issues.Count >= limit.Value) break;
        }

        // Filter by type
        if (issueType != null)
        {
            var type = issueType.ToLowerInvariant() switch
            {
                "format" or "f" => IssueType.Format,
                "content" or "c" => IssueType.Content,
                "structure" or "s" => IssueType.Structure,
                _ => (IssueType?)null
            };
            if (type.HasValue)
                issues = issues.Where(i => i.Type == type.Value).ToList();
        }

        return limit.HasValue ? issues.Take(limit.Value).ToList() : issues;
    }
}
