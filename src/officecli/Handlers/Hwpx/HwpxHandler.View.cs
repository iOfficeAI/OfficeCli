using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== View Layer ====================

    public string ViewAsText(int? startLine = null, int? endLine = null,
                              int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            var rawText = ExtractParagraphText(para);
            var text = HwpxKorean.Normalize(rawText);

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                int remaining = CountRemainingParagraphs(lineNum);
                if (remaining > 0)
                    sb.AppendLine($"... ({remaining} more lines)");
                break;
            }

            sb.AppendLine($"{lineNum}. {text}");
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null,
                                   int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;
            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                int remaining = CountRemainingParagraphs(lineNum);
                if (remaining > 0)
                    sb.AppendLine($"... ({remaining} more lines)");
                break;
            }

            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";
            var styleInfo = GetParagraphStyleInfo(para);
            var runs = ExtractAnnotatedRuns(para);
            var text = string.Join("", runs.Select(r => r.Text));
            text = HwpxKorean.Normalize(text);

            // Build annotation prefix
            var annotations = new List<string>();
            if (!string.IsNullOrEmpty(styleInfo.HeadingLevel))
                annotations.Add($"h{styleInfo.HeadingLevel}");
            if (styleInfo.Alignment != "LEFT")
                annotations.Add(styleInfo.Alignment.ToLowerInvariant());

            var prefix = annotations.Count > 0 ? $"[{string.Join(",", annotations)}] " : "";
            sb.AppendLine($"{lineNum}. {path} {prefix}{text}");
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var styleInfo = GetParagraphStyleInfo(para);
            if (string.IsNullOrEmpty(styleInfo.HeadingLevel)) continue;

            var level = int.Parse(styleInfo.HeadingLevel);
            var indent = new string(' ', (level - 1) * 2);
            var text = HwpxKorean.Normalize(ExtractParagraphText(para));
            var preview = text.Length > 80 ? text[..80] + "…" : text;
            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";

            sb.AppendLine($"{indent}h{level}: {preview} ({path})");
        }

        return sb.Length > 0 ? sb.ToString().TrimEnd() : "(no headings found)";
    }

    public string ViewAsStats()
    {
        int totalParas = 0, totalTables = 0, totalChars = 0, totalWords = 0;
        int totalImages = 0;

        foreach (var sec in _doc.Sections)
        {
            totalParas += sec.Paragraphs.Count;
            totalTables += sec.Tables.Count;
            totalImages += sec.Root.Descendants(HwpxNs.Hp + "img").Count();

            foreach (var p in sec.Paragraphs)
            {
                var text = HwpxKorean.Normalize(ExtractParagraphText(p));
                totalChars += text.Length;
                totalWords += CountWords(text);
            }
        }

        var sb = new StringBuilder();
        sb.AppendLine($"Sections:   {_doc.Sections.Count}");
        sb.AppendLine($"Paragraphs: {totalParas}");
        sb.AppendLine($"Tables:     {totalTables}");
        sb.AppendLine($"Images:     {totalImages}");
        sb.AppendLine($"Characters: {totalChars}");
        sb.AppendLine($"Words:      {totalWords}");

        // Page info — iterate ALL sections for aggregate stats; use first secPr for page size reference
        foreach (var sec in _doc.Sections)
        {
            var secPr = sec.Root.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault();
            var pagePr = secPr?.Element(HwpxNs.Hp + "pagePr");
            if (pagePr != null)
            {
                var width = (int?)pagePr.Attribute("width") ?? 0;
                var height = (int?)pagePr.Attribute("height") ?? 0;
                sb.AppendLine($"Page size:  {FormatHwpUnit(width)} × {FormatHwpUnit(height)}");
                break; // Report first section's page size; add per-section loop if needed
            }
        }

        return sb.ToString().TrimEnd();
    }

    public JsonNode ViewAsStatsJson()
    {
        int totalParas = 0, totalTables = 0, totalChars = 0, totalWords = 0;
        int totalImages = 0;

        foreach (var sec in _doc.Sections)
        {
            totalParas += sec.Paragraphs.Count;
            totalTables += sec.Tables.Count;
            totalImages += sec.Root.Descendants(HwpxNs.Hp + "img").Count();

            foreach (var p in sec.Paragraphs)
            {
                var text = HwpxKorean.Normalize(ExtractParagraphText(p));
                totalChars += text.Length;
                totalWords += CountWords(text);
            }
        }

        return new JsonObject
        {
            ["sections"] = _doc.Sections.Count,
            ["paragraphs"] = totalParas,
            ["tables"] = totalTables,
            ["images"] = totalImages,
            ["characters"] = totalChars,
            ["words"] = totalWords,
        };
    }

    public JsonNode ViewAsOutlineJson()
    {
        var items = new JsonArray();

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var styleInfo = GetParagraphStyleInfo(para);
            if (string.IsNullOrEmpty(styleInfo.HeadingLevel)) continue;

            var level = int.Parse(styleInfo.HeadingLevel);
            var text = HwpxKorean.Normalize(ExtractParagraphText(para));
            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";

            items.Add(new JsonObject
            {
                ["level"] = level,
                ["text"] = text,
                ["path"] = path,
            });
        }

        return items;
    }

    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null,
                                    int? maxLines = null, HashSet<string>? cols = null)
    {
        var lines = new JsonArray();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;
            if (maxLines.HasValue && emitted >= maxLines.Value) break;

            var text = HwpxKorean.Normalize(ExtractParagraphText(para));
            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";

            lines.Add(new JsonObject
            {
                ["line"] = lineNum,
                ["path"] = path,
                ["text"] = text,
            });
            emitted++;
        }

        return new JsonObject
        {
            ["lines"] = lines,
            ["totalLines"] = _doc.AllParagraphs().Count(),
        };
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueId = 0;

        // Check for empty paragraphs
        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var text = ExtractParagraphText(para);
            if (string.IsNullOrWhiteSpace(text))
            {
                // Skip — empty paragraphs are normal spacing
                continue;
            }

            // Check for PUA characters (corruption indicator)
            if (text.Any(c => c >= '\uE000' && c <= '\uF8FF'))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"HWPX-{++issueId:D3}",
                    Type = IssueType.Content,
                    Severity = IssueSeverity.Warning,
                    Path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]",
                    Message = "Paragraph contains Private Use Area characters",
                    Context = text[..Math.Min(text.Length, 50)]
                });
            }
        }

        // Check for tables with inconsistent column counts
        foreach (var (section, tbl, tblIdx) in _doc.AllTables())
        {
            var rows = tbl.Elements(HwpxNs.Hp + "tr").ToList();
            if (rows.Count == 0) continue;

            var expectedCols = (int?)tbl.Attribute("colCnt") ?? -1;
            foreach (var (row, rowIdx) in rows.Select((r, i) => (r, i)))
            {
                // Sum colSpan values (handles merged cells); GetCellAddr is defined in this partial class
                var colSpanSum = row.Elements(HwpxNs.Hp + "tc")
                    .Sum(tc => (int?)GetCellAddr(tc).ColSpan ?? 1);
                if (expectedCols >= 0 && colSpanSum != expectedCols)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"HWPX-{++issueId:D3}",
                        Type = IssueType.Structure,
                        Severity = IssueSeverity.Error,
                        Path = $"/section[{section.Index + 1}]/tbl[{tblIdx + 1}]/tr[{rowIdx + 1}]",
                        Message = $"Row colSpan sum {colSpanSum} != expected {expectedCols}",
                        Context = null
                    });
                }
            }
        }

        // Check for missing header.xml
        if (_doc.Header == null)
        {
            issues.Add(new DocumentIssue
            {
                Id = $"HWPX-{++issueId:D3}",
                Type = IssueType.Structure,
                Severity = IssueSeverity.Warning,
                Path = "/",
                Message = "Document missing header.xml (style definitions unavailable)",
                Context = null
            });
        }

        // Filter by type
        if (issueType != null)
        {
            var filterType = Enum.Parse<IssueType>(issueType, ignoreCase: true);
            issues = issues.Where(i => i.Type == filterType).ToList();
        }

        // Apply limit
        if (limit.HasValue)
            issues = issues.Take(limit.Value).ToList();

        return issues;
    }

    /// <summary>Extract all text from a paragraph's hp:run/hp:t elements.</summary>
    private static string ExtractParagraphText(XElement para)
    {
        var runs = para.Elements(HwpxNs.Hp + "run");
        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            foreach (var t in run.Elements(HwpxNs.Hp + "t"))
            {
                sb.Append(t.Value);
            }
            // Handle line breaks
            if (run.Element(HwpxNs.Hp + "lineBreak") != null)
                sb.Append('\n');
            if (run.Element(HwpxNs.Hp + "tab") != null)
                sb.Append('\t');
        }
        return sb.ToString();
    }

    /// <summary>Extract runs with formatting annotations.</summary>
    private static List<(string Text, Dictionary<string, string> Format)> ExtractAnnotatedRuns(XElement para)
    {
        var result = new List<(string, Dictionary<string, string>)>();
        foreach (var run in para.Elements(HwpxNs.Hp + "run"))
        {
            var text = string.Join("", run.Elements(HwpxNs.Hp + "t").Select(t => t.Value));
            if (string.IsNullOrEmpty(text)) continue;

            var format = new Dictionary<string, string>();
            var charPrIdRef = run.Attribute("charPrIDRef")?.Value;
            if (charPrIdRef != null)
                format["charPrIDRef"] = charPrIdRef;

            result.Add((text, format));
        }
        return result;
    }

    /// <summary>Get paragraph style info from attributes and header.xml lookup.</summary>
    private (string? HeadingLevel, string Alignment) GetParagraphStyleInfo(XElement para)
    {
        var styleIdRef = para.Attribute("styleIDRef")?.Value;
        var paraPrIdRef = para.Attribute("paraPrIDRef")?.Value;

        string? headingLevel = null;
        string alignment = "LEFT";

        // Look up style in header.xml
        if (_doc.Header != null && styleIdRef != null)
        {
            var style = _doc.Header.Root!.Descendants(HwpxNs.Hh + "style")
                .FirstOrDefault(s => s.Attribute("id")?.Value == styleIdRef);
            if (style != null)
            {
                var name = style.Attribute("name")?.Value ?? "";
                // Korean heading styles: "개요 1", "개요 2", etc.
                var headingMatch = System.Text.RegularExpressions.Regex.Match(name, @"개요\s*(\d+)");
                if (headingMatch.Success)
                    headingLevel = headingMatch.Groups[1].Value;
                // English heading styles
                var engMatch = System.Text.RegularExpressions.Regex.Match(name, @"(?i)heading\s*(\d+)");
                if (engMatch.Success)
                    headingLevel = engMatch.Groups[1].Value;
            }
        }

        // Look up paragraph properties for alignment
        if (_doc.Header != null && paraPrIdRef != null)
        {
            var paraPr = _doc.Header.Root!.Descendants(HwpxNs.Hh + "paraPr")
                .FirstOrDefault(p => p.Attribute("id")?.Value == paraPrIdRef);
            if (paraPr != null)
            {
                // Real HWPX: alignment is a child element <hh:align horizontal="LEFT"/>
                var alignEl = paraPr.Element(HwpxNs.Hh + "align");
                alignment = alignEl?.Attribute("horizontal")?.Value ?? "LEFT";
            }
        }

        return (headingLevel, alignment);
    }

    private int CountRemainingParagraphs(int currentLine)
    {
        int total = _doc.AllParagraphs().Count();
        return Math.Max(0, total - currentLine);
    }

    private static int CountWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0;
        // Korean: each syllable cluster counts as a word boundary
        // Simple heuristic: split on whitespace, count non-empty
        return text.Split(Array.Empty<char>(), StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private static string FormatHwpUnit(int hwpUnit)
    {
        var mm = hwpUnit / 283.46;
        return $"{mm:0.#}mm";
    }
}
