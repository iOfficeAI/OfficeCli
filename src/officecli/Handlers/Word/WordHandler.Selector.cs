// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Selector ====================

    private record SelectorPart(string? Element, Dictionary<string, string> Attributes, string? ContainsText, SelectorPart? ChildSelector);

    private static SelectorPart ParseSelector(string selector)
    {
        // Support: element[attr=value] > child[attr=value]
        var childParts = selector.Split('>').Select(s => s.Trim()).ToArray();

        SelectorPart? childSelector = null;
        if (childParts.Length > 1)
        {
            childSelector = ParseSingleSelector(childParts[1]);
        }

        var main = ParseSingleSelector(childParts[0]);
        return main with { ChildSelector = childSelector };
    }

    private static SelectorPart ParseSingleSelector(string selector)
    {
        var attrs = new Dictionary<string, string>();
        string? element = null;
        string? containsText = null;

        // Extract element name (before any [ or : modifier)
        var firstMod = selector.Length;
        var bracketIdx = selector.IndexOf('[');
        if (bracketIdx >= 0 && bracketIdx < firstMod) firstMod = bracketIdx;
        var colonIdx = selector.IndexOf(':');
        if (colonIdx >= 0 && colonIdx < firstMod) firstMod = colonIdx;

        element = selector[..firstMod].Trim();
        if (string.IsNullOrEmpty(element)) element = null;

        // Parse [attr=value] attributes
        if (bracketIdx >= 0)
        {
            var attrPart = selector[bracketIdx..];
            foreach (System.Text.RegularExpressions.Match m in
                System.Text.RegularExpressions.Regex.Matches(attrPart, @"\[(\w+)(!?=)([^\]]+)\]"))
            {
                var key = m.Groups[1].Value;
                var op = m.Groups[2].Value;
                var val = m.Groups[3].Value.Trim('\'', '"');
                attrs[key] = (op == "!=" ? "!" : "") + val;
            }
        }

        // Parse :contains("text") pseudo-selector
        if (selector.Contains(":contains("))
        {
            var idx = selector.IndexOf(":contains(");
            var endIdx = selector.IndexOf(')', idx + 10);
            if (endIdx >= 0)
                containsText = selector[(idx + 10)..endIdx].Trim('\'', '"');
        }

        // Parse :empty pseudo-selector
        if (selector.Contains(":empty"))
        {
            attrs["__empty"] = "true";
        }

        // Parse :no-alt pseudo-selector
        if (selector.Contains(":no-alt"))
        {
            attrs["__no-alt"] = "true";
        }

        return new SelectorPart(element, attrs, containsText, null);
    }

    private bool MatchesSelector(Paragraph para, SelectorPart selector, int lineNum)
    {
        // If selector targets runs (has child selector), only match parent paragraph
        if (selector.ChildSelector != null)
        {
            // Check paragraph-level attributes
            if (selector.Element != null && selector.Element != "p" && selector.Element != "paragraph")
                return false;
            return MatchesParagraphAttrs(para, selector.Attributes);
        }

        if (selector.Element != null && selector.Element != "p" && selector.Element != "paragraph")
            return false;

        if (!MatchesParagraphAttrs(para, selector.Attributes))
            return false;

        if (selector.Attributes.ContainsKey("__empty"))
        {
            return string.IsNullOrWhiteSpace(GetParagraphText(para));
        }

        if (selector.ContainsText != null)
        {
            return GetParagraphText(para).Contains(selector.ContainsText);
        }

        return true;
    }

    private bool MatchesParagraphAttrs(Paragraph para, Dictionary<string, string> attrs)
    {
        foreach (var (key, rawVal) in attrs)
        {
            if (key == "__empty") continue;
            bool negate = rawVal.StartsWith("!");
            var val = negate ? rawVal[1..] : rawVal;

            string? actual = key.ToLowerInvariant() switch
            {
                "style" => GetStyleName(para),
                "alignment" => para.ParagraphProperties?.Justification?.Val?.HasValue == true
                    ? para.ParagraphProperties.Justification.Val.Value.ToString() : null,
                "firstlineindent" => para.ParagraphProperties?.Indentation?.FirstLine?.Value,
                "numid" => para.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.HasValue == true
                    ? para.ParagraphProperties.NumberingProperties.NumberingId.Val.Value.ToString() : null,
                "numlevel" or "ilvl" => para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.HasValue == true
                    ? para.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value.ToString() : null,
                "liststyle" => GetParagraphListStyle(para),
                _ => GenericXmlQuery.GetAttributeValue(para, key)
                     ?? (para.ParagraphProperties != null ? GenericXmlQuery.GetAttributeValue(para.ParagraphProperties, key) : null)
            };

            // For style, also match against styleId (e.g., "Heading1" vs display name "heading 1")
            bool matches;
            if (key.Equals("style", StringComparison.OrdinalIgnoreCase))
            {
                var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase)
                       || string.Equals(styleId, val, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase);
            }
            if (negate ? matches : !matches) return false;
        }
        return true;
    }

    private static bool MatchesRunSelector(Run run, Paragraph parent, SelectorPart selector)
    {
        if (selector.Element != null && selector.Element != "r" && selector.Element != "run")
            return false;

        foreach (var (key, rawVal) in selector.Attributes)
        {
            bool negate = rawVal.StartsWith("!");
            var val = negate ? rawVal[1..] : rawVal;

            string? actual = key.ToLowerInvariant() switch
            {
                "font" => GetRunFont(run),
                "size" => GetRunFontSize(run),
                "bold" => run.RunProperties?.Bold != null ? "true" : "false",
                "italic" => run.RunProperties?.Italic != null ? "true" : "false",
                _ => GenericXmlQuery.GetAttributeValue(run, key)
                     ?? (run.RunProperties != null ? GenericXmlQuery.GetAttributeValue(run.RunProperties, key) : null)
            };

            bool matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase);
            if (negate ? matches : !matches) return false;
        }

        if (selector.ContainsText != null)
        {
            return GetRunText(run).Contains(selector.ContainsText);
        }

        return true;
    }

    private string GetHeaderRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var headerPart = _doc.MainDocumentPart?.HeaderParts.ElementAtOrDefault(idx);
        return headerPart?.Header?.OuterXml ?? $"(header[{idx}] not found)";
    }

    private string GetFooterRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var footerPart = _doc.MainDocumentPart?.FooterParts.ElementAtOrDefault(idx);
        return footerPart?.Footer?.OuterXml ?? $"(footer[{idx}] not found)";
    }

    private string GetChartRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var chartPart = _doc.MainDocumentPart?.ChartParts.ElementAtOrDefault(idx);
        return chartPart?.ChartSpace?.OuterXml ?? $"(chart[{idx}] not found)";
    }
}
