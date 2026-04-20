// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

// Real-world docx files from legacy editors (WPS, older Word, third-party tools)
// sometimes carry attribute values that violate the OOXML schema — e.g.
// `<w:b w:val="yes"/>` or `<w:jc w:val="bogus"/>`. Native Word is lenient,
// but DocumentFormat.OpenXml throws FormatException the moment any reader
// accesses `.Val.Value` on the typed property. Since the crash is lazy, it
// surfaces unpredictably deep inside rendering code (HtmlPreview.Css,
// styling, etc.) rather than at open time.
//
// This sanitizer walks raw XML attributes (no typed conversion) right after
// Open, repairs or strips the offending values, and lets every downstream
// reader operate normally. Corresponds to KNOWN_ISSUES §9.
internal static class WordStrictAttributeSanitizer
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly HashSet<string> OnOffValid = new(StringComparer.OrdinalIgnoreCase)
        { "true", "false", "on", "off", "0", "1" };

    // Elements whose `w:val` attribute is an OnOff. Invalid values → strip val
    // (the element's mere presence means "true", matching Word's behavior).
    private static readonly HashSet<string> OnOffElements = new(StringComparer.Ordinal)
    {
        "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "vanish", "specVanish", "webHidden", "noProof",
        "emboss", "imprint", "outline", "shadow", "snapToGrid",
        "contextualSpacing", "kinsoku", "overflowPunct", "topLinePunct",
        "autoSpaceDE", "autoSpaceDN", "wordWrap",
        "suppressAutoHyphens", "suppressLineNumbers", "suppressOverlap",
        "widowControl", "keepNext", "keepLines", "pageBreakBefore",
        "hidden", "cantSplit", "tblHeader",
        "bookFoldPrinting", "bookFoldRevPrinting",
        "evenAndOddHeaders", "titlePg",
    };

    // Elements whose `w:val` is an enum. Invalid values → strip the whole
    // element (default behavior of the parent kicks in).
    private static readonly Dictionary<string, HashSet<string>> EnumElements = new(StringComparer.Ordinal)
    {
        ["jc"] = new(StringComparer.Ordinal)
        {
            "left", "center", "right", "both", "start", "end",
            "distribute", "mediumKashida", "lowKashida", "highKashida",
            "thaiDistribute", "numTab",
        },
        ["vAlign"] = new(StringComparer.Ordinal) { "top", "center", "bottom", "both" },
        ["textDirection"] = new(StringComparer.Ordinal)
            { "lrTb", "tbRl", "btLr", "lrTbV", "tbRlV", "tbLrV", "rl", "lr" },
    };

    public static void Sanitize(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main == null) return;

        // Wrap each part access: `main.Document` getter throws if the file
        // isn't actually WordML (e.g. xlsx opened as docx). Existing tests
        // document that WordHandler silently tolerates wrong-format opens,
        // so we mirror that by skipping parts we can't load.
        TrySanitize(() => main.Document);
        TrySanitize(() => main.StyleDefinitionsPart?.Styles);
        TrySanitize(() => main.NumberingDefinitionsPart?.Numbering);
        TrySanitize(() => main.FootnotesPart?.Footnotes);
        TrySanitize(() => main.EndnotesPart?.Endnotes);
        TrySanitize(() => main.DocumentSettingsPart?.Settings);
        foreach (var h in main.HeaderParts) TrySanitize(() => h.Header);
        foreach (var f in main.FooterParts) TrySanitize(() => f.Footer);
    }

    private static void TrySanitize(Func<OpenXmlPartRootElement?> getRoot)
    {
        OpenXmlPartRootElement? root;
        try { root = getRoot(); }
        catch { return; }
        if (root != null) SanitizePart(root);
    }

    private static void SanitizePart(OpenXmlPartRootElement root)
    {
        // Snapshot first — we may mutate (remove elements) during sanitize.
        var nodes = root.Descendants<OpenXmlElement>().ToList();
        var toRemove = new List<OpenXmlElement>();

        foreach (var elem in nodes)
        {
            if (elem.NamespaceUri != W) continue;
            var name = elem.LocalName;

            if (OnOffElements.Contains(name))
            {
                var raw = ReadValAttribute(elem);
                if (raw != null && !OnOffValid.Contains(raw))
                {
                    // Strip val — bare element = true, matching Word's
                    // lenient handling of `<w:b w:val="yes"/>`.
                    elem.RemoveAttribute("val", W);
                }
            }
            else if (EnumElements.TryGetValue(name, out var valid))
            {
                var raw = ReadValAttribute(elem);
                if (raw != null && !valid.Contains(raw))
                {
                    toRemove.Add(elem);
                }
            }
        }

        foreach (var elem in toRemove)
        {
            elem.Parent?.RemoveChild(elem);
        }
    }

    private static string? ReadValAttribute(OpenXmlElement elem)
    {
        foreach (var a in elem.GetAttributes())
        {
            if (a.LocalName == "val" && (a.NamespaceUri == W || a.NamespaceUri == ""))
                return a.Value;
        }
        return null;
    }
}
