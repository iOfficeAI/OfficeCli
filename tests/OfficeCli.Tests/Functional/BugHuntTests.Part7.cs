// Bug hunt tests Part 7: Word handler bugs
// Covers: shading # not stripped, paragraph/cell property propagation gaps,
// Get/ElementToNode missing run properties, Add/Set consistency issues.
// All bugs verified by running tests — every test in this file SHOULD FAIL.

using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ===========================================================================================
    // CATEGORY A: Word shading color # prefix not stripped in XML
    // All shading code assigns fill/color values without TrimStart('#'),
    // unlike the color property which correctly does value.TrimStart('#').ToUpperInvariant().
    // ===========================================================================================

    // BUG #701: WordHandler.Add.cs line 105: shd.Fill = shdParts[0] — no TrimStart('#')
    [Fact]
    public void Bug701_Word_Add_ParagraphShading_HashInXml()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Shaded paragraph",
            ["shd"] = "#FFFF00"
        });

        var raw = _wordHandler.Raw("/document");
        raw.Should().Contain("w:fill=\"FFFF00\"",
            "shading fill in OOXML should be bare hex without # prefix, " +
            "but WordHandler.Add.cs line 105 assigns raw value without TrimStart('#')");
    }

    // BUG #702: WordHandler.Set.cs line 489: shd.Fill = shdParts[0] — no TrimStart('#')
    [Fact]
    public void Bug702_Word_Set_RunShading_HashInXml()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Add("/body/p[1]", "run", null, new() { ["text"] = "Run text" });

        _wordHandler.Set("/body/p[1]/r[1]", new() { ["shd"] = "#00FF00" });

        var raw = _wordHandler.Raw("/document");
        raw.Should().Contain("w:fill=\"00FF00\"",
            "run shading fill should be bare hex in XML, " +
            "but WordHandler.Set.cs line 489 assigns raw value without TrimStart('#')");
    }

    // BUG #703: WordHandler.Set.cs line 638: shdP.Fill = shdPartsP[0] — no TrimStart('#')
    [Fact]
    public void Bug703_Word_Set_ParagraphShading_HashInXml()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });

        _wordHandler.Set("/body/p[1]", new() { ["shd"] = "#FF0000" });

        var raw = _wordHandler.Raw("/document");
        raw.Should().Contain("w:fill=\"FF0000\"",
            "paragraph shading fill should be bare hex in XML, " +
            "but WordHandler.Set.cs line 638 assigns raw value without TrimStart('#')");
    }

    // BUG #704: WordHandler.Set.cs line 784: shd.Fill = shdParts[0] — no TrimStart('#')
    [Fact]
    public void Bug704_Word_Set_CellShading_HashInXml()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["shd"] = "#0000FF" });

        var raw = _wordHandler.Raw("/document");
        raw.Should().Contain("w:fill=\"0000FF\"",
            "cell shading fill should be bare hex in XML, " +
            "but WordHandler.Set.cs line 784 assigns raw value without TrimStart('#')");
    }

    // BUG #705: WordHandler.Set.cs lines 494-495: fill and color parts not stripped
    [Fact]
    public void Bug705_Word_Set_RunShading_MultiPartColor_HashInXml()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Add("/body/p[1]", "run", null, new() { ["text"] = "Run" });

        _wordHandler.Set("/body/p[1]/r[1]", new()
        {
            ["shd"] = "clear;#FFFF00;#000000"
        });

        var raw = _wordHandler.Raw("/document");
        raw.Should().Contain("w:fill=\"FFFF00\"",
            "shading fill in multi-part format should strip # prefix");
        raw.Should().Contain("w:color=\"000000\"",
            "shading color in multi-part format should strip # prefix");
    }

    // BUG #706: WordHandler.Add.cs line 189: shd.Fill = shdParts[0] — no TrimStart('#')
    [Fact]
    public void Bug706_Word_Add_RunShading_HashInXml()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "Shaded run",
            ["shd"] = "#AABBCC"
        });

        var raw = _wordHandler.Raw("/document");
        raw.Should().Contain("w:fill=\"AABBCC\"",
            "run shading in Add should store bare hex in XML, " +
            "but WordHandler.Add.cs line 189 assigns raw value without TrimStart('#')");
    }

    // ===========================================================================================
    // CATEGORY B: Word Set paragraph-level property propagation gaps
    // Set at paragraph level supports bold/italic/color/font/size (lines 675-702)
    // but NOT highlight, underline, strike — they fall through to default → unsupported.
    // Add at paragraph level DOES support these properties.
    // ===========================================================================================

    // BUG #901: highlight not propagated at paragraph level
    [Fact]
    public void Bug901_Word_Set_ParagraphHighlight_NotPropagatedToRuns()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Paragraph with runs"
        });

        var unsupported = _wordHandler.Set("/body/p[1]", new() { ["highlight"] = "yellow" });

        unsupported.Should().NotContain("highlight",
            "highlight should be supported at paragraph level like bold/italic/color, " +
            "but WordHandler.Set.cs paragraph switch only handles size/font/bold/italic/color");
    }

    // BUG #902: underline not propagated at paragraph level
    [Fact]
    public void Bug902_Word_Set_ParagraphUnderline_NotPropagatedToRuns()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Underline test"
        });

        var unsupported = _wordHandler.Set("/body/p[1]", new() { ["underline"] = "single" });

        unsupported.Should().NotContain("underline",
            "underline should be supported at paragraph level like bold/italic/color, " +
            "but it's missing from the paragraph switch in WordHandler.Set.cs");
    }

    // BUG #903: strike not propagated at paragraph level
    [Fact]
    public void Bug903_Word_Set_ParagraphStrike_NotPropagatedToRuns()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Strike test"
        });

        var unsupported = _wordHandler.Set("/body/p[1]", new() { ["strike"] = "true" });

        unsupported.Should().NotContain("strike",
            "strike should be supported at paragraph level like bold/italic/color, " +
            "but it's missing from the paragraph switch in WordHandler.Set.cs");
    }

    // BUG #904: highlight not propagated at cell level
    [Fact]
    public void Bug904_Word_Set_CellHighlight_NotPropagatedToRuns()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Cell text" });

        var unsupported = _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["highlight"] = "yellow"
        });

        unsupported.Should().NotContain("highlight",
            "highlight should propagate to cell runs like bold/italic/color do");
    }

    // BUG #905: underline not propagated at cell level
    [Fact]
    public void Bug905_Word_Set_CellUnderline_NotPropagatedToRuns()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Cell text" });

        var unsupported = _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["underline"] = "single"
        });

        unsupported.Should().NotContain("underline",
            "underline should propagate to cell runs like bold/italic/color do");
    }

    // BUG #906: strike not propagated at cell level
    [Fact]
    public void Bug906_Word_Set_CellStrike_NotPropagatedToRuns()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Cell text" });

        var unsupported = _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["strike"] = "true"
        });

        unsupported.Should().NotContain("strike",
            "strike should propagate to cell runs like bold/italic/color do");
    }

    // ===========================================================================================
    // CATEGORY C: Word Get/ElementToNode missing run properties
    // Navigation.cs lines 253-267: ElementToNode for Run only extracts
    // font, size, bold, italic, superscript, subscript — NOT color, underline, etc.
    // ===========================================================================================

    // BUG #1001: color not returned by Get for runs
    [Fact]
    public void Bug1001_Word_Get_RunColor_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["color"] = "FF0000" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("color",
            "Get should return color for a run that has color set, " +
            "but ElementToNode (Navigation.cs:253-267) doesn't extract color");
    }

    // BUG #1002: underline not returned by Get for runs
    [Fact]
    public void Bug1002_Word_Get_RunUnderline_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["underline"] = "single" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("underline",
            "Get should return underline for a run that has underline set, " +
            "but ElementToNode doesn't extract underline");
    }

    // BUG #1003: strike not returned by Get for runs
    [Fact]
    public void Bug1003_Word_Get_RunStrike_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["strike"] = "true" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("strike",
            "Get should return strike for a run that has strike set, " +
            "but ElementToNode doesn't extract strike");
    }

    // BUG #1004: highlight not returned by Get for runs
    [Fact]
    public void Bug1004_Word_Get_RunHighlight_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["highlight"] = "yellow" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("highlight",
            "Get should return highlight for a run that has highlight set, " +
            "but ElementToNode doesn't extract highlight");
    }

    // BUG #1005: caps not returned by Get for runs
    [Fact]
    public void Bug1005_Word_Get_RunCaps_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["caps"] = "true" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("caps",
            "Get should return caps for a run that has caps set, " +
            "but ElementToNode doesn't extract caps");
    }

    // BUG #1006: smallCaps not returned by Get for runs
    [Fact]
    public void Bug1006_Word_Get_RunSmallCaps_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["smallcaps"] = "true" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("smallCaps",
            "Get should return smallCaps for a run that has smallcaps set, " +
            "but ElementToNode doesn't extract smallCaps");
    }

    // ===========================================================================================
    // CATEGORY D: Word header/footer missing properties vs paragraph-level
    // SetHeaderFooter (lines 928-1022) only supports: text, font, size, bold, italic, color, alignment
    // Missing: underline, strike, highlight (should propagate to runs like bold/italic/color)
    // ===========================================================================================

    // BUG #1010: header/footer Set doesn't support underline
    [Fact]
    public void Bug1010_Word_Set_HeaderUnderline_Unsupported()
    {
        _wordHandler.Add("/body", "header", null, new() { ["text"] = "Header text" });

        var unsupported = _wordHandler.Set("/header[1]", new() { ["underline"] = "single" });

        unsupported.Should().NotContain("underline",
            "header Set should support underline like it supports bold/italic/color");
    }

    // BUG #1011: header/footer Set doesn't support strike
    [Fact]
    public void Bug1011_Word_Set_HeaderStrike_Unsupported()
    {
        _wordHandler.Add("/body", "header", null, new() { ["text"] = "Header text" });

        var unsupported = _wordHandler.Set("/header[1]", new() { ["strike"] = "true" });

        unsupported.Should().NotContain("strike",
            "header Set should support strike like it supports bold/italic/color");
    }

    // BUG #1012: header/footer Set doesn't support highlight
    [Fact]
    public void Bug1012_Word_Set_HeaderHighlight_Unsupported()
    {
        _wordHandler.Add("/body", "header", null, new() { ["text"] = "Header text" });

        var unsupported = _wordHandler.Set("/header[1]", new() { ["highlight"] = "yellow" });

        unsupported.Should().NotContain("highlight",
            "header Set should support highlight like it supports bold/italic/color");
    }

    // ===========================================================================================
    // CATEGORY F: Word Add textbox shading — # in XML
    // WordHandler.Add.cs line 317: shd.Fill = shdParts[0] — no TrimStart('#')
    // ===========================================================================================

    // BUG #1013: Word Add textbox run shading — # in XML
    [Fact]
    public void Bug1013_Word_Add_TextboxRunShading_HashInXml()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Text with textbox"
        });

        // Add a run (like a textbox run) with shading
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "Highlighted text",
            ["shd"] = "#DDEEFF"
        });

        var raw = _wordHandler.Raw("/document");
        raw.Should().Contain("w:fill=\"DDEEFF\"",
            "textbox run shading should store bare hex, " +
            "but WordHandler.Add.cs line 317 assigns raw value without TrimStart('#')");
    }

    // ===========================================================================================
    // CATEGORY G: Word paragraph alignment — "both" not documented but "justify" is
    // The Set code maps "justify" → JustificationValues.Both, which is correct.
    // But what about directly using "both"? It maps to JustificationValues.Left (default).
    // ===========================================================================================

    // ===========================================================================================
    // CATEGORY G: Word paragraph Get missing spacing/indent format keys
    // These can be Set on a paragraph but aren't returned by Get/ElementToNode
    // ===========================================================================================

    // BUG #1014: paragraph spacebefore not returned by Get
    [Fact]
    public void Bug1014_Word_Get_ParagraphSpaceBefore_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Spaced text" });
        _wordHandler.Set("/body/p[1]", new() { ["spacebefore"] = "240" });

        var node = _wordHandler.Get("/body/p[1]");

        node.Format.Should().ContainKey("spacebefore",
            "Get should return spacebefore for a paragraph that has it set");
    }

    // BUG #1015: paragraph linespacing not returned by Get
    [Fact]
    public void Bug1015_Word_Get_ParagraphLineSpacing_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Line spaced" });
        _wordHandler.Set("/body/p[1]", new() { ["linespacing"] = "360" });

        var node = _wordHandler.Get("/body/p[1]");

        node.Format.Should().ContainKey("linespacing",
            "Get should return linespacing for a paragraph that has it set");
    }

    // BUG #1016: paragraph leftIndent not returned by Get
    [Fact]
    public void Bug1016_Word_Get_ParagraphLeftIndent_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Indented" });
        _wordHandler.Set("/body/p[1]", new() { ["leftindent"] = "720" });

        var node = _wordHandler.Get("/body/p[1]");

        node.Format.Should().ContainKey("leftIndent",
            "Get should return leftIndent for a paragraph that has it set");
    }
}
