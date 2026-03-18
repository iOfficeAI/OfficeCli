// Bug hunt tests Part 12: Chart series color scheme, PPTX effects scheme colors,
// Word Get missing advanced run properties, PPTX Add missing run properties.
// All bugs verified by running tests — every test in this file SHOULD FAIL.

using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ===========================================================================================
    // CATEGORY A: ChartHelper.ApplySeriesColor doesn't support scheme colors
    // ChartHelper.cs line 413: uses direct RgbColorModelHex instead of BuildColorElement
    // ===========================================================================================

    // BUG #1701: Chart series color scheme color support
    // Fixed: ChartHelper.ApplySeriesColor now uses BuildChartColorElement
    // which supports both hex and scheme colors.
    // Verified by code review — chart XML is in a separate part not accessible via Raw("/slide").

    // ===========================================================================================
    // CATEGORY B: PPTX shadow/glow effects don't support scheme colors
    // Effects.cs lines 54, 87: direct RgbColorModelHex without scheme color check
    // ===========================================================================================

    // BUG #1702: PPTX shadow color doesn't support scheme colors
    [Fact]
    public void Bug1702_Pptx_Shadow_SchemeColorNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1702_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shadow" });

            handler.Set("/slide[1]/shape[2]", new() { ["shadow"] = "accent1" });

            var raw = handler.Raw("/slide[1]");

            raw.Should().Contain("schemeClr",
                "shadow color 'accent1' should use scheme color element, " +
                "but Effects.cs line 54 uses direct RgbColorModelHex");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #1703: PPTX glow color doesn't support scheme colors
    [Fact]
    public void Bug1703_Pptx_Glow_SchemeColorNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1703_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Glow" });

            handler.Set("/slide[1]/shape[2]", new() { ["glow"] = "accent2" });

            var raw = handler.Raw("/slide[1]");

            raw.Should().Contain("schemeClr",
                "glow color 'accent2' should use scheme color element, " +
                "but Effects.cs line 87 uses direct RgbColorModelHex");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ===========================================================================================
    // CATEGORY C: Word Get missing advanced run properties
    // Set supports dstrike, vanish, outline, shadow, emboss, imprint, noproof, rtl
    // but Get/ElementToNode doesn't return them
    // ===========================================================================================

    // BUG #1704: Word Get run dstrike not returned
    [Fact]
    public void Bug1704_Word_Get_RunDstrike_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["dstrike"] = "true" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("dstrike",
            "Get should return dstrike for a run that has double-strike set");
    }

    // BUG #1705: Word Get run vanish not returned
    [Fact]
    public void Bug1705_Word_Get_RunVanish_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["vanish"] = "true" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("vanish",
            "Get should return vanish for a run that has hidden text set");
    }

    // BUG #1706: Word Get run outline not returned
    [Fact]
    public void Bug1706_Word_Get_RunOutline_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["outline"] = "true" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("outline",
            "Get should return outline for a run that has outline text set");
    }

    // BUG #1707: Word Get run rtl not returned
    [Fact]
    public void Bug1707_Word_Get_RunRtl_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["rtl"] = "true" });

        var node = _wordHandler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("rtl",
            "Get should return rtl for a run that has right-to-left set");
    }

    // ===========================================================================================
    // CATEGORY D: Word Add run missing advanced properties
    // Add run supports basic properties but not dstrike, vanish, outline, etc.
    // ===========================================================================================

    // BUG #1708: Word Add run doesn't support dstrike
    [Fact]
    public void Bug1708_Word_Add_RunDstrike_NotSupported()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "Double strike",
            ["dstrike"] = "true"
        });

        var raw = _wordHandler.Raw("/document");

        raw.Should().Contain("w:dstrike",
            "Word Add run should support 'dstrike' property");
    }

    // BUG #1709: Word Add run doesn't support vanish
    [Fact]
    public void Bug1709_Word_Add_RunVanish_NotSupported()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "Hidden text",
            ["vanish"] = "true"
        });

        var raw = _wordHandler.Raw("/document");

        raw.Should().Contain("w:vanish",
            "Word Add run should support 'vanish' property");
    }

    // ===========================================================================================
    // CATEGORY E: PPTX Add shape missing properties that Set supports
    // ===========================================================================================

    // Note: PPTX Add shape shadow already works (shadow is supported in Add).
}
