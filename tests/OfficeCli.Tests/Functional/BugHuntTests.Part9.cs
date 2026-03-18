// Bug hunt tests Part 9: Word Add color # missing, hyperlink color ignored,
// Excel/PPTX missing properties, URI handling bugs.
// All bugs verified by running tests — every test in this file SHOULD FAIL.

using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ===========================================================================================
    // CATEGORY A: Word Add style color missing TrimStart('#')
    // Add.cs line 1112: Val = sColor.ToUpperInvariant() — no TrimStart('#')
    // Compare with Set line 312 which correctly does TrimStart('#').ToUpperInvariant()
    // ===========================================================================================

    // BUG #1403: Word Add style color doesn't strip #
    [Fact]
    public void Bug1403_Word_Add_StyleColor_HashNotStripped()
    {
        _wordHandler.Add("/styles", "style", null, new()
        {
            ["name"] = "TestStyle",
            ["id"] = "TestStyleId",
            ["type"] = "paragraph",
            ["color"] = "#0000FF"
        });

        var raw = _wordHandler.Raw("/styles");

        raw.Should().NotContain("#0000FF",
            "style color should strip # prefix, " +
            "but Add.cs line 1112 does ToUpperInvariant() without TrimStart('#')");
    }

    // ===========================================================================================
    // CATEGORY B: Word Add hyperlink color hardcoded, ignores user input
    // Add.cs line 786: hlRProps.Color = new Color { Val = "0563C1" }
    // User-provided "color" property is never checked
    // ===========================================================================================

    // BUG #1404: Word Add hyperlink ignores user color
    [Fact]
    public void Bug1404_Word_Add_Hyperlink_ColorIgnored()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });

        _wordHandler.Add("/body/p[1]", "hyperlink", null, new()
        {
            ["url"] = "https://example.com",
            ["text"] = "Click here",
            ["color"] = "FF0000"
        });

        var raw = _wordHandler.Raw("/document");

        // BUG: color is hardcoded to 0563C1, user's FF0000 is ignored
        // Add.cs line 786 sets color before checking user properties
        // No properties.TryGetValue("color", ...) to override
        raw.Should().Contain("FF0000",
            "hyperlink color should use user-provided value 'FF0000', " +
            "but Add.cs line 786 hardcodes color to '0563C1', ignoring user input");
    }

    // ===========================================================================================
    // CATEGORY C: Excel comment Set "ref" not supported
    // ===========================================================================================

    // BUG #1410: Excel Set comment ref not supported
    [Fact]
    public void Bug1410_Excel_Set_CommentRef_NotSupported()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Data" });
        _excelHandler.Add("/Sheet1", "comment", null, new()
        {
            ["ref"] = "A1",
            ["text"] = "Note",
            ["author"] = "Test"
        });

        var unsupported = _excelHandler.Set("/Sheet1/comment[1]", new()
        {
            ["ref"] = "B1"
        });

        unsupported.Should().NotContain("ref",
            "Excel comment Set should support changing the cell reference");
    }

    // ===========================================================================================
    // CATEGORY D: Word Get paragraph style not returned
    // ===========================================================================================

    // BUG #1411: Word Get paragraph style not in format
    [Fact]
    public void Bug1411_Word_Get_ParagraphStyle_NotReturned()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Styled text"
        });
        _wordHandler.Set("/body/p[1]", new() { ["style"] = "Heading1" });

        var node = _wordHandler.Get("/body/p[1]");

        node.Format.Should().ContainKey("style",
            "Get should return style for a paragraph that has a custom style set");
    }

    // ===========================================================================================
    // CATEGORY E: Excel Set autofilter not supported at sheet level
    // ===========================================================================================

    // BUG #1413: Excel Set autofilter missing from sheet-level Set
    [Fact]
    public void Bug1413_Excel_Set_AutoFilter_NotSupported()
    {
        _excelHandler.Add("/Sheet1", "row", null, new() { ["values"] = "A,B,C" });
        _excelHandler.Add("/Sheet1", "row", null, new() { ["values"] = "1,2,3" });

        _excelHandler.Set("/Sheet1", new() { ["autofilter"] = "A1:C2" });

        var raw = _excelHandler.Raw("/Sheet1");

        raw.Should().Contain("autoFilter",
            "setting autofilter on a sheet should create an AutoFilter element");
    }

    // ===========================================================================================
    // CATEGORY F: PPTX Add shape doesn't support "opacity"
    // Set supports opacity via Alpha element, but Add doesn't
    // ===========================================================================================

    // BUG #1415: PPTX Add shape opacity not supported
    [Fact]
    public void Bug1415_Pptx_Add_Shape_OpacityNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1415_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });

            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Semi-transparent",
                ["fill"] = "FF0000",
                ["opacity"] = "0.5"
            });

            var node = handler.Get("/slide[1]/shape[2]");

            node.Format.Should().ContainKey("opacity",
                "PPTX Add shape should support 'opacity' during creation");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ===========================================================================================
    // CATEGORY G: Word Set run "link" with non-URL throws UriFormatException
    // Should handle gracefully or provide clear error
    // ===========================================================================================

    // BUG #1416: Word Set run link with non-URL crashes
    [Fact]
    public void Bug1416_Word_Set_RunLink_NonUrlThrowsUriFormatException()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });

        var act = () => _wordHandler.Set("/body/p[1]/r[1]", new()
        {
            ["link"] = "not-a-url"
        });

        // BUG: throws System.UriFormatException with unhelpful message
        // Should either accept relative URIs or throw clear ArgumentException
        act.Should().NotThrow<UriFormatException>(
            "setting a link with a non-URL value should not throw UriFormatException, " +
            "it should either accept relative paths or throw a clear ArgumentException");
    }

    // ===========================================================================================
    // CATEGORY H: Excel Set named range scope not changeable
    // ===========================================================================================

    // BUG #1421: Excel Set named range "scope" not supported
    [Fact]
    public void Bug1421_Excel_Set_NamedRange_ScopeNotSupported()
    {
        _excelHandler.Add("/", "namedrange", null, new()
        {
            ["name"] = "TestRange",
            ["ref"] = "Sheet1!A1:A10"
        });

        var unsupported = _excelHandler.Set("/namedrange[1]", new()
        {
            ["scope"] = "Sheet1"
        });

        unsupported.Should().NotContain("scope",
            "Excel named range Set should support changing the scope");
    }
}
