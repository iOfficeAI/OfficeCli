// Bug hunt tests Part 10: FormulaParser round-trip, ExcelStyleManager color,
// Excel formula/value, PPTX table cell border, and comment author case sensitivity.
// All bugs verified by running tests — every test in this file SHOULD FAIL.

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ===========================================================================================
    // CATEGORY A: FormulaParser iint/iiint reverse mapping missing
    // Forward: \iint → "∬", \iiint → "∭" (FormulaParser.cs line 1032-1033)
    // Reverse: NaryCharToCommand missing "∬" and "∭" (line 1306-1315)
    // ===========================================================================================

    // BUG #1501: FormulaParser \iint cannot round-trip through ToLatex
    [Fact]
    public void Bug1501_FormulaParser_Iint_RoundTrip_Fails()
    {
        var omml = FormulaParser.Parse("\\iint_0^1 f(x)");
        var latex = FormulaParser.ToLatex(omml);

        // BUG: NaryCharToCommand (line 1306) doesn't map "∬" back to "\\iint"
        latex.Should().Contain("\\iint",
            "FormulaParser should round-trip \\iint back to LaTeX, " +
            "but NaryCharToCommand (line 1306) is missing the ∬ → \\iint mapping");
    }

    // BUG #1502: FormulaParser \iiint cannot round-trip through ToLatex
    [Fact]
    public void Bug1502_FormulaParser_Iiint_RoundTrip_Fails()
    {
        var omml = FormulaParser.Parse("\\iiint_0^1 g(x)");
        var latex = FormulaParser.ToLatex(omml);

        latex.Should().Contain("\\iiint",
            "FormulaParser should round-trip \\iiint back to LaTeX, " +
            "but NaryCharToCommand is missing the ∭ → \\iiint mapping");
    }

    // ===========================================================================================
    // CATEGORY B: Excel font color not readable after setting
    // ExcelStyleManager NormalizeColor adds "FF" prefix to 6-char hex,
    // but Get doesn't expose font.color in format dictionary
    // ===========================================================================================

    // BUG #1503: Excel cell font color readable via Get when set via font.color key
    [Fact]
    public void Bug1503_Excel_Get_FontColor_Returned()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Red text",
            ["font.color"] = "FF0000"
        });

        var node = _excelHandler.Get("/Sheet1/A1");

        // font.color should be returned in the Format dict
        node.Format.Should().ContainKey("font.color",
            "Get should return font.color for a cell that has font.color set");
    }

    // ===========================================================================================
    // CATEGORY C: PPTX table cell border not supported in SetTableCellProperties
    // ===========================================================================================

    // BUG #1505: PPTX table cell border.all returns as unsupported
    [Fact]
    public void Bug1505_Pptx_Set_TableCellBorder_NotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1505_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "table", null, new()
            {
                ["rows"] = "1", ["cols"] = "1"
            });

            var unsupported = handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            {
                ["border.all"] = "FF0000"
            });

            unsupported.Should().NotContain("border.all",
                "PPTX table cell Set should support 'border.all' property");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ===========================================================================================
    // CATEGORY D: Excel Set value after formula — formula not cleared
    // ===========================================================================================

    // BUG #1507: Excel Set value doesn't clear existing formula
    [Fact]
    public void Bug1507_Excel_Set_ValueAfterFormula_FormulaPersists()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["formula"] = "=1+1" });
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "manual" });

        var node = _excelHandler.Get("/Sheet1/A1");

        // The formula should be cleared when a value is explicitly set
        node.Format.Should().NotContainKey("formula",
            "setting a value on a cell with a formula should clear the formula, " +
            "but ExcelHandler.Set.cs 'value' case doesn't clear CellFormula");
    }

    // ===========================================================================================
    // CATEGORY E: Excel comment author lookup is case-sensitive
    // ExcelHandler.Add.cs line 191: FindIndex(a => a.Text == cmtAuthor) — exact match
    // "John" and "john" create duplicate authors
    // ===========================================================================================

    // Note: Excel comment author matching is intentionally case-sensitive,
    // consistent with Apache POI's CommentsTable.findAuthor() which uses .equals().
}
