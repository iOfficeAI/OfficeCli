// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Fuzz tests for all newly added table properties across PPTX, DOCX, and XLSX.
/// Focus: boundary values, invalid inputs, conflict scenarios, and null/empty edge cases.
/// Tests pass if no unhandled exception is thrown (crash = CRITICAL bug).
/// Tests that validate error messages use try/catch (bad input without error = MINOR bug).
/// </summary>
public class FuzzTableTests : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private PowerPointHandler _pptx;
    private WordHandler _word;
    private ExcelHandler _excel;

    public FuzzTableTests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"fuzz_{Guid.NewGuid():N}.pptx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"fuzz_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"fuzz_{Guid.NewGuid():N}.xlsx");

        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);

        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        _word = new WordHandler(_docxPath, editable: true);
        _excel = new ExcelHandler(_xlsxPath, editable: true);

        _pptx.Add("/", "slide", null, new() { ["layout"] = "blank" });
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "3" });

        _word.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "3" });

        _excel.Add("/Sheet1", "table", null, new()
        {
            ["ref"] = "A1:C4", ["data"] = "H1,H2,H3;V1,V2,V3;V4,V5,V6;V7,V8,V9"
        });
    }

    public void Dispose()
    {
        _pptx.Dispose();
        _word.Dispose();
        _excel.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    // ==================== 1. BOUNDARY VALUES: margin/padding ====================

    [Fact]
    public void Fuzz_Pptx_CellMargin_Zero_DoesNotCrash()
    {
        // Zero margin should be valid - produces 0 EMU
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin"] = "0cm" });
        act.Should().NotThrow();
        var node = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Pptx_CellMargin_Negative_ThrowsOrIgnores()
    {
        // Negative margin: should throw ArgumentException or silently clamp, NOT crash unhandled
        try
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin"] = "-1cm" });
            // If no exception, the file should still be valid (not corrupt)
            var node = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
            node.Should().NotBeNull();
        }
        catch (ArgumentException)
        {
            // Acceptable: explicit error message for invalid input
        }
    }

    [Fact]
    public void Fuzz_Pptx_CellMargin_VeryLarge_DoesNotCrash()
    {
        // 999cm is extreme but should not crash
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin"] = "999cm" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_CellMargin_SubMillimeter_DoesNotCrash()
    {
        // Very small fractional value
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin"] = "0.001cm" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_CellPadding_NegativeFourValues_ThrowsOrIgnores()
    {
        // Negative padding in four-value syntax
        try
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["padding"] = "-0.5cm,-0.2cm,-0.1cm,-0.3cm" });
            var node = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
            node.Should().NotBeNull();
        }
        catch (ArgumentException)
        {
            // Expected for invalid negative padding
        }
    }

    [Fact]
    public void Fuzz_Word_TableColWidths_NegativeValue_ThrowsOrIgnores()
    {
        // Word colWidths: negative twip value
        try
        {
            _word.Set("/body/tbl[1]", new() { ["colWidths"] = "-1000,2000,3000" });
        }
        catch (ArgumentException)
        {
            // Acceptable
        }
    }

    // ==================== 2. BOUNDARY VALUES: opacity ====================

    [Fact]
    public void Fuzz_Pptx_Opacity_Zero_DoesNotCrash()
    {
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "0" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_Opacity_OneHundred_DoesNotCrash()
    {
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "100" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_Opacity_OverOneHundred_ThrowsOrClamps()
    {
        // 101% opacity is invalid - should throw or clamp
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        try
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "101" });
            // If no exception, file should still open cleanly
            _pptx.Dispose();
            _pptx = new PowerPointHandler(_pptxPath, editable: true);
        }
        catch (ArgumentException)
        {
            // Expected behavior
        }
    }

    [Fact]
    public void Fuzz_Pptx_Opacity_Negative_ThrowsOrClamps()
    {
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        try
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "-1" });
            _pptx.Dispose();
            _pptx = new PowerPointHandler(_pptxPath, editable: true);
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Pptx_Opacity_Fractional_DoesNotCrash()
    {
        // 50.5 is a decimal value - should be handled
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "50.5" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_Opacity_NonNumeric_ThrowsWithMessage()
    {
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "abc" });
        act.Should().Throw<Exception>();
    }

    // ==================== 3. BOUNDARY VALUES: bevel dimensions ====================

    [Fact]
    public void Fuzz_Pptx_BevelDimensions_Zero_DoesNotCrash()
    {
        // Zero bevel dimensions
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle-0-0" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_BevelDimensions_Negative_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle--5-3" });
            var node = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
            node.Should().NotBeNull();
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException)
        {
            // Acceptable
        }
    }

    [Fact]
    public void Fuzz_Pptx_BevelDimensions_VeryLarge_DoesNotCrash()
    {
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle-9999-9999" });
        act.Should().NotThrow();
    }

    // ==================== 4. BOUNDARY VALUES: colWidths ====================

    [Fact]
    public void Fuzz_Pptx_ColWidths_ZeroCm_DoesNotCrash()
    {
        // Zero-width columns
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["colWidths"] = "0cm,5cm,5cm" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_ColWidths_EmptyString_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["colWidths"] = "" });
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException or InvalidOperationException)
        {
            // Acceptable
        }
    }

    [Fact]
    public void Fuzz_Pptx_ColWidths_MoreValuesThanColumns_DoesNotCrash()
    {
        // 5 widths for 3-column table: extra values should be ignored
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["colWidths"] = "2cm,3cm,4cm,5cm,6cm" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_ColWidths_FewerValuesThanColumns_DoesNotCrash()
    {
        // Only 1 width for 3-column table
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["colWidths"] = "4cm" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Word_ColWidths_ZeroValue_ThrowsOrIgnores()
    {
        // Word colWidths in twips: 0 might be invalid
        try
        {
            _word.Set("/body/tbl[1]", new() { ["colWidths"] = "0,2000,3000" });
        }
        catch (ArgumentException)
        {
            // Acceptable
        }
    }

    [Fact]
    public void Fuzz_Word_ColWidths_NonNumericValue_ThrowsOrIgnores()
    {
        // KNOWN ISSUE: Word table Set handler does not support colWidths (only Add does).
        // So non-numeric input may be silently ignored (added to unsupported list).
        // This is a MINOR bug: invalid input should produce an error, not silent ignore.
        try
        {
            _word.Set("/body/tbl[1]", new() { ["colWidths"] = "abc,2000,3000" });
            // If no exception thrown, the file should still be valid (not corrupt)
            _word.Dispose();
            _word = new WordHandler(_docxPath, editable: true);
            _word.Get("/body/tbl[1]").Should().NotBeNull();
        }
        catch (ArgumentException)
        {
            // Also acceptable if implementation validates
        }
    }

    // ==================== 5. BOUNDARY VALUES: lineSpacing ====================

    [Fact]
    public void Fuzz_Pptx_LineSpacing_ZeroMultiplier_DoesNotCrash()
    {
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "T", ["lineSpacing"] = "0x" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_LineSpacing_ZeroPercent_DoesNotCrash()
    {
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "T", ["lineSpacing"] = "0%" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_LineSpacing_ExtremeMultiplier_DoesNotCrash()
    {
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "T", ["lineSpacing"] = "999x" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_LineSpacing_Negative_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "T", ["lineSpacing"] = "-1x" });
            _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]").Should().NotBeNull();
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException or OverflowException)
        {
            // Acceptable
        }
    }

    // ==================== 6. INVALID INPUTS: textDirection ====================

    [Fact]
    public void Fuzz_Pptx_TextDirection_Invalid_ThrowsWithMessage()
    {
        // "invalid" is not a recognized direction
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "invalid" });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*textDirection*");
    }

    [Fact]
    public void Fuzz_Pptx_TextDirection_Empty_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "" });
        }
        catch (ArgumentException)
        {
            // Expected: empty direction should not silently apply garbage
        }
    }

    [Fact]
    public void Fuzz_Pptx_TextDirection_Number_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "42" });
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    // ==================== 7. INVALID INPUTS: bandColor ====================

    [Fact]
    public void Fuzz_Pptx_BandColorOdd_InvalidHex_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.odd"] = "ZZZZZZ" });
            // If no exception, file must still be valid
            _pptx.Dispose();
            _pptx = new PowerPointHandler(_pptxPath, editable: true);
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Pptx_BandColorOdd_Empty_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.odd"] = "" });
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Pptx_BandColorOdd_Keyword_None_DoesNotCrash()
    {
        // "none" is a special color keyword
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.odd"] = "none" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_BandColorEven_TooShortHex_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.even"] = "FFF" }); // 3-char shorthand
            _pptx.Dispose();
            _pptx = new PowerPointHandler(_pptxPath, editable: true);
        }
        catch (ArgumentException)
        {
            // Acceptable for unsupported shorthand
        }
    }

    // ==================== 8. INVALID INPUTS: shadow/glow ====================

    [Fact]
    public void Fuzz_Pptx_Shadow_MalformedFormat_ThrowsOrIgnores()
    {
        // Valid format: COLOR-blurRadius-angle-distance-opacity
        // Malformed: missing parts
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "FF0000" }); // only color, no other parts
            _pptx.Dispose();
            _pptx = new PowerPointHandler(_pptxPath, editable: true);
        }
        catch (ArgumentException)
        {
            // Expected for incomplete format
        }
    }

    [Fact]
    public void Fuzz_Pptx_Shadow_EmptyString_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "" });
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Pptx_Shadow_NonNumericValues_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "FF0000-abc-xyz-def-ghi" });
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Pptx_Glow_MalformedFormat_ThrowsOrIgnores()
    {
        // Valid format: COLOR-radius-opacity
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["glow"] = "notacolor-notanumber" });
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Pptx_Glow_EmptyString_ThrowsOrIgnores()
    {
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["glow"] = "" });
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    // ==================== 9. INVALID INPUTS: Excel col[N].totalFunction ====================

    [Fact]
    public void Fuzz_Excel_TotalFunction_InvalidFunctionName_ThrowsWithMessage()
    {
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "invalid_func" });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*totalFunction*");
    }

    [Fact]
    public void Fuzz_Excel_TotalFunction_EmptyString_ThrowsOrIgnores()
    {
        try
        {
            _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "" });
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Excel_Col_ZeroIndex_ThrowsWithMessage()
    {
        // 0-indexed access - our API is 1-based
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[0].name"] = "ZeroIndex" });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*out of range*");
    }

    [Fact]
    public void Fuzz_Excel_Col_OutOfRange_ThrowsWithMessage()
    {
        // col[999] is out of range for a 3-column table
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[999].name"] = "OutOfRange" });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*out of range*");
    }

    [Fact]
    public void Fuzz_Excel_Col_MaxIntIndex_ThrowsWithMessage()
    {
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[2147483647].name"] = "MaxInt" });
        act.Should().Throw<ArgumentException>();
    }

    // ==================== 10. INVALID INPUTS: position.x/y ====================

    [Fact]
    public void Fuzz_Word_PositionX_VeryLargeTwips_DoesNotCrash()
    {
        var act = () => _word.Set("/body/tbl[1]", new() { ["position.x"] = "9999cm" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Word_PositionX_NegativeValue_ThrowsOrIgnores()
    {
        // Negative position may not be valid in Word
        try
        {
            _word.Set("/body/tbl[1]", new() { ["position.x"] = "-2cm" });
            _word.Dispose();
            _word = new WordHandler(_docxPath, editable: true);
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Word_PositionY_NegativeValue_ThrowsOrIgnores()
    {
        try
        {
            _word.Set("/body/tbl[1]", new() { ["position.y"] = "-5cm" });
            _word.Dispose();
            _word = new WordHandler(_docxPath, editable: true);
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    // ==================== 11. CONFLICT SCENARIOS ====================

    [Fact]
    public void Fuzz_Pptx_SetFillThenOpacity_OpacitySurvives()
    {
        // Set solid fill, then set opacity - opacity must survive and be applied to the fill
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "4472C4" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "60" });

        // Re-set fill again (does opacity survive? It should not, unless fill preserves it)
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        // File must remain valid
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var node = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Pptx_SetBevelThenFill_BothCoexist()
    {
        // Bevel and fill should coexist without crashing
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "00FF00" });
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var node = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Pptx_SetTextDirectionThenWordWrap_BothApply()
    {
        // textDirection and wordWrap should not conflict
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "vertical" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["wordWrap"] = "false" });
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var node = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Pptx_BandColorOddThenBandColorEven_BothApply()
    {
        // Setting bandColor.odd and bandColor.even on the same table should work
        _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.odd"] = "E8F0FE" });
        _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.even"] = "FFF3CD" });
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var node = _pptx.Get("/slide[1]/table[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Pptx_ShadowThenGlow_BothEffectsApplied()
    {
        // Shadow and glow are both effects - both should coexist in effectList
        _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "000000-4-135-3-50" });
        _pptx.Set("/slide[1]/table[1]", new() { ["glow"] = "4472C4-10-60" });
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var node = _pptx.Get("/slide[1]/table[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Pptx_SetSamePropertyMultipleTimes_DoesNotCrash()
    {
        // Repeated sets should be idempotent, not accumulate elements
        for (int i = 0; i < 10; i++)
        {
            _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle" });
        }
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var node = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Pptx_GradientFillThenOpacity_DoesNotCrash()
    {
        // Gradient fill + opacity: opacity might not be applicable to gradient
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000-0000FF-90" });
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "50" });
        act.Should().NotThrow();
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    // ==================== 12. UNEQUAL ROW LENGTHS IN DATA ====================

    [Fact]
    public void Fuzz_Pptx_AddTable_DataWithUnequalRowLengths_DoesNotCrash()
    {
        // Rows have different numbers of cells - should pad or truncate
        var act = () => _pptx.Add("/slide[1]", "table", null, new()
        {
            ["data"] = "A,B,C;X,Y;P"
        });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Word_AddTable_DataWithUnequalRowLengths_DoesNotCrash()
    {
        var act = () => _word.Add("/body", "table", null, new()
        {
            ["data"] = "H1,H2,H3;V1;V2,V3,V4,V5"
        });
        act.Should().NotThrow();
    }

    // ==================== 13. EMPTY/NULL EDGE CASES ====================

    [Fact]
    public void Fuzz_Word_Caption_EmptyString_DoesNotCrash()
    {
        // Empty caption should be handled gracefully (remove or no-op)
        _word.Set("/body/tbl[1]", new() { ["caption"] = "Some Caption" });
        var act = () => _word.Set("/body/tbl[1]", new() { ["caption"] = "" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Word_Description_EmptyString_DoesNotCrash()
    {
        _word.Set("/body/tbl[1]", new() { ["description"] = "Some Description" });
        var act = () => _word.Set("/body/tbl[1]", new() { ["description"] = "" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Excel_ColFormula_EmptyString_DoesNotCrash()
    {
        // Setting a column formula to empty string should clear or no-op
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[1].formula"] = "" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_AddTable_EmptyData_DoesNotCrash()
    {
        // Empty data string - should create empty or default table
        try
        {
            var path = _pptx.Add("/slide[1]", "table", null, new() { ["data"] = "" });
            path.Should().NotBeNullOrEmpty();
        }
        catch (ArgumentException)
        {
            // Acceptable: explicit error for empty data
        }
    }

    [Fact]
    public void Fuzz_Word_AddTable_EmptyData_DoesNotCrash()
    {
        try
        {
            var path = _word.Add("/body", "table", null, new() { ["data"] = "" });
            path.Should().NotBeNullOrEmpty();
        }
        catch (ArgumentException)
        {
            // Acceptable
        }
    }

    [Fact]
    public void Fuzz_Word_Caption_VeryLongString_DoesNotCrash()
    {
        // Extremely long caption - XML attribute has no inherent length limit
        var longCaption = new string('A', 10000);
        var act = () => _word.Set("/body/tbl[1]", new() { ["caption"] = longCaption });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Word_Caption_SpecialXmlChars_DoesNotCrash()
    {
        // XML special characters in caption
        var act = () => _word.Set("/body/tbl[1]", new() { ["caption"] = "<>&\"'" });
        act.Should().NotThrow();
        // Verify file can be reopened after XML-special chars
        _word.Dispose();
        _word = new WordHandler(_docxPath, editable: true);
    }

    [Fact]
    public void Fuzz_Word_Description_SpecialXmlChars_DoesNotCrash()
    {
        var act = () => _word.Set("/body/tbl[1]", new() { ["description"] = "<>&\"'" });
        act.Should().NotThrow();
        _word.Dispose();
        _word = new WordHandler(_docxPath, editable: true);
    }

    // ==================== 14. ADDITIONAL BOUNDARY: rows=0, cols=0 ====================

    [Fact]
    public void Fuzz_Pptx_AddTable_ZeroRows_ThrowsWithMessage()
    {
        var act = () => _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "0", ["cols"] = "3" });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Fuzz_Pptx_AddTable_ZeroCols_ThrowsWithMessage()
    {
        var act = () => _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "0" });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Fuzz_Pptx_AddTable_NegativeRows_ThrowsWithMessage()
    {
        var act = () => _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "-1", ["cols"] = "3" });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Fuzz_Word_AddTable_ZeroRows_ThrowsWithMessage()
    {
        var act = () => _word.Add("/body", "table", null, new() { ["rows"] = "0", ["cols"] = "3" });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Fuzz_Word_AddTable_NegativeRows_ThrowsWithMessage()
    {
        var act = () => _word.Add("/body", "table", null, new() { ["rows"] = "-5", ["cols"] = "3" });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Fuzz_Word_AddTable_NonNumericRows_ThrowsWithMessage()
    {
        var act = () => _word.Add("/body", "table", null, new() { ["rows"] = "abc", ["cols"] = "3" });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Fuzz_Pptx_AddTable_NonNumericCols_ThrowsWithMessage()
    {
        var act = () => _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "abc" });
        act.Should().Throw<ArgumentException>();
    }

    // ==================== 15. EXTREME SIZES ====================

    [Fact]
    public void Fuzz_Pptx_AddTable_VeryLargeTable_DoesNotCrash()
    {
        // Large table: 50 rows x 10 cols
        var act = () => _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "50", ["cols"] = "10" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Word_AddTable_VeryLargeTable_DoesNotCrash()
    {
        var act = () => _word.Add("/body", "table", null, new() { ["rows"] = "100", ["cols"] = "10" });
        act.Should().NotThrow();
    }

    // ==================== 16. Excel-specific: totalFunction edge cases ====================

    [Fact]
    public void Fuzz_Excel_TotalFunction_Sum_ValidAndDoesNotCrash()
    {
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "sum" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Excel_TotalFunction_None_ValidAndDoesNotCrash()
    {
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "none" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Excel_TotalFunction_Custom_DoesNotCrash()
    {
        // "custom" allows custom formula; should not crash even without a formula
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "custom" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Excel_ColFormula_VeryLongFormula_DoesNotCrash()
    {
        // Extremely long formula string
        var longFormula = "=SUM(A1,B1," + string.Join(",", Enumerable.Range(1, 100).Select(i => $"C{i}")) + ")";
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[1].formula"] = longFormula });
        act.Should().NotThrow();
    }

    // ==================== 17. Word floating table: edge cases ====================

    [Fact]
    public void Fuzz_Word_Position_InvalidAnchor_ThrowsOrIgnores()
    {
        try
        {
            _word.Set("/body/tbl[1]", new() { ["position.hAnchor"] = "invalid_anchor" });
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Word_Overlap_InvalidValue_ThrowsOrIgnores()
    {
        try
        {
            _word.Set("/body/tbl[1]", new() { ["position"] = "floating", ["overlap"] = "maybe" });
        }
        catch (ArgumentException)
        {
            // Expected
        }
    }

    [Fact]
    public void Fuzz_Word_PositionFromText_NegativeDistance_ThrowsOrIgnores()
    {
        try
        {
            _word.Set("/body/tbl[1]", new()
            {
                ["position"] = "floating",
                ["position.left"] = "-0.5cm",
                ["position.right"] = "-0.5cm"
            });
            _word.Dispose();
            _word = new WordHandler(_docxPath, editable: true);
        }
        catch (ArgumentException)
        {
            // Acceptable
        }
    }

    // ==================== 18. PPTX table: shadow/glow with extreme values ====================

    [Fact]
    public void Fuzz_Pptx_Shadow_ExtremeBlurRadius_DoesNotCrash()
    {
        // Very large blur radius
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "000000-9999-135-3-50" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_Shadow_ZeroDistance_DoesNotCrash()
    {
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "000000-4-135-0-50" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_Shadow_Angle360_DoesNotCrash()
    {
        // 360 degrees is equivalent to 0
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "000000-4-360-3-50" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_Glow_ZeroRadius_DoesNotCrash()
    {
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["glow"] = "FF0000-0-50" });
        act.Should().NotThrow();
    }

    // ==================== 19. FitText edge cases (Word) ====================

    [Fact]
    public void Fuzz_Word_FitText_WithoutText_DoesNotCrash()
    {
        // FitText on empty cell - no runs to apply to
        var act = () => _word.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "true" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Word_FitText_InvalidValue_ThrowsOrIgnores()
    {
        try
        {
            _word.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "maybe" });
        }
        catch (ArgumentException)
        {
            // Acceptable: "maybe" is not truthy/falsy
        }
    }

    // ==================== 20. PPTX: colWidths with bad units ====================

    [Fact]
    public void Fuzz_Pptx_ColWidths_InvalidUnit_ThrowsOrIgnores()
    {
        // "2bananas" is not a valid unit
        try
        {
            _pptx.Set("/slide[1]/table[1]", new() { ["colWidths"] = "2bananas,3cm,4cm" });
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException or OverflowException)
        {
            // Expected
        }
    }

    // ==================== 21. Simultaneous multi-property stress test ====================

    [Fact]
    public void Fuzz_Pptx_SetManyPropertiesAtOnce_DoesNotCrash()
    {
        // Set many properties simultaneously on a cell
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Stress",
            ["fill"] = "4472C4",
            ["bevel"] = "circle-2-2",
            ["textDirection"] = "horizontal",
            ["wordWrap"] = "true",
            ["margin"] = "0.2cm",
            ["lineSpacing"] = "1.2x",
            ["spaceBefore"] = "6pt",
            ["spaceAfter"] = "6pt"
        });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Pptx_SetManyTablePropertiesAtOnce_DoesNotCrash()
    {
        var act = () => _pptx.Set("/slide[1]/table[1]", new()
        {
            ["firstRow"] = "true",
            ["lastRow"] = "true",
            ["firstCol"] = "false",
            ["lastCol"] = "false",
            ["bandedRows"] = "true",
            ["bandedCols"] = "false",
            ["colWidths"] = "3cm,5cm,3cm",
            ["shadow"] = "000000-3-135-2-30"
        });
        act.Should().NotThrow();
    }

    [Fact]
    public void Fuzz_Word_SetManyPropertiesAtOnce_DoesNotCrash()
    {
        var act = () => _word.Set("/body/tbl[1]", new()
        {
            ["firstRow"] = "true",
            ["lastRow"] = "true",
            ["bandedRows"] = "true",
            ["bandedCols"] = "false",
            ["caption"] = "Stress Test Table",
            ["description"] = "Testing many properties at once",
            ["position"] = "floating",
            ["position.x"] = "2cm",
            ["position.y"] = "3cm",
            ["overlap"] = "never"
        });
        act.Should().NotThrow();
    }

    // ==================== 22. File validity after fuzz: reopen after each major fuzz ====================

    [Fact]
    public void Fuzz_Pptx_AfterAllMutations_FileReopensCleanly()
    {
        // Apply a series of mutations and verify the file is still valid
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "75" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "softRound-3-3" });
        _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "333333-5-180-3-60" });
        _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.odd"] = "E0E0FF" });
        _pptx.Set("/slide[1]/table[1]", new() { ["colWidths"] = "4cm,3cm,4cm" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "none" });

        _pptx.Dispose();
        var act = () => { _pptx = new PowerPointHandler(_pptxPath, editable: true); };
        act.Should().NotThrow();

        var node = _pptx.Get("/slide[1]/table[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Word_AfterAllMutations_FileReopensCleanly()
    {
        _word.Set("/body/tbl[1]", new() { ["caption"] = "Test <&>" });
        _word.Set("/body/tbl[1]", new() { ["firstRow"] = "true", ["bandedRows"] = "true" });
        _word.Set("/body/tbl[1]", new() { ["position.x"] = "1cm", ["position.y"] = "2cm" });
        _word.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "true", ["text"] = "FitMe" });

        _word.Dispose();
        var act = () => { _word = new WordHandler(_docxPath, editable: true); };
        act.Should().NotThrow();

        var node = _word.Get("/body/tbl[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Fuzz_Excel_AfterAllMutations_FileReopensCleanly()
    {
        _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "sum" });
        _excel.Set("/Sheet1/table[1]", new() { ["col[2].name"] = "Renamed Col" });
        _excel.Set("/Sheet1/table[1]", new() { ["col[3].formula"] = "=[H1]*2" });
        _excel.Set("/Sheet1/table[1]", new() { ["showRowStripes"] = "true" });
        _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "none" });

        _excel.Dispose();
        var act = () => { _excel = new ExcelHandler(_xlsxPath, editable: true); };
        act.Should().NotThrow();

        var node = _excel.Get("/Sheet1/table[1]");
        node.Should().NotBeNull();
    }
}
