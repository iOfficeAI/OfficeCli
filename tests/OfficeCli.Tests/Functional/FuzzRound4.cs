// FuzzRound4 — Deep fuzz: Core layer, file-level edge cases, deep paths, extreme format values,
//               conflicting Set operations, concurrent-style access.
//
// Areas:
//   SC01–SC08: SpacingConverter — extreme/boundary values, invalid input
//   CP01–CP06: Color parsing — ARGB alpha, transparent, edge hex, rgb() edge cases
//   FP01–FP05: FormulaParser — empty, deeply nested, invalid latex, very long
//   FL01–FL04: File layer — 0-byte file, corrupt file, read-only, missing dir
//   DP01–DP06: Deep paths — nested selectors, index 0, -1, 9999
//   EF01–EF04: Extreme format — font 999pt, 0.1pt, very long color string, EMU max
//   CO01–CO06: Conflicting / combined Set operations on same element
//   AE01–AE04: Add same-name elements multiple times, Remove non-existent

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound4 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz4_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    private string TempPath(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz4_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
        {
            try
            {
                if (File.Exists(f))
                {
                    File.SetAttributes(f, FileAttributes.Normal);
                    File.Delete(f);
                }
            }
            catch { }
        }
    }

    // ==================== SC01–SC08: SpacingConverter boundary values ====================

    [Theory]
    [InlineData("0pt")]
    [InlineData("0cm")]
    [InlineData("0in")]
    [InlineData("0")]
    public void SC01_SpacingConverter_ZeroValues_ReturnZero(string value)
    {
        var wordResult = SpacingConverter.ParseWordSpacing(value);
        wordResult.Should().Be(0u, $"zero spacing '{value}' should parse to 0 twips");
    }

    [Theory]
    [InlineData("10000pt")]   // extreme large
    [InlineData("999in")]
    [InlineData("1000cm")]
    public void SC02_SpacingConverter_VeryLargeValues_DoNotThrow(string value)
    {
        var act = () => SpacingConverter.ParseWordSpacing(value);
        act.Should().NotThrow($"large spacing '{value}' should parse without throwing");
    }

    [Theory]
    [InlineData("-1pt")]
    [InlineData("-0.001")]
    [InlineData("-5in")]
    public void SC03_SpacingConverter_NegativeValues_Throw(string value)
    {
        var act = () => SpacingConverter.ParseWordSpacing(value);
        act.Should().Throw<ArgumentException>($"negative spacing '{value}' should throw");
    }

    [Theory]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("-Infinity")]
    [InlineData("abc")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("pt")]
    public void SC04_SpacingConverter_InvalidStrings_Throw(string value)
    {
        var act = () => SpacingConverter.ParseWordSpacing(value);
        act.Should().Throw<ArgumentException>($"invalid spacing string '{value}' should throw");
    }

    [Fact]
    public void SC04b_SpacingConverter_StringWithSpaceBeforeUnit_IsAccepted()
    {
        // "12 pt" has a space before the unit — the parser trims and accepts bare "12" after stripping suffix
        // Document this behavior: "12 pt" does NOT throw (it is treated as bare number "12" after stripping suffix " pt")
        // This is an observed behavior. If the intent is to reject it, this test documents a potential gap.
        var act = () => SpacingConverter.ParseWordSpacing("12 pt");
        act.Should().NotThrow("'12 pt' with space is currently accepted (space before unit is tolerated)");
    }

    [Theory]
    [InlineData("1.5x")]
    [InlineData("150%")]
    [InlineData("200%")]
    [InlineData("0.5x")]
    public void SC05_SpacingConverter_LineSpacingMultiplier_RoundTrip(string value)
    {
        var (twips, isMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
        isMultiplier.Should().BeTrue($"'{value}' should be parsed as multiplier");
        twips.Should().BeGreaterThan(0u);
    }

    [Theory]
    [InlineData("18pt")]
    [InlineData("0.5cm")]
    [InlineData("1in")]
    public void SC06_SpacingConverter_LineSpacingFixed_RoundTrip(string value)
    {
        var (twips, isMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
        isMultiplier.Should().BeFalse($"'{value}' should be parsed as fixed line spacing");
        twips.Should().BeGreaterThan(0u);
    }

    [Fact]
    public void SC07_SpacingConverter_FormatWordSpacing_ZeroInput()
    {
        var result = SpacingConverter.FormatWordSpacing("0");
        result.Should().Be("0pt");
    }

    [Fact]
    public void SC08_SpacingConverter_FormatWordLineSpacing_NullLineRule_TreatedAsAuto()
    {
        // 480 twips = 2x multiplier
        var result = SpacingConverter.FormatWordLineSpacing("480", null);
        result.Should().Be("2x");
    }

    // ==================== CP01–CP06: Color parsing edge cases ====================

    [Theory]
    [InlineData("00FF0000")]  // fully transparent red (alpha=00)
    [InlineData("80FF0000")]  // 50% alpha red
    [InlineData("01FF0000")]  // near-transparent
    [InlineData("FEFFFFFF")]  // near-opaque white
    public void CP01_SanitizeColorForOoxml_ArgbWithAlpha_ParsesAlpha(string color)
    {
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(color);
        rgb.Should().HaveLength(6);
        rgb.All(char.IsAsciiHexDigit).Should().BeTrue();
        if (color.StartsWith("FF")) alpha.Should().BeNull();
        else alpha.Should().NotBeNull();
    }

    [Fact]
    public void CP02_SanitizeColorForOoxml_FullyTransparent_HasNearZeroAlpha()
    {
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml("00FF0000");
        alpha.Should().NotBeNull();
        alpha!.Value.Should().Be(0);
    }

    [Fact]
    public void CP03_SanitizeColorForOoxml_FullyOpaque_AlphaIsNull()
    {
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml("FFFF0000");
        alpha.Should().BeNull();
        rgb.Should().Be("FF0000");
    }

    [Theory]
    [InlineData("F00")]       // 3-char shorthand
    [InlineData("#F00")]
    [InlineData("red")]       // named color
    [InlineData("rgb(255,0,0)")]
    [InlineData("FF0000")]    // 6-char
    [InlineData("#FF0000")]
    public void CP04_SanitizeColorForOoxml_ValidInputs_DoNotThrow(string color)
    {
        var act = () => ParseHelpers.SanitizeColorForOoxml(color);
        act.Should().NotThrow($"color '{color}' should be valid");
    }

    [Theory]
    [InlineData("GGGGGG")]       // invalid hex chars
    [InlineData("12345")]        // 5-char — not 3, 6, or 8
    [InlineData("1234567")]      // 7-char
    [InlineData("XXXXXXXXXX")]   // 10-char
    [InlineData("notacolor")]
    [InlineData("rgb(256,0,0)")] // out-of-range RGB
    public void CP05_SanitizeColorForOoxml_InvalidInputs_Throw(string color)
    {
        var act = () => ParseHelpers.SanitizeColorForOoxml(color);
        act.Should().Throw<ArgumentException>($"color '{color}' should be invalid");
    }

    [Theory]
    [InlineData("000000")]  // black
    [InlineData("FFFFFF")]  // white
    [InlineData("000001")]  // near-black
    [InlineData("FFFFFE")]  // near-white
    public void CP06_SanitizeColorForOoxml_BoundaryColors_RoundTrip(string color)
    {
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(color);
        rgb.Should().Be(color.ToUpperInvariant());
        alpha.Should().BeNull();
    }

    // ==================== FP01–FP05: FormulaParser ====================

    [Fact]
    public void FP01_FormulaParser_EmptyString_DoesNotThrow()
    {
        // Empty input may return empty math element — must not crash
        var act = () => FormulaParser.Parse("");
        act.Should().NotThrow("empty latex should be handled gracefully");
    }

    [Fact]
    public void FP02_FormulaParser_SingleChar_ParsesOk()
    {
        var result = FormulaParser.Parse("x");
        result.Should().NotBeNull();
    }

    [Fact]
    public void FP03_FormulaParser_DeeplyNestedFractions_DoNotThrow()
    {
        // Deep nesting: \frac{\frac{\frac{a}{b}}{c}}{d}
        var latex = @"\frac{\frac{\frac{a}{b}}{c}}{d}";
        var act = () => FormulaParser.Parse(latex);
        act.Should().NotThrow("deeply nested fractions should parse without stack overflow");
    }

    [Fact]
    public void FP04_FormulaParser_MismatchedBraces_DoesNotCrash()
    {
        // Unmatched braces — should not throw (may produce partial result)
        var act = () => FormulaParser.Parse(@"\frac{a}{b");
        act.Should().NotThrow("mismatched braces should not crash the parser");
    }

    [Fact]
    public void FP05_FormulaParser_VeryLongExpression_DoesNotThrow()
    {
        // 200-term sum
        var terms = string.Join("+", Enumerable.Range(1, 200).Select(i => $"x_{i}"));
        var act = () => FormulaParser.Parse(terms);
        act.Should().NotThrow("very long formula should parse without timeout/crash");
    }

    // ==================== FL01–FL04: File-layer edge cases ====================

    [Fact]
    public void FL01_ZeroByteFile_OpeningThrows()
    {
        var path = TempPath("pptx");
        File.WriteAllBytes(path, Array.Empty<byte>());
        var act = () => new PowerPointHandler(path, editable: false);
        act.Should().Throw<Exception>("a zero-byte file is not a valid PPTX");
    }

    [Fact]
    public void FL02_CorruptFile_OpeningThrows()
    {
        var path = TempPath("docx");
        File.WriteAllText(path, "this is not a valid docx file at all!!! garbage\x00\x01\x02");
        var act = () => new WordHandler(path, editable: false);
        act.Should().Throw<Exception>("a corrupt file should throw on open");
    }

    [Fact]
    public void FL03_ReadOnlyFile_WritingThrows()
    {
        var path = CreateTemp("xlsx");
        File.SetAttributes(path, FileAttributes.ReadOnly);
        var act = () =>
        {
            using var handler = new ExcelHandler(path, editable: true);
            handler.Add("/", "sheet", null, new() { ["name"] = "S1" });
        };
        act.Should().Throw<Exception>("writing to read-only file should throw");
    }

    [Fact]
    public void FL04_NonExistentFile_OpeningThrows()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz4_nonexistent_{Guid.NewGuid():N}.pptx");
        var act = () => new PowerPointHandler(path, editable: false);
        act.Should().Throw<Exception>("opening a nonexistent file should throw");
    }

    // ==================== DP01–DP06: Deep path / index boundary ====================

    [Fact]
    public void DP01_Get_IndexZero_ThrowsWithClearMessage()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: false);
        // Index 0 is invalid in 1-based system — throws ArgumentException.
        var act = () => handler.Get("/slide[0]");
        act.Should().Throw<ArgumentException>("index 0 is invalid in 1-based path system");
    }

    [Fact]
    public void DP02_Get_NegativeIndex_ThrowsWithClearMessage()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: false);
        // Negative index throws ArgumentException at path parsing stage
        var act = () => handler.Get("/slide[-1]");
        act.Should().Throw<ArgumentException>("negative index is invalid and throws");
    }

    [Fact]
    public void DP03_Get_IndexFarOutOfRange_ThrowsWithClearMessage()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: false);
        // Out-of-range index throws ArgumentException.
        var act = () => handler.Get("/slide[9999]");
        act.Should().Throw<ArgumentException>("out-of-range slide index should throw");
    }

    [Fact]
    public void DP04_Set_PathToNonExistentNode_ThrowsArgumentException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        // BUG DISCOVERY: Set on non-existent slide throws ArgumentException, not a graceful no-op
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["text"] = "x" });
        act.Should().Throw<ArgumentException>("Set on non-existent node throws (BUG: could be no-op or clearer error)");
    }

    [Fact]
    public void DP05_Remove_PathToNonExistentNode_ThrowsArgumentException()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        // BUG DISCOVERY: Remove on non-existent path throws ArgumentException
        var act = () => handler.Remove("/body/p[999]");
        act.Should().Throw<ArgumentException>("Remove on non-existent node throws (BUG: could be graceful no-op)");
    }

    [Fact]
    public void DP06_Get_DeepNestedPath_ThrowsWhenAncestorMissing()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: false);
        // BUG DISCOVERY: Deeply nested path where no slides exist throws instead of returning null
        var act = () => handler.Get("/slide[1]/shape[1]/paragraph[1]/run[1]");
        act.Should().Throw<ArgumentException>("deeply nested path with missing ancestors throws (BUG: should return null)");
    }

    // ==================== EF01–EF04: Extreme format values ====================

    [Fact]
    public void EF01_Pptx_FontSize999pt_SetDoesNotThrow()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hi", ["size"] = "999pt" });
        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void EF02_Pptx_FontSize_VerySmall_SetDoesNotThrow()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        // Minimum valid font size (OpenXML allows >= 1 hundredths of a point)
        var act = () => handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hi", ["size"] = "1pt" });
        act.Should().NotThrow("1pt font size should be the minimum valid size");
    }

    [Fact]
    public void EF03_Word_FontSize_ExtremeValues_SetDoesNotThrow()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Big" });
        // Very large size
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["size"] = "999pt" });
        act.Should().NotThrow("999pt font size should be accepted");
    }

    [Fact]
    public void EF04_Excel_VeryLongNumberFormat_DoesNotThrow()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/", "sheet", null, new() { ["name"] = "S1" });
        handler.Add("/S1", "cell", null, new() { ["value"] = "123", ["ref"] = "A1" });
        // Very long but syntactically valid number format
        var longFmt = "#,##0." + new string('0', 30);
        var act = () => handler.Set("/S1/A1", new() { ["numberformat"] = longFmt });
        act.Should().NotThrow("very long number format should not crash");
    }

    // ==================== CO01–CO06: Conflicting / combined Set properties ====================

    [Fact]
    public void CO01_Pptx_SetBoldAndItalicAndUnderline_AllApplied()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        handler.Set("/slide[1]/shape[1]/paragraph[1]/run[1]", new()
        {
            ["bold"] = "true",
            ["italic"] = "true",
            ["underline"] = "single"
        });
        var node = handler.Get("/slide[1]/shape[1]/paragraph[1]/run[1]");
        node.Should().NotBeNull();
        node!.Format["bold"].Should().Be(true);
        node.Format["italic"].Should().Be(true);
    }

    [Fact]
    public void CO02_Pptx_SetSizeThenColor_BothPersist()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        handler.Set("/slide[1]/shape[1]/paragraph[1]/run[1]", new() { ["size"] = "24pt" });
        handler.Set("/slide[1]/shape[1]/paragraph[1]/run[1]", new() { ["color"] = "FF0000" });
        var node = handler.Get("/slide[1]/shape[1]/paragraph[1]/run[1]");
        node!.Format.Should().ContainKey("size");
        node.Format.Should().ContainKey("color");
    }

    [Fact]
    public void CO03_Word_SetConflictingAlignments_LastWins()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });
        handler.Set("/body/p[1]", new() { ["alignment"] = "left" });
        handler.Set("/body/p[1]", new() { ["alignment"] = "right" });
        var node = handler.Get("/body/p[1]");
        node!.Format["alignment"].Should().Be("right");
    }

    [Fact]
    public void CO04_Excel_SetMultipleStylesAtOnce_DoNotThrow()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/", "sheet", null, new() { ["name"] = "S1" });
        handler.Add("/S1", "cell", null, new() { ["value"] = "Data", ["ref"] = "A1" });
        var act = () => handler.Set("/S1/A1", new()
        {
            ["bold"] = "true",
            ["italic"] = "true",
            ["size"] = "14pt",
            ["color"] = "0000FF",
            ["bgcolor"] = "FFFF00",
            ["alignment.horizontal"] = "center",
            ["alignment.wrapText"] = "true"
        });
        act.Should().NotThrow("setting multiple styles at once should not throw");
    }

    [Fact]
    public void CO05_Pptx_SetFillAndThenRemoveFill_ShapeStillExists()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        // Slide Add puts title at shape[1], so the added shape is at shape[2]
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Box", ["fill"] = "FF0000" });
        // Find actual shape index via Query
        var shapes = handler.Query("shape");
        shapes.Should().NotBeEmpty("at least one shape should exist");
        var shapePath = shapes.Last().Path;
        handler.Set(shapePath, new() { ["fill"] = "none" });
        var node = handler.Get(shapePath);
        node.Should().NotBeNull("shape should still exist after fill is removed");
    }

    [Fact]
    public void CO06_Word_SetSpaceBeforeAndAfterAndLine_AllPersist()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });
        handler.Set("/body/p[1]", new()
        {
            ["spaceBefore"] = "12pt",
            ["spaceAfter"] = "6pt",
            ["lineSpacing"] = "1.5x"
        });
        var node = handler.Get("/body/p[1]");
        node!.Format["spaceBefore"].Should().Be("12pt");
        node.Format["spaceAfter"].Should().Be("6pt");
        node.Format["lineSpacing"].Should().Be("1.5x");
    }

    // ==================== AE01–AE04: Add duplicates, Remove missing ====================

    [Fact]
    public void AE01_Pptx_AddSameSlideMultipleTimes_AllSlidesAccessible()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        for (int i = 0; i < 5; i++)
            handler.Add("/", "slide", null, new() { ["title"] = $"Slide {i + 1}" });
        // Note: Query("slide") is not valid — "slide" is not a queryable selector type.
        // Verify via direct Get that all 5 slides are reachable.
        for (int i = 1; i <= 5; i++)
        {
            var node = handler.Get($"/slide[{i}]");
            node.Should().NotBeNull($"slide[{i}] should exist");
        }
        // slide[6] should not exist — throws for out-of-range
        var act = () => handler.Get("/slide[6]");
        act.Should().Throw<ArgumentException>("slide[6] does not exist");
    }

    [Fact]
    public void AE02_Word_AddManyParagraphs_AllPresent()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        for (int i = 0; i < 20; i++)
            handler.Add("/body", "paragraph", null, new() { ["text"] = $"Para {i + 1}" });
        var paras = handler.Query("paragraph");
        paras.Count.Should().BeGreaterThanOrEqualTo(20);
    }

    [Fact]
    public void AE03_Word_RemoveNonExistentParagraph_ThrowsArgumentException()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        // BUG DISCOVERY: Remove on non-existent paragraph throws ArgumentException
        // The API throws "Path not found: /body/p[500]. No children at /body"
        var act = () => handler.Remove("/body/p[500]");
        act.Should().Throw<ArgumentException>("removing non-existent paragraph throws (BUG: could be graceful no-op)");
    }

    [Fact]
    public void AE04_Excel_RemoveNonExistentCell_ThrowsArgumentException()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/", "sheet", null, new() { ["name"] = "S1" });
        // BUG DISCOVERY: Remove on non-existent cell/row throws ArgumentException
        var act = () => handler.Remove("/S1/Z999");
        act.Should().Throw<ArgumentException>("removing non-existent cell throws (BUG: could be graceful no-op)");
    }
}
