// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Tests.Functional;

// ==================== PPTX Table Enhanced Tests ====================

public class PptxTableEnhancedTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxTableEnhancedTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
        _handler.Add("/", "slide", null, new() { ["layout"] = "blank" });
        _handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "4", ["cols"] = "3", ["style"] = "medium2"
        });
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    // ==================== 1. TableLook Flags ====================

    [Fact]
    public void TableLook_DefaultAfterAdd_FirstRowAndBandRowAreTrue()
    {
        var node = _handler.Get("/slide[1]/table[1]");
        // Default from Add: FirstRow=true, BandRow=true
        node.Format["firstRow"].Should().Be(true);
        node.Format["bandedRows"].Should().Be(true);
    }

    [Fact]
    public void Set_AllTableLookFlags_On_ReadBack()
    {
        _handler.Set("/slide[1]/table[1]", new()
        {
            ["firstRow"] = "true", ["lastRow"] = "true",
            ["firstCol"] = "true", ["lastCol"] = "true",
            ["bandedRows"] = "true", ["bandedCols"] = "true"
        });
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["firstRow"].Should().Be(true);
        node.Format["lastRow"].Should().Be(true);
        node.Format["firstCol"].Should().Be(true);
        node.Format["lastCol"].Should().Be(true);
        node.Format["bandedRows"].Should().Be(true);
        node.Format["bandedCols"].Should().Be(true);
    }

    [Fact]
    public void Set_AllTableLookFlags_Off_ReadBack()
    {
        _handler.Set("/slide[1]/table[1]", new()
        {
            ["firstRow"] = "false", ["lastRow"] = "false",
            ["firstCol"] = "false", ["lastCol"] = "false",
            ["bandedRows"] = "false", ["bandedCols"] = "false"
        });
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["firstRow"].Should().Be(false);
        node.Format["lastRow"].Should().Be(false);
        node.Format["firstCol"].Should().Be(false);
        node.Format["lastCol"].Should().Be(false);
        node.Format["bandedRows"].Should().Be(false);
        node.Format["bandedCols"].Should().Be(false);
    }

    [Fact]
    public void Set_TableLookFlags_Toggle_UpdatesCorrectly()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["firstRow"] = "false" });
        _handler.Get("/slide[1]/table[1]").Format["firstRow"].Should().Be(false);
        _handler.Set("/slide[1]/table[1]", new() { ["firstRow"] = "true" });
        _handler.Get("/slide[1]/table[1]").Format["firstRow"].Should().Be(true);
    }

    [Fact]
    public void Set_TableLookFlags_Persist()
    {
        _handler.Set("/slide[1]/table[1]", new()
        {
            ["firstRow"] = "false", ["lastRow"] = "true",
            ["firstCol"] = "true", ["bandedCols"] = "true"
        });
        Reopen();
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["firstRow"].Should().Be(false);
        node.Format["lastRow"].Should().Be(true);
        node.Format["firstCol"].Should().Be(true);
        node.Format["bandedCols"].Should().Be(true);
    }

    // ==================== 2. Column Widths ====================

    [Fact]
    public void Set_ColWidths_DifferentValues_AppliesCorrectly()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["colWidths"] = "2cm,6cm,4cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var gridCols = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .TableGrid!.Elements<Drawing.GridColumn>().ToList();
        // 2cm≈720000, 6cm≈2160000, 4cm≈1440000
        gridCols[0].Width!.Value.Should().BeInRange(710000, 730000);
        gridCols[1].Width!.Value.Should().BeInRange(2150000, 2170000);
        gridCols[2].Width!.Value.Should().BeInRange(1430000, 1450000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColWidths_SingleValue_AppliesAllColumns()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["colWidths"] = "4cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var gridCols = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .TableGrid!.Elements<Drawing.GridColumn>().ToList();
        foreach (var gc in gridCols)
            gc.Width!.Value.Should().BeInRange(1430000, 1450000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColWidths_Persist()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["colWidths"] = "3cm,5cm,3cm" });
        Reopen();
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var gridCols = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .TableGrid!.Elements<Drawing.GridColumn>().ToList();
        gridCols[1].Width!.Value.Should().BeGreaterThan(gridCols[0].Width!.Value);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 3. Table Shadow ====================

    [Fact]
    public void Set_TableShadow_CreatesOuterShadow()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["shadow"] = "000000-4-135-3-50" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tblPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .GetFirstChild<Drawing.TableProperties>()!;
        tblPr.GetFirstChild<Drawing.EffectList>()!
            .GetFirstChild<Drawing.OuterShadow>().Should().NotBeNull();
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TableShadow_None_RemovesEffect()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["shadow"] = "000000-4-135-3-50" });
        _handler.Set("/slide[1]/table[1]", new() { ["shadow"] = "none" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tblPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .GetFirstChild<Drawing.TableProperties>()!;
        var effectList = tblPr.GetFirstChild<Drawing.EffectList>();
        if (effectList != null)
            effectList.GetFirstChild<Drawing.OuterShadow>().Should().BeNull();
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TableShadow_Persist()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["shadow"] = "333333-6-180-4-70" });
        Reopen();
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tblPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .GetFirstChild<Drawing.TableProperties>()!;
        tblPr.GetFirstChild<Drawing.EffectList>()!
            .GetFirstChild<Drawing.OuterShadow>().Should().NotBeNull();
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 4. Table Glow ====================

    [Fact]
    public void Set_TableGlow_CreatesGlowEffect()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["glow"] = "4472C4-10-60" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tblPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .GetFirstChild<Drawing.TableProperties>()!;
        tblPr.GetFirstChild<Drawing.EffectList>()!
            .GetFirstChild<Drawing.Glow>().Should().NotBeNull();
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TableGlow_None_RemovesEffect()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["glow"] = "4472C4-10-60" });
        _handler.Set("/slide[1]/table[1]", new() { ["glow"] = "none" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tblPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .GetFirstChild<Drawing.TableProperties>()!;
        var effectList = tblPr.GetFirstChild<Drawing.EffectList>();
        if (effectList != null)
            effectList.GetFirstChild<Drawing.Glow>().Should().BeNull();
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 5. Custom Banded Row Colors ====================

    [Fact]
    public void Set_BandColorOdd_OnlyAffectsOddRows()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["bandColor.odd"] = "E8F0FE" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var rows = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .Elements<Drawing.TableRow>().ToList();
        // Rows 0,2 (odd in 0-based) should have fill
        rows[0].Elements<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>().Should().NotBeNull();
        rows[2].Elements<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>().Should().NotBeNull();
        // Rows 1,3 (even in 0-based) should not have fill
        rows[1].Elements<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>().Should().BeNull();
        rows[3].Elements<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>().Should().BeNull();
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_BandColorEven_OnlyAffectsEvenRows()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["bandColor.even"] = "FFF3CD" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var rows = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .Elements<Drawing.TableRow>().ToList();
        // Rows 1,3 should have fill
        rows[1].Elements<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>().Should().NotBeNull();
        rows[3].Elements<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>().Should().NotBeNull();
        // Rows 0,2 should not
        rows[0].Elements<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>().Should().BeNull();
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 6. Autofit ====================

    [Fact]
    public void Set_Autofit_WidensColumnWithLongestText()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "X" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[2]", new() { ["text"] = "This column has much more text content" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[3]", new() { ["text"] = "Med" });
        _handler.Set("/slide[1]/table[1]", new() { ["autofit"] = "true" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var gridCols = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.Table>().First()
            .TableGrid!.Elements<Drawing.GridColumn>().ToList();
        gridCols[1].Width!.Value.Should().BeGreaterThan(gridCols[0].Width!.Value);
        gridCols[1].Width!.Value.Should().BeGreaterThan(gridCols[2].Width!.Value);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 7. Data Import ====================

    [Fact]
    public void Add_Table_WithInlineData_CorrectDimensions()
    {
        var path = _handler.Add("/slide[1]", "table", null, new()
        {
            ["data"] = "Name,Age,City;Alice,30,NYC;Bob,25,LA;Charlie,35,SF"
        });
        path.Should().Be("/slide[1]/table[2]");
        var node = _handler.Get(path, depth: 2);
        node.Format["rows"].Should().Be(4);
        node.Format["cols"].Should().Be(3);
    }

    [Fact]
    public void Add_Table_WithInlineData_CellsPopulated()
    {
        _handler.Add("/slide[1]", "table", null, new()
        {
            ["data"] = "H1,H2;V1,V2"
        });
        var node = _handler.Get("/slide[1]/table[2]", depth: 2);
        node.Children[0].Children[0].Text.Should().Be("H1");
        node.Children[0].Children[1].Text.Should().Be("H2");
        node.Children[1].Children[0].Text.Should().Be("V1");
        node.Children[1].Children[1].Text.Should().Be("V2");
    }

    // ==================== 8. Cell Margin / Padding ====================

    [Fact]
    public void Set_CellMargin_SingleValue_AllSides()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin"] = "0.3cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tcPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TableCellProperties!;
        var expected = 108000; // 0.3cm ≈ 108000 EMU
        tcPr.LeftMargin!.Value.Should().BeInRange(expected - 5000, expected + 5000);
        tcPr.RightMargin!.Value.Should().BeInRange(expected - 5000, expected + 5000);
        tcPr.TopMargin!.Value.Should().BeInRange(expected - 5000, expected + 5000);
        tcPr.BottomMargin!.Value.Should().BeInRange(expected - 5000, expected + 5000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_CellPadding_TwoValues_HorizontalVertical()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["padding"] = "0.5cm,0.2cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tcPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TableCellProperties!;
        tcPr.LeftMargin!.Value.Should().BeInRange(175000, 185000);  // 0.5cm
        tcPr.RightMargin!.Value.Should().BeInRange(175000, 185000); // 0.5cm
        tcPr.TopMargin!.Value.Should().BeInRange(68000, 76000);     // 0.2cm
        tcPr.BottomMargin!.Value.Should().BeInRange(68000, 76000);  // 0.2cm
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_CellMargin_IndividualSides()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["margin.left"] = "0.1cm", ["margin.top"] = "0.2cm",
            ["margin.right"] = "0.3cm", ["margin.bottom"] = "0.4cm"
        });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tcPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TableCellProperties!;
        tcPr.LeftMargin!.Value.Should().BeLessThan(tcPr.TopMargin!.Value);
        tcPr.TopMargin!.Value.Should().BeLessThan(tcPr.RightMargin!.Value);
        tcPr.RightMargin!.Value.Should().BeLessThan(tcPr.BottomMargin!.Value);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 9. Text Direction ====================

    [Fact]
    public void Set_TextDirection_Vertical270()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "vertical" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TableCellProperties!
            .Vertical!.Value.Should().Be(Drawing.TextVerticalValues.Vertical270);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TextDirection_Vertical90()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "vertical90" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TableCellProperties!
            .Vertical!.Value.Should().Be(Drawing.TextVerticalValues.Vertical);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TextDirection_Stacked()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "stacked" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TableCellProperties!
            .Vertical!.Value.Should().Be(Drawing.TextVerticalValues.WordArtVertical);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TextDirection_Horizontal_Resets()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "vertical" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "horizontal" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tcPr = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TableCellProperties!;
        tcPr.Vertical!.Value.Should().Be(Drawing.TextVerticalValues.Horizontal);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 10. Word Wrap ====================

    [Fact]
    public void Set_WordWrap_False_ThenTrue_Toggles()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["wordWrap"] = "false" });
        _handler.Dispose();
        using var doc1 = PresentationDocument.Open(_path, false);
        doc1.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TextBody!
            .GetFirstChild<Drawing.BodyProperties>()!.Wrap!.Value
            .Should().Be(Drawing.TextWrappingValues.None);
        doc1.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);

        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["wordWrap"] = "true" });
        _handler.Dispose();
        using var doc2 = PresentationDocument.Open(_path, false);
        doc2.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TextBody!
            .GetFirstChild<Drawing.BodyProperties>()!.Wrap!.Value
            .Should().Be(Drawing.TextWrappingValues.Square);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 11. Line Spacing ====================

    [Theory]
    [InlineData("1.5x", true, 150000)]
    [InlineData("200%", true, 200000)]
    public void Set_LineSpacing_Multiplier(string input, bool isPercent, int expectedVal)
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Test", ["lineSpacing"] = input });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var para = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TextBody!
            .Elements<Drawing.Paragraph>().First();
        var ls = para.ParagraphProperties!.GetFirstChild<Drawing.LineSpacing>()!;
        if (isPercent)
            ls.GetFirstChild<Drawing.SpacingPercent>()!.Val!.Value.Should().Be(expectedVal);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_LineSpacing_Fixed_Points()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Test", ["lineSpacing"] = "18pt" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var para = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TextBody!
            .Elements<Drawing.Paragraph>().First();
        var ls = para.ParagraphProperties!.GetFirstChild<Drawing.LineSpacing>()!;
        ls.GetFirstChild<Drawing.SpacingPoints>()!.Val!.Value.Should().Be(1800);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 12. Space Before / After ====================

    [Fact]
    public void Set_SpaceBefore_Cm()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "T", ["spaceBefore"] = "0.5cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var para = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TextBody!
            .Elements<Drawing.Paragraph>().First();
        var sb = para.ParagraphProperties!.GetFirstChild<Drawing.SpaceBefore>()!;
        // 0.5cm ≈ 14.17pt → 1417 hundredths
        sb.GetFirstChild<Drawing.SpacingPoints>()!.Val!.Value.Should().BeInRange(1410, 1425);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_SpaceAfter_Pt()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "T", ["spaceAfter"] = "10pt" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var para = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First().TextBody!
            .Elements<Drawing.Paragraph>().First();
        var sa = para.ParagraphProperties!.GetFirstChild<Drawing.SpaceAfter>()!;
        sa.GetFirstChild<Drawing.SpacingPoints>()!.Val!.Value.Should().Be(1000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 13. Cell Opacity ====================

    [Fact]
    public void Set_CellOpacity_AddsAlphaElement()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "4472C4" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "75" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var solid = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>()!;
        solid.GetFirstChild<Drawing.RgbColorModelHex>()!
            .GetFirstChild<Drawing.Alpha>()!.Val!.Value.Should().Be(75000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_CellOpacity_FullOpaque()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "100" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var solid = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.SolidFill>()!;
        solid.GetFirstChild<Drawing.RgbColorModelHex>()!
            .GetFirstChild<Drawing.Alpha>()!.Val!.Value.Should().Be(100000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== 14. Cell Bevel ====================

    [Fact]
    public void Set_CellBevel_Circle()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.Cell3DProperties>()!
            .GetFirstChild<Drawing.Bevel>()!.Preset!.Value.Should().Be(Drawing.BevelPresetValues.Circle);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_CellBevel_SoftRound()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "softRound" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.Cell3DProperties>()!
            .GetFirstChild<Drawing.Bevel>()!.Preset!.Value.Should().Be(Drawing.BevelPresetValues.SoftRound);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_CellBevel_WithDimensions()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle-4-3" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var bevel = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.Cell3DProperties>()!
            .GetFirstChild<Drawing.Bevel>()!;
        bevel.Width!.Value.Should().Be(50800);  // 4pt * 12700
        bevel.Height!.Value.Should().Be(38100); // 3pt * 12700
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_CellBevel_None_Removes()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "none" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First()
            .TableCellProperties!.GetFirstChild<Drawing.Cell3DProperties>()
            .Should().BeNull();
        _handler = new PowerPointHandler(_path, editable: true);
    }
}

// ==================== DOCX Table Enhanced Tests ====================

public class DocxTableEnhancedTests : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public DocxTableEnhancedTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
        _handler.Add("/body", "table", null, new()
        {
            ["rows"] = "4", ["cols"] = "3", ["style"] = "TableGrid"
        });
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== 15. TableLook Flags ====================

    [Fact]
    public void Set_FirstRow_True()
    {
        _handler.Set("/body/tbl[1]", new() { ["firstRow"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableLook>().First().FirstRow!.Value.Should().BeTrue();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_LastRow_True()
    {
        _handler.Set("/body/tbl[1]", new() { ["lastRow"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableLook>().First().LastRow!.Value.Should().BeTrue();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_FirstCol_True()
    {
        _handler.Set("/body/tbl[1]", new() { ["firstCol"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableLook>().First().FirstColumn!.Value.Should().BeTrue();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_LastCol_True()
    {
        _handler.Set("/body/tbl[1]", new() { ["lastCol"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableLook>().First().LastColumn!.Value.Should().BeTrue();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_BandedRows_True_SetsNoHBandFalse()
    {
        _handler.Set("/body/tbl[1]", new() { ["bandedRows"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableLook>().First();
        tblLook.NoHorizontalBand!.Value.Should().BeFalse();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_BandedRows_False_SetsNoHBandTrue()
    {
        _handler.Set("/body/tbl[1]", new() { ["bandedRows"] = "false" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableLook>().First();
        tblLook.NoHorizontalBand!.Value.Should().BeTrue();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_BandedCols_True_SetsNoVBandFalse()
    {
        _handler.Set("/body/tbl[1]", new() { ["bandedCols"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableLook>().First();
        tblLook.NoVerticalBand!.Value.Should().BeFalse();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TableLook_AllFlags_Persist()
    {
        _handler.Set("/body/tbl[1]", new()
        {
            ["firstRow"] = "true", ["lastRow"] = "true",
            ["firstCol"] = "true", ["lastCol"] = "true",
            ["bandedRows"] = "true", ["bandedCols"] = "true"
        });
        Reopen();
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableLook>().First();
        tblLook.FirstRow!.Value.Should().BeTrue();
        tblLook.LastRow!.Value.Should().BeTrue();
        tblLook.FirstColumn!.Value.Should().BeTrue();
        tblLook.LastColumn!.Value.Should().BeTrue();
        tblLook.NoHorizontalBand!.Value.Should().BeFalse();
        tblLook.NoVerticalBand!.Value.Should().BeFalse();
        _handler = new WordHandler(_path, editable: true);
    }

    // ==================== 16. Caption & Description ====================

    [Fact]
    public void Set_Caption_CreatesElement()
    {
        _handler.Set("/body/tbl[1]", new() { ["caption"] = "Revenue Table" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableCaption>().First().Val!.Value
            .Should().Be("Revenue Table");
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_Caption_Update_ReplacesExisting()
    {
        _handler.Set("/body/tbl[1]", new() { ["caption"] = "Old" });
        _handler.Set("/body/tbl[1]", new() { ["caption"] = "New" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var captions = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableCaption>().ToList();
        captions.Should().HaveCount(1);
        captions[0].Val!.Value.Should().Be("New");
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_Caption_Empty_RemovesElement()
    {
        _handler.Set("/body/tbl[1]", new() { ["caption"] = "Temp" });
        _handler.Set("/body/tbl[1]", new() { ["caption"] = "" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableCaption>().Should().BeEmpty();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_Description_CreatesElement()
    {
        _handler.Set("/body/tbl[1]", new() { ["description"] = "Quarterly sales data by region" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableDescription>().First().Val!.Value
            .Should().Be("Quarterly sales data by region");
        _handler = new WordHandler(_path, editable: true);
    }

    // ==================== 17. Floating Table Position ====================

    [Fact]
    public void Set_Position_Floating_EnablesFloating()
    {
        _handler.Set("/body/tbl[1]", new() { ["position"] = "floating" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().Should().NotBeEmpty();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_Position_None_DisablesFloating()
    {
        _handler.Set("/body/tbl[1]", new() { ["position"] = "floating" });
        _handler.Set("/body/tbl[1]", new() { ["position"] = "none" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().Should().BeEmpty();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_PositionX_AbsoluteTwips()
    {
        _handler.Set("/body/tbl[1]", new() { ["position.x"] = "3cm" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tpp = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().First();
        tpp.TablePositionX!.Value.Should().BeInRange(1698, 1704); // 3cm ≈ 1701 twips
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_PositionX_Alignment()
    {
        _handler.Set("/body/tbl[1]", new() { ["position.x"] = "center" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tpp = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().First();
        tpp.TablePositionXAlignment!.Value.Should().Be(W.HorizontalAlignmentValues.Center);
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_PositionY_AbsoluteAndAlignment()
    {
        _handler.Set("/body/tbl[1]", new() { ["position.y"] = "5cm" });
        _handler.Dispose();
        using var doc1 = WordprocessingDocument.Open(_path, false);
        doc1.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().First()
            .TablePositionY!.Value.Should().BeInRange(2833, 2838);
        doc1.Dispose();
        _handler = new WordHandler(_path, editable: true);

        _handler.Set("/body/tbl[1]", new() { ["position.y"] = "bottom" });
        _handler.Dispose();
        using var doc2 = WordprocessingDocument.Open(_path, false);
        doc2.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().First()
            .TablePositionYAlignment!.Value.Should().Be(W.VerticalAlignmentValues.Bottom);
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_PositionAnchors_SetCorrectly()
    {
        _handler.Set("/body/tbl[1]", new()
        {
            ["position.hAnchor"] = "margin", ["position.vAnchor"] = "text"
        });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tpp = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().First();
        tpp.HorizontalAnchor!.Value.Should().Be(W.HorizontalAnchorValues.Margin);
        tpp.VerticalAnchor!.Value.Should().Be(W.VerticalAnchorValues.Text);
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_PositionFromText_AllSides()
    {
        _handler.Set("/body/tbl[1]", new()
        {
            ["position"] = "floating",
            ["position.left"] = "0.5cm",
            ["position.right"] = "0.5cm",
            ["position.top"] = "0.3cm",
            ["position.bottom"] = "0.3cm"
        });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tpp = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().First();
        tpp.LeftFromText!.Value.Should().BeInRange((short)280, (short)290);
        tpp.RightFromText!.Value.Should().BeInRange((short)280, (short)290);
        tpp.TopFromText!.Value.Should().BeInRange((short)168, (short)172);
        tpp.BottomFromText!.Value.Should().BeInRange((short)168, (short)172);
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_FloatingTable_Persist()
    {
        _handler.Set("/body/tbl[1]", new()
        {
            ["position.x"] = "2cm", ["position.y"] = "4cm",
            ["position.hAnchor"] = "page", ["position.vAnchor"] = "page"
        });
        Reopen();
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tpp = doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TablePositionProperties>().First();
        tpp.TablePositionX!.Value.Should().BeInRange(1130, 1140);
        tpp.TablePositionY!.Value.Should().BeInRange(2265, 2275);
        _handler = new WordHandler(_path, editable: true);
    }

    // ==================== 18. Overlap ====================

    [Fact]
    public void Set_Overlap_Never()
    {
        _handler.Set("/body/tbl[1]", new() { ["position"] = "floating", ["overlap"] = "never" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableOverlap>().First()
            .Val!.Value.Should().Be(W.TableOverlapValues.Never);
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_Overlap_Overlap()
    {
        _handler.Set("/body/tbl[1]", new() { ["position"] = "floating", ["overlap"] = "overlap" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.TableOverlap>().First()
            .Val!.Value.Should().Be(W.TableOverlapValues.Overlap);
        _handler = new WordHandler(_path, editable: true);
    }

    // ==================== 19. Data Import ====================

    [Fact]
    public void Add_Table_WithInlineData_PopulatesCells()
    {
        _handler.Add("/body", "table", null, new()
        {
            ["data"] = "Product,Price;Widget,9.99;Gadget,19.99"
        });
        var node = _handler.Get("/body/tbl[2]", depth: 2);
        node.Type.Should().Be("table");
        node.Children[0].Children[0].Text.Should().Be("Product");
        node.Children[0].Children[1].Text.Should().Be("Price");
        node.Children[1].Children[0].Text.Should().Be("Widget");
    }

    // ==================== 20. FitText ====================

    [Fact]
    public void Set_FitText_True_WithText_AddsFitTextToRun()
    {
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Long text content" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.FitText>().Should().NotBeEmpty();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_FitText_False_Removes()
    {
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Text" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "true" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "false" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<W.FitText>().Should().BeEmpty();
        _handler = new WordHandler(_path, editable: true);
    }
}

// ==================== XLSX Table Enhanced Tests ====================

public class ExcelTableEnhancedTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelTableEnhancedTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Name" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "Revenue" });
        _handler.Set("/Sheet1/C1", new() { ["value"] = "Quantity" });
        for (int r = 2; r <= 5; r++)
        {
            _handler.Set($"/Sheet1/A{r}", new() { ["value"] = $"Item{r - 1}" });
            _handler.Set($"/Sheet1/B{r}", new() { ["value"] = (r * 100).ToString() });
            _handler.Set($"/Sheet1/C{r}", new() { ["value"] = (r * 10).ToString() });
        }
        _handler.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:C5", ["name"] = "TestData" });
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== 21. Style Options ====================

    [Fact]
    public void Set_ShowRowStripes_False_ReadBack()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["showRowStripes"] = "false" });
        _handler.Get("/Sheet1/table[1]").Format["showRowStripes"].Should().Be(false);
    }

    [Fact]
    public void Set_ShowColumnStripes_True_ReadBack()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["showColumnStripes"] = "true" });
        _handler.Get("/Sheet1/table[1]").Format["showColumnStripes"].Should().Be(true);
    }

    [Fact]
    public void Set_ShowFirstColumn_True_ReadBack()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["showFirstColumn"] = "true" });
        _handler.Get("/Sheet1/table[1]").Format["showFirstColumn"].Should().Be(true);
    }

    [Fact]
    public void Set_ShowLastColumn_True_ReadBack()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["showLastColumn"] = "true" });
        _handler.Get("/Sheet1/table[1]").Format["showLastColumn"].Should().Be(true);
    }

    [Fact]
    public void Set_AllStyleOptions_Persist()
    {
        _handler.Set("/Sheet1/table[1]", new()
        {
            ["showRowStripes"] = "false", ["showColumnStripes"] = "true",
            ["showFirstColumn"] = "true", ["showLastColumn"] = "true"
        });
        Reopen();
        var node = _handler.Get("/Sheet1/table[1]");
        node.Format["showRowStripes"].Should().Be(false);
        node.Format["showColumnStripes"].Should().Be(true);
        node.Format["showFirstColumn"].Should().Be(true);
        node.Format["showLastColumn"].Should().Be(true);
    }

    [Fact]
    public void Set_BandedRows_Alias_Works()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["bandedRows"] = "false" });
        _handler.Get("/Sheet1/table[1]").Format["showRowStripes"].Should().Be(false);
    }

    // ==================== 22. Column Name ====================

    [Fact]
    public void Set_ColumnName_UpdatesAndPersists()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[1].name"] = "EmployeeName" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var col = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().First();
        col.Name!.Value.Should().Be("EmployeeName");
        _handler = new ExcelHandler(_path, editable: true);
    }

    // ==================== 23. Total Function ====================

    [Fact]
    public void Set_ColumnTotalFunction_Sum()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[2].totalFunction"] = "sum" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        doc.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ElementAt(1)
            .TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Sum);
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnTotalFunction_Count()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[2].totalFunction"] = "count" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        doc.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ElementAt(1)
            .TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Count);
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnTotalFunction_Average()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[2].totalFunction"] = "average" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        doc.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ElementAt(1)
            .TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Average);
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnTotalFunction_Max()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[2].totalFunction"] = "max" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        doc.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ElementAt(1)
            .TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Maximum);
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnTotalFunction_Min()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[2].totalFunction"] = "min" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        doc.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ElementAt(1)
            .TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Minimum);
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnTotalFunction_None()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[2].totalFunction"] = "sum" });
        _handler.Set("/Sheet1/table[1]", new() { ["col[2].totalFunction"] = "none" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        doc.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ElementAt(1)
            .TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.None);
        _handler = new ExcelHandler(_path, editable: true);
    }

    // ==================== 24. Total Label ====================

    [Fact]
    public void Set_ColumnTotalLabel_SetsLabel()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[1].totalLabel"] = "Grand Total" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var col = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().First();
        col.TotalsRowLabel!.Value.Should().Be("Grand Total");
        _handler = new ExcelHandler(_path, editable: true);
    }

    // ==================== 25. Column Formula ====================

    [Fact]
    public void Set_ColumnFormula_SetsCalculatedColumnFormula()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[3].formula"] = "[Revenue]*2" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var col = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ElementAt(2);
        col.CalculatedColumnFormula.Should().NotBeNull();
        col.CalculatedColumnFormula!.Text.Should().Be("[Revenue]*2");
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnFormula_Persist()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[3].formula"] = "[Revenue]+[Quantity]" });
        Reopen();
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var col = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ElementAt(2);
        col.CalculatedColumnFormula!.Text.Should().Be("[Revenue]+[Quantity]");
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_MultipleColumnProperties_Combined()
    {
        _handler.Set("/Sheet1/table[1]", new()
        {
            ["col[1].totalLabel"] = "Total",
            ["col[2].totalFunction"] = "sum",
            ["col[3].totalFunction"] = "average"
        });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var cols = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!
            .GetFirstChild<TableColumns>()!.Elements<TableColumn>().ToList();
        cols[0].TotalsRowLabel!.Value.Should().Be("Total");
        cols[1].TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Sum);
        cols[2].TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Average);
        _handler = new ExcelHandler(_path, editable: true);
    }
}
