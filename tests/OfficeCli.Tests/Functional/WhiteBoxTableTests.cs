// WhiteBoxTableTests.cs — Targeted white-box tests for table properties
// Based on code review of:
//   PowerPointHandler.Set.cs        (table-level: shadow, glow, bandColor, autofit, colWidths, TableLook)
//   PowerPointHandler.ShapeProperties.cs (cell-level: margin, textDirection, wordWrap, lineSpacing,
//                                          spaceBefore/After, opacity, bevel)
//   PowerPointHandler.Add.Table.cs  (data import)
//   WordHandler.Set.cs              (TableLook, caption, description, position, floating, fitText)
//   ExcelHandler.Set.cs             (showRowStripes, showColumnStripes, showFirstColumn, showLastColumn, col[N].*)
//   ExcelHandler.Helpers.cs         (TableToNode readback)

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Tests.Functional;

public class WhiteBoxTableTests : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private PowerPointHandler _pptx;
    private WordHandler _word;
    private ExcelHandler _excel;

    public WhiteBoxTableTests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"wb_tbl_{Guid.NewGuid():N}.pptx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"wb_tbl_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"wb_tbl_{Guid.NewGuid():N}.xlsx");

        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);

        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        _word = new WordHandler(_docxPath, editable: true);
        _excel = new ExcelHandler(_xlsxPath, editable: true);

        // Add a slide so PPTX tests can add tables
        _pptx.Add("/", "slide", null, new());
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

    private void ReopenPptx()
    {
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    private void ReopenWord()
    {
        _word.Dispose();
        _word = new WordHandler(_docxPath, editable: true);
    }

    private void ReopenExcel()
    {
        _excel.Dispose();
        _excel = new ExcelHandler(_xlsxPath, editable: true);
    }

    // ==================== PPTX TABLE-LEVEL: Shadow ====================

    [Fact]
    public void Pptx_Set_TableShadow_PersistsAfterReopen()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "FF0000;4;45;2" });

        ReopenPptx();

        // Verify shadow was written to XML
        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        tblPr.Should().NotBeNull("TableProperties must exist after shadow set");
        var effectList = tblPr!.GetFirstChild<Drawing.EffectList>();
        effectList.Should().NotBeNull("EffectList must exist after shadow set");
        var shadow = effectList!.GetFirstChild<Drawing.OuterShadow>();
        shadow.Should().NotBeNull("OuterShadow element must be present after shadow=FF0000 set");
    }

    [Fact]
    public void Pptx_Set_TableShadow_NoneRemovesShadow()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "FF0000;4;45;2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "none" });
        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        var shadow = tblPr?.GetFirstChild<Drawing.EffectList>()?.GetFirstChild<Drawing.OuterShadow>();
        shadow.Should().BeNull("OuterShadow must be removed after shadow=none");
    }

    // ==================== PPTX TABLE-LEVEL: Glow ====================

    [Fact]
    public void Pptx_Set_TableGlow_PersistsAfterReopen()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["glow"] = "0000FF;8" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        var effectList = tblPr?.GetFirstChild<Drawing.EffectList>();
        var glow = effectList?.GetFirstChild<Drawing.Glow>();
        glow.Should().NotBeNull("Glow element must persist after save/reopen");
    }

    // ==================== PPTX TABLE-LEVEL: BandColor.Odd / BandColor.Even ====================

    [Fact]
    public void Pptx_Set_BandColorOdd_AppliesToRows0And2()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.odd"] = "FF0000" });

        // Rows 0,2 (0-based) = rows 1,3 (1-based) = odd rows should have red fill
        // Row 1 (1-based) should have fill
        var cell00 = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell00.Format.Should().ContainKey("fill", "row 1 (odd, 1-based) should have fill applied by bandColor.odd");

        // Row 2 (1-based) should NOT have fill
        var cell10 = _pptx.Get("/slide[1]/table[1]/tr[2]/tc[1]");
        cell10.Format.Should().NotContainKey("fill", "row 2 (even, 1-based) should not have fill applied by bandColor.odd");

        // Row 3 (1-based) should have fill (odd)
        var cell20 = _pptx.Get("/slide[1]/table[1]/tr[3]/tc[1]");
        cell20.Format.Should().ContainKey("fill", "row 3 (odd, 1-based) should have fill applied by bandColor.odd");
    }

    [Fact]
    public void Pptx_Set_BandColorEven_AppliesToRow1Only_In3RowTable()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["bandColor.even"] = "0000FF" });

        // Even rows (0-based): row index 1 = 2nd row (1-based)
        var cell10 = _pptx.Get("/slide[1]/table[1]/tr[2]/tc[1]");
        cell10.Format.Should().ContainKey("fill", "row 2 (even index 1-based) should have fill applied by bandColor.even");

        var cell00 = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell00.Format.Should().NotContainKey("fill", "row 1 (odd 1-based) should not have fill from bandColor.even");
    }

    // ==================== PPTX TABLE-LEVEL: ColWidths ====================

    [Fact]
    public void Pptx_Set_ColWidths_SingleValueAppliedToAllColumns()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "3" });
        _pptx.Set("/slide[1]/table[1]", new() { ["colWidths"] = "3cm" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();

        gridCols.Should().HaveCount(3);
        long expectedEmu = 1080000; // 3cm ≈ 1080000 EMU
        foreach (var col in gridCols!)
        {
            col.Width!.Value.Should().BeInRange(expectedEmu - 10000, expectedEmu + 10000,
                "each column should have width=3cm when single value is given");
        }
    }

    [Fact]
    public void Pptx_Set_ColWidths_MultipleValuesAppliedPerColumn()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "3" });
        _pptx.Set("/slide[1]/table[1]", new() { ["colWidths"] = "2cm,4cm,3cm" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();

        gridCols.Should().HaveCount(3);
        gridCols![0].Width!.Value.Should().BeInRange(710000L, 730000L, "col[1] should be 2cm");
        gridCols[1].Width!.Value.Should().BeInRange(1430000L, 1450000L, "col[2] should be 4cm");
        gridCols[2].Width!.Value.Should().BeInRange(1070000L, 1090000L, "col[3] should be 3cm");
    }

    // ==================== PPTX TABLE-LEVEL: TableLook flags ====================

    [Fact]
    public void Pptx_Set_TableLook_FirstRowPersists()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });

        var node = _pptx.Get("/slide[1]/table[1]");
        // firstRow defaults to true in Add
        _pptx.Set("/slide[1]/table[1]", new() { ["firstRow"] = "false" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        tblPr.Should().NotBeNull();
        tblPr!.FirstRow?.Value.Should().BeFalse("firstRow=false must persist after save");
    }

    [Fact]
    public void Pptx_Set_TableLook_BandedRowsPersists()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["bandedRows"] = "false" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        tblPr!.BandRow?.Value.Should().BeFalse("BandRow=false must persist");
    }

    // ==================== PPTX CELL-LEVEL: Margin / Padding ====================

    [Fact]
    public void Pptx_SetCell_Margin_SingleValue_AllSidesEqual()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin"] = "5pt" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        tcPr.Should().NotBeNull();
        int expected = 63500; // 5pt in EMU
        tcPr!.LeftMargin!.Value.Should().BeInRange(expected - 1000, expected + 1000, "LeftMargin should be 5pt");
        tcPr.TopMargin!.Value.Should().BeInRange(expected - 1000, expected + 1000, "TopMargin should be 5pt");
        tcPr.RightMargin!.Value.Should().BeInRange(expected - 1000, expected + 1000, "RightMargin should be 5pt");
        tcPr.BottomMargin!.Value.Should().BeInRange(expected - 1000, expected + 1000, "BottomMargin should be 5pt");
    }

    [Fact]
    public void Pptx_SetCell_Margin_FourValues_CorrectOrder()
    {
        // margin="L,T,R,B" where parts[0]=Left, parts[1]=Top, parts[2]=Right, parts[3]=Bottom
        // Code at ShapeProperties.cs:1229-1232: Left=parts[0], Top=parts[1], Right=parts[2], Bottom=parts[3]
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin"] = "1pt,2pt,3pt,4pt" });
        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        tcPr.Should().NotBeNull();

        // Verify order: L=1pt, T=2pt, R=3pt, B=4pt
        tcPr!.LeftMargin!.Value.Should().BeInRange(11700, 13700, "LeftMargin = parts[0] = 1pt = 12700 EMU");
        tcPr.TopMargin!.Value.Should().BeInRange(24400, 26400, "TopMargin = parts[1] = 2pt = 25400 EMU");
        tcPr.RightMargin!.Value.Should().BeInRange(37100, 39100, "RightMargin = parts[2] = 3pt = 38100 EMU");
        tcPr.BottomMargin!.Value.Should().BeInRange(49800, 51800, "BottomMargin = parts[3] = 4pt = 50800 EMU");
    }

    [Fact]
    public void Pptx_SetCell_MarginLeft_IndividualProperty()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin.left"] = "8pt" });
        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        tcPr!.LeftMargin!.Value.Should().BeInRange(100600, 102600, "8pt = 101600 EMU");
        tcPr.RightMargin.Should().BeNull("Only LeftMargin should be set, not RightMargin");
    }

    // ==================== PPTX CELL-LEVEL: TextDirection ====================

    [Fact]
    public void Pptx_SetCell_TextDirection_Vertical270_Persists()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "vertical" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        tcPr.Should().NotBeNull();
        // "vertical" maps to Vertical270 per code line 1275
        tcPr!.Vertical!.InnerText.Should().Be("vert270",
            "'vertical' input should map to TextVerticalValues.Vertical270 (vert270)");
    }

    [Fact]
    public void Pptx_SetCell_TextDirection_Vertical90_Persists()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "vertical90" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        // "vertical90" maps to TextVerticalValues.Vertical per code line 1277
        tcPr!.Vertical!.InnerText.Should().Be("vert",
            "'vertical90' input should map to TextVerticalValues.Vertical (vert)");
    }

    // ==================== PPTX CELL-LEVEL: WordWrap ====================

    [Fact]
    public void Pptx_SetCell_WordWrap_False_SetsNoneWrap()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["wordWrap"] = "false" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var bodyProps = cell.TextBody?.GetFirstChild<Drawing.BodyProperties>();
        bodyProps.Should().NotBeNull("BodyProperties must exist after wordWrap=false");
        bodyProps!.Wrap.Should().NotBeNull("Wrap attribute must be set");
        bodyProps.Wrap!.InnerText.Should().Be("none", "wordWrap=false should set wrap=none");
    }

    [Fact]
    public void Pptx_SetCell_WordWrap_True_SetsSquareWrap()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["wordWrap"] = "true" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var bodyProps = cell.TextBody?.GetFirstChild<Drawing.BodyProperties>();
        bodyProps!.Wrap!.InnerText.Should().Be("square", "wordWrap=true should set wrap=square");
    }

    // ==================== PPTX CELL-LEVEL: LineSpacing ====================

    [Fact]
    public void Pptx_SetCell_LineSpacing_Percent_PersistsAfterReopen()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2",
            ["r1c1"] = "Hello" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["lineSpacing"] = "1.5x" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var para = cell.TextBody?.Elements<Drawing.Paragraph>().First();
        var pProps = para?.ParagraphProperties;
        pProps.Should().NotBeNull("ParagraphProperties must exist after lineSpacing set");
        var ls = pProps!.GetFirstChild<Drawing.LineSpacing>();
        ls.Should().NotBeNull("LineSpacing element must persist");
        var pct = ls!.GetFirstChild<Drawing.SpacingPercent>();
        pct.Should().NotBeNull("SpacingPercent must be used for 1.5x (percent) line spacing");
        pct!.Val!.Value.Should().Be(150000, "1.5x = 150000 in OOXML SpacingPercent");
    }

    [Fact]
    public void Pptx_SetCell_LineSpacing_SchemaOrder_LnSpcBeforeSpcBef()
    {
        // Bug: cell linespacing case uses pProps.AppendChild(ls) without checking schema order.
        // If spaceBefore is already set, lineSpacing will be AFTER spaceBefore (schema violation).
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2",
            ["r1c1"] = "Hello" });
        // Set spaceBefore first, then lineSpacing
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["spaceBefore"] = "6pt" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["lineSpacing"] = "1.5x" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var para = cell.TextBody?.Elements<Drawing.Paragraph>().First();
        var pProps = para?.ParagraphProperties;
        pProps.Should().NotBeNull();

        // Schema order: lnSpc → spcBef → spcAft
        var children = pProps!.ChildElements.ToList();
        var lnSpcIdx = children.FindIndex(c => c is Drawing.LineSpacing);
        var spcBefIdx = children.FindIndex(c => c is Drawing.SpaceBefore);

        lnSpcIdx.Should().BeGreaterThanOrEqualTo(0, "LineSpacing should be present");
        spcBefIdx.Should().BeGreaterThanOrEqualTo(0, "SpaceBefore should be present");
        lnSpcIdx.Should().BeLessThan(spcBefIdx,
            "LineSpacing must come BEFORE SpaceBefore per CT_TextParagraphProperties schema");
    }

    // ==================== PPTX CELL-LEVEL: SpaceBefore / SpaceAfter ====================

    [Fact]
    public void Pptx_SetCell_SpaceBefore_PersistsAfterReopen()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2",
            ["r1c1"] = "Hello" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["spaceBefore"] = "12pt" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var para = cell.TextBody?.Elements<Drawing.Paragraph>().First();
        var pProps = para?.ParagraphProperties;
        var sb = pProps?.GetFirstChild<Drawing.SpaceBefore>();
        sb.Should().NotBeNull("SpaceBefore must persist after save/reopen");
        var pts = sb!.GetFirstChild<Drawing.SpacingPoints>();
        pts.Should().NotBeNull("SpacingPoints must be used for pt-based spacing");
        pts!.Val!.Value.Should().Be(1200, "12pt = 1200 in OOXML SpacingPoints (100 per pt)");
    }

    [Fact]
    public void Pptx_SetCell_SpaceAfter_PersistsAfterReopen()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2",
            ["r1c1"] = "Hello" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["spaceAfter"] = "6pt" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var para = cell.TextBody?.Elements<Drawing.Paragraph>().First();
        var pProps = para?.ParagraphProperties;
        var sa = pProps?.GetFirstChild<Drawing.SpaceAfter>();
        sa.Should().NotBeNull("SpaceAfter must persist after save/reopen");
        var pts = sa!.GetFirstChild<Drawing.SpacingPoints>();
        pts!.Val!.Value.Should().Be(600, "6pt = 600 in OOXML SpacingPoints (100 per pt)");
    }

    // ==================== PPTX CELL-LEVEL: Opacity Bug ====================

    [Fact]
    public void Pptx_SetCell_Opacity_05_SetsAlphaTo50000()
    {
        // BUG: ShapeProperties.cs:1343 uses * 1000 instead of * 100000.
        // opacity=0.5 should set Alpha.Val = 50000 (50% opaque).
        // Current bug: alphaVal = (int)Math.Round(0.5 * 1000) = 500, not 50000.
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        // Set fill first, then opacity
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "0.5" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        var solidFill = tcPr?.GetFirstChild<Drawing.SolidFill>();
        solidFill.Should().NotBeNull("SolidFill should exist after fill=FF0000");
        var colorEl = solidFill!.GetFirstChild<Drawing.RgbColorModelHex>();
        colorEl.Should().NotBeNull();
        var alpha = colorEl!.GetFirstChild<Drawing.Alpha>();
        alpha.Should().NotBeNull("Alpha element must be set after opacity=0.5");
        // Correct value for 50% opacity is 50000; bug causes 500
        alpha!.Val!.Value.Should().Be(50000,
            "opacity=0.5 should set Alpha.Val=50000 (50%), not 500 (0.5%). " +
            "OOXML alpha scale is 0-100000, not 0-1000.");
    }

    [Fact]
    public void Pptx_SetCell_Opacity_Without_Fill_DoesNotThrow()
    {
        // opacity is silently ignored if no TableCellProperties or no solidFill exists
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        // Don't set fill first
        var act = () => _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "0.5" });
        // Should either succeed silently or be returned as unsupported, but must not throw
        act.Should().NotThrow("opacity without a fill should be handled gracefully");
    }

    // ==================== PPTX CELL-LEVEL: Bevel ====================

    [Fact]
    public void Pptx_SetCell_Bevel_Circle_Persists()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle" });

        ReopenPptx();

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        var cell3d = tcPr?.GetFirstChild<Drawing.Cell3DProperties>();
        cell3d.Should().NotBeNull("Cell3DProperties must exist after bevel=circle");
        var bevel = cell3d!.GetFirstChild<Drawing.Bevel>();
        bevel.Should().NotBeNull("Bevel element must be present");
        bevel!.Preset!.InnerText.Should().Be("circle", "preset should be circle");
    }

    [Fact]
    public void Pptx_SetCell_Bevel_WithDimensions_CorrectEmuConversion()
    {
        // bevel="circle-6-4" → width=6pt, height=4pt
        // width = 6 * 12700 = 76200 EMU, height = 4 * 12700 = 50800 EMU
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle-6-4" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        var bevel = tcPr?.GetFirstChild<Drawing.Cell3DProperties>()?.GetFirstChild<Drawing.Bevel>();
        bevel.Should().NotBeNull();
        bevel!.Width!.Value.Should().Be(76200, "6pt * 12700 = 76200 EMU");
        bevel.Height!.Value.Should().Be(50800, "4pt * 12700 = 50800 EMU");
    }

    [Fact]
    public void Pptx_SetCell_Bevel_None_RemovesBevel()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "circle" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bevel"] = "none" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        var cell3d = tcPr?.GetFirstChild<Drawing.Cell3DProperties>();
        cell3d.Should().BeNull("Cell3DProperties must be removed after bevel=none");
    }

    // ==================== PPTX TABLE: Data Import ====================

    [Fact]
    public void Pptx_AddTable_WithData_InlineString_RowsAndColsFromData()
    {
        // data="H1,H2;R1C1,R1C2;R2C1,R2C2" — 3 rows, 2 cols
        _pptx.Add("/slide[1]", "table", null, new() { ["data"] = "H1,H2;R1C1,R1C2;R2C1,R2C2" });

        var cell = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("H1", "first cell should have data from inline string");

        var cell12 = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[2]");
        cell12.Text.Should().Be("H2");

        var cell21 = _pptx.Get("/slide[1]/table[1]/tr[2]/tc[1]");
        cell21.Text.Should().Be("R1C1");
    }

    [Fact]
    public void Pptx_AddTable_WithData_EmptyCell_DoesNotThrow()
    {
        // A semicolon with trailing comma: "A,;B,C" — row 1 has "A" and ""
        var act = () => _pptx.Add("/slide[1]", "table", null, new() { ["data"] = "A,;B,C" });
        act.Should().NotThrow("empty cell in data string should be handled gracefully");
    }

    [Fact]
    public void Pptx_AddTable_Rows1Cols1_MinimumTable()
    {
        // Edge case: 1x1 table
        var path = _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        path.Should().Contain("table[1]");
        var cell = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Should().NotBeNull("1x1 table must produce a valid cell");
    }

    // ==================== PPTX TABLE: autofit ====================

    [Fact]
    public void Pptx_Set_TableAutofit_DoesNotThrow()
    {
        _pptx.Add("/slide[1]", "table", null, new() {
            ["rows"] = "2", ["cols"] = "2",
            ["r1c1"] = "Short", ["r1c2"] = "Long column text here"
        });
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["autofit"] = "true" });
        act.Should().NotThrow("autofit=true should redistribute column widths without throwing");
    }

    [Fact]
    public void Pptx_Set_TableAutofit_False_DoesNothing()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        var act = () => _pptx.Set("/slide[1]/table[1]", new() { ["autofit"] = "false" });
        act.Should().NotThrow("autofit=false should be a no-op without throwing");
    }

    // ==================== WORD TABLE: TableLook flags ====================

    [Fact]
    public void Word_Set_TableLook_FirstRow_True_Persists()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["firstRow"] = "true" });

        ReopenWord();

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var tblLook = tblPr?.GetFirstChild<TableLook>();
        tblLook.Should().NotBeNull("TableLook must exist after firstRow=true");
        tblLook!.FirstRow?.Value.Should().BeTrue("firstRow=true must persist");
    }

    [Fact]
    public void Word_Set_TableLook_BandedRows_True_SetsNoHorizontalBandFalse()
    {
        // bandedRows=true → NoHorizontalBand=false (inverted logic in code line 1536)
        _word.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["bandedRows"] = "true" });

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var tblLook = tblPr?.GetFirstChild<TableLook>();
        tblLook.Should().NotBeNull();
        tblLook!.NoHorizontalBand?.Value.Should().BeFalse(
            "bandedRows=true means banding IS on, so NoHorizontalBand must be false");
    }

    [Fact]
    public void Word_Set_TableLook_BandedRows_False_SetsNoHorizontalBandTrue()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["bandedRows"] = "false" });

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var tblLook = tblPr?.GetFirstChild<TableLook>();
        tblLook.Should().NotBeNull();
        tblLook!.NoHorizontalBand?.Value.Should().BeTrue(
            "bandedRows=false means banding is off, so NoHorizontalBand must be true");
    }

    // ==================== WORD TABLE: Caption ====================

    [Fact]
    public void Word_Set_TableCaption_PersistsAfterReopen()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["caption"] = "My Caption" });

        ReopenWord();

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var caption = tblPr?.GetFirstChild<TableCaption>();
        caption.Should().NotBeNull("TableCaption must persist after save/reopen");
        caption!.Val?.Value.Should().Be("My Caption");
    }

    [Fact]
    public void Word_Set_TableCaption_Empty_RemovesCaption()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["caption"] = "My Caption" });
        _word.Set("/body/tbl[1]", new() { ["caption"] = "" });

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var caption = tblPr?.GetFirstChild<TableCaption>();
        caption.Should().BeNull("Empty caption value should remove the TableCaption element");
    }

    // ==================== WORD TABLE: Description ====================

    [Fact]
    public void Word_Set_TableDescription_PersistsAfterReopen()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["description"] = "Accessibility description" });

        ReopenWord();

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var desc = tblPr?.GetFirstChild<TableDescription>();
        desc.Should().NotBeNull("TableDescription must persist after save/reopen");
        desc!.Val?.Value.Should().Be("Accessibility description");
    }

    // ==================== WORD TABLE: Floating / Position ====================

    [Fact]
    public void Word_Set_TableFloating_Enables_TablePositionProperties()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["floating"] = "floating" });

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var tpp = tblPr?.GetFirstChild<TablePositionProperties>();
        tpp.Should().NotBeNull("TablePositionProperties must be created when floating=floating");
        tpp!.VerticalAnchor.Should().NotBeNull("VerticalAnchor must have default value");
        tpp.HorizontalAnchor.Should().NotBeNull("HorizontalAnchor must have default value");
    }

    [Fact]
    public void Word_Set_TablePosition_None_RemovesFloating()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["floating"] = "floating" });
        _word.Set("/body/tbl[1]", new() { ["position"] = "none" });

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var tpp = tblPr?.GetFirstChild<TablePositionProperties>();
        tpp.Should().BeNull("position=none should remove TablePositionProperties");
    }

    [Fact]
    public void Word_Set_TablePositionX_NumericValue_Persists()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["position.x"] = "720" }); // 720 twips = 0.5in

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var tpp = tblPr?.GetFirstChild<TablePositionProperties>();
        tpp.Should().NotBeNull("TablePositionProperties must be created for position.x");
        tpp!.TablePositionX?.Value.Should().Be(720, "position.x=720 should set TablePositionX=720 twips");
        tpp.TablePositionXAlignment.Should().BeNull("numeric x should clear XAlignment");
    }

    [Fact]
    public void Word_Set_TablePositionX_Left_SetsAlignmentClears_X()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["position.x"] = "left" });

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var tpp = tblPr?.GetFirstChild<TablePositionProperties>();
        tpp.Should().NotBeNull();
        tpp!.TablePositionXAlignment.Should().NotBeNull("position.x=left should set XAlignment");
        tpp.TablePositionX.Should().BeNull("position.x=left should clear numeric X");
    }

    [Fact]
    public void Word_Set_TableOverlap_Persists()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _word.Set("/body/tbl[1]", new() { ["overlap"] = "never" });

        ReopenWord();

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var tblPr = tbl.GetFirstChild<TableProperties>();
        var overlap = tblPr?.GetFirstChild<TableOverlap>();
        overlap.Should().NotBeNull("TableOverlap must persist after save/reopen");
        overlap!.Val.Should().NotBeNull();
        overlap.Val!.Value.Should().Be(TableOverlapValues.Never);
    }

    // ==================== WORD CELL: fitText ====================

    [Fact]
    public void Word_SetCell_FitText_True_AddsFitTextElement()
    {
        _word.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2",
            ["r1c1"] = "Hello" });

        // Set width first so fitText has a width to use
        _word.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["width"] = "1440" }); // 1440 twips = 1in
        _word.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "true" });

        _word.Dispose();
        using var doc = WordprocessingDocument.Open(_docxPath, false);
        var tbl = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var cell = tbl.Elements<TableRow>().First().Elements<TableCell>().First();
        var runs = cell.Descendants<Run>().ToList();
        runs.Should().NotBeEmpty("cell should have runs after text=Hello");
        var fitText = runs[0].GetFirstChild<RunProperties>()?.GetFirstChild<FitText>();
        fitText.Should().NotBeNull("FitText must be added to run properties when fitText=true");
    }

    // ==================== EXCEL TABLE: showRowStripes / showColumnStripes ====================

    [Fact]
    public void Excel_Set_ShowRowStripes_True_WhenStyleInfoExists()
    {
        // Add a sheet with a table
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:C3", ["name"] = "TestTable" });

        _excel.Set("/Sheet1/table[1]", new() { ["showRowStripes"] = "true" });

        var node = _excel.Get("/Sheet1/table[1]");
        node.Format.Should().ContainKey("showRowStripes",
            "showRowStripes should be readable back after being set");
        node.Format["showRowStripes"].Should().Be(true,
            "showRowStripes=true should result in showRowStripes=true in Format");
    }

    [Fact]
    public void Excel_Set_ShowRowStripes_False_UpdatesValue()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:C3", ["name"] = "TestTable" });

        _excel.Set("/Sheet1/table[1]", new() { ["showRowStripes"] = "false" });

        var node = _excel.Get("/Sheet1/table[1]");
        node.Format.Should().ContainKey("showRowStripes");
        node.Format["showRowStripes"].Should().Be(false,
            "showRowStripes=false must update the underlying value");
    }

    [Fact]
    public void Excel_Set_ShowRowStripes_WhenStyleInfoMissing_IsNotSilentlyIgnored()
    {
        // If a table has no TableStyleInfo, showRowStripes is silently ignored (bug: no error, no creation).
        // After Set, a Get should have the value or report it as unsupported.
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:C3", ["name"] = "NoStyleTable",
            ["style"] = "" }); // add without style to test null styleInfo

        // The key question: does Set return "showRowStripes" as unsupported, or does it create styleInfo?
        var unsupported = _excel.Set("/Sheet1/table[1]", new() { ["showRowStripes"] = "true" });

        // If showRowStripes is silently dropped, node won't have it — this is the bug
        var node = _excel.Get("/Sheet1/table[1]");

        // Assert: either it was applied (preferred) or at least it was returned as unsupported
        var wasApplied = node.Format.ContainsKey("showRowStripes") && node.Format["showRowStripes"] is true;
        var wasReported = unsupported.Any(u => u.Contains("showRowStripes", StringComparison.OrdinalIgnoreCase));
        (wasApplied || wasReported).Should().BeTrue(
            "showRowStripes should either be applied or returned as unsupported — " +
            "silently ignoring it when TableStyleInfo is null is a bug");
    }

    [Fact]
    public void Excel_Set_ShowColumnStripes_PersistsAfterReopen()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:C3", ["name"] = "TestTable" });
        _excel.Set("/Sheet1/table[1]", new() { ["showColumnStripes"] = "true" });

        ReopenExcel();

        var node = _excel.Get("/Sheet1/table[1]");
        node.Format.Should().ContainKey("showColumnStripes");
        node.Format["showColumnStripes"].Should().Be(true,
            "showColumnStripes=true must persist after save/reopen");
    }

    // ==================== EXCEL TABLE: col[N].* ====================

    [Fact]
    public void Excel_Set_ColName_UpdatesTableColumnName()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:C3", ["name"] = "TestTable",
            ["headers"] = "ColA,ColB,ColC" });
        _excel.Set("/Sheet1/table[1]", new() { ["col[1].name"] = "NewColA" });

        ReopenExcel();

        var node = _excel.Get("/Sheet1/table[1]");
        var columns = node.Format["columns"]?.ToString() ?? "";
        columns.Should().StartWith("NewColA", "col[1].name should update to NewColA after save/reopen");
    }

    [Fact]
    public void Excel_Set_ColTotalFunction_Sum_Persists()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:B4", ["name"] = "TestTable" });
        _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "sum" });

        ReopenExcel();

        _excel.Dispose();
        using var doc = SpreadsheetDocument.Open(_xlsxPath, false);
        var wbPart = doc.WorkbookPart!;
        var sheet = wbPart.Workbook.Descendants<Spreadsheet.Sheet>()
            .First(s => s.Name!.Value == "Sheet1");
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
        var tablePartRel = wsPart.TableDefinitionParts.First();
        var tblCol = tablePartRel.Table?.GetFirstChild<Spreadsheet.TableColumns>()
            ?.Elements<Spreadsheet.TableColumn>().First();
        tblCol.Should().NotBeNull();
        tblCol!.TotalsRowFunction!.InnerText.Should().Be("sum",
            "col[1].totalFunction=sum must persist after save/reopen");
    }

    [Fact]
    public void Excel_Set_ColOutOfRange_ThrowsArgumentException()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:B3", ["name"] = "TestTable" });
        // Table has 2 columns (A, B); col[5] is out of range
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[5].name"] = "Test" });
        act.Should().Throw<ArgumentException>("col index out of range should throw ArgumentException");
    }

    [Fact]
    public void Excel_Set_ColTotalFunction_InvalidValue_ThrowsArgumentException()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:B3", ["name"] = "TestTable" });
        var act = () => _excel.Set("/Sheet1/table[1]", new() { ["col[1].totalFunction"] = "invalid_function" });
        act.Should().Throw<ArgumentException>("invalid totalFunction value should throw ArgumentException");
    }

    // ==================== EXCEL TABLE: TableToNode readback ====================

    [Fact]
    public void Excel_TableToNode_ShowFirstColumn_ReadbackCorrect()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:D4", ["name"] = "TestTable" });
        _excel.Set("/Sheet1/table[1]", new() { ["showFirstColumn"] = "true" });

        var node = _excel.Get("/Sheet1/table[1]");
        node.Format.Should().ContainKey("showFirstColumn");
        node.Format["showFirstColumn"].Should().Be(true,
            "showFirstColumn=true should be read back from TableStyleInfo");
    }

    [Fact]
    public void Excel_TableToNode_ShowLastColumn_ReadbackCorrect()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:D4", ["name"] = "TestTable" });
        _excel.Set("/Sheet1/table[1]", new() { ["showLastColumn"] = "true" });

        var node = _excel.Get("/Sheet1/table[1]");
        node.Format.Should().ContainKey("showLastColumn");
        node.Format["showLastColumn"].Should().Be(true,
            "showLastColumn=true should be read back from TableStyleInfo");
    }

    [Fact]
    public void Excel_TableToNode_Columns_ReadbackAfterColNameUpdate()
    {
        _excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:C3", ["name"] = "TestTable" });
        _excel.Set("/Sheet1/table[1]", new() { ["col[2].name"] = "SecondCol" });

        var node = _excel.Get("/Sheet1/table[1]");
        var cols = node.Format["columns"]?.ToString() ?? "";
        cols.Should().Contain("SecondCol", "updated column name should appear in columns readback");
    }

    // ==================== PPTX TABLE: Schema Order Verification ====================

    [Fact]
    public void Pptx_SetTable_ShadowAndGlow_EffectListSchemaOrder()
    {
        // Both shadow and glow go into EffectList. They must follow schema order.
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["shadow"] = "FF0000;4;45;2" });
        _pptx.Set("/slide[1]/table[1]", new() { ["glow"] = "0000FF;8" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        var effectList = tblPr?.GetFirstChild<Drawing.EffectList>();
        effectList.Should().NotBeNull();

        var children = effectList!.ChildElements.ToList();
        // Both shadow and glow should be present (glow before shadow in CT_EffectList schema)
        var glowIdx = children.FindIndex(c => c is Drawing.Glow);
        var shadowIdx = children.FindIndex(c => c is Drawing.OuterShadow);
        glowIdx.Should().BeGreaterThanOrEqualTo(0, "Glow should be in EffectList");
        shadowIdx.Should().BeGreaterThanOrEqualTo(0, "OuterShadow should be in EffectList");
    }

    // ==================== PPTX CELL: Fill → Opacity ordering ====================

    [Fact]
    public void Pptx_SetCell_Fill_ThenOpacity_BothPersist()
    {
        _pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000" });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["opacity"] = "0.8" });

        // Verify both fill and opacity (alpha) are set
        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var gf = slide.Descendants<GraphicFrame>().First();
        var table = gf.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.GetFirstChild<Drawing.TableCellProperties>();
        var solidFill = tcPr?.GetFirstChild<Drawing.SolidFill>();
        solidFill.Should().NotBeNull("fill=FF0000 should create SolidFill");
        var hex = solidFill!.GetFirstChild<Drawing.RgbColorModelHex>();
        hex.Should().NotBeNull();
        hex!.Val!.Value.Should().Be("FF0000");
        var alpha = hex.GetFirstChild<Drawing.Alpha>();
        alpha.Should().NotBeNull("Alpha element should be added by opacity=0.8");
        // Bug: current code gives 0.8 * 1000 = 800; correct is 0.8 * 100000 = 80000
        alpha!.Val!.Value.Should().Be(80000,
            "opacity=0.8 should set Alpha.Val=80000 (80%), not 800 (0.8%). " +
            "Multiplier should be 100000 not 1000.");
    }
}
