// Hardcoded Bug Tests Round 1 — bugs found by scanning for hardcoded values
//
// Bug 1 (CRITICAL): ParsePresetShape case mismatch — uturnArrow/circularArrow unreachable
// Bug 2 (HIGH): Equation shapeId collision — only counted Shape+Picture, missed GraphicFrame etc.
// Bug 3 (HIGH): DataBar default min/max used Number(0,1) instead of Min/Max auto-range

using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

public class HardcodedBugTests_Round1 : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _pptHandler;

    public HardcodedBugTests_Round1()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _pptHandler = new PowerPointHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _pptHandler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    // ==================== Bug 1: ParsePresetShape case mismatch ====================

    [Fact]
    public void Bug1_UTurnArrow_ShouldBeCreatedSuccessfully()
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "Arrow Test" });

        var path = _pptHandler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "uturnArrow",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "5cm", ["height"] = "4cm",
            ["fill"] = "FF0000"
        });

        path.Should().NotBeNullOrEmpty();
        var node = _pptHandler.Get(path);
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("preset");
    }

    [Fact]
    public void Bug1_CircularArrow_ShouldBeCreatedSuccessfully()
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "Arrow Test" });

        var path = _pptHandler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "circularArrow",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "5cm", ["height"] = "4cm",
            ["fill"] = "00B050"
        });

        path.Should().NotBeNullOrEmpty();
        var node = _pptHandler.Get(path);
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("preset");
    }

    [Theory]
    [InlineData("uturnarrow")]
    [InlineData("UTURNARROW")]
    [InlineData("uturnArrow")]
    [InlineData("UTurnArrow")]
    public void Bug1_UTurnArrow_CaseInsensitive(string presetName)
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "Case Test" });

        var path = _pptHandler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = presetName,
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "3cm"
        });

        path.Should().NotBeNullOrEmpty();
    }

    // ==================== Bug 2: Equation shapeId collision ====================

    [Fact]
    public void Bug2_EquationId_ShouldNotCollideWithExistingElements()
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "ID Test" });

        // Add shape, table, chart to create diverse element types
        _pptHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape1", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        _pptHandler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "5cm", ["y"] = "1cm", ["width"] = "8cm", ["height"] = "4cm"
        });
        _pptHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Test",
            ["categories"] = "A,B", ["data"] = "S:1,2"
        });

        // Now add equation — should not collide with existing IDs
        var eqPath = _pptHandler.Add("/slide[1]", "equation", null, new()
        {
            ["formula"] = "x^2 + y^2 = z^2"
        });

        eqPath.Should().NotBeNullOrEmpty();

        // Verify all elements on the slide have unique IDs
        var doc = (DocumentFormat.OpenXml.Packaging.PresentationDocument)_pptHandler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_pptHandler)!;
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        shapeTree.Should().NotBeNull();

        var ids = shapeTree!.ChildElements
            .Select(e => e.Descendants<DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties>()
                .FirstOrDefault()?.Id?.Value)
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .ToList();

        ids.Should().OnlyHaveUniqueItems("all shape IDs on the slide should be unique");
    }

    // ==================== Bug 3: DataBar default min/max ====================

    [Fact]
    public void Bug3_DataBar_DefaultMinMax_ShouldUseAutoRange()
    {
        var xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        try
        {
            BlankDocCreator.Create(xlsxPath);
            using var handler = new ExcelHandler(xlsxPath, editable: true);

            // Add some data with values > 1
            for (int i = 1; i <= 5; i++)
                handler.Set($"/Sheet1/A{i}", new() { ["value"] = (i * 20).ToString() });

            // Add databar without specifying min/max
            handler.Add("/Sheet1", "databar", null, new() { ["sqref"] = "A1:A5" });

            // Verify via XML that the CFVO types are Min/Max, not Number
            var doc = (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument)handler.GetType()
                .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
                .GetValue(handler)!;
            var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
            var cf = sheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>().First();
            var dataBar = cf.Descendants<DocumentFormat.OpenXml.Spreadsheet.DataBar>().First();
            var cfvos = dataBar.Elements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObject>().ToList();

            cfvos.Should().HaveCount(2);
            cfvos[0].Type!.InnerText.Should().NotBe("num",
                "default min should use auto-range Min type, not fixed Number");
            cfvos[1].Type!.InnerText.Should().NotBe("num",
                "default max should use auto-range Max type, not fixed Number");
        }
        finally
        {
            if (File.Exists(xlsxPath)) File.Delete(xlsxPath);
        }
    }

    [Fact]
    public void Bug3_DataBar_ExplicitMinMax_ShouldUseNumberType()
    {
        var xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        try
        {
            BlankDocCreator.Create(xlsxPath);
            using var handler = new ExcelHandler(xlsxPath, editable: true);

            // Add databar with explicit min/max
            handler.Add("/Sheet1", "databar", null, new()
            {
                ["sqref"] = "A1:A5", ["min"] = "0", ["max"] = "100"
            });

            var doc = (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument)handler.GetType()
                .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
                .GetValue(handler)!;
            var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
            var cf = sheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>().First();
            var dataBar = cf.Descendants<DocumentFormat.OpenXml.Spreadsheet.DataBar>().First();
            var cfvos = dataBar.Elements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObject>().ToList();

            cfvos[0].Type!.Value.Should().Be(
                DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Number,
                "explicit min should use Number type");
            cfvos[0].Val!.Value.Should().Be("0");
            cfvos[1].Type!.Value.Should().Be(
                DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Number,
                "explicit max should use Number type");
            cfvos[1].Val!.Value.Should().Be("100");
        }
        finally
        {
            if (File.Exists(xlsxPath)) File.Delete(xlsxPath);
        }
    }
}
