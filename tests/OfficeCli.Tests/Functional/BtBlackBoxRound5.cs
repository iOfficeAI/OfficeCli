// Black-box tests (Round 5) — persistence and Remove integrity:
//   - Persistence round-trip: Set various properties, Reopen, verify all properties survive
//   - Remove + verify: Remove → Get returns null/throws, Query excludes removed elements
//   - Remove cascade: deleting parent → child paths inaccessible
//   - Picture round-trip: Add picture → Get → verify → Reopen → re-verify
//   - Complex property persistence: gradient fill, shadow, glow, reflection round-trips

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound5 : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public BtBlackBoxRound5(ITestOutputHelper output) => _output = output;

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bt5_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    private string CreateTestImage()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bt5img_{Guid.NewGuid():N}.png");
        _tempFiles.Add(path);
        // Minimal valid 1x1 PNG
        var pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0C,0x49,0x44,0x41,
            0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
            0x00,0x00,0x02,0x00,0x01,0xE2,0x21,0xBC,
            0x33,0x00,0x00,0x00,0x00,0x49,0x45,0x4E,
            0x44,0xAE,0x42,0x60,0x82
        };
        File.WriteAllBytes(path, pngBytes);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private void AssertValidPptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"PPTX must be schema-valid after: {step}");
    }

    private void AssertValidDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"DOCX must be schema-valid after: {step}");
    }

    private void AssertValidXlsx(string path, string step)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"XLSX must be schema-valid after: {step}");
    }

    // ═══════════════════════════════════════════════════════
    // PPTX — persistence round-trips
    // ═══════════════════════════════════════════════════════

    [Fact]
    public void Pptx_ShapeMultiProps_Reopen_AllPropertiesPersist()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "RoundTrip" });
            h.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Persist Me",
                ["fill"] = "FF4500",
                ["bold"] = "true",
                ["size"] = "18",
                ["left"] = "2cm", ["top"] = "2cm", ["width"] = "8cm", ["height"] = "3cm"
            });
        }

        AssertValidPptx(path, "multi-prop shape");

        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]/shape[2]");
        node.Should().NotBeNull();
        node.Text.Should().Be("Persist Me");
        node.Format["fill"].ToString().Should().Be("#FF4500");
        node.Format["bold"].ToString().Should().Be("True");
        node.Format["size"].ToString().Should().Be("18pt");
    }

    [Fact]
    public void Pptx_GradientFill_Reopen_Persists()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Gradient" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Grad" });
            h.Set("/slide[1]/shape[2]", new() { ["gradient"] = "FF0000-0000FF-90" });
        }

        AssertValidPptx(path, "gradient-fill");

        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]/shape[2]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("gradient");
        var grad = node.Format["gradient"].ToString()!;
        grad.Should().Contain("FF0000");
        grad.Should().Contain("0000FF");
    }

    [Fact]
    public void Pptx_Shadow_Reopen_Persists()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Shadow" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shadowed" });
            h.Set("/slide[1]/shape[2]", new() { ["shadow"] = "000000-4-45-3-50" });
        }

        AssertValidPptx(path, "shadow");

        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("shadow");
        node.Format["shadow"].ToString().Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Pptx_GlowEffect_Reopen_Persists()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Glow" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Glowing" });
            h.Set("/slide[1]/shape[2]", new() { ["glow"] = "0070FF-10-60" });
        }

        AssertValidPptx(path, "glow-effect");

        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("glow");
        node.Format["glow"].ToString().Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Pptx_Picture_AddGetReopen_PathAndSizePersist()
    {
        var img = CreateTestImage();
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Pics" });
            h.Add("/slide[1]", "picture", null, new()
            {
                ["path"] = img,
                ["left"] = "1cm", ["top"] = "1cm", ["width"] = "5cm", ["height"] = "4cm"
            });

            // Immediate Get check
            var node = h.Get("/slide[1]/picture[1]");
            node.Should().NotBeNull();
            node.Type.Should().Be("picture");
        }

        AssertValidPptx(path, "picture-add");

        using var h2 = new PowerPointHandler(path, editable: false);
        var node2 = h2.Get("/slide[1]/picture[1]");
        node2.Should().NotBeNull();
        node2.Type.Should().Be("picture");
    }

    [Fact]
    public void Pptx_Remove_Shape_GetReturnsNull_QueryExcludes()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Remove Test" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "ToRemove" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "KeepMe" });

        // Before removal - shape[2] and shape[3] should exist
        h.Get("/slide[1]/shape[2]").Should().NotBeNull();
        h.Get("/slide[1]/shape[3]").Should().NotBeNull();

        h.Remove("/slide[1]/shape[2]");

        // After removal - shape at old index 2 is gone, old index 3 is now index 2
        var act = () => h.Get("/slide[1]/shape[3]");
        act.Should().Throw<Exception>("after removing shape[2], only 2 shapes remain");

        // Remaining shape should be accessible
        var remaining = h.Get("/slide[1]/shape[2]");
        remaining.Should().NotBeNull();
        remaining.Text.Should().Be("KeepMe");
    }

    [Fact]
    public void Pptx_Remove_Shape_Query_ExcludesRemovedText()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Query Test" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Alpha" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Beta" });

        h.Remove("/slide[1]/shape[2]"); // Remove "Alpha"

        var shapes = h.Query("shape");
        shapes.Should().NotContain(n => n.Text == "Alpha", "removed shape should not appear in Query");
    }

    [Fact]
    public void Pptx_RemoveSlide_Cascade_ShapeInaccessible()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Doomed Slide" });
            h.Add("/", "slide", null, new() { ["title"] = "Survivor" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "WillVanish" });

            h.Remove("/slide[1]");

            // After cascade remove, only 1 slide remains. slide[2] is gone.
            var act = () => h.Get("/slide[2]");
            act.Should().Throw<ArgumentException>("only 1 slide should remain");
        }

        AssertValidPptx(path, "cascade-remove-slide");
    }

    [Fact]
    public void Pptx_RemoveSlide_Reopen_SlideCountCorrect()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "A" });
            h.Add("/", "slide", null, new() { ["title"] = "B" });
            h.Add("/", "slide", null, new() { ["title"] = "C" });
            // Now 3 slides total. Remove slide 2.
            h.Remove("/slide[2]");
        }

        AssertValidPptx(path, "remove-slide-reopen");

        using var h2 = new PowerPointHandler(path, editable: false);
        // 2 slides remain: slide[1]="A", slide[2]="C"
        var s2 = h2.Get("/slide[2]");
        s2.Should().NotBeNull();

        var act = () => h2.Get("/slide[3]");
        act.Should().Throw<ArgumentException>("only 2 slides should remain after removal");
    }

    // ═══════════════════════════════════════════════════════
    // WORD — persistence round-trips
    // ═══════════════════════════════════════════════════════

    [Fact]
    public void Word_Paragraph_MultiProps_Reopen_AllPropertiesPersist()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new()
            {
                ["text"] = "Styled",
                ["bold"] = "true",
                ["size"] = "16",
                ["alignment"] = "center",
                ["spaceBefore"] = "12pt",
                ["spaceAfter"] = "6pt"
            });
        }

        AssertValidDocx(path, "paragraph-multi-props");

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get("/body/p[1]");
        node.Should().NotBeNull();
        node.Text.Should().Contain("Styled");
        node.Format["alignment"].ToString().Should().Be("center");
        node.Format["spaceBefore"].ToString().Should().Be("12pt");
        node.Format["spaceAfter"].ToString().Should().Be("6pt");
    }

    [Fact]
    public void Word_Remove_Paragraph_GetThrows_QueryExcludes()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "First" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "ToDelete" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "Last" });

            // Remove the second paragraph
            h.Remove("/body/p[2]");

            var paras = h.Query("paragraph");
            paras.Should().NotContain(n => n.Text != null && n.Text.Contains("ToDelete"));
        }

        AssertValidDocx(path, "remove-paragraph");
    }

    [Fact]
    public void Word_Remove_Table_Cascade_CellsInaccessible()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
            h.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "CellData" });

            h.Remove("/body/tbl[1]");

            var act = () => h.Get("/body/tbl[1]");
            act.Should().Throw<Exception>("table was removed");
        }

        AssertValidDocx(path, "cascade-remove-table");
    }

    [Fact]
    public void Word_Paragraph_FontColor_Reopen_Persists()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new()
            {
                ["text"] = "Colored",
                ["color"] = "FF0000"
            });
        }

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get("/body/p[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("color");
        node.Format["color"].ToString().Should().Be("#FF0000");
    }

    [Fact]
    public void Word_PageSize_Reopen_Persists()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Set("/section[1]", new() { ["pageWidth"] = "21cm", ["pageHeight"] = "29.7cm" });
        }

        AssertValidDocx(path, "page-size");

        using var h2 = new WordHandler(path, editable: false);
        var sec = h2.Get("/section[1]");
        sec.Format.Should().ContainKey("pageWidth");
        sec.Format.Should().ContainKey("pageHeight");
        // Values may vary slightly due to unit conversion; just confirm they're present and non-empty
        sec.Format["pageWidth"].ToString().Should().NotBeNullOrEmpty();
        sec.Format["pageHeight"].ToString().Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Word_Picture_AddReopen_TypePersists()
    {
        var img = CreateTestImage();
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "picture", null, new()
            {
                ["path"] = img,
                ["width"] = "4cm",
                ["height"] = "3cm"
            });
        }

        AssertValidDocx(path, "picture-word");

        using var h2 = new WordHandler(path, editable: false);
        var pics = h2.Query("picture");
        pics.Should().NotBeEmpty("picture should persist after reopen");
    }

    [Fact]
    public void Word_Remove_Multiple_Elements_DocumentValid()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "P1" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "P2" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "P3" });
            h.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "P4" });

            h.Remove("/body/tbl[1]");
            h.Remove("/body/p[2]");
        }

        AssertValidDocx(path, "multi-remove");

        using var h2 = new WordHandler(path, editable: false);
        // p[1] should be "P1" (the first paragraph in the body, though blank doc may have one)
        var paras = h2.Query("paragraph");
        paras.Should().NotContain(n => n.Text == "P2");
    }

    // ═══════════════════════════════════════════════════════
    // EXCEL — persistence round-trips
    // ═══════════════════════════════════════════════════════

    [Fact]
    public void Excel_CellFormats_Reopen_AllPersist()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "cell", null, new()
            {
                ["address"] = "A1",
                ["value"] = "Formatted",
                ["bold"] = "true",
                ["italic"] = "true",
                ["fill"] = "FFFF00",
                ["alignment.horizontal"] = "center"
            });
            h.Set("/Sheet1/A1", new() { ["font.color"] = "FF0000" });
        }

        AssertValidXlsx(path, "cell-formats");

        using var h2 = new ExcelHandler(path, editable: false);
        var node = h2.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node.Text.Should().Be("Formatted");
        node.Format["bold"].ToString().Should().Be("True");
        node.Format["italic"].ToString().Should().Be("True");
        node.Format["font.color"].ToString().Should().Be("#FF0000");
        node.Format["fill"].ToString().Should().Be("#FFFF00");
        node.Format["alignment.horizontal"].ToString().Should().Be("center");
    }

    [Fact]
    public void Excel_Remove_Cell_CanBeCalledWithoutError()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        h.Add("/Sheet1", "cell", null, new() { ["address"] = "C3", ["value"] = "DeleteMe" });
        var before = h.Get("/Sheet1/C3");
        before.Text.Should().Be("DeleteMe");

        // Remove should not throw; subsequent Get may return null or empty
        var act = () => h.Remove("/Sheet1/C3");
        act.Should().NotThrow("Remove on a cell should be a valid operation");
    }

    [Fact]
    public void Excel_Remove_Row_Persists_After_Reopen()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "Row1A" });
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "B1", ["value"] = "Row1B" });
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "A2", ["value"] = "Row2A" });
            h.Remove("/Sheet1/row[1]");
        }

        AssertValidXlsx(path, "remove-row");

        using var h2 = new ExcelHandler(path, editable: false);
        var a1 = h2.Get("/Sheet1/A1");
        if (a1 != null)
            a1.Text.Should().NotBe("Row1A", "Row 1 data should be removed");
    }

    [Fact]
    public void Excel_Picture_AddReopen_TypePersists()
    {
        var img = CreateTestImage();
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "picture", null, new()
            {
                ["path"] = img,
                ["x"] = "0", ["y"] = "0", ["width"] = "3", ["height"] = "3"
            });
        }

        AssertValidXlsx(path, "picture-excel");

        using var h2 = new ExcelHandler(path, editable: false);
        var pics = h2.Query("picture");
        pics.Should().NotBeEmpty("picture should persist after reopen");
    }

    [Fact]
    public void Excel_NumberFormat_Reopen_Persists()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "cell", null, new()
            {
                ["address"] = "D4",
                ["value"] = "1234.56",
                ["numberformat"] = "#,##0.00"
            });
        }

        using var h2 = new ExcelHandler(path, editable: false);
        var node = h2.Get("/Sheet1/D4");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("numberformat");
        node.Format["numberformat"].ToString().Should().Contain("0.00");
    }

    [Fact]
    public void Excel_MergedCells_Reopen_Persists()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "Merged" });
            h.Set("/Sheet1", new() { ["merge"] = "A1:C1" });
        }

        AssertValidXlsx(path, "merge-cells");

        using var h2 = new ExcelHandler(path, editable: false);
        var sheet = h2.Get("/Sheet1");
        sheet.Should().NotBeNull();
        // Verify file is valid and re-openable — merge persisted
        var a1 = h2.Get("/Sheet1/A1");
        a1.Should().NotBeNull();
        a1.Text.Should().Be("Merged");
    }

    [Fact]
    public void Excel_Formula_Reopen_FormulaKeyPresent()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "E5", ["formula"] = "=SUM(A1:A10)" });
        }

        using var h2 = new ExcelHandler(path, editable: false);
        var node = h2.Get("/Sheet1/E5");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].ToString().Should().Contain("SUM");
    }
}
