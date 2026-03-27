// Black-box tests (Round 4) — deep boundary cases and complex scenarios:
//   - Excel: multi-sheet (add/remove/rename), cross-sheet reference, freeze pane, autofilter
//   - Word: table colspan/vmerge, section columns
//   - PPTX: group shape, hyperlink on shape, slide reorder, slide delete
//   - Persistence: reopen verifies all properties survive
//   - Remove: sheet delete cleans references, slide/shape ordinal consistency

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound4 : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public BtBlackBoxRound4(ITestOutputHelper output) => _output = output;

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bt4_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) File.Delete(f);
    }

    private void AssertValidXlsx(string path, string step)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"XLSX must be schema-valid after: {step}");
    }

    private void AssertValidDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"DOCX must be schema-valid after: {step}");
    }

    private void AssertValidPptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"PPTX must be schema-valid after: {step}");
    }

    // ═══════════════════════════════════════════════════════
    // EXCEL — multi-sheet operations
    // ═══════════════════════════════════════════════════════

    [Fact]
    public void Excel_AddSecondSheet_GetReturnsSheetType()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        var newPath = h.Add("/", "sheet", null, new() { ["name"] = "Summary" });
        newPath.Should().NotBeNullOrEmpty();

        var node = h.Get("/Summary");
        node.Should().NotBeNull();
        node.Type.Should().Be("sheet");
    }

    [Fact]
    public void Excel_AddSheet_WriteCell_Reopen_Persists()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/", "sheet", null, new() { ["name"] = "Data" });
            h.Add("/Data", "cell", null, new() { ["address"] = "A1", ["value"] = "Persisted" });
        }

        using var h2 = new ExcelHandler(path, editable: false);
        var node = h2.Get("/Data/A1");
        node.Should().NotBeNull();
        node.Text.Should().Be("Persisted");
    }

    [Fact]
    public void Excel_RenameSheet_GetByNewName_Works()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "B2", ["value"] = "renamed" });
            h.Set("/Sheet1", new() { ["name"] = "Renamed" });
        }

        AssertValidXlsx(path, "rename-sheet");

        using var h2 = new ExcelHandler(path, editable: false);
        var node = h2.Get("/Renamed/B2");
        node.Should().NotBeNull();
        node.Text.Should().Be("renamed");
    }

    [Fact]
    public void Excel_RemoveSheet_CannotRemoveLast()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        var act = () => h.Remove("/Sheet1");
        act.Should().Throw<InvalidOperationException>()
            .WithMessage("*last sheet*");
    }

    [Fact]
    public void Excel_RemoveSheet_SecondSheet_CleansProperly()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/", "sheet", null, new() { ["name"] = "Temp" });
            h.Add("/Temp", "cell", null, new() { ["address"] = "A1", ["value"] = "bye" });
            h.Remove("/Temp");
        }

        AssertValidXlsx(path, "remove-sheet");

        using var h2 = new ExcelHandler(path, editable: false);
        var act = () => h2.Get("/Temp");
        act.Should().Throw<Exception>();
    }

    [Fact]
    public void Excel_CrossSheetFormula_ReturnsFormulaText()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        h.Add("/", "sheet", null, new() { ["name"] = "Source" });
        h.Add("/Source", "cell", null, new() { ["address"] = "A1", ["value"] = "42" });

        // Cross-sheet formula referencing Source!A1
        h.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["formula"] = "=Source!A1+1" });

        var node = h.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].ToString().Should().Contain("Source");
    }

    [Fact]
    public void Excel_FreezePane_SetAndGet_Persists()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Set("/Sheet1", new() { ["freeze"] = "B2" });
        }

        using var h2 = new ExcelHandler(path, editable: false);
        var node = h2.Get("/Sheet1");
        node.Format.Should().ContainKey("freeze");
        node.Format["freeze"].ToString().Should().Be("B2");
    }

    [Fact]
    public void Excel_FreezePane_Remove_NoLongerPresent()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        h.Set("/Sheet1", new() { ["freeze"] = "A2" });
        h.Set("/Sheet1", new() { ["freeze"] = "none" });

        var node = h.Get("/Sheet1");
        // Either the key is absent or value is empty after removal
        if (node.Format.ContainsKey("freeze"))
            node.Format["freeze"].ToString().Should().BeNullOrEmpty();
    }

    [Fact]
    public void Excel_AutoFilter_SetAndGet()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "Header" });
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "B1", ["value"] = "Value" });
            h.Set("/Sheet1", new() { ["autofilter"] = "A1:B5" });
        }

        AssertValidXlsx(path, "autofilter");

        using var h2 = new ExcelHandler(path, editable: false);
        var node = h2.Get("/Sheet1");
        node.Format.Should().ContainKey("autoFilter");
        node.Format["autoFilter"].ToString().Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Excel_MultiSheet_QueryReturnsAll()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        h.Add("/", "sheet", null, new() { ["name"] = "Alpha" });
        h.Add("/", "sheet", null, new() { ["name"] = "Beta" });

        var sheets = h.Query("sheet");
        sheets.Count.Should().BeGreaterThanOrEqualTo(3); // Sheet1 + Alpha + Beta
    }

    [Fact]
    public void Excel_RemoveSheet_NamedRangeRefCleaned()
    {
        var path = CreateTemp("xlsx");
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/", "sheet", null, new() { ["name"] = "Volatile" });
            h.Add("/Volatile", "cell", null, new() { ["address"] = "A1", ["value"] = "data" });

            // Named range pointing at the soon-to-be-deleted sheet
            h.Add("/", "namedrange", null, new() { ["name"] = "VolRef", ["ref"] = "Volatile!A1" });

            // Remove the sheet — named range referencing it should also be cleaned up
            h.Remove("/Volatile");
        }

        AssertValidXlsx(path, "remove-sheet-named-range-cleanup");
    }

    // ═══════════════════════════════════════════════════════
    // WORD — complex table, section columns
    // ═══════════════════════════════════════════════════════

    [Fact]
    public void Word_TableColspan_SetAndGet()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "3" });

            // Set colspan=3 on first cell (merge across all 3 columns)
            h.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["colspan"] = "3" });

            var cell = h.Get("/body/tbl[1]/tr[1]/tc[1]");
            cell.Should().NotBeNull();
            cell.Format.Should().ContainKey("gridSpan");
            cell.Format["gridSpan"].ToString().Should().Be("3");
        }

        AssertValidDocx(path, "table-colspan");
    }

    [Fact]
    public void Word_TableVmerge_SetAndGet()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });

            // Start merge at row 1
            h.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["vmerge"] = "restart" });
            // Continue merge at row 2
            h.Set("/body/tbl[1]/tr[2]/tc[1]", new() { ["vmerge"] = "continue" });

            var cell1 = h.Get("/body/tbl[1]/tr[1]/tc[1]");
            cell1.Format.Should().ContainKey("vmerge");
            cell1.Format["vmerge"].ToString().Should().Be("restart");
        }

        AssertValidDocx(path, "table-vmerge");
    }

    [Fact]
    public void Word_Table_AddTextToCell_GetReturnsText()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        h.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "CellA1" });

        var cell = h.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Contain("CellA1");
    }

    [Fact]
    public void Word_RemoveTable_DocumentIsValid()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Before" });
            h.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "After" });

            h.Remove("/body/tbl[1]");

            var act = () => h.Get("/body/tbl[1]");
            act.Should().Throw<Exception>();
        }

        AssertValidDocx(path, "remove-table");
    }

    [Fact]
    public void Word_SectionColumns_SetAndGet()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Set("/section[1]", new() { ["columns"] = "2" });

        var secNode = h.Get("/section[1]");
        secNode.Format.Should().ContainKey("columns");
        secNode.Format["columns"].ToString().Should().Be("2");
    }

    [Fact]
    public void Word_SectionColumns_Reopen_Persists()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Set("/section[1]", new() { ["columns"] = "3" });
        }

        AssertValidDocx(path, "section-columns");

        using var h2 = new WordHandler(path, editable: false);
        var sec = h2.Get("/section[1]");
        sec.Format["columns"].ToString().Should().Be("3");
    }

    [Fact]
    public void Word_TableColspan_Reopen_Persists()
    {
        var path = CreateTemp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "4" });
            h.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["colspan"] = "2" });
        }

        using var h2 = new WordHandler(path, editable: false);
        var cell = h2.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Format.Should().ContainKey("gridSpan");
        cell.Format["gridSpan"].ToString().Should().Be("2");
    }

    // ═══════════════════════════════════════════════════════
    // PPTX — group shape, hyperlink, slide management
    // ═══════════════════════════════════════════════════════

    [Fact]
    public void Pptx_GroupShape_CreateAndGet()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Group Test" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape1", ["left"] = "1cm", ["top"] = "1cm", ["width"] = "3cm", ["height"] = "2cm" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape2", ["left"] = "5cm", ["top"] = "1cm", ["width"] = "3cm", ["height"] = "2cm" });

            var grpPath = h.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });
            grpPath.Should().Contain("group");

            var grpNode = h.Get(grpPath);
            grpNode.Should().NotBeNull();
            grpNode.Type.Should().Be("group");
        }

        AssertValidPptx(path, "group-shape");
    }

    [Fact]
    public void Pptx_Hyperlink_SetOnShape_GetReturnsLink()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Link Slide" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Click me" });

            // shape[1] is title placeholder, our added shape is shape[2]
            h.Set("/slide[1]/shape[2]", new() { ["link"] = "https://example.com" });

            var node = h.Get("/slide[1]/shape[2]");
            node.Format.Should().ContainKey("link");
            // Uri normalization may append a trailing slash to bare-domain URLs
            node.Format["link"].ToString().Should().StartWith("https://example.com");
        }

        AssertValidPptx(path, "shape-hyperlink");
    }

    [Fact]
    public void Pptx_Hyperlink_Reopen_Persists()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Persist Link" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Link" });
            h.Set("/slide[1]/shape[2]", new() { ["link"] = "https://persist.test" });
        }

        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("link");
        // Uri normalization may append a trailing slash to bare-domain URLs
        node.Format["link"].ToString().Should().StartWith("https://persist.test");
    }

    [Fact]
    public void Pptx_SlideReorder_MoveLastToFirst()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Slide A" });
            h.Add("/", "slide", null, new() { ["title"] = "Slide B" });
            h.Add("/", "slide", null, new() { ["title"] = "Slide C" });

            // Move slide 3 to position 1 (index 0)
            var newPath = h.Move("/slide[3]", null, 0);
            newPath.Should().Be("/slide[1]");

            // Verify slide at position 1 exists
            var s1 = h.Get("/slide[1]", depth: 1);
            s1.Should().NotBeNull();
        }

        AssertValidPptx(path, "slide-reorder");
    }

    [Fact]
    public void Pptx_DeleteSlide_RemainingCount()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "First" });
            h.Add("/", "slide", null, new() { ["title"] = "Second" });
            h.Add("/", "slide", null, new() { ["title"] = "Third" });

            // Remove slide 2 (3 slides total → 2 remaining)
            h.Remove("/slide[2]");

            // Verify slide 3 is now accessible as slide 2
            var s2 = h.Get("/slide[2]", depth: 1);
            s2.Should().NotBeNull();

            // Verify slide 3 no longer exists
            var act = () => h.Get("/slide[3]");
            act.Should().Throw<ArgumentException>("slide[3] should not exist after deletion");
        }

        AssertValidPptx(path, "delete-slide");
    }

    [Fact]
    public void Pptx_DeleteSlide_CannotDeleteLast()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        // Blank doc has 1 slide; attempting to remove it should throw
        var act = () => h.Remove("/slide[1]");
        act.Should().Throw<Exception>();
    }

    [Fact]
    public void Pptx_MultiSlide_ContentPersistsAcrossReopen()
    {
        var path = CreateTemp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Slide One" });
            h.Add("/", "slide", null, new() { ["title"] = "Slide Two" });
            // Add shapes — shape[1] is the title placeholder, our shapes start at [2]
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
            h.Add("/slide[2]", "shape", null, new() { ["text"] = "World" });
        }

        using var h2 = new PowerPointHandler(path, editable: false);
        var s1shapes = h2.Query("slide[1] > shape");
        s1shapes.Should().Contain(n => n.Text == "Hello");

        var s2shapes = h2.Query("slide[2] > shape");
        s2shapes.Should().Contain(n => n.Text == "World");
    }

    [Fact]
    public void Pptx_HyperlinkRemove_NoLongerInFormat()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Remove Link" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Linked" });
        h.Set("/slide[1]/shape[2]", new() { ["link"] = "https://remove.me" });
        h.Set("/slide[1]/shape[2]", new() { ["link"] = "none" });

        var node = h.Get("/slide[1]/shape[2]");
        if (node.Format.ContainsKey("link"))
            node.Format["link"].ToString().Should().BeNullOrEmpty();
    }
}
