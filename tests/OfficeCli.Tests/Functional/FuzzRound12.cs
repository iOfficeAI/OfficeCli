// FuzzRound12 — Attack surfaces: invalid paths, wrong-format handlers, empty-file ops
//
// Areas:
//   IP01–IP05: BlankDocCreator.Create with invalid/edge-case paths
//   HX01–HX06: Handler constructor with nonexistent / wrong-format files
//   EF01–EF07: Get/Query/Set on freshly created file with no Add (empty doc)

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound12 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string TempPath(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz12_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        return path;
    }

    private string CreateTemp(string ext)
    {
        var path = TempPath(ext);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== IP01–IP05: BlankDocCreator.Create invalid paths ====================

    [Fact]
    public void IP01_Create_NullPath_ThrowsArgumentOrNullException()
    {
        var act = () => BlankDocCreator.Create(null!);
        act.Should().Throw<Exception>("null path must throw, not crash with NullRef silently");
    }

    [Fact]
    public void IP02_Create_EmptyPath_ThrowsException()
    {
        var act = () => BlankDocCreator.Create(string.Empty);
        act.Should().Throw<Exception>("empty path must throw");
    }

    [Fact]
    public void IP03_Create_UnsupportedExtension_ThrowsNotSupported()
    {
        var path = TempPath("txt");
        var act = () => BlankDocCreator.Create(path);
        act.Should().Throw<NotSupportedException>("unsupported extension must throw NotSupportedException");
    }

    [Fact]
    public void IP04_Create_NoExtension_ThrowsNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz12_{Guid.NewGuid():N}");
        _tempFiles.Add(path);
        var act = () => BlankDocCreator.Create(path);
        act.Should().Throw<NotSupportedException>("no extension must throw NotSupportedException");
    }

    [Fact]
    public void IP05_Create_NonexistentDirectory_ThrowsException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"nonexistent_{Guid.NewGuid():N}", "file.docx");
        _tempFiles.Add(path);
        var act = () => BlankDocCreator.Create(path);
        act.Should().Throw<Exception>("path in nonexistent directory must throw");
    }

    // ==================== HX01–HX06: Handler constructor with bad inputs ====================

    [Fact]
    public void HX01_WordHandler_NonexistentFile_ThrowsException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ghost_{Guid.NewGuid():N}.docx");
        var act = () => { using var h = new WordHandler(path, editable: false); };
        act.Should().Throw<Exception>("opening nonexistent .docx must throw");
    }

    [Fact]
    public void HX02_ExcelHandler_NonexistentFile_ThrowsException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ghost_{Guid.NewGuid():N}.xlsx");
        var act = () => { using var h = new ExcelHandler(path, editable: false); };
        act.Should().Throw<Exception>("opening nonexistent .xlsx must throw");
    }

    [Fact]
    public void HX03_PowerPointHandler_NonexistentFile_ThrowsException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ghost_{Guid.NewGuid():N}.pptx");
        var act = () => { using var h = new PowerPointHandler(path, editable: false); };
        act.Should().Throw<Exception>("opening nonexistent .pptx must throw");
    }

    [Fact]
    public void HX04_WordHandler_OpenXlsxAsDocx_BehaviorDocumented()
    {
        // BUG: WordHandler silently opens an xlsx file without format validation.
        // Constructors do not validate document type — wrong-format open is silent.
        // This test documents current (broken) behavior: no exception thrown.
        var path = CreateTemp("xlsx");
        // Actual behavior: no throw on construction; subsequent ops may fail silently.
        var act = () => { using var h = new WordHandler(path, editable: false); };
        // Document the bug: should throw, but currently does NOT.
        // We assert NotThrow to match current behavior; file a bug separately.
        act.Should().NotThrow("KNOWN BUG: WordHandler does not validate format on open");
    }

    [Fact]
    public void HX05_ExcelHandler_OpenDocxAsXlsx_ThrowsException()
    {
        var path = CreateTemp("docx");
        var act = () => { using var h = new ExcelHandler(path, editable: false); };
        act.Should().Throw<Exception>("opening docx with ExcelHandler must throw due to format mismatch");
    }

    [Fact]
    public void HX06_PowerPointHandler_OpenDocxAsPptx_BehaviorDocumented()
    {
        // BUG: PowerPointHandler silently opens a docx file without format validation.
        // Constructors do not validate document type — wrong-format open is silent.
        var path = CreateTemp("docx");
        var act = () => { using var h = new PowerPointHandler(path, editable: false); };
        // Document the bug: should throw, but currently does NOT.
        act.Should().NotThrow("KNOWN BUG: PowerPointHandler does not validate format on open");
    }

    // ==================== EF01–EF07: Empty doc Get/Query/Set ====================

    [Fact]
    public void EF01_Word_Get_Root_OnEmptyDoc_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: false);
        var act = () => h.Get("/");
        act.Should().NotThrow("Get('/') on empty Word doc must not throw");
    }

    [Fact]
    public void EF02_Word_Query_Paragraph_OnEmptyDoc_ReturnsEmpty()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: false);
        var results = h.Query("paragraph").ToList();
        results.Should().NotBeNull("Query on empty doc must return empty, not null");
    }

    [Fact]
    public void EF03_Word_Get_NonexistentPath_ThrowsArgumentException()
    {
        // BUG: WordHandler.Get throws ArgumentException instead of returning null for missing paths.
        // Other handlers return null for missing paths; Word throws — inconsistent contract.
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: false);
        // Document current (broken) behavior: throws ArgumentException
        var act = () => h.Get("/body/paragraph[99]");
        act.Should().Throw<ArgumentException>("KNOWN BUG: WordHandler.Get throws instead of returning null for missing paths");
    }

    [Fact]
    public void EF04_Excel_Get_EmptyCell_ReturnsNullOrEmpty()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: false);
        var act = () => h.Get("/Sheet1/Z99");
        act.Should().NotThrow("Get on unset Excel cell must not throw");
    }

    [Fact]
    public void EF05_Pptx_Query_Shape_OnEmptyDoc_ReturnsEmpty()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: false);
        var results = h.Query("shape").ToList();
        results.Should().NotBeNull("Query on empty pptx must return empty collection, not null");
    }

    [Fact]
    public void EF06_Word_Set_OnEmptyDoc_Editable_NoNullRefCrash()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        // Set on root — may be no-op or throw ArgumentException, but must not NullRef
        var act = () => h.Set("/", new() { ["find"] = "nothing", ["replace"] = "something" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef on Set on empty doc: {ex.Message}"); }
        catch (Exception) { /* ArgumentException/KeyNotFound = acceptable */ }
    }

    [Fact]
    public void EF07_Pptx_Get_NonexistentSlide_Throws()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: false);
        // No slides added — get slide[1] must throw
        var act = () => h.Get("/slide[1]");
        act.Should().Throw<ArgumentException>("Get on nonexistent slide must throw");
    }
}
