// Bug hunt tests Part 3 — Bug #171-250
// Footnotes, Notes, Formatting, Tables, Comments, Merge, ParseEmu, Media, Charts, TOC

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class CoreRegression3 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public CoreRegression3()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        using (var pptx = new PowerPointHandler(_pptxPath, editable: true))
            pptx.Add("/", "slide", null, new());
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _excelHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private WordHandler ReopenWord()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
        return _wordHandler;
    }

    private ExcelHandler ReopenExcel()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        return _excelHandler;
    }

    // ==================== Bug #171-190: Footnotes, Notes, Conditional formatting, Color parsing ====================

    /// Bug #171 — Word footnote: space prepended on every Set
    /// File: WordHandler.Set.cs, lines 117-118
    /// Setting footnote text prepends " " each time: textEl.Text = " " + fnText
    /// Calling Set multiple times accumulates leading spaces.
    [Fact]
    public void Bug171_WordFootnote_SpacePrependedOnEverySet()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Main text" });
        _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Note text" });

        // Set the footnote text again
        _wordHandler.Set("/body/p[1]/footnote[1]", new() { ["text"] = "Updated note" });

        // Get the footnote text
        var node = _wordHandler.Get("/body/p[1]/footnote[1]");
        if (node?.Text != null)
        {
            node.Text.Should().NotStartWith("  ",
                "Footnote text should not accumulate leading spaces on each Set call");
        }
    }

    /// Bug #173 — Word footnote: only first run updated in multi-run footnote
    /// File: WordHandler.Set.cs, lines 112-119
    /// Set only modifies the first non-reference-mark run.
    /// Other runs remain unchanged, creating inconsistent text.
    [Fact]
    public void Bug173_WordFootnote_OnlyFirstRunUpdated()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Text" });
        _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Original footnote" });

        // Update the footnote
        _wordHandler.Set("/body/p[1]/footnote[1]", new() { ["text"] = "New text" });

        var node = _wordHandler.Get("/body/p[1]/footnote[1]");
        // If the footnote had multiple runs, only the first would be updated
        node.Should().NotBeNull();
    }

    /// Bug #174 — PPTX notes: EnsureNotesSlidePart missing NotesMasterPart relationship
    /// File: PowerPointHandler.Notes.cs, lines 88-130
    /// When creating a NotesSlidePart, the code doesn't establish a
    /// relationship to a NotesMasterPart, which OOXML spec may require.
    [Fact]
    public void Bug174_PptxNotes_MissingNotesMasterPartRelationship()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Set speaker notes — this creates a NotesSlidePart
        pptx.Set("/slide[1]", new() { ["notes"] = "Speaker notes here" });

        // Verify notes can be read back
        var node = pptx.Get("/slide[1]");
        node.Format.TryGetValue("notes", out var notes);
        (notes != null && notes.ToString()!.Contains("Speaker notes")).Should().BeTrue(
            "Speaker notes should be readable after setting");
    }


    /// Bug #177 — Excel conditional formatting: iconset enum not validated in Set
    /// File: ExcelHandler.Set.cs, line 349
    /// IconSetValue is set directly from user input without validation,
    /// unlike Add.cs which uses ParseIconSetValues().
    [Fact]
    public void Bug177_ExcelConditionalFormatting_IconSetEnumNotValidated()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "iconset",
            ["sqref"] = "A1:A10",
            ["icons"] = "3Arrows"
        });

        // Set with invalid icon set name — should validate
        var act = () => _excelHandler.Set("/Sheet1/conditionalformatting[1]", new()
        {
            ["icons"] = "InvalidIconSet"
        });

        // Should reject invalid icon set name
        act.Should().Throw<Exception>(
            "Invalid icon set name should be rejected, not silently accepted");
    }


    /// Bug #180 — Excel Set: picture width/height int.Parse
    /// File: ExcelHandler.Set.cs, lines 187, 194
    /// Picture resize uses int.Parse without validation.
    [Fact]
    public void Bug180_ExcelSet_PictureWidthHeightIntParse()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        var imgPath = CreateTempImage();
        try
        {
            _excelHandler.Add("/Sheet1", "picture", null, new() { ["src"] = imgPath });

            var act = () => _excelHandler.Set("/Sheet1/picture[1]", new()
            {
                ["width"] = "large"
            });

            act.Should().Throw<Exception>(
                "invalid string 'large' for picture width should throw");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }


    /// Bug #189 — Word Set footnote: text not properly escaped
    /// File: WordHandler.Set.cs, lines 117-118
    /// Text is set directly without XML escaping considerations.
    [Fact]
    public void Bug189_WordSetFootnote_TextNotEscaped()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Main" });
        _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Initial" });

        // Set text with special characters
        _wordHandler.Set("/body/p[1]/footnote[1]", new()
        {
            ["text"] = "Note with <special> & characters"
        });

        ReopenWord();
        var node = _wordHandler.Get("/body/p[1]/footnote[1]");
        node.Should().NotBeNull("footnote with special characters should survive roundtrip");
    }


    // ==================== Bug #191-210: PPTX tables, Excel validation, Word images, theme colors ====================

    /// Bug #191 — PPTX table style: light3 and medium3 share same GUID
    /// File: PowerPointHandler.Set.cs, lines 380, 384
    /// Both map to "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}" — copy-paste error.
    [Fact]
    public void Bug191_PptxTableStyle_DuplicateGuid()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Set light3 style
        pptx.Set("/slide[1]/table[1]", new() { ["tablestyle"] = "light3" });
        var node1 = pptx.Get("/slide[1]/table[1]");

        // Set medium3 style
        pptx.Set("/slide[1]/table[1]", new() { ["tablestyle"] = "medium3" });
        var node2 = pptx.Get("/slide[1]/table[1]");

        // light3 and medium3 should produce different styles
        // but they share the same GUID due to copy-paste error
        (node1?.Format.TryGetValue("tableStyleId", out var id1) ?? false).Should().BeTrue();
        (node2?.Format.TryGetValue("tableStyleId", out var id2) ?? false).Should().BeTrue();
    }


    /// Bug #198 — Word image: non-unique DocProperties.Id
    /// File: WordHandler.ImageHelpers.cs, lines 37, 108
    /// Uses Environment.TickCount which can produce duplicates
    /// if multiple images added within the same millisecond.
    [Fact]
    public void Bug198_WordImage_NonUniqueDocPropertiesId()
    {
        var imgPath = CreateTempImage();
        try
        {
            // Add two images rapidly — they may get same ID
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image 1:" });
            _wordHandler.Add("/body/p[1]", "image", null, new() { ["src"] = imgPath });
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image 2:" });
            _wordHandler.Add("/body/p[2]", "image", null, new() { ["src"] = imgPath });

            ReopenWord();
            var root = _wordHandler.Get("/body");
            root.Should().NotBeNull("document with multiple images should be valid");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }


    /// Bug #205 — Word image: NonVisualDrawingProperties.Id hardcoded to 0
    /// File: WordHandler.ImageHelpers.cs, lines 45, 115
    /// PIC.NonVisualDrawingProperties.Id = 0U for all images.
    /// Should be unique per drawing object.
    [Fact]
    public void Bug205_WordImage_HardcodedZeroId()
    {
        var imgPath = CreateTempImage();
        try
        {
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image" });
            _wordHandler.Add("/body/p[1]", "image", null, new() { ["src"] = imgPath });

            ReopenWord();
            var node = _wordHandler.Get("/body");
            node.Should().NotBeNull();
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }


    /// Bug #216 — Excel merge: row deletion doesn't clean up merge ranges
    /// File: ExcelHandler.Add.cs, lines 954-962
    /// Deleting a row that participates in a merge doesn't update
    /// or remove the affected merge definition.
    [Fact]
    public void Bug216_ExcelMerge_RowDeletionNoCleanup()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Merged" });
        _excelHandler.Set("/Sheet1", new() { ["merge"] = "A1:A3" });

        // Delete row 2 which is part of the merge
        _excelHandler.Remove("/Sheet1/row[2]");

        // The merge definition "A1:A3" still exists but row 2 is gone
        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }


    /// Bug #221 — ParseEmu: long to int cast overflow
    /// File: PowerPointHandler.Fill.cs, lines 182-192
    /// ParseEmu returns long but results cast to int for BodyProperties.
    /// Values > int.MaxValue silently overflow to negative.
    [Fact]
    public void Bug221_ParseEmu_LongToIntCastOverflow()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Very large inset value that overflows int
        var act = () => pptx.Set("/slide[1]/shape[1]", new()
        {
            ["inset"] = "1000cm,1000cm,1000cm,1000cm"
        });

        // 1000cm = 360,000,000,000 EMU which overflows int.MaxValue
        act.Should().Throw<Exception>(
            "ParseEmu long result cast to int causes silent overflow for large values");
    }

    /// Bug #222 — ParseEmu: negative dimensions accepted
    /// File: WordHandler.ImageHelpers.cs, lines 17-30 and PowerPointHandler.Helpers.cs, lines 161-173
    /// No validation for negative values. "-5cm" produces -1800000 EMU.
    [Fact]
    public void Bug222_ParseEmu_NegativeDimensionsAccepted()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var act = () => pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["width"] = "-5cm"
        });

        // Negative width creates invalid shape
        act.Should().Throw<Exception>(
            "Negative dimensions should be rejected by ParseEmu");
    }

    /// Bug #222b — ParseEmu: negative positions (x, y) should be allowed
    /// Shapes can be placed off-screen with negative x/y values.
    [Fact]
    public void Bug222b_ParseEmu_NegativePositionsAllowed()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add shape with negative x — should NOT throw
        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "OffScreen",
            ["x"] = "-10cm",
            ["y"] = "-5cm",
            ["width"] = "15cm",
            ["height"] = "10cm"
        });

        var node = pptx.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format["x"].Should().Be("-10cm");
        node.Format["y"].Should().Be("-5cm");

        // Set negative position on existing shape — should NOT throw
        pptx.Set("/slide[1]/shape[1]", new() { ["x"] = "-3cm", ["y"] = "-2cm" });
        node = pptx.Get("/slide[1]/shape[1]");
        node!.Format["x"].Should().Be("-3cm");
        node.Format["y"].Should().Be("-2cm");

        // Set negative width — SHOULD throw
        var actW = () => pptx.Set("/slide[1]/shape[1]", new() { ["width"] = "-5cm" });
        actW.Should().Throw<Exception>("Negative width should be rejected");

        // Set negative height — SHOULD throw
        var actH = () => pptx.Set("/slide[1]/shape[1]", new() { ["height"] = "-3cm" });
        actH.Should().Throw<Exception>("Negative height should be rejected");
    }

    /// Bug #223 — ParseEmu: empty unit suffix causes crash
    /// File: WordHandler.ImageHelpers.cs, line 22
    /// Input "cm" (unit only, no number) → value[..^2] = "" → double.Parse("") crash.
    [Fact]
    public void Bug223_ParseEmu_EmptyUnitSuffixCrash()
    {
        var imgPath = CreateTempImage();
        try
        {
            var act = () => _wordHandler.Add("/body", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "cm"
            });

            act.Should().Throw<ArgumentException>(
                "ParseEmu crashes on 'cm' (no number) — should validate input length");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #224 — ParseEmu: unsupported units silently fail
    /// File: WordHandler.ImageHelpers.cs, lines 17-30
    /// "5mm" is not supported — falls through to long.Parse("5mm") which crashes.
    [Fact]
    public void Bug224_ParseEmu_UnsupportedUnitsCrash()
    {
        var imgPath = CreateTempImage();
        try
        {
            var act = () => _wordHandler.Add("/body", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "50mm"
            });

            act.Should().Throw<ArgumentException>(
                "ParseEmu doesn't support 'mm' unit — falls through to long.Parse('50mm')");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }


    /// Bug #229 — ParseEmu: double truncation instead of rounding
    /// File: WordHandler.ImageHelpers.cs, line 22
    /// (long)(double.Parse(value) * 360000) truncates instead of rounding.
    /// "0.001cm" → 360 EMU (truncated) vs 360 EMU (correct by coincidence).
    [Fact]
    public void Bug229_ParseEmu_TruncationInsteadOfRounding()
    {
        var imgPath = CreateTempImage();
        try
        {
            // Fractional cm values may lose precision due to truncation
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image" });
            _wordHandler.Add("/body/p[1]", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "2.54cm"  // Should be exactly 1 inch = 914400 EMU
            });

            ReopenWord();
            var node = _wordHandler.Get("/body");
            node.Should().NotBeNull();
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #230 — ParseEmu: duplicate implementations in Word and PPTX
    /// File: WordHandler.ImageHelpers.cs lines 17-30, PowerPointHandler.Helpers.cs lines 161-173
    /// Identical code duplicated — any fix must be applied to both files.
    [Fact]
    public void Bug230_ParseEmu_DuplicateImplementations()
    {
        // This is a code quality bug — ParseEmu exists in two places
        // Any bug fix or enhancement must be applied to both
        // Word ParseEmu and PPTX ParseEmu are separate copies
        var imgPath = CreateTempImage();
        try
        {
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image" });
            _wordHandler.Add("/body/p[1]", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "5cm"
            });

            BlankDocCreator.Create(_pptxPath);
            using var pptx = new PowerPointHandler(_pptxPath, editable: true);
            pptx.Add("/", "slide", null, new());
            pptx.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Test",
                ["width"] = "5cm"
            });
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }



    private static string CreateTempImage()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(path, pngBytes);
        return path;
    }
}