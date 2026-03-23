// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt round 21-35: Four targeted bugs found via white-box code review.
/// All tests are expected to FAIL until the bugs are fixed.
/// </summary>
public class PptxRound21Tests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext = ".pptx")
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // =========================================================================
    // Bug 1 — `get /theme` crashes with JSON serialization error
    //
    // Root cause: GetThemeNode() stores StringValue (OpenXML typed attribute)
    // for headingFont and bodyFont instead of plain string.
    // PowerPointHandler.Theme.cs lines ~177-180:
    //   var majorLatin = fontScheme?.MajorFont?.GetFirstChild<Drawing.LatinFont>()?.Typeface;
    //   if (!string.IsNullOrEmpty(majorLatin)) node.Format["headingFont"] = majorLatin;
    // `Typeface` is `StringValue`, not `string`. JSON serialization via
    // AppJsonContext fails because StringValue is not a registered type.
    // Fix: call `.Value` or cast to string before storing.
    // =========================================================================

    [Fact]
    public void Bug1_GetTheme_ShouldNotThrowJsonSerializationError()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: false);

        // Get the theme node — should not throw
        var node = handler.Get("/theme");

        node.Should().NotBeNull();
        node.Type.Should().Be("theme");
    }

    [Fact]
    public void Bug1_GetTheme_HeadingFontShouldBeSerializableToJson()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: false);

        var node = handler.Get("/theme");

        // If headingFont is present, it must be a plain string so JSON serialization works.
        // Storing StringValue (OpenXML typed) causes NotSupportedException in the source-gen
        // JsonSerializerContext used by the CLI output formatter.
        // Use OutputFormatter.FormatNode which uses AppJsonContext (source-gen context) internally.
        Action serialize = () =>
        {
            var json = OfficeCli.Core.OutputFormatter.FormatNode(node, OfficeCli.Core.OutputFormat.Json);
            json.Should().NotBeNullOrEmpty();
        };

        serialize.Should().NotThrow("Format values must be JSON-serializable plain types, not OpenXML StringValue");
    }

    [Fact]
    public void Bug1_GetTheme_HeadingFontValueShouldBePlainString()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: false);

        var node = handler.Get("/theme");

        if (node.Format.ContainsKey("headingFont"))
        {
            var val = node.Format["headingFont"];
            val.Should().BeOfType<string>(
                "headingFont must be stored as plain string, not DocumentFormat.OpenXml.StringValue");
        }

        if (node.Format.ContainsKey("bodyFont"))
        {
            var val = node.Format["bodyFont"];
            val.Should().BeOfType<string>(
                "bodyFont must be stored as plain string, not DocumentFormat.OpenXml.StringValue");
        }
    }

    // =========================================================================
    // Bug 2 — `slideSize` only recognizes widescreen/A4
    //
    // Root cause: PowerPointHandler.Set.cs slideSize switch statement only
    // handles "16:9"/"widescreen", "4:3"/"standard", "16:10", and "a4".
    // Missing: "letter", "b4", "b5", "35mm", "overhead", "banner", "custom",
    // and other SlideSizeValues enum members.
    // Fix: add missing preset cases, or map the full SlideSizeValues enum.
    // =========================================================================

    [Fact]
    public void Bug2_SetSlideSize_Letter_ShouldBeRecognized()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        // "letter" is a valid PowerPoint preset (SlideSizeValues.Letter)
        // but the switch statement doesn't handle it → falls into default → added to unsupported list
        // Fix: the unsupported list should be empty (no unrecognized keys).
        var unsupported = handler.Set("/", new() { ["slideSize"] = "letter" });
        unsupported.Should().BeEmpty(
            "letter is a valid PPTX slide size preset and should not be listed as unsupported");

        // Verify dimensions match US Letter landscape: 10x7.5 inches = 9144000 x 6858000 EMU
        var node = handler.Get("/");
        node.Format.Should().ContainKey("slideWidth");
        // Letter landscape width is 10 inches = 9144000 EMU → "25.4cm" or similar
        node.Format["slideWidth"].ToString().Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Bug2_SetSlideSize_B4_ShouldBeRecognized()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        // "b4" is SlideSizeValues.B4IsoBiggerThanA4Landscape
        // Currently unrecognized and silently ignored (added to unsupported list)
        var unsupported = handler.Set("/", new() { ["slideSize"] = "b4" });
        unsupported.Should().BeEmpty(
            "b4 is a valid PPTX slide size preset and should not be listed as unsupported");
    }

    [Fact]
    public void Bug2_SetSlideSize_35mm_ShouldBeRecognized()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        // "35mm" is SlideSizeValues.ThirtyFiveMm
        // Currently unrecognized and silently ignored (added to unsupported list)
        var unsupported = handler.Set("/", new() { ["slideSize"] = "35mm" });
        unsupported.Should().BeEmpty(
            "35mm is a valid PPTX slide size preset and should not be listed as unsupported");
    }

    [Fact]
    public void Bug2_SetSlideSize_Overhead_ShouldBeRecognized()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        // "overhead" is SlideSizeValues.Overhead
        // Currently unrecognized and silently ignored (added to unsupported list)
        var unsupported = handler.Set("/", new() { ["slideSize"] = "overhead" });
        unsupported.Should().BeEmpty(
            "overhead is a valid PPTX slide size preset and should not be listed as unsupported");
    }

    // =========================================================================
    // Bug 3 — `autofit` not returned in Get when no explicit autofit is set
    //
    // Root cause: PowerPointHandler.NodeBuilder.cs ShapeToNode() only adds
    // "autoFit" to Format when one of the three child elements (NormalAutoFit,
    // ShapeAutoFit, NoAutoFit) is explicitly present in the BodyProperties.
    // When none is present (the default for newly created shapes), "autoFit"
    // is missing entirely from Format — caller cannot distinguish between
    // "unknown", "default", and "none".
    // Additionally, the default behavior in PowerPoint for text overflow when
    // no autofit child is present is to overflow (equivalent to "none"), but
    // Get returns no key at all.
    // Fix: when no autofit element is present, explicitly report "autoFit" = "none"
    // (matching PowerPoint's default behavior).
    // =========================================================================

    [Fact]
    public void Bug3_Get_Shape_WithNoAutoFitSet_ShouldReturnDefaultAutoFit()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        // Add shape with NO autofit property — default behavior
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Default shape" });

        var node = handler.Get("/slide[1]/shape[1]");
        // When no autofit element is set, Get should still report a default value
        // so callers can distinguish "unknown" from "explicitly set to none"
        node.Format.Should().ContainKey("autoFit",
            "Get should always report autoFit even when not explicitly set, " +
            "so callers can detect the default state");
    }

    [Fact]
    public void Bug3_Add_Shape_WithAutoFitNone_ShouldBeReturnedAsNone()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "AutoFit none",
            ["autofit"] = "none"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("autoFit",
            "autoFit=none should be stored and returned by Get");
        node.Format["autoFit"].Should().Be("none");
    }

    [Fact]
    public void Bug3_Set_Shape_AutoFit_Normal_ShouldPersistAcrossReopen()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });
            handler.Set("/slide[1]/shape[1]", new() { ["autofit"] = "normal" });
        }

        // Reopen and verify persistence
        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("autoFit",
            "autoFit must be persisted across file close/reopen");
        node.Format["autoFit"].Should().Be("normal");
    }

    // =========================================================================
    // Bug 4 — Cannot remove notes
    //
    // Root cause: PowerPointHandler.Mutations.cs Remove() uses a regex that
    // only matches:
    //   /slide[N]            (remove whole slide)
    //   /slide[N]/element[M] (remove element from slide)
    // The path "/slide[N]/notes" does NOT match this pattern (word+no-index),
    // so it throws ArgumentException immediately:
    //   "Invalid path: /slide[1]/notes. Expected format: /slide[N] or /slide[N]/element[M]"
    // Fix: add a specific check for /slide[N]/notes before the main regex, and
    // remove/clear the NotesSlidePart accordingly.
    // =========================================================================

    [Fact]
    public void Bug4_Remove_Notes_ShouldNotThrowArgumentException()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide with notes" });
        // Add notes to the slide
        handler.Add("/slide[1]", "notes", null, new() { ["text"] = "Speaker note content" });

        // Verify notes were added
        var notesNode = handler.Get("/slide[1]/notes");
        notesNode.Text.Should().Contain("Speaker note content");

        // Remove should NOT throw — currently throws ArgumentException
        var act = () => handler.Remove("/slide[1]/notes");
        act.Should().NotThrow<ArgumentException>(
            "Remove should support /slide[N]/notes path to clear speaker notes");
    }

    [Fact]
    public void Bug4_Remove_Notes_ShouldClearNotesText()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide with notes" });
        handler.Add("/slide[1]", "notes", null, new() { ["text"] = "My speaker notes" });

        // Verify notes exist
        var before = handler.Get("/slide[1]/notes");
        before.Text.Should().Contain("My speaker notes");

        // Remove notes
        handler.Remove("/slide[1]/notes");

        // After removal, notes should not exist (Get returns null)
        var after = handler.Get("/slide[1]/notes");
        after.Should().BeNull(
            "After Remove(/slide[1]/notes), Get should return null");
    }

    [Fact]
    public void Bug4_Remove_Notes_ShouldPersistClearAcrossReopen()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
            handler.Add("/slide[1]", "notes", null, new() { ["text"] = "Temporary notes" });
            handler.Remove("/slide[1]/notes");
        }

        // Reopen and verify notes are gone
        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/notes");
        node.Should().BeNull(
            "Notes removal must persist after file is closed and reopened");
    }
}
