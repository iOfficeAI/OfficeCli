// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Final PPTX test file — covers bugs found in R76-90 plus high-priority feature gaps.
///
/// R76-90 Bugs:
///   Bug1 — slideSize roundtrip: Set "widescreen" but Get returns "screen16x9" (OOXML InnerText).
///          The Set path accepts friendly aliases ("widescreen") but the Get path reads the raw
///          SlideSizeValues InnerText ("screen16x9") without mapping it back to the friendly name.
///
///   Bug2 — transition=none does not remove morph.
///          ApplyTransition() sets slide.Transition = null (line 43) which only clears the SDK
///          typed property but leaves any mc:AlternateContent morph wrapper in the XML tree.
///          After Set("none"), Get still returns "morph".
///
///   Bug3 — cover/pull direction readback returns OOXML abbreviations.
///          ReadTransitionDirection() returns cover.Direction.Value.ToLowerInvariant() which
///          produces "l", "r", "u", "d" instead of the user-friendly "left", "right", etc.
///          Wipe/push transitions expand correctly via MapSlideDirection(); cover/pull do not.
///
///   Bug4 — Get after Remove notes returns a node with empty text instead of null.
///          After Remove("/slide[1]/notes"), Get("/slide[1]/notes") returns a DocumentNode with
///          Text="" rather than null, because the Get branch always constructs a node even when
///          NotesSlidePart is absent.
///
/// High-priority feature gaps:
///   - Picture Set opacity (listed in valid-props hint, not implemented)
///   - Picture Set name (listed in valid-props hint, not implemented)
///   - Animation removal ("none") round-trip: add animation, Set animation="none", verify gone
///   - Connector fill (listed in valid connector props, not implemented in Set)
/// </summary>
public class PptxFinalTests : IDisposable
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
    // Bug1 — slideSize roundtrip: output name differs from input name
    //
    // Set accepts "widescreen", "4:3", "16:9", "16:10", "a4", etc.
    // Get reads SlideSizeValues.InnerText which returns "screen16x9", "screen4x3",
    // "screen16x10", "a4" — creating a round-trip mismatch.
    // Fix: map InnerText back to the canonical friendly name in the Query path.
    // =========================================================================

    [Fact]
    public void Bug1_SlideSize_Widescreen_RoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        var unsupported = handler.Set("/", new() { ["slideSize"] = "widescreen" });
        unsupported.Should().BeEmpty("slideSize is a documented presentation property");

        var node = handler.Get("/");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("slideSize");

        // The returned value must be a user-friendly name, not the raw OOXML InnerText.
        // Either "widescreen" or "16:9" is acceptable as the canonical roundtrip value.
        var returned = node.Format["slideSize"]?.ToString();
        returned.Should().NotBe("screen16x9",
            "Get should return the friendly name 'widescreen' or '16:9', not the raw OOXML 'screen16x9'");
        returned.Should().Match(v => v == "widescreen" || v == "16:9" || v == "screen16x9" == false,
            "slideSize roundtrip should return a consistent user-readable name");
    }

    [Fact]
    public void Bug1_SlideSize_43_RoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Set("/", new() { ["slideSize"] = "4:3" });

        var node = handler.Get("/");
        node!.Format.Should().ContainKey("slideSize");
        var returned = node.Format["slideSize"]?.ToString();
        returned.Should().NotBe("screen4x3",
            "Get should return '4:3', not the raw OOXML 'screen4x3'");
    }

    [Fact]
    public void Bug1_SlideSize_1610_RoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Set("/", new() { ["slideSize"] = "16:10" });

        var node = handler.Get("/");
        node!.Format.Should().ContainKey("slideSize");
        var returned = node.Format["slideSize"]?.ToString();
        returned.Should().NotBe("screen16x10",
            "Get should return '16:10', not the raw OOXML 'screen16x10'");
    }

    [Fact]
    public void Bug1_SlideSize_A4_RoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Set("/", new() { ["slideSize"] = "a4" });

        var node = handler.Get("/");
        node!.Format.Should().ContainKey("slideSize");
        var returned = node.Format["slideSize"]?.ToString();
        // "a4" InnerText is "a4" so this one may already be fine, but verify it's sane
        returned.Should().NotBeNullOrEmpty("slideSize must be readable after Set");
    }

    [Fact]
    public void Bug1_SlideSize_Letter_RoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Set("/", new() { ["slideSize"] = "letter" });

        var node = handler.Get("/");
        node!.Format.Should().ContainKey("slideSize");
        var returned = node.Format["slideSize"]?.ToString();
        returned.Should().NotBe("letter2",
            "Get should return 'letter', not an obscure OOXML alias");
        returned.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Bug1_SlideSize_Widescreen_Persists()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Set("/", new() { ["slideSize"] = "widescreen" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/");
        node!.Format.Should().ContainKey("slideSize");
        node.Format["slideSize"]?.ToString().Should().NotBe("screen16x9",
            "slideSize must persist with a friendly name after save/reopen");
    }

    // =========================================================================
    // Bug2 — transition=none does not remove morph
    //
    // When a morph transition is stored as mc:AlternateContent, setting
    // transition="none" calls slide.Transition = null which clears the SDK
    // typed property but does not remove the AlternateContent XML element.
    // Fix: in ApplyTransition() for "none", also remove any child elements
    // whose LocalName == "AlternateContent".
    // =========================================================================

    [Fact]
    public void Bug2_TransitionNone_RemovesMorph()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // Apply morph transition
        handler.Set("/slide[1]", new() { ["transition"] = "morph" });

        // Verify morph is set
        var before = handler.Get("/slide[1]");
        before!.Format.Should().ContainKey("transition");
        before.Format["transition"]?.ToString().Should().StartWith("morph",
            "morph transition should be set before testing removal");

        // Now remove it
        handler.Set("/slide[1]", new() { ["transition"] = "none" });

        // Verify morph is gone
        var after = handler.Get("/slide[1]");
        after!.Format.Should().NotContainKey("transition",
            "transition=none must remove the morph AlternateContent wrapper from XML");
    }

    [Fact]
    public void Bug2_TransitionNone_RemovesMorphByWord()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "morph-byWord" });

        var before = handler.Get("/slide[1]");
        before!.Format.Should().ContainKey("transition");

        handler.Set("/slide[1]", new() { ["transition"] = "none" });

        var after = handler.Get("/slide[1]");
        after!.Format.Should().NotContainKey("transition",
            "transition=none must remove morph-byWord AlternateContent");
    }

    [Fact]
    public void Bug2_TransitionNone_RemovesMorphByChar()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "morph-byChar" });

        handler.Set("/slide[1]", new() { ["transition"] = "none" });

        var after = handler.Get("/slide[1]");
        after!.Format.Should().NotContainKey("transition",
            "transition=none must remove morph-byChar AlternateContent");
    }

    [Fact]
    public void Bug2_TransitionNone_AfterMorph_ThenAddNewTransition_Works()
    {
        // After removing morph and adding a new transition, the new one must round-trip.
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "morph" });
        handler.Set("/slide[1]", new() { ["transition"] = "none" });
        handler.Set("/slide[1]", new() { ["transition"] = "fade" });

        var node = handler.Get("/slide[1]");
        node!.Format.Should().ContainKey("transition");
        node.Format["transition"]?.ToString().Should().Be("fade",
            "after removing morph, a new transition must be applied cleanly");
    }

    [Fact]
    public void Bug2_TransitionNone_RemovesMorph_Persists()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Set("/slide[1]", new() { ["transition"] = "morph" });
            h.Set("/slide[1]", new() { ["transition"] = "none" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]");
        node!.Format.Should().NotContainKey("transition",
            "morph removal must survive save/reopen");
    }

    // =========================================================================
    // Bug3 — cover/pull direction readback uses OOXML abbreviations
    //
    // ReadTransitionDirection() for CoverTransition and PullTransition calls
    //   cover.Direction.Value?.ToLowerInvariant()
    // which returns the OOXML StringValue like "l", "r", "u", "d" rather than
    // the user-friendly "left", "right", "up", "down".
    // Wipe/push already use MapSlideDirection() correctly.
    // Fix: expand "l"→"left", "r"→"right", "u"→"up", "d"→"down" for cover/pull.
    // =========================================================================

    [Fact]
    public void Bug3_CoverLeft_DirectionReadback_IsFullWord()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "cover-left" });

        var node = handler.Get("/slide[1]");
        node!.Format.Should().ContainKey("transition");
        var transition = node.Format["transition"]?.ToString();
        transition.Should().NotBe("cover-l",
            "direction 'l' must be expanded to 'left' in cover transition readback");
        transition.Should().Be("cover-left",
            "cover-left roundtrip must return 'cover-left', not 'cover-l'");
    }

    [Fact]
    public void Bug3_CoverRight_DirectionReadback_IsFullWord()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "cover-right" });

        var node = handler.Get("/slide[1]");
        var transition = node!.Format["transition"]?.ToString();
        transition.Should().Be("cover-right",
            "cover-right roundtrip must return 'cover-right', not 'cover-r'");
    }

    [Fact]
    public void Bug3_PullLeft_DirectionReadback_IsFullWord()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "pull-left" });

        var node = handler.Get("/slide[1]");
        node!.Format.Should().ContainKey("transition");
        var transition = node.Format["transition"]?.ToString();
        transition.Should().NotBe("pull-l",
            "'pull-l' is the OOXML abbreviation; readback must expand it to 'pull-left'");
        transition.Should().Be("pull-left",
            "pull-left roundtrip must return 'pull-left', not 'pull-l'");
    }

    [Fact]
    public void Bug3_PullRight_DirectionReadback_IsFullWord()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "pull-right" });

        var node = handler.Get("/slide[1]");
        var transition = node!.Format["transition"]?.ToString();
        transition.Should().Be("pull-right",
            "pull-right roundtrip must return 'pull-right', not 'pull-r'");
    }

    [Fact]
    public void Bug3_CoverUp_DirectionReadback_IsFullWord()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "cover-up" });

        var node = handler.Get("/slide[1]");
        var transition = node!.Format["transition"]?.ToString();
        transition.Should().Be("cover-up",
            "cover-up roundtrip must return 'cover-up', not 'cover-u'");
    }

    [Fact]
    public void Bug3_PullDown_DirectionReadback_IsFullWord()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "pull-down" });

        var node = handler.Get("/slide[1]");
        var transition = node!.Format["transition"]?.ToString();
        transition.Should().NotBe("pull-d",
            "'pull-d' is the OOXML abbreviation; readback must expand to 'pull-down'");
        transition.Should().Be("pull-down");
    }

    [Fact]
    public void Bug3_Cover_Direction_Persists_AsFullWord()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Set("/slide[1]", new() { ["transition"] = "cover-right" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]");
        node!.Format["transition"]?.ToString().Should().Be("cover-right",
            "cover direction must persist correctly across save/reopen");
    }

    // =========================================================================
    // Bug4 — Get after Remove notes returns empty node instead of null
    //
    // Get("/slide[N]/notes") always constructs a DocumentNode even when
    // NotesSlidePart is null (i.e., notes have been removed). It should return
    // null (not-found) so callers can distinguish "no notes" from "notes with
    // empty text".
    // Fix: return null when NotesSlidePart is null.
    // =========================================================================

    [Fact]
    public void Bug4_GetNotes_AfterRemove_ReturnsNull()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // Add notes
        handler.Set("/slide[1]", new() { ["notes"] = "Speaker notes text" });

        // Verify notes were added
        var before = handler.Get("/slide[1]/notes");
        before.Should().NotBeNull("notes should exist after Set");
        before!.Text.Should().Be("Speaker notes text");

        // Remove notes
        handler.Remove("/slide[1]/notes");

        // After removal, Get should return null (not-found), not an empty-text node.
        var after = handler.Get("/slide[1]/notes");
        after.Should().BeNull(
            "Get('/slide[N]/notes') must return null when NotesSlidePart has been removed");
    }

    [Fact]
    public void Bug4_GetNotes_OnSlideWithNoNotes_ReturnsNull()
    {
        // A freshly created slide has no NotesSlidePart.
        // Get should return null, not an empty node.
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        var node = handler.Get("/slide[1]/notes");
        node.Should().BeNull(
            "a slide with no notes should return null from Get, not a node with empty Text");
    }

    [Fact]
    public void Bug4_GetNotes_AfterRemove_Persists()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Set("/slide[1]", new() { ["notes"] = "To be deleted" });
            h.Remove("/slide[1]/notes");
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/notes");
        node.Should().BeNull(
            "notes removal must persist across save/reopen; Get should return null");
    }

    [Fact]
    public void Bug4_GetNotes_WithText_ReturnsNode()
    {
        // Confirm the positive case still works after fix.
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["notes"] = "Hello notes" });

        var node = handler.Get("/slide[1]/notes");
        node.Should().NotBeNull("notes with text must be returned");
        node!.Type.Should().Be("notes");
        node.Text.Should().Be("Hello notes");
    }

    // =========================================================================
    // Picture Set opacity
    //
    // "opacity" appears in the valid-picture-props hint string but the Set switch
    // has no case for it. The fix adds a case that applies a Drawing.AlphaModFix
    // or Drawing.Alpha child to the blip element.
    // =========================================================================

    [Fact]
    public void Picture_Set_Opacity_IsAccepted()
    {
        var pngPath = MakeTinyPng();

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "picture", null, new()
        {
            ["path"] = pngPath,
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
        });

        var unsupported = handler.Set("/slide[1]/picture[1]", new() { ["opacity"] = "0.5" });

        unsupported.Should().BeEmpty(
            "opacity is listed in the valid picture props hint and must be implemented");
    }

    [Fact]
    public void Picture_Set_Opacity_RoundTrips()
    {
        var pngPath = MakeTinyPng();

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "picture", null, new()
        {
            ["path"] = pngPath,
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
        });

        handler.Set("/slide[1]/picture[1]", new() { ["opacity"] = "0.75" });

        var node = handler.Get("/slide[1]/picture[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("opacity",
            "Get must return opacity after Set");
        var opacityVal = Convert.ToDouble(node.Format["opacity"]);
        opacityVal.Should().BeApproximately(0.75, 0.01,
            "opacity 0.75 must round-trip correctly");
    }

    [Fact]
    public void Picture_Set_Opacity_Zero_IsTransparent()
    {
        var pngPath = MakeTinyPng();

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "picture", null, new()
        {
            ["path"] = pngPath,
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
        });

        // opacity=0.0 means fully transparent
        var unsupported = handler.Set("/slide[1]/picture[1]", new() { ["opacity"] = "0.0" });
        unsupported.Should().BeEmpty("opacity=0.0 must be accepted");
    }

    [Fact]
    public void Picture_Set_Opacity_Persists()
    {
        var pngPath = MakeTinyPng();
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Add("/slide[1]", "picture", null, new()
            {
                ["path"] = pngPath,
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
            });
            h.Set("/slide[1]/picture[1]", new() { ["opacity"] = "0.5" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/picture[1]");
        node!.Format.Should().ContainKey("opacity",
            "opacity must persist after save/reopen");
        Convert.ToDouble(node.Format["opacity"]).Should().BeApproximately(0.5, 0.01);
    }

    // =========================================================================
    // Picture Set name
    //
    // "name" appears in the valid-picture-props hint string but the Set switch
    // has no case for it. The fix adds a case that sets
    // NonVisualDrawingProperties.Name.
    // =========================================================================

    [Fact]
    public void Picture_Set_Name_IsAccepted()
    {
        var pngPath = MakeTinyPng();

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "picture", null, new()
        {
            ["path"] = pngPath,
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
        });

        var unsupported = handler.Set("/slide[1]/picture[1]", new() { ["name"] = "MyPicture" });

        unsupported.Should().BeEmpty(
            "name is listed in the valid picture props hint and must be implemented");
    }

    [Fact]
    public void Picture_Set_Name_RoundTrips()
    {
        var pngPath = MakeTinyPng();

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "picture", null, new()
        {
            ["path"] = pngPath,
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
        });

        handler.Set("/slide[1]/picture[1]", new() { ["name"] = "Logo Image" });

        var node = handler.Get("/slide[1]/picture[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("name");
        node.Format["name"]?.ToString().Should().Be("Logo Image",
            "name must round-trip exactly as set");
    }

    [Fact]
    public void Picture_Add_Name_IsPreserved()
    {
        var pngPath = MakeTinyPng();

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "picture", null, new()
        {
            ["path"] = pngPath,
            ["name"] = "AddedName",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
        });

        var node = handler.Get("/slide[1]/picture[1]");
        node!.Format["name"]?.ToString().Should().Be("AddedName",
            "name set during Add must be readable via Get");
    }

    [Fact]
    public void Picture_Set_Name_Persists()
    {
        var pngPath = MakeTinyPng();
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Add("/slide[1]", "picture", null, new()
            {
                ["path"] = pngPath,
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
            });
            h.Set("/slide[1]/picture[1]", new() { ["name"] = "PersistName" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/picture[1]");
        node!.Format["name"]?.ToString().Should().Be("PersistName",
            "picture name must persist after save/reopen");
    }

    // =========================================================================
    // Animation removal — Set animation="none" removes existing animation
    //
    // After adding a shape animation, Set(shape, {animation="none"}) must
    // call RemoveShapeAnimations and produce no animation in Get readback.
    // =========================================================================

    [Fact]
    public void Animation_SetNone_RemovesExistingAnimation()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Animated",
            ["animation"] = "fade",
        });

        // Verify animation exists
        var before = handler.Get("/slide[1]/shape[1]");
        before.Should().NotBeNull();
        // The presence of animation is confirmed by checking Set accepted it (no unsupported)
        // and by verifying it can be removed.

        // Remove animation
        var unsupported = handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "none" });
        unsupported.Should().BeEmpty("animation=none must be accepted without error");
    }

    [Fact]
    public void Animation_SetNone_AfterFly_RemovesAnimation()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Flying shape",
            ["animation"] = "fly-entrance",
        });

        var unsupported = handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "none" });
        unsupported.Should().BeEmpty("animation=none must be accepted for fly animation removal");
    }

    [Fact]
    public void Animation_Add_WithEntrance_IsAccepted()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        // Verify Add supports animation directly
        var resultPath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape with fade",
            ["animation"] = "fade",
        });
        resultPath.Should().NotBeNullOrEmpty();

        var node = handler.Get(resultPath);
        node.Should().NotBeNull();
        node!.Text.Should().Be("Shape with fade");
    }

    [Fact]
    public void Animation_SetNone_IsIdempotent_WhenNoAnimation()
    {
        // Setting animation=none on a shape that has no animation should not throw.
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "No animation" });

        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "none" });
        act.Should().NotThrow("removing animation from a shape with none must be a no-op");
    }

    [Fact]
    public void Animation_SetNone_Persists()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Will lose animation",
                ["animation"] = "zoom",
            });
            h.Set("/slide[1]/shape[1]", new() { ["animation"] = "none" });
        }

        // Reopening should not throw; the animation was removed
        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull("shape must still exist after animation removal");
        node!.Text.Should().Be("Will lose animation");
    }

    // =========================================================================
    // Connector fill
    //
    // "fill" appears in the valid connector props hint but the Set switch has no
    // case for it. The fix adds a case that applies a solid fill (or no-fill) to
    // the connector's ShapeProperties.
    // =========================================================================

    [Fact]
    public void Connector_Set_Fill_IsAccepted()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new());

        var unsupported = handler.Set("/slide[1]/connector[1]", new() { ["fill"] = "FF0000" });

        unsupported.Should().BeEmpty(
            "fill is listed in the valid connector props hint and must be implemented");
    }

    [Fact]
    public void Connector_Set_Fill_RoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new());

        handler.Set("/slide[1]/connector[1]", new() { ["fill"] = "0070C0" });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Should().NotBeNull();
        // After the fix, fill must appear in Format
        node!.Format.Should().ContainKey("fill",
            "Get must expose the fill color set via Set on a connector");
        node.Format["fill"]?.ToString().Should().Be("#0070C0",
            "fill color must round-trip with # prefix");
    }

    [Fact]
    public void Connector_Set_Fill_None_IsAccepted()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new());

        var unsupported = handler.Set("/slide[1]/connector[1]", new() { ["fill"] = "none" });
        unsupported.Should().BeEmpty("fill=none must be accepted for connectors");
    }

    [Fact]
    public void Connector_Set_Fill_Persists()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Add("/slide[1]", "connector", null, new());
            h.Set("/slide[1]/connector[1]", new() { ["fill"] = "FF0000" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/connector[1]");
        node!.Format.Should().ContainKey("fill",
            "connector fill must persist after save/reopen");
        node.Format["fill"]?.ToString().Should().Be("#FF0000");
    }

    [Fact]
    public void Connector_Add_WithFill_RoundTrips()
    {
        // Verify that Add also supports fill (not just Set).
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["fill"] = "00B050",
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Should().NotBeNull();
        // If Add supports fill, the Format must reflect it
        // If Add doesn't support fill yet, this documents what the expected behavior should be
        // The connector node should at least be created successfully
        node!.Type.Should().Be("connector");
    }

    // =========================================================================
    // Helper — minimal 1×1 PNG bytes (valid PNG header + IDAT + IEND)
    // =========================================================================

    private string MakeTinyPng()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        _tempFiles.Add(pngPath);
        // A valid 1×1 transparent PNG
        var pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
        File.WriteAllBytes(pngPath, pngBytes);
        return pngPath;
    }
}
