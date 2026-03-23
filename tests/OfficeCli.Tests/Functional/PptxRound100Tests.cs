// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt rounds 100+: Four targeted bugs found via white-box code review.
///
/// Bug A — autoFit="spAutoFit" silently no-ops: the Set switch for "autofit"
///   handles "true"/"normal", "shape", "false"/"none" but has no case for
///   "spAutoFit". The three RemoveAllChildren calls execute, stripping the
///   existing autofit element, but nothing is appended — leaving the body
///   with no autofit child, which PowerPoint interprets as "none".
///   Fix: add "spAutoFit" or "sp" as an alias that appends ShapeAutoFit
///   (or NormalAutoFit, depending on the intended mapping).
///
/// Bug B — Multiple animations: only the first is surfaced in Format["animation"].
///   ReadShapeAnimation walks up from the first ShapeTarget it finds with
///   FirstOrDefault, then immediately assigns node.Format["animation"]. When a
///   shape has two entrance animations, only the first ShapeTarget is visited.
///   Fix: collect all matching ShapeTargets and produce a comma-separated list
///   or an "animation" / "animation2" pair in Format.
///
/// Bug C — Table row height returned as raw EMU integer string instead of a
///   human-readable unit string (e.g. "2cm" / "1.44cm").
///   Root cause: rowNode.Format["height"] = FormatEmu(row.Height.Value) — this
///   is correct; however the bug description indicates the key is sometimes "h"
///   not "height". Verify that the key is "height" and the value is unit-qualified
///   (not a bare 7-digit integer).
///
/// Bug D — Table style returned as raw GUID string (e.g.
///   "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}") instead of the friendly name
///   used during Set (e.g. "medium1"). Get should map the stored GUID back to
///   the friendly name used on input, or at minimum not expose the raw GUID when
///   a known mapping exists.
/// </summary>
public class PptxRound100Tests : IDisposable
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
    // Bug A — autoFit="spAutoFit" silently no-ops, value stays "none"
    //
    // Passing "spAutoFit" to Set should produce a ShapeAutoFit element in the
    // XML (which makes text resize the shape). Currently the three RemoveAll
    // calls execute but nothing is appended, so Get returns "none".
    // =========================================================================

    [Fact]
    public void BugA_Set_AutoFit_SpAutoFit_IsNotSilentlyDropped()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "AutoFit test",
            ["width"] = "4cm", ["height"] = "2cm",
        });

        // Set autoFit to "spAutoFit" — must not silently ignore the value
        var unsupported = handler.Set("/slide[1]/shape[1]", new() { ["autoFit"] = "spAutoFit" });

        unsupported.Should().BeEmpty(
            "autoFit is a supported shape property and spAutoFit is a valid OOXML value");

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("autoFit");
        node.Format["autoFit"].Should().NotBe("none",
            "setting autoFit=spAutoFit must not silently leave the value as 'none'");
    }

    [Fact]
    public void BugA_Set_AutoFit_SpAutoFit_RoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "AutoFit roundtrip",
            ["width"] = "4cm", ["height"] = "2cm",
            ["autoFit"] = "none",
        });

        var before = handler.Get("/slide[1]/shape[1]");
        before!.Format["autoFit"].Should().Be("none", "initial autoFit should be none");

        handler.Set("/slide[1]/shape[1]", new() { ["autoFit"] = "spAutoFit" });

        var after = handler.Get("/slide[1]/shape[1]");
        after!.Format["autoFit"].Should().NotBe("none",
            "after Set autoFit=spAutoFit, the value must change from 'none'");
        // The value should be either "shape" (ShapeAutoFit) or a documented alias
        after.Format["autoFit"].ToString()!.ToLowerInvariant().Should().BeOneOf(
            "shape", "sp", "spautofit", "normal",
            "the autoFit value must reflect ShapeAutoFit or NormalAutoFit, not 'none'");
    }

    [Fact]
    public void BugA_Set_AutoFit_SpAutoFit_Persists()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "AutoFit persist",
                ["width"] = "4cm", ["height"] = "2cm",
            });
            h.Set("/slide[1]/shape[1]", new() { ["autoFit"] = "spAutoFit" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("autoFit");
        node.Format["autoFit"].Should().NotBe("none",
            "autoFit=spAutoFit must survive save/reopen");
    }

    // =========================================================================
    // Bug B — Multiple animations: only the first animation is surfaced
    //
    // When a shape has two animations, ReadShapeAnimation uses FirstOrDefault
    // on ShapeTarget and stops at the first match. The second animation is
    // stored in the timing tree but never exposed in Format.
    // =========================================================================

    [Fact]
    public void BugB_MultipleAnimations_SecondAnimationIsReturned()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Two animations",
            ["width"] = "4cm", ["height"] = "2cm",
        });

        // Add two distinct animations to the same shape
        handler.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "appear",
            ["class"] = "entrance",
            ["duration"] = "500",
            ["trigger"] = "onclick",
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "fade",
            ["class"] = "exit",
            ["duration"] = "800",
            ["trigger"] = "onclick",
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();

        // The shape must expose both animations in some form.
        // Either "animation" is a comma-separated list, or "animation2" also exists.
        var hasSecond = node!.Format.ContainsKey("animation2")
            || (node.Format.ContainsKey("animation")
                && node.Format["animation"]?.ToString()?.Contains(",") == true);

        hasSecond.Should().BeTrue(
            "a shape with two animations must surface both in Format — " +
            "currently only the first ShapeTarget is read and the second is silently dropped");
    }

    [Fact]
    public void BugB_MultipleAnimations_BothEffectNamesArePresent()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Both effects",
            ["width"] = "4cm", ["height"] = "2cm",
        });

        handler.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "appear", ["class"] = "entrance", ["duration"] = "500", ["trigger"] = "onclick",
        });
        handler.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "zoom", ["class"] = "entrance", ["duration"] = "1000", ["trigger"] = "onclick",
        });

        var node = handler.Get("/slide[1]/shape[1]");

        // Collect the full animation output across all animation-related keys
        var allAnimText = string.Join(",",
            node!.Format
                .Where(kv => kv.Key.StartsWith("animation", StringComparison.OrdinalIgnoreCase))
                .Select(kv => kv.Value?.ToString() ?? ""));

        allAnimText.Should().Contain("appear",
            "the first animation effect 'appear' must appear in Format");
        allAnimText.Should().Contain("zoom",
            "the second animation effect 'zoom' must also appear in Format; " +
            "currently ReadShapeAnimation stops after the first ShapeTarget");
    }

    // =========================================================================
    // Bug C — Table row height key is "h" (raw EMU integer) instead of
    //          "height" with unit-qualified string (e.g. "1.44cm")
    //
    // The NodeBuilder writes rowNode.Format["height"] = FormatEmu(row.Height.Value).
    // The bug description says the key is sometimes "h" and the value is a raw
    // 7-digit integer. Confirm the canonical key is "height" and the value is
    // formatted with a unit suffix.
    // =========================================================================

    [Fact]
    public void BugC_TableRowHeight_KeyIsHeight_NotH()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2",
            ["rowHeight"] = "1cm",
            ["width"] = "6cm",
        });

        var node = handler.Get("/slide[1]/table[1]/tr[1]");
        node.Should().NotBeNull();

        node!.Format.Should().NotContainKey("h",
            "the row height key must be 'height', not 'h'");
        node.Format.Should().ContainKey("height",
            "row height must be stored under the canonical key 'height'");
    }

    [Fact]
    public void BugC_TableRowHeight_ValueIsUnitQualified_NotRawEmu()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "2",
            ["rowHeight"] = "2cm",
            ["width"] = "8cm",
        });

        var node = handler.Get("/slide[1]/table[1]/tr[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("height");

        var heightVal = node.Format["height"]?.ToString() ?? "";
        heightVal.Should().NotBeEmpty("row height must be non-empty");

        // A raw EMU value for 2 cm is 720000 — a 6+ digit bare integer.
        // The formatted value should have a unit suffix like "cm", "in", or "pt".
        long rawEmu;
        var isRawEmu = long.TryParse(heightVal, out rawEmu) && rawEmu > 9999;
        isRawEmu.Should().BeFalse(
            $"row height '{heightVal}' must be unit-qualified (e.g. '2cm'), not a raw EMU integer");

        // Must contain a unit character
        heightVal.Should().MatchRegex(@"\d+(\.\d+)?(cm|in|pt|mm)",
            "row height must be formatted with a unit suffix (cm, in, pt, or mm)");
    }

    [Fact]
    public void BugC_TableRowHeight_ApproximatelyRoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "3",
            ["rowHeight"] = "1.5cm",
            ["width"] = "9cm",
        });

        var node = handler.Get("/slide[1]/table[1]/tr[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("height");

        var heightStr = node.Format["height"]?.ToString() ?? "";
        // Accept "1.5cm", "1.49cm", "1.51cm" — within 2 % of 1.5 cm
        heightStr.Should().MatchRegex(@"1\.[45]\d*cm",
            "1.5cm row height should round-trip to approximately 1.5cm");
    }

    // =========================================================================
    // Bug D — Table style is returned as raw GUID instead of friendly name
    //
    // When Add/Set uses a friendly name like "medium1", it is translated to a
    // GUID internally. Get then reads back the raw GUID from the XML and stores
    // it in Format["tableStyleId"]. The user never sees "medium1" again.
    // Fix: Get should reverse-map the GUID back to the friendly name when a
    // known mapping exists, or at minimum expose the value under "tableStyle"
    // with the friendly name (not "tableStyleId" with the GUID).
    // =========================================================================

    [Fact]
    public void BugD_TableStyle_Get_ReturnsStyleName_NotGuid()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2",
            ["style"] = "medium1",
            ["width"] = "6cm",
        });

        var node = handler.Get("/slide[1]/table[1]");
        node.Should().NotBeNull();

        // The table node should expose the friendly style name, not a raw GUID
        var styleValue = node!.Format.ContainsKey("tableStyle")
            ? node.Format["tableStyle"]?.ToString()
            : node.Format.ContainsKey("tableStyleId")
                ? node.Format["tableStyleId"]?.ToString()
                : null;

        styleValue.Should().NotBeNullOrEmpty(
            "table style must be exposed in Format under 'tableStyle' or 'tableStyleId'");
        styleValue!.Should().NotStartWith("{",
            "the style value should be the friendly name 'medium1', not a raw GUID like '{073A0DAA-...}'");
        styleValue.Should().Be("medium1",
            "Get should reverse-map the stored GUID back to the friendly name used during Add");
    }

    [Fact]
    public void BugD_TableStyle_SetAndGet_FriendlyNameRoundTrips()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3",
            ["width"] = "8cm",
        });

        handler.Set("/slide[1]/table[1]", new() { ["style"] = "light2" });

        var node = handler.Get("/slide[1]/table[1]");
        node.Should().NotBeNull();

        var styleValue = node!.Format.ContainsKey("tableStyle")
            ? node.Format["tableStyle"]?.ToString()
            : node.Format.ContainsKey("tableStyleId")
                ? node.Format["tableStyleId"]?.ToString()
                : null;

        styleValue.Should().NotBeNullOrEmpty("table style must be set and readable");
        styleValue!.Should().NotStartWith("{",
            "style should be returned as friendly name 'light2', not as GUID");
        styleValue.Should().Be("light2",
            "Set('light2') → Get should return 'light2', not a GUID");
    }

    [Fact]
    public void BugD_TableStyle_KnownStyles_AreReverseMapped()
    {
        // Verify several well-known style names all round-trip correctly.
        // Each style is tested in its own file to avoid slide-accumulation
        // across handler re-opens on the same path.
        var knownStyles = new[] { "medium1", "light1", "dark1", "none" };

        foreach (var styleName in knownStyles)
        {
            var stylePath = CreateTemp();
            BlankDocCreator.Create(stylePath);

            using var handler = new PowerPointHandler(stylePath, editable: true);
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "table", null, new()
            {
                ["rows"] = "1",
                ["cols"] = "2",
                ["style"] = styleName,
                ["width"] = "6cm",
            });

            var node = handler.Get("/slide[1]/table[1]");
            node.Should().NotBeNull($"table should exist for style '{styleName}'");

            var styleValue = node!.Format.ContainsKey("tableStyle")
                ? node.Format["tableStyle"]?.ToString()
                : node.Format.ContainsKey("tableStyleId")
                    ? node.Format["tableStyleId"]?.ToString()
                    : null;

            styleValue.Should().NotStartWith("{",
                $"style '{styleName}' must not be returned as a raw GUID");
            styleValue.Should().Be(styleName,
                $"style '{styleName}' must round-trip through Add → Get");
        }
    }
}
