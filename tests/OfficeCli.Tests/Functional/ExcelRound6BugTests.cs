// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for Round 6 bugs reported by Agent B.
///
/// CONFIRMED BUGS (all tests here are expected to FAIL until fixed):
///
///   Bug 1 (HIGH) — remove /Sheet1/comment[1], /Sheet1/validation[1], /Sheet1/cf[1] all fail
///            Error: "Cell comment[1] not found" / "Cell validation[1] not found" / "Cell cf[1] not found"
///            Root cause: ExcelHandler.Remove.cs has no dispatch blocks for comment[N],
///            validation[N], or cf[N] path segments. After the shape[N]/picture[N] blocks,
///            the remaining segment falls through to FindCell(), which treats the segment as
///            a raw cell reference — finds nothing — and throws "Cell X not found".
///            Fix: Add three dispatch blocks (before the cell fallback) for:
///              - comment[N]: locate WorksheetCommentsPart.Comments.CommentList children by
///                1-based index and call .Remove()
///              - validation[N]: locate DataValidations children by 1-based index and call .Remove()
///              - cf[N]: locate ConditionalFormatting elements by 1-based index and call .Remove()
///
///   Bug 2 (HIGH) — remove /namedrange[1] fails with "Sheet not found: namedrange[1]"
///            Root cause: NormalizeExcelPath does not rewrite the path, so Remove splits on '/'
///            and treats "namedrange[1]" as the sheet name, then calls SheetNotFoundException.
///            Fix: In the Remove method, check segments[0] against a namedrange[N] regex before
///            calling FindWorksheet; if it matches, remove the Nth DefinedName from the workbook.
///
///   Bug 3 (MEDIUM) — 3-color gradient middle color silently dropped
///            Input: "fill=gradient;FF0000;FFFF00;0000FF;45" (3 colors + angle)
///            Result: only 2 gradient stops stored; middle color FFFF00 is lost.
///            Root cause: ExcelStyleManager.cs line 130 converts the semicolon format to dash
///            format via: fillColor.TrimStart("gradient;".ToCharArray()).Replace(';', '-')
///            This converts "gradient;FF0000;FFFF00;0000FF;45" to "FF0000-FFFF00-0000FF-45".
///            In GetOrCreateGradientFill, the angle-detection heuristic checks
///            colors.Last().Length &lt;= 3 (line 504), and "45" has length 2, so it is correctly
///            stripped as the angle — leaving colors = ["FF0000", "FFFF00", "0000FF"], which
///            is 3 colors. BUT the Get readback in ExcelHandler.Helpers.cs (line 373) only reads
///            stops[0] and stops[^1] and formats them as "gradient;C1;C2;deg" — it discards the
///            middle stop entirely. So the readback is always 2-color even when 3 stops were stored.
///            Fix: The readback loop must include all intermediate stops in the format string:
///            "gradient;C1;C2;...;Cn;deg".
/// </summary>
public class ExcelRound6BugTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelRound6BugTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 1a — remove /Sheet1/comment[1] fails
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Remove must accept a comment[N] path without throwing.
    /// Currently falls through to FindCell("comment[1]") → "Cell comment[1] not found".
    /// </summary>
    [Fact]
    public void RemoveComment_DoesNotThrow()
    {
        // Arrange
        _handler.Add("/Sheet1", "comment", null, new()
        {
            ["ref"] = "B2",
            ["text"] = "Review this"
        });

        var before = _handler.Get("/Sheet1/comment[1]");
        before.Should().NotBeNull("comment must exist before removal");

        // Act — currently throws "Cell comment[1] not found"
        var act = () => _handler.Remove("/Sheet1/comment[1]");
        act.Should().NotThrow("Remove must accept a comment[N] path");
    }

    /// <summary>
    /// After removing a comment, Get("/Sheet1/comment[1]") must return null.
    /// </summary>
    [Fact]
    public void RemoveComment_CommentIsGoneAfterRemoval()
    {
        _handler.Add("/Sheet1", "comment", null, new()
        {
            ["ref"] = "C3",
            ["text"] = "Temporary note"
        });

        _handler.Remove("/Sheet1/comment[1]");

        var after = _handler.Get("/Sheet1/comment[1]");
        after.Should().BeNull("comment should not exist after removal");
    }

    /// <summary>
    /// When two comments exist and only the first is removed, the second must survive.
    /// </summary>
    [Fact]
    public void RemoveComment_SecondCommentSurvivesRemovalOfFirst()
    {
        _handler.Add("/Sheet1", "comment", null, new()
        {
            ["ref"] = "A1",
            ["text"] = "First"
        });
        _handler.Add("/Sheet1", "comment", null, new()
        {
            ["ref"] = "A2",
            ["text"] = "Second"
        });

        _handler.Remove("/Sheet1/comment[1]");

        // The second comment must survive and become comment[1]
        var remaining = _handler.Get("/Sheet1/comment[1]");
        remaining.Should().NotBeNull("second comment must survive removal of first");
        remaining!.Text.Should().Be("Second", "the remaining comment must be the second one");
    }

    /// <summary>
    /// Comment removal must persist across file reopen.
    /// </summary>
    [Fact]
    public void RemoveComment_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "comment", null, new()
        {
            ["ref"] = "D4",
            ["text"] = "Ephemeral"
        });

        _handler.Remove("/Sheet1/comment[1]");
        Reopen();

        var after = _handler.Get("/Sheet1/comment[1]");
        after.Should().BeNull("comment removal must persist after reopen");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 1b — remove /Sheet1/validation[1] fails
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Remove must accept a validation[N] path without throwing.
    /// Currently falls through to FindCell("validation[1]") → "Cell validation[1] not found".
    /// </summary>
    [Fact]
    public void RemoveValidation_DoesNotThrow()
    {
        // Arrange
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "B2",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "100"
        });

        var before = _handler.Get("/Sheet1/validation[1]");
        before.Should().NotBeNull("validation must exist before removal");

        // Act — currently throws "Cell validation[1] not found"
        var act = () => _handler.Remove("/Sheet1/validation[1]");
        act.Should().NotThrow("Remove must accept a validation[N] path");
    }

    /// <summary>
    /// After removing a validation, Get("/Sheet1/validation[1]") must return null.
    /// </summary>
    [Fact]
    public void RemoveValidation_ValidationIsGoneAfterRemoval()
    {
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "C3",
            ["type"] = "list",
            ["formula1"] = "\"Yes,No\""
        });

        _handler.Remove("/Sheet1/validation[1]");

        var after = _handler.Get("/Sheet1/validation[1]");
        after.Should().BeNull("validation should not exist after removal");
    }

    /// <summary>
    /// When two validations exist and only the first is removed, the second must survive.
    /// </summary>
    [Fact]
    public void RemoveValidation_SecondValidationSurvivesRemovalOfFirst()
    {
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1",
            ["type"] = "list",
            ["formula1"] = "\"Yes,No\""
        });
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A2",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "10"
        });

        _handler.Remove("/Sheet1/validation[1]");

        var remaining = _handler.Get("/Sheet1/validation[1]");
        remaining.Should().NotBeNull("second validation must survive removal of first");
    }

    /// <summary>
    /// Validation removal must persist across file reopen.
    /// </summary>
    [Fact]
    public void RemoveValidation_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "E5",
            ["type"] = "list",
            ["formula1"] = "\"A,B,C\""
        });

        _handler.Remove("/Sheet1/validation[1]");
        Reopen();

        var after = _handler.Get("/Sheet1/validation[1]");
        after.Should().BeNull("validation removal must persist after reopen");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 1c — remove /Sheet1/cf[1] fails
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Remove must accept a cf[N] path without throwing.
    /// Currently falls through to FindCell("cf[1]") → "Cell cf[1] not found".
    /// </summary>
    [Fact]
    public void RemoveCf_DoesNotThrow()
    {
        // Arrange
        _handler.Add("/Sheet1", "databar", null, new()
        {
            ["sqref"] = "A1:A10"
        });

        var before = _handler.Get("/Sheet1/cf[1]");
        before.Should().NotBeNull("cf must exist before removal");

        // Act — currently throws "Cell cf[1] not found"
        var act = () => _handler.Remove("/Sheet1/cf[1]");
        act.Should().NotThrow("Remove must accept a cf[N] path");
    }

    /// <summary>
    /// After removing a cf rule, Get("/Sheet1/cf[1]") must return null.
    /// </summary>
    [Fact]
    public void RemoveCf_CfIsGoneAfterRemoval()
    {
        _handler.Add("/Sheet1", "colorscale", null, new()
        {
            ["sqref"] = "B1:B10"
        });

        _handler.Remove("/Sheet1/cf[1]");

        var after = _handler.Get("/Sheet1/cf[1]");
        after.Should().BeNull("cf rule should not exist after removal");
    }

    /// <summary>
    /// When two CF rules exist and only the first is removed, the second must survive.
    /// </summary>
    [Fact]
    public void RemoveCf_SecondCfSurvivesRemovalOfFirst()
    {
        _handler.Add("/Sheet1", "databar", null, new()
        {
            ["sqref"] = "A1:A10"
        });
        _handler.Add("/Sheet1", "colorscale", null, new()
        {
            ["sqref"] = "B1:B10"
        });

        _handler.Remove("/Sheet1/cf[1]");

        var remaining = _handler.Get("/Sheet1/cf[1]");
        remaining.Should().NotBeNull("second cf rule must survive removal of first");
    }

    /// <summary>
    /// CF rule removal must persist across file reopen.
    /// </summary>
    [Fact]
    public void RemoveCf_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "databar", null, new()
        {
            ["sqref"] = "C1:C5"
        });

        _handler.Remove("/Sheet1/cf[1]");
        Reopen();

        var after = _handler.Get("/Sheet1/cf[1]");
        after.Should().BeNull("cf removal must persist after reopen");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 2 — remove /namedrange[1] fails with "Sheet not found: namedrange[1]"
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Remove must accept a /namedrange[N] path without throwing.
    /// Currently the path is split on '/' and "namedrange[1]" is treated as a sheet name,
    /// causing: InvalidOperationException: Sheet not found: namedrange[1]
    /// </summary>
    [Fact]
    public void RemoveNamedRange_DoesNotThrow()
    {
        // Arrange
        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "SalesData",
            ["ref"] = "Sheet1!$A$1:$A$10"
        });

        var before = _handler.Get("/namedrange[1]");
        before.Should().NotBeNull("namedrange must exist before removal");

        // Act — currently throws "Sheet not found: namedrange[1]"
        var act = () => _handler.Remove("/namedrange[1]");
        act.Should().NotThrow("Remove must accept a /namedrange[N] path");
    }

    /// <summary>
    /// After removing a named range, Get("/namedrange[1]") must return null.
    /// </summary>
    [Fact]
    public void RemoveNamedRange_NamedRangeIsGoneAfterRemoval()
    {
        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "TempRange",
            ["ref"] = "Sheet1!$B$1:$B$5"
        });

        _handler.Remove("/namedrange[1]");

        var after = _handler.Get("/namedrange[1]");
        after.Should().BeNull("named range should not exist after removal");
    }

    /// <summary>
    /// When two named ranges exist and only the first is removed, the second must survive
    /// and become namedrange[1].
    /// </summary>
    [Fact]
    public void RemoveNamedRange_SecondNamedRangeSurvivesRemovalOfFirst()
    {
        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "FirstRange",
            ["ref"] = "Sheet1!$A$1:$A$5"
        });
        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "SecondRange",
            ["ref"] = "Sheet1!$B$1:$B$5"
        });

        _handler.Remove("/namedrange[1]");

        var remaining = _handler.Get("/namedrange[1]");
        remaining.Should().NotBeNull("second named range must survive removal of first");
        remaining!.Format["name"].Should().Be("SecondRange",
            "the surviving named range must be the second one (SecondRange)");
    }

    /// <summary>
    /// Named range removal must persist across file reopen.
    /// </summary>
    [Fact]
    public void RemoveNamedRange_PersistsAfterReopen()
    {
        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "EphemeralRange",
            ["ref"] = "Sheet1!$C$1"
        });

        _handler.Remove("/namedrange[1]");
        Reopen();

        var after = _handler.Get("/namedrange[1]");
        after.Should().BeNull("named range removal must persist after reopen");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 3 — 3-color gradient middle color silently dropped in readback
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Setting a 3-color gradient via semicolon format must store all 3 colors.
    /// The Get readback currently returns only start and end colors (stops[0] and stops[^1]),
    /// silently discarding the middle stop.
    /// </summary>
    [Fact]
    public void ThreeColorGradient_MiddleColorIsPreservedInReadback()
    {
        // Arrange: set a 3-color gradient: red → yellow → blue at 45°
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1",
            ["value"] = "test",
            ["fill"] = "gradient;FF0000;FFFF00;0000FF;45"
        });

        // Act
        var node = _handler.Get("/Sheet1/A1");

        // Assert: the returned fill format must include all 3 colors
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("fill");
        var fill = node.Format["fill"]?.ToString() ?? "";

        // The fill value must contain all three hex colors
        fill.Should().Contain("FF0000", "the first (red) color must appear in the fill readback");
        fill.Should().Contain("FFFF00", "the middle (yellow) color must NOT be silently dropped");
        fill.Should().Contain("0000FF", "the last (blue) color must appear in the fill readback");
    }

    /// <summary>
    /// Setting a 3-color gradient via Set (not Add) must also preserve the middle color.
    /// </summary>
    [Fact]
    public void ThreeColorGradient_MiddleColorPreservedAfterSet()
    {
        // Arrange: add a plain cell first
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "B2",
            ["value"] = "data"
        });

        // Act: apply 3-color gradient via Set
        _handler.Set("/Sheet1/B2", new()
        {
            ["fill"] = "gradient;00FF00;FFFFFF;FF0000;90"
        });

        var node = _handler.Get("/Sheet1/B2");
        node.Should().NotBeNull();
        var fill = node!.Format["fill"]?.ToString() ?? "";

        fill.Should().Contain("00FF00", "the first (green) color must appear in the fill readback");
        fill.Should().Contain("FFFFFF", "the middle (white) color must NOT be silently dropped");
        fill.Should().Contain("FF0000", "the last (red) color must appear in the fill readback");
    }

    /// <summary>
    /// A 3-color gradient must persist all 3 colors across file reopen.
    /// </summary>
    [Fact]
    public void ThreeColorGradient_MiddleColorPersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "C3",
            ["value"] = "persist",
            ["fill"] = "gradient;FF0000;FFFF00;0000FF;0"
        });

        Reopen();

        var node = _handler.Get("/Sheet1/C3");
        node.Should().NotBeNull();
        var fill = node!.Format["fill"]?.ToString() ?? "";

        fill.Should().Contain("FF0000", "red must persist after reopen");
        fill.Should().Contain("FFFF00", "yellow (middle) must persist after reopen");
        fill.Should().Contain("0000FF", "blue must persist after reopen");
    }

    /// <summary>
    /// The number of color stops reported by Get must equal the number of colors passed in.
    /// A 3-color gradient input must produce a 3-stop readback, not a 2-stop one.
    /// </summary>
    [Fact]
    public void ThreeColorGradient_ReadbackHasThreeSegments()
    {
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "D4",
            ["value"] = "count",
            ["fill"] = "gradient;FF0000;FFFF00;0000FF;45"
        });

        var node = _handler.Get("/Sheet1/D4");
        var fill = node!.Format["fill"]?.ToString() ?? "";

        // Expected format: "gradient;#FF0000;#FFFF00;#0000FF;45"
        // Split by ';' → ["gradient", "#FF0000", "#FFFF00", "#0000FF", "45"] = 5 parts
        // The color segments are parts 1..(n-2), leaving part[0]="gradient" and part[n-1]=angle.
        var parts = fill.Split(';');
        // There should be at least 5 parts: "gradient", c1, c2, c3, angle
        parts.Length.Should().BeGreaterThanOrEqualTo(5,
            "a 3-color gradient must produce at least 5 semicolon-delimited parts: gradient;C1;C2;C3;angle");
    }
}
