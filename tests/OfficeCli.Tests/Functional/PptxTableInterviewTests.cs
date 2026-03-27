// PptxTableInterviewTests.cs — Tests for table interview-discovered bugs:
//   Bug 1: Add row does not update GraphicFrame container height (cy)
//   Bug 2: underline and strikethrough not returned by Get on table cells
//   Bug 3: table style accepts invalid values without validation

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

public class PptxTableInterviewTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTempPptx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 1: Add row should update GraphicFrame container height (cy)
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public void AddRow_UpdatesContainerHeight()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });

        // Get initial height
        var node1 = handler.Get("/slide[1]/table[1]");
        var heightBefore = node1.Format["height"]?.ToString();
        heightBefore.Should().NotBeNullOrEmpty();

        // Add a row
        handler.Add("/slide[1]/table[1]", "row", null, new());

        // Height should have increased
        var node2 = handler.Get("/slide[1]/table[1]");
        var heightAfter = node2.Format["height"]?.ToString();
        heightAfter.Should().NotBeNullOrEmpty();

        // Parse EMU values for comparison
        var cyBefore = EmuConverter.ParseEmu(heightBefore!);
        var cyAfter = EmuConverter.ParseEmu(heightAfter!);
        cyAfter.Should().BeGreaterThan(cyBefore, "adding a row should increase the table container height");
    }

    [Fact]
    public void AddRow_ContainerHeightMatchesSumOfRowHeights()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2", ["rowHeight"] = "1cm" });

        // Add two more rows
        handler.Add("/slide[1]/table[1]", "row", null, new());
        handler.Add("/slide[1]/table[1]", "row", null, new());

        // Now 4 rows, each 1cm = 360000 EMU
        var node = handler.Get("/slide[1]/table[1]");
        var heightStr = node.Format["height"]?.ToString();
        var cy = EmuConverter.ParseEmu(heightStr!);

        // 4 rows * 360000 EMU = 1440000 EMU
        cy.Should().Be(4 * 360000, "container height should equal sum of all row heights");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 2: underline and strikethrough should be returned by Get on cells
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public void SetUnderline_IsReturnedByGet()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1", ["data"] = "Hello"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["underline"] = "single" });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("underline");
        node.Format["underline"].Should().Be("single");
    }

    [Fact]
    public void SetStrikethrough_IsReturnedByGet()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1", ["data"] = "Hello"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["strike"] = "single" });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("strike");
        node.Format["strike"].Should().Be("single");
    }

    [Fact]
    public void SetDoubleUnderline_IsReturnedByGet()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1", ["data"] = "Hello"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["underline"] = "double" });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("underline");
        node.Format["underline"].Should().Be("double");
    }

    [Fact]
    public void SetDoubleStrike_IsReturnedByGet()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1", ["data"] = "Hello"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["strike"] = "double" });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("strike");
        node.Format["strike"].Should().Be("double");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 3: table style should reject invalid values
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public void SetTableStyle_ValidName_Succeeds()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        handler.Set("/slide[1]/table[1]", new() { ["style"] = "medium2" });

        var node = handler.Get("/slide[1]/table[1]");
        node.Format["style"].Should().Be("medium2");
    }

    [Fact]
    public void SetTableStyle_InvalidName_ThrowsArgumentException()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        var act = () => handler.Set("/slide[1]/table[1]", new() { ["style"] = "nonexistent" });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*Invalid table style*");
    }

    [Fact]
    public void AddTable_InvalidStyle_ThrowsArgumentException()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        var act = () => handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1", ["style"] = "nonexistent"
        });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*Invalid table style*");
    }

    [Fact]
    public void SetTableStyle_DirectGuid_Succeeds()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        // Direct GUID should still work
        handler.Set("/slide[1]/table[1]", new() { ["style"] = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}" });

        var node = handler.Get("/slide[1]/table[1]");
        node.Format["style"].Should().Be("medium1");
    }
}
