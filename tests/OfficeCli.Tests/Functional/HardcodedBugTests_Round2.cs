// Hardcoded Bug Tests Round 2
//
// Bug 1 (CRITICAL): 3D chart types dead code — column3d/bar3d never matched in switch
// Bug 2 (HIGH): Excel shape Set size didn't support "pt" suffix
// Bug 5 (MEDIUM): table cell alignment duplicate key (alignment removed, keep align only)
// Bug 8 (MEDIUM): Excel shape align swallowed "justify"
// Bug 9 (MEDIUM): paragraph-level align returned raw XML codes instead of friendly names
// Bug 12 (LOW): shape rotation readback format inconsistency (now uses 0.## formatter)

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class HardcodedBugTests_Round2 : IDisposable
{
    private readonly string _pptxPath;
    private PowerPointHandler _pptHandler;

    public HardcodedBugTests_Round2()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_pptxPath);
        _pptHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _pptHandler.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    // ==================== Bug 1: 3D chart types ====================

    [Fact]
    public void Bug1_Column3d_ShouldCreateBar3DChart()
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "3D Chart" });

        var path = _pptHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column3d",
            ["title"] = "3D Sales",
            ["categories"] = "Q1,Q2,Q3",
            ["data"] = "Revenue:100,200,300"
        });

        path.Should().NotBeNullOrEmpty();
        var node = _pptHandler.Get(path);
        node.Should().NotBeNull();
        node.Type.Should().Be("chart");
    }

    [Fact]
    public void Bug1_Bar3d_ShouldCreateBar3DChart()
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "3D Bar" });

        var path = _pptHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar3d",
            ["title"] = "3D Bars",
            ["categories"] = "A,B,C",
            ["data"] = "S1:10,20,30"
        });

        path.Should().NotBeNullOrEmpty();
    }

    // ==================== Bug 2: Excel shape size with pt suffix ====================

    [Fact]
    public void Bug2_ExcelShape_SetSize_WithPtSuffix()
    {
        var xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        try
        {
            BlankDocCreator.Create(xlsxPath);
            using var handler = new ExcelHandler(xlsxPath, editable: true);

            handler.Add("/Sheet1", "shape", null, new()
            {
                ["text"] = "Test", ["x"] = "100", ["y"] = "100",
                ["width"] = "300", ["height"] = "200"
            });

            // Should not throw with "pt" suffix
            handler.Set("/Sheet1/shape[1]", new() { ["size"] = "14pt" });

            var node = handler.Get("/Sheet1/shape[1]");
            node.Format["size"].Should().Be("14pt");
        }
        finally
        {
            if (File.Exists(xlsxPath)) File.Delete(xlsxPath);
        }
    }

    // ==================== Bug 5: table cell alignment key ====================

    [Fact]
    public void Bug5_TableCell_ShouldNotHaveDuplicateAlignmentKey()
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "Table" });
        _pptHandler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "10cm", ["height"] = "5cm"
        });
        _pptHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Centered", ["align"] = "center"
        });

        var tableNode = _pptHandler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("align");
        cellNode.Format["align"].Should().Be("center");
        cellNode.Format.Should().NotContainKey("alignment",
            "canonical key is 'align', should not have duplicate 'alignment'");
    }

    // ==================== Bug 8: Excel shape justify alignment ====================

    [Fact]
    public void Bug8_ExcelShape_JustifyAlignment()
    {
        var xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        try
        {
            BlankDocCreator.Create(xlsxPath);
            using var handler = new ExcelHandler(xlsxPath, editable: true);

            handler.Add("/Sheet1", "shape", null, new()
            {
                ["text"] = "Justify test", ["x"] = "100", ["y"] = "100",
                ["width"] = "300", ["height"] = "200"
            });

            // Should not throw — justify should be accepted
            var act = () => handler.Set("/Sheet1/shape[1]", new() { ["align"] = "justify" });
            act.Should().NotThrow("justify should be a valid alignment value");
        }
        finally
        {
            if (File.Exists(xlsxPath)) File.Delete(xlsxPath);
        }
    }

    // ==================== Bug 9: paragraph align friendly names ====================

    [Fact]
    public void Bug9_ParagraphAlign_ShouldReturnFriendlyNames()
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "Align Test" });
        _pptHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line1",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "5cm"
        });
        _pptHandler.Add("/slide[1]/shape[1]", "paragraph", null, new() { ["text"] = "Line2" });
        _pptHandler.Add("/slide[1]/shape[1]", "paragraph", null, new() { ["text"] = "Line3" });

        _pptHandler.Set("/slide[1]/shape[1]/paragraph[2]", new() { ["align"] = "center" });
        _pptHandler.Set("/slide[1]/shape[1]/paragraph[3]", new() { ["align"] = "right" });

        var node = _pptHandler.Get("/slide[1]/shape[1]", depth: 2);
        node.Children.Count.Should().BeGreaterOrEqualTo(3);

        var para2 = node.Children[1];
        para2.Format.Should().ContainKey("align");
        para2.Format["align"].Should().Be("center",
            "paragraph align should return 'center' not 'ctr'");

        var para3 = node.Children[2];
        para3.Format.Should().ContainKey("align");
        para3.Format["align"].Should().Be("right",
            "paragraph align should return 'right' not 'r'");
    }

    // ==================== Bug 12: rotation format consistency ====================

    [Fact]
    public void Bug12_ShapeRotation_ShouldBeCleanNumber()
    {
        _pptHandler.Add("/", "slide", null, new() { ["title"] = "Rotation" });
        _pptHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Rotated",
            ["x"] = "5cm", ["y"] = "5cm", ["width"] = "5cm", ["height"] = "3cm"
        });
        _pptHandler.Set("/slide[1]/shape[1]", new() { ["rotation"] = "45" });

        var node = _pptHandler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("rotation");
        node.Format["rotation"].Should().Be("45",
            "rotation should be clean '45', not '45.0000000000001'");
    }
}
