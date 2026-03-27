// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Black-box tests for chart ManualLayout (plotArea/title/legend x/y/w/h).
/// Tests cover Excel, PPTX, multi-element set, ViewAsHtml, case-insensitivity,
/// invalid values, layout preservation after content Set, and missing legend.
/// </summary>
public class BtChartLayoutTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private ExcelHandler _excel;
    private PowerPointHandler _pptx;

    public BtChartLayoutTests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _excel.Dispose();
        _pptx.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void ReopenExcel() { _excel.Dispose(); _excel = new ExcelHandler(_xlsxPath, editable: true); }
    private void ReopenPptx() { _pptx.Dispose(); _pptx = new PowerPointHandler(_pptxPath, editable: true); }

    // ==================== Scenario 1: Excel chart supports layout properties ====================

    [Fact]
    public void Excel_Chart_Set_PlotAreaLayout_GetReadsBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Sales", ["data"] = "S1:10,20,30"
        });

        _excel.Set("/Sheet1/chart[1]", new()
        {
            ["plotArea.x"] = "0.1",
            ["plotArea.y"] = "0.1",
            ["plotArea.w"] = "0.7",
            ["plotArea.h"] = "0.7"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("plotArea.x");
        node.Format["plotArea.x"].ToString().Should().Be("0.1");
        node.Format.Should().ContainKey("plotArea.y");
        node.Format["plotArea.y"].ToString().Should().Be("0.1");
        node.Format.Should().ContainKey("plotArea.w");
        node.Format["plotArea.w"].ToString().Should().Be("0.7");
        node.Format.Should().ContainKey("plotArea.h");
        node.Format["plotArea.h"].ToString().Should().Be("0.7");
    }

    [Fact]
    public void Excel_Chart_Set_PlotAreaLayout_PersistsAfterReopen()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "Revenue", ["data"] = "S1:5,10,15"
        });
        _excel.Set("/Sheet1/chart[1]", new()
        {
            ["plotArea.x"] = "0.15", ["plotArea.y"] = "0.2",
            ["plotArea.w"] = "0.65", ["plotArea.h"] = "0.6"
        });

        ReopenExcel();

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Should().NotBeNull();
        node!.Format["plotArea.x"].ToString().Should().Be("0.15");
        node.Format["plotArea.y"].ToString().Should().Be("0.2");
        node.Format["plotArea.w"].ToString().Should().Be("0.65");
        node.Format["plotArea.h"].ToString().Should().Be("0.6");
    }

    [Fact]
    public void Excel_Chart_Set_TitleLayout_GetReadsBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Trend", ["data"] = "S1:1,2,3"
        });

        _excel.Set("/Sheet1/chart[1]", new()
        {
            ["title.x"] = "0.3", ["title.y"] = "0.05",
            ["title.w"] = "0.4", ["title.h"] = "0.08"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("title.x");
        node.Format["title.x"].ToString().Should().Be("0.3");
        node.Format.Should().ContainKey("title.y");
        node.Format["title.y"].ToString().Should().Be("0.05");
    }

    [Fact]
    public void Excel_Chart_Set_LegendLayout_GetReadsBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Data", ["legend"] = "true", ["data"] = "S1:1,2,3"
        });

        _excel.Set("/Sheet1/chart[1]", new()
        {
            ["legend.x"] = "0.8", ["legend.y"] = "0.3",
            ["legend.w"] = "0.15", ["legend.h"] = "0.25"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("legend.x");
        node.Format["legend.x"].ToString().Should().Be("0.8");
        node.Format.Should().ContainKey("legend.y");
        node.Format["legend.y"].ToString().Should().Be("0.3");
    }

    // ==================== Scenario 2: Multi-element layout set together ====================

    [Fact]
    public void Pptx_Chart_Set_MultiElement_PlotArea_Title_Legend_Together()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Layout Test" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Multi", ["legend"] = "true", ["data"] = "S1:1,2,3"
        });

        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotArea.x"] = "0.1",
            ["plotArea.y"] = "0.15",
            ["plotArea.w"] = "0.65",
            ["plotArea.h"] = "0.7",
            ["title.x"] = "0.2",
            ["title.y"] = "0.02",
            ["title.w"] = "0.6",
            ["title.h"] = "0.1",
            ["legend.x"] = "0.78",
            ["legend.y"] = "0.4",
            ["legend.w"] = "0.18",
            ["legend.h"] = "0.2"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        node!.Format["plotArea.x"].ToString().Should().Be("0.1");
        node.Format["plotArea.y"].ToString().Should().Be("0.15");
        node.Format["plotArea.w"].ToString().Should().Be("0.65");
        node.Format["plotArea.h"].ToString().Should().Be("0.7");
        node.Format["title.x"].ToString().Should().Be("0.2");
        node.Format["title.y"].ToString().Should().Be("0.02");
        node.Format["legend.x"].ToString().Should().Be("0.78");
        node.Format["legend.y"].ToString().Should().Be("0.4");
    }

    // ==================== Scenario 3: ViewAsHtml does not crash after layout Set ====================

    [Fact]
    public void Pptx_Chart_ViewAsHtml_AfterLayoutSet_DoesNotThrow()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "HTML Test" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "Chart", ["data"] = "S1:5,10,15"
        });

        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotArea.x"] = "0.1",
            ["plotArea.y"] = "0.1",
            ["plotArea.w"] = "0.75",
            ["plotArea.h"] = "0.75"
        });

        var act = () => _pptx.ViewAsHtml();
        act.Should().NotThrow();
        var html = _pptx.ViewAsHtml();
        html.Should().NotBeNullOrEmpty();
    }

    // ==================== Scenario 4: Case-insensitive keys ====================

    [Fact]
    public void Pptx_Chart_Set_Layout_CaseMixed_IsAccepted()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Case Test" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "CaseChart", ["legend"] = "true", ["data"] = "S1:1,2,3"
        });

        // Use uppercase and mixed case keys — should not throw, should apply
        var act = () => _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["PlotArea.X"] = "0.12",
            ["PLOTAREA.Y"] = "0.12",
            ["Title.Y"] = "0.03",
            ["LEGEND.X"] = "0.79"
        });
        act.Should().NotThrow();

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        // At least one of the case-insensitive keys should have been applied
        node!.Format.Should().ContainKey("plotArea.x");
        node.Format["plotArea.x"].ToString().Should().Be("0.12");
    }

    // ==================== Scenario 5: Invalid values are tolerated gracefully ====================

    [Fact]
    public void Pptx_Chart_Set_Layout_NonNumericValue_DoesNotThrow()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Invalid Test" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "InvalidChart", ["data"] = "S1:1,2,3"
        });

        var act = () => _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotArea.x"] = "notanumber",
            ["plotArea.y"] = "",
            ["plotArea.w"] = "abc"
        });
        act.Should().NotThrow();

        // After invalid Set, layout keys should not be written
        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        node!.Format.Should().NotContainKey("plotArea.x");
    }

    // ==================== Scenario 6: Layout preserved after content Set ====================

    [Fact]
    public void Pptx_Chart_Set_Layout_ThenSetTitle_LayoutPreserved()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Preserve Test" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "OriginalTitle", ["data"] = "S1:3,6,9"
        });

        // Set layout first
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotArea.x"] = "0.2",
            ["plotArea.y"] = "0.2",
            ["plotArea.w"] = "0.6",
            ["plotArea.h"] = "0.6"
        });

        // Now change title content
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["title"] = "UpdatedTitle"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        // plotArea layout should still be present
        node!.Format.Should().ContainKey("plotArea.x");
        node.Format["plotArea.x"].ToString().Should().Be("0.2");
        node.Format.Should().ContainKey("plotArea.w");
        node.Format["plotArea.w"].ToString().Should().Be("0.6");
    }

    [Fact]
    public void Excel_Chart_Set_Layout_ThenSetTitle_LayoutPreserved()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Before", ["data"] = "S1:1,2,3"
        });

        _excel.Set("/Sheet1/chart[1]", new()
        {
            ["plotArea.x"] = "0.12", ["plotArea.w"] = "0.72"
        });
        _excel.Set("/Sheet1/chart[1]", new()
        {
            ["title"] = "After"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("plotArea.x");
        node.Format["plotArea.x"].ToString().Should().Be("0.12");
    }

    // ==================== Scenario 7: Set legend.x when chart has no legend ====================

    [Fact]
    public void Pptx_Chart_Set_LegendLayout_NoLegend_DoesNotThrow()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "No Legend Test" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "NoLegendChart", ["legend"] = "false", ["data"] = "S1:1,2,3"
        });

        // Setting legend layout on a chart with no legend element
        var act = () => _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["legend.x"] = "0.8",
            ["legend.y"] = "0.3"
        });
        act.Should().NotThrow();

        // No legend.x should be readable since legend element didn't exist
        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        node!.Format.Should().NotContainKey("legend.x");
    }

    [Fact]
    public void Excel_Chart_Set_LegendLayout_NoLegend_DoesNotThrow()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "NoLegend", ["legend"] = "false", ["data"] = "S1:2,4,6"
        });

        var act = () => _excel.Set("/Sheet1/chart[1]", new()
        {
            ["legend.x"] = "0.8"
        });
        act.Should().NotThrow();

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Should().NotBeNull();
        node!.Format.Should().NotContainKey("legend.x");
    }

    // ==================== Scenario 8: Boundary values ====================

    [Fact]
    public void Pptx_Chart_Set_PlotAreaLayout_BoundaryValues_ZeroAndOne()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Boundary" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "BoundaryChart", ["data"] = "S1:1,2,3"
        });

        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotArea.x"] = "0",
            ["plotArea.y"] = "0",
            ["plotArea.w"] = "1",
            ["plotArea.h"] = "1"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("plotArea.x");
        node.Format["plotArea.x"].ToString().Should().Be("0");
        node.Format["plotArea.w"].ToString().Should().Be("1");
    }

    // ==================== Scenario 9: PPTX plotArea layout persists after reopen ====================

    [Fact]
    public void Pptx_Chart_Set_PlotAreaLayout_PersistsAfterReopen()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Persist" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "PersistChart", ["data"] = "S1:1,2,3"
        });

        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotArea.x"] = "0.11",
            ["plotArea.y"] = "0.22",
            ["plotArea.w"] = "0.66",
            ["plotArea.h"] = "0.55"
        });

        ReopenPptx();

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        node!.Format["plotArea.x"].ToString().Should().Be("0.11");
        node.Format["plotArea.y"].ToString().Should().Be("0.22");
        node.Format["plotArea.w"].ToString().Should().Be("0.66");
        node.Format["plotArea.h"].ToString().Should().Be("0.55");
    }

    // ==================== Scenario 10: Overwrite layout values ====================

    [Fact]
    public void Pptx_Chart_Set_PlotAreaLayout_Overwrite_UpdatesValue()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Overwrite" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "OverwriteChart", ["data"] = "S1:1,2,3"
        });

        _pptx.Set("/slide[1]/chart[1]", new() { ["plotArea.x"] = "0.1" });
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotArea.x"] = "0.25" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        node!.Format["plotArea.x"].ToString().Should().Be("0.25");
    }
}
