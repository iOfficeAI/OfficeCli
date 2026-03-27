// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for enhanced chart features:
/// #1 Data label position/font, #2 Gridlines/plot area, #3 Per-series styling,
/// #4 Chart style ID, #5 Transparency, #6 Gradient fill, #7 Secondary axis
/// </summary>
public class ChartEnhancedTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private ExcelHandler _excel;
    private PowerPointHandler _pptx;

    public ChartEnhancedTests()
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

    private string AddExcelChart(Dictionary<string, string> extraProps = null)
    {
        var props = new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["title"] = "Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        };
        if (extraProps != null) foreach (var kv in extraProps) props[kv.Key] = kv.Value;
        return _excel.Add("/Sheet1", "chart", null, props);
    }

    private string AddPptxChart(Dictionary<string, string> extraProps = null)
    {
        _pptx.Add("/", "slide", null, new());
        var props = new Dictionary<string, string>
        {
            ["chartType"] = "line",
            ["title"] = "Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        };
        if (extraProps != null) foreach (var kv in extraProps) props[kv.Key] = kv.Value;
        return _pptx.Add("/slide[1]", "chart", null, props);
    }

    // ==================== #1 Data Label Position/Font ====================

    [Fact]
    public void Excel_Set_DataLabelPosition_Center()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dataLabels"] = "value", ["labelPos"] = "center" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dataLabels"].Should().Be("value");
        node.Format["labelPos"].Should().Be("ctr");

        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 0);
        node2.Format["labelPos"].Should().Be("ctr");
    }

    [Fact]
    public void Excel_Set_DataLabelPosition_OutsideEnd()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dataLabels"] = "value", ["labelPos"] = "outsideEnd" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["labelPos"].Should().Be("outEnd");
    }

    [Fact]
    public void Excel_Set_LabelFont()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dataLabels"] = "value", ["labelFont"] = "12:FF0000:true" });

        // Just verify no exception and data labels still present
        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dataLabels"].Should().Be("value");
    }

    [Fact]
    public void Pptx_Set_DataLabelPosition()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["dataLabels"] = "value,category", ["labelPos"] = "top" });

        var node = _pptx.Get(chartPath, depth: 0);
        ((string)node.Format["dataLabels"]).Should().Contain("value");
        node.Format["labelPos"].Should().Be("t");
    }

    // ==================== #2 Gridlines / Plot Area ====================

    [Fact]
    public void Excel_Set_Gridlines_WithColor()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["gridlines"] = "CCCCCC:0.5" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["gridlines"].Should().Be("true");

        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 0);
        node2.Format["gridlines"].Should().Be("true");
    }

    [Fact]
    public void Excel_Set_Gridlines_None_RemovesThem()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["gridlines"] = "none" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format.Should().ContainKey("gridlines");
        node.Format["gridlines"].Should().Be("false");
    }

    [Fact]
    public void Excel_Set_MinorGridlines()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["minorGridlines"] = "DDDDDD:0.3:dot" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["minorGridlines"].Should().Be("true");
    }

    [Fact]
    public void Excel_Set_PlotFill()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["plotFill"] = "F0F0F0" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["plotFill"].Should().Be("#F0F0F0");

        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 0);
        node2.Format["plotFill"].Should().Be("#F0F0F0");
    }

    [Fact]
    public void Excel_Set_ChartFill()
    {
        var chartPath = AddExcelChart();
        // chartFill sets the chart area (outer) background
        _excel.Set(chartPath, new() { ["chartFill"] = "FFFFFF" });
        // Just verify no exception
        var node = _excel.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
    }

    // ==================== #3 Per-Series Styling ====================

    [Fact]
    public void Excel_Set_LineWidth()
    {
        var chartPath = AddExcelChart(new() { ["chartType"] = "line" });
        _excel.Set(chartPath, new() { ["lineWidth"] = "2.5" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children.Should().NotBeEmpty();
        var series1 = node.Children[0];
        series1.Format["lineWidth"].Should().Be(2.5);
    }

    [Fact]
    public void Excel_Set_LineDash()
    {
        var chartPath = AddExcelChart(new() { ["chartType"] = "line" });
        _excel.Set(chartPath, new() { ["lineDash"] = "dash" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format["lineDash"].Should().Be("sysDash");
    }

    [Fact]
    public void Excel_Set_Marker()
    {
        var chartPath = AddExcelChart(new() { ["chartType"] = "line" });
        _excel.Set(chartPath, new() { ["marker"] = "diamond:8:FF0000" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format["marker"].Should().Be("diamond");
        node.Children[0].Format["markerSize"].Should().Be((byte)8);

        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 1);
        node2.Children[0].Format["marker"].Should().Be("diamond");
    }

    [Fact]
    public void Excel_Set_MarkerSize()
    {
        var chartPath = AddExcelChart(new() { ["chartType"] = "line" });
        _excel.Set(chartPath, new() { ["marker"] = "circle", ["markerSize"] = "10" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format["marker"].Should().Be("circle");
        node.Children[0].Format["markerSize"].Should().Be((byte)10);
    }

    [Fact]
    public void Pptx_Set_LineWidth_And_Dash()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["lineWidth"] = "3", ["lineDash"] = "dot" });

        var node = _pptx.Get(chartPath, depth: 1);
        node.Children[0].Format["lineWidth"].Should().Be(3.0);
        node.Children[0].Format["lineDash"].Should().Be("sysDot");
    }

    // ==================== #4 Chart Style ID ====================

    [Fact]
    public void Excel_Set_StyleId()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["style"] = "26" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["style"].Should().Be((byte)26);

        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 0);
        node2.Format["style"].Should().Be((byte)26);
    }

    [Fact]
    public void Pptx_Set_StyleId()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["style"] = "10" });

        var node = _pptx.Get(chartPath, depth: 0);
        node.Format["style"].Should().Be((byte)10);
    }

    // ==================== #5 Transparency ====================

    [Fact]
    public void Excel_Set_Transparency()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["transparency"] = "30" });

        var node = _excel.Get(chartPath, depth: 1);
        // 30% transparency = 70% opacity = 70000 alpha
        node.Children[0].Format["alpha"].Should().Be(70000);

        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 1);
        node2.Children[0].Format["alpha"].Should().Be(70000);
    }

    [Fact]
    public void Excel_Set_Opacity()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["opacity"] = "50" });

        var node = _excel.Get(chartPath, depth: 1);
        // 50% opacity = 50000 alpha
        node.Children[0].Format["alpha"].Should().Be(50000);
    }

    // ==================== #6 Gradient Fill ====================

    [Fact]
    public void Excel_Set_Gradient_TwoColor()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["gradient"] = "FF0000-0000FF" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format["gradient"].Should().Be("true");

        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 1);
        node2.Children[0].Format["gradient"].Should().Be("true");
    }

    [Fact]
    public void Excel_Set_Gradient_ThreeColor_WithAngle()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["gradient"] = "FF0000-00FF00-0000FF:90" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format["gradient"].Should().Be("true");
    }

    [Fact]
    public void Excel_Set_Gradients_PerSeries()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["gradients"] = "FF0000-0000FF;00FF00-FFFF00" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format["gradient"].Should().Be("true");
        node.Children[1].Format["gradient"].Should().Be("true");
    }

    [Fact]
    public void Pptx_Set_Gradient()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["gradient"] = "4472C4-ED7D31:45" });

        var node = _pptx.Get(chartPath, depth: 1);
        node.Children[0].Format["gradient"].Should().Be("true");
    }

    // ==================== #7 Secondary Axis ====================

    [Fact]
    public void Excel_Set_SecondaryAxis()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["secondary"] = "2" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["secondaryAxis"].Should().Be("true");

        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 0);
        node2.Format["secondaryAxis"].Should().Be("true");
    }

    [Fact]
    public void Pptx_Set_SecondaryAxis()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["secondary"] = "2" });

        var node = _pptx.Get(chartPath, depth: 0);
        node.Format["secondaryAxis"].Should().Be("true");
    }

    // ==================== Combined Styling ====================

    [Fact]
    public void Excel_Combined_LineChart_FullStyling()
    {
        var chartPath = AddExcelChart(new() { ["chartType"] = "line" });

        // Apply multiple styles at once
        _excel.Set(chartPath, new()
        {
            ["dataLabels"] = "value",
            ["labelPos"] = "top",
            ["gridlines"] = "E0E0E0:0.3",
            ["plotFill"] = "FAFAFA",
            ["lineWidth"] = "2",
            ["lineDash"] = "solid",
            ["marker"] = "circle:6",
            ["style"] = "10"
        });

        var node = _excel.Get(chartPath, depth: 1);
        node.Format["dataLabels"].Should().Be("value");
        node.Format["labelPos"].Should().Be("t");
        node.Format["gridlines"].Should().Be("true");
        node.Format.Should().ContainKey("plotFill");
        node.Format["plotFill"].Should().Be("#FAFAFA");
        node.Format["style"].Should().Be((byte)10);
        node.Children[0].Format["lineWidth"].Should().Be(2.0);
        node.Children[0].Format["marker"].Should().Be("circle");
        node.Children[0].Format["markerSize"].Should().Be((byte)6);

        // Persistence
        ReopenExcel();
        var node2 = _excel.Get(chartPath, depth: 1);
        node2.Format["style"].Should().Be((byte)10);
        node2.Format.Should().ContainKey("plotFill");
        node2.Children[0].Format["lineWidth"].Should().Be(2.0);
    }

    [Fact]
    public void Excel_ColumnChart_GradientAndTransparency()
    {
        var chartPath = AddExcelChart();

        _excel.Set(chartPath, new()
        {
            ["gradient"] = "4472C4-83B9E8",
            ["dataLabels"] = "value",
            ["labelPos"] = "outsideEnd",
            ["gridlines"] = "EEEEEE:0.25:dot"
        });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format["gradient"].Should().Be("true");
        node.Format["dataLabels"].Should().Be("value");
        node.Format["labelPos"].Should().Be("outEnd");
    }

    [Fact]
    public void Excel_DualAxis_WithDifferentStyles()
    {
        var chartPath = AddExcelChart();

        // Move series 2 to secondary axis
        _excel.Set(chartPath, new() { ["secondary"] = "2" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["secondaryAxis"].Should().Be("true");
        // Both series should still exist
        node.Format["seriesCount"].Should().BeOfType<int>().Which.Should().BeGreaterOrEqualTo(2);
    }
}
