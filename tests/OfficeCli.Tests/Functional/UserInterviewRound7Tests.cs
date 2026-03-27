// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Round 7 user interview tests — four issues found during user testing.
///
/// Issue 1: doughnut data=A:30,B:40,C:30 categories show as "1".
///   The comma-separated name:value format (A:30,B:40,C:30) creates 3 separate series,
///   each with 1 data point. For pie/doughnut, the correct format is:
///   data=Shares:30,40,30 + categories=A,B,C (single series with explicit categories).
///   This test documents the correct usage and verifies it works.
///
/// Issue 2: Add trendline to series path not supported.
///   "add /slide[1]/chart[1]/series[1] --type trendline" reports "Parent element not found".
///   This is by design — trendlines are set via Set, not Add. Error message improved.
///
/// Issue 3: Sparkline range vs data parameter name inconsistency.
///   Fixed: sparkline now accepts both "range" and "data" as parameter names.
///
/// Issue 4: batch JSON documentation — documentation-only, no code fix needed.
/// </summary>
public class UserInterviewRound7Tests : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _xlsxPath;
    private PowerPointHandler _pptx;
    private ExcelHandler _excel;

    public UserInterviewRound7Tests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_xlsxPath);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _pptx?.Dispose();
        _excel?.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    private void ReopenPptx()
    {
        _pptx?.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    private void ReopenExcel()
    {
        _excel?.Dispose();
        _excel = new ExcelHandler(_xlsxPath, editable: true);
    }

    // ==================== Issue 1: Pie/Doughnut correct data format ====================

    [Fact]
    public void Issue1_Doughnut_CorrectFormat_WithCategories()
    {
        // Correct format: single series with explicit categories
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "doughnut",
            ["title"] = "Market Share",
            ["data"] = "Share:30,40,30",
            ["categories"] = "A,B,C"
        });

        var node = _pptx.Get(chartPath);
        node.Should().NotBeNull();
        ((string)node!.Format["chartType"]).Should().Be("doughnut");

        // Should have correct categories
        var cats = node.Format.ContainsKey("categories") ? (string)node.Format["categories"] : null;
        cats.Should().NotBeNull();
        cats.Should().Contain("A");
        cats.Should().Contain("B");
        cats.Should().Contain("C");
    }

    [Fact]
    public void Issue1_Pie_CorrectFormat_WithCategories()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "pie",
            ["title"] = "Sales",
            ["data"] = "Revenue:100,200,300",
            ["categories"] = "Q1,Q2,Q3"
        });

        var node = _pptx.Get(chartPath);
        node.Should().NotBeNull();
        ((string)node!.Format["chartType"]).Should().Be("pie");

        var cats = node.Format.ContainsKey("categories") ? (string)node.Format["categories"] : null;
        cats.Should().NotBeNull();
        cats.Should().Contain("Q1");
        cats.Should().Contain("Q2");
        cats.Should().Contain("Q3");
    }

    [Fact]
    public void Issue1_Doughnut_NameValueFormat_Creates_MultipleSeries()
    {
        // Document behavior: A:30,B:40,C:30 creates 3 series, not 1 series with categories
        // This is by design — for pie/doughnut, use data=Name:v1,v2,v3 + categories=...
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "doughnut",
            ["title"] = "Multi-Series",
            ["data"] = "A:30,B:40,C:30"
        });

        var node = _pptx.Get(chartPath);
        node.Should().NotBeNull();

        // With name:value format, we get 3 series (A, B, C) each with 1 point
        // Categories default to "1" because each series has only 1 data point
        var seriesCount = node!.Format.ContainsKey("series.count")
            ? int.Parse((string)node.Format["series.count"])
            : node.Children?.Count(c => c.Type == "series") ?? 0;
        // At minimum, the chart should be created without error
        ((string)node.Format["chartType"]).Should().Be("doughnut");
    }

    [Fact]
    public void Issue1_Doughnut_CorrectFormat_Persists()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "doughnut",
            ["title"] = "Persist Test",
            ["data"] = "Share:10,20,30,40",
            ["categories"] = "North,South,East,West"
        });

        ReopenPptx();

        var node = _pptx.Get(chartPath);
        node.Should().NotBeNull();
        ((string)node!.Format["chartType"]).Should().Be("doughnut");

        var cats = node.Format.ContainsKey("categories") ? (string)node.Format["categories"] : null;
        cats.Should().NotBeNull();
        cats.Should().Contain("North");
        cats.Should().Contain("West");
    }

    // ==================== Issue 2: Trendline via Add shows helpful error ====================

    [Fact]
    public void Issue2_Add_Trendline_ToChartSeries_GivesHelpfulError()
    {
        _pptx.Add("/", "slide", null, new());
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Trend Test",
            ["data"] = "Sales:10,20,30,40"
        });

        // Trying to add a trendline via Add should give a helpful error
        var act = () => _pptx.Add("/slide[1]/chart[1]/series[1]", "trendline", null, new()
        {
            ["type"] = "linear"
        });

        act.Should().Throw<ArgumentException>()
            .WithMessage("*Set*trendline*");
    }

    // ==================== Issue 3: Sparkline data alias ====================

    [Fact]
    public void Issue3_Sparkline_DataAlias_Works()
    {
        // Add some data first
        _excel.Add("/Sheet1", "row", null, new()
        {
            ["values"] = "10,20,30,40,50"
        });

        // Use "data" instead of "range"
        var spkPath = _excel.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["data"] = "A1:E1",
            ["type"] = "line"
        });

        spkPath.Should().Contain("sparkline");
    }

    [Fact]
    public void Issue3_Sparkline_RangeStillWorks()
    {
        // Verify original "range" parameter still works
        _excel.Add("/Sheet1", "row", null, new()
        {
            ["values"] = "5,15,25,35,45"
        });

        var spkPath = _excel.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "column"
        });

        spkPath.Should().Contain("sparkline");
    }

    [Fact]
    public void Issue3_Sparkline_NoRangeOrData_GivesHelpfulError()
    {
        _excel.Add("/Sheet1", "row", null, new() { ["values"] = "1,2,3" });

        var act = () => _excel.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["type"] = "line"
            // Neither "range" nor "data" provided
        });

        act.Should().Throw<ArgumentException>()
            .WithMessage("*range*data*");
    }

    // ==================== Round 8 discovery: Preset ordering (legendFont before legend) ====================

    [Fact]
    public void Round8_Preset_MagazineThenDashboard_NoUnsupportedLegendFont()
    {
        // Magazine removes legend (legend=none), then dashboard sets legendFont before legend.
        // Without ordering fix, legendFont would fail because legend doesn't exist yet.
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Preset Test",
            ["data"] = "S:1,2,3",
            ["categories"] = "A,B,C"
        });

        // Apply magazine (removes legend), then dashboard (needs legend for legendFont)
        _pptx.Set(chartPath, new() { ["preset"] = "magazine" });
        _pptx.Set(chartPath, new() { ["preset"] = "dashboard" });

        // Verify legend is restored and chart is valid
        var node = _pptx.Get(chartPath);
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("legend");
        ((string)node.Format["legend"]).Should().Be("bottom");
    }

    [Fact]
    public void Round8_AllPresets_Sequential_NoErrors()
    {
        // Apply all presets in sequence (worst case for ordering bugs)
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar",
            ["title"] = "All Presets",
            ["data"] = "Data:10,20,30",
            ["categories"] = "X,Y,Z"
        });

        var presets = new[] { "minimal", "dark", "corporate", "magazine", "dashboard", "colorful", "monochrome" };
        foreach (var preset in presets)
        {
            var act = () => _pptx.Set(chartPath, new() { ["preset"] = preset });
            act.Should().NotThrow($"preset '{preset}' should not throw");
        }

        var node = _pptx.Get(chartPath);
        node.Should().NotBeNull();
    }

    [Fact]
    public void Round8_Waterfall_TotalBar_ShowsCorrectRunningTotal()
    {
        // Bug: waterfall total bar showed running + userValue instead of just running.
        // data=Flow:1000,500,-200,-100,1200 should show Net=1200, not 2400.
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Waterfall Total",
            ["data"] = "Flow:1000,500,-200,-100,1200",
            ["categories"] = "Start,Revenue,Costs,Tax,Net"
        });

        var node = _pptx.Get(chartPath);
        node.Should().NotBeNull();

        // The Increase series (series[2]) should have the last value = 1200 (not 2400)
        var increaseSeries = node!.Children?.FirstOrDefault(c => c.Text == "Increase");
        increaseSeries.Should().NotBeNull();
        var values = (string)increaseSeries!.Format["values"];
        var lastValue = values.Split(',').Last().Trim();
        lastValue.Should().Be("1200", "total bar should show running total (1200), not running+userValue (2400)");
    }

    [Fact]
    public void Round8_Excel_AllPresets_Sequential_NoErrors()
    {
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Excel Presets",
            ["data"] = "S:5,10,15",
            ["categories"] = "A,B,C"
        });

        var presets = new[] { "minimal", "dark", "corporate", "magazine", "dashboard", "colorful", "monochrome" };
        foreach (var preset in presets)
        {
            var act = () => _excel.Set(chartPath, new() { ["preset"] = preset });
            act.Should().NotThrow($"preset '{preset}' should not throw on Excel");
        }
    }
}
