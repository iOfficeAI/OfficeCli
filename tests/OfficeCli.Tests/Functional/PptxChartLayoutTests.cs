// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// White-box tests for chart ManualLayout: plotArea, title, legend, trendlineLabel, displayUnitsLabel.
/// Covers SetManualLayoutProperty / ReadManualLayout helpers and ChartSetter/ChartReader integration.
/// </summary>
public class PptxChartLayoutTests : IDisposable
{
    private readonly string _pptxPath;
    private PowerPointHandler _pptx;

    public PptxChartLayoutTests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_pptxPath);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _pptx.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void Reopen()
    {
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    private string AddColumnChart(string title = "Sales", bool withLegend = true)
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "S" });
        var props = new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["title"] = title,
            ["data"] = "Q1:10,20,30;Q2:15,25,35",
            ["categories"] = "Jan,Feb,Mar"
        };
        if (withLegend)
            props["legend"] = "bottom";
        return _pptx.Add("/slide[1]", "chart", null, props);
    }

    // ==================== plotArea.x/y/w/h ====================

    [Fact]
    public void PlotArea_SetXY_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.1", ["plotarea.y"] = "0.2" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("plotArea.x");
        node.Format.Should().ContainKey("plotArea.y");
        ((string)node.Format["plotArea.x"]).Should().Be("0.1");
        ((string)node.Format["plotArea.y"]).Should().Be("0.2");
    }

    [Fact]
    public void PlotArea_SetWH_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.w"] = "0.75", ["plotarea.h"] = "0.65" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("plotArea.w");
        node.Format.Should().ContainKey("plotArea.h");
        ((string)node.Format["plotArea.w"]).Should().Be("0.75");
        ((string)node.Format["plotArea.h"]).Should().Be("0.65");
    }

    [Fact]
    public void PlotArea_PartialSet_OnlySetKeysReturned()
    {
        AddColumnChart();
        // Only set x and w, not y and h
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.05", ["plotarea.w"] = "0.8" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("plotArea.x");
        node.Format.Should().ContainKey("plotArea.w");
        node.Format.Should().NotContainKey("plotArea.y");
        node.Format.Should().NotContainKey("plotArea.h");
    }

    [Fact]
    public void PlotArea_SetZero_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.0", ["plotarea.y"] = "0.0" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("plotArea.x");
        ((string)node.Format["plotArea.x"]).Should().Be("0");
    }

    [Fact]
    public void PlotArea_SetOne_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.w"] = "1.0", ["plotarea.h"] = "1.0" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.w"]).Should().Be("1");
        ((string)node.Format["plotArea.h"]).Should().Be("1");
    }

    [Fact]
    public void PlotArea_HighPrecision_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.123456" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.x"]).Should().Be("0.123456");
    }

    [Fact]
    public void PlotArea_OverwriteValue_UpdatesCorrectly()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.1" });
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.3" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.x"]).Should().Be("0.3");
    }

    [Fact]
    public void PlotArea_LayoutTarget_IsInner()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.1" });
        _pptx.Dispose();

        // Verify LayoutTarget=Inner in raw XML
        using var pkg = PresentationDocument.Open(_pptxPath, false);
        var slide = pkg.PresentationPart!.SlideParts.First();
        var chartPart = slide.ChartParts.First();
        var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
        var layout = plotArea.GetFirstChild<C.Layout>();
        layout.Should().NotBeNull();
        var ml = layout!.GetFirstChild<C.ManualLayout>();
        ml.Should().NotBeNull();
        var layoutTarget = ml!.GetFirstChild<C.LayoutTarget>();
        layoutTarget.Should().NotBeNull();
        layoutTarget!.Val!.Value.Should().Be(C.LayoutTargetValues.Inner);

        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    [Fact]
    public void PlotArea_LeftModeAndTopMode_AreEdge()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.1" });
        _pptx.Dispose();

        using var pkg = PresentationDocument.Open(_pptxPath, false);
        var slide = pkg.PresentationPart!.SlideParts.First();
        var chartPart = slide.ChartParts.First();
        var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
        var ml = plotArea.GetFirstChild<C.Layout>()!.GetFirstChild<C.ManualLayout>()!;
        var leftMode = ml.GetFirstChild<C.LeftMode>();
        var topMode = ml.GetFirstChild<C.TopMode>();
        leftMode.Should().NotBeNull();
        topMode.Should().NotBeNull();
        leftMode!.Val!.Value.Should().Be(C.LayoutModeValues.Edge);
        topMode!.Val!.Value.Should().Be(C.LayoutModeValues.Edge);

        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    [Fact]
    public void PlotArea_NegativeValue_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "-0.1" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("plotArea.x");
        ((string)node.Format["plotArea.x"]).Should().Be("-0.1");
    }

    [Fact]
    public void PlotArea_GreaterThanOne_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.w"] = "1.5" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.w"]).Should().Be("1.5");
    }

    [Fact]
    public void PlotArea_AllFour_Persist()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotarea.x"] = "0.1",
            ["plotarea.y"] = "0.15",
            ["plotarea.w"] = "0.8",
            ["plotarea.h"] = "0.7"
        });
        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.x"]).Should().Be("0.1");
        ((string)node.Format["plotArea.y"]).Should().Be("0.15");
        ((string)node.Format["plotArea.w"]).Should().Be("0.8");
        ((string)node.Format["plotArea.h"]).Should().Be("0.7");
    }

    // ==================== title.x/y/w/h ====================

    [Fact]
    public void TitleLayout_SetXY_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.x"] = "0.3", ["title.y"] = "0.05" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("title.x");
        node.Format.Should().ContainKey("title.y");
        ((string)node.Format["title.x"]).Should().Be("0.3");
        ((string)node.Format["title.y"]).Should().Be("0.05");
    }

    [Fact]
    public void TitleLayout_SetWH_ReadBack()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.w"] = "0.4", ["title.h"] = "0.1" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["title.w"]).Should().Be("0.4");
        ((string)node.Format["title.h"]).Should().Be("0.1");
    }

    [Fact]
    public void TitleLayout_NoLayoutTarget()
    {
        // title does NOT get LayoutTarget=Inner
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.x"] = "0.3" });
        _pptx.Dispose();

        using var pkg = PresentationDocument.Open(_pptxPath, false);
        var slide = pkg.PresentationPart!.SlideParts.First();
        var chartPart = slide.ChartParts.First();
        var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var titleEl = chart.GetFirstChild<C.Title>()!;
        var ml = titleEl.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>();
        ml.Should().NotBeNull();
        var layoutTarget = ml!.GetFirstChild<C.LayoutTarget>();
        layoutTarget.Should().BeNull("title ManualLayout should not have LayoutTarget");

        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    [Fact]
    public void TitleLayout_LeftModeAndTopMode_AreEdge()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.x"] = "0.3" });
        _pptx.Dispose();

        using var pkg = PresentationDocument.Open(_pptxPath, false);
        var slide = pkg.PresentationPart!.SlideParts.First();
        var chartPart = slide.ChartParts.First();
        var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var titleEl = chart.GetFirstChild<C.Title>()!;
        var ml = titleEl.GetFirstChild<C.Layout>()!.GetFirstChild<C.ManualLayout>()!;
        ml.GetFirstChild<C.LeftMode>()!.Val!.Value.Should().Be(C.LayoutModeValues.Edge);
        ml.GetFirstChild<C.TopMode>()!.Val!.Value.Should().Be(C.LayoutModeValues.Edge);

        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    [Fact]
    public void TitleLayout_OverwriteValue_UpdatesCorrectly()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.x"] = "0.2" });
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.x"] = "0.5" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["title.x"]).Should().Be("0.5");
    }

    [Fact]
    public void TitleLayout_WithoutTitle_ReturnsUnsupported()
    {
        // No title on chart — setting title layout should gracefully fail (add to unsupported, not throw)
        _pptx.Add("/", "slide", null, new() { ["title"] = "S" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["data"] = "Q1:10,20"
            // deliberately no title
        });

        // Should not throw
        var act = () => _pptx.Set("/slide[1]/chart[1]", new() { ["title.x"] = "0.3" });
        act.Should().NotThrow();
    }

    [Fact]
    public void TitleLayout_AllFour_Persist()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["title.x"] = "0.25",
            ["title.y"] = "0.02",
            ["title.w"] = "0.5",
            ["title.h"] = "0.08"
        });
        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["title.x"]).Should().Be("0.25");
        ((string)node.Format["title.y"]).Should().Be("0.02");
        ((string)node.Format["title.w"]).Should().Be("0.5");
        ((string)node.Format["title.h"]).Should().Be("0.08");
    }

    // ==================== legend.x/y/w/h ====================

    [Fact]
    public void LegendLayout_SetXY_ReadBack()
    {
        AddColumnChart(withLegend: true);
        _pptx.Set("/slide[1]/chart[1]", new() { ["legend.x"] = "0.7", ["legend.y"] = "0.4" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("legend.x");
        node.Format.Should().ContainKey("legend.y");
        ((string)node.Format["legend.x"]).Should().Be("0.7");
        ((string)node.Format["legend.y"]).Should().Be("0.4");
    }

    [Fact]
    public void LegendLayout_SetWH_ReadBack()
    {
        AddColumnChart(withLegend: true);
        _pptx.Set("/slide[1]/chart[1]", new() { ["legend.w"] = "0.2", ["legend.h"] = "0.15" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend.w"]).Should().Be("0.2");
        ((string)node.Format["legend.h"]).Should().Be("0.15");
    }

    [Fact]
    public void LegendLayout_NoLayoutTarget()
    {
        AddColumnChart(withLegend: true);
        _pptx.Set("/slide[1]/chart[1]", new() { ["legend.x"] = "0.7" });
        _pptx.Dispose();

        using var pkg = PresentationDocument.Open(_pptxPath, false);
        var slide = pkg.PresentationPart!.SlideParts.First();
        var chartPart = slide.ChartParts.First();
        var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var legendEl = chart.GetFirstChild<C.Legend>()!;
        var ml = legendEl.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>();
        ml.Should().NotBeNull();
        ml!.GetFirstChild<C.LayoutTarget>().Should().BeNull("legend ManualLayout should not have LayoutTarget");

        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    [Fact]
    public void LegendLayout_OverwriteValue_UpdatesCorrectly()
    {
        AddColumnChart(withLegend: true);
        _pptx.Set("/slide[1]/chart[1]", new() { ["legend.x"] = "0.6" });
        _pptx.Set("/slide[1]/chart[1]", new() { ["legend.x"] = "0.8" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend.x"]).Should().Be("0.8");
    }

    [Fact]
    public void LegendLayout_WithoutLegend_ReturnsUnsupported()
    {
        // No legend — setting legend layout should not throw
        _pptx.Add("/", "slide", null, new() { ["title"] = "S" });
        _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Sales",
            ["data"] = "Q1:10,20"
        });

        var act = () => _pptx.Set("/slide[1]/chart[1]", new() { ["legend.x"] = "0.5" });
        act.Should().NotThrow();
    }

    [Fact]
    public void LegendLayout_AllFour_Persist()
    {
        AddColumnChart(withLegend: true);
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["legend.x"] = "0.7",
            ["legend.y"] = "0.4",
            ["legend.w"] = "0.25",
            ["legend.h"] = "0.18"
        });
        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend.x"]).Should().Be("0.7");
        ((string)node.Format["legend.y"]).Should().Be("0.4");
        ((string)node.Format["legend.w"]).Should().Be("0.25");
        ((string)node.Format["legend.h"]).Should().Be("0.18");
    }

    // ==================== Combined: plotArea + title + legend ====================

    [Fact]
    public void AllThreeLayoutAreas_SetTogether_ReadBackIndependently()
    {
        AddColumnChart(withLegend: true);
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotarea.x"] = "0.1",
            ["plotarea.y"] = "0.15",
            ["plotarea.w"] = "0.75",
            ["plotarea.h"] = "0.65",
            ["title.x"] = "0.3",
            ["title.y"] = "0.02",
            ["legend.x"] = "0.7",
            ["legend.y"] = "0.4"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.x"]).Should().Be("0.1");
        ((string)node.Format["plotArea.y"]).Should().Be("0.15");
        ((string)node.Format["title.x"]).Should().Be("0.3");
        ((string)node.Format["title.y"]).Should().Be("0.02");
        ((string)node.Format["legend.x"]).Should().Be("0.7");
        ((string)node.Format["legend.y"]).Should().Be("0.4");
    }

    // ==================== trendlineLabel (no trendline = graceful skip) ====================

    [Fact]
    public void TrendlineLabel_WithoutTrendline_DoesNotThrow()
    {
        AddColumnChart();
        var act = () => _pptx.Set("/slide[1]/chart[1]", new() { ["trendlinelabel.x"] = "0.4" });
        act.Should().NotThrow();
    }

    [Fact]
    public void TrendlineLabel_WithoutTrendline_NotInFormat()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["trendlinelabel.x"] = "0.4" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().NotContainKey("trendlineLabel.x");
    }

    // ==================== displayUnitsLabel (no displayUnits = graceful skip) ====================

    [Fact]
    public void DisplayUnitsLabel_WithoutDisplayUnits_DoesNotThrow()
    {
        AddColumnChart();
        var act = () => _pptx.Set("/slide[1]/chart[1]", new() { ["displayunitslabel.x"] = "0.5" });
        act.Should().NotThrow();
    }

    [Fact]
    public void DisplayUnitsLabel_WithoutDisplayUnits_NotInFormat()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["displayunitslabel.x"] = "0.5" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().NotContainKey("displayUnitsLabel.x");
    }

    // ==================== Precision: 0.######  ====================

    [Fact]
    public void PlotArea_SixDecimalPrecision_RoundTrip()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.654321" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.x"]).Should().Be("0.654321");
    }

    [Fact]
    public void PlotArea_TrailingZerosTrimmed()
    {
        // "0.500000" should be trimmed to "0.5" by "0.######" format
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.5" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.x"]).Should().Be("0.5");
        // Must NOT be "0.500000"
        ((string)node.Format["plotArea.x"]).Should().NotBe("0.500000");
    }

    // ==================== Multiple overwrites: no duplicate ManualLayout children ====================

    [Fact]
    public void PlotArea_MultipleOverwrites_NoOrphanChildren()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.1" });
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.2" });
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "0.3" });
        _pptx.Dispose();

        using var pkg = PresentationDocument.Open(_pptxPath, false);
        var slide = pkg.PresentationPart!.SlideParts.First();
        var chartPart = slide.ChartParts.First();
        var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
        var ml = plotArea.GetFirstChild<C.Layout>()!.GetFirstChild<C.ManualLayout>()!;

        // Only one C.Left child should exist
        ml.Elements<C.Left>().Count().Should().Be(1);
        ((double)ml.GetFirstChild<C.Left>()!.Val!).Should().BeApproximately(0.3, 1e-9);

        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    // ==================== Invalid value (non-numeric) ====================

    [Fact]
    public void PlotArea_InvalidValue_DoesNotThrow()
    {
        AddColumnChart();
        var act = () => _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "notanumber" });
        act.Should().NotThrow();
    }

    [Fact]
    public void PlotArea_InvalidValue_NotInFormat()
    {
        AddColumnChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotarea.x"] = "notanumber" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().NotContainKey("plotArea.x");
    }
}
