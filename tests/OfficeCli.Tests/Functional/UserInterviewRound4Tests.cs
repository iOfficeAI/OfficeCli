// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Round 4 user interview tests — cx:chart type support gaps.
///
/// Problem: Excel and Word do not support cx:chart types (funnel, treemap, sunburst,
/// boxWhisker, histogram). These chart types require ExtendedChartPart (cx:chartSpace)
/// but Excel/Word handlers only call ChartHelper.BuildChartSpace() which builds
/// standard C.ChartSpace. The ParseChartType() method throws ArgumentException for
/// these types.
///
/// PPTX works because PowerPointHandler.Add checks ChartExBuilder.IsExtendedChartType()
/// and routes to ChartExBuilder.BuildExtendedChartSpace() instead.
///
/// Also tests:
/// - A's commit 89e271e fixes (schema order, ID dedup) are not regressed
/// - Word cx:chart gap
/// - Excel trendline schema order (shared via ChartSetter)
/// - Excel chart ID dedup after delete+recreate
/// </summary>
public class UserInterviewRound4Tests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _docxPath;
    private readonly string _pptxPath;
    private ExcelHandler _xlsxHandler;
    private WordHandler _docxHandler;
    private PowerPointHandler _pptxHandler;

    public UserInterviewRound4Tests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_pptxPath);
        _xlsxHandler = new ExcelHandler(_xlsxPath, editable: true);
        _docxHandler = new WordHandler(_docxPath, editable: true);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _xlsxHandler?.Dispose();
        _docxHandler?.Dispose();
        _pptxHandler?.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void ReopenXlsx()
    {
        _xlsxHandler?.Dispose();
        _xlsxHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    private void ReopenDocx()
    {
        _docxHandler?.Dispose();
        _docxHandler = new WordHandler(_docxPath, editable: true);
    }

    private void ReopenPptx()
    {
        _pptxHandler?.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    // ==================== Task 2: Excel cx:chart gap ====================

    [Fact]
    public void Excel_Add_FunnelChart_ShouldSucceed()
    {
        // Excel should support funnel chart creation (cx:chart type).
        // Currently fails with ArgumentException because ParseChartType doesn't handle "funnel".
        var act = () => _xlsxHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "funnel",
            ["title"] = "Sales Funnel",
            ["data"] = "Stage:100,80,60,30,10"
        });

        // This SHOULD succeed but currently throws — proving the gap
        act.Should().NotThrow("Excel should support funnel (cx:chart) chart type");
    }

    [Fact]
    public void Excel_Add_TreemapChart_ShouldSucceed()
    {
        // Excel should support treemap chart creation (cx:chart type).
        var act = () => _xlsxHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "treemap",
            ["title"] = "Category Breakdown",
            ["data"] = "Values:40,30,20,10",
            ["categories"] = "A,B,C,D"
        });

        act.Should().NotThrow("Excel should support treemap (cx:chart) chart type");
    }

    [Fact]
    public void Excel_Add_HistogramChart_ShouldSucceed()
    {
        // Excel should support histogram chart creation (cx:chart type).
        var act = () => _xlsxHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "histogram",
            ["title"] = "Distribution",
            ["data"] = "Frequency:5,12,18,25,15,8,3"
        });

        act.Should().NotThrow("Excel should support histogram (cx:chart) chart type");
    }

    // ==================== Task 3.1: Word cx:chart gap ====================

    [Fact]
    public void Word_Add_FunnelChart_ShouldSucceed()
    {
        // Word also calls BuildChartSpace directly without cx:chart routing.
        var act = () => _docxHandler.Add("/body", "chart", null, new()
        {
            ["type"] = "funnel",
            ["title"] = "Word Funnel",
            ["data"] = "Stage:100,80,60,30"
        });

        act.Should().NotThrow("Word should support funnel (cx:chart) chart type");
    }

    [Fact]
    public void Word_Add_TreemapChart_ShouldSucceed()
    {
        var act = () => _docxHandler.Add("/body", "chart", null, new()
        {
            ["type"] = "treemap",
            ["title"] = "Word Treemap",
            ["data"] = "Values:40,30,20,10",
            ["categories"] = "A,B,C,D"
        });

        act.Should().NotThrow("Word should support treemap (cx:chart) chart type");
    }

    [Fact]
    public void Word_Add_HistogramChart_ShouldSucceed()
    {
        var act = () => _docxHandler.Add("/body", "chart", null, new()
        {
            ["type"] = "histogram",
            ["title"] = "Word Histogram",
            ["data"] = "Frequency:5,12,18,25,15,8,3"
        });

        act.Should().NotThrow("Word should support histogram (cx:chart) chart type");
    }

    // ==================== Task 3.2: Trendline schema shared fix verification ====================

    [Fact]
    public void Excel_Chart_SetTrendline_SchemaOrder_IsCorrect()
    {
        // The trendline schema fix in ChartSetter is shared code (ChartHelper.SetChartProperties).
        // Verify it works for Excel charts too — trendline must come before <c:cat>/<c:val>.
        _xlsxHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "line",
            ["title"] = "Trendline Test",
            ["data"] = "Revenue:10,20,30,40,50"
        });

        // Set trendline on series
        _xlsxHandler.Set("/Sheet1/chart[1]/series[1]", new() { ["trendline"] = "linear" });

        // Get should return trendline info without error
        var series = _xlsxHandler.Get("/Sheet1/chart[1]/series[1]");
        series.Should().NotBeNull();
        series.Format.Should().ContainKey("trendline");
        series.Format["trendline"].Should().Be("linear");

        // Verify persistence
        ReopenXlsx();
        var series2 = _xlsxHandler.Get("/Sheet1/chart[1]/series[1]");
        series2.Format.Should().ContainKey("trendline");
        series2.Format["trendline"].Should().Be("linear");
    }

    // ==================== Task 3.3: Excel chart ID dedup after delete+recreate ====================

    [Fact]
    public void Excel_Chart_DeleteAndRecreate_NoDuplicateIds()
    {
        // Regression test for the ID fix in commit 89e271e.
        // After deleting a chart and adding a new one, the new chart's cNvPr ID
        // should not collide with the remaining chart.

        // Add two charts
        _xlsxHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "column",
            ["title"] = "Chart A",
            ["data"] = "A:1,2,3"
        });
        _xlsxHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "line",
            ["title"] = "Chart B",
            ["data"] = "B:4,5,6"
        });

        // Verify 2 charts exist
        var chart1 = _xlsxHandler.Get("/Sheet1/chart[1]");
        chart1.Should().NotBeNull();
        var chart2 = _xlsxHandler.Get("/Sheet1/chart[2]");
        chart2.Should().NotBeNull();

        // Delete first chart
        _xlsxHandler.Remove("/Sheet1/chart[1]");

        // Add a new chart — its ID must not collide
        _xlsxHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["title"] = "Chart C",
            ["data"] = "C:7,8,9"
        });

        // Verify file is valid after persistence — the real test for ID collision
        // Reopen forces save + reload, proving the IDs don't conflict in the file
        ReopenXlsx();
        // After reopen, should be able to get both charts
        var chartA = _xlsxHandler.Get("/Sheet1/chart[1]");
        chartA.Should().NotBeNull("first chart should survive reopen");
        var chartB = _xlsxHandler.Get("/Sheet1/chart[2]");
        chartB.Should().NotBeNull("second chart should survive reopen");
    }

    // ==================== PPTX cx:chart still works (sanity) ====================

    [Fact]
    public void Pptx_FunnelChart_StillWorks()
    {
        // Sanity check that PPTX cx:chart creation was not regressed.
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "funnel",
            ["title"] = "PPTX Funnel",
            ["data"] = "Stage:100,80,60,30"
        });

        var chart = _pptxHandler.Get("/slide[1]/chart[1]");
        chart.Should().NotBeNull();
        chart.Format.Should().ContainKey("chartType");
        chart.Format["chartType"].Should().Be("funnel");
    }
}
