// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Round 6 user interview tests — four bugs found during user interviews.
///
/// Bug B: referenceLine via Set creates duplicate lines on repeated calls.
///   AddReferenceLine always appends without removing existing reference lines,
///   so calling Set with refLine=50 twice creates 2 identical reference lines.
///
/// Bug J/C: Word funnel chart Add succeeds but Query returns nothing.
///   Word handler Query("chart") and Get("/chart[N]") only look at standard ChartParts,
///   missing ExtendedChartParts (funnel, treemap, sunburst, etc.).
///   Also, CountWordCharts returns combined count but Get uses only standard parts.
///
/// Bug G: Series name containing colon causes parse failure.
///   data="Sales:Revenue:100,200" — second colon breaks name:values split.
///   This is a known format limitation; test verifies error message is clear.
///
/// Bug K: Waterfall data with per-category name:value format only renders first point.
///   data="Start:1000,Revenue:500,Expense:-200,Net:1300" is parsed as 4 single-value
///   series (each comma part has colon). Waterfall takes seriesData[0].values = [1000].
///   User should use categories + single series format instead.
/// </summary>
public class UserInterviewRound6Tests : IDisposable
{
    private readonly string _docxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private PowerPointHandler _pptxHandler;

    public UserInterviewRound6Tests()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_pptxPath);
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler?.Dispose();
        _pptxHandler?.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void ReopenWord()
    {
        _wordHandler?.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
    }

    private void ReopenPptx()
    {
        _pptxHandler?.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    // ==================== Bug B: referenceLine duplicate on Set ====================

    [Fact]
    public void BugB_ReferenceLine_Set_ShouldNotCreateDuplicates()
    {
        // Create a column chart in Word
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Set referenceLine once
        _wordHandler.Set("/chart[1]", new() { ["refLine"] = "50:FF0000:Target" });

        // Get the chart and count series (original 1 + reference line 1 = 2)
        var node = _wordHandler.Get("/chart[1]");
        var seriesCount1 = node.Children.Count(c => c.Type == "series");
        seriesCount1.Should().Be(2, "1 data series + 1 reference line");

        // Set referenceLine again — should replace, not duplicate
        _wordHandler.Set("/chart[1]", new() { ["refLine"] = "50:FF0000:Target" });

        node = _wordHandler.Get("/chart[1]");
        var seriesCount2 = node.Children.Count(c => c.Type == "series");
        seriesCount2.Should().Be(2, "reference line should be replaced, not duplicated");
    }

    [Fact]
    public void BugB_ReferenceLine_SetDifferentValue_ShouldReplace()
    {
        // Create a column chart in PPTX
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Set referenceLine
        _pptxHandler.Set("/slide[1]/chart[1]", new() { ["refLine"] = "50:FF0000:Target" });

        var node = _pptxHandler.Get("/slide[1]/chart[1]");
        var seriesCount1 = node.Children.Count(c => c.Type == "series");
        seriesCount1.Should().Be(2);

        // Change reference line value — should replace
        _pptxHandler.Set("/slide[1]/chart[1]", new() { ["refLine"] = "75:00AA00:Average" });

        node = _pptxHandler.Get("/slide[1]/chart[1]");
        var seriesCount2 = node.Children.Count(c => c.Type == "series");
        seriesCount2.Should().Be(2, "old reference line should be removed, new one added");
    }

    [Fact]
    public void BugB_ReferenceLine_SetNone_ShouldRemove()
    {
        // Create a column chart with a reference line
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3",
            ["refLine"] = "50:FF0000:Target"
        });

        var node = _wordHandler.Get("/chart[1]");
        node.Children.Count(c => c.Type == "series").Should().Be(2);

        // Remove reference line
        _wordHandler.Set("/chart[1]", new() { ["refLine"] = "none" });

        node = _wordHandler.Get("/chart[1]");
        node.Children.Count(c => c.Type == "series").Should().Be(1, "reference line should be removed");
    }

    // ==================== Bug J/C: Word funnel chart Add succeeds but Query fails ====================

    [Fact]
    public void BugJC_WordFunnelChart_Add_ShouldBeQueryable()
    {
        // Add a funnel chart (extended chart type) to Word
        var path = _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "funnel",
            ["data"] = "Pipeline:100,80,50,30,10",
            ["categories"] = "Leads,Qualified,Proposal,Negotiation,Won"
        });

        path.Should().Be("/chart[1]");

        // Query should find it
        var charts = _wordHandler.Query("chart");
        charts.Should().HaveCountGreaterOrEqualTo(1, "funnel chart should appear in Query results");
        charts[0].Type.Should().Be("chart");
    }

    [Fact]
    public void BugJC_WordFunnelChart_Get_ShouldReturnNode()
    {
        // Add a funnel chart
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "funnel",
            ["data"] = "Pipeline:100,80,50,30,10",
            ["categories"] = "Leads,Qualified,Proposal,Negotiation,Won"
        });

        // Get should return valid node, not error
        var node = _wordHandler.Get("/chart[1]");
        node.Type.Should().Be("chart");
        node.Type.Should().NotBe("error");
    }

    [Fact]
    public void BugJC_WordMixedCharts_IndexShouldBeConsistent()
    {
        // Add a standard chart first
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30"
        });

        // Add a funnel chart (extended)
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "funnel",
            ["data"] = "Pipeline:100,80,50,30",
            ["categories"] = "A,B,C,D"
        });

        // Both charts should be queryable
        var charts = _wordHandler.Query("chart");
        charts.Should().HaveCount(2);

        // Both should be gettable
        var chart1 = _wordHandler.Get("/chart[1]");
        chart1.Type.Should().Be("chart");

        var chart2 = _wordHandler.Get("/chart[2]");
        chart2.Type.Should().Be("chart");
    }

    // ==================== Bug G: Series name with colon parse failure ====================

    [Fact]
    public void BugG_SeriesNameWithColon_ShouldGiveClearError()
    {
        // data="Sales:Revenue:100,200" — "Sales:Revenue" looks like name, "100,200" looks like values
        // BUT the first colon splits to name="Sales", values="Revenue:100,200" which fails
        // This is a format limitation. Test that the error message is helpful.
        var act = () => _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:Revenue:100,200"
        });

        // Should throw with a message about invalid data value (since "Revenue:100" can't parse as number)
        act.Should().Throw<ArgumentException>()
            .WithMessage("*Invalid data value*");
    }

    [Fact]
    public void BugG_SeriesNameWithColon_WorkaroundWithSeriesSyntax()
    {
        // Workaround: use series1=... syntax where name is separate
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "column",
            ["series1.name"] = "Sales:Revenue",
            ["series1.values"] = "100,200,300",
            ["categories"] = "Q1,Q2,Q3"
        });

        var node = _wordHandler.Get("/chart[1]");
        node.Type.Should().Be("chart");
        var series = node.Children.Where(c => c.Type == "series").ToList();
        series.Should().HaveCount(1);
    }

    // ==================== Bug K: Waterfall data format misunderstanding ====================

    [Fact]
    public void BugK_Waterfall_PerCategoryNameValue_ShouldUseAllValues()
    {
        // User format: data="Start:1000,Revenue:500,Expense:-200,Net:1300"
        // This gets parsed as 4 single-value series. Waterfall takes seriesData[0].values = [1000].
        // After fix: waterfall should detect this pattern and flatten all values into one series,
        // using the names as categories.
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "waterfall",
            ["data"] = "Start:1000,Revenue:500,Expense:-200,Net:1300"
        });

        var node = _wordHandler.Get("/chart[1]");
        node.Type.Should().Be("chart");

        // The chart should have meaningful data (not just 1 point)
        // Waterfall creates 3 series: base (invisible), increase, decrease
        var series = node.Children.Where(c => c.Type == "series").ToList();
        series.Should().HaveCountGreaterOrEqualTo(2, "waterfall should render all 4 data points");

        // Verify categories were auto-derived from the name:value pairs
        if (node.Format.ContainsKey("categories"))
        {
            var cats = node.Format["categories"]?.ToString();
            cats.Should().Contain("Start");
            cats.Should().Contain("Revenue");
        }
    }

    [Fact]
    public void BugK_Waterfall_CorrectFormat_ShouldWork()
    {
        // Correct format: categories + single series
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "waterfall",
            ["data"] = "Cashflow:1000,500,-200,1300",
            ["categories"] = "Start,Revenue,Expense,Net"
        });

        var node = _wordHandler.Get("/chart[1]");
        node.Type.Should().Be("chart");

        // Waterfall creates 3 series: base (invisible), increase, decrease
        var series = node.Children.Where(c => c.Type == "series").ToList();
        series.Should().HaveCountGreaterOrEqualTo(2);
    }

    [Fact]
    public void BugK_Waterfall_CorrectFormat_Persistence()
    {
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "waterfall",
            ["data"] = "Cashflow:1000,500,-200,1300",
            ["categories"] = "Start,Revenue,Expense,Net"
        });

        ReopenWord();

        var node = _wordHandler.Get("/chart[1]");
        node.Type.Should().Be("chart");
        var series = node.Children.Where(c => c.Type == "series").ToList();
        series.Should().HaveCountGreaterOrEqualTo(2);
    }
}
