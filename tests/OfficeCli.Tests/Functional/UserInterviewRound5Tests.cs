// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Round 5 user interview tests — three bugs found by user A.
///
/// Bug 1 (Critical): Doughnut chart --json serialization crash.
///   get "/slide[N]/chart[1]" --json crashes with "JsonTypeInfo metadata for type 'System.Byte'"
///   when chartType=doughnut. Root cause: ChartReader stores holeSize as byte (from C.HoleSize.Val
///   which is ByteValue), but AppJsonContext only registers int/long/short/uint/double/string/bool.
///   System.Byte is not registered in the source-generated JSON context.
///
/// Bug 2 (Medium): Trendline not visible at chart level Get.
///   After Add chart with trendline=linear, get chart doesn't show trendline in Format.
///   Only get chart/series[1] shows it. Agent-unfriendly: agent sets property but can't verify
///   at the natural chart level.
///
/// Bug 3 (Minor): PRESET=corporate uppercase key false positive.
///   set --prop PRESET=corporate reports success but actually doesn't take effect.
///   Should either work (case-insensitive) or report unsupported.
/// </summary>
public class UserInterviewRound5Tests : IDisposable
{
    private readonly string _pptxPath;
    private PowerPointHandler _handler;

    public UserInterviewRound5Tests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_pptxPath);
        _handler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _handler?.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void Reopen()
    {
        _handler?.Dispose();
        _handler = new PowerPointHandler(_pptxPath, editable: true);
    }

    // ==================== Bug 1: Doughnut chart JSON serialization crash ====================

    [Fact]
    public void Bug1_DoughnutChart_JsonSerialization_ShouldNotCrash()
    {
        // Setup: create a doughnut chart (which has holeSize stored as byte)
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "doughnut",
            ["title"] = "Sales Distribution",
            ["data"] = "Revenue:30,25,20,15,10"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("chartType");
        node.Format["chartType"].Should().Be("doughnut");

        // Bug: holeSize is stored as System.Byte in Format dictionary.
        // AppJsonContext doesn't have [JsonSerializable(typeof(byte))], so
        // FormatNode with JSON output crashes with:
        // "Metadata for type 'System.Byte' was not provided"
        var act = () => OutputFormatter.FormatNode(node, OutputFormat.Json);
        act.Should().NotThrow("doughnut chart JSON serialization should handle byte holeSize");
    }

    [Fact]
    public void Bug1_DoughnutChart_HoleSize_ShouldBeIntNotByte()
    {
        // Verify the root cause: holeSize is stored as byte instead of int
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "doughnut",
            ["title"] = "Doughnut",
            ["data"] = "Values:40,30,20,10"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();

        // holeSize should exist for doughnut charts
        if (node.Format.ContainsKey("holeSize"))
        {
            // Bug: the value is System.Byte but should be int for JSON serialization compatibility
            var holeSize = node.Format["holeSize"];
            holeSize.Should().NotBeNull();
            // Should be int, not byte — byte crashes the JSON serializer
            holeSize.Should().BeOfType<int>(
                "holeSize should be stored as int, not byte, to be compatible with AppJsonContext");
        }
    }

    [Fact]
    public void Bug1_DoughnutChart_JsonSerialization_Persists()
    {
        // Verify JSON serialization works after reopen too
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "doughnut",
            ["title"] = "Persistent Doughnut",
            ["data"] = "Q:25,25,25,25"
        });

        Reopen();

        var node = _handler.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();

        var act = () => OutputFormatter.FormatNode(node, OutputFormat.Json);
        act.Should().NotThrow("doughnut chart JSON serialization should work after reopen");
    }

    // ==================== Bug 2: Trendline not visible at chart level Get ====================

    [Fact]
    public void Bug2_Trendline_ShouldBeVisibleAtChartLevel()
    {
        // Setup: create a line chart and add trendline
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "line",
            ["title"] = "Trend Analysis",
            ["data"] = "Revenue:10,20,30,40,50",
            ["trendline"] = "linear"
        });

        // Verify trendline is visible at series level (works)
        var series = _handler.Get("/slide[1]/chart[1]/series[1]");
        series.Should().NotBeNull();
        series.Format.Should().ContainKey("trendline", "series-level should show trendline");
        series.Format["trendline"].Should().Be("linear");

        // Bug: trendline is NOT visible at chart level
        // Agent sets trendline=linear but when it does get chart[1], trendline doesn't appear
        var chart = _handler.Get("/slide[1]/chart[1]");
        chart.Should().NotBeNull();
        chart.Format.Should().ContainKey("trendline",
            "chart-level Get should surface trendline for agent discoverability — " +
            "currently only visible at series[1] level which is agent-unfriendly");
    }

    [Fact]
    public void Bug2_Trendline_SetOnSeries_ShouldShowAtChartLevel()
    {
        // Even if trendline is set after chart creation via Set on series,
        // chart-level Get should still surface it
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "line",
            ["title"] = "Post-Set Trendline",
            ["data"] = "Sales:5,15,25,35,45"
        });

        // Set trendline on series
        _handler.Set("/slide[1]/chart[1]/series[1]", new() { ["trendline"] = "linear" });

        // Verify at series level (should work)
        var series = _handler.Get("/slide[1]/chart[1]/series[1]");
        series.Format.Should().ContainKey("trendline");

        // Bug: chart-level should also show trendline info
        var chart = _handler.Get("/slide[1]/chart[1]");
        chart.Format.Should().ContainKey("trendline",
            "chart-level Get should aggregate trendline from series for agent visibility");
    }

    // ==================== Bug 3: PRESET=corporate uppercase false positive ====================

    [Fact]
    public void Bug3_Preset_UppercaseKey_ShouldApplyOrReportUnsupported()
    {
        // Setup: create a chart
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["title"] = "Business Report",
            ["data"] = "Q1:40,Q2:55,Q3:70,Q4:85"
        });

        // The corporate preset sets legend=right, title.bold=true, title.color=44546A, etc.
        // Set with UPPERCASE key — should work since ChartSetter does key.ToLowerInvariant()
        var unsupported = _handler.Set("/slide[1]/chart[1]", new()
        {
            ["PRESET"] = "corporate"
        });

        // If preset was applied, unsupported should not contain "PRESET" itself
        // (sub-properties may appear if chart lacks certain elements, but the key should be consumed)
        unsupported.Should().NotContain("PRESET",
            "PRESET key should be recognized (case-insensitive) by ChartSetter");

        // Verify that the preset actually took effect by checking at least one property
        var chart = _handler.Get("/slide[1]/chart[1]");
        chart.Should().NotBeNull();

        // Corporate preset sets legend=right
        chart.Format.Should().ContainKey("legend",
            "corporate preset should set legend position to 'right'");
        chart.Format["legend"].Should().Be("right",
            "corporate preset legend should be 'right'");
    }

    [Fact]
    public void Bug3_Preset_LowercaseKey_ShouldApply()
    {
        // Sanity: lowercase preset=corporate should definitely work
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["title"] = "Baseline",
            ["data"] = "A:10,B:20,C:30"
        });

        var unsupported = _handler.Set("/slide[1]/chart[1]", new()
        {
            ["preset"] = "corporate"
        });

        unsupported.Should().NotContain("preset");

        var chart = _handler.Get("/slide[1]/chart[1]");
        chart.Format.Should().ContainKey("legend");
        chart.Format["legend"].Should().Be("right",
            "corporate preset should apply legend=right");

        // Verify title styling from preset
        chart.Format.Should().ContainKey("title");
    }

    [Fact]
    public void Bug3_Preset_MixedCaseValue_ShouldApply()
    {
        // Verify value is also case-insensitive: Corporate, CORPORATE
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "bar",
            ["title"] = "Mixed Case Test",
            ["data"] = "X:10,20,30"
        });

        // Mixed case value
        var unsupported = _handler.Set("/slide[1]/chart[1]", new()
        {
            ["PRESET"] = "Corporate"
        });

        unsupported.Should().NotContain("PRESET");

        var chart = _handler.Get("/slide[1]/chart[1]");
        chart.Format.Should().ContainKey("legend");
        chart.Format["legend"].Should().Be("right",
            "Corporate (mixed case value) should resolve to corporate preset");
    }
}
