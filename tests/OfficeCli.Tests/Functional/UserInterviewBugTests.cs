// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests proving bugs discovered during user interview trial session.
/// These tests are expected to FAIL until the bugs are fixed.
///
/// Bug 1 — Get returns /slide[N]/chart[M]/series[K] paths, but Set does not handle series sub-paths.
/// Bug 2 — Get("/slide[99]/chart[1]") returns null instead of throwing, causing NullReferenceException downstream.
/// Bug 3 — style=999 silently accepted; valid range is 1-48.
/// Bug 4 — overlap=500 silently accepted; valid range is -100 to 100.
/// </summary>
public class UserInterviewBugTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public UserInterviewBugTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler?.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen()
    {
        _handler?.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== Bug 1: series path Get/Set inconsistency ====================

    /// <summary>
    /// Get returns paths like /slide[1]/chart[1]/series[1], but Set on that path
    /// throws "Element not found" because Set only handles /slide[N]/chart[M].
    /// If Get returns a path, Set should accept it.
    /// </summary>
    [Fact]
    public void Bug1_SeriesPath_GetReturns_SetShouldAccept()
    {
        // Arrange: create a chart with data
        _handler.Add("/", "slide", null, new() { ["title"] = "Chart Test" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Act: Get the chart with depth to see series children
        var chartNode = _handler.Get("/slide[1]/chart[1]", depth: 2);
        chartNode.Should().NotBeNull("chart should exist");

        // Find a series child path
        var seriesChild = chartNode.Children?.FirstOrDefault(c => c.Type == "series");
        seriesChild.Should().NotBeNull("chart should have series children in Get output");
        var seriesPath = seriesChild!.Path;
        seriesPath.Should().Contain("series[", "Get should return series sub-paths");

        // Bug: Set on the series path should work, but throws "Element not found"
        var act = () => _handler.Set(seriesPath, new() { ["color"] = "FF0000" });
        act.Should().NotThrow("Set should accept paths that Get returns");
    }

    // ==================== Bug 2: null reference on non-existent slide ====================

    /// <summary>
    /// Get("/slide[99]/chart[1]") returns null instead of throwing ArgumentException.
    /// This causes NullReferenceException when the CLI tries to format the result.
    /// Should throw a clear error message instead.
    /// </summary>
    [Fact]
    public void Bug2_NonExistentSlide_ShouldThrowNotNull()
    {
        // Arrange: create one slide (so slide[99] doesn't exist)
        _handler.Add("/", "slide", null, new() { ["title"] = "Only Slide" });

        // Act & Assert: accessing a non-existent slide should throw, not return null
        var act = () => _handler.Get("/slide[99]/chart[1]");
        act.Should().Throw<ArgumentException>("non-existent slide should throw a clear error, not return null");
    }

    // ==================== Bug 3: style value range validation ====================

    /// <summary>
    /// Chart style values are 1-48, but Set accepts style=999 without error.
    /// The value is cast to byte, potentially producing garbage.
    /// </summary>
    [Fact]
    public void Bug3_StyleOutOfRange_ShouldThrow()
    {
        // Arrange
        _handler.Add("/", "slide", null, new() { ["title"] = "Chart Test" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30"
        });

        // Act & Assert: style=999 should be rejected
        var act = () => _handler.Set("/slide[1]/chart[1]", new() { ["style"] = "999" });
        act.Should().Throw<ArgumentException>("style value 999 is out of valid range 1-48");
    }

    // ==================== Bug 4: overlap value range validation ====================

    /// <summary>
    /// Overlap range is -100 to 100, but Set accepts overlap=500 without error.
    /// The value is cast to sbyte, potentially producing garbage.
    /// </summary>
    [Fact]
    public void Bug4_OverlapOutOfRange_ShouldThrow()
    {
        // Arrange
        _handler.Add("/", "slide", null, new() { ["title"] = "Chart Test" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30"
        });

        // Act & Assert: overlap=500 should be rejected
        var act = () => _handler.Set("/slide[1]/chart[1]", new() { ["overlap"] = "500" });
        act.Should().Throw<ArgumentException>("overlap value 500 is out of valid range -100 to 100");
    }
}
