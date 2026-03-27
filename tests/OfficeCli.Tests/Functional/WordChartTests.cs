// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Word chart tests: docPr id uniqueness, per-series smooth readback, multi-chart lifecycle.
/// </summary>
public class WordChartTests : IDisposable
{
    private readonly string _docxPath;
    private WordHandler _word;

    public WordChartTests()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_docxPath);
        _word = new WordHandler(_docxPath, editable: true);
    }

    public void Dispose()
    {
        _word.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    private void Reopen() { _word.Dispose(); _word = new WordHandler(_docxPath, editable: true); }

    // ==================== Bug 1: Multiple charts have unique docPr ids ====================

    [Fact]
    public void Add_MultipleCharts_DocPrIdsAreUnique()
    {
        // Add two charts
        _word.Add("/", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Chart A", ["data"] = "S1:1,2,3"
        });
        _word.Add("/", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Chart B", ["data"] = "S2:4,5,6"
        });

        // Both charts should be accessible
        var c1 = _word.Get("/chart[1]");
        var c2 = _word.Get("/chart[2]");
        c1.Type.Should().Be("chart");
        c2.Type.Should().Be("chart");

        // Verify docPr ids are unique by checking the raw XML
        var body = GetDocumentBody();
        var docProps = body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().ToList();
        var ids = docProps.Select(dp => dp.Id?.Value).Where(id => id.HasValue).Select(id => id!.Value).ToList();
        ids.Should().OnlyHaveUniqueItems("each wp:docPr must have a unique id");
    }

    [Fact]
    public void Add_ThreeCharts_AllHaveDistinctDocPrIds()
    {
        for (int i = 1; i <= 3; i++)
        {
            _word.Add("/", "chart", null, new()
            {
                ["chartType"] = "column", ["title"] = $"Chart {i}", ["data"] = $"S{i}:1,2,3"
            });
        }

        var body = GetDocumentBody();
        var ids = body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
            .Select(dp => dp.Id?.Value)
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .ToList();

        ids.Count.Should().BeGreaterOrEqualTo(3);
        ids.Should().OnlyHaveUniqueItems("docPr ids must be unique across all charts");
    }

    [Fact]
    public void Add_MultipleCharts_Reopen_DocPrIdsStillUnique()
    {
        _word.Add("/", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Chart A", ["data"] = "S1:10,20"
        });
        _word.Add("/", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "Chart B", ["data"] = "S2:30,40"
        });

        Reopen();

        // Add a third chart after reopen
        _word.Add("/", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Chart C", ["data"] = "S3:50,60"
        });

        var body = GetDocumentBody();
        var ids = body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
            .Select(dp => dp.Id?.Value)
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .ToList();

        ids.Count.Should().BeGreaterOrEqualTo(3);
        ids.Should().OnlyHaveUniqueItems("docPr ids must remain unique after reopen");
    }

    // ==================== Issue 3: Per-series smooth readback ====================

    [Fact]
    public void Set_SeriesSmooth_IsReadBack()
    {
        _word.Add("/", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Smooth Test", ["data"] = "S1:1,2,3;S2:4,5,6"
        });

        // Set smooth on series 1
        _word.Set("/chart[1]/series[1]", new() { ["smooth"] = "true" });

        var s1 = _word.Get("/chart[1]/series[1]");
        s1.Format.Should().ContainKey("smooth");
        ((string)s1.Format["smooth"]).Should().Be("true");
    }

    [Fact]
    public void Set_SeriesSmooth_Persists_AfterReopen()
    {
        _word.Add("/", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Smooth Persist", ["data"] = "S1:1,2,3"
        });
        _word.Set("/chart[1]/series[1]", new() { ["smooth"] = "true" });

        Reopen();

        var s1 = _word.Get("/chart[1]/series[1]");
        s1.Format.Should().ContainKey("smooth");
        ((string)s1.Format["smooth"]).Should().Be("true");
    }

    [Fact]
    public void Set_ChartLevelSmooth_IsReadBack()
    {
        _word.Add("/", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Chart Smooth", ["data"] = "S1:1,2,3"
        });

        _word.Set("/chart[1]", new() { ["smooth"] = "true" });

        var chart = _word.Get("/chart[1]");
        chart.Format.Should().ContainKey("smooth");
        ((string)chart.Format["smooth"]).Should().Be("true");
    }

    // ==================== Helper ====================

    private DocumentFormat.OpenXml.Wordprocessing.Body GetDocumentBody()
    {
        // Access internal document via reflection to inspect raw XML
        var docField = typeof(WordHandler).GetField("_doc",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        var doc = (DocumentFormat.OpenXml.Packaging.WordprocessingDocument)docField!.GetValue(_word)!;
        return doc.MainDocumentPart!.Document!.Body!;
    }
}
