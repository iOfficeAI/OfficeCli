// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Final round bug fixes:
/// 1. scatter series.outline spPr position — use GetOrCreateSeriesShapeProperties
/// 2. gridlines=none Get readback — explicit "false" when axis exists but no gridlines
/// 3. Excel waterfall chart type readback — should read "waterfall" not "column_stacked"
/// </summary>
public class UserInterviewFinalTests : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _xlsxPath;
    private PowerPointHandler _pptx;
    private ExcelHandler _excel;
    private readonly ITestOutputHelper _output;

    public UserInterviewFinalTests(ITestOutputHelper output)
    {
        _output = output;
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_xlsxPath);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _pptx.Dispose();
        _excel.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    private void ReopenPptx() { _pptx.Dispose(); _pptx = new PowerPointHandler(_pptxPath, editable: true); }
    private void ReopenExcel() { _excel.Dispose(); _excel = new ExcelHandler(_xlsxPath, editable: true); }

    private List<ValidationErrorInfo> ValidatePptx(string path)
    {
        using var doc = PresentationDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[PPTX] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        return errors;
    }

    private List<ValidationErrorInfo> ValidateXlsx(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[XLSX] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        return errors;
    }

    // ==================== Bug 1: scatter series.outline spPr position ====================

    [Fact]
    public void Pptx_ScatterChart_SeriesOutline_SpPrPosition_Valid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["title"] = "Scatter Outline Test",
            ["data"] = "S1:1,2,3",
            ["categories"] = "10,20,30"
        });

        // Set series.outline — must not create mispositioned spPr
        _pptx.Set(chartPath, new() { ["series.outline"] = "FF0000:1.5" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("scatter series spPr must be in correct schema position");
    }

    [Fact]
    public void Pptx_ScatterChart_SeriesOutline_SpPrBeforeMarker()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["title"] = "Scatter SpPr Order",
            ["data"] = "S1:1,2,3",
            ["categories"] = "10,20,30"
        });
        _pptx.Set(chartPath, new() { ["series.outline"] = "0000FF:2" });

        // Verify spPr is before marker in the scatter series XML
        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        var chartPart = doc.PresentationPart!.SlideParts.First()
            .ChartParts.First();
        var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var plotArea = chart.GetFirstChild<C.PlotArea>()!;
        var scatterChart = plotArea.GetFirstChild<C.ScatterChart>()!;
        var ser = scatterChart.Elements<C.ScatterChartSeries>().First();

        var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
        spPr.Should().NotBeNull("series should have spPr");

        var marker = ser.GetFirstChild<C.Marker>();
        if (marker != null && spPr != null)
        {
            // spPr should come before marker in child order
            var children = ser.ChildElements.ToList();
            var spPrIdx = children.IndexOf(spPr);
            var markerIdx = children.IndexOf(marker);
            spPrIdx.Should().BeLessThan(markerIdx, "spPr must come before marker in CT_ScatterSer");
        }
    }

    // ==================== Bug 2: gridlines=none Get readback ====================

    [Fact]
    public void Pptx_Chart_GridlinesNone_ReturnsExplicitFalse()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Gridlines Test",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });

        // Remove gridlines
        _pptx.Set(chartPath, new() { ["gridlines"] = "none" });

        var node = _pptx.Get(chartPath);
        node.Format.Should().ContainKey("gridlines");
        node.Format["gridlines"].Should().Be("false",
            "when gridlines are removed, Get should return gridlines=false for UX clarity");
    }

    [Fact]
    public void Pptx_Chart_GridlinesTrue_ReturnsTrue()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Gridlines Present",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C",
            ["gridlines"] = "true"
        });

        var node = _pptx.Get(chartPath);
        node.Format.Should().ContainKey("gridlines");
        node.Format["gridlines"].Should().Be("true");
    }

    [Fact]
    public void Pptx_Chart_GridlinesNone_Persistence()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Gridlines Persist",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C",
            ["gridlines"] = "true"
        });

        // Set gridlines then remove
        _pptx.Set(chartPath, new() { ["gridlines"] = "none" });
        ReopenPptx();

        var node = _pptx.Get(chartPath);
        node.Format.Should().ContainKey("gridlines");
        node.Format["gridlines"].Should().Be("false");
    }

    // ==================== Bug 3: Excel waterfall chart type readback ====================

    [Fact]
    public void Excel_Waterfall_ChartType_ReadsAsWaterfall()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Waterfall Test",
            ["data"] = "Cashflow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("chartType");
        node.Format["chartType"].Should().Be("waterfall",
            "Excel waterfall chart should read back as 'waterfall', not 'column_stacked'");
    }

    [Fact]
    public void Pptx_Waterfall_ChartType_ReadsAsWaterfall()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Waterfall PPTX",
            ["data"] = "Cashflow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit"
        });

        var node = _pptx.Get(chartPath);
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("chartType");
        node.Format["chartType"].Should().Be("waterfall",
            "PPTX waterfall chart should read back as 'waterfall'");
    }

    [Fact]
    public void Excel_Waterfall_ChartType_Persistence()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Waterfall Persist",
            ["data"] = "Cashflow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit"
        });

        ReopenExcel();

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Should().NotBeNull();
        node.Format["chartType"].Should().Be("waterfall");
    }
}
