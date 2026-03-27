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
/// Round 11 schema order bug fixes:
/// 1. series.shadow + series.outline effectLst/ln ordering in spPr
/// 2. majorTickMark / minorTickMark / tickLabelPos position in axis
/// 3. LineChart smooth position (before axId)
/// 4. Scatter trendline position (before xVal/yVal)
/// </summary>
public class UserInterviewRound11Tests : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _xlsxPath;
    private PowerPointHandler _pptx;
    private ExcelHandler _excel;
    private readonly ITestOutputHelper _output;

    public UserInterviewRound11Tests(ITestOutputHelper output)
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

    // ==================== Bug 1: series.shadow then series.outline — effectLst after ln ====================

    [Fact]
    public void Pptx_SeriesShadow_ThenOutline_SchemaOrderValid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Shadow+Outline",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        // Set shadow first, then outline — effectLst must come after ln in spPr
        _pptx.Set(chartPath, new() { ["series.shadow"] = "000000-3-315-2-50" });
        _pptx.Set(chartPath, new() { ["series.outline"] = "FF0000-1" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("series spPr must have ln before effectLst per DrawingML schema");
    }

    [Fact]
    public void Pptx_SeriesShadow_ThenOutline_ElementOrder_LnBeforeEffectLst()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Shadow+Outline Order",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        _pptx.Set(chartPath, new() { ["series.shadow"] = "000000-3-315-2-50" });
        _pptx.Set(chartPath, new() { ["series.outline"] = "FF0000-1" });

        // Verify element order via reflection: ln must come before effectLst
        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        var chartParts = doc.PresentationPart!.SlideParts.SelectMany(s => s.ChartParts);
        foreach (var cp in chartParts)
        {
            var serElements = cp.ChartSpace.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser");
            foreach (var ser in serElements)
            {
                var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                if (spPr == null) continue;
                var children = spPr.ChildElements.Select(c => c.LocalName).ToList();
                var lnIdx = children.IndexOf("ln");
                var effIdx = children.IndexOf("effectLst");
                if (lnIdx >= 0 && effIdx >= 0)
                    lnIdx.Should().BeLessThan(effIdx, "ln must come before effectLst in spPr");
            }
        }
    }

    [Fact]
    public void Pptx_PerSeriesDotted_ShadowThenOutline_SchemaOrderValid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Dotted Shadow+Outline",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        // Per-series dotted key shadow then outline
        _pptx.Set(chartPath, new() { ["series1.shadow"] = "000000-3-315-2-50" });
        _pptx.Set(chartPath, new() { ["series1.outline"] = "0000FF" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("per-series spPr must have ln before effectLst");
    }

    // ==================== Bug 2: majorTickMark / minorTickMark / tickLabelPos position ====================

    [Fact]
    public void Pptx_MajorTickMark_SchemaOrderValid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Tick Marks",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        _pptx.Set(chartPath, new() { ["majorTickMark"] = "outside" });
        _pptx.Set(chartPath, new() { ["minorTickMark"] = "inside" });
        _pptx.Set(chartPath, new() { ["tickLabelPos"] = "low" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("majorTickMark/minorTickMark/tickLabelPos must be in schema order within axis");
    }

    [Fact]
    public void Excel_MajorTickMark_SchemaOrderValid()
    {
        _excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Cat" });
        _excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "10" });
        _excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "20" });
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Tick Marks",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        _excel.Set(chartPath, new() { ["majorTickMark"] = "cross" });
        _excel.Set(chartPath, new() { ["minorTickMark"] = "outside" });
        _excel.Set(chartPath, new() { ["tickLabelPos"] = "high" });

        _excel.Dispose();
        var errors = ValidateXlsx(_xlsxPath);
        errors.Should().BeEmpty("majorTickMark/minorTickMark/tickLabelPos must be in schema order within axis");
    }

    [Fact]
    public void Pptx_TickMark_ElementOrder_InAxis()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Tick Order",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        // Set tick properties — these should be placed before spPr/crossAx in axis
        _pptx.Set(chartPath, new() { ["majorTickMark"] = "outside" });
        _pptx.Set(chartPath, new() { ["minorTickMark"] = "inside" });
        _pptx.Set(chartPath, new() { ["tickLabelPos"] = "nextTo" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        foreach (var cp in doc.PresentationPart!.SlideParts.SelectMany(s => s.ChartParts))
        {
            foreach (var ax in cp.ChartSpace.Descendants<C.ValueAxis>())
            {
                var children = ax.ChildElements.Select(c => c.LocalName).ToList();
                var mjIdx = children.IndexOf("majorTickMark");
                var mnIdx = children.IndexOf("minorTickMark");
                var tlIdx = children.IndexOf("tickLblPos");
                var crossIdx = children.IndexOf("crossAx");
                var spPrIdx = children.IndexOf("spPr");

                if (mjIdx >= 0 && mnIdx >= 0)
                    mjIdx.Should().BeLessThan(mnIdx, "majorTickMark must come before minorTickMark");
                if (mnIdx >= 0 && tlIdx >= 0)
                    mnIdx.Should().BeLessThan(tlIdx, "minorTickMark must come before tickLblPos");
                if (tlIdx >= 0 && crossIdx >= 0)
                    tlIdx.Should().BeLessThan(crossIdx, "tickLblPos must come before crossAx");
                if (tlIdx >= 0 && spPrIdx >= 0)
                    tlIdx.Should().BeLessThan(spPrIdx, "tickLblPos must come before spPr");
            }
        }
    }

    // ==================== Bug 3: LineChart smooth before axId ====================

    [Fact]
    public void Pptx_LineChart_Smooth_SchemaOrderValid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Smooth Line",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        _pptx.Set(chartPath, new() { ["smooth"] = "true" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("smooth must come before axId in CT_LineChart");
    }

    [Fact]
    public void Pptx_LineChart_Smooth_BeforeAxId()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Smooth Order",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });
        _pptx.Set(chartPath, new() { ["smooth"] = "true" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        foreach (var cp in doc.PresentationPart!.SlideParts.SelectMany(s => s.ChartParts))
        {
            foreach (var lc in cp.ChartSpace.Descendants<C.LineChart>())
            {
                var children = lc.ChildElements.Select(c => c.LocalName).ToList();
                var smoothIdx = children.IndexOf("smooth");
                var axIdIdx = children.IndexOf("axId");
                if (smoothIdx >= 0 && axIdIdx >= 0)
                    smoothIdx.Should().BeLessThan(axIdIdx, "smooth must come before axId in LineChart");
            }
        }
    }

    [Fact]
    public void Excel_LineChart_Smooth_SchemaOrderValid()
    {
        _excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Cat" });
        _excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "10" });
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Smooth Line",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        _excel.Set(chartPath, new() { ["smooth"] = "true" });

        _excel.Dispose();
        var errors = ValidateXlsx(_xlsxPath);
        errors.Should().BeEmpty("smooth must come before axId in CT_LineChart");
    }

    // ==================== Bug 4: Scatter trendline before xVal/yVal ====================

    [Fact]
    public void Pptx_ScatterTrendline_SchemaOrderValid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["title"] = "Scatter Trendline",
            ["data"] = "S1:10,20,30",
            ["categories"] = "1,2,3"
        });
        _pptx.Set(chartPath, new() { ["trendline"] = "linear" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("trendline must come before xVal/yVal in scatter series");
    }

    [Fact]
    public void Pptx_ScatterTrendline_BeforeXValYVal()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["title"] = "Scatter Trendline Order",
            ["data"] = "S1:10,20,30",
            ["categories"] = "1,2,3"
        });
        _pptx.Set(chartPath, new() { ["trendline"] = "linear" });

        _pptx.Dispose();
        using var doc = PresentationDocument.Open(_pptxPath, false);
        foreach (var cp in doc.PresentationPart!.SlideParts.SelectMany(s => s.ChartParts))
        {
            foreach (var ser in cp.ChartSpace.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser" && e.Parent is C.ScatterChart))
            {
                var children = ser.ChildElements.Select(c => c.LocalName).ToList();
                var trendIdx = children.IndexOf("trendline");
                var xValIdx = children.IndexOf("xVal");
                var yValIdx = children.IndexOf("yVal");
                if (trendIdx >= 0 && xValIdx >= 0)
                    trendIdx.Should().BeLessThan(xValIdx, "trendline must come before xVal");
                if (trendIdx >= 0 && yValIdx >= 0)
                    trendIdx.Should().BeLessThan(yValIdx, "trendline must come before yVal");
            }
        }
    }

    [Fact]
    public void Pptx_PerSeriesDotted_ScatterTrendline_SchemaOrderValid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["title"] = "Per-Series Scatter Trendline",
            ["data"] = "S1:10,20,30",
            ["categories"] = "1,2,3"
        });
        // Use per-series dotted key
        _pptx.Set(chartPath, new() { ["series1.trendline"] = "polynomial" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("per-series trendline must come before xVal/yVal in scatter series");
    }

    [Fact]
    public void Pptx_PerSeriesDotted_Smooth_SchemaOrderValid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["title"] = "Per-Series Scatter Smooth",
            ["data"] = "S1:10,20,30",
            ["categories"] = "1,2,3"
        });
        // Use per-series dotted key
        _pptx.Set(chartPath, new() { ["series1.smooth"] = "true" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("per-series smooth must be at correct position in scatter series");
    }

    // ==================== Combined: multiple properties stress test ====================

    [Fact]
    public void Pptx_AllFixedProperties_Combined_SchemaOrderValid()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Combined Stress Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });
        // Set all properties that had schema order bugs
        _pptx.Set(chartPath, new() { ["series.shadow"] = "000000-3-315-2-50" });
        _pptx.Set(chartPath, new() { ["series.outline"] = "FF0000-1" });
        _pptx.Set(chartPath, new() { ["majorTickMark"] = "outside" });
        _pptx.Set(chartPath, new() { ["minorTickMark"] = "inside" });
        _pptx.Set(chartPath, new() { ["tickLabelPos"] = "low" });
        _pptx.Set(chartPath, new() { ["smooth"] = "true" });

        _pptx.Dispose();
        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("all fixed properties combined must produce valid OOXML");
    }
}
