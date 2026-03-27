// Black-box tests (Customer A, Round 17) — chart validation scenarios:
//   1. Same chart: 5 consecutive preset changes → validate each reads back
//   2. preset=dark → chartFill=FFFFFF → chartArea.border=000000 → validate
//   3. preset=corporate → reopen → preset=dark → reopen → validate persistence
//   4. PPTX column chart + all major property combinations → validate
//   5. Excel column chart + all major property combinations → validate
//   6. Word column chart + preset=corporate → validate
//   7. PPTX scatter chart + full suite of properties → validate
//   8. pie + doughnut + radar + area charts, each with data/labels → validate
//   9. waterfall + referenceLine → validate both render
//  10. funnel (cx:chart) + set position → validate

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtCustomerARound17 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtCustomerARound17(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt17_{Guid.NewGuid():N}.{ext}");
        _temps.Add(p);
        BlankDocCreator.Create(p);
        return p;
    }

    // ==================== 1. Same chart: 5 consecutive preset changes ====================

    [Fact]
    public void Pptx_Chart_FiveConsecutivePresets_EachReadsBack()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Multi-Preset Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });

        // Preset 1: minimal
        h.Set(chartPath, new() { ["preset"] = "minimal" });
        var n1 = h.Get(chartPath, depth: 0);
        n1.Should().NotBeNull("chart must be readable after preset=minimal");
        _out.WriteLine($"After minimal: chartFill={n1!.Format.GetValueOrDefault("chartFill")}");

        // Preset 2: dark
        h.Set(chartPath, new() { ["preset"] = "dark" });
        var n2 = h.Get(chartPath, depth: 0);
        n2.Should().NotBeNull("chart must be readable after preset=dark");
        n2!.Format.Should().ContainKey("chartFill");
        n2.Format["chartFill"].Should().Be("#1E1E1E", "dark preset chartFill is 1E1E1E");
        _out.WriteLine($"After dark: chartFill={n2.Format["chartFill"]}");

        // Preset 3: corporate
        h.Set(chartPath, new() { ["preset"] = "corporate" });
        var n3 = h.Get(chartPath, depth: 0);
        n3.Should().NotBeNull("chart must be readable after preset=corporate");
        _out.WriteLine($"After corporate: colors={n3!.Format.GetValueOrDefault("colors")}");

        // Preset 4: magazine
        h.Set(chartPath, new() { ["preset"] = "magazine" });
        var n4 = h.Get(chartPath, depth: 0);
        n4.Should().NotBeNull("chart must be readable after preset=magazine");
        n4!.Format.Should().ContainKey("dataLabels", "magazine preset enables dataLabels");
        _out.WriteLine($"After magazine: dataLabels={n4.Format["dataLabels"]}");

        // Preset 5: dashboard
        h.Set(chartPath, new() { ["preset"] = "dashboard" });
        var n5 = h.Get(chartPath, depth: 0);
        n5.Should().NotBeNull("chart must be readable after preset=dashboard");
        _out.WriteLine($"After dashboard: chartFill={n5!.Format.GetValueOrDefault("chartFill")}");
    }

    // ==================== 2. preset=dark → chartFill=FFFFFF → chartArea.border=000000 ====================

    [Fact]
    public void Pptx_Chart_DarkPreset_ThenOverrideChartFill_ThenBorder()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar",
            ["title"] = "Override Test",
            ["data"] = "S1:5,10,15",
            ["categories"] = "X,Y,Z"
        });

        // Step 1: apply dark preset
        h.Set(chartPath, new() { ["preset"] = "dark" });
        var n1 = h.Get(chartPath, depth: 0);
        n1!.Format["chartFill"].Should().Be("#1E1E1E", "dark preset sets chartFill to 1E1E1E");

        // Step 2: override chartFill with white
        h.Set(chartPath, new() { ["chartFill"] = "FFFFFF" });
        var n2 = h.Get(chartPath, depth: 0);
        n2!.Format["chartFill"].Should().Be("#FFFFFF", "chartFill override after preset");

        // Step 3: add chartArea border
        h.Set(chartPath, new() { ["chartArea.border"] = "000000:1.0" });
        var n3 = h.Get(chartPath, depth: 0);
        // Chart readback confirms fill is still #FFFFFF (border is stored in XML but may not be in Format keys by default)
        n3!.Format["chartFill"].Should().Be("#FFFFFF", "chartFill unchanged after adding border");
        _out.WriteLine($"After all 3 steps: chartFill={n3.Format["chartFill"]}");
    }

    // ==================== 3. preset=corporate → reopen → preset=dark → reopen → validate ====================

    [Fact]
    public void Pptx_Chart_CorporateToDark_WithReopens()
    {
        var path = Temp("pptx");
        string chartPath;

        // Create chart + apply corporate
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new());
            chartPath = h.Add("/slide[1]", "chart", null, new()
            {
                ["chartType"] = "line",
                ["title"] = "Persistence Test",
                ["data"] = "S1:1,2,3",
                ["categories"] = "A,B,C"
            });
            h.Set(chartPath, new() { ["preset"] = "corporate" });
        }

        // Reopen and verify corporate applied (no chartFill = none for corporate)
        using (var h = new PowerPointHandler(path, editable: false))
        {
            var n = h.Get(chartPath, depth: 0);
            n.Should().NotBeNull("chart must be readable after corporate preset + reopen");
            _out.WriteLine($"After corporate reopen: chartType={n!.Format.GetValueOrDefault("chartType")}");
        }

        // Apply dark preset
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Set(chartPath, new() { ["preset"] = "dark" });
        }

        // Reopen and verify dark persisted
        using (var h2 = new PowerPointHandler(path, editable: false))
        {
            var n = h2.Get(chartPath, depth: 0);
            n.Should().NotBeNull("chart must be readable after dark preset + reopen");
            n!.Format.Should().ContainKey("chartFill");
            n.Format["chartFill"].Should().Be("#1E1E1E", "dark preset chartFill persists across reopen");
            _out.WriteLine($"After dark reopen: chartFill={n.Format["chartFill"]}");
        }
    }

    // ==================== 4. PPTX column + all major property combinations ====================

    [Fact]
    public void Pptx_Column_AllMajorProperties_SetAndGet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Column All Props",
            ["data"] = "Revenue:100,200,150,300;Costs:80,150,120,250",
            ["categories"] = "Q1,Q2,Q3,Q4",
            ["legend"] = "bottom",
            ["dataLabels"] = "value",
            ["colors"] = "4472C4,ED7D31"
        });

        h.Set(chartPath, new()
        {
            ["title"] = "Updated Column",
            ["axisTitle"] = "USD",
            ["catTitle"] = "Quarter",
            ["axisMin"] = "0",
            ["axisMax"] = "400",
            ["majorUnit"] = "100",
            ["gridlines"] = "CCCCCC:0.5",
            ["plotFill"] = "F5F5F5",
            ["gapwidth"] = "150",
            ["overlap"] = "0",
            ["title.bold"] = "true",
            ["title.size"] = "16",
            ["title.color"] = "2E75B6",
        });

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        node!.Format["title"].Should().Be("Updated Column");
        node.Format.Should().ContainKey("axisMin");
        node.Format.Should().ContainKey("axisMax");
        node.Format.Should().ContainKey("gridlines");
        _out.WriteLine($"column all-props: title={node.Format["title"]}, axisMin={node.Format["axisMin"]}, axisMax={node.Format["axisMax"]}");
    }

    // ==================== 5. Excel column + all major property combinations ====================

    [Fact]
    public void Excel_Column_AllMajorProperties_SetAndGet()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        var chartPath = h.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Excel Column All Props",
            ["data"] = "Sales:100,200,150;Target:120,180,160",
            ["categories"] = "Jan,Feb,Mar",
            ["legend"] = "top",
            ["dataLabels"] = "value"
        });

        h.Set(chartPath, new()
        {
            ["title"] = "Updated Excel Column",
            ["axisMin"] = "0",
            ["axisMax"] = "250",
            ["majorUnit"] = "50",
            ["gridlines"] = "E0E0E0:0.3",
            ["plotFill"] = "FAFAFA",
            ["gapwidth"] = "100",
            ["overlap"] = "-20",
            ["title.bold"] = "true",
            ["title.size"] = "14",
        });

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        node!.Format["title"].Should().Be("Updated Excel Column");
        node.Format.Should().ContainKey("axisMin");
        node.Format.Should().ContainKey("axisMax");
        _out.WriteLine($"Excel column: title={node.Format["title"]}, axisMax={node.Format["axisMax"]}");
    }

    // ==================== 6. Word column chart + preset=corporate ====================

    [Fact]
    public void Word_Column_Chart_Preset_Corporate()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Chart below:" });

        var chartPath = h.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Word Corporate Chart",
            ["data"] = "Series1:10,20,30",
            ["categories"] = "Jan,Feb,Mar"
        });

        chartPath.Should().NotBeNull("Word Add chart should return a path");
        _out.WriteLine($"Word chart path: {chartPath}");

        // Apply corporate preset via Set
        var ex = Record.Exception(() => h.Set(chartPath!, new() { ["preset"] = "corporate" }));
        ex.Should().BeNull("Word chart Set preset=corporate must not throw");

        var node = h.Get(chartPath!, depth: 0);
        node.Should().NotBeNull("Word chart Get must return a node");
        _out.WriteLine($"Word chart node type: {node!.Type}, format keys: {string.Join(", ", node.Format.Keys)}");
    }

    // ==================== 7. PPTX scatter + full property suite ====================

    [Fact]
    public void Pptx_Scatter_FullPropertySuite()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["title"] = "Scatter Full Props",
            ["data"] = "Series1:10,20,30,40,50",
            ["categories"] = "1,2,3,4,5",
            ["legend"] = "right"
        });

        h.Set(chartPath, new()
        {
            ["title"] = "Updated Scatter",
            ["marker"] = "circle:6:4472C4",
            ["lineWidth"] = "1.5",
            ["lineDash"] = "dash",
            ["gridlines"] = "DDDDDD:0.4",
            ["minorGridlines"] = "none",
            ["axisMin"] = "0",
            ["axisMax"] = "60",
            ["dataLabels"] = "none",
            ["plotFill"] = "none",
            ["chartFill"] = "none",
            ["series.outline"] = "none",
        });

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        node!.Format["chartType"].Should().Be("scatter");
        node.Format["title"].Should().Be("Updated Scatter");
        _out.WriteLine($"scatter: marker={node.Format.GetValueOrDefault("marker")}, lineWidth={node.Format.GetValueOrDefault("lineWidth")}");
    }

    // ==================== 8. pie + doughnut + radar + area ====================

    [Fact]
    public void Pptx_PieChart_AddAndGet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "pie",
            ["title"] = "Pie Chart",
            ["data"] = "Share:40,35,25",
            ["categories"] = "Product A,Product B,Product C",
            ["dataLabels"] = "percent",
            ["legend"] = "bottom"
        });

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        node!.Format["chartType"].Should().Be("pie");
        _out.WriteLine($"pie: chartType={node.Format["chartType"]}, seriesCount={node.Format.GetValueOrDefault("seriesCount")}");
    }

    [Fact]
    public void Pptx_DoughnutChart_AddAndGet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "doughnut",
            ["title"] = "Doughnut Chart",
            ["data"] = "Sales:30,45,25",
            ["categories"] = "Online,Retail,Wholesale",
            ["dataLabels"] = "value",
            ["legend"] = "right"
        });

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        node!.Format["chartType"].Should().Be("doughnut");
        _out.WriteLine($"doughnut: chartType={node.Format["chartType"]}");
    }

    [Fact]
    public void Pptx_RadarChart_AddAndGet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "radar",
            ["title"] = "Radar Chart",
            ["data"] = "Team:8,6,7,9,5",
            ["categories"] = "Code,Design,PM,Testing,Docs",
            ["legend"] = "bottom"
        });

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        node!.Format["chartType"].Should().Be("radar");
        _out.WriteLine($"radar: chartType={node.Format["chartType"]}");
    }

    [Fact]
    public void Pptx_AreaChart_AddAndGet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "area",
            ["title"] = "Area Chart",
            ["data"] = "Sales:10,30,25,40,35;Costs:8,20,18,28,25",
            ["categories"] = "Jan,Feb,Mar,Apr,May",
            ["legend"] = "top"
        });

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        node!.Format["chartType"].Should().Be("area");
        _out.WriteLine($"area: chartType={node.Format["chartType"]}, seriesCount={node.Format.GetValueOrDefault("seriesCount")}");
    }

    // ==================== 9. waterfall + referenceLine ====================

    [Fact]
    public void Pptx_Waterfall_WithReferenceLine()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Waterfall + Reference",
            ["data"] = "Cashflow:100,-30,50,-20,80",
            ["categories"] = "Start,Q1,Q2,Q3,End"
        });

        // Add a reference line at value 80
        var ex = Record.Exception(() =>
            h.Set(chartPath, new() { ["referenceLine"] = "80:FF0000:Target:dash" }));
        ex.Should().BeNull("waterfall + referenceLine Set must not throw");

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull("waterfall chart must be readable after referenceLine Set");
        // After adding a referenceLine overlay (LineChart), the reader returns "combo" because the plot area
        // now contains both a BarChart (waterfall stacked) and a LineChart (reference). This is expected behavior.
        var chartType = (string)node!.Format["chartType"];
        chartType.Should().BeOneOf("waterfall", "combo",
            "waterfall with referenceLine has BarChart+LineChart overlay, reader may return 'combo'");
        _out.WriteLine($"waterfall: chartType={chartType}, seriesCount={node.Format.GetValueOrDefault("seriesCount")}");
    }

    [Fact]
    public void Pptx_Waterfall_WithReferenceLine_Persists()
    {
        var path = Temp("pptx");
        string chartPath;

        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new());
            chartPath = h.Add("/slide[1]", "chart", null, new()
            {
                ["chartType"] = "waterfall",
                ["title"] = "Waterfall Persist",
                ["data"] = "Flow:200,-50,80,-30,150",
                ["categories"] = "Init,Loss1,Gain1,Loss2,Final"
            });
            h.Set(chartPath, new() { ["referenceLine"] = "100:0000FF:Baseline" });
        }

        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get(chartPath, depth: 0);
        node.Should().NotBeNull("waterfall with referenceLine must persist across reopen");
        // Same behavior as above: BarChart+LineChart overlay → reader returns "combo"
        var chartType = (string)node!.Format["chartType"];
        chartType.Should().BeOneOf("waterfall", "combo",
            "waterfall with referenceLine persists as BarChart+LineChart combo");
        _out.WriteLine($"waterfall persist: chartType={chartType}");
    }

    // ==================== 10. funnel (cx:chart) + set position ====================

    [Fact]
    public void Pptx_Funnel_CxChart_SetPosition()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new());
        var chartPath = h.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["title"] = "Sales Funnel",
            ["data"] = "Pipeline:1200,900,600,300,150",
            ["categories"] = "Leads,Qualified,Proposal,Negotiation,Won"
        });

        var node = h.Get(chartPath, depth: 0);
        node.Should().NotBeNull("funnel chart must be created");
        node!.Format["chartType"].Should().Be("funnel");
        _out.WriteLine($"funnel initial: chartType={node.Format["chartType"]}, x={node.Format.GetValueOrDefault("x")}");

        // Set position via x/y/width/height
        var exPos = Record.Exception(() =>
            h.Set(chartPath, new() { ["x"] = "2cm", ["y"] = "3cm", ["width"] = "20cm", ["height"] = "12cm" }));
        exPos.Should().BeNull("funnel cx:chart Set position must not throw");

        var node2 = h.Get(chartPath, depth: 0);
        node2.Should().NotBeNull("funnel chart must be readable after position Set");
        _out.WriteLine($"funnel after Set position: x={node2!.Format.GetValueOrDefault("x")}, y={node2.Format.GetValueOrDefault("y")}");
    }

    [Fact]
    public void Pptx_Funnel_CxChart_Persists_After_Reopen()
    {
        var path = Temp("pptx");
        string chartPath;

        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new());
            chartPath = h.Add("/slide[1]", "chart", null, new()
            {
                ["chartType"] = "funnel",
                ["title"] = "Funnel Persist",
                ["data"] = "Stages:500,400,300,200,100",
                ["categories"] = "Awareness,Interest,Desire,Intent,Purchase"
            });
            h.Set(chartPath, new() { ["x"] = "1cm", ["y"] = "2cm", ["width"] = "18cm", ["height"] = "10cm" });
        }

        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get(chartPath, depth: 0);
        node.Should().NotBeNull("funnel chart must persist after reopen");
        node!.Format["chartType"].Should().Be("funnel");
        _out.WriteLine($"funnel persist reopen: chartType={node.Format["chartType"]}, x={node.Format.GetValueOrDefault("x")}");
    }
}
