// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

/// <summary>
/// Shared chart build/read/set logic used by PPTX, Excel, and Word handlers.
/// All methods operate on ChartPart / C.Chart / C.PlotArea — independent of host document type.
/// </summary>
internal static class ChartHelper
{
    // ==================== Parse Helpers ====================

    internal static (string kind, bool is3D, bool stacked, bool percentStacked) ParseChartType(string chartType)
    {
        var ct = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
        var is3D = ct.EndsWith("3d") || ct.Contains("3d");
        ct = ct.Replace("3d", "");

        var stacked = ct.Contains("stacked") && !ct.Contains("percent");
        var percentStacked = ct.Contains("percentstacked") || ct.Contains("pstacked");
        ct = ct.Replace("percentstacked", "").Replace("pstacked", "").Replace("stacked", "");

        var kind = ct switch
        {
            "bar" => "bar",
            "column" or "col" => "column",
            "line" => "line",
            "pie" => "pie",
            "doughnut" or "donut" => "doughnut",
            "area" => "area",
            "scatter" or "xy" => "scatter",
            _ => ct
        };

        return (kind, is3D, stacked, percentStacked);
    }

    internal static List<(string name, double[] values)> ParseSeriesData(Dictionary<string, string> properties)
    {
        var result = new List<(string name, double[] values)>();

        if (properties.TryGetValue("data", out var dataStr))
        {
            foreach (var seriesPart in dataStr.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var colonIdx = seriesPart.IndexOf(':');
                if (colonIdx < 0) continue;
                var name = seriesPart[..colonIdx].Trim();
                var valStr = seriesPart[(colonIdx + 1)..].Trim();
                if (string.IsNullOrEmpty(valStr))
                    throw new ArgumentException($"Series '{name}' has no data values. Expected format: 'Name:1,2,3'");
                var vals = valStr.Split(',')
                    .Select(v => double.Parse(v.Trim(), System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                result.Add((name, vals));
            }
            return result;
        }

        for (int i = 1; i <= 20; i++)
        {
            if (!properties.TryGetValue($"series{i}", out var seriesStr)) break;
            var colonIdx = seriesStr.IndexOf(':');
            if (colonIdx < 0)
            {
                var vals = seriesStr.Split(',').Select(v => double.Parse(v.Trim(), System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                result.Add(($"Series {i}", vals));
            }
            else
            {
                var name = seriesStr[..colonIdx].Trim();
                var vals = seriesStr[(colonIdx + 1)..].Split(',')
                    .Select(v => double.Parse(v.Trim(), System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                result.Add((name, vals));
            }
        }

        return result;
    }

    internal static string[]? ParseCategories(Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("categories", out var catStr)) return null;
        return catStr.Split(',').Select(c => c.Trim()).ToArray();
    }

    internal static string[]? ParseSeriesColors(Dictionary<string, string> properties)
    {
        if (properties.TryGetValue("colors", out var colorsStr))
            return colorsStr.Split(',').Select(c => c.Trim()).ToArray();
        return null;
    }

    // ==================== Build ChartSpace ====================

    internal static C.ChartSpace BuildChartSpace(
        string chartType,
        string? title,
        string[]? categories,
        List<(string name, double[] values)> seriesData,
        Dictionary<string, string> properties)
    {
        var (kind, is3D, stacked, percentStacked) = ParseChartType(chartType);

        var chartSpace = new C.ChartSpace();
        var chart = new C.Chart();

        if (!string.IsNullOrEmpty(title))
            chart.AppendChild(BuildChartTitle(title));

        if (categories == null && seriesData.Count > 0)
        {
            var maxLen = seriesData.Max(s => s.values.Length);
            categories = Enumerable.Range(1, maxLen).Select(i => i.ToString()).ToArray();
        }

        var plotArea = new C.PlotArea(new C.Layout());
        uint catAxisId = 1;
        uint valAxisId = 2;

        OpenXmlCompositeElement? chartElement;
        bool needsAxes = true;

        var colors = ParseSeriesColors(properties);

        switch (kind)
        {
            case "bar":
                chartElement = BuildBarChart(C.BarDirectionValues.Bar, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "column":
                chartElement = BuildBarChart(C.BarDirectionValues.Column, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "line":
                chartElement = BuildLineChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "area":
                chartElement = BuildAreaChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "pie":
                chartElement = BuildPieChart(categories, seriesData);
                needsAxes = false;
                break;
            case "doughnut":
                chartElement = BuildDoughnutChart(categories, seriesData);
                needsAxes = false;
                break;
            case "scatter":
                chartElement = BuildScatterChart(categories, seriesData, catAxisId, valAxisId);
                break;
            case "combo":
            {
                int splitAt = 1;
                if (properties.TryGetValue("combosplit", out var splitStr))
                    splitAt = int.Parse(splitStr);
                splitAt = Math.Min(splitAt, seriesData.Count);

                var barData = seriesData.Take(splitAt).ToList();
                var lineData = seriesData.Skip(splitAt).ToList();

                var comboBar = new C.BarChart(
                    new C.BarDirection { Val = C.BarDirectionValues.Column },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false }
                );
                for (int ci = 0; ci < barData.Count; ci++)
                {
                    var clr = colors != null && ci < colors.Length ? colors[ci] : DefaultSeriesColors[ci % DefaultSeriesColors.Length];
                    comboBar.AppendChild(BuildBarSeries((uint)ci, barData[ci].name, categories, barData[ci].values, clr));
                }
                comboBar.AppendChild(new C.AxisId { Val = catAxisId });
                comboBar.AppendChild(new C.AxisId { Val = valAxisId });
                plotArea.AppendChild(comboBar);

                if (lineData.Count > 0)
                {
                    var comboLine = new C.LineChart(
                        new C.Grouping { Val = C.GroupingValues.Standard },
                        new C.VaryColors { Val = false }
                    );
                    for (int ci = 0; ci < lineData.Count; ci++)
                    {
                        var sIdx = (uint)(splitAt + ci);
                        var cIdx = splitAt + ci;
                        var clr = colors != null && cIdx < colors.Length ? colors[cIdx] : DefaultSeriesColors[cIdx % DefaultSeriesColors.Length];
                        comboLine.AppendChild(BuildLineSeries(sIdx, lineData[ci].name, categories, lineData[ci].values, clr));
                    }
                    comboLine.AppendChild(new C.ShowMarker { Val = true });
                    comboLine.AppendChild(new C.AxisId { Val = catAxisId });
                    comboLine.AppendChild(new C.AxisId { Val = valAxisId });
                    plotArea.AppendChild(comboLine);
                }
                chartElement = null;
                break;
            }
            default:
                throw new ArgumentException(
                    $"Unknown chart type: '{kind}'. Supported: column, bar, line, pie, doughnut, area, scatter, combo. " +
                    "Add 'stacked' or 'percentstacked' suffix for variants (e.g. columnstacked).");
        }

        if (chartElement != null)
            plotArea.AppendChild(chartElement);

        if (needsAxes)
        {
            if (kind == "scatter")
            {
                plotArea.AppendChild(BuildValueAxis(catAxisId, valAxisId, C.AxisPositionValues.Bottom));
                plotArea.AppendChild(BuildValueAxis(valAxisId, catAxisId, C.AxisPositionValues.Left));
            }
            else
            {
                plotArea.AppendChild(BuildCategoryAxis(catAxisId, valAxisId));
                plotArea.AppendChild(BuildValueAxis(valAxisId, catAxisId, C.AxisPositionValues.Left));
            }
        }

        chart.AppendChild(plotArea);

        var showLegend = properties.GetValueOrDefault("legend", "true");
        if (!showLegend.Equals("false", StringComparison.OrdinalIgnoreCase) &&
            !showLegend.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            var legendPos = showLegend.ToLowerInvariant() switch
            {
                "top" or "t" => C.LegendPositionValues.Top,
                "left" or "l" => C.LegendPositionValues.Left,
                "right" or "r" => C.LegendPositionValues.Right,
                "bottom" or "b" => C.LegendPositionValues.Bottom,
                _ => C.LegendPositionValues.Bottom
            };
            chart.AppendChild(new C.Legend(
                new C.LegendPosition { Val = legendPos },
                new C.Overlay { Val = false }
            ));
        }

        chart.AppendChild(new C.PlotVisibleOnly { Val = true });
        chart.AppendChild(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        chartSpace.AppendChild(chart);
        return chartSpace;
    }

    /// <summary>
    /// Keys that BuildChartSpace doesn't handle directly but SetChartProperties does.
    /// After saving ChartSpace to a ChartPart, call SetChartProperties with these to apply them.
    /// </summary>
    internal static readonly HashSet<string> DeferredAddKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "datalabels", "labels",
        "axistitle", "vtitle", "cattitle", "htitle",
        "axismin", "min", "axismax", "max",
        "majorunit", "minorunit",
        "axisnumfmt", "axisnumberformat"
    };

    // ==================== Chart Type Builders ====================

    internal static C.BarChart BuildBarChart(
        C.BarDirectionValues direction, bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
    {
        var grouping = percentStacked ? C.BarGroupingValues.PercentStacked
            : stacked ? C.BarGroupingValues.Stacked
            : C.BarGroupingValues.Clustered;

        var barChart = new C.BarChart(
            new C.BarDirection { Val = direction },
            new C.BarGrouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            barChart.AppendChild(BuildBarSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
        }

        barChart.AppendChild(new C.GapWidth { Val = (ushort)150 });
        if (stacked || percentStacked)
            barChart.AppendChild(new C.Overlap { Val = 100 });
        barChart.AppendChild(new C.AxisId { Val = catAxisId });
        barChart.AppendChild(new C.AxisId { Val = valAxisId });
        return barChart;
    }

    internal static C.LineChart BuildLineChart(
        bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
    {
        var grouping = percentStacked ? C.GroupingValues.PercentStacked
            : stacked ? C.GroupingValues.Stacked
            : C.GroupingValues.Standard;

        var lineChart = new C.LineChart(
            new C.Grouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            lineChart.AppendChild(BuildLineSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
        }

        lineChart.AppendChild(new C.ShowMarker { Val = true });
        lineChart.AppendChild(new C.AxisId { Val = catAxisId });
        lineChart.AppendChild(new C.AxisId { Val = valAxisId });
        return lineChart;
    }

    internal static C.AreaChart BuildAreaChart(
        bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
    {
        var grouping = percentStacked ? C.GroupingValues.PercentStacked
            : stacked ? C.GroupingValues.Stacked
            : C.GroupingValues.Standard;

        var areaChart = new C.AreaChart(
            new C.Grouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            areaChart.AppendChild(BuildAreaSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
        }

        areaChart.AppendChild(new C.AxisId { Val = catAxisId });
        areaChart.AppendChild(new C.AxisId { Val = valAxisId });
        return areaChart;
    }

    internal static C.PieChart BuildPieChart(
        string[]? categories, List<(string name, double[] values)> seriesData)
    {
        var pieChart = new C.PieChart(new C.VaryColors { Val = true });
        if (seriesData.Count > 0)
            pieChart.AppendChild(BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values));
        return pieChart;
    }

    internal static C.DoughnutChart BuildDoughnutChart(
        string[]? categories, List<(string name, double[] values)> seriesData)
    {
        var chart = new C.DoughnutChart(new C.VaryColors { Val = true });
        if (seriesData.Count > 0)
            chart.AppendChild(BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values));
        chart.AppendChild(new C.HoleSize { Val = 50 });
        return chart;
    }

    internal static C.ScatterChart BuildScatterChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId)
    {
        var scatterChart = new C.ScatterChart(
            new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
            new C.VaryColors { Val = false }
        );

        double[]? xValues = null;
        if (categories != null)
            xValues = categories.Select(c => double.TryParse(c, out var v) ? v : 0).ToArray();

        for (int i = 0; i < seriesData.Count; i++)
        {
            scatterChart.AppendChild(BuildScatterSeries((uint)i, seriesData[i].name,
                xValues, seriesData[i].values));
        }

        scatterChart.AppendChild(new C.AxisId { Val = catAxisId });
        scatterChart.AppendChild(new C.AxisId { Val = valAxisId });
        return scatterChart;
    }

    // ==================== Default Series Colors ====================

    internal static readonly string[] DefaultSeriesColors =
    {
        "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
        "264478", "9B4A22", "636363", "BF8F00", "3A75A8", "4E8538"
    };

    // ==================== Series Color ====================

    internal static void ApplySeriesColor(OpenXmlCompositeElement series, string color)
    {
        series.RemoveAllChildren<C.ChartShapeProperties>();
        var spPr = new C.ChartShapeProperties();
        spPr.AppendChild(new Drawing.SolidFill(
            new Drawing.RgbColorModelHex { Val = color.TrimStart('#').ToUpperInvariant() }));
        var serText = series.GetFirstChild<C.SeriesText>();
        if (serText != null)
            serText.InsertAfterSelf(spPr);
        else
            series.PrependChild(spPr);
    }

    // ==================== Series Builders ====================

    internal static C.BarChartSeries BuildBarSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.BarChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.LineChartSeries BuildLineSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.LineChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.AreaChartSeries BuildAreaSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.AreaChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.PieChartSeries BuildPieSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.PieChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.ScatterChartSeries BuildScatterSeries(uint idx, string name,
        double[]? xValues, double[] yValues)
    {
        var series = new C.ScatterChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );

        if (xValues != null)
        {
            var xLit = new C.NumberLiteral(new C.PointCount { Val = (uint)xValues.Length });
            for (int i = 0; i < xValues.Length; i++)
                xLit.AppendChild(new C.NumericPoint(new C.NumericValue(xValues[i].ToString("G"))) { Index = (uint)i });
            series.AppendChild(new C.XValues(xLit));
        }

        var yLit = new C.NumberLiteral(new C.PointCount { Val = (uint)yValues.Length });
        for (int i = 0; i < yValues.Length; i++)
            yLit.AppendChild(new C.NumericPoint(new C.NumericValue(yValues[i].ToString("G"))) { Index = (uint)i });
        series.AppendChild(new C.YValues(yLit));

        return series;
    }

    // ==================== Data Builders ====================

    internal static C.CategoryAxisData BuildCategoryData(string[] categories)
    {
        var strLit = new C.StringLiteral(new C.PointCount { Val = (uint)categories.Length });
        for (int i = 0; i < categories.Length; i++)
            strLit.AppendChild(new C.StringPoint(new C.NumericValue(categories[i])) { Index = (uint)i });
        return new C.CategoryAxisData(strLit);
    }

    internal static C.Values BuildValues(double[] values)
    {
        var numLit = new C.NumberLiteral(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Length }
        );
        for (int i = 0; i < values.Length; i++)
            numLit.AppendChild(new C.NumericPoint(new C.NumericValue(values[i].ToString("G"))) { Index = (uint)i });
        return new C.Values(numLit);
    }

    // ==================== Axis Builders ====================

    internal static C.CategoryAxis BuildCategoryAxis(uint axisId, uint crossAxisId)
    {
        return new C.CategoryAxis(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.MajorTickMark { Val = C.TickMarkValues.Outside },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 }
        );
    }

    internal static C.ValueAxis BuildValueAxis(uint axisId, uint crossAxisId, C.AxisPositionValues position)
    {
        return new C.ValueAxis(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = position },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.MajorTickMark { Val = C.TickMarkValues.Outside },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between }
        );
    }

    // ==================== Title Builder ====================

    internal static C.Title BuildChartTitle(string titleText)
    {
        return new C.Title(
            new C.ChartText(
                new C.RichText(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties(
                            new Drawing.DefaultRunProperties { FontSize = 1400, Bold = true }
                        ),
                        new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US", FontSize = 1400, Bold = true },
                            new Drawing.Text(titleText)
                        )
                    )
                )
            ),
            new C.Overlay { Val = false }
        );
    }

    // ==================== Chart Readback ====================

    internal static void ReadChartProperties(C.Chart chart, DocumentNode node, int depth)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return;

        var chartType = DetectChartType(plotArea);
        if (chartType != null) node.Format["chartType"] = chartType;

        var titleEl = chart.GetFirstChild<C.Title>();
        var titleText = titleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (titleText != null) node.Format["title"] = titleText;

        var legend = chart.GetFirstChild<C.Legend>();
        if (legend != null)
        {
            var pos = legend.GetFirstChild<C.LegendPosition>()?.Val?.HasValue == true
                ? legend.GetFirstChild<C.LegendPosition>()!.Val!.InnerText : "b";
            node.Format["legend"] = pos;
        }

        var dataLabels = plotArea.Descendants<C.DataLabels>().FirstOrDefault();
        if (dataLabels != null)
        {
            var parts = new List<string>();
            if (dataLabels.GetFirstChild<C.ShowValue>()?.Val?.Value == true) parts.Add("value");
            if (dataLabels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true) parts.Add("category");
            if (dataLabels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true) parts.Add("series");
            if (dataLabels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true) parts.Add("percent");
            if (parts.Count > 0) node.Format["dataLabels"] = string.Join(",", parts);
        }

        // Axis titles
        var valAxis = plotArea.GetFirstChild<C.ValueAxis>();
        var valAxisTitle = valAxis?.GetFirstChild<C.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (valAxisTitle != null) node.Format["axisTitle"] = valAxisTitle;

        var catAxis = plotArea.GetFirstChild<C.CategoryAxis>();
        var catAxisTitle = catAxis?.GetFirstChild<C.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (catAxisTitle != null) node.Format["catTitle"] = catAxisTitle;

        // Axis scale
        var scaling = valAxis?.GetFirstChild<C.Scaling>();
        var minVal = scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value;
        if (minVal != null) node.Format["axisMin"] = minVal;
        var maxVal = scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value;
        if (maxVal != null) node.Format["axisMax"] = maxVal;

        var majorUnit = valAxis?.GetFirstChild<C.MajorUnit>()?.Val?.Value;
        if (majorUnit != null) node.Format["majorUnit"] = majorUnit;
        var minorUnit = valAxis?.GetFirstChild<C.MinorUnit>()?.Val?.Value;
        if (minorUnit != null) node.Format["minorUnit"] = minorUnit;

        var axisNumFmt = valAxis?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
        if (axisNumFmt != null && axisNumFmt != "General") node.Format["axisNumFmt"] = axisNumFmt;

        var seriesCount = CountSeries(plotArea);
        node.Format["seriesCount"] = seriesCount;

        var cats = ReadCategories(plotArea);
        if (cats != null) node.Format["categories"] = string.Join(",", cats);

        if (depth > 0)
        {
            var seriesList = ReadAllSeries(plotArea);
            for (int i = 0; i < seriesList.Count; i++)
            {
                var (sName, sValues) = seriesList[i];
                var seriesNode = new DocumentNode
                {
                    Path = $"{node.Path}/series[{i + 1}]",
                    Type = "series",
                    Text = sName
                };
                seriesNode.Format["name"] = sName;
                seriesNode.Format["values"] = string.Join(",", sValues.Select(v => v.ToString("G")));
                var serEl = plotArea.Descendants<OpenXmlCompositeElement>()
                    .Where(e => e.LocalName == "ser").ElementAtOrDefault(i);
                var serColor = serEl?.GetFirstChild<C.ChartShapeProperties>()
                    ?.GetFirstChild<Drawing.SolidFill>();
                if (serColor != null)
                {
                    var colorVal = ReadColorFromFill(serColor);
                    if (colorVal != null) seriesNode.Format["color"] = colorVal;
                }
                node.Children.Add(seriesNode);
            }
            node.ChildCount = seriesList.Count;
        }
        else
        {
            node.ChildCount = seriesCount;
        }
    }

    internal static string? DetectChartType(C.PlotArea plotArea)
    {
        var chartTypeCount = plotArea.ChildElements
            .Count(e => e is C.BarChart or C.LineChart or C.PieChart or C.AreaChart
                or C.ScatterChart or C.DoughnutChart or C.Bar3DChart or C.Line3DChart or C.Pie3DChart);
        if (chartTypeCount > 1) return "combo";

        if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart bar)
        {
            var dir = bar.GetFirstChild<C.BarDirection>()?.Val?.Value;
            var grp = bar.GetFirstChild<C.BarGrouping>()?.Val?.InnerText;
            var prefix = dir == C.BarDirectionValues.Bar ? "bar" : "column";
            if (grp == "stacked") return $"{prefix}_stacked";
            if (grp == "percentStacked") return $"{prefix}_percentStacked";
            return prefix;
        }
        if (plotArea.GetFirstChild<C.LineChart>() != null) return "line";
        if (plotArea.GetFirstChild<C.PieChart>() != null) return "pie";
        if (plotArea.GetFirstChild<C.DoughnutChart>() != null) return "doughnut";
        if (plotArea.GetFirstChild<C.AreaChart>() != null) return "area";
        if (plotArea.GetFirstChild<C.ScatterChart>() != null) return "scatter";
        if (plotArea.GetFirstChild<C.Bar3DChart>() != null) return "bar3d";
        if (plotArea.GetFirstChild<C.Line3DChart>() != null) return "line3d";
        if (plotArea.GetFirstChild<C.Pie3DChart>() != null) return "pie3d";
        return null;
    }

    internal static int CountSeries(C.PlotArea plotArea)
    {
        return plotArea.Descendants<C.Index>()
            .Count(idx => idx.Parent?.LocalName == "ser");
    }

    internal static string[]? ReadCategories(C.PlotArea plotArea)
    {
        var catData = plotArea.Descendants<C.CategoryAxisData>().FirstOrDefault();
        if (catData == null) return null;

        var strLit = catData.GetFirstChild<C.StringLiteral>();
        if (strLit != null)
        {
            return strLit.Elements<C.StringPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? "")
                .ToArray();
        }

        var strRef = catData.GetFirstChild<C.StringReference>();
        var strCache = strRef?.GetFirstChild<C.StringCache>();
        if (strCache != null)
        {
            return strCache.Elements<C.StringPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? "")
                .ToArray();
        }

        return null;
    }

    internal static List<(string name, double[] values)> ReadAllSeries(C.PlotArea plotArea)
    {
        var result = new List<(string name, double[] values)>();

        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent != null &&
                (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart"))))
        {
            var serText = ser.GetFirstChild<C.SeriesText>();
            var name = serText?.Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? "?";

            var values = ReadNumericData(ser.GetFirstChild<C.Values>())
                ?? ReadNumericData(ser.Elements<OpenXmlCompositeElement>()
                    .FirstOrDefault(e => e.LocalName == "yVal"))
                ?? Array.Empty<double>();

            result.Add((name, values));
        }

        return result;
    }

    internal static double[]? ReadNumericData(OpenXmlCompositeElement? valElement)
    {
        if (valElement == null) return null;

        var numLit = valElement.GetFirstChild<C.NumberLiteral>();
        if (numLit != null)
        {
            return numLit.Elements<C.NumericPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => double.TryParse(p.GetFirstChild<C.NumericValue>()?.Text, out var v) ? v : 0)
                .ToArray();
        }

        var numRef = valElement.GetFirstChild<C.NumberReference>();
        var numCache = numRef?.GetFirstChild<C.NumberingCache>();
        if (numCache != null)
        {
            return numCache.Elements<C.NumericPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => double.TryParse(p.GetFirstChild<C.NumericValue>()?.Text, out var v) ? v : 0)
                .ToArray();
        }

        return null;
    }

    internal static string? ReadColorFromFill(Drawing.SolidFill? solidFill)
    {
        if (solidFill == null) return null;
        var rgb = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (rgb != null) return rgb;
        var scheme = solidFill.GetFirstChild<Drawing.SchemeColor>()?.Val;
        if (scheme?.HasValue == true) return scheme.InnerText;
        return null;
    }

    // ==================== Chart Set ====================

    internal static void UpdateSeriesData(C.PlotArea plotArea, List<(string name, double[] values)> newData)
    {
        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();

        for (int i = 0; i < Math.Min(newData.Count, allSer.Count); i++)
        {
            var ser = allSer[i];
            var (sName, sVals) = newData[i];

            var serText = ser.GetFirstChild<C.SeriesText>();
            if (serText != null)
            {
                serText.RemoveAllChildren();
                serText.AppendChild(new C.NumericValue(sName));
            }

            var valEl = ser.GetFirstChild<C.Values>();
            if (valEl != null)
            {
                valEl.RemoveAllChildren();
                var builtVals = BuildValues(sVals);
                foreach (var child in builtVals.ChildElements.ToList())
                    valEl.AppendChild(child.CloneNode(true));
            }
        }
    }

    internal static List<string> SetChartProperties(ChartPart chartPart, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var chartSpace = chartPart.ChartSpace;
        var chart = chartSpace?.GetFirstChild<C.Chart>();
        if (chart == null) { unsupported.AddRange(properties.Keys); return unsupported; }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "title":
                    chart.RemoveAllChildren<C.Title>();
                    if (!string.IsNullOrEmpty(value) && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        chart.PrependChild(BuildChartTitle(value));
                    break;

                case "legend":
                    chart.RemoveAllChildren<C.Legend>();
                    if (!value.Equals("false", StringComparison.OrdinalIgnoreCase) &&
                        !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var pos = value.ToLowerInvariant() switch
                        {
                            "top" or "t" => C.LegendPositionValues.Top,
                            "left" or "l" => C.LegendPositionValues.Left,
                            "right" or "r" => C.LegendPositionValues.Right,
                            _ => C.LegendPositionValues.Bottom
                        };
                        var plotVisOnly = chart.GetFirstChild<C.PlotVisibleOnly>();
                        var insertBefore = plotVisOnly as OpenXmlElement ?? chart.LastChild;
                        chart.InsertBefore(new C.Legend(
                            new C.LegendPosition { Val = pos },
                            new C.Overlay { Val = false }
                        ), insertBefore);
                    }
                    break;

                case "datalabels" or "labels":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var chartTypeEl in plotArea2.ChildElements
                        .Where(e => e.LocalName.Contains("Chart") || e.LocalName.Contains("chart")))
                    {
                        chartTypeEl.RemoveAllChildren<C.DataLabels>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var dl = new C.DataLabels();
                            var parts = value.ToLowerInvariant().Split(',').Select(s => s.Trim()).ToHashSet();
                            dl.AppendChild(new C.ShowLegendKey { Val = false });
                            dl.AppendChild(new C.ShowValue { Val = parts.Contains("value") || parts.Contains("true") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowCategoryName { Val = parts.Contains("category") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowSeriesName { Val = parts.Contains("series") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowPercent { Val = parts.Contains("percent") || parts.Contains("all") });
                            // Insert dLbls before gapWidth/overlap/axId per schema order
                            var dlInsertBefore = chartTypeEl.GetFirstChild<C.GapWidth>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.Overlap>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.AxisId>();
                            if (dlInsertBefore != null)
                                chartTypeEl.InsertBefore(dl, dlInsertBefore);
                            else
                                chartTypeEl.AppendChild(dl);
                        }
                    }
                    break;
                }

                case "colors":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var colorList = value.Split(',').Select(c => c.Trim()).ToArray();
                    var allSer = plotArea2.Descendants<OpenXmlCompositeElement>()
                        .Where(e => e.LocalName == "ser").ToList();
                    for (int ci = 0; ci < Math.Min(colorList.Length, allSer.Count); ci++)
                        ApplySeriesColor(allSer[ci], colorList[ci]);
                    break;
                }

                case "axistitle" or "vtitle":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.Title>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var insertAfter = (OpenXmlElement?)valAxis.GetFirstChild<C.MinorGridlines>()
                            ?? (OpenXmlElement?)valAxis.GetFirstChild<C.MajorGridlines>()
                            ?? valAxis.GetFirstChild<C.AxisPosition>();
                        if (insertAfter != null) valAxis.InsertAfter(BuildChartTitle(value), insertAfter);
                    }
                    break;
                }

                case "cattitle" or "htitle":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var catAxis = plotArea2?.GetFirstChild<C.CategoryAxis>();
                    if (catAxis == null) { unsupported.Add(key); break; }
                    catAxis.RemoveAllChildren<C.Title>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var insertAfter = (OpenXmlElement?)catAxis.GetFirstChild<C.MinorGridlines>()
                            ?? (OpenXmlElement?)catAxis.GetFirstChild<C.MajorGridlines>()
                            ?? catAxis.GetFirstChild<C.AxisPosition>();
                        if (insertAfter != null) catAxis.InsertAfter(BuildChartTitle(value), insertAfter);
                    }
                    break;
                }

                case "axismin" or "min":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAxis?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.MinAxisValue>();
                    scaling.AppendChild(new C.MinAxisValue { Val = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) });
                    break;
                }

                case "axismax" or "max":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAxis?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.MaxAxisValue>();
                    var maxEl = new C.MaxAxisValue { Val = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) };
                    // Schema order: logBase?, orientation, max?, min? — insert max after orientation
                    var orient = scaling.GetFirstChild<C.Orientation>();
                    if (orient != null) orient.InsertAfterSelf(maxEl);
                    else scaling.PrependChild(maxEl);
                    break;
                }

                case "majorunit":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.MajorUnit>();
                    valAxis.AppendChild(new C.MajorUnit { Val = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) });
                    break;
                }

                case "minorunit":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.MinorUnit>();
                    valAxis.AppendChild(new C.MinorUnit { Val = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) });
                    break;
                }

                case "axisnumfmt" or "axisnumberformat":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.NumberingFormat>();
                    var nf = new C.NumberingFormat { FormatCode = value, SourceLinked = false };
                    // Schema order: ...title, numFmt, majorTickMark... — insert before majorTickMark
                    var nfInsertBefore = valAxis.GetFirstChild<C.MajorTickMark>();
                    if (nfInsertBefore != null) valAxis.InsertBefore(nf, nfInsertBefore);
                    else valAxis.AppendChild(nf);
                    break;
                }

                case "categories":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var newCats = value.Split(',').Select(c => c.Trim()).ToArray();
                    foreach (var catData in plotArea2.Descendants<C.CategoryAxisData>())
                    {
                        catData.RemoveAllChildren();
                        catData.AppendChild(BuildCategoryData(newCats).FirstChild!.CloneNode(true));
                    }
                    break;
                }

                case "data":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var newSeries = ParseSeriesData(new Dictionary<string, string> { ["data"] = value });
                    UpdateSeriesData(plotArea2, newSeries);
                    break;
                }

                default:
                    if (key.StartsWith("series", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(key[6..], out var seriesIdx))
                    {
                        var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                        if (plotArea2 == null) { unsupported.Add(key); break; }
                        var allSer = plotArea2.Descendants<OpenXmlCompositeElement>()
                            .Where(e => e.LocalName == "ser").ToList();
                        if (seriesIdx < 1 || seriesIdx > allSer.Count) { unsupported.Add(key); break; }
                        var ser = allSer[seriesIdx - 1];

                        var colonIdx = value.IndexOf(':');
                        double[] vals;
                        if (colonIdx >= 0)
                        {
                            var sName = value[..colonIdx].Trim();
                            vals = value[(colonIdx + 1)..].Split(',').Select(v => double.Parse(v.Trim(), System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                            var serText = ser.GetFirstChild<C.SeriesText>();
                            if (serText != null)
                            {
                                serText.RemoveAllChildren();
                                serText.AppendChild(new C.NumericValue(sName));
                            }
                        }
                        else
                        {
                            vals = value.Split(',').Select(v => double.Parse(v.Trim(), System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                        }

                        var valEl = ser.GetFirstChild<C.Values>();
                        if (valEl != null)
                        {
                            valEl.RemoveAllChildren();
                            var builtVals = BuildValues(vals);
                            foreach (var child in builtVals.ChildElements.ToList())
                                valEl.AppendChild(child.CloneNode(true));
                        }
                        var yValEl = ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal");
                        if (yValEl != null)
                        {
                            yValEl.RemoveAllChildren();
                            var numLit = new C.NumberLiteral(
                                new C.FormatCode("General"),
                                new C.PointCount { Val = (uint)vals.Length });
                            for (int vi = 0; vi < vals.Length; vi++)
                                numLit.AppendChild(new C.NumericPoint(new C.NumericValue(vals[vi].ToString("G"))) { Index = (uint)vi });
                            yValEl.AppendChild(numLit);
                        }
                    }
                    else
                    {
                        unsupported.Add(key);
                    }
                    break;
            }
        }

        chartSpace!.Save();
        return unsupported;
    }
}
