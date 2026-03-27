// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

/// <summary>
/// Advanced chart features: reference lines, conditional coloring, waterfall simulation.
/// </summary>
internal static partial class ChartHelper
{
    // ==================== Reference Line ====================

    /// <summary>
    /// Add a reference (target/average) line to a chart by inserting a hidden line series.
    /// Format: "value" or "value:color" or "value:color:label" or "value:color:label:dash"
    /// e.g. "50", "75:FF0000", "100:00AA00:Target", "80:0000FF:Average:dash"
    /// </summary>
    internal static void AddReferenceLine(C.Chart chart, string spec)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return;

        // Remove any existing reference line series before adding a new one
        RemoveExistingReferenceLines(plotArea);

        var parts = spec.Split(':');
        if (!double.TryParse(parts[0].Trim(),
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var refValue))
            throw new ArgumentException(
                $"Invalid referenceLine value '{parts[0]}'. Expected: number or number:color:label:dash (e.g. '50:FF0000:Target:dash').");

        var color = parts.Length > 1 ? parts[1].Trim() : "FF0000";
        var label = parts.Length > 2 ? parts[2].Trim() : $"Ref ({refValue})";
        var dash = parts.Length > 3 ? parts[3].Trim() : "dash";

        // Find max data point count from existing series (after removing old ref lines)
        var existingSerCount = CountSeries(plotArea);
        var maxDataPoints = 0;
        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
        {
            var vals = ser.GetFirstChild<C.Values>();
            var numLit = vals?.GetFirstChild<C.NumberLiteral>();
            var ptCount = numLit?.GetFirstChild<C.PointCount>()?.Val?.Value ?? 0;
            if ((int)ptCount > maxDataPoints) maxDataPoints = (int)ptCount;
            var numRef = vals?.GetFirstChild<C.NumberReference>();
            var cacheCount = numRef?.GetFirstChild<C.NumberingCache>()?.GetFirstChild<C.PointCount>()?.Val?.Value ?? 0;
            if ((int)cacheCount > maxDataPoints) maxDataPoints = (int)cacheCount;
        }
        if (maxDataPoints == 0) maxDataPoints = 3;

        // Create a flat line series (all values = refValue)
        var refValues = Enumerable.Repeat(refValue, maxDataPoints).ToArray();
        var seriesIdx = (uint)existingSerCount;

        // Find or create a LineChart in the plot area for the reference line
        var lineChart = plotArea.GetFirstChild<C.LineChart>();
        if (lineChart == null)
        {
            // Create a new line chart overlay — shares axes with existing chart
            uint catAxisId = 1, valAxisId = 2;
            // Try to find existing axis IDs
            var existingCatAx = plotArea.GetFirstChild<C.CategoryAxis>()?.GetFirstChild<C.AxisId>()?.Val?.Value;
            var existingValAx = plotArea.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.AxisId>()?.Val?.Value;
            if (existingCatAx != null) catAxisId = existingCatAx.Value;
            if (existingValAx != null) valAxisId = existingValAx.Value;

            lineChart = new C.LineChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false }
            );
            lineChart.AppendChild(new C.ShowMarker { Val = false });
            lineChart.AppendChild(new C.AxisId { Val = catAxisId });
            lineChart.AppendChild(new C.AxisId { Val = valAxisId });

            // Insert before axes
            var firstAxis = plotArea.Elements<C.CategoryAxis>().FirstOrDefault() as OpenXmlElement
                ?? plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            if (firstAxis != null)
                plotArea.InsertBefore(lineChart, firstAxis);
            else
                plotArea.AppendChild(lineChart);
        }

        // Build the reference line series
        var refSer = new C.LineChartSeries();
        refSer.AppendChild(new C.Index { Val = seriesIdx });
        refSer.AppendChild(new C.Order { Val = seriesIdx });
        refSer.AppendChild(new C.SeriesText(new C.NumericValue(label)));

        // Style: colored dashed line, no markers
        var spPr = new C.ChartShapeProperties();
        var outline = new Drawing.Outline { Width = 19050 }; // 1.5pt
        var sf = new Drawing.SolidFill();
        sf.AppendChild(BuildChartColorElement(color));
        outline.AppendChild(sf);
        outline.AppendChild(new Drawing.PresetDash { Val = ParseDashStyle(dash) });
        spPr.AppendChild(outline);
        refSer.AppendChild(spPr);

        // No marker
        refSer.AppendChild(new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.None }));

        // Flat data — same value repeated
        var numLitRef = new C.NumberLiteral(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)refValues.Length });
        for (int i = 0; i < refValues.Length; i++)
            numLitRef.AppendChild(new C.NumericPoint(
                new C.NumericValue(refValue.ToString("G"))) { Index = (uint)i });
        refSer.AppendChild(new C.Values(numLitRef));

        // Insert before ShowMarker
        var showMarkerEl = lineChart.GetFirstChild<C.ShowMarker>();
        if (showMarkerEl != null)
            lineChart.InsertBefore(refSer, showMarkerEl);
        else
        {
            var axisIdEl = lineChart.GetFirstChild<C.AxisId>();
            if (axisIdEl != null)
                lineChart.InsertBefore(refSer, axisIdEl);
            else
                lineChart.AppendChild(refSer);
        }
    }

    /// <summary>
    /// Remove existing reference line series from a plot area.
    /// A reference line series is identified as a LineChartSeries in a LineChart
    /// where all data points have the same value (flat line), the series has a dashed
    /// outline style, and the marker is set to None.
    /// </summary>
    internal static void RemoveExistingReferenceLines(C.PlotArea plotArea)
    {
        var lineChart = plotArea.GetFirstChild<C.LineChart>();
        if (lineChart == null) return;

        var toRemove = new List<C.LineChartSeries>();
        foreach (var ser in lineChart.Elements<C.LineChartSeries>())
        {
            // Check for reference line markers: no marker (None) and dashed outline
            var marker = ser.GetFirstChild<C.Marker>();
            var markerSymbol = marker?.GetFirstChild<C.Symbol>()?.Val?.Value;
            if (markerSymbol != C.MarkerStyleValues.None) continue;

            var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
            var outline = spPr?.GetFirstChild<Drawing.Outline>();
            var hasDash = outline?.GetFirstChild<Drawing.PresetDash>() != null;
            if (!hasDash) continue;

            // Check if all values are the same (flat line = reference line)
            var vals = ser.GetFirstChild<C.Values>();
            var numLit = vals?.GetFirstChild<C.NumberLiteral>();
            if (numLit != null)
            {
                var points = numLit.Elements<C.NumericPoint>().Select(p => p.InnerText).Distinct().ToList();
                if (points.Count == 1)
                    toRemove.Add(ser);
            }
        }

        foreach (var ser in toRemove)
            ser.Remove();

        // If the LineChart is now empty (no series left), remove it entirely
        if (!lineChart.Elements<C.LineChartSeries>().Any())
            lineChart.Remove();
    }

    // ==================== Conditional Coloring ====================

    /// <summary>
    /// Apply conditional coloring to data points based on value thresholds.
    /// Format: "threshold:belowColor:aboveColor" or "low:lowColor:mid:midColor:high:highColor"
    /// Simple: "0:FF0000:00AA00" — below 0 = red, above 0 = green
    /// Three-tier: "0:FF0000:50:FFAA00:100:00AA00" — red/orange/green zones
    /// </summary>
    internal static void ApplyColorRule(C.PlotArea plotArea, string spec)
    {
        var parts = spec.Split(':');
        if (parts.Length < 3)
            throw new ArgumentException(
                $"Invalid colorRule '{spec}'. Expected: threshold:belowColor:aboveColor (e.g. '0:FF0000:00AA00') " +
                "or low:lowColor:mid:midColor:high:highColor (e.g. '0:FF0000:50:FFAA00:100:00AA00').");

        var rules = new List<(double threshold, string color)>();
        string topColor;

        if (parts.Length == 3)
        {
            // Simple two-zone: threshold:belowColor:aboveColor
            if (!double.TryParse(parts[0], System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var t))
                throw new ArgumentException($"Invalid threshold '{parts[0]}' in colorRule. Expected a number.");
            rules.Add((t, parts[1].Trim()));
            topColor = parts[2].Trim();
        }
        else
        {
            // Multi-zone: t1:c1:t2:c2:...:cN
            for (int i = 0; i < parts.Length - 1; i += 2)
            {
                if (!double.TryParse(parts[i], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var t))
                    throw new ArgumentException($"Invalid threshold '{parts[i]}' in colorRule.");
                rules.Add((t, parts[i + 1].Trim()));
            }
            topColor = parts.Length % 2 == 1 ? parts[^1].Trim() : rules[^1].color;
            if (parts.Length % 2 == 0)
                rules.RemoveAt(rules.Count - 1); // Last pair has no "above" — use as topColor
        }

        // Apply to each data point in each series
        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
        {
            var values = ReadNumericData(ser.GetFirstChild<C.Values>())
                ?? ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal"));
            if (values == null) continue;

            for (int pi = 0; pi < values.Length; pi++)
            {
                var val = values[pi];
                string pointColor = topColor;
                foreach (var (threshold, color) in rules)
                {
                    if (val < threshold) { pointColor = color; break; }
                    pointColor = color; // at or above this threshold, use this color
                }
                // If above all thresholds, use topColor
                if (rules.Count > 0 && val >= rules[^1].threshold)
                    pointColor = topColor;

                ApplyDataPointColor(ser, pi, pointColor);
            }
        }
    }

    // ==================== Waterfall Chart (Stacked Bar Simulation) ====================

    /// <summary>
    /// Build a waterfall chart using stacked bar technique:
    /// - Invisible "base" series for the running total
    /// - Visible "increase" series (positive changes) and "decrease" series (negative changes)
    /// - Last bar shows the total
    ///
    /// Input: categories and a single series of change values.
    /// e.g. categories=Revenue,Cost,Tax,Profit  data=Cashflow:100,-30,-15,55
    /// The last value can be auto-calculated as the total if "auto" or omitted.
    /// </summary>
    internal static C.ChartSpace BuildWaterfallChart(
        string? title,
        string[]? categories,
        double[] values,
        string? increaseColor,
        string? decreaseColor,
        string? totalColor,
        Dictionary<string, string> properties)
    {
        increaseColor ??= "4472C4"; // blue
        decreaseColor ??= "FF0000"; // red
        totalColor ??= "2E75B6";    // dark blue

        var n = values.Length;
        var baseVals = new double[n];
        var incVals = new double[n];
        var decVals = new double[n];

        double running = 0;
        for (int i = 0; i < n; i++)
        {
            var v = values[i];
            if (i == n - 1 && properties.GetValueOrDefault("waterfallTotal", "true")
                .Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                // Last bar = total (starts from 0)
                baseVals[i] = 0;
                incVals[i] = running + v;
                decVals[i] = 0;
            }
            else if (v >= 0)
            {
                baseVals[i] = running;
                incVals[i] = v;
                decVals[i] = 0;
                running += v;
            }
            else
            {
                baseVals[i] = running + v; // base drops by |v|
                incVals[i] = 0;
                decVals[i] = -v;
                running += v;
            }
        }

        categories ??= Enumerable.Range(1, n).Select(i => i.ToString()).ToArray();

        var chartSpace = new C.ChartSpace();
        var chart = new C.Chart();
        if (!string.IsNullOrEmpty(title))
            chart.AppendChild(BuildChartTitle(title));

        var plotArea = new C.PlotArea(new C.Layout());
        uint catAxisId = 1, valAxisId = 2;

        var barChart = new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Stacked },
            new C.VaryColors { Val = false }
        );

        // Series 0: invisible base
        var baseSer = BuildBarSeries(0, "Base", categories, baseVals, null);
        // Make base series invisible: no fill, no border
        baseSer.RemoveAllChildren<C.ChartShapeProperties>();
        var baseSpPr = new C.ChartShapeProperties();
        baseSpPr.AppendChild(new Drawing.NoFill());
        var baseOutline = new Drawing.Outline();
        baseOutline.AppendChild(new Drawing.NoFill());
        baseSpPr.AppendChild(baseOutline);
        baseSer.InsertAfter(baseSpPr, baseSer.GetFirstChild<C.SeriesText>());
        barChart.AppendChild(baseSer);

        // Series 1: increase (positive values)
        barChart.AppendChild(BuildBarSeries(1, "Increase", categories, incVals, increaseColor));

        // Series 2: decrease (negative values)
        barChart.AppendChild(BuildBarSeries(2, "Decrease", categories, decVals, decreaseColor));

        barChart.AppendChild(new C.GapWidth { Val = 80 });
        barChart.AppendChild(new C.Overlap { Val = 100 });
        barChart.AppendChild(new C.AxisId { Val = catAxisId });
        barChart.AppendChild(new C.AxisId { Val = valAxisId });

        plotArea.AppendChild(barChart);
        plotArea.AppendChild(BuildCategoryAxis(catAxisId, valAxisId));
        plotArea.AppendChild(BuildValueAxis(valAxisId, catAxisId, C.AxisPositionValues.Left));

        chart.AppendChild(plotArea);

        // Hide base series from legend
        var legend = new C.Legend(
            new C.LegendPosition { Val = C.LegendPositionValues.Bottom },
            new C.Overlay { Val = false }
        );
        // Delete legend entry for base series (index 0)
        // CT_Legend schema order: legendPos, legendEntry+, layout, overlay — insert after legendPos
        var leBase = new C.LegendEntry();
        leBase.AppendChild(new C.Index { Val = 0 });
        leBase.AppendChild(new C.Delete { Val = true });
        var legendPosEl = legend.GetFirstChild<C.LegendPosition>();
        if (legendPosEl != null)
            legendPosEl.InsertAfterSelf(leBase);
        else
            legend.PrependChild(leBase);
        chart.AppendChild(legend);

        chart.AppendChild(new C.PlotVisibleOnly { Val = true });
        chart.AppendChild(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        chartSpace.AppendChild(chart);

        // Color the total bar differently (last data point of increase series)
        if (properties.GetValueOrDefault("waterfallTotal", "true")
            .Equals("true", StringComparison.OrdinalIgnoreCase) && n > 0)
        {
            var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser").ToList();
            if (allSer.Count >= 2)
                ApplyDataPointColor(allSer[1], n - 1, totalColor);
        }

        return chartSpace;
    }

    // ==================== Flexible Combo Chart ====================

    /// <summary>
    /// Build a combo chart with per-series chart type assignment.
    /// comboTypes property: "column,column,line,area" — one type per series.
    /// </summary>
    internal static void RebuildComboChart(C.Chart chart, string comboTypes)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return;

        var typeList = comboTypes.Split(',').Select(t => t.Trim().ToLowerInvariant()).ToArray();

        // Read all existing series data
        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();

        if (allSer.Count == 0) return;

        // Read series data
        var seriesInfo = new List<(OpenXmlCompositeElement original, string targetType)>();
        for (int i = 0; i < allSer.Count; i++)
        {
            var targetType = i < typeList.Length ? typeList[i] : typeList[^1];
            seriesInfo.Add((allSer[i], targetType));
        }

        // Find axis IDs
        uint catAxisId = plotArea.GetFirstChild<C.CategoryAxis>()?.GetFirstChild<C.AxisId>()?.Val?.Value ?? 1;
        uint valAxisId = plotArea.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.AxisId>()?.Val?.Value ?? 2;

        // Remove existing chart type elements (but keep axes, layout, etc.)
        foreach (var ct in plotArea.ChildElements
            .Where(e => e.LocalName.EndsWith("Chart") || e.LocalName.EndsWith("chart"))
            .OfType<OpenXmlCompositeElement>().ToList())
        {
            ct.Remove();
        }

        // Group series by target chart type
        var groups = seriesInfo.GroupBy(s => s.targetType).ToList();
        foreach (var group in groups)
        {
            OpenXmlCompositeElement chartTypeEl;
            switch (group.Key)
            {
                case "bar":
                    chartTypeEl = new C.BarChart(
                        new C.BarDirection { Val = C.BarDirectionValues.Bar },
                        new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                        new C.VaryColors { Val = false });
                    break;
                case "column" or "col":
                    chartTypeEl = new C.BarChart(
                        new C.BarDirection { Val = C.BarDirectionValues.Column },
                        new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                        new C.VaryColors { Val = false });
                    break;
                case "line":
                    chartTypeEl = new C.LineChart(
                        new C.Grouping { Val = C.GroupingValues.Standard },
                        new C.VaryColors { Val = false });
                    break;
                case "area":
                    chartTypeEl = new C.AreaChart(
                        new C.Grouping { Val = C.GroupingValues.Standard },
                        new C.VaryColors { Val = false });
                    break;
                case "scatter":
                    chartTypeEl = new C.ScatterChart(
                        new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                        new C.VaryColors { Val = false });
                    break;
                default:
                    chartTypeEl = new C.LineChart(
                        new C.Grouping { Val = C.GroupingValues.Standard },
                        new C.VaryColors { Val = false });
                    break;
            }

            foreach (var (original, _) in group)
            {
                chartTypeEl.AppendChild(original.CloneNode(true));
            }

            chartTypeEl.AppendChild(new C.AxisId { Val = catAxisId });
            chartTypeEl.AppendChild(new C.AxisId { Val = valAxisId });

            // Insert before axes
            var firstAxis = plotArea.Elements<C.CategoryAxis>().FirstOrDefault() as OpenXmlElement
                ?? plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            if (firstAxis != null)
                plotArea.InsertBefore(chartTypeEl, firstAxis);
            else
                plotArea.AppendChild(chartTypeEl);
        }
    }
}
