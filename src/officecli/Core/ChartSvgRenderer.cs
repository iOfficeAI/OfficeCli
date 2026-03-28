// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Shared chart SVG rendering logic used by both PowerPoint and Excel HTML preview.
/// </summary>
internal class ChartSvgRenderer
{
    // Default chart colors matching Office theme accent colors
    public static readonly string[] DefaultColors = [
        "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
        "#264478", "#9E480E", "#636363", "#997300", "#255E91", "#43682B"
    ];

    // Chart styling — configurable per chart instance
    public string ValueColor { get; set; } = "#D0D8E0";
    public string CatColor { get; set; } = "#C8D0D8";
    public string AxisColor { get; set; } = "#B0B8C0";
    public string GridColor { get; set; } = "#333";
    public string AxisLineColor { get; set; } = "#555";
    public int ValFontPx { get; set; } = 9;
    public int CatFontPx { get; set; } = 9;

    public static string HtmlEncode(string text) =>
        text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
            .Replace("\"", "&quot;").Replace("'", "&#39;");

    public void RenderBarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph,
        bool horizontal, bool stacked = false, bool percentStacked = false,
        double? ooxmlMax = null, double? ooxmlMin = null, double? ooxmlMajorUnit = null,
        int? ooxmlGapWidth = null, int valFontSize = 9, int catFontSize = 9,
        bool showDataLabels = false)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;
        if (percentStacked) stacked = true;

        double maxVal;
        if (percentStacked) maxVal = 100;
        else if (stacked)
        {
            maxVal = 0;
            for (int c = 0; c < catCount; c++)
            {
                var sum = series.Sum(s => c < s.values.Length ? s.values[c] : 0);
                if (sum > maxVal) maxVal = sum;
            }
        }
        else maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;

        double niceMax, tickStep;
        int nTicks;
        if (!percentStacked)
        {
            if (ooxmlMax.HasValue && ooxmlMajorUnit.HasValue)
            {
                niceMax = ooxmlMax.Value;
                tickStep = ooxmlMajorUnit.Value;
                nTicks = (int)Math.Round(niceMax / tickStep);
            }
            else (niceMax, tickStep, nTicks) = ComputeNiceAxis(ooxmlMax ?? maxVal);
        }
        else { niceMax = 100; nTicks = 5; tickStep = 20; }

        if (horizontal)
        {
            var hLabelMargin = 50;
            var plotOx = ox + hLabelMargin;
            var plotPw = pw - hLabelMargin;
            var groupH = (double)ph / Math.Max(catCount, 1);
            var gapPct = (ooxmlGapWidth ?? 150) / 100.0;
            double barH, gap;
            if (stacked) { barH = groupH / (1 + gapPct); gap = (groupH - barH) / 2; }
            else { barH = groupH / (serCount + gapPct); gap = barH * gapPct / 2; }

            for (int t = 1; t <= nTicks; t++)
            {
                var gx = plotOx + (double)plotPw * t / nTicks;
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            for (int c = 0; c < catCount; c++)
            {
                var dataIdx = catCount - 1 - c;
                double stackX = 0;
                var catSum = percentStacked ? series.Sum(s => dataIdx < s.values.Length ? s.values[dataIdx] : 0) : 1;
                for (int s = 0; s < serCount; s++)
                {
                    var rawVal = dataIdx < series[s].values.Length ? series[s].values[dataIdx] : 0;
                    var val = percentStacked && catSum > 0 ? (rawVal / catSum) * 100 : rawVal;
                    var barW = (val / niceMax) * plotPw;
                    if (stacked)
                    {
                        var bx = plotOx + (stackX / niceMax) * plotPw;
                        var by = oy + c * groupH + gap;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        stackX += val;
                    }
                    else
                    {
                        var bx = plotOx;
                        var by = oy + c * groupH + gap + (serCount - 1 - s) * barH;
                        sb.AppendLine($"        <rect x=\"{bx}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                    }
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var dataIdx = catCount - 1 - c;
                var label = dataIdx < categories.Length ? categories[dataIdx] : "";
                var ly = oy + c * groupH + groupH / 2;
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{ly:0.#}\" fill=\"{CatColor}\" font-size=\"{catFontSize}\" text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= nTicks; t++)
            {
                var val = tickStep * t;
                var label = percentStacked ? $"{(int)val}%" : (val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}");
                var tx = plotOx + (double)plotPw * t / nTicks;
                sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{AxisColor}\" font-size=\"{valFontSize}\" text-anchor=\"middle\">{label}</text>");
            }
        }
        else
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var gapPct = (ooxmlGapWidth ?? 150) / 100.0;
            double barW, gap;
            if (stacked) { barW = groupW / (1 + gapPct); gap = (groupW - barW) / 2; }
            else { barW = groupW / (serCount + gapPct); gap = barW * gapPct / 2; }

            for (int t = 1; t <= nTicks; t++)
            {
                var gy = oy + ph - (double)ph * t / nTicks;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            for (int c = 0; c < catCount; c++)
            {
                double stackY = 0;
                var catSum = percentStacked ? series.Sum(s => c < s.values.Length ? s.values[c] : 0) : 1;
                for (int s = 0; s < serCount; s++)
                {
                    var rawVal = c < series[s].values.Length ? series[s].values[c] : 0;
                    var val = percentStacked && catSum > 0 ? (rawVal / catSum) * 100 : rawVal;
                    var barH = (val / niceMax) * ph;
                    if (stacked)
                    {
                        var bx = ox + c * groupW + gap;
                        var by = oy + ph - (stackY / niceMax) * ph - barH;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        stackY += val;
                    }
                    else
                    {
                        var bx = ox + c * groupW + gap + s * barW;
                        var by = oy + ph - barH;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        if (showDataLabels)
                        {
                            var vlabel = rawVal % 1 == 0 ? $"{(int)rawVal}" : $"{rawVal:0.#}";
                            sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by - 3:0.#}\" fill=\"{ValueColor}\" font-size=\"8\" text-anchor=\"middle\">{vlabel}</text>");
                        }
                    }
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var lx = ox + c * groupW + groupW / 2;
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{catFontSize}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= nTicks; t++)
            {
                var val = tickStep * t;
                var label = percentStacked ? $"{(int)val}%" : (val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}");
                var ty = oy + ph - (double)ph * t / nTicks;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{valFontSize}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    public void RenderLineChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool showDataLabels = false)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var (niceMax, tickStep, nTicks) = ComputeNiceAxis(maxVal);

        for (int t = 1; t <= nTicks; t++)
        {
            var gy = oy + ph - (double)ph * t / nTicks;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\" stroke-dasharray=\"none\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        for (int s = 0; s < series.Count; s++)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (series[s].values[c] / niceMax) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s]}\" stroke-width=\"2\"/>");
                for (int p = 0; p < points.Count; p++)
                {
                    var parts = points[p].Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s]}\"/>");
                    if (showDataLabels)
                    {
                        var val = series[s].values[p];
                        var vlabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                        sb.AppendLine($"        <text x=\"{parts[0]}\" y=\"{double.Parse(parts[1]) - 6:0.#}\" fill=\"{ValueColor}\" font-size=\"7\" text-anchor=\"middle\">{vlabel}</text>");
                    }
                }
            }
        }
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        for (int t = 0; t <= nTicks; t++)
        {
            var val = tickStep * t;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / nTicks;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public void RenderPieChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, double holeRatio = 0.0, bool showDataLabels = false)
    {
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.42;
        var innerR = r * holeRatio;
        var startAngle = -Math.PI / 2;

        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var endAngle = startAngle + sliceAngle;
            var color = i < colors.Count ? colors[i] : DefaultColors[i % DefaultColors.Length];

            if (values.Length == 1 && holeRatio <= 0)
                sb.AppendLine($"        <circle cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" r=\"{r:0.#}\" fill=\"{color}\" opacity=\"0.85\"/>");
            else if (holeRatio > 0)
            {
                var ox1 = cx + r * Math.Cos(startAngle); var oy1 = cy + r * Math.Sin(startAngle);
                var ox2 = cx + r * Math.Cos(endAngle); var oy2 = cy + r * Math.Sin(endAngle);
                var ix1 = cx + innerR * Math.Cos(endAngle); var iy1 = cy + innerR * Math.Sin(endAngle);
                var ix2 = cx + innerR * Math.Cos(startAngle); var iy2 = cy + innerR * Math.Sin(startAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {ox1:0.#},{oy1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {ox2:0.#},{oy2:0.#} L {ix1:0.#},{iy1:0.#} A {innerR:0.#},{innerR:0.#} 0 {largeArc},0 {ix2:0.#},{iy2:0.#} Z\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
            else
            {
                var x1 = cx + r * Math.Cos(startAngle); var y1 = cy + r * Math.Sin(startAngle);
                var x2 = cx + r * Math.Cos(endAngle); var y2 = cy + r * Math.Sin(endAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {cx:0.#},{cy:0.#} L {x1:0.#},{y1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
            startAngle = endAngle;
        }
        if (showDataLabels)
        {
            var labelAngle = -Math.PI / 2;
            var labelR = holeRatio > 0 ? r * (1 + holeRatio) / 2 : r * 0.65;
            for (int i = 0; i < values.Length; i++)
            {
                var sliceAngle = 2 * Math.PI * values[i] / total;
                var midAngle = labelAngle + sliceAngle / 2;
                var lx = cx + labelR * Math.Cos(midAngle);
                var ly = cy + labelR * Math.Sin(midAngle);
                var pct = values[i] / total * 100;
                var label = pct >= 5 ? $"{pct:0}%" : "";
                if (!string.IsNullOrEmpty(label))
                    sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"#fff\" font-size=\"9\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"central\">{label}</text>");
                labelAngle += sliceAngle;
            }
        }
    }

    public void RenderAreaChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool stacked = false)
    {
        if (series.Count == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        if (catCount == 0) return;

        var cumulative = new double[series.Count, catCount];
        for (int c = 0; c < catCount; c++)
        {
            double runningSum = 0;
            for (int s = 0; s < series.Count; s++)
            {
                var val = c < series[s].values.Length ? series[s].values[c] : 0;
                runningSum += stacked ? val : 0;
                cumulative[s, c] = stacked ? runningSum : val;
            }
        }
        var maxVal = 0.0;
        if (stacked) { for (int c = 0; c < catCount; c++) maxVal = Math.Max(maxVal, cumulative[series.Count - 1, c]); }
        else maxVal = series.SelectMany(s => s.values).DefaultIfEmpty(0).Max();
        if (maxVal <= 0) maxVal = 1;
        var (niceMax, tickInterval, tickCount) = ComputeNiceAxis(maxVal);

        for (int t = 1; t <= tickCount; t++)
        {
            var gy = oy + ph - (double)ph * t / tickCount;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        if (stacked)
        {
            for (int s = series.Count - 1; s >= 0; s--)
            {
                var topPoints = new List<string>();
                var bottomPoints = new List<string>();
                for (int c = 0; c < catCount; c++)
                {
                    var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                    topPoints.Add($"{px:0.#},{oy + ph - (cumulative[s, c] / niceMax) * ph:0.#}");
                    var bottomVal = s > 0 ? cumulative[s - 1, c] : 0;
                    bottomPoints.Add($"{px:0.#},{oy + ph - (bottomVal / niceMax) * ph:0.#}");
                }
                bottomPoints.Reverse();
                sb.AppendLine($"        <polygon points=\"{string.Join(" ", topPoints)} {string.Join(" ", bottomPoints)}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
            }
        }
        else
        {
            var renderOrder = Enumerable.Range(0, series.Count).OrderByDescending(s => series[s].values.DefaultIfEmpty(0).Max()).ToList();
            foreach (var s in renderOrder)
            {
                var topPoints = new List<string>();
                for (int c = 0; c < catCount; c++)
                {
                    var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                    var val = c < series[s].values.Length ? series[s].values[c] : 0;
                    topPoints.Add($"{px:0.#},{oy + ph - (val / niceMax) * ph:0.#}");
                }
                var firstX = ox + (catCount > 1 ? 0 : pw / 2.0);
                var lastIdx = Math.Min(series[s].values.Length - 1, catCount - 1);
                var lastX = ox + (catCount > 1 ? (double)pw * lastIdx / (catCount - 1) : pw / 2.0);
                sb.AppendLine($"        <polygon points=\"{firstX:0.#},{oy + ph} {string.Join(" ", topPoints)} {lastX:0.#},{oy + ph}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
            }
        }
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        for (int t = 0; t <= tickCount; t++)
        {
            var val = tickInterval * t;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / tickCount;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public void RenderRadarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, int catLabelFontSize = 0)
    {
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        if (catCount < 3) return;
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;

        var labelSize = catLabelFontSize > 0 ? catLabelFontSize : 11;
        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.33;

        for (int ring = 1; ring <= 5; ring++)
        {
            var rr = r * ring / 5;
            var gridPoints = new List<string>();
            for (int c = 0; c < catCount; c++)
            {
                var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                gridPoints.Add($"{cx + rr * Math.Cos(angle):0.#},{cy + rr * Math.Sin(angle):0.#}");
            }
            sb.AppendLine($"        <polygon points=\"{string.Join(" ", gridPoints)}\" fill=\"none\" stroke=\"#ccc\" stroke-width=\"0.5\"/>");
        }
        for (int c = 0; c < catCount; c++)
        {
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            sb.AppendLine($"        <line x1=\"{cx:0.#}\" y1=\"{cy:0.#}\" x2=\"{cx + r * Math.Cos(angle):0.#}\" y2=\"{cy + r * Math.Sin(angle):0.#}\" stroke=\"#ccc\" stroke-width=\"0.5\"/>");
        }
        for (int s = 0; s < series.Count; s++)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                var val = series[s].values[c] / maxVal * r;
                points.Add($"{cx + val * Math.Cos(angle):0.#},{cy + val * Math.Sin(angle):0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polygon points=\"{string.Join(" ", points)}\" fill=\"{colors[s]}\" fill-opacity=\"0.2\" stroke=\"{colors[s]}\" stroke-width=\"2\"/>");
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s]}\"/>");
                }
            }
        }
        foreach (var frac in new[] { 0.2, 0.4, 0.6, 0.8, 1.0 })
        {
            var val = maxVal * frac;
            var tickLabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{cx + 2:0.#}\" y=\"{cy - r * frac:0.#}\" fill=\"{AxisColor}\" font-size=\"8\" dominant-baseline=\"middle\">{tickLabel}</text>");
        }
        var labelOffset = Math.Max(18, r * 0.15);
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            var lx = cx + (r + labelOffset) * Math.Cos(angle);
            var ly = cy + (r + labelOffset) * Math.Sin(angle);
            var anchor = Math.Abs(Math.Cos(angle)) < 0.1 ? "middle" : (Math.Cos(angle) > 0 ? "start" : "end");
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"{CatColor}\" font-size=\"{labelSize}\" text-anchor=\"{anchor}\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
        }
    }

    public void RenderBubbleChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> series, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph)
    {
        var bubbleSeries = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent?.LocalName == "bubbleChart").ToList();

        var allX = new List<double>(); var allY = new List<double>(); var allSize = new List<double>();
        var seriesData = new List<(double[] x, double[] y, double[] size)>();

        for (int s = 0; s < bubbleSeries.Count; s++)
        {
            var ser = bubbleSeries[s];
            var xVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "xVal")) ?? [];
            var yVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal")) ?? [];
            var sizeVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "bubbleSize")) ?? yVals;
            seriesData.Add((xVals, yVals, sizeVals));
            allX.AddRange(xVals); allY.AddRange(yVals); allSize.AddRange(sizeVals);
        }
        if (seriesData.Count == 0)
        {
            foreach (var s in series)
            {
                var xVals = Enumerable.Range(0, s.values.Length).Select(i => (double)i).ToArray();
                seriesData.Add((xVals, s.values, s.values));
                allX.AddRange(xVals); allY.AddRange(s.values); allSize.AddRange(s.values);
            }
        }
        if (allY.Count == 0) return;
        var minX = allX.Min(); var maxX = allX.Max(); if (maxX <= minX) maxX = minX + 1;
        var minY = allY.Min(); var maxY = allY.Max(); if (maxY <= minY) maxY = minY + 1;
        var maxSz = allSize.Count > 0 ? allSize.Max() : 1; if (maxSz <= 0) maxSz = 1;
        var bubbleScaleEl = plotArea.Descendants<BubbleScale>().FirstOrDefault();
        var bubbleScale = bubbleScaleEl?.Val?.HasValue == true ? bubbleScaleEl.Val.Value / 100.0 : 1.0;
        var maxRadius = Math.Min(pw, ph) * 0.12 * bubbleScale;

        for (int t = 1; t <= 4; t++)
        {
            var gy = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        for (int s = 0; s < seriesData.Count; s++)
        {
            var (xVals, yVals, sizeVals) = seriesData[s];
            var count = Math.Min(xVals.Length, yVals.Length);
            for (int i = 0; i < count; i++)
            {
                var bx = ox + ((xVals[i] - minX) / (maxX - minX)) * pw;
                var by = oy + ph - ((yVals[i] - minY) / (maxY - minY)) * ph;
                var sz = i < sizeVals.Length ? sizeVals[i] : yVals[i];
                var r = Math.Sqrt(Math.Max(0, sz) / maxSz) * maxRadius + maxRadius * 0.15;
                sb.AppendLine($"        <circle cx=\"{bx:0.#}\" cy=\"{by:0.#}\" r=\"{r:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.6\"/>");
            }
        }
        for (int t = 0; t <= 4; t++)
        {
            var val = minX + (maxX - minX) * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox + (double)pw * t / 4:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{label}</text>");
        }
        for (int t = 0; t <= 4; t++)
        {
            var val = minY + (maxY - minY) * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{oy + ph - (double)ph * t / 4:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public void RenderComboChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> seriesList, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph)
    {
        var barIndices = new HashSet<int>();
        var lineIndices = new HashSet<int>();
        var areaIndices = new HashSet<int>();
        var idx = 0;
        foreach (var chartEl in plotArea.ChildElements)
        {
            var serElements = chartEl.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser").ToList();
            if (serElements.Count == 0) continue;
            var localName = chartEl.LocalName.ToLowerInvariant();
            var isBar = localName.Contains("bar");
            var isArea = localName.Contains("area");
            foreach (var _ in serElements)
            {
                if (isBar) barIndices.Add(idx);
                else if (isArea) areaIndices.Add(idx);
                else lineIndices.Add(idx);
                idx++;
            }
        }

        var allValues = seriesList.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var rawMax = allValues.Max(); if (rawMax <= 0) rawMax = 1;
        var (maxVal, _, _) = ComputeNiceAxis(rawMax);
        var catCount = Math.Max(categories.Length, seriesList.Max(s => s.values.Length));

        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        var barSeries = barIndices.Where(i => i < seriesList.Count).ToList();
        if (barSeries.Count > 0)
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = groupW * 0.5 / barSeries.Count;
            var gap = (groupW - barSeries.Count * barW) / 2;
            for (int bi = 0; bi < barSeries.Count; bi++)
            {
                var s = barSeries[bi];
                for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
                {
                    var val = seriesList[s].values[c];
                    var barH = (val / maxVal) * ph;
                    sb.AppendLine($"        <rect x=\"{ox + c * groupW + gap + bi * barW:0.#}\" y=\"{oy + ph - barH:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                }
            }
        }
        foreach (var s in areaIndices.Where(i => i < seriesList.Count))
        {
            var points = new List<string>();
            for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                points.Add($"{px:0.#},{oy + ph - (seriesList[s].values[c] / maxVal) * ph:0.#}");
            }
            if (points.Count > 0)
            {
                var firstX = ox + (catCount > 1 ? 0 : pw / 2.0);
                var lastX = ox + (catCount > 1 ? (double)pw * (seriesList[s].values.Length - 1) / (catCount - 1) : pw / 2.0);
                sb.AppendLine($"        <polygon points=\"{firstX:0.#},{oy + ph} {string.Join(" ", points)} {lastX:0.#},{oy + ph}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.3\"/>");
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s % colors.Count]}\" stroke-width=\"2\"/>");
            }
        }
        foreach (var s in lineIndices.Where(i => i < seriesList.Count))
        {
            var points = new List<string>();
            for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                points.Add($"{px:0.#},{oy + ph - (seriesList[s].values[c] / maxVal) * ph:0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s % colors.Count]}\" stroke-width=\"2.5\"/>");
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s % colors.Count]}\"/>");
                }
            }
        }
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (double)pw * c / Math.Max(catCount, 1) + (double)pw / Math.Max(catCount, 1) / 2;
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{oy + ph - (double)ph * t / 4:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public void RenderStockChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> series, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max(); var minVal = allValues.Min();
        if (maxVal <= minVal) maxVal = minVal + 1;
        var range = maxVal - minVal;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        var upColor = "#2ECC71"; var downColor = "#E74C3C";
        var stockChart = plotArea.GetFirstChild<StockChart>();
        if (stockChart != null)
        {
            var upFill = stockChart.Descendants<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "upBars")
                ?.Descendants<Drawing.SolidFill>().FirstOrDefault()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (upFill != null) upColor = $"#{upFill}";
            var downFill = stockChart.Descendants<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "downBars")
                ?.Descendants<Drawing.SolidFill>().FirstOrDefault()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (downFill != null) downColor = $"#{downFill}";
        }

        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        var groupW = (double)pw / Math.Max(catCount, 1);
        if (series.Count >= 4)
        {
            for (int c = 0; c < catCount; c++)
            {
                var open = c < series[0].values.Length ? series[0].values[c] : 0;
                var high = c < series[1].values.Length ? series[1].values[c] : 0;
                var low = c < series[2].values.Length ? series[2].values[c] : 0;
                var close = c < series[3].values.Length ? series[3].values[c] : 0;
                var ccx = ox + c * groupW + groupW / 2;
                var yHigh = oy + ph - ((high - minVal) / range) * ph;
                var yLow = oy + ph - ((low - minVal) / range) * ph;
                var yOpen = oy + ph - ((open - minVal) / range) * ph;
                var yClose = oy + ph - ((close - minVal) / range) * ph;
                var color = close >= open ? upColor : downColor;
                var barW = groupW * 0.5;
                sb.AppendLine($"        <line x1=\"{ccx:0.#}\" y1=\"{yHigh:0.#}\" x2=\"{ccx:0.#}\" y2=\"{yLow:0.#}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
                var bodyTop = Math.Min(yOpen, yClose); var bodyH = Math.Max(Math.Abs(yOpen - yClose), 1);
                sb.AppendLine($"        <rect x=\"{ccx - barW / 2:0.#}\" y=\"{bodyTop:0.#}\" width=\"{barW:0.#}\" height=\"{bodyH:0.#}\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
        }
        else { RenderLineChartSvg(sb, series, categories, colors, ox, oy, pw, ph); return; }

        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            sb.AppendLine($"        <text x=\"{ox + c * groupW + groupW / 2:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        for (int t = 0; t <= 4; t++)
        {
            var val = minVal + range * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{oy + ph - (double)ph * t / 4:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public static (double niceMax, double tickStep, int nTicks) ComputeNiceAxis(double maxVal)
    {
        if (maxVal <= 0) maxVal = 1;
        var mag = Math.Pow(10, Math.Floor(Math.Log10(maxVal)));
        var res = maxVal / mag;
        var tickStep = res <= 1.5 ? 0.2 * mag : res <= 4 ? 0.5 * mag : res <= 8 ? 1.0 * mag : 2.0 * mag;
        var niceMax = Math.Ceiling(maxVal / tickStep) * tickStep;
        if (niceMax < maxVal * 1.05) niceMax += tickStep;
        var nTicks = (int)Math.Round(niceMax / tickStep);
        if (nTicks < 2) nTicks = 2;
        return (niceMax, tickStep, nTicks);
    }
}
