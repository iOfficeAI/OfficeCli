// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddChart(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var chartMainPart = _doc.MainDocumentPart!;

        // Parse chart data
        var chartType = properties.FirstOrDefault(kv =>
            kv.Key.Equals("charttype", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Equals("type", StringComparison.OrdinalIgnoreCase)).Value
            ?? "column";
        var chartTitle = properties.GetValueOrDefault("title");
        var categories = Core.ChartHelper.ParseCategories(properties);
        var seriesData = Core.ChartHelper.ParseSeriesData(properties);

        if (seriesData.Count == 0)
            throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                "or series1=\"Revenue:100,200,300\"");

        // Dimensions (default: 15cm x 10cm)
        long chartCx = properties.TryGetValue("width", out var chartWStr) ? ParseEmu(chartWStr) : 5400000;
        long chartCy = properties.TryGetValue("height", out var chStr) ? ParseEmu(chStr) : 3600000;

        var docPropId = NextDocPropId();
        var chartName = chartTitle ?? $"Chart {docPropId}";

        // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
        if (Core.ChartExBuilder.IsExtendedChartType(chartType))
        {
            var cxChartSpace = Core.ChartExBuilder.BuildExtendedChartSpace(
                chartType, chartTitle, categories, seriesData, properties);
            var extChartPart = chartMainPart.AddNewPart<ExtendedChartPart>();
            extChartPart.ChartSpace = cxChartSpace;
            extChartPart.ChartSpace.Save();

            var cxRelId = chartMainPart.GetIdOfPart(extChartPart);
            var cxChartRef = new DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId { Id = cxRelId };

            var cxInline = new DW.Inline(
                new DW.Extent { Cx = chartCx, Cy = chartCy },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = docPropId, Name = chartName },
                new DW.NonVisualGraphicFrameDrawingProperties(),
                new A.Graphic(
                    new A.GraphicData(cxChartRef)
                    { Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            };

            var cxRun = new Run(new Drawing(cxInline));
            Paragraph cxPara;
            if (parent is Paragraph existingCxPara)
            {
                existingCxPara.AppendChild(cxRun);
                cxPara = existingCxPara;
            }
            else
            {
                cxPara = new Paragraph(cxRun);
                AssignParaId(cxPara);
                AppendToParent(parent, cxPara);
            }

            var totalCharts = CountWordCharts(chartMainPart);
            return $"/chart[{totalCharts}]";
        }

        // Create ChartPart and build chart
        var chartPart = chartMainPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = Core.ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);

        // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
        // Must be called BEFORE Save() so the in-memory DOM is still available
        var deferredProps = properties
            .Where(kv => Core.ChartHelper.DeferredAddKeys.Contains(kv.Key))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        if (deferredProps.Count > 0)
            Core.ChartHelper.SetChartProperties(chartPart, deferredProps);
        else
            chartPart.ChartSpace.Save();

        var chartRelId = chartMainPart.GetIdOfPart(chartPart);

        // Build Drawing/Inline with ChartReference
        var inline = new DW.Inline(
            new DW.Extent { Cx = chartCx, Cy = chartCy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = docPropId, Name = chartName },
            new DW.NonVisualGraphicFrameDrawingProperties(),
            new A.Graphic(
                new A.GraphicData(
                    new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = chartRelId }
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
            )
        )
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        var chartRun = new Run(new Drawing(inline));
        Paragraph chartPara;
        if (parent is Paragraph existingChartPara)
        {
            existingChartPara.AppendChild(chartRun);
            chartPara = existingChartPara;
        }
        else
        {
            chartPara = new Paragraph(chartRun);
            AssignParaId(chartPara);
            AppendToParent(parent, chartPara);
        }

        var totalChartIdx = CountWordCharts(chartMainPart);
        return $"/chart[{totalChartIdx}]";
    }

    private string AddPicture(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("path", out var imgPath) && !properties.TryGetValue("src", out imgPath))
            throw new ArgumentException("'path' (or 'src') property is required for picture type");

        var (imgStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
        using var imgStreamDispose = imgStream;

        var mainPart = _doc.MainDocumentPart!;
        var imagePart = mainPart.AddImagePart(imgPartType);
        imagePart.FeedData(imgStream);
        var relId = mainPart.GetIdOfPart(imagePart);

        // Determine dimensions with auto aspect ratio
        bool hasExplicitWidth = properties.TryGetValue("width", out var widthStr);
        bool hasExplicitHeight = properties.TryGetValue("height", out var heightStr);
        long cxEmu = hasExplicitWidth ? ParseEmu(widthStr!) : 5486400; // default: 6 inches
        long cyEmu = hasExplicitHeight ? ParseEmu(heightStr!) : 3657600; // default: 4 inches

        // Auto-calculate missing dimension from image pixel aspect ratio
        if (!hasExplicitWidth || !hasExplicitHeight)
        {
            var dims = OfficeCli.Core.ImageSource.TryGetDimensions(imgStream);
            if (dims is { Width: > 0, Height: > 0 })
            {
                var (pixW, pixH) = dims.Value;
                if (hasExplicitWidth)
                    cyEmu = (long)(cxEmu * (double)pixH / pixW);
                else if (hasExplicitHeight)
                    cxEmu = (long)(cyEmu * (double)pixW / pixH);
                else
                    cyEmu = (long)(cxEmu * (double)pixH / pixW);
            }
        }

        var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

        var imgDocPropId = NextDocPropId();
        Run imgRun;
        if (properties.TryGetValue("anchor", out var anchorVal) && IsTruthy(anchorVal))
        {
            var wrapType = properties.GetValueOrDefault("wrap", "none");
            long hPos = properties.TryGetValue("hposition", out var hPosStr) ? ParseEmu(hPosStr) : 0;
            long vPos = properties.TryGetValue("vposition", out var vPosStr) ? ParseEmu(vPosStr) : 0;
            var hRel = properties.TryGetValue("hrelative", out var hRelStr)
                ? ParseHorizontalRelative(hRelStr)
                : DW.HorizontalRelativePositionValues.Margin;
            var vRel = properties.TryGetValue("vrelative", out var vRelStr)
                ? ParseVerticalRelative(vRelStr)
                : DW.VerticalRelativePositionValues.Margin;
            var behind = properties.TryGetValue("behindtext", out var behindStr) && IsTruthy(behindStr);
            imgRun = CreateAnchorImageRun(relId, cxEmu, cyEmu, altText, wrapType, hPos, vPos, hRel, vRel, behind, imgDocPropId);
        }
        else
        {
            imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText, imgDocPropId);
        }

        string resultPath;
        Paragraph imgPara;
        if (parent is Paragraph existingPara)
        {
            var runCount = existingPara.Elements<Run>().Count();
            if (index.HasValue && index.Value < runCount)
            {
                var refRun = existingPara.Elements<Run>().ElementAt(index.Value);
                existingPara.InsertBefore(imgRun, refRun);
            }
            else
            {
                existingPara.AppendChild(imgRun);
            }
            imgPara = existingPara;
            var imgRunIdx = existingPara.Elements<Run>().ToList().IndexOf(imgRun) + 1;
            resultPath = $"{parentPath}/r[{imgRunIdx}]";
        }
        else if (parent is TableCell imgCell)
        {
            // Insert image into existing first paragraph if empty, otherwise create new paragraph
            var firstCellPara = imgCell.Elements<Paragraph>().FirstOrDefault();
            if (firstCellPara != null && !firstCellPara.Elements<Run>().Any())
            {
                firstCellPara.AppendChild(imgRun);
                imgPara = firstCellPara;
            }
            else
            {
                imgPara = new Paragraph(imgRun);
                AssignParaId(imgPara);
                imgPara.PrependChild(new ParagraphProperties(
                    new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }));
                imgCell.AppendChild(imgPara);
            }
            var imgPIdx = imgCell.Elements<Paragraph>().ToList().IndexOf(imgPara) + 1;
            resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
        }
        else
        {
            imgPara = new Paragraph(imgRun);
            AssignParaId(imgPara);
            // Prevent fixed line spacing (inherited from Normal style) from clipping the image
            imgPara.PrependChild(new ParagraphProperties(
                new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }));
            // Use ChildElements for index lookup to match ResolveAnchorPosition
            // which computes indices against ChildElements (not just Paragraphs)
            var allChildren = parent.ChildElements.ToList();
            if (index.HasValue && index.Value < allChildren.Count)
            {
                var refElement = allChildren[index.Value];
                parent.InsertBefore(imgPara, refElement);
                var imgPIdx = parent.Elements<Paragraph>().ToList().IndexOf(imgPara) + 1;
                resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
            }
            else
            {
                AppendToParent(parent, imgPara);
                var imgPIdx = parent.Elements<Paragraph>().Count();
                resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
            }
        }
        return resultPath;
    }
}
