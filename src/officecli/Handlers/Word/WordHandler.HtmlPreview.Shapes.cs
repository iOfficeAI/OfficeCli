// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Drawing with Overlaid Images ====================

    private void RenderDrawingWithOverlaidImages(StringBuilder sb, Drawing groupDrawing, List<Drawing> overlaidImages)
    {
        if (overlaidImages.Count == 0)
        {
            RenderDrawingHtml(sb, groupDrawing, null);
            return;
        }

        RenderDrawingHtml(sb, groupDrawing, overlaidImages);
    }

    // ==================== Drawing Rendering (images, groups, shapes) ====================

    /// <summary>Check if a paragraph contains drawings with actual text box content (txbxContent).</summary>
    private static bool HasTextBoxContent(Paragraph para)
    {
        foreach (var run in para.Elements<Run>())
        {
            var drawing = run.GetFirstChild<Drawing>() ?? run.Descendants<Drawing>().FirstOrDefault();
            if (drawing != null && HasTextBox(drawing))
                return true;
        }
        return false;
    }

    /// <summary>Check if paragraph contains any drawing that renders as block-level HTML (text box, chart, shape).</summary>
    private static bool HasBlockLevelDrawing(Paragraph para)
    {
        // Check all descendants (including inside mc:AlternateContent)
        foreach (var drawing in para.Descendants<Drawing>())
        {
            if (HasGroupOrShape(drawing)) return true;
            if (drawing.Descendants().Any(e => e.LocalName == "chart")) return true;
        }
        // Also check for text box content via localName (catches mc:AlternateContent cases)
        if (para.Descendants().Any(e => e.LocalName == "txbxContent"))
            return true;
        return false;
    }

    /// <summary>Find VML horizontal rule shape in a paragraph (w:pict > v:rect/v:line with o:hr="t").</summary>
    private static OpenXmlElement? FindVmlHorizontalRule(Paragraph para)
    {
        // Search all descendants to handle both direct w:pict and mc:AlternateContent wrapping
        foreach (var pict in para.Descendants().Where(e => e.LocalName == "pict"))
        {
            var hrShape = pict.ChildElements.FirstOrDefault(c =>
                (c.LocalName == "rect" || c.LocalName == "line") &&
                c.GetAttributes().Any(a => a.LocalName == "hr" && a.Value == "t"));
            if (hrShape != null) return hrShape;
        }
        return null;
    }

    /// <summary>Check if a paragraph contains a VML horizontal rule.</summary>
    private static bool IsVmlHorizontalRule(Paragraph para) => FindVmlHorizontalRule(para) != null;

    /// <summary>Render a VML horizontal rule as an HTML hr element.</summary>
    private static void RenderVmlHorizontalRule(StringBuilder sb, Paragraph para)
    {
        var shape = FindVmlHorizontalRule(para)!;

        // Color from fillcolor attribute
        var fillColor = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "fillcolor").Value ?? "#a0a0a0";
        if (!fillColor.StartsWith("#")) fillColor = "#" + fillColor;

        // Height from VML style (e.g. style="width:0;height:1.5pt")
        var heightPx = 1.5;
        var vmlStyle = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value;
        if (vmlStyle != null)
        {
            var hMatch = System.Text.RegularExpressions.Regex.Match(vmlStyle, @"height:\s*([\d.]+)pt");
            if (hMatch.Success && double.TryParse(hMatch.Groups[1].Value, out var hPt))
                heightPx = hPt;
        }

        // Width percentage from o:hrpct (value in tenths of a percent, e.g. 1000 = 100%)
        var widthCss = "100%";
        var hrpct = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "hrpct").Value;
        if (hrpct != null && int.TryParse(hrpct, out var pctVal) && pctVal > 0 && pctVal < 1000)
            widthCss = $"{pctVal / 10.0:0.#}%";

        // Alignment from o:hralign
        var align = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "hralign").Value ?? "center";
        var marginCss = align switch
        {
            "left" => "margin:0.5em auto 0.5em 0",
            "right" => "margin:0.5em 0 0.5em auto",
            _ => "margin:0.5em auto"
        };

        sb.AppendLine($"<hr style=\"border:none;border-top:{heightPx:0.#}px solid {fillColor};width:{widthCss};{marginCss}\">");
    }

    /// <summary>Check if a drawing contains groups or shapes (for rendering).</summary>
    private static bool HasGroupOrShape(Drawing drawing)
    {
        return drawing.Descendants().Any(e => e.LocalName == "wgp" || e.LocalName == "wsp");
    }

    /// <summary>Check if a drawing contains actual text box content with text (not empty decorative shapes).</summary>
    private static bool HasTextBox(Drawing drawing)
    {
        foreach (var txbx in drawing.Descendants().Where(e => e.LocalName == "txbxContent"))
        {
            // Check if any paragraph inside has actual text
            if (txbx.Descendants<Text>().Any(t => !string.IsNullOrWhiteSpace(t.Text)))
                return true;
        }
        return false;
    }

    private void RenderDrawingHtml(StringBuilder sb, Drawing drawing, List<Drawing>? floatImages = null)
    {
        // Check for chart (c:chart inside a:graphicData)
        var chartRef = drawing.Descendants().FirstOrDefault(e => e.LocalName == "chart" &&
            e.GetAttributes().Any(a => a.LocalName == "id"));
        if (chartRef != null)
        {
            RenderChartHtml(sb, drawing, chartRef);
            return;
        }

        // Check for groups/shapes first (text boxes, decorated shapes)
        var group = drawing.Descendants().FirstOrDefault(e => e.LocalName == "wgp");
        if (group != null)
        {
            // Get overall extent from wp:inline or wp:anchor
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            long groupWidthEmu = extent?.Cx?.Value ?? 0;
            long groupHeightEmu = extent?.Cy?.Value ?? 0;

            if (groupWidthEmu > 0 && groupHeightEmu > 0)
            {
                RenderGroupHtml(sb, group, groupWidthEmu, groupHeightEmu, floatImages);
                return;
            }
        }

        // Check for standalone shape (wsp without group)
        var shape = drawing.Descendants().FirstOrDefault(e => e.LocalName == "wsp");
        if (shape != null)
        {
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            long shapeWidth = extent?.Cx?.Value ?? 0;
            long shapeHeight = extent?.Cy?.Value ?? 0;
            if (shapeWidth > 0 && shapeHeight > 0)
            {
                // Full-page shapes → render as background layer
                if (IsFullPageSize(shapeWidth, shapeHeight))
                {
                    var fillCss = ResolveShapeFillCss(shape.Elements().FirstOrDefault(e => e.LocalName == "spPr"));
                    if (!string.IsNullOrEmpty(fillCss))
                        sb.Append($"<div style=\"position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;{fillCss}\"></div>");
                    return;
                }
                // Standalone shape — render as inline block, not absolute positioned
                RenderStandaloneShapeHtml(sb, shape, shapeWidth, shapeHeight, floatImages);
                return;
            }
        }

        // Fall back to image rendering
        RenderImageHtml(sb, drawing);
    }

    private void RenderImageHtml(StringBuilder sb, Drawing drawing)
    {
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value == null) return;

        // Prefer the SVG extension rel if present (Office 2019+ keeps a PNG
        // raster in Embed plus an SVG via a:extLst/asvg:svgBlip). PNG fallback
        // is often a 1×1 transparent pixel that renders as a blank, so SVG
        // wins for modern documents that embed vector art.
        string blipRelId = blip.Embed.Value;
        var svgBlip = blip.Descendants().FirstOrDefault(e => e.LocalName == "svgBlip");
        if (svgBlip != null)
        {
            var svgRel = svgBlip.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "embed" || a.LocalName == "link").Value;
            if (!string.IsNullOrEmpty(svgRel))
                blipRelId = svgRel;
        }
        var dataUri = LoadImageAsDataUri(blipRelId);
        if (dataUri == null) return;

        try
        {

            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault()
                ?? drawing.Descendants<A.Extents>().FirstOrDefault() as OpenXmlElement;
            long imgCxEmu = 0, imgCyEmu = 0;
            if (extent is DW.Extent dwExt) { imgCxEmu = dwExt.Cx?.Value ?? 0; imgCyEmu = dwExt.Cy?.Value ?? 0; }
            else if (extent is A.Extents aExt) { imgCxEmu = aExt.Cx?.Value ?? 0; imgCyEmu = aExt.Cy?.Value ?? 0; }

            var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
            var alt = docProps?.Description?.Value ?? docProps?.Name?.Value ?? "image";

            // Detect full-page background images → render as absolute background
            if (IsFullPageSize(imgCxEmu, imgCyEmu))
            {
                sb.Append($"<div style=\"position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;overflow:hidden\">");
                sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncodeAttr(alt)}\" style=\"width:100%;height:100%;object-fit:cover\">");
                sb.Append("</div>");
                return;
            }

            var widthPx = imgCxEmu / 9525;
            var heightPx = imgCyEmu / 9525;
            string widthAttr = widthPx > 0 ? $" width=\"{widthPx}\"" : "";
            string heightAttr = heightPx > 0 ? $" height=\"{heightPx}\"" : "";

            // Detect anchored/floating positioning
            var anchor = drawing.Descendants<DW.Anchor>().FirstOrDefault();
            var floatCss = "";
            if (anchor != null)
            {
                var hPos = anchor.GetFirstChild<DW.HorizontalPosition>();
                var hAlign = hPos?.Descendants().FirstOrDefault(e => e.LocalName == "align")?.InnerText;
                var hPosFrom = hPos?.RelativeFrom?.Value;

                // wrapTopAndBottom → centered block image (no text beside it)
                var wrapTopBottom = anchor.Elements().Any(e => e.LocalName == "wrapTopAndBottom");
                if (wrapTopBottom)
                {
                    floatCss = "display:block;margin:8px auto";
                }
                // wrapSquare / wrapTight → float left or right
                else if (anchor.Elements().Any(e => e.LocalName == "wrapSquare" || e.LocalName == "wrapTight"))
                {
                    var isRight = hAlign == "right"
                        || hPosFrom == DW.HorizontalRelativePositionValues.RightMargin;
                    // Also check posOffset — if offset > half page width, float right
                    if (!isRight && hPos != null)
                    {
                        var offsetEl = hPos.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
                        if (offsetEl != null && long.TryParse(offsetEl.InnerText, out var offsetEmu))
                        {
                            var halfPageEmu = (long)(GetPageLayout().WidthPt * 12700); // pt to EMU
                            isRight = offsetEmu > halfPageEmu;
                        }
                    }
                    // #7b: use the anchor's distT/distB/distL/distR for the
                    // float margin instead of a hardcoded 8px. The emu→pt
                    // conversion keeps spacing in line with what Word paints.
                    var distT = (long)(anchor.DistanceFromTop?.Value ?? 0) / 12700.0;
                    var distB = (long)(anchor.DistanceFromBottom?.Value ?? 0) / 12700.0;
                    var distL = (long)(anchor.DistanceFromLeft?.Value ?? 0) / 12700.0;
                    var distR = (long)(anchor.DistanceFromRight?.Value ?? 0) / 12700.0;
                    // Floor the "inside" margin (right for float:left, left for
                    // float:right) so text always has breathing room.
                    if (isRight)
                    {
                        if (distL < 6) distL = 6;
                    }
                    else
                    {
                        if (distR < 6) distR = 6;
                    }
                    floatCss = isRight
                        ? $"float:right;margin:{distT:0.#}pt {distR:0.#}pt {distB:0.#}pt {distL:0.#}pt"
                        : $"float:left;margin:{distT:0.#}pt {distR:0.#}pt {distB:0.#}pt {distL:0.#}pt";

                    // Anchored at top of margin — emit marker for relocation to page start
                    var vPos = anchor.GetFirstChild<DW.VerticalPosition>();
                    var vAlign = vPos?.Descendants().FirstOrDefault(e => e.LocalName == "align")?.InnerText;
                    var vFrom = vPos?.RelativeFrom?.Value;
                    if (vAlign == "top" && vFrom == DW.VerticalRelativePositionValues.Margin)
                    {
                        var fc = isRight ? "float:right;margin:0 0 8px 8px" : "float:left;margin:0 8px 8px 0";
                        var cropVal = GetCropPercents(drawing);
                        var imgHtml = new StringBuilder();
                        if (cropVal.HasValue)
                            RenderCroppedImage(imgHtml, dataUri, widthPx, heightPx, cropVal.Value.l, cropVal.Value.t, cropVal.Value.r, cropVal.Value.b, HtmlEncodeAttr(alt), fc);
                        else
                            imgHtml.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncodeAttr(alt)}\" width=\"{widthPx}\" height=\"{heightPx}\" style=\"max-width:100%;height:auto;{fc}\">");
                        var markerId = $"TOP_ANCHOR_{_ctx.TopAnchoredImages.Count}";
                        _ctx.TopAnchoredImages.Add((markerId, imgHtml.ToString()));
                        sb.Append($"<!--{markerId}-->");
                        return;
                    }
                }
            }

            // Crop support: container-based cropping
            var crop = GetCropPercents(drawing);
            // #7a001: when the image's native width exceeds the page body's
            // content width, drop `max-width:100%` so the image paints at
            // native size and overflows the margin the way Word does.
            // Otherwise `max-width:100%` + explicit width + flex-column parent
            // can collapse the layout slot to zero.
            var pgLayout = GetPageLayout();
            var contentWidthPt = pgLayout.WidthPt - pgLayout.MarginLeftPt - pgLayout.MarginRightPt;
            var imgWidthPt = widthPx * 72.0 / 96.0; // 96 DPI → pt
            var overflows = widthPx > 0 && imgWidthPt > contentWidthPt;
            var styleParts = overflows
                ? new List<string> { $"width:{imgWidthPt:0.#}pt", "height:auto" }
                : new List<string> { "max-width:100%", "height:auto" };
            if (!string.IsNullOrEmpty(floatCss)) styleParts.Add(floatCss);

            // Picture effects from pic:spPr — rotation, flip, border, shadow
            var spPr = drawing.Descendants().FirstOrDefault(e => e.LocalName == "spPr");
            var effectCss = spPr != null ? GetPictureEffectsCss(spPr) : "";
            if (!string.IsNullOrEmpty(effectCss)) styleParts.Add(effectCss);

            if (crop.HasValue)
            {
                RenderCroppedImage(sb, dataUri, widthPx, heightPx, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, HtmlEncodeAttr(alt), floatCss + (string.IsNullOrEmpty(effectCss) ? "" : ";" + effectCss));
            }
            else
            {
                sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncodeAttr(alt)}\"{widthAttr}{heightAttr} style=\"{string.Join(";", styleParts)}\">");
            }
        }
        catch
        {
            sb.Append("<span class=\"img-error\">[Image]</span>");
        }
    }

    /// <summary>
    /// Extract CSS for picture visual effects from a:xfrm (rotation, flip),
    /// a:ln (border), and a:effectLst (shadow/glow). All live under pic:spPr.
    /// </summary>
    private static string GetPictureEffectsCss(OpenXmlElement spPr)
    {
        var parts = new List<string>();

        // Rotation + flip from a:xfrm
        var xfrm = spPr.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
        if (xfrm != null)
        {
            var rot = xfrm.GetAttributes().FirstOrDefault(a => a.LocalName == "rot").Value;
            var flipH = xfrm.GetAttributes().FirstOrDefault(a => a.LocalName == "flipH").Value;
            var flipV = xfrm.GetAttributes().FirstOrDefault(a => a.LocalName == "flipV").Value;

            var transforms = new List<string>();
            if (long.TryParse(rot, out var rotVal) && rotVal != 0)
            {
                // OOXML rotation is in 60000ths of a degree
                var deg = rotVal / 60000.0;
                transforms.Add($"rotate({deg:0.##}deg)");
            }
            if (flipH == "1" || flipH == "true") transforms.Add("scaleX(-1)");
            if (flipV == "1" || flipV == "true") transforms.Add("scaleY(-1)");
            if (transforms.Count > 0)
                parts.Add($"transform:{string.Join(" ", transforms)}");
        }

        // Border from a:ln
        var ln = spPr.Elements().FirstOrDefault(e => e.LocalName == "ln");
        if (ln != null)
        {
            var wAttr = ln.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
            double borderPx = 1;
            if (long.TryParse(wAttr, out var wEmu) && wEmu > 0)
                borderPx = Math.Max(1, wEmu / 9525.0); // EMU → px
            var solidFill = ln.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
            var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
            var colorHex = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            var borderColor = !string.IsNullOrEmpty(colorHex) ? $"#{colorHex}" : "#000";
            parts.Add($"border:{borderPx:0.##}px solid {borderColor}");
        }

        // Outer shadow from a:effectLst/a:outerShdw — map to box-shadow
        var effectLst = spPr.Elements().FirstOrDefault(e => e.LocalName == "effectLst");
        var outerShdw = effectLst?.Elements().FirstOrDefault(e => e.LocalName == "outerShdw");
        if (outerShdw != null)
        {
            // blurRad, dist, dir (60000ths of a degree) — simplified offset projection
            var blurAttr = outerShdw.GetAttributes().FirstOrDefault(a => a.LocalName == "blurRad").Value;
            var distAttr = outerShdw.GetAttributes().FirstOrDefault(a => a.LocalName == "dist").Value;
            var dirAttr = outerShdw.GetAttributes().FirstOrDefault(a => a.LocalName == "dir").Value;
            double blurPx = long.TryParse(blurAttr, out var blurEmu) ? blurEmu / 9525.0 : 4;
            double distPx = long.TryParse(distAttr, out var distEmu) ? distEmu / 9525.0 : 4;
            double dirDeg = long.TryParse(dirAttr, out var dirVal) ? dirVal / 60000.0 : 45;
            var offX = distPx * Math.Cos(dirDeg * Math.PI / 180);
            var offY = distPx * Math.Sin(dirDeg * Math.PI / 180);
            var shdwFill = outerShdw.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
            var shdwHex = shdwFill?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value ?? "000000";
            parts.Add($"box-shadow:{offX:0.#}px {offY:0.#}px {blurPx:0.#}px #{shdwHex}");
        }

        return string.Join(";", parts);
    }

    /// <summary>
    /// Get crop percentages from a:srcRect.
    /// Values are in 1/1000 of a percent (e.g., 25000 = 25%).
    /// Negative values mean extend (treated as 0).
    /// Returns (left, top, right, bottom) as CSS percentages, or null if no crop.
    /// </summary>
    private static (double l, double t, double r, double b)? GetCropPercents(OpenXmlElement container)
    {
        var srcRect = container.Descendants().FirstOrDefault(e => e.LocalName == "srcRect");
        if (srcRect == null) return null;

        var l = Math.Max(0, GetIntAttr(srcRect, "l") / 1000.0);
        var t = Math.Max(0, GetIntAttr(srcRect, "t") / 1000.0);
        var r = Math.Max(0, GetIntAttr(srcRect, "r") / 1000.0);
        var b = Math.Max(0, GetIntAttr(srcRect, "b") / 1000.0);

        if (l == 0 && t == 0 && r == 0 && b == 0) return null;
        return (l, t, r, b);
    }

    /// <summary>
    /// Render a cropped image using a container div with overflow:hidden.
    /// The image is scaled to its original size and positioned to show only the cropped region.
    /// </summary>
    private static void RenderCroppedImage(StringBuilder sb, string dataUri, long displayWidthPx, long displayHeightPx,
        double cropL, double cropT, double cropR, double cropB, string alt, string extraStyle = "")
    {
        // The display size is the cropped result size.
        // Original image visible fraction: (1 - cropL/100 - cropR/100) horizontally, (1 - cropT/100 - cropB/100) vertically.
        var fracW = 1.0 - cropL / 100.0 - cropR / 100.0;
        var fracH = 1.0 - cropT / 100.0 - cropB / 100.0;
        if (fracW <= 0) fracW = 1; if (fracH <= 0) fracH = 1;

        // Original image size in CSS
        var imgW = displayWidthPx / fracW;
        var imgH = displayHeightPx / fracH;
        // Offset to show the cropped region
        var offsetX = -imgW * (cropL / 100.0);
        var offsetY = -imgH * (cropT / 100.0);

        var containerStyle = $"display:inline-block;width:{displayWidthPx}px;height:{displayHeightPx}px;overflow:hidden";
        if (!string.IsNullOrEmpty(extraStyle)) containerStyle += $";{extraStyle}";
        sb.Append($"<div style=\"{containerStyle}\">");
        sb.Append($"<img src=\"{dataUri}\" alt=\"{alt}\" style=\"width:{imgW:0}px;height:{imgH:0}px;margin-left:{offsetX:0}px;margin-top:{offsetY:0}px\">");
        sb.Append("</div>");
    }

    private static int GetIntAttr(OpenXmlElement el, string attrName)
    {
        var val = el.GetAttributes().FirstOrDefault(a => a.LocalName == attrName).Value;
        return val != null && int.TryParse(val, out var v) ? v : 0;
    }

    /// <summary>Load an image part by relationship ID and return as a base64 data URI.</summary>
    private string? LoadImageAsDataUri(string relId)
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return null;
        return HtmlPreviewHelper.PartToDataUri(mainPart, relId);
    }

    // ==================== Group / Shape Rendering ====================

    private void RenderGroupHtml(StringBuilder sb, OpenXmlElement group, long groupWidthEmu, long groupHeightEmu,
        List<Drawing>? floatImages = null)
    {
        var widthPx = groupWidthEmu / 9525;
        var heightPx = groupHeightEmu / 9525;

        // Get the group's child coordinate space from grpSpPr > xfrm
        long chOffX = 0, chOffY = 0, chExtCx = groupWidthEmu, chExtCy = groupHeightEmu;
        var grpSpPr = group.Elements().FirstOrDefault(e => e.LocalName == "grpSpPr");
        var grpXfrm = grpSpPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
        if (grpXfrm != null)
        {
            var chOff = grpXfrm.Elements().FirstOrDefault(e => e.LocalName == "chOff");
            var chExt = grpXfrm.Elements().FirstOrDefault(e => e.LocalName == "chExt");
            if (chOff != null)
            {
                chOffX = GetLongAttr(chOff, "x");
                chOffY = GetLongAttr(chOff, "y");
            }
            if (chExt != null)
            {
                chExtCx = GetLongAttr(chExt, "cx");
                chExtCy = GetLongAttr(chExt, "cy");
            }
        }

        sb.Append($"<div class=\"wg\" style=\"position:relative;width:{widthPx}px;height:{heightPx}px;display:inline-block;overflow:hidden\">");

        // Render each child element (shapes, pictures, nested groups)
        foreach (var child in group.Elements())
        {
            if (child.LocalName is "wsp" or "pic" or "grpSp")
            {
                // Get transform from xfrm (may be in spPr or grpSpPr)
                var xfrm = child.Descendants().FirstOrDefault(e => e.LocalName == "xfrm");
                long offX = 0, offY = 0, extCx = 0, extCy = 0;
                if (xfrm != null)
                {
                    var off = xfrm.Elements().FirstOrDefault(e => e.LocalName == "off");
                    var ext = xfrm.Elements().FirstOrDefault(e => e.LocalName == "ext");
                    if (off != null) { offX = GetLongAttr(off, "x"); offY = GetLongAttr(off, "y"); }
                    if (ext != null) { extCx = GetLongAttr(ext, "cx"); extCy = GetLongAttr(ext, "cy"); }
                }

                // Pass floatImages to first text box shape, then clear
                RenderShapeHtml(sb, child, offX - chOffX, offY - chOffY, extCx, extCy, chExtCx, chExtCy, floatImages);
                floatImages = null; // only inject into first shape
            }
        }

        sb.Append("</div>");
    }

    private void RenderStandaloneShapeHtml(StringBuilder sb, OpenXmlElement shape, long widthEmu, long heightEmu,
        List<Drawing>? floatImages)
    {
        // Standalone shapes use inline positioning with pixel dimensions
        RenderShapeHtml(sb, shape, 0, 0, widthEmu, heightEmu, widthEmu, heightEmu, floatImages, standalone: true);
    }

    /// <summary>
    /// Render a shape element (wsp, pic, grpSp) with either absolute (inside group) or inline (standalone) positioning.
    /// </summary>
    private void RenderShapeHtml(StringBuilder sb, OpenXmlElement shape, long offX, long offY,
        long extCx, long extCy, long coordSpaceCx, long coordSpaceCy,
        List<Drawing>? floatImages = null, bool standalone = false)
    {
        // Common shape properties
        var spPr = shape.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        var fillCss = ResolveShapeFillCss(spPr);
        var borderCss = ResolveShapeBorderCss(spPr);
        var txbx = shape.LocalName == "pic" ? null
            : shape.Descendants().FirstOrDefault(e => e.LocalName == "txbxContent");

        // Build positioning style
        string style;
        if (standalone)
        {
            var widthPx = extCx / 9525;
            var heightPx = extCy / 9525;
            style = $"display:inline-block;width:{widthPx}px;min-height:{heightPx}px;vertical-align:top";

            // Rotation on standalone shapes too (was only applied inside groups)
            var sXfrm = spPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
            var sRot = GetLongAttr(sXfrm, "rot");
            if (sRot != 0) style += $";transform:rotate({sRot / 60000.0:0.##}deg)";
        }
        else
        {
            double leftPct = coordSpaceCx > 0 ? (double)offX / coordSpaceCx * 100 : 0;
            double topPct = coordSpaceCy > 0 ? (double)offY / coordSpaceCy * 100 : 0;
            double widthPct = coordSpaceCx > 0 ? (double)extCx / coordSpaceCx * 100 : 100;
            double heightPct = coordSpaceCy > 0 ? (double)extCy / coordSpaceCy * 100 : 100;
            style = $"position:absolute;left:{leftPct:0.##}%;top:{topPct:0.##}%;width:{widthPct:0.##}%;height:{heightPct:0.##}%";

            // Rotation (only for positioned shapes inside groups)
            var xfrm = spPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
            var rot = GetLongAttr(xfrm, "rot");
            if (rot != 0) style += $";transform:rotate({rot / 60000.0:0.##}deg)";
        }

        // prstGeom → border-radius for ellipse, round rect, etc.
        var prstGeom = spPr?.Elements().FirstOrDefault(e => e.LocalName == "prstGeom");
        var prst = prstGeom?.GetAttributes().FirstOrDefault(a => a.LocalName == "prst").Value;
        if (prst == "ellipse" || prst == "oval")
            style += ";border-radius:50%";
        else if (prst == "roundRect")
            style += ";border-radius:12px";

        // #7a: for complex preset geometries (line, arrows, callouts) the
        // background/border approach collapses to a plain rect. Render
        // those as inline SVG overlays using the shape's fill/border colors.
        var svgPrst = prst is "line" or "straightConnector1"
            or "rightArrow" or "leftArrow" or "upArrow" or "downArrow"
            or "wedgeRoundRectCallout";
        if (svgPrst)
        {
            // Defer fill/border to the SVG so the host div stays transparent.
            style += ";overflow:visible";
        }
        else
        {
            if (!string.IsNullOrEmpty(fillCss)) style += $";{fillCss}";
            if (!string.IsNullOrEmpty(borderCss)) style += $";{borderCss}";
        }

        // Body properties: text layout + padding
        var bodyPr = shape.Elements().FirstOrDefault(e => e.LocalName == "bodyPr");
        // Vertical text anchor applies to both standalone and positioned shapes
        var vAnchor = bodyPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "anchor").Value;
        if (vAnchor == "ctr") style += ";display:flex;align-items:center";
        else if (vAnchor == "b") style += ";display:flex;align-items:flex-end";

        var lIns = GetLongAttr(bodyPr, "lIns", 91440);
        var tIns = GetLongAttr(bodyPr, "tIns", 45720);
        var rIns = GetLongAttr(bodyPr, "rIns", 91440);
        var bIns = GetLongAttr(bodyPr, "bIns", 45720);
        style += $";padding:{tIns / 9525}px {rIns / 9525}px {bIns / 9525}px {lIns / 9525}px";

        sb.Append($"<div style=\"{style}\">");

        // #7a: paint the geometry via inline SVG overlay when the preset
        // needs real polygon/path geometry (line, arrows, callouts).
        if (svgPrst)
        {
            var svgFill = ExtractCssColor(fillCss, "background-color")
                ?? ExtractFirstGradientColor(fillCss)
                ?? "transparent";
            var (borderColor, borderWidth) = ExtractBorderParts(borderCss);
            RenderPrstGeomSvg(sb, prst!, svgFill, borderColor ?? "#000", borderWidth ?? 1);
        }

        if (txbx != null)
        {
            // Render text box content (standard Word paragraphs)
            sb.Append("<div style=\"width:100%\">");

            // Inject pending float images into this text box
            if (floatImages != null && floatImages.Count > 0)
            {
                foreach (var imgDrawing in floatImages)
                {
                    var imgBlip = imgDrawing.Descendants<A.Blip>().FirstOrDefault();
                    if (imgBlip?.Embed?.Value == null) continue;
                    var imgDataUri = LoadImageAsDataUri(imgBlip.Embed.Value);
                    if (imgDataUri == null) continue;
                    try
                    {
                        var imgExtent = imgDrawing.Descendants<DW.Extent>().FirstOrDefault();
                        var imgW = imgExtent?.Cx?.Value > 0 ? imgExtent.Cx.Value / 9525 : 100;
                        var imgH = imgExtent?.Cy?.Value > 0 ? imgExtent.Cy.Value / 9525 : 100;
                        // Read distT/distB/distL/distR for image margins (EMU)
                        var inline = imgDrawing.Descendants<DW.Inline>().FirstOrDefault();
                        var anchor = imgDrawing.Descendants<DW.Anchor>().FirstOrDefault();
                        long distT = 0, distB = 0, distL = 0, distR = 0;
                        if (inline != null)
                        {
                            distT = (long)(inline.DistanceFromTop?.Value ?? 0);
                            distB = (long)(inline.DistanceFromBottom?.Value ?? 0);
                            distL = (long)(inline.DistanceFromLeft?.Value ?? 0);
                            distR = (long)(inline.DistanceFromRight?.Value ?? 0);
                        }
                        else if (anchor != null)
                        {
                            distT = (long)(anchor.DistanceFromTop?.Value ?? 0);
                            distB = (long)(anchor.DistanceFromBottom?.Value ?? 0);
                            distL = (long)(anchor.DistanceFromLeft?.Value ?? 0);
                            distR = (long)(anchor.DistanceFromRight?.Value ?? 0);
                        }
                        var marginCss = $"margin:{distT/9525}px {distR/9525}px {distB/9525}px {distL/9525}px";
                        var crop = GetCropPercents(imgDrawing);
                        if (crop.HasValue)
                        {
                            sb.Append($"<div style=\"float:left;{marginCss}\">");
                            RenderCroppedImage(sb, imgDataUri, imgW, imgH, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, "");
                            sb.Append("</div>");
                        }
                        else
                        {
                            sb.Append($"<img src=\"{imgDataUri}\" style=\"float:left;width:{imgW}px;height:{imgH}px;object-fit:cover;{marginCss}\">");
                        }
                    }
                    catch { }
                }
                floatImages = null;
            }

            foreach (var para in txbx.Descendants<Paragraph>())
            {
                RenderParagraphHtml(sb, para);
            }
            sb.Append("</div>");
        }
        else
        {
            // Check for image inside shape
            var embedAttr = FindEmbedInDescendants(shape);
            if (embedAttr != null)
            {
                var dataUri = LoadImageAsDataUri(embedAttr);
                if (dataUri != null)
                    sb.Append($"<img src=\"{dataUri}\" style=\"width:100%;height:100%;object-fit:contain\">");
            }
        }

        sb.Append("</div>");
    }

    // ==================== #7a prstGeom SVG helpers ====================

    /// <summary>
    /// Pull a CSS property's color value out of strings like
    /// <c>background-color:#FF0000</c> or
    /// <c>background:linear-gradient(...)</c>. Returns null if not present.
    /// </summary>
    private static string? ExtractCssColor(string css, string prop)
    {
        if (string.IsNullOrEmpty(css)) return null;
        var m = System.Text.RegularExpressions.Regex.Match(
            css, $@"{prop}\s*:\s*(#[0-9A-Fa-f]{{3,8}}|[a-zA-Z]+)");
        return m.Success ? m.Groups[1].Value : null;
    }

    // Pull the first hex color out of a `background:linear-gradient(...)`
    // / `background-image:linear-gradient(...)` rule so SVG prstGeom shapes
    // don't degrade to transparent when only a gradient fill is available.
    private static string? ExtractFirstGradientColor(string css)
    {
        if (string.IsNullOrEmpty(css)) return null;
        if (css.IndexOf("gradient", StringComparison.OrdinalIgnoreCase) < 0) return null;
        var m = System.Text.RegularExpressions.Regex.Match(
            css, @"#[0-9A-Fa-f]{3,8}");
        return m.Success ? m.Value : null;
    }

    private static (string? color, double? width) ExtractBorderParts(string css)
    {
        if (string.IsNullOrEmpty(css)) return (null, null);
        // e.g. "border:1.5px solid #336699"
        var m = System.Text.RegularExpressions.Regex.Match(
            css, @"border\s*:\s*([\d.]+)px\s+\w+\s+(#[0-9A-Fa-f]{3,8}|[a-zA-Z]+)");
        if (!m.Success) return (null, null);
        return (m.Groups[2].Value,
            double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var w) ? w : 1);
    }

    /// <summary>
    /// Emit an inline SVG overlay rendering the given preset geometry.
    /// The SVG uses viewBox="0 0 100 100" and preserveAspectRatio="none"
    /// so it stretches to the host div's full size.
    /// </summary>
    private static void RenderPrstGeomSvg(
        StringBuilder sb, string prst, string fill, string stroke, double strokeW)
    {
        // Normalize stroke width to viewBox coordinates: at 100-unit viewBox
        // and typical host size ~150px, 1px ≈ 0.67 units. Keep as-is since
        // preserveAspectRatio=none scales X/Y differently anyway; ok for
        // approximation.
        // Display:block + width/height:100% makes the SVG fill the host
        // <div> without needing position:absolute (which would anchor to
        // the nearest positioned ancestor and cause all shapes on a page
        // to stack on top of each other).
        sb.Append(
            "<svg style=\"display:block;width:100%;height:100%;overflow:visible\" " +
            "viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\" xmlns=\"http://www.w3.org/2000/svg\">");
        var sw = strokeW.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        switch (prst)
        {
            case "line":
            case "straightConnector1":
                // Diagonal from top-left to bottom-right.
                sb.Append($"<line x1=\"0\" y1=\"0\" x2=\"100\" y2=\"100\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "rightArrow":
                // Classic block arrow pointing right: body 0..70, head 70..100.
                sb.Append($"<polygon points=\"0,30 70,30 70,10 100,50 70,90 70,70 0,70\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "leftArrow":
                sb.Append($"<polygon points=\"100,30 30,30 30,10 0,50 30,90 30,70 100,70\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "downArrow":
                sb.Append($"<polygon points=\"30,0 70,0 70,70 90,70 50,100 10,70 30,70\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "upArrow":
                sb.Append($"<polygon points=\"30,100 70,100 70,30 90,30 50,0 10,30 30,30\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "wedgeRoundRectCallout":
                // Rounded rect (80% height) + triangular pointer down-left.
                // Rect corners rounded at 10 units; pointer tip at (15, 95).
                sb.Append($"<path d=\"M 10,0 L 90,0 Q 100,0 100,10 L 100,70 Q 100,80 90,80 L 45,80 L 15,95 L 30,80 L 10,80 Q 0,80 0,70 L 0,10 Q 0,0 10,0 Z\" " +
                          $"fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
        }
        sb.Append("</svg>");
    }

}
