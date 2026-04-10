// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Runtime.Versioning;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Image Helpers ====================

    private static long ParseEmu(string value) => Core.EmuConverter.ParseEmu(value);

    private uint NextDocPropId()
    {
        uint maxId = 0;
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            foreach (var dp in body.Descendants<DW.DocProperties>())
            {
                if (dp.Id?.HasValue == true && dp.Id.Value > maxId)
                    maxId = dp.Id.Value;
            }
        }
        return maxId + 1;
    }

    private static Run CreateImageRun(string relationshipId, long cx, long cy, string altText, uint docPropId)
    {
        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = docPropId, Name = altText, Description = altText },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }
            ),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = docPropId, Name = altText },
                            new PIC.NonVisualPictureDrawingProperties()
                        ),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                            new A.Stretch(new A.FillRectangle())
                        ),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = cx, Cy = cy }
                            ),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                        )
                    )
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
            )
        )
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        return new Run(new Drawing(inline));
    }

    private static Run CreateAnchorImageRun(string relationshipId, long cx, long cy, string altText,
        string wrap, long hPos, long vPos,
        DW.HorizontalRelativePositionValues hRel, DW.VerticalRelativePositionValues vRel,
        bool behindText, uint docPropId)
    {
        OpenXmlElement wrapElement = wrap.ToLowerInvariant() switch
        {
            "square" => new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
            "tight" => new DW.WrapTight(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "through" => new DW.WrapThrough(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "topandbottom" or "topbottom" => new DW.WrapTopBottom(),
            "none" => new DW.WrapNone() as OpenXmlElement,
            _ => throw new ArgumentException($"Invalid wrap value: '{wrap}'. Valid values: none, square, tight, through, topandbottom.")
        };

        var anchorDocPropId = docPropId;
        var anchor = new DW.Anchor(
            new DW.SimplePosition { X = 0, Y = 0 },
            new DW.HorizontalPosition(new DW.PositionOffset(hPos.ToString()))
                { RelativeFrom = hRel },
            new DW.VerticalPosition(new DW.PositionOffset(vPos.ToString()))
                { RelativeFrom = vRel },
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            wrapElement,
            new DW.DocProperties { Id = anchorDocPropId, Name = altText, Description = altText },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = anchorDocPropId, Name = altText },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle })
                    )
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
            )
        )
        {
            BehindDoc = behindText,
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 114300U,
            DistanceFromRight = 114300U,
            SimplePos = false,
            RelativeHeight = 1U,
            AllowOverlap = true,
            LayoutInCell = true,
            Locked = false
        };

        return new Run(new Drawing(anchor));
    }

    private static DW.HorizontalRelativePositionValues ParseHorizontalRelative(string value) =>
        value.ToLowerInvariant() switch
        {
            "page" => DW.HorizontalRelativePositionValues.Page,
            "column" => DW.HorizontalRelativePositionValues.Column,
            "character" => DW.HorizontalRelativePositionValues.Character,
            "margin" => DW.HorizontalRelativePositionValues.Margin,
            _ => throw new ArgumentException($"Invalid horizontal relative position: '{value}'. Valid values: margin, page, column, character.")
        };

    private static DW.VerticalRelativePositionValues ParseVerticalRelative(string value) =>
        value.ToLowerInvariant() switch
        {
            "page" => DW.VerticalRelativePositionValues.Page,
            "paragraph" => DW.VerticalRelativePositionValues.Paragraph,
            "line" => DW.VerticalRelativePositionValues.Line,
            "margin" => DW.VerticalRelativePositionValues.Margin,
            _ => throw new ArgumentException($"Invalid vertical relative position: '{value}'. Valid values: margin, page, paragraph, line.")
        };

    private static string GetDrawingInfo(Drawing drawing)
    {
        var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();

        var parts = new List<string>();
        if (docProps?.Description?.Value is string desc && !string.IsNullOrEmpty(desc))
            parts.Add($"alt=\"{desc}\"");
        else if (docProps?.Name?.Value is string name && !string.IsNullOrEmpty(name))
            parts.Add($"name=\"{name}\"");
        if (extent != null)
        {
            var wCm = extent.Cx != null ? $"{extent.Cx.Value / 360000.0:F1}cm" : "?";
            var hCm = extent.Cy != null ? $"{extent.Cy.Value / 360000.0:F1}cm" : "?";
            parts.Add($"{wCm}×{hCm}");
        }
        return parts.Count > 0 ? string.Join(", ", parts) : "unknown";
    }

    private static DocumentNode CreateImageNode(Drawing drawing, Run run, string path)
    {
        var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();

        var node = new DocumentNode
        {
            Path = path,
            Type = "picture",
            Text = docProps?.Description?.Value ?? docProps?.Name?.Value ?? ""
        };
        if (docProps?.Id?.HasValue == true) node.Format["id"] = docProps.Id.Value;
        if (docProps?.Name?.Value != null) node.Format["name"] = docProps.Name.Value;
        if (extent?.Cx != null) node.Format["width"] = $"{extent.Cx.Value / 360000.0:F1}cm";
        if (extent?.Cy != null) node.Format["height"] = $"{extent.Cy.Value / 360000.0:F1}cm";
        if (docProps?.Description?.Value != null) node.Format["alt"] = docProps.Description.Value;

        // Detect wrap type and position from inline/anchor
        var inlineEl = drawing.GetFirstChild<DW.Inline>();
        var anchorEl = drawing.GetFirstChild<DW.Anchor>();
        if (inlineEl != null)
        {
            node.Format["wrap"] = "inline";
        }
        else if (anchorEl != null)
        {
            node.Format["wrap"] = DetectWrapType(anchorEl);
            if (anchorEl.BehindDoc?.Value == true)
                node.Format["behindText"] = true;

            var hPos = anchorEl.GetFirstChild<DW.HorizontalPosition>();
            if (hPos != null)
            {
                var offset = hPos.GetFirstChild<DW.PositionOffset>();
                if (offset != null && long.TryParse(offset.Text, out var hEmu))
                    node.Format["hPosition"] = $"{hEmu / 360000.0:F1}cm";
                if (hPos.RelativeFrom?.HasValue == true)
                    node.Format["hRelative"] = hPos.RelativeFrom.InnerText;
            }

            var vPos = anchorEl.GetFirstChild<DW.VerticalPosition>();
            if (vPos != null)
            {
                var offset = vPos.GetFirstChild<DW.PositionOffset>();
                if (offset != null && long.TryParse(offset.Text, out var vEmu))
                    node.Format["vPosition"] = $"{vEmu / 360000.0:F1}cm";
                if (vPos.RelativeFrom?.HasValue == true)
                    node.Format["vRelative"] = vPos.RelativeFrom.InnerText;
            }
        }

        return node;
    }

    private static string DetectWrapType(DW.Anchor anchor)
    {
        if (anchor.GetFirstChild<DW.WrapNone>() != null) return "none";
        if (anchor.GetFirstChild<DW.WrapSquare>() != null) return "square";
        if (anchor.GetFirstChild<DW.WrapTight>() != null) return "tight";
        if (anchor.GetFirstChild<DW.WrapThrough>() != null) return "through";
        if (anchor.GetFirstChild<DW.WrapTopBottom>() != null) return "topandbottom";
        return "none";
    }

    private static void ReplaceWrapElement(DW.Anchor anchor, string wrapType)
    {
        // Remove existing wrap element
        anchor.GetFirstChild<DW.WrapNone>()?.Remove();
        anchor.GetFirstChild<DW.WrapSquare>()?.Remove();
        anchor.GetFirstChild<DW.WrapTight>()?.Remove();
        anchor.GetFirstChild<DW.WrapThrough>()?.Remove();
        anchor.GetFirstChild<DW.WrapTopBottom>()?.Remove();

        OpenXmlElement newWrap = wrapType.ToLowerInvariant() switch
        {
            "square" => new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
            "tight" => new DW.WrapTight(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "through" => new DW.WrapThrough(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "topandbottom" or "topbottom" => new DW.WrapTopBottom(),
            "none" => new DW.WrapNone(),
            _ => throw new ArgumentException($"Invalid wrap value: '{wrapType}'. Valid values: none, square, tight, through, topandbottom.")
        };

        // Insert wrap after EffectExtent (standard OOXML order)
        var effectExtent = anchor.GetFirstChild<DW.EffectExtent>();
        if (effectExtent != null)
            effectExtent.InsertAfterSelf(newWrap);
        else
            anchor.PrependChild(newWrap);
    }

    private DocumentNode CreateOleNode(EmbeddedObject oleObj, Run run, string path)
    {
        var node = new DocumentNode
        {
            Path = path,
            Type = "ole",
            Text = ""
        };
        node.Format["objectType"] = "ole";

        // Extract ProgID from o:OLEObject
        var oleElement = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
        if (oleElement != null)
        {
            var progId = oleElement.GetAttributes().FirstOrDefault(a => a.LocalName == "ProgID").Value;
            if (progId != null)
            {
                node.Format["progId"] = progId;
                node.Text = progId;
            }
        }

        // Extract dimensions from v:shape style
        var shape = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "shape");
        if (shape != null)
        {
            var style = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value;
            if (style != null)
                ParseVmlStyle(style, node);
        }

        // Extract preview image from v:imagedata (Windows only — requires GDI+)
        var (previewPath, previewContentType) = OperatingSystem.IsWindowsVersionAtLeast(6, 1)
            ? ExtractOlePreviewImage(oleObj, path)
            : (null, null);
        if (previewPath != null)
        {
            node.Format["previewImage"] = previewPath;
            if (previewContentType != null)
                node.Format["previewContentType"] = previewContentType;
        }

        return node;
    }

    /// <summary>
    /// Extract the OLE preview image (EMF/WMF) from v:imagedata, convert to PNG,
    /// and save to temp directory. Returns (pngPath, originalContentType) or (null, null).
    /// </summary>
    [SupportedOSPlatform("windows6.1")]
    private (string? path, string? contentType) ExtractOlePreviewImage(EmbeddedObject oleObj, string nodePath)
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return (null, null);

        // Find v:imagedata element and its r:id
        var shape = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "shape");
        if (shape == null) return (null, null);

        var imageData = shape.Descendants().FirstOrDefault(e => e.LocalName == "imagedata");
        if (imageData == null) return (null, null);

        var rId = imageData.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value;
        if (string.IsNullOrEmpty(rId)) return (null, null);

        try
        {
            var imgPart = mainPart.GetPartById(rId);
            using var stream = imgPart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            ms.Position = 0;

            var contentType = imgPart.ContentType ?? "";
            var isMetafile = contentType.Contains("emf") || contentType.Contains("wmf")
                          || contentType.Contains("metafile");

            // Build a stable file name from the node path
            var safeId = nodePath.Replace("/", "_").Replace("[", "").Replace("]", "").TrimStart('_');
            var pngPath = Path.Combine(Path.GetTempPath(), $"officecli_ole_{safeId}.png");

            if (isMetafile)
            {
                // Convert EMF/WMF to PNG using System.Drawing (Windows GDI+)
                using var img = System.Drawing.Image.FromStream(ms);
                img.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
            }
            else if (contentType.Contains("png"))
            {
                using var fs = new FileStream(pngPath, FileMode.Create);
                ms.CopyTo(fs);
            }
            else
            {
                // JPEG or other raster — convert to PNG for consistency
                using var img = System.Drawing.Image.FromStream(ms);
                img.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
            }

            return (pngPath, contentType);
        }
        catch
        {
            return (null, null);
        }
    }

    private static void ParseVmlStyle(string style, DocumentNode node)
    {
        foreach (var part in style.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var kv = part.Split(':', 2);
            if (kv.Length != 2) continue;
            var k = kv[0].Trim().ToLowerInvariant();
            var v = kv[1].Trim();
            if (k == "width") node.Format["width"] = ConvertPtToCm(v);
            else if (k == "height") node.Format["height"] = ConvertPtToCm(v);
        }
    }

    private static string ConvertPtToCm(string ptValue)
    {
        // Handle values like "385.45pt"
        var num = ptValue.Replace("pt", "").Replace("in", "").Trim();
        if (double.TryParse(num, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var val))
        {
            if (ptValue.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                return $"{val * 2.54 / 72.0:F1}cm";
            if (ptValue.EndsWith("in", StringComparison.OrdinalIgnoreCase))
                return $"{val * 2.54:F1}cm";
        }
        return ptValue; // return as-is if unparseable
    }
}
