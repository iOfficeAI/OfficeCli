// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private static void InsertFillElement(ShapeProperties spPr, OpenXmlElement fillElement)
    {
        // Schema order: xfrm → prstGeom → fill → ln → effectLst
        var prstGeom = spPr.GetFirstChild<Drawing.PresetGeometry>();
        if (prstGeom != null)
            spPr.InsertAfter(fillElement, prstGeom);
        else
        {
            var xfrm = spPr.Transform2D;
            if (xfrm != null)
                spPr.InsertAfter(fillElement, xfrm);
            else
                spPr.PrependChild(fillElement);
        }
    }

    private static void ApplyShapeFill(ShapeProperties spPr, string value)
    {
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        spPr.RemoveAllChildren<Drawing.PatternFill>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
            InsertFillElement(spPr, new Drawing.NoFill());
        else
        {
            var solidFill = new Drawing.SolidFill();
            solidFill.Append(new Drawing.RgbColorModelHex { Val = value.TrimStart('#').ToUpperInvariant() });
            InsertFillElement(spPr, solidFill);
        }
    }

    /// <summary>
    /// Apply gradient fill to ShapeProperties.
    /// Linear:  "color1-color2[-angle]"       e.g. "FF0000-0000FF", "FF0000-0000FF-90"
    /// Radial:  "radial:color1-color2"         e.g. "radial:4B0082-1E90FF"
    /// Radial with focus: "radial:color1-color2-tl" (tl/tr/bl/br/center)
    /// </summary>
    private static void ApplyGradientFill(ShapeProperties spPr, string value)
    {
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        InsertFillElement(spPr, BuildGradientFill(value));
    }

    /// <summary>
    /// Apply image (blip) fill to a shape.
    /// Format: file path to image, e.g. "/tmp/bg.png"
    /// </summary>
    private static void ApplyShapeImageFill(ShapeProperties spPr, string imagePath, SlidePart part)
    {
        if (!File.Exists(imagePath))
            throw new ArgumentException($"Image file not found: {imagePath}");

        var ext = Path.GetExtension(imagePath).ToLowerInvariant();
        var partType = ext switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tif" or ".tiff" => ImagePartType.Tiff,
            ".emf" => ImagePartType.Emf,
            ".wmf" => ImagePartType.Wmf,
            _ => throw new ArgumentException($"Unsupported image format: {ext}")
        };

        var imagePart = part.AddImagePart(partType);
        using (var stream = File.OpenRead(imagePath))
            imagePart.FeedData(stream);
        var relId = part.GetIdOfPart(imagePart);

        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        spPr.RemoveAllChildren<Drawing.BlipFill>();
        spPr.RemoveAllChildren<Drawing.PatternFill>();

        var blipFill = new Drawing.BlipFill();
        blipFill.Append(new Drawing.Blip { Embed = relId });
        blipFill.Append(new Drawing.Stretch(new Drawing.FillRectangle()));
        InsertFillElement(spPr, blipFill);
    }

    /// <summary>
    /// Apply text margin (padding) to a BodyProperties element.
    /// Supports: single value "0.5cm" (all sides), or "left,top,right,bottom" e.g. "0.5cm,0.3cm,0.5cm,0.3cm"
    /// </summary>
    private static void ApplyTextMargin(Drawing.BodyProperties bodyPr, string value)
    {
        var parts = value.Split(',');
        if (parts.Length == 1)
        {
            var emu = ParseEmu(parts[0]);
            bodyPr.LeftInset = (int)emu;
            bodyPr.TopInset = (int)emu;
            bodyPr.RightInset = (int)emu;
            bodyPr.BottomInset = (int)emu;
        }
        else if (parts.Length == 4)
        {
            bodyPr.LeftInset = (int)ParseEmu(parts[0].Trim());
            bodyPr.TopInset = (int)ParseEmu(parts[1].Trim());
            bodyPr.RightInset = (int)ParseEmu(parts[2].Trim());
            bodyPr.BottomInset = (int)ParseEmu(parts[3].Trim());
        }
        else
        {
            throw new ArgumentException("margin must be a single value or 4 comma-separated values (left,top,right,bottom)");
        }
    }

    private static Drawing.TextAlignmentTypeValues ParseTextAlignment(string value) =>
        value.ToLowerInvariant() switch
        {
            "left" or "l" => Drawing.TextAlignmentTypeValues.Left,
            "center" or "c" => Drawing.TextAlignmentTypeValues.Center,
            "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
            "justify" or "j" => Drawing.TextAlignmentTypeValues.Justified,
            _ => throw new ArgumentException($"Invalid align: {value}. Use: left, center, right, justify")
        };

    /// <summary>
    /// Apply list style (bullet/numbered) to ParagraphProperties.
    /// Values: "bullet" or "•", "numbered" or "1", "alpha" or "a", "roman" or "i", "none"
    /// </summary>
    private static void ApplyListStyle(Drawing.ParagraphProperties pProps, string value)
    {
        pProps.RemoveAllChildren<Drawing.CharacterBullet>();
        pProps.RemoveAllChildren<Drawing.AutoNumberedBullet>();
        pProps.RemoveAllChildren<Drawing.NoBullet>();
        pProps.RemoveAllChildren<Drawing.BulletFont>();

        switch (value.ToLowerInvariant())
        {
            case "bullet" or "•" or "disc":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "•" });
                break;
            case "dash" or "-" or "–":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "–" });
                break;
            case "arrow" or ">" or "→":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "→" });
                break;
            case "check" or "✓":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "✓" });
                break;
            case "star" or "★":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "★" });
                break;
            case "numbered" or "number" or "1":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.ArabicPeriod });
                break;
            case "alpha" or "a":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod });
                break;
            case "alphaupper" or "A":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.AlphaUpperCharacterPeriod });
                break;
            case "roman" or "i":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.RomanLowerCharacterPeriod });
                break;
            case "romanupper" or "I":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.RomanUpperCharacterPeriod });
                break;
            case "none" or "false":
                pProps.AppendChild(new Drawing.NoBullet());
                break;
            default:
                if (value.Length <= 2)
                    pProps.AppendChild(new Drawing.CharacterBullet { Char = value });
                else
                    throw new ArgumentException($"Invalid list style: {value}. Use: bullet, numbered, alpha, roman, none, or a single character");
                break;
        }
    }

    private static Drawing.ShapeTypeValues ParsePresetShape(string name) =>
        name.ToLowerInvariant() switch
        {
            "rect" or "rectangle" => Drawing.ShapeTypeValues.Rectangle,
            "roundrect" or "roundedrectangle" => Drawing.ShapeTypeValues.RoundRectangle,
            "ellipse" or "oval" => Drawing.ShapeTypeValues.Ellipse,
            "triangle" => Drawing.ShapeTypeValues.Triangle,
            "rtriangle" or "righttriangle" => Drawing.ShapeTypeValues.RightTriangle,
            "diamond" => Drawing.ShapeTypeValues.Diamond,
            "parallelogram" => Drawing.ShapeTypeValues.Parallelogram,
            "trapezoid" => Drawing.ShapeTypeValues.Trapezoid,
            "pentagon" => Drawing.ShapeTypeValues.Pentagon,
            "hexagon" => Drawing.ShapeTypeValues.Hexagon,
            "heptagon" => Drawing.ShapeTypeValues.Heptagon,
            "octagon" => Drawing.ShapeTypeValues.Octagon,
            "star4" => Drawing.ShapeTypeValues.Star4,
            "star5" => Drawing.ShapeTypeValues.Star5,
            "star6" => Drawing.ShapeTypeValues.Star6,
            "star8" => Drawing.ShapeTypeValues.Star8,
            "star10" => Drawing.ShapeTypeValues.Star10,
            "star12" => Drawing.ShapeTypeValues.Star12,
            "star16" => Drawing.ShapeTypeValues.Star16,
            "star24" => Drawing.ShapeTypeValues.Star24,
            "star32" => Drawing.ShapeTypeValues.Star32,
            "rightarrow" or "rarrow" => Drawing.ShapeTypeValues.RightArrow,
            "leftarrow" or "larrow" => Drawing.ShapeTypeValues.LeftArrow,
            "uparrow" => Drawing.ShapeTypeValues.UpArrow,
            "downarrow" => Drawing.ShapeTypeValues.DownArrow,
            "leftrightarrow" or "lrarrow" => Drawing.ShapeTypeValues.LeftRightArrow,
            "updownarrow" or "udarrow" => Drawing.ShapeTypeValues.UpDownArrow,
            "chevron" => Drawing.ShapeTypeValues.Chevron,
            "homeplat" or "homeplate" => Drawing.ShapeTypeValues.HomePlate,
            "plus" or "cross" => Drawing.ShapeTypeValues.Plus,
            "heart" => Drawing.ShapeTypeValues.Heart,
            "cloud" => Drawing.ShapeTypeValues.Cloud,
            "lightning" or "lightningbolt" => Drawing.ShapeTypeValues.LightningBolt,
            "sun" => Drawing.ShapeTypeValues.Sun,
            "moon" => Drawing.ShapeTypeValues.Moon,
            "arc" => Drawing.ShapeTypeValues.Arc,
            "donut" => Drawing.ShapeTypeValues.Donut,
            "nosmoking" or "blockarc" => Drawing.ShapeTypeValues.NoSmoking,
            "cube" => Drawing.ShapeTypeValues.Cube,
            "can" or "cylinder" => Drawing.ShapeTypeValues.Can,
            "line" => Drawing.ShapeTypeValues.Line,
            "decagon" => Drawing.ShapeTypeValues.Decagon,
            "dodecagon" => Drawing.ShapeTypeValues.Dodecagon,
            "ribbon" => Drawing.ShapeTypeValues.Ribbon,
            "ribbon2" => Drawing.ShapeTypeValues.Ribbon2,
            "callout1" => Drawing.ShapeTypeValues.Callout1,
            "callout2" => Drawing.ShapeTypeValues.Callout2,
            "callout3" => Drawing.ShapeTypeValues.Callout3,
            "wedgeroundrectcallout" or "callout" => Drawing.ShapeTypeValues.WedgeRoundRectangleCallout,
            "wedgeellipsecallout" => Drawing.ShapeTypeValues.WedgeEllipseCallout,
            "cloudcallout" => Drawing.ShapeTypeValues.CloudCallout,
            "flowchartprocess" or "process" => Drawing.ShapeTypeValues.FlowChartProcess,
            "flowchartdecision" or "decision" => Drawing.ShapeTypeValues.FlowChartDecision,
            "flowchartterminator" or "terminator" => Drawing.ShapeTypeValues.FlowChartTerminator,
            "flowchartdocument" => Drawing.ShapeTypeValues.FlowChartDocument,
            "flowcharttinputoutput" or "io" => Drawing.ShapeTypeValues.FlowChartInputOutput,
            "brace" or "leftbrace" => Drawing.ShapeTypeValues.LeftBrace,
            "rightbrace" => Drawing.ShapeTypeValues.RightBrace,
            "leftbracket" => Drawing.ShapeTypeValues.LeftBracket,
            "rightbracket" => Drawing.ShapeTypeValues.RightBracket,
            "smileyface" or "smiley" => Drawing.ShapeTypeValues.SmileyFace,
            "foldedcorner" => Drawing.ShapeTypeValues.FoldedCorner,
            "frame" => Drawing.ShapeTypeValues.Frame,
            "gear6" => Drawing.ShapeTypeValues.Gear6,
            "gear9" => Drawing.ShapeTypeValues.Gear9,
            "notchedrightarrow" => Drawing.ShapeTypeValues.NotchedRightArrow,
            "bentuparrow" => Drawing.ShapeTypeValues.BentUpArrow,
            "curvedrightarrow" => Drawing.ShapeTypeValues.CurvedRightArrow,
            "stripedrightarrow" => Drawing.ShapeTypeValues.StripedRightArrow,
            "uturnArrow" => Drawing.ShapeTypeValues.UTurnArrow,
            "circularArrow" => Drawing.ShapeTypeValues.CircularArrow,
            _ => throw new ArgumentException(
                $"Unknown preset shape: '{name}'. Common presets: rect, roundRect, ellipse, triangle, diamond, " +
                "pentagon, hexagon, star5, rightArrow, leftArrow, chevron, plus, heart, cloud, cube, can, line, " +
                "callout, process, decision, smiley, frame, gear6")
        };
}
