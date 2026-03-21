// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Shared helpers for building Drawing-namespace text/shape effects (a:effectLst children).
/// Used by both PPTX and Excel handlers to avoid code duplication.
/// Word uses a different namespace (w14) and has its own implementation.
/// </summary>
public static class DrawingEffectsHelper
{
    /// <summary>
    /// Build an OuterShadow element from a value string.
    /// Format: "COLOR[-BLUR[-ANGLE[-DIST[-OPACITY]]]]"
    /// Defaults: blur=4pt, angle=45°, dist=3pt, opacity=40%
    /// </summary>
    public static Drawing.OuterShadow BuildOuterShadow(string value, Func<string, OpenXmlElement> colorBuilder)
    {
        value = value.Replace(';', '-');
        var parts = value.Split('-');
        var blurPt = ParseParam(parts, 1, 4.0, "shadow blur");
        var angleDeg = ParseParam(parts, 2, 45.0, "shadow angle");
        var distPt = ParseParam(parts, 3, 3.0, "shadow distance");
        var opacity = ParseParam(parts, 4, 40.0, "shadow opacity");

        var shadow = new Drawing.OuterShadow
        {
            BlurRadius = (long)(blurPt * 12700),
            Distance = (long)(distPt * 12700),
            Direction = (int)(angleDeg * 60000),
            Alignment = Drawing.RectangleAlignmentValues.TopLeft,
            RotateWithShape = false
        };
        var clr = colorBuilder(parts[0]);
        clr.AppendChild(new Drawing.Alpha { Val = (int)(opacity * 1000) });
        shadow.AppendChild(clr);
        return shadow;
    }

    /// <summary>
    /// Build a Glow element from a value string.
    /// Format: "COLOR[-RADIUS[-OPACITY]]"
    /// Defaults: radius=8pt, opacity=75%
    /// </summary>
    public static Drawing.Glow BuildGlow(string value, Func<string, OpenXmlElement> colorBuilder)
    {
        value = value.Replace(';', '-');
        var parts = value.Split('-');
        var radiusPt = ParseParam(parts, 1, 8.0, "glow radius");
        var opacity = ParseParam(parts, 2, 75.0, "glow opacity");

        var glow = new Drawing.Glow { Radius = (long)(radiusPt * 12700) };
        var clr = colorBuilder(parts[0]);
        clr.AppendChild(new Drawing.Alpha { Val = (int)(opacity * 1000) });
        glow.AppendChild(clr);
        return glow;
    }

    /// <summary>
    /// Build a Reflection element from a value string.
    /// Values: "tight"/"small", "half"/"true", "full", or numeric percentage.
    /// </summary>
    public static Drawing.Reflection BuildReflection(string value)
    {
        int endPos = value.ToLowerInvariant() switch
        {
            "tight" or "small" => 55000,
            "true" or "half" => 90000,
            "full" => 100000,
            _ => int.TryParse(value, out var pct) ? (int)Math.Min((long)pct * 1000, 100000) : 90000
        };

        return new Drawing.Reflection
        {
            BlurRadius = 6350,
            StartOpacity = 52000,
            StartPosition = 0,
            EndAlpha = 300,
            EndPosition = endPos,
            Distance = 0,
            Direction = 5400000,
            VerticalRatio = -100000,
            Alignment = Drawing.RectangleAlignmentValues.BottomLeft,
            RotateWithShape = false
        };
    }

    /// <summary>
    /// Build a SoftEdge element from a value string (radius in points).
    /// </summary>
    public static Drawing.SoftEdge BuildSoftEdge(string value)
    {
        if (!double.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var radiusPt)
            || double.IsNaN(radiusPt) || double.IsInfinity(radiusPt))
            throw new ArgumentException($"Invalid 'softedge' value '{value}'. Expected a numeric radius in points.");
        return new Drawing.SoftEdge { Radius = (long)(radiusPt * 12700) };
    }

    /// <summary>
    /// Get or create EffectList in correct schema position within Drawing.RunProperties.
    /// CT_TextCharacterProperties order: ln → fill → effectLst → highlight → ... → latin → ea → ...
    /// </summary>
    public static Drawing.EffectList EnsureRunEffectList(Drawing.RunProperties rPr)
    {
        var existing = rPr.GetFirstChild<Drawing.EffectList>();
        if (existing != null) return existing;

        var effectList = new Drawing.EffectList();
        var insertBefore = (OpenXmlElement?)rPr.GetFirstChild<Drawing.Highlight>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.UnderlineFollowsText>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.Underline>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.UnderlineFillText>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.UnderlineFill>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.LatinFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.EastAsianFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.ComplexScriptFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.SymbolFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.HyperlinkOnClick>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.HyperlinkOnMouseOver>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.ExtensionList>();
        if (insertBefore != null)
            rPr.InsertBefore(effectList, insertBefore);
        else
            rPr.AppendChild(effectList);
        return effectList;
    }

    /// <summary>
    /// Insert a fill element at the correct schema position in Drawing.RunProperties.
    /// CT_TextCharacterProperties order: ln → fill → effectLst → ... → latin → ea → ...
    /// </summary>
    public static void InsertFillInRunProperties(Drawing.RunProperties rPr, OpenXmlElement fillElement)
    {
        var insertBefore = (OpenXmlElement?)rPr.GetFirstChild<Drawing.EffectList>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.EffectDag>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.Highlight>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.LatinFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.EastAsianFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.ComplexScriptFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.SymbolFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.HyperlinkOnClick>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.HyperlinkOnMouseOver>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.ExtensionList>();
        if (insertBefore != null)
            rPr.InsertBefore(fillElement, insertBefore);
        else
            rPr.AppendChild(fillElement);
    }

    /// <summary>
    /// Apply a text effect to a Drawing.Run's RunProperties effectLst.
    /// Handles create/remove logic. Returns false if value is "none".
    /// </summary>
    public static void ApplyTextEffect<T>(Drawing.Run run, string value, Func<T> builder) where T : OpenXmlElement
    {
        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
        var effectList = EnsureRunEffectList(rPr);
        effectList.RemoveAllChildren<T>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) rPr.RemoveChild(effectList);
            return;
        }
        effectList.AppendChild(builder());
    }

    /// <summary>
    /// Standard color builder for Drawing effects: sanitizes hex, creates RgbColorModelHex with optional alpha.
    /// Use instead of duplicating the lambda pattern inline.
    /// </summary>
    public static OpenXmlElement BuildRgbColor(string colorValue)
    {
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(colorValue);
        var clr = new Drawing.RgbColorModelHex { Val = rgb };
        if (alpha.HasValue) clr.AppendChild(new Drawing.Alpha { Val = alpha.Value });
        return clr;
    }

    // --- Private helpers ---

    private static double ParseParam(string[] parts, int index, double defaultValue, string paramName)
    {
        if (parts.Length <= index) return defaultValue;
        if (!double.TryParse(parts[index], System.Globalization.CultureInfo.InvariantCulture, out var val)
            || double.IsNaN(val) || double.IsInfinity(val))
            throw new ArgumentException($"Invalid {paramName} value: '{parts[index]}'.");
        return val;
    }
}
