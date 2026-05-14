// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Cross-handler builders for DrawingML (`a:` namespace) color/fill elements.
/// Used by both PowerPointHandler (slide shapes, runs) and ExcelHandler (drawing-layer
/// shapes, chart series). Word's <c>w:</c> namespace has its own run-property color
/// model and does not share this helper.
/// </summary>
internal static class DrawingColorBuilder
{
    /// <summary>
    /// Parse a color string and return the appropriate DrawingML color element:
    /// <c>a:srgbClr</c> (with optional <c>a:alpha</c>) for hex/named colors,
    /// or <c>a:schemeClr</c> for theme color names (accent1..6, dk1/lt1/tx1/bg1/hlink/...).
    /// </summary>
    internal static OpenXmlElement BuildColorElement(string value)
    {
        var schemeColor = TryParseSchemeColor(value);
        if (schemeColor.HasValue)
            return new Drawing.SchemeColor { Val = schemeColor.Value };

        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(value);
        var colorEl = new Drawing.RgbColorModelHex { Val = rgb };
        if (alpha.HasValue)
            colorEl.AppendChild(new Drawing.Alpha { Val = alpha.Value });
        return colorEl;
    }

    /// <summary>
    /// Build an <c>a:solidFill</c> element with the appropriate color child (RGB or scheme).
    /// </summary>
    internal static Drawing.SolidFill BuildSolidFill(string colorValue)
    {
        var solidFill = new Drawing.SolidFill();
        solidFill.Append(BuildColorElement(colorValue));
        return solidFill;
    }

    /// <summary>
    /// Try to parse a theme/scheme color name. Returns null if the input is a hex RGB value.
    /// </summary>
    internal static Drawing.SchemeColorValues? TryParseSchemeColor(string value)
    {
        return value.ToLowerInvariant().TrimStart('#') switch
        {
            "accent1" => Drawing.SchemeColorValues.Accent1,
            "accent2" => Drawing.SchemeColorValues.Accent2,
            "accent3" => Drawing.SchemeColorValues.Accent3,
            "accent4" => Drawing.SchemeColorValues.Accent4,
            "accent5" => Drawing.SchemeColorValues.Accent5,
            "accent6" => Drawing.SchemeColorValues.Accent6,
            "dk1" or "dark1" => Drawing.SchemeColorValues.Dark1,
            "dk2" or "dark2" => Drawing.SchemeColorValues.Dark2,
            "lt1" or "light1" => Drawing.SchemeColorValues.Light1,
            "lt2" or "light2" => Drawing.SchemeColorValues.Light2,
            "tx1" or "text1" => Drawing.SchemeColorValues.Text1,
            "tx2" or "text2" => Drawing.SchemeColorValues.Text2,
            "bg1" or "background1" => Drawing.SchemeColorValues.Background1,
            "bg2" or "background2" => Drawing.SchemeColorValues.Background2,
            "hlink" or "hyperlink" => Drawing.SchemeColorValues.Hyperlink,
            "folhlink" or "followedhyperlink" => Drawing.SchemeColorValues.FollowedHyperlink,
            _ => null
        };
    }
}
