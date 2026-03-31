// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Shared theme color resolution. Builds a scheme-color-name → hex dictionary
/// from an OOXML ColorScheme. Used by both PowerPoint and Word handlers.
/// </summary>
public static class ThemeColorResolver
{
    /// <summary>
    /// Build a map of scheme color names to hex values from a ColorScheme.
    /// </summary>
    /// <param name="colorScheme">The theme's ColorScheme element.</param>
    /// <param name="includePptAliases">
    /// If true, adds PPT-specific aliases: text1, text2, background1, background2.
    /// Word uses a smaller alias set.
    /// </param>
    public static Dictionary<string, string> BuildColorMap(
        Drawing.ColorScheme? colorScheme, bool includePptAliases = false)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (colorScheme == null) return map;

        void Add(string name, OpenXmlCompositeElement? color)
        {
            if (color == null) return;
            var rgb = color.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            var sys = color.GetFirstChild<Drawing.SystemColor>();
            var srgb = sys?.LastColor?.Value ?? sys?.Val?.InnerText;
            var hex = rgb ?? srgb;
            if (hex != null) map[name] = hex;
        }

        Add("dk1", colorScheme.Dark1Color);
        Add("dk2", colorScheme.Dark2Color);
        Add("lt1", colorScheme.Light1Color);
        Add("lt2", colorScheme.Light2Color);
        Add("accent1", colorScheme.Accent1Color);
        Add("accent2", colorScheme.Accent2Color);
        Add("accent3", colorScheme.Accent3Color);
        Add("accent4", colorScheme.Accent4Color);
        Add("accent5", colorScheme.Accent5Color);
        Add("accent6", colorScheme.Accent6Color);
        Add("hlink", colorScheme.Hyperlink);
        Add("folHlink", colorScheme.FollowedHyperlinkColor);

        // Aliases shared by both PPT and Word
        if (map.TryGetValue("dk1", out var dk1)) { map["tx1"] = dk1; map["dark1"] = dk1; }
        if (map.TryGetValue("dk2", out var dk2)) { map["dark2"] = dk2; }
        if (map.TryGetValue("lt1", out var lt1)) { map["bg1"] = lt1; map["light1"] = lt1; }
        if (map.TryGetValue("lt2", out var lt2)) { map["bg2"] = lt2; map["light2"] = lt2; }

        // PPT-specific aliases
        if (includePptAliases)
        {
            if (dk1 != null) map["text1"] = dk1;
            if (dk2 != null) { map["text2"] = dk2; map["tx2"] = dk2; }
            if (lt1 != null) map["background1"] = lt1;
            if (lt2 != null) map["background2"] = lt2;
        }

        return map;
    }
}
