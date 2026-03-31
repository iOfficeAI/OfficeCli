// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Shared RGB↔HSL color space conversion and OOXML color transform helpers.
/// Extracted from PowerPointHandler.HtmlPreview.Css and WordHandler.HtmlPreview.Css
/// to eliminate duplication.
/// </summary>
public static class ColorMath
{
    /// <summary>Convert RGB (0-255) to HSL (h: 0-1, s: 0-1, l: 0-1).</summary>
    public static void RgbToHsl(int r, int g, int b, out double h, out double s, out double l)
    {
        var rf = r / 255.0;
        var gf = g / 255.0;
        var bf = b / 255.0;
        var max = Math.Max(rf, Math.Max(gf, bf));
        var min = Math.Min(rf, Math.Min(gf, bf));
        var delta = max - min;

        l = (max + min) / 2.0;

        if (delta < 1e-10)
        {
            h = 0;
            s = 0;
            return;
        }

        s = l < 0.5 ? delta / (max + min) : delta / (2.0 - max - min);

        if (Math.Abs(max - rf) < 1e-10)
            h = ((gf - bf) / delta + (gf < bf ? 6 : 0)) / 6.0;
        else if (Math.Abs(max - gf) < 1e-10)
            h = ((bf - rf) / delta + 2) / 6.0;
        else
            h = ((rf - gf) / delta + 4) / 6.0;
    }

    /// <summary>Convert HSL (h: 0-1, s: 0-1, l: 0-1) to RGB (0-255).</summary>
    public static void HslToRgb(double h, double s, double l, out int r, out int g, out int b)
    {
        if (s < 1e-10)
        {
            r = g = b = (int)Math.Round(l * 255);
            return;
        }

        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;

        r = (int)Math.Round(HueToRgb(p, q, h + 1.0 / 3) * 255);
        g = (int)Math.Round(HueToRgb(p, q, h) * 255);
        b = (int)Math.Round(HueToRgb(p, q, h - 1.0 / 3) * 255);
    }

    /// <summary>Helper for HSL→RGB conversion.</summary>
    internal static double HueToRgb(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 1.0 / 2) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    /// <summary>
    /// Apply OOXML lumMod/lumOff color transform in HSL space.
    /// lumMod and lumOff are in 0–100000 units (percentage × 1000).
    /// Formula: newL = clamp(L × lumMod/100000 + lumOff/100000, 0, 1)
    /// </summary>
    public static string ApplyLumModOff(string hex, int lumMod, int lumOff)
    {
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

        RgbToHsl(r, g, b, out var h, out var s, out var l);
        l = Math.Clamp(l * (lumMod / 100000.0) + (lumOff / 100000.0), 0, 1);
        HslToRgb(h, s, l, out r, out g, out b);

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    /// <summary>
    /// Apply OOXML DrawingML color transforms: tint, shade, lumMod, lumOff, alpha.
    /// All values in 0–100000 units (percentage × 1000). Pass null to skip a transform.
    /// Input hex is 6-char without '#' prefix. Output includes '#' prefix (or rgba() if alpha &lt; 100000).
    /// </summary>
    public static string ApplyTransforms(string hex, int? tint = null, int? shade = null,
        int? lumMod = null, int? lumOff = null, int? alpha = null)
    {
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

        // OOXML spec: tint blends toward white, shade blends toward black
        if (tint.HasValue)
        {
            var t = tint.Value / 100000.0;
            r = (int)(r + (255 - r) * (1 - t));
            g = (int)(g + (255 - g) * (1 - t));
            b = (int)(b + (255 - b) * (1 - t));
        }

        if (shade.HasValue)
        {
            var s = shade.Value / 100000.0;
            r = (int)(r * s);
            g = (int)(g * s);
            b = (int)(b * s);
        }

        // OOXML spec: lumMod/lumOff operate in HSL space
        if (lumMod.HasValue || lumOff.HasValue)
        {
            var mod = (lumMod ?? 100000) / 100000.0;
            var off = (lumOff ?? 0) / 100000.0;
            RgbToHsl(r, g, b, out var h, out var s, out var l);
            l = Math.Clamp(l * mod + off, 0, 1);
            HslToRgb(h, s, l, out r, out g, out b);
        }

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);

        if (alpha.HasValue && alpha.Value < 100000)
            return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";

        return $"#{r:X2}{g:X2}{b:X2}";
    }
}
