// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;

namespace OfficeCli.Core;

/// <summary>
/// Swap the current thread's culture to <see cref="CultureInfo.InvariantCulture"/>
/// for the lifetime of the scope, restoring the original culture on Dispose.
///
/// Use to wrap CSS / HTML / SVG generation paths: under locales like de-DE the
/// default formatting of <c>double</c> produces <c>141,73</c> with a comma
/// decimal separator, which is invalid CSS and breaks the preview entirely.
/// Wrapping each public renderer entry point is preferable to auditing every
/// interpolated number deep in the rendering tree.
/// </summary>
internal readonly struct InvariantCultureScope : IDisposable
{
    private readonly CultureInfo _previous;
    private InvariantCultureScope(CultureInfo previous) { _previous = previous; }

    public static InvariantCultureScope Enter()
    {
        var prev = System.Threading.Thread.CurrentThread.CurrentCulture;
        System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        return new InvariantCultureScope(prev);
    }

    public void Dispose()
    {
        System.Threading.Thread.CurrentThread.CurrentCulture = _previous;
    }
}

/// <summary>
/// Shared unit conversion utilities for HTML preview rendering.
/// All methods convert to points (pt) — the natural unit of the OOXML coordinate system.
///
/// Key relationships (all exact integer ratios):
///   1 pt = 20 twips        (Word)
///   1 pt = 12700 EMU       (PowerPoint / Excel drawings)
///   1 pt = 2 half-points   (font sizes)
///
/// Using pt avoids the precision loss inherent in converting to cm or px:
///   EMU → cm: 360000 EMU/cm produces irrational values for most inputs
///   twips → px: 1440 twips/inch × 96 DPI involves floating-point rounding
/// </summary>
internal static class Units
{
    /// <summary>Convert Word twips to points. 1 pt = 20 twips (exact).</summary>
    public static double TwipsToPt(int twips) => twips / 20.0;

    /// <summary>Convert Word twips (string) to points. Returns 0 for unparseable input.</summary>
    public static double TwipsToPt(string twipsStr)
    {
        if (!int.TryParse(twipsStr, out var twips)) return 0;
        return twips / 20.0;
    }

    /// <summary>Format Word twips (string) to CSS pt value, e.g. "36pt".</summary>
    public static string TwipsToPtStr(string twipsStr)
    {
        return $"{TwipsToPt(twipsStr):0.##}pt";
    }

    /// <summary>Convert EMU to points. 1 pt = 12700 EMU (exact).</summary>
    public static double EmuToPt(long emu) => Math.Round(emu / 12700.0, 2);

    /// <summary>Convert half-points to points. 1 pt = 2 half-points (exact).</summary>
    public static double HalfPointsToPt(int hp) => hp / 2.0;
}
