// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;

namespace OfficeCli.Core;

/// <summary>
/// Shared parsing helpers for handler property values.
/// Accepts flexible user input (e.g. "true", "yes", "1", "on" for booleans;
/// "24pt" or "24" for font sizes).
/// </summary>
public static class ParseHelpers
{
    /// <summary>
    /// Accepts "true", "1", "yes", "on" (case-insensitive) as truthy.
    /// </summary>
    public static bool IsTruthy(string value) =>
        value.ToLowerInvariant() is "true" or "1" or "yes" or "on";

    /// <summary>
    /// Parse a font size string, stripping optional "pt" suffix.
    /// Supports integers and fractional values (e.g. "24", "10.5", "24pt").
    /// Returns double to preserve fractional sizes for correct unit conversion.
    /// </summary>
    public static double ParseFontSize(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed[..^2].Trim();
        return double.Parse(trimmed, CultureInfo.InvariantCulture);
    }
}
