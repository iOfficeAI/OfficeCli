// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;

namespace OfficeCli.Handlers;

internal static partial class HwpxKorean
{
    public static string Normalize(string text)
    {
        text = StripPuaChars(text);
        text = StripShapeAltText(text);
        text = NormalizeKoreanSpacing(text);
        return text;
    }

    public static string StripPuaChars(string text)
        => string.Concat(text.Where(c => c < '\uE000' || c > '\uF8FF'));

    public static string StripShapeAltText(string text)
        => ShapeAltTextRegex().Replace(text, "");

    public static string NormalizeKoreanSpacing(string text)
    {
        // Fix uniform-distribution spacing between Korean syllables
        text = KoreanSpacingRegex().Replace(text, "$1$2");
        // Remove zero-width joiners between jamo
        text = text.Replace("\u200D", "");
        return text;
    }

    [GeneratedRegex(@"사각형입니다\.")]
    private static partial Regex ShapeAltTextRegex();

    [GeneratedRegex(@"(\p{IsHangulSyllables}) +(\p{IsHangulSyllables})")]
    private static partial Regex KoreanSpacingRegex();
}
