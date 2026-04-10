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
        // Fix uniform-distribution spacing (균등 분할): "현 장 대 응" → "현장대응"
        // Only collapse when 3+ consecutive single Korean syllables are space-separated.
        // Preserves normal word spacing like "인사 발령 통보".
        text = UniformDistRegex().Replace(text, m => m.Value.Replace(" ", ""));
        // Remove zero-width joiners between jamo
        text = text.Replace("\u200D", "");
        return text;
    }

    [GeneratedRegex(@"사각형입니다\.")]
    private static partial Regex ShapeAltTextRegex();

    [GeneratedRegex(@"(?<!\p{IsHangulSyllables})\p{IsHangulSyllables}(?: \p{IsHangulSyllables}){2,}(?!\p{IsHangulSyllables})")]
    private static partial Regex UniformDistRegex();
}
