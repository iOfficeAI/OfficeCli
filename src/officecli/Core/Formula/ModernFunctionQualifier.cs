// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Prefixes Excel 2016+ dynamic-array and "modern" function names with
/// <c>_xlfn.</c> when emitting OOXML. Excel refuses to resolve bare
/// post-2016 function names (e.g. <c>SEQUENCE(5)</c> → <c>#NAME?</c>)
/// unless the XML formula uses the namespaced form (<c>_xlfn.SEQUENCE(5)</c>).
/// Excel strips the prefix back out when displaying the formula to the user,
/// so the round-trip is transparent.
///
/// Also handles <c>_xlfn._xlws.</c> (worksheet-only namespace) for FILTER
/// and <c>_xlfn.ANCHORARRAY</c> for spilled-range references (<c>A1#</c> stays
/// user-facing; the XML serialization is a separate concern handled by Excel).
/// </summary>
public static class ModernFunctionQualifier
{
    // Functions that need just _xlfn.
    // Source: MS-XLSX / Excel 2016+ dynamic-array + modern function catalogue.
    private static readonly HashSet<string> XlfnFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "SEQUENCE", "SORT", "SORTBY", "UNIQUE",
        "XLOOKUP", "XMATCH",
        "LET", "LAMBDA",
        "IFS", "SWITCH",
        "MAXIFS", "MINIFS",
        "CONCAT", "TEXTJOIN",
        "STOCKHISTORY",
        "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
        "TAKE", "DROP",
        "CHOOSECOLS", "CHOOSEROWS",
        "ARRAYTOTEXT", "VALUETOTEXT",
        "TOCOL", "TOROW",
        "WRAPCOLS", "WRAPROWS",
        "EXPAND",
        "ANCHORARRAY",
    };

    // Functions that need _xlfn._xlws. (dynamic-array, worksheet-only)
    private static readonly HashSet<string> XlwsFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "FILTER",
    };

    // Match a bare function name (identifier followed by '('), not preceded by
    // a '.' or alphanumeric (so _xlfn.SEQUENCE and MYSEQUENCE are skipped),
    // and not inside a quoted string literal.
    private static readonly Regex FunctionCallRegex = new(
        @"(?<![A-Za-z0-9_\.])([A-Za-z_][A-Za-z0-9_]*)\s*\(",
        RegexOptions.Compiled);

    /// <summary>
    /// Returns the formula with Excel 2016+ modern function names qualified
    /// with <c>_xlfn.</c> / <c>_xlfn._xlws.</c> as required by OOXML. Leaves
    /// already-qualified names, older functions, quoted string literals, and
    /// non-function identifiers untouched.
    /// </summary>
    public static string Qualify(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return formula;

        // Walk the string and only rewrite identifiers outside quoted strings.
        // Excel formula strings are bounded by '"' with '""' as an escape.
        var sb = new System.Text.StringBuilder(formula.Length + 32);
        int i = 0;
        while (i < formula.Length)
        {
            char c = formula[i];
            if (c == '"')
            {
                // Copy the entire string literal verbatim.
                sb.Append(c);
                i++;
                while (i < formula.Length)
                {
                    sb.Append(formula[i]);
                    if (formula[i] == '"')
                    {
                        // escaped "" → consume both, stay in string
                        if (i + 1 < formula.Length && formula[i + 1] == '"')
                        {
                            sb.Append('"');
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    i++;
                }
                continue;
            }

            // Outside a string: scan for an identifier-call.
            // Use regex-on-substring is awkward; instead detect manually.
            if (IsIdentStart(c) && (i == 0 || !IsIdentPrev(formula[i - 1])))
            {
                int start = i;
                while (i < formula.Length && IsIdentCont(formula[i])) i++;
                // Skip whitespace then check for '('
                int j = i;
                while (j < formula.Length && formula[j] == ' ') j++;
                if (j < formula.Length && formula[j] == '(')
                {
                    var name = formula.Substring(start, i - start);
                    if (XlwsFunctions.Contains(name))
                        sb.Append("_xlfn._xlws.").Append(name);
                    else if (XlfnFunctions.Contains(name))
                        sb.Append("_xlfn.").Append(name);
                    else
                        sb.Append(name);
                }
                else
                {
                    sb.Append(formula, start, i - start);
                }
                continue;
            }

            sb.Append(c);
            i++;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Inverse of <see cref="Qualify"/> for readback: strips the
    /// <c>_xlfn.</c> / <c>_xlfn._xlws.</c> prefix so users see canonical
    /// function names instead of the OOXML-internal namespaced form.
    /// </summary>
    public static string Unqualify(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return formula;
        // Longer prefix first so we don't leave _xlws. stragglers.
        var s = formula.Replace("_xlfn._xlws.", "", StringComparison.Ordinal);
        s = s.Replace("_xlfn.", "", StringComparison.Ordinal);
        return s;
    }

    private static bool IsIdentStart(char c) => char.IsLetter(c) || c == '_';
    private static bool IsIdentCont(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.';
    // Prev char that would mean we're in the middle of an existing identifier
    // (incl. already-qualified `_xlfn.NAME`).
    private static bool IsIdentPrev(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.';
}
