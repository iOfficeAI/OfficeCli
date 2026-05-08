// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Direction of insertion or deletion that triggered a formula reference shift.
/// Insert directions shift refs by +1; delete directions shift refs by -1
/// and collapse refs that landed on the deleted index into <c>#REF!</c>.
/// </summary>
public enum FormulaShiftDirection
{
    /// <summary>A column was inserted; cell-ref columns at or past insertIdx shift right by 1.</summary>
    ColumnsRight,
    /// <summary>A row was inserted; cell-ref rows at or past insertIdx shift down by 1.</summary>
    RowsDown,
    /// <summary>A column was deleted; refs to that column collapse to <c>#REF!</c>, refs past it shift left by 1.</summary>
    ColumnsLeft,
    /// <summary>A row was deleted; refs to that row collapse to <c>#REF!</c>, refs past it shift up by 1.</summary>
    RowsUp,
}

/// <summary>
/// Rewrites Excel formula text after a column or row was inserted, so that
/// references that previously pointed to a moved cell continue to point to
/// the same cell.
///
/// <para>This is the regex-based "good enough" implementation (Path A). It
/// handles the common ~90% of formulas: A1 / $A$1 / $A1 / A$1 single refs,
/// A1:B5 ranges, sheet-qualified refs (Sheet2!A1, 'Sheet With Spaces'!A1),
/// and skips string literals and structured-ref bracket content. It does
/// NOT handle: cross-workbook refs ([Book]Sheet!A1), R1C1 notation,
/// whole-column (A:A) or whole-row (1:1) refs, or structured table refs
/// (Table1[Col1]) — those pass through verbatim.</para>
///
/// <para>The public API is intentionally minimal so a future tokenizer-based
/// implementation (Path B) can replace the body of <see cref="Shift"/>
/// without touching call sites or tests.</para>
/// </summary>
public static class FormulaRefShifter
{
    // One regex matches either a single A1 ref or a range, optionally
    // sheet-qualified. Whole-col / whole-row refs are NOT matched here —
    // they require digits in r1, which is mandatory in this pattern.
    //
    // Capture groups:
    //   sheet  — optional sheet name (with surrounding quotes preserved)
    //   c1, r1 — first cell column letters (with optional leading $) and row digits
    //   c2, r2 — range end (or empty for single-cell)
    private static readonly Regex CellRefPattern = new(
        @"(?<![\w.])" +
        @"(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?" +
        @"(?<c1>\$?[A-Z]{1,3})(?<r1>\$?\d+)" +
        @"(?::(?<c2>\$?[A-Z]{1,3})(?<r2>\$?\d+))?" +
        // (?![\w(]) — also reject when followed by '(' so that function names
        // shaped like `LOG10` / `ATAN2` (col-letters + row-digits) are not
        // misread as cell refs. Cell refs are never followed by '('.
        @"(?![\w(])",
        RegexOptions.Compiled);

    /// <summary>
    /// Rewrite cell references in a formula by remapping their row numbers
    /// through an arbitrary <paramref name="oldToNewRow"/> mapping. Used by
    /// move/reorder operations where the change is not a uniform +1/-1 shift
    /// but a permutation of row indices (e.g. moving row 3 before row 2
    /// produces the map {1→1, 2→3, 3→2}).
    ///
    /// <para>Refs whose row is not in the map pass through unchanged. For
    /// ranges, both endpoints are remapped; if the result inverts (start &gt;
    /// end) the original range is returned unchanged. Sheet-scope, string-
    /// literal, and structured-ref skip rules are identical to <see cref="Shift"/>.</para>
    /// </summary>
    public static string ApplyRowRenumberMap(
        string formula,
        string currentSheet,
        string modifiedSheet,
        IReadOnlyDictionary<int, int> oldToNewRow)
    {
        if (string.IsNullOrEmpty(formula) || oldToNewRow.Count == 0) return formula;
        return WalkFormulaTokens(formula, chunk =>
            RenumberRefsInChunk(chunk, currentSheet, modifiedSheet, oldToNewRow));
    }

    /// <summary>
    /// Outer tokenize-skip walker shared by every <c>FormulaRefShifter</c>
    /// public entry point. Streams the formula char-by-char, copying string
    /// literals (with the Excel <c>""</c> doubling escape) and bracket
    /// content (structured refs like <c>Table1[Col1]</c>; cross-workbook
    /// prefixes like <c>[Book2]Sheet1!A1</c>) verbatim. Hands every other
    /// contiguous chunk to <paramref name="chunkProcessor"/>, which runs
    /// the per-match cell-ref rewrite for that semantic (shift / renumber /
    /// copy-delta) and returns the rewritten chunk.
    /// </summary>
    private static string WalkFormulaTokens(string formula, Func<string, string> chunkProcessor)
    {
        var sb = new StringBuilder(formula.Length);
        int i = 0;
        while (i < formula.Length)
        {
            char ch = formula[i];
            if (ch == '"')
            {
                sb.Append(ch); i++;
                while (i < formula.Length)
                {
                    sb.Append(formula[i]);
                    if (formula[i] == '"')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '"')
                        { sb.Append(formula[i + 1]); i += 2; continue; }
                        i++; break;
                    }
                    i++;
                }
            }
            else if (ch == '[')
            {
                int depth = 0;
                while (i < formula.Length)
                {
                    char c = formula[i];
                    sb.Append(c);
                    if (c == '[') depth++;
                    else if (c == ']') { depth--; if (depth == 0) { i++; break; } }
                    i++;
                }
            }
            else
            {
                int start = i;
                while (i < formula.Length && formula[i] != '"' && formula[i] != '[') i++;
                sb.Append(chunkProcessor(formula.AsSpan(start, i - start).ToString()));
            }
        }
        return sb.ToString();
    }

    private static string RenumberRefsInChunk(
        string chunk, string currentSheet, string modifiedSheet,
        IReadOnlyDictionary<int, int> oldToNewRow)
    {
        return CellRefPattern.Replace(chunk, m =>
        {
            var sheetGroup = m.Groups["sheet"].Value;
            string targetSheet = string.IsNullOrEmpty(sheetGroup)
                ? currentSheet
                : (sheetGroup.StartsWith('\'') && sheetGroup.EndsWith('\'')
                    ? sheetGroup[1..^1].Replace("''", "'")
                    : sheetGroup);
            if (!targetSheet.Equals(modifiedSheet, StringComparison.OrdinalIgnoreCase))
                return m.Value;

            string c1 = m.Groups["c1"].Value;
            string r1 = m.Groups["r1"].Value;
            string c2 = m.Groups["c2"].Value;
            string r2 = m.Groups["r2"].Value;
            bool isRange = !string.IsNullOrEmpty(c2);
            string sheetPrefix = string.IsNullOrEmpty(sheetGroup) ? "" : sheetGroup + "!";

            string newR1 = RemapRow(r1, oldToNewRow);
            if (!isRange) return $"{sheetPrefix}{c1}{newR1}";

            string newR2 = RemapRow(r2, oldToNewRow);
            // The range covers a contiguous SET of rows [r1..r2]. After
            // renumber, that set must remain contiguous (and represent
            // the same row content) for the new range to be a faithful
            // rewrite. If the mapped set is not contiguous or doesn't
            // match [min..max] of the new endpoints, fall back to the
            // original text rather than write a misleading ref.
            int Parse(string s) => int.Parse(s.StartsWith('$') ? s[1..] : s);
            int oldR1 = Parse(r1), oldR2 = Parse(r2);
            int newR1Int = Parse(newR1), newR2Int = Parse(newR2);
            if (!RangeRemapStillContiguous(oldR1, oldR2, newR1Int, newR2Int, oldToNewRow))
                return m.Value;

            return $"{sheetPrefix}{c1}{newR1}:{c2}{newR2}";
        });
    }

    private static bool RangeRemapStillContiguous(
        int oldStart, int oldEnd, int newStart, int newEnd,
        IReadOnlyDictionary<int, int> map)
    {
        if (oldStart > oldEnd) return false;
        int newMin = Math.Min(newStart, newEnd);
        int newMax = Math.Max(newStart, newEnd);
        // Build the mapped set and check it equals [newMin..newMax] exactly.
        var mappedSet = new HashSet<int>();
        for (int i = oldStart; i <= oldEnd; i++)
        {
            int mapped = map.TryGetValue(i, out var n) ? n : i;
            mappedSet.Add(mapped);
        }
        if (mappedSet.Count != (newMax - newMin + 1)) return false;
        for (int i = newMin; i <= newMax; i++)
            if (!mappedSet.Contains(i)) return false;
        return newStart <= newEnd;
    }

    private static string RemapRow(string rowPart, IReadOnlyDictionary<int, int> map)
    {
        bool abs = rowPart.StartsWith('$');
        int oldNum = int.Parse(abs ? rowPart[1..] : rowPart);
        if (!map.TryGetValue(oldNum, out var newNum)) return rowPart;
        return (abs ? "$" : "") + newNum;
    }

    /// <summary>
    /// Column-axis variant of <see cref="ApplyRowRenumberMap"/>. Same skip
    /// rules, sheet scope, and contiguity guard. Map keys/values are 1-based
    /// column indices (A=1, B=2, ...).
    /// </summary>
    public static string ApplyColRenumberMap(
        string formula,
        string currentSheet,
        string modifiedSheet,
        IReadOnlyDictionary<int, int> oldToNewCol)
    {
        if (string.IsNullOrEmpty(formula) || oldToNewCol.Count == 0) return formula;
        return WalkFormulaTokens(formula, chunk =>
            RenumberColRefsInChunk(chunk, currentSheet, modifiedSheet, oldToNewCol));
    }

    private static string RenumberColRefsInChunk(
        string chunk, string currentSheet, string modifiedSheet,
        IReadOnlyDictionary<int, int> oldToNewCol)
    {
        return CellRefPattern.Replace(chunk, m =>
        {
            var sheetGroup = m.Groups["sheet"].Value;
            string targetSheet = string.IsNullOrEmpty(sheetGroup)
                ? currentSheet
                : (sheetGroup.StartsWith('\'') && sheetGroup.EndsWith('\'')
                    ? sheetGroup[1..^1].Replace("''", "'")
                    : sheetGroup);
            if (!targetSheet.Equals(modifiedSheet, StringComparison.OrdinalIgnoreCase))
                return m.Value;

            string c1 = m.Groups["c1"].Value;
            string r1 = m.Groups["r1"].Value;
            string c2 = m.Groups["c2"].Value;
            string r2 = m.Groups["r2"].Value;
            bool isRange = !string.IsNullOrEmpty(c2);
            string sheetPrefix = string.IsNullOrEmpty(sheetGroup) ? "" : sheetGroup + "!";

            string newC1 = RemapCol(c1, oldToNewCol);
            if (!isRange) return $"{sheetPrefix}{newC1}{r1}";

            string newC2 = RemapCol(c2, oldToNewCol);
            int Idx(string s) => ColumnLettersToIndex(s.StartsWith('$') ? s[1..] : s);
            int oldC1Idx = Idx(c1), oldC2Idx = Idx(c2);
            int newC1Idx = Idx(newC1), newC2Idx = Idx(newC2);
            if (!RangeRemapStillContiguous(oldC1Idx, oldC2Idx, newC1Idx, newC2Idx, oldToNewCol))
                return m.Value;

            return $"{sheetPrefix}{newC1}{r1}:{newC2}{r2}";
        });
    }

    private static string RemapCol(string colPart, IReadOnlyDictionary<int, int> map)
    {
        bool abs = colPart.StartsWith('$');
        string letters = abs ? colPart[1..] : colPart;
        int oldIdx = ColumnLettersToIndex(letters);
        if (!map.TryGetValue(oldIdx, out var newIdx)) return colPart;
        return (abs ? "$" : "") + IndexToColumnLetters(newIdx);
    }

    /// <summary>
    /// Shift relative cell references in a formula by a (deltaCol, deltaRow)
    /// vector. Models Excel's "copy formula" semantics: refs without a $
    /// marker shift by the delta, refs with $ stay absolute. Used when a
    /// row or column is copied to a new position — the cloned formulas keep
    /// their relative spatial relationships but their literal text needs to
    /// reflect the new anchor cell.
    ///
    /// <para>Sheet-scope, string-literal, and structured-ref skip rules are
    /// identical to <see cref="Shift"/>. A ref whose absolute resulting row
    /// or column would be &lt;= 0 collapses to <c>#REF!</c>.</para>
    /// </summary>
    public static string ApplyCopyDelta(
        string formula,
        string currentSheet,
        string modifiedSheet,
        int deltaCol,
        int deltaRow)
    {
        if (string.IsNullOrEmpty(formula) || (deltaCol == 0 && deltaRow == 0)) return formula;
        return WalkFormulaTokens(formula, chunk =>
            DeltaShiftRefsInChunk(chunk, currentSheet, modifiedSheet, deltaCol, deltaRow));
    }

    private static string DeltaShiftRefsInChunk(
        string chunk, string currentSheet, string modifiedSheet,
        int deltaCol, int deltaRow)
    {
        return CellRefPattern.Replace(chunk, m =>
        {
            var sheetGroup = m.Groups["sheet"].Value;
            string targetSheet = string.IsNullOrEmpty(sheetGroup)
                ? currentSheet
                : (sheetGroup.StartsWith('\'') && sheetGroup.EndsWith('\'')
                    ? sheetGroup[1..^1].Replace("''", "'")
                    : sheetGroup);
            if (!targetSheet.Equals(modifiedSheet, StringComparison.OrdinalIgnoreCase))
                return m.Value;

            string c1 = m.Groups["c1"].Value;
            string r1 = m.Groups["r1"].Value;
            string c2 = m.Groups["c2"].Value;
            string r2 = m.Groups["r2"].Value;
            bool isRange = !string.IsNullOrEmpty(c2);
            string sheetPrefix = string.IsNullOrEmpty(sheetGroup) ? "" : sheetGroup + "!";

            string? newC1 = DeltaShiftCol(c1, deltaCol);
            string? newR1 = DeltaShiftRow(r1, deltaRow);
            if (newC1 == null || newR1 == null) return "#REF!";

            if (!isRange) return $"{sheetPrefix}{newC1}{newR1}";

            string? newC2 = DeltaShiftCol(c2, deltaCol);
            string? newR2 = DeltaShiftRow(r2, deltaRow);
            if (newC2 == null || newR2 == null) return "#REF!";
            return $"{sheetPrefix}{newC1}{newR1}:{newC2}{newR2}";
        });
    }

    private static string? DeltaShiftCol(string colPart, int delta)
    {
        bool abs = colPart.StartsWith('$');
        if (abs || delta == 0) return colPart;
        int idx = ColumnLettersToIndex(colPart);
        int newIdx = idx + delta;
        if (newIdx < 1) return null;
        return IndexToColumnLetters(newIdx);
    }

    private static string? DeltaShiftRow(string rowPart, int delta)
    {
        bool abs = rowPart.StartsWith('$');
        if (abs || delta == 0) return rowPart;
        int num = int.Parse(rowPart);
        int newNum = num + delta;
        if (newNum < 1) return null;
        return newNum.ToString();
    }

    /// <summary>
    /// Returns the formula text rewritten so that any references targeting
    /// <paramref name="modifiedSheet"/> at or past <paramref name="insertIdx"/>
    /// are shifted by 1 in <paramref name="direction"/>. Refs targeting other
    /// sheets, references inside string literals, and references inside
    /// structured-ref brackets are returned untouched.
    /// </summary>
    /// <param name="formula">Formula text without a leading '=' (matching how
    /// the Excel handler stores <c>CellFormula</c> content).</param>
    /// <param name="currentSheet">Sheet that contains the formula. Used to
    /// resolve unqualified refs.</param>
    /// <param name="modifiedSheet">Sheet on which the insert happened. Refs
    /// shift only when their resolved sheet equals this.</param>
    /// <param name="direction">Whether a column or row was inserted.</param>
    /// <param name="insertIdx">1-based column index (for ColumnsRight) or
    /// 1-based row index (for RowsDown) at which the insert happened.</param>
    /// <returns>The rewritten formula text. Returns the input unchanged when
    /// no refs match the shift criteria.</returns>
    public static string Shift(
        string formula,
        string currentSheet,
        string modifiedSheet,
        FormulaShiftDirection direction,
        int insertIdx)
    {
        if (string.IsNullOrEmpty(formula)) return formula;
        return WalkFormulaTokens(formula, chunk =>
            ShiftRefsInChunk(chunk, currentSheet, modifiedSheet, direction, insertIdx));
    }

    private static string ShiftRefsInChunk(
        string chunk, string currentSheet, string modifiedSheet,
        FormulaShiftDirection direction, int insertIdx)
    {
        return CellRefPattern.Replace(chunk, m =>
        {
            var sheetGroup = m.Groups["sheet"].Value;
            string targetSheet;
            if (string.IsNullOrEmpty(sheetGroup))
            {
                targetSheet = currentSheet;
            }
            else if (sheetGroup.StartsWith('\'') && sheetGroup.EndsWith('\''))
            {
                targetSheet = sheetGroup[1..^1].Replace("''", "'");
            }
            else
            {
                targetSheet = sheetGroup;
            }

            if (!targetSheet.Equals(modifiedSheet, StringComparison.OrdinalIgnoreCase))
                return m.Value;

            string c1 = m.Groups["c1"].Value;
            string r1 = m.Groups["r1"].Value;
            string c2 = m.Groups["c2"].Value;
            string r2 = m.Groups["r2"].Value;

            string sheetPrefix = string.IsNullOrEmpty(sheetGroup) ? "" : sheetGroup + "!";
            bool isRange = !string.IsNullOrEmpty(c2);

            // For each axis (col, row), compute the new value or null=#REF!.
            // For Insert directions, shifts never produce #REF!. For Delete
            // directions, an endpoint exactly on the deleted index either
            // collapses (single ref → #REF!), keeps the same row/col number
            // when it is the start endpoint of a range (now points to the
            // next row/col), or moves up/left when it is the end endpoint
            // (range shrinks by 1).
            var (newC1, newC2, colRef) = ShiftColAxis(c1, c2, isRange, direction, insertIdx);
            var (newR1, newR2, rowRef) = ShiftRowAxis(r1, r2, isRange, direction, insertIdx);
            if (colRef || rowRef) return "#REF!";

            if (!isRange)
                return $"{sheetPrefix}{newC1}{newR1}";

            return $"{sheetPrefix}{newC1}{newR1}:{newC2}{newR2}";
        });
    }

    private static (string newC1, string newC2, bool refError) ShiftColAxis(
        string c1, string c2, bool isRange, FormulaShiftDirection direction, int insertIdx)
    {
        switch (direction)
        {
            case FormulaShiftDirection.ColumnsRight:
                return (ShiftColPart(c1, insertIdx),
                        isRange ? ShiftColPart(c2, insertIdx) : c2,
                        false);
            case FormulaShiftDirection.ColumnsLeft:
                var (nc1, nc2, refErr) = DeleteShiftAxis(
                    c1, c2, isRange, insertIdx,
                    parseIdx: ColumnLettersToIndex,
                    formatIdx: (idx, abs) => (abs ? "$" : "") + IndexToColumnLetters(idx),
                    parseAbs: s => s.StartsWith('$'),
                    parseDigits: s => s.StartsWith('$') ? s[1..] : s);
                return (nc1, nc2, refErr);
            default:
                return (c1, c2, false);
        }
    }

    private static (string newR1, string newR2, bool refError) ShiftRowAxis(
        string r1, string r2, bool isRange, FormulaShiftDirection direction, int insertIdx)
    {
        switch (direction)
        {
            case FormulaShiftDirection.RowsDown:
                return (ShiftRowPart(r1, insertIdx),
                        isRange ? ShiftRowPart(r2, insertIdx) : r2,
                        false);
            case FormulaShiftDirection.RowsUp:
                var (nr1, nr2, refErr) = DeleteShiftAxis(
                    r1, r2, isRange, insertIdx,
                    parseIdx: s => int.Parse(s),
                    formatIdx: (idx, abs) => (abs ? "$" : "") + idx,
                    parseAbs: s => s.StartsWith('$'),
                    parseDigits: s => s.StartsWith('$') ? s[1..] : s);
                return (nr1, nr2, refErr);
            default:
                return (r1, r2, false);
        }
    }

    /// <summary>
    /// Shared delete-direction logic for both row and column axes. Returns
    /// the new endpoint strings and a refError flag set when the ref must
    /// collapse to <c>#REF!</c>.
    /// </summary>
    private static (string n1, string n2, bool refError) DeleteShiftAxis(
        string p1, string p2, bool isRange, int deletedIdx,
        Func<string, int> parseIdx,
        Func<int, bool, string> formatIdx,
        Func<string, bool> parseAbs,
        Func<string, string> parseDigits)
    {
        bool abs1 = parseAbs(p1);
        int idx1 = parseIdx(parseDigits(p1));

        if (!isRange)
        {
            if (idx1 == deletedIdx) return (p1, p2, true);
            if (idx1 > deletedIdx) return (formatIdx(idx1 - 1, abs1), p2, false);
            return (p1, p2, false);
        }

        bool abs2 = parseAbs(p2);
        int idx2 = parseIdx(parseDigits(p2));

        // Endpoint at deleted index: as start, stays at deletedIdx (now points
        // to the next survivor); as end, becomes deletedIdx-1 (range shrinks).
        int newIdx1 = idx1 == deletedIdx ? deletedIdx
                    : idx1 > deletedIdx ? idx1 - 1
                    : idx1;
        int newIdx2 = idx2 == deletedIdx ? deletedIdx - 1
                    : idx2 > deletedIdx ? idx2 - 1
                    : idx2;

        // Range collapsed past zero or inverted (e.g. A3:A3 with row 3 deleted).
        if (newIdx1 > newIdx2 || newIdx2 < 1) return (p1, p2, true);

        return (formatIdx(newIdx1, abs1), formatIdx(newIdx2, abs2), false);
    }

    private static string ShiftColPart(string colPart, int insertColIdx)
    {
        bool isAbs = colPart.StartsWith('$');
        string letters = isAbs ? colPart[1..] : colPart;
        int idx = ColumnLettersToIndex(letters);
        if (idx < insertColIdx) return colPart;
        return (isAbs ? "$" : "") + IndexToColumnLetters(idx + 1);
    }

    private static string ShiftRowPart(string rowPart, int insertRow)
    {
        bool isAbs = rowPart.StartsWith('$');
        int num = int.Parse(isAbs ? rowPart[1..] : rowPart);
        if (num < insertRow) return rowPart;
        return (isAbs ? "$" : "") + (num + 1);
    }

    // Local copies — keep Core/ free of Handlers/ dependencies so the shifter
    // can be used by any handler or tested in isolation.
    private static int ColumnLettersToIndex(string letters)
    {
        int idx = 0;
        foreach (char c in letters)
            idx = idx * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        return idx;
    }

    private static string IndexToColumnLetters(int idx)
    {
        var sb = new StringBuilder();
        while (idx > 0)
        {
            int rem = (idx - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            idx = (idx - 1) / 26;
        }
        return sb.ToString();
    }
}
