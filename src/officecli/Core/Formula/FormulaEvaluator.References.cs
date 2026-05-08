// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Unresolved cell or area reference, kept first-class so that OFFSET / INDIRECT
/// can manipulate the reference itself instead of receiving a dereferenced value.
/// Single-cell refs use Width=Height=1.
/// </summary>
internal record RefArg(string? Sheet, int Col, int Row, int Width, int Height);

internal partial class FormulaEvaluator
{
    /// <summary>
    /// Convert a token-level range expression like "A1:B3" (or "A:A", "1:1")
    /// to a RefArg. Sheet-prefixed forms pass the sheet name in via parameter.
    /// </summary>
    private RefArg? BuildRefFromRange(string? sheet, string rangeExpr)
    {
        var parts = rangeExpr.Split(':');
        if (parts.Length != 2) return null;
        var left = StripDollar(parts[0]);
        var right = StripDollar(parts[1]);

        var leftColOnly = Regex.IsMatch(left, @"^[A-Z]+$", RegexOptions.IgnoreCase);
        var rightColOnly = Regex.IsMatch(right, @"^[A-Z]+$", RegexOptions.IgnoreCase);
        var leftRowOnly = Regex.IsMatch(left, @"^\d+$");
        var rightRowOnly = Regex.IsMatch(right, @"^\d+$");

        int r1, r2, c1, c2;
        if (leftColOnly && rightColOnly)
        {
            c1 = ColToIndex(left.ToUpperInvariant());
            c2 = ColToIndex(right.ToUpperInvariant());
            var target = GetSheetDataFor(sheet);
            if (target == null) return null;
            var (minRow, maxRow) = GetPopulatedRowRange(target);
            if (maxRow == 0) return null;
            r1 = minRow; r2 = maxRow;
        }
        else if (leftRowOnly && rightRowOnly)
        {
            r1 = int.Parse(left); r2 = int.Parse(right);
            var target = GetSheetDataFor(sheet);
            if (target == null) return null;
            var (minCol, maxCol) = GetPopulatedColRange(target);
            if (maxCol == 0) return null;
            c1 = minCol; c2 = maxCol;
        }
        else
        {
            var (col1, row1) = ParseRef(left);
            var (col2, row2) = ParseRef(right);
            c1 = ColToIndex(col1); c2 = ColToIndex(col2);
            r1 = row1; r2 = row2;
        }

        var colMin = Math.Min(c1, c2); var colMax = Math.Max(c1, c2);
        var rowMin = Math.Min(r1, r2); var rowMax = Math.Max(r1, r2);
        return new RefArg(sheet, colMin, rowMin, colMax - colMin + 1, rowMax - rowMin + 1);
    }

    /// <summary>
    /// Parse a reference string (e.g. "A1", "Sheet1!B2", "A1:C3") into a RefArg.
    /// Used by INDIRECT to convert its evaluated string argument into a reference.
    /// </summary>
    private RefArg? ParseRefString(string s)
    {
        s = s.Trim();
        if (string.IsNullOrEmpty(s)) return null;
        string? sheet = null;
        var bang = s.IndexOf('!');
        if (bang > 0)
        {
            sheet = s[..bang].Trim('\'');
            s = s[(bang + 1)..];
        }
        s = StripDollar(s);
        if (s.Contains(':')) return BuildRefFromRange(sheet, s);
        if (IsCellRef(s))
        {
            var (col, row) = ParseRef(s);
            return new RefArg(sheet, ColToIndex(col), row, 1, 1);
        }
        return null;
    }

    /// <summary>
    /// Resolve a RefArg to the actual cell values. Single-cell → scalar
    /// FormulaResult; multi-cell → FormulaResult.Area wrapping a RangeData.
    /// </summary>
    private FormulaResult? ResolveRef(RefArg r)
    {
        if (r.Width == 1 && r.Height == 1)
        {
            var cellRef = $"{IndexToCol(r.Col)}{r.Row}";
            return r.Sheet != null
                ? ResolveSheetCellResult($"{r.Sheet}!{cellRef}")
                : ResolveCellResult(cellRef);
        }
        var cells = new FormulaResult?[r.Height, r.Width];
        for (int dr = 0; dr < r.Height; dr++)
            for (int dc = 0; dc < r.Width; dc++)
            {
                var cellRef = $"{IndexToCol(r.Col + dc)}{r.Row + dr}";
                cells[dr, dc] = r.Sheet != null
                    ? ResolveSheetCellResult($"{r.Sheet}!{cellRef}")
                    : ResolveCellResult(cellRef);
            }
        return FormulaResult.Area(new RangeData(cells));
    }

    /// <summary>
    /// OFFSET(reference, rows, cols, [height], [width]).
    /// Returns the value at the offset position (single cell) or an Area result
    /// (multi-cell). Outer functions like SUM/AVERAGE consume the Area through
    /// the IsRange handling in helpers.
    /// </summary>
    private FormulaResult? EvalOffset(List<object> args)
    {
        if (args.Count < 3 || args.Count > 5) return FormulaResult.Error("#VALUE!");
        if (args[0] is not RefArg baseRef) return FormulaResult.Error("#VALUE!");

        int rowOffset = (int)((args[1] as FormulaResult)?.AsNumber() ?? 0);
        int colOffset = (int)((args[2] as FormulaResult)?.AsNumber() ?? 0);
        int height = baseRef.Height;
        int width = baseRef.Width;
        if (args.Count >= 4 && args[3] is FormulaResult hArg) height = (int)hArg.AsNumber();
        if (args.Count >= 5 && args[4] is FormulaResult wArg) width = (int)wArg.AsNumber();
        if (height == 0 || width == 0) return FormulaResult.Error("#REF!");

        var newRow = baseRef.Row + rowOffset;
        var newCol = baseRef.Col + colOffset;
        if (height < 0) { newRow += height + 1; height = -height; }
        if (width < 0) { newCol += width + 1; width = -width; }
        if (newRow < 1 || newCol < 1) return FormulaResult.Error("#REF!");

        return ResolveRef(new RefArg(baseRef.Sheet, newCol, newRow, width, height));
    }

    /// <summary>
    /// INDIRECT(ref_text). Only the A1-style form is supported (the [a1] argument
    /// is accepted but ignored — R1C1 syntax is not implemented).
    /// </summary>
    private FormulaResult? EvalIndirect(List<object> args)
    {
        if (args.Count < 1) return FormulaResult.Error("#VALUE!");
        var s = (args[0] as FormulaResult)?.AsString();
        if (string.IsNullOrEmpty(s)) return FormulaResult.Error("#REF!");
        var refArg = ParseRefString(s);
        if (refArg == null) return FormulaResult.Error("#REF!");
        return ResolveRef(refArg);
    }
}
