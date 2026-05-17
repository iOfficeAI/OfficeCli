// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // Format keys that must NOT be emitted: derived (Get computes from cache),
    // diagnostic (relIds, cNvPr ids that resolve per package), or coordinate-
    // system (only meaningful in the source document). Same role as
    // WordBatchEmitter.SkipKeys.
    // CONSISTENCY(emit-filter-mirror): see WordBatchEmitter.Filters.cs:14.
    private static readonly HashSet<string> PptxSkipKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        // Internal relationship id — unstable across packages, see WordBatchEmitter.
        "relId",
        // Cached display content for unevaluated fields. The `evaluated`
        // protocol surfaces this for diagnostic Get only; replay would
        // re-emit an a:fld with stale text.
        "evaluated",
        // Aggregate child counts surface only on the Get tree (ChildCount).
        "shapeCount", "layoutCount",
        // Per-presentation metadata that auto-restamps (last-modified-by /
        // revision / created / modified). Mirrors Word's stance on
        // similar metadata.
        "revision", "lastModifiedBy", "created", "modified",
        // Default font + slide dimensions live at the root presentation
        // node, not slide-level — they roll up into a single root `set /`
        // bag in PR2 (or are already set on the blank-doc baseline).
        "defaultFont",
    };

    private static Dictionary<string, string> FilterEmittableProps(Dictionary<string, object?> raw)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (key, val) in raw)
        {
            if (PptxSkipKeys.Contains(key)) continue;
            // CONSISTENCY(effective-X-mirror): docx WordBatchEmitter.Filters.cs
            // applies the same `effective.*` prefix filter — those are read-only
            // cascade snapshots, never user-settable.
            if (key.StartsWith("effective.", StringComparison.OrdinalIgnoreCase)) continue;
            if (val == null) continue;
            string s = val switch
            {
                bool b => b ? "true" : "false",
                _ => val.ToString() ?? ""
            };
            if (s.Length > 0) result[key] = s;
        }

        return result;
    }
}
