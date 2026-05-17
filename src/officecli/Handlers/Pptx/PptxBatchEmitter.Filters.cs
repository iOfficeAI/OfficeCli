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
        // OOXML cNvPr id — auto-accumulated by GenerateUniqueShapeId on Add.
        // Emitting the source id would force Add to honor it, which works for
        // free-form shapes but collides on slides where the layout already
        // contributes a placeholder at the same id slot (animations/video deck).
        // Mirrors docx WordBatchEmitter.Filters.cs treating paraId as derived:
        // emit uses positional /slide[N]/shape[K] for downstream addressing.
        "id",
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
        // Get emits both fill=gradient (type marker) and gradient=<spec> (params).
        // ApplyShapeFill would try parsing "gradient" as a color and reject; the
        // spec via gradient= already drives the fill. Same logic for pattern.
        if (result.TryGetValue("fill", out var fillVal))
        {
            if (fillVal.Equals("gradient", StringComparison.OrdinalIgnoreCase) && result.ContainsKey("gradient"))
                result.Remove("fill");
            else if (fillVal.Equals("pattern", StringComparison.OrdinalIgnoreCase) && result.ContainsKey("pattern"))
                result.Remove("fill");
        }

        return result;
    }
}
