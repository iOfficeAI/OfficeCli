// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(chart-data-string): mirrors WordBatchEmitter.Charts.cs —
    // emit a semantic `data="Name1:v1,v2;Name2:v3,v4"` string reconstructed
    // from series children that AddChart re-builds at replay. The embedded
    // xlsx (ppt/embeddings/Microsoft_Excel_Worksheet.xlsx) is lossy on
    // round-trip: formulas, conditional formatting, defined names from the
    // source workbook are dropped. Same trade-off as docx — chart visual
    // round-trips, chart workbook does not.
    private static void EmitChart(PowerPointHandler ppt, DocumentNode chartNode,
                                  string parentSlidePath, List<BatchItem> items,
                                  SlideEmitContext ctx)
    {
        // depth=1 so series children materialize with their name/values.
        var fullChart = ppt.Get(chartNode.Path, depth: 1);
        var props = FilterEmittableProps(fullChart.Format);
        // Strip Get-only keys AddChart neither expects nor accepts.
        props.Remove("id");
        props.Remove("seriesCount");

        // Reconstruct AddChart's data="Name1:v1,v2;..." input from the
        // series children (each carries `name` + `values` Format keys).
        var seriesParts = new List<string>();
        if (fullChart.Children != null)
        {
            foreach (var s in fullChart.Children)
            {
                if (s.Type != "series") continue;
                if (!s.Format.TryGetValue("name", out var nObj) || nObj == null) continue;
                if (!s.Format.TryGetValue("values", out var vObj) || vObj == null) continue;
                var name = nObj.ToString() ?? "";
                var vals = vObj.ToString() ?? "";
                if (name.Length == 0 || vals.Length == 0) continue;
                seriesParts.Add($"{name}:{vals}");
            }
        }
        if (seriesParts.Count > 0)
            props["data"] = string.Join(";", seriesParts);

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "chart",
            Props = props.Count > 0 ? props : null,
        });
    }
}
