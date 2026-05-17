// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // PR2: emit speaker notes as a single `add notes parent=/slide[N]` row
    // carrying the concatenated body text. AddNotes accepts only `text=` /
    // `direction=` / `lang=` for the notes body — there is no typed Add
    // path for arbitrary shapes on a notesSlide, so the docx-mirror
    // "walk spTree as shape tree" approach has no replay surface to land
    // on. Emit only the body-placeholder text + direction/lang carried by
    // the `/slide[N]/notes` Get node. Mirrors AddNotes's input vocabulary.
    private static void EmitNotes(PowerPointHandler ppt, string slidePath,
                                  List<BatchItem> items, SlideEmitContext ctx)
    {
        var slideMatch = System.Text.RegularExpressions.Regex.Match(slidePath, @"^/slide\[(\d+)\]$");
        if (!slideMatch.Success) return;
        var slideIdx = int.Parse(slideMatch.Groups[1].Value);
        if (!ppt.SlideHasNotes(slideIdx)) return;

        DocumentNode notes;
        try { notes = ppt.Get($"{slidePath}/notes"); }
        catch { return; }
        if (notes.Type == "error") return;

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrEmpty(notes.Text))
            props["text"] = notes.Text!;
        // Direction/lang are surfaced on the notes node Format bag by
        // AddNotes round-trip; forward them through the same canonical
        // keys AddNotes accepts.
        foreach (var key in new[] { "direction", "lang" })
        {
            if (notes.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString() ?? "";
                if (s.Length > 0) props[key] = s;
            }
        }
        if (props.Count == 0) return;

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = slidePath,
            Type = "notes",
            Props = props,
        });
    }
}
