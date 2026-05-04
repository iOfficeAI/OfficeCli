// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Handlers;

namespace OfficeCli.Core;

/// <summary>
/// Walks an opened handler's document tree and emits a sequence of BatchItem
/// rows that, when replayed against a blank document of the same format,
/// reconstruct the original document.
///
/// <para>
/// This is the core of the `officecli dump --format batch` pipeline. The
/// emit relies on the OOXML schema reflection fallback in
/// <see cref="TypedAttributeFallback"/> + <see cref="GenericXmlQuery"/>:
/// any leaf property that Get reads can be re-applied via Add/Set, so
/// emit just transcribes Format keys directly without per-property
/// allowlisting.
/// </para>
///
/// <para>
/// Scope (v0.5): docx body paragraphs (with run formatting) + tables (single
/// paragraph + single run per cell, common case). Resources (styles,
/// numbering, theme, headers, footers, sections, comments, footnotes,
/// endnotes) and richer cell contents are NOT yet emitted — follow-up
/// passes will add them.
/// </para>
/// </summary>
public static class BatchEmitter
{
    /// <summary>Emit a batch sequence for a Word document.</summary>
    public static List<BatchItem> EmitWord(WordHandler word)
    {
        var items = new List<BatchItem>();
        EmitBody(word, items);
        return items;
    }

    private static void EmitBody(WordHandler word, List<BatchItem> items)
    {
        var bodyNode = word.Get("/body");
        if (bodyNode.Children == null) return;

        int pIndex = 0, tblIndex = 0;
        foreach (var child in bodyNode.Children)
        {
            switch (child.Type)
            {
                case "paragraph":
                case "p":
                    pIndex++;
                    EmitParagraph(word, child.Path, "/body", pIndex, items, autoPresent: false);
                    break;
                case "table":
                    tblIndex++;
                    EmitTable(word, child.Path, tblIndex, items);
                    break;
                case "section":
                case "sectPr":
                    // The body always carries one trailing sectPr that the
                    // blank document already provides; for v0.5 we rely on
                    // that default and skip emitting section properties.
                    // Section emit is a follow-up.
                    break;
                default:
                    // Unknown body-level child types (sdt, etc.) — skip for v0.5.
                    break;
            }
        }
    }

    /// <summary>
    /// Emit a paragraph at the target index under <paramref name="parentPath"/>.
    /// When <paramref name="autoPresent"/> is true, the parent already has a
    /// pre-existing paragraph at that index (e.g. an auto-created table cell
    /// paragraph); we issue a `set` instead of a fresh `add` so the existing
    /// paragraph gets reused rather than duplicated.
    /// </summary>
    private static void EmitParagraph(WordHandler word, string sourcePath, string parentPath,
                                      int targetIndex, List<BatchItem> items, bool autoPresent)
    {
        var pNode = word.Get(sourcePath);
        var props = FilterEmittableProps(pNode.Format);
        var runs = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "run" || c.Type == "r")
            .ToList();

        // Single-run / no-run paragraph: collapse run formatting into the
        // paragraph's prop bag (the schema-reflection layer accepts run-level
        // keys on a paragraph and routes them through ApplyRunFormatting).
        if (runs.Count <= 1)
        {
            if (runs.Count == 1)
            {
                var runProps = FilterEmittableProps(runs[0].Format);
                foreach (var (k, v) in runProps)
                {
                    if (!props.ContainsKey(k)) props[k] = v;
                }
                if (!string.IsNullOrEmpty(runs[0].Text))
                    props["text"] = runs[0].Text!;
            }

            if (autoPresent)
            {
                // Replace the auto-created paragraph in place — only push the
                // set when there is something to apply, otherwise the empty
                // skeleton is already correct.
                if (props.Count > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = $"{parentPath}/p[{targetIndex}]",
                        Props = props
                    });
                }
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = parentPath,
                    Type = "p",
                    Props = props.Count > 0 ? props : null
                });
            }
            return;
        }

        // Multi-run paragraph: emit/set the paragraph empty first, then add
        // each run as an explicit child.
        if (autoPresent)
        {
            if (props.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{parentPath}/p[{targetIndex}]",
                    Props = props
                });
            }
        }
        else
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "p",
                Props = props.Count > 0 ? props : null
            });
        }

        var paraTargetPath = $"{parentPath}/p[{targetIndex}]";
        foreach (var run in runs)
        {
            var rProps = FilterEmittableProps(run.Format);
            if (!string.IsNullOrEmpty(run.Text))
                rProps["text"] = run.Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "r",
                Props = rProps.Count > 0 ? rProps : null
            });
        }
    }

    private static void EmitTable(WordHandler word, string sourcePath, int targetIndex, List<BatchItem> items)
    {
        var tableNode = word.Get(sourcePath);
        var rows = (tableNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "row")
            .ToList();
        if (rows.Count == 0) return;

        // Pull cell count from the first row. Column count emitted by Get
        // (Format["cols"]) reflects the gridCol count, which can drift from
        // actual cells if the table has merges; the row's own count is the
        // safer bet for shape during replay.
        int cellsInFirstRow = 0;
        var row0 = word.Get(rows[0].Path);
        if (row0.Children != null)
            cellsInFirstRow = row0.Children.Count(c => c.Type == "cell");
        if (cellsInFirstRow == 0) return;

        var tableProps = FilterEmittableProps(tableNode.Format);
        tableProps["rows"] = rows.Count.ToString();
        tableProps["cols"] = cellsInFirstRow.ToString();
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "table",
            Props = tableProps
        });

        var tablePath = $"/body/tbl[{targetIndex}]";
        for (int r = 0; r < rows.Count; r++)
        {
            var rowNode = word.Get(rows[r].Path);
            var cells = (rowNode.Children ?? new List<DocumentNode>())
                .Where(c => c.Type == "cell")
                .ToList();
            for (int c = 0; c < cells.Count; c++)
            {
                var cellNode = word.Get(cells[c].Path);
                var cellTargetPath = $"{tablePath}/tr[{r + 1}]/tc[{c + 1}]";

                // Each cell carries auto-generated paragraphs (Add table seeds
                // one empty paragraph per cell). Update the first one in place
                // and append further paragraphs as fresh adds.
                var cellParas = (cellNode.Children ?? new List<DocumentNode>())
                    .Where(x => x.Type == "paragraph" || x.Type == "p")
                    .ToList();
                for (int p = 0; p < cellParas.Count; p++)
                {
                    EmitParagraph(word, cellParas[p].Path, cellTargetPath, p + 1, items,
                                  autoPresent: p == 0);
                }
            }
        }
    }

    // Format keys that must NOT be emitted: derived (computed by Get, not
    // user-set), unstable (regenerate on save), or coordinate-system
    // (paths that only make sense in the source document).
    private static readonly HashSet<string> SkipKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "basedOn.path",
        "paraId", "textId", "rsidR", "rsidRDefault", "rsidRPr", "rsidP", "rsidTr",
    };

    private static Dictionary<string, string> FilterEmittableProps(Dictionary<string, object?> raw)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (SkipKeys.Contains(key)) continue;
            if (key.StartsWith("effective.", StringComparison.OrdinalIgnoreCase)) continue;
            if (key.EndsWith(".cs.source", StringComparison.OrdinalIgnoreCase)) continue;

            // BORDER subattr asymmetry: Get exposes `border.top: single` AND
            // `border.top.sz: 4` / `border.top.color: 808080` as separate keys,
            // but Set's case table stops at the 2-segment level — the 3-segment
            // sub-attribute keys would be misrouted through ApplyTableBorders'
            // dotted fallback and crash on `Invalid border style: '4'`. Drop
            // them here as a known lossy projection until Set grows the
            // matching cases (border width / color readback survive only via
            // the main `border.*` style key for now).
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase) &&
                key.Count(ch => ch == '.') >= 2)
            {
                continue;
            }

            if (val == null) continue;
            var s = val switch
            {
                bool b => b ? "true" : "false",
                _ => val.ToString() ?? ""
            };
            if (s.Length > 0) result[key] = s;
        }
        return result;
    }
}
