// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(emit-resources-mirror): mirrors WordBatchEmitter.Resources.cs
    // — each whole-part-XML block emits as a single raw-set replace. Theme /
    // master / layout / notesMaster carry rich structured XML (clrScheme,
    // fontScheme, txStyles, fmtScheme, …) that has no typed Set vocabulary; the
    // natural operation is "swap the whole block". Replay's raw-set overwrites
    // whatever the blank deck stamped during BlankDocCreator.

    private static void EmitThemeRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        // Pptx Raw("/theme") returns the presentation-level theme part (first
        // master's theme). Multi-master decks have additional theme parts
        // attached to each master, but the existing Raw/RawSet surface only
        // addresses the primary one — keep parity until per-master theme
        // raw-set lands. Skip silently when the source has none.
        string xml;
        try { xml = ppt.Raw("/theme"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<") || xml == "(no theme)")
            return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/theme",
            Xpath = "/a:theme",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitNotesMasterRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        if (!ppt.HasNotesMaster) return;
        string xml;
        try { xml = ppt.Raw("/notesMaster"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/notesMaster",
            Xpath = "/p:notesMaster",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitMasterRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        var n = ppt.SlideMasterCount;
        for (int i = 1; i <= n; i++)
        {
            string xml;
            try { xml = ppt.Raw($"/slideMaster[{i}]"); }
            catch { continue; }
            if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) continue;

            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = $"/slideMaster[{i}]",
                Xpath = "/p:sldMaster",
                Action = "replace",
                Xml = xml
            });
        }
    }

    private static void EmitLayoutRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        var n = ppt.SlideLayoutCount;
        for (int i = 1; i <= n; i++)
        {
            string xml;
            try { xml = ppt.Raw($"/slideLayout[{i}]"); }
            catch { continue; }
            if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) continue;

            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = $"/slideLayout[{i}]",
                Xpath = "/p:sldLayout",
                Action = "replace",
                Xml = xml
            });
        }
    }
}
