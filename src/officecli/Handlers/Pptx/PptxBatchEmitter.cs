// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// CONSISTENCY(emit-X-mirror): scaffold mirrors WordBatchEmitter.cs — same
// public entry shape (full-doc + subtree overloads), same Get-driven
// transcription, same partial-class split (entry / Filters / Shape / Notes).
//
// PR1 scope (text-only): slide / shape / textbox / title / connector /
// group / placeholder + paragraph + run. Tables, pictures, charts, notes
// bodies, layout/master/theme raw — PR2.
public static partial class PptxBatchEmitter
{
    /// <summary>
    /// Carry-state for one emit run. Mirrors WordBatchEmitter.BodyEmitContext
    /// but trimmed for PR1 (no footnote/endnote/chart cursors yet —
    /// PowerPoint has no notes-with-numbering concept; chart/table content
    /// lands in PR2).
    /// </summary>
    internal sealed record SlideEmitContext(
        List<UnsupportedWarning> Unsupported);

    /// <summary>
    /// Captured at emit time when a slide carries content we cannot round-trip
    /// through the existing handler vocabulary (animations, SmartArt, OLE,
    /// video/audio, exotic transitions). The slide itself is emitted; the
    /// unsupported element is dropped silently from `items` but recorded
    /// here so the CLI can surface a warning bundle to the caller.
    /// </summary>
    public sealed record UnsupportedWarning(string Element, string SlidePath, string Reason);

    /// <summary>
    /// Emit a full PowerPoint document as a sequence of BatchItem rows.
    /// Returns the items plus any unsupported-element warnings.
    /// </summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitPptx(PowerPointHandler ppt)
    {
        var items = new List<BatchItem>();
        var ctx = new SlideEmitContext(new List<UnsupportedWarning>());

        // CONSISTENCY(slide-order): always iterate via the handler's
        // GetSlideParts() (sldIdLst-driven). Walking SlideParts off the
        // package returns parts in zip URI order — `slide12.xml` sorts
        // before `slide3.xml`, scrambling user-visible order.
        var slideTree = ppt.Get("/");
        if (slideTree.Children == null) return (items, ctx.Unsupported);

        int slideNum = 0;
        foreach (var slideNode in slideTree.Children)
        {
            if (slideNode.Type != "slide") continue;
            slideNum++;
            EmitSlide(ppt, slideNode, slideNum, items, ctx);
        }

        return (items, ctx.Unsupported);
    }

    /// <summary>
    /// Emit a subtree of a PowerPoint document. Supported subtree paths:
    /// `/slide[N]`. Other paths fall through to a NotImplementedException
    /// for now — PR3 will widen the entry surface when the CLI is wired up.
    /// </summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitPptx(
        PowerPointHandler ppt, string path)
    {
        if (string.IsNullOrEmpty(path))
            throw new CliException("dump path cannot be empty. Use '/' for the full document or /slide[N].")
                { Code = "invalid_path" };
        if (path == "/") return EmitPptx(ppt);

        var items = new List<BatchItem>();
        var ctx = new SlideEmitContext(new List<UnsupportedWarning>());

        var slideMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/slide\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (slideMatch.Success)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            DocumentNode slideNode;
            try { slideNode = ppt.Get(path); }
            catch (Exception ex)
            {
                throw new CliException($"dump path not found: {path} ({ex.Message})") { Code = "path_not_found" };
            }
            EmitSlide(ppt, slideNode, idx, items, ctx);
            return (items, ctx.Unsupported);
        }

        throw new CliException(
            $"dump path not supported: {path}. Supported: /, /slide[N]")
            { Code = "unsupported_path" };
    }

    private static void EmitSlide(PowerPointHandler ppt, DocumentNode slideNode, int slideNum,
                                  List<BatchItem> items, SlideEmitContext ctx)
    {
        var slidePath = slideNode.Path;
        ProbeUnsupportedOnSlide(ppt, slidePath, ctx);

        // Pull the full slide node so layout / hidden / background etc. surface
        // even when the entry passed us a depth-truncated tree from "/".
        var fullSlide = ppt.Get(slidePath);
        var slideProps = FilterEmittableProps(fullSlide.Format);

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/",
            Type = "slide",
            Props = slideProps.Count > 0 ? slideProps : null,
        });

        // ShapeToNode tags placeholder shapes as plain "textbox"/"title". To
        // emit them as `add placeholder` we cross-reference each shape's cNvPr
        // id with the slide's Query("placeholder") result.
        var placeholderById = new Dictionary<string, DocumentNode>(StringComparer.Ordinal);
        foreach (var ph in ppt.Query("placeholder"))
        {
            if (!ph.Path.StartsWith(slidePath + "/", StringComparison.Ordinal)) continue;
            if (ph.Format.TryGetValue("id", out var phId) && phId != null)
                placeholderById[phId.ToString()!] = ph;
        }

        // Children: walk shape-tree level. Get already routed group/connector/
        // textbox/title/equation into typed nodes, so just iterate and dispatch.
        if (fullSlide.Children == null) return;
        foreach (var child in fullSlide.Children)
        {
            // Placeholder dispatch first — overrides textbox/title type.
            if ((child.Type == "textbox" || child.Type == "title" || child.Type == "shape")
                && child.Format.TryGetValue("id", out var cid) && cid != null
                && placeholderById.TryGetValue(cid.ToString()!, out var phNode))
            {
                EmitPlaceholder(ppt, phNode, slidePath, items, ctx);
                continue;
            }
            switch (child.Type)
            {
                case "textbox":
                case "title":
                case "shape":
                case "equation":
                    EmitShape(ppt, child, slidePath, items, ctx);
                    break;
                case "placeholder":
                    EmitPlaceholder(ppt, child, slidePath, items, ctx);
                    break;
                case "connector":
                    EmitConnector(ppt, child, slidePath, items, ctx);
                    break;
                case "group":
                    EmitGroup(ppt, child, slidePath, items, ctx);
                    break;
                case "table":
                case "picture":
                case "chart":
                case "ole":
                case "video":
                case "audio":
                case "3dmodel":
                case "model3d":
                case "zoom":
                    // PR2 scope — emit unsupported marker so the caller knows
                    // these were seen but not transcribed yet. Not an animation
                    // / SmartArt warning per se but conveys the same gap.
                    ctx.Unsupported.Add(new UnsupportedWarning(
                        Element: child.Type ?? "unknown",
                        SlidePath: slidePath,
                        Reason: "deferred to PR2"));
                    break;
                default:
                    ctx.Unsupported.Add(new UnsupportedWarning(
                        Element: child.Type ?? "unknown",
                        SlidePath: slidePath,
                        Reason: "unrecognized child type"));
                    break;
            }
        }

        // Notes body content — stub for PR1. Notes part presence does not
        // surface in the slide subtree's children today (notes live under
        // /slide[N]/notes); PR2 will reach in and emit them.
        EmitNotes(ppt, slidePath, items, ctx);
    }

    // Touch the raw slide XML to find content that has no handler vocabulary
    // yet. Each match adds an UnsupportedWarning entry; we never throw.
    private static void ProbeUnsupportedOnSlide(PowerPointHandler ppt, string slidePath,
                                                SlideEmitContext ctx)
    {
        string xml;
        try { xml = ppt.Raw(slidePath); }
        catch { return; }

        // <p:timing> = slide animation. Cheapest substring test is sufficient —
        // the element name is unique within slide XML.
        if (xml.Contains("<p:timing", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("animation", slidePath,
                "<p:timing> animation tree present"));

        // SmartArt sits inside a graphicFrame as a dgm:relIds element.
        if (xml.Contains("dgm:relIds", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("smartArt", slidePath,
                "diagram (SmartArt) graphic frame present"));

        // OLE / video / audio / 3D — element names are distinctive enough.
        if (xml.Contains("<p:oleObj", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("oleObj", slidePath,
                "embedded OLE object present"));
        if (xml.Contains("<p:video", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("video", slidePath, "video element present"));
        if (xml.Contains("<p:audio", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("audio", slidePath, "audio element present"));
        if (xml.Contains("p:model3d", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("model3D", slidePath, "3D model present"));

        // Exotic transitions. Morph is most common; conveyor/ferris/honeycomb/
        // gallery live under p:transition's p15: extension list. Sniff the
        // transition element if present and tag by extension hint.
        // Vanilla transitions (fade/push/wipe/cut) already round-trip via
        // the `transition` prop, so they are NOT unsupported.
        var tIdx = xml.IndexOf("<p:transition", StringComparison.Ordinal);
        if (tIdx >= 0)
        {
            var tEnd = xml.IndexOf("</p:transition>", tIdx, StringComparison.Ordinal);
            var tSlice = tEnd > tIdx ? xml.Substring(tIdx, tEnd - tIdx) : xml.Substring(tIdx);
            if (tSlice.Contains("p159:morph", StringComparison.Ordinal)
                || tSlice.Contains("p15:morph", StringComparison.Ordinal)
                || tSlice.Contains("<p159:morph", StringComparison.Ordinal))
            {
                ctx.Unsupported.Add(new UnsupportedWarning("transition.morph", slidePath,
                    "morph transition uses p15: extension"));
            }
        }
    }
}
