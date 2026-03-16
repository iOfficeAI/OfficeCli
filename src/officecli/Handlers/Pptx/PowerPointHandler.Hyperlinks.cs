// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Hyperlink helpers ====================

    /// <summary>
    /// Apply a hyperlink URL to all runs in a shape. Pass "none" or "" to remove.
    /// </summary>
    private static void ApplyShapeHyperlink(SlidePart slidePart, Shape shape, string url)
    {
        var allRuns = shape.Descendants<Drawing.Run>().ToList();
        if (allRuns.Count == 0) return;

        if (string.IsNullOrEmpty(url) || url.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var run in allRuns)
                run.RunProperties?.GetFirstChild<Drawing.HyperlinkOnClick>()?.Remove();
            return;
        }

        var rel = slidePart.AddHyperlinkRelationship(new Uri(url), isExternal: true);
        foreach (var run in allRuns)
        {
            var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
            rProps.RemoveAllChildren<Drawing.HyperlinkOnClick>();
            rProps.InsertAt(new Drawing.HyperlinkOnClick { Id = rel.Id }, 0);
        }
    }

    /// <summary>
    /// Apply a hyperlink URL to a single run. Pass "none" or "" to remove.
    /// </summary>
    private static void ApplyRunHyperlink(SlidePart slidePart, Drawing.Run run, string url)
    {
        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
        rProps.RemoveAllChildren<Drawing.HyperlinkOnClick>();

        if (!string.IsNullOrEmpty(url) && !url.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            var rel = slidePart.AddHyperlinkRelationship(new Uri(url), isExternal: true);
            rProps.InsertAt(new Drawing.HyperlinkOnClick { Id = rel.Id }, 0);
        }
    }

    /// <summary>
    /// Read the hyperlink URL from a run's RunProperties. Returns null if no hyperlink.
    /// </summary>
    private static string? ReadRunHyperlinkUrl(Drawing.Run run, OpenXmlPart part)
    {
        var id = run.RunProperties?.GetFirstChild<Drawing.HyperlinkOnClick>()?.Id?.Value;
        if (id == null) return null;
        try
        {
            var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
            return rel?.Uri?.ToString();
        }
        catch { return null; }
    }
}
