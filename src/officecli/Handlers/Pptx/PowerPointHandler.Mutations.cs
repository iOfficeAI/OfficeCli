// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    public string? Remove(string path)
    {
        var slideMatch = Regex.Match(path, @"^/slide\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (!slideMatch.Success)
            throw new ArgumentException($"Invalid path: {path}. Expected format: /slide[N] or /slide[N]/element[M] (e.g. /slide[1], /slide[1]/shape[2])");

        var slideIdx = int.Parse(slideMatch.Groups[1].Value);

        if (!slideMatch.Groups[2].Success)
        {
            // Remove entire slide
            var presentationPart = _doc.PresentationPart
                ?? throw new InvalidOperationException("Presentation not found");
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = presentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");

            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideIds.Count})");

            var slideId = slideIds[slideIdx - 1];
            var relId = slideId.RelationshipId?.Value;
            slideId.Remove();
            if (relId != null)
                presentationPart.DeletePart(presentationPart.GetPartById(relId));
            presentation.Save();
            return null;
        }

        // Remove element from slide
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shapes");

        var elementType = slideMatch.Groups[2].Value;
        var elementIdx = int.Parse(slideMatch.Groups[3].Value);

        if (elementType == "shape")
        {
            var shapes = shapeTree.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found");
            var shapeToRemove = shapes[elementIdx - 1];
            var shapeId = shapeToRemove.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0;
            if (shapeId > 0)
                RemoveShapeAnimations(GetSlide(slidePart), (uint)shapeId);
            shapeToRemove.Remove();
        }
        else if (elementType is "picture" or "pic" or "video" or "audio")
        {
            List<Picture> pics;
            if (elementType is "video")
                pics = shapeTree.Elements<Picture>()
                    .Where(p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<Drawing.VideoFromFile>() != null).ToList();
            else if (elementType is "audio")
                pics = shapeTree.Elements<Picture>()
                    .Where(p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<Drawing.AudioFromFile>() != null).ToList();
            else
                pics = shapeTree.Elements<Picture>().ToList();

            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {pics.Count})");

            var pic = pics[elementIdx - 1];
            RemovePictureWithCleanup(slidePart, shapeTree, pic);
        }
        else if (elementType == "table")
        {
            var tables = shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > tables.Count)
                throw new ArgumentException($"Table {elementIdx} not found");
            tables[elementIdx - 1].Remove();
        }
        else if (elementType == "chart")
        {
            var charts = shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<C.ChartReference>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > charts.Count)
                throw new ArgumentException($"Chart {elementIdx} not found");
            var chartGf = charts[elementIdx - 1];
            // Clean up ChartPart
            var chartRef = chartGf.Descendants<C.ChartReference>().FirstOrDefault();
            if (chartRef?.Id?.Value != null)
            {
                try { slidePart.DeletePart(chartRef.Id.Value); } catch { }
            }
            chartGf.Remove();
        }
        else if (elementType is "connector" or "connection")
        {
            var connectors = shapeTree.Elements<ConnectionShape>().ToList();
            if (elementIdx < 1 || elementIdx > connectors.Count)
                throw new ArgumentException($"Connector {elementIdx} not found");
            connectors[elementIdx - 1].Remove();
        }
        else if (elementType == "group")
        {
            // Ungroup: move children back to parent shape tree, then remove group
            var groups = shapeTree.Elements<GroupShape>().ToList();
            if (elementIdx < 1 || elementIdx > groups.Count)
                throw new ArgumentException($"Group {elementIdx} not found");
            var group = groups[elementIdx - 1];
            // Recursively clean up any pictures inside the group before ungrouping
            var children = group.ChildElements
                .Where(e => e is Shape or Picture or ConnectionShape or GraphicFrame or GroupShape)
                .ToList();
            foreach (var child in children)
            {
                child.Remove();
                shapeTree.AppendChild(child);
            }
            group.Remove();
        }
        else if (elementType is "zoom" or "slidezoom")
        {
            var zoomElements = GetZoomElements(shapeTree);
            if (elementIdx < 1 || elementIdx > zoomElements.Count)
                throw new ArgumentException($"Zoom {elementIdx} not found (total: {zoomElements.Count})");
            var zmAc = zoomElements[elementIdx - 1];
            // Clean up image relationship if not referenced by other elements
            var zmBlip = zmAc.Descendants().FirstOrDefault(d => d.LocalName == "blip");
            if (zmBlip != null)
            {
                var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var embedAttr = zmBlip.GetAttribute("embed", rNs);
                if (!string.IsNullOrEmpty(embedAttr.Value))
                {
                    var relId = embedAttr.Value;
                    // Check if any other element references this image
                    zmAc.Remove();
                    var slideXml = GetSlide(slidePart).OuterXml;
                    if (!slideXml.Contains(relId))
                    {
                        try { slidePart.DeletePart(relId); } catch { }
                    }
                    GetSlide(slidePart).Save();
                    return null;
                }
            }
            zmAc.Remove();
        }
        else
        {
            throw new ArgumentException($"Unknown element type: {elementType}. Supported: shape, picture, video, audio, table, chart, connector/connection, group, zoom");
        }

        GetSlide(slidePart).Save();
        return null;
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var slideParts = GetSlideParts().ToList();

        // Case 1: Move entire slide (reorder)
        var slideOnlyMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success)
        {
            var slideIdx = int.Parse(slideOnlyMatch.Groups[1].Value);
            var movePresentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = movePresentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");
            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideIds.Count})");

            var slideId = slideIds[slideIdx - 1];
            slideId.Remove();

            if (index.HasValue)
            {
                var remaining = slideIdList.Elements<SlideId>().ToList();
                if (index.Value >= 0 && index.Value < remaining.Count)
                    remaining[index.Value].InsertBeforeSelf(slideId);
                else
                    slideIdList.AppendChild(slideId);
            }
            else
            {
                slideIdList.AppendChild(slideId);
            }

            movePresentation.Save();
            var newSlideIds = slideIdList.Elements<SlideId>().ToList();
            var newIdx = newSlideIds.IndexOf(slideId) + 1;
            return $"/slide[{newIdx}]";
        }

        // Case 2: Move element within/across slides
        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);

        // Determine target
        string effectiveParentPath;
        SlidePart tgtSlidePart;
        ShapeTree tgtShapeTree;

        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within same parent
            tgtSlidePart = srcSlidePart;
            tgtShapeTree = GetSlide(srcSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
            var srcSlideIdx = slideParts.IndexOf(srcSlidePart) + 1;
            effectiveParentPath = $"/slide[{srcSlideIdx}]";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
            if (!tgtSlideMatch.Success)
                throw new ArgumentException($"Target must be a slide: /slide[N]");
            var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
            if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {tgtSlideIdx} not found (total: {slideParts.Count})");
            tgtSlidePart = slideParts[tgtSlideIdx - 1];
            tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
        }

        // Copy relationships BEFORE removing from source (so rel IDs are still accessible)
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(srcElement, srcSlidePart, tgtSlidePart);

        srcElement.Remove();

        InsertAtPosition(tgtShapeTree, srcElement, index);

        GetSlide(srcSlidePart).Save();
        if (srcSlidePart != tgtSlidePart)
            GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(effectiveParentPath, srcElement, tgtShapeTree);
    }

    public (string NewPath1, string NewPath2) Swap(string path1, string path2)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var slideParts = GetSlideParts().ToList();

        // Case 1: Swap two slides
        var slide1Match = Regex.Match(path1, @"^/slide\[(\d+)\]$");
        var slide2Match = Regex.Match(path2, @"^/slide\[(\d+)\]$");
        if (slide1Match.Success && slide2Match.Success)
        {
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = presentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");
            var slideIds = slideIdList.Elements<SlideId>().ToList();
            var idx1 = int.Parse(slide1Match.Groups[1].Value);
            var idx2 = int.Parse(slide2Match.Groups[1].Value);
            if (idx1 < 1 || idx1 > slideIds.Count) throw new ArgumentException($"Slide {idx1} not found (total: {slideIds.Count})");
            if (idx2 < 1 || idx2 > slideIds.Count) throw new ArgumentException($"Slide {idx2} not found (total: {slideIds.Count})");
            if (idx1 == idx2) return (path1, path2);

            SwapXmlElements(slideIds[idx1 - 1], slideIds[idx2 - 1]);
            presentation.Save();
            return ($"/slide[{idx2}]", $"/slide[{idx1}]");
        }

        // Case 2: Swap two elements within the same slide
        var (slide1Part, elem1) = ResolveSlideElement(path1, slideParts);
        var (slide2Part, elem2) = ResolveSlideElement(path2, slideParts);
        if (slide1Part != slide2Part)
            throw new ArgumentException("Cannot swap elements on different slides");

        SwapXmlElements(elem1, elem2);
        GetSlide(slide1Part).Save();

        var slideIdx = slideParts.IndexOf(slide1Part) + 1;
        var parentPath = $"/slide[{slideIdx}]";
        var shapeTree = GetSlide(slide1Part).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");
        var newPath1 = ComputeElementPath(parentPath, elem1, shapeTree);
        var newPath2 = ComputeElementPath(parentPath, elem2, shapeTree);
        return (newPath1, newPath2);
    }

    internal static void SwapXmlElements(OpenXmlElement a, OpenXmlElement b)
    {
        if (a == b || a.Parent == null || b.Parent == null) return;
        var parent = a.Parent;
        var aNext = a.NextSibling();
        var bNext = b.NextSibling();

        a.Remove();
        b.Remove();

        if (aNext == b)
        {
            // A was directly before B: [... A B ...] → [... B A ...]
            if (bNext != null)
                bNext.InsertBeforeSelf(b);
            else
                parent.AppendChild(b);
            b.InsertAfterSelf(a);
        }
        else if (bNext == a)
        {
            // B was directly before A: [... B A ...] → [... A B ...]
            if (aNext != null)
                aNext.InsertBeforeSelf(a);
            else
                parent.AppendChild(a);
            a.InsertBeforeSelf(b);
        }
        else
        {
            // Non-adjacent: insert each where the other was
            if (aNext != null)
                aNext.InsertBeforeSelf(b);
            else
                parent.AppendChild(b);
            if (bNext != null)
                bNext.InsertBeforeSelf(a);
            else
                parent.AppendChild(a);
        }
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var slideParts = GetSlideParts().ToList();

        // Whole-slide clone: --from /slide[N] to /
        var slideCloneMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]$");
        if (slideCloneMatch.Success && (targetParentPath is "/" or "" or "/presentation"))
        {
            return CloneSlide(slideCloneMatch, slideParts, index);
        }

        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);
        var clone = srcElement.CloneNode(true);

        var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
        if (!tgtSlideMatch.Success)
            throw new ArgumentException($"Target must be a slide: /slide[N]");
        var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
        if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {tgtSlideIdx} not found (total: {slideParts.Count})");

        var tgtSlidePart = slideParts[tgtSlideIdx - 1];
        var tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        // Copy relationships if across slides
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(clone, srcSlidePart, tgtSlidePart);

        InsertAtPosition(tgtShapeTree, clone, index);
        GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(targetParentPath, clone, tgtShapeTree);
    }

    /// <summary>
    /// Clone an entire slide with all its content, relationships (images, charts, media),
    /// layout link, background, notes, and transitions.
    /// Pattern follows POI's createSlide(layout) + importContent(srcSlide).
    /// </summary>
    private string CloneSlide(Match slideMatch, List<SlidePart> slideParts, int? index)
    {
        var srcSlideIdx = int.Parse(slideMatch.Groups[1].Value);
        if (srcSlideIdx < 1 || srcSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {srcSlideIdx} not found (total: {slideParts.Count})");

        var srcSlidePart = slideParts[srcSlideIdx - 1];
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var presentation = presentationPart.Presentation
            ?? throw new InvalidOperationException("No presentation");

        // 1. Create new SlidePart
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // 2. Copy slide layout relationship (link to same layout as source)
        var srcLayoutPart = srcSlidePart.SlideLayoutPart;
        if (srcLayoutPart != null)
            newSlidePart.AddPart(srcLayoutPart);

        // 3. Deep-clone the Slide XML
        var srcSlide = GetSlide(srcSlidePart);
        newSlidePart.Slide = (Slide)srcSlide.CloneNode(true);

        // 4. Copy all referenced parts (images, charts, embedded objects, media)
        CopySlideParts(srcSlidePart, newSlidePart);

        // 5. Copy notes slide if present
        if (srcSlidePart.NotesSlidePart != null)
        {
            var srcNotesPart = srcSlidePart.NotesSlidePart;
            var newNotesPart = newSlidePart.AddNewPart<NotesSlidePart>();
            newNotesPart.NotesSlide = srcNotesPart.NotesSlide != null
                ? (NotesSlide)srcNotesPart.NotesSlide.CloneNode(true)
                : new NotesSlide();
            // Link notes to the new slide
            newNotesPart.AddPart(newSlidePart);
        }

        newSlidePart.Slide.Save();

        // 6. Register in SlideIdList at the correct position
        var slideIdList = presentation.GetFirstChild<SlideIdList>()
            ?? presentation.AppendChild(new SlideIdList());
        var maxId = slideIdList.Elements<SlideId>().Any()
            ? slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255) + 1
            : 256;
        var relId = presentationPart.GetIdOfPart(newSlidePart);
        var newSlideId = new SlideId { Id = maxId, RelationshipId = relId };

        if (index.HasValue && index.Value < slideIdList.Elements<SlideId>().Count())
        {
            var refSlide = slideIdList.Elements<SlideId>().ElementAtOrDefault(index.Value);
            if (refSlide != null)
                slideIdList.InsertBefore(newSlideId, refSlide);
            else
                slideIdList.AppendChild(newSlideId);
        }
        else
        {
            slideIdList.AppendChild(newSlideId);
        }

        presentation.Save();

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        var insertedIdx = slideIds.FindIndex(s => s.RelationshipId?.Value == relId) + 1;
        return $"/slide[{insertedIdx}]";
    }

    /// <summary>
    /// Copy all sub-parts (images, charts, media, etc.) from source to target slide,
    /// remapping relationship IDs in the cloned XML.
    /// </summary>
    private static void CopySlideParts(SlidePart source, SlidePart target)
    {
        // Build a map of old rId → new rId for all parts that need copying
        var rIdMap = new Dictionary<string, string>();

        foreach (var part in source.Parts)
        {
            // Skip SlideLayoutPart (already linked above)
            if (part.OpenXmlPart is SlideLayoutPart) continue;
            // Skip NotesSlidePart (handled separately)
            if (part.OpenXmlPart is NotesSlidePart) continue;

            try
            {
                // Try to add the same part (shares the underlying data)
                var newRelId = target.CreateRelationshipToPart(part.OpenXmlPart);
                if (newRelId != part.RelationshipId)
                    rIdMap[part.RelationshipId] = newRelId;
            }
            catch
            {
                // If sharing fails, deep-copy the part data
                try
                {
                    var newPart = target.AddNewPart<OpenXmlPart>(part.OpenXmlPart.ContentType, part.RelationshipId);
                    using var stream = part.OpenXmlPart.GetStream();
                    newPart.FeedData(stream);
                }
                catch { /* Best effort — some parts may not be copyable */ }
            }
        }

        // Also copy external relationships (hyperlinks, media links)
        foreach (var extRel in source.ExternalRelationships)
        {
            try
            {
                target.AddExternalRelationship(extRel.RelationshipType, extRel.Uri, extRel.Id);
            }
            catch { }
        }
        foreach (var hyperRel in source.HyperlinkRelationships)
        {
            try
            {
                target.AddHyperlinkRelationship(hyperRel.Uri, hyperRel.IsExternal, hyperRel.Id);
            }
            catch { }
        }

        // Remap any changed relationship IDs in the slide XML
        if (rIdMap.Count > 0 && target.Slide != null)
        {
            RemapRelationshipIds(target.Slide, rIdMap);
            target.Slide.Save();
        }
    }

    /// <summary>
    /// Update all r:id references in the XML tree when relationship IDs changed during copy.
    /// </summary>
    private static void RemapRelationshipIds(OpenXmlElement root, Dictionary<string, string> rIdMap)
    {
        var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        foreach (var el in root.Descendants().Prepend(root).ToList())
        {
            foreach (var attr in el.GetAttributes().ToList())
            {
                if (attr.NamespaceUri != rNsUri || attr.Value == null) continue;
                if (rIdMap.TryGetValue(attr.Value, out var newId))
                {
                    el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newId));
                }
            }
        }
    }

    private (SlidePart slidePart, OpenXmlElement element) ResolveSlideElement(string path, List<SlidePart> slideParts)
    {
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\]$");
        if (!match.Success)
            throw new ArgumentException($"Invalid element path: {path}. Expected /slide[N]/element[M]");

        var slideIdx = int.Parse(match.Groups[1].Value);
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);

        OpenXmlElement element = elementType switch
        {
            "shape" => shapeTree.Elements<Shape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Shape {elementIdx} not found"),
            "picture" or "pic" => shapeTree.Elements<Picture>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Picture {elementIdx} not found"),
            "connector" or "connection" => shapeTree.Elements<ConnectionShape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Connector {elementIdx} not found"),
            "table" => shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Table {elementIdx} not found"),
            "chart" => shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<C.ChartReference>().Any()).ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Chart {elementIdx} not found"),
            "group" => shapeTree.Elements<GroupShape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Group {elementIdx} not found"),
            _ => shapeTree.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase))
                .ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"{elementType} {elementIdx} not found")
        };

        return (slidePart, element);
    }

    private static void CopyRelationships(OpenXmlElement element, SlidePart sourcePart, SlidePart targetPart)
    {
        var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var allElements = element.Descendants().Prepend(element);

        foreach (var el in allElements.ToList())
        {
            foreach (var attr in el.GetAttributes().ToList())
            {
                if (attr.NamespaceUri != rNsUri) continue;

                var oldRelId = attr.Value;
                if (string.IsNullOrEmpty(oldRelId)) continue;

                try
                {
                    var referencedPart = sourcePart.GetPartById(oldRelId);
                    string newRelId;
                    try
                    {
                        newRelId = targetPart.GetIdOfPart(referencedPart);
                    }
                    catch (ArgumentException)
                    {
                        newRelId = targetPart.CreateRelationshipToPart(referencedPart);
                    }

                    if (newRelId != oldRelId)
                    {
                        el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newRelId));
                    }
                }
                catch (ArgumentOutOfRangeException) { /* Not a valid relationship ID, skip */ }
            }
        }
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue && parent is ShapeTree)
        {
            // Skip structural elements (nvGrpSpPr, grpSpPr) that must stay at the beginning
            var contentChildren = parent.ChildElements
                .Where(e => e is not NonVisualGroupShapeProperties && e is not GroupShapeProperties)
                .ToList();
            if (index.Value >= 0 && index.Value < contentChildren.Count)
                contentChildren[index.Value].InsertBeforeSelf(element);
            else if (contentChildren.Count > 0)
                contentChildren.Last().InsertAfterSelf(element);
            else
                parent.AppendChild(element);
        }
        else if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                parent.AppendChild(element);
        }
        else
        {
            parent.AppendChild(element);
        }
    }

    private static string ComputeElementPath(string parentPath, OpenXmlElement element, ShapeTree shapeTree)
    {
        // Map back to semantic type names
        string typeName;
        int typeIdx;
        if (element is Shape)
        {
            typeName = "shape";
            typeIdx = shapeTree.Elements<Shape>().ToList().IndexOf((Shape)element) + 1;
        }
        else if (element is Picture)
        {
            typeName = "picture";
            typeIdx = shapeTree.Elements<Picture>().ToList().IndexOf((Picture)element) + 1;
        }
        else if (element is ConnectionShape)
        {
            typeName = "connector";
            typeIdx = shapeTree.Elements<ConnectionShape>().ToList().IndexOf((ConnectionShape)element) + 1;
        }
        else if (element is GroupShape)
        {
            typeName = "group";
            typeIdx = shapeTree.Elements<GroupShape>().ToList().IndexOf((GroupShape)element) + 1;
        }
        else if (element is GraphicFrame gf)
        {
            if (gf.Descendants<Drawing.Table>().Any())
            {
                typeName = "table";
                typeIdx = shapeTree.Elements<GraphicFrame>()
                    .Where(f => f.Descendants<Drawing.Table>().Any())
                    .ToList().IndexOf(gf) + 1;
            }
            else if (gf.Descendants<C.ChartReference>().Any())
            {
                typeName = "chart";
                typeIdx = shapeTree.Elements<GraphicFrame>()
                    .Where(f => f.Descendants<C.ChartReference>().Any())
                    .ToList().IndexOf(gf) + 1;
            }
            else
            {
                typeName = element.LocalName;
                typeIdx = shapeTree.ChildElements
                    .Where(e => e.LocalName == element.LocalName)
                    .ToList().IndexOf(element) + 1;
            }
        }
        else
        {
            typeName = element.LocalName;
            typeIdx = shapeTree.ChildElements
                .Where(e => e.LocalName == element.LocalName)
                .ToList().IndexOf(element) + 1;
        }
        return $"{parentPath}/{typeName}[{typeIdx}]";
    }
}
