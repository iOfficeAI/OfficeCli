// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler : IDocumentHandler
{
    private readonly PresentationDocument _doc;
    private readonly string _filePath;
    private HashSet<uint> _usedShapeIds = new();
    private uint _nextShapeId = 10000;
    public int LastFindMatchCount { get; internal set; }

    public PowerPointHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _doc = PresentationDocument.Open(filePath, editable);
        if (editable)
            InitShapeIdCounter();
    }

    /// <summary>
    /// Get the slide dimensions from the presentation. Falls back to 16:9 (33.867cm × 19.05cm).
    /// </summary>
    private (long width, long height) GetSlideSize()
    {
        var sldSz = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>();
        return (sldSz?.Cx?.Value ?? SlideSizeDefaults.Widescreen16x9Cx, sldSz?.Cy?.Value ?? SlideSizeDefaults.Widescreen16x9Cy);
    }

    // ==================== Raw Layer ====================

    // CONSISTENCY(zip-uri-lookup): see ExcelHandler.cs / RawXmlHelper —
    // any partPath ending in `.xml` is resolved as a literal zip URI via
    // the package's part tree, no per-handler alias table needed.

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var presentationPart = _doc.PresentationPart;
        if (presentationPart == null) return "(empty)";

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var xml = RawXmlHelper.TryReadByZipUri(_doc, _filePath, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/presentation, /slide[N], /slideMaster[N]) for stable identification.");
            return xml;
        }

        if (partPath == "/" || partPath == "/presentation")
            return presentationPart.Presentation?.OuterXml ?? "(empty)";

        if (partPath == "/theme")
            return presentationPart.ThemePart?.Theme?.OuterXml ?? "(no theme)";

        var slideMatch = Regex.Match(partPath, @"^/slide\[(\d+)\]$");
        if (slideMatch.Success)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx >= 1 && idx <= slideParts.Count)
                return GetSlide(slideParts[idx - 1]).OuterXml;
            throw new ArgumentException($"slide[{idx}] not found (total: {slideParts.Count})");
        }

        // CONSISTENCY(raw-rawset-symmetry): RawSet supports master/layout/noteSlide;
        // Raw must too, otherwise users can't read back what they just wrote.
        var masterMatch = Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$");
        if (masterMatch.Success)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1 || idx > masters.Count)
                throw new ArgumentException($"slideMaster[{idx}] not found (total: {masters.Count})");
            return masters[idx - 1].SlideMaster?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }

        var layoutMatch = Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$");
        if (layoutMatch.Success)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1 || idx > layouts.Count)
                throw new ArgumentException($"slideLayout[{idx}] not found (total: {layouts.Count})");
            return layouts[idx - 1].SlideLayout?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }

        var noteMatch = Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$");
        if (noteMatch.Success)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"slide[{idx}] not found (total: {slideParts.Count})");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            return notesPart.NotesSlide?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }

        // CONSISTENCY(raw-rawset-symmetry): /notesMaster surfaces the
        // presentation-level NotesMasterPart's XML so PptxBatchEmitter can
        // raw-set it on replay (mirrors theme/master/layout treatment).
        if (partPath == "/notesMaster")
        {
            return presentationPart.NotesMasterPart?.NotesMaster?.OuterXml
                ?? throw new ArgumentException("No notes master part");
        }

        throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /theme, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N], /notesMaster");
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/presentation, /slide[N], /slideMaster[N]) for stable identification.");
            RawXmlHelper.Execute(part, xpath, action, xml);
            return;
        }

        OpenXmlPartRootElement rootElement;

        if (partPath is "/" or "/presentation")
        {
            rootElement = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
        }
        else if (partPath == "/theme")
        {
            rootElement = presentationPart.ThemePart?.Theme
                ?? throw new ArgumentException("No theme part");
        }
        else if (Regex.Match(partPath, @"^/slide\[(\d+)\]$") is { Success: true } slideMatch)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found (total: {slideParts.Count})");
            rootElement = GetSlide(slideParts[idx - 1]);
        }
        else if (Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$") is { Success: true } masterMatch)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1 || idx > masters.Count)
                throw new ArgumentException($"SlideMaster {idx} not found (total: {masters.Count})");
            rootElement = masters[idx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }
        else if (Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$") is { Success: true } layoutMatch)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1)
                throw new ArgumentException($"SlideLayout {idx} not found (total: {layouts.Count})");
            if (idx > layouts.Count)
            {
                // BUG-J: Replay scenario — source deck has N layouts (e.g. 11) but
                // the blank target only stamped K (5). EmitMasterRaw already
                // replaced the master's sldLayoutIdLst to reference all N rIds,
                // but the SlideLayoutPart objects for K+1..N don't exist yet, so
                // raw-set fails with "SlideLayout {idx} not found". Auto-grow the
                // missing layout parts under the appropriate master based on the
                // post-master-replace sldLayoutIdLst, then re-resolve.
                GrowSlideLayoutParts(idx);
                layouts = presentationPart.SlideMasterParts
                    .SelectMany(m => m.SlideLayoutParts).ToList();
                if (idx > layouts.Count)
                    throw new ArgumentException($"SlideLayout {idx} not found (total: {layouts.Count})");
            }
            rootElement = layouts[idx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }
        else if (Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$") is { Success: true } noteMatch)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found (total: {slideParts.Count})");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            rootElement = notesPart.NotesSlide
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }
        else if (partPath == "/notesMaster")
        {
            rootElement = presentationPart.NotesMasterPart?.NotesMaster
                ?? throw new ArgumentException("No notes master part");
        }
        else
        {
            throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /theme, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N], /notesMaster");
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        // BUG-R43: raw-set may have inserted/removed shape XML directly (incl.
        // cNvPr ids). The cached _usedShapeIds set is now stale, so the next
        // Add() can hand out an id that already exists in the tree, producing
        // duplicate cNvPr ids that PowerPoint silently rejects. Rebuild the
        // shape-id index from the live tree after every raw-set.
        InitShapeIdCounter();
        // BUG-R5-01: silent — CLI wrappers print their own structured message.
        _ = affected;
    }

    // BUG-J: Auto-grow SlideLayoutPart objects so raw-set replay can target
    // /slideLayout[targetGlobalIdx] even when the blank deck only stamped a
    // subset. The master's sldLayoutIdLst (already replaced by EmitMasterRaw)
    // is the source of truth: each entry holds the rId that the new
    // SlideLayoutPart must register with so the master-layout relationship
    // matches. We walk masters in declaration order, compute each master's
    // declared layout count from sldLayoutIdLst, and create missing parts
    // under whichever master's range contains targetGlobalIdx. Newly created
    // parts get a stub root (the imminent replace overwrites it).
    private void GrowSlideLayoutParts(int targetGlobalIdx)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");
        var masters = presentationPart.SlideMasterParts.ToList();
        int seen = 0;
        foreach (var mp in masters)
        {
            var declared = mp.SlideMaster?.SlideLayoutIdList?.Elements<SlideLayoutId>().ToList()
                ?? new List<SlideLayoutId>();
            int declaredCount = declared.Count;
            int existingCount = mp.SlideLayoutParts.Count();
            // This master "owns" global indices (seen+1)..(seen+declaredCount).
            int rangeStart = seen + 1;
            int rangeEnd = seen + declaredCount;
            if (targetGlobalIdx >= rangeStart && targetGlobalIdx <= rangeEnd)
            {
                // Create missing parts for slots existingCount+1 .. declaredCount.
                for (int slot = existingCount; slot < declaredCount; slot++)
                {
                    var declaredId = declared[slot];
                    var rId = declaredId.RelationshipId?.Value;
                    SlideLayoutPart newPart;
                    if (!string.IsNullOrEmpty(rId) && !mp.Parts.Any(p => p.RelationshipId == rId))
                    {
                        newPart = mp.AddNewPart<SlideLayoutPart>(rId);
                    }
                    else
                    {
                        // Either rId missing in sldLayoutIdLst or already taken
                        // (corruption guard) — let OpenXml allocate a new one
                        // and patch the sldLayoutIdLst entry to match.
                        newPart = mp.AddNewPart<SlideLayoutPart>();
                        var newRid = mp.GetIdOfPart(newPart);
                        declaredId.RelationshipId = newRid;
                    }
                    // Stub root — the raw-set replace immediately rewrites it.
                    newPart.SlideLayout = new SlideLayout(
                        new CommonSlideData(
                            new ShapeTree(
                                new NonVisualGroupShapeProperties(
                                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                                    new NonVisualGroupShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties()),
                                new GroupShapeProperties()))
                    ) { Type = SlideLayoutValues.Blank };
                    newPart.SlideLayout.Save();
                    // Layouts must point back to their master.
                    newPart.AddPart(mp);
                }
                if (mp.SlideMaster != null) mp.SlideMaster.Save();
                return;
            }
            seen += declaredCount;
        }
        // targetGlobalIdx is beyond every master's declared range; caller will
        // raise the canonical "not found" error after we return.
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a SlidePart
                var slideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException(
                        "Chart must be added under a slide: add-part <file> '/slide[N]' --type chart");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide index {slideIdx} out of range");

                var slidePart = slideParts[slideIdx - 1];
                var chartPart = slidePart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
                var relId = slidePart.GetIdOfPart(chartPart);

                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = slidePart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/slide[{slideIdx}]/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose() => _doc.Dispose();

    // Internal accessors used by PptxBatchEmitter (resource enumeration).
    // Keep the PresentationPart itself private; expose only the counts and
    // a binary getter that the emitter needs.
    internal int SlideMasterCount =>
        _doc.PresentationPart?.SlideMasterParts.Count() ?? 0;
    internal int SlideLayoutCount =>
        _doc.PresentationPart?.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).Count() ?? 0;
    internal bool HasNotesMaster =>
        _doc.PresentationPart?.NotesMasterPart != null;
    internal bool SlideHasNotes(int slideIdx)
    {
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return false;
        return parts[slideIdx - 1].NotesSlidePart != null;
    }

    // Resolve a /slide[N]/picture[M] path's image bytes for base64-inline emit.
    // Mirrors WordHandler.GetImageBinary's contract: returns null if the path
    // does not resolve to a Picture with an embedded ImagePart.
    public (byte[] Bytes, string ContentType)? GetImageBinary(string picturePath)
    {
        // Accept both `picture[N]` positional and `picture[@id=N]` cNvPr-id
        // segment forms (BuildElementPathSegment emits @id= when the shape
        // carries a cNvPr id, which Pictures always do).
        var m = Regex.Match(picturePath,
            @"^/slide\[(\d+)\]/(?:.+/)?picture\[(?:@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var idOrIdx = int.Parse(m.Groups[2].Value);
        var byId = picturePath.Contains("@id=", StringComparison.Ordinal);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return null;
        var pictures = shapeTree.Descendants<Picture>().ToList();
        Picture? pic = null;
        if (byId)
        {
            pic = pictures.FirstOrDefault(p =>
            {
                var pid = p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
                return pid.HasValue && pid.Value == (uint)idOrIdx;
            });
        }
        else
        {
            if (idOrIdx >= 1 && idOrIdx <= pictures.Count) pic = pictures[idOrIdx - 1];
        }
        if (pic == null) return null;
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        var embedId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return null;
        try
        {
            var part = slidePart.GetPartById(embedId);
            using var src = part.GetStream();
            using var ms = new MemoryStream();
            src.CopyTo(ms);
            return (ms.ToArray(), part.ContentType);
        }
        catch { return null; }
    }

    // ==================== Private Helpers ====================

    private static Slide GetSlide(SlidePart part) =>
        part.Slide ?? throw new InvalidOperationException("Corrupt file: slide data missing");

    private IEnumerable<SlidePart> GetSlideParts()
    {
        var presentation = _doc.PresentationPart?.Presentation;
        var slideIdList = presentation?.GetFirstChild<SlideIdList>();
        if (slideIdList == null) yield break;

        foreach (var slideId in slideIdList.Elements<SlideId>())
        {
            var relId = slideId.RelationshipId?.Value;
            if (relId == null) continue;
            yield return (SlidePart)_doc.PresentationPart!.GetPartById(relId);
        }
    }

}
