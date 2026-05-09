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
        var presentationPart = _doc.PresentationPart;
        if (presentationPart == null) return "(empty)";

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"because it ends in .xml, but no matching part exists in the package. " +
                    $"Use semantic paths (/presentation, /slide[N], /slideMaster[N]) for stable identification.");
            return RawXmlHelper.ReadPartXml(part);
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

        throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /theme, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N]");
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"because it ends in .xml, but no matching part exists in the package. " +
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
            if (idx < 1 || idx > layouts.Count)
                throw new ArgumentException($"SlideLayout {idx} not found (total: {layouts.Count})");
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
        else
        {
            throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /theme, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N]");
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
