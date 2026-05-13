// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public string? Remove(string path)
    {
        // CONSISTENCY(container-remove-guard): reject removal of required
        // structural container elements up front. Without this guard,
        // `remove /body` / `remove /styles` etc. fall through to
        // NavigateToElement + element.Remove() and permanently corrupt
        // the document (body cleared, styles/numbering NRE). AI agents
        // mis-dispatching a remove command should never be able to nuke
        // the file.
        if (IsProtectedContainerPath(path))
            throw new ArgumentException(
                $"Cannot remove container element '{path}': it is a required structural element of the document.");

        // CONSISTENCY(container-remove-guard): the last <w:sectPr> inside
        // <w:body> is required by the OOXML schema — removing it corrupts the
        // document so that Word refuses to open it on next launch. Matches
        // `/body/sectPr` and the indexed form `/body/sectPr[N]`.
        if (IsProtectedSectPrPath(path))
            throw new ArgumentException(
                "Cannot remove '/body/sectPr': it is required by the Word document body (last section properties). Removing it corrupts the document.");

        // Handle /watermark removal
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase))
        {
            RemoveWatermarkHeaders();
            _doc.MainDocumentPart?.Document?.Save();
            return null;
        }

        // /styles/<id> removal is idempotent: missing-style is a soft
        // success. Dump→batch emits a "clear blank's auto-stamped Normal
        // style" preamble so dump∘replay∘dump is a fixed point even when
        // the source doc lacks a Normal style (blank's leftover would
        // otherwise show up on dump-2). On replay of that emit, when
        // the target already lacks Normal (e.g. dump-2's replay onto a
        // future bare target), Path-not-found would abort an entire
        // batch — even though the post-state ("Normal absent") matches
        // intent. Narrow scope to /styles/<id> only; generic remove
        // (e.g. /body/p[N] out-of-range) keeps strict semantics.
        var styleRemoveMatch = Regex.Match(path, @"^/styles/([^/]+)/?$", RegexOptions.IgnoreCase);
        if (styleRemoveMatch.Success)
        {
            var styleKey = styleRemoveMatch.Groups[1].Value;
            var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            var target = stylesPart?.Elements<Style>().FirstOrDefault(s =>
                string.Equals(s.StyleId?.Value, styleKey, StringComparison.Ordinal)
                || string.Equals(s.StyleName?.Val?.Value, styleKey, StringComparison.Ordinal));
            if (target == null) return null; // soft success on missing
            target.Remove();
            stylesPart!.Save();
            return null;
        }

        // BUG-R10-03: support /header[N]/ole[M] and /footer[N]/ole[M] shorthand
        // in Remove, mirroring the Get shorthand added in Round 9. Users
        // cannot easily discover the underlying run path, so without this
        // intercept the shorthand path crashed with "Path not found" on
        // Remove even though Get accepted it.
        // CONSISTENCY(ole-shorthand-remove): also handle /body/ole[N] — the body
        // OLE actual path is /body/p[N]/r[M], not /body/ole[N], so the normal
        // path parser hits "No ole found at /body" just like header/footer.
        // The root-level /ole[N] shorthand (added in BUG-R11-03 for Get) is
        // handled by the regex below which allows an absent <parent> group.
        var wordOleShortMatch = Regex.Match(
            path,
            @"^(?<parent>/body|/header\[\d+\]|/footer\[\d+\])?/(?:ole|object|embed)\[(?<idx>\d+)\]$",
            RegexOptions.IgnoreCase);
        if (wordOleShortMatch.Success)
        {
            var wOleIdx = int.Parse(wordOleShortMatch.Groups["idx"].Value);
            var wOleParent = wordOleShortMatch.Groups["parent"].Success && wordOleShortMatch.Groups["parent"].Value.Length > 0
                ? wordOleShortMatch.Groups["parent"].Value
                : "/body";
            var allOles = Query("ole")
                .Where(n => n.Path.StartsWith(wOleParent + "/", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (wOleIdx < 1 || wOleIdx > allOles.Count)
                throw new ArgumentException(
                    $"OLE object {wOleIdx} not found at {wOleParent} (available: {allOles.Count}).");
            // Recurse into Remove with the resolved run path (e.g.
            // /body/p[1]/r[1] or /header[1]/p[1]/r[1]) so the normal
            // run/OLE cleanup runs on the correct part.
            return Remove(allOles[wOleIdx - 1].Path);
        }

        // Virtual table column path — strip gridCol + per-row tc.
        var colRemoveMatch = Regex.Match(path, @"^/body/tbl\[(\d+)\]/col\[(\d+)\]$");
        if (colRemoveMatch.Success)
        {
            RemoveTableColumn(colRemoveMatch);
            _doc.MainDocumentPart?.Document?.Save();
            return null;
        }

        var parts = ParsePath(path);

        // Handle header/footer removal by deleting the part itself
        if (parts.Count == 1 && parts[0].Name.ToLowerInvariant() is "header" or "footer")
        {
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var idx = (parts[0].Index ?? 1) - 1;
            var isHeader = parts[0].Name.ToLowerInvariant() == "header";

            // Track removed ref types so we can mirror the add-time settings/sectPr
            // writes performed by AddHeader/AddFooter (round23 A) and TitlePage write
            // for type=first. Without this, settings.xml keeps a stale
            // <w:evenAndOddHeaders/> and the sectPr keeps a stale <w:titlePg/>.
            bool removedAnyEven = false;
            var sectPrsWithFirstRemoved = new List<SectionProperties>();

            if (isHeader)
            {
                var headerPart = mainPart.HeaderParts.ElementAtOrDefault(idx)
                    ?? throw new ArgumentException($"Path not found: {path}");
                // Remove header references from section properties
                var partId = mainPart.GetIdOfPart(headerPart);
                foreach (var sectProps in mainPart.Document?.Body?.Descendants<SectionProperties>() ?? Enumerable.Empty<SectionProperties>())
                {
                    var refs = sectProps.Elements<HeaderReference>().Where(r => r.Id?.Value == partId).ToList();
                    foreach (var r in refs)
                    {
                        if (r.Type?.Value == HeaderFooterValues.Even) removedAnyEven = true;
                        if (r.Type?.Value == HeaderFooterValues.First) sectPrsWithFirstRemoved.Add(sectProps);
                        r.Remove();
                    }
                }
                // Clean up ImageParts referenced only by this header
                CleanupImageParts(mainPart, headerPart.Header?.Descendants<A.Blip>(), headerPart);
                mainPart.DeletePart(headerPart);
            }
            else
            {
                var footerPart = mainPart.FooterParts.ElementAtOrDefault(idx)
                    ?? throw new ArgumentException($"Path not found: {path}");
                var partId = mainPart.GetIdOfPart(footerPart);
                foreach (var sectProps in mainPart.Document?.Body?.Descendants<SectionProperties>() ?? Enumerable.Empty<SectionProperties>())
                {
                    var refs = sectProps.Elements<FooterReference>().Where(r => r.Id?.Value == partId).ToList();
                    foreach (var r in refs)
                    {
                        if (r.Type?.Value == HeaderFooterValues.Even) removedAnyEven = true;
                        if (r.Type?.Value == HeaderFooterValues.First) sectPrsWithFirstRemoved.Add(sectProps);
                        r.Remove();
                    }
                }
                // Clean up ImageParts referenced only by this footer
                CleanupImageParts(mainPart, footerPart.Footer?.Descendants<A.Blip>(), footerPart);
                mainPart.DeletePart(footerPart);
            }

            // Doc-level: when the last even-typed Header/FooterReference goes away,
            // the doc-level <w:evenAndOddHeaders/> in settings.xml must go too.
            // Scan every remaining sectPr (header AND footer refs) so that an even
            // header removal triggered while an even footer still exists keeps it.
            if (removedAnyEven)
            {
                bool anyEvenLeft = (mainPart.Document?.Body?.Descendants<SectionProperties>() ?? Enumerable.Empty<SectionProperties>())
                    .Any(sp => sp.Elements<HeaderReference>().Any(r => r.Type?.Value == HeaderFooterValues.Even)
                            || sp.Elements<FooterReference>().Any(r => r.Type?.Value == HeaderFooterValues.Even));
                if (!anyEvenLeft)
                {
                    var settingsPart = mainPart.DocumentSettingsPart;
                    if (settingsPart?.Settings != null)
                    {
                        settingsPart.Settings.RemoveAllChildren<EvenAndOddHeaders>();
                        settingsPart.Settings.Save();
                    }
                }
            }

            // Per-sectPr: <w:titlePg/> only matters when at least one first-typed
            // Header or Footer is still attached. Once the last first-typed ref on
            // a given sectPr is gone, strip TitlePage from that sectPr alone —
            // sibling sectPrs that still carry a first ref keep theirs.
            foreach (var sp in sectPrsWithFirstRemoved.Distinct())
            {
                bool firstRefStill = sp.Elements<HeaderReference>().Any(r => r.Type?.Value == HeaderFooterValues.First)
                                  || sp.Elements<FooterReference>().Any(r => r.Type?.Value == HeaderFooterValues.First);
                if (!firstRefStill)
                    sp.RemoveAllChildren<TitlePage>();
            }

            mainPart.Document?.Save();
            return null;
        }

        // Handle TOC removal
        if (parts.Count == 1 && parts[0].Name.ToLowerInvariant() == "toc")
        {
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var tocIdx = parts[0].Index ?? 1;
            var tocParas = FindTocParagraphs();
            if (tocIdx < 1 || tocIdx > tocParas.Count)
                throw new ArgumentException($"TOC {tocIdx} not found (total: {tocParas.Count})");

            var tocPara = tocParas[tocIdx - 1];

            // Also remove preceding TOCHeading title paragraph if present
            var prevSibling = tocPara.PreviousSibling<Paragraph>();
            if (prevSibling != null)
            {
                var styleId = prevSibling.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && styleId.Equals("TOCHeading", StringComparison.OrdinalIgnoreCase))
                    prevSibling.Remove();
            }

            tocPara.Remove();
            mainPart.Document?.Save();
            return null;
        }

        // Handle footnote/endnote removal
        if (parts.Count == 1 && parts[0].Name.ToLowerInvariant() == "footnote")
        {
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var fnId = parts[0].Index ?? 1;
            var fn = mainPart.FootnotesPart?.Footnotes?
                .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId)
                ?? throw new ArgumentException($"Path not found: {path}");
            // Remove footnote reference from body
            foreach (var fnRef in mainPart.Document!.Descendants<FootnoteReference>()
                .Where(r => r.Id?.Value == fnId).ToList())
                fnRef.Parent?.Remove();
            fn.Remove();
            mainPart.FootnotesPart?.Footnotes?.Save();
            mainPart.Document?.Save();
            return null;
        }
        if (parts.Count == 1 && parts[0].Name.ToLowerInvariant() == "endnote")
        {
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var enId = parts[0].Index ?? 1;
            var en = mainPart.EndnotesPart?.Endnotes?
                .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId)
                ?? throw new ArgumentException($"Path not found: {path}");
            // Remove endnote reference from body
            foreach (var enRef in mainPart.Document!.Descendants<EndnoteReference>()
                .Where(r => r.Id?.Value == enId).ToList())
                enRef.Parent?.Remove();
            en.Remove();
            mainPart.EndnotesPart?.Endnotes?.Save();
            mainPart.Document?.Save();
            return null;
        }

        // Handle /chart[N] removal
        var chartRemoveMatch = Regex.Match(path, @"^/chart\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (chartRemoveMatch.Success)
        {
            var chartIdx = int.Parse(chartRemoveMatch.Groups[1].Value);
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var chartParts = mainPart.ChartParts.ToList();
            if (chartIdx < 1 || chartIdx > chartParts.Count)
                throw new ArgumentException($"Chart index {chartIdx} out of range (1..{chartParts.Count})");
            var chartPart = chartParts[chartIdx - 1];
            var relId = mainPart.GetIdOfPart(chartPart);
            // Find and remove the Run containing the ChartReference in the body
            var chartRef = mainPart.Document?.Body?
                .Descendants<C.ChartReference>()
                .FirstOrDefault(cr => cr.Id?.Value == relId);
            if (chartRef != null)
            {
                var run = chartRef.Ancestors<Run>().FirstOrDefault();
                run?.Remove();
            }
            mainPart.DeletePart(chartPart);
            mainPart.Document?.Save();
            return null;
        }

        var element = NavigateToElement(parts, out var ctx)
            ?? throw new ArgumentException($"Path not found: {path}" + (ctx != null ? $". {ctx}" : ""));

        // Clean up ImageParts referenced by any inline/anchor pictures in the element
        var mainPart2 = _doc.MainDocumentPart;
        if (mainPart2 != null)
        {
            foreach (var blip in element.Descendants<A.Blip>())
            {
                var embedId = blip.Embed?.Value;
                if (!string.IsNullOrEmpty(embedId))
                {
                    // Count how many times this embedId is referenced across body + headers + footers
                    var refCount = mainPart2.Document!.Descendants<A.Blip>()
                        .Count(b => b.Embed?.Value == embedId);
                    foreach (var hp in mainPart2.HeaderParts)
                        refCount += hp.Header?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;
                    foreach (var fp in mainPart2.FooterParts)
                        refCount += fp.Footer?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;
                    if (refCount <= 1)
                    {
                        try { mainPart2.DeletePart(embedId); } catch { }
                    }
                }
            }

            // Clean up embedded-object and VML imagedata parts referenced by an
            // EmbeddedObject inside the element being removed. Mirrors the blip
            // cleanup above. Without this, removing a Word OLE leaves both
            // the backing payload part (o:OLEObject r:id) and the custom icon
            // part (v:imagedata r:id) as orphans.
            // BUG-R10-03: OLE inside a HeaderPart/FooterPart stores its rel
            // on the header/footer part itself, so resolve the hosting part
            // from the element's ancestor chain and delete from there.
            foreach (var embObj in element.Descendants<EmbeddedObject>())
            {
                OpenXmlPart hostPart = mainPart2;
                if (embObj.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Header>().FirstOrDefault() is { } hdr)
                    hostPart = (OpenXmlPart?)mainPart2.HeaderParts.FirstOrDefault(p => p.Header == hdr) ?? mainPart2;
                else if (embObj.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Footer>().FirstOrDefault() is { } ftr)
                    hostPart = (OpenXmlPart?)mainPart2.FooterParts.FirstOrDefault(p => p.Footer == ftr) ?? mainPart2;

                // v:imagedata r:id → icon ImagePart
                foreach (var vimg in embObj.Descendants().Where(e => e.LocalName == "imagedata"))
                {
                    var imgRid = vimg.GetAttributes().FirstOrDefault(a => a.LocalName == "id"
                        && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
                    if (!string.IsNullOrEmpty(imgRid))
                    {
                        try { hostPart.DeletePart(imgRid); } catch { }
                    }
                }

                // o:OLEObject r:id → backing embedded payload part
                foreach (var oleEl in embObj.Descendants().Where(e => e.LocalName == "OLEObject"))
                {
                    var oleRid = oleEl.GetAttributes().FirstOrDefault(a => a.LocalName == "id"
                        && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
                    if (!string.IsNullOrEmpty(oleRid))
                    {
                        try { hostPart.DeletePart(oleRid); } catch { }
                    }
                }
            }

            // BUG-R3-09: clean up dead HyperlinkRelationship entries.
            // Each w:hyperlink carries an r:id pointing at a HyperlinkRelationship
            // (an external rel, NOT a part). Deleting the containing element
            // leaves the rel as an orphan that Word silently tolerates but
            // round-tripping tools and validators flag.
            //
            // edge case: same rId may be referenced by multiple hyperlinks. Use
            // a reference count so we only delete rels still uniquely owned by
            // the element being removed.
            var hyperlinksInElement = element.Descendants<Hyperlink>().ToList();
            if (hyperlinksInElement.Count > 0)
            {
                // Collect unique rIds referenced by hyperlinks inside the element.
                var rIdsToCheck = hyperlinksInElement
                    .Select(h => h.Id?.Value)
                    .Where(id => !string.IsNullOrEmpty(id))
                    .Distinct()
                    .ToList();
                foreach (var rId in rIdsToCheck)
                {
                    // Count references in body + headers + footers OUTSIDE of
                    // the element being removed. Deleting `element` would drop
                    // all in-element refs, so any remaining out-of-element ref
                    // means the rel is still live elsewhere.
                    int outsideRefs = 0;
                    foreach (var hl in mainPart2.Document!.Descendants<Hyperlink>())
                    {
                        if (hl.Id?.Value != rId) continue;
                        // Check whether `hl` lives inside `element` (skip self-refs;
                        // those go away with the removal).
                        bool inside = false;
                        for (var anc = (OpenXmlElement?)hl; anc != null; anc = anc.Parent)
                        {
                            if (ReferenceEquals(anc, element)) { inside = true; break; }
                        }
                        if (!inside) outsideRefs++;
                    }
                    foreach (var hp in mainPart2.HeaderParts)
                        outsideRefs += hp.Header?.Descendants<Hyperlink>()
                            .Count(h => h.Id?.Value == rId) ?? 0;
                    foreach (var fp in mainPart2.FooterParts)
                        outsideRefs += fp.Footer?.Descendants<Hyperlink>()
                            .Count(h => h.Id?.Value == rId) ?? 0;
                    if (outsideRefs == 0)
                    {
                        try { mainPart2.DeleteReferenceRelationship(rId!); } catch { }
                    }
                }
            }

            // CONSISTENCY(ref-cleanup): mirror BUG-R3-09 hyperlink cleanup for
            // comments. A removed paragraph that hosted a CommentReference (or
            // a CommentRangeStart/End pair) leaves the matching <w:comment id=N>
            // in comments.xml as an orphan — Word ignores it but validators and
            // round-trip tools flag it, and the sidebar shows ghost comments.
            // For each comment id referenced inside `element`, count outside
            // refs (CommentReference/CommentRangeStart/CommentRangeEnd in the
            // body that are NOT inside `element`); if zero, drop the matching
            // <w:comment> entry.
            var commentIdsInElement = element.Descendants<CommentReference>()
                .Select(cr => cr.Id?.Value)
                .Concat(element.Descendants<CommentRangeStart>().Select(rs => rs.Id?.Value))
                .Concat(element.Descendants<CommentRangeEnd>().Select(re => re.Id?.Value))
                .Where(id => !string.IsNullOrEmpty(id))
                .Distinct()
                .ToList();
            if (commentIdsInElement.Count > 0
                && mainPart2.WordprocessingCommentsPart?.Comments is { } commentsRoot)
            {
                bool IsInside(OpenXmlElement? n)
                {
                    for (var anc = n; anc != null; anc = anc.Parent)
                        if (ReferenceEquals(anc, element)) return true;
                    return false;
                }
                var bodyRoot = mainPart2.Document?.Body;
                foreach (var cid in commentIdsInElement)
                {
                    int outsideRefs = 0;
                    if (bodyRoot != null)
                    {
                        outsideRefs += bodyRoot.Descendants<CommentReference>()
                            .Count(cr => cr.Id?.Value == cid && !IsInside(cr));
                        outsideRefs += bodyRoot.Descendants<CommentRangeStart>()
                            .Count(rs => rs.Id?.Value == cid && !IsInside(rs));
                        outsideRefs += bodyRoot.Descendants<CommentRangeEnd>()
                            .Count(re => re.Id?.Value == cid && !IsInside(re));
                    }
                    if (outsideRefs == 0)
                    {
                        var orphan = commentsRoot.Elements<Comment>()
                            .FirstOrDefault(c => c.Id?.Value == cid);
                        orphan?.Remove();
                    }
                }
                commentsRoot.Save();
            }
        }

        // If removing a Comment, also clean up dangling references in the body
        if (element is Comment comment && comment.Id?.Value is string commentId)
        {
            var body2 = _doc.MainDocumentPart?.Document?.Body;
            if (body2 != null)
            {
                foreach (var rs in body2.Descendants<CommentRangeStart>()
                    .Where(r => r.Id?.Value == commentId).ToList())
                    rs.Remove();
                foreach (var re in body2.Descendants<CommentRangeEnd>()
                    .Where(r => r.Id?.Value == commentId).ToList())
                    re.Remove();
                foreach (var cr in body2.Descendants<CommentReference>()
                    .Where(r => r.Id?.Value == commentId).ToList())
                    cr.Parent?.Remove(); // Remove the containing Run
            }
        }

        // CONSISTENCY(ref-cleanup): mirror Comment cleanup above — removing a
        // NumberingInstance must clear dangling numId references from any
        // paragraph numPr in body/headers/footers/footnotes/endnotes.
        if (element is NumberingInstance numInst && numInst.NumberID?.Value is int removedNumId)
        {
            var mainPart3 = _doc.MainDocumentPart;
            if (mainPart3 != null)
            {
                IEnumerable<OpenXmlElement> roots = new OpenXmlElement?[]
                {
                    mainPart3.Document?.Body,
                    mainPart3.FootnotesPart?.Footnotes,
                    mainPart3.EndnotesPart?.Endnotes,
                }
                .Concat(mainPart3.HeaderParts.Select(h => (OpenXmlElement?)h.Header))
                .Concat(mainPart3.FooterParts.Select(f => (OpenXmlElement?)f.Footer))
                .Where(e => e != null)!;

                foreach (var root in roots)
                {
                    foreach (var numPr in root.Descendants<NumberingProperties>().ToList())
                    {
                        if (numPr.NumberingId?.Val?.Value == removedNumId)
                        {
                            numPr.Remove();
                        }
                    }
                }
            }
        }

        // If removing an oMathPara (M.Paragraph) whose parent w:p has no other
        // meaningful content, remove the wrapper w:p too to avoid zombie paragraphs.
        var wrapperPara = (element is M.Paragraph && element.Parent is Paragraph wp
            && wp.ChildElements.All(c => c == element || c is ParagraphProperties))
            ? wp : null;

        // Refresh textId on parent paragraph if removing a child element (e.g. run)
        var parentPara = element.Ancestors<Paragraph>().FirstOrDefault();

        // CONSISTENCY(tblGrid-sync): when a TableCell is removed via the generic
        // /body/tbl[T]/tr[R]/tc[C] path, the virtual /col[N] path's helper
        // (RemoveTableColumn) is bypassed. After removal, if no row has a cell
        // occupying that column slot anymore, prune the corresponding gridCol
        // so Get() reports correct cols/colWidths and Word doesn't see a stale
        // grid wider than any row. Match RemoveTableColumn's behaviour but
        // applied per-cell.
        // BUG-R2-table-merge BUG-6a: a table with 0 rows is invalid OOXML —
        // Word errors / repairs the file on open. Reject removal of the only
        // remaining row up-front; users must remove the table itself.
        if (element is TableRow lastRowChk
            && lastRowChk.Parent is Table lastRowTbl
            && lastRowTbl.Elements<TableRow>().Count() == 1)
        {
            throw new ArgumentException(
                "Cannot remove the last row of a table. Remove the table itself instead (a table with 0 rows is invalid OOXML).");
        }

        // BUG-R2-table-merge BUG-4a: when the removed row contains any
        // <w:vMerge w:val="restart"/> anchor, every same-column continuation
        // cell in subsequent rows is left orphaned (Word renders it
        // invisible). Snapshot the affected (row, colSlot) pairs before
        // removal so the post-Remove pass can clear continuations.
        List<(Table tbl, int colSlot, int afterRowIdx)>? orphanRestartFixups = null;
        if (element is TableRow vmRow && vmRow.Parent is Table vmTbl)
        {
            int slot = 0;
            foreach (var vmCell in vmRow.Elements<TableCell>())
            {
                int span = vmCell.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
                var vmProp = vmCell.TableCellProperties?.GetFirstChild<VerticalMerge>();
                if (vmProp != null
                    && (vmProp.Val == null || vmProp.Val.Value == MergedCellValues.Restart))
                {
                    var allRows = vmTbl.Elements<TableRow>().ToList();
                    int removedIdx = allRows.IndexOf(vmRow);
                    orphanRestartFixups ??= new();
                    orphanRestartFixups.Add((vmTbl, slot, removedIdx));
                }
                slot += span;
            }
        }

        TableRow? tcRow = null;
        Table? tcTable = null;
        int tcColStart = -1, tcColSpan = 1;
        if (element is TableCell tcRem && element.Parent is TableRow row && row.Parent is Table tbl)
        {
            tcRow = row;
            tcTable = tbl;
            // Compute the gridCol starting index occupied by this cell, summing
            // gridSpan of preceding cells (a merged cell occupies multiple slots).
            tcColStart = 0;
            foreach (var sib in row.Elements<TableCell>())
            {
                if (ReferenceEquals(sib, tcRem)) break;
                tcColStart += sib.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
            }
            tcColSpan = tcRem.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
        }

        // Section removal: cascade-clean Header/Footer parts that this sectPr was the
        // sole reference holder for. Without this, remove /section[N] orphans
        // word/headerN.xml + its rel in document.xml.rels — strict OOXML validators
        // and file-bloat scanners will flag the doc. Only fires for /section[N] or
        // /body/sectPr[N] paths to avoid touching normal paragraph removal.
        var isSectionRemoval = System.Text.RegularExpressions.Regex.IsMatch(
            path, @"^/(?:section|body/sectPr)\[\d+\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (isSectionRemoval && _doc.MainDocumentPart is { } mpForSec)
        {
            // The sectPr that is being removed lives either inside the carrier
            // paragraph's pPr (mid-doc) or directly under body (final). Resolve it
            // from the navigated element first; fall back to the body-level sectPr.
            SectionProperties? targetSectPr =
                (element is Paragraph navP) ? navP.ParagraphProperties?.GetFirstChild<SectionProperties>() : null;
            targetSectPr ??= element as SectionProperties;

            if (targetSectPr != null)
            {
                var headerRefIds = targetSectPr.Elements<HeaderReference>()
                    .Select(r => r.Id?.Value).Where(id => !string.IsNullOrEmpty(id)).ToList();
                var footerRefIds = targetSectPr.Elements<FooterReference>()
                    .Select(r => r.Id?.Value).Where(id => !string.IsNullOrEmpty(id)).ToList();

                bool OtherRefs<T>(string relId) where T : OpenXmlElement
                    => mpForSec.Document?.Body?.Descendants<SectionProperties>()
                        .Where(sp => !ReferenceEquals(sp, targetSectPr))
                        .Any(sp => sp.Elements<T>().Any(r =>
                            (r as HeaderReference)?.Id?.Value == relId
                            || (r as FooterReference)?.Id?.Value == relId)) ?? false;

                foreach (var hid in headerRefIds)
                {
                    if (hid == null) continue;
                    if (OtherRefs<HeaderReference>(hid)) continue;
                    var hp = mpForSec.HeaderParts.FirstOrDefault(p => mpForSec.GetIdOfPart(p) == hid);
                    if (hp != null)
                    {
                        CleanupImageParts(mpForSec, hp.Header?.Descendants<A.Blip>(), hp);
                        mpForSec.DeletePart(hp);
                    }
                }
                foreach (var fid in footerRefIds)
                {
                    if (fid == null) continue;
                    if (OtherRefs<FooterReference>(fid)) continue;
                    var fp = mpForSec.FooterParts.FirstOrDefault(p => mpForSec.GetIdOfPart(p) == fid);
                    if (fp != null)
                    {
                        CleanupImageParts(mpForSec, fp.Footer?.Descendants<A.Blip>(), fp);
                        mpForSec.DeletePart(fp);
                    }
                }
            }
        }

        element.Remove();

        // BUG-R2-table-merge BUG-4a: clear orphan vmerge=continue cells whose
        // restart anchor was just removed. The first remaining row at the
        // removed-row's slot becomes the new "stranded" row; promote its cell
        // to a normal cell (or restart) by removing the <w:vMerge/> child.
        if (orphanRestartFixups != null)
        {
            foreach (var (fxTbl, fxSlot, removedIdx) in orphanRestartFixups)
            {
                var rowsAfter = fxTbl.Elements<TableRow>().ToList();
                for (int ri = removedIdx; ri < rowsAfter.Count; ri++)
                {
                    var targetCell = ResolveCellAtSlot(rowsAfter[ri], fxSlot);
                    if (targetCell == null) break;
                    var tcPrFx = targetCell.TableCellProperties;
                    var vm = tcPrFx?.GetFirstChild<VerticalMerge>();
                    if (vm == null) break;
                    bool isContinue = vm.Val == null
                        || vm.Val.Value == MergedCellValues.Continue;
                    if (!isContinue) break;
                    vm.Remove();
                }
            }
        }

        // CONSISTENCY(tblGrid-sync): after TableCell removal, scan all rows; if
        // any column slot in [tcColStart, tcColStart+tcColSpan) is now unoccupied
        // by every row, drop the corresponding gridCol(s). Otherwise leave the
        // grid alone (column still in use by other rows — partial removal is a
        // ragged-row case which we don't auto-shrink).
        if (tcTable != null && tcColStart >= 0 && tcColSpan >= 1)
        {
            var gridCols = tcTable.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
            if (gridCols != null && gridCols.Count > 0)
            {
                // For each affected slot, check if any remaining row has a cell occupying it.
                bool SlotOccupied(int slotIdx)
                {
                    foreach (var r in tcTable.Elements<TableRow>())
                    {
                        int acc = 0;
                        foreach (var c in r.Elements<TableCell>())
                        {
                            int span = c.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
                            if (slotIdx >= acc && slotIdx < acc + span) return true;
                            acc += span;
                        }
                    }
                    return false;
                }
                // Walk highest slot first to keep indices stable.
                for (int slot = tcColStart + tcColSpan - 1; slot >= tcColStart; slot--)
                {
                    if (slot >= 0 && slot < gridCols.Count && !SlotOccupied(slot))
                    {
                        gridCols[slot].Remove();
                    }
                }
            }
        }

        wrapperPara?.Remove();

        if (parentPara != null)
            parentPara.TextId = GenerateParaId();

        _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments?.Save();
        _doc.MainDocumentPart?.Document?.Save();
        // BUG-R10-03: if we removed a run inside a header/footer, the
        // Save() above only persists the main document part. Also save
        // every header/footer part so the removal actually lands on disk.
        if (_doc.MainDocumentPart != null)
        {
            foreach (var hp in _doc.MainDocumentPart.HeaderParts)
                hp.Header?.Save();
            foreach (var fp in _doc.MainDocumentPart.FooterParts)
                fp.Footer?.Save();
        }
        return null;
    }

    // CONSISTENCY(container-remove-guard): hardcoded list of root-level
    // container paths that must never be removed. Kept in sync (in spirit)
    // with schema entries marked `"container": true` under
    // schemas/help/docx/*.json (document, body, styles, numbering). /settings
    // is also blocked: docSettings are part of the main document part and
    // removing that part destroys the document.
    private static readonly HashSet<string> ProtectedContainerPaths = new(StringComparer.OrdinalIgnoreCase)
    {
        "/body",
        "/document",
        "/styles",
        "/numbering",
        "/settings",
    };

    private static bool IsProtectedContainerPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return false;
        return ProtectedContainerPaths.Contains(path.TrimEnd('/'));
    }

    // CONSISTENCY(container-remove-guard): /body/sectPr needs regex match
    // because it commonly appears with an index (e.g. /body/sectPr[1]). The
    // flat HashSet in ProtectedContainerPaths would require enumerating every
    // index variant, so this is kept as its own predicate.
    private static readonly Regex ProtectedSectPrRegex = new(
        @"^/body/sectPr(?:\[\d+\])?/?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static bool IsProtectedSectPrPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return false;
        return ProtectedSectPrRegex.IsMatch(path);
    }

    /// <summary>
    /// Clean up ImageParts in a header/footer part that are not referenced elsewhere.
    /// </summary>
    private static void CleanupImageParts(MainDocumentPart mainPart, IEnumerable<A.Blip>? blips, OpenXmlPart ownerPart)
    {
        if (blips == null) return;
        foreach (var blip in blips.ToList())
        {
            var embedId = blip.Embed?.Value;
            if (string.IsNullOrEmpty(embedId)) continue;

            // Count references across body + all headers + all footers (excluding the part being deleted)
            var refCount = mainPart.Document?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;
            foreach (var hp in mainPart.HeaderParts.Where(p => p != ownerPart))
                refCount += hp.Header?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;
            foreach (var fp in mainPart.FooterParts.Where(p => p != ownerPart))
                refCount += fp.Footer?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;

            if (refCount == 0)
            {
                try { mainPart.DeletePart(embedId); } catch { }
            }
        }
    }

    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position)
    {
        // Virtual table column path — same-table only. OOXML has no <w:col>
        // element; the move is a (gridCol + per-row tc) shuffle in lockstep.
        var colMoveMatch = Regex.Match(sourcePath, @"^/body/tbl\[(\d+)\]/col\[(\d+)\]$");
        if (colMoveMatch.Success)
        {
            return MoveTableColumn(colMoveMatch, position, targetParentPath);
        }

        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        // Infer --to from --after/--before full path if not specified
        var anchorFullPath = position?.After ?? position?.Before;
        if (string.IsNullOrEmpty(targetParentPath) && anchorFullPath != null && anchorFullPath.StartsWith("/"))
        {
            var lastSlash = anchorFullPath.LastIndexOf('/');
            if (lastSlash > 0)
                targetParentPath = anchorFullPath[..lastSlash];
        }

        // Resolve after/before anchor BEFORE removing the element
        OpenXmlElement? afterAnchor = null, beforeAnchor = null;
        if (position?.After != null)
        {
            var anchorPath = position.After;
            if (!anchorPath.StartsWith("/"))
                anchorPath = (targetParentPath ?? "/body").TrimEnd('/') + "/" + anchorPath;
            afterAnchor = NavigateToElement(ParsePath(anchorPath))
                ?? throw new ArgumentException($"After anchor not found: {position.After}");
        }
        else if (position?.Before != null)
        {
            var anchorPath = position.Before;
            if (!anchorPath.StartsWith("/"))
                anchorPath = (targetParentPath ?? "/body").TrimEnd('/') + "/" + anchorPath;
            beforeAnchor = NavigateToElement(ParsePath(anchorPath))
                ?? throw new ArgumentException($"Before anchor not found: {position.Before}");
        }

        // Determine target parent
        string effectiveParentPath;
        OpenXmlElement targetParent;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within current parent
            targetParent = element.Parent
                ?? throw new InvalidOperationException("Element has no parent");
            // Compute parent path by removing last segment
            var lastSlash = sourcePath.LastIndexOf('/');
            effectiveParentPath = lastSlash > 0 ? sourcePath[..lastSlash] : "/body";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            if (targetParentPath is "/" or "" or "/body")
                targetParent = _doc.MainDocumentPart!.Document!.Body!;
            else
            {
                var tgtParts = ParsePath(targetParentPath);
                targetParent = NavigateToElement(tgtParts)
                    ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
            }
        }

        // CONSISTENCY(word-schema): w:r cannot be a direct child of w:body.
        // Reject obviously invalid parent/child combinations rather than
        // produce malformed XML that breaks downstream queries.
        if (element.LocalName == "r" && targetParent.LocalName == "body")
            throw new ArgumentException(
                "Cannot move a run (w:r) directly to /body. Runs must live inside a paragraph.");

        // CONSISTENCY(word-schema): w:p cannot be nested inside w:p.
        // Without this guard, `move /body/p[1] --to /body/p[3]` happily
        // appends the source paragraph as a child of the target paragraph,
        // producing schema-invalid <w:p><w:p>...</w:p></w:p>. Users almost
        // always meant "place after", so steer them toward --after.
        if (element.LocalName == "p" && targetParent.LocalName == "p")
            throw new ArgumentException(
                "Cannot move a paragraph into another paragraph (would create invalid <w:p><w:p>). " +
                "To place after a paragraph, use `--after <path>`; to reorder within /body, omit --to.");
        // Same guard for moving table/tbl into a paragraph.
        if ((element.LocalName == "tbl" || element.LocalName == "table") && targetParent.LocalName == "p")
            throw new ArgumentException(
                "Cannot move a table into a paragraph. Use `--after <paragraph-path>` to place it after.");

        element.Remove();

        // Insert at the resolved position
        if (afterAnchor != null)
        {
            afterAnchor.InsertAfterSelf(element);
        }
        else if (beforeAnchor != null)
        {
            beforeAnchor.InsertBeforeSelf(element);
        }
        else if (position?.Index is int index)
        {
            var sameTypeSiblings = targetParent.ChildElements
                .Where(e => e.LocalName == element.LocalName).ToList();
            if (index >= 0 && index < sameTypeSiblings.Count)
                sameTypeSiblings[index].InsertBeforeSelf(element);
            else
                AppendToParent(targetParent, element);
        }
        else
        {
            targetParent.AppendChild(element);
        }

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == element.LocalName).ToList();
        var newIdx = siblings.IndexOf(element) + 1;
        return $"{effectiveParentPath}/{element.LocalName}[{newIdx}]";
    }

    public (string NewPath1, string NewPath2) Swap(string path1, string path2)
    {
        var parts1 = ParsePath(path1);
        var elem1 = NavigateToElement(parts1)
            ?? throw new ArgumentException($"Element not found: {path1}");
        var parts2 = ParsePath(path2);
        var elem2 = NavigateToElement(parts2)
            ?? throw new ArgumentException($"Element not found: {path2}");

        if (elem1.Parent != elem2.Parent)
            throw new ArgumentException("Cannot swap elements with different parents");

        PowerPointHandler.SwapXmlElements(elem1, elem2);
        _doc.MainDocumentPart?.Document?.Save();

        // Recompute paths
        var parent = elem1.Parent!;
        var lastSlash = path1.LastIndexOf('/');
        var parentPath = lastSlash > 0 ? path1[..lastSlash] : "/body";

        var siblings1 = parent.ChildElements.Where(e => e.LocalName == elem1.LocalName).ToList();
        var newIdx1 = siblings1.IndexOf(elem1) + 1;
        var siblings2 = parent.ChildElements.Where(e => e.LocalName == elem2.LocalName).ToList();
        var newIdx2 = siblings2.IndexOf(elem2) + 1;
        return ($"{parentPath}/{elem1.LocalName}[{newIdx1}]", $"{parentPath}/{elem2.LocalName}[{newIdx2}]");
    }

    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position)
    {
        // Virtual table column clone — same-table only.
        var colCopyMatch = Regex.Match(sourcePath, @"^/body/tbl\[(\d+)\]/col\[(\d+)\]$");
        if (colCopyMatch.Success)
        {
            return CopyTableColumn(colCopyMatch, position, targetParentPath);
        }

        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        // Bookmarks are a start/end pair spanning arbitrary content; the
        // virtual `/bookmark[@name=X]` selector (and any bare bookmarkStart/
        // bookmarkEnd path) points at one marker only, so a naive clone
        // produces a never-closed bookmark. Reject with a direction to clone
        // the containing paragraph or range instead.
        if (element is BookmarkStart or BookmarkEnd)
        {
            throw new ArgumentException(
                $"Cannot clone '{sourcePath}': bookmarks span content via a start/end pair. Clone the containing paragraph (or a range) instead.");
        }

        // Part-scoped elements: <w:footnote>, <w:endnote>, <w:comment> live
        // in their own XML parts. Cloning the raw element into main-document
        // body produces schema-invalid OOXML (body can only reference these
        // via <w:footnoteReference>, <w:endnoteReference>, <w:commentReference>
        // and commentRangeStart/End). This rejection is clone-specific — the
        // legitimate `add --type footnote/endnote/comment --prop text=...`
        // path uses dedicated helpers that insert a reference at the target
        // and append the content to the correct part.
        if (element is Footnote)
        {
            throw new ArgumentException(
                $"Cannot clone '{sourcePath}': <w:footnote> belongs in /word/footnotes.xml. Use `add --type footnote --prop text=...` to create a new footnote, or clone a paragraph containing a footnoteReference.");
        }
        if (element is Endnote)
        {
            throw new ArgumentException(
                $"Cannot clone '{sourcePath}': <w:endnote> belongs in /word/endnotes.xml. Use `add --type endnote --prop text=...` to create a new endnote, or clone a paragraph containing an endnoteReference.");
        }
        if (element is Comment)
        {
            throw new ArgumentException(
                $"Cannot clone '{sourcePath}': <w:comment> belongs in /word/comments.xml. Use `add --type comment --prop text=...` to create a new comment, or clone a paragraph containing commentRangeStart/End.");
        }
        // Equation content (<m:oMathPara>, <m:oMath>) lives inside a <w:p>
        // (paragraph) and is not itself a valid direct child of <w:body>.
        // Cloning a bare oMathPara/oMath into /body produces schema-invalid
        // OOXML. Direct users to clone the containing paragraph instead.
        if (element is DocumentFormat.OpenXml.Math.Paragraph)
        {
            throw new ArgumentException(
                $"Cannot clone '{sourcePath}': equation content lives inside a paragraph; clone /body/p[N] instead.");
        }
        if (element is DocumentFormat.OpenXml.Math.OfficeMath)
        {
            throw new ArgumentException(
                $"Cannot clone '{sourcePath}': equation content lives inside a paragraph; clone /body/p[N] instead.");
        }

        OpenXmlElement targetParent;
        if (targetParentPath is "/" or "" or "/body")
        {
            targetParent = _doc.MainDocumentPart!.Document!.Body!;
            targetParentPath = "/body";
        }
        else if (targetParentPath == "/styles")
        {
            var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
                ?? throw new ArgumentException("Target parent not found: /styles");
            targetParent = stylesPart.Styles
                ?? throw new ArgumentException("Target parent not found: /styles");
        }
        else
        {
            var tgtParts = ParsePath(targetParentPath);
            targetParent = NavigateToElement(tgtParts)
                ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
        }

        // Reject self-clone (source == targetParent) and
        // ancestor-into-descendant (cloning /body into /body/... would stack
        // the body inside itself). The node-level check here complements the
        // LocalName-based ValidateParentChild below, catching cases where the
        // shapes would nominally match but the operation is still degenerate.
        if (ReferenceEquals(element, targetParent))
            throw new ArgumentException(
                $"Cannot clone '{sourcePath}' into itself.");
        if (targetParent.Ancestors().Contains(element))
            throw new ArgumentException(
                $"Cannot clone '{sourcePath}' into one of its own descendants ('{targetParentPath}').");

        // Map OOXML local name to the type token ValidateParentChild expects
        // (mirrors the dispatcher in Add.cs).
        var typeToken = MapLocalNameToAddType(element.LocalName);
        ValidateParentChild(targetParent, targetParentPath, typeToken);

        var clone = element.CloneNode(true);

        // Regenerate paraIds on cloned paragraphs to ensure uniqueness
        var clonedParas = clone is Paragraph cp
            ? new[] { cp }
            : clone.Descendants<Paragraph>().ToArray();
        foreach (var p in clonedParas)
        {
            p.ParagraphId = GenerateParaId();
            p.TextId = GenerateParaId();
        }

        // Regenerate bookmark ids/names so a cloned paragraph containing
        // <w:bookmarkStart>/<w:bookmarkEnd> doesn't introduce duplicate
        // numeric ids or duplicate names (the latter silently breaks
        // hyperlink/ref resolution, the former is a schema violation).
        var docBody = _doc.MainDocumentPart?.Document?.Body;
        if (docBody != null)
        {
            var existingIds = docBody.Descendants<BookmarkStart>()
                .Where(b => !ReferenceEquals(b, clone) && !b.Ancestors().Contains(clone))
                .Select(b => int.TryParse(b.Id?.Value, out var id) ? id : 0);
            var existingNames = new HashSet<string>(
                docBody.Descendants<BookmarkStart>()
                    .Where(b => !ReferenceEquals(b, clone) && !b.Ancestors().Contains(clone))
                    .Select(b => b.Name?.Value ?? "")
                    .Where(n => n.Length > 0));
            var nextId = existingIds.Any() ? existingIds.Max() + 1 : 1;

            // Collect pairs inside the clone (by matching old Id).
            var startsInClone = clone is BookmarkStart bsSelf
                ? new[] { bsSelf }
                : clone.Descendants<BookmarkStart>().ToArray();
            var endsInClone = clone is BookmarkEnd beSelf
                ? new[] { beSelf }
                : clone.Descendants<BookmarkEnd>().ToArray();

            foreach (var bs in startsInClone)
            {
                var oldId = bs.Id?.Value;
                var newId = nextId++.ToString();
                bs.Id = newId;
                var name = bs.Name?.Value ?? "";
                if (string.IsNullOrEmpty(name) || existingNames.Contains(name))
                {
                    var baseName = string.IsNullOrEmpty(name) ? "bm" : name;
                    var candidate = $"{baseName}_{newId}";
                    while (existingNames.Contains(candidate))
                        candidate = $"{baseName}_{nextId++}";
                    bs.Name = candidate;
                    existingNames.Add(candidate);
                }
                else
                {
                    existingNames.Add(name);
                }
                // Retarget matching ends.
                if (oldId != null)
                {
                    foreach (var be in endsInClone)
                    {
                        if (be.Id?.Value == oldId)
                            be.Id = newId;
                    }
                }
            }
        }

        // Regenerate revision ids on cloned <w:ins>/<w:del> elements so the
        // clone doesn't collide with the source (or any other in-doc) w:id.
        // Semantic validators reject duplicate ins/del ids and Word treats
        // two elements with the same id as a single tracked change.
        if (docBody != null)
        {
            var existingRevIds = new HashSet<int>(
                docBody.Descendants<InsertedRun>()
                    .Where(e => !ReferenceEquals(e, clone) && !e.Ancestors().Contains(clone))
                    .Select(e => int.TryParse(e.Id?.Value, out var i) ? i : -1)
                    .Where(i => i >= 0));
            foreach (var i in docBody.Descendants<DeletedRun>()
                .Where(e => !ReferenceEquals(e, clone) && !e.Ancestors().Contains(clone))
                .Select(e => int.TryParse(e.Id?.Value, out var i) ? i : -1)
                .Where(i => i >= 0))
            {
                existingRevIds.Add(i);
            }
            var nextRevId = existingRevIds.Count > 0 ? existingRevIds.Max() + 1 : 1;

            var insInClone = clone is InsertedRun irSelf
                ? new[] { irSelf }
                : clone.Descendants<InsertedRun>().ToArray();
            var delInClone = clone is DeletedRun drSelf
                ? new[] { drSelf }
                : clone.Descendants<DeletedRun>().ToArray();
            foreach (var ir in insInClone)
            {
                ir.Id = (nextRevId++).ToString();
            }
            foreach (var dr in delInClone)
            {
                dr.Id = (nextRevId++).ToString();
            }
        }

        // Regenerate wp:docPr/@id on cloned drawings. <wp:docPr> requires
        // document-unique numeric ids; cloning a paragraph containing a
        // chart/picture/shape duplicates the id and fails validation.
        // Matching pic:cNvPr (inside DrawingML picture) carries the same id
        // by convention (see CreateImageRun / AddChart), so keep them in sync.
        {
            var docPrInClone = clone is DW.DocProperties dpSelf
                ? new[] { dpSelf }
                : clone.Descendants<DW.DocProperties>().ToArray();
            var nextDocPrId = NextDocPropId();
            foreach (var dp in docPrInClone)
            {
                var oldId = dp.Id?.Value;
                var newId = nextDocPrId++;
                dp.Id = newId;

                // Update matching pic:cNvPr ids within the same drawing subtree.
                var drawingAncestor = (OpenXmlElement?)dp.Parent;
                while (drawingAncestor != null && drawingAncestor is not Drawing)
                    drawingAncestor = drawingAncestor.Parent;
                var scope = (OpenXmlElement?)drawingAncestor ?? dp.Parent ?? clone;
                if (scope != null)
                {
                    foreach (var picCnv in scope.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>())
                    {
                        if (oldId == null || picCnv.Id?.Value == oldId)
                            picCnv.Id = newId;
                    }
                    foreach (var wpsCnv in scope.Descendants<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape>()
                        .SelectMany(s => s.Descendants<DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties>()))
                    {
                        if (oldId == null || wpsCnv.Id?.Value == oldId)
                            wpsCnv.Id = newId;
                    }
                }
            }
        }

        // Handle find: anchor sentinel up front — Add() uses AddAtFindPosition
        // to split the paragraph at a text-match point, but CopyFrom has no
        // analogous split-based insertion path. The common case (e.g. cloning
        // a paragraph before/after a find: anchor) is well served by
        // resolving the anchor to the containing paragraph at the targetParent
        // level and inserting the clone as that paragraph's before/after
        // sibling.
        if (position != null)
        {
            var anchorPath = position.After ?? position.Before;
            if (anchorPath != null && anchorPath.StartsWith("find:", StringComparison.OrdinalIgnoreCase))
            {
                var findValue = anchorPath["find:".Length..];
                var (pattern, isRegex) = ParseFindPattern(findValue);
                if (string.IsNullOrEmpty(pattern))
                    throw new ArgumentException("find: pattern must not be empty.");
                var hit = FindParagraphContainingText(targetParent, targetParentPath, pattern, isRegex)
                    ?? throw new ArgumentException(
                        $"Text '{findValue}' not found in any paragraph under {targetParentPath}.");
                var paragraphs = targetParent.Elements<Paragraph>().ToList();
                var anchorIdx = paragraphs.IndexOf(hit.Para);
                if (anchorIdx < 0)
                    throw new ArgumentException($"find: anchor resolved outside {targetParentPath}.");

                if (position.After != null)
                    hit.Para.InsertAfterSelf(clone);
                else
                    hit.Para.InsertBeforeSelf(clone);

                _doc.MainDocumentPart?.Document?.Save();
                var fSiblings = targetParent.ChildElements.Where(e => e.LocalName == clone.LocalName).ToList();
                var fNewIdx = fSiblings.IndexOf(clone) + 1;
                return $"{targetParentPath}/{clone.LocalName}[{fNewIdx}]";
            }
        }

        // Resolve --after/--before to a concrete int index in targetParent,
        // mirroring what Add() does. Without this, CopyFrom silently ignored
        // anchor-based positions and always appended.
        var index = ResolveAnchorPosition(targetParent, targetParentPath, position);

        InsertAtPosition(targetParent, clone, index);

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == clone.LocalName).ToList();
        var newIdx = siblings.IndexOf(clone) + 1;
        return $"{targetParentPath}/{clone.LocalName}[{newIdx}]";
    }

    /// <summary>
    /// Map an OpenXML LocalName to the type token ValidateParentChild expects
    /// (the same tokens the Add() dispatcher uses). Unknown names fall
    /// through to the local name itself, which produces no special rejection
    /// in ValidateParentChild — matching pre-fix behaviour for exotic types.
    /// </summary>
    private static string MapLocalNameToAddType(string localName) =>
        localName.ToLowerInvariant() switch
        {
            "p" => "paragraph",
            "tbl" => "table",
            "tr" => "row",
            "tc" => "cell",
            "r" => "run",
            "body" => "body",
            // Keep "sectpr" distinct from "section": the former represents a raw
            // <w:sectPr> element being cloned (only valid as body-level singleton)
            // and is rejected by ValidateParentChild; the latter is the user-level
            // Add verb that creates a paragraph carrying a section break.
            "sectpr" => "sectpr",
            "sdt" => "sdt",
            "hyperlink" => "hyperlink",
            "bookmarkstart" or "bookmarkend" => "bookmark",
            // Part-scoped elements — ValidateParentChild rejects these wholesale
            // when the target parent is body/paragraph/cell, preventing raw
            // <w:footnote>/<w:endnote>/<w:comment> from being cloned into
            // main-document content.
            "footnote" => "footnote",
            "endnote" => "endnote",
            "comment" => "comment",
            _ => localName.ToLowerInvariant()
        };

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        // Paragraphs require pPr-aware insertion so an index 0 (which resolves
        // to the <w:pPr> child when present) does not shove content in front
        // of the paragraph properties.
        if (parent is Paragraph para)
        {
            InsertIntoParagraph(para, element, index);
            return;
        }

        if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                AppendToParent(parent, element);
        }
        else
        {
            AppendToParent(parent, element);
        }
    }

    // ==================== Track Changes ====================

    /// <summary>
    /// Accept all tracked changes in the document.
    /// - w:ins (InsertedRun): unwrap — keep inner content, remove wrapper
    /// - w:del (DeletedRun): remove entire element
    /// - w:rPrChange (RunPropertiesChange): remove change marker, keep current formatting
    /// - w:pPrChange (ParagraphPropertiesChange): remove change marker, keep current formatting
    /// - w:sectPrChange (SectionPropertiesChange): remove change marker
    /// - w:tblPrChange (TablePropertyExceptionChange): remove change marker
    /// - w:trPr/w:ins (table row insertion): keep row, remove marker
    /// </summary>
    private int AcceptAllChanges()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 0;

        int count = 0;

        // Accept w:ins — unwrap (keep inner content)
        foreach (var ins in body.Descendants<InsertedRun>().ToList())
        {
            var parent = ins.Parent;
            if (parent == null) { ins.Remove(); count++; continue; }
            foreach (var child in ins.ChildElements.ToList())
                parent.InsertBefore(child.CloneNode(true), ins);
            ins.Remove();
            count++;
        }

        // Accept w:del — remove entirely (deletions are discarded)
        foreach (var del in body.Descendants<DeletedRun>().ToList())
        {
            del.Remove();
            count++;
        }

        // Accept w:rPrChange — remove the change element, keep current run properties
        foreach (var rPrChange in body.Descendants<RunPropertiesChange>().ToList())
        {
            rPrChange.Remove();
            count++;
        }

        // Accept w:pPrChange — remove the change element, keep current paragraph properties
        foreach (var pPrChange in body.Descendants<ParagraphPropertiesChange>().ToList())
        {
            pPrChange.Remove();
            count++;
        }

        // Accept w:sectPrChange — remove the change element
        foreach (var sectPrChange in body.Descendants<SectionPropertiesChange>().ToList())
        {
            sectPrChange.Remove();
            count++;
        }

        // Accept table property changes
        foreach (var tblPrChange in body.Descendants<TablePropertiesChange>().ToList())
        {
            tblPrChange.Remove();
            count++;
        }

        // Accept table row property changes (w:trPr containing w:ins)
        foreach (var trPr in body.Descendants<TableRowProperties>().ToList())
        {
            var trIns = trPr.GetFirstChild<InsertedRun>();
            if (trIns != null) { trIns.Remove(); count++; }
        }

        // Accept w:moveTo / w:moveFrom
        foreach (var moveFrom in body.Descendants<MoveFromRun>().ToList())
        {
            moveFrom.Remove();
            count++;
        }
        foreach (var moveTo in body.Descendants<MoveToRun>().ToList())
        {
            var parent = moveTo.Parent;
            if (parent == null) { moveTo.Remove(); count++; continue; }
            foreach (var child in moveTo.ChildElements.ToList())
                parent.InsertBefore(child.CloneNode(true), moveTo);
            moveTo.Remove();
            count++;
        }

        // Remove move range markers
        foreach (var marker in body.Descendants<MoveFromRangeStart>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveFromRangeEnd>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveToRangeStart>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveToRangeEnd>().ToList()) marker.Remove();

        _doc.MainDocumentPart?.Document?.Save();
        return count;
    }

    /// <summary>
    /// Reject all tracked changes in the document.
    /// - w:ins (InsertedRun): remove entire element (discard insertion)
    /// - w:del (DeletedRun): unwrap — restore content, convert w:delText to w:t
    /// - w:rPrChange: restore original formatting from inside the change element
    /// - w:pPrChange: restore original paragraph properties
    /// - w:sectPrChange: restore original section properties
    /// </summary>
    private int RejectAllChanges()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 0;

        int count = 0;

        // Reject w:ins — remove entirely (discard insertions)
        foreach (var ins in body.Descendants<InsertedRun>().ToList())
        {
            ins.Remove();
            count++;
        }

        // Reject w:del — unwrap, convert w:delText to w:t
        foreach (var del in body.Descendants<DeletedRun>().ToList())
        {
            var parent = del.Parent;
            if (parent == null) { del.Remove(); count++; continue; }
            foreach (var child in del.ChildElements.ToList())
            {
                var clone = child.CloneNode(true);
                // Convert DeletedText elements to Text elements
                foreach (var delText in clone.Descendants<DeletedText>().ToList())
                {
                    var text = new Text(delText.Text);
                    if (delText.Space != null)
                        text.Space = delText.Space;
                    delText.Parent?.ReplaceChild(text, delText);
                }
                parent.InsertBefore(clone, del);
            }
            del.Remove();
            count++;
        }

        // Reject w:rPrChange — restore original run properties
        foreach (var rPrChange in body.Descendants<RunPropertiesChange>().ToList())
        {
            var rPr = rPrChange.Parent as RunProperties;
            if (rPr != null)
            {
                var originalProps = rPrChange.GetFirstChild<PreviousRunProperties>();
                if (originalProps != null)
                {
                    // Replace current run properties with original ones
                    var run = rPr.Parent;
                    if (run != null)
                    {
                        var newRPr = new RunProperties();
                        foreach (var child in originalProps.ChildElements.ToList())
                            newRPr.AppendChild(child.CloneNode(true));
                        run.ReplaceChild(newRPr, rPr);
                    }
                }
                else
                {
                    rPrChange.Remove();
                }
            }
            else
            {
                rPrChange.Remove();
            }
            count++;
        }

        // Reject w:pPrChange — restore original paragraph properties
        foreach (var pPrChange in body.Descendants<ParagraphPropertiesChange>().ToList())
        {
            var pPr = pPrChange.Parent as ParagraphProperties;
            if (pPr != null)
            {
                var originalProps = pPrChange.GetFirstChild<ParagraphPropertiesExtended>();
                if (originalProps != null)
                {
                    var para = pPr.Parent;
                    if (para != null)
                    {
                        var newPPr = new ParagraphProperties();
                        foreach (var child in originalProps.ChildElements.ToList())
                            newPPr.AppendChild(child.CloneNode(true));
                        para.ReplaceChild(newPPr, pPr);
                    }
                }
                else
                {
                    pPrChange.Remove();
                }
            }
            else
            {
                pPrChange.Remove();
            }
            count++;
        }

        // Reject w:sectPrChange — restore original section properties
        foreach (var sectPrChange in body.Descendants<SectionPropertiesChange>().ToList())
        {
            var sectPr = sectPrChange.Parent as SectionProperties;
            if (sectPr != null)
            {
                var originalProps = sectPrChange.GetFirstChild<PreviousSectionProperties>();
                if (originalProps != null)
                {
                    var parent = sectPr.Parent;
                    if (parent != null)
                    {
                        var newSectPr = new SectionProperties();
                        foreach (var child in originalProps.ChildElements.ToList())
                            newSectPr.AppendChild(child.CloneNode(true));
                        parent.ReplaceChild(newSectPr, sectPr);
                    }
                }
                else
                {
                    sectPrChange.Remove();
                }
            }
            else
            {
                sectPrChange.Remove();
            }
            count++;
        }

        // Reject table property changes — restore original table properties
        foreach (var tblPrChange in body.Descendants<TablePropertiesChange>().ToList())
        {
            var tblPr = tblPrChange.Parent as TableProperties;
            if (tblPr != null)
            {
                var originalProps = tblPrChange.GetFirstChild<PreviousTableProperties>();
                if (originalProps != null)
                {
                    var tbl = tblPr.Parent;
                    if (tbl != null)
                    {
                        var newTblPr = new TableProperties();
                        foreach (var child in originalProps.ChildElements.ToList())
                            newTblPr.AppendChild(child.CloneNode(true));
                        tbl.ReplaceChild(newTblPr, tblPr);
                    }
                }
                else
                {
                    tblPrChange.Remove();
                }
            }
            else
            {
                tblPrChange.Remove();
            }
            count++;
        }

        // Reject w:moveTo — remove (discard the move target)
        foreach (var moveTo in body.Descendants<MoveToRun>().ToList())
        {
            moveTo.Remove();
            count++;
        }
        // Reject w:moveFrom — unwrap (restore original position)
        foreach (var moveFrom in body.Descendants<MoveFromRun>().ToList())
        {
            var parent = moveFrom.Parent;
            if (parent == null) { moveFrom.Remove(); count++; continue; }
            foreach (var child in moveFrom.ChildElements.ToList())
                parent.InsertBefore(child.CloneNode(true), moveFrom);
            moveFrom.Remove();
            count++;
        }

        // Remove move range markers
        foreach (var marker in body.Descendants<MoveFromRangeStart>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveFromRangeEnd>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveToRangeStart>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveToRangeEnd>().ToList()) marker.Remove();

        _doc.MainDocumentPart?.Document?.Save();
        return count;
    }

    // -------- Word virtual table-column ops --------
    //
    // OOXML has no <w:col> child of <w:tbl>; columns are implicit (gridCol +
    // per-row tc). These helpers synthesize Remove/Move/CopyFrom for the
    // virtual `/body/tbl[N]/col[C]` path. Same-table only — cross-table is
    // rejected because grid widths and row counts differ ambiguously.

    private (Table table, TableGrid grid) ResolveBodyTable(int tableIdx)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");
        var tables = body.Elements<Table>().ToList();
        if (tableIdx < 1 || tableIdx > tables.Count)
            throw new ArgumentException($"Table {tableIdx} not found at /body (total: {tables.Count})");
        var table = tables[tableIdx - 1];
        var grid = table.GetFirstChild<TableGrid>()
            ?? throw new InvalidOperationException("Table has no <w:tblGrid>");
        return (table, grid);
    }

    // Resolve the TableCell occupying a specific gridCol slot in a row,
    // accounting for gridSpan-merged cells. Returns null if the row's total
    // span is shorter than slot+1.
    private static TableCell? ResolveCellAtSlot(TableRow trow, int slot)
    {
        int acc = 0;
        foreach (var c in trow.Elements<TableCell>())
        {
            int span = c.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
            if (slot >= acc && slot < acc + span) return c;
            acc += span;
        }
        return null;
    }

    private static void GuardNoMergesInColumn(Table table, int colIdx, string action)
    {
        // gridSpan/vMerge in the affected column slot would silently break.
        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            if (colIdx - 1 < cells.Count)
            {
                var tc = cells[colIdx - 1];
                var tcPr = tc.GetFirstChild<TableCellProperties>();
                var span = tcPr?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
                if (span > 1 || tcPr?.GetFirstChild<VerticalMerge>() != null)
                    throw new ArgumentException(
                        $"Cannot {action} column {colIdx}: a row contains a merged cell (gridSpan/vMerge) " +
                        "spanning that column. Unmerge before performing column-level operations.");
            }
        }
    }

    private void RemoveTableColumn(Match colMatch)
    {
        var tableIdx = int.Parse(colMatch.Groups[1].Value);
        var colIdx = int.Parse(colMatch.Groups[2].Value);
        var (table, grid) = ResolveBodyTable(tableIdx);
        var gridCols = grid.Elements<GridColumn>().ToList();
        if (colIdx < 1 || colIdx > gridCols.Count)
            throw new ArgumentException($"Column {colIdx} not found (total: {gridCols.Count})");

        GuardNoMergesInColumn(table, colIdx, "remove");

        gridCols[colIdx - 1].Remove();
        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            if (colIdx - 1 < cells.Count)
                cells[colIdx - 1].Remove();
        }
    }

    private static int? ResolveSameTableColumnAnchor(InsertPosition? position, int tableIdx, int? sourceColIdx)
    {
        if (position?.After == null && position?.Before == null)
            return position?.Index;
        var anchorPath = position.After ?? position.Before!;
        var anchorMatch = Regex.Match(anchorPath, @"^/body/tbl\[(\d+)\]/col\[(\d+)\]$");
        if (!anchorMatch.Success || int.Parse(anchorMatch.Groups[1].Value) != tableIdx)
            throw new ArgumentException(
                $"Column anchor must be a column in the same table: /body/tbl[{tableIdx}]/col[N]. Got: {anchorPath}");
        var anchorColIdx = int.Parse(anchorMatch.Groups[2].Value);
        if (sourceColIdx.HasValue && anchorColIdx == sourceColIdx.Value)
            return -1; // self-anchor sentinel
        var target = position.After != null ? anchorColIdx : anchorColIdx - 1; // 0-based
        if (sourceColIdx.HasValue && sourceColIdx.Value < anchorColIdx) target -= 1;
        return target;
    }

    private string MoveTableColumn(Match colMatch, InsertPosition? position, string? targetParentPath)
    {
        var tableIdx = int.Parse(colMatch.Groups[1].Value);
        var colIdx = int.Parse(colMatch.Groups[2].Value);

        if (!string.IsNullOrEmpty(targetParentPath))
        {
            var expected = $"/body/tbl[{tableIdx}]";
            if (!string.Equals(targetParentPath, expected, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Cross-table column move is not supported. Source column's table is {expected}; target was {targetParentPath}.");
        }

        var (table, grid) = ResolveBodyTable(tableIdx);
        var gridCols = grid.Elements<GridColumn>().ToList();
        if (colIdx < 1 || colIdx > gridCols.Count)
            throw new ArgumentException($"Column {colIdx} not found (total: {gridCols.Count})");

        GuardNoMergesInColumn(table, colIdx, "move");

        var targetIdx = ResolveSameTableColumnAnchor(position, tableIdx, colIdx);
        if (targetIdx == -1)
            return $"/body/tbl[{tableIdx}]/col[{colIdx}]";

        var movingGridCol = gridCols[colIdx - 1];
        movingGridCol.Remove();
        var movingCells = new List<TableCell>();
        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            if (colIdx - 1 < cells.Count)
            {
                movingCells.Add(cells[colIdx - 1]);
                cells[colIdx - 1].Remove();
            }
            else
            {
                movingCells.Add(new TableCell(new Paragraph()));
            }
        }

        var remainingGridCols = grid.Elements<GridColumn>().ToList();
        if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < remainingGridCols.Count)
            remainingGridCols[targetIdx.Value].InsertBeforeSelf(movingGridCol);
        else
            grid.AppendChild(movingGridCol);

        int rowIdx = 0;
        foreach (var row in table.Elements<TableRow>())
        {
            var rowCells = row.Elements<TableCell>().ToList();
            var movingCell = movingCells[rowIdx++];
            if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < rowCells.Count)
                rowCells[targetIdx.Value].InsertBeforeSelf(movingCell);
            else
                row.AppendChild(movingCell);
        }

        _doc.MainDocumentPart?.Document?.Save();
        var newGridCols = grid.Elements<GridColumn>().ToList();
        var newColIdx = newGridCols.IndexOf(movingGridCol) + 1;
        return $"/body/tbl[{tableIdx}]/col[{newColIdx}]";
    }

    private string CopyTableColumn(Match colMatch, InsertPosition? position, string? targetParentPath)
    {
        var tableIdx = int.Parse(colMatch.Groups[1].Value);
        var colIdx = int.Parse(colMatch.Groups[2].Value);

        if (!string.IsNullOrEmpty(targetParentPath))
        {
            var expected = $"/body/tbl[{tableIdx}]";
            if (!string.Equals(targetParentPath, expected, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Cross-table column copy is not supported. Source column's table is {expected}; target was {targetParentPath}.");
        }

        var (table, grid) = ResolveBodyTable(tableIdx);
        var gridCols = grid.Elements<GridColumn>().ToList();
        if (colIdx < 1 || colIdx > gridCols.Count)
            throw new ArgumentException($"Column {colIdx} not found (total: {gridCols.Count})");

        GuardNoMergesInColumn(table, colIdx, "copy");

        var targetIdx = ResolveSameTableColumnAnchor(position, tableIdx, sourceColIdx: null);

        var clonedGridCol = (GridColumn)gridCols[colIdx - 1].CloneNode(true);
        var clonedCells = new List<TableCell>();
        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            clonedCells.Add(colIdx - 1 < cells.Count
                ? (TableCell)cells[colIdx - 1].CloneNode(true)
                : new TableCell(new Paragraph()));
        }

        var siblingsGrid = grid.Elements<GridColumn>().ToList();
        if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < siblingsGrid.Count)
            siblingsGrid[targetIdx.Value].InsertBeforeSelf(clonedGridCol);
        else
            grid.AppendChild(clonedGridCol);

        int rowIdx = 0;
        foreach (var row in table.Elements<TableRow>())
        {
            var rowCells = row.Elements<TableCell>().ToList();
            var clone = clonedCells[rowIdx++];
            if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < rowCells.Count)
                rowCells[targetIdx.Value].InsertBeforeSelf(clone);
            else
                row.AppendChild(clone);
        }

        // Re-assign paraId to all cloned paragraphs to avoid duplicates.
        foreach (var clonedCell in clonedCells)
            foreach (var p in clonedCell.Descendants<Paragraph>())
                AssignParaId(p);

        _doc.MainDocumentPart?.Document?.Save();
        var newGridCols = grid.Elements<GridColumn>().ToList();
        var newColIdx = newGridCols.IndexOf(clonedGridCol) + 1;
        return $"/body/tbl[{tableIdx}]/col[{newColIdx}]";
    }
}
