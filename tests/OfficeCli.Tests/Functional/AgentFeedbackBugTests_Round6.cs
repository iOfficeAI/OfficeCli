// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests that prove bugs found during Agent A Round 6 testing.
/// These tests are expected to FAIL until the bugs are fixed.
///
/// BUG-5: Set url/text on a hyperlink node silently fails.
///        SetElement() has no `element is Hyperlink` branch, so targeting
///        a hyperlink path (e.g. /body/p[1]/hyperlink[1]) with properties
///        like url, text, href, or link falls through all type checks and
///        does nothing. The "link" property is only handled for Run elements
///        (line ~861), not for Hyperlink elements.
///        Root cause: WordHandler.Set.cs SetElement() only handles
///        BookmarkStart, SdtBlock/SdtRun, Run, Paragraph, TableCell,
///        TableRow, Table. No Hyperlink case exists.
///
/// BUG-6: Set link=<new-url> on a hyperlink node silently fails.
///        The "link" property handler exists only in the `element is Run`
///        branch (line ~861). When targeting a Hyperlink element directly
///        (e.g. /body/p[1]/hyperlink[1]), there is no matching branch,
///        so the property is silently dropped and the URL remains unchanged.
///
/// BUG-8: After adding a header, setting orientation on the section
///        inserts PageSize via EnsureSectPrPageSize() which prepends
///        the element when no SectionType exists. This places PageSize
///        BEFORE HeaderReference, violating OOXML schema order
///        (HeaderReference must precede SectionType/PageSize/PageMargin).
///        Word may silently ignore or corrupt such files.
///        Root cause: EnsureSectPrPageSize() does not account for
///        HeaderReference/FooterReference children when inserting.
/// </summary>
public class AgentFeedbackBugTests_Round6 : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public AgentFeedbackBugTests_Round6()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== BUG-5: Hyperlink URL/text Set silently fails ====================

    /// <summary>
    /// BUG-5a: Setting "url" on a hyperlink node should update the hyperlink target.
    /// Currently, Set on a hyperlink path falls through all type checks in SetElement()
    /// and silently does nothing — the URL remains unchanged.
    /// </summary>
    [Fact]
    public void Set_Hyperlink_Url_ShouldUpdateTarget()
    {
        // Add a paragraph and a hyperlink
        _handler.Add("/", "paragraph", null, new() { ["text"] = "before" });
        _handler.Add("/body/p[1]", "hyperlink", null, new()
        {
            ["url"] = "https://example.com",
            ["text"] = "Click here"
        });

        // Verify initial state
        var node = _handler.Get("/body/p[1]/hyperlink[1]");
        node.Should().NotBeNull();
        node!.Format["link"].Should().Be("https://example.com/");

        // Set a new URL on the hyperlink node
        var unsupported = _handler.Set("/body/p[1]/hyperlink[1]", new()
        {
            ["url"] = "https://newsite.com"
        });

        // The "url" property should NOT be reported as unsupported
        unsupported.Should().NotContain(u => u.Contains("url"),
            "url should be a recognized property for hyperlink elements");

        // The hyperlink should now point to the new URL
        var updated = _handler.Get("/body/p[1]/hyperlink[1]");
        updated.Should().NotBeNull();
        ((string)updated!.Format["link"]).Should().Contain("newsite.com",
            "hyperlink URL should be updated after Set");
    }

    /// <summary>
    /// BUG-5b: Setting "text" on a hyperlink node should update the display text.
    /// Currently silently ignored because SetElement has no Hyperlink handler.
    /// </summary>
    [Fact]
    public void Set_Hyperlink_Text_ShouldUpdateDisplayText()
    {
        _handler.Add("/", "paragraph", null, new() { ["text"] = "before" });
        _handler.Add("/body/p[1]", "hyperlink", null, new()
        {
            ["url"] = "https://example.com",
            ["text"] = "Original"
        });

        var node = _handler.Get("/body/p[1]/hyperlink[1]");
        node!.Text.Should().Be("Original");

        // Set new text on the hyperlink
        var unsupported = _handler.Set("/body/p[1]/hyperlink[1]", new()
        {
            ["text"] = "Updated Link Text"
        });

        unsupported.Should().NotContain(u => u.Contains("text"),
            "text should be a recognized property for hyperlink elements");

        var updated = _handler.Get("/body/p[1]/hyperlink[1]");
        updated!.Text.Should().Be("Updated Link Text",
            "hyperlink display text should change after Set");
    }

    // ==================== BUG-6: Set link (change URL) on hyperlink node silently fails ====================

    /// <summary>
    /// BUG-6: Setting "link" on a hyperlink node to change its URL should update
    /// the hyperlink relationship. The "link" property is only handled in the
    /// `element is Run` branch (line ~861), not for Hyperlink elements directly.
    /// When you target /body/p[1]/hyperlink[1] with link=https://newurl.com,
    /// the property is silently ignored — the URL remains unchanged.
    /// This is the same root cause as Bug 5 (no Hyperlink branch in SetElement)
    /// but tests the "link" property key specifically (vs "url"/"text").
    /// </summary>
    [Fact]
    public void Set_Link_OnHyperlinkNode_ShouldUpdateUrl()
    {
        _handler.Add("/", "paragraph", null, new() { ["text"] = "prefix " });
        _handler.Add("/body/p[1]", "hyperlink", null, new()
        {
            ["url"] = "https://example.com",
            ["text"] = "linked text"
        });

        // Verify hyperlink exists with original URL
        var hlNode = _handler.Get("/body/p[1]/hyperlink[1]");
        hlNode.Should().NotBeNull();
        hlNode!.Text.Should().Be("linked text");
        ((string)hlNode.Format["link"]).Should().Contain("example.com");

        // Set a new URL using "link" property on the hyperlink node
        var unsupported = _handler.Set("/body/p[1]/hyperlink[1]", new()
        {
            ["link"] = "https://updated-url.com"
        });

        // "link" should be handled, not reported as unsupported
        unsupported.Should().NotContain(u => u.Contains("link"),
            "link should be a recognized property for hyperlink elements");

        // The hyperlink should now point to the new URL
        var updated = _handler.Get("/body/p[1]/hyperlink[1]");
        updated.Should().NotBeNull();
        ((string)updated!.Format["link"]).Should().Contain("updated-url.com",
            "hyperlink URL should change after Set with 'link' property");

        // Text should remain unchanged
        updated.Text.Should().Be("linked text",
            "hyperlink display text should not change when only URL is updated");
    }

    // ==================== BUG-8: Header + orientation causes schema error ====================

    /// <summary>
    /// BUG-8: After adding a header, setting orientation on the section inserts
    /// PageSize before HeaderReference, violating OOXML schema order.
    /// The schema requires: HeaderReference*, FooterReference*, ... SectionType,
    /// PageSize, PageMargin, ...
    /// EnsureSectPrPageSize() prepends PageSize when no SectionType exists,
    /// but doesn't account for HeaderReference already being in the section.
    /// This causes Word to report file corruption or silently ignore the page setup.
    /// </summary>
    [Fact]
    public void Set_Orientation_AfterAddHeader_ShouldMaintainSchemaOrder()
    {
        // Add a header first — this puts a HeaderReference in SectionProperties
        _handler.Add("/", "header", null, new()
        {
            ["text"] = "My Header",
            ["type"] = "default"
        });

        // Now set orientation — this calls EnsureSectPrPageSize which may
        // insert PageSize before the HeaderReference
        _handler.Set("/section[1]", new()
        {
            ["orientation"] = "landscape"
        });

        // Close handler so we can inspect raw XML
        _handler.Dispose();

        // Verify schema order in the XML: HeaderReference must come before PageSize
        using (var docPkg = DocumentFormat.OpenXml.Packaging.WordprocessingDocument
            .Open(_path, false))
        {
            var body = docPkg.MainDocumentPart?.Document?.Body;
            var sectPr = body?.Elements<SectionProperties>().LastOrDefault();
            sectPr.Should().NotBeNull();

            var children = sectPr!.ChildElements.ToList();
            var headerRefIndex = children.FindIndex(e => e is HeaderReference);
            var pageSizeIndex = children.FindIndex(e => e is PageSize);

            headerRefIndex.Should().BeGreaterThanOrEqualTo(0, "HeaderReference should exist");
            pageSizeIndex.Should().BeGreaterThanOrEqualTo(0, "PageSize should exist");

            headerRefIndex.Should().BeLessThan(pageSizeIndex,
                "HeaderReference must appear before PageSize in SectionProperties " +
                "per OOXML schema order. Currently PageSize is inserted before " +
                "HeaderReference, causing schema violation.");
        }

        // Reopen and verify orientation and header survive
        _handler = new WordHandler(_path, editable: true);

        var section = _handler.Get("/section[1]");
        section.Should().NotBeNull();
        ((string)section!.Format["orientation"]).Should().Be("landscape",
            "orientation should be landscape after Set");

        var header = _handler.Get("/header[1]");
        header.Should().NotBeNull("header should still exist after setting orientation");
    }

    /// <summary>
    /// BUG-8b: Same schema order issue but verified with both header and footer,
    /// then changing orientation. All references must precede PageSize.
    /// </summary>
    [Fact]
    public void Set_Orientation_AfterAddHeaderAndFooter_FileRemainsValid()
    {
        // Add header and footer
        _handler.Add("/", "header", null, new()
        {
            ["text"] = "Header Text"
        });
        _handler.Add("/", "footer", null, new()
        {
            ["text"] = "Footer Text"
        });

        // Set orientation
        _handler.Set("/section[1]", new()
        {
            ["orientation"] = "landscape"
        });

        // Close handler so we can inspect raw XML
        _handler.Dispose();

        // Verify schema order: HeaderReference < FooterReference < PageSize
        using (var docPkg = DocumentFormat.OpenXml.Packaging.WordprocessingDocument
            .Open(_path, false))
        {
            var sectPr = docPkg.MainDocumentPart?.Document?.Body?
                .Elements<SectionProperties>().LastOrDefault();
            sectPr.Should().NotBeNull();

            var children = sectPr!.ChildElements.ToList();
            var headerRefIdx = children.FindIndex(e => e is HeaderReference);
            var footerRefIdx = children.FindIndex(e => e is FooterReference);
            var pageSizeIdx = children.FindIndex(e => e is PageSize);

            headerRefIdx.Should().BeGreaterThanOrEqualTo(0);
            footerRefIdx.Should().BeGreaterThanOrEqualTo(0);
            pageSizeIdx.Should().BeGreaterThanOrEqualTo(0);

            headerRefIdx.Should().BeLessThan(pageSizeIdx,
                "HeaderReference must precede PageSize per OOXML schema");
            footerRefIdx.Should().BeLessThan(pageSizeIdx,
                "FooterReference must precede PageSize per OOXML schema");
        }

        // Reopen and verify all elements survive
        _handler = new WordHandler(_path, editable: true);

        var section = _handler.Get("/section[1]");
        section.Should().NotBeNull();
        ((string)section!.Format["orientation"]).Should().Be("landscape");

        var header = _handler.Get("/header[1]");
        header.Should().NotBeNull("header should survive orientation change");
        header!.Text.Should().Be("Header Text");

        var footer = _handler.Get("/footer[1]");
        footer.Should().NotBeNull("footer should survive orientation change");
        footer!.Text.Should().Be("Footer Text");
    }
}
