// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for Excel enhancement Phase 1-3:
/// Subscript/Superscript, Hyperlink Add, Cell/Sheet Protection,
/// Array Formulas, Print Settings, Page Breaks, Rich Text Runs,
/// New CF types, Sorting
/// </summary>
public class ExcelPhase1Tests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelPhase1Tests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen() { _handler.Dispose(); _handler = new ExcelHandler(_path, editable: true); }

    // ==================== 1.1 Subscript / Superscript ====================

    [Fact]
    public void Set_Superscript_GetReturnsFontSuperscript()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "H2O" });
        _handler.Set("/Sheet1/A1", new() { ["superscript"] = "true" });

        var node = _handler.Get("/Sheet1/A1", 0);
        // Both canonical (font.superscript) and shorthand (superscript) should be present
        node.Format["font.superscript"].Should().Be(true);
        node.Format["superscript"].Should().Be(true);
    }

    [Fact]
    public void Set_Subscript_GetReturnsSubscript()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "H2O" });
        _handler.Set("/Sheet1/A1", new() { ["subscript"] = "true" });

        var node = _handler.Get("/Sheet1/A1", 0);
        node.Format["font.subscript"].Should().Be(true);
        node.Format["subscript"].Should().Be(true);
    }

    [Fact]
    public void Set_Superscript_PersistsAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "x2", ["superscript"] = "true" });

        Reopen();

        var node = _handler.Get("/Sheet1/A1", 0);
        node.Format["superscript"].Should().Be(true);
    }

    [Fact]
    public void Set_Superscript_False_RemovesSuperscript()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "x2", ["superscript"] = "true" });
        _handler.Set("/Sheet1/A1", new() { ["superscript"] = "false" });

        var node = _handler.Get("/Sheet1/A1", 0);
        node.Format.Should().NotContainKey("superscript");
    }

    // ==================== 1.2 Hyperlink Add ====================

    [Fact]
    public void Add_Cell_WithLink_GetReturnsLink()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Click", ["link"] = "https://example.com" });

        var node = _handler.Get("/Sheet1/A1", 0);
        node.Format.Should().ContainKey("link");
        node.Format["link"].ToString().Should().StartWith("https://example.com");
    }

    [Fact]
    public void Add_Cell_WithLink_PersistsAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Click", ["link"] = "https://example.com" });

        Reopen();

        var node = _handler.Get("/Sheet1/A1", 0);
        node.Format.Should().ContainKey("link");
    }

    // ==================== 1.3 Cell / Sheet Protection ====================

    [Fact]
    public void Set_Cell_Locked_GetReturnsLocked()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Secret" });
        _handler.Set("/Sheet1/A1", new() { ["locked"] = "false" });

        var node = _handler.Get("/Sheet1/A1", 0);
        node.Format["locked"].Should().Be(false);
    }

    [Fact]
    public void Set_Cell_FormulaHidden_GetReturnsFormulaHidden()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Secret" });
        _handler.Set("/Sheet1/A1", new() { ["formulahidden"] = "true" });

        var node = _handler.Get("/Sheet1/A1", 0);
        node.Format["formulahidden"].Should().Be(true);
    }

    [Fact]
    public void Set_SheetProtect_GetReturnsProtect()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Set("/Sheet1", new() { ["protect"] = "true" });

        var node = _handler.Get("/Sheet1", 0);
        node.Format["protect"].Should().Be(true);
    }

    [Fact]
    public void Set_Cell_Locked_PersistsAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Data" });
        _handler.Set("/Sheet1/A1", new() { ["locked"] = "false" });
        _handler.Set("/Sheet1", new() { ["protect"] = "true" });

        Reopen();

        var cellNode = _handler.Get("/Sheet1/A1", 0);
        cellNode.Format["locked"].Should().Be(false);
        var sheetNode = _handler.Get("/Sheet1", 0);
        sheetNode.Format["protect"].Should().Be(true);
    }

    // ==================== 1.4 Array Formulas ====================

    [Fact]
    public void Add_Cell_WithArrayFormula_GetReturnsArrayFormula()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "2" });
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "B1", ["arrayformula"] = "A1:A2*2"
        });

        var node = _handler.Get("/Sheet1/B1", 0);
        node.Format.Should().ContainKey("arrayformula");
        node.Format["arrayformula"].Should().Be(true);
    }

    [Fact]
    public void Set_Cell_ArrayFormula_PersistsAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _handler.Set("/Sheet1/A1", new() { ["arrayformula"] = "SUM(B1:B10)" });

        Reopen();

        var node = _handler.Get("/Sheet1/A1", 0);
        node.Format["arrayformula"].Should().Be(true);
    }

    // ==================== 2.1 Print Settings ====================

    [Fact]
    public void Set_PrintArea_GetReturnsRange()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Set("/Sheet1", new() { ["printArea"] = "$A$1:$D$10" });

        var node = _handler.Get("/Sheet1", 0);
        node.Format.Should().ContainKey("printArea");
        node.Format["printArea"].ToString().Should().Contain("$A$1:$D$10");
    }

    [Fact]
    public void Set_Orientation_Landscape_PersistsAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Set("/Sheet1", new() { ["orientation"] = "landscape" });

        Reopen();

        var node = _handler.Get("/Sheet1", 0);
        node.Format["orientation"].Should().Be("landscape");
    }

    [Fact]
    public void Set_FitToPage_GetReturnsFitToPage()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Set("/Sheet1", new() { ["fitToPage"] = "1x2" });

        var node = _handler.Get("/Sheet1", 0);
        node.Format["fitToPage"].Should().Be("1x2");
    }

    [Fact]
    public void Set_HeaderFooter_GetReturnsText()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Set("/Sheet1", new() { ["header"] = "&CPage &P", ["footer"] = "&LConfidential" });

        var node = _handler.Get("/Sheet1", 0);
        node.Format["header"].Should().Be("&CPage &P");
        node.Format["footer"].Should().Be("&LConfidential");
    }

    // ==================== 2.2 Page Breaks ====================

    [Fact]
    public void Add_RowPageBreak_GetReturnsBreak()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "data" });
        _handler.Add("/Sheet1", "pagebreak", null, new() { ["row"] = "5" });

        var node = _handler.Get("/Sheet1", 0);
        node.Format.Should().ContainKey("rowBreaks");
        node.Format["rowBreaks"].Should().Be("5");
    }

    [Fact]
    public void Add_ColPageBreak_PersistsAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "colbreak", null, new() { ["col"] = "C" });

        Reopen();

        var node = _handler.Get("/Sheet1", 0);
        node.Format.Should().ContainKey("colBreaks");
        node.Format["colBreaks"].Should().Be("3");
    }

    // ==================== 2.3 Rich Text Runs ====================

    [Fact]
    public void Add_SingleRun_CreatesRichTextCell()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "" });
        var runPath = _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "Hello", ["bold"] = "true" });

        runPath.Should().Be("/Sheet1/A1/run[1]");

        var node = _handler.Get("/Sheet1/A1", 1);
        node.Children.Should().HaveCountGreaterOrEqualTo(1);
        node.Children[0].Text.Should().Be("Hello");
        node.Children[0].Format["bold"].Should().Be(true);
    }

    [Fact]
    public void Add_MultipleRuns_DifferentFormats()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "" });
        _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "Normal " });
        _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "Bold", ["bold"] = "true", ["color"] = "FF0000" });

        var node = _handler.Get("/Sheet1/A1", 1);
        node.Children.Should().HaveCount(2);
        node.Children[0].Text.Should().Be("Normal ");
        node.Children[1].Text.Should().Be("Bold");
        node.Children[1].Format["bold"].Should().Be(true);
        node.Children[1].Format["color"].Should().Be("#FF0000");
    }

    [Fact]
    public void Set_RunBold_ModifiesExistingRun()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "" });
        _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "Test" });

        _handler.Set("/Sheet1/A1/run[1]", new() { ["bold"] = "true", ["color"] = "0000FF" });

        var node = _handler.Get("/Sheet1/A1", 1);
        node.Children[0].Format["bold"].Should().Be(true);
        node.Children[0].Format["color"].Should().Be("#0000FF");
    }

    [Fact]
    public void Add_Runs_PersistAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "" });
        _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "Hello ", ["bold"] = "true" });
        _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "World", ["italic"] = "true" });

        Reopen();

        var node = _handler.Get("/Sheet1/A1", 1);
        node.Children.Should().HaveCount(2);
        node.Children[0].Text.Should().Be("Hello ");
        node.Children[0].Format["bold"].Should().Be(true);
        node.Children[1].Text.Should().Be("World");
        node.Children[1].Format["italic"].Should().Be(true);
    }

    [Fact]
    public void Add_Run_WithSuperscript()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "" });
        _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "x" });
        _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "2", ["superscript"] = "true" });

        var node = _handler.Get("/Sheet1/A1", 1);
        node.Children.Should().HaveCount(2);
        node.Children[1].Format["superscript"].Should().Be(true);
    }

    // ==================== 3.1 Conditional Formatting New Types ====================

    [Fact]
    public void Add_CF_TopN_GetReturnsCFType()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        for (int i = 1; i <= 5; i++)
            _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = $"A{i}", ["value"] = (i * 10).ToString() });

        _handler.Add("/Sheet1", "topn", null, new() { ["sqref"] = "A1:A5", ["rank"] = "3", ["fill"] = "00FF00" });

        var node = _handler.Get("/Sheet1/cf[1]", 0);
        node.Format["cfType"].Should().Be("topN");
        ((uint)node.Format["rank"]).Should().Be(3);
    }

    [Fact]
    public void Add_CF_AboveAverage_PersistsAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "100" });
        _handler.Add("/Sheet1", "aboveaverage", null, new() { ["sqref"] = "A1:A10", ["fill"] = "FFFF00" });

        Reopen();

        var node = _handler.Get("/Sheet1/cf[1]", 0);
        node.Format["cfType"].Should().Be("aboveAverage");
    }

    [Fact]
    public void Add_CF_DuplicateValues_Works()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "hello" });
        _handler.Add("/Sheet1", "duplicatevalues", null, new() { ["sqref"] = "A1:A10", ["fill"] = "FF0000" });

        var node = _handler.Get("/Sheet1/cf[1]", 0);
        node.Format["cfType"].Should().Be("duplicateValues");
    }

    [Fact]
    public void Add_CF_UniqueValues_Works()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "hello" });
        _handler.Add("/Sheet1", "uniquevalues", null, new() { ["sqref"] = "A1:A10", ["fill"] = "00FF00" });

        var node = _handler.Get("/Sheet1/cf[1]", 0);
        node.Format["cfType"].Should().Be("uniqueValues");
    }

    [Fact]
    public void Add_CF_ContainsText_Works()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "hello world" });
        _handler.Add("/Sheet1", "containstext", null, new() { ["sqref"] = "A1:A10", ["text"] = "hello", ["fill"] = "0000FF" });

        var node = _handler.Get("/Sheet1/cf[1]", 0);
        node.Format["cfType"].Should().Be("containsText");
        node.Format["text"].Should().Be("hello");
    }

    [Fact]
    public void Add_CF_DateOccurring_Works()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "2024-01-01" });
        _handler.Add("/Sheet1", "dateoccurring", null, new() { ["sqref"] = "A1:A10", ["period"] = "thismonth" });

        var node = _handler.Get("/Sheet1/cf[1]", 0);
        node.Format["cfType"].Should().Be("timePeriod");
        node.Format["period"].Should().Be("thisMonth");
    }

    // ==================== 3.2 Sorting ====================

    [Fact]
    public void Set_Sort_SingleColumn_ActuallySortsData()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "3" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A3", ["value"] = "2" });

        _handler.Set("/Sheet1", new() { ["sort"] = "A:asc" });

        // Data should actually be reordered
        _handler.Get("/Sheet1/A1", 0).Text.Should().Be("1");
        _handler.Get("/Sheet1/A2", 0).Text.Should().Be("2");
        _handler.Get("/Sheet1/A3", 0).Text.Should().Be("3");

        // Sort state metadata should also be present
        var node = _handler.Get("/Sheet1", 0);
        node.Format["sort"].Should().Be("A:asc");
    }

    [Fact]
    public void Set_Sort_MultiColumn_PersistsAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "a" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "1" });

        _handler.Set("/Sheet1", new() { ["sort"] = "A:asc,B:desc" });

        Reopen();

        var node = _handler.Get("/Sheet1", 0);
        node.Format["sort"].Should().Be("A:asc,B:desc");
    }

    // ==================== Agent Round-Trip: Get output keys work as Set input ====================

    [Fact]
    public void Agent_RoundTrip_GetOutputKeysWorkAsSetInput()
    {
        // An agent should be able to: Get → read Format dict → pass keys back to Set
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test", ["bold"] = "true", ["superscript"] = "true" });

        var node = _handler.Get("/Sheet1/A1", 0);

        // Extract keys from Get output and use them as Set input on another cell
        var propsFromGet = new Dictionary<string, string>();
        if (node.Format.ContainsKey("bold")) propsFromGet["bold"] = "true";
        if (node.Format.ContainsKey("superscript")) propsFromGet["superscript"] = "true";
        propsFromGet["value"] = "copy";

        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1" });
        _handler.Set("/Sheet1/B1", propsFromGet);

        var copy = _handler.Get("/Sheet1/B1", 0);
        copy.Format["bold"].Should().Be(true);
        copy.Format["superscript"].Should().Be(true);
        copy.Text.Should().Be("copy");
    }

    // ==================== CF Dispatcher: Add("cf", { type: "topn" }) works ====================

    [Fact]
    public void Add_CF_Via_Dispatcher_TopN()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "100" });

        // Use unified "cf" type with "type" property — agent-friendly pattern
        _handler.Add("/Sheet1", "cf", null, new() { ["type"] = "topn", ["sqref"] = "A1:A10", ["rank"] = "5" });

        var node = _handler.Get("/Sheet1/cf[1]", 0);
        node.Format["cfType"].Should().Be("topN");
    }

    [Fact]
    public void Remove_PageBreak_ByPath()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "pagebreak", null, new() { ["row"] = "5" });
        _handler.Add("/Sheet1", "pagebreak", null, new() { ["row"] = "10" });

        // Get by path
        var pb = _handler.Get("/Sheet1/rowbreak[1]", 0);
        ((uint)pb.Format["row"]).Should().Be(5u);

        // Remove first break
        _handler.Remove("/Sheet1/rowbreak[1]");

        var node = _handler.Get("/Sheet1", 0);
        node.Format["rowBreaks"].Should().Be("10");
    }

    [Fact]
    public void RichText_Cell_HasRichtextFlag()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "" });
        _handler.Add("/Sheet1/A1", "run", null, new() { ["text"] = "Hello", ["bold"] = "true" });

        var node = _handler.Get("/Sheet1/A1", 1);
        node.Format["richtext"].Should().Be(true);
    }

    [Fact]
    public void Locked_AlwaysOutputsState()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "data", ["locked"] = "true" });

        var node = _handler.Get("/Sheet1/A1", 0);
        // Should output locked=true explicitly (not omit it)
        node.Format.Should().ContainKey("locked");
        node.Format["locked"].Should().Be(true);
    }

    [Fact]
    public void Add_CF_Via_Dispatcher_ContainsText()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "hello" });

        _handler.Add("/Sheet1", "cf", null, new() { ["type"] = "containstext", ["sqref"] = "A1:A10", ["text"] = "hello" });

        var node = _handler.Get("/Sheet1/cf[1]", 0);
        node.Format["cfType"].Should().Be("containsText");
    }
}
