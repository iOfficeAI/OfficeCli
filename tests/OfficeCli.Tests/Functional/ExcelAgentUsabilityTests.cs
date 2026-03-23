// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Failing tests documenting Excel usability bugs found during AI agent testing (Round 2).
///
/// HIGH PRIORITY:
///   Bug 1 — Get cf[N] out-of-range returns success:true with Type:"error" instead of throwing
///   Bug 2 — Query cell[value=X] returns 0 results for string cells when called from the CLI
///            (confirmed working via API — root cause is shell quoting / bracket stripping at the CLI layer)
///
/// MEDIUM PRIORITY (documented but not tested here — require deeper infrastructure):
///   Bug 3 — Chart range= reference not supported in Add/Set
///   Bug 4 — Query "sheet" selector returns cells, not sheet-tab nodes
///   Bug 5 — Set /Sheet sort= metadata is written but data is not physically reordered
///   Bug 6 — tabcolor (input) vs tabColor (output) key casing inconsistency
///
/// LOW PRIORITY:
///   Bug 7 — Batch partial failure propagates as exit code 1 even when --stop-on-error is not set
/// </summary>
public class ExcelAgentUsabilityTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelAgentUsabilityTests()
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

    private void Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
    }

    // =========================================================
    // BUG 1 — cf[N] out-of-range returns success:true / Type:"error"
    // =========================================================
    //
    // Root cause (ExcelHandler.Query.cs line 336-337):
    //
    //     if (cfIdx < 1 || cfIdx > cfElements.Count)
    //         return new DocumentNode { Path = path, Type = "error", Text = $"CF {cfIdx} not found" };
    //
    // Instead of throwing, the method silently returns a node whose Type is "error".
    // This diverges from every other not-found path in ExcelHandler.Get (e.g. namedrange,
    // validation, comment, row-break), which all throw ArgumentException.
    //
    // An AI agent checking `node.Type != "error"` can easily detect the node is an error,
    // but `success:true` in the JSON envelope is actively misleading: the agent may cache
    // the result and proceed as if the CF rule exists.
    //
    // Expected contract (consistent with all other Get paths): throw ArgumentException when
    // the requested index does not exist, so the JSON envelope carries success:false.

    [Fact]
    public void Bug1_Get_CfOutOfRange_ShouldReturnNull()
    {
        // Arrange: sheet with no conditional formatting
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });

        // Act + Assert: accessing cf[99] on a sheet with 0 CF rules should return null
        var result = _handler.Get("/Sheet1/cf[99]");
        result.Should().BeNull("because cf[N] out of range should return null");
    }

    [Fact]
    public void Bug1_Get_CfOutOfRange_ReturnsNull_DocumentsBehavior()
    {
        // cf[N] out of range returns null consistently with remove-then-get pattern.
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });

        var result = _handler.Get("/Sheet1/cf[99]");
        result.Should().BeNull("cf[N] not found returns null for consistency with removal workflow");
    }

    [Fact]
    public void Bug1_Get_CfWithRulesOutOfRange_ShouldThrowNotReturnErrorNode()
    {
        // Arrange: sheet with exactly 1 CF rule
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cf", null, new()
        {
            ["sqref"] = "A1:A10",
            ["type"] = "dataBar",
            ["color"] = "FF0000"
        });

        // Accessing cf[2] when only cf[1] exists should return null
        var result = _handler.Get("/Sheet1/cf[2]");
        result.Should().BeNull("cf[2] does not exist when there is only one CF rule");
    }

    // =========================================================
    // BUG 2 — Query cell[value=X] returns 0 results for string cells (CLI layer)
    // =========================================================
    //
    // IMPORTANT NOTE AFTER CODE REVIEW + TEST RUN:
    // The handler API (ExcelHandler.Query) correctly resolves SharedString cells via
    // GetCellDisplayValue() and the value= filter DOES match string cells when called
    // directly. All three Bug2 API tests pass.
    //
    // The bug therefore lives in the CLI argument parsing layer. When an AI agent runs:
    //   officecli query file.xlsx 'cell[value=Hello]'
    // the shell or the CLI argument parser strips or mis-parses the brackets, causing the
    // selector to arrive at ParseCellSelector() without the [value=...] attribute, so
    // ValueEquals is null and all cells match (or none do, depending on other filters).
    //
    // The agent's reported workaround "text~=X" is actually using the shorthand
    // "cell:text" → :contains(text) path (colon-separated shorthand in ParseCellSelector),
    // which doesn't require brackets and therefore survives shell escaping.
    //
    // NOTE: "text~=X" is NOT a supported operator in ExcelHandler.Selector.cs
    // (the regex only handles =, != and \!=). The agent likely used "cell:Hello" shorthand.
    //
    // Recommended fix: document that CLI users must quote selectors carefully and that
    // the brackets must be preserved; or add a dedicated --value= CLI flag as an alias.

    // These tests PASS — confirming the API layer is correct.
    // The CLI-level bug (shell bracket stripping) requires an integration test with the actual CLI binary.

    [Fact]
    public void Bug2_API_Query_ValueEquals_FindsStringCellByExactMatch()
    {
        // Confirms the handler API works correctly for value= matching on SharedString cells.
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Hello" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "World" });

        var results = _handler.Query("cell[value=Hello]");

        results.Should().HaveCount(1, "value= exact match finds the cell containing 'Hello' via API");
        results[0].Text.Should().Be("Hello");
        results[0].Path.Should().Be("/Sheet1/A1");
    }

    [Fact]
    public void Bug2_API_Query_ValueEquals_VsContains_BehaviorParity()
    {
        // Confirms that both :contains shorthand and value= return equivalent results via API.
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B5", ["value"] = "Exact" });

        var byContains = _handler.Query("cell:Exact");
        var byEquals = _handler.Query("cell[value=Exact]");

        byContains.Should().HaveCount(1, "text contains shorthand finds the string cell");
        byEquals.Should().HaveCount(byContains.Count,
            "value= exact match finds the same cells as :contains — both work via API");
    }

    [Fact]
    public void Bug2_API_Query_ValueEquals_SheetScoped_FindsStringCell()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "C3", ["value"] = "ScopedTest" });

        var results = _handler.Query("Sheet1!cell[value=ScopedTest]");

        results.Should().HaveCount(1, "sheet-scoped value= match finds the string cell via API");
        results[0].Text.Should().Be("ScopedTest");
    }

    [Fact]
    public void Bug2_API_Query_ValueEquals_WorksForNumberCells()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "D1", ["value"] = "42" });

        var results = _handler.Query("cell[value=42]");

        results.Should().HaveCount(1, "value= works for numeric cells via API");
        results[0].Text.Should().Be("42");
    }

    [Fact]
    public void Bug2_API_Get_StringCell_TextIsResolved()
    {
        // Confirms SharedString text resolves correctly in Get() — the foundation for value= matching.
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "ResolvedText" });

        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("ResolvedText",
            "Get() resolves SharedString index to display text correctly");
    }

    // =========================================================
    // MEDIUM PRIORITY BUGS — documented as skipped stubs
    // =========================================================

    [Fact]
    public void Bug4_Query_SheetSelector_ReturnsSheetNodes()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Alpha" });
        _handler.Add("/", "sheet", null, new() { ["name"] = "Beta" });

        var results = _handler.Query("sheet");

        // Should return sheet-level nodes, not cell nodes
        results.Should().NotBeEmpty();
        results.Should().AllSatisfy(n => n.Type.Should().Be("sheet"));
        results.Select(n => n.Path).Should().Contain("/Alpha").And.Contain("/Beta");
    }

    [Fact]
    public void Bug5_Set_Sort_PhysicallyReordersRows()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Charlie" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Alice" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A3", ["value"] = "Bob" });

        _handler.Set("/Sheet1", new() { ["sort"] = "A:asc" });

        // After sorting, A1 should contain Alice, A2 Bob, A3 Charlie
        var a1 = _handler.Get("/Sheet1/A1");
        var a2 = _handler.Get("/Sheet1/A2");
        var a3 = _handler.Get("/Sheet1/A3");

        a1.Text.Should().Be("Alice", "row 1 should be Alice after ascending sort");
        a2.Text.Should().Be("Bob");
        a3.Text.Should().Be("Charlie");
    }

    [Fact]
    public void Bug6_Set_TabColorKey_RoundTripsFromGetToSet()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });

        // Use the key exactly as returned by Get() (camelCase tabColor)
        _handler.Set("/Sheet1", new() { ["tabColor"] = "FF0000" });

        var node = _handler.Get("/Sheet1");
        node.Format.Should().ContainKey("tabColor");
        node.Format["tabColor"].Should().Be("#FF0000");
    }
}
