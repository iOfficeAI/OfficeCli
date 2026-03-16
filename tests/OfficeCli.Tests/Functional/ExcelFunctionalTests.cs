// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for XLSX: each test creates a blank file, adds elements,
/// queries them, and modifies them — exercising the full Create→Add→Get→Set lifecycle.
/// </summary>
public class ExcelFunctionalTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelFunctionalTests()
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

    // Reopen the file to verify persistence
    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Sheet lifecycle ====================

    [Fact]
    public void BlankFile_HasSheet1()
    {
        var node = _handler.Get("/");
        node.Children.Should().Contain(c => c.Path == "/Sheet1");
    }

    [Fact]
    public void AddSheet_ReturnsPath()
    {
        var path = _handler.Add("/", "sheet", null,
            new Dictionary<string, string> { ["name"] = "Sales" });
        path.Should().Be("/Sales");
    }

    [Fact]
    public void AddSheet_Get_ReturnsSheetType()
    {
        _handler.Add("/", "sheet", null, new Dictionary<string, string> { ["name"] = "Report" });
        var node = _handler.Get("/Report");
        node.Type.Should().Be("sheet");
    }

    [Fact]
    public void AddSheet_Multiple_AllVisible()
    {
        _handler.Add("/", "sheet", null, new Dictionary<string, string> { ["name"] = "Alpha" });
        _handler.Add("/", "sheet", null, new Dictionary<string, string> { ["name"] = "Beta" });

        var root = _handler.Get("/");
        var sheetPaths = root.Children.Select(c => c.Path).ToList();
        sheetPaths.Should().Contain("/Alpha");
        sheetPaths.Should().Contain("/Beta");
    }

    // ==================== Cell lifecycle ====================

    [Fact]
    public void AddCell_NumberValue_TextIsReadBack()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "42" });

        var node = _handler.Get("/Sheet1/A1");
        node.Type.Should().Be("cell");
        node.Text.Should().Be("42");
    }

    [Fact]
    public void AddCell_StringValue_TextIsReadBack()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "B2", ["value"] = "Hello", ["type"] = "string" });

        var node = _handler.Get("/Sheet1/B2");
        node.Text.Should().Be("Hello");
    }

    [Fact]
    public void AddCell_Formula_FormulaIsReadBack()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "C1", ["formula"] = "A1+B1" });

        var node = _handler.Get("/Sheet1/C1");
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].Should().Be("A1+B1");
    }

    [Fact]
    public void AddCell_MultipleCells_AllReadBack()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "B1", ["value"] = "20" });
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "C1", ["value"] = "30" });

        _handler.Get("/Sheet1/A1").Text.Should().Be("10");
        _handler.Get("/Sheet1/B1").Text.Should().Be("20");
        _handler.Get("/Sheet1/C1").Text.Should().Be("30");
    }

    // ==================== Set: modify cell properties ====================

    [Fact]
    public void SetCell_Value_ValueIsUpdated()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "old" });

        _handler.Set("/Sheet1/A1", new Dictionary<string, string> { ["value"] = "new" });

        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("new");
    }

    [Fact]
    public void SetCell_Bold_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "100" });

        var act = () => _handler.Set("/Sheet1/A1",
            new Dictionary<string, string> { ["font.bold"] = "true" });
        act.Should().NotThrow();

        // Cell value should still be intact
        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("100");
    }

    [Fact]
    public void SetCell_Fill_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "styled" });

        var act = () => _handler.Set("/Sheet1/A1",
            new Dictionary<string, string> { ["fill"] = "4472C4" });
        act.Should().NotThrow();
    }

    [Fact]
    public void SetCell_NumFmt_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "0.5" });

        var act = () => _handler.Set("/Sheet1/A1",
            new Dictionary<string, string> { ["numFmt"] = "0.00%" });
        act.Should().NotThrow();
    }

    [Fact]
    public void SetCell_Alignment_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "text" });

        var act = () => _handler.Set("/Sheet1/A1",
            new Dictionary<string, string> { ["alignment.horizontal"] = "center" });
        act.Should().NotThrow();
    }

    // ==================== Query ====================

    [Fact]
    public void GetSheet_WithCells_ChildrenContainCells()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "1" });
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "B1", ["value"] = "2" });

        var sheet = _handler.Get("/Sheet1", depth: 2);
        var allCells = sheet.Children.SelectMany(row => row.Children).ToList();
        allCells.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void GetRange_ReturnsAllCellsInRange()
    {
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "1" });
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "B1", ["value"] = "2" });
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "C1", ["value"] = "3" });

        var range = _handler.Get("/Sheet1/A1:C1");
        range.Children.Should().HaveCount(3);
    }

    [Fact]
    public void GetRange_ChildrenHaveCorrectValues()
    {
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "B1", ["value"] = "20" });

        var range = _handler.Get("/Sheet1/A1:B1");
        var values = range.Children.Select(c => c.Text).ToList();
        values.Should().Contain("10");
        values.Should().Contain("20");
    }

    // ==================== Persistence ====================

    [Fact]
    public void AddCell_Persist_SurvivesReopenFile()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "persistent" });

        Reopen();
        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("persistent");
    }

    [Fact]
    public void AddSheet_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "sheet", null, new Dictionary<string, string> { ["name"] = "Saved" });

        Reopen();
        var root = _handler.Get("/");
        root.Children.Should().Contain(c => c.Path == "/Saved");
    }

    // ==================== Row lifecycle ====================

    [Fact]
    public void AddRow_RowIsQueryable()
    {
        _handler.Add("/Sheet1", "row", null,
            new Dictionary<string, string> { ["cols"] = "3" });

        var sheet = _handler.Get("/Sheet1", depth: 1);
        sheet.Children.Should().HaveCountGreaterThanOrEqualTo(1);
        sheet.Children.Should().Contain(c => c.Type == "row");
    }

    // ==================== XLSX Hyperlinks ====================

    [Fact]
    public void CellLink_Lifecycle()
    {
        // 1. Set cell value + link
        _handler.Set("/Sheet1/A1", new Dictionary<string, string>
        {
            ["value"] = "Visit us",
            ["link"] = "https://first.com"
        });

        // 2. Get + Verify
        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("Visit us");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://first.com");

        // 3. Set updated link + Verify
        _handler.Set("/Sheet1/A1", new Dictionary<string, string> { ["link"] = "https://updated.com" });
        node = _handler.Get("/Sheet1/A1");
        ((string)node.Format["link"]).Should().StartWith("https://updated.com");

        // 4. Remove link + Verify
        _handler.Set("/Sheet1/A1", new Dictionary<string, string> { ["link"] = "none" });
        node = _handler.Get("/Sheet1/A1");
        node.Format.Should().NotContainKey("link");
    }

    [Fact]
    public void CellLink_Persist_SurvivesReopenFile()
    {
        _handler.Set("/Sheet1/B1", new Dictionary<string, string>
        {
            ["value"] = "Link cell",
            ["link"] = "https://original.com"
        });
        _handler.Set("/Sheet1/B1", new Dictionary<string, string> { ["link"] = "https://persist.com" });

        var handler2 = Reopen();
        var node = handler2.Get("/Sheet1/B1");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://persist.com");
    }
}
