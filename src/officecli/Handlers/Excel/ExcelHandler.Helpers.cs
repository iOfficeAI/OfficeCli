// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Private Helpers ====================

    private static Worksheet GetSheet(WorksheetPart part) =>
        part.Worksheet ?? throw new InvalidOperationException("Corrupt file: worksheet data missing");

    private Workbook GetWorkbook() =>
        _doc.WorkbookPart?.Workbook ?? throw new InvalidOperationException("Corrupt file: workbook missing");

    private List<(string Name, WorksheetPart Part)> GetWorksheets()
    {
        var result = new List<(string, WorksheetPart)>();
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return result;

        var sheets = workbook.GetFirstChild<Sheets>();
        if (sheets == null) return result;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var id = sheet.Id?.Value;
            if (id == null) continue;
            var part = (WorksheetPart)_doc.WorkbookPart!.GetPartById(id);
            result.Add((name, part));
        }

        return result;
    }

    private WorksheetPart? FindWorksheet(string sheetName)
    {
        foreach (var (name, part) in GetWorksheets())
        {
            if (name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                return part;
        }
        return null;
    }

    private string GetCellDisplayValue(Cell cell)
    {
        var value = cell.CellValue?.Text ?? "";

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sst?.SharedStringTable != null && int.TryParse(value, out int idx))
            {
                var item = sst.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                return item?.InnerText ?? value;
            }
        }

        // Formula cells without cached value: show the formula
        if (string.IsNullOrEmpty(value) && cell.CellFormula != null
            && !string.IsNullOrEmpty(cell.CellFormula.Text))
        {
            return $"={cell.CellFormula.Text}";
        }

        return value;
    }

    private List<DocumentNode> GetSheetChildNodes(string sheetName, SheetData sheetData, int depth)
    {
        var children = new List<DocumentNode>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = row.RowIndex?.Value ?? 0;
            var rowNode = new DocumentNode
            {
                Path = $"/{sheetName}/row[{rowIdx}]",
                Type = "row",
                ChildCount = row.Elements<Cell>().Count()
            };

            if (depth > 0)
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    rowNode.Children.Add(CellToNode(sheetName, cell));
                }
            }

            children.Add(rowNode);
        }
        return children;
    }

    private DocumentNode CellToNode(string sheetName, Cell cell, WorksheetPart? part = null)
    {
        var cellRef = cell.CellReference?.Value ?? "?";
        var value = GetCellDisplayValue(cell);
        var formula = cell.CellFormula?.Text;
        var type = cell.DataType?.Value.ToString() ?? "Number";

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{cellRef}",
            Type = "cell",
            Text = value,
            Preview = cellRef
        };

        node.Format["type"] = type;
        if (formula != null) node.Format["formula"] = formula;
        if (string.IsNullOrEmpty(value)) node.Format["empty"] = true;

        // Hyperlink readback
        if (part != null)
        {
            var hyperlink = GetSheet(part).GetFirstChild<Hyperlinks>()?.Elements<Hyperlink>()
                .FirstOrDefault(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
            if (hyperlink?.Id?.Value != null)
            {
                try
                {
                    var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id.Value);
                    if (rel != null) node.Format["link"] = rel.Uri.ToString();
                }
                catch { }
            }
        }

        return node;
    }

    private DocumentNode GetCellRange(string sheetName, SheetData sheetData, string range, int depth)
    {
        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException($"Invalid range: {range}");

        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{range}",
            Type = "range",
            Preview = range
        };

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = row.RowIndex?.Value ?? 0;
            if (rowIdx < startRow || rowIdx > endRow) continue;

            foreach (var cell in row.Elements<Cell>())
            {
                var (colName, _) = ParseCellReference(cell.CellReference?.Value ?? "A1");
                var colIdx = ColumnNameToIndex(colName);
                if (colIdx < ColumnNameToIndex(startCol) || colIdx > ColumnNameToIndex(endCol)) continue;

                node.Children.Add(CellToNode(sheetName, cell));
            }
        }

        node.ChildCount = node.Children.Count;
        return node;
    }

    private static Cell? FindCell(SheetData sheetData, string cellRef)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                    return cell;
            }
        }
        return null;
    }

    private static Cell FindOrCreateCell(SheetData sheetData, string cellRef)
    {
        var (colName, rowIdx) = ParseCellReference(cellRef);

        // Find or create row
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
        if (row == null)
        {
            row = new Row { RowIndex = (uint)rowIdx };
            // Insert in order
            var after = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < rowIdx);
            if (after != null)
                after.InsertAfterSelf(row);
            else
                sheetData.InsertAt(row, 0);
        }

        // Find or create cell
        var cell = row.Elements<Cell>().FirstOrDefault(c =>
            c.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellRef.ToUpperInvariant() };
            // Insert in column order
            var afterCell = row.Elements<Cell>().LastOrDefault(c =>
            {
                var (cn, _) = ParseCellReference(c.CellReference?.Value ?? "A1");
                return ColumnNameToIndex(cn) < ColumnNameToIndex(colName);
            });
            if (afterCell != null)
                afterCell.InsertAfterSelf(cell);
            else
                row.InsertAt(cell, 0);
        }

        return cell;
    }
}
