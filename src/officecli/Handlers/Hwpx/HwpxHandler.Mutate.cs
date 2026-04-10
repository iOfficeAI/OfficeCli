// File: src/officecli/Handlers/Hwpx/HwpxHandler.Mutate.cs
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Add ====================

    /// <summary>
    /// Add a new element under the parent at the given path.
    /// Returns the path of the newly created element.
    /// </summary>
    /// <param name="parentPath">Path to the parent element.</param>
    /// <param name="type">Element type: "paragraph", "table", "run" (lowercase).</param>
    /// <param name="index">Optional 1-based insertion index. null = append.</param>
    /// <param name="properties">Optional properties for the new element.</param>
    public string Add(string parentPath, string type, int? index,
                      Dictionary<string, string> properties)
    {
        var parent = ResolvePath(parentPath);

        var newElement = type.ToLowerInvariant() switch
        {
            "paragraph" or "p" => CreateParagraph(properties),
            "table" or "tbl"   => CreateTable(properties),
            "run"              => CreateRun(properties),
            "row" or "tr"      => CreateRow(parent, properties),
            "cell" or "tc"     => CreateCell(parent, properties),
            _ => throw new CliException($"Unsupported element type: {type}")
        };

        // Insert at index or append
        if (index.HasValue)
        {
            var targetName = newElement.Name;
            var siblings = parent.Elements(targetName).ToList();
            var insertIdx = index.Value - 1; // convert 1-based to 0-based

            if (insertIdx <= 0 || siblings.Count == 0)
            {
                // Insert as first child of this type
                var firstOfType = parent.Elements(targetName).FirstOrDefault();
                if (firstOfType != null)
                    firstOfType.AddBeforeSelf(newElement);
                else
                    parent.Add(newElement);
            }
            else if (insertIdx >= siblings.Count)
            {
                // Append after last sibling of this type
                siblings.Last().AddAfterSelf(newElement);
            }
            else
            {
                // Insert before the element currently at this index
                siblings[insertIdx].AddBeforeSelf(newElement);
            }
        }
        else
        {
            parent.Add(newElement);
        }

        _dirty = true;
        SaveSection(parent);
        return BuildPath(newElement);
    }

    // ==================== Element Factories ====================

    /// <summary>
    /// Create a new paragraph element with optional text content.
    /// Props: "text" → paragraph text, "styleidref" → style ID, "charpridref" → char property ID.
    /// </summary>
    private XElement CreateParagraph(Dictionary<string, string>? props)
    {
        var id = NewId();
        var text = props?.GetValueOrDefault("text") ?? "";
        var styleIdRef = props?.GetValueOrDefault("styleidref") ?? props?.GetValueOrDefault("styleIDRef") ?? "0";
        var charPrIdRef = props?.GetValueOrDefault("charpridref") ?? props?.GetValueOrDefault("charPrIDRef") ?? "0";
        var paraPrIdRef = props?.GetValueOrDefault("parapridref") ?? props?.GetValueOrDefault("paraPrIDRef") ?? "0";

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", id),
            new XAttribute("styleIDRef", styleIdRef),
            new XAttribute("paraPrIDRef", paraPrIdRef),
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", charPrIdRef),
                new XElement(HwpxNs.Hp + "t", text)
            )
        );
    }

    /// <summary>
    /// Create a new table element with full DOCX-parity features.
    ///
    /// CRITICAL: Every &lt;hp:tc&gt; MUST have ALL of the following children
    /// in this exact order, or Hancom will crash on open:
    ///   - &lt;hp:subList vertAlign="CENTER" ...&gt;&lt;hp:p .../&gt;&lt;/hp:subList&gt;
    ///   - &lt;hp:cellAddr colAddr="C" rowAddr="R"/&gt;
    ///   - &lt;hp:cellSpan colSpan="1" rowSpan="1"/&gt;
    ///   - &lt;hp:cellSz width="W" height="H"/&gt;
    ///   - &lt;hp:cellMargin left="510" right="510" top="141" bottom="141"/&gt;
    ///
    /// Props:
    ///   "rows"           → row count (default 2)
    ///   "cols"           → col count (default 2)
    ///   "width"          → total table width in HWPML units (default 42520 ≈ A4 body)
    ///   "data"           → cell data: "H1,H2;R1C1,R1C2" or CSV file path
    ///   "colWidths"      → per-column widths: "10000,15000,17520"
    ///   "merge"          → merge spec: "startRow,startCol,endRow,endCol;..."
    ///   "borderFillIDRef"→ table-level border fill ID (default "1")
    /// </summary>
    private XElement CreateTable(Dictionary<string, string>? props)
    {
        var id = NewId();

        // Parse data: "H1,H2;R1C1,R1C2;R2C1,R2C2" or CSV file path
        string[][]? tableData = null;
        if (props?.TryGetValue("data", out var dataStr) == true && !string.IsNullOrEmpty(dataStr))
        {
            if (File.Exists(dataStr))
                tableData = File.ReadAllLines(dataStr)
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(l => l.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
            else
                tableData = dataStr.Split(';')
                    .Select(r => r.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
        }

        // Determine dimensions
        int rows, cols;
        if (tableData != null)
        {
            rows = tableData.Length;
            cols = tableData.Max(r => r.Length);
            // Allow explicit overrides to be larger
            if (int.TryParse(props?.GetValueOrDefault("rows"), out var r2) && r2 > rows) rows = r2;
            if (int.TryParse(props?.GetValueOrDefault("cols"), out var c2) && c2 > cols) cols = c2;
        }
        else
        {
            rows = int.TryParse(props?.GetValueOrDefault("rows"), out var r) && r > 0 ? r : 2;
            cols = int.TryParse(props?.GetValueOrDefault("cols"), out var c) && c > 0 ? c : 2;
        }

        var totalWidth = int.TryParse(props?.GetValueOrDefault("width"), out var w) && w > 0 ? w : 42520;
        var defaultCellWidth = totalWidth / Math.Max(cols, 1);

        // Parse per-column widths: "10000,15000,17520"
        int[]? colWidthArr = null;
        if ((props?.TryGetValue("colwidths", out var cwStr) == true
            || props?.TryGetValue("colWidths", out cwStr) == true)
            && !string.IsNullOrEmpty(cwStr))
        {
            colWidthArr = cwStr.Split(',')
                .Select(s => int.TryParse(s.Trim(), out var v) ? v : defaultCellWidth)
                .ToArray();
        }

        var borderFillRef = props?.GetValueOrDefault("borderfillid")
            ?? props?.GetValueOrDefault("borderFillIDRef") ?? "1";

        var tbl = new XElement(HwpxNs.Hp + "tbl",
            new XAttribute("id", id),
            new XAttribute("colCnt", cols.ToString()),
            new XAttribute("rowCnt", rows.ToString()),
            new XAttribute("cellSpacing", "0"),
            new XAttribute("borderFillIDRef", borderFillRef)
        );

        // Column widths
        for (int col = 0; col < cols; col++)
        {
            var cw = colWidthArr != null && col < colWidthArr.Length ? colWidthArr[col] : defaultCellWidth;
            tbl.Add(new XElement(HwpxNs.Hp + "colSz",
                new XAttribute("width", cw.ToString())));
        }

        // Parse merge instructions: "0,0,0,3;1,0,2,0" (startRow,startCol,endRow,endCol)
        var mergedCells = new HashSet<(int row, int col)>();
        var spanMap = new Dictionary<(int row, int col), (int rowSpan, int colSpan)>();
        if (props?.TryGetValue("merge", out var mergeStr) == true && !string.IsNullOrEmpty(mergeStr))
        {
            foreach (var m in mergeStr.Split(';'))
            {
                var parts = m.Trim().Split(',');
                if (parts.Length == 4
                    && int.TryParse(parts[0].Trim(), out var sr)
                    && int.TryParse(parts[1].Trim(), out var sc)
                    && int.TryParse(parts[2].Trim(), out var er)
                    && int.TryParse(parts[3].Trim(), out var ec))
                {
                    for (int mr = sr; mr <= er; mr++)
                        for (int mc = sc; mc <= ec; mc++)
                        {
                            if (mr == sr && mc == sc)
                                spanMap[(mr, mc)] = (er - sr + 1, ec - sc + 1);
                            else
                                mergedCells.Add((mr, mc));
                        }
                }
            }
        }

        var cellHeight = 1000;

        // Build rows and cells
        for (int row = 0; row < rows; row++)
        {
            var tr = new XElement(HwpxNs.Hp + "tr");

            for (int col = 0; col < cols; col++)
            {
                // Skip cells covered by a merge
                if (mergedCells.Contains((row, col)))
                    continue;

                var cellId = NewId();
                var rowSpan = 1;
                var colSpan = 1;
                if (spanMap.TryGetValue((row, col), out var span))
                {
                    rowSpan = span.rowSpan;
                    colSpan = span.colSpan;
                }

                // Cell width = sum of spanned columns
                var cellW = 0;
                for (int ci = col; ci < col + colSpan && ci < cols; ci++)
                    cellW += colWidthArr != null && ci < colWidthArr.Length ? colWidthArr[ci] : defaultCellWidth;

                // Cell text from data prop or positional prop
                var cellText = "";
                if (tableData != null && row < tableData.Length && col < tableData[row].Length)
                    cellText = tableData[row][col];
                else if (props?.TryGetValue($"r{row + 1}c{col + 1}", out var rc) == true)
                    cellText = rc;

                // Per-cell borderFillIDRef: check "r{row+1}c{col+1}borderfillid" prop, fallback to table-level
                var cellBorderFill = props?.GetValueOrDefault($"r{row + 1}c{col + 1}borderfillid")
                    ?? props?.GetValueOrDefault($"r{row + 1}borderfillid")
                    ?? borderFillRef;

                // Per-cell vertAlign: check "r{row+1}c{col+1}valign" prop
                var cellVertAlign = props?.GetValueOrDefault($"r{row + 1}c{col + 1}valign")
                    ?? props?.GetValueOrDefault("valign") ?? "CENTER";

                var tc = BuildCell(row, col, rowSpan, colSpan, cellW, cellHeight,
                    cellText, cellBorderFill, isHeader: row == 0, vertAlign: cellVertAlign);

                tr.Add(tc);
            }

            tbl.Add(tr);
        }

        return tbl;
    }

    /// <summary>
    /// Create a new run element with optional text content.
    /// Props: "text" → run text, "charpridref" → char property ID.
    /// </summary>
    private XElement CreateRun(Dictionary<string, string>? props)
    {
        var text = props?.GetValueOrDefault("text") ?? "";
        var charPrIdRef = props?.GetValueOrDefault("charpridref") ?? props?.GetValueOrDefault("charPrIDRef") ?? "0";

        return new XElement(HwpxNs.Hp + "run",
            new XAttribute("charPrIDRef", charPrIdRef),
            new XElement(HwpxNs.Hp + "t", text)
        );
    }

    /// <summary>
    /// Create a new table row with cells. Parent MUST be a &lt;hp:tbl&gt;.
    /// Props: "cols" → cell count (default from table colCnt),
    ///        "c1", "c2", ... → cell text for each column.
    /// </summary>
    private XElement CreateRow(XElement parent, Dictionary<string, string>? props)
    {
        if (parent.Name.LocalName != "tbl")
            throw new CliException("Rows can only be added to a table element");

        var colCnt = int.TryParse(parent.Attribute("colCnt")?.Value, out var cc) ? cc : 1;
        var cols = int.TryParse(props?.GetValueOrDefault("cols"), out var c) && c > 0 ? c : colCnt;
        var existingRows = parent.Elements(HwpxNs.Hp + "tr").Count();

        // Get column widths from existing colSz elements
        var colSizes = parent.Elements(HwpxNs.Hp + "colSz")
            .Select(e => int.TryParse(e.Attribute("width")?.Value, out var w) ? w : 42520 / cols)
            .ToList();

        var tr = new XElement(HwpxNs.Hp + "tr");

        for (int col = 0; col < cols; col++)
        {
            var cellText = props?.GetValueOrDefault($"c{col + 1}") ?? "";
            var cellWidth = col < colSizes.Count ? colSizes[col] : 42520 / cols;

            tr.Add(BuildCell(existingRows, col, 1, 1, cellWidth, 1000, cellText, "1"));
        }

        // Update table rowCnt
        parent.SetAttributeValue("rowCnt", (existingRows + 1).ToString());

        return tr;
    }

    /// <summary>
    /// Create a new table cell. Parent MUST be a &lt;hp:tr&gt;.
    /// Props: "text" → cell text, "width" → cell width.
    /// </summary>
    private XElement CreateCell(XElement parent, Dictionary<string, string>? props)
    {
        if (parent.Name.LocalName != "tr")
            throw new CliException("Cells can only be added to a table row element");

        var existingCells = parent.Elements(HwpxNs.Hp + "tc").Count();
        var text = props?.GetValueOrDefault("text") ?? "";
        var cellWidth = int.TryParse(props?.GetValueOrDefault("width"), out var w) && w > 0 ? w : 10000;

        // Determine row address from parent's position in the table
        var tbl = parent.Parent;
        var rowAddr = tbl?.Elements(HwpxNs.Hp + "tr").ToList().IndexOf(parent) ?? 0;

        return BuildCell(rowAddr, existingCells, 1, 1, cellWidth, 1000, text, "1");
    }

    // ==================== Remove ====================

    /// <summary>
    /// Remove the element at the given path.
    /// Returns null on success. Throws CliException if not found.
    /// </summary>
    public string? Remove(string path)
    {
        var element = ResolvePath(path);

        // Capture parent for SaveSection before detaching
        var parent = element.Parent
            ?? throw new CliException($"Cannot remove root element at: {path}");

        element.Remove();
        _dirty = true;
        SaveSection(parent);
        return null;
    }

    // ==================== Move ====================

    /// <summary>
    /// Move an element from sourcePath to a new position under targetParentPath.
    ///
    /// CORRECT detach-then-insert pattern:
    /// 1. Resolve targetParentPath FIRST (validate before modifying the tree).
    /// 2. Resolve sourcePath.
    /// 3. Detach: call source.Remove() BEFORE re-inserting.
    ///    Bad pattern: target.Add(source) when source is still parented —
    ///    XLinq silently moves it, but only within the same XDocument.
    ///    Cross-section moves fail silently without detach.
    /// 4. Insert at the specified index under the target parent.
    /// </summary>
    /// <returns>New path of the moved element.</returns>
    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        if (string.IsNullOrEmpty(targetParentPath))
            throw new CliException("Target parent path is required for move");

        // 1. Resolve target FIRST (before tree modification)
        var target = ResolvePath(targetParentPath);

        // 2. Resolve source
        var source = ResolvePath(sourcePath);
        var sourceParent = source.Parent;

        // 3. Detach source — NEVER re-parent directly
        source.Remove();

        // 4. Insert at position
        if (index.HasValue)
        {
            var siblings = target.Elements(source.Name).ToList();
            var insertIdx = index.Value - 1;

            if (insertIdx <= 0 || siblings.Count == 0)
            {
                var firstOfType = target.Elements(source.Name).FirstOrDefault();
                if (firstOfType != null)
                    firstOfType.AddBeforeSelf(source);
                else
                    target.Add(source);
            }
            else if (insertIdx >= siblings.Count)
            {
                siblings.Last().AddAfterSelf(source);
            }
            else
            {
                siblings[insertIdx].AddBeforeSelf(source);
            }
        }
        else
        {
            target.Add(source);
        }

        _dirty = true;

        // Save both affected sections
        if (sourceParent != null)
            SaveSection(sourceParent);
        SaveSection(target);

        return BuildPath(source);
    }

    // ==================== CopyFrom ====================

    /// <summary>
    /// Deep-clone the element at sourcePath and insert the copy under targetParentPath.
    /// Assigns a new id attribute to the clone to avoid duplicate IDs.
    /// </summary>
    /// <returns>Path of the newly created copy.</returns>
    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var source = ResolvePath(sourcePath);
        var target = ResolvePath(targetParentPath);

        // Deep clone
        var clone = new XElement(source);

        // Assign new IDs to clone and all descendants with id attributes
        AssignNewIds(clone);

        // Insert at position
        if (index.HasValue)
        {
            var siblings = target.Elements(clone.Name).ToList();
            var insertIdx = index.Value - 1;

            if (insertIdx <= 0 || siblings.Count == 0)
            {
                var firstOfType = target.Elements(clone.Name).FirstOrDefault();
                if (firstOfType != null)
                    firstOfType.AddBeforeSelf(clone);
                else
                    target.Add(clone);
            }
            else if (insertIdx >= siblings.Count)
            {
                siblings.Last().AddAfterSelf(clone);
            }
            else
            {
                siblings[insertIdx].AddBeforeSelf(clone);
            }
        }
        else
        {
            target.Add(clone);
        }

        _dirty = true;
        SaveSection(target);
        return BuildPath(clone);
    }

    /// <summary>
    /// Recursively assign new IDs to an element and all descendants
    /// that have an "id" attribute. Prevents duplicate IDs in the document.
    /// </summary>
    private void AssignNewIds(XElement element)
    {
        if (element.Attribute("id") != null)
        {
            element.SetAttributeValue("id", NewId());
        }

        foreach (var child in element.Elements())
        {
            AssignNewIds(child);
        }
    }

    // ==================== Helpers ====================

    /// <summary>
    /// Build a Hancom-compatible <hp:tc> element with correct child ordering:
    /// subList → cellAddr → cellSpan → cellSz → cellMargin.
    /// This matches the structure produced by Hancom Office (2011 namespace).
    /// </summary>
    private XElement BuildCell(int rowAddr, int colAddr, int rowSpan, int colSpan,
                               int width, int height, string text,
                               string borderFillIDRef, bool isHeader = false,
                               string vertAlign = "CENTER")
    {
        return new XElement(HwpxNs.Hp + "tc",
            new XAttribute("name", ""),
            new XAttribute("header", isHeader ? "1" : "0"),
            new XAttribute("hasMargin", "0"),
            new XAttribute("protect", "0"),
            new XAttribute("editable", "0"),
            new XAttribute("dirty", "0"),
            new XAttribute("borderFillIDRef", borderFillIDRef),
            new XElement(HwpxNs.Hp + "subList",
                new XAttribute("id", ""),
                new XAttribute("textDirection", "HORIZONTAL"),
                new XAttribute("lineWrap", "BREAK"),
                new XAttribute("vertAlign", vertAlign),
                new XAttribute("linkListIDRef", "0"),
                new XAttribute("linkListNextIDRef", "0"),
                new XAttribute("textWidth", "0"),
                new XAttribute("textHeight", "0"),
                new XAttribute("hasTextRef", "0"),
                new XAttribute("hasNumRef", "0"),
                new XElement(HwpxNs.Hp + "p",
                    new XAttribute("id", NewId()),
                    new XAttribute("styleIDRef", "0"),
                    new XAttribute("paraPrIDRef", "0"),
                    new XElement(HwpxNs.Hp + "run",
                        new XAttribute("charPrIDRef", "0"),
                        new XElement(HwpxNs.Hp + "t", text)))),
            new XElement(HwpxNs.Hp + "cellAddr",
                new XAttribute("colAddr", colAddr.ToString()),
                new XAttribute("rowAddr", rowAddr.ToString())),
            new XElement(HwpxNs.Hp + "cellSpan",
                new XAttribute("colSpan", colSpan.ToString()),
                new XAttribute("rowSpan", rowSpan.ToString())),
            new XElement(HwpxNs.Hp + "cellSz",
                new XAttribute("width", width.ToString()),
                new XAttribute("height", height.ToString())),
            new XElement(HwpxNs.Hp + "cellMargin",
                new XAttribute("left", "510"),
                new XAttribute("right", "510"),
                new XAttribute("top", "141"),
                new XAttribute("bottom", "141"))
        );
    }

    /// <summary>
    /// Generate a unique ID string.
    /// Format: "p" + first 8 hex chars of a new GUID.
    /// Short enough for readability, collision-resistant for single-document scope.
    /// </summary>
    private string NewId()
    {
        return "p" + Guid.NewGuid().ToString("N")[..8];
    }
}
