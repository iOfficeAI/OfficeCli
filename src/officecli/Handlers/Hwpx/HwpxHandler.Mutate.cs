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
    /// Create a new table element.
    ///
    /// CRITICAL: Every &lt;hp:tc&gt; MUST have ALL of the following children
    /// or Hancom will crash on open:
    ///   - &lt;hp:cellAddr rowAddr="R" colAddr="C" rowSpan="1" colSpan="1"/&gt;
    ///   - &lt;hp:cellSz width="W" height="H"/&gt;
    ///   - &lt;hp:cellMargin left="360" right="360" top="180" bottom="180"/&gt;
    ///   - &lt;hp:subList&gt;&lt;hp:p .../&gt;&lt;/hp:subList&gt;  (at least one paragraph)
    ///
    /// Props: "rows" → row count (default 2), "cols" → col count (default 2),
    ///        "width" → total table width in HWPML units (default 42520 ≈ A4 body width).
    /// </summary>
    private XElement CreateTable(Dictionary<string, string>? props)
    {
        var id = NewId();
        var rows = int.TryParse(props?.GetValueOrDefault("rows"), out var r) && r > 0 ? r : 2;
        var cols = int.TryParse(props?.GetValueOrDefault("cols"), out var c) && c > 0 ? c : 2;   // guard: cols=0 → DivideByZeroException
        var totalWidth = int.TryParse(props?.GetValueOrDefault("width"), out var w) && w > 0 ? w : 42520;
        var cellWidth = totalWidth / cols;
        var cellHeight = 1000; // default cell height

        var tbl = new XElement(HwpxNs.Hp + "tbl",
            new XAttribute("id", id),
            new XAttribute("colCnt", cols.ToString()),
            new XAttribute("rowCnt", rows.ToString()),
            new XAttribute("cellSpacing", "0"),
            new XAttribute("borderFillIDRef", "1")
        );

        // Column widths
        for (int col = 0; col < cols; col++)
        {
            tbl.Add(new XElement(HwpxNs.Hp + "colSz",
                new XAttribute("width", cellWidth.ToString())));
        }

        // Rows and cells
        for (int row = 0; row < rows; row++)
        {
            var tr = new XElement(HwpxNs.Hp + "tr");

            for (int col = 0; col < cols; col++)
            {
                var cellId = NewId();
                var tc = new XElement(HwpxNs.Hp + "tc",
                    new XAttribute("id", cellId),

                    // 1. Cell address — MANDATORY
                    new XElement(HwpxNs.Hp + "cellAddr",
                        new XAttribute("rowAddr", row.ToString()),
                        new XAttribute("colAddr", col.ToString()),
                        new XAttribute("rowSpan", "1"),
                        new XAttribute("colSpan", "1")),

                    // 2. Cell size — MANDATORY
                    new XElement(HwpxNs.Hp + "cellSz",
                        new XAttribute("width", cellWidth.ToString()),
                        new XAttribute("height", cellHeight.ToString())),

                    // 3. Cell margin — MANDATORY
                    new XElement(HwpxNs.Hp + "cellMargin",
                        new XAttribute("left", "360"),
                        new XAttribute("right", "360"),
                        new XAttribute("top", "180"),
                        new XAttribute("bottom", "180")),

                    // 4. SubList with at least one paragraph — MANDATORY
                    new XElement(HwpxNs.Hp + "subList",
                        new XElement(HwpxNs.Hp + "p",
                            new XAttribute("id", NewId()),
                            new XAttribute("styleIDRef", "0"),
                            new XAttribute("paraPrIDRef", "0"),
                            new XElement(HwpxNs.Hp + "run",
                                new XAttribute("charPrIDRef", "0"),
                                new XElement(HwpxNs.Hp + "t", ""))))
                );

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
    /// Generate a unique ID string.
    /// Format: "p" + first 8 hex chars of a new GUID.
    /// Short enough for readability, collision-resistant for single-document scope.
    /// </summary>
    private string NewId()
    {
        return "p" + Guid.NewGuid().ToString("N")[..8];
    }
}
