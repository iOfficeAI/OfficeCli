// Bug hunt tests Part 8: PPTX and Excel handler bugs
// Covers: PPTX Add/Set color handling, table cell scheme color readback,
// gradient scheme color readback, Excel validation operator missing,
// Excel table headerRow/totalRow missing from Set, cross-handler consistency.
// All bugs verified by running tests — every test in this file SHOULD FAIL.

using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ===========================================================================================
    // CATEGORY A: PPTX Add shape color handling — # not stripped, scheme colors not supported
    // PowerPointHandler.Add.cs line 162 uses direct RgbColorModelHex instead of BuildSolidFill
    // ===========================================================================================

    // BUG #801: Add shape text color doesn't strip # prefix
    [Fact]
    public void Bug801_Pptx_Add_ShapeTextColor_HashNotStripped()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug801_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });

            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Red text",
                ["color"] = "#FF0000"
            });

            var node = handler.Get("/slide[1]/shape[2]");
            var color = node.Format.GetValueOrDefault("color")?.ToString() ?? "";

            color.Should().Be("FF0000",
                "PPTX Add shape text color should strip # prefix like BuildColorElement does, " +
                "but Add.cs line 162 uses raw ToUpperInvariant() without TrimStart('#')");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #802: Add shape text color doesn't support scheme colors
    [Fact]
    public void Bug802_Pptx_Add_ShapeTextColor_SchemeColorNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug802_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });

            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Themed text",
                ["color"] = "accent1"
            });

            var raw = handler.Raw("/slide[1]");

            raw.Should().Contain("schemeClr",
                "PPTX Add shape text color should support scheme colors like 'accent1', " +
                "but Add.cs line 162 always creates RgbColorModelHex instead of using BuildColorElement");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #803: Add shape line color doesn't support scheme colors
    [Fact]
    public void Bug803_Pptx_Add_ShapeLineColor_SchemeColorNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug803_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });

            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Bordered shape",
                ["line"] = "accent2"
            });

            var raw = handler.Raw("/slide[1]");

            raw.Should().Contain("schemeClr",
                "PPTX Add shape line color should support scheme colors like 'accent2', " +
                "but Add.cs line 306 always creates RgbColorModelHex");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #804: Add vs Set color inconsistency
    [Fact]
    public void Bug804_Pptx_AddVsSet_ColorConsistency()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug804_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });

            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Test",
                ["color"] = "#FF0000"
            });

            var addNode = handler.Get("/slide[1]/shape[2]");
            var addColor = addNode.Format.GetValueOrDefault("color")?.ToString() ?? "";

            handler.Set("/slide[1]/shape[2]", new() { ["color"] = "#FF0000" });

            var setNode = handler.Get("/slide[1]/shape[2]");
            var setColor = setNode.Format.GetValueOrDefault("color")?.ToString() ?? "";

            addColor.Should().Be(setColor,
                "Add and Set should produce the same color value for the same input, " +
                "but Add doesn't strip # while Set does (via BuildSolidFill)");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ===========================================================================================
    // CATEGORY B: PPTX table cell color/fill scheme color issues
    // ===========================================================================================

    // BUG #1102: Gradient stops with scheme colors stored as invalid hex
    // BuildGradientFill line 271 uses RgbColorModelHex for ALL colors,
    // so scheme names like "accent1" are stored as hex "ACCENT1" which is invalid OOXML
    [Fact]
    public void Bug1102_Pptx_GradientStop_SchemeColorStoredAsInvalidHex()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1102_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Test" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Gradient" });

            handler.Set("/slide[1]/shape[2]", new()
            {
                ["gradient"] = "accent1-accent2"
            });

            var raw = handler.Raw("/slide[1]");

            // BUG: scheme color names are stored as <a:srgbClr val="ACCENT1"/>
            // instead of <a:schemeClr val="accent1"/>
            // BuildGradientFill (Background.cs line 271) always uses RgbColorModelHex
            raw.Should().Contain("schemeClr",
                "gradient stops with scheme color names should use schemeClr elements, " +
                "but BuildGradientFill always uses RgbColorModelHex");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #1103: Table cell text color doesn't strip # prefix
    [Fact]
    public void Bug1103_Pptx_Set_TableCellTextColor_HashNotStripped()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1103_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });
            handler.Add("/slide[1]", "table", null, new()
            {
                ["rows"] = "1", ["cols"] = "1"
            });

            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            {
                ["text"] = "Colored cell",
                ["color"] = "#FF0000"
            });

            var raw = handler.Raw("/slide[1]");

            raw.Should().NotContain("val=\"#FF0000\"",
                "cell text color should strip # prefix before storing, " +
                "but ShapeProperties.cs line 516 does value.ToUpperInvariant() without TrimStart('#')");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #1104: Table cell text color doesn't support scheme colors
    [Fact]
    public void Bug1104_Pptx_Set_TableCellTextColor_SchemeColorNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1104_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });
            handler.Add("/slide[1]", "table", null, new()
            {
                ["rows"] = "1", ["cols"] = "1"
            });

            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            {
                ["text"] = "Themed",
                ["color"] = "accent1"
            });

            var raw = handler.Raw("/slide[1]");

            raw.Should().Contain("schemeClr",
                "table cell text color should support scheme colors, " +
                "but ShapeProperties.cs line 516 always uses RgbColorModelHex");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #1105: Table cell fill doesn't support scheme colors in Set
    [Fact]
    public void Bug1105_Pptx_Set_TableCellFill_SchemeColorNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1105_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });
            handler.Add("/slide[1]", "table", null, new()
            {
                ["rows"] = "1", ["cols"] = "1"
            });

            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            {
                ["fill"] = "accent3"
            });

            var raw = handler.Raw("/slide[1]");

            // ShapeProperties.cs line 537 uses direct RgbColorModelHex instead of BuildSolidFill
            raw.Should().Contain("schemeClr",
                "table cell fill should support scheme colors via BuildSolidFill, " +
                "but ShapeProperties.cs line 537 only creates RgbColorModelHex");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ===========================================================================================
    // CATEGORY C: Excel handler bugs — missing properties in Set
    // ===========================================================================================

    // BUG #1201: Excel validation Set missing "operator" property
    [Fact]
    public void Bug1201_Excel_Set_Validation_OperatorNotSupported()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1:A10",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "100"
        });

        var unsupported = _excelHandler.Set("/Sheet1/validation[1]", new()
        {
            ["operator"] = "greaterThan"
        });

        // BUG: "operator" is returned as unsupported
        // ExcelHandler.Set.cs validation section (lines 90-147) has no case for "operator"
        // but Add.cs supports it
        unsupported.Should().NotContain("operator",
            "Excel validation Set should support 'operator' property, " +
            "but it's missing from the switch in ExcelHandler.Set.cs");
    }

    // BUG #1202: Excel table Set missing "headerRow" property
    [Fact]
    public void Bug1202_Excel_Set_Table_HeaderRowNotSupported()
    {
        _excelHandler.Add("/Sheet1", "row", null, new() { ["values"] = "Name,Age" });
        _excelHandler.Add("/Sheet1", "row", null, new() { ["values"] = "Alice,30" });
        _excelHandler.Add("/Sheet1", "table", null, new()
        {
            ["ref"] = "A1:B2",
            ["name"] = "Table1",
            ["displayname"] = "Table1"
        });

        var unsupported = _excelHandler.Set("/Sheet1/table[1]", new()
        {
            ["headerRow"] = "false"
        });

        // BUG: "headerRow" is unsupported in Set
        unsupported.Should().NotContain("headerRow",
            "Excel table Set should support 'headerRow' property");
    }

    // BUG #1203: Excel table Set missing "totalRow" property
    [Fact]
    public void Bug1203_Excel_Set_Table_TotalRowNotSupported()
    {
        _excelHandler.Add("/Sheet1", "row", null, new() { ["values"] = "Name,Age" });
        _excelHandler.Add("/Sheet1", "row", null, new() { ["values"] = "Alice,30" });
        _excelHandler.Add("/Sheet1", "table", null, new()
        {
            ["ref"] = "A1:B2",
            ["name"] = "Table2",
            ["displayname"] = "Table2"
        });

        var unsupported = _excelHandler.Set("/Sheet1/table[1]", new()
        {
            ["totalRow"] = "true"
        });

        // BUG: "totalRow" is unsupported in Set
        unsupported.Should().NotContain("totalRow",
            "Excel table Set should support 'totalRow' property");
    }

    // ===========================================================================================
    // CATEGORY D: PPTX shape Add — missing properties that Set supports
    // ===========================================================================================

    // BUG #1301: PPTX Add shape doesn't support "underline"
    [Fact]
    public void Bug1301_Pptx_Add_Shape_UnderlineNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1301_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });

            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Underlined",
                ["underline"] = "sng"
            });

            var node = handler.Get("/slide[1]/shape[2]");

            // Check if underline was applied during Add
            node.Format.Should().ContainKey("underline",
                "PPTX Add shape should support 'underline' property during creation, " +
                "since Set supports it via ShapeProperties.cs");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #1302: PPTX Add shape doesn't support "strikethrough"
    [Fact]
    public void Bug1302_Pptx_Add_Shape_StrikethroughNotSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1302_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });

            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Struck",
                ["strikethrough"] = "sngStrike"
            });

            var node = handler.Get("/slide[1]/shape[2]");

            node.Format.Should().ContainKey("strikethrough",
                "PPTX Add shape should support 'strikethrough' property during creation, " +
                "since Set supports it via ShapeProperties.cs");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ===========================================================================================
    // CATEGORY E: PPTX table cell property gaps
    // SetTableCellProperties only supports: text, font, size, bold, italic, color, fill, align,
    // gridspan/colspan, margin, valign, border.*, underline, strikethrough.
    // Missing: highlight (unlike Word which has it at run level)
    // ===========================================================================================

    // BUG #1303: PPTX table cell Set doesn't support "underline" propagation
    [Fact]
    public void Bug1303_Pptx_Set_TableCell_UnderlineSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1303_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });
            handler.Add("/slide[1]", "table", null, new()
            {
                ["rows"] = "1", ["cols"] = "1"
            });
            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Test" });

            var unsupported = handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            {
                ["underline"] = "sng"
            });

            unsupported.Should().NotContain("underline",
                "PPTX table cell Set should support 'underline'");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // BUG #1304: PPTX table cell Set doesn't support "strikethrough" propagation
    [Fact]
    public void Bug1304_Pptx_Set_TableCell_StrikethroughSupported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1304_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });
            handler.Add("/slide[1]", "table", null, new()
            {
                ["rows"] = "1", ["cols"] = "1"
            });
            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Test" });

            var unsupported = handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            {
                ["strikethrough"] = "sngStrike"
            });

            unsupported.Should().NotContain("strikethrough",
                "PPTX table cell Set should support 'strikethrough'");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ===========================================================================================
    // CATEGORY F: PPTX shape rotation format inconsistency
    // NodeBuilder reads rotation as "X°" (with degree symbol), but Set expects plain number
    // This means Get → Set round-trip would fail if user feeds Get value back to Set
    // ===========================================================================================

    // BUG #1306: PPTX rotation format doesn't round-trip
    [Fact]
    public void Bug1306_Pptx_Rotation_FormatDoesntRoundTrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug1306_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "T" });
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Rotated",
                ["rotation"] = "45"
            });

            var node = handler.Get("/slide[1]/shape[2]");
            var rotation = node.Format.GetValueOrDefault("rotation")?.ToString() ?? "";

            // NodeBuilder.cs line 326: node.Format["rotation"] = $"{xfrm.Rotation.Value / 60000.0}°";
            // This adds a ° suffix. But Set expects a plain number.
            // If user does: Set(path, { rotation = Get(path).Format["rotation"] }), it would fail.

            // The rotation value should be usable as input to Set
            rotation.Should().NotContain("°",
                "rotation format should be a plain number that can be fed back to Set, " +
                "but NodeBuilder.cs line 326 appends '°' suffix making round-trip impossible");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
