using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using OfficeCli.Core;
using OfficeCli.Handlers;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeCli.Tests;

public class OleAndImageTests : IDisposable
{
    private readonly string _testDir;

    public OleAndImageTests()
    {
        _testDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officecli_tests_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(_testDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_testDir, true); } catch { }
    }

    private string CreateTestDocx(Action<WordprocessingDocument> configure)
    {
        var path = System.IO.Path.Combine(_testDir, $"test_{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        configure(doc);
        return path;
    }

    /// <summary>Creates an inline image Drawing element for testing.</summary>
    private static Run CreateInlineImageRun(MainDocumentPart mainPart, uint docPropId = 1)
    {
        // Add a tiny 1x1 PNG as image part
        var imgPart = mainPart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(CreateMinimalPng()))
            imgPart.FeedData(ms);
        var relId = mainPart.GetIdOfPart(imgPart);

        long cx = 3600000; // 10cm
        long cy = 1800000; // 5cm
        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = docPropId, Name = "test_image.png", Description = "Test inline image" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = docPropId, Name = "test_image.png" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                    )
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
            )
        )
        { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U };
        return new Run(new Drawing(inline));
    }

    /// <summary>Creates a floating (anchor) image Drawing element with specified wrap.</summary>
    private static Run CreateAnchorImageRun(MainDocumentPart mainPart, string wrapType, uint docPropId = 2)
    {
        var imgPart = mainPart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(CreateMinimalPng()))
            imgPart.FeedData(ms);
        var relId = mainPart.GetIdOfPart(imgPart);

        long cx = 2160000; // 6cm
        long cy = 1440000; // 4cm
        long hPos = 720000; // 2cm
        long vPos = 360000; // 1cm

        OpenXmlElement wrapElement = wrapType switch
        {
            "square" => new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
            "tight" => new DW.WrapTight(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "none" => new DW.WrapNone(),
            _ => new DW.WrapNone()
        };

        var anchor = new DW.Anchor(
            new DW.SimplePosition { X = 0, Y = 0 },
            new DW.HorizontalPosition(new DW.PositionOffset(hPos.ToString()))
                { RelativeFrom = DW.HorizontalRelativePositionValues.Column },
            new DW.VerticalPosition(new DW.PositionOffset(vPos.ToString()))
                { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph },
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            wrapElement,
            new DW.DocProperties { Id = docPropId, Name = "anchor_image.png", Description = "Floating image" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = docPropId, Name = "anchor_image.png" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                    )
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
            )
        )
        {
            BehindDoc = false,
            DistanceFromTop = 0U, DistanceFromBottom = 0U,
            DistanceFromLeft = 114300U, DistanceFromRight = 114300U,
            SimplePos = false, RelativeHeight = 1U,
            AllowOverlap = true, LayoutInCell = true, Locked = false
        };
        return new Run(new Drawing(anchor));
    }

    /// <summary>Creates a minimal OLE embedded object (simulates Visio.Drawing.11).</summary>
    private static Run CreateOleObjectRun(string progId = "Visio.Drawing.11", string width = "385.45pt", string height = "397.75pt")
    {
        // Build raw OLE object XML using OpenXmlUnknownElement for VML/OLE parts
        var shapeXml = $"<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                       $"style=\"width:{width};height:{height}\" />";
        var oleXml = $"<o:OLEObject xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
                     $"ProgID=\"{progId}\" />";

        var shape = new OpenXmlUnknownElement("v", "shape", "urn:schemas-microsoft-com:vml");
        shape.SetAttribute(new OpenXmlAttribute("style", "", $"width:{width};height:{height}"));

        var oleEl = new OpenXmlUnknownElement("o", "OLEObject", "urn:schemas-microsoft-com:office:office");
        oleEl.SetAttribute(new OpenXmlAttribute("ProgID", "", progId));

        var embeddedObject = new EmbeddedObject();
        embeddedObject.AppendChild(shape);
        embeddedObject.AppendChild(oleEl);

        return new Run(embeddedObject);
    }

    private static byte[] CreateMinimalPng()
    {
        // Minimal valid 1x1 white PNG
        return Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==");
    }

    // ===================== Tests =====================

    [Fact]
    public void Query_Picture_DetectsInlineImage()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var para = new Paragraph(CreateInlineImageRun(doc.MainDocumentPart!));
            body.AppendChild(para);
        });

        using var handler = new WordHandler(path, false);
        var results = handler.Query("picture");

        Assert.Single(results);
        Assert.Equal("picture", results[0].Type);
        Assert.Equal("inline", results[0].Format["wrap"]);
    }

    [Fact]
    public void Query_Picture_DetectsAnchorImageWithWrapType()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(CreateAnchorImageRun(doc.MainDocumentPart!, "square")));
        });

        using var handler = new WordHandler(path, false);
        var results = handler.Query("picture");

        Assert.Single(results);
        Assert.Equal("picture", results[0].Type);
        Assert.Equal("square", results[0].Format["wrap"]);
        Assert.Equal("2.0cm", results[0].Format["hPosition"]);
        Assert.Equal("1.0cm", results[0].Format["vPosition"]);
        Assert.Equal("column", results[0].Format["hRelative"]);
        Assert.Equal("paragraph", results[0].Format["vRelative"]);
    }

    [Fact]
    public void Query_Picture_DetectsOleObject()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(CreateOleObjectRun()));
        });

        using var handler = new WordHandler(path, false);
        var results = handler.Query("picture");

        Assert.Single(results);
        Assert.Equal("ole", results[0].Type);
        Assert.Equal("Visio.Drawing.11", results[0].Format["progId"]);
    }

    [Fact]
    public void Query_Picture_ReturnsBothDrawingAndOle()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(CreateInlineImageRun(doc.MainDocumentPart!, 1)));
            body.AppendChild(new Paragraph(CreateOleObjectRun()));
            body.AppendChild(new Paragraph(CreateOleObjectRun("Excel.Sheet.12", "200pt", "150pt")));
        });

        using var handler = new WordHandler(path, false);
        var results = handler.Query("picture");

        Assert.Equal(3, results.Count);
        Assert.Equal("picture", results[0].Type);
        Assert.Equal("ole", results[1].Type);
        Assert.Equal("ole", results[2].Type);
        Assert.Equal("Excel.Sheet.12", results[2].Format["progId"]);
    }

    [Fact]
    public void Query_Ole_OnlyReturnsOleObjects()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(CreateInlineImageRun(doc.MainDocumentPart!, 1)));
            body.AppendChild(new Paragraph(CreateOleObjectRun()));
            body.AppendChild(new Paragraph(CreateOleObjectRun()));
        });

        using var handler = new WordHandler(path, false);
        var results = handler.Query("ole");

        Assert.Equal(2, results.Count);
        Assert.All(results, r => Assert.Equal("ole", r.Type));
    }

    [Fact]
    public void Query_Object_IsAliasForOle()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(CreateOleObjectRun()));
        });

        using var handler = new WordHandler(path, false);
        var oleResults = handler.Query("ole");
        var objectResults = handler.Query("object");

        Assert.Single(oleResults);
        Assert.Single(objectResults);
        Assert.Equal(oleResults[0].Format["progId"], objectResults[0].Format["progId"]);
    }

    [Fact]
    public void Query_Ole_ExtractsDimensions()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(CreateOleObjectRun("Visio.Drawing.11", "385.45pt", "397.75pt")));
        });

        using var handler = new WordHandler(path, false);
        var results = handler.Query("ole");

        Assert.Single(results);
        Assert.Equal("ole", results[0].Format["objectType"]);
        // 385.45pt * 2.54/72 = ~13.6cm
        var width = results[0].Format["width"]?.ToString();
        Assert.NotNull(width);
        Assert.EndsWith("cm", width);
    }

    [Fact]
    public void View_Outline_IncludesOleCount()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(CreateInlineImageRun(doc.MainDocumentPart!, 1)));
            body.AppendChild(new Paragraph(CreateOleObjectRun()));
            body.AppendChild(new Paragraph(CreateOleObjectRun()));
        });

        using var handler = new WordHandler(path, false);
        var json = handler.ViewAsOutlineJson();

        Assert.Equal(1, (int)json["images"]!);
        Assert.Equal(2, (int)json["oleObjects"]!);
    }

    [Fact]
    public void View_Outline_NoOleField_WhenZero()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(new Run(new Text("Hello"))));
        });

        using var handler = new WordHandler(path, false);
        var json = handler.ViewAsOutlineJson();

        Assert.Null(json["oleObjects"]);
    }

    [Fact]
    public void Query_Picture_WrapNone_BehindText()
    {
        var path = CreateTestDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            // Create anchor with WrapNone and BehindDoc=true
            var imgPart = doc.MainDocumentPart!.AddImagePart(ImagePartType.Png);
            using (var ms = new MemoryStream(CreateMinimalPng()))
                imgPart.FeedData(ms);
            var relId = doc.MainDocumentPart!.GetIdOfPart(imgPart);

            long cx = 2160000, cy = 1440000;
            var anchor = new DW.Anchor(
                new DW.SimplePosition { X = 0, Y = 0 },
                new DW.HorizontalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.HorizontalRelativePositionValues.Page },
                new DW.VerticalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.VerticalRelativePositionValues.Page },
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.WrapNone(),
                new DW.DocProperties { Id = 1, Name = "bg" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 1, Name = "bg" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(new A.Blip { Embed = relId }, new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            ) { BehindDoc = true, SimplePos = false, RelativeHeight = 1U, AllowOverlap = true, LayoutInCell = true, Locked = false };

            body.AppendChild(new Paragraph(new Run(new Drawing(anchor))));
        });

        using var handler = new WordHandler(path, false);
        var results = handler.Query("picture");

        Assert.Single(results);
        Assert.Equal("none", results[0].Format["wrap"]);
        Assert.Equal(true, results[0].Format["behindText"]);
        Assert.Equal("page", results[0].Format["hRelative"]);
        Assert.Equal("page", results[0].Format["vRelative"]);
    }
}
