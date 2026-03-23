// PptxProactiveBugTests.cs — Proactive bug scan for PPTX handler issues.
//
// Bug 1: Picture Set lists "name" as valid prop but doesn't implement the case.
//         Set(picture, { name = "MyPic" }) silently returns "name" as unsupported.
//
// Bug 2: Picture Set lists "opacity" as valid prop but doesn't implement the case.
//         Set(picture, { opacity = "0.5" }) silently returns "opacity" as unsupported.
//
// Bug 3: Add shape doesn't delegate "image"/"imagefill" to SetRunOrShapeProperties.
//         The effectKeys set is missing these keys, so Add(shape, { image = "..." }) ignores them.

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxProactiveScanBugTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTempPptx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 1: Picture Set "name" not implemented
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public void PptxProactive_SetPictureName_ShouldUpdateName()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        // Create a small 1x1 red PNG in memory and save to temp file
        var imgPath = CreateTempPng();
        handler.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath, ["name"] = "OrigPic" });

        // Verify original name
        var node = handler.Get("/slide[1]/picture[1]");
        node.Format["name"].Should().Be("OrigPic");

        // Set new name — this should NOT return "name" as unsupported
        var unsupported = handler.Set("/slide[1]/picture[1]", new() { ["name"] = "RenamedPic" });
        unsupported.Should().BeEmpty("name should be a supported picture property");

        // Verify name updated
        var node2 = handler.Get("/slide[1]/picture[1]");
        node2.Format["name"].Should().Be("RenamedPic");
    }

    [Fact]
    public void PptxProactive_SetPictureName_PersistsAfterReopen()
    {
        var path = CreateTempPptx();
        var imgPath = CreateTempPng();

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath, ["name"] = "Pic1" });
            handler.Set("/slide[1]/picture[1]", new() { ["name"] = "PersistPic" });
        }

        // Reopen and verify
        using (var handler2 = new PowerPointHandler(path, editable: false))
        {
            var node = handler2.Get("/slide[1]/picture[1]");
            node.Format["name"].Should().Be("PersistPic");
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 2: Picture Set "opacity" not implemented
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public void PptxProactive_SetPictureOpacity_ShouldNotReturnUnsupported()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var imgPath = CreateTempPng();
        handler.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });

        // Set opacity — should NOT return "opacity" as unsupported
        var unsupported = handler.Set("/slide[1]/picture[1]", new() { ["opacity"] = "0.5" });
        unsupported.Should().BeEmpty("opacity should be a supported picture property");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 3: Add shape missing "image"/"imagefill" in effectKeys
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public void PptxProactive_AddShapeWithImageFill_ShouldApplyImageFill()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var imgPath = CreateTempPng();
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "ImageFilled",
            ["image"] = imgPath
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // If image fill was applied, the "image" format key should be present
        node.Format.Should().ContainKey("image", "image fill should be applied during Add");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Helper: create a minimal 1x1 PNG
    // ─────────────────────────────────────────────────────────────────────────

    private string CreateTempPng()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        _tempFiles.Add(pngPath);
        // Minimal valid 1x1 red PNG (67 bytes)
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, // 8-bit RGB
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x03, 0x00, 0x01, 0x36, 0x28, 0x19,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(pngPath, pngBytes);
        return pngPath;
    }
}
