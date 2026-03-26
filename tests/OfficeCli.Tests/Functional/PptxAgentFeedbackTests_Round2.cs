// Agent Feedback Round 2 — Bug tests discovered by Agent A's PPT testing (second round)
// Tests for video timing validation, group children/format, chart readback correctness.
//
// Selected bugs (by priority):
//   Bug 15 (CRITICAL): video add creates invalid bldLst in animation timing
//   Bug 5/7 (HIGH): group children always empty, member shapes inaccessible
//   Bug 8 (MEDIUM): group format missing x/y/width/height on direct Get
//   Bug 16 (MEDIUM): title.bold=false removes key instead of returning false
//   Bug 1 (MEDIUM): chart legend returns raw OOXML codes (b/r/t/l) not readable values

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Tests.Functional;

public class PptxAgentFeedbackTests_Round2 : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxAgentFeedbackTests_Round2()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Reflection Helpers ====================

    private PresentationDocument GetDoc()
    {
        return (PresentationDocument)_handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler)!;
    }

    private string AddBarChart(string title = "Sales")
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "S" });
        return _handler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = title,
            ["categories"] = "Q1,Q2,Q3",
            ["data"] = "Revenue:100,200,300"
        });
    }

    // ==================== Bug 15: video add creates invalid bldLst in animation timing ====================
    // Adding a video calls EnsureTimingTree which always creates an empty bldLst.
    // An empty bldLst violates OOXML schema (requires at least one child element).
    // Media playback timing should not create bldLst at all.

    [Fact]
    public void Bug15_VideoAdd_ShouldNotCreateInvalidBldLst()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Video Slide" });

        // Create a minimal dummy video file for the test
        var videoPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.mp4");
        try
        {
            // Write minimal bytes -- just enough to pass file existence check
            File.WriteAllBytes(videoPath, new byte[64]);

            _handler.Add("/slide[1]", "video", null, new() { ["path"] = videoPath });

            // Validate the document
            var errors = _handler.Validate();
            var bldErrors = errors.Where(e =>
                e.Description.Contains("bldLst", StringComparison.OrdinalIgnoreCase) ||
                e.Description.Contains("BuildList", StringComparison.OrdinalIgnoreCase) ||
                e.Description.Contains("bldP", StringComparison.OrdinalIgnoreCase)).ToList();

            bldErrors.Should().BeEmpty(
                "video add should not create invalid bldLst/bldP elements in animation timing");

            // Additionally, if bldLst exists, it should not be empty
            var doc = GetDoc();
            var slidePart = doc.PresentationPart!.SlideParts.First();
            var timing = slidePart.Slide.GetFirstChild<Timing>();
            if (timing?.BuildList != null)
            {
                // An empty bldLst is itself invalid per OOXML schema
                timing.BuildList.HasChildren.Should().BeTrue(
                    "if bldLst exists, it must not be empty (empty bldLst violates OOXML schema)");
            }
        }
        finally
        {
            if (File.Exists(videoPath)) File.Delete(videoPath);
        }
    }

    [Fact]
    public void Bug15_VideoAdd_DocumentShouldPassFullValidation()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Video Validation" });

        var videoPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.mp4");
        try
        {
            File.WriteAllBytes(videoPath, new byte[64]);
            _handler.Add("/slide[1]", "video", null, new() { ["path"] = videoPath });

            var errors = _handler.Validate();
            // Filter to slide-level errors to focus on the video timing issue
            var timingErrors = errors.Where(e =>
                e.Part != null && e.Part.Contains("slide", StringComparison.OrdinalIgnoreCase)).ToList();

            timingErrors.Should().BeEmpty(
                "adding a video should not introduce any validation errors on the slide");
        }
        finally
        {
            if (File.Exists(videoPath)) File.Delete(videoPath);
        }
    }

    // ==================== Bug 5/7: group children always empty [] ====================
    // Get(/slide[N]/group[M]) returns a node with ChildCount > 0 but Children list is empty.
    // Member shapes inside a group are inaccessible through the DOM API.

    [Fact]
    public void Bug5_GroupChildren_ShouldNotBeEmpty()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Group Test" });

        // Add two shapes to group
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape A",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "5cm", ["height"] = "3cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape B",
            ["x"] = "8cm", ["y"] = "2cm", ["width"] = "5cm", ["height"] = "3cm"
        });

        // Group them
        var groupPath = _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });
        groupPath.Should().NotBeNullOrEmpty();

        // Get group node -- ChildCount should be > 0
        var groupNode = _handler.Get(groupPath);
        groupNode.Should().NotBeNull();
        groupNode.Type.Should().Be("group");
        groupNode.ChildCount.Should().BeGreaterThan(0,
            "group contains member shapes, so ChildCount should reflect that");

        // The actual Children list should also be populated, not empty
        groupNode.Children.Should().NotBeEmpty(
            "group members should be accessible via Children list, not just counted");
        groupNode.Children.Count.Should().Be(groupNode.ChildCount,
            "Children.Count should match ChildCount");
    }

    [Fact]
    public void Bug7_GroupChildren_MemberShapes_ShouldBeAccessibleByPath()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Group Access" });

        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Alpha",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "2cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Beta",
            ["x"] = "6cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "2cm"
        });

        var groupPath = _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });

        // Children of the group should have valid Path values
        var groupNode = _handler.Get(groupPath);
        groupNode.Children.Should().NotBeEmpty("group should expose its member shapes");

        // Each child should have a path that includes the group path prefix
        foreach (var child in groupNode.Children)
        {
            child.Path.Should().StartWith(groupPath,
                "group member paths should be nested under the group path");
            child.Type.Should().NotBeNullOrEmpty(
                "group member nodes should have a type");
        }
    }

    // ==================== Bug 8: group format missing x/y/width/height on direct Get ====================
    // Get(/slide[N]/group[M]) does not populate x/y/width/height in Format,
    // even though the group has a TransformGroup with valid position/size.
    // The NodeBuilder code path (slide children listing) does populate these,
    // but the direct Get path does not -- inconsistent behavior.

    [Fact]
    public void Bug8_GroupFormat_ShouldIncludePositionAndSize()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Group Format" });

        // Add two shapes with known positions
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A",
            ["x"] = "2cm", ["y"] = "3cm", ["width"] = "5cm", ["height"] = "4cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "B",
            ["x"] = "8cm", ["y"] = "3cm", ["width"] = "5cm", ["height"] = "4cm"
        });

        // Group them
        var groupPath = _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });

        // Direct Get should return position/size in Format
        var groupNode = _handler.Get(groupPath);
        groupNode.Should().NotBeNull();
        groupNode.Type.Should().Be("group");

        // These should be present -- the group has a bounding box from TransformGroup
        groupNode.Format.Should().ContainKey("x",
            "group node should include 'x' position in Format when accessed via Get");
        groupNode.Format.Should().ContainKey("y",
            "group node should include 'y' position in Format when accessed via Get");
        groupNode.Format.Should().ContainKey("width",
            "group node should include 'width' in Format when accessed via Get");
        groupNode.Format.Should().ContainKey("height",
            "group node should include 'height' in Format when accessed via Get");
    }

    [Fact]
    public void Bug8_GroupFormat_DirectGet_ShouldMatchSlideChildrenListing()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Group Consistency" });

        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "X",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Y",
            ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });

        var groupPath = _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });

        // Get via slide children listing (NodeBuilder path)
        var slideNode = _handler.Get("/slide[1]");
        var groupFromSlide = slideNode.Children.FirstOrDefault(c => c.Type == "group");
        groupFromSlide.Should().NotBeNull("group should appear in slide children");

        // Get via direct path
        var groupDirect = _handler.Get(groupPath);

        // Both should have the same format keys for position/size
        if (groupFromSlide!.Format.ContainsKey("x"))
        {
            groupDirect.Format.Should().ContainKey("x",
                "direct Get should return same format keys as slide children listing");
            groupDirect.Format["x"].Should().Be(groupFromSlide.Format["x"],
                "direct Get and slide listing should return same x value");
        }
    }

    // ==================== Bug 16: title.bold=false removes key rather than returning false ====================
    // ChartReader only writes title.bold when bold==true (line 44-45 in ChartReader.cs).
    // When bold is explicitly set to false, Get returns no title.bold key at all.
    // Expected: title.bold should be "false" when explicitly set to false.

    [Fact]
    public void Bug16_ChartTitleBoldFalse_ShouldReadBackAsFalse()
    {
        var chartPath = AddBarChart("Bold Test");

        // First set bold to true
        _handler.Set(chartPath, new() { ["title.bold"] = "true" });
        var node1 = _handler.Get(chartPath);
        node1.Format.Should().ContainKey("title.bold");
        node1.Format["title.bold"].Should().Be("true");

        // Now set bold to false explicitly
        _handler.Set(chartPath, new() { ["title.bold"] = "false" });
        var node2 = _handler.Get(chartPath);

        // title.bold should be present with value "false", not absent
        node2.Format.Should().ContainKey("title.bold",
            "title.bold should be readable as 'false' after being explicitly set to false, not silently dropped");
        node2.Format["title.bold"].Should().Be("false",
            "explicitly setting bold=false should read back as 'false'");
    }

    [Fact]
    public void Bug16_ChartTitleBoldFalse_PersistsAfterReopen()
    {
        var chartPath = AddBarChart("Bold Persist");

        _handler.Set(chartPath, new() { ["title.bold"] = "true" });
        _handler.Set(chartPath, new() { ["title.bold"] = "false" });

        Reopen();

        var node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("title.bold",
            "title.bold=false should persist after save/reopen");
        node.Format["title.bold"].Should().Be("false");
    }

    // ==================== Bug 1: chart legend returns raw OOXML codes ====================
    // Get returns legend position as raw OOXML abbreviations: "b", "r", "t", "l"
    // instead of human-readable values: "bottom", "right", "top", "left".
    // This violates the canonical value format rule (user-facing values should be readable).

    [Fact]
    public void Bug1_ChartLegend_ShouldReturnReadablePosition_NotRawCode()
    {
        var chartPath = AddBarChart("Legend Test");

        // Set legend to right position
        _handler.Set(chartPath, new() { ["legend"] = "right" });
        var node = _handler.Get(chartPath);

        node.Format.Should().ContainKey("legend");
        // The value should be "right", not the raw OOXML "r"
        node.Format["legend"].Should().Be("right",
            "legend position should be human-readable 'right', not raw OOXML code 'r'");
    }

    [Fact]
    public void Bug1_ChartLegend_AllPositions_ShouldBeReadable()
    {
        var chartPath = AddBarChart("Legend Positions");

        // Test each standard position
        var positions = new Dictionary<string, string>
        {
            ["top"] = "top",
            ["bottom"] = "bottom",
            ["left"] = "left",
            ["right"] = "right",
        };

        foreach (var (input, expected) in positions)
        {
            _handler.Set(chartPath, new() { ["legend"] = input });
            var node = _handler.Get(chartPath);

            node.Format.Should().ContainKey("legend");
            node.Format["legend"].Should().Be(expected,
                $"legend position '{input}' should read back as '{expected}', not as raw OOXML code");
        }
    }
}
