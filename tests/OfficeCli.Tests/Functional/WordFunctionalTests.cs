// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for DOCX: each test creates a blank file, adds elements,
/// queries them, and modifies them — exercising the full Create→Add→Get→Set lifecycle.
/// </summary>
public class WordFunctionalTests : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public WordFunctionalTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== DOCX Hyperlinks ====================

    [Fact]
    public void Hyperlink_Lifecycle()
    {
        // 1. Add paragraph + hyperlink
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>());
        var path = _handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://first.com",
            ["text"] = "Click here"
        });
        path.Should().Be("/body/p[1]/hyperlink[1]");

        // 2. Get + Verify type, url, text
        var node = _handler.Get("/body/p[1]/hyperlink[1]");
        node.Type.Should().Be("hyperlink");
        node.Text.Should().Be("Click here");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://first.com");

        // 3. Verify paragraph text contains link text
        var para = _handler.Get("/body/p[1]");
        para.Text.Should().Contain("Click here");

        // 4. Query + Verify
        var results = _handler.Query("hyperlink");
        results.Should().Contain(n => n.Type == "hyperlink" && n.Text == "Click here");

        // 5. Set (update URL via run) + Verify
        _handler.Set("/body/p[1]/r[1]", new Dictionary<string, string> { ["link"] = "https://updated.com" });
        node = _handler.Get("/body/p[1]/hyperlink[1]");
        ((string)node.Format["link"]).Should().StartWith("https://updated.com");
    }

    [Fact]
    public void Hyperlink_Persist_SurvivesReopenFile()
    {
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>());
        _handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://original.com",
            ["text"] = "My link"
        });
        _handler.Set("/body/p[1]/r[1]", new Dictionary<string, string> { ["link"] = "https://persist.com" });

        var handler2 = Reopen();
        var node = handler2.Get("/body/p[1]/hyperlink[1]");
        node.Text.Should().Be("My link");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://persist.com");
    }
}
