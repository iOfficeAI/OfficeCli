// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Common interface for all document types (Word/Excel/PowerPoint).
/// Each handler implements the three-layer architecture:
///   - Semantic layer: view (text/annotated/outline/stats/issues)
///   - Query layer: get, query, set
///   - Raw layer: raw XML access
/// </summary>
public interface IDocumentHandler : IDisposable
{
    // === Semantic Layer ===
    string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null);
    string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null);
    string ViewAsOutline();
    string ViewAsStats();

    // === Structured JSON variants (for --json mode) ===
    System.Text.Json.Nodes.JsonNode ViewAsStatsJson();
    System.Text.Json.Nodes.JsonNode ViewAsOutlineJson();
    System.Text.Json.Nodes.JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null);
    List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null);

    // === Query Layer ===
    DocumentNode Get(string path, int depth = 1);
    List<DocumentNode> Query(string selector);
    /// <summary>
    /// Returns list of prop names that were not applied (unsupported for this element type).
    /// </summary>
    List<string> Set(string path, Dictionary<string, string> properties);
    string Add(string parentPath, string type, int? index, Dictionary<string, string> properties);
    /// <summary>
    /// Remove element at path. Returns an optional warning message (e.g. formula cells affected by shift).
    /// </summary>
    string? Remove(string path);
    string Move(string sourcePath, string? targetParentPath, int? index);
    string CopyFrom(string sourcePath, string targetParentPath, int? index);

    // === Raw Layer ===
    string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null);
    void RawSet(string partPath, string xpath, string action, string? xml);

    /// <summary>
    /// Create a new part (chart, header, footer, etc.) and return its relationship ID and accessible path.
    /// </summary>
    (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null);

    /// <summary>
    /// Validate the document against OpenXML schema and return any errors.
    /// </summary>
    List<ValidationError> Validate();
}

public record ValidationError(string ErrorType, string Description, string? Path, string? Part);
