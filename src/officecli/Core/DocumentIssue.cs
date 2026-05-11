// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Serialization;

namespace OfficeCli.Core;

public enum IssueType
{
    Format,
    Content,
    Structure
}

public enum IssueSeverity
{
    Error,
    Warning,
    Info
}

public class DocumentIssue
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";
    [JsonPropertyName("type")]
    public IssueType Type { get; set; }
    /// <summary>
    /// Machine-readable issue subtype. Stable identifier agents can match on
    /// (snake_case). Doubles as the value accepted by `view issues --type`
    /// for narrow filtering. Examples: formula_not_evaluated,
    /// field_not_evaluated, slide_field_not_evaluated,
    /// chart_series_ref_missing_sheet, chart_cache_stale,
    /// definedname_broken. Distinct from the broad <see cref="Type"/> enum
    /// (Format / Content / Structure) which buckets the issue category.
    /// </summary>
    [JsonPropertyName("subtype")]
    public string? Subtype { get; set; }
    [JsonPropertyName("severity")]
    public IssueSeverity Severity { get; set; }
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";
    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
    [JsonPropertyName("context")]
    public string? Context { get; set; }
    [JsonPropertyName("suggestion")]
    public string? Suggestion { get; set; }
}
