// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeCli.Core;

public enum OutputFormat
{
    Text,
    Json
}

public class ViewResult
{
    public string View { get; set; } = "";
    public string Content { get; set; } = "";
}

public class NodesResult
{
    public int Matches { get; set; }
    public List<DocumentNode> Results { get; set; } = new();
}

public class IssuesResult
{
    public int Count { get; set; }
    public List<DocumentIssue> Issues { get; set; } = new();
}

[JsonSourceGenerationOptions(
    WriteIndented = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(ViewResult))]
[JsonSerializable(typeof(NodesResult))]
[JsonSerializable(typeof(IssuesResult))]
[JsonSerializable(typeof(DocumentNode))]
[JsonSerializable(typeof(List<DocumentNode>))]
[JsonSerializable(typeof(List<DocumentIssue>))]
[JsonSerializable(typeof(Dictionary<string, object?>))]
[JsonSerializable(typeof(bool))]
[JsonSerializable(typeof(int))]
[JsonSerializable(typeof(long))]
[JsonSerializable(typeof(short))]
[JsonSerializable(typeof(uint))]
[JsonSerializable(typeof(double))]
[JsonSerializable(typeof(string))]
internal partial class AppJsonContext : JsonSerializerContext;

public static class OutputFormatter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        TypeInfoResolver = AppJsonContext.Default
    };

    public static string FormatView(string view, string content, OutputFormat format)
    {
        return format switch
        {
            OutputFormat.Json => JsonSerializer.Serialize(new ViewResult { View = view, Content = content }, AppJsonContext.Default.ViewResult),
            _ => content
        };
    }

    public static string FormatNode(DocumentNode node, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(node, AppJsonContext.Default.DocumentNode);

        return FormatNodeAsText(node, 0);
    }

    public static string FormatNodes(List<DocumentNode> nodes, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(new NodesResult { Matches = nodes.Count, Results = nodes }, AppJsonContext.Default.NodesResult);

        var sb = new StringBuilder();
        sb.AppendLine($"Matches: {nodes.Count}");
        foreach (var node in nodes)
        {
            sb.AppendLine($"  {node.Path}: {node.Text ?? node.Preview ?? node.Type}");
            foreach (var (key, val) in node.Format)
                sb.AppendLine($"    {key}: {val}");
        }
        return sb.ToString().TrimEnd();
    }

    public static string FormatIssues(List<DocumentIssue> issues, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(new IssuesResult { Count = issues.Count, Issues = issues }, AppJsonContext.Default.IssuesResult);

        var sb = new StringBuilder();
        sb.AppendLine($"Found {issues.Count} issue(s):");
        sb.AppendLine();

        var grouped = issues.GroupBy(i => i.Type);
        foreach (var group in grouped)
        {
            var typeName = group.Key switch
            {
                IssueType.Format => "Format Issues",
                IssueType.Content => "Content Issues",
                IssueType.Structure => "Structure Issues",
                _ => "Other"
            };
            sb.AppendLine($"{typeName} ({group.Count()}):");

            foreach (var issue in group)
            {
                var severity = issue.Severity switch
                {
                    IssueSeverity.Error => "ERROR",
                    IssueSeverity.Warning => "WARN",
                    _ => "INFO"
                };
                sb.AppendLine($"  [{issue.Id}] {issue.Path}: {issue.Message}");
                if (issue.Context != null)
                    sb.AppendLine($"       Context: \"{issue.Context}\"");
                if (issue.Suggestion != null)
                    sb.AppendLine($"       Suggestion: {issue.Suggestion}");
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private static string FormatNodeAsText(DocumentNode node, int indent)
    {
        var sb = new StringBuilder();
        var prefix = new string(' ', indent * 2);

        sb.Append($"{prefix}{node.Path} ({node.Type})");
        if (node.Text != null) sb.Append($" \"{Truncate(node.Text, 60)}\"");
        if (node.Style != null) sb.Append($" [{node.Style}]");
        if (node.ChildCount > 0 && node.Children.Count == 0) sb.Append($" ({node.ChildCount} children)");
        sb.AppendLine();

        foreach (var (key, val) in node.Format)
            sb.AppendLine($"{prefix}  {key}: {val}");

        foreach (var child in node.Children)
            sb.Append(FormatNodeAsText(child, indent + 1));

        return sb.ToString();
    }

    private static string Truncate(string s, int maxLen)
    {
        return s.Length > maxLen ? s[..maxLen] + "..." : s;
    }
}
