// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// <see cref="IDocumentHandler"/> implementation that delegates every call to a
/// running format-handler plugin via <see cref="FormatHandlerSession"/>. Per
/// docs/plugin-protocol.md §2.3, this is what wraps the plugin so existing
/// get/view/query pipelines work transparently on foreign formats.
///
/// Scope: read-path (ViewAs*, Get, Query, Validate) and mutation
/// (Set/Add/Remove/Move/CopyFrom/Raw/RawSet/AddPart/TryExtractBinary)
/// are all proxied. Plugins that don't implement a given verb should
/// reply with error code <c>unsupported_command</c> per docs/plugin-protocol.md §5.3.
/// </summary>
internal sealed class FormatHandlerProxy : IDocumentHandler
{
    private readonly FormatHandlerSession _session;

    public FormatHandlerProxy(FormatHandlerSession session) { _session = session; }

    // ----- Semantic layer (text views) -----------------------------------

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewString("text", startLine, endLine, maxLines, cols);

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewString("annotated", startLine, endLine, maxLines, cols);

    public string ViewAsOutline()
        => SendViewString("outline");

    public string ViewAsStats()
        => SendViewString("stats");

    public JsonNode ViewAsStatsJson() => SendViewJson("stats");
    public JsonNode ViewAsOutlineJson() => SendViewJson("outline");
    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewJson("text", startLine, endLine, maxLines, cols);

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var args = new JsonObject { ["mode"] = "issues" };
        if (issueType != null) args["type"] = issueType;
        if (limit.HasValue) args["limit"] = limit.Value;
        var result = _session.Send("command", "view", args);
        if (result is null) return new List<DocumentIssue>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListDocumentIssue) ?? new List<DocumentIssue>();
    }

    // ----- Query layer --------------------------------------------------

    public DocumentNode Get(string path, int depth = 1)
    {
        var result = _session.Send("command", "get", new JsonObject
        {
            ["path"] = path,
            ["depth"] = depth,
        });
        if (result is null)
            return new DocumentNode { Path = path, Type = "error", Text = "Plugin returned null result." };
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.DocumentNode)
            ?? new DocumentNode { Path = path, Type = "error", Text = "Plugin result did not deserialize as DocumentNode." };
    }

    public List<DocumentNode> Query(string selector)
    {
        var result = _session.Send("command", "query", new JsonObject { ["selector"] = selector });
        if (result is null) return new List<DocumentNode>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListDocumentNode) ?? new List<DocumentNode>();
    }

    public List<ValidationError> Validate()
    {
        var result = _session.Send("command", "validate", new JsonObject());
        if (result is null) return new List<ValidationError>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListValidationError) ?? new List<ValidationError>();
    }

    // ----- Mutation layer ----------------------------------------------

    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var args = new JsonObject { ["path"] = path };
        var props = PropsToJson(properties);
        var result = _session.Send("command", "set", args, props);
        if (result is null) return new List<string>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListString) ?? new List<string>();
    }

    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var args = new JsonObject
        {
            ["parent_path"] = parentPath,
            ["type"] = type,
        };
        if (position is not null) args["position"] = PositionToJson(position);
        var props = PropsToJson(properties);
        var result = _session.Send("command", "add", args, props);
        return result?.GetValue<string>() ?? "";
    }

    public string? Remove(string path)
    {
        var result = _session.Send("command", "remove", new JsonObject { ["path"] = path });
        return result?.GetValue<string>();
    }

    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position)
    {
        var args = new JsonObject { ["source_path"] = sourcePath };
        if (targetParentPath is not null) args["target_parent_path"] = targetParentPath;
        if (position is not null) args["position"] = PositionToJson(position);
        var result = _session.Send("command", "move", args);
        return result?.GetValue<string>() ?? "";
    }

    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position)
    {
        var args = new JsonObject
        {
            ["source_path"] = sourcePath,
            ["target_parent_path"] = targetParentPath,
        };
        if (position is not null) args["position"] = PositionToJson(position);
        var result = _session.Send("command", "copy", args);
        return result?.GetValue<string>() ?? "";
    }

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var args = new JsonObject { ["part_path"] = partPath };
        if (startRow.HasValue) args["start_row"] = startRow.Value;
        if (endRow.HasValue) args["end_row"] = endRow.Value;
        if (cols != null && cols.Count > 0) args["cols"] = string.Join(",", cols);
        var result = _session.Send("command", "raw", args);
        return result?.GetValue<string>() ?? "";
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var args = new JsonObject
        {
            ["part_path"] = partPath,
            ["xpath"] = xpath,
            ["action"] = action,
        };
        if (xml is not null) args["xml"] = xml;
        _session.Send("command", "raw-set", args);
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var args = new JsonObject
        {
            ["parent_part_path"] = parentPartPath,
            ["part_type"] = partType,
        };
        var props = properties is not null ? PropsToJson(properties) : null;
        var result = _session.Send("command", "add-part", args, props)?.AsObject();
        if (result is null)
            throw new CliException("Format-handler add-part returned null.") { Code = "protocol_mismatch" };
        var relId = result["rel_id"]?.GetValue<string>() ?? "";
        var partPath = result["part_path"]?.GetValue<string>() ?? "";
        return (relId, partPath);
    }

    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount)
    {
        contentType = null;
        byteCount = 0;
        try
        {
            var result = _session.Send("command", "extract-binary", new JsonObject
            {
                ["path"] = path,
                ["dest_path"] = destPath,
            })?.AsObject();
            if (result is null) return false;
            var found = result["found"]?.GetValue<bool>() ?? false;
            if (!found) return false;
            contentType = result["content_type"]?.GetValue<string>();
            byteCount = result["byte_count"]?.GetValue<long>() ?? 0;
            return true;
        }
        catch (CliException ex) when (ex.Code == "unsupported_command")
        {
            return false;
        }
    }

    public void Dispose() => _session.Dispose();

    // ----- Helpers ------------------------------------------------------

    private string SendViewString(string mode, int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var args = BuildViewArgs(mode, startLine, endLine, maxLines, cols);
        var result = _session.Send("command", "view", args);
        return result?.GetValue<string>() ?? "";
    }

    private JsonNode SendViewJson(string mode, int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var args = BuildViewArgs(mode, startLine, endLine, maxLines, cols);
        args["format"] = "json";
        var result = _session.Send("command", "view", args);
        return result ?? new JsonObject();
    }

    private static JsonObject BuildViewArgs(string mode, int? startLine, int? endLine, int? maxLines, HashSet<string>? cols)
    {
        var args = new JsonObject { ["mode"] = mode };
        if (startLine.HasValue) args["start"] = startLine.Value;
        if (endLine.HasValue) args["end"] = endLine.Value;
        if (maxLines.HasValue) args["max-lines"] = maxLines.Value;
        if (cols != null && cols.Count > 0) args["cols"] = string.Join(",", cols);
        return args;
    }

    private static JsonObject PropsToJson(Dictionary<string, string> properties)
    {
        var obj = new JsonObject();
        foreach (var kv in properties)
            obj[kv.Key] = kv.Value;
        return obj;
    }

    private static JsonObject PositionToJson(InsertPosition pos)
    {
        var obj = new JsonObject();
        if (pos.Index.HasValue) obj["index"] = pos.Index.Value;
        if (pos.After != null) obj["after"] = pos.After;
        if (pos.Before != null) obj["before"] = pos.Before;
        return obj;
    }
}
