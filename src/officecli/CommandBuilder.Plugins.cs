// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Core.Plugins;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildPluginsCommand(Option<bool> jsonOption)
    {
        var pluginsCommand = new Command("plugins", "Manage and inspect installed plugins");
        pluginsCommand.Add(BuildPluginsListCommand(jsonOption));
        pluginsCommand.Add(BuildPluginsInfoCommand(jsonOption));
        return pluginsCommand;
    }

    private static Command BuildPluginsListCommand(Option<bool> jsonOption)
    {
        var cmd = new Command("list", "List plugins discoverable in the standard search paths");
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var plugins = PluginRegistry.EnumerateAll();

            if (json)
            {
                using var stream = new MemoryStream();
                using (var w = new Utf8JsonWriter(stream))
                {
                    w.WriteStartArray();
                    foreach (var p in plugins)
                    {
                        w.WriteStartObject();
                        w.WriteString("name", p.Manifest.Name);
                        w.WriteString("version", p.Manifest.Version);
                        w.WriteNumber("protocol", p.Manifest.Protocol);
                        w.WritePropertyName("kinds");
                        JsonSerializer.Serialize(w, p.Manifest.Kinds, PluginJsonContext.Default.ListString);
                        w.WritePropertyName("extensions");
                        JsonSerializer.Serialize(w, p.Manifest.Extensions, PluginJsonContext.Default.ListString);
                        if (p.Manifest.Tier is { } tier) w.WriteString("tier", tier);
                        w.WriteString("path", p.ExecutablePath);
                        w.WriteEndObject();
                    }
                    w.WriteEndArray();
                }
                Console.WriteLine(OutputFormatter.WrapEnvelope(Encoding.UTF8.GetString(stream.ToArray())));
                return 0;
            }

            if (plugins.Count == 0)
            {
                Console.WriteLine("No plugins installed.");
                Console.WriteLine("");
                Console.WriteLine("Plugins extend officecli to support additional formats (.doc, .hwpx, .pdf export, ...).");
                Console.WriteLine("See: docs/plugin-protocol.md for installation paths.");
                return 0;
            }

            // Plain-text table.
            var rows = plugins
                .Select(p => new
                {
                    Name = p.Manifest.Name,
                    Version = p.Manifest.Version,
                    Kinds = string.Join(",", p.Manifest.Kinds),
                    Exts = string.Join(",", p.Manifest.Extensions),
                    Path = p.ExecutablePath,
                })
                .ToList();

            int wName = Math.Max(4, rows.Max(r => r.Name.Length));
            int wVer = Math.Max(7, rows.Max(r => r.Version.Length));
            int wKinds = Math.Max(5, rows.Max(r => r.Kinds.Length));
            int wExts = Math.Max(11, rows.Max(r => r.Exts.Length));

            Console.WriteLine($"{"NAME".PadRight(wName)}  {"VERSION".PadRight(wVer)}  {"KINDS".PadRight(wKinds)}  {"EXTENSIONS".PadRight(wExts)}  PATH");
            foreach (var r in rows)
                Console.WriteLine($"{r.Name.PadRight(wName)}  {r.Version.PadRight(wVer)}  {r.Kinds.PadRight(wKinds)}  {r.Exts.PadRight(wExts)}  {r.Path}");

            return 0;
        }, json); });

        return cmd;
    }

    private static Command BuildPluginsInfoCommand(Option<bool> jsonOption)
    {
        var nameArg = new Argument<string>("name")
        {
            Description = "Plugin name or path to its executable",
        };

        var cmd = new Command("info", "Show the full manifest for a single plugin");
        cmd.Add(nameArg);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var target = result.GetValue(nameArg) ?? "";
            var resolved = ResolveByNameOrPath(target);
            if (resolved is null)
                throw new CliException($"Plugin not found: '{target}'")
                {
                    Code = "plugin_not_found",
                    Suggestion = "Run `officecli plugins list` to see installed plugins, or provide the absolute path to the plugin executable.",
                };

            // Re-read the manifest raw rather than re-serializing from our typed
            // class: this preserves any extra fields the plugin emits beyond
            // what PluginManifest knows about, so `plugins info` is faithful to
            // the plugin's actual --info output.
            using var p = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = resolved.ExecutablePath,
                    Arguments = "--info",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                }
            };
            p.Start();
            var rawManifest = p.StandardOutput.ReadToEnd();
            if (!p.WaitForExit(5000)) { try { p.Kill(true); } catch { } }

            if (json)
            {
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["path"] = resolved.ExecutablePath,
                    ["manifest"] = System.Text.Json.Nodes.JsonNode.Parse(rawManifest),
                };
                Console.WriteLine(OutputFormatter.WrapEnvelope(envelope.ToJsonString()));
                return 0;
            }

            Console.WriteLine($"Path: {resolved.ExecutablePath}");
            Console.WriteLine();
            // Pretty-print the manifest JSON via Utf8JsonWriter (AOT-safe,
            // unlike JsonSerializer.Serialize(JsonElement) which trips IL2026).
            try
            {
                using var doc = JsonDocument.Parse(rawManifest);
                using var ms = new MemoryStream();
                using (var w = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true }))
                    doc.RootElement.WriteTo(w);
                Console.WriteLine(Encoding.UTF8.GetString(ms.ToArray()));
            }
            catch
            {
                Console.WriteLine(rawManifest);
            }
            return 0;
        }, json); });

        return cmd;
    }

    private static ResolvedPlugin? ResolveByNameOrPath(string target)
    {
        // Path mode: absolute or relative path that exists.
        if (target.Contains(Path.DirectorySeparatorChar) || target.Contains(Path.AltDirectorySeparatorChar) || File.Exists(target))
        {
            var full = Path.GetFullPath(target);
            if (File.Exists(full) && PluginRegistry.TryReadManifest(full, out var m))
                return new ResolvedPlugin(full, m);
            return null;
        }

        // Name mode: search the full enumeration for a manifest whose name matches.
        var all = PluginRegistry.EnumerateAll();
        return all.FirstOrDefault(p =>
            string.Equals(p.Manifest.Name, target, StringComparison.OrdinalIgnoreCase));
    }
}
