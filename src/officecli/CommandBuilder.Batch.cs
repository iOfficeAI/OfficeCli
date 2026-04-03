// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildBatchCommand(Option<bool> jsonOption)
    {
        var batchFileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var batchInputOpt = new Option<FileInfo?>("--input") { Description = "JSON file containing batch commands. If omitted, reads from stdin" };
        var batchCommandsOpt = new Option<string?>("--commands") { Description = "Inline JSON array of batch commands (alternative to --input or stdin)" };
        var batchForceOpt = new Option<bool>("--force") { Description = "Continue execution even if a command fails (default: stop on first error)" };
        var batchCommand = new Command("batch", "Execute multiple commands from a JSON array (one open/save cycle)");
        batchCommand.Add(batchFileArg);
        batchCommand.Add(batchInputOpt);
        batchCommand.Add(batchCommandsOpt);
        batchCommand.Add(batchForceOpt);
        batchCommand.Add(jsonOption);

        batchCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(batchFileArg)!;
            var inputFile = result.GetValue(batchInputOpt);
            var inlineCommands = result.GetValue(batchCommandsOpt);
            var stopOnError = !result.GetValue(batchForceOpt);

            string jsonText;
            if (inlineCommands != null)
            {
                jsonText = inlineCommands;
            }
            else if (inputFile != null)
            {
                if (!inputFile.Exists)
                {
                    throw new FileNotFoundException($"Input file not found: {inputFile.FullName}");
                }
                jsonText = File.ReadAllText(inputFile.FullName);
            }
            else
            {
                // Read from stdin
                jsonText = Console.In.ReadToEnd();
            }

            // Pre-validate: check for unknown JSON fields before deserializing
            using var jsonDoc = System.Text.Json.JsonDocument.Parse(jsonText);
            if (jsonDoc.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                int ri = 0;
                foreach (var elem in jsonDoc.RootElement.EnumerateArray())
                {
                    if (elem.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        var unknown = new List<string>();
                        foreach (var prop in elem.EnumerateObject())
                        {
                            if (!BatchItem.KnownFields.Contains(prop.Name))
                                unknown.Add(prop.Name);
                        }
                        if (unknown.Count > 0)
                            throw new ArgumentException($"batch item[{ri}]: unknown field(s) {string.Join(", ", unknown.Select(f => $"\"{f}\""))}. Valid fields: command, parent, path, type, from, index, to, props, selector, text, mode, depth, part, xpath, action, xml");
                    }
                    ri++;
                }
            }

            var items = System.Text.Json.JsonSerializer.Deserialize<List<BatchItem>>(jsonText, BatchJsonContext.Default.ListBatchItem);
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("No commands found in input.");
            }

            // If a resident process is running, forward each command to it
            if (ResidentClient.TryConnect(file.FullName, out _))
            {
                var results = new List<BatchResult>();
                for (int bi = 0; bi < items.Count; bi++)
                {
                    var item = items[bi];
                    var req = item.ToResidentRequest();
                    req.Json = json;
                    var response = ResidentClient.TrySend(file.FullName, req);
                    if (response == null)
                    {
                        results.Add(new BatchResult { Index = bi, Success = false, Item = item, Error = "Failed to send to resident" });
                        if (stopOnError) break;
                        continue;
                    }
                    var success = response.ExitCode == 0;
                    results.Add(new BatchResult { Index = bi, Success = success, Item = !success ? item : null, Output = response.Stdout, Error = response.Stderr });
                    if (!success && stopOnError) break;
                }
                PrintBatchResults(results, json, items.Count);
                if (results.Any(r => !r.Success))
                    throw new InvalidOperationException($"Batch completed with {results.Count(r => !r.Success)} error(s)");
                return 0;
            }

            // Non-resident: open file once, execute all commands, save once
            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var batchResults = new List<BatchResult>();
            for (int bi = 0; bi < items.Count; bi++)
            {
                var item = items[bi];
                try
                {
                    var output = ExecuteBatchItem(handler, item, json);
                    batchResults.Add(new BatchResult { Index = bi, Success = true, Output = output });
                }
                catch (Exception ex)
                {
                    batchResults.Add(new BatchResult { Index = bi, Success = false, Item = item, Error = ex.Message });
                    if (stopOnError) break;
                }
            }
            PrintBatchResults(batchResults, json, items.Count);
            if (batchResults.Any(r => r.Success))
                NotifyWatch(handler, file.FullName, null);
            if (batchResults.Any(r => !r.Success))
                throw new InvalidOperationException($"Batch completed with {batchResults.Count(r => !r.Success)} error(s)");
            return 0;
        }, json); });

        return batchCommand;
    }
}
