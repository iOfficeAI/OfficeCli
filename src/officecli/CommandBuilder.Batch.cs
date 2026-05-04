// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildBatchCommand(Option<bool> jsonOption)
    {
        var batchFileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var batchInputOpt = new Option<FileInfo?>("--input") { Description = "JSON file containing batch commands. If omitted, reads from stdin" };
        var batchCommandsOpt = new Option<string?>("--commands") { Description = "Inline JSON array of batch commands (alternative to --input or stdin)" };
        // BUG-R4-BT2: default flipped to continue-on-error. A 700-command
        // dump replay losing 80% of the document on the first failing item
        // (e.g. one unsupported prop) is a far worse default than reporting
        // the failure and letting the rest of the batch through. Errors are
        // still surfaced individually (BatchResult.Error) and the overall
        // exit code is 1 if any item failed, so callers can still tell
        // "everything succeeded". `--stop-on-error` opts back into the
        // strict abort-on-first-failure flow for callers who depend on it.
        var batchForceOpt = new Option<bool>("--force") { Description = "Deprecated alias for the default continue-on-error mode (kept for compatibility)" };
        var batchStopOpt = new Option<bool>("--stop-on-error") { Description = "Abort the batch as soon as any command fails (default: continue and report per-item errors)" };
        var batchCommand = new Command("batch", "Execute multiple commands from a JSON array (one open/save cycle)");
        batchCommand.Add(batchFileArg);
        batchCommand.Add(batchInputOpt);
        batchCommand.Add(batchCommandsOpt);
        batchCommand.Add(batchForceOpt);
        batchCommand.Add(batchStopOpt);
        batchCommand.Add(jsonOption);

        batchCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(batchFileArg)!;
            var inputFile = result.GetValue(batchInputOpt);
            var inlineCommands = result.GetValue(batchCommandsOpt);
            // Default: continue on error. --stop-on-error flips it to strict.
            // --force still acts as the docx-protection bypass (matches set
            // --force semantics) but no longer doubles as the continue-on-
            // error switch.
            var stopOnError = result.GetValue(batchStopOpt);
            var forceFlag = result.GetValue(batchForceOpt);

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

            var items = System.Text.Json.JsonSerializer.Deserialize<List<BatchItem>>(jsonText, BatchJsonContext.Default.ListBatchItem) ?? new();
            // BUG-R40-B11: explicit null entries (e.g. `[null]`) deserialize
            // to a List<BatchItem> with a null slot and trip a NRE deeper in
            // ExecuteBatchItem. Reject up-front with a recognizable error
            // pointing at the offending index.
            for (int ni = 0; ni < items.Count; ni++)
            {
                if (items[ni] == null)
                    throw new ArgumentException(
                        $"batch item[{ni}] is null. Each entry must be a JSON object (e.g. {{\"command\":\"get\",\"path\":\"/\"}}).");
            }
            if (items.Count == 0)
            {
                PrintBatchResults(new List<BatchResult>(), json, 0);
                return 0;
            }

            // BUG-FUZZER-R6-03: batch must honour the same .docx document
            // protection check that `set` enforces. Without this, a protected
            // doc could be silently modified via
            //   officecli batch protected.docx --commands '[{"command":"set",...}]'
            // even though the same set issued via the standalone `set` command
            // would be rejected. We piggy-back on `--force` (which already
            // means "ignore safety guards" for the continue-on-error path) so
            // agents that need to override protection use the same flag they
            // already know from `set --force`.
            // CONSISTENCY(docx-protection): if you change the protection
            // semantics, also update CommandBuilder.Set.cs at the matching
            // CheckDocxProtection call site.
            var force = forceFlag;
            if (!force && file.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var batchItem in items)
                {
                    // Only mutation commands need the protection gate. Read
                    // commands (get/query/view) are unaffected by document
                    // protection — protection blocks writes, not reads.
                    var cmdLower = (batchItem.Command ?? "").ToLowerInvariant();
                    if (cmdLower is not ("set" or "add" or "remove" or "raw-set"))
                        continue;
                    // Property-bag protection-changing op is its own escape
                    // hatch (mirrors set's isProtectionChange exemption).
                    if (batchItem.Props != null && batchItem.Props.Keys.Any(k =>
                        k.Equals("protection", StringComparison.OrdinalIgnoreCase)))
                        continue;
                    var path = batchItem.Path ?? "";
                    var rc = CheckDocxProtection(file.FullName, path, json);
                    if (rc != 0) return rc;
                }
            }

            // If a resident process is running, send the entire batch as a
            // single "batch" command so it executes in one open/save cycle
            // inside the resident process (same semantics as non-resident mode).
            if (ResidentClient.TryConnect(file.FullName, out _))
            {
                var req = new ResidentRequest
                {
                    Command = "batch",
                    Json = json,
                    Args =
                    {
                        ["batchJson"] = jsonText,
                        ["force"] = force.ToString(),
                        ["stopOnError"] = stopOnError.ToString()
                    }
                };
                // CONSISTENCY(resident-two-step): long connectTimeoutMs so the
                // batch waits for its turn in the main-pipe queue instead of
                // silently timing out under load. Matches TryResident in
                // CommandBuilder.cs.
                var response = ResidentClient.TrySend(file.FullName, req, maxRetries: 3, connectTimeoutMs: 30000);
                if (response == null)
                {
                    Console.Error.WriteLine($"Resident for {file.Name} is running but the batch could not be delivered (main pipe busy or unresponsive). Retry, or run 'officecli close {file.Name}' and try again.");
                    return 3;
                }
                // The resident returns the formatted batch output directly
                if (!string.IsNullOrEmpty(response.Stdout))
                    Console.Write(response.Stdout);
                if (!string.IsNullOrEmpty(response.Stderr))
                    Console.Error.Write(response.Stderr);
                return response.ExitCode;
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
            return batchResults.Any(r => !r.Success) ? 1 : 0;
        }, json); });

        return batchCommand;
    }
}
