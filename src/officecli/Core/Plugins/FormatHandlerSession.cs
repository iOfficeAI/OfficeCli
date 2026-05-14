// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// Owns a running format-handler plugin process and the named-pipe channel
/// used to talk to it. Per docs/plugin-protocol.md §2.3 / §5.3 / §6.
///
/// Lifecycle:
///   1. Main creates a random pipe name and a <see cref="NamedPipeServerStream"/> for it.
///   2. Main spawns the plugin with <c>open &lt;file&gt; --pipe &lt;name&gt;</c>.
///   3. The plugin connects to the pipe; main accepts.
///   4. Requests/responses flow as JSON-line envelopes (§6.3) for the
///      session's duration.
///   5. <see cref="Dispose"/> sends <c>close</c>, waits briefly for the plugin
///      to exit, and force-kills the process tree on timeout.
///
/// The session is single-threaded: callers must serialize access. The proxy
/// IDocumentHandler wraps each public method in a lock if used concurrently.
/// </summary>
internal sealed class FormatHandlerSession : IDisposable
{
    private readonly string _filePath;
    private readonly ResolvedPlugin _plugin;
    private readonly string _pipeName;
    private NamedPipeServerStream? _pipe;
    private StreamReader? _reader;
    private StreamWriter? _writer;
    private Process? _proc;
    private bool _disposed;
    private readonly object _ioLock = new();

    public ResolvedPlugin Plugin => _plugin;
    public string PipeName => _pipeName;

    public FormatHandlerSession(string filePath, ResolvedPlugin plugin)
    {
        _filePath = Path.GetFullPath(filePath);
        _plugin = plugin;
        _pipeName = $"officecli-fmt-{Guid.NewGuid():N}";
    }

    public void Start(int connectTimeoutMs = 15000)
    {
        // Create the server endpoint BEFORE spawning the plugin so it can
        // connect immediately on startup.
        _pipe = new NamedPipeServerStream(
            _pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: 1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous);

        var psi = new ProcessStartInfo
        {
            FileName = _plugin.ExecutablePath,
            ArgumentList = { "open", _filePath, "--pipe", _pipeName },
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        var selfPath = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(selfPath))
            psi.Environment["OFFICECLI_BIN"] = selfPath;

        _proc = Process.Start(psi)
            ?? throw new CliException($"Failed to start format-handler plugin '{_plugin.Manifest.Name}'.")
                { Code = "plugin_spawn_failed" };

        // Race the pipe connection against process exit and a wall-clock
        // timeout. Without the exit branch, an immediate plugin crash would
        // make the caller wait out the full connect timeout for nothing —
        // the real cause (exit code + stderr) is swallowed.
        var connectTask = _pipe.WaitForConnectionAsync();
        var exitTask = _proc.WaitForExitAsync();
        var timeoutTask = Task.Delay(connectTimeoutMs);
        var winner = Task.WhenAny(connectTask, exitTask, timeoutTask)
            .GetAwaiter().GetResult();

        if (winner == exitTask)
        {
            var stderr = "";
            try { stderr = _proc.StandardError.ReadToEnd(); } catch { }
            throw new CliException(
                $"Format-handler plugin '{_plugin.Manifest.Name}' exited (code {_proc.ExitCode}) before connecting to pipe: {Truncate(stderr, 500)}")
            { Code = "plugin_spawn_failed" };
        }
        if (winner == timeoutTask)
        {
            TryKill();
            throw new CliException(
                $"Format-handler plugin '{_plugin.Manifest.Name}' did not connect to pipe within {connectTimeoutMs}ms.")
            { Code = "plugin_pipe_timeout" };
        }
        // connectTask won — propagate any exception that may have completed it
        connectTask.GetAwaiter().GetResult();

        // Buffered reader/writer over the pipe. Raw newline-delimited UTF-8.
        // We don't reuse the resident-pipe helpers here because the pipe lives
        // for the whole session and we want async streams for clean shutdown.
        _reader = new StreamReader(_pipe, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, bufferSize: 8192, leaveOpen: true);
        _writer = new StreamWriter(_pipe, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), bufferSize: 8192, leaveOpen: true)
        {
            AutoFlush = true,
            NewLine = "\n",
        };
    }

    /// <summary>
    /// Send a request envelope and synchronously wait for the matching reply.
    /// Throws <see cref="CliException"/> on protocol error, pipe failure, or
    /// plugin-reported error responses.
    /// </summary>
    public JsonNode? Send(string msgType, string? command, JsonObject? args = null, JsonObject? props = null)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(FormatHandlerSession));
        if (_writer is null || _reader is null)
            throw new InvalidOperationException("Session not started.");

        var request = new JsonObject
        {
            ["protocol"] = 1,
            ["msg_type"] = msgType,
        };
        if (command is not null) request["command"] = command;
        if (args is not null) request["args"] = args;
        if (props is not null) request["props"] = props;

        lock (_ioLock)
        {
            try
            {
                _writer.WriteLine(request.ToJsonString());
                var line = _reader.ReadLine();
                if (line is null)
                    throw new CliException(
                        $"Format-handler plugin '{_plugin.Manifest.Name}' closed the pipe unexpectedly (no reply to {msgType}/{command ?? ""}).")
                    { Code = "plugin_pipe_closed" };

                var reply = JsonNode.Parse(line)?.AsObject()
                    ?? throw new CliException("Format-handler reply is not a JSON object.")
                        { Code = "protocol_mismatch" };

                var replyType = reply["msg_type"]?.GetValue<string>() ?? "";
                if (replyType == "ok")
                    return reply["result"];
                if (replyType == "error")
                {
                    var err = reply["error"]?.AsObject();
                    var code = err?["code"]?.GetValue<string>() ?? "plugin_error";
                    var msg = err?["message"]?.GetValue<string>() ?? "(no message)";
                    throw new CliException(
                        $"Format-handler plugin '{_plugin.Manifest.Name}' reported error on {command ?? msgType}: {msg}")
                    { Code = code };
                }
                throw new CliException(
                    $"Format-handler plugin '{_plugin.Manifest.Name}' replied with unknown msg_type '{replyType}'.")
                { Code = "protocol_mismatch" };
            }
            catch (IOException ex)
            {
                throw new CliException(
                    $"Format-handler plugin '{_plugin.Manifest.Name}' pipe I/O failed: {ex.Message}", ex)
                { Code = "plugin_pipe_io" };
            }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        // Ask the plugin to shut down cleanly. If it doesn't acknowledge
        // within a short window we drop through to stream disposal and the
        // hard-kill timer below — a hung plugin must not be allowed to wedge
        // Dispose() indefinitely.
        try
        {
            if (_writer is not null && _reader is not null && _pipe?.IsConnected == true)
            {
                var close = new JsonObject
                {
                    ["protocol"] = 1,
                    ["msg_type"] = "close",
                };
                _writer.WriteLine(close.ToJsonString());
                using var ackCts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
                try
                {
                    _reader.ReadLineAsync(ackCts.Token).AsTask().GetAwaiter().GetResult();
                }
                catch { /* ack timeout / EOF / IO — proceed to forced teardown */ }
            }
        }
        catch { /* shutting down; ignore */ }

        try { _reader?.Dispose(); } catch { }
        try { _writer?.Dispose(); } catch { }
        try { _pipe?.Dispose(); } catch { }

        if (_proc is not null)
        {
            try
            {
                if (!_proc.WaitForExit(5000))
                    TryKill();
            }
            catch { TryKill(); }
            try { _proc.Dispose(); } catch { }
        }
    }

    private void TryKill()
    {
        try { _proc?.Kill(entireProcessTree: true); } catch { }
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";
}
