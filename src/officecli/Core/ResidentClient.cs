// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Text;

namespace OfficeCli.Core;

public static class ResidentClient
{
    /// <summary>
    /// Check if a resident is running for this file (without consuming a connection).
    /// Just tries to connect briefly.
    /// </summary>
    public static bool TryConnect(string filePath, out string pipeName)
    {
        pipeName = ResidentServer.GetPipeName(filePath);
        try
        {
            using var client = new NamedPipeClientStream(".", pipeName + "-ping", PipeDirection.InOut);
            client.Connect(100); // 100ms timeout

            using var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true);
            using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

            // Ping to verify it's the right file
            var pingRequest = new ResidentRequest { Command = "__ping__" };
            var json = System.Text.Json.JsonSerializer.Serialize(pingRequest, ResidentJsonContext.Default.ResidentRequest);
            writer.WriteLine(json);

            var responseLine = reader.ReadLine();
            if (responseLine == null) return false;

            var response = System.Text.Json.JsonSerializer.Deserialize<ResidentResponse>(responseLine, ResidentJsonContext.Default.ResidentResponse);
            if (response == null) return false;

            // Stdout contains the file path when responding to ping
            if (string.IsNullOrEmpty(response.Stdout)) return false;
            var residentFilePath = Path.GetFullPath(response.Stdout);
            var requestedFilePath = Path.GetFullPath(filePath);
            return string.Equals(residentFilePath, requestedFilePath, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Send a command to the resident server in a single connection.
    /// Returns null if no resident is running or the file doesn't match.
    /// </summary>
    public static ResidentResponse? TrySend(string filePath, ResidentRequest request, int maxRetries = 2)
    {
        var pipeName = ResidentServer.GetPipeName(filePath);
        for (int attempt = 0; attempt <= maxRetries; attempt++)
        {
            try
            {
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(1000); // 1s timeout (was 200ms — too short under load)

                using var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true);
                using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

                var json = System.Text.Json.JsonSerializer.Serialize(request, ResidentJsonContext.Default.ResidentRequest);
                writer.WriteLine(json);

                var responseLine = reader.ReadLine();
                if (responseLine == null) continue;

                var response = System.Text.Json.JsonSerializer.Deserialize<ResidentResponse>(responseLine, ResidentJsonContext.Default.ResidentResponse);
                if (response != null) return response;
            }
            catch
            {
                if (attempt == maxRetries) return null;
                Thread.Sleep(50 * (attempt + 1)); // brief backoff before retry
            }
        }
        return null;
    }

    /// <summary>
    /// Send a close command to the resident server.
    /// </summary>
    public static bool SendClose(string filePath)
    {
        // Send close via the dedicated ping pipe (always responsive)
        var pipeName = ResidentServer.GetPipeName(filePath) + "-ping";
        try
        {
            using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
            client.Connect(200);

            using var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true);
            using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

            var request = new ResidentRequest { Command = "__close__" };
            var json = System.Text.Json.JsonSerializer.Serialize(request, ResidentJsonContext.Default.ResidentRequest);
            writer.WriteLine(json);

            var responseLine = reader.ReadLine();
            if (responseLine == null) return false;

            var response = System.Text.Json.JsonSerializer.Deserialize<ResidentResponse>(responseLine, ResidentJsonContext.Default.ResidentResponse);
            return response != null && response.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }
}
