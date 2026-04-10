// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Xml.Linq;
using System.Text.Json.Nodes;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler : IDocumentHandler
{
    private readonly HwpxDocument _doc;
    private readonly string _filePath;
    private readonly bool _editable;
    private readonly Stream _stream;
    private bool _dirty;

    public HwpxHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _editable = editable;
        _stream = File.Open(filePath, FileMode.Open,
            editable ? FileAccess.ReadWrite : FileAccess.Read);
        var archive = new ZipArchive(_stream,
            editable ? ZipArchiveMode.Update : ZipArchiveMode.Read);
        _doc = LoadDocument(archive);
    }

    private static HwpxDocument LoadDocument(ZipArchive archive)
    {
        var doc = new HwpxDocument { Archive = archive };

        // Load header.xml — MUST use LoadAndNormalize (not XDocument.Load) for HWPML 2016 compat
        var headerEntry = archive.GetEntry("Contents/header.xml");
        if (headerEntry != null)
        {
            using var stream = headerEntry.Open();
            doc.Header = LoadAndNormalize(stream);
        }

        // Discover sections from content.hpf manifest
        var hpfEntry = archive.GetEntry("Contents/content.hpf");
        if (hpfEntry != null)
        {
            using var hpfStream = hpfEntry.Open();
            var hpf = LoadAndNormalize(hpfStream);
            var items = hpf.Descendants()
                .Where(e => e.Name.LocalName == "item"
                    && (e.Attribute("media-type")?.Value?.Contains("section") ?? false))
                .Select(e => e.Attribute("href")?.Value)
                .Where(h => h != null);

            int idx = 0;
            foreach (var href in items)
            {
                var entryPath = href!.StartsWith("Contents/") ? href : $"Contents/{href}";
                var entry = archive.GetEntry(entryPath);
                if (entry == null) continue;
                using var s = entry.Open();
                doc.Sections.Add(new HwpxSection
                {
                    Index = idx++,
                    Document = LoadAndNormalize(s)
                });
            }
        }

        // Fallback: try section0.xml, section1.xml, ...
        if (doc.Sections.Count == 0)
        {
            for (int i = 0; i < 100; i++)
            {
                var entry = archive.GetEntry($"Contents/section{i}.xml");
                if (entry == null) break;
                using var s = entry.Open();
                doc.Sections.Add(new HwpxSection
                {
                    Index = i,
                    Document = LoadAndNormalize(s)
                });
            }
        }

        if (doc.Sections.Count == 0)
            throw new InvalidOperationException("No sections found in HWPX document");

        return doc;
    }

    // --- Helper: read ZIP entry, normalize HWPML 2016→2011 namespaces, then parse ---
    private static XDocument LoadAndNormalize(Stream stream)
    {
        using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
        var raw = reader.ReadToEnd();
        foreach (var (old, canonical) in HwpxNs.LegacyToCanonical)
            raw = raw.Replace(old, canonical, StringComparison.Ordinal);
        return XDocument.Parse(raw);
    }

    public void Dispose()
    {
        _doc.Archive.Dispose();
        _stream.Dispose();
    }
}
