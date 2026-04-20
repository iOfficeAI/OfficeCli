// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using System.Xml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Helpers for embedding SVG images into OOXML documents.
///
/// OOXML requires a dual representation for SVG:
///   - The main a:blip/@r:embed points to a raster fallback (PNG) so older
///     Office versions render something.
///   - An a:blip/a:extLst/a:ext[@uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"]
///     contains an asvg:svgBlip whose r:embed points to the SVG part.
/// Modern Office (2016+) picks up the SVG; older versions fall back to the PNG.
/// </summary>
internal static class SvgImageHelper
{
    public const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    public const string SvgNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    public const string RelsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// 1×1 transparent PNG used as the default raster fallback when the
    /// caller does not supply an explicit fallback image. Modern Office
    /// renders the SVG directly; this placeholder is only what older
    /// viewers see.
    /// </summary>
    public static byte[] TransparentPng1x1 { get; } = new byte[]
    {
        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
        0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
        0x89,0x00,0x00,0x00,0x0D,0x49,0x44,0x41,
        0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
        0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
        0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
        0x42,0x60,0x82
    };

    /// <summary>
    /// Append (or replace) the Office SVG extension on a a:blip element,
    /// wiring it to the SVG image part's relationship id.
    /// </summary>
    public static void AppendSvgExtension(A.Blip blip, string svgRelId)
    {
        if (blip is null) throw new ArgumentNullException(nameof(blip));
        if (string.IsNullOrEmpty(svgRelId)) throw new ArgumentException("svgRelId required", nameof(svgRelId));

        var extList = blip.GetFirstChild<A.BlipExtensionList>();
        if (extList == null)
        {
            extList = new A.BlipExtensionList();
            blip.AppendChild(extList);
        }

        // Drop any pre-existing SVG extension first — we only want one.
        var existing = extList.Elements<A.BlipExtension>()
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, SvgExtensionUri, StringComparison.OrdinalIgnoreCase));
        existing?.Remove();

        var ext = new A.BlipExtension { Uri = SvgExtensionUri };
        var svgBlip = new DocumentFormat.OpenXml.OpenXmlUnknownElement(
            "asvg", "svgBlip", SvgNamespace);
        svgBlip.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
            "r", "embed", RelsNamespace, svgRelId));
        ext.AppendChild(svgBlip);
        extList.AppendChild(ext);
    }

    /// <summary>
    /// Return the r:embed rel id from the SVG extension on this blip, or
    /// null if the blip has no SVG extension.
    /// </summary>
    public static string? GetSvgRelId(A.Blip blip)
    {
        if (blip is null) return null;
        var extList = blip.GetFirstChild<A.BlipExtensionList>();
        if (extList == null) return null;
        foreach (var ext in extList.Elements<A.BlipExtension>())
        {
            if (!string.Equals(ext.Uri?.Value, SvgExtensionUri, StringComparison.OrdinalIgnoreCase))
                continue;
            // asvg:svgBlip is stored as a non-strongly-typed child; walk
            // descendants by LocalName to find the r:embed attribute.
            foreach (var child in ext.ChildElements)
            {
                if (child.LocalName != "svgBlip") continue;
                foreach (var attr in child.GetAttributes())
                {
                    if (attr.LocalName == "embed" && attr.NamespaceUri == RelsNamespace)
                        return attr.Value;
                }
            }
        }
        return null;
    }

    /// <summary>
    /// Try to parse pixel dimensions from an SVG document's &lt;svg&gt; root.
    /// Handles width/height attributes (px, pt, in, cm, mm, or bare numbers)
    /// and falls back to the viewBox's width/height. The stream position is
    /// restored on return. Returns null if parsing fails.
    /// </summary>
    public static (int Width, int Height)? TryGetSvgDimensions(Stream stream)
    {
        if (stream is null || !stream.CanSeek) return null;

        var startPos = stream.Position;
        try
        {
            stream.Position = 0;
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null,
                IgnoreWhitespace = true,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                CloseInput = false
            };
            using var reader = XmlReader.Create(stream, settings);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName != "svg") continue;

                var w = reader.GetAttribute("width");
                var h = reader.GetAttribute("height");
                var vb = reader.GetAttribute("viewBox");

                double? wd = ParseSvgLength(w);
                double? hd = ParseSvgLength(h);

                if ((wd is null || hd is null) && !string.IsNullOrEmpty(vb))
                {
                    var vbParts = vb.Split(new[] { ' ', ',', '\t', '\n', '\r' },
                        StringSplitOptions.RemoveEmptyEntries);
                    if (vbParts.Length == 4
                        && double.TryParse(vbParts[2], System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var vbW)
                        && double.TryParse(vbParts[3], System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var vbH))
                    {
                        wd ??= vbW;
                        hd ??= vbH;
                    }
                }

                if (wd is > 0 && hd is > 0)
                    return ((int)Math.Round(wd.Value), (int)Math.Round(hd.Value));
                return null;
            }
            return null;
        }
        catch
        {
            return null;
        }
        finally
        {
            try { stream.Position = startPos; } catch (IOException) { }
        }
    }

    private static readonly Regex _svgLengthRegex =
        new(@"^\s*([+-]?\d+(?:\.\d+)?)\s*(px|pt|in|cm|mm|pc|em|ex|%)?\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static double? ParseSvgLength(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var m = _svgLengthRegex.Match(value);
        if (!m.Success) return null;
        if (!double.TryParse(m.Groups[1].Value,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture,
            out var n))
            return null;
        var unit = m.Groups[2].Success ? m.Groups[2].Value.ToLowerInvariant() : "px";
        // Convert to pixels at 96dpi so aspect-ratio calculations in
        // ImageSource.TryGetDimensions land on the same scale as PNG/JPEG.
        return unit switch
        {
            "px" or "" => n,
            "pt" => n * 96.0 / 72.0,
            "in" => n * 96.0,
            "cm" => n * 96.0 / 2.54,
            "mm" => n * 96.0 / 25.4,
            "pc" => n * 16.0,
            "em" or "ex" => n * 16.0,
            "%" => null,  // needs viewport context — fall back to viewBox
            _ => n
        };
    }

    /// <summary>
    /// Sniff whether the byte stream looks like SVG XML. Used to recover
    /// when a caller resolved the source but didn't tell us the content
    /// type up front.
    /// </summary>
    public static bool LooksLikeSvg(byte[] bytes)
    {
        if (bytes is null || bytes.Length < 5) return false;
        // Skip leading whitespace + BOM.
        int i = 0;
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) i = 3;
        while (i < bytes.Length && (bytes[i] == ' ' || bytes[i] == '\t'
            || bytes[i] == '\r' || bytes[i] == '\n')) i++;
        // Look for <?xml or <svg or <!DOCTYPE svg within the first 256 bytes.
        var head = System.Text.Encoding.UTF8.GetString(bytes,
            i, Math.Min(256, bytes.Length - i)).ToLowerInvariant();
        return head.Contains("<svg") || (head.StartsWith("<?xml") && head.Contains("<svg"));
    }
}
