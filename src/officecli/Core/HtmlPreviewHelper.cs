// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

/// <summary>
/// Shared helpers for HTML preview rendering across PowerPoint, Word, and Excel handlers.
/// </summary>
internal static class HtmlPreviewHelper
{
    /// <summary>
    /// Load an OpenXML part by its relationship ID and return the content as a base64 data URI.
    /// EMF/WMF images are automatically converted to PNG for browser compatibility.
    /// Returns null if the part cannot be found or read.
    /// </summary>
    public static string? PartToDataUri(OpenXmlPart parentPart, string relId)
    {
        try
        {
            var part = parentPart.GetPartById(relId);
            using var stream = part.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var contentType = part.ContentType ?? "image/png";

            // Convert EMF/WMF to PNG for browser rendering
            if (contentType.Contains("emf", StringComparison.OrdinalIgnoreCase) ||
                contentType.Contains("wmf", StringComparison.OrdinalIgnoreCase))
            {
                var pngUri = ConvertMetafileToPngDataUri(ms.ToArray());
                if (pngUri != null) return pngUri;
            }

            return $"data:{contentType};base64,{Convert.ToBase64String(ms.ToArray())}";
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Convert EMF/WMF metafile bytes to a PNG data URI using System.Drawing (Windows GDI+).
    /// Returns null if conversion fails.
    /// </summary>
    private static string? ConvertMetafileToPngDataUri(byte[] metafileBytes)
    {
        try
        {
            using var emfStream = new MemoryStream(metafileBytes);
            using var metafile = new System.Drawing.Imaging.Metafile(emfStream);

            int width = metafile.Width;
            int height = metafile.Height;
            if (width <= 0 || height <= 0) return null;

            // Scale up for readability: use 2x for small images
            int scale = (width < 400 || height < 300) ? 2 : 1;
            int bmpW = width * scale;
            int bmpH = height * scale;

            using var bitmap = new System.Drawing.Bitmap(bmpW, bmpH);
            bitmap.SetResolution(96, 96);
            using (var g = System.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(System.Drawing.Color.White);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(metafile, 0, 0, bmpW, bmpH);
            }

            using var pngStream = new MemoryStream();
            bitmap.Save(pngStream, System.Drawing.Imaging.ImageFormat.Png);
            return $"data:image/png;base64,{Convert.ToBase64String(pngStream.ToArray())}";
        }
        catch
        {
            return null;
        }
    }
}
