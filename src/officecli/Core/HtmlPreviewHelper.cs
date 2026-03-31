// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

/// <summary>
/// Shared helpers for HTML preview rendering across PowerPoint, Word, and Excel handlers.
/// </summary>
public static class HtmlPreviewHelper
{
    /// <summary>
    /// Load an OpenXML part by its relationship ID and return the content as a base64 data URI.
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
            return $"data:{contentType};base64,{Convert.ToBase64String(ms.ToArray())}";
        }
        catch
        {
            return null;
        }
    }
}
