// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OfficeCli.Core;

/// <summary>
/// Shared helper for raw XML operations (read/write via XPath).
/// This enables AI to perform any OpenXML operation by manipulating XML directly.
/// </summary>
internal static class RawXmlHelper
{
    /// <summary>
    /// Perform a raw XML operation on a document part's root element.
    /// </summary>
    /// <param name="rootElement">The OpenXml root element (e.g. Document, Worksheet, Slide)</param>
    /// <param name="xpath">XPath expression to locate target element(s)</param>
    /// <param name="action">Operation: append, prepend, insertbefore, insertafter, replace, remove, setattr</param>
    /// <param name="xml">XML fragment for append/prepend/insert/replace, or attr=value for setattr</param>
    /// <returns>Number of elements affected</returns>
    public static int Execute(OpenXmlPartRootElement rootElement, string xpath, string action, string? xml)
    {
        var (xDoc, affected) = ExecuteOnXmlString(rootElement.OuterXml, xpath, action, xml);
        if (affected == 0) return 0;
        // Write modified XML back to the OpenXml element.
        // Propagate namespace declarations from root to direct child elements,
        // so that each child's ToString() produces self-contained XML with
        // all necessary namespace bindings (otherwise inherited namespaces are lost).
        var rootNsAttrs = xDoc.Root!.Attributes()
            .Where(a => a.IsNamespaceDeclaration).ToList();
        foreach (var child in xDoc.Root.Elements())
        {
            foreach (var nsAttr in rootNsAttrs)
            {
                if (child.Attribute(nsAttr.Name) == null)
                    child.SetAttributeValue(nsAttr.Name, nsAttr.Value);
            }
        }
        rootElement.InnerXml = string.Concat(xDoc.Root.Nodes().Select(n => n.ToString()));
        return affected;
    }

    /// <summary>
    /// Apply a raw XML operation directly on a part's stream (no SDK typed
    /// root needed). Used for arbitrary XML parts addressed by zip URI —
    /// sheet1.xml, footnotes.xml, customXml/item1.xml, etc.
    /// </summary>
    public static int Execute(OpenXmlPart part, string xpath, string action, string? xml)
    {
        var sourceXml = ReadPartXml(part);
        var (xDoc, affected) = ExecuteOnXmlString(sourceXml, xpath, action, xml);
        if (affected == 0) return 0;
        WritePartXml(part, xDoc.ToString(SaveOptions.DisableFormatting));
        return affected;
    }

    private static (XDocument xDoc, int affected) ExecuteOnXmlString(
        string sourceXml, string xpath, string action, string? xml)
    {
        var xDoc = XDocument.Parse(sourceXml);
        var nsManager = BuildNamespaceManager(xDoc);

        var nodes = xDoc.XPathSelectElements(xpath, nsManager).ToList();
        if (nodes.Count == 0)
        {
            Console.Error.WriteLine($"raw-set: XPath matched no elements: {xpath}");
            Console.Error.WriteLine("Hint: auto-registered namespace prefixes: " +
                string.Join(", ", CommonNamespaces.Keys.Order()) +
                ". No xmlns declarations needed in --xml fragments.");
            return (xDoc, 0);
        }

        int affected = 0;

        foreach (var node in nodes)
        {
            switch (action.ToLowerInvariant())
            {
                case "append":
                    if (xml == null) throw new ArgumentException("--xml is required for append");
                    var appendFragment = ParseFragment(xml, xDoc);
                    foreach (var el in appendFragment)
                        node.Add(el);
                    affected++;
                    break;

                case "prepend":
                    if (xml == null) throw new ArgumentException("--xml is required for prepend");
                    var prependFragment = ParseFragment(xml, xDoc);
                    foreach (var el in prependFragment.AsEnumerable().Reverse())
                        node.AddFirst(el);
                    affected++;
                    break;

                case "insertbefore" or "before":
                    if (xml == null) throw new ArgumentException("--xml is required for insertbefore");
                    var beforeFragment = ParseFragment(xml, xDoc);
                    foreach (var el in beforeFragment.AsEnumerable().Reverse())
                        node.AddBeforeSelf(el);
                    affected++;
                    break;

                case "insertafter" or "after":
                    if (xml == null) throw new ArgumentException("--xml is required for insertafter");
                    var afterFragment = ParseFragment(xml, xDoc);
                    foreach (var el in afterFragment)
                        node.AddAfterSelf(el);
                    affected++;
                    break;

                case "replace":
                    if (xml == null) throw new ArgumentException("--xml is required for replace");
                    var replaceFragment = ParseFragment(xml, xDoc);
                    node.ReplaceWith(replaceFragment.ToArray());
                    affected++;
                    break;

                case "remove" or "delete":
                    node.Remove();
                    affected++;
                    break;

                case "setattr":
                    if (xml == null) throw new ArgumentException("--xml is required for setattr (format: name=value)");
                    var eqIdx = xml.IndexOf('=');
                    if (eqIdx <= 0) throw new ArgumentException("setattr format: name=value");
                    var attrName = xml[..eqIdx];
                    var attrValue = xml[(eqIdx + 1)..];

                    // Handle namespaced attributes (e.g. w:val)
                    var colonIdx = attrName.IndexOf(':');
                    if (colonIdx > 0)
                    {
                        var prefix = attrName[..colonIdx];
                        var localName = attrName[(colonIdx + 1)..];
                        var ns = nsManager.LookupNamespace(prefix);
                        if (ns != null)
                            node.SetAttributeValue(XName.Get(localName, ns), attrValue);
                        else
                            node.SetAttributeValue(attrName, attrValue);
                    }
                    else
                    {
                        node.SetAttributeValue(attrName, attrValue);
                    }
                    affected++;
                    break;

                default:
                    throw new ArgumentException($"Unknown action: {action}. Supported: append, prepend, insertbefore, insertafter, replace, remove, setattr");
            }
        }

        return (xDoc, affected);
    }

    private static readonly Dictionary<string, string> CommonNamespaces = new()
    {
        ["w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ["r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ["a"] = "http://schemas.openxmlformats.org/drawingml/2006/main",
        ["p"] = "http://schemas.openxmlformats.org/presentationml/2006/main",
        ["x"] = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ["wp"] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        ["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ["c"] = "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ["xdr"] = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        ["wps"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        ["wp14"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        ["v"] = "urn:schemas-microsoft-com:vml",
    };

    private static List<XElement> ParseFragment(string xml, XDocument contextDoc)
    {
        // Collect namespace declarations from the context document
        var nsDict = new Dictionary<string, string>(CommonNamespaces);
        string? defaultNs = null;

        if (contextDoc.Root != null)
        {
            // Inherit the default namespace from the document root so that
            // unprefixed elements (e.g. <mergeCells>) are parsed into the
            // correct namespace (e.g. spreadsheetml) instead of empty namespace.
            var rootNsName = contextDoc.Root.Name.NamespaceName;
            if (!string.IsNullOrEmpty(rootNsName))
                defaultNs = rootNsName;

            foreach (var attr in contextDoc.Root.Attributes().Where(a => a.IsNamespaceDeclaration))
            {
                var prefix = attr.Name.Namespace == XNamespace.Xmlns ? attr.Name.LocalName : "";
                if (!string.IsNullOrEmpty(prefix))
                    nsDict[prefix] = attr.Value;
            }
        }

        var prefixedNs = string.Join(" ", nsDict.Select(kv => $"xmlns:{kv.Key}=\"{kv.Value}\""));
        var defaultNsDecl = !string.IsNullOrEmpty(defaultNs) ? $"xmlns=\"{defaultNs}\"" : "";
        var wrappedXml = $"<_root {defaultNsDecl} {prefixedNs}>{xml}</_root>";
        var parsed = XDocument.Parse(wrappedXml);
        return parsed.Root!.Elements().ToList();
    }

    private static XmlNamespaceManager BuildNamespaceManager(XDocument xDoc)
    {
        var nsManager = new XmlNamespaceManager(new NameTable());

        if (xDoc.Root != null)
        {
            foreach (var attr in xDoc.Root.Attributes().Where(a => a.IsNamespaceDeclaration))
            {
                var prefix = attr.Name.LocalName;
                if (attr.Name.Namespace == XNamespace.Xmlns)
                {
                    nsManager.AddNamespace(prefix, attr.Value);
                }
                else if (attr.Name == "xmlns")
                {
                    // Default namespace — assign a usable prefix
                    nsManager.AddNamespace("default", attr.Value);
                }
            }
        }

        // Ensure common OpenXML namespaces are available
        TryAddNamespace(nsManager, "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        TryAddNamespace(nsManager, "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        TryAddNamespace(nsManager, "a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        TryAddNamespace(nsManager, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
        TryAddNamespace(nsManager, "x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        TryAddNamespace(nsManager, "wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        TryAddNamespace(nsManager, "mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        TryAddNamespace(nsManager, "c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        TryAddNamespace(nsManager, "xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");

        return nsManager;
    }

    /// <summary>
    /// Validate an OpenXmlPackage and return structured errors.
    /// </summary>
    public static List<ValidationError> ValidateDocument(OpenXmlPackage package)
    {
        var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        // BUG-R6-08: documents containing w:numPicBullet can trip an NRE
        // inside SDK validation when one of its child accessors hits a
        // null. Materialise per-error with try/catch so a single problem
        // entry doesn't bring the whole `validate` command down. Surface
        // the exception as a synthetic ValidationError instead of
        // bubbling out as a process-level crash.
        var errors = new List<ValidationError>();
        IEnumerable<DocumentFormat.OpenXml.Validation.ValidationErrorInfo> raw;
        try
        {
            raw = validator.Validate(package);
        }
        catch (Exception ex)
        {
            errors.Add(new ValidationError(
                "ValidatorException",
                $"Validator threw before producing results: {ex.GetType().Name}: {ex.Message}",
                null, null));
            return errors;
        }
        // The IEnumerable is lazy — iterate with try/catch so one bad
        // error entry does not abort the rest.
        using var enumerator = raw.GetEnumerator();
        while (true)
        {
            DocumentFormat.OpenXml.Validation.ValidationErrorInfo? e = null;
            try
            {
                if (!enumerator.MoveNext()) break;
                e = enumerator.Current;
            }
            catch (NullReferenceException nre)
            {
                errors.Add(new ValidationError(
                    "ValidatorNullReference",
                    $"SDK validator hit a null while inspecting next error: {nre.Message}",
                    null, null));
                continue;
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "ValidatorException",
                    $"SDK validator threw while inspecting next error: {ex.GetType().Name}: {ex.Message}",
                    null, null));
                continue;
            }
            if (e == null) continue;
            try
            {
                errors.Add(new ValidationError(
                    e.ErrorType.ToString(),
                    e.Description,
                    e.Path?.XPath,
                    e.Part?.Uri.ToString()));
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "ValidatorException",
                    $"Failed to materialise validation error: {ex.GetType().Name}: {ex.Message}",
                    null, null));
            }
        }
        return errors;
    }

    private static void TryAddNamespace(XmlNamespaceManager nsManager, string prefix, string uri)
    {
        if (string.IsNullOrEmpty(nsManager.LookupNamespace(prefix)))
        {
            nsManager.AddNamespace(prefix, uri);
        }
    }

    // ==================== Zip-URI part lookup ====================
    //
    // Rule: any partPath ending in `.xml` is treated as a literal zip-internal
    // URI (e.g. `/xl/worksheets/sheet1.xml`, `/word/footnotes.xml`,
    // `/ppt/slides/slide1.xml`). We walk the entire part tree of the package
    // and match against `OpenXmlPart.Uri.OriginalString`.
    //
    // This supersedes the per-handler hand-curated alias tables, which could
    // never be complete (only covered global parts like /xl/workbook.xml).
    // Semantic paths (`/Sheet1`, `/workbook`, `/document`, `/header[1]`) still
    // route through the handler's own switch — only `.xml`-suffixed inputs
    // hit this lookup.

    /// <summary>
    /// Returns true if `partPath` should be resolved as a literal zip-internal
    /// URI rather than a semantic short name. Trims surrounding whitespace
    /// before checking the .xml suffix so a stray trailing space/newline
    /// doesn't silently route the input to the semantic-path dispatcher
    /// (which would emit a misleading "Available: ..." list).
    /// </summary>
    public static bool IsZipUriPath(string partPath) =>
        partPath != null && partPath.AsSpan().TrimEnd().EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Walk the entire part tree of a package and return the part whose
    /// `Uri.OriginalString` matches `partPath` (with leading-slash
    /// normalization). Returns null if no part matches.
    /// </summary>
    public static OpenXmlPart? FindPartByZipUri(OpenXmlPackage package, string partPath)
    {
        // Trim surrounding whitespace and normalize leading slash.
        partPath = partPath.Trim();
        var target = partPath.StartsWith('/') ? partPath : "/" + partPath;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        return Walk(package);

        OpenXmlPart? Walk(OpenXmlPartContainer container)
        {
            foreach (var rel in container.Parts)
            {
                var p = rel.OpenXmlPart;
                var uri = p.Uri.OriginalString;
                if (!seen.Add(uri)) continue;
                if (string.Equals(uri, target, StringComparison.OrdinalIgnoreCase))
                    return p;
                var nested = Walk(p);
                if (nested != null) return nested;
            }
            return null;
        }
    }

    /// <summary>
    /// Read a part's XML content as a string. Prefer the typed
    /// <see cref="OpenXmlPartRootElement"/> when available (preserves
    /// canonical SDK serialization); fall back to the underlying stream for
    /// untyped XML parts (e.g. CustomXml).
    ///
    /// Output omits the &lt;?xml ?&gt; prolog uniformly so that:
    ///   raw /workbook              (semantic path, typed OuterXml)
    ///   raw /xl/workbook.xml       (zip URI, typed OuterXml)
    ///   raw /customXml/item1.xml   (zip URI, untyped stream)
    /// all produce element-only output. The semantic short-name path has
    /// always done this; this method extends the convention to untyped
    /// parts so zip-URI calls don't randomly include the prolog depending
    /// on whether the SDK strongly-typed the target part.
    /// </summary>
    public static string ReadPartXml(OpenXmlPart part)
    {
        if (part.RootElement is OpenXmlPartRootElement root && root != null)
            return root.OuterXml;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        var content = reader.ReadToEnd();
        return StripXmlProlog(content);
    }

    private static string StripXmlProlog(string xml)
    {
        var s = xml.AsSpan().TrimStart();
        // BOM stripping (StreamReader usually handles this, but defensive)
        if (s.Length > 0 && s[0] == '﻿') s = s[1..];
        if (s.StartsWith("<?xml"))
        {
            var end = s.IndexOf("?>", StringComparison.Ordinal);
            if (end >= 0) s = s[(end + 2)..].TrimStart();
        }
        return s.ToString();
    }

    /// <summary>
    /// Write XML content into a part's stream, replacing prior contents.
    /// </summary>
    public static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }
}
