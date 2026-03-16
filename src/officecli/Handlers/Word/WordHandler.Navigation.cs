// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Navigation ====================

    private DocumentNode GetRootNode(int depth)
    {
        var node = new DocumentNode { Path = "/", Type = "document" };
        var children = new List<DocumentNode>();

        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/body",
                Type = "body",
                ChildCount = mainPart.Document.Body.ChildElements.Count
            });
        }

        if (mainPart?.StyleDefinitionsPart != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/styles",
                Type = "styles",
                ChildCount = mainPart.StyleDefinitionsPart.Styles?.ChildElements.Count ?? 0
            });
        }

        int headerIdx = 0;
        if (mainPart?.HeaderParts != null)
        {
            foreach (var _ in mainPart.HeaderParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/header[{headerIdx + 1}]",
                    Type = "header"
                });
                headerIdx++;
            }
        }

        int footerIdx = 0;
        if (mainPart?.FooterParts != null)
        {
            foreach (var _ in mainPart.FooterParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/footer[{footerIdx + 1}]",
                    Type = "footer"
                });
                footerIdx++;
            }
        }

        if (mainPart?.NumberingDefinitionsPart != null)
        {
            children.Add(new DocumentNode { Path = "/numbering", Type = "numbering" });
        }

        node.Children = children;
        node.ChildCount = children.Count;
        return node;
    }

    private record PathSegment(string Name, int? Index);

    private static List<PathSegment> ParsePath(string path)
    {
        var segments = new List<PathSegment>();
        var parts = path.Trim('/').Split('/');

        foreach (var part in parts)
        {
            var bracketIdx = part.IndexOf('[');
            if (bracketIdx >= 0)
            {
                var name = part[..bracketIdx];
                var indexStr = part[(bracketIdx + 1)..^1];
                segments.Add(new PathSegment(name, int.Parse(indexStr)));
            }
            else
            {
                segments.Add(new PathSegment(part, null));
            }
        }

        return segments;
    }

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments)
    {
        if (segments.Count == 0) return null;

        var first = segments[0];
        OpenXmlElement? current = first.Name.ToLowerInvariant() switch
        {
            "body" => _doc.MainDocumentPart?.Document?.Body,
            "styles" => _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles,
            "header" => _doc.MainDocumentPart?.HeaderParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Header,
            "footer" => _doc.MainDocumentPart?.FooterParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Footer,
            "numbering" => _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering,
            "settings" => _doc.MainDocumentPart?.DocumentSettingsPart?.Settings,
            "comments" => _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments,
            _ => null
        };

        for (int i = 1; i < segments.Count && current != null; i++)
        {
            var seg = segments[i];
            IEnumerable<OpenXmlElement> children;
            if (current is Body body2 && (seg.Name.ToLowerInvariant() == "p" || seg.Name.ToLowerInvariant() == "tbl"))
            {
                // Flatten sdt containers when navigating body-level paragraphs/tables
                children = seg.Name.ToLowerInvariant() == "p"
                    ? GetBodyElements(body2).OfType<Paragraph>().Cast<OpenXmlElement>()
                    : GetBodyElements(body2).OfType<Table>().Cast<OpenXmlElement>();
            }
            else
            {
                children = seg.Name.ToLowerInvariant() switch
                {
                    "p" => current.Elements<Paragraph>().Cast<OpenXmlElement>(),
                    "r" => current.Descendants<Run>()
                        .Where(r => r.GetFirstChild<CommentReference>() == null)
                        .Cast<OpenXmlElement>(),
                    "tbl" => current.Elements<Table>().Cast<OpenXmlElement>(),
                    "tr" => current.Elements<TableRow>().Cast<OpenXmlElement>(),
                    "tc" => current.Elements<TableCell>().Cast<OpenXmlElement>(),
                    _ => current.ChildElements.Where(e => e.LocalName == seg.Name).Cast<OpenXmlElement>()
                };
            }

            current = seg.Index.HasValue
                ? children.ElementAtOrDefault(seg.Index.Value - 1)
                : children.FirstOrDefault();
        }

        return current;
    }

    private DocumentNode ElementToNode(OpenXmlElement element, string path, int depth)
    {
        var node = new DocumentNode { Path = path, Type = element.LocalName };

        if (element is Paragraph para)
        {
            node.Type = "paragraph";
            node.Text = GetParagraphText(para);
            node.Style = GetStyleName(para);
            node.Preview = node.Text?.Length > 50 ? node.Text[..50] + "..." : node.Text;
            node.ChildCount = GetAllRuns(para).Count();

            var pProps = para.ParagraphProperties;
            if (pProps != null)
            {
                if (pProps.Justification?.Val?.Value != null)
                    node.Format["alignment"] = pProps.Justification.Val.Value.ToString();
                if (pProps.Indentation?.FirstLine?.Value != null)
                    node.Format["firstLineIndent"] = pProps.Indentation.FirstLine.Value;
            }

            if (depth > 0)
            {
                int runIdx = 0;
                foreach (var run in GetAllRuns(para))
                {
                    node.Children.Add(ElementToNode(run, $"{path}/r[{runIdx + 1}]", depth - 1));
                    runIdx++;
                }
            }
        }
        else if (element is Run run)
        {
            node.Type = "run";
            node.Text = GetRunText(run);
            var font = GetRunFont(run);
            if (font != null) node.Format["font"] = font;
            var size = GetRunFontSize(run);
            if (size != null) node.Format["size"] = size;
            if (run.RunProperties?.Bold != null) node.Format["bold"] = true;
            if (run.RunProperties?.Italic != null) node.Format["italic"] = true;
        }
        else if (element is Hyperlink hyperlink)
        {
            node.Type = "hyperlink";
            node.Text = string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text));
            var relId = hyperlink.Id?.Value;
            if (relId != null)
            {
                try
                {
                    var rel = _doc.MainDocumentPart?.HyperlinkRelationships
                        .FirstOrDefault(r => r.Id == relId);
                    if (rel != null) node.Format["link"] = rel.Uri.ToString();
                }
                catch { }
            }
        }
        else if (element is Table table)
        {
            node.Type = "table";
            node.ChildCount = table.Elements<TableRow>().Count();
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            node.Format["cols"] = firstRow?.Elements<TableCell>().Count() ?? 0;

            if (depth > 0)
            {
                int rowIdx = 0;
                foreach (var row in table.Elements<TableRow>())
                {
                    var rowNode = new DocumentNode
                    {
                        Path = $"{path}/tr[{rowIdx + 1}]",
                        Type = "row",
                        ChildCount = row.Elements<TableCell>().Count()
                    };
                    if (depth > 1)
                    {
                        int cellIdx = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            var cellNode = new DocumentNode
                            {
                                Path = $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]",
                                Type = "cell",
                                Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                                ChildCount = cell.Elements<Paragraph>().Count()
                            };
                            if (depth > 2)
                            {
                                int pIdx = 0;
                                foreach (var cellPara in cell.Elements<Paragraph>())
                                {
                                    cellNode.Children.Add(ElementToNode(cellPara, $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]/p[{pIdx + 1}]", depth - 3));
                                    pIdx++;
                                }
                            }
                            rowNode.Children.Add(cellNode);
                            cellIdx++;
                        }
                    }
                    node.Children.Add(rowNode);
                    rowIdx++;
                }
            }
        }
        else
        {
            // Generic fallback: collect XML attributes and child val patterns
            foreach (var attr in element.GetAttributes())
                node.Format[attr.LocalName] = attr.Value;
            foreach (var child in element.ChildElements)
            {
                if (child.ChildElements.Count == 0)
                {
                    foreach (var attr in child.GetAttributes())
                    {
                        if (attr.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                        {
                            node.Format[child.LocalName] = attr.Value;
                            break;
                        }
                    }
                }
            }

            var innerText = element.InnerText;
            if (!string.IsNullOrEmpty(innerText))
                node.Text = innerText.Length > 200 ? innerText[..200] + "..." : innerText;
            if (string.IsNullOrEmpty(innerText))
            {
                var outerXml = element.OuterXml;
                node.Preview = outerXml.Length > 200 ? outerXml[..200] + "..." : outerXml;
            }

            node.ChildCount = element.ChildElements.Count;
            if (depth > 0)
            {
                var typeCounters = new Dictionary<string, int>();
                foreach (var child in element.ChildElements)
                {
                    var name = child.LocalName;
                    typeCounters.TryGetValue(name, out int idx);
                    node.Children.Add(ElementToNode(child, $"{path}/{name}[{idx + 1}]", depth - 1));
                    typeCounters[name] = idx + 1;
                }
            }
        }

        return node;
    }
}
