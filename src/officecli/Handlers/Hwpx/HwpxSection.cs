// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Xml.Linq;

namespace OfficeCli.Handlers;

internal class HwpxSection
{
    public int Index { get; init; }
    public XDocument Document { get; init; } = null!;
    public XElement Root => Document.Root!;
    public List<XElement> Paragraphs => Root.Elements(HwpxNs.Hp + "p").ToList();
    public List<XElement> Tables => Root.Elements(HwpxNs.Hp + "tbl").ToList();
}
