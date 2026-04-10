// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Xml.Linq;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    private XElement? FindCharPr(string idRef)
    {
        return _doc.Header?.Root?
            .Descendants(HwpxNs.Hh + "charPr")
            .FirstOrDefault(e => e.Attribute("id")?.Value == idRef);
    }

    private static double GetFontSizePt(XElement charPr)
        => ((double?)charPr.Attribute("height") ?? 1000) / 100.0;
}
