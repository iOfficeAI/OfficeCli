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

    /// <summary>
    /// Extract cell address and span from multiple possible formats:
    /// 1. Modern: &lt;hp:cellAddr colAddr rowAddr/&gt; + &lt;hp:cellSpan colSpan rowSpan/&gt; (separate elements)
    /// 2. Combined: &lt;hp:cellAddr colAddr rowAddr colSpan rowSpan/&gt; (span attrs on cellAddr)
    /// 3. Legacy: attributes directly on &lt;hp:tc&gt;
    /// </summary>
    internal static (int Row, int Col, int RowSpan, int ColSpan) GetCellAddr(XElement tc)
    {
        var cellAddr = tc.Element(HwpxNs.Hp + "cellAddr");
        if (cellAddr != null)
        {
            int row = (int?)cellAddr.Attribute("rowAddr") ?? 0;
            int col = (int?)cellAddr.Attribute("colAddr") ?? 0;

            // Try separate <hp:cellSpan> element first (Hancom native format)
            var cellSpan = tc.Element(HwpxNs.Hp + "cellSpan");
            if (cellSpan != null)
            {
                return (row, col,
                    (int?)cellSpan.Attribute("rowSpan") ?? 1,
                    (int?)cellSpan.Attribute("colSpan") ?? 1);
            }

            // Fallback: span attrs on cellAddr itself
            return (row, col,
                (int?)cellAddr.Attribute("rowSpan") ?? 1,
                (int?)cellAddr.Attribute("colSpan") ?? 1);
        }

        // Fallback: attributes directly on <hp:tc>
        return (
            (int?)tc.Attribute("rowAddr") ?? 0,
            (int?)tc.Attribute("colAddr") ?? 0,
            (int?)tc.Attribute("rowSpan") ?? 1,
            (int?)tc.Attribute("colSpan") ?? 1
        );
    }
}
