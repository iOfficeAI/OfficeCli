// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Populate Format dictionary on the root DocumentNode with document-level settings.
    /// Called from GetRootNode().
    /// </summary>
    private void PopulateDocSettings(DocumentNode node)
    {
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>()
            ?? _doc.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().LastOrDefault();

        // ==================== DocGrid ====================
        if (sectPr != null)
        {
            var grid = sectPr.GetFirstChild<DocGrid>();
            if (grid != null)
            {
                if (grid.Type?.Value != null)
                    node.Format["docGrid.type"] = grid.Type.InnerText;
                if (grid.LinePitch?.Value != null)
                    node.Format["docGrid.linePitch"] = grid.LinePitch.Value;
                if (grid.CharacterSpace?.Value != null)
                    node.Format["docGrid.charSpace"] = grid.CharacterSpace.Value;
            }

            // ==================== Columns ====================
            // CONSISTENCY(root-vs-section-readback): canonical column keys must match
            // BuildSectionNode so `get /` and `get /section[N]` round-trip the
            // same key names. Schema canonical: `columns`, `columnSpace` (with
            // legacy aliases `columns.count`, `columns.space` accepted on
            // Add/Set, dropped on Get per CLAUDE.md "Get should normalize to
            // the canonical key only"). EqualWidth / separator have no schema
            // canonical alias yet so they keep the dotted form.
            var cols = sectPr.GetFirstChild<Columns>();
            if (cols != null)
            {
                if (cols.ColumnCount?.Value != null)
                    node.Format["columns"] = (int)cols.ColumnCount.Value;
                if (cols.Space?.Value != null && uint.TryParse(cols.Space.Value, out var colSpaceTwips))
                    node.Format["columnSpace"] = FormatTwipsToCm(colSpaceTwips);
                if (cols.EqualWidth?.Value != null)
                    node.Format["columns.equalWidth"] = cols.EqualWidth.Value;
                if (cols.Separator?.Value == true)
                    node.Format["columns.separator"] = true;
            }

            // ==================== SectionType ====================
            var sectType = sectPr.GetFirstChild<SectionType>();
            if (sectType?.Val?.Value != null)
                node.Format["section.type"] = sectType.Val.InnerText;
        }

        // ==================== CJK Layout (from DocDefaults ParagraphProperties) ====================
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        var pPrBase = stylesPart?.Styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        if (pPrBase != null)
        {
            if (pPrBase.AutoSpaceDE != null)
                node.Format["autoSpaceDE"] = pPrBase.AutoSpaceDE.Val?.Value ?? true;
            if (pPrBase.AutoSpaceDN != null)
                node.Format["autoSpaceDN"] = pPrBase.AutoSpaceDN.Val?.Value ?? true;
            if (pPrBase.Kinsoku != null)
                node.Format["kinsoku"] = pPrBase.Kinsoku.Val?.Value ?? true;
            if (pPrBase.OverflowPunctuation != null)
                node.Format["overflowPunct"] = pPrBase.OverflowPunctuation.Val?.Value ?? true;
        }

        if (settings == null) return;

        // ==================== CharacterSpacingControl ====================
        var charSpacing = settings.GetFirstChild<CharacterSpacingControl>();
        if (charSpacing?.Val?.Value != null)
            node.Format["charSpacingControl"] = charSpacing.Val.InnerText;

        // ==================== Print / Display ====================
        if (settings.GetFirstChild<DisplayBackgroundShape>() != null)
            node.Format["displayBackgroundShape"] = true;
        if (settings.GetFirstChild<DoNotDisplayPageBoundaries>() != null)
            node.Format["doNotDisplayPageBoundaries"] = true;
        if (settings.GetFirstChild<PrintFormsData>() != null)
            node.Format["printFormsData"] = true;
        if (settings.GetFirstChild<PrintPostScriptOverText>() != null)
            node.Format["printPostScriptOverText"] = true;
        if (settings.GetFirstChild<PrintFractionalCharacterWidth>() != null)
            node.Format["printFractionalCharacterWidth"] = true;
        if (settings.GetFirstChild<RemovePersonalInformation>() != null)
            node.Format["removePersonalInformation"] = true;
        if (settings.GetFirstChild<RemoveDateAndTime>() != null)
            node.Format["removeDateAndTime"] = true;

        // ==================== Font Embedding ====================
        if (settings.GetFirstChild<EmbedTrueTypeFonts>() != null)
            node.Format["embedFonts"] = true;
        if (settings.GetFirstChild<EmbedSystemFonts>() != null)
            node.Format["embedSystemFonts"] = true;
        if (settings.GetFirstChild<SaveSubsetFonts>() != null)
            node.Format["saveSubsetFonts"] = true;

        // ==================== Layout Flags ====================
        if (settings.GetFirstChild<MirrorMargins>() != null)
            node.Format["mirrorMargins"] = true;
        if (settings.GetFirstChild<GutterAtTop>() != null)
            node.Format["gutterAtTop"] = true;
        if (settings.GetFirstChild<BookFoldPrinting>() != null)
            node.Format["bookFoldPrinting"] = true;
        if (settings.GetFirstChild<BookFoldReversePrinting>() != null)
            node.Format["bookFoldReversePrinting"] = true;
        var bookFoldSheets = settings.GetFirstChild<BookFoldPrintingSheets>();
        if (bookFoldSheets?.Val?.Value != null)
            node.Format["bookFoldPrintingSheets"] = (int)bookFoldSheets.Val.Value;
        if (settings.GetFirstChild<EvenAndOddHeaders>() != null)
            node.Format["evenAndOddHeaders"] = true;
        if (settings.GetFirstChild<AutoHyphenation>() != null)
            node.Format["autoHyphenation"] = true;
        var defTabStop = settings.GetFirstChild<DefaultTabStop>();
        if (defTabStop?.Val?.Value != null)
            node.Format["defaultTabStop"] = FormatTwipsToCm((uint)defTabStop.Val.Value);

        // ==================== Theme Font Languages ====================
        // CONSISTENCY(locale-readback): `--locale ar-SA` writes
        // settings/themeFontLang on Set; `Get /` must surface the same
        // value so locale round-trips. Mirror R5-1 run-level lang.* keys
        // (lang.latin / lang.ea / lang.cs) at doc-level. The bare
        // `locale` key is the bidi-priority single-string view (the
        // value Set most recently received via --locale); when only
        // val/eastAsia are set, fall back to those.
        var themeFontLang = settings.GetFirstChild<ThemeFontLanguages>();
        if (themeFontLang != null)
        {
            if (themeFontLang.Val?.Value != null)
                node.Format["lang.latin"] = themeFontLang.Val.Value;
            if (themeFontLang.EastAsia?.Value != null)
                node.Format["lang.ea"] = themeFontLang.EastAsia.Value;
            if (themeFontLang.Bidi?.Value != null)
                node.Format["lang.cs"] = themeFontLang.Bidi.Value;
            // Single-string `locale` view: bidi takes priority (matches
            // how --locale ar-SA writes <w:themeFontLang w:bidi="ar-SA"/>),
            // then val (Latin), then eastAsia.
            var localeStr = themeFontLang.Bidi?.Value
                ?? themeFontLang.Val?.Value
                ?? themeFontLang.EastAsia?.Value;
            if (localeStr != null)
                node.Format["locale"] = localeStr;
        }
    }
}
