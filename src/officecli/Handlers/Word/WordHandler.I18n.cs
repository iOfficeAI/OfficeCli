// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

// SCOPE: Word-only i18n write/read helpers. This file consolidates two
// duplicated patterns previously scattered across Set.cs, Set.Element.cs,
// Add.Text.cs, Add.Structure.cs, and Navigation.cs:
//
//   1) The RTL cascade — `direction=rtl` requires <w:bidi/> on pPr +
//      <w:rtl/> on the paragraph mark rPr + <w:rtl/> on every run rPr.
//      Word's UI writes all three; missing any of them produces a mixed-bidi
//      paragraph that renders incorrectly for Arabic/Hebrew fonts.
//
//   2) Complex-script (CS) run readback — font.cs / size.cs / bold.cs /
//      italic.cs were read at two sites in Navigation with subtly different
//      fallback semantics. ReadComplexScriptRunFormatting unifies them.
//
// DO NOT add: locale → font mapping (lives in Core/LocaleFontRegistry),
// HTML preview lang/CSS fallback (lives in HtmlPreview.* and there has only
// one call site each), themeFontLang stamping (lives in BlankDocCreator,
// single site). Those don't have duplication worth abstracting.
//
// Pptx/Excel handlers have similar patterns but are intentionally NOT
// covered here — wait until two handlers actually share, then promote to
// Core/. This file stays Word-only.
public partial class WordHandler
{
    /// <summary>
    /// Apply the full RTL cascade (<w:bidi/> + paragraph-mark <w:rtl/> +
    /// every run's <w:rtl/>) to <paramref name="paragraph"/>. Idempotent and
    /// reversible: pass <paramref name="rtl"/>=false to clear the cascade.
    ///
    /// <para>
    /// CONSISTENCY(rtl-cascade): a paragraph-level <w:bidi/> alone only flips
    /// layout (page side, mark anchor); it does NOT reverse the run-internal
    /// character order. Word's UI also writes <w:rtl/> on every run and on
    /// the paragraph mark when the user toggles paragraph direction — this
    /// helper mirrors that so a single direction=rtl produces a fully
    /// Arabic-correct paragraph. Used by all paragraph-level callers (Set,
    /// SetElement, Add header/footer, table cell).
    /// </para>
    ///
    /// <para>
    /// One deliberate exclusion: <c>StyleRunProperties</c> in
    /// Add.Structure.cs:498-500 stamps <w:rFonts> only and intentionally
    /// omits <w:rtl/> due to schema-order constraints there. That site stays
    /// hand-rolled — do not redirect through this helper.
    /// </para>
    /// </summary>
    private void ApplyDirectionCascade(Paragraph paragraph, bool rtl)
    {
        var pProps = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());

        if (rtl)
        {
            pProps.BiDi = new BiDi();
        }
        else
        {
            // R18-fuzz-2 + R19-fuzz-1/2: when ANY inherited source carries
            // bidi=true (enclosing section, paragraph-style chain, docDefaults,
            // numbering lvl pPr), simply removing pPr.bidi leaves the paragraph
            // inheriting RTL — the user's explicit ltr override would be
            // silently lost. Emit <w:bidi w:val="false"/> to override
            // inheritance. When no inherited bidi exists, just remove pPr.bidi
            // (canonical clean state).
            pProps.RemoveAllChildren<BiDi>();
            if (HasInheritedBidi(paragraph))
            {
                pProps.BiDi = new BiDi { Val = new OnOffValue(false) };
            }
        }

        var markRPr = pProps.ParagraphMarkRunProperties
            ?? EnsureParagraphMarkRunPropertiesInSchemaOrder(pProps);
        ApplyRunFormatting(markRPr, "direction", rtl ? "rtl" : "ltr");

        foreach (var run in paragraph.Descendants<Run>())
        {
            var rPr = EnsureRunProperties(run);
            ApplyRunFormatting(rPr, "direction", rtl ? "rtl" : "ltr");
        }
    }

    /// <summary>
    /// True iff <paramref name="paragraph"/> would inherit RTL from any
    /// source above its direct pPr.bidi: the enclosing section's sectPr,
    /// the linked paragraph-style basedOn chain, docDefaults pPrDefault,
    /// or its numbering lvl pPr. Used by direction=ltr handlers to decide
    /// whether to emit <w:bidi w:val="0"/> (cancel inheritance) or simply
    /// clear (no inherited RTL — canonical clean state).
    /// </summary>
    private bool HasInheritedBidi(Paragraph paragraph)
    {
        // Section
        var owningSect = FindOwningSectionProperties(paragraph);
        if (BidiOnOffOrDefaultTrue(owningSect?.GetFirstChild<BiDi>()) == true) return true;

        // Paragraph-style chain (basedOn walk)
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null && StyleChainHasBidi(styleId)) return true;

        // docDefaults pPrDefault.bidi
        var pPrDefault = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults
            ?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        if (BidiOnOffOrDefaultTrue(pPrDefault?.GetFirstChild<BiDi>()) == true) return true;

        // Numbering lvl pPr.bidi (R9-1 layer)
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        var numId = numPr?.NumberingId?.Val?.Value;
        var ilvl = numPr?.NumberingLevelReference?.Val?.Value;
        if (numId.HasValue && ilvl.HasValue)
        {
            var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            var inst = numbering?.Elements<NumberingInstance>()
                .FirstOrDefault(n => n.NumberID?.Value == numId.Value);
            var absId = inst?.AbstractNumId?.Val?.Value;
            var abs = absId.HasValue
                ? numbering!.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == absId.Value)
                : null;
            var lvl = abs?.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == ilvl.Value);
            var lvlBidi = lvl?.PreviousParagraphProperties?.GetFirstChild<BiDi>();
            if (BidiOnOffOrDefaultTrue(lvlBidi) == true) return true;
        }
        return false;
    }

    /// <summary>
    /// True iff the basedOn chain rooted at <paramref name="styleId"/>
    /// contains a style whose pPr.bidi resolves to true (CT_OnOff defaults
    /// true when no Val is set). Returns false on cycles, missing styles,
    /// or explicit bidi=false.
    /// </summary>
    private bool StyleChainHasBidi(string styleId)
    {
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        var styles = stylesPart?.Styles?.Elements<Style>().ToList();
        if (styles == null) return false;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        while (current != null && seen.Add(current))
        {
            var s = styles.FirstOrDefault(x => x.StyleId?.Value == current);
            if (s == null) return false;
            var bidi = s.StyleParagraphProperties?.GetFirstChild<BiDi>();
            var on = BidiOnOffOrDefaultTrue(bidi);
            if (on == true) return true;
            // Explicit false on a closer style does NOT cancel further-up
            // inheritance walking (Word's resolver picks the nearest explicit
            // value); but for our purposes, an explicit false anywhere in the
            // chain means the paragraph inheriting from that style does not
            // get RTL via this chain — short-circuit.
            if (on == false) return false;
            current = s.BasedOn?.Val?.Value;
        }
        return false;
    }

    private static bool? BidiOnOffOrDefaultTrue(BiDi? bidi)
    {
        if (bidi == null) return null;
        var on = TryReadOnOff(bidi.Val);
        // <w:bidi/> with no Val defaults to true under CT_OnOff.
        return on ?? true;
    }

    /// <summary>
    /// Insert a fresh <see cref="ParagraphMarkRunProperties"/> into
    /// <paramref name="pProps"/> at the schema-correct position. CT_PPrBase
    /// places rPr after the body of pPr children but before <c>sectPr</c> and
    /// <c>pPrChange</c>. Naively appending makes Word's validator reject the
    /// document when a pPrChange is already present (R18-bt-2).
    /// </summary>
    private static ParagraphMarkRunProperties EnsureParagraphMarkRunPropertiesInSchemaOrder(ParagraphProperties pProps)
    {
        var rPr = new ParagraphMarkRunProperties();
        // Insert before the first sectPr / pPrChange child if any; otherwise append.
        OpenXmlElement? successor = null;
        foreach (var child in pProps.ChildElements)
        {
            if (child is SectionProperties || child is ParagraphPropertiesChange)
            {
                successor = child;
                break;
            }
        }
        if (successor != null)
            successor.InsertBeforeSelf(rPr);
        else
            pProps.AppendChild(rPr);
        return rPr;
    }

    /// <summary>
    /// Read complex-script run formatting (<w:rFonts cs/>, <w:szCs/>,
    /// <w:bCs/>, <w:iCs/>) into <paramref name="format"/>. Mirrors the
    /// canonical keys font.cs / size.cs / bold.cs / italic.cs.
    ///
    /// <para>
    /// Two-arg form lets the paragraph readback site fall back from the
    /// first run's rPr to the paragraph-mark rPr (covers paragraphs that
    /// have CS flags on the mark but no runs yet). Run-level callers pass
    /// <paramref name="fallback"/>=null.
    /// </para>
    ///
    /// <para>
    /// Skips keys that already exist in <paramref name="format"/> so callers
    /// can layer this on top of other readers without overwriting.
    /// </para>
    /// </summary>
    private static void ReadComplexScriptRunFormatting(
        OpenXmlCompositeElement? primary,
        OpenXmlCompositeElement? fallback,
        IDictionary<string, object?> format)
    {
        // font.cs — only set by ApplyRunFormatting; falls under <w:rFonts>.
        var rFontsP = primary?.GetFirstChild<RunFonts>();
        var rFontsF = fallback?.GetFirstChild<RunFonts>();
        var fontCs = !string.IsNullOrEmpty(rFontsP?.ComplexScript?.Value)
            ? rFontsP!.ComplexScript!.Value
            : (!string.IsNullOrEmpty(rFontsF?.ComplexScript?.Value)
                ? rFontsF!.ComplexScript!.Value
                : null);
        if (fontCs != null && !format.ContainsKey("font.cs"))
            format["font.cs"] = fontCs;

        // size.cs — half-points, formatted as "Npt".
        var szCsEl = primary?.GetFirstChild<FontSizeComplexScript>()
            ?? fallback?.GetFirstChild<FontSizeComplexScript>();
        if (szCsEl?.Val?.Value is string szCsVal
            && int.TryParse(szCsVal, out var szCsHalfPt)
            && !format.ContainsKey("size.cs"))
        {
            format["size.cs"] = $"{szCsHalfPt / 2.0:0.##}pt";
        }

        // bold.cs / italic.cs — boolean flags.
        var bCsEl = primary?.GetFirstChild<BoldComplexScript>()
            ?? fallback?.GetFirstChild<BoldComplexScript>();
        if (bCsEl != null && !format.ContainsKey("bold.cs"))
            format["bold.cs"] = true;

        var iCsEl = primary?.GetFirstChild<ItalicComplexScript>()
            ?? fallback?.GetFirstChild<ItalicComplexScript>();
        if (iCsEl != null && !format.ContainsKey("italic.cs"))
            format["italic.cs"] = true;
    }
}
