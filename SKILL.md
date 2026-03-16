---
name: officecli
description: Create, analyze, proofread, and modify Office documents (.docx, .xlsx, .pptx) using the officecli CLI tool. Use when the user wants to create, inspect, check formatting, find issues, add charts, or modify Office documents.
---

# officecli

AI-friendly CLI for .docx, .xlsx, .pptx.

**First, check if officecli is available:**
```bash
officecli --version
```
If the command is not found, install it:
```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCli/main/install.sh | bash
```
For Windows (PowerShell):
```powershell
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCli/main/install.ps1 | iex
```

**Strategy:** L1 (read) → L2 (DOM edit) → L3 (raw XML). Always prefer higher layers. Add `--json` for structured output.

**Performance:** Use `open <file>`/`close <file>` when running multiple commands on the same file to avoid repeated loading.

**Batch:** For 3+ operations, plan all changes first, generate a single script (work backwards on inserts), execute once.

---

## L1: Create, Read & Inspect

```bash
officecli create <file>          # create blank .docx/.xlsx/.pptx (type inferred from extension)
officecli view <file> outline|stats|issues|text|annotated [--start N --end N] [--max-lines N] [--cols A,B]
officecli get <file> '<path>' --depth 2 [--json]
officecli query <file> 'paragraph[style=Normal] > run[font!=宋体]'
```

**get paths:** Any XML localName works. Common paths: `/body/p[3]`, `/Sheet1/A1`, `/slide[1]/shape[1]`, `/slide[1]/table[1]/tr[1]/tc[1]`, `/slide[1]/placeholder[title]`. Use `--depth N` to expand children.

**view modes:** `outline` (structure), `stats` (statistics with style inheritance), `issues` (`--type format|content|structure`, `--limit N`), `text` (plain with line numbers), `annotated` (with formatting)

**query selectors:** `[attr=value]`, `[attr!=value]`, `:contains("text")`, `:empty`, `:has(formula)`, `:no-alt`. Built-in types: `paragraph`, `run`, `picture`, `equation`, `cell`, `table`, `placeholder`. Falls back to generic XML element name (e.g. `wsp`, `a:ln`, `srgbClr[val=0070C0]`).

For large documents, ALWAYS use `--max-lines` or `--start`/`--end` to limit output.

---

## L2: DOM Operations

### set — `officecli set <file> <path> --prop key=value [--prop ...]`

**Any XML attribute is settable via element path** (found via `get --depth N`) — even attributes not currently present. Use this before reaching for L3.

| Target | Path example | Properties |
|--------|-------------|------------|
| Word run | `/body/p[3]/r[1]` | `text`, `font`, `size`, `bold`, `italic`, `caps`, `smallCaps`, `strike`, `dstrike`, `vanish`, `outline`, `shadow`, `emboss`, `imprint`, `noProof`, `rtl`, `highlight`, `color`, `underline`, `shd`, ... |
| Word run image | `/body/p[5]/r[1]` | `alt`, `width`, `height` (cm/in/pt/px), ... |
| Word paragraph | `/body/p[3]` | `style`, `alignment`, `firstLineIndent`, `shd`, `spaceBefore`, `spaceAfter`, `lineSpacing`, `numId`, `numLevel`/`ilvl`, `listStyle`(=bullet\|numbered), ... |
| Word table cell | `/body/tbl[1]/tr[1]/tc[1]` | `text`, `font`, `size`, `bold`, `italic`, `color`, `shd`, `alignment`, `valign`(top\|center\|bottom), `width`, `vmerge`, `gridspan`, ... |
| Word table row | `/body/tbl[1]/tr[1]` | `height`, `header`(bool), ... |
| Word table | `/body/tbl[1]` | `alignment`, `width`, ... |
| Word document | `/` | `defaultFont`, `pageBackground`, `pageWidth`, `pageHeight`, `marginTop/Bottom/Left/Right`, ... |
| Excel cell | `/Sheet1/A1` | `value`, `formula`, `clear`, `font.bold/italic/strike/underline/color/size/name`, `fill`(hex RGB), `alignment.horizontal/vertical/wrapText`, `numFmt`, `link`(url/none) |
| PPT slide | `/slide[1]` | `background`(hex/gradient/`image:path`/none), `transition`(fade/wipe/push/…), `advanceTime`(ms), `advanceClick`(bool) |
| PPT notes | `/slide[1]/notes` | `text` |
| PPT shape | `/slide[1]/shape[1]` | `text`(`\n` for line breaks), `name`, `font`, `size`, `bold`, `italic`, `underline`, `strikethrough`, `color`, `fill`, `gradient`(e.g. `FF0000-0000FF-90`), `line`, `lineWidth`, `lineDash`, `preset`, `margin`, `align`, `valign`, `list`(bullet/numbered/alpha/roman), `lineSpacing`, `spaceBefore`, `spaceAfter`, `rotation`, `opacity`, `autoFit`, `shadow`(hex), `glow`(hex), `reflection`(true/none), `animation`(effect-class-ms), `link`(url/none), `x`, `y`, `width`, `height` |
| PPT table | `/slide[1]/table[1]` | `x`, `y`, `width`, `height`, `name`, ... |
| PPT table row | `/slide[1]/table[1]/tr[1]` | `height`; other props (text, bold, fill, ...) apply to all cells in row |
| PPT table cell | `/slide[1]/table[1]/tr[1]/tc[1]` | `text`, `font`, `size`, `bold`, `italic`, `color`, `fill`, `align`, `gridspan`/`colspan`, `rowspan`, `vmerge`, `hmerge` |
| PPT placeholder | `/slide[1]/placeholder[title]` | Same as shape. Types: `title`, `body`, `subtitle`, `date`, `footer`, `slidenum`. Auto-created from layout if missing. |

Composite props (`pBdr`, `tabs`, `lang`, `bdr`) → use L3 (`raw-set --action setattr`).

### add — `officecli add <file> <parent> --type <type> [--index N] [--prop ...]` or `--from <path>`

| Format | Types & props |
|--------|--------------|
| Word | `paragraph`(text,font,size,bold,style,alignment,...), `run`(text,font,size,bold,italic,...), `table`(rows,cols), `picture`(path,width,height,alt,...), `equation`(formula,mode), `comment`(text,author,initials,date,...) |
| Excel | `sheet`(name), `row`(cols), `cell`(ref,value,formula,...), `databar`(sqref,min,max,color,...) |
| PPT | `slide`(title,text), `shape`(text,font,size,name,bold,italic,underline,strikethrough,color,fill,gradient,line,lineWidth,lineDash,preset,margin,align,valign,list,lineSpacing,spaceBefore,spaceAfter,rotation,opacity,autoFit,shadow,glow,reflection,animation,link,x,y,width,height), `table`(rows,cols,x,y,width,height; cells: gridspan,rowspan,vmerge,hmerge), `picture`(path,width,height,x,y,alt), `equation`(formula,x,y,width,height) |

Dimensions: raw EMU or suffixed `cm`/`in`/`pt`/`px`. Equation formula: LaTeX subset. `--from <path>` clones an existing element (cross-part relationships handled automatically).

### move — `officecli move <file> <path> [--to <parent>] [--index N]`

### remove — `officecli remove <file> '<path>'`

---

## L3: Raw XML

Use for charts, borders, or any structure L2 cannot express. **No xmlns needed** — prefixes auto-registered: `w`, `a`, `p`, `x`, `r`, `c`, `xdr`, `wp`, `wps`, `mc`, `wp14`, `v`

```bash
officecli raw <file> /document                     # Word: /styles, /numbering, /settings, /header[N], /footer[N]
officecli raw <file> /Sheet1 --start 1 --end 100 --cols A,B   # Excel: /styles, /sharedstrings, /<Sheet>/drawing, /<Sheet>/chart[N]
officecli raw <file> /slide[1]                     # PPT: /presentation, /slideMaster[N], /slideLayout[N]
officecli raw-set <file> /document --xpath "//w:body/w:p[1]" --action replace --xml '<w:p>...</w:p>'
# actions: append, prepend, insertbefore, insertafter, replace, remove, setattr
officecli add-part <file> /Sheet1 --type chart     # returns relId for use with raw-set
officecli add-part <file> / --type header|footer   # Word only
```

---

## Notes

- Paths are **1-based** (XPath convention), quote brackets: `'/body/p[3]'`
- `--index` is **0-based** (array convention): `--index 0` = first position
- After modifications, verify with `validate` and/or `view issues`
- `raw-set`/`add-part` auto-validate after execution
- `view stats`/`annotated` resolve style inheritance (docDefaults → basedOn → direct)
