# Mosaic Corporate + Sunset Circle

## Style Overview

morph-template v31 — Mosaic Corporate + Sunset Circle。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Navy | #282E3A | 主/辅色块、背景或强调 |
| Navy2 | #1A2030 | 主/辅色块、背景或强调 |
| Orange | #F0980A | 主/辅色块、背景或强调 |
| Orange2 | #C07808 | 主/辅色块、背景或强调 |
| Sky | #A8C8DC | 主/辅色块、背景或强调 |
| Sky2 | #7AAAC0 | 主/辅色块、背景或强调 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI Black 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Segoe UI 18-28pt | 分节标题与过渡句 |
| Body | Segoe UI 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **Technique**: rect mosaic partition
- **Technique**: gradient ellipse as hero visual
- **Technique**: data blocks with %
- **Technique**: Bold morph: !!sun (gradient circle) travels across slides changing size+position
- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `v31_build.py`.

**Recommended slides to read for understanding core design techniques**:

- **Techniques: rect mosaic partition, gradient ellipse as hero visual, data blocks with %**
- **S1 HERO — left content / right mosaic grid + sunset circle**
- **S2 ABOUT — NAVY bg, !!sun large center-right, content left**
- **S3 STATS — mosaic 3-column percentage blocks**

No need to read all — skim 2-3 representative slides.
