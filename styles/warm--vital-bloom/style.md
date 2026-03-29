# Vital Bloom: Wellness Brand

## Style Overview

morph-template v35 — Vital Bloom: Wellness Brand。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Cream | #F5F0E8 | 主/辅色块、背景或强调 |
| Cobalt | #1A3FD8 | 主/辅色块、背景或强调 |
| Lime | #7DC840 | 主/辅色块、背景或强调 |
| Lilac | #9B8FD8 | 主/辅色块、背景或强调 |
| Sand | #E8D8B8 | 主/辅色块、背景或强调 |
| Dark | #0A1830 | 主/辅色块、背景或强调 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Segoe UI Black 18-28pt | 分节标题与过渡句 |
| Body | Segoe UI Black 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **Technique**: starburst (fan of rotated thin rects)
- **Technique**: large organic blob ellipses
- **Technique**: Bold morph: !!bloom (large ellipse) shifts position, size, color each slide
- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `v35_build.py`.

**Recommended slides to read for understanding core design techniques**:

- **S1 HERO — large COBALT blob left, starburst top-right, cream BG**
- **hero title on blob**
- **S2 ABOUT — bloom shrinks to top-right, starburst center, stat bubbles**
- **S3 PROGRAMS — bloom bottom-left (LIME), starburst top-right, 3 cards**

No need to read all — skim 2-3 representative slides.
