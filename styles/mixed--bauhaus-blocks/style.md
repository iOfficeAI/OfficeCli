# Bauhaus Color Block

## Style Overview

morph-template v29 — Bauhaus Color Block。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Forest | #1D5C38 | 主/辅色块、背景或强调 |
| Amber | #F4C040 | 主/辅色块、背景或强调 |
| Tangerine | #E06828 | 主/辅色块、背景或强调 |
| Teal | #1B6060 | 主/辅色块、背景或强调 |
| Dark | #1E1818 | 主/辅色块、背景或强调 |
| Cream | #F0EBE0 | 主/辅色块、背景或强调 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Segoe UI Black 18-28pt | 分节标题与过渡句 |
| Body | Segoe UI Black 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **Technique**: colored rect mosaic
- **Technique**: 3-stacked circles
- **Technique**: vertical bar cluster
- **Technique**: Bold morph: !!panel shifts: right-block→top-stripe→left-col→top-band→accent-bar→full-slide→full-FOREST
- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `v29_build.py`.

**Recommended slides to read for understanding core design techniques**:

- **7 slides — re-use: branding / creative studio / portfolio**
- **S1 HERO — mosaic: left content / right color grid**
- **S2 STATS — 2×2 color cards, !!panel → thin DARK top stripe**
- **S3 FEATURES — FOREST left column, feature rows right, !!panel → left col**

No need to read all — skim 2-3 representative slides.
