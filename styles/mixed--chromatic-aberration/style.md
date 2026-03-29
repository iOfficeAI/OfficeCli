# Chromatic Aberration

## Style Overview

morph-template v27 — Chromatic Aberration。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Dark | #050814 | 主/辅色块、背景或强调 |
| Dark2 | #0A1030 | 主/辅色块、背景或强调 |
| Cyan | #00F5E4 | 主/辅色块、背景或强调 |
| Pink | #FF0066 | 主/辅色块、背景或强调 |
| White | #FFFFFF | 主/辅色块、背景或强调 |
| Dim | #334466 | 主/辅色块、背景或强调 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI Black 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Segoe UI 18-28pt | 分节标题与过渡句 |
| Body | Segoe UI 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **Technique**: Bold morph: !!cyan-layer and !!pink-layer are ghost copies of the title text,
- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `v27_build.py`.

**Recommended slides to read for understanding core design techniques**:

- **offset horizontally (and vertically on S5). As slides advance, aberration**
- **SPREADS → EXPANDS → COLLAPSES → shifts axis → reconverges on CTA.**
- **S1 HERO — tight aberration (offset ±0.3cm)**
- **S2 PRODUCT — aberration spreads wider (±1.5cm)**

No need to read all — skim 2-3 representative slides.
