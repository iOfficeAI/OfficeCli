# Obsidian Amber / Finance & Investment

## Style Overview

morph-template v21 — Obsidian Amber / Finance & Investment。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Black | #070708 | 主/辅色块、背景或强调 |
| Char | #111214 | 主/辅色块、背景或强调 |
| Amber | #FF8C00 | 主/辅色块、背景或强调 |
| Gold | #FFD040 | 主/辅色块、背景或强调 |
| Dgold | #2A1800 | 主/辅色块、背景或强调 |
| White | #FFFFFF | 主/辅色块、背景或强调 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI Black 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Segoe UI 18-28pt | 分节标题与过渡句 |
| Body | Segoe UI 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **Technique**: near-black BG
- **Technique**: amber corner glow
- **Technique**: huge ghost percentage numbers
- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `v21_build.py`.

**Recommended slides to read for understanding core design techniques**:

- **═══ S1 HERO — amber glow corner, ghost "01", textFill fade title ═══**
- **t(1, "MERIDIAN CAPITAL  ·  INVESTOR OVERVIEW", "2cm", "2.5cm", "22cm", "0.9cm", "Segoe UI", 10, DIM),**
- **═══ S2 BIG NUMBER — single massive stat, centered, pure minimal ═══**
- **═══ S3 SPLIT — narrow warm amber panel left, charcoal right ═══**

No need to read all — skim 2-3 representative slides.
