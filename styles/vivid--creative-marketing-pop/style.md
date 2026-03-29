# Creative Marketing Pop

## Style Overview

!/bin/bash。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Primary | #1A1A1A | 主要标题与高对比信息 |
| Secondary | #F5F5F5 | 大面积背景 |
| Accent | #FF6A00 | 关键数字与强调元素 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Arial 18-28pt | 分节标题与过渡句 |
| Body | Arial 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `creative-marketing.sh`.

**Recommended slides to read for understanding core design techniques**:

- **── S1  bg=#F5E0C0  no transition ────────────────────────────**
- **echo "S1"**
- **── S2  bg=#1A1A1A  transition=morph ─────────────────────────**
- **echo "S2"**

No need to read all — skim 2-3 representative slides.
