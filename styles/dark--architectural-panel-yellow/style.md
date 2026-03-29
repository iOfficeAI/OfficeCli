# Architectural Panel Yellow

## Style Overview

!/bin/bash。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Dark | #1C2B3A | 主/辅色块、背景或强调 |
| Panel | #B8D4E0 | 主/辅色块、背景或强调 |
| Img Bg | #7FABBF | 主/辅色块、背景或强调 |
| Yellow | #F4C430 | 主/辅色块、背景或强调 |
| Gray | #666666 | 主/辅色块、背景或强调 |
| Lgray | #9A9A9A | 主/辅色块、背景或强调 |

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

Complete build script available in `build_arch_ppt.sh`.

**Recommended slides to read for understanding core design techniques**:

- **SLIDE 1 — TITLE SLIDE  (panel RIGHT)**
- **echo "  Slide 1: Title..."**
- **--prop 'name=!!hero-img' --prop text="[ Architecture Image ]" \**
- **--prop 'name=!!cta-btn' --prop text="Get Started  →" \**

No need to read all — skim 2-3 representative slides.
