# Horizon Ventures: VC Fund Deck

## Style Overview

morph-template v39 — Horizon Ventures: VC Fund Deck。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Skyb | #C8E8FF | 主/辅色块、背景或强调 |
| Skyb2 | #A0D4F8 | 主/辅色块、背景或强调 |
| Purple | #8B5CF6 | 主/辅色块、背景或强调 |
| Purpl2 | #C4B5FD | 主/辅色块、背景或强调 |
| Orange | #F97316 | 主/辅色块、背景或强调 |
| Orange2 | #FED7AA | 主/辅色块、背景或强调 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Segoe UI Black 18-28pt | 分节标题与过渡句 |
| Body | Segoe UI Black 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **Technique**: glassmorphism card (semi-transparent roundRect)
- **Technique**: gradient 3D spheres
- **Technique**: Bold morph: !!glass (frosted card) shifts position each slide
- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `v39_build.py`.

**Recommended slides to read for understanding core design techniques**:

- **S1 HERO — large sphere cluster right, glass card center-left, fund name**
- **S2 INVESTMENT THESIS — glass card left, spheres repositioned, thesis points**
- **t(2, "We invest at the intersection\nof infrastructure, AI, and\nhuman-centred systems.",**
- **S3 PORTFOLIO — glass card top-wide, sphere cluster, company list**

No need to read all — skim 2-3 representative slides.
