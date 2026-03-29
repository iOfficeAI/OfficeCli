# Firmwise SaaS / Efficiency

## Style Overview

morph-template v6 — Firmwise SaaS / Efficiency。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Bg | #F0F2F9 | 主/辅色块、背景或强调 |
| C1 | #5B4FFF | 主/辅色块、背景或强调 |
| C2 | #9B8BFF | 主/辅色块、背景或强调 |
| C3 | #E8E5FF | 主/辅色块、背景或强调 |
| Fg | #1A1A2E | 主/辅色块、背景或强调 |
| White | #FFFFFF | 主/辅色块、背景或强调 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Segoe UI Black 18-28pt | 分节标题与过渡句 |
| Body | Segoe UI Black 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `v6_build.py`.

**Recommended slides to read for understanding core design techniques**:

- **S1 — HERO  (light bg, big purple headline, 3 chamfered preview cards)**
- **3 chamfered stat cards (morph actor: expand on S3)**
- **S2 — PROBLEM  (same layout, new text → morph S1→S2 = content swap)**
- **S3 — FEATURES  (cards move up & expand → hero stat layout)**

No need to read all — skim 2-3 representative slides.
