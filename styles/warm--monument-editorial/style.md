# Monument Editorial

## Style Overview

morph-template v28 — Monument Editorial。

- **Scene**: 适合商业提案、品牌叙事、产品发布或数据汇报
- **Mood**: 视觉对比清晰，强调节奏与重点信息
- **Tone**: 结构化版式 + 明确强调色

## Color Palette

| Name | Hex | Usage |
| ---- | ---- | ---- |
| Paper | #F0EBE0 | 主/辅色块、背景或强调 |
| Clay | #1A1410 | 主/辅色块、背景或强调 |
| Terra | #C85030 | 主/辅色块、背景或强调 |
| Rust | #8B2E10 | 主/辅色块、背景或强调 |
| Lgray | #B0A898 | 主/辅色块、背景或强调 |
| Mgray | #6A5E52 | 主/辅色块、背景或强调 |

## Typography

| Element | Font | Description |
| ---- | ---- | ---- |
| Main Title | Segoe UI Black 48-72pt | 封面/章节页主标题，作为视觉锚点 |
| Subtitle | Segoe UI 18-28pt | 分节标题与过渡句 |
| Body | Segoe UI 12-18pt | 正文、注释、标签与说明 |

## Design Techniques

- **Technique**: Bold morph: !!block (terracotta filled rect) SHAPE-SHIFTS between slides —
- **形状层级**：区分背景层 / 信息层 / 强调层，避免装饰元素压住正文。
- **遮挡 layout**：主视觉做 Morph actor，正文与数据卡保持稳定锚点。
- **配色选择**：控制为 1 主色 + 1-2 辅色 + 1 强调色，强调色用于 KPI/CTA。
- **结构节奏**：在冲击页与信息页之间交替，保证叙事推进。

## Reference Script

Complete build script available in `v28_build.py`.

**Recommended slides to read for understanding core design techniques**:

- **thin left strip → top band → right half panel → thin bottom strip → center square → full-slide CTA**
- **S1 HERO — !!block: thin left vertical strip**
- **S2 PHILOSOPHY — !!block: top horizontal band**
- **S3 PORTFOLIO — !!block: right half panel (vertical split)**

No need to read all — skim 2-3 representative slides.
