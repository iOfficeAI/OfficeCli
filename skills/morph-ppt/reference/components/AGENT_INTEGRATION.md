# AGENT_INTEGRATION.md

给 `morph-ppt` skill 维护者的接入说明。
本文件描述如何让 Design Expert 使用 `COMPONENT_LIBRARY.md` 进行分层组合生成。

---

## 组件库文件清单

```
goodcase-2/reference/components/
├── COMPONENT_LIBRARY.md   # 人类可读描述 + officecli CLI 命令片段（agent 设计时读这个）
├── components.py          # 可 import 的 Python 函数库（Python build 脚本用这个）
└── AGENT_INTEGRATION.md   # 本文件，接入说明
```

**两种使用方式：**

| 场景 | 用哪个 |
|------|--------|
| morph-ppt skill（生成 bash 脚本） | 读 `COMPONENT_LIBRARY.md`，复制 CLI 片段 |
| Python build 脚本 | `from components import *`，直接调用函数 |

每个 PPT 版本文件夹（v6、v15、v31 等）内都有对应的 `build.py` 脚本，记录了该版本的完整视觉实现，供组件库持续补充参考。

---

## 需要在 skill 里改的地方

只需改 Phase 3 的两处：

### 1. 新增参考文档读取

在 Phase 3 开始时，让 agent 读取组件库：

```
在生成前，先读取 `goodcase-2/reference/components/COMPONENT_LIBRARY.md`，
了解可用的视觉积木（L0~L4 Element + Technique）。
```

### 2. 修改设计指令

将原来的"按 style.md 复用布局"改为"按组件库分层组合"：

**原指令逻辑（旧）：**
> 参考 style.md 的布局方案，复现其视觉风格。

**新指令逻辑：**
> 按以下流程为每页选择组件：
> 1. 从 L1 选 1 个主结构
> 2. 从 L0 选 1~2 个背景质感
> 3. 从 L2 选 1~2 个点缀装饰
> 4. 数据页从 L3 选内容组件
> 5. 最多叠加 1 个 L4 排版特效
> 6. 从 Technique 选 1 个 Morph 主策略
>
> 不同页面可以组合来自不同版本的组件（跨风格组合）。
> 组件库提供了 4 个推荐 Combo，可直接使用或作为起点变体。

---

## 硬约束（直接复制进 skill prompt）

```
设计硬约束：
- 每页最多 1 个 L4 特效
- 每页最多 1 个 Morph 主技法
- 若正文对比度不足，先去掉 L4，再减少 L2
- 连续两页生成失败，降级为 L1 + L3 + 最多 1 个 L2
- Morph Actor 必须使用 !! 前缀命名，单引号包裹
```

---

## 归一化规则（解决组件兼容性问题）

不同组件来源版本坐标不一，混搭前须遵守以下基准：

| 参数 | 规范值 |
|------|--------|
| 画布尺寸 | 33.87cm × 19.05cm |
| 安全边距 | 上下左右各 1.5cm |
| 内容区 | x: 1.5~32.37cm，y: 1.5~17.55cm |
| 正文字号 | 18~28pt（标题 36~60pt） |
| 最小留白 | 相邻元素间距 ≥ 0.5cm |
| L0 点/噪点型 opacity | ≤ 0.15（grain、halftone） |
| L0 渐变光晕型 opacity | 不限（corner_glow、gradient_sweep 依赖 softedge 控制柔化，不靠 opacity） |
| L2 层 opacity | ≤ 0.35（避免压正文） |

**跨版本组件混搭时**：
- L1 主结构坐标优先，L0/L2 围绕 L1 调整
- 若组件与 L1 发生重叠，缩小 L2 尺寸（不移动 L1）
- 文字区域（正文所在 x/y 范围）内不放 L2/L4

---

## Morph Actor 命名规范（解决跨页 actor 冲突）

统一命名格式：`!!<layer>-<role>-<slot>`

| layer | 含义 |
|-------|------|
| `bg` | L0 背景质感 |
| `struct` | L1 主结构 |
| `deco` | L2 装饰 |
| `fx` | L4 特效 |

**slot** 为同类第几个（1/2/3...），同一 PPT 内不重名。

示例：
```bash
# L0 背景光晕
--prop 'name=!!bg-glow-1'

# L1 主球体
--prop 'name=!!struct-sphere-1'

# L2 装饰星光
--prop 'name=!!deco-spark-1'

# L4 鬼影数字
--prop 'name=!!fx-ghost-1'
```

**每页生成前**，先列出该页计划使用的所有 actor 名称，确认无重名。跨页同一 actor 必须完全相同的名称（Morph 靠名称匹配跨页形变）。

---

## 默认 Combo 模板（解决首轮质量波动）

**Agent 默认使用 Combo A**，失败时按顺序降级：

### Combo A（深色科技极光）—— 默认首选

```
L0: corner_glow(!!bg-glow-1) + grain
L1: aurora_sphere(!!struct-sphere-1) + glass_card(!!struct-card-1)
L2: sparkle(!!deco-spark-1)
L4: ghost_number(!!fx-ghost-1)
Morph: softedge_layering
```

```bash
# L0 角落光晕（渐变光晕不设 opacity，用 softedge 控制柔化）
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!bg-glow-1' --prop preset=ellipse --prop fill=A8C8DC --prop gradient=A8C8DC-F0980A-90 --prop softedge=25 --prop x=13.5cm --prop y=3cm --prop width=12cm --prop height=12cm
# L1 极光球体（3层）
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!struct-sphere-1' --prop preset=ellipse --prop fill=FF6622 --prop gradient=FF6622-8822EE-135 --prop opacity=0.90 --prop softedge=25 --prop x=16cm --prop y=0cm --prop width=18cm --prop height=19cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=FF6622 --prop gradient=FF6622-8822EE-135 --prop opacity=0.55 --prop softedge=10 --prop x=19cm --prop y=2cm --prop width=12cm --prop height=15cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=FF6622 --prop gradient=FF6622-8822EE-135 --prop opacity=0.30 --prop softedge=4 --prop x=21cm --prop y=4cm --prop width=8cm --prop height=11cm
# L1 玻璃卡
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!struct-card-1' --prop preset=roundRect --prop fill=FFFFFF --prop opacity=0.18 --prop x=3cm --prop y=3cm --prop width=14cm --prop height=10cm
# L2 星光
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!deco-spark-1' --prop preset=plus --prop fill=FFFFFF --prop x=30cm --prop y=2cm --prop width=1.5cm --prop height=1.5cm
# L4 鬼影数字（文字区域外）
officecli add "$DECK" '/slide[N]' --type text --prop 'name=!!fx-ghost-1' --prop text=2026 --prop size=180 --prop color=FFFFFF --prop opacity=0.08 --prop x=15cm --prop y=1cm --prop width=18cm --prop height=10cm
```

---

### 失败降级模板（Combo F）

连续两页失败时降级至最简结构：

```
L1: split_panel 或 twin_blocks
L3: editorial_stat 或 stat_3col
L2: sparkle（仅 1 个）
无 L4，无 L0
```

```bash
# L1 左右分屏
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=F5F2EC --prop x=0cm --prop y=0cm --prop width=17cm --prop height=19.05cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=282E3A --prop x=17cm --prop y=0cm --prop width=16.87cm --prop height=19.05cm
# L2 单星光
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!deco-spark-1' --prop preset=plus --prop fill=FFFFFF --prop x=30cm --prop y=2cm --prop width=1.2cm --prop height=1.2cm
```

---

## 规则优先级（冲突时以此为准）

```
硬约束  >  可读性检查  >  Combo 示例参数
```

示例参数仅为参考起点，若与硬约束冲突，以硬约束为准。

---

## 生成后自动检查清单（每页执行）

生成每页后，按顺序检查以下 3 项，任一不通过则按指示修复：

```
CHECK 1 — 单页 L4 数量
  ✗ 当前页 L4 形状数量 > 1 → 保留最后一个，删除其余

CHECK 2 — Morph actor 命名/重名
  ✗ 当前页存在与前页同名但语义不同的 !! actor → 重命名（slot 编号+1）
  ✗ 同一页有两个相同 !! 名称 → 删除重复的

CHECK 3 — 文字区域装饰侵入
  正文文字区域 = style.md 指定区域，或默认 x: 2~16cm, y: 3~16cm
  ✗ 该区域内有 L2 形状（opacity > 0.05）→ 移出区域或删除
  ✗ 该区域内有 L4 形状 → 删除
  ✗ 该区域内有 L0 点/噪点型（grain/halftone）opacity > 0.08 → 降低至 0.05

连续 2 页 CHECK 3 不通过 → 启用 Combo F 降级模板
```

---

## 最小验收标准

一页 PPT 视为"生成成功"，需同时满足：

| 指标 | 标准 |
|------|------|
| L4 数量 | ≤ 1 |
| Morph actor 重名 | 0 |
| 文字区域被 L2/L4 覆盖 | 无 |
| 背景/正文对比色组合合法 | 深底浅字 或 浅底深字 |
| L1 主结构存在 | 是 |

全部通过 = 成功。否则按检查清单修复后重新验证。

---

## style.md 的角色变化

接入组件库后，`style.md` 不再是"布局蓝本"，而是：
- 提供配色方案（用于替换组件命令中的颜色示例值）
- 提供整体气质定位（决定选哪个 Combo 作为起点）

两者互补，不冲突。

---

## 不需要改的地方

- OfficeCLI 调用链
- 4 Phase 主流程
- Planner / Quality Reviewer 角色边界
- morph-helpers.sh 工具函数
