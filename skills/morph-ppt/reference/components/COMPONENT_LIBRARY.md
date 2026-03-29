# COMPONENT_LIBRARY.md

用途：给 Design Expert / 其他 Agent 提供可组合的 PPT 视觉积木（Element + Technique）。

适配对象：`morph-ppt` skill Phase 3，配合 `officecli` CLI 使用。

命令格式约定：所有片段均为真实 officecli CLI 语法，可直接复制进 bash 脚本。
- `$DECK` = 目标文件路径
- `N` = 幻灯片编号（从 1 开始）
- `name` 含 `!!` 前缀的为 Morph 跨页 Actor，需用单引号包裹

---

## Agent 使用协议（先读）

1. 每页先选 `L1` 主结构（1 个）。
2. 再补 `L0` 背景质感（1-2 个）。
3. 再补 `L2` 点缀（1-2 个）。
4. 数据页按需加入 `L3`。
5. `L4` 最多 1 个，不可与可读性冲突。
6. Morph 每页只用 1 个主技法。

硬约束：
- 单页最多 1 个 L4。
- 单页最多 1 个 Morph 主技法。
- 若正文对比不足，先关 L4，再减 L2。
- 连续两页视觉失败，降级为 `L1 + L3 + 1 个 L2`。

---

## Layer Matrix

- L0 背景质感：`grain` `halftone` `corner_glow` `gradient_sweep`
- L1 主结构：`twin_blocks` `glass_card` `organic_blob` `aurora_sphere` `split_panel` `mosaic_layout`
- L2 装饰点缀：`asterisk` `sparkle` `starburst` `ray_fan` `orbital_rings` `arcs` `arrow` `ring_chart` `bar_cluster` `stacked_circles`
- L3 数据内容：`editorial_stat` `stat_3col` `bar_chart`
- L4 排版特效：`ghost_number` `textFill_fade` `chromatic_aberration`

---

## L0 — Background Texture

### grain — 低透明噪点
- 视觉：稀疏小圆点制造胶片颗粒感。
- 适合：深色/复古/编辑风。
- 不适合：数据页（会干扰刻度）。
- 来源：`v15-pink-editorial` `商业风/v19-sage-grain`
- 实现（11 个点，坐标来自 v15，按需铺满 8~12 个）：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=FFFFFF --prop opacity=0.04 --prop x=5cm --prop y=3cm --prop width=0.9cm --prop height=0.9cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=FFFFFF --prop opacity=0.04 --prop x=12.5cm --prop y=6.2cm --prop width=0.5cm --prop height=0.5cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=FFFFFF --prop opacity=0.04 --prop x=25.5cm --prop y=14cm --prop width=0.8cm --prop height=0.8cm
```
- 规则：点径 0.4~0.9cm；opacity 0.03~0.06；不命名（不参与 Morph）。
- 搭配：`split_panel` `textFill_fade`。

---

### halftone — 半调点阵
- 视觉：角落点阵，提供印刷感与秩序。
- 适合：活力、运动、wellness。
- 不适合：高密度文字页。
- 来源：`好看的图形/v35`（halftone 函数，gap=1.2cm，dot=0.22cm）
- 实现（右下角区域示例，按需平移到其他角落）：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=1A3FD8 --prop opacity=0.10 --prop x=24cm --prop y=13.5cm --prop width=0.22cm --prop height=0.22cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=1A3FD8 --prop opacity=0.10 --prop x=25.2cm --prop y=13.5cm --prop width=0.22cm --prop height=0.22cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=1A3FD8 --prop opacity=0.10 --prop x=24cm --prop y=14.7cm --prop width=0.22cm --prop height=0.22cm
```
- 规则：gap 1.1~1.5cm；opacity <= 0.14；覆盖 4~8 列 × 3~5 行。
- 搭配：`starburst` `organic_blob`。

---

### corner_glow — 角落光晕
- 视觉：大渐变椭圆半透明发光，置于画面角落。
- 适合：暗色科技、CTA 页。
- 不适合：浅底细字页。
- 来源：`v31`（`!!sun` 渐变圆）`柔滑边缘/v42-softedge`
- 实现（来自 v31 S1，SKY→ORANGE 渐变，12cm 圆）：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!sun' --prop preset=ellipse --prop fill=A8C8DC --prop gradient=A8C8DC-F0980A-90 --prop x=13.5cm --prop y=3cm --prop width=12cm --prop height=12cm
```
- 规则：覆盖面积 25%~55%；深色背景下加 `--prop softedge=15` 更柔和。
- 搭配：`glass_card` `ghost_number`。

---

### gradient_sweep — 横向渐变扫带
- 视觉：大面积低透明矩形渐变带，横穿画面。
- 适合：大字标题页、编辑风。
- 不适合：多图并排页。
- 来源：`v15-pink-editorial`（`!!num-sweep`，S1）
- 实现（来自 v15 S1）：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!num-sweep' --prop preset=rect --prop fill=C85080 --prop gradient=C85080-160B33-90 --prop opacity=0.35 --prop x=0cm --prop y=3cm --prop width=33.87cm --prop height=12cm
```
- 规则：高度 8~12cm；opacity 0.2~0.4；渐变角度 90 或 270。
- 搭配：`editorial_stat` `textFill_fade`。

---

## L1 — Main Structure

### twin_blocks — 双卡结构
- 视觉：左右等宽卡片，对比信息块。
- 适合：对比、前后、方案 A/B。
- 不适合：超长段落。
- 来源：`商业风/v19-sage-grain`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=F2F0E8 --prop x=2cm --prop y=4cm --prop width=14.5cm --prop height=13.5cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=F2F0E8 --prop x=17.5cm --prop y=4cm --prop width=14.5cm --prop height=13.5cm
```
- 规则：两卡间距 0.8~1.2cm；左右宽度对称。
- 搭配：`sparkle` `editorial_stat`。

---

### glass_card — 玻璃卡
- 视觉：半透明圆角卡，常放在球体前景。
- 适合：投资、科技、金融。
- 不适合：低对比背景。
- 来源：`v39`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=roundRect --prop fill=FFFFFF --prop opacity=0.18 --prop x=3cm --prop y=3cm --prop width=14cm --prop height=10cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=roundRect --prop fill=FFFFFF --prop opacity=0.08 --prop x=3.4cm --prop y=3.4cm --prop width=13.2cm --prop height=9.2cm
```
- 规则：外层 opacity 0.14~0.22；内层比外层各边缩进约 0.4cm。
- 搭配：`corner_glow` `bar_chart`。

---

### organic_blob — 有机大团块
- 视觉：大椭圆团块 + 内层次级椭圆，形成有机感。
- 适合：品牌、wellness、教育。
- 不适合：严肃审计报告。
- 来源：`好看的图形/v35`（`!!bloom`，S1）
- 实现（来自 v35 S1）：
```bash
# 主 blob，命名为 Morph Actor
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!bloom' --prop preset=ellipse --prop fill=1A3FD8 --prop opacity=0.92 --prop x=0cm --prop y=2cm --prop width=17cm --prop height=13cm
# 内层叠加（约外层 0.53 尺寸，偏移 1cm）
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=7DC840 --prop opacity=0.55 --prop x=1cm --prop y=8.5cm --prop width=9cm --prop height=6cm
```
- 规则：内层尺寸约外层 0.45~0.65；跨页通过 `!!bloom` 变色/位移形成 Morph。
- 搭配：`starburst` `halftone`。

---

### aurora_sphere — 极光球体
- 视觉：3 层渐变椭圆 + 柔化边缘，外软内实。
- 适合：深色科技、作品集。
- 不适合：浅底页。
- 来源：`柔滑边缘/v42-softedge`（aurora 函数，r=10，cx=26，cy=6）
- 实现（来自 v42 S1）：
```bash
# 外层 — 背景最大最软（softedge = r×2.5 = 25pt）
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!aurora' --prop preset=ellipse --prop fill=FF6622 --prop gradient=FF6622-8822EE-135 --prop opacity=0.90 --prop softedge=25 --prop x=16cm --prop y=0cm --prop width=18cm --prop height=19cm
# 中层（r×0.70，softedge = 10pt）
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=8822EE --prop gradient=8822EE-8822EE-195 --prop opacity=0.55 --prop softedge=10 --prop x=19cm --prop y=0cm --prop width=14cm --prop height=14cm
# 内核（r×0.38，softedge = 4pt）
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=FFB830 --prop opacity=0.30 --prop softedge=4 --prop x=22.2cm --prop y=2.2cm --prop width=7.6cm --prop height=7.6cm
```
- 规则：三层尺寸比 1.00 : 0.70 : 0.38；softedge 比 2.5 : 1.0 : 0.4（以 r pt 为单位）。
- 搭配：`sparkle` `textFill_fade`。

---

### split_panel — 分栏拼接
- 视觉：左右双色分栏，内容区与视觉区分工。
- 适合：叙事切换、内容-视觉分工。
- 不适合：需要居中聚焦的页面。
- 来源：`v31`（S6）`好看的图形/v35`
- 实现（来自 v31 S6，浅左暗右）：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=F5F2EC --prop x=0cm --prop y=0cm --prop width=17cm --prop height=19.05cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=282E3A --prop x=17cm --prop y=0cm --prop width=16.87cm --prop height=19.05cm
```
- 规则：主文案区占比 35%~55%；分割线与整体网格对齐。
- 搭配：`mosaic_layout` `bar_chart`。

---

### mosaic_layout — 马赛克矩阵
- 视觉：多矩形拼块形成结构秩序。
- 适合：B2B/企业/工程感。
- 不适合：柔和情绪风。
- 来源：`v31`（S1 右侧网格）
- 实现（来自 v31 S1，左白右网格）：
```bash
# 白色左侧内容面板
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=FFFFFF --prop x=0cm --prop y=0cm --prop width=13.5cm --prop height=19.05cm
# 深色顶栏
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=282E3A --prop x=13.5cm --prop y=0cm --prop width=20.37cm --prop height=4.5cm
# 天蓝大块
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=A8C8DC --prop x=13.5cm --prop y=4.5cm --prop width=10.2cm --prop height=9.5cm
# 橙色强调块
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=F0980A --prop x=23.7cm --prop y=4.5cm --prop width=10.17cm --prop height=4.5cm
# 暖色底块
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=E8D0A0 --prop x=23.7cm --prop y=9cm --prop width=10.17cm --prop height=5cm
# 深色底栏
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A2030 --prop x=13.5cm --prop y=14cm --prop width=20.37cm --prop height=5.05cm
```
- 规则：网格边对齐；错位仅在 1 个方向；色块数 4~6 个。
- 搭配：`traveling_orb`（`!!sun`）`stat_3col`。

---

## L2 — Decoration

### asterisk — 星号符点
- 来源：`商业风/v19-sage-grain`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop text="✱" --prop font="Segoe UI Black" --prop size=22 --prop color=B8FF43 --prop fill=none --prop x=2cm --prop y=2cm --prop width=1cm --prop height=1cm
```
- 规则：每页 1~3 个；不压正文区。

---

### sparkle — 十字闪光
- 视觉：水平+垂直细矩形 + ±45° 斜线组成十字。
- 来源：`柔滑边缘/v42-softedge`（sparkle 函数，cx=4，cy=3.5）
- 实现：
```bash
# 横线（named，参与 Morph）
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!spark-h' --prop preset=rect --prop fill=FFFFFF --prop x=2.5cm --prop y=3.43cm --prop width=3cm --prop height=0.14cm
# 竖线（named，参与 Morph）
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!spark-v' --prop preset=rect --prop fill=FFFFFF --prop x=3.93cm --prop y=2cm --prop width=0.14cm --prop height=3cm
# 斜线 +45°（辅助，低 opacity）
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=FFFFFF --prop opacity=0.55 --prop x=3.1cm --prop y=3.45cm --prop width=1.8cm --prop height=0.07cm --prop rotation=45
# 斜线 -45°
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=FFFFFF --prop opacity=0.55 --prop x=3.1cm --prop y=3.45cm --prop width=1.8cm --prop height=0.07cm --prop rotation=-45
```
- 规则：主十字 3cm 长、0.14cm 厚；斜线 1.8cm 长、0.07cm 厚；不压正文。

---

### starburst — 放射扇叶
- 视觉：多条细矩形以同一中心点放射排列。
- 来源：`好看的图形/v35`（starburst 函数，8 rays，cx=28，cy=5，length=8，thick=0.28）
- 实现（8 条，角度间隔 22.5°，所有矩形 x = cx - length/2，y = cy - thick/2）：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A3FD8 --prop opacity=0.65 --prop x=24cm --prop y=4.86cm --prop width=8cm --prop height=0.28cm --prop rotation=0
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A3FD8 --prop opacity=0.65 --prop x=24cm --prop y=4.86cm --prop width=8cm --prop height=0.28cm --prop rotation=22
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A3FD8 --prop opacity=0.65 --prop x=24cm --prop y=4.86cm --prop width=8cm --prop height=0.28cm --prop rotation=45
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A3FD8 --prop opacity=0.65 --prop x=24cm --prop y=4.86cm --prop width=8cm --prop height=0.28cm --prop rotation=67
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A3FD8 --prop opacity=0.65 --prop x=24cm --prop y=4.86cm --prop width=8cm --prop height=0.28cm --prop rotation=90
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A3FD8 --prop opacity=0.65 --prop x=24cm --prop y=4.86cm --prop width=8cm --prop height=0.28cm --prop rotation=112
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A3FD8 --prop opacity=0.65 --prop x=24cm --prop y=4.86cm --prop width=8cm --prop height=0.28cm --prop rotation=135
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=1A3FD8 --prop opacity=0.65 --prop x=24cm --prop y=4.86cm --prop width=8cm --prop height=0.28cm --prop rotation=157
```
- 规则：rays 6~12；角度间隔 = 180/rays；所有矩形锚点以中心对齐。

---

### ray_fan — 定向光线扇
- 视觉：从一个点向某方向放射多条细矩形，模拟舞台光/阳光束，透明度从内到外渐弱。
- 与 `starburst` 的区别：starburst 是 360° 均匀放射；ray_fan 有方向性，有角度范围。
- 适合：科技/渐变背景/封面视觉。不适合：内容密集页。
- 来源：`好看的图形/v30`（ray_fan 函数）
- 实现（来自 v30 S1，左下角光源，base_angle=300，spread=50，10 条）：
```bash
# 10 条光线，角度 300°~350°，透明度从 0.18 递减
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=B4FF20 --prop opacity=0.18 --prop x=0cm --prop y=19cm --prop width=28cm --prop height=0.25cm --prop rotation=300
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=B4FF20 --prop opacity=0.13 --prop x=0cm --prop y=19cm --prop width=28cm --prop height=0.25cm --prop rotation=305
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=B4FF20 --prop opacity=0.08 --prop x=0cm --prop y=19cm --prop width=28cm --prop height=0.25cm --prop rotation=310
```
- 规则：count 8~14；每条透明度递减约 0.05；厚度 0.18~0.30cm；光源常置于角落。
- 搭配：`halftone` `aurora_sphere`。

---

### orbital_rings — 轨道环装饰
- 视觉：两个互相垂直的扁椭圆叠加，模拟行星轨道感。
- 适合：科技/AI/宇宙风。不适合：商务严肃风。
- 来源：`好看的图形/v30`（S1/S5 交叉椭圆）
- 实现（来自 v30 S5，中心 cx=17，cy=9.5）：
```bash
# 横向扁椭圆
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=B4FF20 --prop opacity=0.20 --prop x=10cm --prop y=6cm --prop width=14cm --prop height=7cm
# 纵向扁椭圆（宽窄互换，模拟垂直轨道）
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=D870FF --prop opacity=0.18 --prop x=13cm --prop y=3cm --prop width=7cm --prop height=14cm
```
- 规则：两椭圆长轴相同；短轴约长轴 0.4~0.6；颜色建议用对比色。
- 搭配：`sphere` `ray_fan`。

---

### arcs — 弧线轨道
- 来源：`柔滑边缘/v42-softedge`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=arc --prop fill=FFFFFF --prop opacity=0.35 --prop x=24cm --prop y=2cm --prop width=8cm --prop height=8cm
```
- 规则：仅作辅助装饰；opacity 0.20~0.45。

---

### arrow — 导向箭头
- 来源：`v26 - 副本`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rightArrow --prop fill=F0980A --prop x=2cm --prop y=16cm --prop width=3.5cm --prop height=1.2cm
```
- 规则：每页最多 2 个；必须指向信息块。

---

### ring_chart — 环形图
- 视觉：外圆 + 同色背景内圆镂空模拟环形。
- 来源：`柔滑边缘/v44-2_build.py`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=FF9A5A --prop x=22cm --prop y=6cm --prop width=6cm --prop height=6cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=FFFFFF --prop x=23.3cm --prop y=7.3cm --prop width=3.4cm --prop height=3.4cm
```
- 规则：内孔直径约外圆 0.45~0.65；内圆颜色需与背景一致。

---

### bar_cluster — 微型柱簇
- 视觉：3~4 根不同高度的渐变柱，作为数据感装饰。
- 来源：`v39`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=7AAAC0 --prop gradient=7AAAC0-1A2030-180 --prop x=18cm --prop y=12cm --prop width=1.6cm --prop height=4.2cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=F0980A --prop gradient=F0980A-1A2030-180 --prop x=20cm --prop y=10.5cm --prop width=1.6cm --prop height=5.7cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=A8C8DC --prop gradient=A8C8DC-1A2030-180 --prop x=22cm --prop y=8.6cm --prop width=1.6cm --prop height=7.6cm
```
- 规则：柱宽一致（1.6cm）；间距 0.4~0.8cm；所有柱共用同一 baseline（y + height = 16.2cm）。

---

### stacked_circles — 叠层圆组
- 视觉：2~3 层椭圆叠加，透明度递减形成深度感。
- 来源：`v15-pink-editorial`（`!!accent-dot`）`好看的图形/v35`
- 实现（来自 v15 S1）：
```bash
# 外层（低透明，命名参与 Morph）
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!accent-dot' --prop preset=ellipse --prop fill=FF8DB8 --prop gradient=FF8DB8-C85080-135 --prop opacity=0.28 --prop x=28cm --prop y=13.5cm --prop width=5.5cm --prop height=5.5cm
# 内层（高实体）
officecli add "$DECK" '/slide[N]' --type shape --prop preset=ellipse --prop fill=C85080 --prop opacity=0.85 --prop x=29.5cm --prop y=15cm --prop width=3cm --prop height=3cm
```
- 规则：2~3 层；外层 opacity 0.15~0.35；内层 opacity 0.70~0.95。

---

## L3 — Data Content

### editorial_stat — 巨型统计字
- 视觉：超大数字铺满页面 + 一句说明文字。
- 来源：`v15-pink-editorial`（73%、99.2%，S1/S2）
- 实现（来自 v15 S2，99.2%）：
```bash
# 巨型数字
officecli add "$DECK" '/slide[N]' --type shape --prop text="99.2%" --prop font="Segoe UI Black" --prop size=160 --prop color=FFFFFF --prop bold=true --prop fill=none --prop x=1cm --prop y=2cm --prop width=32cm --prop height=12cm
# 说明文字
officecli add "$DECK" '/slide[N]' --type shape --prop text="customer retention rate across enterprise clients." --prop font="Segoe UI Black" --prop size=22 --prop color=F5E8F0 --prop bold=true --prop fill=none --prop x=2cm --prop y=15cm --prop width=22cm --prop height=3cm
```
- 规则：主数字字号 120~220；`fill=none`；不与 L4 叠用（对比度下降）。

---

### stat_3col — 三列指标块
- 视觉：3 个并列彩色矩形卡，每卡含大数字+标签+说明。
- 来源：`v31`（S3 三列百分比）
- 实现（来自 v31 S3）：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=A8C8DC --prop x=2cm --prop y=4.5cm --prop width=9.5cm --prop height=13cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=F0980A --prop x=12.5cm --prop y=4.5cm --prop width=9.5cm --prop height=13cm
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=282E3A --prop x=23cm --prop y=4.5cm --prop width=9.87cm --prop height=13cm
```
- 规则：三列 y 相同；宽度差 < 0.5cm；内容文字另行 add 在各列矩形之上。

---

### bar_chart — 柱状图
- 视觉：4~6 根渐变柱 + 底部标签。
- 来源：`v39`
- 实现（单柱示例，复制并调整 x/y/height）：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop preset=rect --prop fill=9B8FD8 --prop gradient=9B8FD8-0F1530-180 --prop x=17.5cm --prop y=11.2cm --prop width=2.5cm --prop height=6cm
officecli add "$DECK" '/slide[N]' --type shape --prop text="Q1" --prop font="Segoe UI" --prop size=10 --prop color=8899AA --prop fill=none --prop x=17.5cm --prop y=17.5cm --prop width=2.5cm --prop height=0.8cm
```
- 规则：柱宽一致；间距 0.4~0.8cm；所有柱共用同一 baseline。

---

## L4 — Typography FX

### ghost_number — 幽灵数字
- 视觉：低透明大号数字/年份作背景层。
- 来源：`v6` `v15-pink-editorial`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop text="2026" --prop font="Segoe UI Black" --prop size=180 --prop color=FFFFFF --prop opacity=0.08 --prop fill=none --prop x=15cm --prop y=1cm --prop width=18cm --prop height=10cm
```
- 规则：opacity 0.05~0.12；必须先 add（置于文字层之下）；仅用于标题/封面页。

---

### textFill_fade — 文字渐隐填充
- 视觉：标题文字从亮色向背景色渐变消隐。
- 来源：`商业风/v19-sage-grain`
- 实现：
```bash
officecli add "$DECK" '/slide[N]' --type shape --prop text="LESS IS MORE." --prop font="Segoe UI Black" --prop size=56 --prop color=FFFFFF --prop bold=true --prop fill=none --prop textFill=FFFFFF-0F1A12-0 --prop x=2cm --prop y=4.5cm --prop width=28cm --prop height=4.5cm
```
- 规则：仅用于主标题；正文禁用；`textFill` 终点颜色需接近背景色。

---

### chromatic_aberration — 色像差字层
- 视觉：同一文字三层叠加（粉/青/白），错位模拟 RGB 分色。
- 来源：`v27`（S1，tight 状态，offset ±0.3cm）
- 实现（来自 v27 S1）：
```bash
# 粉层，向左偏移 0.3cm（先 add，置底）
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!pink-layer' --prop text="NOVA SYSTEMS." --prop font="Segoe UI Black" --prop size=68 --prop color=FF0066 --prop bold=true --prop opacity=0.45 --prop fill=none --prop x=1.7cm --prop y=3cm --prop width=30cm --prop height=12cm
# 青层，向右偏移 0.3cm
officecli add "$DECK" '/slide[N]' --type shape --prop 'name=!!cyan-layer' --prop text="NOVA SYSTEMS." --prop font="Segoe UI Black" --prop size=68 --prop color=00F5E4 --prop bold=true --prop opacity=0.40 --prop fill=none --prop x=2.3cm --prop y=3cm --prop width=30cm --prop height=12cm
# 白层，居中实体（最后 add，置顶）
officecli add "$DECK" '/slide[N]' --type shape --prop text="NOVA SYSTEMS." --prop font="Segoe UI Black" --prop size=68 --prop color=FFFFFF --prop bold=true --prop fill=none --prop x=2cm --prop y=3cm --prop width=30cm --prop height=12cm
```
- 规则：三层文本内容必须完全一致；白层最后 add；配合 `spreading_aberration` 跨页扩散。

---

## Technique — Surface / Morph / Rhythm

### softedge_layering
- 意图：用柔化边缘做空间层次，不靠复杂图形。
- 前置：存在 2~3 层椭圆或团块（如 `aurora_sphere`）。
- 步骤：背景层 softedge 最大，中层次之，前景最小或 0。
- 参数：outer:middle:inner = r×2.5 : r×1.0 : r×0.4（pt）。
- 禁用：浅色背景页（softedge 在浅色上无明显效果）。
- 来源：`柔滑边缘/v42-softedge` `柔滑边缘/v44-2`

### gradient_fill
- 意图：同一形体注入方向性光感。
- 前置：主视觉块面积 > 20%。
- 步骤：双色梯度 + 固定角度（90/135/180）；跨页只改角度或颜色一项。
- 禁用：极细条形（高度 < 0.3cm）。
- 来源：`v15-pink-editorial` `v31` `v39`

### opacity_stack
- 意图：用透明度制造前后景纵深。
- 前置：至少 2 层同类形体。
- 步骤：前层 0.75~0.95；中层 0.35~0.65；底层 0.08~0.30。
- 禁用：单层页面。
- 来源：`好看的图形/v35` `v39` `柔滑边缘/v42`

### shape_shifting
- 意图：同名 Actor 跨页变形（位置/大小/颜色）形成叙事。
- 前置：元素设置固定 `!!name`（如 `!!bloom`）。
- 步骤：每页只变 1~2 维（位置 > 尺寸 > 颜色）。
- 禁用：超过 3 个 Actor 同时 shifting。
- 来源：`好看的图形/v35`（`!!bloom` 跨 7 页）`v6`

### twin_journey
- 意图：双 Actor 协同位移，形成对话感。
- 前置：两个同类主元素均有 `!!name`。
- 步骤：A 向主视区移动时，B 反向退场；终页收敛。
- 禁用：单 Actor 页、内容密集页。
- 来源：`柔滑边缘/v44-2`

### traveling_orb
- 意图：单球跨页漫游，作为视觉导航锚点。
- 前置：存在命名渐变圆（如 `!!sun`）。
- 步骤：S1 中等 → S2 放大 → S3 缩小移位 → 末页铺满或退场。
- 禁用：与 `shape_shifting` 同时使用（冲突）。
- 来源：`v31`（`!!sun` 跨 7 页漫游）

### spreading_aberration
- 意图：色像差由轻到强再回归清晰，形成焦点拉伸感。
- 前置：三层同字，命名 `!!pink-layer` / `!!cyan-layer`。
- 步骤：S1 ±0.3cm → S2 ±1.5cm → S3/4 ±4cm → 末页 offset 归零。
- 禁用：超过 4 页（过长则疲劳）。
- 来源：`v27`

### glow_reposition
- 意图：角落光晕跨页换位，制造节奏感。
- 前置：有 `corner_glow` 或 `aurora_sphere`，命名 `!!sun` / `!!aurora`。
- 步骤：右上 → 左下 → 中右 → 全幅铺满。
- 禁用：浅底页。
- 来源：`v31` `柔滑边缘/v42`

### slide_inversion
- 意图：CTA 页做配色反转，强调行动力。
- 前置：前文有稳定主色/辅色。
- 步骤：背景与按钮颜色对调，文案对比升高。
- 禁用：第一页（无前文铺垫则反转无意义）。
- 来源：`v31` `商业风/v19`

### alternating_bg
- 意图：深浅页交替，降低视觉疲劳。
- 前置：至少 4 页。
- 步骤：深 → 浅 → 深 → 浅；标题位置尽量恒定。
- 禁用：少于 4 页。
- 来源：`商业风/v19` `v31`

### diagonal_energy
- 意图：斜向构图提高动势。
- 前置：页面不是纯数据页。
- 步骤：斜线 / 斜矩形 / 斜向星芒三选一（`--prop rotation=` 实现）。
- 禁用：数据密集页（斜向干扰阅读顺序）。
- 来源：`好看的图形/v35` `v26 - 副本`

---

## 推荐组合（可直接给 Agent）

### Combo A — Dark Tech Aurora
- L0: `grain`
- L1: `aurora_sphere`（`!!aurora`）
- L2: `sparkle`（`!!spark-h` / `!!spark-v`）
- L3: `bar_chart`
- L4: `textFill_fade`
- Technique: `softedge_layering` + `traveling_orb`

### Combo B — Corporate Mosaic
- L0: `corner_glow`（`!!sun`）
- L1: `mosaic_layout`
- L2: `arrow`
- L3: `stat_3col`
- L4: `ghost_number`
- Technique: `alternating_bg` + `slide_inversion`

### Combo C — Editorial Energy
- L0: `gradient_sweep`（`!!num-sweep`）
- L1: `split_panel`
- L2: `starburst`
- L3: `editorial_stat`
- L4: `chromatic_aberration`（`!!pink-layer` / `!!cyan-layer`）
- Technique: `spreading_aberration`

### Combo D — Wellness Organic
- L0: `halftone`
- L1: `organic_blob`（`!!bloom`）
- L2: `stacked_circles`（`!!accent-dot`）
- L3: `bar_chart`
- L4: `ghost_number`
- Technique: `shape_shifting` + `opacity_stack`

---

## 维护规则

1. 新增元素必须附来源版本 ID。
2. 每个元素至少提供 2 行可执行的 officecli 命令。
3. 每个技法必须写禁用场景。
4. 命令中颜色值为示例色，使用时根据当前 deck 配色替换。
5. 每次升级记录版本号。

当前版本：`v0.2`（修正命令格式为真实 officecli CLI；参数经 v15/v27/v31/v35/v42 build 脚本核对）。
