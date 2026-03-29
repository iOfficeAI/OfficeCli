"""
components.py — Morph PPT 视觉积木库
=====================================
提炼自 goodcase-2 各版本 build 脚本，供 agent 直接 import 调用。

使用方式：
    import sys, os
    sys.path.insert(0, "/path/to/your/project")   # morph_base.py 所在目录
    sys.path.insert(0, "/path/to/goodcase-2/reference/components")
    from morph_base import *
    from components import *

    # 然后在 batch() 里直接使用：
    batch(DECK, grain(1) + aurora(1, cx=26, cy=6, r=10, c1="FF6622", c2="8822EE", angle=135))

所有函数返回 list[dict]，即 batch() 所需的 ops 列表。
颜色参数均为不带 # 的 hex 字符串，如 "FF6622"。
位置参数为 float（cm），内部自动转为 "Xcm" 字符串。
"""

import math


# ─────────────────────────────────────────────────────────────────
# L0 — Background Texture
# ─────────────────────────────────────────────────────────────────

def grain(slide, color="FFFFFF", opacity="0.04", n=12):
    """
    [L0] grain — 胶片颗粒噪点
    来源：v15-pink-editorial / v19-sage-grain

    散落的小圆点，制造低透明噪点质感。
    适合：深色/复古/编辑风。不适合：数据页。

    Args:
        slide: 幻灯片编号
        color: 点的颜色（默认白色）
        opacity: 透明度，建议 "0.03"~"0.06"
        n: 使用前 n 个预设位置（最多 15）
    """
    # 来自 v19 grain 函数（15 个位置，覆盖全页）
    pos = [
        (4.5,  2.5,  0.8), (11.2, 5.8,  0.5), (19.4, 1.8,  0.7),
        (27.5, 7.2,  0.6), (7.2,  13.5, 0.4), (25.0, 13.2, 0.8),
        (14.5, 10.8, 0.5), (2.8,  9.0,  0.6), (30.2, 3.5,  0.5),
        (1.2,  5.5,  0.7), (22.0, 16.2, 0.4), (9.5,  17.0, 0.5),
        (16.8, 3.2,  0.6), (28.8, 11.5, 0.5), (5.5,  16.0, 0.4),
    ]
    return [s(slide, preset="ellipse", fill=color, opacity=opacity,
              x=f"{x}cm", y=f"{y}cm", width=f"{r}cm", height=f"{r}cm")
            for x, y, r in pos[:min(n, len(pos))]]


def halftone(slide, x0=24.0, y0=13.5, cols=8, rows=5,
             gap=1.2, dot_size=0.22, color="1A3FD8", opacity="0.10"):
    """
    [L0] halftone — 半调点阵
    来源：v35（halftone 函数）

    角落规则点阵，印刷感装饰。
    适合：活力/wellness/运动风。不适合：高密度文字页。

    Args:
        x0, y0: 起始坐标（左上角，cm）
        cols, rows: 列数、行数
        gap: 点间距（cm），建议 1.1~1.5
        dot_size: 点直径（cm），建议 0.18~0.25
        opacity: 建议 <= "0.14"
    """
    items = []
    for r in range(rows):
        for c in range(cols):
            x = round(x0 + c * gap, 2)
            y = round(y0 + r * gap, 2)
            if x < 0 or y < 0 or x + dot_size > 34 or y + dot_size > 20:
                continue
            items.append(s(slide, preset="ellipse", fill=color,
                           x=f"{x}cm", y=f"{y}cm",
                           width=f"{dot_size}cm", height=f"{dot_size}cm",
                           opacity=opacity))
    return items


def gradient_sweep(slide, y=3.0, height=12.0,
                   c1="C85080", c2="160B33", angle=90,
                   opacity="0.35", name="!!num-sweep"):
    """
    [L0] gradient_sweep — 横向渐变扫带
    来源：v15-pink-editorial（!!num-sweep，S1）

    大面积低透明渐变矩形，横穿页面制造编辑感。
    适合：大字标题页。不适合：多图并排页。

    Args:
        y: 矩形顶部位置（cm）
        height: 矩形高度（cm），建议 8~12
        c1, c2: 渐变起止色
        angle: 渐变角度，建议 90 或 270
        opacity: 建议 "0.20"~"0.40"
        name: Morph Actor 名（含 !!）
    """
    return [s(slide, name=name, preset="rect", fill=c1,
              gradient=f"{c1}-{c2}-{angle}", opacity=opacity,
              x="0cm", y=f"{y}cm", width="33.87cm", height=f"{height}cm")]


def corner_glow(slide, cx=19.5, cy=9.0, size=12.0,
                c1="A8C8DC", c2="F0980A", angle=90,
                opacity="1.0", name="!!sun"):
    """
    [L0] corner_glow — 角落光晕（渐变圆）
    来源：v31（!!sun）

    大渐变椭圆作为跨页漫游的视觉锚点。
    适合：暗色科技/企业风。不适合：浅底细字页。

    Args:
        cx, cy: 圆心坐标（cm）
        size: 直径（cm）
        c1, c2: 渐变颜色（上→下）
        angle: 渐变角度
        opacity: 整体透明度
        name: Morph Actor 名（!!sun 跨页漫游）
    """
    x = round(max(0.0, cx - size / 2), 2)
    y = round(max(0.0, cy - size / 2), 2)
    return [s(slide, name=name, preset="ellipse", fill=c1,
              gradient=f"{c1}-{c2}-{angle}",
              x=f"{x}cm", y=f"{y}cm",
              width=f"{size}cm", height=f"{size}cm",
              opacity=opacity)]


# ─────────────────────────────────────────────────────────────────
# L1 — Main Structure
# ─────────────────────────────────────────────────────────────────

def aurora_sphere(slide, cx=26.0, cy=6.0, r=10.0,
                  c1="FF6622", c2="8822EE", angle=135,
                  opacity="0.90", name="!!aurora"):
    """
    [L1] aurora_sphere — 极光球体（3 层渐变椭圆 + 柔化边缘）
    来源：v42-softedge（aurora 函数）

    三层叠加，背景最软，前景最实，制造空间层次。
    适合：深色科技/作品集。不适合：浅底页。

    Softedge 规则（来自 v42）：
        outer = r × 2.5 pt
        middle = r × 1.0 pt
        inner = r × 0.4 pt

    Args:
        cx, cy: 球心坐标（cm）
        r: 外层半径（cm）
        c1, c2: 外层渐变色（c2 也用于中层）
        angle: 渐变角度
        opacity: 外层透明度
        name: 外层 Morph Actor 名
    """
    AMBER = "FFB830"

    x1 = round(max(0.0, cx - r), 2)
    y1 = round(max(0.0, cy - r), 2)
    w1 = round(min(r * 2, 33.87 - x1), 2)
    h1 = round(min(r * 2, 19.05 - y1), 2)
    se_outer = round(r * 2.5, 1)

    r2 = round(r * 0.70, 2)
    x2 = round(max(0.0, cx - r2), 2)
    y2 = round(max(0.0, cy - r2), 2)
    se_middle = round(r * 1.0, 1)

    r3 = round(r * 0.38, 2)
    x3 = round(max(0.0, cx - r3), 2)
    y3 = round(max(0.0, cy - r3), 2)
    se_inner = round(r * 0.4, 1)

    return [
        # 外层 — 背景，最大最软
        s(slide, name=name, preset="ellipse", fill=c1,
          gradient=f"{c1}-{c2}-{angle}",
          x=f"{x1}cm", y=f"{y1}cm", width=f"{w1}cm", height=f"{h1}cm",
          opacity=opacity, softedge=str(se_outer)),
        # 中层 — 中景
        s(slide, preset="ellipse", fill=c2,
          gradient=f"{c2}-{c2}-{(angle + 60) % 360}",
          x=f"{x2}cm", y=f"{y2}cm",
          width=f"{r2*2:.2f}cm", height=f"{r2*2:.2f}cm",
          opacity="0.55", softedge=str(se_middle)),
        # 内核 — 前景，最小最实
        s(slide, preset="ellipse", fill=AMBER,
          x=f"{x3}cm", y=f"{y3}cm",
          width=f"{r3*2:.2f}cm", height=f"{r3*2:.2f}cm",
          opacity="0.30", softedge=str(se_inner)),
    ]


def organic_blob(slide, x=0.0, y=2.0, w=17.0, h=13.0,
                 color="1A3FD8", inner_color="7DC840",
                 opacity="0.92", name="!!bloom"):
    """
    [L1] organic_blob — 有机大团块
    来源：v35（!!bloom）

    大椭圆 + 内层次级椭圆，制造有机层次感。
    适合：品牌/wellness/教育。不适合：严肃报告。

    内层尺寸约外层 0.53，偏移约 1cm。

    Args:
        x, y, w, h: 外层椭圆位置和尺寸（cm）
        color: 外层颜色
        inner_color: 内层叠加色
        opacity: 外层透明度
        name: Morph Actor 名（!!bloom 跨页变形）
    """
    inner_w = round(w * 0.53, 2)
    inner_h = round(h * 0.46, 2)
    inner_x = round(x + 1.0, 2)
    inner_y = round(y + h - inner_h - 0.5, 2)

    return [
        s(slide, name=name, preset="ellipse", fill=color,
          x=f"{x}cm", y=f"{y}cm", width=f"{w}cm", height=f"{h}cm",
          opacity=opacity),
        s(slide, preset="ellipse", fill=inner_color,
          x=f"{inner_x}cm", y=f"{inner_y}cm",
          width=f"{inner_w}cm", height=f"{inner_h}cm",
          opacity="0.55"),
    ]


def glass_card(slide, x=1.5, y=2.0, w=19.0, h=15.0, name="!!glass"):
    """
    [L1] glass_card — 玻璃卡（Glassmorphism）
    来源：v39（glass_card 函数）

    半透明圆角矩形，模拟毛玻璃效果，常覆于球体前景。
    适合：科技/投资/金融风。不适合：低对比背景。

    Args:
        x, y, w, h: 卡片位置和尺寸（cm）
        name: Morph Actor 名（!!glass 跨页位移）
    """
    return [
        # 外层主体
        s(slide, name=name, preset="roundRect", fill="FFFFFF",
          x=f"{x}cm", y=f"{y}cm", width=f"{w}cm", height=f"{h}cm",
          opacity="0.22"),
        # 内层暗边效果
        s(slide, preset="roundRect", fill="FFFFFF",
          x=f"{x}cm", y=f"{y}cm", width=f"{w}cm", height=f"{h}cm",
          opacity="0.07"),
    ]


def sphere(slide, cx=26.0, cy=9.5, r=9.0,
           color="8B5CF6", light_color="C4B5FD",
           opacity="0.70", name=None):
    """
    [L1] sphere — 3D 渐变球体
    来源：v39（sphere 函数）

    单层渐变椭圆模拟立体球感（光源在左上）。
    多个 sphere() 叠加可形成球簇效果。

    Args:
        cx, cy: 球心坐标（cm）
        r: 半径（cm）
        color: 主色（暗部）
        light_color: 高光色（亮部，渐变起点）
        opacity: 透明度
        name: 可选的 Morph Actor 名
    """
    x = round(max(0.0, cx - r), 2)
    y = round(max(0.0, cy - r), 2)
    kw = {"name": name} if name else {}
    return [s(slide, preset="ellipse", fill=color,
              gradient=f"{light_color}-{color}-135",
              x=f"{x}cm", y=f"{y}cm",
              width=f"{r*2:.2f}cm", height=f"{r*2:.2f}cm",
              opacity=opacity, **kw)]


def split_panel(slide, split_x=17.0,
                left_color="F5F2EC", right_color="282E3A"):
    """
    [L1] split_panel — 左右分栏
    来源：v31（S6）

    左右双色矩形，内容区与视觉区分工。
    适合：叙事切换/内容-视觉分工。不适合：居中聚焦页。

    Args:
        split_x: 分割线 x 坐标（cm），建议 13~20
        left_color: 左侧颜色
        right_color: 右侧颜色
    """
    right_w = round(33.87 - split_x, 2)
    return [
        s(slide, preset="rect", fill=left_color,
          x="0cm", y="0cm", width=f"{split_x}cm", height="19.05cm"),
        s(slide, preset="rect", fill=right_color,
          x=f"{split_x}cm", y="0cm",
          width=f"{right_w}cm", height="19.05cm"),
    ]


def mosaic_layout(slide, colors=None):
    """
    [L1] mosaic_layout — 马赛克矩阵（左白右网格）
    来源：v31（S1）

    白色左侧面板 + 右侧 4 色拼块网格，B2B/企业感。
    适合：工程/企业/基础设施。不适合：柔和情绪风。

    Args:
        colors: dict，可覆盖默认色，key: navy/sky/orange/warm/dark
                默认 {"navy":"282E3A","sky":"A8C8DC","orange":"F0980A",
                       "warm":"E8D0A0","dark":"1A2030"}
    """
    c = {"navy": "282E3A", "sky": "A8C8DC", "orange": "F0980A",
         "warm": "E8D0A0", "dark": "1A2030"}
    if colors:
        c.update(colors)
    return [
        s(slide, preset="rect", fill="FFFFFF",
          x="0cm", y="0cm", width="13.5cm", height="19.05cm"),
        s(slide, preset="rect", fill=c["navy"],
          x="13.5cm", y="0cm", width="20.37cm", height="4.5cm"),
        s(slide, preset="rect", fill=c["sky"],
          x="13.5cm", y="4.5cm", width="10.2cm", height="9.5cm"),
        s(slide, preset="rect", fill=c["orange"],
          x="23.7cm", y="4.5cm", width="10.17cm", height="4.5cm"),
        s(slide, preset="rect", fill=c["warm"],
          x="23.7cm", y="9cm", width="10.17cm", height="5cm"),
        s(slide, preset="rect", fill=c["dark"],
          x="13.5cm", y="14cm", width="20.37cm", height="5.05cm"),
    ]


def twin_blocks(slide, x1=2.0, x2=17.5, y=4.0, w=14.5, h=13.5,
                c1="F2F0E8", c2="F2F0E8"):
    """
    [L1] twin_blocks — 双卡结构
    来源：v19-sage-grain（S2）

    左右两张等宽卡片，用于对比/A-B 方案页。
    适合：对比/前后/方案。不适合：超长段落。

    Args:
        x1, x2: 左右卡片的 x 坐标
        y, w, h: 共用的 y/宽/高
        c1, c2: 左右卡片颜色
    """
    return [
        s(slide, preset="rect", fill=c1,
          x=f"{x1}cm", y=f"{y}cm", width=f"{w}cm", height=f"{h}cm"),
        s(slide, preset="rect", fill=c2,
          x=f"{x2}cm", y=f"{y}cm", width=f"{w}cm", height=f"{h}cm"),
    ]


# ─────────────────────────────────────────────────────────────────
# L2 — Decoration
# ─────────────────────────────────────────────────────────────────

def sparkle(slide, cx=30.5, cy=3.5, size=1.5,
            color="FFFFFF", name="!!spark"):
    """
    [L2] sparkle — 十字闪光
    来源：v19（sparkle 函数）/ v42（sparkle 函数，带斜线）

    水平+垂直细矩形组成十字，可作为 Morph Actor 跨页位移。

    Args:
        cx, cy: 中心坐标（cm）
        size: 臂长（cm），建议 1.2~2.5
        color: 颜色
        name: Morph Actor 名（!!spark）
    """
    th = 0.09  # 厚度
    return [
        s(slide, name=name, preset="rect", fill=color,
          x=f"{round(cx - size/2, 2)}cm",
          y=f"{round(cy - th/2, 2)}cm",
          width=f"{size}cm", height=f"{th}cm"),
        s(slide, preset="rect", fill=color,
          x=f"{round(cx - th/2, 2)}cm",
          y=f"{round(max(0.0, cy - size/2), 2)}cm",
          width=f"{th}cm", height=f"{size}cm"),
    ]


def sparkle_full(slide, cx=4.0, cy=3.5, color="FFFFFF", name_h="!!spark-h", name_v="!!spark-v"):
    """
    [L2] sparkle_full — 十字闪光（含 ±45° 斜线，v42 完整版）
    来源：v42（sparkle 函数）

    主十字 3cm 长 0.14cm 厚，斜线 1.8cm 长 0.07cm 厚。

    Args:
        cx, cy: 中心坐标（cm）
        color: 颜色
        name_h, name_v: 横/竖线的 Morph Actor 名
    """
    return [
        s(slide, name=name_h, preset="rect", fill=color,
          x=f"{round(cx-1.5, 2)}cm", y=f"{round(cy-0.07, 2)}cm",
          width="3cm", height="0.14cm"),
        s(slide, name=name_v, preset="rect", fill=color,
          x=f"{round(cx-0.07, 2)}cm", y=f"{round(max(0.0, cy-1.5), 2)}cm",
          width="0.14cm", height="3cm"),
        s(slide, preset="rect", fill=color,
          x=f"{round(cx-0.9, 2)}cm", y=f"{round(cy-0.035, 2)}cm",
          width="1.8cm", height="0.07cm", rotation="45", opacity="0.55"),
        s(slide, preset="rect", fill=color,
          x=f"{round(cx-0.9, 2)}cm", y=f"{round(cy-0.035, 2)}cm",
          width="1.8cm", height="0.07cm", rotation="-45", opacity="0.55"),
    ]


def starburst(slide, cx=28.0, cy=5.0, rays=8,
              length=8.0, thick=0.28,
              color="1A3FD8", opacity="0.65"):
    """
    [L2] starburst — 放射扇叶
    来源：v35（starburst 函数）

    多条细矩形以同一中心点放射排列。
    角度间隔 = 180 / rays。

    Args:
        cx, cy: 中心坐标（cm）
        rays: 射线数量，建议 6~12
        length: 每条矩形长度（cm）
        thick: 每条矩形厚度（cm）
        color: 颜色
        opacity: 透明度
    """
    items = []
    for i in range(rays):
        angle = round(i * (180 / rays))
        x = round(max(0.0, cx - length / 2), 2)
        y = round(max(0.0, cy - thick / 2), 2)
        items.append(s(slide, preset="rect", fill=color,
                       x=f"{x}cm", y=f"{y}cm",
                       width=f"{length}cm", height=f"{thick}cm",
                       rotation=str(angle), opacity=opacity))
    return items


def stacked_circles(slide, cx=30.5, cy=16.0, r_outer=5.5,
                    c_outer="FF8DB8", c_inner="C85080",
                    name="!!accent-dot"):
    """
    [L2] stacked_circles — 叠层圆组
    来源：v15（!!accent-dot）

    外层低透明 + 内层高实体，制造深度感装饰。

    Args:
        cx, cy: 中心坐标（cm）
        r_outer: 外圆半径（cm）
        c_outer: 外圆颜色
        c_inner: 内圆颜色
        name: 外层 Morph Actor 名
    """
    r_inner = round(r_outer * 0.50, 2)
    return [
        s(slide, name=name, preset="ellipse", fill=c_outer,
          gradient=f"{c_outer}-{c_inner}-135", opacity="0.28",
          x=f"{round(cx - r_outer, 2)}cm", y=f"{round(cy - r_outer, 2)}cm",
          width=f"{r_outer*2}cm", height=f"{r_outer*2}cm"),
        s(slide, preset="ellipse", fill=c_inner, opacity="0.85",
          x=f"{round(cx - r_inner, 2)}cm", y=f"{round(cy - r_inner, 2)}cm",
          width=f"{r_inner*2}cm", height=f"{r_inner*2}cm"),
    ]


def ray_fan(slide, cx=4.0, cy=19.05, count=10,
            length=28.0, thick=0.25, color="B4FF20",
            base_angle=300, spread=50, max_opacity=0.18):
    """
    [L2] ray_fan — 定向光线扇
    来源：v30（ray_fan 函数）

    从一个点向某方向放射出多条细矩形，模拟舞台光/阳光束。
    与 starburst 的区别：starburst 是 360° 均匀放射；
    ray_fan 有角度范围（base_angle ± spread/2），且透明度从内到外渐弱。

    适合：科技/渐变背景/封面视觉。不适合：内容密集页。

    Args:
        cx, cy: 光源中心坐标（cm），常放在角落如 (4, 19.05) 左下角
        count: 光线数量，建议 8~14
        length: 每条光线长度（cm），建议 20~30
        thick: 光线厚度（cm），建议 0.18~0.30
        color: 光线颜色
        base_angle: 扇形起始角度（度），如 300 = 左上方
        spread: 扇形角度跨度（度），如 50 表示 ±25°
        max_opacity: 最亮一条的透明度，后续每条递减 0.05
    """
    ops = []
    for i in range(count):
        angle = base_angle + spread * i / max(count - 1, 1)
        x = round(max(0.0, cx - length / 2), 2)
        y = round(max(0.0, cy - thick / 2), 2)
        op = round(max(max_opacity * (1 - i * 0.05), 0.02), 3)
        ops.append(s(slide, preset="rect", fill=color,
                     x=f"{x}cm", y=f"{y}cm",
                     width=f"{length}cm", height=f"{thick}cm",
                     rotation=str(round(angle)), opacity=str(op)))
    return ops


def orbital_rings(slide, cx=26.0, cy=8.0,
                  r_outer=8.0, r_h_ratio=0.5,
                  c1="B4FF20", c2="D870FF",
                  op1="0.20", op2="0.18"):
    """
    [L2] orbital_rings — 轨道环装饰
    来源：v30（S1/S5 两个交叉椭圆）

    两个互相垂直的扁椭圆，模拟行星轨道感。
    一横一竖叠加，制造空间/科技氛围。

    适合：科技/AI/宇宙风。不适合：商务严肃风。

    Args:
        cx, cy: 轨道中心（cm）
        r_outer: 外轨道半径（cm）
        r_h_ratio: 扁率（短轴/长轴），建议 0.4~0.6
        c1: 横向椭圆颜色
        c2: 纵向椭圆颜色
        op1, op2: 两个椭圆的透明度
    """
    w = round(r_outer * 2, 2)
    h = round(r_outer * 2 * r_h_ratio, 2)
    x = round(cx - r_outer, 2)
    y_h = round(cy - h / 2, 2)   # 横向椭圆
    x_v = round(cx - h / 2, 2)   # 纵向椭圆
    y_v = round(cy - r_outer, 2)
    return [
        # 横向扁椭圆
        s(slide, preset="ellipse", fill=c1, opacity=op1,
          x=f"{x}cm", y=f"{y_h}cm",
          width=f"{w}cm", height=f"{h}cm"),
        # 纵向扁椭圆（旋转 90°）
        s(slide, preset="ellipse", fill=c2, opacity=op2,
          x=f"{x_v}cm", y=f"{y_v}cm",
          width=f"{h}cm", height=f"{w}cm"),
    ]


def bar_cluster(slide, x0=18.0, baseline_y=16.2,
                bars_data=None, colors=None):
    """
    [L2] bar_cluster — 微型柱簇（装饰性数据感）
    来源：v39（bars 函数变体）

    3~5 根不同高度的渐变柱，纯视觉装饰。

    Args:
        x0: 第一根柱的 x 坐标（cm）
        baseline_y: 所有柱的底部 y（cm）
        bars_data: list of float，每根柱的高度（cm）
                   默认 [4.2, 5.7, 7.6]
        colors: list of str，颜色列表，默认蓝/橙/天蓝
    """
    if bars_data is None:
        bars_data = [4.2, 5.7, 7.6]
    if colors is None:
        colors = ["7AAAC0", "F0980A", "A8C8DC"]
    dark = "1A2030"
    w, gap = 1.6, 0.8
    items = []
    for i, h in enumerate(bars_data):
        x = round(x0 + i * (w + gap), 2)
        y = round(baseline_y - h, 2)
        c = colors[i % len(colors)]
        items.append(s(slide, preset="rect", fill=c,
                       gradient=f"{c}-{dark}-180",
                       x=f"{x}cm", y=f"{y}cm",
                       width=f"{w}cm", height=f"{h}cm"))
    return items


# ─────────────────────────────────────────────────────────────────
# L3 — Data Content
# ─────────────────────────────────────────────────────────────────

def editorial_stat(slide, number="99.2%", caption="caption text here.",
                   x=1.0, y=2.0, color="FFFFFF", caption_color="F5E8F0",
                   size=160):
    """
    [L3] editorial_stat — 巨型统计字
    来源：v15（73%、99.2% 等）

    超大数字铺满页面 + 一句说明文字。
    不与 L4 叠用（对比度下降）。

    Args:
        number: 显示的数字/文字（如 "99.2%"、"10×"）
        caption: 说明文字
        x, y: 数字左上角坐标（cm）
        color: 数字颜色
        caption_color: 说明文字颜色
        size: 数字字号，建议 120~220
    """
    return [
        t(slide, number,
          f"{x}cm", f"{y}cm", "32cm", "12cm",
          "Segoe UI Black", size, color, bold=True),
        t(slide, caption,
          f"{x+1}cm", f"{round(y+13, 2)}cm", "22cm", "3cm",
          "Segoe UI Black", 22, caption_color, bold=True),
    ]


def stat_3col(slide, y=4.5, h=13.0,
              colors=None, labels=None, numbers=None, descs=None):
    """
    [L3] stat_3col — 三列指标块
    来源：v31（S3）

    3 个并列彩色矩形卡，每卡含大数字 + 标签 + 说明。

    Args:
        y: 三列顶部 y 坐标（cm）
        h: 高度（cm），建议 10~14
        colors: 三列颜色，默认 ["A8C8DC","F0980A","282E3A"]
        labels: 三列标签，如 ["EFFICIENCY","QUALITY","REACH"]
        numbers: 三列数字，如 ["150%","100%","40+"]
        descs: 三列描述（可选）
    """
    if colors is None:
        colors = ["A8C8DC", "F0980A", "282E3A"]
    xs = [2.0, 12.5, 23.0]
    ws = [9.5, 9.5, 9.87]
    text_color = ["FFFFFF", "FFFFFF", "FFFFFF"]
    items = []
    for i in range(3):
        items.append(s(slide, preset="rect", fill=colors[i],
                       x=f"{xs[i]}cm", y=f"{y}cm",
                       width=f"{ws[i]}cm", height=f"{h}cm"))
        if numbers and i < len(numbers):
            items.append(t(slide, numbers[i],
                           f"{xs[i]+0.5}cm", f"{y+0.7}cm",
                           f"{ws[i]-0.5}cm", "4cm",
                           "Segoe UI Black", 52, text_color[i], bold=True))
        if labels and i < len(labels):
            items.append(t(slide, labels[i],
                           f"{xs[i]+0.5}cm", f"{y+4.7}cm",
                           f"{ws[i]-0.5}cm", "1.8cm",
                           "Segoe UI Black", 12, text_color[i], bold=True))
        if descs and i < len(descs):
            items.append(t(slide, descs[i],
                           f"{xs[i]+0.5}cm", f"{y+6.8}cm",
                           f"{ws[i]-0.5}cm", "4cm",
                           "Segoe UI", 11, text_color[i], opacity="0.80"))
    return items


def bar_chart(slide, x0=17.5, baseline_y=17.5,
              values=None, labels=None,
              colors=None, w=2.5, gap=0.5, max_h=7.0):
    """
    [L3] bar_chart — 柱状图
    来源：v39（bars 函数）

    按比例缩放的渐变柱 + 底部标签。

    Args:
        x0: 第一根柱的 x 坐标（cm）
        baseline_y: 柱底 y 坐标（cm）
        values: list of float，数值（按比例缩放）
        labels: list of str，每柱底部标签
        colors: list of str，颜色列表
        w: 柱宽（cm）
        gap: 柱间距（cm）
        max_h: 最高柱的高度（cm）
    """
    if values is None:
        values = [1.5, 0.9, 2.3, 0.7, 1.8]
    if labels is None:
        labels = ["Q1", "Q2", "Q3", "Q4", "Q5"]
    if colors is None:
        colors = ["9B8FD8", "8B5CF6", "C4B5FD", "6D28D9", "A78BFA"]
    dark = "0F1530"
    max_v = max(values)
    items = []
    for i, (val, label) in enumerate(zip(values, labels)):
        h = round(max_h * val / max_v, 2)
        x = round(x0 + i * (w + gap), 2)
        y = round(baseline_y - h, 2)
        c = colors[i % len(colors)]
        items.append(s(slide, preset="roundRect", fill=c,
                       gradient=f"{c}-{dark}-180",
                       x=f"{x}cm", y=f"{y}cm",
                       width=f"{w}cm", height=f"{h}cm", opacity="0.85"))
        items.append(t(slide, label,
                       f"{x}cm", f"{baseline_y+0.2:.2f}cm",
                       f"{w}cm", "1cm", "Segoe UI", 10, "8899AA",
                       align="center"))
    return items


# ─────────────────────────────────────────────────────────────────
# L4 — Typography FX
# ─────────────────────────────────────────────────────────────────

def ghost_number(slide, text="2026", x=15.0, y=1.0,
                 size=180, color="FFFFFF", opacity="0.08"):
    """
    [L4] ghost_number — 幽灵数字
    来源：v6 / v15

    低透明大号数字作为背景层，必须先 add（置于文字层之下）。
    仅用于标题/封面页。opacity 建议 0.05~0.12。

    Args:
        text: 显示内容（年份/数字/字母）
        x, y: 左上角坐标（cm）
        size: 字号，建议 160~220
        color: 颜色
        opacity: 透明度
    """
    return [t(slide, text,
              f"{x}cm", f"{y}cm", "20cm", "12cm",
              "Segoe UI Black", size, color,
              bold=True, opacity=opacity)]


def text_fill_fade(slide, text="LESS IS MORE.", x=2.0, y=4.5,
                   color="FFFFFF", bg_color="0F1A12",
                   size=56, w=28.0):
    """
    [L4] text_fill_fade — 文字渐隐填充
    来源：v19（textFill 参数）

    标题文字从亮色向背景色渐变消隐。
    仅用于主标题，正文禁用。

    Args:
        text: 标题文字
        x, y: 位置（cm）
        color: 文字起始色
        bg_color: 渐变终点色（应与背景色接近）
        size: 字号
        w: 文字框宽度（cm）
    """
    return [t(slide, text,
              f"{x}cm", f"{y}cm", f"{w}cm", "4.5cm",
              "Segoe UI Black", size, color,
              bold=True, textFill=f"{color}-{bg_color}-0")]


def chromatic_aberration(slide, text="NOVA SYSTEMS.", x=2.0, y=3.0,
                         size=68, offset=0.3,
                         pink="FF0066", cyan="00F5E4",
                         name_pink="!!pink-layer", name_cyan="!!cyan-layer"):
    """
    [L4] chromatic_aberration — 色像差字层
    来源：v27（S1，offset ±0.3cm 为 tight 状态）

    三层同字叠加（粉/青/白），错位模拟 RGB 分色。
    配合 spreading_aberration 技法可跨页扩散（±0.3 → ±1.5 → ±4.0 → 0）。

    命名约定：
        粉层 = !!pink-layer（左偏移）
        青层 = !!cyan-layer（右偏移）
        白层 = 无名（置顶）

    Args:
        text: 三层共用文字（必须完全一致）
        x, y: 白层（基准）位置（cm）
        size: 字号
        offset: 粉/青层的水平偏移量（cm），跨页时修改此值
        pink, cyan: 粉/青层颜色
        name_pink, name_cyan: Morph Actor 名
    """
    return [
        # 粉层 — 向左偏移，先 add（置底）
        t(slide, text,
          f"{round(x - offset, 2)}cm", f"{y}cm", "30cm", "12cm",
          "Segoe UI Black", size, pink,
          bold=True, opacity="0.45", name=name_pink),
        # 青层 — 向右偏移
        t(slide, text,
          f"{round(x + offset, 2)}cm", f"{y}cm", "30cm", "12cm",
          "Segoe UI Black", size, cyan,
          bold=True, opacity="0.40", name=name_cyan),
        # 白层 — 居中实体，最后 add（置顶）
        t(slide, text,
          f"{x}cm", f"{y}cm", "30cm", "12cm",
          "Segoe UI Black", size, "FFFFFF", bold=True),
    ]


# ─────────────────────────────────────────────────────────────────
# 推荐组合（Combo）
# ─────────────────────────────────────────────────────────────────

def combo_dark_tech_aurora(slide, title="YOUR TITLE HERE.",
                           subtitle="Supporting message goes here.",
                           accent="B8FF43"):
    """
    Combo A — Dark Tech Aurora
    L0: grain  L1: aurora_sphere  L2: sparkle_full  L4: text_fill_fade
    Technique: softedge_layering + traveling_orb
    背景色建议：050510
    """
    ops = []
    ops += grain(slide, opacity="0.04")
    ops += aurora_sphere(slide, cx=26, cy=6, r=10,
                         c1="FF6622", c2="8822EE", angle=135)
    ops += sparkle_full(slide, cx=4, cy=3.5, color=accent)
    ops += text_fill_fade(slide, title, x=2, y=4.5,
                          color="FFFFFF", bg_color="050510", size=52)
    ops += [t(slide, subtitle, "2cm", "10cm", "20cm", "3cm",
              "Segoe UI", 16, "8899BB")]
    return ops


def combo_corporate_mosaic(slide, title="OUR TECHNOLOGIES\nYOUR SOLUTIONS.",
                           body="Supporting description text."):
    """
    Combo B — Corporate Mosaic
    L0: corner_glow  L1: mosaic_layout  L3: stat_3col ready  L4: ghost_number
    Technique: alternating_bg + slide_inversion
    背景色建议：F5F2EC
    """
    ops = []
    ops += mosaic_layout(slide)
    ops += corner_glow(slide, cx=19.5, cy=9.0, size=12,
                       c1="A8C8DC", c2="F0980A", angle=90)
    ops += ghost_number(slide, text="01", x=14, y=5, size=180,
                        color="282E3A", opacity="0.06")
    ops += [
        t(slide, title, "2cm", "3cm", "10cm", "6cm",
          "Segoe UI Black", 28, "282E3A", bold=True),
        t(slide, body, "2cm", "9.5cm", "10cm", "4cm",
          "Segoe UI", 13, "282E3A", opacity="0.70"),
    ]
    return ops


def combo_editorial_energy(slide, number="73%",
                            caption="of businesses report better outcomes."):
    """
    Combo C — Editorial Energy
    L0: gradient_sweep  L1: split_panel  L2: starburst  L3: editorial_stat  L4: chromatic_aberration
    Technique: spreading_aberration
    背景色建议：160B33-7B2D52-135（渐变）
    """
    ops = []
    ops += gradient_sweep(slide, y=3, height=12,
                          c1="C85080", c2="160B33", angle=90, opacity="0.35")
    ops += starburst(slide, cx=28, cy=5, rays=8,
                     length=8, thick=0.28, color="FF8DB8", opacity="0.50")
    ops += editorial_stat(slide, number=number, caption=caption,
                          color="FFFFFF", caption_color="F5E8F0", size=160)
    ops += grain(slide, opacity="0.04")
    return ops


def combo_wellness_organic(slide, title="Move.\nBreathe.\nBecome.",
                            body="Your daily ritual for mindful movement."):
    """
    Combo D — Wellness Organic
    L0: halftone  L1: organic_blob  L2: starburst  L4: ghost_number
    Technique: shape_shifting + opacity_stack
    背景色建议：F5F0E8
    """
    ops = []
    ops += halftone(slide, x0=24, y0=13.5, cols=8, rows=5,
                    color="1A3FD8", opacity="0.10")
    ops += starburst(slide, cx=28, cy=5, rays=8,
                     length=8, thick=0.28, color="1A3FD8", opacity="0.65")
    ops += organic_blob(slide, x=0, y=2, w=17, h=13,
                        color="1A3FD8", inner_color="7DC840")
    ops += [
        t(slide, title, "1.5cm", "4.2cm", "14cm", "10cm",
          "Segoe UI Black", 50, "E8ECFF", bold=True),
        t(slide, body, "19cm", "4cm", "13cm", "4.5cm",
          "Segoe UI", 16, "0A1830"),
    ]
    return ops
