#!/bin/bash
# Budget Review PPT v2 — 严格遵循 pptx-design.md v7 规范
# 坐标系：1cm = 360000 EMU，直接用 cm 值计算
# 三层 actor 体系 + SVG 图标 + Morph 动效
set -euo pipefail

OUT="${1:-$HOME/Desktop/officecli_pptx_suite/budget_review_v2.pptx}"
rm -f "$OUT"
officecli create "$OUT"

# ══════════════════════════════════════════════════════
# cm → EMU 换算函数
cm() { echo $(( ${1%.*} * 360000 + (${1#*.} * 360000 / 10) )) 2>/dev/null || echo $(( $1 * 360000 )); }

# 画布：33.87cm × 19.05cm
W=12192000; H=6858000
M=432000       # margin 1.2cm
FW=11328000    # 可用宽度 (W - 2M)

# ══════════════════════════════════════════════════════
# 配色 — Dark Tech + 金/绿
BG=080A0E; BG2=0F1318; CARD=141A20
PRI=C9A84C   # 金色
ACC=5A9E48   # 绿色
AC2=8BC94A   # 亮绿
TIT=F2F0E8   # 标题文字
BOD=A8A090   # 正文
MUT=6A6458   # 辅助
LIN=252820   # 边框

# ══════════════════════════════════════════════════════
# SVG icon paths (归一化 0-100 坐标)
ICO_CHECK='M 15,50 L 38,72 L 82,22'
ICO_ARROW='M 20,50 L 70,50 M 52,30 L 75,50 L 52,70'
ICO_DIAMOND='M 50,8 L 92,50 L 50,92 L 8,50 Z'
ICO_STAR='M 50,10 L 61,37 L 90,37 L 67,57 L 75,84 L 50,67 L 25,84 L 33,57 L 10,37 L 39,37 Z'
ICO_ALERT='M 50,12 L 88,82 L 12,82 Z M 50,42 L 50,62 M 50,70 L 50,74'
ICO_BAR='M 10,80 L 10,50 L 28,50 L 28,80 M 40,80 L 40,25 L 58,25 L 58,80 M 70,80 L 70,55 L 88,55 L 88,80'

# ══════════════════════════════════════════════════════
# Helper: icon + text bullet
# add_bullet SLIDE PFX ICON ICOL TEXT SIZE TCOL X Y W H
add_bullet() {
  local s=$1 p=$2 ico=$3 ic=$4 txt=$5 sz=$6 tc=$7
  local x=$8 y=$9 w=${10} h=${11:-648000}  # 默认行高 1.8cm
  local isz=540000  # 1.5cm icon
  local tx=$((x + isz + 180000))
  local tw=$((w - isz - 180000))
  officecli add "$OUT" "$s" --type shape \
    --prop name="${p}_ico" --prop geometry="$ico" \
    --prop fill=none --prop line=$ic --prop lineWidth=6 \
    --prop x=$x --prop y=$((y + 54000)) --prop width=$isz --prop height=$isz
  officecli add "$OUT" "$s" --type shape \
    --prop name="${p}_txt" --prop text="$txt" \
    --prop font=Calibri --prop size=$sz --prop color=$tc \
    --prop fill=none --prop line=none \
    --prop x=$tx --prop y=$y --prop width=$tw --prop height=$h
}

echo "=== Budget Review v2 ==="

# ══════════════════════════════════════════════════════
# SLIDE 1 — hero
# Scene: 右侧大圆溢出 + 左上角矩形 + 右下中圆 + 菱形 + 小点 + 顶条
officecli add "$OUT" '/' --type slide --prop background=$BG
S='/slide[1]'

# Scene actors (6个，不同大小散落)
officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=7560000 --prop y=0 --prop width=7200000 --prop height=7200000   # 右侧 20cm 圆溢出

officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.15 \
  --prop x=9360000 --prop y=3600000 --prop width=4320000 --prop height=4320000  # 右下 12cm

officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.07 \
  --prop x=0 --prop y=0 --prop width=2160000 --prop height=2160000  # 左上 6cm

officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$AC2 --prop line=none --prop opacity=0.18 \
  --prop x=360000 --prop y=4680000 --prop width=1440000 --prop height=1440000  # 左下菱形 4cm

officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.6 \
  --prop x=6192000 --prop y=2016000 --prop width=360000 --prop height=360000  # 小圆点 1cm

officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=108000  # 顶部金条 0.3cm

# Headline actors
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=792000 --prop width=2520000 --prop height=108000  # 标签条 7cm×0.3cm
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag_txt --prop text='BUDGET REVIEW  ·  FY 2024' \
  --prop font=Calibri --prop size=11 --prop bold=true --prop color=$BG \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=792000 --prop width=2520000 --prop height=108000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Budget Review' \
  --prop font='Aptos Display' --prop size=68 --prop bold=false --prop color=$PRI \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1188000 --prop width=5400000 --prop height=2160000  # 15cm 宽

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub --prop text='Strategic Financial Analysis & Planning' \
  --prop font=Aptos --prop size=22 --prop bold=false --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=3492000 --prop width=5400000 --prop height=648000

# Content actors
officecli add "$OUT" "$S" --type shape \
  --prop name=co_line --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=4320000 --prop width=1800000 --prop height=72000  # 5cm 横线

officecli add "$OUT" "$S" --type shape \
  --prop name=co_org --prop text='Finance & Strategy Division' \
  --prop font=Aptos --prop size=15 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=4500000 --prop width=3600000 --prop height=432000

# 图片占位（右侧半屏）
officecli add "$OUT" "$S" --type shape \
  --prop name=co_imgph --prop preset=roundRect --prop fill=$CARD --prop line=$LIN --prop lineWidth=2 \
  --prop x=6480000 --prop y=360000 --prop width=5520000 --prop height=6120000  # 18cm×17cm
officecli add "$OUT" "$S" --type shape \
  --prop name=co_imgph_txt --prop text='[ 流光背景图片 ]' \
  --prop font=Aptos --prop size=14 --prop color=$LIN \
  --prop fill=none --prop line=none \
  --prop x=7920000 --prop y=3240000 --prop width=2880000 --prop height=432000

echo "  ✓ S1: hero"

# ══════════════════════════════════════════════════════
# SLIDE 2 — statement（大字冲击，观点页）
# Scene actors 大幅换位
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[2]' --prop transition=morph
S='/slide[2]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.10 \
  --prop x=0 --prop y=0 --prop width=5400000 --prop height=5400000  # 左上大圆

officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.10 \
  --prop x=9720000 --prop y=3960000 --prop width=4320000 --prop height=4320000  # 右下

officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$AC2 --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H  # 全屏淡色底

officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.20 \
  --prop x=9000000 --prop y=1440000 --prop width=1800000 --prop height=1800000

officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$AC2 --prop line=none --prop opacity=0.5 \
  --prop x=1080000 --prop y=5400000 --prop width=504000 --prop height=504000  # 左下小点

officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$ACC --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=108000

# Headline — 大号居中
officecli add "$OUT" "$S" --type shape \
  --prop name=co_accent --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=2016000 --prop width=1440000 --prop height=144000  # 金色短横线

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title \
  --prop text='Effective budget reviews drive alignment — not just accountability.' \
  --prop font='Aptos Display' --prop size=40 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=2268000 --prop width=9360000 --prop height=2520000  # 26cm 宽

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub \
  --prop text='When every stakeholder shares the same financial context, decisions follow naturally.' \
  --prop font=Aptos --prop size=20 --prop italic=true --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=4896000 --prop width=9000000 --prop height=720000

echo "  ✓ S2: statement"

# ══════════════════════════════════════════════════════
# SLIDE 3 — pillars（4 列卡片：概览）
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[3]' --prop transition=morph
S='/slide[3]'

# Scene actors — 结构化横条背景
officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=1800000 --prop width=$W --prop height=5400000  # 横条

officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.08 \
  --prop x=9000000 --prop y=3600000 --prop width=5400000 --prop height=5400000  # 右下大圆溢出

officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=0 --prop width=4680000 --prop height=1620000  # 左上矩形

officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$AC2 --prop line=none --prop opacity=0.12 \
  --prop x=360000 --prop y=5760000 --prop width=1080000 --prop height=1080000

officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.4 \
  --prop x=11520000 --prop y=720000 --prop width=360000 --prop height=360000

officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=108000

# Headline
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=1800000 --prop height=108000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag_txt --prop text='OVERVIEW' \
  --prop font=Aptos --prop size=10 --prop bold=true --prop color=$BG \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=360000 --prop width=1800000 --prop height=108000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Previous Budget Period Overview' \
  --prop font='Aptos Display' --prop size=36 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=540000 --prop width=$FW --prop height=1080000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub \
  --prop text='Establishing the baseline and strategic framework that guided allocation decisions.' \
  --prop font=Aptos --prop size=15 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1692000 --prop width=$FW --prop height=504000

# 4 cards — 每张约 8cm × 8.5cm，间距 0.4cm
CW=2736000; CH=3060000; CY=2304000; CG=108000
CTITLES=('Budget Allocated' 'Period Covered' 'Primary Objectives' 'Major Initiatives')
CSUBS=('Total funding approved for the period' 'Timeframe under review' 'Key goals and strategic priorities' 'Significant projects funded')

for i in 0 1 2 3; do
  CX=$((M + i*(CW + CG)))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_card${i}" --prop preset=roundRect \
    --prop fill=$CARD --prop line=$LIN --prop lineWidth=1 \
    --prop x=$CX --prop y=$CY --prop width=$CW --prop height=$CH
  # 顶部色条
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_strip${i}" --prop preset=rect --prop fill=$PRI --prop line=none \
    --prop x=$CX --prop y=$CY --prop width=$CW --prop height=108000
  # 数字徽章
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_num_bg${i}" --prop preset=ellipse \
    --prop fill=$PRI --prop line=none --prop opacity=0.15 \
    --prop x=$((CX + 144000)) --prop y=$((CY + 216000)) --prop width=720000 --prop height=720000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_num${i}" --prop text=$((i+1)) \
    --prop font=Aptos --prop size=20 --prop bold=true --prop color=$PRI \
    --prop fill=none --prop line=none \
    --prop x=$((CX + 144000)) --prop y=$((CY + 216000)) --prop width=720000 --prop height=720000
  # 标题
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_ctitle${i}" --prop text="${CTITLES[$i]}" \
    --prop font='Aptos Display' --prop size=19 --prop bold=false --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=$((CX + 144000)) --prop y=$((CY + 1044000)) --prop width=$((CW - 288000)) --prop height=540000
  # 子文本（带 check 图标）
  add_bullet "$S" "co_cb${i}" "$ICO_CHECK" "$ACC" "${CSUBS[$i]}" 13 "$BOD" \
    $((CX + 144000)) $((CY + 1656000)) $((CW - 288000)) 540000
done

officecli add "$OUT" "$S" --type shape \
  --prop name=co_footer \
  --prop text='This section provides context for the financial period being reviewed, establishing the baseline for our analysis.' \
  --prop font=Aptos --prop size=12 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=5616000 --prop width=$FW --prop height=720000

echo "  ✓ S3: pillars (overview)"

# ══════════════════════════════════════════════════════
# SLIDE 4 — evidence（3 KPI 非对称）
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[4]' --prop transition=morph
S='/slide[4]'

# Scene actors — 大圆做数据背景（opacity 豁免规则）
officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.30 \
  --prop x=360000 --prop y=900000 --prop width=5760000 --prop height=5760000  # 左侧 16cm 大圆（数据背景）

officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.20 \
  --prop x=5400000 --prop y=2160000 --prop width=3960000 --prop height=3960000  # 中右 11cm

officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=1800000

officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$AC2 --prop line=none --prop opacity=0.12 \
  --prop x=11160000 --prop y=5400000 --prop width=1440000 --prop height=1440000

officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.5 \
  --prop x=360000 --prop y=6120000 --prop width=504000 --prop height=504000

officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$ACC --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=108000

# Headline
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$ACC --prop line=none \
  --prop x=$M --prop y=360000 --prop width=2880000 --prop height=108000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag_txt --prop text='BUDGET OVERVIEW' \
  --prop font=Aptos --prop size=10 --prop bold=true --prop color=$BG \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=360000 --prop width=2880000 --prop height=108000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Budget Allocation Summary' \
  --prop font='Aptos Display' --prop size=36 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=540000 --prop width=$FW --prop height=1080000

# 3 KPI（非对称：大 + 中 + 小）
# KPI 1：最大（左）
officecli add "$OUT" "$S" --type shape \
  --prop name=co_k1 --prop text='$4.2M' \
  --prop font='Aptos Display' --prop size=72 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1980000 --prop width=3600000 --prop height=1800000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_k1_lbl --prop text='Total Budget Allocated' \
  --prop font=Aptos --prop size=17 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=3816000 --prop width=3600000 --prop height=504000
add_bullet "$S" co_k1_note "$ICO_BAR" "$PRI" 'Across 5 departments, 12 projects' 13 "$MUT" \
  $M 4392000 3600000 540000

# KPI 2：中
officecli add "$OUT" "$S" --type shape \
  --prop name=co_k2 --prop text='$3.8M' \
  --prop font='Aptos Display' --prop size=54 --prop bold=false --prop color=$AC2 \
  --prop fill=none --prop line=none \
  --prop x=4320000 --prop y=2340000 --prop width=3240000 --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_k2_lbl --prop text='Total Budget Spent' \
  --prop font=Aptos --prop size=17 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=4320000 --prop y=3816000 --prop width=3240000 --prop height=504000
add_bullet "$S" co_k2_note "$ICO_ARROW" "$AC2" '90.5% utilization rate' 13 "$MUT" \
  4320000 4392000 3240000 540000

# KPI 3：小
officecli add "$OUT" "$S" --type shape \
  --prop name=co_k3 --prop text='9.5%' \
  --prop font='Aptos Display' --prop size=44 --prop bold=false --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=7920000 --prop y=2700000 --prop width=2880000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_k3_lbl --prop text='Strategic Reserve' \
  --prop font=Aptos --prop size=17 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=7920000 --prop y=3816000 --prop width=2880000 --prop height=504000
add_bullet "$S" co_k3_note "$ICO_ALERT" "$MUT" 'Maintained as planned' 13 "$MUT" \
  7920000 4392000 2880000 540000

# 底部 insight
officecli add "$OUT" "$S" --type shape \
  --prop name=co_iline --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=5580000 --prop width=$FW --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_insight \
  --prop text='Key Insight: Utilization exceeded 85% target across all departments, with strategic reserve maintained.' \
  --prop font=Aptos --prop size=13 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=5724000 --prop width=$FW --prop height=540000

echo "  ✓ S4: evidence (KPIs)"

# ══════════════════════════════════════════════════════
# SLIDE 5 — comparison（左右分栏：文字 + 饼图）
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[5]' --prop transition=morph
S='/slide[5]'

SP=6192000  # 分割点 17.2cm（参考 demo slide 3 的图片起始位置）
DW=144000   # 分隔线 0.4cm

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=0 --prop width=$SP --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=rect --prop fill=$ACC --prop line=none --prop opacity=0.05 \
  --prop x=$((SP+DW)) --prop y=0 --prop width=$((W-SP-DW)) --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$SP --prop y=720000 --prop width=$DW --prop height=$((H-1440000))
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.08 \
  --prop x=0 --prop y=0 --prop width=3600000 --prop height=3600000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$AC2 --prop line=none --prop opacity=0.12 \
  --prop x=$((SP+DW+2160000)) --prop y=5400000 --prop width=1080000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$SP --prop height=108000

# Headline（左栏）
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=108000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag_txt --prop text='SPENDING ANALYSIS' \
  --prop font=Aptos --prop size=10 --prop bold=true --prop color=$BG \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=108000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Expense Breakdown' \
  --prop font='Aptos Display' --prop size=34 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=540000 --prop width=$((SP-M-360000)) --prop height=1080000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub --prop text='How resources were allocated across major spending categories.' \
  --prop font=Aptos --prop size=14 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1692000 --prop width=$((SP-M-360000)) --prop height=504000

# 4 category rows（左栏）
CAT_TEXTS=(
  'Category 1 — largest portion  (35%)'
  'Category 2 — operational costs  (28%)'
  'Category 3 — support functions  (22%)'
  'Category 4 — strategic investments  (15%)'
)
CAT_COLS=($PRI $AC2 $ACC $MUT)
for i in 0 1 2 3; do
  RY=$((2304000 + i*1044000))
  RW=$((SP - M - 360000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_rbar${i}" --prop preset=rect --prop fill="${CAT_COLS[$i]}" --prop line=none \
    --prop x=$M --prop y=$RY --prop width=144000 --prop height=828000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_rtxt${i}" --prop text="${CAT_TEXTS[$i]}" \
    --prop font=Aptos --prop size=14 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=$((M+288000)) --prop y=$((RY+180000)) --prop width=$((RW-288000)) --prop height=432000
done

# 右栏：饼图
CX=$((SP+DW+360000)); CW2=$((W-SP-DW-720000))
officecli add "$OUT" "$S" --type chart \
  --prop chartType=pie \
  --prop categories='Category 1,Category 2,Category 3,Category 4' \
  --prop series1='35,28,22,15' \
  --prop colors="$PRI,8BC94A,5A9E48,6A6458" \
  --prop x=$CX --prop y=540000 --prop width=$CW2 --prop height=5760000 \
  --prop legend=true --prop dataLabels=true

echo "  ✓ S5: comparison (expense)"

# ══════════════════════════════════════════════════════
# SLIDE 6 — evidence（差异柱图，全宽）
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[6]' --prop transition=morph
S='/slide[6]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.08 \
  --prop x=8640000 --prop y=3240000 --prop width=5400000 --prop height=5400000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.08 \
  --prop x=0 --prop y=0 --prop width=2160000 --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$AC2 --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=108000 --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.15 \
  --prop x=1440000 --prop y=5400000 --prop width=1440000 --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$ACC --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=108000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$ACC --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=108000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag_txt --prop text='VARIANCE ANALYSIS' \
  --prop font=Aptos --prop size=10 --prop bold=true --prop color=$BG \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=108000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Budget vs. Actual Variance' \
  --prop font='Aptos Display' --prop size=36 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=540000 --prop width=$FW --prop height=1080000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub \
  --prop text='Planned versus actual expenditures across departments, highlighting key deviations.' \
  --prop font=Aptos --prop size=15 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1692000 --prop width=$FW --prop height=504000

officecli add "$OUT" "$S" --type chart \
  --prop chartType=bar \
  --prop title='Department Budget vs Actual ($)' \
  --prop categories='Operations,Marketing,R&D,HR,IT' \
  --prop series1='1200000,800000,600000,400000,500000' \
  --prop series2='1050000,920000,580000,390000,460000' \
  --prop colors="$PRI,$ACC" \
  --prop x=$M --prop y=2304000 --prop width=$FW --prop height=3960000 \
  --prop legend=true

officecli add "$OUT" "$S" --type shape \
  --prop name=co_note \
  --prop text='Note: Marketing over-spend reflects accelerated Q3 campaign. R&D and IT within approved parameters.' \
  --prop font=Aptos --prop size=12 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=6372000 --prop width=$FW --prop height=360000

echo "  ✓ S6: evidence (variance)"

# ══════════════════════════════════════════════════════
# SLIDE 7 — pillars（前瞻 3 列，与 S6 全宽结构不同）
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[7]' --prop transition=morph
S='/slide[7]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.07 \
  --prop x=4176000 --prop y=0 --prop width=7200000 --prop height=7200000  # 顶部中间大圆溢出
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.08 \
  --prop x=0 --prop y=3240000 --prop width=3600000 --prop height=3600000  # 左下
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=9360000 --prop y=720000 --prop width=2880000 --prop height=$((H-1440000))
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop preset=rect --prop fill=$ACC --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=108000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$AC2 --prop line=none --prop opacity=0.15 \
  --prop x=11520000 --prop y=5400000 --prop width=1080000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.5 \
  --prop x=360000 --prop y=1080000 --prop width=504000 --prop height=504000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$ACC --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=108000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag_txt --prop text='FORWARD OUTLOOK' \
  --prop font=Aptos --prop size=10 --prop bold=true --prop color=$BG \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=108000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Strategic Priorities for Next Period' \
  --prop font='Aptos Display' --prop size=34 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=540000 --prop width=7920000 --prop height=1080000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub \
  --prop text="Building on this period's learnings to drive more efficient resource allocation." \
  --prop font=Aptos --prop size=15 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1692000 --prop width=7920000 --prop height=504000

# 3 priority columns，每个约 10.8cm × 8.5cm，间距 0.5cm
PCW=3744000; PCH=3060000; PCY=2304000; PCG=180000
PTITLES=('Operational Excellence' 'Strategic Investment' 'Risk Management')
PTEXTS=(
  'Streamline core processes and reduce overhead by 12% through targeted efficiency programs.'
  'Increase R&D allocation by 15% to accelerate innovation pipeline and competitive positioning.'
  'Maintain 10% strategic reserve and implement quarterly variance review cadence.'
)
PCOLORS=($PRI $AC2 $MUT)
PICONS=("$ICO_ARROW" "$ICO_STAR" "$ICO_ALERT")

for i in 0 1 2; do
  PX=$((M + i*(PCW + PCG)))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_pc${i}" --prop preset=roundRect \
    --prop fill=$CARD --prop line=$LIN --prop lineWidth=1 \
    --prop x=$PX --prop y=$PCY --prop width=$PCW --prop height=$PCH
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_pstrip${i}" --prop preset=rect --prop fill="${PCOLORS[$i]}" --prop line=none \
    --prop x=$PX --prop y=$PCY --prop width=$PCW --prop height=144000
  # icon
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_pico${i}" --prop geometry="${PICONS[$i]}" \
    --prop fill=none --prop line="${PCOLORS[$i]}" --prop lineWidth=7 \
    --prop x=$((PX + 216000)) --prop y=$((PCY + 288000)) --prop width=720000 --prop height=720000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_ptit${i}" --prop text="${PTITLES[$i]}" \
    --prop font='Aptos Display' --prop size=19 --prop bold=false --prop color="${PCOLORS[$i]}" \
    --prop fill=none --prop line=none \
    --prop x=$((PX + 216000)) --prop y=$((PCY + 1080000)) --prop width=$((PCW - 432000)) --prop height=540000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_ptxt${i}" --prop text="${PTEXTS[$i]}" \
    --prop font=Aptos --prop size=13 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=$((PX + 216000)) --prop y=$((PCY + 1692000)) --prop width=$((PCW - 432000)) --prop height=1224000
done

echo "  ✓ S7: pillars (outlook)"

# ══════════════════════════════════════════════════════
# SLIDE 8 — cta（行动号召，scene 回归散落与 hero 呼应）
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[8]' --prop transition=morph
S='/slide[8]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.10 \
  --prop x=0 --prop y=0 --prop width=3600000 --prop height=3600000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=9360000 --prop y=3960000 --prop width=4320000 --prop height=4320000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$AC2 --prop line=none --prop opacity=0.04 \
  --prop x=4320000 --prop y=0 --prop width=3600000 --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=108000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.20 \
  --prop x=720000 --prop y=3060000 --prop width=1440000 --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.4 \
  --prop x=11520000 --prop y=720000 --prop width=360000 --prop height=360000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=1800000 --prop height=108000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag_txt --prop text='NEXT STEPS' \
  --prop font=Aptos --prop size=10 --prop bold=true --prop color=$BG \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=360000 --prop width=1800000 --prop height=108000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Action Items & Decision Points' \
  --prop font='Aptos Display' --prop size=38 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=540000 --prop width=$FW --prop height=1080000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub \
  --prop text='Clear ownership and timelines for each priority item emerging from this budget review.' \
  --prop font=Aptos --prop size=16 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1692000 --prop width=$FW --prop height=504000

# 3 action rows
ATITS=('Approve Q1 budget reallocation proposals' 'Implement variance tracking dashboard' 'Conduct departmental budget alignment workshops')
AOWNERS=('CFO / Finance Committee' 'FP&A Team / IT' 'Department Heads / HR')
ADATES=('Q1 2025' 'Q2 2025' 'Q2 2025')
ACOLORS=($PRI $AC2 $MUT)

for i in 0 1 2; do
  AY=$((2304000 + i*1440000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_abg${i}" --prop preset=rect \
    --prop fill=$CARD --prop line=$LIN --prop lineWidth=1 \
    --prop x=$M --prop y=$AY --prop width=$FW --prop height=1260000
  # 左色条
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_abar${i}" --prop preset=rect --prop fill="${ACOLORS[$i]}" --prop line=none \
    --prop x=$M --prop y=$AY --prop width=144000 --prop height=1260000
  # check icon
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_aico${i}" --prop geometry="$ICO_CHECK" \
    --prop fill=none --prop line="${ACOLORS[$i]}" --prop lineWidth=7 \
    --prop x=$((M+360000)) --prop y=$((AY+360000)) --prop width=540000 --prop height=540000
  # title
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_atit${i}" --prop text="${ATITS[$i]}" \
    --prop font='Aptos Display' --prop size=20 --prop bold=false --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=$((M+1080000)) --prop y=$((AY+180000)) --prop width=8640000 --prop height=504000
  # owner（diamond icon）
  add_bullet "$S" "co_aown${i}" "$ICO_DIAMOND" "${ACOLORS[$i]}" "${AOWNERS[$i]}" 13 "$MUT" \
    $((M+1080000)) $((AY+756000)) 5040000 432000
  # date badge
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_adtbg${i}" --prop preset=rect \
    --prop fill="${ACOLORS[$i]}" --prop line=none --prop opacity=0.15 \
    --prop x=$((M+FW-2520000)) --prop y=$((AY+396000)) --prop width=2520000 --prop height=468000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_adt${i}" --prop text="${ADATES[$i]}" \
    --prop font=Aptos --prop size=15 --prop bold=true --prop color="${ACOLORS[$i]}" \
    --prop fill=none --prop line=none \
    --prop x=$((M+FW-2520000)) --prop y=$((AY+396000)) --prop width=2520000 --prop height=468000
done

# closing
officecli add "$OUT" "$S" --type shape \
  --prop name=co_cline --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=6480000 --prop width=$FW --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_closing \
  --prop text='Together, data-driven decisions create accountability and drive sustainable financial performance.' \
  --prop font='Aptos Display' --prop size=16 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=6588000 --prop width=$FW --prop height=216000

echo "  ✓ S8: cta (action items)"

# ══════════════════════════════════════════════════════
officecli validate "$OUT" >/dev/null
echo ""
echo "✅  $OUT"
officecli view "$OUT" outline
