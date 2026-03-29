#!/bin/bash
# Strategy Map PPT — 深蓝+青色 + Morph
# 风格：Corporate Deep Blue，适合战略规划/OKR/年度目标
set -euo pipefail

OUT="${1:-$HOME/Desktop/officecli_pptx_suite/strategy_map.pptx}"
rm -f "$OUT"
officecli create "$OUT"

W=12192000; H=6858000
M=432000; FW=11328000

# ── 颜色：深海蓝 + 青色 ──
BG=020B18       # 深海蓝背景
BG2=071525      # 次背景
CARD=0A1E32     # 卡片色
PRI=00C2D4      # 青色（主强调）
ACC=00E5FF      # 亮青
AC2=0077A8      # 深青
TIT=E8F4F8      # 标题
BOD=7BAEC4      # 正文
MUT=3A6070      # 辅助
LIN=0F2840      # 边框

ICO_CHECK='M 15,50 L 38,72 L 82,22'
ICO_ARROW='M 20,50 L 70,50 M 52,30 L 75,50 L 52,70'
ICO_DIAMOND='M 50,8 L 92,50 L 50,92 L 8,50 Z'
ICO_BAR='M 10,80 L 10,50 L 28,50 L 28,80 M 40,80 L 40,25 L 58,25 L 58,80 M 70,80 L 70,55 L 88,55 L 88,80'
ICO_STAR='M 50,10 L 61,37 L 90,37 L 67,57 L 75,84 L 50,67 L 25,84 L 33,57 L 10,37 L 39,37 Z'

add_bullet() {
  local s=$1 p=$2 ico=$3 ic=$4 txt=$5 sz=$6 tc=$7
  local x=$8 y=$9 w=${10} h=${11:-648000}
  local isz=504000
  local tx=$((x + isz + 180000))
  local tw=$((w - isz - 180000))
  officecli add "$OUT" "$s" --type shape \
    --prop name="${p}_ico" --prop geometry="$ico" \
    --prop fill=none --prop line=$ic --prop lineWidth=6 \
    --prop x=$x --prop y=$((y + 54000)) --prop width=$isz --prop height=$isz
  officecli add "$OUT" "$s" --type shape \
    --prop name="${p}_txt" --prop text="$txt" \
    --prop font=Aptos --prop size=$sz --prop color=$tc \
    --prop fill=none --prop line=none \
    --prop x=$tx --prop y=$y --prop width=$tw --prop height=$h
}

echo "=== Strategy Map ==="

# ══ SLIDE 1 — hero ══
officecli add "$OUT" '/' --type slide --prop background=$BG
S='/slide[1]'

# Scene: 左下大圆 + 右上中圆 + 网格光感 + 顶条
officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.1 \
  --prop x=0 --prop y=2520000 --prop width=7200000 --prop height=7200000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.08 \
  --prop x=8640000 --prop y=0 --prop width=5400000 --prop height=5400000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.03 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.15 \
  --prop x=10800000 --prop y=4680000 --prop width=1440000 --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.6 \
  --prop x=5040000 --prop y=2160000 --prop width=360000 --prop height=360000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=72000

# Headline
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=1080000 --prop width=3600000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='2025 Strategy Map' \
  --prop font='Aptos Display' --prop size=66 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1260000 --prop width=8640000 --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub --prop text='Aligning teams, priorities, and outcomes for breakthrough growth' \
  --prop font=Aptos --prop size=20 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=3600000 --prop width=7200000 --prop height=720000

# Bottom: 4 strategic pillars preview
SPILLS=('Growth' 'Efficiency' 'Innovation' 'Culture')
for i in 0 1 2 3; do
  SX=$((M + i * 2808000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_sp${i}" --prop preset=rect \
    --prop fill=$CARD --prop line=$PRI --prop lineWidth=2 \
    --prop x=$SX --prop y=4680000 --prop width=2520000 --prop height=1800000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_sn${i}" --prop text="0$((i+1))" \
    --prop font='Aptos Display' --prop size=28 --prop bold=false --prop color=$PRI \
    --prop fill=none --prop line=none \
    --prop x=$((SX+180000)) --prop y=4860000 --prop width=720000 --prop height=720000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_sl${i}" --prop text="${SPILLS[$i]}" \
    --prop font=Aptos --prop size=14 --prop bold=true --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=$((SX+180000)) --prop y=5652000 --prop width=2160000 --prop height=432000
done

echo "  ✓ S1: hero"

# ══ SLIDE 2 — statement（Vision）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[2]' --prop transition=morph
S='/slide[2]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=3600000 --prop y=0 --prop width=10800000 --prop height=10800000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=4320000 --prop width=3600000 --prop height=3600000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$AC2 --prop line=none --prop opacity=0.06 \
  --prop x=0 --prop y=0 --prop width=2160000 --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$ACC --prop line=none --prop opacity=0.2 \
  --prop x=1440000 --prop y=5400000 --prop width=1080000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.5 \
  --prop x=9360000 --prop y=5400000 --prop width=504000 --prop height=504000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=72000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=1440000 --prop width=1440000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Our North Star' \
  --prop font='Aptos Display' --prop size=60 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1620000 --prop width=10800000 --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub --prop text='To be the operating system every high-performance team runs on by 2027.' \
  --prop font=Aptos --prop size=22 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=3960000 --prop width=9000000 --prop height=1080000

echo "  ✓ S2: statement (vision)"

# ══ SLIDE 3 — pillars（4大战略支柱）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[3]' --prop transition=morph
S='/slide[3]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.07 \
  --prop x=9360000 --prop y=2160000 --prop width=5400000 --prop height=5400000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=1440000 --prop width=$W --prop height=2
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=4320000 --prop width=2520000 --prop height=2520000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=5760000 --prop y=5760000 --prop width=1080000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.4 \
  --prop x=10440000 --prop y=720000 --prop width=360000 --prop height=360000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=72000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='4 Strategic Pillars' \
  --prop font='Aptos Display' --prop size=40 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=7200000 --prop height=1080000

PTITLES=('01 Growth' '02 Efficiency' '03 Innovation' '04 Culture')
PBODIES=('Expand TAM 3× through new segments and geographic expansion' 'Cut unit costs 25% via automation and workflow consolidation' 'Ship 4 AI-native features per quarter, 2 patents filed' 'Top-decile eNPS, 90-day onboarding success for all new hires')
PGOALS=('$65M ARR' '-25% COGS' '4 features/Q' 'eNPS 70+')
for i in 0 1 2 3; do
  CY=$((1800000 + i * 1188000))
  # Left accent bar
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_bar${i}" --prop preset=rect --prop fill=$PRI --prop line=none \
    --prop x=$M --prop y=$CY --prop width=72000 --prop height=900000
  # Title
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_tit${i}" --prop text="${PTITLES[$i]}" \
    --prop font='Aptos Display' --prop size=22 --prop bold=false --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=648000 --prop y=$CY --prop width=5040000 --prop height=540000
  # Body
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_bod${i}" --prop text="${PBODIES[$i]}" \
    --prop font=Aptos --prop size=14 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=648000 --prop y=$((CY+504000)) --prop width=6480000 --prop height=432000
  # Goal badge
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_goal${i}" --prop preset=roundRect \
    --prop fill=$AC2 --prop line=$PRI --prop lineWidth=2 \
    --prop x=8640000 --prop y=$((CY+108000)) --prop width=2880000 --prop height=576000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_gv${i}" --prop text="${PGOALS[$i]}" \
    --prop font='Aptos Display' --prop size=20 --prop bold=false --prop color=$ACC \
    --prop fill=none --prop line=none \
    --prop x=8640000 --prop y=$((CY+180000)) --prop width=2880000 --prop height=432000
done

echo "  ✓ S3: pillars (strategic)"

# ══ SLIDE 4 — evidence（OKRs）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[4]' --prop transition=morph
S='/slide[4]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.08 \
  --prop x=8100000 --prop y=1800000 --prop width=6480000 --prop height=6480000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=0 --prop width=2520000 --prop height=2520000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=$((H-1080000)) --prop width=$W --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$ACC --prop line=none --prop opacity=0.15 \
  --prop x=4320000 --prop y=5760000 --prop width=1080000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.5 \
  --prop x=2160000 --prop y=3960000 --prop width=252000 --prop height=252000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=72000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=1800000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Q1 OKRs' \
  --prop font='Aptos Display' --prop size=44 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=5400000 --prop height=1080000

# OKR progress bars
OKRS=('Reach $48M ARR' 'Launch 3 enterprise features' 'Hire 12 senior engineers' 'Reduce churn < 2%' 'Expand to APAC')
OKRPCT=(78 65 91 88 42)
for i in 0 1 2 3 4; do
  OY=$((1800000 + i * 936000))
  # Label
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_okrl${i}" --prop text="${OKRS[$i]}" \
    --prop font=Aptos --prop size=15 --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=$M --prop y=$OY --prop width=5400000 --prop height=432000
  # BG bar
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_okrbg${i}" --prop preset=rect --prop fill=$CARD --prop line=none \
    --prop x=$M --prop y=$((OY+468000)) --prop width=6480000 --prop height=252000
  # Progress bar
  PW=$((6480000 * OKRPCT[$i] / 100))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_okrfg${i}" --prop preset=rect --prop fill=$PRI --prop line=none \
    --prop x=$M --prop y=$((OY+468000)) --prop width=$PW --prop height=252000
  # Percentage
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_okrp${i}" --prop text="${OKRPCT[$i]}%" \
    --prop font='Aptos Display' --prop size=16 --prop bold=false --prop color=$ACC \
    --prop fill=none --prop line=none \
    --prop x=7200000 --prop y=$OY --prop width=1440000 --prop height=432000
done

echo "  ✓ S4: evidence (OKRs)"

# ══ SLIDE 5 — comparison（执行路线图）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[5]' --prop transition=morph
S='/slide[5]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.07 \
  --prop x=7560000 --prop y=2520000 --prop width=5400000 --prop height=5400000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=0 --prop width=2520000 --prop height=2520000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=2160000 --prop width=$W --prop height=2
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=10440000 --prop y=360000 --prop width=1080000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.4 \
  --prop x=5760000 --prop y=5760000 --prop width=360000 --prop height=360000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=72000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=2520000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Execution Roadmap' \
  --prop font='Aptos Display' --prop size=44 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=7200000 --prop height=1080000

# Timeline: 4 quarters
QS=('Q1 2025' 'Q2 2025' 'Q3 2025' 'Q4 2025')
QCOLS=($PRI $AC2 $AC2 $AC2)
for i in 0 1 2 3; do
  QX=$((M + i * 2844000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_qhdr${i}" --prop preset=rect \
    --prop fill="${QCOLS[$i]}" --prop line=none \
    --prop x=$QX --prop y=1800000 --prop width=2520000 --prop height=324000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_ql${i}" --prop text="${QS[$i]}" \
    --prop font=Aptos --prop size=13 --prop bold=true --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=$((QX+144000)) --prop y=1872000 --prop width=2160000 --prop height=252000
done

# Roadmap items
RM_Q=(0 0 1 1 2 2 3 3)
RM_T=('Enterprise tier launch' 'APAC market entry' 'AI features GA' 'SOC2 Type II renewal' 'Mobile app v2' 'Partner program launch' '$65M ARR target' 'Series C preparation')
for i in 0 1 2 3 4 5 6 7; do
  QI=${RM_Q[$i]}
  RX=$((M + QI * 2844000 + 144000))
  RI=$((i % 2))
  RY=$((2304000 + RI * 1620000 + (i/2 % 2) * 0))
  RY=$((2304000 + (i/8 * 0) + ((i - QI*2)) * 972000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_rm${i}" --prop preset=roundRect \
    --prop fill=$CARD --prop line=$PRI --prop lineWidth=1 \
    --prop x=$((M + QI * 2844000 + 72000)) --prop y=$((2304000 + (i - QI*2)*1116000)) \
    --prop width=2340000 --prop height=900000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_rmt${i}" --prop text="${RM_T[$i]}" \
    --prop font=Aptos --prop size=13 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=$((M + QI * 2844000 + 216000)) \
    --prop y=$((2412000 + (i - QI*2)*1116000)) \
    --prop width=2052000 --prop height=648000
done

echo "  ✓ S5: comparison (roadmap)"

# ══ SLIDE 6 — evidence（资源分配）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[6]' --prop transition=morph
S='/slide[6]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.08 \
  --prop x=7920000 --prop y=0 --prop width=6480000 --prop height=6480000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=4680000 --prop width=2520000 --prop height=2520000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$AC2 --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=0 --prop width=360000 --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=5760000 --prop y=5760000 --prop width=1080000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.5 \
  --prop x=10800000 --prop y=3960000 --prop width=360000 --prop height=360000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=72000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Resource Allocation' \
  --prop font='Aptos Display' --prop size=44 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=7200000 --prop height=1080000

# Left: headcount by team
HC_TEAMS=('Engineering' 'Sales' 'Marketing' 'Product' 'Operations' 'G&A')
HC_VALS=(42 28 18 12 15 8)
for i in 0 1 2 3 4 5; do
  HY=$((1800000 + i * 828000))
  HW=$((HC_VALS[$i] * 120000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_hcl${i}" --prop text="${HC_TEAMS[$i]}" \
    --prop font=Aptos --prop size=13 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=$M --prop y=$HY --prop width=2160000 --prop height=432000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_hcbg${i}" --prop preset=rect --prop fill=$CARD --prop line=none \
    --prop x=2880000 --prop y=$((HY+108000)) --prop width=5040000 --prop height=252000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_hcfg${i}" --prop preset=rect --prop fill=$PRI --prop line=none \
    --prop x=2880000 --prop y=$((HY+108000)) --prop width=$HW --prop height=252000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_hcv${i}" --prop text="${HC_VALS[$i]}" \
    --prop font='Aptos Display' --prop size=15 --prop bold=false --prop color=$ACC \
    --prop fill=none --prop line=none \
    --prop x=8280000 --prop y=$HY --prop width=720000 --prop height=432000
done

# Right: budget donut chart
officecli add "$OUT" "$S" --type chart \
  --prop chartType=pie \
  --prop categories='Engineering,Sales,Marketing,Product,Ops,G&A' \
  --prop series1='35,23,15,10,12,5' \
  --prop colors="00C2D4,0077A8,00E5FF,3A6070,071525,0A1E32" \
  --prop x=9000000 --prop y=1440000 --prop width=2880000 --prop height=2880000 \
  --prop legend=true --prop dataLabels=true

echo "  ✓ S6: evidence (resources)"

# ══ SLIDE 7 — pillars（风险与缓释）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[7]' --prop transition=morph
S='/slide[7]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.07 \
  --prop x=8640000 --prop y=3240000 --prop width=5040000 --prop height=5040000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=rect --prop fill=$MUT --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=1440000 --prop width=$W --prop height=2
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=5040000 --prop width=2160000 --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=11160000 --prop y=720000 --prop width=1080000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.4 \
  --prop x=5400000 --prop y=2160000 --prop width=216000 --prop height=216000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=72000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3600000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Risks & Mitigation' \
  --prop font='Aptos Display' --prop size=40 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=7200000 --prop height=1080000

RISKS=('Market slowdown' 'Key talent attrition' 'Competitive pressure' 'Regulatory changes')
MITS=('Diversify segments; 3-month cash reserve' '401k match + equity refresh every 18mo' 'Patent key IP; deepen integrations' 'Legal review every quarter; SOC2 maintained')
LEVELS=('Medium' 'High' 'Medium' 'Low')
LCOLS=("E8A020" "E83520" "E8A020" "20D45A")
for i in 0 1 2 3; do
  RY=$((1800000 + i * 1188000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_rlvl${i}" --prop preset=roundRect \
    --prop fill="${LCOLS[$i]}" --prop line=none \
    --prop x=$M --prop y=$((RY+108000)) --prop width=1440000 --prop height=432000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_rlvlt${i}" --prop text="${LEVELS[$i]}" \
    --prop font=Aptos --prop size=12 --prop bold=true --prop color=000000 \
    --prop fill=none --prop line=none \
    --prop x=$M --prop y=$((RY+180000)) --prop width=1440000 --prop height=288000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_rr${i}" --prop text="${RISKS[$i]}" \
    --prop font='Aptos Display' --prop size=18 --prop bold=false --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=2016000 --prop y=$RY --prop width=3960000 --prop height=504000
  add_bullet "$S" rm${i} "$ICO_ARROW" "$PRI" "${MITS[$i]}" 13 "$BOD" \
    2016000 $((RY+576000)) 9720000 504000
done

echo "  ✓ S7: pillars (risks)"

# ══ SLIDE 8 — cta（行动号召）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[8]' --prop transition=morph
S='/slide[8]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.15 \
  --prop x=3600000 --prop y=0 --prop width=10800000 --prop height=10800000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.08 \
  --prop x=0 --prop y=4680000 --prop width=3600000 --prop height=3600000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.06 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$ACC --prop line=none --prop opacity=0.2 \
  --prop x=10440000 --prop y=5040000 --prop width=1440000 --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.5 \
  --prop x=1440000 --prop y=720000 --prop width=504000 --prop height=504000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=72000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=1440000 --prop width=2520000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Execute. Measure. Win.' \
  --prop font='Aptos Display' --prop size=62 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1620000 --prop width=10800000 --prop height=2520000

add_bullet "$S" a1 "$ICO_CHECK" "$PRI" 'All team leads confirm Q1 OKR ownership by March 25' 17 "$BOD" \
  $M 4320000 $FW 540000
add_bullet "$S" a2 "$ICO_CHECK" "$PRI" 'Weekly strategy sync every Monday 9am — CEO + VPs' 17 "$BOD" \
  $M 4968000 $FW 540000
add_bullet "$S" a3 "$ICO_CHECK" "$PRI" 'Mid-quarter review scheduled April 15' 17 "$BOD" \
  $M 5616000 $FW 540000

echo "  ✓ S8: cta"

# ══ 验证 ══
echo ""
officecli validate "$OUT"
officecli view outline "$OUT"
echo ""
echo "✅  $OUT"
