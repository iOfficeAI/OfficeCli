#!/bin/bash
# Investor Deck — 白底极简 + 品红强调 + Morph
# 风格：Light Minimal Editorial，适合融资路演
set -euo pipefail

OUT="${1:-$HOME/Desktop/officecli_pptx_suite/investor_deck.pptx}"
rm -f "$OUT"
officecli create "$OUT"

W=12192000; H=6858000
M=432000; FW=11328000

# ── 颜色：白底极简 + 品红 ──
BG=FAFAFA      # 近白背景
BG2=F0F0F0     # 浅灰
CARD=FFFFFF    # 卡片白
PRI=E8215A     # 品红（主强调）
ACC=FF6B9D     # 浅品红
AC2=1A1A2E     # 深蓝黑（文字）
TIT=0D0D0D     # 标题
BOD=4A4A5A     # 正文
MUT=9A9AAA     # 辅助
LIN=E0E0E8     # 边框

ICO_CHECK='M 15,50 L 38,72 L 82,22'
ICO_ARROW='M 20,50 L 70,50 M 52,30 L 75,50 L 52,70'
ICO_DIAMOND='M 50,8 L 92,50 L 50,92 L 8,50 Z'
ICO_BAR='M 10,80 L 10,50 L 28,50 L 28,80 M 40,80 L 40,25 L 58,25 L 58,80 M 70,80 L 70,55 L 88,55 L 88,80'
ICO_CIRCLE='M 50,10 A 40,40 0 1,1 49.9,10 Z'

add_bullet() {
  local s=$1 p=$2 ico=$3 ic=$4 txt=$5 sz=$6 tc=$7
  local x=$8 y=$9 w=${10} h=${11:-648000}
  local isz=504000
  local tx=$((x + isz + 180000))
  local tw=$((w - isz - 180000))
  officecli add "$OUT" "$s" --type shape \
    --prop name="${p}_ico" --prop geometry="$ico" \
    --prop fill=none --prop line=$ic --prop lineWidth=7 \
    --prop x=$x --prop y=$((y + 54000)) --prop width=$isz --prop height=$isz
  officecli add "$OUT" "$s" --type shape \
    --prop name="${p}_txt" --prop text="$txt" \
    --prop font=Aptos --prop size=$sz --prop color=$tc \
    --prop fill=none --prop line=none \
    --prop x=$tx --prop y=$y --prop width=$tw --prop height=$h
}

echo "=== Investor Deck ==="

# ══ SLIDE 1 — hero ══
officecli add "$OUT" '/' --type slide --prop background=$BG
S='/slide[1]'

# Scene: 右侧大品红圆（低透明） + 左上小圆 + 底部横条
officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.06 \
  --prop x=7920000 --prop y=0 --prop width=6480000 --prop height=6480000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.08 \
  --prop x=9720000 --prop y=3960000 --prop width=3240000 --prop height=3240000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=$((H-72000)) --prop width=$W --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=5400000 --prop y=5400000 --prop width=900000 --prop height=900000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.5 \
  --prop x=4320000 --prop y=1080000 --prop width=252000 --prop height=252000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=54000

# Headline
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=756000 --prop width=216000 --prop height=756000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Investor Presentation' \
  --prop font='Aptos Display' --prop size=64 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=720000 --prop y=720000 --prop width=6480000 --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub --prop text='Series B — Strategic Growth Plan 2025' \
  --prop font=Aptos --prop size=22 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=720000 --prop y=2952000 --prop width=5760000 --prop height=648000

# Content: 3 key numbers bottom-left
officecli add "$OUT" "$S" --type shape \
  --prop name=co_n1 --prop text='$42M' \
  --prop font='Aptos Display' --prop size=40 --prop bold=false --prop color=$PRI \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=3960000 --prop width=2160000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_l1 --prop text='ARR' \
  --prop font=Aptos --prop size=13 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=4968000 --prop width=2160000 --prop height=432000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_n2 --prop text='3.2×' \
  --prop font='Aptos Display' --prop size=40 --prop bold=false --prop color=$PRI \
  --prop fill=none --prop line=none \
  --prop x=2880000 --prop y=3960000 --prop width=2160000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_l2 --prop text='YoY Growth' \
  --prop font=Aptos --prop size=13 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=2880000 --prop y=4968000 --prop width=2160000 --prop height=432000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_n3 --prop text='94%' \
  --prop font='Aptos Display' --prop size=40 --prop bold=false --prop color=$PRI \
  --prop fill=none --prop line=none \
  --prop x=5580000 --prop y=3960000 --prop width=2160000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=co_l3 --prop text='Net Retention' \
  --prop font=Aptos --prop size=13 --prop color=$MUT \
  --prop fill=none --prop line=none \
  --prop x=5580000 --prop y=4968000 --prop width=2160000 --prop height=432000

echo "  ✓ S1: hero"

# ══ SLIDE 2 — statement（问题陈述）══
officecli add "$OUT" '/' --type slide --prop background=$AC2
officecli set "$OUT" '/slide[2]' --prop transition=morph
S='/slide[2]'

# Scene: 深色背景 + 品红大圆 + 几何
officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.15 \
  --prop x=7200000 --prop y=1800000 --prop width=7200000 --prop height=7200000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.08 \
  --prop x=0 --prop y=3600000 --prop width=3600000 --prop height=3600000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.06 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$ACC --prop line=none --prop opacity=0.2 \
  --prop x=10800000 --prop y=5400000 --prop width=1440000 --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.6 \
  --prop x=1800000 --prop y=1080000 --prop width=360000 --prop height=360000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=54000

# Headline: 大引述
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=900000 --prop width=3240000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='The market is broken.' \
  --prop font='Aptos Display' --prop size=58 --prop bold=false --prop color=FAFAFA \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1080000 --prop width=7560000 --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub --prop text='Enterprises waste $2.4T annually on fragmented, disconnected workflows. We fix that.' \
  --prop font=Aptos --prop size=20 --prop color=$ACC \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=3456000 --prop width=7200000 --prop height=1080000

echo "  ✓ S2: statement"

# ══ SLIDE 3 — pillars（产品三大支柱）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[3]' --prop transition=morph
S='/slide[3]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.05 \
  --prop x=8640000 --prop y=2160000 --prop width=5400000 --prop height=5400000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=rect --prop fill=$LIN --prop line=none \
  --prop x=0 --prop y=1440000 --prop width=$W --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=4320000 --prop width=2160000 --prop height=2160000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.1 \
  --prop x=11520000 --prop y=360000 --prop width=720000 --prop height=720000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.5 \
  --prop x=6480000 --prop y=720000 --prop width=216000 --prop height=216000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=54000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=2880000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Why We Win' \
  --prop font='Aptos Display' --prop size=44 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=6480000 --prop height=1080000

# 3 pillars as cards
PTITLES=('Unified Platform' 'AI-Native' 'Enterprise Scale')
PBODIES=('One workspace for all teams — no more context switching' 'Intelligent automation that learns your workflow patterns' 'SOC2 Type II, 99.99% uptime, deploys in 24 hours')
PICONS=("$ICO_ARROW" "$ICO_CHECK" "$ICO_BAR")
for i in 0 1 2; do
  CX=$((M + i * 3744000))
  # card bg
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_card${i}" --prop preset=roundRect \
    --prop fill=$CARD --prop line=$LIN --prop lineWidth=2 \
    --prop x=$CX --prop y=1872000 --prop width=3420000 --prop height=3600000
  # accent top bar
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_bar${i}" --prop preset=rect --prop fill=$PRI --prop line=none \
    --prop x=$CX --prop y=1872000 --prop width=3420000 --prop height=108000
  # number
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_num${i}" --prop text="0$((i+1))" \
    --prop font='Aptos Display' --prop size=52 --prop bold=false --prop color=$PRI \
    --prop fill=none --prop line=none \
    --prop x=$((CX+216000)) --prop y=2160000 --prop width=1440000 --prop height=1440000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_tit${i}" --prop text="${PTITLES[$i]}" \
    --prop font='Aptos Display' --prop size=20 --prop bold=false --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=$((CX+216000)) --prop y=3528000 --prop width=2880000 --prop height=540000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_bod${i}" --prop text="${PBODIES[$i]}" \
    --prop font=Aptos --prop size=14 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=$((CX+216000)) --prop y=4140000 --prop width=2880000 --prop height=1080000
done

echo "  ✓ S3: pillars (product)"

# ══ SLIDE 4 — evidence（Traction 数据）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[4]' --prop transition=morph
S='/slide[4]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.06 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=1620000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.07 \
  --prop x=9360000 --prop y=2520000 --prop width=3960000 --prop height=3960000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=5040000 --prop width=1800000 --prop height=1800000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.12 \
  --prop x=5760000 --prop y=5400000 --prop width=720000 --prop height=720000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.5 \
  --prop x=3600000 --prop y=3600000 --prop width=216000 --prop height=216000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=54000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=2520000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Traction' \
  --prop font='Aptos Display' --prop size=52 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=5400000 --prop height=1260000

# 4 big numbers
NUMS=('850+' '$42M' '94%' '3.2×')
LABELS=('Enterprise Customers' 'Annual Recurring Revenue' 'Net Revenue Retention' 'YoY Growth')
for i in 0 1 2 3; do
  NX=$((M + i * 2808000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_div${i}" --prop preset=rect --prop fill=$LIN --prop line=none \
    --prop x=$NX --prop y=1980000 --prop width=2016000 --prop height=2 \
    2>/dev/null || true
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_n${i}" --prop text="${NUMS[$i]}" \
    --prop font='Aptos Display' --prop size=56 --prop bold=false --prop color=$PRI \
    --prop fill=none --prop line=none \
    --prop x=$NX --prop y=2160000 --prop width=2520000 --prop height=1440000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_l${i}" --prop text="${LABELS[$i]}" \
    --prop font=Aptos --prop size=14 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=$NX --prop y=3672000 --prop width=2520000 --prop height=504000
done

# Revenue growth chart
officecli add "$OUT" "$S" --type chart \
  --prop chartType=line \
  --prop categories='Q1 23,Q2 23,Q3 23,Q4 23,Q1 24,Q2 24,Q3 24,Q4 24' \
  --prop series1='8,12,18,26,32,38,42,42' \
  --prop colors="$PRI" \
  --prop x=$M --prop y=4320000 --prop width=$FW --prop height=2160000 \
  --prop legend=false --prop dataLabels=false

echo "  ✓ S4: evidence (traction)"

# ══ SLIDE 5 — comparison（竞争对手对比）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[5]' --prop transition=morph
S='/slide[5]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.05 \
  --prop x=7560000 --prop y=0 --prop width=5400000 --prop height=5400000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=5040000 --prop width=1800000 --prop height=1800000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.1 \
  --prop x=6480000 --prop y=5400000 --prop width=900000 --prop height=900000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.5 \
  --prop x=10800000 --prop y=1440000 --prop width=252000 --prop height=252000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=54000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=3240000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Competitive Advantage' \
  --prop font='Aptos Display' --prop size=44 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=8640000 --prop height=1080000

# Comparison table header
officecli add "$OUT" "$S" --type shape \
  --prop name=co_hdr --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=1800000 --prop width=$FW --prop height=540000
HCOLS=('Feature' 'Us' 'Competitor A' 'Competitor B')
HXPOS=(432000 3600000 6768000 9936000)
HWIDS=(2880000 2880000 2880000 2160000)
for i in 0 1 2 3; do
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_h${i}" --prop text="${HCOLS[$i]}" \
    --prop font=Aptos --prop size=13 --prop bold=true --prop color=FAFAFA \
    --prop fill=none --prop line=none \
    --prop x="${HXPOS[$i]}" --prop y=1872000 --prop width="${HWIDS[$i]}" --prop height=360000
done

ROWS=('Unified workspace' 'AI automation' 'Real-time sync' 'SOC2 compliant' 'Open API')
RCOLORS=($PRI MUT MUT MUT MUT)
for i in 0 1 2 3 4; do
  RY=$((2520000 + i * 756000))
  RC=$((i % 2))
  if [ $RC -eq 0 ]; then
    officecli add "$OUT" "$S" --type shape \
      --prop name="co_rbg${i}" --prop preset=rect --prop fill=$BG2 --prop line=none \
      --prop x=$M --prop y=$RY --prop width=$FW --prop height=720000
  fi
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_rlbl${i}" --prop text="${ROWS[$i]}" \
    --prop font=Aptos --prop size=14 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=612000 --prop y=$((RY+180000)) --prop width=2880000 --prop height=360000
  # Us: ✓  Others: partial
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_us${i}" --prop text='✓' \
    --prop font=Aptos --prop size=18 --prop bold=true --prop color=$PRI \
    --prop fill=none --prop line=none \
    --prop x=3960000 --prop y=$((RY+144000)) --prop width=900000 --prop height=432000
done

echo "  ✓ S5: comparison (competitive)"

# ══ SLIDE 6 — evidence（用钱计划）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[6]' --prop transition=morph
S='/slide[6]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.05 \
  --prop x=0 --prop y=2520000 --prop width=4320000 --prop height=4320000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.07 \
  --prop x=9360000 --prop y=0 --prop width=3600000 --prop height=3600000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=$((H-1440000)) --prop width=$W --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.1 \
  --prop x=5760000 --prop y=360000 --prop width=720000 --prop height=720000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.5 \
  --prop x=2160000 --prop y=5760000 --prop width=252000 --prop height=252000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=54000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=2520000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='Use of Funds' \
  --prop font='Aptos Display' --prop size=44 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=5400000 --prop height=1080000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub --prop text='$20M Series B — 24-month runway to profitability' \
  --prop font=Aptos --prop size=18 --prop color=$BOD \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1656000 --prop width=6480000 --prop height=504000

# Use of funds bar chart
officecli add "$OUT" "$S" --type chart \
  --prop chartType=bar \
  --prop categories='Product R&D,Sales & Mktg,Ops & Infra,G&A,Reserve' \
  --prop series1='40,30,15,10,5' \
  --prop colors="$PRI,$ACC,E8215A,9A9AAA,F0F0F0" \
  --prop x=$M --prop y=2340000 --prop width=6480000 --prop height=3960000 \
  --prop legend=true --prop dataLabels=true

# Right: milestones
add_bullet "$S" m1 "$ICO_CHECK" "$PRI" 'Q1: Launch Enterprise tier — 100 new logos' 15 "$BOD" \
  7200000 2340000 4560000 540000
add_bullet "$S" m2 "$ICO_CHECK" "$PRI" 'Q2: Expand APAC — 3 new markets' 15 "$BOD" \
  7200000 3060000 4560000 540000
add_bullet "$S" m3 "$ICO_CHECK" "$PRI" 'Q3: AI features GA — 40% upsell target' 15 "$BOD" \
  7200000 3780000 4560000 540000
add_bullet "$S" m4 "$ICO_CHECK" "$PRI" 'Q4: Reach $65M ARR milestone' 15 "$BOD" \
  7200000 4500000 4560000 540000

echo "  ✓ S6: evidence (use of funds)"

# ══ SLIDE 7 — pillars（团队）══
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[7]' --prop transition=morph
S='/slide[7]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.05 \
  --prop x=8640000 --prop y=3240000 --prop width=4320000 --prop height=4320000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=rect --prop fill=$LIN --prop line=none \
  --prop x=0 --prop y=1440000 --prop width=$W --prop height=2
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.04 \
  --prop x=0 --prop y=4320000 --prop width=1800000 --prop height=1800000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$PRI --prop line=none --prop opacity=0.1 \
  --prop x=11520000 --prop y=5040000 --prop width=720000 --prop height=720000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.5 \
  --prop x=5400000 --prop y=3600000 --prop width=216000 --prop height=216000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=54000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=360000 --prop width=1800000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text='The Team' \
  --prop font='Aptos Display' --prop size=44 --prop bold=false --prop color=$TIT \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=504000 --prop width=5400000 --prop height=1080000

TNAMES=('Sarah Chen' 'Marcus Williams' 'Priya Patel' 'James Rodriguez')
TROLES=('CEO & Co-founder' 'CTO & Co-founder' 'VP Sales' 'VP Engineering')
TBIOS=('Ex-Salesforce CPO, Stanford MBA' 'Ex-Google SWE, MIT CS' 'Ex-Stripe, $200M ARR experience' 'Ex-Databricks, scaled 0→200 eng')
for i in 0 1 2 3; do
  TX=$((M + i * 2880000))
  # avatar placeholder
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_av${i}" --prop preset=ellipse \
    --prop fill=$BG2 --prop line=$PRI --prop lineWidth=3 \
    --prop x=$((TX+720000)) --prop y=1800000 --prop width=1440000 --prop height=1440000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_avlbl${i}" --prop text='Photo' \
    --prop font=Aptos --prop size=11 --prop color=$MUT \
    --prop fill=none --prop line=none \
    --prop x=$((TX+720000)) --prop y=2016000 --prop width=1440000 --prop height=1008000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_tn${i}" --prop text="${TNAMES[$i]}" \
    --prop font='Aptos Display' --prop size=16 --prop bold=false --prop color=$TIT \
    --prop fill=none --prop line=none \
    --prop x=$TX --prop y=3384000 --prop width=2700000 --prop height=504000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_tr${i}" --prop text="${TROLES[$i]}" \
    --prop font=Aptos --prop size=12 --prop bold=true --prop color=$PRI \
    --prop fill=none --prop line=none \
    --prop x=$TX --prop y=3888000 --prop width=2700000 --prop height=360000
  officecli add "$OUT" "$S" --type shape \
    --prop name="co_tb${i}" --prop text="${TBIOS[$i]}" \
    --prop font=Aptos --prop size=12 --prop color=$BOD \
    --prop fill=none --prop line=none \
    --prop x=$TX --prop y=4320000 --prop width=2700000 --prop height=504000
done

echo "  ✓ S7: pillars (team)"

# ══ SLIDE 8 — cta（投资邀请）══
officecli add "$OUT" '/' --type slide --prop background=$AC2
officecli set "$OUT" '/slide[8]' --prop transition=morph
S='/slide[8]'

officecli add "$OUT" "$S" --type shape \
  --prop name=sa1 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.2 \
  --prop x=5760000 --prop y=0 --prop width=9000000 --prop height=9000000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa2 --prop preset=ellipse --prop fill=$PRI --prop line=none --prop opacity=0.1 \
  --prop x=0 --prop y=3600000 --prop width=3600000 --prop height=3600000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa3 --prop preset=rect --prop fill=$PRI --prop line=none --prop opacity=0.06 \
  --prop x=0 --prop y=0 --prop width=$W --prop height=2520000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa4 --prop geometry="$ICO_DIAMOND" \
  --prop fill=$ACC --prop line=none --prop opacity=0.25 \
  --prop x=10800000 --prop y=5040000 --prop width=1440000 --prop height=1440000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa5 --prop preset=ellipse --prop fill=$ACC --prop line=none --prop opacity=0.6 \
  --prop x=2160000 --prop y=1440000 --prop width=360000 --prop height=360000
officecli add "$OUT" "$S" --type shape \
  --prop name=sa6 --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=54000

officecli add "$OUT" "$S" --type shape \
  --prop name=hl_tag --prop preset=rect --prop fill=$PRI --prop line=none \
  --prop x=$M --prop y=1440000 --prop width=2880000 --prop height=72000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_title --prop text="Let's build the future together." \
  --prop font='Aptos Display' --prop size=52 --prop bold=false --prop color=FAFAFA \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=1620000 --prop width=9000000 --prop height=2520000
officecli add "$OUT" "$S" --type shape \
  --prop name=hl_sub --prop text='Raising $20M Series B · Lead investor sought · Close Q2 2025' \
  --prop font=Aptos --prop size=20 --prop color=$ACC \
  --prop fill=none --prop line=none \
  --prop x=$M --prop y=4320000 --prop width=7200000 --prop height=720000

# Contact info
add_bullet "$S" ct1 "$ICO_ARROW" "$PRI" 'invest@company.com' 18 "FAFAFA" \
  $M 5328000 4320000 540000
add_bullet "$S" ct2 "$ICO_ARROW" "$PRI" 'company.com/investors' 18 "FAFAFA" \
  5040000 5328000 4320000 540000

echo "  ✓ S8: cta"

# ══ 验证 ══
echo ""
officecli validate "$OUT"
officecli view outline "$OUT"
echo ""
echo "✅  $OUT"
