#!/bin/bash
# Budget Review PPT - Morph Transition Edition
# Layout: W=12192000, H=6858000 (16:9 standard)
# Design: Dark editorial, gold/green palette, morph split-panel effect
set -euo pipefail

OUT="${1:-$HOME/Desktop/officecli_pptx_suite/budget_review_morph.pptx}"
rm -f "$OUT"
officecli create "$OUT"

# ─────────────────────────────────────────────
# Constants (EMU)  16:9 = 12192000 x 6858000
W=12192000; H=6858000; M=457200
LEFT_W=5486400          # ~45% width
DIV_W=14400             # gold divider bar width
DIV_X=5486400
RIGHT_X=$((DIV_X + DIV_W))
RIGHT_W=$((W - RIGHT_X))
FULL_W=$((W - 2*M))    # 11277600 usable

# ─────────────────────────────────────────────
# SLIDE 1 — COVER
officecli add "$OUT" '/' --type slide --prop background=080808
S1='/slide[1]'

officecli add "$OUT" "$S1" --type shape \
  --prop name='!!panel_left' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=0 --prop y=0 --prop width=$LEFT_W --prop height=$H

officecli add "$OUT" "$S1" --type shape \
  --prop name='!!panel_divider' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$DIV_X --prop y=0 --prop width=$DIV_W --prop height=$H

# Right: image placeholder (dark green with label)
officecli add "$OUT" "$S1" --type shape \
  --prop name='!!panel_right' --prop preset=rect --prop fill=111E0A --prop line=none \
  --prop x=$RIGHT_X --prop y=0 --prop width=$RIGHT_W --prop height=$H

officecli add "$OUT" "$S1" --type shape \
  --prop name='img_placeholder_cover' --prop preset=rect --prop fill=162810 --prop line=243D1C --prop lineWidth=3 \
  --prop x=$((RIGHT_X + 120000)) --prop y=120000 --prop width=$((RIGHT_W - 240000)) --prop height=$((H - 240000))
officecli add "$OUT" "$S1" --type shape \
  --prop name='img_ph_label' --prop text='📷 绿色流光背景图片' --prop font=Calibri --prop size=16 \
  --prop color=2E5A22 --prop fill=none --prop line=none \
  --prop x=$((RIGHT_X + 1800000)) --prop y=$((H/2 - 180000)) --prop width=3000000 --prop height=360000

# Cover tag
officecli add "$OUT" "$S1" --type shape \
  --prop name='cover_tag' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=700000 --prop width=800000 --prop height=110000
officecli add "$OUT" "$S1" --type shape \
  --prop name='cover_tag_txt' --prop text='BUDGET REVIEW' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=080808 --prop fill=none --prop line=none \
  --prop x=$M --prop y=700000 --prop width=800000 --prop height=110000

# Main headline
officecli add "$OUT" "$S1" --type shape \
  --prop name='!!headline' --prop text='Budget Review' --prop font=Georgia --prop size=64 \
  --prop bold=false --prop color=C9A84C --prop fill=none --prop line=none \
  --prop x=$M --prop y=2200000 --prop width=4600000 --prop height=900000

# Subline
officecli add "$OUT" "$S1" --type shape \
  --prop name='!!subline' --prop text='Strategic Financial Analysis & Planning' --prop font=Calibri --prop size=20 \
  --prop bold=false --prop color=777777 --prop fill=none --prop line=none \
  --prop x=$M --prop y=3220000 --prop width=4600000 --prop height=400000

# Year badge
officecli add "$OUT" "$S1" --type shape \
  --prop name='!!year_badge' --prop text='FY 2024' --prop font=Calibri --prop size=14 \
  --prop bold=true --prop color=080808 --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=6100000 --prop width=520000 --prop height=230000

echo "  ✓ Slide 1: Cover"

# ─────────────────────────────────────────────
# SLIDE 2 — PREVIOUS BUDGET PERIOD OVERVIEW
# Morph effect: panel_left expands to 85%, divider & right panel shift right
officecli add "$OUT" '/' --type slide --prop background=080808
officecli set "$OUT" '/slide[2]' --prop transition=morph
S2='/slide[2]'

P2_LEFT=$((W * 85 / 100))   # 10363200
DIV2_X=$P2_LEFT
RIGHT2_X=$((P2_LEFT + DIV_W))
RIGHT2_W=$((W - RIGHT2_X))

officecli add "$OUT" "$S2" --type shape \
  --prop name='!!panel_left' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=0 --prop y=0 --prop width=$P2_LEFT --prop height=$H

officecli add "$OUT" "$S2" --type shape \
  --prop name='!!panel_divider' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$DIV2_X --prop y=0 --prop width=$DIV_W --prop height=$H

officecli add "$OUT" "$S2" --type shape \
  --prop name='!!panel_right' --prop preset=rect --prop fill=111E0A --prop line=none \
  --prop x=$RIGHT2_X --prop y=0 --prop width=$RIGHT2_W --prop height=$H

# Section label
officecli add "$OUT" "$S2" --type shape \
  --prop name='s2_tag_bg' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=480000 --prop width=700000 --prop height=110000
officecli add "$OUT" "$S2" --type shape \
  --prop name='s2_tag_txt' --prop text='OVERVIEW' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=080808 --prop fill=none --prop line=none \
  --prop x=$M --prop y=480000 --prop width=700000 --prop height=110000

officecli add "$OUT" "$S2" --type shape \
  --prop name='!!headline' --prop text='Previous Budget Period Overview' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=F0F0F0 --prop fill=none --prop line=none \
  --prop x=$M --prop y=660000 --prop width=9600000 --prop height=640000

# 4 info cards: 2 cols × 2 rows
CARD_W=4900000; CARD_H=1560000; CARD_GAP=200000
COL2_X=$((M + CARD_W + CARD_GAP))

add_card() {
  local idx=$1 col_x=$2 row_y=$3 title=$4 sub=$5
  officecli add "$OUT" "$S2" --type shape \
    --prop name="card${idx}_bg" --prop preset=rect --prop fill=111111 --prop line=1C3A14 --prop lineWidth=1 \
    --prop x=$col_x --prop y=$row_y --prop width=$CARD_W --prop height=$CARD_H
  officecli add "$OUT" "$S2" --type shape \
    --prop name="card${idx}_num" --prop text=$idx --prop font=Calibri --prop size=18 \
    --prop bold=true --prop color=C9A84C --prop fill=1C3A14 --prop line=none \
    --prop x=$((col_x + 140000)) --prop y=$((row_y + 200000)) --prop width=480000 --prop height=480000
  officecli add "$OUT" "$S2" --type shape \
    --prop name="card${idx}_title" --prop text="$title" --prop font=Georgia --prop size=21 \
    --prop bold=false --prop color=EEEEEE --prop fill=none --prop line=none \
    --prop x=$((col_x + 720000)) --prop y=$((row_y + 200000)) --prop width=$((CARD_W - 840000)) --prop height=380000
  officecli add "$OUT" "$S2" --type shape \
    --prop name="card${idx}_sub" --prop text="$sub" --prop font=Calibri --prop size=15 \
    --prop bold=false --prop color=777777 --prop fill=none --prop line=none \
    --prop x=$((col_x + 720000)) --prop y=$((row_y + 640000)) --prop width=$((CARD_W - 840000)) --prop height=340000
}

add_card 1 $M 1600000 'Budget Allocated' 'Total funding approved for the period'
add_card 2 $COL2_X 1600000 'Period Covered' 'Timeframe under review'
add_card 3 $M 3280000 'Primary Objectives' 'Key goals and strategic priorities'
add_card 4 $COL2_X 3280000 'Major Initiatives' 'Significant projects funded'

officecli add "$OUT" "$S2" --type shape \
  --prop name='s2_footer' \
  --prop text='This section provides context for the financial period being reviewed, establishing the baseline for our analysis and highlighting the strategic framework that guided budget allocation decisions.' \
  --prop font=Calibri --prop size=13 --prop color=444444 --prop fill=none --prop line=none \
  --prop x=$M --prop y=5900000 --prop width=9600000 --prop height=400000

echo "  ✓ Slide 2: Overview"

# ─────────────────────────────────────────────
# SLIDE 3 — BUDGET ALLOCATION (3 KPI cards, full-width dark)
officecli add "$OUT" '/' --type slide --prop background=080808
officecli set "$OUT" '/slide[3]' --prop transition=morph
S3='/slide[3]'

# Divider moves to far right (off-screen feel)
officecli add "$OUT" "$S3" --type shape \
  --prop name='!!panel_left' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S3" --type shape \
  --prop name='!!panel_divider' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$((W - DIV_W)) --prop y=0 --prop width=$DIV_W --prop height=$H
officecli add "$OUT" "$S3" --type shape \
  --prop name='!!panel_right' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=$W --prop y=0 --prop width=100 --prop height=$H

officecli add "$OUT" "$S3" --type shape \
  --prop name='s3_tag_bg' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=480000 --prop width=1050000 --prop height=110000
officecli add "$OUT" "$S3" --type shape \
  --prop name='s3_tag_txt' --prop text='BUDGET OVERVIEW' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=080808 --prop fill=none --prop line=none \
  --prop x=$M --prop y=480000 --prop width=1050000 --prop height=110000

officecli add "$OUT" "$S3" --type shape \
  --prop name='!!headline' --prop text='Budget Allocation Summary' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=F0F0F0 --prop fill=none --prop line=none \
  --prop x=$M --prop y=660000 --prop width=$FULL_W --prop height=640000

officecli add "$OUT" "$S3" --type shape \
  --prop name='!!subline' \
  --prop text='Comprehensive breakdown of how funds were distributed across organizational units and strategic initiatives during the review period.' \
  --prop font=Calibri --prop size=16 --prop color=777777 --prop fill=none --prop line=none \
  --prop x=$M --prop y=1380000 --prop width=$FULL_W --prop height=340000

# 3 KPI cards
KPI_W=3520000; KPI_H=2600000; KPI_Y=2000000; KPI_GAP=248000
KPI2_X=$((M + KPI_W + KPI_GAP))
KPI3_X=$((M + (KPI_W + KPI_GAP)*2))

for kpi in 1 2 3; do
  case $kpi in
    1) KPI_X=$M;      KVAL='$4.2M'; KLABEL='Total Budget Allocated';  KCOLOR=C9A84C; ACCENT=C9A84C ;;
    2) KPI_X=$KPI2_X; KVAL='$3.8M'; KLABEL='Total Budget Spent';      KCOLOR=7AB648; ACCENT=3A6A25 ;;
    3) KPI_X=$KPI3_X; KVAL='90.5%'; KLABEL='Utilization Rate';        KCOLOR=EEEEEE; ACCENT=555555 ;;
  esac
  officecli add "$OUT" "$S3" --type shape \
    --prop name="kpi${kpi}_bg" --prop preset=rect --prop fill=111111 --prop line=none \
    --prop x=$KPI_X --prop y=$KPI_Y --prop width=$KPI_W --prop height=$KPI_H
  officecli add "$OUT" "$S3" --type shape \
    --prop name="kpi${kpi}_top" --prop preset=rect --prop fill=$ACCENT --prop line=none \
    --prop x=$KPI_X --prop y=$KPI_Y --prop width=$KPI_W --prop height=16000
  officecli add "$OUT" "$S3" --type shape \
    --prop name="kpi${kpi}_val" --prop text="$KVAL" --prop font=Georgia --prop size=56 \
    --prop bold=false --prop color=$KCOLOR --prop fill=none --prop line=none \
    --prop x=$((KPI_X + 240000)) --prop y=$((KPI_Y + 300000)) --prop width=$((KPI_W - 480000)) --prop height=900000
  officecli add "$OUT" "$S3" --type shape \
    --prop name="kpi${kpi}_label" --prop text="$KLABEL" --prop font=Calibri --prop size=17 \
    --prop bold=false --prop color=AAAAAA --prop fill=none --prop line=none \
    --prop x=$((KPI_X + 240000)) --prop y=$((KPI_Y + 1300000)) --prop width=$((KPI_W - 480000)) --prop height=380000
done

officecli add "$OUT" "$S3" --type shape \
  --prop name='insight' \
  --prop text='▸  Key Insight: Utilization exceeded 85% target threshold across all departments, with strategic reserve maintained.' \
  --prop font=Calibri --prop size=14 --prop color=555555 --prop fill=none --prop line=none \
  --prop x=$M --prop y=6300000 --prop width=$FULL_W --prop height=300000

echo "  ✓ Slide 3: Budget Allocation"

# ─────────────────────────────────────────────
# SLIDE 4 — EXPENSE BREAKDOWN (left text + right pie chart)
officecli add "$OUT" '/' --type slide --prop background=080808
officecli set "$OUT" '/slide[4]' --prop transition=morph
S4='/slide[4]'

SPLIT4=6400000
DIV4_X=$SPLIT4; RIGHT4_X=$((SPLIT4 + DIV_W)); RIGHT4_W=$((W - RIGHT4_X))

officecli add "$OUT" "$S4" --type shape \
  --prop name='!!panel_left' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=0 --prop y=0 --prop width=$SPLIT4 --prop height=$H
officecli add "$OUT" "$S4" --type shape \
  --prop name='!!panel_divider' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$DIV4_X --prop y=0 --prop width=$DIV_W --prop height=$H
officecli add "$OUT" "$S4" --type shape \
  --prop name='!!panel_right' --prop preset=rect --prop fill=0E1E08 --prop line=none \
  --prop x=$RIGHT4_X --prop y=0 --prop width=$RIGHT4_W --prop height=$H

officecli add "$OUT" "$S4" --type shape \
  --prop name='s4_tag_bg' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=480000 --prop width=1150000 --prop height=110000
officecli add "$OUT" "$S4" --type shape \
  --prop name='s4_tag_txt' --prop text='SPENDING ANALYSIS' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=080808 --prop fill=none --prop line=none \
  --prop x=$M --prop y=480000 --prop width=1150000 --prop height=110000

officecli add "$OUT" "$S4" --type shape \
  --prop name='!!headline' --prop text='Expense Breakdown' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=F0F0F0 --prop fill=none --prop line=none \
  --prop x=$M --prop y=660000 --prop width=5600000 --prop height=640000
officecli add "$OUT" "$S4" --type shape \
  --prop name='!!subline' \
  --prop text='Understanding how resources were allocated across different categories provides insight into organizational priorities and spending patterns.' \
  --prop font=Calibri --prop size=14 --prop color=777777 --prop fill=none --prop line=none \
  --prop x=$M --prop y=1420000 --prop width=5600000 --prop height=420000

officecli add "$OUT" "$S4" --type shape \
  --prop name='cat_heading' --prop text='Major Spending Categories' --prop font=Georgia --prop size=22 \
  --prop bold=false --prop color=DDDDDD --prop fill=none --prop line=none \
  --prop x=$M --prop y=2000000 --prop width=5600000 --prop height=380000

officecli add "$OUT" "$S4" --type shape \
  --prop name='cat_sub' \
  --prop text='This distribution reflects strategic priorities and operational requirements for the review period.' \
  --prop font=Calibri --prop size=13 --prop color=555555 --prop fill=none --prop line=none \
  --prop x=$M --prop y=2430000 --prop width=5600000 --prop height=320000

# 4 category rows
CAT1='Category 1 represents the largest portion of expenses'
CAT2='Category 2 accounts for significant operational costs'
CAT3='Category 3 includes essential support functions'
CAT4='Category 4 covers additional strategic investments'
idx=0
for cat_text in "$CAT1" "$CAT2" "$CAT3" "$CAT4"; do
  idx=$((idx+1))
  ITEM_Y=$((2860000 + (idx-1)*820000))
  officecli add "$OUT" "$S4" --type shape \
    --prop name="catbox${idx}" --prop preset=rect --prop fill=none --prop line=252525 --prop lineWidth=1 \
    --prop x=$M --prop y=$ITEM_Y --prop width=5600000 --prop height=640000
  officecli add "$OUT" "$S4" --type shape \
    --prop name="cattext${idx}" --prop text="$cat_text" --prop font=Calibri --prop size=14 \
    --prop color=BBBBBB --prop fill=none --prop line=none \
    --prop x=$((M + 200000)) --prop y=$((ITEM_Y + 200000)) --prop width=5200000 --prop height=280000
done

# Pie chart on right panel
officecli add "$OUT" "$S4" --type chart \
  --prop chartType=pie \
  --prop categories='Category 1,Category 2,Category 3,Category 4' \
  --prop series1='35,28,22,15' \
  --prop colors='4A7A35,3A6A25,5A8A45,2A4A15' \
  --prop x=$((RIGHT4_X + 200000)) --prop y=700000 \
  --prop width=$((RIGHT4_W - 400000)) --prop height=5400000 \
  --prop legend=true --prop dataLabels=true

echo "  ✓ Slide 4: Expense Breakdown"

# ─────────────────────────────────────────────
# SLIDE 5 — VARIANCE ANALYSIS (bar chart, full width)
officecli add "$OUT" '/' --type slide --prop background=080808
officecli set "$OUT" '/slide[5]' --prop transition=morph
S5='/slide[5]'

officecli add "$OUT" "$S5" --type shape \
  --prop name='!!panel_left' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S5" --type shape \
  --prop name='!!panel_divider' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$((W - DIV_W)) --prop y=0 --prop width=$DIV_W --prop height=$H
officecli add "$OUT" "$S5" --type shape \
  --prop name='!!panel_right' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=$W --prop y=0 --prop width=100 --prop height=$H

officecli add "$OUT" "$S5" --type shape \
  --prop name='s5_tag_bg' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=480000 --prop width=1150000 --prop height=110000
officecli add "$OUT" "$S5" --type shape \
  --prop name='s5_tag_txt' --prop text='VARIANCE ANALYSIS' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=080808 --prop fill=none --prop line=none \
  --prop x=$M --prop y=480000 --prop width=1150000 --prop height=110000

officecli add "$OUT" "$S5" --type shape \
  --prop name='!!headline' --prop text='Budget vs. Actual Variance' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=F0F0F0 --prop fill=none --prop line=none \
  --prop x=$M --prop y=660000 --prop width=$FULL_W --prop height=640000
officecli add "$OUT" "$S5" --type shape \
  --prop name='!!subline' \
  --prop text='Detailed comparison of planned versus actual expenditures, highlighting areas of significant deviation and their underlying causes.' \
  --prop font=Calibri --prop size=16 --prop color=777777 --prop fill=none --prop line=none \
  --prop x=$M --prop y=1360000 --prop width=$FULL_W --prop height=340000

officecli add "$OUT" "$S5" --type chart \
  --prop chartType=bar \
  --prop title='Department Budget vs Actual ($)' \
  --prop categories='Operations,Marketing,R&D,HR,IT' \
  --prop series1='1200000,800000,600000,400000,500000' \
  --prop series2='1050000,920000,580000,390000,460000' \
  --prop colors='C9A84C,3A6A25' \
  --prop x=$M --prop y=1900000 --prop width=$FULL_W --prop height=4400000 \
  --prop legend=true --prop dataLabels=false

officecli add "$OUT" "$S5" --type shape \
  --prop name='chart_note' \
  --prop text='Note: Positive variance in Marketing reflects accelerated Q3 campaign spend. R&D and IT remained within approved parameters.' \
  --prop font=Calibri --prop size=13 --prop color=444444 --prop fill=none --prop line=none \
  --prop x=$M --prop y=6400000 --prop width=$FULL_W --prop height=260000

echo "  ✓ Slide 5: Variance Analysis"

# ─────────────────────────────────────────────
# SLIDE 6 — FORWARD OUTLOOK (50% split, priorities + image placeholder)
officecli add "$OUT" '/' --type slide --prop background=080808
officecli set "$OUT" '/slide[6]' --prop transition=morph
S6='/slide[6]'

SPLIT6=$((W / 2))  # 6096000
DIV6_X=$SPLIT6; RIGHT6_X=$((SPLIT6 + DIV_W)); RIGHT6_W=$((W - RIGHT6_X))

officecli add "$OUT" "$S6" --type shape \
  --prop name='!!panel_left' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=0 --prop y=0 --prop width=$SPLIT6 --prop height=$H
officecli add "$OUT" "$S6" --type shape \
  --prop name='!!panel_divider' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$DIV6_X --prop y=0 --prop width=$DIV_W --prop height=$H
officecli add "$OUT" "$S6" --type shape \
  --prop name='!!panel_right' --prop preset=rect --prop fill=0C1A08 --prop line=none \
  --prop x=$RIGHT6_X --prop y=0 --prop width=$RIGHT6_W --prop height=$H

# Image placeholder right
officecli add "$OUT" "$S6" --type shape \
  --prop name='img_ph_outlook' --prop preset=rect --prop fill=131F0F --prop line=1E3218 --prop lineWidth=2 \
  --prop x=$((RIGHT6_X + 100000)) --prop y=100000 \
  --prop width=$((RIGHT6_W - 200000)) --prop height=$((H - 200000))
officecli add "$OUT" "$S6" --type shape \
  --prop name='img_ph_outlook_txt' --prop text='📷 流光图片占位' --prop font=Calibri --prop size=14 \
  --prop color=253D1A --prop fill=none --prop line=none \
  --prop x=$((RIGHT6_X + 1000000)) --prop y=$((H/2 - 180000)) --prop width=4000000 --prop height=360000

officecli add "$OUT" "$S6" --type shape \
  --prop name='s6_tag_bg' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=480000 --prop width=1050000 --prop height=110000
officecli add "$OUT" "$S6" --type shape \
  --prop name='s6_tag_txt' --prop text='FORWARD OUTLOOK' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=080808 --prop fill=none --prop line=none \
  --prop x=$M --prop y=480000 --prop width=1050000 --prop height=110000

officecli add "$OUT" "$S6" --type shape \
  --prop name='!!headline' --prop text='Strategic Priorities for Next Period' --prop font=Georgia --prop size=36 \
  --prop bold=false --prop color=F0F0F0 --prop fill=none --prop line=none \
  --prop x=$M --prop y=700000 --prop width=5200000 --prop height=700000
officecli add "$OUT" "$S6" --type shape \
  --prop name='!!subline' \
  --prop text="Building on this period's learnings to drive more efficient and impactful resource allocation." \
  --prop font=Calibri --prop size=15 --prop color=777777 --prop fill=none --prop line=none \
  --prop x=$M --prop y=1520000 --prop width=5200000 --prop height=380000

P_TITLES=('Operational Excellence' 'Strategic Investment' 'Risk Management')
P_TEXTS=('Streamline core processes and reduce overhead by 12% through targeted efficiency programs.' 'Increase R&D allocation by 15% to accelerate innovation pipeline and competitive positioning.' 'Maintain 10% strategic reserve and implement quarterly variance review cadence.')
P_COLORS=(C9A84C 7AB648 888888)

for i in 0 1 2; do
  PY=$((2100000 + i*1500000))
  officecli add "$OUT" "$S6" --type shape \
    --prop name="p$((i+1))_bar" --prop preset=rect --prop fill="${P_COLORS[$i]}" --prop line=none \
    --prop x=$M --prop y=$PY --prop width=110000 --prop height=1200000
  officecli add "$OUT" "$S6" --type shape \
    --prop name="p$((i+1))_title" --prop text="${P_TITLES[$i]}" --prop font=Georgia --prop size=22 \
    --prop bold=false --prop color="${P_COLORS[$i]}" --prop fill=none --prop line=none \
    --prop x=$((M + 270000)) --prop y=$PY --prop width=5000000 --prop height=400000
  officecli add "$OUT" "$S6" --type shape \
    --prop name="p$((i+1))_text" --prop text="${P_TEXTS[$i]}" --prop font=Calibri --prop size=14 \
    --prop bold=false --prop color=999999 --prop fill=none --prop line=none \
    --prop x=$((M + 270000)) --prop y=$((PY + 460000)) --prop width=5000000 --prop height=420000
done

echo "  ✓ Slide 6: Forward Outlook"

# ─────────────────────────────────────────────
# SLIDE 7 — ACTION ITEMS (full width, 3 action rows)
officecli add "$OUT" '/' --type slide --prop background=080808
officecli set "$OUT" '/slide[7]' --prop transition=morph
S7='/slide[7]'

officecli add "$OUT" "$S7" --type shape \
  --prop name='!!panel_left' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S7" --type shape \
  --prop name='!!panel_divider' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$((W - DIV_W)) --prop y=0 --prop width=$DIV_W --prop height=$H
officecli add "$OUT" "$S7" --type shape \
  --prop name='!!panel_right' --prop preset=rect --prop fill=080808 --prop line=none \
  --prop x=$W --prop y=0 --prop width=100 --prop height=$H

officecli add "$OUT" "$S7" --type shape \
  --prop name='s7_tag_bg' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=480000 --prop width=700000 --prop height=110000
officecli add "$OUT" "$S7" --type shape \
  --prop name='s7_tag_txt' --prop text='NEXT STEPS' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=080808 --prop fill=none --prop line=none \
  --prop x=$M --prop y=480000 --prop width=700000 --prop height=110000

officecli add "$OUT" "$S7" --type shape \
  --prop name='!!headline' --prop text='Action Items & Decision Points' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=F0F0F0 --prop fill=none --prop line=none \
  --prop x=$M --prop y=660000 --prop width=$FULL_W --prop height=640000
officecli add "$OUT" "$S7" --type shape \
  --prop name='!!subline' \
  --prop text='Clear ownership and timelines for each priority item emerging from this budget review.' \
  --prop font=Calibri --prop size=16 --prop color=777777 --prop fill=none --prop line=none \
  --prop x=$M --prop y=1360000 --prop width=$FULL_W --prop height=340000

A_TITLES=('Approve Q1 budget reallocation proposals' 'Implement variance tracking dashboard' 'Conduct departmental budget alignment workshops')
A_OWNERS=('CFO / Finance Committee' 'FP&A Team / IT' 'Department Heads / HR')
A_DATES=('Q1 2025' 'Q2 2025' 'Q2 2025')
A_COLORS=(C9A84C 3A6A25 666666)

for i in 0 1 2; do
  AY=$((1980000 + i*1560000))
  officecli add "$OUT" "$S7" --type shape \
    --prop name="a$((i+1))_bg" --prop preset=rect --prop fill=111111 --prop line=none \
    --prop x=$M --prop y=$AY --prop width=$FULL_W --prop height=1360000
  officecli add "$OUT" "$S7" --type shape \
    --prop name="a$((i+1))_dot" --prop preset=ellipse --prop fill="${A_COLORS[$i]}" --prop line=none \
    --prop x=$((M + 200000)) --prop y=$((AY + 440000)) --prop width=440000 --prop height=440000
  officecli add "$OUT" "$S7" --type shape \
    --prop name="a$((i+1))_title" --prop text="${A_TITLES[$i]}" --prop font=Georgia --prop size=21 \
    --prop bold=false --prop color=EEEEEE --prop fill=none --prop line=none \
    --prop x=$((M + 840000)) --prop y=$((AY + 280000)) --prop width=9000000 --prop height=380000
  officecli add "$OUT" "$S7" --type shape \
    --prop name="a$((i+1))_owner" --prop text="${A_OWNERS[$i]}" --prop font=Calibri --prop size=14 \
    --prop bold=false --prop color=666666 --prop fill=none --prop line=none \
    --prop x=$((M + 840000)) --prop y=$((AY + 720000)) --prop width=7000000 --prop height=300000
  officecli add "$OUT" "$S7" --type shape \
    --prop name="a$((i+1))_date" --prop text="${A_DATES[$i]}" --prop font=Calibri --prop size=15 \
    --prop bold=true --prop color="${A_COLORS[$i]}" --prop fill=none --prop line=none \
    --prop x=$((M + FULL_W - 900000)) --prop y=$((AY + 520000)) --prop width=800000 --prop height=300000
done

officecli add "$OUT" "$S7" --type shape \
  --prop name='closing_line' --prop preset=rect --prop fill=C9A84C --prop line=none \
  --prop x=$M --prop y=6600000 --prop width=$FULL_W --prop height=14400
officecli add "$OUT" "$S7" --type shape \
  --prop name='closing' \
  --prop text='Together, data-driven decisions create accountability and drive sustainable financial performance across the organization.' \
  --prop font=Georgia --prop size=16 --prop bold=false --prop color=555555 --prop fill=none --prop line=none \
  --prop x=$M --prop y=6640000 --prop width=$FULL_W --prop height=180000

echo "  ✓ Slide 7: Action Items"

# ─────────────────────────────────────────────
officecli validate "$OUT" >/dev/null
echo ""
echo "✅ Generated: $OUT"
echo "   Slides: 7  |  Theme: Dark Editorial  |  Transition: Morph"
echo "   Image placeholders: Slide 1 (cover right), Slide 6 (outlook right)"
echo "   To add real images: officecli set <file> /slide[N]/picture[1] --prop path=<image.jpg>"
