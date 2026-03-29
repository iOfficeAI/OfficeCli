#!/bin/bash
# Product Launch Pitch PPT - Morph Edition
# Theme: Pure black + electric purple/violet accent, bold modern
# Morph: geometric shapes expand from center, dramatic transformations
set -euo pipefail

OUT="${1:-$HOME/Desktop/officecli_pptx_suite/product_launch_morph.pptx}"
rm -f "$OUT"
officecli create "$OUT"

W=12192000; H=6858000; M=457200
FULL_W=$((W - 2*M))
# Theme colors
BG=050508          # near-black
BG2=0D0D18         # dark navy-black
ACCENT=7B5EA7      # electric purple
ACCENT2=9B7EC7     # lighter purple
ACCENT3=C4B5E8     # pale lavender
WHITE=F8F8FF       # blue-white
GRAY=888899
DARK_LINE=1A1A2E

# ────────────────────────────────────────────────────────
# SLIDE 1 — COVER: Bold center layout
officecli add "$OUT" '/' --type slide --prop background=$BG
S='/slide[1]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_full' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H

# Large purple circle (geometric hero)
officecli add "$OUT" "$S" --type shape \
  --prop name='!!hero_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.15 \
  --prop x=$((W/2 - 3000000)) --prop y=$((H/2 - 3000000)) --prop width=6000000 --prop height=6000000

# Smaller solid circle
officecli add "$OUT" "$S" --type shape \
  --prop name='!!accent_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.40 \
  --prop x=$((W*3/4 - 800000)) --prop y=$((H/4 - 800000)) --prop width=1600000 --prop height=1600000

# Top accent line
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_accent' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=21600

# Image placeholder (right side)
officecli add "$OUT" "$S" --type shape \
  --prop name='!!img_panel' --prop preset=rect --prop fill=0A0A14 --prop line=$DARK_LINE --prop lineWidth=2 \
  --prop x=$((W*55/100)) --prop y=400000 \
  --prop width=$((W*40/100)) --prop height=$((H - 800000))
officecli add "$OUT" "$S" --type shape \
  --prop name='img_cover_label' --prop text='[ 产品图片占位 ]' --prop font=Calibri --prop size=15 \
  --prop color=1E1E35 --prop fill=none --prop line=none \
  --prop x=$((W*55/100 + 800000)) --prop y=$((H/2 - 180000)) --prop width=3000000 --prop height=360000

# Left content
officecli add "$OUT" "$S" --type shape \
  --prop name='launch_tag' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=1000000 --prop width=1100000 --prop height=120000
officecli add "$OUT" "$S" --type shape \
  --prop name='launch_tag_txt' --prop text='PRODUCT LAUNCH' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=1000000 --prop width=1100000 --prop height=120000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!headline' --prop text='Introducing' --prop font=Georgia --prop size=36 \
  --prop bold=false --prop italic=true --prop color=$ACCENT3 --prop fill=none --prop line=none \
  --prop x=$M --prop y=1300000 --prop width=4500000 --prop height=500000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!product_name' --prop text='Nova Platform' --prop font=Georgia --prop size=72 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=1800000 --prop width=6000000 --prop height=1000000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!tagline' --prop text='Intelligence at scale. Clarity at speed.' --prop font=Calibri --prop size=22 \
  --prop bold=false --prop color=$ACCENT2 --prop fill=none --prop line=none \
  --prop x=$M --prop y=2950000 --prop width=5000000 --prop height=400000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!launch_date' --prop text='Available Q1 2025' --prop font=Calibri --prop size=18 \
  --prop bold=false --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=3500000 --prop width=3000000 --prop height=340000

echo "  ✓ Slide 1: Cover"

# ────────────────────────────────────────────────────────
# SLIDE 2 — THE PROBLEM (full dark, centered)
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[2]' --prop transition=morph
S='/slide[2]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_full' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name='!!hero_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.08 \
  --prop x=$((W - 2000000)) --prop y=$((H - 2000000)) --prop width=4000000 --prop height=4000000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!accent_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.20 \
  --prop x=0 --prop y=0 --prop width=1000000 --prop height=1000000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_accent' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=21600
officecli add "$OUT" "$S" --type shape \
  --prop name='!!img_panel' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=$((W*55/100)) --prop y=400000 --prop width=$((W*40/100)) --prop height=$((H - 800000))

officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=420000 --prop width=800000 --prop height=110000
officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag_txt' --prop text='THE PROBLEM' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=420000 --prop width=800000 --prop height=110000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!headline' --prop text='Teams are drowning in data noise' --prop font=Georgia --prop size=44 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=620000 --prop width=$FULL_W --prop height=700000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!tagline' \
  --prop text='Modern organizations generate more data than ever — yet decision-making is slower, more fragmented, and more expensive.' \
  --prop font=Calibri --prop size=18 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=1440000 --prop width=$FULL_W --prop height=420000

# 3 pain point cards
PAINS=('73% of analysts spend over 4 hours/day just preparing data' 'Average decision lag: 11 days from insight to action' 'Tool fragmentation costs enterprises $340K/year in wasted effort')
PAIN_ICONS=('⏱' '⚡' '💸')
for i in 0 1 2; do
  PX=$((M + i*(3700000 + 150000)))
  officecli add "$OUT" "$S" --type shape \
    --prop name="pain${i}_bg" --prop preset=rect --prop fill=$BG2 --prop line=$DARK_LINE --prop lineWidth=1 \
    --prop x=$PX --prop y=2100000 --prop width=3700000 --prop height=3400000
  officecli add "$OUT" "$S" --type shape \
    --prop name="pain${i}_top" --prop preset=rect --prop fill=$ACCENT --prop line=none \
    --prop x=$PX --prop y=2100000 --prop width=3700000 --prop height=16000
  officecli add "$OUT" "$S" --type shape \
    --prop name="pain${i}_icon" --prop text="${PAIN_ICONS[$i]}" --prop font=Calibri --prop size=48 \
    --prop color=$ACCENT2 --prop fill=none --prop line=none \
    --prop x=$((PX + 200000)) --prop y=2280000 --prop width=800000 --prop height=800000
  officecli add "$OUT" "$S" --type shape \
    --prop name="pain${i}_txt" --prop text="${PAINS[$i]}" --prop font=Calibri --prop size=16 \
    --prop color=$WHITE --prop fill=none --prop line=none \
    --prop x=$((PX + 200000)) --prop y=3200000 --prop width=3300000 --prop height=600000
done

echo "  ✓ Slide 2: The Problem"

# ────────────────────────────────────────────────────────
# SLIDE 3 — THE SOLUTION (split, feature highlights)
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[3]' --prop transition=morph
S='/slide[3]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_full' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name='!!hero_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.12 \
  --prop x=0 --prop y=$((H - 3000000)) --prop width=3000000 --prop height=3000000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!accent_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.30 \
  --prop x=$((W - 1200000)) --prop y=$((H/4)) --prop width=1200000 --prop height=1200000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_accent' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=21600
officecli add "$OUT" "$S" --type shape \
  --prop name='!!img_panel' --prop preset=rect --prop fill=0A0A14 --prop line=$DARK_LINE --prop lineWidth=2 \
  --prop x=$((W*55/100)) --prop y=400000 --prop width=$((W*40/100)) --prop height=$((H - 800000))
officecli add "$OUT" "$S" --type shape \
  --prop name='img_solution_label' --prop text='[ 产品界面截图 ]' --prop font=Calibri --prop size=14 \
  --prop color=1E1E35 --prop fill=none --prop line=none \
  --prop x=$((W*55/100 + 800000)) --prop y=$((H/2 - 180000)) --prop width=3000000 --prop height=360000

officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=420000 --prop width=900000 --prop height=110000
officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag_txt' --prop text='THE SOLUTION' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=420000 --prop width=900000 --prop height=110000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!headline' --prop text='Nova Platform' --prop font=Georgia --prop size=48 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=620000 --prop width=5500000 --prop height=640000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!tagline' \
  --prop text='One unified intelligence layer that connects your data, surfaces insights instantly, and accelerates decisions by 10x.' \
  --prop font=Calibri --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=1360000 --prop width=5500000 --prop height=420000

# 4 feature bullets
FEATURES=('Unified data ingestion from 200+ sources in real-time' 'AI-powered anomaly detection and insight surfacing' 'Collaborative decision workspace with audit trail' 'Enterprise-grade security: SOC2, GDPR, ISO 27001')
FEAT_COLORS=($ACCENT2 $ACCENT $ACCENT3 $GRAY)
for i in 0 1 2 3; do
  FY=$((2000000 + i*1100000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="feat${i}_line" --prop preset=rect --prop fill="${FEAT_COLORS[$i]}" --prop line=none \
    --prop x=$M --prop y=$FY --prop width=100000 --prop height=720000
  officecli add "$OUT" "$S" --type shape \
    --prop name="feat${i}_txt" --prop text="${FEATURES[$i]}" --prop font=Calibri --prop size=16 \
    --prop color=$WHITE --prop fill=none --prop line=none \
    --prop x=$((M + 250000)) --prop y=$((FY + 100000)) --prop width=5200000 --prop height=400000
done

echo "  ✓ Slide 3: The Solution"

# ────────────────────────────────────────────────────────
# SLIDE 4 — TRACTION & METRICS
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[4]' --prop transition=morph
S='/slide[4]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_full' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name='!!hero_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.10 \
  --prop x=2096000 --prop y=0 --prop width=8000000 --prop height=8000000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!accent_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.25 \
  --prop x=$((W - 800000)) --prop y=0 --prop width=1600000 --prop height=1600000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_accent' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=21600
officecli add "$OUT" "$S" --type shape \
  --prop name='!!img_panel' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=$((W*55/100)) --prop y=400000 --prop width=$((W*40/100)) --prop height=$((H - 800000))

officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=420000 --prop width=800000 --prop height=110000
officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag_txt' --prop text='TRACTION' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=420000 --prop width=800000 --prop height=110000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!headline' --prop text='Early Results Speak' --prop font=Georgia --prop size=48 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=620000 --prop width=$FULL_W --prop height=640000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!tagline' --prop text='Beta program results from 12 enterprise customers across 6 industries.' \
  --prop font=Calibri --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=1360000 --prop width=$FULL_W --prop height=340000

# Big stat + chart
STATS_VALS=('10x' '340K' '94%' '18 days')
STATS_LABELS=('Faster decision-making' 'Annual savings per team' 'User satisfaction score' 'Avg. time-to-value')
STATS_COLORS=($ACCENT $ACCENT2 $ACCENT3 $WHITE)

for i in 0 1 2 3; do
  SX=$((M + i*(2700000 + 160000)))
  officecli add "$OUT" "$S" --type shape \
    --prop name="stat${i}_bg" --prop preset=rect --prop fill=$BG2 --prop line=$DARK_LINE --prop lineWidth=1 \
    --prop x=$SX --prop y=1920000 --prop width=2700000 --prop height=2200000
  officecli add "$OUT" "$S" --type shape \
    --prop name="stat${i}_top" --prop preset=rect --prop fill="${STATS_COLORS[$i]}" --prop line=none \
    --prop x=$SX --prop y=1920000 --prop width=2700000 --prop height=16000
  officecli add "$OUT" "$S" --type shape \
    --prop name="stat${i}_val" --prop text="${STATS_VALS[$i]}" --prop font=Georgia --prop size=60 \
    --prop bold=false --prop color="${STATS_COLORS[$i]}" --prop fill=none --prop line=none \
    --prop x=$((SX + 200000)) --prop y=2080000 --prop width=2300000 --prop height=900000
  officecli add "$OUT" "$S" --type shape \
    --prop name="stat${i}_label" --prop text="${STATS_LABELS[$i]}" --prop font=Calibri --prop size=15 \
    --prop color=$GRAY --prop fill=none --prop line=none \
    --prop x=$((SX + 200000)) --prop y=3060000 --prop width=2300000 --prop height=400000
done

# Customer logos placeholder
officecli add "$OUT" "$S" --type shape \
  --prop name='logos_label' --prop text='Trusted by early beta customers including:' \
  --prop font=Calibri --prop size=15 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=4400000 --prop width=$FULL_W --prop height=300000

officecli add "$OUT" "$S" --type shape \
  --prop name='logos_ph' --prop preset=rect --prop fill=$BG2 --prop line=$DARK_LINE --prop lineWidth=1 \
  --prop x=$M --prop y=4780000 --prop width=$FULL_W --prop height=700000
officecli add "$OUT" "$S" --type shape \
  --prop name='logos_ph_txt' --prop text='[ Company A ]     [ Company B ]     [ Company C ]     [ Company D ]     [ Company E ]' \
  --prop font=Calibri --prop size=16 --prop color=333355 --prop fill=none --prop line=none \
  --prop x=$((M + 300000)) --prop y=5020000 --prop width=$((FULL_W - 600000)) --prop height=300000

echo "  ✓ Slide 4: Traction"

# ────────────────────────────────────────────────────────
# SLIDE 5 — CALL TO ACTION
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[5]' --prop transition=morph
S='/slide[5]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_full' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
# Large centered purple gradient feel
officecli add "$OUT" "$S" --type shape \
  --prop name='!!hero_circle' --prop preset=ellipse --prop fill=$ACCENT --prop line=none --prop opacity=0.25 \
  --prop x=1596000 --prop y=0 --prop width=9000000 --prop height=9000000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!accent_circle' --prop preset=ellipse --prop fill=$ACCENT2 --prop line=none --prop opacity=0.50 \
  --prop x=$((W/2 - 2000000)) --prop y=$((H/2 - 2000000)) --prop width=4000000 --prop height=4000000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_accent' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=21600
officecli add "$OUT" "$S" --type shape \
  --prop name='!!img_panel' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=$((W*55/100)) --prop y=400000 --prop width=$((W*40/100)) --prop height=$((H - 800000))

officecli add "$OUT" "$S" --type shape \
  --prop name='!!headline' --prop text='Ready to transform how your team decides?' --prop font=Georgia --prop size=44 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=1600000 --prop width=$FULL_W --prop height=800000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!tagline' \
  --prop text='Join 500+ teams already on the waitlist. Early access launches Q1 2025.' \
  --prop font=Calibri --prop size=20 --prop color=$ACCENT3 --prop fill=none --prop line=none \
  --prop x=$M --prop y=2600000 --prop width=$FULL_W --prop height=400000

# CTA boxes
officecli add "$OUT" "$S" --type shape \
  --prop name='cta1_bg' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=3300000 --prop width=2800000 --prop height=600000
officecli add "$OUT" "$S" --type shape \
  --prop name='cta1_txt' --prop text='Request Early Access' --prop font=Calibri --prop size=18 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=3380000 --prop width=2800000 --prop height=420000

officecli add "$OUT" "$S" --type shape \
  --prop name='cta2_bg' --prop preset=rect --prop fill=none --prop line=$ACCENT --prop lineWidth=2 \
  --prop x=$((M + 3000000)) --prop y=3300000 --prop width=2800000 --prop height=600000
officecli add "$OUT" "$S" --type shape \
  --prop name='cta2_txt' --prop text='Schedule a Demo' --prop font=Calibri --prop size=18 \
  --prop bold=false --prop color=$ACCENT2 --prop fill=none --prop line=none \
  --prop x=$((M + 3000000)) --prop y=3380000 --prop width=2800000 --prop height=420000

officecli add "$OUT" "$S" --type shape \
  --prop name='contact' --prop text='nova@example.com  |  nova-platform.io' \
  --prop font=Calibri --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=4200000 --prop width=$FULL_W --prop height=320000

echo "  ✓ Slide 5: Call to Action"

officecli validate "$OUT" >/dev/null
echo ""
echo "✅ Generated: $OUT"
echo "   Slides: 5  |  Theme: Dark Purple / Bold Modern  |  Transition: Morph"
