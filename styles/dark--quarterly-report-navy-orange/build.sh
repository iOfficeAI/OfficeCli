#!/bin/bash
# Quarterly Business Report PPT - Morph Edition
# Theme: Navy blue + white + orange accent, corporate minimal
# Morph: horizontal bar slides across, content blocks expand
set -euo pipefail

OUT="${1:-$HOME/Desktop/officecli_pptx_suite/quarterly_report_morph.pptx}"
rm -f "$OUT"
officecli create "$OUT"

W=12192000; H=6858000; M=457200
FULL_W=$((W - 2*M))
# Theme colors
BG=0D1B2A        # deep navy
BG2=132233       # slightly lighter navy
ACCENT=E8793A    # warm orange
ACCENT2=F5A05A   # light orange
WHITE=F2F4F7
GRAY=8899AA
DARK_LINE=1E3048

# ────────────────────────────────────────────────────────
# SLIDE 1 — COVER: diagonal split (top-left navy, bottom-right white)
officecli add "$OUT" '/' --type slide --prop background=$BG
S='/slide[1]'

# Full navy background
officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_main' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H

# Orange accent top bar
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_bar' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=28800

# Image placeholder (right half)
officecli add "$OUT" "$S" --type shape \
  --prop name='!!img_panel' --prop preset=rect --prop fill=132233 --prop line=$DARK_LINE --prop lineWidth=2 \
  --prop x=$((W/2)) --prop y=0 --prop width=$((W/2)) --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name='img_label_cover' --prop text='[ 封面图片占位 ]' --prop font=Calibri --prop size=16 \
  --prop color=1E3048 --prop fill=none --prop line=none \
  --prop x=$((W/2 + 1200000)) --prop y=$((H/2 - 180000)) --prop width=3500000 --prop height=360000

# Left content
officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_quarter' --prop text='Q3 2024' --prop font=Calibri --prop size=18 \
  --prop bold=true --prop color=$ACCENT --prop fill=none --prop line=none \
  --prop x=$M --prop y=1200000 --prop width=3000000 --prop height=360000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_title' --prop text='Quarterly Business Report' --prop font=Georgia --prop size=54 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=1700000 --prop width=5500000 --prop height=1400000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_sub' --prop text='Performance Review & Strategic Update' --prop font=Calibri --prop size=20 \
  --prop bold=false --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=3200000 --prop width=5500000 --prop height=400000

# Divider line
officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_line' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=3740000 --prop width=800000 --prop height=14400

# Company/dept
officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_org' --prop text='Finance & Strategy Division' --prop font=Calibri --prop size=16 \
  --prop bold=false --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=3900000 --prop width=3000000 --prop height=300000

echo "  ✓ Slide 1: Cover"

# ────────────────────────────────────────────────────────
# SLIDE 2 — EXECUTIVE SUMMARY (4 highlight boxes)
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[2]' --prop transition=morph
S='/slide[2]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_main' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_bar' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=28800

officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=420000 --prop width=1200000 --prop height=110000
officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag_txt' --prop text='EXECUTIVE SUMMARY' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=420000 --prop width=1200000 --prop height=110000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_title' --prop text='Q3 Performance Highlights' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=620000 --prop width=$FULL_W --prop height=640000

# 4 highlight metric boxes
BOX_W=2600000; BOX_H=1800000; BOX_Y=1600000; BOX_GAP=192000
HB_VALS=('$12.4M' '+8.3%' '94%' '3 New')
HB_LABELS=('Total Revenue' 'YoY Growth' 'Customer Retention' 'Product Lines Launched')
HB_COLORS=($ACCENT 3A8A5A 4A6AA8 8A4AB8)

for i in 0 1 2 3; do
  BX=$((M + i*(BOX_W + BOX_GAP)))
  officecli add "$OUT" "$S" --type shape \
    --prop name="hb$((i+1))_bg" --prop preset=rect --prop fill=$BG2 --prop line=$DARK_LINE --prop lineWidth=1 \
    --prop x=$BX --prop y=$BOX_Y --prop width=$BOX_W --prop height=$BOX_H
  officecli add "$OUT" "$S" --type shape \
    --prop name="hb$((i+1))_top" --prop preset=rect --prop fill="${HB_COLORS[$i]}" --prop line=none \
    --prop x=$BX --prop y=$BOX_Y --prop width=$BOX_W --prop height=16000
  officecli add "$OUT" "$S" --type shape \
    --prop name="hb$((i+1))_val" --prop text="${HB_VALS[$i]}" --prop font=Georgia --prop size=40 \
    --prop bold=false --prop color="${HB_COLORS[$i]}" --prop fill=none --prop line=none \
    --prop x=$((BX + 200000)) --prop y=$((BOX_Y + 300000)) --prop width=$((BOX_W - 400000)) --prop height=700000
  officecli add "$OUT" "$S" --type shape \
    --prop name="hb$((i+1))_label" --prop text="${HB_LABELS[$i]}" --prop font=Calibri --prop size=15 \
    --prop bold=false --prop color=$GRAY --prop fill=none --prop line=none \
    --prop x=$((BX + 200000)) --prop y=$((BOX_Y + 1100000)) --prop width=$((BOX_W - 400000)) --prop height=340000
done

# Summary paragraph
officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_sub' \
  --prop text='Q3 2024 delivered strong results across all key performance indicators, with revenue growth outpacing industry benchmarks and operational efficiency improvements across all business units.' \
  --prop font=Calibri --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=3700000 --prop width=$FULL_W --prop height=400000

# 3 key highlights list
HIGHS=('Revenue exceeded quarterly target by $1.2M (+10.7%)' 'Operating margin improved to 23.4% from 21.1% in Q2' 'Customer acquisition cost reduced by 18% through digital campaigns')
for i in 0 1 2; do
  HY=$((4280000 + i*640000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="hl_dot$((i+1))" --prop preset=ellipse --prop fill=$ACCENT --prop line=none \
    --prop x=$M --prop y=$((HY + 120000)) --prop width=220000 --prop height=220000
  officecli add "$OUT" "$S" --type shape \
    --prop name="hl_txt$((i+1))" --prop text="${HIGHS[$i]}" --prop font=Calibri --prop size=15 \
    --prop color=$WHITE --prop fill=none --prop line=none \
    --prop x=$((M + 380000)) --prop y=$HY --prop width=$((FULL_W - 380000)) --prop height=380000
done

echo "  ✓ Slide 2: Executive Summary"

# ────────────────────────────────────────────────────────
# SLIDE 3 — REVENUE ANALYSIS (line/bar chart)
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[3]' --prop transition=morph
S='/slide[3]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_main' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_bar' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=28800
officecli add "$OUT" "$S" --type shape \
  --prop name='!!img_panel' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=$((W/2)) --prop y=0 --prop width=$((W/2)) --prop height=$H

officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=420000 --prop width=1100000 --prop height=110000
officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag_txt' --prop text='REVENUE ANALYSIS' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=420000 --prop width=1100000 --prop height=110000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_title' --prop text='Revenue Performance' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=620000 --prop width=$FULL_W --prop height=640000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_sub' \
  --prop text='Quarterly revenue trends across product lines and geographic regions.' \
  --prop font=Calibri --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=1320000 --prop width=$FULL_W --prop height=340000

officecli add "$OUT" "$S" --type chart \
  --prop chartType=line \
  --prop title='Quarterly Revenue 2023-2024 ($M)' \
  --prop categories='Q1 2023,Q2 2023,Q3 2023,Q4 2023,Q1 2024,Q2 2024,Q3 2024' \
  --prop series1='8.2,9.1,10.3,11.4,10.8,11.5,12.4' \
  --prop series2='7.5,8.8,9.5,10.2,10.0,10.8,11.6' \
  --prop colors="$ACCENT,3A6AA8" \
  --prop x=$M --prop y=1900000 --prop width=$FULL_W --prop height=4400000 \
  --prop legend=true --prop dataLabels=false

officecli add "$OUT" "$S" --type shape \
  --prop name='chart_note' \
  --prop text='Series 1: Actual Revenue  |  Series 2: Target Revenue' \
  --prop font=Calibri --prop size=13 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=6400000 --prop width=$FULL_W --prop height=260000

echo "  ✓ Slide 3: Revenue Analysis"

# ────────────────────────────────────────────────────────
# SLIDE 4 — DEPARTMENT SCORECARD (table-style layout)
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[4]' --prop transition=morph
S='/slide[4]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_main' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_bar' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=28800
officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_line' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=28800 --prop height=$H

officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=420000 --prop width=1300000 --prop height=110000
officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag_txt' --prop text='DEPARTMENT SCORECARD' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=420000 --prop width=1300000 --prop height=110000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_title' --prop text='Performance by Department' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=620000 --prop width=$FULL_W --prop height=640000

# Table header
ROW_H=520000; TABLE_Y=1600000; COL_WIDTHS=(3200000 2000000 2000000 2000000 2000000)
COL_XS=($M $((M+3200000)) $((M+5200000)) $((M+7200000)) $((M+9200000)))
HDR_LABELS=('Department' 'Budget' 'Actual' 'Variance' 'Status')

for i in 0 1 2 3 4; do
  officecli add "$OUT" "$S" --type shape \
    --prop name="hdr_bg$((i+1))" --prop preset=rect --prop fill=$ACCENT --prop line=none \
    --prop x="${COL_XS[$i]}" --prop y=$TABLE_Y --prop width="${COL_WIDTHS[$i]}" --prop height=$ROW_H
  officecli add "$OUT" "$S" --type shape \
    --prop name="hdr_txt$((i+1))" --prop text="${HDR_LABELS[$i]}" --prop font=Calibri --prop size=15 \
    --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
    --prop x=$((${COL_XS[$i]} + 120000)) --prop y=$((TABLE_Y + 140000)) --prop width="${COL_WIDTHS[$i]}" --prop height=300000
done

# Data rows
DEPTS=('Sales' 'Marketing' 'Operations' 'R&D' 'HR')
BUDGETS=('$4,200K' '$1,500K' '$2,800K' '$2,000K' '$800K')
ACTUALS=('$4,650K' '$1,720K' '$2,650K' '$1,980K' '$780K')
VARIANCES=('+$450K' '-$220K' '+$150K' '+$20K' '+$20K')
STATUSES=('✓ On Track' '⚠ Over' '✓ On Track' '✓ On Track' '✓ On Track')
STATUS_COLORS=(3A8A5A C85A3A 3A8A5A 3A7A9A 3A8A5A)

for r in 0 1 2 3 4; do
  RY=$((TABLE_Y + ROW_H + r*(ROW_H + 14400)))
  BG_ROW=$([ $((r % 2)) -eq 0 ] && echo $BG2 || echo $BG)
  officecli add "$OUT" "$S" --type shape \
    --prop name="row${r}_bg" --prop preset=rect --prop fill=$BG_ROW --prop line=$DARK_LINE --prop lineWidth=1 \
    --prop x=$M --prop y=$RY --prop width=$FULL_W --prop height=$ROW_H
  officecli add "$OUT" "$S" --type shape \
    --prop name="row${r}_dept" --prop text="${DEPTS[$r]}" --prop font=Calibri --prop size=16 \
    --prop color=$WHITE --prop fill=none --prop line=none \
    --prop x=$((M + 200000)) --prop y=$((RY + 140000)) --prop width=2800000 --prop height=300000
  officecli add "$OUT" "$S" --type shape \
    --prop name="row${r}_bud" --prop text="${BUDGETS[$r]}" --prop font=Calibri --prop size=15 \
    --prop color=$GRAY --prop fill=none --prop line=none \
    --prop x=$((M + 3200000 + 120000)) --prop y=$((RY + 140000)) --prop width=1800000 --prop height=300000
  officecli add "$OUT" "$S" --type shape \
    --prop name="row${r}_act" --prop text="${ACTUALS[$r]}" --prop font=Calibri --prop size=15 \
    --prop color=$WHITE --prop fill=none --prop line=none \
    --prop x=$((M + 5200000 + 120000)) --prop y=$((RY + 140000)) --prop width=1800000 --prop height=300000
  officecli add "$OUT" "$S" --type shape \
    --prop name="row${r}_var" --prop text="${VARIANCES[$r]}" --prop font=Calibri --prop size=15 \
    --prop color=$ACCENT2 --prop fill=none --prop line=none \
    --prop x=$((M + 7200000 + 120000)) --prop y=$((RY + 140000)) --prop width=1800000 --prop height=300000
  officecli add "$OUT" "$S" --type shape \
    --prop name="row${r}_sts" --prop text="${STATUSES[$r]}" --prop font=Calibri --prop size=14 \
    --prop bold=true --prop color="${STATUS_COLORS[$r]}" --prop fill=none --prop line=none \
    --prop x=$((M + 9200000 + 120000)) --prop y=$((RY + 140000)) --prop width=1800000 --prop height=300000
done

echo "  ✓ Slide 4: Department Scorecard"

# ────────────────────────────────────────────────────────
# SLIDE 5 — OUTLOOK & NEXT STEPS
officecli add "$OUT" '/' --type slide --prop background=$BG
officecli set "$OUT" '/slide[5]' --prop transition=morph
S='/slide[5]'

officecli add "$OUT" "$S" --type shape \
  --prop name='!!bg_main' --prop preset=rect --prop fill=$BG --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=$H
officecli add "$OUT" "$S" --type shape \
  --prop name='!!top_bar' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=$W --prop height=28800
officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_line' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=0 --prop y=0 --prop width=28800 --prop height=$H

officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag' --prop preset=rect --prop fill=$ACCENT --prop line=none \
  --prop x=$M --prop y=420000 --prop width=700000 --prop height=110000
officecli add "$OUT" "$S" --type shape \
  --prop name='sec_tag_txt' --prop text='Q4 OUTLOOK' --prop font=Calibri --prop size=9 \
  --prop bold=true --prop color=$BG --prop fill=none --prop line=none \
  --prop x=$M --prop y=420000 --prop width=700000 --prop height=110000

officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_title' --prop text='Q4 Priorities & Targets' --prop font=Georgia --prop size=42 \
  --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
  --prop x=$M --prop y=620000 --prop width=$FULL_W --prop height=640000
officecli add "$OUT" "$S" --type shape \
  --prop name='!!cover_sub' \
  --prop text='Strategic focus areas and measurable targets for Q4 2024 to close the year strong.' \
  --prop font=Calibri --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none \
  --prop x=$M --prop y=1360000 --prop width=$FULL_W --prop height=340000

# Left col: 3 targets
TARGETS=('Revenue Target: $13.5M' 'New Customer Acquisitions: 250+' 'Operating Margin: ≥24%')
TARGET_SUBS=('Achieve $13.5M quarterly revenue through expanded enterprise sales pipeline.' 'Drive 250+ net new customer acquisitions via Q4 campaign blitz.' 'Maintain operating margin at or above 24% through cost discipline.')
for i in 0 1 2; do
  TY=$((1900000 + i*1500000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="tgt${i}_num" --prop text="0$((i+1))" --prop font=Georgia --prop size=36 \
    --prop bold=false --prop color=$ACCENT --prop fill=none --prop line=none \
    --prop x=$M --prop y=$TY --prop width=600000 --prop height=500000
  officecli add "$OUT" "$S" --type shape \
    --prop name="tgt${i}_title" --prop text="${TARGETS[$i]}" --prop font=Georgia --prop size=22 \
    --prop bold=false --prop color=$WHITE --prop fill=none --prop line=none \
    --prop x=$((M + 700000)) --prop y=$TY --prop width=5000000 --prop height=380000
  officecli add "$OUT" "$S" --type shape \
    --prop name="tgt${i}_text" --prop text="${TARGET_SUBS[$i]}" --prop font=Calibri --prop size=14 \
    --prop color=$GRAY --prop fill=none --prop line=none \
    --prop x=$((M + 700000)) --prop y=$((TY + 440000)) --prop width=5000000 --prop height=360000
  officecli add "$OUT" "$S" --type shape \
    --prop name="tgt${i}_line" --prop preset=rect --prop fill=$DARK_LINE --prop line=none \
    --prop x=$M --prop y=$((TY + 1340000)) --prop width=$((FULL_W/2)) --prop height=7200
done

# Right col: Q4 initiatives box
officecli add "$OUT" "$S" --type shape \
  --prop name='init_bg' --prop preset=rect --prop fill=$BG2 --prop line=$DARK_LINE --prop lineWidth=1 \
  --prop x=$((M + FULL_W/2 + 200000)) --prop y=1900000 --prop width=$((FULL_W/2 - 200000)) --prop height=4400000
officecli add "$OUT" "$S" --type shape \
  --prop name='init_title' --prop text='Key Initiatives' --prop font=Georgia --prop size=24 \
  --prop bold=false --prop color=$ACCENT --prop fill=none --prop line=none \
  --prop x=$((M + FULL_W/2 + 400000)) --prop y=2100000 --prop width=4800000 --prop height=400000

INITS=('Launch enterprise subscription tier' 'Expand APAC sales team to 12 reps' 'Deploy AI-powered customer support' 'Complete ISO 27001 certification')
for i in 0 1 2 3; do
  IY=$((2640000 + i*700000))
  officecli add "$OUT" "$S" --type shape \
    --prop name="init_dot$((i+1))" --prop preset=rect --prop fill=$ACCENT --prop line=none \
    --prop x=$((M + FULL_W/2 + 400000)) --prop y=$((IY + 180000)) --prop width=120000 --prop height=120000
  officecli add "$OUT" "$S" --type shape \
    --prop name="init_txt$((i+1))" --prop text="${INITS[$i]}" --prop font=Calibri --prop size=15 \
    --prop color=$WHITE --prop fill=none --prop line=none \
    --prop x=$((M + FULL_W/2 + 620000)) --prop y=$IY --prop width=4600000 --prop height=360000
done

echo "  ✓ Slide 5: Outlook"

officecli validate "$OUT" >/dev/null
echo ""
echo "✅ Generated: $OUT"
echo "   Slides: 5  |  Theme: Navy Corporate  |  Transition: Morph"
