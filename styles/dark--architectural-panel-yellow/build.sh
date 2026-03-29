#!/bin/bash
set +H  # disable !! history expansion
set -e

F="/Users/visher/Desktop/architectural_business_plan.pptx"

# Design tokens (hex no #)
DARK="1C2B3A"
PANEL="B8D4E0"
IMG_BG="7FABBF"
YELLOW="F4C430"
GRAY="666666"
LGRAY="9A9A9A"
WHITE="FFFFFF"
CARD="EFF6FA"
CARD_LINE="D2E8F2"

# Slide size: 33.87 x 19.05 cm (widescreen 16:9)

add_shape() {
  officecli add "$F" "$1" --type shape "${@:2}"
}
add_conn() {
  officecli add "$F" "$1" --type connector "${@:2}"
}
add_slide() {
  officecli add "$F" / --type slide "${@}"
}

echo "Creating $F..."
rm -f "$F"
officecli create "$F"

# ============================================================
# SLIDE 1 — TITLE SLIDE  (panel RIGHT)
# ============================================================
echo "  Slide 1: Title..."
add_slide --prop background=$WHITE

# Right blue panel
add_shape '/slide[1]' \
  --prop 'name=!!bg-panel' --prop preset=rect \
  --prop x=18cm --prop y=0cm --prop width=15.9cm --prop height=19.1cm \
  --prop fill=$PANEL --prop line=none

# Image placeholder
add_shape '/slide[1]' \
  --prop 'name=!!hero-img' --prop text="[ Architecture Image ]" \
  --prop x=18.6cm --prop y=0.6cm --prop width=14.8cm --prop height=17.8cm \
  --prop fill=$IMG_BG --prop line=none --prop color=$WHITE \
  --prop size=14 --prop align=center --prop valign=center

# "Your Project" label top-left
add_shape '/slide[1]' \
  --prop 'name=!!top-label' --prop text="Your Project" \
  --prop x=1.3cm --prop y=0.65cm --prop width=6cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

# Year tag
add_shape '/slide[1]' \
  --prop 'name=!!year-tag' --prop text="— 2025" \
  --prop x=10.5cm --prop y=0.65cm --prop width=3.5cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

# "Business Plan" right label
add_shape '/slide[1]' \
  --prop 'name=!!biz-label' --prop text="Business Plan" \
  --prop x=13.5cm --prop y=0.65cm --prop width=4cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none --prop align=right

# Yellow decorative star
add_shape '/slide[1]' \
  --prop 'name=!!deco-star' --prop text="✦" \
  --prop x=1.3cm --prop y=3.8cm --prop width=1.8cm --prop height=1.8cm \
  --prop size=30 --prop color=$YELLOW --prop fill=none --prop line=none

# Main title
add_shape '/slide[1]' \
  --prop text="Architectural\nBusiness Plan" \
  --prop x=1.3cm --prop y=4.5cm --prop width=16cm --prop height=6cm \
  --prop size=62 --prop bold=true --prop color=$DARK \
  --prop fill=none --prop line=none --prop lineSpacing=1.05

# Subtitle
add_shape '/slide[1]' \
  --prop text="Lorem ipsum dolor sit amet, consectetur adipiscing elit,\nsed do eiusmod tempor incididunt ut labore et dolore\nmagna aliqua. Ut enim ad minim veniam." \
  --prop x=1.3cm --prop y=11cm --prop width=16cm --prop height=2.7cm \
  --prop size=10.5 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.4

# Get Started button
add_shape '/slide[1]' \
  --prop 'name=!!cta-btn' --prop text="Get Started  →" \
  --prop x=1.3cm --prop y=14.2cm --prop width=5cm --prop height=1.3cm \
  --prop size=10.5 --prop bold=true --prop color=$WHITE \
  --prop fill=$DARK --prop line=none --prop preset=roundRect \
  --prop align=center --prop valign=center

# Stat divider
add_conn '/slide[1]' \
  --prop x=6cm --prop y=16.2cm --prop width=0cm --prop height=2.4cm \
  --prop line=CCCCCC --prop lineWidth=0.5pt

# Stats
add_shape '/slide[1]' \
  --prop 'name=!!stat1-num' --prop text="450+" \
  --prop x=1.3cm --prop y=16cm --prop width=4.3cm --prop height=1.2cm \
  --prop size=34 --prop bold=true --prop color=$DARK --prop fill=none --prop line=none

add_shape '/slide[1]' \
  --prop 'name=!!stat1-lbl' --prop text="Projects Completed" \
  --prop x=1.3cm --prop y=17.2cm --prop width=4.5cm --prop height=0.9cm \
  --prop size=8.5 --prop color=$GRAY --prop fill=none --prop line=none

add_shape '/slide[1]' \
  --prop 'name=!!stat2-num' --prop text="230+" \
  --prop x=6.7cm --prop y=16cm --prop width=4cm --prop height=1.2cm \
  --prop size=34 --prop bold=true --prop color=$DARK --prop fill=none --prop line=none

add_shape '/slide[1]' \
  --prop 'name=!!stat2-lbl' --prop text="Awards Won" \
  --prop x=6.7cm --prop y=17.2cm --prop width=4cm --prop height=0.9cm \
  --prop size=8.5 --prop color=$GRAY --prop fill=none --prop line=none


# ============================================================
# SLIDE 2 — OUR SPECIALIZED OFFERINGS (panel LEFT, morph)
# ============================================================
echo "  Slide 2: Our Specialized Offerings..."
add_slide --prop background=$WHITE --prop transition=morph

add_shape '/slide[2]' \
  --prop 'name=!!bg-panel' --prop preset=rect \
  --prop x=0cm --prop y=0cm --prop width=12cm --prop height=19.1cm \
  --prop fill=$PANEL --prop line=none

add_shape '/slide[2]' \
  --prop 'name=!!hero-img' --prop text="[ Architecture Image ]" \
  --prop x=0.5cm --prop y=0.5cm --prop width=11cm --prop height=18.1cm \
  --prop fill=$IMG_BG --prop line=none --prop color=$WHITE \
  --prop size=14 --prop align=center --prop valign=center

add_shape '/slide[2]' \
  --prop 'name=!!top-label' --prop text="Your Project" \
  --prop x=13cm --prop y=0.65cm --prop width=6cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

add_shape '/slide[2]' \
  --prop 'name=!!year-tag' --prop text="— 2025" \
  --prop x=25cm --prop y=0.65cm --prop width=3.5cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

add_shape '/slide[2]' \
  --prop 'name=!!biz-label' --prop text="Business Plan" \
  --prop x=28.5cm --prop y=0.65cm --prop width=5cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none --prop align=right

add_shape '/slide[2]' \
  --prop 'name=!!deco-star' --prop text="✦" \
  --prop x=13cm --prop y=2.7cm --prop width=1.6cm --prop height=1.6cm \
  --prop size=24 --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[2]' \
  --prop text="Our Specialized\nOfferings" \
  --prop x=14.8cm --prop y=2.4cm --prop width=18.5cm --prop height=4.5cm \
  --prop size=50 --prop bold=true --prop color=$DARK \
  --prop fill=none --prop line=none

add_shape '/slide[2]' \
  --prop text="We bring architectural vision to life through innovative design,\nprecision engineering and sustainable solutions for every client." \
  --prop x=13cm --prop y=7.5cm --prop width=20.5cm --prop height=2cm \
  --prop size=10.5 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.4

# 3 service cards
add_shape '/slide[2]' \
  --prop text="01\nResidential\nDesign" \
  --prop x=13cm --prop y=10cm --prop width=5.5cm --prop height=3.4cm \
  --prop size=11 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[2]' \
  --prop text="02\nCommercial\nProjects" \
  --prop x=19.2cm --prop y=10cm --prop width=5.5cm --prop height=3.4cm \
  --prop size=11 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[2]' \
  --prop text="03\nUrban\nPlanning" \
  --prop x=25.4cm --prop y=10cm --prop width=5.5cm --prop height=3.4cm \
  --prop size=11 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[2]' \
  --prop 'name=!!stat1-num' --prop text="450+" \
  --prop x=13cm --prop y=15.4cm --prop width=4.3cm --prop height=1.2cm \
  --prop size=34 --prop bold=true --prop color=$DARK --prop fill=none --prop line=none

add_shape '/slide[2]' \
  --prop 'name=!!stat1-lbl' --prop text="Projects Completed" \
  --prop x=13cm --prop y=16.6cm --prop width=5cm --prop height=0.9cm \
  --prop size=8.5 --prop color=$GRAY --prop fill=none --prop line=none

add_shape '/slide[2]' \
  --prop 'name=!!stat2-num' --prop text="230+" \
  --prop x=18.7cm --prop y=15.4cm --prop width=4cm --prop height=1.2cm \
  --prop size=34 --prop bold=true --prop color=$DARK --prop fill=none --prop line=none

add_shape '/slide[2]' \
  --prop 'name=!!stat2-lbl' --prop text="Awards Won" \
  --prop x=18.7cm --prop y=16.6cm --prop width=4cm --prop height=0.9cm \
  --prop size=8.5 --prop color=$GRAY --prop fill=none --prop line=none

add_shape '/slide[2]' \
  --prop 'name=!!cta-btn' --prop text="Learn More  →" \
  --prop x=13cm --prop y=17.7cm --prop width=5cm --prop height=1.1cm \
  --prop size=10.5 --prop bold=true --prop color=$WHITE \
  --prop fill=$DARK --prop line=none --prop preset=roundRect \
  --prop align=center --prop valign=center


# ============================================================
# SLIDE 3 — VISION & MISSION STATEMENT (panel RIGHT, morph)
# ============================================================
echo "  Slide 3: Vision & Mission..."
add_slide --prop background=$WHITE --prop transition=morph

add_shape '/slide[3]' \
  --prop 'name=!!bg-panel' --prop preset=rect \
  --prop x=21.9cm --prop y=0cm --prop width=12cm --prop height=19.1cm \
  --prop fill=$PANEL --prop line=none

add_shape '/slide[3]' \
  --prop 'name=!!hero-img' --prop text="[ Architecture Image ]" \
  --prop x=22.4cm --prop y=0.5cm --prop width=11cm --prop height=18.1cm \
  --prop fill=$IMG_BG --prop line=none --prop color=$WHITE \
  --prop size=14 --prop align=center --prop valign=center

add_shape '/slide[3]' \
  --prop 'name=!!top-label' --prop text="Your Project" \
  --prop x=1.3cm --prop y=0.65cm --prop width=6cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

add_shape '/slide[3]' \
  --prop 'name=!!year-tag' --prop text="— 2025" \
  --prop x=10.5cm --prop y=0.65cm --prop width=3.5cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

add_shape '/slide[3]' \
  --prop 'name=!!biz-label' --prop text="Business Plan" \
  --prop x=13.5cm --prop y=0.65cm --prop width=4cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none --prop align=right

add_shape '/slide[3]' \
  --prop 'name=!!deco-star' --prop text="✦" \
  --prop x=1.3cm --prop y=2.9cm --prop width=1.6cm --prop height=1.6cm \
  --prop size=24 --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[3]' \
  --prop text="Vision & Mission\nStatement" \
  --prop x=3.1cm --prop y=2.6cm --prop width=18cm --prop height=4.5cm \
  --prop size=50 --prop bold=true --prop color=$DARK \
  --prop fill=none --prop line=none

# Vision
add_shape '/slide[3]' \
  --prop text="Our Vision" \
  --prop x=1.3cm --prop y=8.3cm --prop width=8cm --prop height=1cm \
  --prop size=13 --prop bold=true --prop color=$DARK --prop fill=none --prop line=none

add_shape '/slide[3]' \
  --prop text="To be the leading architectural firm that transforms urban\nlandscapes through innovative, sustainable design that\ninspires communities and endures through generations." \
  --prop x=1.3cm --prop y=9.5cm --prop width=20cm --prop height=2.7cm \
  --prop size=10.5 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.4

# Mission
add_shape '/slide[3]' \
  --prop text="Our Mission" \
  --prop x=1.3cm --prop y=12.8cm --prop width=8cm --prop height=1cm \
  --prop size=13 --prop bold=true --prop color=$DARK --prop fill=none --prop line=none

add_shape '/slide[3]' \
  --prop text="To deliver exceptional architectural solutions that balance\naesthetics, functionality and sustainability, while building\nlasting relationships with clients and communities." \
  --prop x=1.3cm --prop y=14cm --prop width=20cm --prop height=2.7cm \
  --prop size=10.5 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.4

add_shape '/slide[3]' \
  --prop 'name=!!stat-pct' --prop text="25%" \
  --prop x=1.3cm --prop y=16.8cm --prop width=4cm --prop height=1.6cm \
  --prop size=40 --prop bold=true --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[3]' \
  --prop text="Annual growth\nin client base" \
  --prop x=5.7cm --prop y=17.1cm --prop width=6cm --prop height=1.4cm \
  --prop size=9 --prop color=$GRAY --prop fill=none --prop line=none


# ============================================================
# SLIDE 4 — FOUNDATIONS OF OUR BUSINESS (panel LEFT, morph)
# ============================================================
echo "  Slide 4: Foundations..."
add_slide --prop background=$WHITE --prop transition=morph

add_shape '/slide[4]' \
  --prop 'name=!!bg-panel' --prop preset=rect \
  --prop x=0cm --prop y=0cm --prop width=14.7cm --prop height=19.1cm \
  --prop fill=$PANEL --prop line=none

add_shape '/slide[4]' \
  --prop 'name=!!hero-img' --prop text="[ Architecture Image ]" \
  --prop x=0.5cm --prop y=0.5cm --prop width=13.7cm --prop height=18.1cm \
  --prop fill=$IMG_BG --prop line=none --prop color=$WHITE \
  --prop size=14 --prop align=center --prop valign=center

add_shape '/slide[4]' \
  --prop 'name=!!top-label' --prop text="Your Project" \
  --prop x=15.7cm --prop y=0.65cm --prop width=6cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

add_shape '/slide[4]' \
  --prop 'name=!!year-tag' --prop text="— 2025" \
  --prop x=26cm --prop y=0.65cm --prop width=3.5cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

add_shape '/slide[4]' \
  --prop 'name=!!biz-label' --prop text="Business Plan" \
  --prop x=29.5cm --prop y=0.65cm --prop width=4cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none --prop align=right

add_shape '/slide[4]' \
  --prop 'name=!!deco-star' --prop text="✦" \
  --prop x=15.7cm --prop y=2.9cm --prop width=1.6cm --prop height=1.6cm \
  --prop size=24 --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[4]' \
  --prop text="Foundations of\nOur Business" \
  --prop x=17.5cm --prop y=2.6cm --prop width=15.5cm --prop height=4.5cm \
  --prop size=50 --prop bold=true --prop color=$DARK \
  --prop fill=none --prop line=none

add_shape '/slide[4]' \
  --prop text="Our business is built on three core pillars that define\nour approach to every project we undertake." \
  --prop x=15.7cm --prop y=7.5cm --prop width=17.5cm --prop height=2cm \
  --prop size=10.5 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.4

# 3 pillars
add_shape '/slide[4]' \
  --prop text="Innovation\n\nWe constantly push boundaries of architectural design, embracing new technologies and materials." \
  --prop x=15.7cm --prop y=10cm --prop width=5.2cm --prop height=6cm \
  --prop size=10 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[4]' \
  --prop text="Sustainability\n\nEnvironmental responsibility guides every design decision we make for our clients." \
  --prop x=21.4cm --prop y=10cm --prop width=5.2cm --prop height=6cm \
  --prop size=10 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[4]' \
  --prop text="Excellence\n\nWe deliver projects that exceed expectations in quality, function and aesthetic beauty." \
  --prop x=27.1cm --prop y=10cm --prop width=5.2cm --prop height=6cm \
  --prop size=10 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[4]' \
  --prop 'name=!!stat-pct' --prop text="25%" \
  --prop x=15.7cm --prop y=16.8cm --prop width=4cm --prop height=1.6cm \
  --prop size=40 --prop bold=true --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[4]' \
  --prop text="Average ROI for\nclient investments" \
  --prop x=20.3cm --prop y=17.1cm --prop width=7cm --prop height=1.4cm \
  --prop size=9 --prop color=$GRAY --prop fill=none --prop line=none


# ============================================================
# SLIDE 5 — DETAILING THE BUSINESS (panel RIGHT, morph)
# ============================================================
echo "  Slide 5: Detailing the Business..."
add_slide --prop background=$WHITE --prop transition=morph

add_shape '/slide[5]' \
  --prop 'name=!!bg-panel' --prop preset=rect \
  --prop x=21.9cm --prop y=0cm --prop width=12cm --prop height=19.1cm \
  --prop fill=$PANEL --prop line=none

add_shape '/slide[5]' \
  --prop 'name=!!hero-img' --prop text="[ Architecture Image ]" \
  --prop x=22.4cm --prop y=0.5cm --prop width=11cm --prop height=18.1cm \
  --prop fill=$IMG_BG --prop line=none --prop color=$WHITE \
  --prop size=14 --prop align=center --prop valign=center

add_shape '/slide[5]' \
  --prop 'name=!!top-label' --prop text="Your Project" \
  --prop x=1.3cm --prop y=0.65cm --prop width=6cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

add_shape '/slide[5]' \
  --prop 'name=!!year-tag' --prop text="— 2025" \
  --prop x=10.5cm --prop y=0.65cm --prop width=3.5cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none

add_shape '/slide[5]' \
  --prop 'name=!!biz-label' --prop text="Business Plan" \
  --prop x=13.5cm --prop y=0.65cm --prop width=4cm --prop height=0.75cm \
  --prop size=9 --prop color=$LGRAY --prop fill=none --prop line=none --prop align=right

add_shape '/slide[5]' \
  --prop 'name=!!deco-star' --prop text="✦" \
  --prop x=1.3cm --prop y=2.9cm --prop width=1.6cm --prop height=1.6cm \
  --prop size=24 --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[5]' \
  --prop text="Detailing the\nBusiness" \
  --prop x=3.1cm --prop y=2.6cm --prop width=18cm --prop height=4.5cm \
  --prop size=50 --prop bold=true --prop color=$DARK \
  --prop fill=none --prop line=none

add_shape '/slide[5]' \
  --prop text="A comprehensive breakdown of our business model,\noperational strategy and financial projections." \
  --prop x=1.3cm --prop y=7.5cm --prop width=20cm --prop height=2cm \
  --prop size=10.5 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.4

# 3 detail cards
add_shape '/slide[5]' \
  --prop text="Revenue Model\n\n• Project-based fees\n• Retainer services\n• Consultation\n• IP Licensing" \
  --prop x=1.3cm --prop y=10cm --prop width=5.8cm --prop height=7.3cm \
  --prop size=10 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[5]' \
  --prop text="Market Strategy\n\n• Premium positioning\n• Digital marketing\n• Referral programs\n• Awards & PR" \
  --prop x=7.7cm --prop y=10cm --prop width=5.8cm --prop height=7.3cm \
  --prop size=10 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[5]' \
  --prop text="Growth Plan\n\n• 3 new markets\n• Team expansion\n• Tech investment\n• Global reach" \
  --prop x=14.1cm --prop y=10cm --prop width=5.8cm --prop height=7.3cm \
  --prop size=10 --prop color=$DARK --prop fill=$CARD --prop line=$CARD_LINE \
  --prop lineWidth=0.5pt --prop preset=roundRect --prop margin=0.4cm

add_shape '/slide[5]' \
  --prop 'name=!!stat-pct' --prop text="25%" \
  --prop x=1.3cm --prop y=16.8cm --prop width=4cm --prop height=1.6cm \
  --prop size=40 --prop bold=true --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[5]' \
  --prop text="Projected annual\nrevenue growth" \
  --prop x=5.7cm --prop y=17.1cm --prop width=6cm --prop height=1.4cm \
  --prop size=9 --prop color=$GRAY --prop fill=none --prop line=none


# ============================================================
# SLIDE 6 — CLOSING: DELVING DEEPER (dark full-bg, morph)
# ============================================================
echo "  Slide 6: Closing..."
add_slide --prop background=$DARK --prop transition=morph

# Full dark panel (morph from right panel)
add_shape '/slide[6]' \
  --prop 'name=!!bg-panel' --prop preset=rect \
  --prop x=0cm --prop y=0cm --prop width=33.9cm --prop height=19.1cm \
  --prop fill=$DARK --prop line=none

# Image overlay right half
add_shape '/slide[6]' \
  --prop 'name=!!hero-img' --prop text="[ Architecture Image ]" \
  --prop x=16.5cm --prop y=0cm --prop width=17.4cm --prop height=19.1cm \
  --prop fill=253A4D --prop line=none --prop color=3A6070 \
  --prop size=14 --prop align=center --prop valign=center

add_shape '/slide[6]' \
  --prop 'name=!!top-label' --prop text="Your Project" \
  --prop x=1.3cm --prop y=0.65cm --prop width=6cm --prop height=0.75cm \
  --prop size=9 --prop color=4A6A7A --prop fill=none --prop line=none

add_shape '/slide[6]' \
  --prop 'name=!!year-tag' --prop text="— 2025" \
  --prop x=10.5cm --prop y=0.65cm --prop width=3.5cm --prop height=0.75cm \
  --prop size=9 --prop color=4A6A7A --prop fill=none --prop line=none

add_shape '/slide[6]' \
  --prop 'name=!!biz-label' --prop text="Business Plan" \
  --prop x=13.5cm --prop y=0.65cm --prop width=4cm --prop height=0.75cm \
  --prop size=9 --prop color=4A6A7A --prop fill=none --prop line=none --prop align=right

add_shape '/slide[6]' \
  --prop 'name=!!deco-star' --prop text="✦" \
  --prop x=1.3cm --prop y=5.1cm --prop width=2cm --prop height=2cm \
  --prop size=34 --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[6]' \
  --prop text="Delving Deeper\ninto the\nFoundations" \
  --prop x=1.3cm --prop y=4.7cm --prop width=14.5cm --prop height=8cm \
  --prop size=56 --prop bold=true --prop color=$WHITE \
  --prop fill=none --prop line=none --prop lineSpacing=1.05

add_shape '/slide[6]' \
  --prop text="Explore the full scope of our architectural expertise,\nour proven track record and vision for the future." \
  --prop x=1.3cm --prop y=13.1cm --prop width=14cm --prop height=2cm \
  --prop size=11 --prop color=$PANEL --prop fill=none --prop line=none --prop lineSpacing=1.4

add_shape '/slide[6]' \
  --prop 'name=!!cta-btn' --prop text="View Full Plan  →" \
  --prop x=1.3cm --prop y=15.4cm --prop width=6cm --prop height=1.4cm \
  --prop size=10.5 --prop bold=true --prop color=$DARK \
  --prop fill=$YELLOW --prop line=none --prop preset=roundRect \
  --prop align=center --prop valign=center

add_shape '/slide[6]' \
  --prop 'name=!!stat-pct' --prop text="25%" \
  --prop x=1.3cm --prop y=16.8cm --prop width=4cm --prop height=1.6cm \
  --prop size=40 --prop bold=true --prop color=$YELLOW --prop fill=none --prop line=none

add_shape '/slide[6]' \
  --prop text="Growth Rate" \
  --prop x=5.7cm --prop y=17.1cm --prop width=5cm --prop height=1.4cm \
  --prop size=9 --prop color=$PANEL --prop fill=none --prop line=none

echo ""
echo "✓ Done! Saved to: $F"
