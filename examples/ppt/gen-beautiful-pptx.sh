#!/bin/bash
# Generate a visually stunning presentation: "The Art of Design"
# Deep gradient backgrounds, geometric accents, clean typography
set -e

OUT="$(dirname "$0")/beautiful_presentation.pptx"
rm -f "$OUT"
officecli create "$OUT"
officecli open "$OUT"

# Slide dimensions: 12192000 x 6858000 EMU (16:9)

###############################################################################
# SLIDE 1 — Title Slide
###############################################################################
echo "  -> Slide 1: Title"
officecli add "$OUT" /presentation --type slide

# Full-bleed dark gradient background
officecli raw-set "$OUT" /slide[1] --xpath "//p:cSld" --action prepend --xml '
<p:bg>
  <p:bgPr>
    <a:gradFill rotWithShape="0">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="0D1B2A"/></a:gs>
        <a:gs pos="50000"><a:srgbClr val="1B2838"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="0A1628"/></a:gs>
      </a:gsLst>
      <a:lin ang="5400000" scaled="1"/>
    </a:gradFill>
    <a:effectLst/>
  </p:bgPr>
</p:bg>'

# Decorative circle — top right (large, semi-transparent teal)
officecli raw-set "$OUT" /slide[1] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="100" name="Deco Circle 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="8500000" y="-1200000"/><a:ext cx="4800000" cy="4800000"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="00B4D8"><a:alpha val="8000"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

# Decorative circle — bottom left (lavender)
officecli raw-set "$OUT" /slide[1] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="101" name="Deco Circle 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="-800000" y="4500000"/><a:ext cx="3200000" cy="3200000"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="E0AAFF"><a:alpha val="6000"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

# Gradient accent line
officecli raw-set "$OUT" /slide[1] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="102" name="Accent Line"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="4200000"/><a:ext cx="5000000" cy="0"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="28575">
      <a:gradFill>
        <a:gsLst>
          <a:gs pos="0"><a:srgbClr val="00B4D8"/></a:gs>
          <a:gs pos="100000"><a:srgbClr val="E0AAFF"/></a:gs>
        </a:gsLst>
        <a:lin ang="0" scaled="1"/>
      </a:gradFill>
    </a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

# Main title
officecli raw-set "$OUT" /slide[1] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="103" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="1600000"/><a:ext cx="8000000" cy="1200000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="b"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="l"/>
      <a:r>
        <a:rPr lang="en-US" sz="5400" b="1" dirty="0">
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:latin typeface="Segoe UI"/>
        </a:rPr>
        <a:t>The Art of Design</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'

# Subtitle
officecli raw-set "$OUT" /slide[1] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="104" name="Subtitle"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="2900000"/><a:ext cx="8000000" cy="1100000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="l"/>
      <a:r>
        <a:rPr lang="en-US" sz="2000" dirty="0">
          <a:solidFill><a:srgbClr val="90E0EF"/></a:solidFill>
          <a:latin typeface="Segoe UI"/>
        </a:rPr>
        <a:t>Crafting Beautiful Experiences</a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr algn="l"/>
      <a:r>
        <a:rPr lang="en-US" sz="1400" dirty="0" spc="600">
          <a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill>
          <a:latin typeface="Segoe UI"/>
        </a:rPr>
        <a:t>SIMPLICITY  &#xB7;  ELEGANCE  &#xB7;  FUNCTION</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'

# Diamond accent
officecli raw-set "$OUT" /slide[1] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="105" name="Diamond"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm rot="2700000"><a:off x="600000" y="4050000"/><a:ext cx="200000" cy="200000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="00B4D8"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

###############################################################################
# SLIDE 2 — Three Pillars
###############################################################################
echo "  -> Slide 2: Three Pillars"
officecli add "$OUT" /presentation --type slide

officecli raw-set "$OUT" /slide[2] --xpath "//p:cSld" --action prepend --xml '
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0D1B2A"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'

# Section title
officecli raw-set "$OUT" /slide[2] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="200" name="Section Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="400000"/><a:ext cx="10592000" cy="900000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="en-US" sz="3200" b="1" dirty="0">
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:latin typeface="Segoe UI"/>
        </a:rPr>
        <a:t>Three Pillars of Great Design</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'

# Subtitle
officecli raw-set "$OUT" /slide[2] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="201" name="SubLine"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="1200000"/><a:ext cx="10592000" cy="400000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="en-US" sz="1400" dirty="0">
          <a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill>
          <a:latin typeface="Segoe UI"/>
        </a:rPr>
        <a:t>Every exceptional design is built upon these core principles</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'

# Card 1 — Simplicity
officecli raw-set "$OUT" /slide[2] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="210" name="Card1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="900000" y="2000000"/><a:ext cx="3200000" cy="4200000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="152238"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="1E3A5F"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="228600" tIns="228600" rIns="228600" bIns="228600" anchor="t"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4800" dirty="0"><a:solidFill><a:srgbClr val="00B4D8"/></a:solidFill></a:rPr><a:t>&#x25CB;</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" sz="800"/></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Simplicity</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" sz="600"/></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Less is more. Strip away the unnecessary to let the essential shine through.</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Card 2 — Hierarchy
officecli raw-set "$OUT" /slide[2] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="211" name="Card2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="4496000" y="2000000"/><a:ext cx="3200000" cy="4200000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="152238"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="1E3A5F"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="228600" tIns="228600" rIns="228600" bIns="228600" anchor="t"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4800" dirty="0"><a:solidFill><a:srgbClr val="E0AAFF"/></a:solidFill></a:rPr><a:t>&#x25B3;</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" sz="800"/></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Hierarchy</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" sz="600"/></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Guide the eye with size, color, and space. Create a clear visual flow.</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Card 3 — Harmony
officecli raw-set "$OUT" /slide[2] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="212" name="Card3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="8092000" y="2000000"/><a:ext cx="3200000" cy="4200000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="152238"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="1E3A5F"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="228600" tIns="228600" rIns="228600" bIns="228600" anchor="t"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4800" dirty="0"><a:solidFill><a:srgbClr val="FFD166"/></a:solidFill></a:rPr><a:t>&#x25C7;</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" sz="800"/></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Harmony</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" sz="600"/></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Consistent color, type, and layout create a professional, cohesive experience.</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

###############################################################################
# SLIDE 3 — Data Showcase
###############################################################################
echo "  -> Slide 3: Data Showcase"
officecli add "$OUT" /presentation --type slide

officecli raw-set "$OUT" /slide[3] --xpath "//p:cSld" --action prepend --xml '
<p:bg><p:bgPr><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="0D1B2A"/></a:gs><a:gs pos="100000"><a:srgbClr val="152238"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>'

# Title
officecli raw-set "$OUT" /slide[3] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="300" name="DataTitle"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="300000"/><a:ext cx="10592000" cy="700000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="l"/><a:r><a:rPr lang="en-US" sz="2800" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Data-Driven Design</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Gradient accent bar
officecli raw-set "$OUT" /slide[3] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="301" name="Bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="1050000"/><a:ext cx="3000000" cy="50000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00B4D8"/></a:gs><a:gs pos="100000"><a:srgbClr val="E0AAFF"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

# Stat card 1 — 98%
officecli raw-set "$OUT" /slide[3] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="310" name="Stat1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="1500000"/><a:ext cx="3400000" cy="2200000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 6000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="0E2540"/></a:solidFill>
    <a:ln w="19050"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00B4D8"/></a:gs><a:gs pos="100000"><a:srgbClr val="0077B6"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="228600" tIns="182880" rIns="228600" bIns="182880" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="5600" b="1" dirty="0"><a:solidFill><a:srgbClr val="00B4D8"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>98%</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>User Satisfaction</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Stat card 2 — 2.5M
officecli raw-set "$OUT" /slide[3] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="311" name="Stat2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="4500000" y="1500000"/><a:ext cx="3400000" cy="2200000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 6000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="0E2540"/></a:solidFill>
    <a:ln w="19050"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="E0AAFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="9B5DE5"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="228600" tIns="182880" rIns="228600" bIns="182880" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="5600" b="1" dirty="0"><a:solidFill><a:srgbClr val="E0AAFF"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>2.5M</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Monthly Active Users</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Stat card 3 — 47ms
officecli raw-set "$OUT" /slide[3] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="312" name="Stat3"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="8200000" y="1500000"/><a:ext cx="3400000" cy="2200000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 6000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="0E2540"/></a:solidFill>
    <a:ln w="19050"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFD166"/></a:gs><a:gs pos="100000"><a:srgbClr val="F48C06"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="228600" tIns="182880" rIns="228600" bIns="182880" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="5600" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFD166"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>47ms</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Avg Response Time</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Bottom description
officecli raw-set "$OUT" /slide[3] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="320" name="Desc"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="4200000"/><a:ext cx="10592000" cy="2200000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>
    <a:p><a:pPr algn="l"/><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Numbers tell stories. Through thoughtful visual design, every data point</a:t></a:r></a:p>
    <a:p><a:pPr algn="l"/><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>communicates its meaning at first glance.</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

###############################################################################
# SLIDE 4 — Quote Slide
###############################################################################
echo "  -> Slide 4: Quote"
officecli add "$OUT" /presentation --type slide

officecli raw-set "$OUT" /slide[4] --xpath "//p:cSld" --action prepend --xml '
<p:bg><p:bgPr><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="1B2838"/></a:gs><a:gs pos="50000"><a:srgbClr val="0D1B2A"/></a:gs><a:gs pos="100000"><a:srgbClr val="1B2838"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>'

# Large quote mark
officecli raw-set "$OUT" /slide[4] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="400" name="QuoteMark"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1000000" y="800000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>
    <a:p><a:pPr algn="l"/><a:r><a:rPr lang="en-US" sz="12000" dirty="0"><a:solidFill><a:srgbClr val="00B4D8"><a:alpha val="20000"/></a:srgbClr></a:solidFill><a:latin typeface="Georgia"/></a:rPr><a:t>&#x201C;</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Quote text
officecli raw-set "$OUT" /slide[4] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="401" name="Quote"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1500000" y="2000000"/><a:ext cx="9192000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2800" i="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Georgia"/></a:rPr><a:t>Good design is obvious.</a:t></a:r></a:p>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2800" i="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Georgia"/></a:rPr><a:t>Great design is transparent.</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Attribution
officecli raw-set "$OUT" /slide[4] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="402" name="Author"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1500000" y="4200000"/><a:ext cx="9192000" cy="600000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1600" dirty="0"><a:solidFill><a:srgbClr val="00B4D8"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>&#x2014; Joe Sparano</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Decorative line under quote
officecli raw-set "$OUT" /slide[4] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="403" name="QuoteLine"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="5096000" y="5000000"/><a:ext cx="2000000" cy="0"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="19050"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00B4D8"><a:alpha val="0"/></a:srgbClr></a:gs><a:gs pos="50000"><a:srgbClr val="00B4D8"/></a:gs><a:gs pos="100000"><a:srgbClr val="00B4D8"><a:alpha val="0"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

###############################################################################
# SLIDE 5 — Process / Timeline
###############################################################################
echo "  -> Slide 5: Process"
officecli add "$OUT" /presentation --type slide

officecli raw-set "$OUT" /slide[5] --xpath "//p:cSld" --action prepend --xml '
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0D1B2A"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'

# Title
officecli raw-set "$OUT" /slide[5] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="500" name="ProcessTitle"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="300000"/><a:ext cx="10592000" cy="900000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="3200" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Design Process</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Horizontal rainbow connector
officecli raw-set "$OUT" /slide[5] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="501" name="ConnLine"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1800000" y="2800000"/><a:ext cx="8600000" cy="0"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="25400"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00B4D8"/></a:gs><a:gs pos="33000"><a:srgbClr val="E0AAFF"/></a:gs><a:gs pos="66000"><a:srgbClr val="FFD166"/></a:gs><a:gs pos="100000"><a:srgbClr val="06D6A0"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

# Step circles + labels
LABELS=("Research" "Ideate" "Design" "Validate")
COLORS=("00B4D8" "E0AAFF" "FFD166" "06D6A0")
XPOS=(1400000 3600000 5800000 8000000)

for i in 0 1 2 3; do
  X=${XPOS[$i]}
  C=${COLORS[$i]}
  L=${LABELS[$i]}
  N=$((i+1))
  ID=$((510 + i*2))
  ID2=$((511 + i*2))

  officecli raw-set "$OUT" /slide[5] --xpath "//p:cSld/p:spTree" --action append --xml "
<p:sp>
  <p:nvSpPr><p:cNvPr id=\"${ID}\" name=\"Step${N}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x=\"${X}\" y=\"2200000\"/><a:ext cx=\"1200000\" cy=\"1200000\"/></a:xfrm>
    <a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val=\"${C}\"><a:alpha val=\"15000\"/></a:srgbClr></a:solidFill>
    <a:ln w=\"38100\"><a:solidFill><a:srgbClr val=\"${C}\"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap=\"square\" anchor=\"ctr\"/><a:lstStyle/>
    <a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr lang=\"en-US\" sz=\"2400\" b=\"1\" dirty=\"0\"><a:solidFill><a:srgbClr val=\"${C}\"/></a:solidFill></a:rPr><a:t>0${N}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"

  officecli raw-set "$OUT" /slide[5] --xpath "//p:cSld/p:spTree" --action append --xml "
<p:sp>
  <p:nvSpPr><p:cNvPr id=\"${ID2}\" name=\"Label${N}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x=\"${X}\" y=\"3600000\"/><a:ext cx=\"1200000\" cy=\"800000\"/></a:xfrm>
    <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap=\"square\" anchor=\"t\"/><a:lstStyle/>
    <a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr lang=\"en-US\" sz=\"1800\" b=\"1\" dirty=\"0\"><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill><a:latin typeface=\"Segoe UI\"/></a:rPr><a:t>${L}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"
done

# Bottom text
officecli raw-set "$OUT" /slide[5] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="530" name="Bottom"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="800000" y="5000000"/><a:ext cx="10592000" cy="600000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="8B95A2"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Every step is iterative. From research to validation, we refine until perfection.</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

###############################################################################
# SLIDE 6 — Closing
###############################################################################
echo "  -> Slide 6: Closing"
officecli add "$OUT" /presentation --type slide

officecli raw-set "$OUT" /slide[6] --xpath "//p:cSld" --action prepend --xml '
<p:bg><p:bgPr><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="0A1628"/></a:gs><a:gs pos="50000"><a:srgbClr val="0D1B2A"/></a:gs><a:gs pos="100000"><a:srgbClr val="1B2838"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>'

# Gradient ring
officecli raw-set "$OUT" /slide[6] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="600" name="Ring"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="3596000" y="800000"/><a:ext cx="5000000" cy="5000000"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln w="12700"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00B4D8"><a:alpha val="30000"/></a:srgbClr></a:gs><a:gs pos="50000"><a:srgbClr val="E0AAFF"><a:alpha val="30000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="FFD166"><a:alpha val="30000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

# Thank You
officecli raw-set "$OUT" /slide[6] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="601" name="Thanks"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1500000" y="2200000"/><a:ext cx="9192000" cy="1400000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4800" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Thank You</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Closing subtitle
officecli raw-set "$OUT" /slide[6] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="602" name="ClosingSub"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1500000" y="3600000"/><a:ext cx="9192000" cy="800000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1600" dirty="0"><a:solidFill><a:srgbClr val="90E0EF"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>Design is not just what it looks like &#x2014; it&#x2019;s how it works.</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'

# Three accent diamonds
officecli raw-set "$OUT" /slide[6] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="603" name="D1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm rot="2700000"><a:off x="5850000" y="4700000"/><a:ext cx="120000" cy="120000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="00B4D8"/></a:solidFill><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

officecli raw-set "$OUT" /slide[6] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="604" name="D2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm rot="2700000"><a:off x="6100000" y="4700000"/><a:ext cx="120000" cy="120000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="E0AAFF"/></a:solidFill><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

officecli raw-set "$OUT" /slide[6] --xpath "//p:cSld/p:spTree" --action append --xml '
<p:sp>
  <p:nvSpPr><p:cNvPr id="605" name="D3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm rot="2700000"><a:off x="6350000" y="4700000"/><a:ext cx="120000" cy="120000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FFD166"/></a:solidFill><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>'

officecli close "$OUT"

echo ""
echo "=========================================="
echo "Beautiful presentation generated: $OUT"
echo "=========================================="
officecli view "$OUT" outline
