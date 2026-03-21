#!/bin/bash
# Generate a showcase document with beautiful charts
# Contains 8 chart types: combo chart, 3D bar, scatter+trendline, 3D pie, bubble, stock OHLC, filled radar, multi-ring doughnut
# 4 Sheets: monthly sales, analysis data, stock data, capability assessment

set -e

XLSX="beautiful_charts.xlsx"
echo ""
echo "=========================================="
echo "Generating beautiful charts document: $XLSX"
echo "=========================================="

rm -f "$XLSX"
officecli create "$XLSX"
officecli open "$XLSX"

###############################################################################
# Sheet1: Monthly sales data
###############################################################################
echo "  -> Populating Sheet1: Monthly sales data"

officecli set "$XLSX" '/Sheet1/A1' --prop value="Month" --prop font.bold=true --prop fill=1F4E79 --prop font.color=FFFFFF --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/B1' --prop value="East Sales" --prop font.bold=true --prop fill=2E75B6 --prop font.color=FFFFFF --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/C1' --prop value="South Sales" --prop font.bold=true --prop fill=9DC3E6 --prop font.color=1F4E79 --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/D1' --prop value="North Sales" --prop font.bold=true --prop fill=BDD7EE --prop font.color=1F4E79 --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/E1' --prop value="Total" --prop font.bold=true --prop fill=C55A11 --prop font.color=FFFFFF --prop font.size=11 --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/F1' --prop value="YoY Growth %" --prop font.bold=true --prop fill=548235 --prop font.color=FFFFFF --prop font.size=11 --prop alignment.horizontal=center

MONTHS=("Jan" "Feb" "Mar" "Apr" "May" "Jun" "Jul" "Aug" "Sep" "Oct" "Nov" "Dec")
EAST=(120 135 148 162 155 178 195 210 188 172 165 198)
SOUTH=(95 108 115 128 142 155 168 175 160 148 135 158)
NORTH=(88 92 105 118 125 138 145 152 140 130 122 142)
TOTAL=(303 335 368 408 422 471 508 537 488 450 422 498)
GROWTH=(5.2 8.1 12.3 15.6 10.2 18.5 22.1 25.3 16.8 11.2 7.5 19.8)

for i in $(seq 0 11); do
    row=$((i + 2))
    officecli set "$XLSX" "/Sheet1/A${row}" --prop "value=${MONTHS[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/B${row}" --prop "value=${EAST[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/C${row}" --prop "value=${SOUTH[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/D${row}" --prop "value=${NORTH[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/E${row}" --prop "value=${TOTAL[$i]}" --prop 'numFmt=#,##0' --prop font.bold=true --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/F${row}" --prop "value=${GROWTH[$i]}" --prop 'numFmt=0.0"%"' --prop alignment.horizontal=center
done

echo "  Done: Sheet1 data"

###############################################################################
# Sheet2: Scatter/bubble chart data
###############################################################################
echo "  -> Populating Sheet2: Analysis data"

officecli add "$XLSX" / --type sheet --prop name=Analysis

officecli set "$XLSX" '/Analysis/A1' --prop value="Ad Spend (10K)" --prop font.bold=true --prop fill=7030A0 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Analysis/B1' --prop value="Sales (10K)" --prop font.bold=true --prop fill=7030A0 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Analysis/C1' --prop value="Margin %" --prop font.bold=true --prop fill=7030A0 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Analysis/D1' --prop value="Market Share %" --prop font.bold=true --prop fill=7030A0 --prop font.color=FFFFFF --prop alignment.horizontal=center

AD_SPEND=(10 15 22 28 35 42 50 58 65 72 80 88 95 105 115)
SALES_REV=(45 68 95 120 155 180 220 260 290 335 370 410 445 500 550)
PROFIT=(8.5 10.2 12.1 14.5 16.8 15.2 18.3 20.1 19.5 22.3 21.8 24.5 23.1 26.8 28.2)
MKT_SHARE=(2.1 3.2 4.5 5.8 7.2 8.5 10.1 11.8 12.5 14.2 15.8 17.5 18.2 20.5 22.1)

for i in $(seq 0 14); do
    row=$((i + 2))
    officecli set "$XLSX" "/Analysis/A${row}" --prop "value=${AD_SPEND[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Analysis/B${row}" --prop "value=${SALES_REV[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Analysis/C${row}" --prop "value=${PROFIT[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Analysis/D${row}" --prop "value=${MKT_SHARE[$i]}" --prop alignment.horizontal=center
done

echo "  Done: Sheet2 data"

###############################################################################
# Sheet3: Stock data (with red/green coloring)
###############################################################################
echo "  -> Populating Sheet3: Stock data"

officecli add "$XLSX" / --type sheet --prop name=StockData

officecli set "$XLSX" '/StockData/A1' --prop value="Date" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/B1' --prop value="Open" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/C1' --prop value="High" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/D1' --prop value="Low" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/E1' --prop value="Close" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/StockData/F1' --prop value="Volume (10K)" --prop font.bold=true --prop fill=C00000 --prop font.color=FFFFFF --prop alignment.horizontal=center

DATES=("3/1" "3/2" "3/3" "3/4" "3/5" "3/6" "3/7" "3/8" "3/9" "3/10" "3/11" "3/12" "3/13" "3/14" "3/15" "3/16" "3/17" "3/18" "3/19" "3/20")
OPEN=(52.3 53.1 52.8 54.2 55.1 54.5 56.2 57.8 58.5 57.2 56.8 58.3 59.5 60.2 59.8 61.5 62.3 61.8 63.5 64.2)
HIGH=(53.8 54.2 54.5 55.8 56.3 56.8 58.1 59.2 59.8 58.5 58.2 59.8 61.2 61.5 61.8 63.2 63.8 63.5 65.2 65.8)
LOW=(51.5 52.2 51.8 53.5 54.2 53.8 55.5 56.8 57.2 56.1 55.8 57.5 58.8 59.2 58.5 60.8 61.2 60.5 62.8 63.5)
CLOSE=(53.1 52.8 54.2 55.1 54.5 56.2 57.8 58.5 57.2 56.8 58.3 59.5 60.2 59.8 61.5 62.3 61.8 63.5 64.2 65.1)
VOLUME=(285 312 268 345 298 378 425 468 395 310 352 415 485 442 368 512 548 478 562 598)

for i in $(seq 0 19); do
    row=$((i + 2))

    open=${OPEN[$i]}
    close=${CLOSE[$i]}
    if (( $(echo "$close > $open" | bc -l) )); then
        COLOR="FF0000"; BG="FFF2F2"  # Up: red
    elif (( $(echo "$close < $open" | bc -l) )); then
        COLOR="008000"; BG="F2FFF2"  # Down: green
    else
        COLOR="666666"; BG="F5F5F5"  # Flat: gray
    fi

    officecli set "$XLSX" "/StockData/A${row}" --prop "value=${DATES[$i]}" --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/B${row}" --prop "value=${OPEN[$i]}" --prop 'numFmt=0.00' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/C${row}" --prop "value=${HIGH[$i]}" --prop 'numFmt=0.00' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/D${row}" --prop "value=${LOW[$i]}" --prop 'numFmt=0.00' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/E${row}" --prop "value=${CLOSE[$i]}" --prop 'numFmt=0.00' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
    officecli set "$XLSX" "/StockData/F${row}" --prop "value=${VOLUME[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center --prop "font.color=${COLOR}" --prop "fill=${BG}"
done

echo "  Done: Sheet3 stock data (with red/green coloring)"

###############################################################################
# Sheet4: Capability radar chart data
###############################################################################
echo "  -> Populating Sheet4: Capability assessment"

officecli add "$XLSX" / --type sheet --prop name=Assessment

officecli set "$XLSX" '/Assessment/A1' --prop value="Dimension" --prop font.bold=true --prop fill=002060 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Assessment/B1' --prop value="Product A" --prop font.bold=true --prop fill=0070C0 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Assessment/C1' --prop value="Product B" --prop font.bold=true --prop fill=00B050 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Assessment/D1' --prop value="Product C" --prop font.bold=true --prop fill=FFC000 --prop font.color=000000 --prop alignment.horizontal=center

DIMS=("Performance" "Stability" "Usability" "Security" "Scalability" "Value" "Ecosystem" "Docs")
PA=(92 88 75 95 82 70 85 78)
PB=(78 92 88 80 90 85 72 82)
PC=(85 76 92 72 78 92 88 70)

for i in $(seq 0 7); do
    row=$((i + 2))
    officecli set "$XLSX" "/Assessment/A${row}" --prop "value=${DIMS[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Assessment/B${row}" --prop "value=${PA[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Assessment/C${row}" --prop "value=${PB[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Assessment/D${row}" --prop "value=${PC[$i]}" --prop alignment.horizontal=center
done

echo "  Done: Sheet4 data"

###############################################################################
# Chart 1: Combo chart (bar + line dual axis)
###############################################################################
echo "  -> Chart 1: Combo chart (bar + line dual axis)"

CHART1_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[1]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr rot="0" /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill><a:latin typeface="Microsoft YaHei" /><a:ea typeface="Microsoft YaHei" /></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></a:rPr><a:t>Monthly Sales and YoY Growth Trend</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:barChart>
        <c:barDir val="col" /><c:grouping val="clustered" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill rotWithShape="1"><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="1F4E79" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="2E75B6" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:ln w="0"><a:noFill /></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000" /></a:srgbClr></a:outerShdw></a:effectLst>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill rotWithShape="1"><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="C55A11" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="ED7D31" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:ln w="0"><a:noFill /></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000" /></a:srgbClr></a:outerShdw></a:effectLst>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill rotWithShape="1"><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="548235" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="70AD47" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:ln w="0"><a:noFill /></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000" /></a:srgbClr></a:outerShdw></a:effectLst>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1" /><c:axId val="2" />
      </c:barChart>
      <c:lineChart>
        <c:grouping val="standard" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$F$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="38100" cap="rnd"><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill><a:prstDash val="solid" /><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="circle" /><c:size val="8" />
            <c:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill></a:ln></c:spPr>
          </c:marker>
          <c:dLbls>
            <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
            <c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr>
            <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
            <c:showLegendKey val="0" /><c:showVal val="1" /><c:showCatName val="0" /><c:showSerName val="0" /><c:showPercent val="0" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$F$2:$F$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:marker val="1" />
        <c:axId val="1" /><c:axId val="3" />
      </c:lineChart>
      <c:catAx>
        <c:axId val="1" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF" /></a:solidFill></a:ln></c:spPr>
        <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="404040" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
        <c:crossAx val="2" />
      </c:catAx>
      <c:valAx>
        <c:axId val="2" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="404040" /></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Sales (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF" /></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="1" />
      </c:valAx>
      <c:valAx>
        <c:axId val="3" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="r" />
        <c:title><c:tx><c:rich><a:bodyPr rot="5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>YoY Growth (%)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000" /></a:srgbClr></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="1" /><c:crosses val="max" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" />
      <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="404040" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
    </c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>18</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 1\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART1_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Chart 1 combo chart"

###############################################################################
# Chart 2: 3D bar chart
###############################################################################
echo "  -> Chart 2: 3D bar chart"

CHART2_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[2]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>3D Regional Sales Comparison</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:view3D>
      <c:rotX val="15" /><c:rotY val="20" /><c:depthPercent val="100" /><c:rAngAx val="1" /><c:perspective val="30" />
    </c:view3D>
    <c:plotArea>
      <c:layout />
      <c:bar3DChart>
        <c:barDir val="col" /><c:grouping val="clustered" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="4472C4" /></a:gs>
              <a:gs pos="50000"><a:srgbClr val="5B9BD5" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="9DC3E6" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="ED7D31" /></a:gs>
              <a:gs pos="50000"><a:srgbClr val="F4B183" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="F8CBAD" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="70AD47" /></a:gs>
              <a:gs pos="50000"><a:srgbClr val="A9D18E" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="C5E0B4" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:shape val="cylinder" />
        <c:axId val="10" /><c:axId val="20" /><c:axId val="30" />
      </c:bar3DChart>
      <c:catAx><c:axId val="10" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="20" /></c:catAx>
      <c:valAx><c:axId val="20" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="10" /></c:valAx>
      <c:serAx><c:axId val="30" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="20" /></c:serAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>18</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>37</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"3\" name=\"Chart 2\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART2_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Chart 2 3D bar chart"

###############################################################################
# Chart 3: Scatter plot + trendline (Sheet2)
###############################################################################
echo "  -> Chart 3: Scatter plot + trendline"

CHART3_REL=$(officecli add-part "$XLSX" /Analysis --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Analysis/chart[1]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Ad Spend vs Sales Correlation</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:scatterChart>
        <c:scatterStyle val="lineMarker" />
        <c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Analysis!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="circle" /><c:size val="10" />
            <c:spPr>
              <a:solidFill><a:srgbClr val="7030A0"><a:alpha val="70000" /></a:srgbClr></a:solidFill>
              <a:ln w="19050"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:ln>
              <a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="30000" /></a:srgbClr></a:outerShdw></a:effectLst>
            </c:spPr>
          </c:marker>
          <c:trendline>
            <c:spPr><a:ln w="25400" cap="rnd"><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill><a:prstDash val="dash" /><a:round /></a:ln></c:spPr>
            <c:trendlineType val="linear" />
            <c:dispRSqr val="1" /><c:dispEq val="1" />
          </c:trendline>
          <c:xVal><c:numRef><c:f>Analysis!$A$2:$A$16</c:f></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>Analysis!$B$2:$B$16</c:f></c:numRef></c:yVal>
          <c:smooth val="0" />
        </c:ser>
        <c:axId val="100" /><c:axId val="200" />
      </c:scatterChart>
      <c:valAx>
        <c:axId val="100" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Ad Spend (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF" /></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="200" />
      </c:valAx>
      <c:valAx>
        <c:axId val="200" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Sales (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF" /></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="100" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Analysis/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>16</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 3\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART3_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Chart 3 scatter plot"

###############################################################################
# Chart 4: 3D pie chart (exploded)
###############################################################################
echo "  -> Chart 4: 3D pie chart (exploded)"

CHART4_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[3]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Annual Regional Sales Share (3D)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:view3D>
      <c:rotX val="30" /><c:rotY val="70" /><c:rAngAx val="0" /><c:perspective val="30" />
    </c:view3D>
    <c:plotArea>
      <c:layout />
      <c:pie3DChart>
        <c:varyColors val="1" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:explosion val="10" />
          <c:dPt><c:idx val="0" />
            <c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="1F4E79" /></a:gs><a:gs pos="100000"><a:srgbClr val="4472C4" /></a:gs></a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000" /></a:srgbClr></a:outerShdw></a:effectLst></c:spPr>
          </c:dPt>
          <c:dPt><c:idx val="1" />
            <c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="C55A11" /></a:gs><a:gs pos="100000"><a:srgbClr val="ED7D31" /></a:gs></a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000" /></a:srgbClr></a:outerShdw></a:effectLst></c:spPr>
          </c:dPt>
          <c:dPt><c:idx val="2" />
            <c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="548235" /></a:gs><a:gs pos="100000"><a:srgbClr val="70AD47" /></a:gs></a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000" /></a:srgbClr></a:outerShdw></a:effectLst></c:spPr>
          </c:dPt>
          <c:dLbls>
            <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
            <c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr>
            <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="1" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$D$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$8:$D$8</c:f></c:numRef></c:val>
        </c:ser>
      </c:pie3DChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>19</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>28</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"4\" name=\"Chart 4\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART4_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Chart 4 3D pie chart"

###############################################################################
# Chart 5: Bubble chart (Sheet2)
###############################################################################
echo "  -> Chart 5: Bubble chart"

CHART5_REL=$(officecli add-part "$XLSX" /Analysis --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Analysis/chart[2]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Spend-Revenue-Market Share Bubble</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:bubbleChart>
        <c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Analysis!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="7030A0"><a:alpha val="60000" /></a:srgbClr></a:solidFill>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="25000" /></a:srgbClr></a:outerShdw></a:effectLst>
          </c:spPr>
          <c:xVal><c:numRef><c:f>Analysis!$A$2:$A$16</c:f></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>Analysis!$B$2:$B$16</c:f></c:numRef></c:yVal>
          <c:bubbleSize><c:numRef><c:f>Analysis!$D$2:$D$16</c:f></c:numRef></c:bubbleSize>
          <c:bubble3D val="1" />
        </c:ser>
        <c:axId val="300" /><c:axId val="400" />
      </c:bubbleChart>
      <c:valAx>
        <c:axId val="300" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Ad Spend (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="400" />
      </c:valAx>
      <c:valAx>
        <c:axId val="400" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Sales (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="300" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Analysis/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>16</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>37</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"3\" name=\"Chart 5\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART5_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Chart 5 bubble chart"

###############################################################################
# Chart 6: Stock OHLC candlestick chart (red up, green down)
###############################################################################
echo "  -> Chart 6: Stock OHLC chart"

CHART6_REL=$(officecli add-part "$XLSX" /StockData --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/StockData/chart[1]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="C00000" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Stock Candlestick Chart (OHLC)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:stockChart>
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>StockData!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="none" /></c:marker>
          <c:cat><c:strRef><c:f>StockData!$A$2:$A$21</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>StockData!$B$2:$B$21</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>StockData!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="none" /></c:marker>
          <c:cat><c:strRef><c:f>StockData!$A$2:$A$21</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>StockData!$C$2:$C$21</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>StockData!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="none" /></c:marker>
          <c:cat><c:strRef><c:f>StockData!$A$2:$A$21</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>StockData!$D$2:$D$21</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>StockData!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="none" /></c:marker>
          <c:cat><c:strRef><c:f>StockData!$A$2:$A$21</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>StockData!$E$2:$E$21</c:f></c:numRef></c:val>
        </c:ser>
        <c:hiLowLines>
          <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="404040" /></a:solidFill></a:ln></c:spPr>
        </c:hiLowLines>
        <c:upDownBars>
          <c:gapWidth val="100" />
          <c:upBars><c:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="C00000" /></a:solidFill></a:ln></c:spPr></c:upBars>
          <c:downBars><c:spPr><a:solidFill><a:srgbClr val="00B050" /></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="006400" /></a:solidFill></a:ln></c:spPr></c:downBars>
        </c:upDownBars>
        <c:axId val="500" /><c:axId val="600" />
      </c:stockChart>
      <c:catAx>
        <c:axId val="500" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:txPr><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="800" /></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
        <c:crossAx val="600" />
      </c:catAx>
      <c:valAx>
        <c:axId val="600" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:numFmt formatCode="0.00" sourceLinked="0" />
        <c:crossAx val="500" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/StockData/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>20</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>22</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 6\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART6_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Chart 6 stock OHLC chart"

###############################################################################
# Chart 7: Filled radar chart (Sheet4)
###############################################################################
echo "  -> Chart 7: Filled radar chart"

CHART7_REL=$(officecli add-part "$XLSX" /Assessment --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Assessment/chart[1]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="002060" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Product Capability Radar Comparison</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:radarChart>
        <c:radarStyle val="filled" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Assessment!$B$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="4472C4"><a:alpha val="40000" /></a:srgbClr></a:solidFill>
            <a:ln w="28575"><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></a:ln>
          </c:spPr>
          <c:cat><c:strRef><c:f>Assessment!$A$2:$A$9</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Assessment!$B$2:$B$9</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Assessment!$C$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="00B050"><a:alpha val="40000" /></a:srgbClr></a:solidFill>
            <a:ln w="28575"><a:solidFill><a:srgbClr val="00B050" /></a:solidFill></a:ln>
          </c:spPr>
          <c:cat><c:strRef><c:f>Assessment!$A$2:$A$9</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Assessment!$C$2:$C$9</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Assessment!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="FFC000"><a:alpha val="40000" /></a:srgbClr></a:solidFill>
            <a:ln w="28575"><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></a:ln>
          </c:spPr>
          <c:cat><c:strRef><c:f>Assessment!$A$2:$A$9</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Assessment!$D$2:$D$9</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="700" /><c:axId val="800" />
      </c:radarChart>
      <c:catAx><c:axId val="700" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="800" /></c:catAx>
      <c:valAx><c:axId val="800" /><c:scaling><c:orientation val="minMax" /><c:max val="100" /><c:min val="0" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:crossAx val="700" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Assessment/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>16</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 7\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART7_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Chart 7 radar chart"

###############################################################################
# Chart 8: Multi-ring doughnut chart (2 nested series)
###############################################################################
echo "  -> Chart 8: Multi-ring doughnut chart"

CHART8_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[4]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Q3 vs Q4 Regional Sales Multi-Ring</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:doughnutChart>
        <c:varyColors val="1" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:v>Q3</c:v></c:tx>
          <c:dPt><c:idx val="0" /><c:spPr><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1" /><c:spPr><a:solidFill><a:srgbClr val="C55A11" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2" /><c:spPr><a:solidFill><a:srgbClr val="548235" /></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
            <c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr>
            <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="0" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$D$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$9:$D$9</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:v>Q4</c:v></c:tx>
          <c:dPt><c:idx val="0" /><c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1" /><c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2" /><c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
            <c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr>
            <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="1" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$D$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$13:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:holeSize val="40" />
      </c:doughnutChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>19</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>28</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>37</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"5\" name=\"Chart 8\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART8_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Chart 8 multi-ring doughnut chart"

###############################################################################
# Validation
###############################################################################
officecli close "$XLSX"

echo ""
echo "=========================================="
echo "Validating file"
echo "=========================================="
officecli validate "$XLSX"
officecli view "$XLSX" outline
echo ""
ls -lh "$XLSX"
echo ""
echo "All done! 8 chart types generated"
