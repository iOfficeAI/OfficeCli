#!/bin/bash
# 1:1 reconstruction of ppt/v2/brand-strategy.pptx
# Shape data extracted via: officecli get brand-strategy.pptx /slide[N] --depth 3 --json
set -e

BASE="$(cd "$(dirname "$0")/.." && pwd)"
PPTX="$BASE/brand-strategy.pptx"

echo "=== brand-strategy v6 (1:1 from v2) ==="
rm -f "$PPTX"
officecli create "$PPTX"

# ── S1  bg=#F5F0E8  no transition ────────────────────────────
echo "S1"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=F5F0E8
jq -n '[
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!blk-navy",  "fill":"162040","x":"19cm", "y":"0cm",   "width":"7cm",    "height":"12cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!blk-blue",  "fill":"1A6BFF","x":"26cm", "y":"0cm",   "width":"8cm",    "height":"8cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!blk-cyan",  "fill":"00C9D4","x":"26cm", "y":"8cm",   "width":"4cm",    "height":"4cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!blk-orange","fill":"F4713A","x":"30cm", "y":"8cm",   "width":"4cm",    "height":"8cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!blk-green", "fill":"7EC8A0","x":"19cm", "y":"12cm",  "width":"4cm",    "height":"7cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!blk-pink",  "fill":"E8749A","x":"23cm", "y":"14cm",  "width":"3cm",    "height":"5.05cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!tag-line",  "text":"BRAND STRATEGY","font":"Arial","size":"11","bold":"true","color":"9A9080","fill":"none","x":"1.6cm","y":"1.2cm","width":"12cm","height":"0.7cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!h-title",   "text":"Brand Refresh 2025","font":"Arial","size":"58","bold":"true","color":"162040","fill":"none","x":"1.6cm","y":"4cm","width":"16cm","height":"7cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!h-sub",     "text":"A new visual language for the next era of growth","font":"Arial","size":"16","color":"6B6355","fill":"none","x":"1.6cm","y":"13.5cm","width":"15cm","height":"2cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!dot-accent","preset":"ellipse","fill":"F4713A","x":"1.6cm","y":"13.2cm","width":"0.5cm","height":"0.5cm"}}
]' | officecli batch "$PPTX"

# ── S2  bg=#162040  transition=morph ─────────────────────────
echo "S2"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=162040
jq -n '[
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!blk-blue",  "fill":"1A6BFF","x":"0cm","y":"0cm",  "width":"6cm", "height":"19.05cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!blk-orange","fill":"F4713A","x":"0cm","y":"0cm",  "width":"6cm", "height":"6cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!blk-cyan",  "fill":"00C9D4","x":"27cm","y":"11cm","width":"7cm", "height":"8.05cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!blk-green", "fill":"7EC8A0","x":"27cm","y":"0cm", "width":"7cm", "height":"5cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!blk-pink",  "fill":"E8749A","x":"27cm","y":"5cm", "width":"7cm", "height":"6cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!blk-navy",  "fill":"0A0F1E","x":"6cm","y":"15cm", "width":"21cm","height":"4.05cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!h-title",   "text":"Clarity beats complexity.","font":"Arial","size":"52","bold":"true","color":"F5F0E8","fill":"none","x":"7.6cm","y":"5.5cm","width":"18cm","height":"8cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!h-sub",     "text":"The strongest brands say less — and mean more.","font":"Arial","size":"17","color":"A09888","fill":"none","x":"7.6cm","y":"13cm","width":"17cm","height":"2cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[2]' --prop transition=morph

# ── S3  bg=#F5F0E8  transition=morph  (Three Pillars) ────────
echo "S3"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=F5F0E8
jq -n '[
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-navy",  "fill":"162040","x":"0cm",   "y":"0cm",     "width":"33.87cm","height":"2.4cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-blue",  "fill":"1A6BFF","x":"1.6cm", "y":"4cm",     "width":"9.2cm",  "height":"13.2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-orange","fill":"F4713A","x":"12.4cm","y":"4cm",     "width":"9.2cm",  "height":"13.2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-cyan",  "fill":"00C9D4","x":"23.2cm","y":"4cm",     "width":"9.2cm",  "height":"13.2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-green", "fill":"7EC8A0","x":"0cm",   "y":"17.05cm", "width":"11cm",   "height":"2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-pink",  "fill":"E8749A","x":"11cm",  "y":"17.05cm", "width":"11cm",   "height":"2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!h-title",   "text":"Three Pillars","font":"Arial","size":"20","bold":"true","color":"F5F0E8","fill":"none","x":"1.6cm","y":"0.4cm","width":"20cm","height":"1.6cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p1-num",  "text":"01","font":"Arial","size":"13","bold":"true","color":"F5F0E8","fill":"none","x":"2.2cm","y":"4.6cm","width":"4cm","height":"1.2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p1-title","text":"Identity","font":"Arial","size":"28","bold":"true","color":"F5F0E8","fill":"none","x":"2.2cm","y":"6cm","width":"8cm","height":"2.4cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p1-body", "text":"A consistent visual system that speaks before words do.","font":"Arial","size":"14","color":"E8E0D4","fill":"none","x":"2.2cm","y":"8.8cm","width":"8cm","height":"5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p2-num",  "text":"02","font":"Arial","size":"13","bold":"true","color":"F5F0E8","fill":"none","x":"13cm","y":"4.6cm","width":"4cm","height":"1.2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p2-title","text":"Voice","font":"Arial","size":"28","bold":"true","color":"F5F0E8","fill":"none","x":"13cm","y":"6cm","width":"8cm","height":"2.4cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p2-body", "text":"The tone and language that builds trust across every touchpoint.","font":"Arial","size":"14","color":"E8E0D4","fill":"none","x":"13cm","y":"8.8cm","width":"8cm","height":"5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p3-num",  "text":"03","font":"Arial","size":"13","bold":"true","color":"F5F0E8","fill":"none","x":"23.8cm","y":"4.6cm","width":"4cm","height":"1.2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p3-title","text":"Experience","font":"Arial","size":"28","bold":"true","color":"F5F0E8","fill":"none","x":"23.8cm","y":"6cm","width":"8cm","height":"2.4cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!p3-body", "text":"Every interaction, designed with purpose and precision.","font":"Arial","size":"14","color":"E8E0D4","fill":"none","x":"23.8cm","y":"8.8cm","width":"8cm","height":"5cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[3]' --prop transition=morph

# ── S4  bg=#F5F0E8  transition=morph  (The Numbers) ──────────
echo "S4"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=F5F0E8
jq -n '[
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!blk-blue",  "fill":"1A6BFF","x":"0cm", "y":"3.6cm", "width":"16cm","height":"9.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!blk-navy",  "fill":"162040","x":"0cm", "y":"13.1cm","width":"16cm","height":"6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!blk-orange","fill":"F4713A","x":"16cm","y":"11.5cm","width":"18cm","height":"7.55cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!blk-cyan",  "fill":"00C9D4","x":"16cm","y":"0cm",   "width":"9cm", "height":"6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!blk-pink",  "fill":"E8749A","x":"25cm","y":"0cm",   "width":"9cm", "height":"6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!blk-green", "fill":"7EC8A0","x":"16cm","y":"6cm",   "width":"18cm","height":"5.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!h-title", "text":"The Numbers","font":"Arial","size":"20","bold":"true","color":"162040","fill":"none","x":"1.6cm","y":"0.6cm","width":"15cm","height":"1.6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d1-num",  "text":"+42%","font":"Arial","size":"52","bold":"true","color":"F5F0E8","fill":"none","x":"1cm","y":"4.4cm","width":"14cm","height":"4cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d1-label","text":"Brand recognition lift after refresh","font":"Arial","size":"14","color":"C8E0F5","fill":"none","x":"1cm","y":"8.2cm","width":"14cm","height":"1.6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d2-num",  "text":"2.8x","font":"Arial","size":"40","bold":"true","color":"162040","fill":"none","x":"17cm","y":"6.6cm","width":"14cm","height":"3cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d2-label","text":"Engagement rate increase","font":"Arial","size":"14","color":"3A5040","fill":"none","x":"17cm","y":"9.4cm","width":"14cm","height":"1.6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d3-num",  "text":"89","font":"Arial","size":"40","bold":"true","color":"F5F0E8","fill":"none","x":"17cm","y":"12.2cm","width":"8cm","height":"3cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d3-label","text":"Net Promoter Score","font":"Arial","size":"14","color":"F5D0C0","fill":"none","x":"17cm","y":"15cm","width":"14cm","height":"1.6cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[4]' --prop transition=morph

# ── S5  bg=#F5F0E8  transition=morph  (CTA) ──────────────────
echo "S5"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=F5F0E8
jq -n '[
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!blk-navy",  "fill":"162040","x":"0cm", "y":"0cm",  "width":"33.87cm","height":"11cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!blk-orange","fill":"F4713A","x":"0cm", "y":"11cm", "width":"8cm",    "height":"8.05cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!blk-blue",  "fill":"1A6BFF","x":"8cm", "y":"11cm", "width":"6cm",    "height":"8.05cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!blk-cyan",  "fill":"00C9D4","x":"14cm","y":"11cm", "width":"4cm",    "height":"8.05cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!blk-green", "fill":"7EC8A0","x":"18cm","y":"11cm", "width":"16cm",   "height":"8.05cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!blk-pink",  "fill":"E8749A","x":"18cm","y":"11cm", "width":"4cm",    "height":"8.05cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!h-title","text":"Start the transformation.","font":"Arial","size":"52","bold":"true","color":"F5F0E8","fill":"none","x":"1.6cm","y":"2cm","width":"28cm","height":"6cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!h-sub",  "text":"Your brand is ready for what comes next.","font":"Arial","size":"18","color":"9A9898","fill":"none","x":"1.6cm","y":"8.4cm","width":"20cm","height":"1.8cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!arrow",  "text":"->","font":"Arial","size":"28","bold":"true","color":"F4713A","fill":"none","x":"1.6cm","y":"12.2cm","width":"3cm","height":"2cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[5]' --prop transition=morph

officecli validate "$PPTX"
officecli view "$PPTX" outline
echo "Done: $PPTX"
