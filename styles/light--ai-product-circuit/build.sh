#!/bin/bash
set -e

BASE="$(cd "$(dirname "$0")/.." && pwd)"
PPTX="$BASE/ai-product.pptx"
TP="$BASE/assets/tech-portrait.jpg"

echo "=== ai-product v4 ==="
rm -f "$PPTX"
officecli create "$PPTX"

# ── S1 HERO (0D1B4B) ─────────────────────────────────────────
echo "S1 hero"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=0D1B4B
jq -n --arg tp "$TP" '[
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"orb-a","preset":"ellipse","fill":"0066FF","opacity":"0.32","x":"12cm","y":"0cm","width":"22cm","height":"17cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"orb-b","preset":"ellipse","fill":"00D4FF","opacity":"0.22","x":"17cm","y":"4cm","width":"14cm","height":"13cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"orb-c","preset":"ellipse","fill":"FF69B4","opacity":"0.18","x":"0cm","y":"10cm","width":"12cm","height":"10cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.16","x":"24cm","y":"0cm","width":"10cm","height":"9cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.20","x":"0cm","y":"0cm","width":"8cm","height":"7cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.05","x":"16cm","y":"5cm","width":"14cm","height":"12cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"photo-1","image":$tp,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"tag","text":"AI PRODUCT","font":"Segoe UI","size":"11","bold":"true","color":"00D4FF","fill":"none","x":"1.6cm","y":"6.5cm","width":"12cm","height":"0.7cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"h-title","text":"Intelligence, Redefined","font":"Segoe UI","size":"52","bold":"true","color":"FFFFFF","fill":"none","x":"1.6cm","y":"7.4cm","width":"18cm","height":"5.5cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"h-sub","text":"The AI platform built for the way humans actually work.","font":"Segoe UI","size":"17","color":"7AB8FF","fill":"none","x":"1.6cm","y":"13.5cm","width":"16cm","height":"2cm"}}
]' | officecli batch "$PPTX"

# ── S2 STATEMENT (060D24) ────────────────────────────────────
echo "S2 statement"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=060D24
jq -n --arg tp "$TP" '[
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"orb-a","preset":"ellipse","fill":"FF69B4","opacity":"0.28","x":"0cm","y":"2cm","width":"18cm","height":"17cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"orb-b","preset":"ellipse","fill":"00D4FF","opacity":"0.20","x":"2cm","y":"5cm","width":"13cm","height":"12cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"orb-c","preset":"ellipse","fill":"0066FF","opacity":"0.22","x":"3cm","y":"0cm","width":"15cm","height":"10cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.15","x":"26cm","y":"10cm","width":"10cm","height":"9cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.18","x":"22cm","y":"14cm","width":"12cm","height":"6cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.05","x":"4cm","y":"6cm","width":"10cm","height":"9cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"photo-1","image":$tp,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"tag","text":"","font":"Segoe UI","size":"11","color":"00D4FF","fill":"none","x":"15.5cm","y":"5cm","width":"4cm","height":"0.7cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"h-title","text":"Every decision, enhanced by AI.","font":"Segoe UI","size":"44","bold":"true","color":"FFFFFF","fill":"none","x":"15.5cm","y":"6cm","width":"17cm","height":"6cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"h-sub","text":"Not a tool. A thinking partner that scales with your ambition.","font":"Segoe UI","size":"17","color":"7AB8FF","fill":"none","x":"15.5cm","y":"12.5cm","width":"16cm","height":"3cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[2]' --prop transition=morph

# ── S3 PILLARS (0D1B4B) ─────────────────────────────────────
echo "S3 pillars"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=0D1B4B
jq -n --arg tp "$TP" '[
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"orb-a","preset":"ellipse","fill":"0066FF","opacity":"0.10","x":"18cm","y":"7cm","width":"16cm","height":"14cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"orb-b","preset":"ellipse","fill":"00D4FF","opacity":"0.08","x":"24cm","y":"0cm","width":"12cm","height":"10cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"orb-c","preset":"ellipse","fill":"FF69B4","opacity":"0.08","x":"26cm","y":"12cm","width":"10cm","height":"8cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.07","x":"0cm","y":"14cm","width":"10cm","height":"6cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.08","x":"0cm","y":"0cm","width":"6cm","height":"5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.04","x":"12cm","y":"6cm","width":"20cm","height":"14cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"photo-1","image":$tp,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.07","x":"1.6cm","y":"4.4cm","width":"9.6cm","height":"11cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.07","x":"12.4cm","y":"4.4cm","width":"9.6cm","height":"11cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.07","x":"23.2cm","y":"4.4cm","width":"9.6cm","height":"11cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"tag","text":"THREE CAPABILITIES","font":"Segoe UI","size":"13","bold":"true","color":"FFFFFF","fill":"none","x":"1.6cm","y":"1cm","width":"20cm","height":"1.6cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"h-title","text":"Predict          Automate          Adapt","font":"Segoe UI","size":"22","bold":"true","color":"FFFFFF","fill":"none","x":"1.6cm","y":"7.5cm","width":"30cm","height":"2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"h-sub","text":"Anticipate outcomes. Eliminate routine. Refine continuously.","font":"Segoe UI","size":"14","color":"7AB8FF","fill":"none","x":"1.6cm","y":"10cm","width":"30cm","height":"3cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[3]' --prop transition=morph

# ── S4 EVIDENCE (060D24) ────────────────────────────────────
echo "S4 evidence"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=060D24
jq -n --arg tp "$TP" '[
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"orb-a","preset":"ellipse","fill":"00D4FF","opacity":"0.28","x":"2cm","y":"2cm","width":"16cm","height":"14cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"orb-b","preset":"ellipse","fill":"FF69B4","opacity":"0.20","x":"0cm","y":"6cm","width":"13cm","height":"13cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"orb-c","preset":"ellipse","fill":"0066FF","opacity":"0.22","x":"4cm","y":"0cm","width":"12cm","height":"10cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.14","x":"22cm","y":"10cm","width":"12cm","height":"10cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.16","x":"24cm","y":"4cm","width":"10cm","height":"7cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.04","x":"8cm","y":"4cm","width":"16cm","height":"13cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"photo-1","image":$tp,"x":"0cm","y":"0cm","width":"16cm","height":"19.05cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"tag","text":"RESULTS","font":"Segoe UI","size":"13","bold":"true","color":"FFFFFF","fill":"none","x":"17.6cm","y":"0.8cm","width":"14cm","height":"1cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"h-title","text":"10x","font":"Segoe UI","size":"72","bold":"true","color":"FFFFFF","fill":"none","x":"17.6cm","y":"2.4cm","width":"14cm","height":"5.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"h-sub","text":"Faster decisions\n\n$4.2M  Annual savings\n\n99.97%  Uptime SLA","font":"Segoe UI","size":"15","color":"7AB8FF","fill":"none","x":"17.6cm","y":"8.5cm","width":"14cm","height":"8cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[4]' --prop transition=morph

# ── S5 CTA (0D1B4B) ─────────────────────────────────────────
echo "S5 CTA"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=0D1B4B
jq -n --arg tp "$TP" '[
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"orb-a","preset":"ellipse","fill":"0066FF","opacity":"0.30","x":"18cm","y":"1cm","width":"18cm","height":"17cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"orb-b","preset":"ellipse","fill":"00D4FF","opacity":"0.22","x":"22cm","y":"4cm","width":"13cm","height":"13cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"orb-c","preset":"ellipse","fill":"FF69B4","opacity":"0.16","x":"28cm","y":"10cm","width":"8cm","height":"8cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.14","x":"22cm","y":"0cm","width":"10cm","height":"8cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.18","x":"0cm","y":"14cm","width":"8cm","height":"6cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.05","x":"20cm","y":"6cm","width":"14cm","height":"12cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"photo-1","image":$tp,"x":"20cm","y":"0cm","width":"10cm","height":"19.05cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"tag","text":"GET STARTED","font":"Segoe UI","size":"11","bold":"true","color":"00D4FF","fill":"none","x":"1.6cm","y":"5.5cm","width":"14cm","height":"0.7cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"h-title","text":"Request Early Access","font":"Segoe UI","size":"46","bold":"true","color":"FFFFFF","fill":"none","x":"1.6cm","y":"6.4cm","width":"17cm","height":"5.5cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"h-sub","text":"Join 200+ enterprises already building with us.","font":"Segoe UI","size":"17","color":"7AB8FF","fill":"none","x":"1.6cm","y":"12.5cm","width":"16cm","height":"2cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"cta","text":"Get started ->","font":"Segoe UI","size":"15","bold":"true","color":"FFFFFF","fill":"0066FF","x":"1.6cm","y":"15.4cm","width":"9cm","height":"1.8cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[5]' --prop transition=morph

officecli validate "$PPTX"
officecli view "$PPTX" outline
echo "Done: $PPTX"
