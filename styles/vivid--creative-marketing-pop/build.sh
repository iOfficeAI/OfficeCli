#!/bin/bash
# 1:1 reconstruction of ppt/v2/creative-marketing.pptx
# Shape data extracted via: officecli get creative-marketing.pptx /slide[N] --depth 3 --json
set -e

BASE="$(cd "$(dirname "$0")/.." && pwd)"
PPTX="$BASE/creative-marketing.pptx"

echo "=== creative-marketing v6 (1:1 from v2) ==="
rm -f "$PPTX"
officecli create "$PPTX"

# ── S1  bg=#F5E0C0  no transition ────────────────────────────
echo "S1"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=F5E0C0
jq -n '[
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!blk-orange","fill":"E8601C","x":"0cm",    "y":"0cm",    "width":"33.87cm","height":"8.5cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!blk-dark",  "fill":"1A1A1A","x":"0cm",    "y":"8.5cm",  "width":"14cm",   "height":"10.55cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!star1","fill":"F5E0C0","opacity":"0.15","rotation":"45","x":"20cm",   "y":"1cm",    "width":"4cm",    "height":"4cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!star2","fill":"1A1A1A","opacity":"0.2", "rotation":"45","x":"26cm",   "y":"3cm",    "width":"3cm",    "height":"3cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!star3","fill":"F5E0C0",                 "rotation":"45","x":"2cm",    "y":"11cm",   "width":"2.5cm",  "height":"2.5cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!star4","fill":"E8601C",                 "rotation":"45","x":"24cm",   "y":"12cm",   "width":"5cm",    "height":"5cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!h-main", "text":"MAKE NOISE",           "font":"Impact","size":"96","bold":"true","color":"F5E0C0","fill":"none","x":"0cm",    "y":"0.4cm",  "width":"33.87cm","height":"7.8cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!h-sub",  "text":"CREATIVE MARKETING",   "font":"Impact","size":"22",             "color":"F5E0C0","fill":"none","x":"15.4cm", "y":"9.2cm",  "width":"18cm",   "height":"2cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"!!h-tag",  "text":"2025 Campaign Strategy","font":"Arial", "size":"14",             "color":"1A1A1A","fill":"none","x":"15.4cm", "y":"14.5cm", "width":"16cm",   "height":"1.4cm"}}
]' | officecli batch "$PPTX"

# ── S2  bg=#1A1A1A  transition=morph ─────────────────────────
echo "S2"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=1A1A1A
jq -n '[
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!blk-orange","fill":"E8601C","x":"0cm",    "y":"9.5cm",  "width":"33.87cm","height":"9.55cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!blk-sand",  "fill":"F5E0C0","x":"0cm",    "y":"0cm",    "width":"12cm",   "height":"9.5cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!star1","fill":"E8601C",                 "rotation":"45","x":"8cm",    "y":"1cm",    "width":"4cm",    "height":"4cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!star2","fill":"F5E0C0","opacity":"0.12","rotation":"45","x":"27cm",   "y":"10cm",   "width":"6cm",    "height":"6cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!star3","fill":"1A1A1A",                 "rotation":"45","x":"2cm",    "y":"13cm",   "width":"3cm",    "height":"3cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!star4","fill":"F5E0C0","opacity":"0.08","rotation":"45","x":"18cm",   "y":"1cm",    "width":"3cm",    "height":"3cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!h-main","text":"ATTENTION IS THE CURRENCY",                               "font":"Impact","size":"62","bold":"true","color":"F5E0C0","fill":"none","x":"1cm",    "y":"10cm",   "width":"32cm",   "height":"8cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"!!h-sub", "text":"In a world of infinite scroll, disruption is the only strategy.","font":"Arial", "size":"16",             "color":"A09080","fill":"none","x":"13.4cm", "y":"2.5cm",  "width":"18cm",   "height":"5cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[2]' --prop transition=morph

# ── S3  bg=#F5E0C0  transition=morph  (3 Channels) ───────────
echo "S3"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=F5E0C0
jq -n '[
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-orange","fill":"E8601C","x":"0cm",    "y":"0cm",    "width":"33.87cm","height":"3.2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-dark1", "fill":"1A1A1A","x":"0cm",    "y":"5.5cm",  "width":"10cm",   "height":"13.55cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-dark2", "fill":"1A1A1A","x":"12cm",   "y":"5.5cm",  "width":"10cm",   "height":"13.55cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!blk-dark3", "fill":"1A1A1A","x":"24cm",   "y":"5.5cm",  "width":"10cm",   "height":"13.55cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!star1","fill":"E8601C","rotation":"45","x":"10cm",   "y":"6.5cm",  "width":"2cm",    "height":"2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!star2","fill":"E8601C","rotation":"45","x":"22cm",   "y":"6.5cm",  "width":"2cm",    "height":"2cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!h-title","text":"3 CHANNELS",  "font":"Impact","size":"24","color":"F5E0C0","fill":"none","x":"1cm",    "y":"0.4cm",  "width":"30cm",   "height":"2.4cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c1-num",  "text":"01",          "font":"Impact","size":"48","color":"E8601C","fill":"none","x":"1.6cm",  "y":"6cm",    "width":"8cm",    "height":"3.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c1-title","text":"SOCIAL",       "font":"Impact","size":"30","color":"F5E0C0","fill":"none","x":"1.6cm",  "y":"9.2cm",  "width":"8cm",    "height":"2.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c1-body", "text":"Organic reach amplified by community. Make every post a conversation starter.","font":"Arial","size":"13","color":"C0A888","fill":"none","x":"1.6cm","y":"11.8cm","width":"8cm","height":"5.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c2-num",  "text":"02",          "font":"Impact","size":"48","color":"E8601C","fill":"none","x":"13.6cm", "y":"6cm",    "width":"8cm",    "height":"3.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c2-title","text":"OOH",          "font":"Impact","size":"30","color":"F5E0C0","fill":"none","x":"13.6cm", "y":"9.2cm",  "width":"8cm",    "height":"2.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c2-body", "text":"Out-of-home placements designed to stop people mid-stride. Bold. Unmissable.","font":"Arial","size":"13","color":"C0A888","fill":"none","x":"13.6cm","y":"11.8cm","width":"8cm","height":"5.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c3-num",  "text":"03",          "font":"Impact","size":"48","color":"E8601C","fill":"none","x":"25.6cm", "y":"6cm",    "width":"8cm",    "height":"3.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c3-title","text":"EVENTS",       "font":"Impact","size":"30","color":"F5E0C0","fill":"none","x":"25.6cm", "y":"9.2cm",  "width":"8cm",    "height":"2.5cm"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"!!c3-body", "text":"Live experiences that turn audiences into advocates. Be the moment people talk about.","font":"Arial","size":"13","color":"C0A888","fill":"none","x":"25.6cm","y":"11.8cm","width":"8cm","height":"5.5cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[3]' --prop transition=morph

# ── S4  bg=#E8601C  transition=morph  (Results) ───────────────
echo "S4"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=E8601C
jq -n '[
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!blk-sand","fill":"F5E0C0","x":"0cm",    "y":"0cm",    "width":"20cm",   "height":"19.05cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!blk-dark","fill":"1A1A1A","x":"0cm",    "y":"12cm",   "width":"20cm",   "height":"7.05cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!star1","fill":"E8601C",                 "rotation":"45","x":"16cm",   "y":"1cm",    "width":"5cm",    "height":"5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!star2","fill":"1A1A1A","opacity":"0.15","rotation":"45","x":"24cm",   "y":"4cm",    "width":"8cm",    "height":"8cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!star3","fill":"F5E0C0","opacity":"0.2", "rotation":"45","x":"26cm",   "y":"12cm",   "width":"6cm",    "height":"6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!star4","fill":"1A1A1A",                 "rotation":"45","x":"2cm",    "y":"14cm",   "width":"2.5cm",  "height":"2.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!h-title", "text":"RESULTS",                               "font":"Impact","size":"20","color":"1A1A1A","fill":"none","x":"1.6cm",  "y":"0.8cm",  "width":"15cm",   "height":"1.6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d1-num",  "text":"380%",                                  "font":"Impact","size":"72","color":"E8601C","fill":"none","x":"1.2cm",  "y":"2.4cm",  "width":"17cm",   "height":"6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d1-label","text":"ROI on social-first campaigns",         "font":"Arial", "size":"16","color":"6A5040","fill":"none","x":"1.6cm",  "y":"8.4cm",  "width":"17cm",   "height":"1.6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d2-num",  "text":"12M+",                                  "font":"Impact","size":"44","color":"F5E0C0","fill":"none","x":"1.6cm",  "y":"12.4cm", "width":"17cm",   "height":"3.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d2-label","text":"Earned media impressions in 90 days",   "font":"Arial", "size":"14","color":"A09888","fill":"none","x":"1.6cm",  "y":"15.6cm", "width":"17cm",   "height":"1.6cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d3-num",  "text":"#1",                                    "font":"Impact","size":"60","color":"1A1A1A","fill":"none","x":"22cm",   "y":"4cm",    "width":"10cm",   "height":"5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"!!d3-label","text":"Trending topic, 3 consecutive weeks",   "font":"Arial", "size":"14","color":"E8601C","fill":"none","x":"22cm",   "y":"9cm",    "width":"10cm",   "height":"2cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[4]' --prop transition=morph

# ── S5  bg=#F5E0C0  transition=morph  (CTA) ──────────────────
echo "S5"
officecli add "$PPTX" '/' --type slide --prop layout=blank --prop background=F5E0C0
jq -n '[
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!blk-dark",  "fill":"1A1A1A","x":"0cm",    "y":"0cm",    "width":"33.87cm","height":"13cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!blk-orange","fill":"E8601C","x":"0cm",    "y":"13cm",   "width":"33.87cm","height":"6.05cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!star1","fill":"E8601C",                 "rotation":"45","x":"0cm",    "y":"3cm",    "width":"5cm",    "height":"5cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!star2","fill":"F5E0C0","opacity":"0.08","rotation":"45","x":"26cm",   "y":"0cm",    "width":"8cm",    "height":"8cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!star3","fill":"1A1A1A",                 "rotation":"45","x":"28cm",   "y":"14cm",   "width":"4cm",    "height":"4cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!star4","fill":"F5E0C0",                 "rotation":"45","x":"8cm",    "y":"14.5cm", "width":"3cm",    "height":"3cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!h-main","text":"GO BOLD OR GO HOME",                                                                 "font":"Impact","size":"68","bold":"true","color":"F5E0C0","fill":"none","x":"1cm",    "y":"2cm",    "width":"32cm",   "height":"9cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"!!h-sub", "text":"The brief is ready. The strategy is set. Now it is time to execute.",               "font":"Arial", "size":"17",             "color":"F5E0C0","fill":"none","x":"1.6cm",  "y":"13.8cm", "width":"24cm",   "height":"2cm"}}
]' | officecli batch "$PPTX"
officecli set "$PPTX" '/slide[5]' --prop transition=morph

officecli validate "$PPTX"
officecli view "$PPTX" outline
echo "Done: $PPTX"
