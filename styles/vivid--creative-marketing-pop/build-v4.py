import json, subprocess, sys, copy

BASE = "/Users/visher/Library/Application Support/AionUI/aionui/claude-temp-1774010336920/ppt/v4"
PPTX = f"{BASE}/creative-marketing.pptx"
BP   = f"{BASE}/assets/bold-portrait.jpg"
BC   = f"{BASE}/assets/bold-color.jpg"

def batch(ops):
    r = subprocess.run(["officecli","batch",PPTX],input=json.dumps(ops),text=True,capture_output=True)
    l=[x for x in r.stdout.splitlines() if "complete" in x]
    print(" ",(l[0] if l else r.stdout[:80]))
    if r.returncode!=0: print("  ERR:",r.stderr[:120])

def add_slide(n,bg): batch([{"command":"add","parent":"/","type":"slide","props":{"layout":"blank","background":bg}}])
def add_shapes(n,shapes):
    ops=[{"command":"add","parent":f"/slide[{n}]","type":"shape","props":copy.deepcopy(s)} for s in shapes]
    batch(ops)
def morph(n): batch([{"command":"set","path":f"/slide[{n}]","props":{"transition":"morph"}}])

subprocess.run(["officecli","create",PPTX],capture_output=True)
print("=== creative-marketing v4 ===")

# MORPH VOCAB: blk-a(orange) blk-b(dark) blk-c(sand)
# star-1 star-2 star-3 (rotation rects)
# photo-1(BP) photo-2(BC)  h-main h-sub
# 核心策略: blk-a(橙)像"活塞"每页剧烈变换形态
# S1:顶条→S2:底条→S3:细顶条→S4:巨左块→S5:底条

# ── S1 HERO (F5E0C0 sand) ──
print("S1 hero")
add_slide(1,"F5E0C0")
add_shapes(1,[
  # blk-a 橙色占满上方
  {"name":"blk-a","fill":"E8601C","x":"0cm",  "y":"0cm",    "width":"33.87cm","height":"9cm"},
  {"name":"blk-b","fill":"1A1A1A","x":"0cm",  "y":"14cm",   "width":"16cm",   "height":"5.05cm"},
  {"name":"blk-c","fill":"F5E0C0","x":"16cm", "y":"14cm",   "width":"17.87cm","height":"5.05cm"},
  {"name":"star-1","fill":"F5E0C0","rotation":"45","x":"19cm","y":"1cm",    "width":"6cm",    "height":"6cm"},
  {"name":"star-2","fill":"1A1A1A","rotation":"45","x":"26cm","y":"2.5cm",  "width":"4cm",    "height":"4cm"},
  {"name":"star-3","fill":"E8601C","rotation":"45","x":"2.5cm","y":"10.5cm","width":"2.5cm",  "height":"2.5cm"},
  {"name":"photo-1","image":BP,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"photo-2","image":BC,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"h-main","text":"MAKE NOISE","font":"Impact","size":"92","bold":"true","color":"F5E0C0","fill":"none","x":"0.5cm","y":"0.5cm","width":"33cm","height":"8cm"},
  {"name":"h-sub", "text":"CREATIVE MARKETING  —  2025 CAMPAIGN","font":"Impact","size":"18","color":"F5E0C0","fill":"none","x":"16.5cm","y":"10cm","width":"16cm","height":"2cm"},
])

# ── S2 STATEMENT (1A1A1A dark) ──
# blk-a 从顶部→掉落到底部 (整个版面翻转!)
# blk-b 从底部左块→移到顶部左块
# photo-1 从角落飞出到上方沙色区
print("S2 statement")
add_slide(2,"1A1A1A")
add_shapes(2,[
  {"name":"blk-a","fill":"E8601C","x":"0cm",   "y":"10.05cm","width":"33.87cm","height":"9cm"},
  {"name":"blk-b","fill":"1A1A1A","x":"0cm",   "y":"0cm",    "width":"12cm",   "height":"10.05cm"},
  {"name":"blk-c","fill":"F5E0C0","x":"12cm",  "y":"0cm",    "width":"21.87cm","height":"10.05cm"},
  {"name":"star-1","fill":"F5E0C0","rotation":"45","x":"29cm","y":"11cm","width":"5cm","height":"5cm"},
  {"name":"star-2","fill":"1A1A1A","rotation":"45","x":"2cm", "y":"12.5cm","width":"3.5cm","height":"3.5cm"},
  {"name":"star-3","fill":"E8601C","rotation":"45","x":"10.5cm","y":"0.5cm","width":"3.5cm","height":"3.5cm"},
  # photo-1 进入沙色区
  {"name":"photo-1","image":BP,"x":"12cm","y":"0cm","width":"9cm","height":"10.05cm"},
  {"name":"photo-2","image":BC,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"h-main","text":"ATTENTION IS THE CURRENCY","font":"Impact","size":"56","bold":"true","color":"F5E0C0","fill":"none","x":"0.5cm","y":"10.5cm","width":"32cm","height":"8cm"},
  {"name":"h-sub", "text":"In a world of infinite scroll, disruption is the only strategy.","font":"Arial","size":"15","color":"9A9080","fill":"none","x":"21cm","y":"2cm","width":"11cm","height":"6cm"},
])
morph(2)

# ── S3 PILLARS (F5E0C0 sand) ──
# blk-a 从底部→压缩成细顶条 (剧烈压缩)
# blk-b 从顶部左块→变成左列深色背景
# photo-1 从沙色区缩小进入左列; photo-2 出现在中列
print("S3 pillars")
add_slide(3,"F5E0C0")
add_shapes(3,[
  {"name":"blk-a","fill":"E8601C","x":"0cm","y":"0cm","width":"33.87cm","height":"3.2cm"},
  {"name":"blk-b","fill":"1A1A1A","x":"0cm","y":"5.5cm","width":"10cm","height":"13.55cm"},
  {"name":"blk-c","fill":"1A1A1A","x":"12cm","y":"5.5cm","width":"10cm","height":"13.55cm"},
  {"name":"star-1","fill":"E8601C","rotation":"45","x":"10.2cm","y":"6.8cm","width":"1.8cm","height":"1.8cm"},
  {"name":"star-2","fill":"E8601C","rotation":"45","x":"22.2cm","y":"6.8cm","width":"1.8cm","height":"1.8cm"},
  {"name":"star-3","fill":"F5E0C0","rotation":"45","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"photo-1","image":BP,"x":"0cm",  "y":"5.5cm","width":"10cm","height":"13.55cm"},
  {"name":"photo-2","image":BC,"x":"12cm", "y":"5.5cm","width":"10cm","height":"13.55cm"},
  {"name":"h-main","text":"3 CHANNELS","font":"Impact","size":"22","color":"F5E0C0","fill":"none","x":"1cm","y":"0.5cm","width":"30cm","height":"2.4cm"},
  {"name":"h-sub", "text":"01 SOCIAL\n\n02 OOH\n\n03 EVENTS","font":"Impact","size":"20","color":"E8601C","fill":"none","x":"23cm","y":"6cm","width":"10cm","height":"12cm"},
])
morph(3)

# ── S4 EVIDENCE (E8601C full orange) ──
# blk-a 从细顶条→爆炸成大左块 (最震撼!)
# photo-1 飞到右侧上方; photo-2 消失
print("S4 evidence")
add_slide(4,"E8601C")
add_shapes(4,[
  {"name":"blk-a","fill":"E8601C","x":"0cm","y":"0cm","width":"19cm","height":"19.05cm"},
  {"name":"blk-b","fill":"1A1A1A","x":"19cm","y":"11cm","width":"14.87cm","height":"8.05cm"},
  {"name":"blk-c","fill":"F5E0C0","x":"19cm","y":"0cm", "width":"14.87cm","height":"11cm"},
  {"name":"star-1","fill":"F5E0C0","rotation":"45","x":"19.5cm","y":"0.5cm","width":"5cm","height":"5cm"},
  {"name":"star-2","fill":"1A1A1A","rotation":"45","x":"29.5cm","y":"12cm", "width":"4cm","height":"4cm"},
  {"name":"star-3","fill":"F5E0C0","rotation":"45","x":"2cm",   "y":"11cm", "width":"3cm","height":"3cm"},
  {"name":"photo-1","image":BP,"x":"19.5cm","y":"0cm","width":"14.37cm","height":"11cm"},
  {"name":"photo-2","image":BC,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"h-main","text":"380%","font":"Impact","size":"72","bold":"true","color":"F5E0C0","fill":"none","x":"0.8cm","y":"2cm","width":"17cm","height":"9cm"},
  {"name":"h-sub", "text":"ROI on social-first campaigns\n\n12M+  Earned impressions\n\n#1    Trending 3 weeks","font":"Arial","size":"15","color":"F5E0C0","fill":"none","x":"0.8cm","y":"12cm","width":"17cm","height":"7cm"},
])
morph(4)

# ── S5 CTA (1A1A1A dark) ──
# blk-a 从大左块→变成底部窄条 (再次翻转)
# blk-b 扩张成顶部2/3 dark区
# photo-2 从角落飞出成大背景图
print("S5 CTA")
add_slide(5,"1A1A1A")
add_shapes(5,[
  {"name":"blk-a","fill":"E8601C","x":"0cm","y":"13cm","width":"33.87cm","height":"6.05cm"},
  {"name":"blk-b","fill":"1A1A1A","x":"0cm","y":"0cm", "width":"33.87cm","height":"13cm"},
  {"name":"blk-c","fill":"F5E0C0","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"star-1","fill":"F5E0C0","rotation":"45","x":"0.5cm", "y":"14cm","width":"5cm","height":"5cm"},
  {"name":"star-2","fill":"1A1A1A","rotation":"45","x":"27cm",  "y":"13.5cm","width":"3.5cm","height":"3.5cm"},
  {"name":"star-3","fill":"F5E0C0","rotation":"45","x":"8.5cm", "y":"14cm","width":"3cm","height":"3cm"},
  {"name":"photo-1","image":BP,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"photo-2","image":BC,"opacity":"0.35","x":"0cm","y":"0cm","width":"33.87cm","height":"13cm"},
  {"name":"h-main","text":"GO BOLD OR GO HOME","font":"Impact","size":"62","bold":"true","color":"F5E0C0","fill":"none","x":"1cm","y":"1cm","width":"32cm","height":"11cm"},
  {"name":"h-sub", "text":"The brief is ready. The strategy is set. Time to execute.","font":"Arial","size":"17","color":"F5E0C0","fill":"none","x":"1.6cm","y":"13.8cm","width":"24cm","height":"2cm"},
])
morph(5)

r=subprocess.run(["officecli","validate",PPTX],capture_output=True,text=True)
print("validate:","PASS" if "no errors" in r.stdout else r.stdout[:100])
r2=subprocess.run(["officecli","view",PPTX,"outline"],capture_output=True,text=True)
print(r2.stdout)
