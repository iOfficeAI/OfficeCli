import json, subprocess, sys, copy

BASE = "/Users/visher/Library/Application Support/AionUI/aionui/claude-temp-1774010336920/ppt/v4"
PPTX = f"{BASE}/ai-product.pptx"
TP   = f"{BASE}/assets/tech-portrait.jpg"   # dark editorial portrait
TC   = f"{BASE}/assets/tech-circuit.jpg"    # circuit board

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
print("=== ai-product v4 ===")

# MORPH VOCAB: orb-a~f(ellipse), photo-1(TP), card-1~3(roundRect), h-title/h-sub/tag
# 核心策略: orb群体每页戏剧性漂移，photo-1从角落爆发再收回，cards聚散

# ── S1 HERO (0D1B4B) ──
# orb群: 聚集在右侧 / photo-1和cards隐于角落
print("S1 hero")
add_slide(1,"0D1B4B")
add_shapes(1,[
  {"name":"orb-a","preset":"ellipse","fill":"0066FF","opacity":"0.32","x":"12cm","y":"0cm","width":"22cm","height":"17cm"},
  {"name":"orb-b","preset":"ellipse","fill":"00D4FF","opacity":"0.22","x":"17cm","y":"4cm","width":"14cm","height":"13cm"},
  {"name":"orb-c","preset":"ellipse","fill":"FF69B4","opacity":"0.18","x":"0cm", "y":"10cm","width":"12cm","height":"10cm"},
  {"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.16","x":"24cm","y":"0cm", "width":"10cm","height":"9cm"},
  {"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.20","x":"0cm", "y":"0cm", "width":"8cm", "height":"7cm"},
  {"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.05","x":"16cm","y":"5cm", "width":"14cm","height":"12cm"},
  {"name":"photo-1","image":TP,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"tag",    "text":"AI PRODUCT","font":"Segoe UI","size":"11","bold":"true","color":"00D4FF","fill":"none","x":"1.6cm","y":"6.5cm","width":"12cm","height":"0.7cm"},
  {"name":"h-title","text":"Intelligence, Redefined","font":"Segoe UI","size":"52","bold":"true","color":"FFFFFF","fill":"none","x":"1.6cm","y":"7.4cm","width":"18cm","height":"5.5cm"},
  {"name":"h-sub",  "text":"The AI platform built for the way humans actually work.","font":"Segoe UI","size":"17","color":"7AB8FF","fill":"none","x":"1.6cm","y":"13.5cm","width":"16cm","height":"2cm"},
])

# ── S2 STATEMENT (060D24) ──
# orb-a 飞到左 (12cm位移!); 色调从蓝→粉主导
print("S2 statement")
add_slide(2,"060D24")
add_shapes(2,[
  {"name":"orb-a","preset":"ellipse","fill":"FF69B4","opacity":"0.28","x":"0cm", "y":"2cm", "width":"18cm","height":"17cm"},
  {"name":"orb-b","preset":"ellipse","fill":"00D4FF","opacity":"0.20","x":"2cm", "y":"5cm", "width":"13cm","height":"12cm"},
  {"name":"orb-c","preset":"ellipse","fill":"0066FF","opacity":"0.22","x":"3cm", "y":"0cm", "width":"15cm","height":"10cm"},
  {"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.15","x":"26cm","y":"10cm","width":"10cm","height":"9cm"},
  {"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.18","x":"22cm","y":"14cm","width":"12cm","height":"6cm"},
  {"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.05","x":"4cm", "y":"6cm", "width":"10cm","height":"9cm"},
  {"name":"photo-1","image":TP,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"tag",    "text":"","font":"Segoe UI","size":"11","color":"00D4FF","fill":"none","x":"15.5cm","y":"5cm","width":"4cm","height":"0.7cm"},
  {"name":"h-title","text":"Every decision, enhanced by AI.","font":"Segoe UI","size":"44","bold":"true","color":"FFFFFF","fill":"none","x":"15.5cm","y":"6cm","width":"17cm","height":"6cm"},
  {"name":"h-sub",  "text":"Not a tool. A thinking partner that scales with your ambition.","font":"Segoe UI","size":"17","color":"7AB8FF","fill":"none","x":"15.5cm","y":"12.5cm","width":"16cm","height":"3cm"},
])
morph(2)

# ── S3 PILLARS (0D1B4B) ──
# orb群退到背景右侧; cards从角落聚拢成三列 (最强聚合morph)
print("S3 pillars")
add_slide(3,"0D1B4B")
add_shapes(3,[
  {"name":"orb-a","preset":"ellipse","fill":"0066FF","opacity":"0.10","x":"18cm","y":"7cm","width":"16cm","height":"14cm"},
  {"name":"orb-b","preset":"ellipse","fill":"00D4FF","opacity":"0.08","x":"24cm","y":"0cm","width":"12cm","height":"10cm"},
  {"name":"orb-c","preset":"ellipse","fill":"FF69B4","opacity":"0.08","x":"26cm","y":"12cm","width":"10cm","height":"8cm"},
  {"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.07","x":"0cm", "y":"14cm","width":"10cm","height":"6cm"},
  {"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.08","x":"0cm", "y":"0cm", "width":"6cm","height":"5cm"},
  {"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.04","x":"12cm","y":"6cm","width":"20cm","height":"14cm"},
  {"name":"photo-1","image":TP,"opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  # cards 从角落飞出聚合成三列
  {"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.07","x":"1.6cm", "y":"4.4cm","width":"9.6cm","height":"11cm"},
  {"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.07","x":"12.4cm","y":"4.4cm","width":"9.6cm","height":"11cm"},
  {"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.07","x":"23.2cm","y":"4.4cm","width":"9.6cm","height":"11cm"},
  {"name":"tag",    "text":"THREE CAPABILITIES","font":"Segoe UI","size":"13","bold":"true","color":"FFFFFF","fill":"none","x":"1.6cm","y":"1cm","width":"20cm","height":"1.6cm"},
  {"name":"h-title","text":"Predict          Automate          Adapt","font":"Segoe UI","size":"22","bold":"true","color":"FFFFFF","fill":"none","x":"1.6cm","y":"7.5cm","width":"30cm","height":"2cm"},
  {"name":"h-sub",  "text":"Anticipate outcomes. Eliminate routine. Refine continuously — all without switching context.","font":"Segoe UI","size":"14","color":"7AB8FF","fill":"none","x":"1.6cm","y":"10cm","width":"30cm","height":"3cm"},
])
morph(3)

# ── S4 EVIDENCE (060D24) ──
# photo-1 从角落爆炸出现，覆盖左侧 (最高潮morph)
# orb群重新聚合在右侧围绕数据区
print("S4 evidence")
add_slide(4,"060D24")
add_shapes(4,[
  {"name":"orb-a","preset":"ellipse","fill":"00D4FF","opacity":"0.28","x":"2cm", "y":"2cm", "width":"16cm","height":"14cm"},
  {"name":"orb-b","preset":"ellipse","fill":"FF69B4","opacity":"0.20","x":"0cm", "y":"6cm", "width":"13cm","height":"13cm"},
  {"name":"orb-c","preset":"ellipse","fill":"0066FF","opacity":"0.22","x":"4cm", "y":"0cm", "width":"12cm","height":"10cm"},
  {"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.14","x":"22cm","y":"10cm","width":"12cm","height":"10cm"},
  {"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.16","x":"24cm","y":"4cm", "width":"10cm","height":"7cm"},
  {"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.04","x":"8cm", "y":"4cm", "width":"16cm","height":"13cm"},
  # photo-1 爆炸出现!
  {"name":"photo-1","image":TP,                   "x":"0cm","y":"0cm","width":"16cm","height":"19.05cm"},
  # cards收缩回角落
  {"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"tag",    "text":"RESULTS","font":"Segoe UI","size":"13","bold":"true","color":"FFFFFF","fill":"none","x":"17.6cm","y":"0.8cm","width":"14cm","height":"1cm"},
  {"name":"h-title","text":"10x","font":"Segoe UI","size":"72","bold":"true","color":"FFFFFF","fill":"none","x":"17.6cm","y":"2.4cm","width":"14cm","height":"5.5cm"},
  {"name":"h-sub",  "text":"Faster decisions\n\n$4.2M  Annual savings per client\n\n99.97%  Uptime SLA","font":"Segoe UI","size":"15","color":"7AB8FF","fill":"none","x":"17.6cm","y":"8.5cm","width":"14cm","height":"8cm"},
])
morph(4)

# ── S5 CTA (0D1B4B) ──
# orb群回归右侧呼应S1; photo-1保留为竖条
print("S5 CTA")
add_slide(5,"0D1B4B")
add_shapes(5,[
  {"name":"orb-a","preset":"ellipse","fill":"0066FF","opacity":"0.30","x":"18cm","y":"1cm","width":"18cm","height":"17cm"},
  {"name":"orb-b","preset":"ellipse","fill":"00D4FF","opacity":"0.22","x":"22cm","y":"4cm","width":"13cm","height":"13cm"},
  {"name":"orb-c","preset":"ellipse","fill":"FF69B4","opacity":"0.16","x":"28cm","y":"10cm","width":"8cm","height":"8cm"},
  {"name":"orb-d","preset":"ellipse","fill":"00FFC8","opacity":"0.14","x":"22cm","y":"0cm", "width":"10cm","height":"8cm"},
  {"name":"orb-e","preset":"ellipse","fill":"7B2FFF","opacity":"0.18","x":"0cm", "y":"14cm","width":"8cm","height":"6cm"},
  {"name":"orb-f","preset":"ellipse","fill":"FFFFFF","opacity":"0.05","x":"20cm","y":"6cm","width":"14cm","height":"12cm"},
  {"name":"photo-1","image":TP,                    "x":"20cm","y":"0cm","width":"10cm","height":"19.05cm"},
  {"name":"card-1","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-2","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"card-3","preset":"roundRect","fill":"FFFFFF","opacity":"0.01","x":"33cm","y":"18.5cm","width":"0.5cm","height":"0.5cm"},
  {"name":"tag",    "text":"GET STARTED","font":"Segoe UI","size":"11","bold":"true","color":"00D4FF","fill":"none","x":"1.6cm","y":"5.5cm","width":"14cm","height":"0.7cm"},
  {"name":"h-title","text":"Request Early Access","font":"Segoe UI","size":"46","bold":"true","color":"FFFFFF","fill":"none","x":"1.6cm","y":"6.4cm","width":"17cm","height":"5.5cm"},
  {"name":"h-sub",  "text":"Join 200+ enterprises already building with us.","font":"Segoe UI","size":"17","color":"7AB8FF","fill":"none","x":"1.6cm","y":"12.5cm","width":"16cm","height":"2cm"},
  {"name":"cta",    "text":"Get started ->","font":"Segoe UI","size":"15","bold":"true","color":"FFFFFF","fill":"0066FF","x":"1.6cm","y":"15.4cm","width":"9cm","height":"1.8cm"},
])
morph(5)

r=subprocess.run(["officecli","validate",PPTX],capture_output=True,text=True)
print("validate:","PASS" if "no errors" in r.stdout else r.stdout[:100])
r2=subprocess.run(["officecli","view",PPTX,"outline"],capture_output=True,text=True)
print(r2.stdout)
