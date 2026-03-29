import json, subprocess, sys, copy

BASE = "/Users/visher/Library/Application Support/AionUI/aionui/claude-temp-1774010336920/ppt/v4"
PPTX = f"{BASE}/brand-refresh.pptx"
P1   = f"{BASE}/assets/portrait1.jpg"
P2   = f"{BASE}/assets/portrait2.jpg"
AB   = f"{BASE}/assets/abstract1.jpg"
TM   = f"{BASE}/assets/team1.jpg"

def batch(ops):
    r = subprocess.run(["officecli","batch",PPTX], input=json.dumps(ops), text=True, capture_output=True)
    line = [l for l in r.stdout.splitlines() if "complete" in l]
    print(" ", line[0] if line else r.stdout[:120])
    if r.returncode != 0: print("  ERR:", r.stderr[:150])

def add_slide(n, bg):
    batch([{"command":"add","parent":"/","type":"slide","props":{"layout":"blank","background":bg}}])

def add_shapes(n, shapes):
    ops = []
    for s in shapes:
        props = copy.deepcopy(s)
        ops.append({"command":"add","parent":f"/slide[{n}]","type":"shape","props":props})
    batch(ops)

def set_morph(n):
    batch([{"command":"set","path":f"/slide[{n}]","props":{"transition":"morph"}}])

subprocess.run(["officecli","create",PPTX], capture_output=True)
print("=== brand-refresh v4 ===")

# ── S1 HERO (cream F5F0E8) ──────────────────────────────────────────────────
# 布局：左侧大标题 | 右侧拼贴马赛克(photo-1 + 6个色块)
# photo-2 隐藏在右下角等待被morph激活
print("S1 hero")
add_slide(1,"F5F0E8")
add_shapes(1,[
  {"name":"photo-1","image":P1,              "x":"15.5cm","y":"0cm",    "width":"10cm",   "height":"13cm"},
  {"name":"blk-a",  "fill":"162040",          "x":"25.5cm","y":"0cm",    "width":"8.37cm", "height":"7cm"},
  {"name":"blk-b",  "fill":"1A6BFF",          "x":"25.5cm","y":"7cm",    "width":"4cm",    "height":"6cm"},
  {"name":"blk-c",  "fill":"F4713A",          "x":"29.5cm","y":"7cm",    "width":"4.37cm", "height":"6cm"},
  {"name":"blk-d",  "fill":"00C9D4",          "x":"15.5cm","y":"13cm",   "width":"5cm",    "height":"6.05cm"},
  {"name":"blk-e",  "fill":"7EC8A0",          "x":"20.5cm","y":"13cm",   "width":"5cm",    "height":"6.05cm"},
  {"name":"blk-f",  "fill":"E8749A",          "x":"25.5cm","y":"13cm",   "width":"8.37cm", "height":"6.05cm"},
  {"name":"photo-2","image":P2,"opacity":"0.01","x":"33cm","y":"18.55cm","width":"0.5cm",  "height":"0.5cm"},
  {"name":"tag",    "text":"BRAND REFRESH 2025","font":"Arial","size":"11","bold":"true","color":"9A9080","fill":"none","x":"1.6cm","y":"7cm",    "width":"13cm","height":"0.7cm"},
  {"name":"h-title","text":"Your Brand, Redefined.","font":"Arial","size":"52","bold":"true","color":"162040","fill":"none","x":"1.6cm","y":"7.8cm","width":"13cm","height":"5.5cm"},
  {"name":"h-sub",  "text":"A new visual language built for how the world sees you now.","font":"Arial","size":"15","color":"6B6355","fill":"none","x":"1.6cm","y":"14cm","width":"13cm","height":"2.5cm"},
])

# ── S2 STATEMENT (dark 162040) ─────────────────────────────────────────────
# photo-1: 右侧→整个左半屏 (巨大位移+放大)
# blk-a 覆盖在photo上做蒙版 (从右上角→同位置)
# blk-b~f: 从右侧拼贴→变成右侧等高水平条带 (完全重排)
print("S2 statement")
add_slide(2,"162040")
add_shapes(2,[
  {"name":"photo-1","image":P1,               "x":"0cm",   "y":"0cm",    "width":"14cm",   "height":"19.05cm"},
  {"name":"blk-a",  "fill":"162040","opacity":"0.58","x":"0cm","y":"0cm", "width":"14cm",   "height":"19.05cm"},
  {"name":"blk-b",  "fill":"1A6BFF",           "x":"22cm",  "y":"0cm",    "width":"11.87cm","height":"3.2cm"},
  {"name":"blk-c",  "fill":"F4713A",           "x":"22cm",  "y":"3.2cm",  "width":"11.87cm","height":"3.2cm"},
  {"name":"blk-d",  "fill":"00C9D4",           "x":"22cm",  "y":"6.4cm",  "width":"11.87cm","height":"3.2cm"},
  {"name":"blk-e",  "fill":"7EC8A0",           "x":"22cm",  "y":"9.6cm",  "width":"11.87cm","height":"3.2cm"},
  {"name":"blk-f",  "fill":"E8749A",           "x":"22cm",  "y":"12.8cm", "width":"11.87cm","height":"6.25cm"},
  {"name":"photo-2","image":P2,"opacity":"0.01","x":"33cm", "y":"18.55cm","width":"0.5cm",  "height":"0.5cm"},
  {"name":"tag",    "text":"","font":"Arial","size":"11","color":"4A5A7A","fill":"none","x":"15.2cm","y":"5cm","width":"4cm","height":"0.7cm"},
  {"name":"h-title","text":"Clarity beats complexity.","font":"Arial","size":"46","bold":"true","color":"F5F0E8","fill":"none","x":"15.2cm","y":"6cm","width":"15.5cm","height":"7cm"},
  {"name":"h-sub",  "text":"The strongest brands say less — and mean more.","font":"Arial","size":"16","color":"7890B8","fill":"none","x":"15.2cm","y":"13.5cm","width":"15cm","height":"2.5cm"},
])
set_morph(2)

# ── S3 PILLARS (cream F5F0E8) ──────────────────────────────────────────────
# blk-a: 从left-full覆盖→压扁成顶栏 (最强视觉morph)
# photo-1: 从left-full→缩成col2图头 (同样震撼)
# photo-2: 从角落→飞出成col1图头 (出现感)
# blk-b~d 作为各列图片上的色彩蒙版; blk-e用team1图作col3
print("S3 pillars")
add_slide(3,"F5F0E8")
add_shapes(3,[
  {"name":"blk-a",  "fill":"162040",           "x":"0cm",   "y":"0cm",    "width":"33.87cm","height":"2.4cm"},
  {"name":"photo-2","image":P2,                "x":"1.6cm", "y":"2.4cm",  "width":"9.6cm",  "height":"8cm"},
  {"name":"photo-1","image":P1,                "x":"12.4cm","y":"2.4cm",  "width":"9.6cm",  "height":"8cm"},
  {"name":"blk-e",  "image":TM,                "x":"22.8cm","y":"2.4cm",  "width":"9.6cm",  "height":"8cm"},
  {"name":"blk-b",  "fill":"162040","opacity":"0.42","x":"1.6cm","y":"2.4cm","width":"9.6cm","height":"8cm"},
  {"name":"blk-c",  "fill":"F4713A","opacity":"0.38","x":"12.4cm","y":"2.4cm","width":"9.6cm","height":"8cm"},
  {"name":"blk-d",  "fill":"00C9D4","opacity":"0.38","x":"22.8cm","y":"2.4cm","width":"9.6cm","height":"8cm"},
  {"name":"blk-f",  "fill":"E8749A","opacity":"0.01","x":"33cm","y":"18.55cm","width":"0.5cm","height":"0.5cm"},
  {"name":"tag",    "text":"THREE PILLARS","font":"Arial","size":"13","bold":"true","color":"F5F0E8","fill":"none","x":"1.6cm","y":"0.5cm","width":"20cm","height":"1.4cm"},
  {"name":"h-title","text":"Identity                    Voice                    Experience","font":"Arial","size":"14","bold":"true","color":"162040","fill":"none","x":"1.6cm","y":"11cm","width":"31cm","height":"1.2cm"},
  {"name":"h-sub",  "text":"A system that speaks\nbefore words do.","font":"Arial","size":"14","color":"6B6355","fill":"none","x":"1.6cm","y":"12.4cm","width":"9.6cm","height":"3.5cm"},
])
set_morph(3)

# ── S4 EVIDENCE (cream F5F0E8) ────────────────────────────────────────────
# photo-1: 从col2图头→膨胀成左侧大面板(换成abstract纹理) — 变换+换图
# blk-b~f: 从色彩蒙版→变成彩色bezier波浪叠在纹理上
# blk-a: 顶栏保持，略变高度
print("S4 evidence")
add_slide(4,"F5F0E8")
add_shapes(4,[
  {"name":"blk-a",  "fill":"162040",           "x":"0cm","y":"0cm","width":"33.87cm","height":"2cm"},
  {"name":"photo-1","image":AB,                "x":"0cm","y":"2cm","width":"19cm","height":"17.05cm"},
  {"name":"blk-b",  "fill":"162040","opacity":"0.78","geometry":"M 0,52 C 22,36 44,66 64,46 C 80,30 92,56 100,42 L 100,100 L 0,100 Z","x":"0cm","y":"2cm","width":"19cm","height":"17.05cm"},
  {"name":"blk-c",  "fill":"1A6BFF","opacity":"0.72","geometry":"M 0,63 C 22,48 44,76 65,57 C 82,44 93,65 100,53 L 100,100 L 0,100 Z","x":"0cm","y":"2cm","width":"19cm","height":"17.05cm"},
  {"name":"blk-d",  "fill":"00C9D4","opacity":"0.68","geometry":"M 0,73 C 22,60 44,84 65,66 C 83,55 93,74 100,63 L 100,100 L 0,100 Z","x":"0cm","y":"2cm","width":"19cm","height":"17.05cm"},
  {"name":"blk-e",  "fill":"7EC8A0","opacity":"0.65","geometry":"M 0,82 C 24,70 46,90 66,75 C 83,65 93,82 100,72 L 100,100 L 0,100 Z","x":"0cm","y":"2cm","width":"19cm","height":"17.05cm"},
  {"name":"blk-f",  "fill":"F4713A","opacity":"0.68","geometry":"M 0,90 C 24,80 46,96 66,84 C 83,76 93,90 100,82 L 100,100 L 0,100 Z","x":"0cm","y":"2cm","width":"19cm","height":"17.05cm"},
  {"name":"photo-2","image":P2,"opacity":"0.01","x":"33cm","y":"18.55cm","width":"0.5cm","height":"0.5cm"},
  {"name":"tag",    "text":"THE NUMBERS","font":"Arial","size":"13","bold":"true","color":"9A9080","fill":"none","x":"20.4cm","y":"0.4cm","width":"12cm","height":"0.8cm"},
  {"name":"h-title","text":"+47%","font":"Arial","size":"64","bold":"true","color":"162040","fill":"none","x":"20.4cm","y":"2.5cm","width":"12cm","height":"5cm"},
  {"name":"h-sub",  "text":"Brand recognition lift after refresh\n\n2.8x  Engagement rate increase\n\n89    Net Promoter Score","font":"Arial","size":"14","color":"6B6355","fill":"none","x":"20.4cm","y":"8cm","width":"12cm","height":"8cm"},
])
set_morph(4)

# ── S5 CTA (dark 162040) ──────────────────────────────────────────────────
# photo-2: 从角落→飞出成右侧整列 (最强出场)
# photo-1: 退出到角落消失
# blk-a~f: 重新围框photo-2，与S1构成呼应
print("S5 CTA")
add_slide(5,"162040")
add_shapes(5,[
  {"name":"photo-2","image":P2,                "x":"21cm",   "y":"0cm",    "width":"9cm",    "height":"19.05cm"},
  {"name":"blk-a",  "fill":"162040","opacity":"0.75","x":"21cm","y":"0cm",  "width":"4cm",    "height":"5.5cm"},
  {"name":"blk-b",  "fill":"1A6BFF",           "x":"21cm",   "y":"5.5cm",  "width":"2.4cm",  "height":"4.5cm"},
  {"name":"blk-c",  "fill":"F4713A",           "x":"29.5cm", "y":"13.5cm", "width":"4.37cm", "height":"5.55cm"},
  {"name":"blk-d",  "fill":"00C9D4",           "x":"29.5cm", "y":"0cm",    "width":"4.37cm", "height":"5cm"},
  {"name":"blk-e",  "fill":"7EC8A0","opacity":"0.01","x":"33cm","y":"18.55cm","width":"0.5cm","height":"0.5cm"},
  {"name":"blk-f",  "fill":"E8749A","opacity":"0.01","x":"33cm","y":"18.55cm","width":"0.5cm","height":"0.5cm"},
  {"name":"photo-1","image":AB,"opacity":"0.01","x":"33cm",  "y":"18.55cm","width":"0.5cm",  "height":"0.5cm"},
  {"name":"tag",    "text":"BRAND STRATEGY","font":"Arial","size":"11","bold":"true","color":"4A5A7A","fill":"none","x":"1.6cm","y":"5.5cm","width":"14cm","height":"0.7cm"},
  {"name":"h-title","text":"Start the transformation.","font":"Arial","size":"46","bold":"true","color":"F5F0E8","fill":"none","x":"1.6cm","y":"6.4cm","width":"17cm","height":"6cm"},
  {"name":"h-sub",  "text":"Let's build something that lasts.","font":"Arial","size":"16","color":"7890B8","fill":"none","x":"1.6cm","y":"13.2cm","width":"16cm","height":"2cm"},
  {"name":"cta",    "text":"Get in touch  ->","font":"Arial","size":"15","bold":"true","color":"F5F0E8","fill":"F4713A","x":"1.6cm","y":"15.6cm","width":"9cm","height":"1.8cm"},
])
set_morph(5)

r = subprocess.run(["officecli","validate",PPTX],capture_output=True,text=True)
ok = "no errors" in r.stdout
print("validate:", "PASS" if ok else r.stdout[:150])
r2 = subprocess.run(["officecli","view",PPTX,"outline"],capture_output=True,text=True)
print(r2.stdout)
