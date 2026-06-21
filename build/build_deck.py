# -*- coding: utf-8 -*-
"""Build Sentinel Capital × Solafune pitch deck.
Design tokens replicated from template/identity_academy_vc課題_A班.pptx.
Grammar per slide: eyebrow(=タイトルのkicker) + Title + KeyMessage band + Body cards.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

HERE = os.path.dirname(os.path.abspath(__file__))
BG_COVER = os.path.join(HERE, "assets", "image1.png")
BG_BODY  = os.path.join(HERE, "assets", "image13.png")
OUT = os.path.join(os.path.dirname(HERE), "outputs", "05_発表スライド.pptx")

# ---- palette ----
CY   = "5BC8FF"   # cyan accent / eyebrow
WH   = "EAF1FF"   # near-white title
MUT  = "9DB0D0"   # muted body
DIM  = "6E7FA6"   # dim grey
GOLD = "F5B454"   # gold accent
PUR  = "9C8CFF"   # purple accent
CARD = "111C3C"   # card fill
CARDLN = "2A3B66" # card border
KMFILL = "13234D" # key-message band fill
FONT = "Arial"
EAFONT = "Meiryo"  # JP fallback (Arial ea empty in template)

prs = Presentation()
prs.slide_width  = Inches(10)
prs.slide_height = Inches(5.625)
BLANK = prs.slide_layouts[6]

def rgb(h): return RGBColor.from_string(h)

def slide(bg=BG_BODY):
    s = prs.slides.add_slide(BLANK)
    pic = s.shapes.add_picture(bg, 0, 0, prs.slide_width, prs.slide_height)
    # send picture to back
    sp = pic._element
    sp.getparent().remove(sp)
    s.shapes._spTree.insert(2, sp)
    return s

def _set_font(run, size, color, bold, font=FONT):
    f = run.font
    f.size = Pt(size); f.bold = bold; f.name = font
    f.color.rgb = rgb(color)
    # set east-asian font
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
    ea.set('typeface', EAFONT)

def text(s, x, y, w, h, runs, size=12, color=MUT, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, line_spacing=1.0, font=FONT, wrap=True):
    """runs: str OR list of (txt,color,bold[,size]) tuples; or list of paragraphs (each a list of tuples)."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    tf.vertical_anchor = anchor
    # normalize to list of paragraphs
    if isinstance(runs, str):
        paras = [[(runs, color, bold, size)]]
    elif runs and isinstance(runs[0], tuple):
        paras = [runs]
    else:
        paras = runs
    for pi, para in enumerate(paras):
        p = tf.paragraphs[0] if pi==0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = line_spacing
        for tup in para:
            txt = tup[0]; c = tup[1] if len(tup)>1 and tup[1] else color
            b = tup[2] if len(tup)>2 and tup[2] is not None else bold
            sz = tup[3] if len(tup)>3 and tup[3] else size
            r = p.add_run(); r.text = txt; _set_font(r, sz, c, b, font)
    return tb

def _shadow(shape):
    spPr = shape._element.spPr
    el = spPr.makeelement(qn('a:effectLst'), {})
    sh = spPr.makeelement(qn('a:outerShdw'),
        {'blurRad':'90000','dist':'30000','dir':'5400000','rotWithShape':'0','algn':'tl'})
    clr = spPr.makeelement(qn('a:srgbClr'), {'val':'000000'})
    al = spPr.makeelement(qn('a:alpha'), {'val':'38000'})
    clr.append(al); sh.append(clr); el.append(sh); spPr.append(el)

def card(s, x, y, w, h, fill=CARD, line=CARDLN, radius=0.07, shadow=True, line_w=1.0):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = rgb(fill)
    if line:
        sh.line.color.rgb = rgb(line); sh.line.width = Pt(line_w)
    else:
        sh.line.fill.background()
    try: sh.adjustments[0] = radius
    except Exception: pass
    sh.shadow.inherit = False
    if shadow: _shadow(sh)
    if sh.has_text_frame:
        sh.text_frame.word_wrap = True
    return sh

def bar(s, x, y, w, h, color, radius=0.18):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = rgb(color); sh.line.fill.background()
    try: sh.adjustments[0]=radius
    except Exception: pass
    sh.shadow.inherit=False
    return sh

def rect(s, x, y, w, h, color):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = rgb(color); sh.line.fill.background()
    sh.shadow.inherit=False
    return sh

def header(s, kicker, title):
    """タイトル zone: cyan eyebrow + near-white title."""
    text(s, 0.6, 0.34, 8.8, 0.28, kicker, size=11, color=CY, bold=True)
    text(s, 0.6, 0.62, 8.8, 0.62, title, size=27, color=WH, bold=True)

def keymsg(s, runs, y=1.34, h=0.62):
    """キーメッセージ band: distinct full-width band, gold left bar + assertion."""
    card(s, 0.6, y, 8.8, h, fill=KMFILL, line="2C4A86", radius=0.10, shadow=False, line_w=1.0)
    rect(s, 0.6, y, 0.07, h, GOLD)
    if isinstance(runs, str): runs=[(runs, WH, True)]
    text(s, 0.86, y, 8.35, h, runs, size=13.5, color=WH, bold=True,
         anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.04)

def footer(s, n, src=None):
    text(s, 0.6, 5.31, 6.5, 0.24, src or "", size=8, color=DIM)
    text(s, 8.7, 5.31, 0.7, 0.24, f"{n:02d}", size=9, color=DIM, align=PP_ALIGN.RIGHT)

def metric(s, x, y, w, h, big, big_c, label):
    card(s, x, y, w, h)
    text(s, x+0.18, y+0.16, w-0.36, 0.62, big, size=30, color=big_c, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.18, y+h-0.42, w-0.36, 0.34, label, size=11, color=MUT, anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================
# 1. COVER
# =====================================================================
s = slide(BG_COVER)
text(s, 0.6, 0.55, 8.0, 0.30, "VC FUND PITCH", size=12, color=CY, bold=True)
text(s, 0.6, 1.30, 9.0, 0.95, "SENTINEL CAPITAL", size=54, color=WH, bold=True)
text(s, 0.6, 2.18, 9.0, 0.80, "PARTNERS", size=54, color=CY, bold=True)
text(s, 0.6, 3.22, 8.8, 0.45, "宇宙から地球を読み、経済安全保障に張る。", size=19, color=WH, bold=False)
text(s, 0.6, 3.70, 8.8, 0.35, "Securing the critical frontier — from orbit.", size=13, color=DIM)
text(s, 0.6, 4.95, 9.0, 0.35,
     [("投資対象  ", DIM, False, 13), ("Solafune", WH, True, 13),
      ("       3号ファンド  ", DIM, False, 13), ("¥150億", GOLD, True, 13)])
text(s, 6.7, 0.56, 2.7, 0.30, "Sentinel Capital Partners｜センチネル・キャピタル", size=8.5, color=DIM, align=PP_ALIGN.RIGHT)

# =====================================================================
# 2. THESIS / EDGE — 3 pillars
# =====================================================================
s = slide()
header(s, "OUR EDGE ・ 勝ち筋", "私たちは、どう勝つか")
keymsg(s, [("衛星を持たない解析レイヤーの1社に、", WH, True),
           ("ファンドの元手1本ぶんを集中して張る", GOLD, True),
           ("——狭く深く。", WH, True)])
pills = [
    ("主権", "SOVEREIGNTY", CY,  "国産・データ主権が不可欠な領域に、政府アンカーで張る"),
    ("集中", "CONCENTRATION", GOLD, "10〜12社に絞り、1社でファンドを返す設計に徹する"),
    ("宇宙×AI", "ORBIT × AI", PUR, "衛星×AIで地球を解析するソフト層に投資する"),
]
x=0.6
for jp,en,c,desc in pills:
    card(s, x, 2.20, 2.78, 2.60)
    rect(s, x, 2.20, 2.78, 0.09, c)
    text(s, x+0.22, 2.52, 2.34, 0.45, [(jp+"  ",WH,True,18),(en,c,True,11)], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.22, 3.20, 2.34, 1.40, desc, size=12.5, color=MUT, line_spacing=1.15)
    x+=2.93
footer(s, 2)

# =====================================================================
# 3. FUND DESIGN — 6 metrics + algebra + team
# =====================================================================
s = slide()
header(s, "FUND DESIGN ・ ファンド設計", "数字で語るファンドの骨格")
keymsg(s, [("150億・集中10社・初回5.5億で持分", WH, True), ("約15%", GOLD, True),
           ("、狙うはネット", WH, True), ("DPI 3.0X", CY, True)])
mw, mh, gx = 2.78, 1.18, 2.93
cells = [
    ("¥150億", CY,  "ファンド総額（3号）"),
    ("3.0X",   GOLD,"目標ネットDPI"),
    ("10–12社",CY,  "集中ポートフォリオ"),
    ("約15%",  GOLD,"エントリー持分（リード）"),
    ("2 / 20", CY,  "管理報酬 / 成功報酬"),
    ("10年",   GOLD,"運用期間"),
]
for i,(b,c,l) in enumerate(cells):
    col=i%3; row=i//3
    metric(s, 0.6+col*gx, 2.18+row*1.30, mw, mh, b, c, l)
# bottom takeaway
card(s, 0.6, 4.82, 8.8, 0.52, fill=KMFILL, line=None, radius=0.12, shadow=False)
text(s, 0.86, 4.82, 8.3, 0.52,
     [("算数：初回5.5億 ÷ 投資後36.5億（Pre 31＋5.5）＝ ", MUT, False, 12.5),
      ("約15%", GOLD, True, 12.5),
      ("。チームは ", MUT, False, 12.5), ("GP2＋アソシ2＋クリアランス保持", WH, True, 12.5),
      ("＝深く回せる件数の天井10〜12社。", MUT, False, 12.5)],
     anchor=MSO_ANCHOR.MIDDLE)
footer(s, 3)

# =====================================================================
# 4. POWER LAW / FUND RETURNER
# =====================================================================
s = slide()
header(s, "HOW VC WINS ・ べき乗則", "1社で元手を返す“ファンドリターナー”")
keymsg(s, [("ネット3.0Xは“全社そこそこ”では作れない——", WH, True),
           ("少数のアウトライヤーを必ず1社仕込む", GOLD, True)])
# power-law bars (left)
card(s, 0.6, 2.18, 4.55, 2.62)
text(s, 0.85, 2.34, 4.05, 0.30, "10社のリターン分布（イメージ）", size=11.5, color=CY, bold=True)
heights = [(1.55,GOLD,"FR"),(1.05,CY,"7-8X"),(0.7,MUT,""),(0.62,MUT,""),(0.4,DIM,""),(0.32,DIM,""),(0.24,DIM,""),(0.16,DIM,""),(0.12,DIM,""),(0.08,DIM,"")]
bx=0.95; bw=0.30; base=4.55
for hgt,c,lab in heights:
    bar(s, bx, base-hgt, bw, hgt, c)
    if lab: text(s, bx-0.07, base-hgt-0.24, bw+0.4, 0.22, lab, size=8.5, color=c, bold=True)
    bx+=0.40
text(s, 0.85, 4.58, 4.05, 0.24, "1〜2社の大当たりが回収の大半を作る", size=10.5, color=MUT)
# FR 4 conditions (right)
card(s, 5.35, 2.18, 4.05, 2.62)
text(s, 5.58, 2.34, 3.6, 0.30, "ファンドリターナーの4条件", size=11.5, color=CY, bold=True)
conds = [
    ("①", "1社で150億超のポジションが射程（高い天井）"),
    ("②", "シード〜Aでリードしエグジット持分10〜12%確保"),
    ("③", "日本のIPOか国内プライム・商社M&Aで回収"),
    ("④", "センサー非依存の解析レイヤーという自社エッジ"),
]
yy=2.70
for num,desc in conds:
    text(s, 5.58, yy, 0.35, 0.40, num, size=13.5, color=GOLD, bold=True)
    text(s, 5.95, yy, 3.30, 0.46, desc, size=11, color=MUT, line_spacing=1.04)
    yy+=0.46
text(s, 5.58, yy+0.00, 3.65, 0.30, [("Solafune＝", MUT, False, 11.5),("4条件すべてに高適合",WH,True,11.5)])
footer(s, 4)

# =====================================================================
# 5. WHY NOW
# =====================================================================
s = slide()
header(s, "WHY NOW ・ なぜ今", "経済安全保障 × 宇宙の時代")
keymsg(s, [("重要鉱物をめぐる主権争いが、", WH, True),
           ("国産インテリジェンス需要を構造的に押し上げる", GOLD, True)])
# left narrative
card(s, 0.6, 2.18, 4.55, 2.62)
text(s, 0.88, 2.36, 4.0, 0.34, "重要鉱物をめぐる主権争い", size=15, color=WH, bold=True)
text(s, 0.88, 2.86, 4.05, 1.9,
     [[("● 中国がレアアース輸出管理を強化、アフリカ等で鉱床を囲い込み", MUT, False, 12)],
      [("● サプライチェーン途絶リスクが現実化（磁石・半導体）", MUT, False, 12)],
      [("→ 国産の資源モニタリング／解析需要が構造拡大", CY, True, 12)]],
     line_spacing=1.25)
# right demand pool
text(s, 5.55, 2.18, 4.0, 0.30, "追い風となる公的需要プール", size=13, color=WH, bold=True)
pool = [("¥8.8兆", CY, "防衛関係費（2026年度）"),
        ("¥1兆", GOLD, "宇宙戦略基金（10年）"),
        ("数千億円", PUR, "経済安保 強靱化予算")]
yy=2.62
for b,c,l in pool:
    card(s, 5.55, yy, 3.85, 0.66, radius=0.16)
    text(s, 5.78, yy, 1.75, 0.66, b, size=24, color=c, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 7.4, yy, 1.9, 0.66, l, size=11, color=MUT, anchor=MSO_ANCHOR.MIDDLE)
    yy+=0.76
# institutional changes strip (2026)
card(s, 0.6, 4.84, 8.8, 0.37, fill=KMFILL, line=None, radius=0.16, shadow=False)
text(s, 0.82, 4.84, 8.45, 0.37,
     [("2026年の制度変化　", GOLD, True, 10), ("① 輸出緩和（同志国市場へ）　② ファストパス調達 約3.5か月　③ 前払い導入＝死の谷を浅く", WH, False, 10)],
     anchor=MSO_ANCHOR.MIDDLE)
footer(s, 5, "出典：防衛省・宇宙戦略基金／Morning Pitch（01前提パック §4・§6）")

# =====================================================================
# 6. THE TARGET — Solafune
# =====================================================================
s = slide()
header(s, "THE TARGET ・ 投資対象", "Solafune ― 地球を読むOS")
keymsg(s, [("衛星画像・公開情報・通信という様々な“目”を束ねるAI＝", WH, True),
           ("Planetary Intelligence OS", CY, True)])
# 3 sources -> OS -> uses
srcs=[("GEOINT","衛星・地理空間"),("OSINT","公開情報・Web"),("SIGINT","通信・電波")]
yy=2.30
for t,d in srcs:
    card(s, 0.6, yy, 2.2, 0.72, radius=0.14)
    text(s, 0.8, yy+0.06, 1.9, 0.32, t, size=14, color=WH, bold=True)
    text(s, 0.8, yy+0.40, 1.9, 0.26, d, size=9.5, color=MUT)
    yy+=0.83
card(s, 3.15, 2.45, 1.95, 1.95, fill="0E2A52", line=CY, radius=0.5)
text(s, 3.15, 2.70, 1.95, 0.40, "Intelligence", size=13, color=WH, bold=True, align=PP_ALIGN.CENTER)
text(s, 3.15, 3.10, 1.95, 0.40, "OS", size=20, color=CY, bold=True, align=PP_ALIGN.CENTER)
uses=["重要鉱物の探査・違法採掘モニタリング（SHIGEN AI）",
      "防衛省・警察庁・内閣府 ＋ アフリカ諸国・国連機関",
      "120カ国超の開発者コミュニティが解析を供給"]
yy=2.30
for u in uses:
    card(s, 5.45, yy, 3.95, 0.72, radius=0.14)
    text(s, 5.68, yy, 3.55, 0.72, [("▸ ",CY,True,11.5),(u,MUT,False,11.5)], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    yy+=0.83
footer(s, 6)

# =====================================================================
# 7. MOAT — value chain
# =====================================================================
s = slide()
header(s, "MOAT ・ 競争優位", "なぜ Solafune が勝てるか")
keymsg(s, [("各社と競合せず“その上”の解析レイヤーに乗る——", WH, True),
           ("資本効率・ネットワーク効果・主権的信認", GOLD, True)])
chain=[("上流","衛星を作る・飛ばす","三菱電機 等",MUT),
       ("中流","データを撮る・売る","Synspective / ICEYE",MUT),
       ("下流","解析 → インテリジェンス","★ Solafune",GOLD)]
x=0.6
for t,d,who,c in chain:
    card(s, x, 2.16, 2.78, 1.18)
    text(s, x+0.2, 2.28, 2.4, 0.32, t, size=15, color=(GOLD if c==GOLD else MUT), bold=True)
    text(s, x+0.2, 2.64, 2.4, 0.30, d, size=10.5, color=MUT)
    text(s, x+0.2, 2.96, 2.4, 0.30, who, size=11.5, color=(GOLD if c==GOLD else WH), bold=True)
    x+=2.93
moats=[("センサー非依存","衛星を持たず複数ソースを統合する解析層＝高い資本効率"),
       ("ネットワーク効果","120カ国の開発者がデータ・人材・アルゴリズムを循環供給"),
       ("主権的信認","政府・国際機関の実績が参入障壁の高い領域での堀になる")]
x=0.6
for t,d in moats:
    card(s, x, 3.50, 2.78, 1.55)
    text(s, x+0.2, 3.66, 2.4, 0.34, [("◆ ",CY,True,12),(t,WH,True,13)])
    text(s, x+0.2, 4.10, 2.42, 0.88, d, size=10.5, color=MUT, line_spacing=1.12)
    x+=2.93
footer(s, 7)

# =====================================================================
# 8. MARKET — bottom-up SOM
# =====================================================================
s = slide()
header(s, "MARKET ・ 市場規模（SOM）", "“誰に・いくらで・何件”で積む")
keymsg(s, [("2032年SOM ", WH, True), ("約101億円", GOLD, True),
           (" は、Exit逆算の必要売上 約88億と一致＝市場と財務が一本の線", WH, True)])
segs=[("A 日本政府","12件×2.5億",30,CY),
      ("B 宇宙基金","5件×3億",15,PUR),
      ("C 同志国MoU","8×50%×1.5億",6,MUT),
      ("D 経済安保","SHIGEN AI",14,GOLD),
      ("E 民間","資源・インフラ監視",20,CY)]
# stacked horizontal bar of 2030=85
text(s, 0.6, 2.16, 4.2, 0.28, "2030年 SOM 積み上げ ＝ 85億円 / 38件", size=11.5, color=CY, bold=True)
total=sum(v for *_,v,_ in [(a,b,v,c) for a,b,v,c in segs])
x=0.6; full=8.8; scale=full/85.0
for name,unit,v,c in segs:
    w=v*scale
    bar(s, x, 2.50, w-0.04, 0.46, c, radius=0.10)
    text(s, x, 2.50, w-0.04, 0.46, f"{v}", size=12, color="0B1220", bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x+=w
# legend cards
x=0.6; cw=1.70
for name,unit,v,c in segs:
    card(s, x, 3.18, cw-0.06, 1.05)
    rect(s, x, 3.18, cw-0.06, 0.07, c)
    text(s, x+0.12, 3.32, cw-0.26, 0.34, name, size=11, color=WH, bold=True)
    text(s, x+0.12, 3.66, cw-0.26, 0.30, unit, size=8.7, color=MUT, line_spacing=1.0)
    text(s, x+0.12, 3.96, cw-0.26, 0.24, f"¥{v}億", size=12.5, color=c, bold=True)
    x+=cw
card(s, 0.6, 4.40, 8.8, 0.62, fill=KMFILL, line=None, radius=0.12, shadow=False)
text(s, 0.86, 4.40, 8.3, 0.62,
     [("2030年 ", MUT, False, 12), ("85億/38件", WH, True, 12),
      (" → 2032年 ", MUT, False, 12), ("101億/46件", GOLD, True, 12),
      ("。シェア仮定なし、単価×件数の積み上げ［仮・要検証／C転換率が最も要実証］", MUT, False, 11)],
     anchor=MSO_ANCHOR.MIDDLE)
footer(s, 8, "出典：01前提パック §2-3（ボトムアップ積算）")

# =====================================================================
# 9. FINANCIALS
# =====================================================================
s = slide()
header(s, "QUANT ・ 定量モデル", "5年で描く成長（仮・要検証）")
keymsg(s, [("売上 4億→85億（CAGR約113%）×粗利 ", WH, True), ("60→72%", GOLD, True),
           ("。KPI閾値クリアで段階投資", WH, True)])
# revenue bars
card(s, 0.6, 2.16, 5.7, 2.88)
text(s, 0.85, 2.30, 5.2, 0.28, "売上高の推移（億円）", size=11.5, color=CY, bold=True)
revs=[("26",4),("27",12),("28",30),("29",55),("30",85)]
maxr=85; bx=1.05; bw=0.72; base=4.62; gap=0.32
for yr,v in revs:
    h=(v/maxr)*1.9
    bar(s, bx, base-h, bw, h, CY if yr!="30" else GOLD)
    text(s, bx-0.1, base-h-0.24, bw+0.2, 0.22, str(v), size=10, color=WH, bold=True, align=PP_ALIGN.CENTER)
    text(s, bx-0.1, base+0.04, bw+0.2, 0.22, "'"+yr, size=9.5, color=MUT, align=PP_ALIGN.CENTER)
    bx+=bw+gap
text(s, 0.85, 4.74, 5.2, 0.24, "日本政府 → 同志国 → 民間（グローバルサウス）の段階展開", size=10, color=MUT)
# right metrics
rights=[("60% → 72%",GOLD,"グロスマージン（ソフト/解析）"),
        ("¥31→36.5億",CY,"投資前Pre→投資後Post-Val"),
        ("年率 約2倍",PUR,"政府リカーリングのフォローKPI")]
yy=2.16
for b,c,l in rights:
    card(s, 6.55, yy, 2.85, 0.90)
    text(s, 6.75, yy+0.08, 2.5, 0.44, b, size=19, color=c, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 6.77, yy+0.54, 2.5, 0.30, l, size=9.7, color=MUT)
    yy+=0.99
footer(s, 9, "出典：02財務モデル.xlsx（PL/前提シート）")

# =====================================================================
# 10. VALUATION & EXIT — waterfall
# =====================================================================
s = slide()
header(s, "EXIT ・ バリュエーション", "1社でファンドを返す")
keymsg(s, [("Exit売上101億 × ", WH, True), ("EV/Sales 17倍", CY, True),
           (" → 企業価値 ", WH, True), ("約1,717億円", GOLD, True),
           ("（持分12%で206億）", WH, True)])
steps=[("売上 ¥101億","2032E",CY),("× 17倍","EV/Sales（保守15–18）",MUT),
       ("¥1,717億","Exit企業価値",GOLD),("× 12%","Exit持分",MUT),("¥206億","ポジション",PUR)]
x=0.6; cw=1.62; gap=0.18
for i,(b,l,c) in enumerate(steps):
    card(s, x, 2.16, cw, 1.12)
    text(s, x+0.1, 2.30, cw-0.2, 0.50, b, size=(15 if "¥" in b and "億" in b else 16), color=c, bold=True, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    text(s, x+0.1, 2.84, cw-0.2, 0.34, l, size=9.3, color=MUT, align=PP_ALIGN.CENTER, line_spacing=1.0)
    if i<len(steps)-1:
        text(s, x+cw-0.02, 2.16, gap+0.1, 1.12, "▸", size=14, color=DIM, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    x+=cw+gap
# EV/Sales derivation
card(s, 0.6, 3.50, 5.55, 1.55)
text(s, 0.85, 3.64, 5.0, 0.30, "EV/Sales 17倍の統計的導出", size=11.5, color=CY, bold=True)
text(s, 0.85, 4.00, 5.35, 1.0,
     [[("“国しか客になれない国産インフラ”＝普通のソフトの約1.7倍の値札", CY, False, 10.5)],
      [("comps中央値10倍 ＋ ソフト+3／粗利+2／主権+2 ＝ ", MUT, False, 10.5), ("17倍", GOLD, True, 10.5)],
      [("Palantirプレ20・現況60.7倍はバブルとして除外", DIM, False, 10)]],
     line_spacing=1.18)
# fund returner box
card(s, 6.35, 3.50, 3.05, 1.55, fill=KMFILL, line=GOLD)
text(s, 6.58, 3.66, 2.6, 0.30, "ファンドリターナー成立", size=11.5, color=GOLD, bold=True)
text(s, 6.58, 4.02, 2.6, 0.5, [("¥206億",WH,True,22),(" ÷ ¥150億",MUT,False,12)], anchor=MSO_ANCHOR.MIDDLE)
text(s, 6.58, 4.56, 2.7, 0.40, [("＝ 対ファンド ",MUT,False,11),("1.37X",CY,True,13),("／総コミMOIC ",MUT,False,11),("12.5X",GOLD,True,13)], line_spacing=1.0)
footer(s, 10, "出典：01前提パック §5・02財務モデル Exit/リターン")

# =====================================================================
# 11. PORTFOLIO — power-law to net 3.0X
# =====================================================================
s = slide()
header(s, "PORTFOLIO ・ ポートフォリオ", "ネット3.0Xは分布で作る")
keymsg(s, [("実在2本（Solafune＋Humanity Brain）に厚く張り → ", WH, True),
           ("グロス4.3X → ネットDPI 3.0X", GOLD, True),
           ("（上位2社で回収の約65%）", WH, True)])
# distribution table-ish cards
rows=[("Solafune","★ FR（本命）",206,GOLD),("Humanity Brain","ニセ情報耐性AI・7-8X｜実在",136,CY),
      ("C・D","3-5X級2社",119,MUT),("E-G","1-2X級3社",43,DIM),
      ("H-J","0-1X級3社",20,DIM)]
x=0.6; cw=1.70
for name,tier,rec,c in rows:
    card(s, x, 2.16, cw-0.06, 1.30)
    rect(s, x, 2.16, cw-0.06, 0.07, c)
    text(s, x+0.12, 2.30, cw-0.26, 0.30, name, size=11.5, color=WH, bold=True)
    text(s, x+0.12, 2.62, cw-0.26, 0.26, tier, size=8.7, color=MUT)
    text(s, x+0.12, 2.92, cw-0.26, 0.44, f"¥{rec}億", size=16, color=c, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    x+=cw
# bottom math row (3 metric)
trio=[("¥123億","投下資本（10社）",CY),("¥524億 / 4.3X","グロス回収",GOLD),("3.0X","ネットDPI（検算一致）",PUR)]
x=0.6
for b,l,c in trio:
    card(s, x, 3.66, 2.88, 0.90)
    text(s, x+0.18, 3.74, 2.5, 0.46, b, size=18, color=c, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.18, 4.20, 2.5, 0.30, l, size=10, color=MUT)
    x+=2.93
text(s, 0.6, 4.66, 8.8, 0.40,
     [("DPI＝財布全体が何倍／MOIC＝この1社が何倍。", MUT, False, 11),
      ("理想の薄い裾より、検証可能な実弾2本に厚く張る集中設計。", WH, True, 11)])
footer(s, 11, "出典：02財務モデル「ポートフォリオ」シート（net DPI 2.995≈3.0X）")

# =====================================================================
# 12. RISK
# =====================================================================
s = slide()
header(s, "RISK ・ リスクと備え", "弱点を、先に潰す")
keymsg(s, [("出口・政府依存・評価・バブルの4論点に、", WH, True),
           ("契約と設計レベルの具体策を当てる", GOLD, True)])
risks=[("出口が未実証","上場済み宇宙SU（Synspective等）が実証。日本IPO波に乗る"),
       ("政府依存・調達変動","マイルストン連動＋同志国横展開でリカーリング化"),
       ("バリュエーション非開示","トランシェ払込・希薄化防止・取締役指名権を確保"),
       ("防衛AIバブル懸念","保守的マルチプル(15–18倍)とExit時期の柔軟化で吸収")]
pos=[(0.6,2.18),(5.1,2.18),(0.6,3.62),(5.1,3.62)]
for (px,py),(t,d) in zip(pos,risks):
    card(s, px, py, 4.30, 1.34)
    text(s, px+0.24, py+0.16, 3.9, 0.36, [("⚠ ",GOLD,True,13),(t,WH,True,14)])
    text(s, px+0.24, py+0.62, 3.85, 0.62, d, size=11.5, color=MUT, line_spacing=1.12)
footer(s, 12)

# =====================================================================
# 13. CLOSING
# =====================================================================
s = slide(BG_COVER)
text(s, 0.6, 0.70, 9.0, 0.35, "SENTINEL CAPITAL PARTNERS  ×  Solafune", size=13, color=CY, bold=True)
text(s, 0.6, 1.55, 9.0, 0.80, "地球を読む者が、", size=40, color=WH, bold=True)
text(s, 0.6, 2.40, 9.0, 0.80, "次の安全保障を制する。", size=40, color=CY, bold=True)
text(s, 0.6, 3.45, 9.0, 0.40, "経済安全保障 × 宇宙 × AI に、狭く深く張る。", size=16, color=MUT)
# decision strip
card(s, 0.6, 4.02, 8.8, 0.74, fill=KMFILL, line=GOLD, radius=0.10, shadow=False)
text(s, 0.86, 4.02, 8.3, 0.74,
     [("投資判断：", GOLD, True, 12.5), ("初回5.5億・持分約15%でリード", WH, True, 12.5),
      ("／Exit 206億・対ファンド1.37X・ネットDPI 3.0X ＝ ", MUT, False, 11.5),
      ("ファンドリターナー候補", CY, True, 12.5)],
     anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.6, 4.98, 9.0, 0.32, "ご清聴ありがとうございました", size=11.5, color=DIM)

prs.save(OUT)
print("saved:", OUT, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))
