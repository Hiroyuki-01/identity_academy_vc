#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""完成版スライド（outputs/05_発表スライド.pdf）を、縦スクロールで見られる
ホームページ（outputs/web/index.html）に変換する。グラフは画像なので“そのまま”。

使い方:  python build/build_web.py
前提:    PyMuPDF（fitz）。未導入なら  python -m pip install PyMuPDF
出力:    outputs/web/index.html ＋ outputs/web/assets/slide-NN.png
"""
import shutil, pathlib, datetime, html
import fitz  # PyMuPDF

ROOT = pathlib.Path(__file__).resolve().parent.parent
PDF = ROOT / "outputs" / "05_発表スライド.pdf"
# index.html は手作りの宇宙デザインHP。こちらは原本スライドの画像ビュー（slides.html）。
WEB = ROOT / "outputs" / "web"
ASSETS = WEB / "assets"
OUT = WEB / "slides.html"
ZOOM = 2.0  # 72dpi × 2 = 144dpi（くっきり）


# 目次ラベル＝完成版デッキ（build_deck.py）の各スライド見出し。デッキは固定の
# 提出物なのでここを正とする。枚数が変わった場合は page_title() にフォールバック。
LABELS = [
    "表紙｜Sentinel Capital × Solafune",
    "勝ち筋｜私たちは、どう勝つか",
    "ファンド設計｜数字で語る骨格",
    "べき乗則｜1社で元手を返すFR",
    "なぜ今｜経済安全保障 × 宇宙",
    "投資対象｜Solafune＝地球を読むOS",
    "競争優位｜なぜ Solafune が勝てるか",
    "市場（SOM）｜誰に・いくらで・何件",
    "定量モデル｜5年の成長と粗利",
    "Exit｜1社でファンドを返す",
    "ポートフォリオ｜ネット3.0Xは分布で",
    "リスク｜弱点を、先に潰す",
    "まとめ｜投資判断",
]


def page_title(page) -> str:
    """フォールバック：ページ内で最も上にあるテキスト（＝見出し）を拾う。"""
    best_txt, best_y = "", 1e9
    for b in page.get_text("dict").get("blocks", []):
        for line in b.get("lines", []):
            for sp in line.get("spans", []):
                t = sp["text"].strip()
                if t and sp["bbox"][1] < best_y:
                    best_txt, best_y = t, sp["bbox"][1]
    return best_txt


def render():
    doc = fitz.open(PDF)
    if ASSETS.exists():
        shutil.rmtree(ASSETS)
    ASSETS.mkdir(parents=True, exist_ok=True)

    # 宇宙デザインHP（index.html）が使う背景もここで配置（rmtreeで消えるため毎回コピー）
    bg = ROOT / "build" / "assets"
    shutil.copy(bg / "image1.png", ASSETS / "hero-bg.png")
    shutil.copy(bg / "image13.png", ASSETS / "body-bg.png")

    sections, toc = [], []
    mat = fitz.Matrix(ZOOM, ZOOM)
    for i, page in enumerate(doc):
        n = i + 1
        fname = f"slide-{n:02d}.png"
        page.get_pixmap(matrix=mat).save(ASSETS / fname)
        label = LABELS[i] if i < len(LABELS) else (page_title(page) or f"Slide {n}")
        toc.append(
            f'<a class="toc-link" href="#s{n}" data-target="s{n}">'
            f'<span class="toc-no">{n:02d}</span>'
            f'<span class="toc-text">{html.escape(label)}</span></a>'
        )
        sections.append(
            f'<section id="s{n}" class="section">'
            f'<div class="sec-head"><span class="sec-no">{n:02d}</span>'
            f'<span class="sec-title">{html.escape(label)}</span></div>'
            f'<img loading="lazy" src="assets/{fname}" alt="スライド{n}：{html.escape(label)}">'
            f"</section>"
        )

    built = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    htmldoc = (
        TEMPLATE.replace("{{SECTIONS}}", "\n".join(sections))
        .replace("{{TOC}}", "\n".join(toc))
        .replace("{{COUNT}}", str(doc.page_count))
        .replace("{{BUILT}}", built)
    )
    OUT.write_text(htmldoc, encoding="utf-8")
    print(f"OK: {doc.page_count} slides -> {OUT}")


TEMPLATE = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Solafune 投資ピッチ — プレビュー</title>
<style>
:root{ --bg:#0b1020; --card:#121a31; --ink:#eaf0ff; --sub:#9fb0d6; --line:#26324f;
  --accent:#5b8cff; }
*{box-sizing:border-box}
html{scroll-behavior:smooth;scroll-padding-top:64px}
body{margin:0;background:var(--bg);color:var(--ink);
  font-family:"Segoe UI","Yu Gothic UI","Hiragino Kaku Gothic ProN",Meiryo,sans-serif;}
header#top{position:fixed;top:0;left:0;right:0;height:52px;display:flex;align-items:center;
  gap:14px;padding:0 20px;background:rgba(14,20,38,.92);backdrop-filter:blur(6px);
  border-bottom:1px solid var(--line);z-index:50}
header#top b{font-size:14px;letter-spacing:.02em}
header#top .meta{color:var(--sub);font-size:12px}
#prog{position:fixed;top:52px;left:0;height:3px;background:var(--accent);width:0;z-index:50}
.layout{max-width:1200px;margin:0 auto;display:grid;grid-template-columns:248px 1fr;
  gap:30px;padding:80px 22px 90px}
nav.toc{position:sticky;top:80px;align-self:start;max-height:calc(100vh - 110px);
  overflow:auto;border:1px solid var(--line);border-radius:14px;padding:10px;background:#101830}
.toc-title{font-size:11px;letter-spacing:.08em;color:var(--sub);padding:4px 8px 8px;text-transform:uppercase}
.toc-link{display:flex;gap:8px;align-items:baseline;padding:6px 8px;border-radius:8px;
  color:var(--sub);text-decoration:none;font-size:13px;line-height:1.35}
.toc-link:hover{background:#1a2746;color:var(--ink)}
.toc-link.active{background:#1a2746;color:#fff;box-shadow:inset 3px 0 0 var(--accent)}
.toc-no{font-size:11px;color:#6f82ad;min-width:20px}
.toc-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
main{min-width:0}
.section{margin-bottom:30px;scroll-margin-top:64px}
.sec-head{display:flex;align-items:baseline;gap:10px;margin-bottom:8px}
.sec-no{font-size:12px;color:var(--sub);border:1px solid var(--line);border-radius:6px;padding:1px 8px}
.sec-title{font-size:15px;color:#dce6ff;font-weight:700}
.section img{width:100%;height:auto;display:block;border:1px solid var(--line);
  border-radius:14px;box-shadow:0 14px 40px rgba(0,0,0,.4);background:#000}
#toTop{position:fixed;right:18px;bottom:18px;z-index:50;background:#1a2746;color:var(--ink);
  border:1px solid var(--line);border-radius:50%;width:44px;height:44px;font-size:18px;
  cursor:pointer;opacity:0;transition:opacity .2s}
#toTop.show{opacity:1}
@media(max-width:880px){.layout{grid-template-columns:1fr}nav.toc{display:none}}
</style>
</head>
<body>
<header id="top"><b>Solafune 投資ピッチ</b>
  <span class="meta">完成版スライド・縦スクロール（提出時はそのままスライド配布）　|　source: outputs/05_発表スライド.pdf　|　全{{COUNT}}枚　|　built {{BUILT}}</span></header>
<div id="prog"></div>
<div class="layout">
  <nav class="toc"><div class="toc-title">目次</div>
{{TOC}}
  </nav>
  <main>
{{SECTIONS}}
  </main>
</div>
<button id="toTop" title="トップへ">↑</button>
<script>
const prog=document.getElementById('prog');
const toTop=document.getElementById('toTop');
function onScroll(){
  const h=document.documentElement;
  const sc=h.scrollTop/((h.scrollHeight-h.clientHeight)||1);
  prog.style.width=(sc*100)+'%';
  toTop.classList.toggle('show',h.scrollTop>400);
}
document.addEventListener('scroll',onScroll,{passive:true});onScroll();
toTop.onclick=()=>window.scrollTo({top:0,behavior:'smooth'});
const links=[...document.querySelectorAll('.toc-link')];
const map=Object.fromEntries(links.map(l=>[l.dataset.target,l]));
const obs=new IntersectionObserver((es)=>{
  es.forEach(e=>{ if(e.isIntersecting){
    links.forEach(l=>l.classList.remove('active'));
    const a=map[e.target.id]; if(a){a.classList.add('active');a.scrollIntoView({block:'nearest'});}
  }});
},{rootMargin:'-45% 0px -50% 0px'});
document.querySelectorAll('main section').forEach(s=>obs.observe(s));
</script>
</body>
</html>
"""

if __name__ == "__main__":
    render()
