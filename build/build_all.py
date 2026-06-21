#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_all.py — 1コマンドで提出物を同期する連鎖ビルド（設計B：ビルド連鎖）。

  python build/build_all.py            # フルチェーン: pptx → PDF → web
  python build/build_all.py --from-pdf # PDFは手動更新済み → web だけ再生成

流れ:
  1) build_deck.py   → outputs/05_発表スライド.pptx   （python-pptx・"内容の正"）
  2) pptx → pdf      → outputs/05_発表スライド.pdf     （LibreOffice or PowerPoint）
  3) build_web.py    → outputs/web/                    （PyMuPDF：PDFをPNG化）

→ build_deck.py を1回直して本コマンドを叩けば、pptx・PDF・web がすべて同期する。
   （"webを直接いじる" 運用にしたい場合は設計A＝単一ソース化が必要。本スクリプトは設計B。）

PDF変換は環境依存:
  - LibreOffice があれば soffice --headless --convert-to pdf（推奨・確実・ヘッドレス可）
  - 無ければ PowerPoint COM（PowerPoint が入った "対話" 環境で動作。非対話セッションでは不可）
  どちらも使えなければ、手動で pptx→PDF をエクスポートしてから
  `python build/build_all.py --from-pdf` を実行（webだけ更新）。
"""
import os
import sys
import shutil
import tempfile
import subprocess
import pathlib

# Windowsの cp932 コンソールでも記号でクラッシュしないよう、出力のエラー処理だけ緩める
# （encodingは変えない＝日本語はそのまま表示される）。
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
OUT = ROOT / "outputs"
PPTX = OUT / "05_発表スライド.pptx"
PDF = OUT / "05_発表スライド.pdf"
PY = sys.executable


def run_py(script: pathlib.Path):
    print(f"\n>> {script.name}")
    r = subprocess.run([PY, str(script)], cwd=str(ROOT))
    if r.returncode != 0:
        sys.exit(f"X {script.name} が失敗しました (exit {r.returncode})")


def find_soffice():
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(c):
            return c
    return shutil.which("soffice") or shutil.which("soffice.exe")


def convert_via_soffice(soffice) -> bool:
    print(">> PDF変換: LibreOffice (headless)")
    subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                    "--outdir", str(OUT), str(PPTX)], check=True)
    return PDF.exists()  # soffice は <同名>.pdf を outdir に出力＝05_発表スライド.pdf


# 日本語ファイル名でCOMが開けないことがあるので、ASCIIの一時パスで変換して戻す。
_PS_CONVERT = r"""
$ErrorActionPreference="Stop"
$src=$args[0]; $dst=$args[1]
$tmp=[System.IO.Path]::GetTempPath()
$tp=Join-Path $tmp ("deck_"+[guid]::NewGuid().ToString("N")+".pptx")
$td=[System.IO.Path]::ChangeExtension($tp,".pdf")
Copy-Item -LiteralPath $src $tp -Force
$ppt=New-Object -ComObject PowerPoint.Application
try{
  try{$ppt.Visible=-1}catch{}
  $pres=$ppt.Presentations.Open($tp,-1,0,-1)   # ReadOnly, NotUntitled, WithWindow=True
  $pres.ExportAsFixedFormat($td,2)             # 2 = ppFixedFormatTypePDF
  $pres.Close()
}finally{ try{$ppt.Quit()}catch{} }
Copy-Item -LiteralPath $td $dst -Force
Remove-Item -LiteralPath $tp,$td -Force -ErrorAction SilentlyContinue
"""


def convert_via_powerpoint() -> bool:
    print(">> PDF変換: PowerPoint COM (PowerShell)")
    f = tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8")
    f.write(_PS_CONVERT)
    f.close()
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", f.name, str(PPTX), str(PDF)],
            capture_output=True, text=True)
        if r.returncode != 0 or not PDF.exists():
            msg = (r.stderr or r.stdout or "").strip()
            if msg:
                print("  ", msg.splitlines()[0][:300])
            return False
        return True
    finally:
        os.unlink(f.name)


def convert_pdf() -> bool:
    so = find_soffice()
    if so:
        try:
            if convert_via_soffice(so):
                return True
        except Exception as e:
            print("  LibreOffice失敗:", e)
    if os.name == "nt":
        try:
            if convert_via_powerpoint():
                return True
        except Exception as e:
            print("  PowerPoint失敗:", e)
    return False


def main():
    from_pdf = "--from-pdf" in sys.argv
    if not from_pdf:
        run_py(HERE / "build_deck.py")
        if not convert_pdf():
            print("\n" + "=" * 66)
            print("X PDF自動変換ができませんでした（LibreOffice未導入＆PowerPoint自動化不可）。")
            print("  → outputs/05_発表スライド.pptx を開き PDF にエクスポートして")
            print("     outputs/05_発表スライド.pdf を上書き保存し、次を実行:")
            print("        python build/build_all.py --from-pdf")
            print("  （恒久的に自動化するなら LibreOffice を入れると本コマンドだけで完結します）")
            print("=" * 66)
            sys.exit(2)
        print("OK: PDF 更新:", PDF)
    else:
        if not PDF.exists():
            sys.exit("X PDFが見つかりません。先に pptx→PDF を手動エクスポートしてください。")
        print(">> --from-pdf: deck/変換をスキップし、既存PDFからwebのみ再生成")

    run_py(HERE / "build_web.py")
    print("\nDONE: 同期完了 — pptx / pdf / web がすべて最新です。")


if __name__ == "__main__":
    main()
