"""Render the visual blocks of docs/index.html to PNGs for README.md (light and dark).

GitHub strips CSS from READMEs, so the README shows these blocks as images and keeps the
text people copy (install prompt, tables) as Markdown. Re-run after editing docs/index.html:

    py -3.14 docs/readme/render.py

Needs Chrome or Edge, bs4 and Pillow. Output: docs/readme/<block>-<light|dark>.png, transparent
background, 2x pixel density, trimmed to the content.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from bs4 import BeautifulSoup
from PIL import Image

HERE = Path(__file__).resolve().parent
PAGE = HERE.parent / "index.html"
WIDTH = 820  # CSS px, about the width of GitHub's README column

BROWSERS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

RENDER_CSS = """
html,body{background:transparent!important}
body{padding:6px!important;margin:0!important;width:%dpx}
.wrap{max-width:none!important;margin:0!important}
header{border-bottom:0!important;padding:0!important}
figure{margin:0!important}
.chips{margin-top:0!important}
""" % (WIDTH,)


def blocks(soup):
    """name -> list of elements rendered together, in page order."""
    q = soup.select_one
    return {
        "hero": [q("header")],
        "diagram": [q("figure .dg")],
        "cards": [q(".cards")],
        "stats": [q(".stats")],
        "checks": [q(".chips"), q(".callout")],
        "flow": [q(".flow")],
    }


def dark_tokens(css):
    m = re.search(r"@media \(prefers-color-scheme: dark\)\{\s*:root\{([^}]*)\}", css)
    if not m:
        sys.exit("dark token block not found in docs/index.html")
    return ":root{%s}" % m.group(1)


def render(html, out, browser):
    tmp = Path(tempfile.mkdtemp(prefix="xl-readme-"))
    try:
        src = tmp / "block.html"
        src.write_text(html, encoding="utf-8")
        shot = tmp / "shot.png"
        subprocess.run([browser, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
                        f"--user-data-dir={tmp / 'profile'}", f"--window-size={WIDTH + 12},2600",
                        "--force-device-scale-factor=2", "--default-background-color=00000000",
                        "--virtual-time-budget=8000", f"--screenshot={shot}", src.as_uri()],
                       check=True, capture_output=True, timeout=120)
        im = Image.open(shot).convert("RGBA")
        box = im.getchannel("A").getbbox()
        if not box:
            sys.exit(f"{out.name}: nothing rendered")
        im.crop(box).save(out, optimize=True)
        print(f"{out.name}: {im.crop(box).size}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    browser = next((b for b in BROWSERS if os.path.isfile(b)), None)
    if not browser:
        sys.exit("Chrome or Edge not found")
    soup = BeautifulSoup(PAGE.read_text(encoding="utf-8"), "html.parser")
    head = "".join(str(t) for t in soup.head.find_all(["meta", "link", "style"]))
    css = soup.head.style.string or ""
    for name, els in blocks(soup).items():
        if not all(els):
            sys.exit(f"block {name!r} not found in docs/index.html")
        body = "".join(str(e) for e in els)
        for theme in ("light", "dark"):
            extra = RENDER_CSS + (dark_tokens(css) if theme == "dark" else "")
            html = (f'<!doctype html><html lang="en"><head>{head}<style>{extra}</style></head>'
                    f'<body><div class="wrap">{body}</div></body></html>')
            render(html, HERE / f"{name}-{theme}.png", browser)


if __name__ == "__main__":
    main()
