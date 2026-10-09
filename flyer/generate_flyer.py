#!/usr/bin/env python3
"""Build the self-contained SVG, print PNG, and A5 PDF flyer."""

from __future__ import annotations

import base64
import shutil
import subprocess
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
WIDTH = 1748
HEIGHT = 2480


def data_uri(path: Path, mime: str) -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def build_svg() -> str:
    central_art = data_uri(ASSETS / "central-art.jpg", "image/jpeg")
    qr_code = data_uri(ASSETS / "qr-code.jpg", "image/jpeg")

    inter_regular = data_uri(
        Path("/usr/share/fonts/truetype/macos/Inter-Regular.ttf"),
        "font/ttf",
    )
    inter_semibold = data_uri(
        Path("/usr/share/fonts/truetype/macos/Inter-SemiBold.ttf"),
        "font/ttf",
    )
    noto_serif_bold = data_uri(
        Path("/usr/share/fonts/truetype/noto/NotoSerif-Bold.ttf"),
        "font/ttf",
    )
    noto_serif_italic = data_uri(
        Path("/usr/share/fonts/truetype/noto/NotoSerif-Italic.ttf"),
        "font/ttf",
    )

    return f"""\
<svg xmlns="http://www.w3.org/2000/svg"
     xmlns:xlink="http://www.w3.org/1999/xlink"
     width="{WIDTH}" height="{HEIGHT}" viewBox="0 0 {WIDTH} {HEIGHT}">
  <title>Листовка High Heels — Juliya Kudris</title>
  <defs>
    <style>
      @font-face {{
        font-family: "Flyer Inter";
        src: url("{inter_regular}") format("truetype");
        font-weight: 400;
      }}
      @font-face {{
        font-family: "Flyer Inter";
        src: url("{inter_semibold}") format("truetype");
        font-weight: 600;
      }}
      @font-face {{
        font-family: "Flyer Serif";
        src: url("{noto_serif_bold}") format("truetype");
        font-weight: 700;
      }}
      @font-face {{
        font-family: "Flyer Serif";
        src: url("{noto_serif_italic}") format("truetype");
        font-style: italic;
      }}
      .sans {{ font-family: "Flyer Inter", Arial, sans-serif; }}
      .serif {{ font-family: "Flyer Serif", Georgia, serif; }}
    </style>

    <linearGradient id="paper" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#fbf7f1"/>
      <stop offset="0.58" stop-color="#f4eadf"/>
      <stop offset="1" stop-color="#ead8ca"/>
    </linearGradient>
    <radialGradient id="topGlow" cx="50%" cy="0%" r="80%">
      <stop offset="0" stop-color="#ffffff" stop-opacity=".9"/>
      <stop offset=".65" stop-color="#f8eee5" stop-opacity=".3"/>
      <stop offset="1" stop-color="#f8eee5" stop-opacity="0"/>
    </radialGradient>
    <linearGradient id="wine" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#2a0612"/>
      <stop offset=".52" stop-color="#4b1027"/>
      <stop offset="1" stop-color="#741d3e"/>
    </linearGradient>
    <linearGradient id="gold" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#e9c18b"/>
      <stop offset=".5" stop-color="#bd8248"/>
      <stop offset="1" stop-color="#f0d2a3"/>
    </linearGradient>
    <linearGradient id="imageShade" x1="0" y1="0" x2="0" y2="1">
      <stop offset=".52" stop-color="#20020d" stop-opacity="0"/>
      <stop offset="1" stop-color="#20020d" stop-opacity=".28"/>
    </linearGradient>
    <filter id="shadow" x="-30%" y="-30%" width="160%" height="170%">
      <feDropShadow dx="0" dy="28" stdDeviation="28"
                    flood-color="#2d0715" flood-opacity=".24"/>
    </filter>
    <filter id="softBlur" x="-30%" y="-30%" width="160%" height="160%">
      <feGaussianBlur stdDeviation="72"/>
    </filter>
    <clipPath id="photoClip">
      <rect x="374" y="450" width="1000" height="1238" rx="30"/>
    </clipPath>
    <clipPath id="panelClip">
      <rect x="48" y="1770" width="1652" height="662" rx="38"/>
    </clipPath>
    <clipPath id="qrClip">
      <rect x="1203" y="1847" width="374" height="374" rx="16"/>
    </clipPath>
  </defs>

  <!-- Warm editorial background -->
  <rect width="{WIDTH}" height="{HEIGHT}" fill="url(#paper)"/>
  <rect width="{WIDTH}" height="1000" fill="url(#topGlow)"/>
  <circle cx="1610" cy="460" r="500" fill="#7a2442" opacity=".055"/>
  <circle cx="52" cy="1570" r="430" fill="#a36f51" opacity=".07"/>
  <ellipse cx="888" cy="1280" rx="720" ry="820"
           fill="#ffffff" opacity=".24" filter="url(#softBlur)"/>

  <!-- Fine print-safe border -->
  <rect x="48" y="48" width="1652" height="2384" rx="38"
        fill="none" stroke="#531128" stroke-width="3" opacity=".72"/>
  <path d="M92 346 H1656" stroke="#b47b47" stroke-width="2" opacity=".54"/>

  <!-- Header -->
  <text x="874" y="220" text-anchor="middle"
        class="serif" font-size="138" font-weight="700"
        letter-spacing="15" fill="#3c0b1d">HIGH HEELS</text>
  <text x="874" y="309" text-anchor="middle"
        class="sans" font-size="43" font-weight="600"
        letter-spacing="25" fill="#a36d3d">ТАНЦЫ</text>

  <!-- Vertical dance marker -->
  <rect x="96" y="565" width="164" height="752" rx="82"
        fill="#f7eee6" fill-opacity=".72"
        stroke="#4a0d22" stroke-width="4"/>
  <circle cx="178" cy="617" r="9" fill="#c18a52"/>
  <text x="178" y="740" text-anchor="middle"
        class="serif" font-size="68" font-weight="700" fill="#4a0d22">D</text>
  <text x="178" y="861" text-anchor="middle"
        class="serif" font-size="68" font-weight="700" fill="#4a0d22">A</text>
  <text x="178" y="982" text-anchor="middle"
        class="serif" font-size="68" font-weight="700" fill="#4a0d22">N</text>
  <text x="178" y="1103" text-anchor="middle"
        class="serif" font-size="68" font-weight="700" fill="#4a0d22">C</text>
  <text x="178" y="1224" text-anchor="middle"
        class="serif" font-size="68" font-weight="700" fill="#4a0d22">E</text>
  <path d="M178 1368 V1588" stroke="#4a0d22" stroke-width="3"/>
  <circle cx="178" cy="1624" r="12" fill="#c18a52"/>

  <!-- Main image card -->
  <rect x="348" y="424" width="1052" height="1290" rx="42"
        fill="#bd854e" opacity=".55"/>
  <rect x="360" y="436" width="1028" height="1266" rx="36"
        fill="#3b091b" filter="url(#shadow)"/>
  <image x="374" y="450" width="1000" height="1238"
         xlink:href="{central_art}"
         preserveAspectRatio="xMidYMid slice" clip-path="url(#photoClip)"/>
  <rect x="374" y="450" width="1000" height="1238"
        fill="url(#imageShade)" clip-path="url(#photoClip)"/>
  <rect x="374" y="450" width="1000" height="1238" rx="30"
        fill="none" stroke="#f2cf9b" stroke-width="3" opacity=".66"/>

  <!-- Small editorial accents -->
  <g transform="translate(1460 648)" opacity=".82">
    <path d="M0 0 H142" stroke="#4a0d22" stroke-width="3"/>
    <path d="M118 -12 L142 0 L118 12" fill="none"
          stroke="#4a0d22" stroke-width="3"/>
    <text x="2" y="52" class="sans" font-size="22" font-weight="600"
          letter-spacing="4" fill="#4a0d22">MOVE</text>
    <text x="2" y="86" class="sans" font-size="22" font-weight="600"
          letter-spacing="4" fill="#4a0d22">FEEL</text>
  </g>

  <!-- Information panel -->
  <rect x="48" y="1770" width="1652" height="662" rx="38"
        fill="url(#wine)"/>
  <ellipse cx="1510" cy="2130" rx="430" ry="440"
           fill="#b35777" opacity=".13" clip-path="url(#panelClip)"/>
  <ellipse cx="1060" cy="1690" rx="520" ry="260"
           fill="#f0c38d" opacity=".06" clip-path="url(#panelClip)"/>

  <text x="126" y="1878" class="sans" font-size="28" font-weight="600"
        letter-spacing="8" fill="#d6a76f">ТРЕНИРОВКИ</text>
  <text x="126" y="1962" class="sans" font-size="55" font-weight="600"
        letter-spacing="2" fill="#fffaf5">УТРЕННИЕ И</text>
  <text x="126" y="2032" class="sans" font-size="55" font-weight="600"
        letter-spacing="2" fill="#fffaf5">ВЕЧЕРНИЕ ТРЕНИРОВКИ</text>
  <path d="M126 2091 H1026" stroke="#d8a970" stroke-width="3" opacity=".72"/>

  <text x="126" y="2160" class="sans" font-size="25" font-weight="600"
        letter-spacing="7" fill="#d6a76f">ЗАПИСЬ ПО ТЕЛЕФОНУ</text>
  <text x="126" y="2240" class="sans" font-size="62" font-weight="600"
        letter-spacing="2" fill="#fffaf5">+7 (900) 549-75-78</text>

  <text x="126" y="2360" class="serif" font-size="69" font-style="italic"
        fill="#e5bd88">Juliya Kudris</text>

  <!-- QR card; the supplied code remains unaltered -->
  <rect x="1177" y="1821" width="426" height="426" rx="30"
        fill="#fff" filter="url(#shadow)"/>
  <image x="1203" y="1847" width="374" height="374"
         xlink:href="{qr_code}"
         preserveAspectRatio="xMidYMid meet" clip-path="url(#qrClip)"
         style="image-rendering:auto"/>
  <text x="1390" y="2311" text-anchor="middle"
        class="sans" font-size="25" font-weight="600"
        letter-spacing="6" fill="#e4bb86">ЗАПИСАТЬСЯ</text>
</svg>
"""


def main() -> None:
    svg = build_svg()
    svg_path = ROOT / "flyer.svg"
    html_path = ROOT / "flyer.html"
    png_path = ROOT / "flyer.png"
    pdf_path = ROOT / "flyer.pdf"

    svg_path.write_text(svg, encoding="utf-8")
    html_path.write_text(
        """<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>High Heels — Juliya Kudris</title>
  <style>
    @page { size: 148mm 210mm; margin: 0; }
    html, body { width: 148mm; height: 210mm; margin: 0; overflow: hidden; }
    body { background: #f4eadf; }
    svg { display: block; width: 148mm; height: 210mm; }
  </style>
</head>
<body>
"""
        + svg
        + "\n</body>\n</html>\n",
        encoding="utf-8",
    )

    chrome = shutil.which("google-chrome") or shutil.which("chromium")
    if not chrome:
        print(f"Created {svg_path}")
        print("Chrome was not found; skipped PNG and PDF rendering.")
        return

    with tempfile.TemporaryDirectory(prefix="flyer-chrome-") as profile:
        common = [
            chrome,
            "--headless=new",
            "--no-sandbox",
            "--disable-gpu",
            "--hide-scrollbars",
            "--disable-background-networking",
            "--disable-component-update",
            "--disable-default-apps",
            "--disable-software-rasterizer",
            "--no-first-run",
            "--run-all-compositor-stages-before-draw",
            "--virtual-time-budget=2000",
            f"--user-data-dir={profile}",
        ]
        subprocess.run(
            common
            + [
                f"--window-size={WIDTH},{HEIGHT}",
                "--force-device-scale-factor=1",
                f"--screenshot={png_path}",
                svg_path.as_uri(),
            ],
            check=True,
            timeout=45,
        )
        subprocess.run(
            common
            + [
                "--no-pdf-header-footer",
                f"--print-to-pdf={pdf_path}",
                html_path.as_uri(),
            ],
            check=True,
            timeout=45,
        )

    print(f"Created {svg_path}")
    print(f"Created {png_path}")
    print(f"Created {pdf_path}")


if __name__ == "__main__":
    main()
