"""
Native Editable PPTX Export Engine for Slide Agent Framework
Đại Nam University ("Đại Nam Style")

Exports HTML presentations to 100% native, fully editable PowerPoint presentations (.pptx):
- Converts DOM elements to native PowerPoint Shapes, TextBoxes, Ovals, Lines, and Images.
- Preserves layout, coordinates (13.333 x 7.5 inches / 16:9), typography (Times New Roman), colors, and borders.
- Allows editing all text, changing shapes, formatting options, and replacing pictures in Microsoft PowerPoint.

Usage:
    python export_editable_pptx.py --file slide_A4_hoan_thien.html --out ./exports --name On_Tap_Bai_A4_DaiNam_Editable
    python export_editable_pptx.py --url http://localhost:8000/slide_A3_hoan_thien.html --out ./exports
"""

import argparse
import asyncio
import os
import sys
import base64
from playwright.async_api import async_playwright

sys.stdout.reconfigure(encoding='utf-8')

async def export_editable_pptx(target_url: str, output_pptx_path: str, converter_js_path: str):
    print(f"\n========================================================")
    print(f"  ĐẠI NAM UNIVERSITY - EDITABLE PPTX EXPORT ENGINE")
    print(f"========================================================")
    print(f"[1/4] Target Presentation: {target_url}")
    print(f"[2/4] Initializing Headless Chromium Engine...")

    # Locate available browser
    executable_path = None
    default_chrome = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
    default_edge = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
    if os.path.exists(default_chrome):
        executable_path = default_chrome
    elif os.path.exists(default_edge):
        executable_path = default_edge

    launch_kwargs = {
        "headless": True,
        "args": ["--allow-file-access-from-files", "--disable-web-security"]
    }
    if executable_path:
        launch_kwargs["executable_path"] = executable_path

    async with async_playwright() as p:
        browser = await p.chromium.launch(**launch_kwargs)
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})

        await page.goto(target_url, wait_until="networkidle")
        await page.wait_for_timeout(600)

        # Inject converter library if not already present
        has_converter = await page.evaluate("() => typeof window.HtmlToPptxConverter !== 'undefined'")
        if not has_converter:
            print(f"       Injecting HtmlToPptxConverter from: {converter_js_path}")
            with open(converter_js_path, "r", encoding="utf-8") as f:
                await page.evaluate(f.read())

        # Count total slides
        total_slides = await page.evaluate("() => document.querySelectorAll('.slide').length")
        print(f"[3/4] Detected {total_slides} slides. Starting DOM-to-PowerPoint standardization...")

        # Execute conversion inside page context
        b64_pptx = await page.evaluate(f"""async () => {{
            const pptx = new PptxGenJS();
            pptx.defineLayout({{ name: 'LAYOUT_16X9_DNU', width: 13.333333, height: 7.5 }});
            pptx.layout = 'LAYOUT_16X9_DNU';
            pptx.title = document.title || 'Đại Nam University Presentation';
            pptx.author = 'Trường Đại học Đại Nam';
            pptx.company = 'Đại học Đại Nam (DNU)';

            const allSlides = document.querySelectorAll('.slide');
            const total = allSlides.length;

            for (let i = 1; i <= total; i++) {{
                allSlides.forEach(s => s.classList.remove('active'));
                const targetSlide = document.getElementById(`slide-${{i}}`) || allSlides[i - 1];
                if (!targetSlide) continue;

                targetSlide.classList.add('active');
                if (window.lucide) window.lucide.createIcons();

                await new Promise(r => setTimeout(r, 60));
                await HtmlToPptxConverter.convertSlideToPptx(targetSlide, pptx, i, total);
            }}

            return await pptx.write('base64');
        }}""")

        await browser.close()

    print(f"[4/4] Writing PowerPoint file (.pptx)...")
    pptx_bytes = base64.b64decode(b64_pptx)
    os.makedirs(os.path.dirname(os.path.abspath(output_pptx_path)), exist_ok=True)
    with open(output_pptx_path, "wb") as f:
        f.write(pptx_bytes)

    print(f"       [SUCCESS] Exported native editable PPTX: {output_pptx_path} ({len(pptx_bytes):,} bytes)")
    print(f"       All texts, shapes, cards, badges, and pictures are fully editable in Microsoft PowerPoint!")
    print(f"========================================================\n")

def main():
    parser = argparse.ArgumentParser(description="Slide Agent Framework - Native Editable PPTX Exporter")
    parser.add_argument("--url", default="", help="HTTP URL to presentation (e.g. http://localhost:8000/slide_A4_hoan_thien.html)")
    parser.add_argument("--file", default="slide_A4_hoan_thien.html", help="Local HTML file path")
    parser.add_argument("--out", default="./exports", help="Output directory")
    parser.add_argument("--name", default="", help="Base name of exported file")
    args = parser.parse_args()

    # Determine target URL
    if args.url:
        target_url = args.url
    else:
        abs_html = os.path.abspath(args.file)
        if not os.path.exists(abs_html):
            print(f"[Error] File not found: {abs_html}")
            sys.exit(1)
        target_url = "file:///" + abs_html.replace("\\", "/")

    # Output filename
    if args.name:
        base_name = args.name
    else:
        base_name = os.path.splitext(os.path.basename(args.file if not args.url else args.url.split('/')[-1]))[0] + "_Editable"

    output_path = os.path.join(args.out, f"{base_name}.pptx")

    # Locate converter.js
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(script_dir, "..", ".."))
    converter_js_path = os.path.join(project_root, "assets", "html_to_pptx_converter.js")
    if not os.path.exists(converter_js_path):
        converter_js_path = os.path.join(script_dir, "html_to_pptx_converter.js")

    asyncio.run(export_editable_pptx(target_url, output_path, converter_js_path))

if __name__ == '__main__':
    main()
