import asyncio
import os
from playwright.async_api import async_playwright

async def capture_html_slides():
    executable_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
    async with async_playwright() as p:
        browser = await p.chromium.launch(executable_path=executable_path, headless=True)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        
        # Test slide_A3_hoan_thien.html
        await page.goto('file:///C:/Users/Giang/Desktop/slideA3A4/slide_A3_hoan_thien.html', wait_until='networkidle')
        await page.wait_for_timeout(1000)
        
        out_dir = 'Slide_Agent_Framework/exports/html_slides_a3'
        os.makedirs(out_dir, exist_ok=True)
        
        for i in range(1, 21):
            await page.evaluate(f'window.goToSlide({i})')
            await page.wait_for_timeout(300)
            
            # Check if visual img is loaded
            status = await page.evaluate('''() => {
                const s = document.querySelector('.slide.active');
                const img = s ? s.querySelector('.quiz-visual-img') : null;
                return {
                    src: img ? img.src : null,
                    complete: img ? img.complete : false,
                    naturalWidth: img ? img.naturalWidth : 0,
                    naturalHeight: img ? img.naturalHeight : 0,
                    display: img ? window.getComputedStyle(img).display : null,
                    visibility: img ? window.getComputedStyle(img).visibility : null,
                    opacity: img ? window.getComputedStyle(img).opacity : null
                };
            }''')
            
            screenshot_path = f'{out_dir}/slide_{i}.png'
            await page.screenshot(path=screenshot_path)
            print(f"Slide {i:2}: img={os.path.basename(status['src'] or '')} natW={status['naturalWidth']} natH={status['naturalHeight']} -> {screenshot_path}")

        # Also test slide_A4_hoan_thien.html
        await page.goto('file:///C:/Users/Giang/Desktop/slideA3A4/slide_A4_hoan_thien.html', wait_until='networkidle')
        await page.wait_for_timeout(1000)
        out_dir_a4 = 'Slide_Agent_Framework/exports/html_slides_a4'
        os.makedirs(out_dir_a4, exist_ok=True)
        for i in range(1, 21):
            await page.evaluate(f'window.goToSlide({i})')
            await page.wait_for_timeout(300)
            status = await page.evaluate('''() => {
                const s = document.querySelector('.slide.active');
                const img = s ? s.querySelector('.quiz-visual-img') : null;
                return {
                    src: img ? img.src : null,
                    complete: img ? img.complete : false,
                    naturalWidth: img ? img.naturalWidth : 0,
                    naturalHeight: img ? img.naturalHeight : 0
                };
            }''')
            screenshot_path = f'{out_dir_a4}/slide_{i}.png'
            await page.screenshot(path=screenshot_path)
            print(f"A4 Slide {i:2}: img={os.path.basename(status['src'] or '')} natW={status['naturalWidth']} natH={status['naturalHeight']}")

        await browser.close()

asyncio.run(capture_html_slides())
