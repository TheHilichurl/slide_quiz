import os, re, glob
from PIL import Image

html_files = glob.glob('*.html') + glob.glob('Slide_Agent_Framework/**/*.html', recursive=True)

all_ok = True
for html_file in sorted(html_files):
    dir_name = os.path.dirname(html_file)
    with open(html_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    srcs = re.findall(r'<img[^>]+src=[\'"]([^\'"]+)[\'"]', content)
    print(f'\nChecking {html_file} ({len(srcs)} img tags):')
    for src in sorted(set(srcs)):
        full_path = os.path.normpath(os.path.join(dir_name, src))
        if not os.path.exists(full_path):
            print(f'  [MISSING] {src} -> {full_path}')
            all_ok = False
        else:
            if src.lower().endswith('.svg'):
                print(f'  [SVG OK] {src} ({os.path.getsize(full_path)} bytes)')
            else:
                try:
                    with Image.open(full_path) as img:
                        print(f'  [IMG OK] {src}: {img.format} {img.size} {img.mode} ({os.path.getsize(full_path)/1024:.1f} KB)')
                except Exception as e:
                    print(f'  [CORRUPT] {src}: {e}')
                    all_ok = False

if all_ok:
    print('\n>>> ALL IMAGES IN ALL HTML FILES EXIST AND ARE VALID! <<<')
else:
    print('\n>>> SOME IMAGES FAILED! <<<')
