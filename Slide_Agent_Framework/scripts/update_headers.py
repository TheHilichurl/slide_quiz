import re, os

files = ['slide_A3_hoan_thien.html', 'slide_A4_hoan_thien.html', 'Slide_Agent_Framework/boilerplate/index.html']

pattern = re.compile(
    r'<div class="header-right">\s*<div class="header-badge">.*?</div>\s*</div>',
    re.DOTALL
)

replacement = '<div class="header-right"></div>'

for fpath in files:
    if not os.path.exists(fpath):
        continue
    with open(fpath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    matches = len(pattern.findall(content))
    content_new = pattern.sub(replacement, content)
    with open(fpath, 'w', encoding='utf-8') as f:
        f.write(content_new)
    print(f'{fpath}: successfully removed badge from {matches} headers.')

print('Done removing header badges!')
