import re, os

files = ['slide_A3_hoan_thien.html', 'slide_A4_hoan_thien.html', 'Slide_Agent_Framework/boilerplate/index.html']

old_pattern = re.compile(
    r'<header class="slide-header">\s*<div class="header-left">\s*<div class="brand-slogan">HỌC ĐỂ THAY ĐỔI</div>\s*</div>\s*<div class="header-right">\s*<img src="assets/dai-nam-logo-ngang\.svg"[^>]*>\s*</div>\s*</header>',
    re.DOTALL
)

new_header = '''<header class="slide-header">
          <div class="header-left">
            <img src="assets/dai-nam-logo-ngang.svg" alt="Trường Đại học Đại Nam" class="official-logo-img">
            <div class="header-divider"></div>
            <div class="brand-slogan">HỌC ĐỂ THAY ĐỔI</div>
          </div>
          <div class="header-right">
            <div class="header-badge">
              <i data-lucide="shield-check" style="width: 18px; height: 18px;"></i>
              <span>HỆ THỐNG ÔN TẬP GDQP-AN</span>
            </div>
          </div>
        </header>'''

for fpath in files:
    if not os.path.exists(fpath):
        continue
    with open(fpath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    matches = len(old_pattern.findall(content))
    content_new = old_pattern.sub(new_header, content)
    with open(fpath, 'w', encoding='utf-8') as f:
        f.write(content_new)
    print(f'{fpath}: successfully updated {matches} slide headers.')

print('Done updating headers!')
