import re, glob, os

files = ['slide_A3_hoan_thien.html', 'slide_A4_hoan_thien.html', 'Slide_Agent_Framework/boilerplate/index.html']

for filepath in files:
    if not os.path.exists(filepath):
        continue
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Find all images
    all_imgs = sorted(list(set(re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', content))))
    
    # 1. Add loading="eager" decoding="async" to quiz-visual-img
    def add_attrs(match):
        tag = match.group(0)
        if 'loading=' not in tag:
            tag = tag[:-1] + ' loading="eager"' + tag[-1]
        if 'decoding=' not in tag:
            tag = tag[:-1] + ' decoding="async"' + tag[-1]
        return tag

    content = re.sub(r'<img[^>]+class=["\'][^"\']*quiz-visual-img[^"\']*["\'][^>]*>', add_attrs, content)

    # 2. Add preload script right before </head> if not already added
    preload_code = f"""  <!-- Immediate Asset Preloading for Zero Latency -->
  <script>
    (function() {{
      const assetUrls = {all_imgs};
      assetUrls.forEach(function(url) {{
        const link = document.createElement('link');
        link.rel = 'preload';
        link.as = 'image';
        link.href = url;
        document.head.appendChild(link);
        const img = new Image();
        img.src = url;
      }});
    }})();
  </script>
</head>"""

    if 'Immediate Asset Preloading' not in content:
        content = content.replace('</head>', preload_code, 1)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'Enhanced {filepath} with eager attributes and preloading {len(all_imgs)} images.')

print('Done!')
