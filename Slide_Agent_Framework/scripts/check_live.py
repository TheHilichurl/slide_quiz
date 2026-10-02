import urllib.request
import time
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = "https://thehilichurl.github.io/slide_quiz/"
print(f"Checking deployment status at: {url}")

for i in range(10):
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            print(f"[SUCCESS] HTTP {resp.status} OK!")
            html = resp.read().decode('utf-8')
            t_start = html.find("<title>")
            t_end = html.find("</title>")
            if t_start != -1 and t_end != -1:
                print(f"Title: {html[t_start+7:t_end]}")
            print("Website is deployed and fully operational!")
            sys.exit(0)
    except Exception as e:
        print(f"[{i+1}/10] Still deploying: {e}. Retrying in 6s...")
        time.sleep(6)
