import subprocess
import json
import urllib.request
import urllib.error
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Get token from git credential
proc = subprocess.Popen(['git', 'credential', 'fill'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
stdout, stderr = proc.communicate(input="protocol=https\nhost=github.com\n")

token = None
for line in stdout.splitlines():
    if line.startswith("password="):
        token = line[len("password="):].strip()

if not token:
    print("[Error] No token found.")
    sys.exit(1)

headers = {
    "Authorization": f"Bearer {token}",
    "User-Agent": "SlideQuizDeployer",
    "Accept": "application/vnd.github.v3+json",
    "Content-Type": "application/json"
}

owner = "TheHilichurl"
repo = "slide_quiz"
pages_url = f"https://api.github.com/repos/{owner}/{repo}/pages"

print(f"[1/3] Enabling GitHub Pages for {owner}/{repo} from branch 'main' (root '/')...")

# Payload to configure GitHub Pages
payload = json.dumps({
    "source": {
        "branch": "main",
        "path": "/"
    }
}).encode('utf-8')

req = urllib.request.Request(pages_url, data=payload, headers=headers, method="POST")

try:
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read().decode('utf-8'))
        print(f"[2/3] GitHub Pages successfully created!")
        print(f"      Status: {data.get('status')}")
        print(f"      Live URL: {data.get('html_url')}")
except urllib.error.HTTPError as e:
    err_body = e.read().decode('utf-8')
    if e.code == 409: # Already enabled
        print(f"[2/3] GitHub Pages is already enabled for this repository.")
    else:
        print(f"[Notice] Pages API response ({e.code}): {err_body}")

# Check current Pages status
get_req = urllib.request.Request(pages_url, headers=headers)
try:
    with urllib.request.urlopen(get_req) as resp:
        data = json.loads(resp.read().decode('utf-8'))
        print(f"[3/3] Verified GitHub Pages endpoint:")
        print(f"      Live Website URL: {data.get('html_url')}")
        print(f"      Build Status: {data.get('status')}")
except urllib.error.HTTPError as e:
    print(f"      Expected Live URL: https://{owner.lower()}.github.io/{repo}/")
