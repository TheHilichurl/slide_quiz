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
username = None
for line in stdout.splitlines():
    if line.startswith("password="):
        token = line[len("password="):].strip()
    elif line.startswith("username="):
        username = line[len("username="):].strip()

if not token:
    print("[Error] Could not find GitHub token from git credential helper.")
    sys.exit(1)

headers = {
    "Authorization": f"Bearer {token}",
    "User-Agent": "SlideQuizDeployer",
    "Accept": "application/vnd.github.v3+json",
    "Content-Type": "application/json"
}

# 1. Verify User
req = urllib.request.Request("https://api.github.com/user", headers=headers)
try:
    with urllib.request.urlopen(req) as resp:
        user_data = json.loads(resp.read().decode('utf-8'))
        login = user_data.get('login')
        print(f"[1/4] Verified user: {login} ({user_data.get('html_url')})")
except urllib.error.HTTPError as e:
    print(f"[Error] Failed to verify user: {e.code} {e.read().decode('utf-8')}")
    sys.exit(1)

# 2. Check or Create Repo
repo_name = "slide_quiz"
repo_url = f"https://api.github.com/repos/{login}/{repo_name}"
repo_data = None

try:
    req = urllib.request.Request(repo_url, headers=headers)
    with urllib.request.urlopen(req) as resp:
        repo_data = json.loads(resp.read().decode('utf-8'))
        print(f"[2/4] Repository '{repo_name}' already exists: {repo_data.get('html_url')}")
except urllib.error.HTTPError as e:
    if e.code == 404:
        print(f"[2/4] Repository '{repo_name}' not found. Creating new repository...")
        create_payload = json.dumps({
            "name": repo_name,
            "description": "He thong On tap Trac nghiem GDQP-AN Bai A3 va A4 - Dai hoc Dai Nam",
            "private": False,
            "has_issues": True,
            "has_wiki": False,
            "auto_init": False
        }).encode('utf-8')
        create_req = urllib.request.Request("https://api.github.com/user/repos", data=create_payload, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(create_req) as resp:
                repo_data = json.loads(resp.read().decode('utf-8'))
                print(f"      Successfully created repository: {repo_data.get('html_url')}")
        except urllib.error.HTTPError as ce:
            print(f"[Error] Could not create repo: {ce.code} {ce.read().decode('utf-8')}")
            sys.exit(1)
    else:
        print(f"[Error] Failed checking repo: {e.code} {e.read().decode('utf-8')}")
        sys.exit(1)

clone_url = repo_data.get('clone_url')
html_url = repo_data.get('html_url')
print(f"      Remote URL: {clone_url}")

# Print summary
print("\n[READY] GitHub repository is ready for git push:")
print(f"        URL: {html_url}")
