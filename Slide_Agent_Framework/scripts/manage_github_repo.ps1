$creds = @"
protocol=https
host=github.com
"@ | git credential fill

$token = ""
foreach ($line in ($creds -split "`r?`n")) {
    if ($line.Trim().StartsWith("password=")) {
        $token = $line.Trim().Substring("password=".Length)
    }
}

if (-not $token) {
    Write-Error "Could not retrieve token."
    exit 1
}

$headers = @{
    "Authorization" = "Bearer $token"
    "User-Agent" = "PowerShell-GitHub-Manager"
    "Accept" = "application/vnd.github.v3+json"
}

Write-Output "[1/4] Verifying GitHub user..."
$user = Invoke-RestMethod -Uri "https://api.github.com/user" -Headers $headers -Method Get
Write-Output "      Logged in as: $($user.login)"

Write-Output "[2/4] Checking/Creating repository 'slide_quiz'..."
$repoName = "slide_quiz"
$repoExists = $false
try {
    $repo = Invoke-RestMethod -Uri "https://api.github.com/repos/$($user.login)/$repoName" -Headers $headers -Method Get
    Write-Output "      Repository already exists: $($repo.html_url)"
    $repoExists = $true
} catch {
    Write-Output "      Repository not found. Creating new repository '$repoName'..."
}

if (-not $repoExists) {
    $body = @{
        name = $repoName
        description = "Hệ Thống Ôn Tập Trắc Nghiệm GDQP-AN Bài A3 & A4 - Trường Đại Học Đại Nam"
        private = $false
        has_issues = $true
        has_wiki = $false
        has_downloads = $true
    } | ConvertTo-Json

    $newRepo = Invoke-RestMethod -Uri "https://api.github.com/user/repos" -Headers $headers -Method Post -Body $body
    Write-Output "      Successfully created repository: $($newRepo.html_url)"
}

Write-Output "[3/4] Enabling GitHub Pages for repository..."
try {
    $pagesBody = @{
        source = @{
            branch = "main"
            path = "/"
        }
    } | ConvertTo-Json
    $pages = Invoke-RestMethod -Uri "https://api.github.com/repos/$($user.login)/$repoName/pages" -Headers $headers -Method Post -Body $pagesBody
    Write-Output "      GitHub Pages enabled: $($pages.html_url)"
} catch {
    Write-Output "      Pages enable notice: $($_.Exception.Message)"
}
