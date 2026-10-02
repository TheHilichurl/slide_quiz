$creds = @"
protocol=https
host=github.com
"@ | git credential fill

$token = ""
foreach ($line in ($creds -split "`n")) {
    if ($line.Trim().StartsWith("password=")) {
        $token = $line.Trim().Substring("password=".Length)
    }
}

if ($token) {
    $token | gh auth login --with-token
    gh auth status
} else {
    Write-Error "Could not retrieve GitHub token."
}
