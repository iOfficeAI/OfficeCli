$repo = "iOfficeAI/OfficeCLI"
$asset = "officecli-win-x64.exe"
$binary = "officecli.exe"

$source = $null

# Step 1: Try downloading from GitHub
$url = "https://github.com/$repo/releases/latest/download/$asset"
$checksumUrl = "https://github.com/$repo/releases/latest/download/SHA256SUMS"
$tempFile = "$env:TEMP\$binary"
Write-Host "Downloading OfficeCLI..."
try {
    Invoke-WebRequest -Uri $url -OutFile $tempFile
    # Verify checksum if available
    $checksumOk = $false
    try {
        $checksumFile = "$env:TEMP\officecli-SHA256SUMS"
        Invoke-WebRequest -Uri $checksumUrl -OutFile $checksumFile
        $checksumContent = Get-Content $checksumFile
        $expectedLine = $checksumContent | Where-Object { $_ -match $asset }
        if ($expectedLine) {
            $expected = ($expectedLine -split '\s+')[0]
            $actual = (Get-FileHash -Path $tempFile -Algorithm SHA256).Hash.ToLower()
            if ($expected -eq $actual) {
                $checksumOk = $true
                Write-Host "Checksum verified."
            } else {
                Write-Host "Checksum mismatch! Expected: $expected, Got: $actual"
                Remove-Item -Force $tempFile, $checksumFile -ErrorAction SilentlyContinue
                exit 1
            }
        }
        Remove-Item -Force $checksumFile -ErrorAction SilentlyContinue
    } catch {
        Write-Host "Checksum file not available, skipping verification."
    }
    $output = & $tempFile --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        $source = $tempFile
        Write-Host "Download verified."
    } else {
        Write-Host "Downloaded file is not a valid OfficeCLI binary."
        Remove-Item -Force $tempFile -ErrorAction SilentlyContinue
    }
} catch {
    Write-Host "Download failed."
}

# Step 2: Fallback to local files
if (-not $source) {
    Write-Host "Looking for local binary..."
    $candidates = @(".\$asset", ".\$binary", ".\bin\$asset", ".\bin\$binary", ".\bin\release\$asset", ".\bin\release\$binary")
    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            $output = & $candidate --version 2>&1
            if ($LASTEXITCODE -eq 0) {
                $source = $candidate
                Write-Host "Found valid binary at $candidate"
                break
            }
        }
    }
}

if (-not $source) {
    Write-Host "Error: Could not find a valid OfficeCLI binary."
    Write-Host "Download manually from: https://github.com/$repo/releases"
    exit 1
}

# Step 3: Install
$existing = Get-Command $binary -ErrorAction SilentlyContinue
if ($existing) {
    $installDir = Split-Path $existing.Source
    Write-Host "Found existing installation at $($existing.Source), upgrading..."
} else {
    $installDir = "$env:LOCALAPPDATA\OfficeCLI"
}

New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Copy-Item -Force $source "$installDir\$binary"

Remove-Item -Force $tempFile -ErrorAction SilentlyContinue

# Add to PATH if not already there
$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($currentPath -notlike "*$installDir*") {
    [Environment]::SetEnvironmentVariable("Path", "$currentPath;$installDir", "User")
    Write-Host "Added $installDir to PATH (restart your terminal to take effect)."
}

# Step 4: Install AI agent skills (first install only)
$skillMarker = "$installDir\.officecli-skills-installed"
if (-not (Test-Path $skillMarker)) {
    $skillTargets = @()
    $tools = @{
        "$env:USERPROFILE\.claude" = "Claude Code"
        "$env:USERPROFILE\.copilot" = "GitHub Copilot"
        "$env:USERPROFILE\.agents" = "Codex CLI"
        "$env:USERPROFILE\.cursor" = "Cursor"
        "$env:USERPROFILE\.windsurf" = "Windsurf"
        "$env:USERPROFILE\.minimax" = "MiniMax CLI"
        "$env:USERPROFILE\.openclaw" = "OpenClaw"
        "$env:USERPROFILE\.nanobot\workspace" = "NanoBot"
        "$env:USERPROFILE\.zeroclaw\workspace" = "ZeroClaw"
        "$env:USERPROFILE\.hermes" = "Hermes Agent"
    }
    foreach ($dir in $tools.Keys) {
        if (Test-Path $dir) {
            $skillTargets += "$dir\skills\officecli"
            Write-Host "$($tools[$dir]) detected."
        }
    }

    if ($skillTargets.Count -gt 0) {
        Write-Host "Downloading officecli skill..."
        $tempSkill = "$env:TEMP\officecli-skill.md"
        try {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/$repo/main/SKILL.md" -OutFile $tempSkill
            foreach ($target in $skillTargets) {
                New-Item -ItemType Directory -Force -Path $target | Out-Null
                Copy-Item -Force $tempSkill "$target\SKILL.md"
                Write-Host "  Installed: $target\SKILL.md"
            }
            Remove-Item -Force $tempSkill -ErrorAction SilentlyContinue
        } catch {}
    }
    New-Item -ItemType File -Force -Path $skillMarker | Out-Null
}

Write-Host "OfficeCLI installed successfully!"
Write-Host "Run 'officecli --help' to get started."
