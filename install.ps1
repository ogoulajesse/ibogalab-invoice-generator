# Installation Script for Iboga Lab Invoice Skill (Windows)
# Copies or downloads the skill into the global Gemini configurations directory

$ErrorActionPreference = "Stop"

$SkillName = "ibogalab-invoice-generator"
$GlobalGeminiPath = Join-Path $HOME ".gemini\config\skills\$SkillName"
$RepoOwner = "ogoulajesse"
$RepoName = "ibogalab-invoice-generator"
$BaseUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/main"

Write-Host "Installing Iboga Lab Invoice Generator skill..." -ForegroundColor Green

# 1. Determine if running locally or remotely
$IsRemote = $true
if ($PSScriptRoot) {
    $LocalSkillFolder = Join-Path $PSScriptRoot "skills\$SkillName"
    if (Test-Path $LocalSkillFolder) {
        $IsRemote = $false
        $SourcePath = $PSScriptRoot
        Write-Host "Local repository detected. Running in offline installation mode." -ForegroundColor Gray
    }
}

if ($IsRemote) {
    Write-Host "Remote installation mode. Downloading skill files from GitHub..." -ForegroundColor Gray
}

# 2. Create target directories
$Dirs = @("", "scripts", "resources", "examples")
foreach ($dir in $Dirs) {
    $targetPath = if ($dir) { Join-Path $GlobalGeminiPath $dir } else { $GlobalGeminiPath }
    if (!(Test-Path $targetPath)) {
        New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
        Write-Host "Created directory: $targetPath" -ForegroundColor Cyan
    }
}

# 3. Copy or download files
# Array of relative file paths within the skill folder
$FilesToCopy = @(
    "SKILL.md",
    "scripts/sync.py",
    "resources/style.css",
    "resources/logo1.png",
    "resources/logo2.png",
    "examples/devis_template.md",
    "examples/facture_template.md"
)

foreach ($file in $FilesToCopy) {
    $destFile = Join-Path $GlobalGeminiPath $file
    
    # Create parent folder for file if it doesn't exist (e.g. scripts/sync.py)
    $parentDir = Split-Path $destFile
    if (!(Test-Path $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }
    
    if ($IsRemote) {
        # Remote mode: Download from Raw Github URLs
        # Replace forward slashes with Windows backslashes for local paths
        $webUrl = "$BaseUrl/skills/$SkillName/$file"
        Write-Host "Downloading $file..." -ForegroundColor Cyan
        try {
            Invoke-WebRequest -Uri $webUrl -OutFile $destFile -UseBasicParsing
        } catch {
            Write-Error "Failed to download file from $webUrl. Ensure you have internet connection."
        }
    } else {
        # Local mode: Copy from offline folder
        $srcFile = Join-Path $SourcePath "skills\$SkillName\$file"
        if (Test-Path $srcFile) {
            Copy-Item $srcFile $destFile -Force
            Write-Host "Copied $file" -ForegroundColor Cyan
        } else {
            Write-Warning "Source file not found: $srcFile"
        }
    }
}

# 4. Check python & dependencies
Write-Host "`nChecking system dependencies..." -ForegroundColor Green
if (Get-Command python -ErrorAction SilentlyContinue) {
    Write-Host "Python is available." -ForegroundColor Cyan
    
    # Try importing docx, markdown, and yaml
    Write-Host "Verifying Python packages..." -ForegroundColor Cyan
    python -c "
try:
    import markdown
    print('  [OK] markdown')
except ImportError:
    print('  [..] markdown missing (will self-install on first run)')
try:
    import docx
    print('  [OK] python-docx')
except ImportError:
    print('  [..] python-docx missing (will self-install on first run)')
try:
    import yaml
    print('  [OK] pyyaml')
except ImportError:
    print('  [..] pyyaml missing (will self-install on first run)')
"
} else {
    Write-Warning "Python was not found in your PATH. Please install Python 3 to run the compiler."
}

# 5. Check for browser
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$chromePathx86 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

if ((Test-Path $chromePath) -or (Test-Path $chromePathx86)) {
    Write-Host "Google Chrome found (available for PDF print)." -ForegroundColor Cyan
} elseif (Test-Path $edgePath) {
    Write-Host "Microsoft Edge found (available for PDF print)." -ForegroundColor Cyan
} else {
    Write-Warning "Neither Google Chrome nor Microsoft Edge was found in their standard installation paths. Headless PDF generation might fail unless a browser is in your PATH."
}

Write-Host "`nSkill installed successfully!" -ForegroundColor Green
Write-Host "You can now ask your Gemini agent in Antigravity or ClaudeCode to generate quotes and invoices." -ForegroundColor Yellow
