#!/bin/bash
# Installation Script for Iboga Lab Invoice Skill (macOS / Linux)
# Copies or downloads the skill into the global Gemini configurations directory

set -e

SKILL_NAME="ibogalab-invoice-generator"
GLOBAL_GEMINI_PATH="$HOME/.gemini/config/skills/$SKILL_NAME"
REPO_OWNER="ogoulajesse"
REPO_NAME="ibogalab-invoice-generator"
BASE_URL="https://raw.githubusercontent.com/$REPO_OWNER/$REPO_NAME/main"

echo -e "\033[0;32mInstalling Iboga Lab Invoice Generator skill...\033[0m"

# 1. Determine if running locally or remotely
IS_REMOTE=true
SOURCE_PATH="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd 2>/dev/null || pwd)"

if [ -d "$SOURCE_PATH/skills/$SKILL_NAME" ]; then
    IS_REMOTE=false
    echo -e "\033[0;36mLocal repository detected. Running in offline installation mode.\033[0m"
else
    echo -e "\033[0;36mRemote installation mode. Downloading skill files from GitHub...\033[0m"
fi

# 2. Create target directories
mkdir -p "$GLOBAL_GEMINI_PATH"
mkdir -p "$GLOBAL_GEMINI_PATH/scripts"
mkdir -p "$GLOBAL_GEMINI_PATH/resources"
mkdir -p "$GLOBAL_GEMINI_PATH/examples"

# 3. Copy or download files
FILES_TO_COPY=(
    "SKILL.md"
    "scripts/sync.py"
    "resources/style.css"
    "resources/logo1.png"
    "resources/logo2.png"
    "examples/devis_template.md"
    "examples/facture_template.md"
)

for file in "${FILES_TO_COPY[@]}"; do
    dest_file="$GLOBAL_GEMINI_PATH/$file"
    parent_dir="$(dirname "$dest_file")"
    
    # Ensure parent folder exists
    mkdir -p "$parent_dir"
    
    if [ "$IS_REMOTE" = true ]; then
        # Remote mode: Download from Raw Github URLs
        web_url="$BASE_URL/skills/$SKILL_NAME/$file"
        echo -e "\033[0;36mDownloading $file...\033[0m"
        curl -fsSL "$web_url" -o "$dest_file" || { echo -e "\033[0;31mError: Failed to download $file\033[0m"; exit 1; }
    else
        # Local mode: Copy from offline folder
        src_file="$SOURCE_PATH/skills/$SKILL_NAME/$file"
        if [ -f "$src_file" ]; then
            cp "$src_file" "$dest_file"
            echo -e "\033[0;36mCopied $file\033[0m"
        else
            echo -e "\033[0;33mWarning: Source file not found: $src_file\033[0m"
        fi
    fi
done

# 4. Check python & dependencies
echo -e "\n\033[0;32mChecking system dependencies...\033[0m"
if command -v python3 &>/dev/null; then
    echo -e "\033[0;36mPython 3 is available.\033[0m"
    
    # Try importing docx, markdown, and yaml
    echo -e "\033[0;36mVerifying Python packages...\033[0m"
    python3 -c "
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
else:
    echo -e "\033[0;31mWarning: Python 3 was not found in your PATH. Please install Python 3 to run the compiler.\033[0m"
fi

# 5. Check for browser
if command -v google-chrome &>/dev/null || command -v google-chrome-stable &>/dev/null; then
    echo -e "\033[0;36mGoogle Chrome found (available for PDF print).\033[0m"
elif [ -d "/Applications/Google Chrome.app" ]; then
    echo -e "\033[0;36mGoogle Chrome app found on macOS (available for PDF print).\033[0m"
elif command -v microsoft-edge &>/dev/null; then
    echo -e "\033[0;36mMicrosoft Edge found (available for PDF print).\033[0m"
else
    echo -e "\033[0;33mWarning: Headless browser not found in standard Unix PATHs. PDF generation may fail unless Chrome/Edge is installed.\033[0m"
fi

echo -e "\n\033[0;32mSkill installed successfully!\033[0m"
echo -e "\033[0;33mYou can now ask your Gemini agent in Antigravity or ClaudeCode to generate quotes and invoices.\033[0m"
