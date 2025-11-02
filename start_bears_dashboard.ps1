# ==========================================
# Bears Weekly Tracker - Start Script (fixed)
# ==========================================

# Build the path safely (works whether OneDrive is used or not)
$projectPath = Join-Path $env:USERPROFILE "OneDrive\Documents\GitHub\bears-tracker"

Write-Host "🐻 Starting Bears Weekly Tracker..."
Write-Host "📁 Project path: $projectPath"

if (!(Test-Path $projectPath)) {
    Write-Host "❌ Project folder not found: $projectPath"
    Write-Host "   Check that the folder exists or adjust this script's path."
    Read-Host "Press Enter to exit"
    exit 1
}

Set-Location $projectPath

# Ensure the main app file exists (correct name!)
$appFile = "bears_dashboard.py"   # <— important! not 'dashboard.py'
if (!(Test-Path $appFile)) {
    Write-Host "❌ File not found: $appFile"
    Write-Host "   Files in folder:"
    Get-ChildItem -Name
    Read-Host "Press Enter to exit"
    exit 1
}

# Check python and venv
if (!(Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "❌ Python not found on PATH."
    Write-Host "   Install Python 3.11+ and reopen this window."
    Read-Host "Press Enter to exit"
    exit 1
}

# Create venv if needed
if (!(Test-Path ".\venv")) {
    Write-Host "⚙️ Creating virtual environment..."
    python -m venv venv
}

# Activate venv
$activate = ".\venv\Scripts\Activate.ps1"
if (Test-Path $activate) {
    Write-Host "✅ Activating virtual environment..."
    & $activate
} else {
    Write-Host "❌ Could not find venv activate script at $activate"
    Read-Host "Press Enter to exit"
    exit 1
}

# Run the app (use python -m to avoid 'streamlit' not found issues)
Write-Host "🚀 Launching Streamlit app..."
python -m streamlit run $appFile

Write-Host ""
Write-Host "✅ If your browser doesn’t open automatically, go to: http://localhost:8501"
Read-Host "Press Enter to close this window"
