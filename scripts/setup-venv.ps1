<#
Creates or recreates a virtual environment in .venv using the newest python available
on the PATH or common Windows install locations, upgrades pip, and installs
packages from requirements.txt.

Usage:
  # Create venv (will error if .venv exists)
  .\scripts\setup-venv.ps1

  # Force recreate (backups current .venv to .venv_backup_YYYYMMDD_HHMMSS)
  .\scripts\setup-venv.ps1 -Force

#>
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$projectRoot = Get-Location
$venvPath = Join-Path $projectRoot '.venv'

function Find-PythonExe {
    # Try common locations first, then rely on 'python' in PATH
    $candidates = @(
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\python3.12.exe",
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\python.exe",
        'python3',
        'python'
    )

    foreach ($c in $candidates) {
        try {
            # capture version output (stderr or stdout depending on implementation)
            $ver = & $c --version 2>&1
            if ($LASTEXITCODE -eq 0 -or $ver) { return $c }
        } catch { }
    }
    return $null
}

if (Test-Path $venvPath) {
    if (-not $Force) {
        Write-Host ".venv already exists. To recreate and back it up run with -Force." -ForegroundColor Yellow
        Write-Host "If you want to recreate now use: .\scripts\setup-venv.ps1 -Force"
        exit 0
    } else {
        $ts = Get-Date -Format yyyyMMdd_HHmmss
        $backup = ".venv_backup_$ts"
        Rename-Item -LiteralPath $venvPath -NewName $backup
        Write-Host "Backed up existing venv to: $backup"
    }
}

$pythonExe = Find-PythonExe
if (-not $pythonExe) {
    Write-Host "No suitable Python interpreter found. Install Python 3.8+ and try again." -ForegroundColor Red
    exit 1
}

# Show discovered Python and its version for clarity
$pyver = (& $pythonExe --version 2>&1) -join ''
Write-Host "Creating virtual environment using: $pythonExe ($pyver)"
& $pythonExe -m venv $venvPath

Write-Host "Upgrading pip in the venv..."
& "$venvPath\Scripts\python.exe" -m pip install --upgrade pip

if (Test-Path "$projectRoot\requirements.txt") {
    Write-Host "Installing requirements from requirements.txt..."
    & "$venvPath\Scripts\python.exe" -m pip install -r "$projectRoot\requirements.txt"
} else {
    Write-Host "No requirements.txt found in project root — skipping package install." -ForegroundColor Yellow
}

Write-Host "Setup complete. Activate the venv with: .\\.venv\\Scripts\\Activate.ps1"
