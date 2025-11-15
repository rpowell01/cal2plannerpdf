<#
Automates fixing pre-commit errors caused by missing Python interpreters (e.g., Python39).
- Cleans pre-commit cache
- Deletes pre-commit cache folder
- Reinstalls pre-commit hooks
#>

Write-Host "Running: pre-commit clean"
pre-commit clean

$cache = Join-Path $env:USERPROFILE ".cache\pre-commit"
if (Test-Path $cache) {
    Write-Host "Deleting pre-commit cache at $cache"
    Remove-Item -Recurse -Force $cache
} else {
    Write-Host "No pre-commit cache directory found at $cache"
}

Write-Host "Reinstalling pre-commit hooks"
pre-commit install

Write-Host "Done. Try your git commit again. If you still see errors, restart your terminal and try again."
