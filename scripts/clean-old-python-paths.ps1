<#
Removes references to Python39 from user and system PATH, and clears PYTHONHOME/PYTHONPATH if they reference Python39.
Run as administrator to update system (Machine) environment variables.
#>

$oldPython = "Python39"
$changed = $false

function Remove-FromPath($scope) {
    $path = [System.Environment]::GetEnvironmentVariable("PATH", $scope)
    if ($path -and $path -like "*${oldPython}*") {
        $newPath = ($path -split ';' | Where-Object { $_ -notlike "*${oldPython}*" }) -join ';'
        [System.Environment]::SetEnvironmentVariable("PATH", $newPath, $scope)
        Write-Host "Removed $oldPython from $scope PATH."
        $GLOBALS:changed = $true
    } else {
        Write-Host "$scope PATH: no $oldPython found."
    }
}

function Clear-IfPythonVar($var, $scope) {
    $val = [System.Environment]::GetEnvironmentVariable($var, $scope)
    if ($val -and $val -like "*${oldPython}*") {
        [System.Environment]::SetEnvironmentVariable($var, $null, $scope)
        Write-Host "Cleared $var from $scope environment variables."
        $GLOBALS:changed = $true
    } else {
    Write-Host "${scope} ${var}: not set or does not reference ${oldPython}."
    }
}

Remove-FromPath "User"
Remove-FromPath "Machine"
Clear-IfPythonVar "PYTHONHOME" "User"
Clear-IfPythonVar "PYTHONHOME" "Machine"
Clear-IfPythonVar "PYTHONPATH" "User"
Clear-IfPythonVar "PYTHONPATH" "Machine"

if ($changed) {
    Write-Host "Environment updated. You may need to restart your terminal or log out/in for changes to take effect." -ForegroundColor Yellow
} else {
    Write-Host "No changes made."
}
