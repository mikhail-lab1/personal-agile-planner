$ErrorActionPreference = "Stop"

$ProjectRoot = "C:\rating-work-git"
$LogPath = Join-Path $ProjectRoot "docs\execution_log.txt"

function Resolve-PythonExecutable {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return $python.Source
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return $py.Source
    }

    throw "Python 3 was not found. Install Python or add it to PATH."
}

function Invoke-Python {
    param([string[]]$Arguments)

    if ((Split-Path -Leaf $script:PythonExe) -ieq "py.exe") {
        & $script:PythonExe -3 @Arguments
    }
    else {
        & $script:PythonExe @Arguments
    }
}

if (-not (Test-Path -LiteralPath $ProjectRoot)) {
    throw "Project folder was not found: $ProjectRoot"
}

Set-Location $ProjectRoot
$script:PythonExe = Resolve-PythonExecutable

New-Item -ItemType Directory -Force -Path (Split-Path $LogPath) | Out-Null

if (Test-Path -LiteralPath $LogPath) {
    Remove-Item -LiteralPath $LogPath -Force
}

Start-Transcript -Path $LogPath -Force | Out-Null

try {
    Write-Host "=== Personal Agile Planner: project checks ==="
    Write-Host "Project folder: $ProjectRoot"
    Write-Host "Python executable: $script:PythonExe"
    Write-Host ""

    Write-Host "=== Current folder ==="
    Get-Location
    Write-Host ""

    Write-Host "=== Git status ==="
    git status
    Write-Host ""

    Write-Host "=== Program help ==="
    Invoke-Python @("-m", "src.agile_planner.main", "--help")
    Write-Host ""

    Write-Host "=== Program run: task list ==="
    Invoke-Python @("-m", "src.agile_planner.main", "list")
    Write-Host ""

    Write-Host "=== Unit tests ==="
    Invoke-Python @("-m", "unittest", "discover", "-s", "tests")
    Write-Host ""

    Write-Host "Checks completed. Log saved to: $LogPath"
}
finally {
    Stop-Transcript | Out-Null
}
