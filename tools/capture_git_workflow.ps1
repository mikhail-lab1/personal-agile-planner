$ErrorActionPreference = "Stop"

$ProjectRoot = "C:\rating-work-git"
$ScreenshotDir = "C:\rating-work-screenshots"
$DocCommitMessage = "Document project requirements and lifecycle"

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

function Initialize-ScreenshotSupport {
    Add-Type -AssemblyName System.Drawing

    if (-not ("WindowCapture.NativeMethods" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace WindowCapture
{
    public struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out Rect rect);
    }
}
"@
    }
}

function Save-ActiveWindowScreenshot {
    param([string]$Path)

    $handle = [WindowCapture.NativeMethods]::GetForegroundWindow()
    $rect = New-Object WindowCapture.Rect
    [WindowCapture.NativeMethods]::GetWindowRect($handle, [ref]$rect) | Out-Null

    $width = [Math]::Max(1, $rect.Right - $rect.Left)
    $height = [Math]::Max(1, $rect.Bottom - $rect.Top)

    $bitmap = New-Object System.Drawing.Bitmap $width, $height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

    try {
        $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }

    Write-Host ""
    Write-Host "Screenshot saved: $Path"
}

function Invoke-CaptureBlock {
    param(
        [string]$Title,
        [string]$FileName,
        [scriptblock]$Action
    )

    Clear-Host
    Write-Host "============================================================"
    Write-Host $Title
    Write-Host "============================================================"
    Write-Host ""
    & $Action
    Write-Host ""
    Start-Sleep -Seconds 2

    $script:ScreenshotIndex++
    $path = Join-Path $ScreenshotDir ("{0:00}_{1}.png" -f $script:ScreenshotIndex, $FileName)
    Save-ActiveWindowScreenshot $path
    Start-Sleep -Seconds 1
}

if (-not (Test-Path -LiteralPath $ProjectRoot)) {
    throw "Project folder was not found: $ProjectRoot"
}

New-Item -ItemType Directory -Force -Path $ScreenshotDir | Out-Null
Initialize-ScreenshotSupport

Set-Location $ProjectRoot
$script:PythonExe = Resolve-PythonExecutable
$script:ScreenshotIndex = 0

Invoke-CaptureBlock "Step 1. Current branch and working tree status" "01_branch_and_status" {
    Get-Location
    git branch --show-current
    git status
}

Invoke-CaptureBlock "Step 2. Commit documentation changes if needed" "02_documentation_commit" {
    $docStatus = git status --porcelain -- README.md docs/project_description.md

    if ($docStatus) {
        Write-Host "Documentation changes found:"
        $docStatus
        Write-Host ""
        git add README.md docs/project_description.md
        git commit -m $DocCommitMessage -- README.md docs/project_description.md
    }
    else {
        Write-Host "No uncommitted documentation changes were found."
    }

    Write-Host ""
    git status --short
}

Invoke-CaptureBlock "Step 3. Git log" "03_git_log" {
    git log --oneline --decorate --graph --all -10
}

Invoke-CaptureBlock "Step 4. Git stash demonstration" "04_git_stash" {
    $demoFile = Join-Path $ProjectRoot ("docs\stash_demo_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    "Temporary file for git stash demonstration. Created at $(Get-Date -Format s)." | Set-Content -LiteralPath $demoFile -Encoding UTF8

    Write-Host "Status before stash:"
    git status --short -- $demoFile
    Write-Host ""

    git stash push -u -m "demo stash for report" -- $demoFile
    Write-Host ""
    git stash list -n 3
    Write-Host ""

    git stash pop
    if (Test-Path -LiteralPath $demoFile) {
        Remove-Item -LiteralPath $demoFile -Force
    }

    Write-Host ""
    Write-Host "Status after stash demonstration:"
    git status --short
}

Invoke-CaptureBlock "Step 5. Branch list" "05_git_branches" {
    git branch
}

Invoke-CaptureBlock "Step 6. Program help and safe run" "06_program_run" {
    Invoke-Python @("-m", "src.agile_planner.main", "--help")
    Write-Host ""
    Invoke-Python @("-m", "src.agile_planner.main", "list")
}

Invoke-CaptureBlock "Step 7. Unit tests" "07_unit_tests" {
    Invoke-Python @("-m", "unittest", "discover", "-s", "tests")
}

Clear-Host
Write-Host "Git workflow capture completed."
Write-Host "Screenshots folder: $ScreenshotDir"
Write-Host "No GitHub remote or push commands were executed."
