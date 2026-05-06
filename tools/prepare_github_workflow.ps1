$ErrorActionPreference = "Stop"

$ProjectRoot = "C:\rating-work-git"
$ScreenshotDir = "C:\rating-work-screenshots"
$CloneRoot = Join-Path $ProjectRoot "cloned-repos"

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
    $path = Join-Path $ScreenshotDir ("github_{0:00}_{1}.png" -f $script:ScreenshotIndex, $FileName)
    Save-ActiveWindowScreenshot $path
    Start-Sleep -Seconds 1
}

function Push-BranchIfExists {
    param([string]$BranchName)

    git show-ref --verify --quiet "refs/heads/$BranchName"
    if ($LASTEXITCODE -eq 0) {
        git push origin $BranchName
    }
    else {
        Write-Host "Branch not found, skipped: $BranchName"
    }
}

if (-not (Test-Path -LiteralPath $ProjectRoot)) {
    throw "Project folder was not found: $ProjectRoot"
}

New-Item -ItemType Directory -Force -Path $ScreenshotDir | Out-Null
New-Item -ItemType Directory -Force -Path $CloneRoot | Out-Null
Initialize-ScreenshotSupport

Set-Location $ProjectRoot
$script:ScreenshotIndex = 0

$repoUrl = Read-Host "Enter public GitHub repository URL"
if ([string]::IsNullOrWhiteSpace($repoUrl)) {
    throw "Repository URL is required."
}

Clear-Host
Write-Host "GitHub workflow command plan"
Write-Host ""
Write-Host "The script will not push anything until you type YES."
Write-Host ""
Write-Host "Planned commands:"
Write-Host "git remote add origin $repoUrl"
Write-Host "git remote -v"
Write-Host "git push -u origin main"
Write-Host "git push origin develop"
Write-Host "git push origin wbs/01-initiation"
Write-Host "git push origin wbs/02-requirements"
Write-Host "git push origin wbs/03-design"
Write-Host "git push origin wbs/04-implementation"
Write-Host "git push origin wbs/05-testing"
Write-Host "git push origin wbs/06-release"
Write-Host "git clone $repoUrl"
Write-Host "git fetch"
Write-Host "git pull"
Write-Host ""

$confirm = Read-Host "Type YES to execute remote and push commands"
if ($confirm -ne "YES") {
    Write-Host "GitHub publishing was cancelled. No remote or push commands were executed."
    exit 0
}

Invoke-CaptureBlock "GitHub Step 1. Remote configuration" "remote_configuration" {
    $existingOrigin = git remote get-url origin 2>$null

    if ($LASTEXITCODE -eq 0 -and $existingOrigin) {
        Write-Host "Existing origin found:"
        $existingOrigin
        git remote set-url origin $repoUrl
    }
    else {
        git remote add origin $repoUrl
    }

    git remote -v
}

Invoke-CaptureBlock "GitHub Step 2. Push branches" "push_branches" {
    git push -u origin main
    Push-BranchIfExists "develop"
    Push-BranchIfExists "wbs/01-initiation"
    Push-BranchIfExists "wbs/02-requirements"
    Push-BranchIfExists "wbs/03-design"
    Push-BranchIfExists "wbs/04-implementation"
    Push-BranchIfExists "wbs/05-testing"
    Push-BranchIfExists "wbs/06-release"
}

Invoke-CaptureBlock "GitHub Step 3. Clone repository" "clone_repository" {
    $cloneDir = Join-Path $CloneRoot ("personal-agile-planner-{0}" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
    git clone $repoUrl $cloneDir
    Write-Host "Clone folder: $cloneDir"
}

Invoke-CaptureBlock "GitHub Step 4. Fetch and pull" "fetch_pull" {
    $latestClone = Get-ChildItem -LiteralPath $CloneRoot -Directory |
        Where-Object { $_.Name -like "personal-agile-planner-*" } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $latestClone) {
        throw "Clone folder was not found."
    }

    git -C $latestClone.FullName fetch
    git -C $latestClone.FullName pull
}

Clear-Host
Write-Host "GitHub workflow completed."
Write-Host "Screenshots folder: $ScreenshotDir"

