#Requires -Version 5.1
param(
    [switch]$SkipDeps
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$VenvPython = Join-Path $Root ".venv\Scripts\python.exe"
$VenvPip = Join-Path $Root ".venv\Scripts\pip.exe"

Write-Host "Installing word-paste-image-fit into $Root"

if (-not (Test-Path $VenvPython)) {
    python -m venv (Join-Path $Root ".venv")
}

if (-not $SkipDeps) {
    & $VenvPip install -r (Join-Path $Root "requirements.txt")
}

& $VenvPython (Join-Path $Root "scripts\generate_sample_image.py")

$StartMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
$ShortcutPath = Join-Path $StartMenu "Word Paste Image Fit.lnk"
$Target = $VenvPython
$ShortcutArgs = "-m wpif"
$WorkDir = Join-Path $Root "src"

$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut($ShortcutPath)
$sc.TargetPath = $Target
$sc.Arguments = $ShortcutArgs
$sc.WorkingDirectory = $WorkDir
$sc.WindowStyle = 1
$sc.Description = "Word Paste Image Fit - live preview and smart paste"
$sc.Save()

Write-Host "Done."
Write-Host "Launch: $ShortcutPath"
Write-Host ("Or:     {0} -m wpif   (from {1})" -f $VenvPython, $WorkDir)
