param(
    [switch]$OneDir,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $projectRoot ".venv\\Scripts\\python.exe"

if (Test-Path $venvPython) {
    $pythonExe = $venvPython
} else {
    $pythonExe = "python"
}

Push-Location $projectRoot
try {
    if (-not $SkipInstall) {
        & $pythonExe -m pip install --upgrade pip pyinstaller
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to install PyInstaller."
        }
    }

    $modeArg = "--onefile"
    if ($OneDir) {
        $modeArg = "--onedir"
    }

    & $pythonExe -m PyInstaller `
        --noconfirm `
        --clean `
        --windowed `
        --name VocabHelper `
        $modeArg `
        --collect-data pykakasi `
        run_vocab_helper.py

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller build failed."
    }

    if ($OneDir) {
        Write-Host "Build complete: dist\\VocabHelper\\VocabHelper.exe"
    } else {
        Write-Host "Build complete: dist\\VocabHelper.exe"
    }
} finally {
    Pop-Location
}
