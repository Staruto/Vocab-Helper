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
        --name VocabHelperGUI `
        $modeArg `
        --collect-data pykakasi `
        run_vocab_helper_gui.py

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller GUI build failed."
    }

    & $pythonExe -m PyInstaller `
        --noconfirm `
        --clean `
        --console `
        --name VocabHelperCLI `
        $modeArg `
        --collect-data pykakasi `
        run_vocab_helper_cli.py

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller CLI build failed."
    }

    if ($OneDir) {
        Write-Host "Build complete: dist\\VocabHelperGUI\\VocabHelperGUI.exe"
        Write-Host "Build complete: dist\\VocabHelperCLI\\VocabHelperCLI.exe"
    } else {
        Write-Host "Build complete: dist\\VocabHelperGUI.exe"
        Write-Host "Build complete: dist\\VocabHelperCLI.exe"
    }
} finally {
    Pop-Location
}
