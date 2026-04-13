$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
$pyinstaller = Join-Path $repoRoot ".venv\Scripts\pyinstaller.exe"
$specPath = Join-Path $repoRoot "qt_redesign\seca_qt_converter.spec"
$distDir = Join-Path $repoRoot "dist"
$appDir = Join-Path $distDir "seca_qt_converter"
$zipPath = Join-Path $distDir "seca_qt_converter_portable.zip"
$installerScript = Join-Path $repoRoot "qt_redesign\seca_qt_converter.iss"
$installerOutput = Join-Path $distDir "seca_qt_converter_setup.exe"

function Find-InnoSetupCompiler {
    $candidates = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

if (-not (Test-Path $python)) {
    throw "Python was not found at $python"
}

if (-not (Test-Path $pyinstaller)) {
    throw "PyInstaller was not found at $pyinstaller"
}

Push-Location $repoRoot
try {
    & $pyinstaller --noconfirm $specPath

    if (Test-Path $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    Compress-Archive -Path $appDir -DestinationPath $zipPath -CompressionLevel Optimal

    $iscc = Find-InnoSetupCompiler
    if ($iscc) {
        if (Test-Path $installerOutput) {
            Remove-Item -LiteralPath $installerOutput -Force
        }
        & $iscc $installerScript
    } else {
        Write-Host "Inno Setup 6 was not found. Skipping installer build."
    }

    Write-Host ""
    Write-Host "Portable app folder: $appDir"
    Write-Host "Portable zip: $zipPath"
    if (Test-Path $installerOutput) {
        Write-Host "Installer exe: $installerOutput"
    }
} finally {
    Pop-Location
}
