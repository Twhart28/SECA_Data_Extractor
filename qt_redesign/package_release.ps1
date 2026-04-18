$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
$pyinstaller = Join-Path $repoRoot ".venv\Scripts\pyinstaller.exe"
$specPath = Join-Path $repoRoot "qt_redesign\seca_data_extractor.spec"
$distDir = Join-Path $repoRoot "dist"
$appDir = Join-Path $distDir "seca_data_extractor"
$buildDir = Join-Path $repoRoot "build\seca_data_extractor"
$zipPath = Join-Path $distDir "seca_data_extractor_portable.zip"
$installerScript = Join-Path $repoRoot "qt_redesign\seca_data_extractor.iss"
$installerOutput = Join-Path $distDir "seca_data_extractor_setup.exe"
$installerPattern = "seca_data_extractor_setup*.exe"

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

function Remove-PathIfExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$AllowFailure
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $true
    }

    try {
        Remove-Item -LiteralPath $Path -Recurse -Force
        return $true
    } catch {
        if ($AllowFailure) {
            Write-Warning "Could not remove $Path. $($_.Exception.Message)"
            return $false
        }
        throw
    }
}

if (-not (Test-Path $python)) {
    throw "Python was not found at $python"
}

if (-not (Test-Path $pyinstaller)) {
    throw "PyInstaller was not found at $pyinstaller"
}

Push-Location $repoRoot
try {
    Remove-PathIfExists -Path $appDir
    Remove-PathIfExists -Path $buildDir

    & $pyinstaller --noconfirm --clean $specPath

    Remove-PathIfExists -Path $zipPath
    Compress-Archive -Path $appDir -DestinationPath $zipPath -CompressionLevel Optimal

    $iscc = Find-InnoSetupCompiler
    if ($iscc) {
        $removedInstaller = Remove-PathIfExists -Path $installerOutput -AllowFailure
        if ($removedInstaller) {
            & $iscc $installerScript
        } else {
            $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
            $fallbackOutputBase = "seca_data_extractor_setup_$timestamp"
            Write-Warning "Previous installer output is locked. Building a timestamped installer instead."
            & $iscc "/F$fallbackOutputBase" $installerScript
        }
    } else {
        Write-Host "Inno Setup 6 was not found. Skipping installer build."
    }

    Write-Host ""
    Write-Host "Portable app folder: $appDir"
    Write-Host "Portable zip: $zipPath"
    $latestInstaller = Get-ChildItem -LiteralPath $distDir -Filter $installerPattern -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($latestInstaller) {
        Write-Host "Installer exe: $($latestInstaller.FullName)"
    }
} finally {
    Pop-Location
}
