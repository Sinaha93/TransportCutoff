[CmdletBinding()]
param(
    [switch]$RunLauncherTests,
    [switch]$RunTests
)

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$previousLocation = Get-Location

try {
    Set-Location $projectRoot

    if ($RunLauncherTests -and $RunTests) {
        throw "-RunLauncherTests and -RunTests cannot be used together."
    }
    if ($RunTests) {
        & python -m pytest -q
        if ($LASTEXITCODE -ne 0) { throw "Test suite failed." }
    }
    elseif ($RunLauncherTests) {
        & python -m pytest tests/unit/test_launcher.py -q
        if ($LASTEXITCODE -ne 0) { throw "Launcher tests failed." }
    }

    & python -m PyInstaller `
        --noconfirm `
        --clean `
        --onedir `
        --name TransportReport `
        --paths $projectRoot `
        app/launcher.py
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed." }

    $bundleRoot = Join-Path $projectRoot "dist\TransportReport"
    $internalRoot = Join-Path $bundleRoot "_internal"

    $dataCopies = @(
        @{ Source = "app\web\templates"; Destination = "app\web\templates" },
        @{ Source = "app\web\static"; Destination = "app\web\static" },
        @{ Source = "app\db\migrations"; Destination = "app\db\migrations" }
    )
    foreach ($copy in $dataCopies) {
        $source = Join-Path $projectRoot $copy.Source
        $destination = Join-Path $internalRoot $copy.Destination
        New-Item -ItemType Directory -Force -Path (Split-Path $destination) | Out-Null
        Copy-Item -LiteralPath $source -Destination $destination -Recurse -Force
    }

    $assets = Join-Path $bundleRoot "assets"
    New-Item -ItemType Directory -Force -Path $assets | Out-Null
    Copy-Item -LiteralPath (Join-Path $projectRoot "assets\report_template.pptx") `
        -Destination (Join-Path $assets "report_template.pptx") -Force

    foreach ($directory in @("data", "imports", "backups", "outputs")) {
        New-Item -ItemType Directory -Force -Path (Join-Path $bundleRoot "runtime\$directory") | Out-Null
    }

    $executable = Join-Path $bundleRoot "TransportReport.exe"
    if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
        throw "Portable executable was not created: $executable"
    }
    Write-Output $executable
}
finally {
    Set-Location $previousLocation
}
