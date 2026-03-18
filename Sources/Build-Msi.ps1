[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$wixPath = Get-Command wix -ErrorAction SilentlyContinue
$buildAppScript = Join-Path $root "Build-AlerteSalaries.ps1"
$wixFile = Join-Path $root "wix\Product.wxs"
$publishDir = Join-Path $root "build\publish"
$deployDir = Join-Path $root "deploy"
$outputDir = Join-Path $root "build\msi"
$outputMsi = Join-Path $outputDir "AlerteSalaries.msi"

if (-not $wixPath) {
    throw "WiX v4 n'est pas installe. Installez WiX puis relancez ce script pour generer le MSI."
}

& $buildAppScript

New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
if (Test-Path $outputMsi) {
    Remove-Item -Path $outputMsi -Force
}

& $wixPath.Source build $wixFile -d PublishDir="$publishDir" -d DeployDir="$deployDir" -o $outputMsi

if ($LASTEXITCODE -ne 0 -or -not (Test-Path $outputMsi)) {
    throw "La generation du MSI a echoue."
}

Write-Host "MSI genere: $outputMsi"
