[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$srcFile = Join-Path $projectRoot "src\AlerteSalaries.cs"
$iconBuildScript = Join-Path $projectRoot "Build-Icon.ps1"
$iconPath = Join-Path $projectRoot "assets\warning.ico"
$deployDir = Join-Path $projectRoot "deploy"
$buildDir = Join-Path $projectRoot "build"
$publishDir = Join-Path $buildDir "publish"
$compiler = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if (-not (Test-Path $srcFile)) {
    throw "Le fichier source '$srcFile' est introuvable."
}

if (-not (Test-Path $compiler)) {
    throw "Le compilateur C# '$compiler' est introuvable."
}

& $iconBuildScript

if (-not (Test-Path $iconPath)) {
    throw "L'icone '$iconPath' est introuvable."
}

New-Item -Path $buildDir -ItemType Directory -Force | Out-Null
New-Item -Path $publishDir -ItemType Directory -Force | Out-Null

$outputExe = Join-Path $publishDir "AlerteSalaries.exe"

& $compiler `
    /nologo `
    /target:winexe `
    /out:$outputExe `
    /win32icon:$iconPath `
    /reference:System.dll `
    /reference:System.Drawing.dll `
    /reference:System.Runtime.Serialization.dll `
    /reference:System.Windows.Forms.dll `
    $srcFile

if ($LASTEXITCODE -ne 0 -or -not (Test-Path $outputExe)) {
    throw "La compilation de l'application a echoue."
}

Copy-Item -Path (Join-Path $projectRoot "README.md") -Destination (Join-Path $publishDir "README.md") -Force
Copy-Item -Path $iconPath -Destination (Join-Path $publishDir "warning.ico") -Force
if (Test-Path $deployDir) {
    Copy-Item -Path (Join-Path $deployDir "*") -Destination $publishDir -Recurse -Force
}

Write-Host "Compilation terminee: $outputExe"
