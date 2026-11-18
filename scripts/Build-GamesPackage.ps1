<#
.SYNOPSIS
Builds the wfc-games WPM package: every game as a Windows program.

.DESCRIPTION
Two stages:

  -Collect   Run on each architecture's own machine with that architecture's
             wfc.exe (-Wfc, from the wfc-<arch> release package): builds each
             game with `wfc build`, statically linked (one file) and
             dynamically linked (loads wfcrt.dll of the WFC installation),
             into <InputRoot>/<arch>/<game>.

  (default)  Assembles wfc-games-any-<version>.zip from the collected
             programs of x86, x64 and arm64: one folder per game with its
             programs, PROVENANCE.md and license, a checksum list, and the
             WFC license (the programs carry the WFC runtime).

Nothing is signed unless -SigningKey is given, and nothing is published.

.PARAMETER Version
Package version (the release tag).

.PARAMETER Collect
Collect stage for -Architecture.

.PARAMETER Architecture
The architecture of -Wfc, for the collect stage.

.PARAMETER Wfc
Path of wfc.exe, for the collect stage.

.PARAMETER WfcLicense
The LICENSE file of WFC, for the assemble stage.

.PARAMETER WfcVersion
The WFC version that built the programs, recorded in the package.

.PARAMETER InputRoot
Where the collect stage writes and the assemble stage reads.

.PARAMETER Output
Folder for the package.

.PARAMETER Wpm
Path of wpm.exe.

.PARAMETER SigningKey
Optional WPM private key file; the package is unsigned without it.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [switch]$Collect,
    [string]$Architecture,
    [string]$Wfc,
    [string]$WfcLicense,
    [string]$WfcVersion = '',
    [string]$InputRoot = 'out/package-input',
    [string]$Output = 'out/packages',
    [string]$Wpm = 'wpm.exe',
    [string]$SigningKey
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$games = @(
    @{ Name = 'captureking'; Project = 'captureking/CaptureKing.vbp' }
    @{ Name = 'clix'; Project = 'clix/src/Clix.vbp' }
    @{ Name = 'connect4'; Project = 'connect4/PConnect4.vbp' }
    @{ Name = 'executioner'; Project = 'executioner/Executioner.vbp' }
    @{ Name = 'pipes'; Project = 'pipes/pipes.vbp' }
    @{ Name = 'restaum'; Project = 'restaum/source/BrainVita.vbp' }
    @{ Name = 'sumboku'; Project = 'sumboku/v1.4.2/SuMBOku_V_1_4/Sudoku.vbp' }
)
$architectures = 'x86', 'x64', 'arm64'

if ($Collect) {
    if (-not $Architecture -or -not $Wfc) {
        throw 'Collecting needs -Architecture and -Wfc.'
    }
    $Wfc = (Resolve-Path $Wfc).Path
    foreach ($game in $games) {
        $source = Join-Path $root $game.Project
        $folder = Join-Path $InputRoot "$Architecture/$($game.Name)"
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
        foreach ($link in 'static', 'dynamic') {
            $suffix = if ($link -eq 'static') { '' } else { '-dynamic' }
            $program = Join-Path $folder `
                "$($game.Name)-$Architecture$suffix.exe"
            & $Wfc --build $source -o $program --gui --link $link
            if ($LASTEXITCODE -ne 0) {
                throw "wfc build ($link) failed for $($game.Name) " +
                    "($Architecture)"
            }
        }
    }
    return
}

if (-not $Version) { throw 'Assembling needs -Version.' }
if (-not $WfcLicense) { throw 'Assembling needs -WfcLicense.' }
$revision = (git -C $root rev-parse HEAD).Trim()
$staging = Join-Path $root 'out/wpm-staging/wfc-games-any'
if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Force -Path $staging, $Output | Out-Null

$sums = New-Object System.Collections.Generic.List[string]
foreach ($game in $games) {
    $folder = Join-Path $staging $game.Name
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    $gameRoot = Join-Path $root $game.Project.Split('/')[0]
    Copy-Item (Join-Path $gameRoot 'PROVENANCE.md') $folder
    foreach ($license in 'LICENSE', 'LICENSE.txt', 'LICENSE.md', 'NOTICE') {
        $path = Join-Path $gameRoot $license
        if (Test-Path $path) { Copy-Item $path $folder }
    }
    foreach ($architecture in $architectures) {
        foreach ($suffix in '', '-dynamic') {
            $name = "$($game.Name)-$architecture$suffix.exe"
            $built = Join-Path $InputRoot "$architecture/$($game.Name)/$name"
            if (-not (Test-Path $built)) { throw "Missing $built" }
            Copy-Item $built $folder
            $hash = (Get-FileHash (Join-Path $folder $name) `
                    -Algorithm SHA256).Hash.ToLowerInvariant()
            $sums.Add("$hash  $($game.Name)/$name")
        }
    }
}
$sums | Set-Content (Join-Path $staging 'SHA256SUMS') -Encoding ascii
Copy-Item $WfcLicense (Join-Path $staging 'WFC-LICENSE')
Copy-Item (Join-Path $root 'README.md') $staging

New-Item -ItemType Directory -Force -Path (Join-Path $staging '.wpm') |
    Out-Null
foreach ($name in 'install.cmd', 'remove.cmd') {
    Set-Content (Join-Path $staging ".wpm/$name") -Encoding ascii -Value @(
        '@echo off'
        'exit /b 0'
    )
}
$built = if ($WfcVersion) { " (built with WFC $WfcVersion)" } else { '' }
@(
    'name=wfc-games'
    "version=$Version"
    'arch=any'
    'debug=false'
    "description=VB6 games of the vb6games corpus as Windows programs$built"
    'maintainer=Jordan Waughtal'
    'homepage=https://github.com/Thewafflication/vb6games'
    'repository=https://github.com/Thewafflication/vb6games'
    'license=see the LICENSE of each game'
    "source-version=$Version"
    "source-revision=$revision"
) | Set-Content (Join-Path $staging '.wpm/package.txt') -Encoding ascii

$arguments = @('build', $staging, $Output)
if ($SigningKey) {
    $arguments += @('--sign', (Resolve-Path $SigningKey).Path)
}
& $Wpm @arguments
if ($LASTEXITCODE -ne 0) { throw 'wpm build failed for wfc-games' }
$package = Join-Path $Output "wfc-games-any-$Version.zip"
if (-not (Test-Path $package)) { throw "Missing package: $package" }
$package
