<#
.SYNOPSIS
Builds the wfc-games-static and wfc-games-dynamic WPM packages.

.DESCRIPTION
Two stages:

  -Collect   Run on each architecture's own machine with that architecture's
             wfc.exe (-Wfc, from the wfc-<arch> release package): builds each
             game with `wfc build`, statically linked (one file) and
             dynamically linked (loads wfcrt.dll of the WFC installation),
             into <InputRoot>/<arch>/<game>.

  (default)  Assembles two packages from the collected programs of x86,
             x64 and arm64, each with one folder per game (its programs,
             PROVENANCE.md and license), a checksum list and the WFC license:
               wfc-games-static-any-<version>.zip   the statically linked
                   programs, which carry the WFC runtime and need nothing
               wfc-games-dynamic-any-<version>.zip  the dynamically linked
                   programs; the package depends on wfc (wfcrt.dll) and on
                   the library packages the games use (wfc-vbrun and the
                   controls of their .vbp files), at the versions of the WFC
                   release index (-WfcIndex)

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

.PARAMETER WfcIndex
index.json of that WFC release, for the versions of the dependencies.

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
    [string]$WfcIndex,
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
# The library package of each control a project names (Object= lines).
$controls = @{
    'mscomctl.ocx' = 'wfc-mscomctl'
    'mscomct2.ocx' = 'wfc-mscomctl2'
    'comdlg32.ocx' = 'wfc-comdlg32'
    'mswinsck.ocx' = 'wfc-mswinsck'
}

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
if (-not $WfcVersion -or -not $WfcIndex) {
    throw 'Assembling needs -WfcVersion and -WfcIndex.'
}
$revision = (git -C $root rev-parse HEAD).Trim()
New-Item -ItemType Directory -Force -Path $Output | Out-Null

# The dependencies of the dynamic programs: wfc, which carries wfcrt.dll,
# wfc-vbrun, and the library of each control the games use, at the versions
# the WFC release depends on.
$released = @{}
$index = Get-Content $WfcIndex -Raw | ConvertFrom-Json
foreach ($entry in @($index.packages)) {
    $released[$entry.name] = $entry.version
}
$needed = [System.Collections.Generic.SortedSet[string]]::new()
[void]$needed.Add('wfc-vbrun')
foreach ($game in $games) {
    $project = Join-Path $root $game.Project
    foreach ($line in Get-Content $project) {
        if ($line -match '^Object=.*;\s*(?<ocx>\S+)\s*$') {
            $library = $controls[$Matches.ocx.ToLowerInvariant()]
            if (-not $library) {
                throw "$($game.Name) uses $($Matches.ocx), which WFC lacks."
            }
            [void]$needed.Add($library)
        }
    }
}
$depends = @("wfc>=$WfcVersion")
foreach ($library in $needed) {
    if (-not $released.ContainsKey($library)) {
        throw "The WFC $WfcVersion index has no $library."
    }
    $depends += "$library>=$($released[$library])"
}

function Build-GamesPackage([string]$Link, [string[]]$Depends) {
    $name = "wfc-games-$Link"
    $suffix = if ($Link -eq 'static') { '' } else { '-dynamic' }
    $staging = Join-Path $root "out/wpm-staging/$name-any"
    if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $staging | Out-Null
    $sums = New-Object System.Collections.Generic.List[string]
    foreach ($game in $games) {
        $folder = Join-Path $staging $game.Name
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
        $gameRoot = Join-Path $root $game.Project.Split('/')[0]
        Copy-Item (Join-Path $gameRoot 'PROVENANCE.md') $folder
        foreach ($license in 'LICENSE', 'LICENSE.txt', 'LICENSE.md',
            'NOTICE') {
            $path = Join-Path $gameRoot $license
            if (Test-Path $path) { Copy-Item $path $folder }
        }
        foreach ($architecture in $architectures) {
            $file = "$($game.Name)-$architecture$suffix.exe"
            $built = Join-Path $InputRoot "$architecture/$($game.Name)/$file"
            if (-not (Test-Path $built)) { throw "Missing $built" }
            Copy-Item $built $folder
            $hash = (Get-FileHash (Join-Path $folder $file) `
                    -Algorithm SHA256).Hash.ToLowerInvariant()
            $sums.Add("$hash  $($game.Name)/$file")
        }
    }
    $sums | Set-Content (Join-Path $staging 'SHA256SUMS') -Encoding ascii
    Copy-Item $WfcLicense (Join-Path $staging 'WFC-LICENSE')
    Copy-Item (Join-Path $root 'README.md') $staging

    New-Item -ItemType Directory -Force -Path (Join-Path $staging '.wpm') |
        Out-Null
    # WPM runs install.cmd and remove.cmd from the unpacked package: they hand
    # over to Install-WfcGames.ps1, which copies the games to Program Files
    # (or the user's Programs folder) and adds the Start menu shortcuts.
    Copy-Item (Join-Path $PSScriptRoot 'Install-WfcGames.ps1') $staging
    $script = '"%~dp0..\Install-WfcGames.ps1"'
    $run = "powershell -NoProfile -ExecutionPolicy Bypass -File $script " +
        "-Source `"%~dp0..`" -Link $Link"
    Set-Content (Join-Path $staging '.wpm/install.cmd') -Encoding ascii `
        -Value @('@echo off', "$run || exit /b 1", 'exit /b 0')
    Set-Content (Join-Path $staging '.wpm/remove.cmd') -Encoding ascii `
        -Value @('@echo off', "$run -Uninstall || exit /b 1", 'exit /b 0')
    $how = if ($Link -eq 'static') {
        'statically linked, they need nothing else'
    } else {
        'dynamically linked to the installed WFC runtime'
    }
    @(
        "name=$name"
        "version=$Version"
        'arch=any'
        'debug=false'
        "description=VB6 games of the vb6games corpus ($how; WFC $WfcVersion)"
        'maintainer=Jordan Waughtal'
        'homepage=https://github.com/Thewafflication/vb6games'
        'repository=https://github.com/Thewafflication/vb6games'
        'license=see the LICENSE of each game'
        "source-version=$Version"
        "source-revision=$revision"
        foreach ($dependency in $Depends) { "depends=$dependency" }
    ) | Set-Content (Join-Path $staging '.wpm/package.txt') -Encoding ascii

    $arguments = @('build', $staging, $Output)
    if ($SigningKey) {
        $arguments += @('--sign', (Resolve-Path $SigningKey).Path)
    }
    & $Wpm @arguments
    if ($LASTEXITCODE -ne 0) { throw "wpm build failed for $name" }
    $package = Join-Path $Output "$name-any-$Version.zip"
    if (-not (Test-Path $package)) { throw "Missing package: $package" }
    $package
}

Build-GamesPackage 'static' @()
Build-GamesPackage 'dynamic' $depends
