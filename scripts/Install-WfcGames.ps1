<#
.SYNOPSIS
Installs or removes the games of a wfc-games package.

.DESCRIPTION
Run by the package's .wpm/install.cmd and remove.cmd. As an administrator the
games go to Program Files\WFC Games\<Static|Dynamic> for every user, with
Start menu shortcuts in the Programs list; without administrator rights
they go to %LOCALAPPDATA%\Programs\WFC Games\<Static|Dynamic> with shortcuts
for the current user. Each shortcut starts the game built for this machine's
architecture (the x86 program on a machine that has no other).

.PARAMETER Source
The unpacked package: one folder per game.

.PARAMETER Link
static or dynamic: which package this is.

.PARAMETER Uninstall
Removes the games and their shortcuts.
#>
[CmdletBinding()]
param(
    [string]$Source = '',
    [Parameter(Mandatory)][ValidateSet('static', 'dynamic')][string]$Link,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
# Windows PowerShell 5.1 leaves $PSCommandPath empty in parameter defaults, so
# the folder of this script is taken here.
if (-not $Source) { $Source = $PSScriptRoot }

$names = [ordered]@{
    captureking = 'Capture King'
    clix = 'Clix'
    connect4 = 'Connect 4'
    executioner = 'Executioner'
    pipes = 'Pipes'
    restaum = 'Resta Um'
    sumboku = 'SuMBOku'
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$administrator = ([Security.Principal.WindowsPrincipal]$identity).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if ($administrator) {
    $programs = $env:ProgramW6432
    if (-not $programs) { $programs = $env:ProgramFiles }
    $start = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs'
} else {
    $programs = Join-Path $env:LOCALAPPDATA 'Programs'
    $start = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
}
$kind = if ($Link -eq 'static') { 'Static' } else { 'Dynamic' }
$games = Join-Path $programs 'WFC Games'
$target = Join-Path $games $kind
# The shortcuts sit directly in the Programs list; older packages put them in
# a "WFC Games" folder, which is cleaned up here.
$menu = $start
$oldMenu = Join-Path $start 'WFC Games'
$label = if ($Link -eq 'static') { '' } else { ' (dynamic)' }
$suffix = if ($Link -eq 'static') { '' } else { '-dynamic' }

function Remove-InstalledGame {
    foreach ($folder in $menu, $oldMenu) {
        foreach ($game in $names.Keys) {
            $shortcut = Join-Path $folder "$($names[$game])$label.lnk"
            if (Test-Path -LiteralPath $shortcut) {
                Remove-Item -LiteralPath $shortcut -Force
            }
        }
    }
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
    foreach ($folder in $oldMenu, $games) {
        if ((Test-Path -LiteralPath $folder) -and
            -not (Get-ChildItem -LiteralPath $folder -Force)) {
            Remove-Item -LiteralPath $folder -Force
        }
    }
}

if ($Uninstall) {
    Remove-InstalledGame
    Write-Output "Removed the $Link games from $target"
    return
}

Remove-InstalledGame
New-Item -ItemType Directory -Force -Path $target | Out-Null
foreach ($name in @($names.Keys) + 'SHA256SUMS', 'WFC-LICENSE',
    'README.md') {
    $path = Join-Path $Source $name
    if (Test-Path -LiteralPath $path) {
        Copy-Item -LiteralPath $path -Destination $target -Recurse -Force
    }
}

$machine = $env:PROCESSOR_ARCHITEW6432
if (-not $machine) { $machine = $env:PROCESSOR_ARCHITECTURE }
$architecture = switch ($machine) {
    'ARM64' { 'arm64' }
    'AMD64' { 'x64' }
    default { 'x86' }
}
$shell = New-Object -ComObject WScript.Shell
foreach ($game in $names.Keys) {
    $folder = Join-Path $target $game
    $program = Join-Path $folder "$game-$architecture$suffix.exe"
    if (-not (Test-Path -LiteralPath $program)) {
        $program = Join-Path $folder "$game-x86$suffix.exe"
    }
    $shortcut = $shell.CreateShortcut(
        (Join-Path $menu "$($names[$game])$label.lnk"))
    $shortcut.TargetPath = $program
    $shortcut.WorkingDirectory = $folder
    $shortcut.IconLocation = "$program,0"
    $shortcut.Description = "$($names[$game]), a VB6 game built with WFC"
    $shortcut.Save()
}
Write-Output "Installed the $Link games to $target"
