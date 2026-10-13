# Connect 4 — provenance

| Item | Value |
| --- | --- |
| Game | Connect 4 (Vier Gewinnt) with undo and redo, written in Visual Basic 6 |
| Author | OlimilO1402 (the repository owner) |
| Source | https://github.com/OlimilO1402/Game_Connect4 |
| Commit | `98585759a5f7b2cecf7bbc5cb0d3bfa6dfa916ad` (2025-10-05) |
| Retrieved | 2026-10-12 by `git clone --depth 1` |
| License | GNU General Public License, version 3 (the repository's `LICENSE`) |
| Credit | OlimilO1402 |

## Contents

The repository's files are kept byte for byte. The project `PConnect4.vbp`
names the class `ColorDialog` in `..\Win_Dialogs\Classes\ColorDialog.cls`,
from the author's separate `Win_Dialogs` repository, which is not in this
corpus. WFC supplies `ColorDialog` (and `FontDialog`, `OpenFileDialog` and
`SaveFileDialog`) itself over its native dialogs when a project names them and
the file is missing.

## Left out

Nothing.

## Use by WFC

WFC's acceptance tests load `PConnect4.vbp` and capture a screenshot of its
form. The program is not modified; gaps found in it are fixed in WFC.
