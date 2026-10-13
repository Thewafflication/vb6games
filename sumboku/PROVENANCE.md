# SuMBOku — provenance

| Item | Value |
| --- | --- |
| Game | SuMBOku, a sudoku game written in Visual Basic 6 (2004 to 2006) |
| Author | OlimilO1402 (the repository owner) |
| Source | https://github.com/OlimilO1402/Game_SuMBOku |
| Commit | `2fc43e2b2805752b354466b42958b59737b17292` (2023-01-29) |
| Retrieved | 2026-10-12 by `git clone --depth 1` |
| License | GNU General Public License, version 3 (the repository's `LICENSE`) |
| Credit | OlimilO1402 |

## Contents

The repository's files are kept byte for byte, except for the omission below.
The top-level project `PSuMBOku.vbp` needs the classes `OpenFileDialog` and
`SaveFileDialog` from the separate `Win_Dialogs` repository, which is not in
this corpus. The archived version in `v1.4.2/SuMBOku_V_1_4/` carries its own
copies, and WFC's tests use its `Sudoku.vbp`.

## Left out

- `v1.4.2/SuMBOku_V_1_4/SuMBOku.GID`, a Windows help cache file.

## Use by WFC

WFC's acceptance tests load `v1.4.2/SuMBOku_V_1_4/Sudoku.vbp` and capture
screenshots of its forms. The program is not modified; gaps found in it are
fixed in WFC.
