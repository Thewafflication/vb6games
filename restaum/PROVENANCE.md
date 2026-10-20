# Resta Um (BrainVita) — provenance

| Item | Value |
| --- | --- |
| Game | Resta Um, a peg-solitaire game in Visual Basic 6 (the README says its logic is thirteen lines) |
| Author | `sourcecode-museum` (the repository's copyright holder) |
| Source | https://github.com/sourcecode-museum/RestaUm |
| Commit | `b70055391afd965ae850d559bcd2a89f65e962e1` (2023-01-31) |
| Retrieved | 2026-10-12 by `git clone --depth 1` |
| License | MIT, Copyright (c) 2020 sourcecode-museum (`LICENSE`) |
| Credit | sourcecode-museum |

## Contents

The repository's files are kept byte for byte, except for the omission below.
The program is `source/BrainVita.vbp` (one form).

## Left out

- `BrainVita.exe`, a compiled binary from the repository (never run by WFC's
  maintainers; WFC builds the program from source).

## Use by WFC

WFC's acceptance tests load `source/BrainVita.vbp` and capture screenshots of
its form. The program is not modified; gaps found in it are fixed in WFC.

## Localization (branch `localization`)

The `localization` branch of this repository adds English, Polish, Chinese
(Simplified) and Spanish text to this game. It changes source/frmBoard.frm and source/BrainVita.vbp, and adds
`modLocale.bas`, which is generated from `localization/restaum.csv`. The
translated parts are the window title and the instructions (the original instructions were Portuguese, so English is a translation too). The language comes from the Windows display
language (English when it is none of the four; `VB6GAMES_LANG` overrides it).
The `master` branch keeps the program as the author wrote it.
