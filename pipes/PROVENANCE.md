# Pipes — provenance

| Item | Value |
| --- | --- |
| Game | Pipes version 1.0, a pipe-connecting puzzle written in Visual Basic |
| Author | Kym and Jonathan Wilson (named on the program's logo form) |
| Source | https://github.com/jonwil/pipes |
| Commit | `df4a8acfdb1c222d28816bf5bb639d8222c37900` (2026-09-23) |
| Retrieved | 2026-10-12 by `git clone --depth 1` |
| License | GNU General Public License, version 3 (the repository's `LICENSE`) |
| Credit | Kym and Jonathan Wilson |

## Contents

The repository's files are kept byte for byte. The repository has no project
file, only `FORM1.FRM`, `LOGO.FRM`, `MODULE1.BAS` and icons; the project file
`pipes.vbp` in this directory was written for WFC (it lists those files and
starts at the logo form) and is the only file not from the repository.

## Left out

Nothing.

## Use by WFC

WFC's acceptance tests load `pipes.vbp` and capture screenshots of its forms.
The program's own files are not modified; gaps found in it are fixed in WFC.

## Localization (branch `localization`)

The `localization` branch of this repository adds English, Polish, Chinese
(Simplified) and Spanish text to this game. It changes FORM1.FRM, LOGO.FRM and PIPES.VBP, and adds
`modLocale.bas`, which is generated from `localization/pipes.csv`. The
translated parts are the window titles, the Quit and OK buttons, the version label and the win and lose messages. The language comes from the Windows display
language (English when it is none of the four; `VB6GAMES_LANG` overrides it).
The `master` branch keeps the program as the author wrote it.
