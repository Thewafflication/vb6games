# Executioner — provenance

| Item | Value |
| --- | --- |
| Game | Executioner, a Hangman-style game written in Visual Basic 6 in 1999 |
| Author | Corey McInroy (named on the splash form, "By Corey McInroy", and in the source comments, 1999); the repository owner is `creptic` |
| Source | https://github.com/creptic/executioner |
| Commit | `0f5e8fda15a8e16f0987f00c41907255dcf1c978` (2026-04-02) |
| Retrieved | 2026-10-12 by `git clone --depth 1` |
| License | GNU General Public License, version 3 (the repository's `LICENSE`) |
| Credit | Corey McInroy (original author); `creptic` published the repository |

## Contents

The repository's files are kept byte for byte (project `Executioner.vbp`, six
forms, a phrase list, GIF/JPG pictures and WAV sounds).

## Left out

Nothing.

## Use by WFC

WFC's acceptance tests load `Executioner.vbp` and capture screenshots of its
forms. The program is not modified; gaps found in it are fixed in WFC.

## Localization (branch `localization`)

The `localization` branch of this repository adds English, Polish, Chinese
(Simplified) and Spanish text to this game. It changes frmGameScreen.frm, frmOptions.frm, EnterName.frm, frmSplash.frm, frmrulez.frm and Executioner.vbp, and adds
`modLocale.bas`, which is generated from `localization/executioner.csv`. The
translated parts are the menus, buttons, labels, tool tips and game messages; the phrase list (`Data/Phrases.txt`) and the pictures that contain text are not translated. The language comes from the Windows display
language (English when it is none of the four; `VB6GAMES_LANG` overrides it).
The `master` branch keeps the program as the author wrote it.
