# Capture King — provenance

| Item | Value |
| --- | --- |
| Game | Capture King, a network-based chess game written in Visual Basic 6.0 |
| Author | Christopher McKirgan (Bridgwater College project, 2007 to 2009) |
| Source | https://sourceforge.net/projects/vb6gamearchive/files/Capture%20King/captureking.zip |
| Retrieved | 2026-10-12 |
| Archive size | 1,780,575 bytes |
| Archive SHA-256 | `8a34ae73e99b1693fe4386bc081ec8a7df9d41a6fee371c5840040cb6623fb16` |
| License | GPL; the archive's readme says the software is free to download, distribute and sell under the GPL with the one condition that Christopher McKirgan is referenced alongside its works. The readme does not name a GPL version. |
| Credit | (c) Christopher McKirgan 2007, Capture King (shown on its splash screen) |

## Contents

The archive's files are kept byte for byte, except for the omissions below.
`CaptureKing.vbp` names `comdlg32.ocx` and `mswinsck.ocx`; WFC provides native
replacements for those two controls.

## Left out

- `CaptureKing.exe`, a compiled binary from the archive (never run by WFC's
  maintainers; WFC builds the program from source).
- `Thumbs.db` and `WS_FTP.LOG` files (Windows and FTP client leftovers).
- The archive's `readme.txt` is on the SourceForge page; it is summarized above
  instead of copied.

## Use by WFC

WFC's acceptance tests load `CaptureKing.vbp` and capture screenshots of its
forms. Gaps found in the program are fixed in WFC, not in the program.

## Modification

One change was made at the maintainer's request, to the layout properties of
`CaptureKing.frm` only (commit after `c0434f0`; the unmodified source is commit
`10282de`): the developer buttons and hover-image holders that the author left
visible on the game form are hidden (the picture buttons call their handlers),
Surrender joins the row of buttons under the board, and the status line moves
below it, so the form looks like the screenshots in the user guide. No code
was changed.

## Localization (branch `localization`)

The `localization` branch of this repository adds English, Polish, Chinese
(Simplified) and Spanish text to this game. It changes CaptureKing.frm, frmChat.frm, frmWelcome.frm, frmPawnPromotion.frm, frmStats.frm, frmAchievments.frm and CaptureKing.vbp, and adds
`modLocale.bas`, which is generated from `localization/captureking.csv`. The
translated parts are the menus, buttons, labels, message boxes and the statistics text; the text that players send to each other over the network is not translated, so that players using different languages stay compatible. The language comes from the Windows display
language (English when it is none of the four; `VB6GAMES_LANG` overrides it).
The `master` branch keeps the program as the author wrote it.
