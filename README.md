# vb6games

Source code of Visual Basic 6 games, the compatibility corpus of
[WFC](https://github.com/Thewafflication/wfc), a reimplementation of the VB6
toolchain and runtime. WFC loads these projects as submodule test cases and
builds demonstration programs from them. The games were added unmodified (see
the first commits of `master`); they are now localized and given a modern,
touch-friendly look, as described below.

Each project lives in its own directory with a `PROVENANCE.md` that records the
source URL, retrieval date, archive checksum, the author's license and credit,
and every file that was left out. **Each project is licensed by its original
author, not by this repository.** Compiled binaries and archive junk
(`Thumbs.db`, FTP logs) are not kept. The changes made here (translations,
colors, layout) are listed below; the games' rules and logic are the authors'.

| Directory | Game | License |
| --- | --- | --- |
| `captureking/` | Capture King, a network chess game by Christopher McKirgan | GPL (see its `PROVENANCE.md`) |
| `connect4/` | Connect 4 with undo and redo by OlimilO1402 | GPL-3.0 |
| `clix/` | Clix, a Jawbreaker-style puzzle by Mili Hunjic | MIT |
| `executioner/` | Executioner, a Hangman game (1999) | GPL-3.0 |
| `pipes/` | Pipes, a pipe-connecting puzzle by Kym and Jonathan Wilson | GPL-3.0 |
| `restaum/` | Resta Um, peg solitaire by sourcecode-museum | MIT |
| `sumboku/` | SuMBOku, a sudoku game by OlimilO1402 | GPL-3.0 |

Source: the SourceForge project
[VB6 Game Archives](https://sourceforge.net/projects/vb6gamearchive/), whose
maintainer states that the code is open source, is not owned by them unless
noted, and that users must follow each original author's license and give
credit.

The five newer projects come from GitHub repositories (URL and commit in each
`PROVENANCE.md`); each repository carries its own license file.

## Localization

The games have English, Polish, Chinese and Spanish text and pick the language
from the Windows display language, with English as the default. See
[`localization/README.md`](localization/README.md).

## Modern look

Every game uses one dark navy and teal palette, raised cards with a thin border
and a soft shadow, and touch-sized buttons. The palette and the helpers are in
[`ui/modTheme.bas`](ui/modTheme.bas); each game carries a copy of that file
(`Module=modTheme` in its project), so a change is made in `ui/` and copied
over. A menu bar is replaced by buttons (the original menus stay, hidden, so
shortcuts and pop-ups still work).

The games are meant to run full screen at any resolution, for example with
WFC's `--fullscreen` option (640x480 up to 1920x1080 and beyond); WFC scales
the program's forms and turns mouse and touch input into the events the
programs expect. WFC release 0.8.0 with its full-screen commits is the
reference; the games also run in ordinary windows.

Menus are replaced by buttons that open floating panels of controls. No game
has a window border or a menu bar; every form fills the screen exactly. Text
is fitted to its control at run time (`ThemeFitText`), so no language or
resolution clips it.

Bitmap sprites come in several sizes (`sprites/32` ... `192`, made by
`python ui/make_sprites.py`, which needs Pillow) and the set nearest the
on-screen size is chosen at run time (`ThSprite` in `ui/modTheme.bas`). A GIF
carries the size the program lays it out at in a comment (`wfc:size=WxH`),
which WFC honors, so a sprite with more pixels is only resampled.

Some art is the original authors' (the Pipes splash picture, the pieces of
Capture King and Executioner, the sprites of Pipes and Resta Um) and is kept.
