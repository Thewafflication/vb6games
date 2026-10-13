# Clix — provenance

| Item | Value |
| --- | --- |
| Game | Clix, a Visual Basic 6 clone of Jawbreaker (a SameGame-style puzzle) |
| Author | Mili Hunjic (written c. 2006, aged 21) |
| Source | https://github.com/milihwork/clix-vb6 |
| Commit | `5baa61b79d067a3d5fb782feee290bccfb0c3389` (2026-04-03) |
| Retrieved | 2026-10-12 by `git clone --depth 1` |
| License | MIT, Copyright (c) 2006-2026 Mili Hunjic (`LICENSE`) |
| Credit | Mili Hunjic; `NOTICE` records that Jawbreaker and related names belong to their owners |

## Contents

The repository's files are kept byte for byte, except for the omission below.
The program is `src/Clix.vbp` (three forms and one module).

## Left out

- `bin/Clix.exe`, a compiled binary from the repository (never run by WFC's
  maintainers; WFC builds the program from source).

## Use by WFC

WFC's acceptance tests load `src/Clix.vbp` and capture screenshots of its
forms. The program is not modified; gaps found in it are fixed in WFC.
