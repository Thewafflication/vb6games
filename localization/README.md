# Localization

This branch (`localization`) adds English, Polish, Chinese (Simplified) and
Spanish text to the games. English is the default.

## How a game chooses its language

Each game has a generated module, `modLocale.bas`, with three functions:

| Function | Result |
| --- | --- |
| `Loc("key")` | the text of a key in the language in use |
| `LocF("key", a, b)` | the same, with `{0}`, `{1}` replaced by the arguments |
| `LocLanguage` | `"en"`, `"pl"`, `"zh"` or `"es"` |

The language is the Windows display language (`GetUserDefaultUILanguage`).
Anything other than Polish, Chinese or Spanish gives English. To force a
language, set the environment variable `VB6GAMES_LANG` to `en`, `pl`, `zh` or
`es` before starting the game.

Forms keep their designed (English) captions. In another language, each form's
`ApplyLocale` sets the captions from the table when the form loads. In
English it does nothing, so English output is the original's (the exceptions
are the parts the author wrote in Croatian, German or Portuguese, which are now
English translations).

## Changing a translation

The text lives in `localization/<game>.csv` (UTF-8, columns
`key,en,pl,zh,es`). `{nl}` is a line break and `{tab}` a tab. After editing:

```
python localization/generate.py
```

`generate.py` writes each game's `modLocale.bas`. The generated files are plain
ASCII, because VB6 source is ANSI; non-ASCII text is stored as UTF-16 code
units that `ChrW` turns back into characters at run time.

## Limits

- The VB6 intrinsic controls are ANSI. On a real Windows system whose
  non-Unicode ("system locale") code page does not cover Polish or Chinese,
  those characters can show as `?`. WFC's controls are Unicode.
- Labels that `AutoSize` in VB6 are not resized by WFC when the caption
  changes, so translations are kept short enough for the designed layouts.
- Text inside pictures, `.frx` resources, help files and the Executioner phrase
  list is not translated.
- Text that Capture King players send to each other over the network is not
  translated.
