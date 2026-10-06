# json2dir-vbs

[json2dir](https://github.com/kitsunoff/awesome-json2dir) in VBScript (Windows Script Host), run with Wine's `cscript`; `["link", …]`/`["script", …]` are rejected (RFC §10) and existing entries are never replaced (Wine cannot detect symlinks).

    wine cscript //nologo json2dir.vbs < document.json

Under Wine, launch through a wrapper that pins `LC_ALL=C.UTF-8` and exports `WINE_UNIX_PID=$$` before `exec wine` (Wine's WSH has no `WScript.StdIn`/`StdErr`).
