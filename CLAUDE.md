# Project Rules

## Language
- PowerShell (loader) + C# (logic / UI)
- C# compiled at runtime via Add-Type (no build step)
- Outlook COM via late binding (dynamic), no reference required
- All source files must be ASCII / English only (no Japanese in code or comments)

## Architecture
- WinForms app (no Excel dependency)
- launch.bat > mailpull.ps1 > C# compile > App.Run()
- Config stored in config.json (app directory)

## Run
- Start: double-click `launch.bat`
- Dev: `powershell -ExecutionPolicy Bypass -File mailpull.ps1`
