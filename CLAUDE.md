# Project Rules

## Language
- PowerShell (loader) + C# (logic / UI)
- C# compiled at runtime via Add-Type (no build step)
- Outlook COM via late binding (dynamic), no reference required
- All source files must be ASCII / English only (no Japanese in code or comments)

## Architecture
- WPF app (no Excel dependency)
- launch.bat > watchbox.ps1 > C# compile > App.Run()
- Config stored in config.json (app directory)
- Profile types: mail (Outlook export), folder (with source_folder = file sync, without = manifest-only)

## Run
- Start: double-click `launch.bat`
- Dev: `powershell -ExecutionPolicy Bypass -File watchbox.ps1`
