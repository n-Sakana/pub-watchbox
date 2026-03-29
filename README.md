# mailpull

Outlook mail archive and search tool. Pulls mail from Outlook via COM, saves to local folders, and provides a built-in viewer with search.

## Requirements

- Windows 10/11
- PowerShell 5.1
- Outlook (desktop app, running)

## Quick Start

1. Double-click `launch.vbs` (or run `launch.bat` for dev)
2. Click **Settings** to configure profiles
3. Click **Pull** to download mail

## Features

- **Multi-profile**: configure multiple account/folder/filter combinations
- **Keyword filter**: AND/OR match on subject, body, sender
- **Auto polling**: periodic pull on a configurable interval
- **3-pane viewer**: folder tree, mail list, body + attachments
- **Realtime search**: filter by subject, sender, body, attachment name
- **Flat output**: option to skip folder hierarchy
- **CSV import**: bulk-create profiles from CSV
- **Background pull**: runs on STA thread, UI stays responsive
- **Cancel**: stop a running pull mid-operation

## File Structure

```
mailpull.ps1          PowerShell loader (compiles C# at runtime)
launch.vbs            Start without terminal window
launch.bat            Start with terminal (dev)
config.json           Settings (auto-generated)
src/
  00_FolderPicker.cs  Modern folder dialog (IFileOpenDialog COM)
  01_App.cs           Entry point + Config (JSON key-value)
  02_Exporter.cs      Outlook COM + export + manifest + search
  03_MonitorForm.cs   Main window (Settings/Viewer/Pull/Auto)
  04_SettingsForm.cs  Profile editor + CSV import
  05_SearchForm.cs    3-pane mail viewer
```

## Export Format

Each mail is saved as a folder:

```
export_root/
  account/
    Inbox/
      20260328_195417_Subject/
        meta.json         Email metadata
        body.txt          Plain text body
        mail.msg          Original message
        attachment.pdf    Attachments
  manifest.csv            Index of all exported mail (BOM UTF-8)
```

## CSV Import Format

```csv
name,account,folder_path,since,filter_mode,filters,export_root,flat_output,poll_seconds
Project A,,\\user@co.jp\Inbox,2025-01-01,and,report;@client.com,C:\mail\projA,1,60
Project B,,\\user@co.jp\Inbox,2025-01-01,or,budget,C:\mail\projB,0,300
```

- `filters`: semicolon-separated keywords
- `filter_mode`: `and` or `or`
- `flat_output`: `1` = no folder hierarchy, `0` = preserve structure
