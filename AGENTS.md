# AGENTS.md - pub-watchbox

Entry note for Codex. For repo-specific details, see [README.md](README.md) and [CLAUDE.md](CLAUDE.md). For cross-cutting topology, see [fin/hub/ARCHITECTURE.md](../../fin/hub/ARCHITECTURE.md).

## Role

- Windows tool that monitors/collects mail / folder and produces `manifest.csv` and `log.csv`
- From `pub/casedesk`'s perspective, the core of external collection and manifest generation

## runtime / connectivity

- runtime: local Windows only
- loader: PowerShell
- logic / UI: runtime-compiled C# WPF
- downstream contract: `manifest.csv`, `log.csv`

## Where to look first

- `watchbox.ps1` - loader
- `src/02_ManifestIO.cs` - manifest I/O
- `src/05_MailScanner.cs` - Outlook mail collection
- `src/06_FolderScanner.cs` - folder collection
- `src/07_ProfileRunner.cs`, `src/07a_EventWatcher.cs` - Pull / Watch execution
- `src/08_MonitorForm.cs`, `src/09_SettingsForm.cs` - UI

## Launch

```bat
launch.bat
```

```powershell
powershell -ExecutionPolicy Bypass -File watchbox.ps1
```

## Guardrails

- Keep code and comments ASCII / English only
- `config.json` is the source of truth
- If you change the `manifest.csv` header, also check `pub/casedesk` at the same time
- Keep the WPF app loosely coupled with the Outlook / folder scanner
