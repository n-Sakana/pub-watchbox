# AGENTS.md - pub-watchbox

Codex 用の入口メモです。repo 固有の詳細は [README.md](README.md) と [CLAUDE.md](CLAUDE.md)、横断トポロジは [fin/hub/ARCHITECTURE.md](../../fin/hub/ARCHITECTURE.md) を参照してください。

## 役割

- mail / folder を監視・収集し、`manifest.csv` と `log.csv` を生成する Windows ツール
- `pub/casedesk` から見ると外部収集と manifest 生成の中核

## runtime / 接続

- runtime: local Windows only
- loader: PowerShell
- logic / UI: runtime-compiled C# WPF
- downstream contract: `manifest.csv`, `log.csv`

## まず見る場所

- `watchbox.ps1` - loader
- `src/02_ManifestIO.cs` - manifest I/O
- `src/05_MailScanner.cs` - Outlook mail 収集
- `src/06_FolderScanner.cs` - folder 収集
- `src/07_ProfileRunner.cs`, `src/07a_EventWatcher.cs` - Pull / Watch 実行
- `src/08_MonitorForm.cs`, `src/09_SettingsForm.cs` - UI

## 起動

```bat
launch.bat
```

```powershell
powershell -ExecutionPolicy Bypass -File watchbox.ps1
```

## ガードレール

- コードとコメントは ASCII / English only を維持する
- `config.json` を正本とする
- `manifest.csv` ヘッダを変えたら `pub/casedesk` も同時に確認する
- WPF app と Outlook / folder scanner を疎結合に保つ
