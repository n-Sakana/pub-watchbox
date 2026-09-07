# AGENTS.md — pub/watchbox

## 役割

watchbox は、Windows上でメールとフォルダを収集・監視し、`manifest.csv` と `log.csv` を生成するローカルアプリです。CaseDeskにとっては外部データ収集とmanifest生成の正本です。

## 現行構成

- loader: PowerShell 5.1
- host: runtime compileしたC#
- main monitor: WPF
- settings / viewer: WPF上のWebView2 + local HTML/CSS/JavaScript
- Outlook: late-bound COM
- 通常起動: `launch.vbs`
- console付き起動: `launch.bat`
- 開発起動: `watchbox.ps1`

「UIはすべてWPF」または「build不要なので外部DLL不要」と説明しません。WebView2のmanaged/native DLLとruntimeが必要です。

## 読む順番

1. [README.md](README.md)
2. `watchbox.ps1`
3. `src/01_App.cs`
4. `src/02_ManifestIO.cs`
5. `src/05_MailScanner.cs`、`src/06_FolderScanner.cs`
6. `src/07_ProfileRunner.cs`、`src/07a_EventWatcher.cs`
7. `src/08_MonitorForm.cs`、`src/08b_WebViewHost.cs`
8. `src/09_SettingsForm.cs`、`src/10_SearchForm.cs`
9. `web/`

## 変更時の原則

- source codeとcommentはASCII / Englishに保つ。
- PowerShell 5.1とruntime C# compilerの制約を守る。
- `config.json` をprofile設定の正本とする。
- mail / folder scannerとUIを密結合させない。
- manifest headerを変えるときはCaseDesk側を同時に確認する。
- WebView2のvirtual host mappingとpostMessage境界で、任意file accessや任意navigationを許さない。
- 通常起動は `launch.vbs`。consoleの表示有無を混同しない。
- Pull、Watch、削除検出、manifest更新を実データで確認する。
