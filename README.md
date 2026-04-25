# watchbox

Windows 上で mail / folder を監視・収集し、`manifest.csv` と `log.csv` を生成するツール。PowerShell ローダーが C# ソースをその場でコンパイルし、WPF UI を起動する。

## 役割

- Outlook からメールを収集する
- フォルダを監視し、必要ならコピー同期する
- output ごとに `manifest.csv` を維持する
- 変更があれば `log.csv` を追記する
- Pull と Watch の 2 モードで動く

CaseDesk から見ると、watchbox は外部収集と manifest 生成の中核。

## 要件

- Windows 10 / 11
- PowerShell 5.1
- Outlook データを扱う場合は Outlook desktop app
- WPF が動く .NET / Windows 環境

## 起動

1. `launch.vbs` をダブルクリックするか、開発中は `launch.bat` を使う
2. Settings で profile を設定する
3. Pull で全 profile を一括実行する
4. Watch で event-driven 監視を開始する

## 動作構成

```text
watchbox.ps1
  ├── src/*.cs を連結
  ├── Add-Type でその場コンパイル
  └── [WatchBox.App]::Run()

WPF UI
  ├── MonitorWindow   メイン画面
  ├── SettingsWindow  profile 設定
  ├── SearchWindow    manifest 検索
  └── ToastPopup      通知
```

## profile

profile は `config.json` に保存する。`type` は `mail` か `folder`。

共通キー:

- `name`
- `type`
- `output_root`
- `notify`
- `log_enabled`
- `since`
- `filter_mode`
- `filters`
- `flat_output`

mail profile 追加キー:

- `account`
- `outlook_folder`

folder profile 追加キー:

- `source_folder`
- `recurse`

### mail profile

- Outlook COM に接続してメールを走査する
- `output_root/manifest.csv` を mail 形式で更新する
- 本文、msg、添付ファイルを出力する
- `knownIds` と `latest received_at` を使って増分走査する

### folder profile

2 つの動作がある。

- `source_folder` あり: source から `output_root` へコピー同期しつつ manifest 更新
- `source_folder` なし: `output_root` を直接走査する manifest-only モード

どちらも削除検出を行い、必要なら manifest から該当行を落とす。

## Pull / Watch

### Pull

`MonitorWindow` の Pull は全 profile を 1 回ずつ実行する。STA thread 上で `ProfileRunner.Run` を回し、mail / folder ごとの結果を集計する。

### Watch

`EventWatcher` が profile に応じて監視を張る。

- folder: `FileSystemWatcher`
- mail: Outlook `ItemAdd` イベント

イベント発火後は該当 profile だけ `ProfileRunner.Run` を呼び直す。2 秒デバウンスあり。

## manifest.csv

### mail

```text
entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text
```

### folder

```text
item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at
```

manifest 型はヘッダで判定する。

## log.csv

`log_enabled=1` のとき、追加・変更・削除を `log.csv` に追記する。CaseDesk 側の Change Log とは別物で、watchbox 側の収集ログ。

## ファイル構成

```text
watchbox/
├── watchbox.ps1
├── launch.bat
├── launch.vbs
├── README.md
└── src/
    ├── 00_FolderPicker.cs
    ├── 01_App.cs
    ├── 02_ManifestIO.cs
    ├── 03_ChangeLog.cs
    ├── 04_SourceScanner.cs
    ├── 05_MailScanner.cs
    ├── 06_FolderScanner.cs
    ├── 07_ProfileRunner.cs
    ├── 07a_EventWatcher.cs
    ├── 08_MonitorForm.cs
    ├── 08a_ToastPopup.cs
    ├── 09_SettingsForm.cs
    └── 10_SearchForm.cs
```

旧 `mailpull` README の説明は現状より狭い。watchbox は mail 専用ではなく、folder profile と event watch まで含む。
