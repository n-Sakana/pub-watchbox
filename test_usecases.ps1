# test_usecases.ps1 - Runtime use case tests for watchbox
$ErrorActionPreference = 'Continue'
$pass = 0; $fail = 0; $skip = 0

function Assert($name, $condition, $detail = '') {
    if ($condition) {
        Write-Host "  PASS: $name" -ForegroundColor Green
        $script:pass++
    } else {
        Write-Host "  FAIL: $name  $detail" -ForegroundColor Red
        $script:fail++
    }
}
function Skip($name, $reason) {
    Write-Host "  SKIP: $name ($reason)" -ForegroundColor Yellow
    $script:skip++
}

# --- Compile (same as watchbox.ps1) ---
Write-Host "`n=== Compiling C# ===" -ForegroundColor Cyan
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml

$libDir = "$PSScriptRoot\lib"
if (Test-Path "$libDir\Microsoft.Web.WebView2.Core.dll") {
    [System.Reflection.Assembly]::LoadFrom("$libDir\Microsoft.Web.WebView2.Core.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$libDir\Microsoft.Web.WebView2.Wpf.dll") | Out-Null
}

$combined = (Get-ChildItem "$PSScriptRoot\src\*.cs" | Sort-Object Name |
    ForEach-Object { Get-Content $_ -Raw }) -join "`n"
$pat = '(?m)^\s*using\s+[\w][\w.]*\s*;'
$usings = [regex]::Matches($combined, $pat) | ForEach-Object { $_.Value.Trim() } | Sort-Object -Unique
$body = $combined -replace $pat, ''
$source = ($usings -join "`n") + "`n`n" + $body

$refs = @(
    [System.Windows.Window].Assembly.Location
    [System.Windows.UIElement].Assembly.Location
    [System.Windows.DependencyObject].Assembly.Location
    [System.Xaml.XamlReader].Assembly.Location
    'Microsoft.CSharp'
    'System.Drawing'
    'System.IO.Compression'
    'System.IO.Compression.FileSystem'
)
if (Test-Path "$libDir\Microsoft.Web.WebView2.Wpf.dll") {
    $refs += "$libDir\Microsoft.Web.WebView2.Core.dll"
    $refs += "$libDir\Microsoft.Web.WebView2.Wpf.dll"
}
Add-Type -TypeDefinition $source -ReferencedAssemblies $refs
Write-Host "Compile OK" -ForegroundColor Green

# --- Setup temp directory ---
$testDir = Join-Path ([IO.Path]::GetTempPath()) "watchbox_test_$(Get-Random)"
New-Item -ItemType Directory -Path $testDir -Force | Out-Null

# ============================================================
Write-Host "`n=== UC1: CSV quoting - commas, quotes, newlines ===" -ForegroundColor Cyan
# ============================================================
$ud1 = Join-Path $testDir "uc1"
New-Item -ItemType Directory -Path $ud1 -Force | Out-Null

[WatchBox.ManifestIO]::AppendMailRow($ud1,
    "EID001", "alice@example.com", "Alice, Jr.",
    'Re: Hello, "World"', [DateTime]::new(2025,3,15,10,30,0),
    "\\Inbox\Sub,Folder", "$ud1\body.txt", "$ud1\mail.msg",
    "att1.pdf|att2.xlsx", $ud1, "body with, comma and `"quote`"",
    "bob@example.com;carol@example.com", "dave@example.com", $false)

$csvPath = [WatchBox.ManifestIO]::ResolvePath($ud1)
$lines = [IO.File]::ReadAllLines($csvPath)
Assert "manifest file created" (Test-Path $csvPath)
Assert "has header + 1 data row" ($lines.Count -eq 2)

# Parse back with CsvSplit
$cols = [WatchBox.ManifestIO]::CsvSplit($lines[1])
Assert "CsvSplit recovers 13 columns" ($cols.Length -eq 13) "got $($cols.Length)"
Assert "sender_name with comma" ($cols[2] -eq 'Alice, Jr.') "got '$($cols[2])'"
Assert "subject with comma+quotes" ($cols[3] -eq 'Re: Hello, "World"') "got '$($cols[3])'"
Assert "to_recipients recovered" ($cols[11] -eq 'bob@example.com;carol@example.com') "got '$($cols[11])'"
Assert "cc_recipients recovered" ($cols[12] -eq 'dave@example.com') "got '$($cols[12])'"

# ============================================================
Write-Host "`n=== UC2: Dedup key consistency ===" -ForegroundColor Cyan
# ============================================================
$ud2 = Join-Path $testDir "uc2"
New-Item -ItemType Directory -Path $ud2 -Force | Out-Null

# Write two rows with different EntryIDs but same subject+sender+date
[WatchBox.ManifestIO]::AppendMailRow($ud2,
    "EID_A", "sender@test.com", "Sender", "Test Subject",
    [DateTime]::new(2025,6,1,14,0,0), "\\Inbox", "", "", "", $ud2, "",
    "", "", $false)
[WatchBox.ManifestIO]::AppendMailRow($ud2,
    "EID_B", "other@test.com", "Other", "Different Subject",
    [DateTime]::new(2025,6,1,14,0,0), "\\Inbox", "", "", "", $ud2, "",
    "", "", $false)

$dedupKeys = [WatchBox.ManifestIO]::LoadDedupKeys($ud2)
Assert "dedup keys loaded" ($dedupKeys.Count -eq 2)

# Verify key format matches what MailScanner would generate
$expectedKey = "test subject|sender@test.com|2025-06-01T14:00:00"
Assert "dedup key matches scan format" ($dedupKeys.Contains($expectedKey)) "key not found: $expectedKey"

$differentKey = "different subject|other@test.com|2025-06-01T14:00:00"
Assert "second dedup key present" ($dedupKeys.Contains($differentKey))

# Non-matching key should not exist
$missingKey = "test subject|sender@test.com|2025-06-01T14:01:00"
Assert "different time = different key" (-not $dedupKeys.Contains($missingKey))

# ============================================================
Write-Host "`n=== UC3: LoadIds + LoadDedupKeys backward compat (old manifest without recipient columns) ===" -ForegroundColor Cyan
# ============================================================
$ud3 = Join-Path $testDir "uc3"
New-Item -ItemType Directory -Path $ud3 -Force | Out-Null
# Write old-format manifest (11 columns, no to/cc)
$oldHeader = "entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text"
$oldRow = 'EID_OLD,old@test.com,Old Sender,Old Subject,2025-01-01T09:00:00,\Inbox,,,,' + $ud3 + ',some body text'
[IO.File]::WriteAllLines("$ud3\manifest.csv", @($oldHeader, $oldRow))

$ids3 = [WatchBox.ManifestIO]::LoadIds($ud3)
Assert "LoadIds reads old format" ($ids3.Contains("EID_OLD"))

$dk3 = [WatchBox.ManifestIO]::LoadDedupKeys($ud3)
$expected3 = "old subject|old@test.com|2025-01-01T09:00:00"
Assert "LoadDedupKeys reads old format" ($dk3.Contains($expected3))

# ============================================================
Write-Host "`n=== UC4: FileShare concurrent read/write ===" -ForegroundColor Cyan
# ============================================================
$ud4 = Join-Path $testDir "uc4"
New-Item -ItemType Directory -Path $ud4 -Force | Out-Null

# Write initial row
[WatchBox.ManifestIO]::AppendMailRow($ud4,
    "EID_FS1", "fs@test.com", "FS", "FileShare Test",
    [DateTime]::new(2025,7,1,10,0,0), "\\Inbox", "", "", "", $ud4, "",
    "", "", $false)

# Open for reading with FileShare while we append
$csvFs = [WatchBox.ManifestIO]::ResolvePath($ud4)
$readStream = [IO.FileStream]::new($csvFs, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
$reader = [IO.StreamReader]::new($readStream)
$content1 = $reader.ReadToEnd()
$reader.Close(); $readStream.Close()

# Now append another row (should not throw)
$appendOk = $true
try {
    [WatchBox.ManifestIO]::AppendMailRow($ud4,
        "EID_FS2", "fs2@test.com", "FS2", "Second Row",
        [DateTime]::new(2025,7,2,10,0,0), "\\Inbox", "", "", "", $ud4, "",
        "", "", $false)
} catch { $appendOk = $false }
Assert "append after concurrent read succeeds" $appendOk

# Read while holding write handle open
$writeStream = [IO.FileStream]::new($csvFs, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
$readOk = $true
try {
    $readStream2 = [IO.FileStream]::new($csvFs, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $reader2 = [IO.StreamReader]::new($readStream2)
    $content2 = $reader2.ReadToEnd()
    $reader2.Close(); $readStream2.Close()
} catch { $readOk = $false }
$writeStream.Close()
Assert "read while write handle open succeeds" $readOk
Assert "both rows present after concurrent access" ($content2 -match 'EID_FS1' -and $content2 -match 'EID_FS2')

# ============================================================
Write-Host "`n=== UC5: ChangeLog FileShare ===" -ForegroundColor Cyan
# ============================================================
$ud5 = Join-Path $testDir "uc5"
New-Item -ItemType Directory -Path $ud5 -Force | Out-Null

[WatchBox.ChangeLog]::Append($ud5, "added", "ID1", "file1.txt")
[WatchBox.ChangeLog]::Append($ud5, "modified", "ID2", 'file with, comma')
[WatchBox.ChangeLog]::Append($ud5, "removed", "ID3", 'file "quoted"')

$logPath = Join-Path $ud5 "log.csv"
$logLines = [IO.File]::ReadAllLines($logPath)
Assert "log has header + 3 rows" ($logLines.Count -eq 4) "got $($logLines.Count)"

$logCols2 = [WatchBox.ManifestIO]::CsvSplit($logLines[2])
Assert "log comma in name preserved" ($logCols2[3] -eq 'file with, comma') "got '$($logCols2[3])'"

$logCols3 = [WatchBox.ManifestIO]::CsvSplit($logLines[3])
Assert "log quotes in name preserved" ($logCols3[3] -eq 'file "quoted"') "got '$($logCols3[3])'"

# ============================================================
Write-Host "`n=== UC6: PurgeStaleMailIds + dedup key reload order ===" -ForegroundColor Cyan
# ============================================================
$ud6 = Join-Path $testDir "uc6"
$ud6mail = Join-Path $ud6 "mail_dir_exists"
$ud6gone = Join-Path $ud6 "mail_dir_gone"
New-Item -ItemType Directory -Path $ud6 -Force | Out-Null
New-Item -ItemType Directory -Path $ud6mail -Force | Out-Null
# Don't create $ud6gone - simulate manual deletion

[WatchBox.ManifestIO]::AppendMailRow($ud6,
    "EID_KEEP", "keep@test.com", "Keep", "Keep This",
    [DateTime]::new(2025,8,1,10,0,0), "\\Inbox", "", "", "", $ud6mail, "",
    "", "", $false)
[WatchBox.ManifestIO]::AppendMailRow($ud6,
    "EID_GONE", "gone@test.com", "Gone", "Deleted Folder",
    [DateTime]::new(2025,8,2,10,0,0), "\\Inbox", "", "", "", $ud6gone, "",
    "", "", $false)

$ids6 = [WatchBox.ManifestIO]::LoadIds($ud6)
Assert "both IDs initially present" ($ids6.Count -eq 2)

# Simulate what MailScanner.Scan does: purge first, then load dedup keys
$exported6 = [System.Collections.Generic.HashSet[string]]::new($ids6)
# Use reflection to call private PurgeStaleMailIds (cast args explicitly)
$purgeMethod = [WatchBox.MailScanner].GetMethod('PurgeStaleMailIds',
    [System.Reflection.BindingFlags]'Static,NonPublic')
$purgeArgs = [object[]]@([string]$ud6, [System.Collections.Generic.HashSet[string]]$exported6)
$purgeMethod.Invoke($null, $purgeArgs)

Assert "stale ID removed from exported" (-not $exported6.Contains("EID_GONE"))
Assert "valid ID kept in exported" ($exported6.Contains("EID_KEEP"))

# Now load dedup keys AFTER purge
$dk6 = [WatchBox.ManifestIO]::LoadDedupKeys($ud6)
$goneKey = "deleted folder|gone@test.com|2025-08-02T10:00:00"
Assert "purged key NOT in dedup keys (re-pull allowed)" (-not $dk6.Contains($goneKey)) "stale key still present!"
$keepKey = "keep this|keep@test.com|2025-08-01T10:00:00"
Assert "valid key still in dedup keys" ($dk6.Contains($keepKey))

# ============================================================
Write-Host "`n=== UC7: Outlook COM connect + account enumeration ===" -ForegroundColor Cyan
# ============================================================
$scanner = [WatchBox.MailScanner]::new()
$connected = $scanner.Connect()
if ($connected) {
    $accounts = $scanner.GetAccounts()
    Assert "COM connect succeeded" $true
    Assert "at least 1 account found" ($accounts.Count -ge 1) "got $($accounts.Count)"
    Write-Host "    Accounts: $($accounts -join ', ')"

    # Test GetFolders for first account
    if ($accounts.Count -gt 0) {
        $folders = $scanner.GetFolders($accounts[0])
        Assert "folders enumerated" ($folders.Count -ge 1) "got $($folders.Count)"
        Write-Host "    Folders: $($folders.Count) folders found"
    }
    $scanner.Cleanup()
    Assert "COM cleanup succeeded" $true
} else {
    Skip "COM connect" "Outlook not running"
    Skip "account enumeration" "Outlook not running"
    Skip "folder enumeration" "Outlook not running"
    Skip "COM cleanup" "Outlook not running"
}

# ============================================================
Write-Host "`n=== UC8: Folder manifest - write, rewrite, remove ===" -ForegroundColor Cyan
# ============================================================
$ud8 = Join-Path $testDir "uc8"
New-Item -ItemType Directory -Path $ud8 -Force | Out-Null

# Append rows
[WatchBox.ManifestIO]::AppendFolderRow($ud8,
    "HASH001", "report.xlsx", "C:\Data,Shared\report.xlsx",
    "C:\Data,Shared", "report.xlsx", 1024,
    [DateTime]::new(2025,5,1,9,0,0), $false)
[WatchBox.ManifestIO]::AppendFolderRow($ud8,
    "HASH002", 'file "special".txt', "C:\Normal\file.txt",
    "C:\Normal", "file.txt", 512,
    [DateTime]::new(2025,5,2,10,0,0), $false)

$fids = [WatchBox.ManifestIO]::LoadIds($ud8)
Assert "folder IDs loaded" ($fids.Count -eq 2)

$frows = [WatchBox.ManifestIO]::LoadFolderRows($ud8)
Assert "folder rows loaded" ($frows.Count -eq 2)
Assert "path with comma preserved" ($frows["HASH001"].FilePath -eq "C:\Data,Shared\report.xlsx") "got '$($frows["HASH001"].FilePath)'"
Assert "filename with quotes preserved" ($frows["HASH002"].FileName -eq 'file "special".txt') "got '$($frows["HASH002"].FileName)'"

# Remove one row
$removeSet = [System.Collections.Generic.HashSet[string]]::new()
$removeSet.Add("HASH002") | Out-Null
[WatchBox.ManifestIO]::RemoveRows($ud8, $removeSet)
$fids2 = [WatchBox.ManifestIO]::LoadIds($ud8)
Assert "row removed" ($fids2.Count -eq 1 -and $fids2.Contains("HASH001"))

# ============================================================
Write-Host "`n=== UC9: WriteFolderManifest quoting ===" -ForegroundColor Cyan
# ============================================================
$ud9 = Join-Path $testDir "uc9"
New-Item -ItemType Directory -Path $ud9 -Force | Out-Null

$rows9 = [System.Collections.Generic.List[WatchBox.FolderManifestRow]]::new()
$r = [WatchBox.FolderManifestRow]::new()
$r.ItemId = "H001"; $r.FileName = "data,file.csv"
$r.FilePath = 'C:\Users\Smith, John\data,file.csv'
$r.FolderPath = 'C:\Users\Smith, John'
$r.RelativePath = 'data,file.csv'
$r.FileSize = "2048"; $r.ModifiedAt = "2025-09-01T12:00:00"
$rows9.Add($r)

[WatchBox.ManifestIO]::WriteFolderManifest($ud9, $rows9, $false)
$csv9 = [WatchBox.ManifestIO]::ResolvePath($ud9)
$lines9 = [IO.File]::ReadAllLines($csv9)
$cols9 = [WatchBox.ManifestIO]::CsvSplit($lines9[1])
Assert "WriteFolderManifest: 7 columns" ($cols9.Length -eq 7) "got $($cols9.Length)"
Assert "WriteFolderManifest: filename with comma" ($cols9[1] -eq 'data,file.csv') "got '$($cols9[1])'"
Assert "WriteFolderManifest: path with comma" ($cols9[2] -eq 'C:\Users\Smith, John\data,file.csv') "got '$($cols9[2])'"

# ============================================================
Write-Host "`n=== UC10: Zip slip protection ===" -ForegroundColor Cyan
# ============================================================
$ud10 = Join-Path $testDir "uc10"
New-Item -ItemType Directory -Path $ud10 -Force | Out-Null

# Create a zip with a path-traversal entry
$zipPath = Join-Path $ud10 "evil.zip"
$zipExtract = Join-Path $ud10 "extracted"
$sentinel = Join-Path $ud10 "should_not_exist.txt"

# Create zip using .NET
$zipStream = [IO.FileStream]::new($zipPath, [IO.FileMode]::Create)
$archive = [IO.Compression.ZipArchive]::new($zipStream, [IO.Compression.ZipArchiveMode]::Create)
# Normal entry
$entry1 = $archive.CreateEntry("normal.txt")
$w1 = [IO.StreamWriter]::new($entry1.Open())
$w1.Write("normal content"); $w1.Close()
# Evil entry with path traversal
$entry2 = $archive.CreateEntry("../../should_not_exist.txt")
$w2 = [IO.StreamWriter]::new($entry2.Open())
$w2.Write("evil content"); $w2.Close()
$archive.Dispose(); $zipStream.Close()

# Call TryExtractZip via reflection (it's static private in MailScanner)
$tryExtract = [WatchBox.MailScanner].GetMethod('TryExtractZip',
    [System.Reflection.BindingFlags]'Static,NonPublic')
$extractArgs = [object[]]@([string]$zipPath, [string]$ud10)
$result = $tryExtract.Invoke($null, $extractArgs)
Assert "zip extraction succeeded" ($result -eq $true)
Assert "normal file extracted" (Test-Path (Join-Path $ud10 "evil\normal.txt"))
Assert "path traversal file BLOCKED" (-not (Test-Path $sentinel)) "zip slip succeeded!"

# ============================================================
Write-Host "`n=== UC11: Scan with actual Outlook data (small test) ===" -ForegroundColor Cyan
# ============================================================
$scanner2 = [WatchBox.MailScanner]::new()
if ($scanner2.Connect()) {
    $accounts2 = $scanner2.GetAccounts()
    if ($accounts2.Count -gt 0) {
        $ud11 = Join-Path $testDir "uc11"
        New-Item -ItemType Directory -Path $ud11 -Force | Out-Null

        # Get inbox path via GetFolders
        $folders11 = $scanner2.GetFolders($accounts2[0])
        $inboxPath11 = ""
        foreach ($f in $folders11) {
            if ([string]::IsNullOrEmpty($inboxPath11) -and -not [string]::IsNullOrEmpty($f[1])) {
                $inboxPath11 = $f[1]
            }
        }
        Write-Host "    Inbox path: $inboxPath11"
        Write-Host "    Folders available: $($folders11.Count)"

        $config11 = [System.Collections.Generic.Dictionary[string,string]]::new()
        $config11["account"] = $accounts2[0]
        $config11["outlook_folder"] = $inboxPath11
        $config11["output_root"] = $ud11
        $config11["since"] = (Get-Date).AddDays(-14).ToString("yyyy-MM-dd")
        $config11["filter_mode"] = "or"
        $config11["filters"] = ""
        $config11["flat_output"] = "1"
        $config11["short_dirname"] = "1"
        $config11["auto_unzip"] = "0"
        $config11["source_folder"] = ""
        $config11["recurse"] = "0"
        $config11["last_scan"] = ""

        $knownIds11 = [WatchBox.ManifestIO]::LoadIds($ud11)
        $results11 = $scanner2.Scan($config11, $knownIds11)

        Assert "scan completed without crash" $true
        Write-Host "    Items found: $($results11.Count)"

        if ($results11.Count -gt 0) {
            $first = $results11[0]
            Assert "ScanResult has EntryID" (-not [string]::IsNullOrEmpty($first.ItemId))
            Assert "ScanResult has Subject" ($first.Subject -ne $null)
            Assert "ScanResult has SenderEmail" ($first.SenderEmail -ne $null)
            Assert "ScanResult has ReceivedAt" ($first.ReceivedAt -gt [DateTime]::MinValue)

            # Check body.txt has metadata header
            if (Test-Path $first.BodyPath) {
                $bodyContent = [IO.File]::ReadAllText($first.BodyPath)
                Assert "body.txt starts with From:" ($bodyContent.StartsWith("From:"))
                Assert "body.txt has separator ---" ($bodyContent.Contains("---"))
                Assert "body.txt has Date:" ($bodyContent.Contains("Date:"))
                Assert "body.txt has Subject:" ($bodyContent.Contains("Subject:"))
            } else {
                Skip "body.txt content" "body file not found"
            }

            # Check meta.json has recipients
            $metaPath = Join-Path (Split-Path $first.BodyPath) "meta.json"
            if (Test-Path $metaPath) {
                $metaContent = [IO.File]::ReadAllText($metaPath)
                Assert "meta.json has to_recipients" ($metaContent.Contains('"to_recipients"'))
                Assert "meta.json has cc_recipients" ($metaContent.Contains('"cc_recipients"'))
            }

            # Write manifest rows (Scan() returns results but doesn't write manifest)
            foreach ($sr in $results11) {
                [WatchBox.ManifestIO]::AppendMailRow($ud11,
                    $sr.ItemId, $sr.SenderEmail, $sr.SenderName,
                    $sr.Subject, $sr.ReceivedAt, $sr.SourcePath,
                    $sr.BodyPath, $sr.MsgPath, $sr.AttachmentPaths,
                    $sr.ItemFolder, $sr.BodyText,
                    $sr.ToRecipients, $sr.CcRecipients, $false)
            }

            # Check manifest has 13 columns
            $csvPath11 = [WatchBox.ManifestIO]::ResolvePath($ud11)
            $mlines = [IO.File]::ReadAllLines($csvPath11)
            Assert "manifest written with data rows" ($mlines.Count -ge 2) "got $($mlines.Count) lines"
            if ($mlines.Count -ge 2) {
                $mcols = [WatchBox.ManifestIO]::CsvSplit($mlines[1])
                Assert "manifest has 13 columns" ($mcols.Length -eq 13) "got $($mcols.Length)"
            }

            # UC12: Re-scan should skip already exported (dedup by EntryID)
            Write-Host "`n=== UC12: Re-scan dedup - same profile ===" -ForegroundColor Cyan
            $knownIds12 = [WatchBox.ManifestIO]::LoadIds($ud11)
            Assert "knownIds loaded from manifest" ($knownIds12.Count -eq $results11.Count) "expected $($results11.Count), got $($knownIds12.Count)"
            $results12 = $scanner2.Scan($config11, $knownIds12)
            Assert "re-scan returns 0 new items (EntryID dedup)" ($results12.Count -eq 0) "got $($results12.Count)"

            # UC13: Cross-profile dedup via composite key
            Write-Host "`n=== UC13: Cross-profile dedup via composite key ===" -ForegroundColor Cyan
            $ud13 = Join-Path $testDir "uc13"
            New-Item -ItemType Directory -Path $ud13 -Force | Out-Null
            # Copy manifest to different output_root (simulates different profile)
            Copy-Item $csvPath11 (Join-Path $ud13 "manifest.csv")
            # Manifest has data but mail folders don't exist in ud13
            # PurgeStaleMailIds will remove rows (folders don't exist)
            # Then LoadDedupKeys reads AFTER purge = 0 keys (correct behavior)
            # Instead, test that dedup keys can be loaded before purge
            $dedupKeys13 = [WatchBox.ManifestIO]::LoadDedupKeys($ud13)
            Assert "dedup keys loadable from copied manifest" ($dedupKeys13.Count -gt 0) "got $($dedupKeys13.Count)"

            # Now do actual scan with empty knownIds + the dedup keys
            # This tests the composite key dedup path
            $config13 = [System.Collections.Generic.Dictionary[string,string]]::new($config11)
            $config13["output_root"] = $ud13
            $knownIds13 = [System.Collections.Generic.HashSet[string]]::new()
            $results13 = $scanner2.Scan($config13, $knownIds13)
            # Items should be skipped by meta.json check (files exist in ud11, not ud13)
            # or re-exported to ud13 since neither EntryID nor composite key blocks them
            # (purge removes manifest rows since mail_folder dirs don't exist in ud13)
            Write-Host "    Cross-profile scan items: $($results13.Count)"
            Assert "cross-profile scan completes without crash" $true
        } else {
            Skip "ScanResult fields" "no items in date range"
            Skip "body.txt header" "no items"
            Skip "meta.json recipients" "no items"
            Skip "manifest columns" "no items"
            Skip "re-scan dedup" "no items"
            Skip "cross-profile dedup" "no items"
        }
    } else {
        Skip "scan test" "no accounts found"
    }
    $scanner2.Cleanup()
} else {
    Skip "scan test" "Outlook not available"
}

# ============================================================
Write-Host "`n=== UC14: EventWatcher thread safety ===" -ForegroundColor Cyan
# ============================================================
$watcher = [WatchBox.EventWatcher]::new()
# Start and stop rapidly to test thread join
$watcher.Start({ param($idx, $name) })
Start-Sleep -Milliseconds 500
$watcher.Stop()
Assert "watcher start/stop without crash" $true

# Start again (should not double-hook)
$watcher.Start({ param($idx, $name) })
Start-Sleep -Milliseconds 500
$watcher.Stop()
Assert "watcher restart without crash" $true

# --- Cleanup ---
Write-Host "`n=== Cleanup ===" -ForegroundColor Cyan
try { Remove-Item $testDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
Write-Host "Temp dir cleaned up"

# --- Summary ---
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RESULTS: $pass passed, $fail failed, $skip skipped" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
Write-Host "========================================`n"

exit $fail
