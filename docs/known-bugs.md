# Known Bugs

Verified against commit `ed82906` (2026-04-01).

---

## 1. Manifest CSV fields not fully quoted

**Severity:** Moderate
**Files:** `src/02_ManifestIO.cs` (AppendMailRow, AppendFolderRow, WriteFolderManifest)

`CsvQuote()` is only applied to `subject`, `bodyText`, and `fileName`. Other fields
(`senderName`, `folderPath`, `filePath`, `mailFolder`, `attachmentPaths`, `relativePath`)
are written unquoted.

If `senderName` contains a comma (e.g. "Smith, John"), `CsvSplit` on read will produce
extra columns and the row will be parsed incorrectly.

**Fix:** Apply `CsvQuote()` to all string fields in `AppendMailRow`, `AppendFolderRow`,
and `WriteFolderManifest`.

---

## 2. manifest_hidden toggle causes read/write file mismatch

**Severity:** Moderate
**Files:** `src/02_ManifestIO.cs` (ResolvePath, WritePath)

When `manifest_hidden` is changed (e.g. 1 -> 0), the old file is not deleted or renamed.

- `ResolvePath` (read): prefers `.manifest.csv` over `manifest.csv`
- `WritePath` (write): uses the current `hide` setting

After toggling hidden -> visible, reads hit the old `.manifest.csv` while writes go to
`manifest.csv`. This causes duplicate entries, missed change detection, and data
inconsistency.

**Fix:** When writing, check if the other-named manifest exists and delete or rename it.

---

## 3. Auto-unzip extracted files not tracked in manifest

**Severity:** Low
**Files:** `src/06_FolderScanner.cs` (ScanWithCopy, lines 122-128)

When `auto_unzip=1`, the zip is extracted and deleted, but the extracted files are never
added to scan results. Additionally, `RewriteFolderManifest` scans the source folder
where the original `.zip` still exists, so the manifest references a zip that is not
present in the output directory.

**Fix:** After extraction, enumerate extracted files and add them to results. Or re-scan
the output directory after extraction.

---

## 4. addProfile / resetAll missing auto_unzip key in JS

**Severity:** Low
**Files:** `web/js/settings.js` (addProfile line ~260, resetAll line ~330)

New profile objects created in JS do not include `auto_unzip`. On save, C# extracts an
empty string instead of "0", writing `""` to config.json. Functionally harmless
(`"" != "1"` is falsy) but inconsistent with the expected `"0"` value.

**Fix:** Add `auto_unzip: '0'` to the new profile template in `addProfile` and `resetAll`.

---

## 5. filter_mode case sensitivity inconsistency in FolderScanner

**Severity:** Low
**Files:** `src/06_FolderScanner.cs` (ParseFilters, line 37)

`FolderScanner` compares `config["filter_mode"] == "and"` without lowercasing.
`MailScanner` lowercases before comparison. If "AND" arrives via CSV import, only
`FolderScanner` misinterprets it as OR mode.

**Fix:** Change to `config["filter_mode"].ToLower() == "and"` or equivalent.
