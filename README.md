# Duplicate File Scanner

PowerShell scripts for finding duplicate files -- in OneDrive cloud storage or on a local drive.

## Scripts

### `dupescan.ps1` -- OneDrive Cloud Scanner

Scans your entire OneDrive via the Microsoft Graph API using QuickXorHash fingerprints. No files are downloaded.

1. **Authenticates** with Microsoft Graph using interactive browser login.
2. **Scans** every folder using breadth-first traversal, collecting file metadata and QuickXorHash values.
3. **Recovers missing hashes** via batch API requests for files that didn't return a hash on the first pass.
4. **Exports** all file records to a timestamped CSV in the script directory.
5. **Optionally groups duplicates** with `-GroupDuplicates`.

### `localscan.ps1` -- Local Drive Scanner

Scans a local directory tree, hashing files to find duplicates. Uses a size pre-filter to skip hashing files with unique sizes, dramatically speeding up large scans.

1. **Enumerates** all files recursively in the target directory.
2. **Groups by file size** -- files with a unique size can't be duplicates, so they're skipped.
3. **Hashes** only files that share a size with at least one other file.
4. **Exports** all file records to a timestamped CSV in the script directory.
5. **Optionally groups duplicates** with `-GroupDuplicates`.

## Prerequisites

- PowerShell 5.1 or later

For the OneDrive scanner only:
- The `Microsoft.Graph.Authentication` module:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

- A Microsoft account with OneDrive access

## Usage

### OneDrive scan

```powershell
# Basic scan
.\dupescan.ps1

# With duplicate grouping
.\dupescan.ps1 -GroupDuplicates
```

### Local drive scan

```powershell
# Basic scan
.\localscan.ps1 -Path "D:\Photos"

# With duplicate grouping
.\localscan.ps1 -Path "C:\Users\me\Documents" -GroupDuplicates

# Use a faster hash algorithm
.\localscan.ps1 -Path "D:\" -Algorithm MD5 -GroupDuplicates

# Exclude directories by wildcard
.\localscan.ps1 -Path "C:\" -Exclude "*\node_modules\*", "*\.git\*" -GroupDuplicates
```

#### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-Path` | Yes | Directory to scan |
| `-GroupDuplicates` | No | Produce a duplicate report and CSV |
| `-Algorithm` | No | Hash algorithm: `SHA256` (default), `SHA1`, or `MD5` |
| `-Exclude` | No | Wildcard patterns to exclude (matched against full path) |

## Output

Both scripts produce CSVs in the script directory. With `-GroupDuplicates`, a second CSV is generated containing only duplicate files grouped by number.

### OneDrive hashes CSV

| Column | Description |
|---|---|
| FullPath | Full OneDrive path |
| FileName | File name |
| QuickXorHash | Microsoft's QuickXorHash fingerprint |
| SizeBytes | File size in bytes |
| LastModified | Last modified timestamp |
| Created | Creation timestamp |
| ItemId | Graph API item ID |

### Local hashes CSV

| Column | Description |
|---|---|
| FullPath | Full local file path |
| FileName | File name |
| FileHash | Hash value (or `ERROR` / empty if skipped) |
| SizeBytes | File size in bytes |
| LastModified | Last modified timestamp |
| Created | Creation timestamp |

### Duplicates CSV (both scripts)

| Column | Description |
|---|---|
| DuplicateGroup | Group number (1 = most copies) |
| FullPath | File path |
| FileName | File name |
| FileHash / QuickXorHash | Hash identifying the duplicate set |
| SizeBytes | File size in bytes |
| LastModified | Last modified timestamp |
| Created | Creation timestamp |

## Duplicate report

When `-GroupDuplicates` is used, the console displays:
- Total number of duplicate groups
- Total duplicate files
- Wasted space in MB
- The top 10 duplicate groups with file paths
