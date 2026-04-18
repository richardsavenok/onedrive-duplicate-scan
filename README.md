# Duplicate File Scanner

PowerShell scripts for finding duplicate files -- in OneDrive cloud storage or on a local drive.

## Scripts

### `dupescan.ps1` -- OneDrive Cloud Scanner

Scans your entire OneDrive via the Microsoft Graph API using QuickXorHash fingerprints. No files are downloaded.

1. **Checks for a checkpoint** -- if a previous scan was interrupted, prompts to resume or start fresh.
2. **Authenticates** with Microsoft Graph using interactive browser login.
3. **Scans** every folder using breadth-first traversal, collecting file metadata and QuickXorHash values. **Saves progress** every 50 folders.
4. **Recovers missing hashes** via batch API requests for files that didn't return a hash on the first pass. Skips items already recovered in a previous partial run.
5. **Exports** all file records to a timestamped CSV in the script directory and clears the checkpoint.
6. **Optionally groups duplicates** with `-GroupDuplicates`.

### `localscan.ps1` -- Local Drive Scanner

Scans a local directory tree, hashing files to find duplicates. Uses a size pre-filter to skip hashing files with unique sizes, dramatically speeding up large scans.

1. **Checks for a checkpoint** -- if a previous scan was interrupted, prompts to resume or start fresh.
2. **Enumerates** all files recursively in the target directory.
3. **Groups by file size** -- files with a unique size can't be duplicates, so they're skipped.
4. **Hashes** only files that share a size with at least one other file, using cached hashes from the checkpoint where available.
5. **Saves progress** periodically so interrupted scans can be resumed without starting over.
6. **Exports** all file records to a timestamped CSV in the script directory and clears the checkpoint.
7. **Optionally groups duplicates** with `-GroupDuplicates`.

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

## Checkpoint / Resume

Both scripts automatically save progress to checkpoint files in the script directory. If a scan is interrupted (Ctrl+C, crash, system restart), the next run detects the checkpoint and prompts to continue or start fresh.

### `localscan.ps1`

Checkpoint files: `localscan_checkpoint.json`, `localscan_checkpoint.csv`

```
Found an existing scan checkpoint:
  Scan path:    D:\Photos
  Algorithm:    SHA256
  Started at:   2026-04-17T10:30:00
  Files cached: 12345

Continue from checkpoint? (Y = continue, N = start fresh)
```

- **Y** -- Loads cached hashes and skips re-hashing files that haven't changed (validated by path, size, and last modified timestamp).
- **N** -- Deletes the checkpoint and starts a fresh scan.

If the scan path or algorithm differs from the checkpoint, the script automatically starts fresh.

### `dupescan.ps1`

Checkpoint files: `dupescan_checkpoint.json`, `dupescan_checkpoint_results.csv`, `dupescan_checkpoint_missing.csv`

```
Found an existing scan checkpoint:
  Drive ID:         b!abc123...
  Started at:       2026-04-17T10:30:00
  Folders scanned:  250
  Files found:      8400
  Queue remaining:  73
  Pass 1 complete:  False

Continue from checkpoint? (Y = continue, N = start fresh)
```

- **Y** -- Restores the BFS queue, collected results, and missing-hash items. Resumes the folder traversal from where it stopped. If pass 1 (BFS) was already complete, skips straight to pass 2 (batch hash recovery), filtering out items already recovered.
- **N** -- Deletes the checkpoint and starts a fresh scan.

If the drive ID differs from the checkpoint, the script automatically starts fresh.

Both scripts delete their checkpoint files on successful completion.

## Duplicate report

When `-GroupDuplicates` is used, the console displays:
- Total number of duplicate groups
- Total duplicate files
- Wasted space in MB
- The top 10 duplicate groups with file paths
