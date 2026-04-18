# OneDrive Duplicate Scanner

A PowerShell script that scans your entire OneDrive cloud storage via the Microsoft Graph API to identify duplicate files using QuickXorHash fingerprints. No files are downloaded -- everything runs against the cloud API.

## What It Does

1. **Authenticates** with Microsoft Graph using interactive browser login.
2. **Scans** every folder in your OneDrive using breadth-first traversal, collecting file metadata and QuickXorHash values.
3. **Recovers missing hashes** via batch API requests for files that didn't return a hash on the first pass.
4. **Exports** all file records (path, name, hash, size, dates) to a timestamped CSV in the script directory.
5. **Optionally groups duplicates** -- when run with `-GroupDuplicates`, produces a second CSV listing duplicate groups and prints a summary with wasted space calculations.

## Prerequisites

- PowerShell 5.1 or later
- The `Microsoft.Graph.Authentication` module:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

- A Microsoft account with OneDrive access

## Usage

### Basic scan (export all file hashes)

```powershell
.\dupescan.ps1
```

Produces `OneDrive_Hashes_<timestamp>.csv` in the script directory.

### Scan with duplicate detection

```powershell
.\dupescan.ps1 -GroupDuplicates
```

Produces both:
- `OneDrive_Hashes_<timestamp>.csv` -- all files with their hashes
- `OneDrive_Duplicates_<timestamp>.csv` -- only duplicate files, grouped and numbered

The console will also display:
- Total number of duplicate groups
- Total wasted space in MB
- The top 10 largest duplicate groups with file paths

## Output

### Hashes CSV columns

| Column | Description |
|---|---|
| FullPath | Full OneDrive path |
| FileName | File name |
| QuickXorHash | Microsoft's QuickXorHash fingerprint |
| SizeBytes | File size in bytes |
| LastModified | Last modified timestamp |
| Created | Creation timestamp |
| ItemId | Graph API item ID |

### Duplicates CSV columns (with `-GroupDuplicates`)

| Column | Description |
|---|---|
| DuplicateGroup | Group number (1 = most copies) |
| FullPath | Full OneDrive path |
| FileName | File name |
| QuickXorHash | Shared hash identifying the duplicate set |
| SizeBytes | File size in bytes |
| LastModified | Last modified timestamp |
| Created | Creation timestamp |
