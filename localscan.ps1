# ================================================
# Local Drive Duplicate Scanner
# Hashes files on disk and identifies duplicates
# ================================================

#Requires -Version 5.1

param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [switch]$GroupDuplicates,

    [ValidateSet('SHA256', 'SHA1', 'MD5')]
    [string]$Algorithm = 'SHA256',

    [string[]]$Exclude
)

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Host "Error: '$Path' is not a valid directory." -ForegroundColor Red
    exit 1
}

$Path = (Resolve-Path -LiteralPath $Path).Path

Write-Host "Starting local scan of: $Path" -ForegroundColor Green
Write-Host "  Hash algorithm: $Algorithm" -ForegroundColor Gray
if ($Exclude) {
    Write-Host "  Excluding: $($Exclude -join ', ')" -ForegroundColor Gray
}

$scanStart = [System.Diagnostics.Stopwatch]::StartNew()

# --- Enumerate files ---
Write-Host "`n[1/2] Enumerating files..." -ForegroundColor Cyan
$sw = [System.Diagnostics.Stopwatch]::StartNew()

$allFiles = Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue |
    Where-Object {
        $file = $_
        if ($Exclude) {
            $skip = $false
            foreach ($pattern in $Exclude) {
                if ($file.FullName -like $pattern) { $skip = $true; break }
            }
            -not $skip
        }
        else { $true }
    }

$fileCount = @($allFiles).Count
Write-Host "  Found $fileCount files in $([math]::Round($sw.Elapsed.TotalSeconds, 1))s" -ForegroundColor Gray

if ($fileCount -eq 0) {
    Write-Host "No files to scan." -ForegroundColor Yellow
    exit 0
}

# --- Hash files (only hash files that share a size with at least one other) ---
Write-Host "`n[2/2] Hashing files..." -ForegroundColor Cyan
$sw.Restart()

$sizeGroups = $allFiles | Where-Object { $_.Length -gt 0 } | Group-Object -Property Length
$uniqueBySize = @($sizeGroups | Where-Object { $_.Count -eq 1 }).Count
$toHash = $sizeGroups | Where-Object { $_.Count -gt 1 }
$hashCandidateCount = ($toHash | Measure-Object -Property Count -Sum).Sum
if (-not $hashCandidateCount) { $hashCandidateCount = 0 }

Write-Host "  $uniqueBySize files have unique sizes (skipping hash)" -ForegroundColor Gray
Write-Host "  $hashCandidateCount files share a size and will be hashed" -ForegroundColor Gray

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$errors = 0
$hashed = 0

foreach ($file in ($sizeGroups | Where-Object { $_.Count -eq 1 } | ForEach-Object { $_.Group[0] })) {
    $null = $results.Add([PSCustomObject]@{
        FullPath     = $file.FullName
        FileName     = $file.Name
        FileHash     = $null
        SizeBytes    = $file.Length
        LastModified = $file.LastWriteTime
        Created      = $file.CreationTime
    })
}

foreach ($group in $toHash) {
    foreach ($file in $group.Group) {
        $hashed++
        if ($hashed % 100 -eq 0 -or $hashed -eq $hashCandidateCount) {
            $elapsed = $scanStart.Elapsed.ToString("hh\:mm\:ss")
            Write-Progress -Activity "Hashing files [$elapsed elapsed]" `
                -Status "$hashed / $hashCandidateCount" `
                -PercentComplete ([math]::Min(100, [math]::Round($hashed / $hashCandidateCount * 100))) `
                -Id 1
        }

        try {
            $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm $Algorithm -ErrorAction Stop).Hash
            $null = $results.Add([PSCustomObject]@{
                FullPath     = $file.FullName
                FileName     = $file.Name
                FileHash     = $hash
                SizeBytes    = $file.Length
                LastModified = $file.LastWriteTime
                Created      = $file.CreationTime
            })
        }
        catch {
            $errors++
            $null = $results.Add([PSCustomObject]@{
                FullPath     = $file.FullName
                FileName     = $file.Name
                FileHash     = "ERROR"
                SizeBytes    = $file.Length
                LastModified = $file.LastWriteTime
                Created      = $file.CreationTime
            })
        }
    }
}

Write-Progress -Activity "Hashing files" -Completed -Id 1
Write-Host "  Hashed $hashed files in $([math]::Round($sw.Elapsed.TotalSeconds, 1))s ($errors errors)" -ForegroundColor Gray

# --- Export to CSV ---
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$csvPath = Join-Path $scriptDir "Local_Hashes_$timestamp.csv"

$results | Sort-Object FullPath | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# --- Final summary ---
$totalElapsed = $scanStart.Elapsed.ToString("hh\:mm\:ss")
Write-Host "`n=== Scan Complete ===" -ForegroundColor Green
Write-Host "  Total time:          $totalElapsed" -ForegroundColor White
Write-Host "  Total files:         $fileCount" -ForegroundColor White
Write-Host "  Files hashed:        $hashed" -ForegroundColor White
Write-Host "  Skipped (unique size): $uniqueBySize" -ForegroundColor White
if ($errors -gt 0) {
    Write-Host "  Errors:              $errors" -ForegroundColor Red
}
Write-Host "  Results saved to:    $csvPath" -ForegroundColor Cyan

if ($GroupDuplicates) {
    $hashableResults = $results | Where-Object { $_.FileHash -and $_.FileHash -ne "ERROR" }
    $duplicateGroups = $hashableResults |
        Group-Object -Property FileHash |
        Where-Object { $_.Count -gt 1 } |
        Sort-Object { $_.Count } -Descending

    if ($duplicateGroups.Count -eq 0) {
        Write-Host "`n  No duplicates found." -ForegroundColor Green
    }
    else {
        $totalDuplicateFiles = ($duplicateGroups | Measure-Object -Property Count -Sum).Sum
        $wastedBytes = ($duplicateGroups | ForEach-Object {
            ($_.Group[0].SizeBytes) * ($_.Count - 1)
        } | Measure-Object -Sum).Sum
        $wastedMB = [math]::Round($wastedBytes / 1MB, 2)

        Write-Host "`n=== Duplicate Report ===" -ForegroundColor Yellow
        Write-Host "  Duplicate groups:      $($duplicateGroups.Count)" -ForegroundColor White
        Write-Host "  Total duplicate files:  $totalDuplicateFiles" -ForegroundColor White
        Write-Host "  Wasted space:          $wastedMB MB" -ForegroundColor White

        $dupRows = [System.Collections.Generic.List[PSCustomObject]]::new()
        $groupNum = 0
        foreach ($group in $duplicateGroups) {
            $groupNum++
            foreach ($file in ($group.Group | Sort-Object FullPath)) {
                $null = $dupRows.Add([PSCustomObject]@{
                    DuplicateGroup = $groupNum
                    FullPath       = $file.FullPath
                    FileName       = $file.FileName
                    FileHash       = $file.FileHash
                    SizeBytes      = $file.SizeBytes
                    LastModified   = $file.LastModified
                    Created        = $file.Created
                })
            }
        }

        $dupCsvPath = Join-Path $scriptDir "Local_Duplicates_$timestamp.csv"
        $dupRows | Export-Csv -Path $dupCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Duplicates saved to:   $dupCsvPath" -ForegroundColor Cyan

        Write-Host "`n--- Top 10 Duplicate Groups ---" -ForegroundColor Yellow
        $shown = 0
        foreach ($group in $duplicateGroups) {
            if ($shown -ge 10) { break }
            $shown++
            $sizeMB = [math]::Round($group.Group[0].SizeBytes / 1MB, 2)
            Write-Host "  Group $shown ($($group.Count) copies, $sizeMB MB each):" -ForegroundColor White
            foreach ($file in ($group.Group | Sort-Object FullPath)) {
                Write-Host "    $($file.FullPath)" -ForegroundColor Gray
            }
        }
        if ($duplicateGroups.Count -gt 10) {
            Write-Host "  ... and $($duplicateGroups.Count - 10) more groups (see CSV for full list)" -ForegroundColor DarkGray
        }
    }
}
else {
    Write-Host "  Tip: Run with -GroupDuplicates to identify and report duplicate files." -ForegroundColor Cyan
}
