# ================================================
# ROBUST OneDrive Full Scan - QuickXorHash + Path
# Uses raw Graph API with batch retry for missing hashes
# Works 100% in the cloud - no local files downloaded
# Supports checkpoint/resume for interrupted scans
# ================================================

#Requires -Version 5.1

param(
    [switch]$GroupDuplicates
)

# --- Helper for safe nested property access (handles mixed PSObject/Hashtable) ---
function Get-NestedValue {
    param($obj, [string[]]$path)
    $current = $obj
    foreach ($key in $path) {
        if ($null -eq $current) { return $null }
        if ($current -is [hashtable] -or $current -is [System.Collections.IDictionary]) {
            $current = $current[$key]
        } else {
            $current = $current.$key
        }
    }
    return $current
}

# --- Helper: append rows to checkpoint CSV ---
function Save-CheckpointBatch {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$Items,
        [string]$FilePath
    )
    if ($Items.Count -eq 0) { return }
    if (-not (Test-Path $FilePath)) {
        $Items | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
    }
    else {
        $Items | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Add-Content -Path $FilePath -Encoding UTF8
    }
}

# --- Helper: save checkpoint metadata + queue ---
function Save-CheckpointMeta {
    param(
        [string]$MetaPath,
        [string]$DriveId,
        [string]$StartedAt,
        [int]$FoldersScanned,
        [int]$FilesFound,
        [int]$TotalItemsSeen,
        [bool]$Pass1Complete,
        [System.Collections.Generic.Queue[hashtable]]$Queue
    )
    $queueArray = @()
    if ($Queue -and $Queue.Count -gt 0) {
        foreach ($qItem in $Queue.ToArray()) {
            $queueArray += @{
                DriveId     = $qItem.DriveId
                ItemId      = $qItem.ItemId
                CurrentPath = $qItem.CurrentPath
            }
        }
    }
    @{
        DriveId          = $DriveId
        StartedAt        = $StartedAt
        FoldersScanned   = $FoldersScanned
        FilesFound       = $FilesFound
        TotalItemsSeen   = $TotalItemsSeen
        Pass1Complete    = $Pass1Complete
        QueueCount       = $queueArray.Count
        Queue            = $queueArray
    } | ConvertTo-Json -Depth 5 | Set-Content $MetaPath -Encoding UTF8
}

# --- Prerequisite check ---
if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
    Write-Host "Missing required module: Microsoft.Graph.Authentication" -ForegroundColor Red
    Write-Host "Install it with:" -ForegroundColor Yellow
    Write-Host "  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -ForegroundColor White
    exit 1
}

# --- Module import (selective - only what we need) ---
Write-Host "[1/3] Loading Microsoft Graph modules..." -ForegroundColor Cyan
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Write-Host "      Modules loaded in $([math]::Round($sw.Elapsed.TotalSeconds, 1))s" -ForegroundColor Gray

# --- Connect ---
Write-Host "[2/3] Connecting to Microsoft Graph (browser auth may open)..." -ForegroundColor Cyan
$sw.Restart()

Connect-MgGraph -Scopes "Files.Read.All", "User.Read" -NoWelcome

Write-Host "      Connected in $([math]::Round($sw.Elapsed.TotalSeconds, 1))s" -ForegroundColor Gray

# --- Get drive info ---
Write-Host "[3/3] Retrieving OneDrive information..." -ForegroundColor Cyan
$sw.Restart()

$drive = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/drive" -OutputType PSObject
$rootDriveId = $drive.id

Write-Host "      Drive ID: $rootDriveId (retrieved in $([math]::Round($sw.Elapsed.TotalSeconds, 1))s)" -ForegroundColor Gray

# --- Checkpoint paths ---
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$checkpointMeta = Join-Path $scriptDir "dupescan_checkpoint.json"
$checkpointResults = Join-Path $scriptDir "dupescan_checkpoint_results.csv"
$checkpointMissing = Join-Path $scriptDir "dupescan_checkpoint_missing.csv"

# --- Check for existing checkpoint ---
$resuming = $false
$pass1Complete = $false
$startedAt = (Get-Date -Format "o")

$global:Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$missingHashItems = [System.Collections.Generic.List[PSCustomObject]]::new()

$queue = [System.Collections.Generic.Queue[hashtable]]::new()
$foldersScanned = 0
$filesFound = 0
$totalItemsSeen = 0

if (Test-Path $checkpointMeta) {
    $meta = Get-Content $checkpointMeta -Raw | ConvertFrom-Json

    Write-Host "`nFound an existing scan checkpoint:" -ForegroundColor Yellow
    Write-Host "  Drive ID:         $($meta.DriveId)" -ForegroundColor White
    Write-Host "  Started at:       $($meta.StartedAt)" -ForegroundColor White
    Write-Host "  Folders scanned:  $($meta.FoldersScanned)" -ForegroundColor White
    Write-Host "  Files found:      $($meta.FilesFound)" -ForegroundColor White
    Write-Host "  Queue remaining:  $($meta.QueueCount)" -ForegroundColor White
    Write-Host "  Pass 1 complete:  $($meta.Pass1Complete)" -ForegroundColor White

    $choice = Read-Host "`nContinue from checkpoint? (Y = continue, N = start fresh)"

    if ($choice -match '^[Yy]') {
        if ($meta.DriveId -ne $rootDriveId) {
            Write-Host "  Drive ID differs from checkpoint. Starting fresh." -ForegroundColor Red
            Remove-Item $checkpointMeta -ErrorAction SilentlyContinue
            Remove-Item $checkpointResults -ErrorAction SilentlyContinue
            Remove-Item $checkpointMissing -ErrorAction SilentlyContinue
        }
        else {
            $resuming = $true
            $startedAt = $meta.StartedAt
            $foldersScanned = [int]$meta.FoldersScanned
            $filesFound = [int]$meta.FilesFound
            $totalItemsSeen = [int]$meta.TotalItemsSeen
            $pass1Complete = [bool]$meta.Pass1Complete

            if (Test-Path $checkpointResults) {
                foreach ($row in (Import-Csv $checkpointResults)) {
                    $null = $global:Results.Add([PSCustomObject]@{
                        FullPath     = $row.FullPath
                        FileName     = $row.FileName
                        QuickXorHash = $row.QuickXorHash
                        SizeBytes    = [long]$row.SizeBytes
                        LastModified = $row.LastModified
                        Created      = $row.Created
                        ItemId       = $row.ItemId
                    })
                }
            }

            if (Test-Path $checkpointMissing) {
                foreach ($row in (Import-Csv $checkpointMissing)) {
                    $null = $missingHashItems.Add([PSCustomObject]@{
                        FullPath     = $row.FullPath
                        FileName     = $row.FileName
                        SizeBytes    = [long]$row.SizeBytes
                        LastModified = $row.LastModified
                        Created      = $row.Created
                        ItemId       = $row.ItemId
                        DriveId      = $row.DriveId
                    })
                }
            }

            if (-not $pass1Complete -and $meta.Queue) {
                foreach ($qItem in $meta.Queue) {
                    $queue.Enqueue(@{
                        DriveId     = $qItem.DriveId
                        ItemId      = $qItem.ItemId
                        CurrentPath = $qItem.CurrentPath
                    })
                }
            }

            Write-Host "  Restored $($global:Results.Count) results, $($missingHashItems.Count) missing-hash items, $($queue.Count) queued folders." -ForegroundColor Green
        }
    }
    else {
        Write-Host "Starting fresh scan." -ForegroundColor Cyan
        Remove-Item $checkpointMeta -ErrorAction SilentlyContinue
        Remove-Item $checkpointResults -ErrorAction SilentlyContinue
        Remove-Item $checkpointMissing -ErrorAction SilentlyContinue
    }
}

# --- Initialize fresh scan if not resuming ---
if (-not $resuming) {
    $queue.Enqueue(@{
        DriveId     = $rootDriveId
        ItemId      = "root"
        CurrentPath = ""
    })

    Save-CheckpointMeta -MetaPath $checkpointMeta -DriveId $rootDriveId -StartedAt $startedAt `
        -FoldersScanned 0 -FilesFound 0 -TotalItemsSeen 0 -Pass1Complete $false -Queue $queue
}

$scanStart = [System.Diagnostics.Stopwatch]::StartNew()

# --- BFS scan (Pass 1) ---
if (-not $pass1Complete) {
    if ($resuming) {
        Write-Host "`nResuming BFS scan ($($queue.Count) folders remaining)..." -ForegroundColor Green
    }
    else {
        Write-Host "`nStarting full OneDrive scan..." -ForegroundColor Green
    }

    $resultsBuffer = [System.Collections.Generic.List[PSCustomObject]]::new()
    $missingBuffer = [System.Collections.Generic.List[PSCustomObject]]::new()
    $foldersSinceCheckpoint = 0
    $checkpointInterval = 50

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $driveId = $current.DriveId
        $itemId = $current.ItemId
        $path = $current.CurrentPath

        $foldersScanned++
        $foldersSinceCheckpoint++
        $elapsed = $scanStart.Elapsed.ToString("hh\:mm\:ss")

        Write-Progress -Activity "Scanning OneDrive [$elapsed elapsed]" `
            -Status "Folder: $path" `
            -CurrentOperation "Files: $filesFound | Folders scanned: $foldersScanned | Queue: $($queue.Count)" `
            -Id 1

        if ($foldersScanned % 50 -eq 0) {
            Write-Host "  Progress: $foldersScanned folders scanned, $filesFound files with hash, $($missingHashItems.Count + $missingBuffer.Count) missing hash, $($queue.Count) queued [$elapsed]" -ForegroundColor DarkGray
        }

        $nextLink = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/children?`$select=id,name,file,folder,size,lastModifiedDateTime,createdDateTime&`$top=999"

        do {
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -OutputType PSObject

                foreach ($item in $response.value) {
                    $totalItemsSeen++
                    $fullPath = if ($path) { "$path/$($item.name)" } else { $item.name }

                    if ($item.file) {
                        $hash = Get-NestedValue $item @('file', 'hashes', 'quickXorHash')

                        if ($hash) {
                            $filesFound++
                            $resultObj = [PSCustomObject]@{
                                FullPath     = "/$fullPath"
                                FileName     = $item.name
                                QuickXorHash = $hash
                                SizeBytes    = $item.size
                                LastModified = $item.lastModifiedDateTime
                                Created      = $item.createdDateTime
                                ItemId       = $item.id
                            }
                            $null = $global:Results.Add($resultObj)
                            $null = $resultsBuffer.Add($resultObj)
                        }
                        elseif ($item.size -gt 0) {
                            $missingObj = [PSCustomObject]@{
                                FullPath     = "/$fullPath"
                                FileName     = $item.name
                                SizeBytes    = $item.size
                                LastModified = $item.lastModifiedDateTime
                                Created      = $item.createdDateTime
                                ItemId       = $item.id
                                DriveId      = $driveId
                            }
                            $null = $missingHashItems.Add($missingObj)
                            $null = $missingBuffer.Add($missingObj)
                        }
                    }

                    if ($item.folder) {
                        $queue.Enqueue(@{
                            DriveId     = $driveId
                            ItemId      = $item.id
                            CurrentPath = $fullPath
                        })
                    }
                }

                $nextLink = $response.'@odata.nextLink'
            }
            catch {
                Write-Warning "Error scanning folder '$path': $($_.Exception.Message)"
                $nextLink = $null
            }
        } while ($nextLink)

        if ($foldersSinceCheckpoint -ge $checkpointInterval) {
            Save-CheckpointBatch $resultsBuffer $checkpointResults
            Save-CheckpointBatch $missingBuffer $checkpointMissing
            $resultsBuffer.Clear()
            $missingBuffer.Clear()

            Save-CheckpointMeta -MetaPath $checkpointMeta -DriveId $rootDriveId -StartedAt $startedAt `
                -FoldersScanned $foldersScanned -FilesFound $filesFound -TotalItemsSeen $totalItemsSeen `
                -Pass1Complete $false -Queue $queue

            $foldersSinceCheckpoint = 0
        }
    }

    if ($resultsBuffer.Count -gt 0 -or $missingBuffer.Count -gt 0) {
        Save-CheckpointBatch $resultsBuffer $checkpointResults
        Save-CheckpointBatch $missingBuffer $checkpointMissing
        $resultsBuffer.Clear()
        $missingBuffer.Clear()
    }

    Save-CheckpointMeta -MetaPath $checkpointMeta -DriveId $rootDriveId -StartedAt $startedAt `
        -FoldersScanned $foldersScanned -FilesFound $filesFound -TotalItemsSeen $totalItemsSeen `
        -Pass1Complete $true -Queue $queue

    Write-Progress -Activity "Scanning OneDrive" -Completed -Id 1

    $elapsed = $scanStart.Elapsed.ToString("hh\:mm\:ss")
    Write-Host "`nBFS scan complete [$elapsed]. Found $filesFound files with hash, $($missingHashItems.Count) files with missing hash." -ForegroundColor Green
}
else {
    Write-Host "`nBFS scan already complete (from checkpoint). $($global:Results.Count) results, $($missingHashItems.Count) missing-hash items." -ForegroundColor Green
}

# --- Batch retry for missing hashes (Pass 2) ---
$existingIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($r in $global:Results) { $null = $existingIds.Add($r.ItemId) }

$pass2Items = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($item in $missingHashItems) {
    if (-not $existingIds.Contains($item.ItemId)) {
        $null = $pass2Items.Add($item)
    }
}

$recovered = 0
$stillMissing = 0
$alreadyRecovered = $missingHashItems.Count - $pass2Items.Count

if ($alreadyRecovered -gt 0) {
    Write-Host "  $alreadyRecovered missing-hash items already recovered in previous run." -ForegroundColor Gray
}

if ($pass2Items.Count -gt 0) {
    Write-Host "`nRetrieving missing hashes via batch requests ($($pass2Items.Count) files)..." -ForegroundColor Yellow

    $batchSize = 20
    $recoveredBuffer = [System.Collections.Generic.List[PSCustomObject]]::new()
    $batchesSinceCheckpoint = 0

    for ($i = 0; $i -lt $pass2Items.Count; $i += $batchSize) {
        $endIndex = [math]::Min($i + $batchSize - 1, $pass2Items.Count - 1)
        $batch = $pass2Items[$i..$endIndex]

        $requests = @()
        $requestIndex = 0
        foreach ($file in $batch) {
            $requests += @{
                id     = "$requestIndex"
                method = "GET"
                url    = "/drives/$($file.DriveId)/items/$($file.ItemId)?`$select=id,file"
            }
            $requestIndex++
        }

        $batchBody = @{ requests = $requests } | ConvertTo-Json -Depth 5

        try {
            $batchResponse = Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/`$batch" `
                -Body $batchBody `
                -ContentType "application/json" `
                -OutputType PSObject

            foreach ($resp in $batchResponse.responses) {
                $idx = [int]$resp.id
                $file = $batch[$idx]

                $status = Get-NestedValue $resp @('status')
                if ($status -eq 429) {
                    $retryAfter = Get-NestedValue $resp @('headers', 'Retry-After')
                    $waitSec = if ($retryAfter) { [int]$retryAfter } else { 5 }
                    Write-Host "  Throttled - waiting ${waitSec}s..." -ForegroundColor DarkYellow
                    Start-Sleep -Seconds $waitSec
                    $stillMissing++
                    continue
                }

                $hash = Get-NestedValue $resp @('body', 'file', 'hashes', 'quickXorHash')

                if ($status -eq 200 -and $hash) {
                    $recovered++
                    $resultObj = [PSCustomObject]@{
                        FullPath     = $file.FullPath
                        FileName     = $file.FileName
                        QuickXorHash = $hash
                        SizeBytes    = $file.SizeBytes
                        LastModified = $file.LastModified
                        Created      = $file.Created
                        ItemId       = $file.ItemId
                    }
                    $null = $global:Results.Add($resultObj)
                    $null = $recoveredBuffer.Add($resultObj)
                }
                else {
                    $stillMissing++
                }
            }
        }
        catch {
            Write-Warning "Batch request failed: $($_.Exception.Message)"
            $stillMissing += $batch.Count
        }

        $batchesSinceCheckpoint++
        if ($batchesSinceCheckpoint -ge 5) {
            Save-CheckpointBatch $recoveredBuffer $checkpointResults
            $recoveredBuffer.Clear()
            $batchesSinceCheckpoint = 0
        }

        $batchProgress = [math]::Min(100, [math]::Round(($i + $batchSize) / $pass2Items.Count * 100))
        Write-Progress -Activity "Retrieving missing hashes" `
            -Status "Recovered: $recovered | Still missing: $stillMissing" `
            -PercentComplete $batchProgress -Id 2
    }

    if ($recoveredBuffer.Count -gt 0) {
        Save-CheckpointBatch $recoveredBuffer $checkpointResults
        $recoveredBuffer.Clear()
    }

    Write-Progress -Activity "Retrieving missing hashes" -Completed -Id 2
    Write-Host "  Recovered $recovered hashes. $stillMissing files still have no hash (zero-byte or server-side issue)." -ForegroundColor Gray
}

# --- Export to CSV ---
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$csvPath = Join-Path $scriptDir "OneDrive_Hashes_$timestamp.csv"

$global:Results | Sort-Object FullPath | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# --- Clean up checkpoint (scan completed successfully) ---
Remove-Item $checkpointMeta -ErrorAction SilentlyContinue
Remove-Item $checkpointResults -ErrorAction SilentlyContinue
Remove-Item $checkpointMissing -ErrorAction SilentlyContinue

# --- Final summary ---
$totalElapsed = $scanStart.Elapsed.ToString("hh\:mm\:ss")
Write-Host "`n=== Scan Complete ===" -ForegroundColor Green
Write-Host "  Total time:          $totalElapsed" -ForegroundColor White
Write-Host "  Folders scanned:     $foldersScanned" -ForegroundColor White
Write-Host "  Total items seen:    $totalItemsSeen" -ForegroundColor White
Write-Host "  Files with hash:     $($global:Results.Count)" -ForegroundColor White
if ($missingHashItems.Count -gt 0) {
    Write-Host "  Hashes recovered:    $($recovered + $alreadyRecovered) (from $($missingHashItems.Count) missing)" -ForegroundColor White
}
Write-Host "  Results saved to:    $csvPath" -ForegroundColor Cyan
Write-Host "  Checkpoint cleared." -ForegroundColor Gray

if ($GroupDuplicates) {
    $duplicateGroups = $global:Results |
        Group-Object -Property QuickXorHash |
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
        Write-Host "  Duplicate groups:    $($duplicateGroups.Count)" -ForegroundColor White
        Write-Host "  Total duplicate files: $totalDuplicateFiles" -ForegroundColor White
        Write-Host "  Wasted space:        $wastedMB MB" -ForegroundColor White

        $dupRows = [System.Collections.Generic.List[PSCustomObject]]::new()
        $groupNum = 0
        foreach ($group in $duplicateGroups) {
            $groupNum++
            foreach ($file in ($group.Group | Sort-Object FullPath)) {
                $null = $dupRows.Add([PSCustomObject]@{
                    DuplicateGroup = $groupNum
                    FullPath       = $file.FullPath
                    FileName       = $file.FileName
                    QuickXorHash   = $file.QuickXorHash
                    SizeBytes      = $file.SizeBytes
                    LastModified   = $file.LastModified
                    Created        = $file.Created
                })
            }
        }

        $dupCsvPath = Join-Path $scriptDir "OneDrive_Duplicates_$timestamp.csv"
        $dupRows | Export-Csv -Path $dupCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Duplicates saved to: $dupCsvPath" -ForegroundColor Cyan

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
