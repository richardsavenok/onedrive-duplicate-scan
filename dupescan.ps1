# ================================================
# ROBUST OneDrive Full Scan - QuickXorHash + Path
# Uses raw Graph API with batch retry for missing hashes
# Works 100% in the cloud - no local files downloaded
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

# --- Initialize scan ---
$global:Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$missingHashItems = [System.Collections.Generic.List[PSCustomObject]]::new()

$queue = [System.Collections.Generic.Queue[hashtable]]::new()
$queue.Enqueue(@{
    DriveId     = $rootDriveId
    ItemId      = "root"
    CurrentPath = ""
})

$foldersScanned = 0
$filesFound = 0
$totalItemsSeen = 0
$scanStart = [System.Diagnostics.Stopwatch]::StartNew()

Write-Host "`nStarting full OneDrive scan..." -ForegroundColor Green

# --- BFS scan (Pass 1) ---
while ($queue.Count -gt 0) {
    $current = $queue.Dequeue()
    $driveId = $current.DriveId
    $itemId = $current.ItemId
    $path = $current.CurrentPath

    $foldersScanned++
    $elapsed = $scanStart.Elapsed.ToString("hh\:mm\:ss")

    Write-Progress -Activity "Scanning OneDrive [$elapsed elapsed]" `
        -Status "Folder: $path" `
        -CurrentOperation "Files: $filesFound | Folders scanned: $foldersScanned | Queue: $($queue.Count)" `
        -Id 1

    if ($foldersScanned % 50 -eq 0) {
        Write-Host "  Progress: $foldersScanned folders scanned, $filesFound files with hash, $($missingHashItems.Count) missing hash, $($queue.Count) queued [$elapsed]" -ForegroundColor DarkGray
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
                        $null = $global:Results.Add([PSCustomObject]@{
                            FullPath       = "/$fullPath"
                            FileName       = $item.name
                            QuickXorHash   = $hash
                            SizeBytes      = $item.size
                            LastModified   = $item.lastModifiedDateTime
                            Created        = $item.createdDateTime
                            ItemId         = $item.id
                        })
                    }
                    elseif ($item.size -gt 0) {
                        $null = $missingHashItems.Add([PSCustomObject]@{
                            FullPath       = "/$fullPath"
                            FileName       = $item.name
                            SizeBytes      = $item.size
                            LastModified   = $item.lastModifiedDateTime
                            Created        = $item.createdDateTime
                            ItemId         = $item.id
                            DriveId        = $driveId
                        })
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
}

Write-Progress -Activity "Scanning OneDrive" -Completed -Id 1

$elapsed = $scanStart.Elapsed.ToString("hh\:mm\:ss")
Write-Host "`nBFS scan complete [$elapsed]. Found $filesFound files with hash, $($missingHashItems.Count) files with missing hash." -ForegroundColor Green

# --- Batch retry for missing hashes (Pass 2) ---
$recovered = 0
$stillMissing = 0

if ($missingHashItems.Count -gt 0) {
    Write-Host "`nRetrieving missing hashes via batch requests ($($missingHashItems.Count) files)..." -ForegroundColor Yellow

    $batchSize = 20

    for ($i = 0; $i -lt $missingHashItems.Count; $i += $batchSize) {
        $endIndex = [math]::Min($i + $batchSize - 1, $missingHashItems.Count - 1)
        $batch = $missingHashItems[$i..$endIndex]

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
                    $null = $global:Results.Add([PSCustomObject]@{
                        FullPath       = $file.FullPath
                        FileName       = $file.FileName
                        QuickXorHash   = $hash
                        SizeBytes      = $file.SizeBytes
                        LastModified   = $file.LastModified
                        Created        = $file.Created
                        ItemId         = $file.ItemId
                    })
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

        $batchProgress = [math]::Min(100, [math]::Round(($i + $batchSize) / $missingHashItems.Count * 100))
        Write-Progress -Activity "Retrieving missing hashes" `
            -Status "Recovered: $recovered | Still missing: $stillMissing" `
            -PercentComplete $batchProgress -Id 2
    }

    Write-Progress -Activity "Retrieving missing hashes" -Completed -Id 2
    Write-Host "  Recovered $recovered hashes. $stillMissing files still have no hash (zero-byte or server-side issue)." -ForegroundColor Gray
}

# --- Export to CSV ---
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$csvPath = Join-Path $scriptDir "OneDrive_Hashes_$timestamp.csv"

$global:Results | Sort-Object FullPath | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# --- Final summary ---
$totalElapsed = $scanStart.Elapsed.ToString("hh\:mm\:ss")
Write-Host "`n=== Scan Complete ===" -ForegroundColor Green
Write-Host "  Total time:          $totalElapsed" -ForegroundColor White
Write-Host "  Folders scanned:     $foldersScanned" -ForegroundColor White
Write-Host "  Total items seen:    $totalItemsSeen" -ForegroundColor White
Write-Host "  Files with hash:     $($global:Results.Count)" -ForegroundColor White
if ($missingHashItems.Count -gt 0) {
    Write-Host "  Hashes recovered:    $recovered (from $($missingHashItems.Count) missing)" -ForegroundColor White
}
Write-Host "  Results saved to:    $csvPath" -ForegroundColor Cyan

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
