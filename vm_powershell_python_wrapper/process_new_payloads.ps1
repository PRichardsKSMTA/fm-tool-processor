# process_new_payloads.ps1
# ====================================================
# 1.  Copy‑verify‑delete JSON payloads from OneDrive → C:\Tasks\Input
# 2.  De‑duplicate identical OP/WEEK payloads (keep earliest)
# 3.  Send one “queue started” Teams notification
# 4.  Run run_payload.py (Python) for each JSON
# 5.  Archive payload + result, or move to Failed\
# 6.  **NEW:** Never recycle failed payloads back to Input
# 7.  **NEW:** Weekly clean‑up – delete items >14 days from Failed, Archive, Logs
# 8.  **NEW:** Processing continues with next queue item after any failure
# ====================================================

# ---------------- Configuration ----------------
$OneDrivePayloadDir = 'C:\Users\PowerAutomateSVC\OneDrive - KSM Business Services Inc\KMSTA Automation\FM Tool Automation\Tasks\Input'
$TasksRoot  = 'C:\Tasks\ExcelAutomation'
$InputDir   = Join-Path $TasksRoot 'Input'
$ArchiveDir = Join-Path $TasksRoot 'Archive'
$FailedDir  = Join-Path $TasksRoot 'Failed'
$LogsDir    = Join-Path $TasksRoot 'Logs'
$LastCleanupDir = Join-Path $TasksRoot 'LastCleanup'

$PythonExe  = 'C:\Tools\Python311\python.exe'
$Wrapper    = Join-Path $TasksRoot 'PowerShellWrapper\run_payload.py'

$EndFlowUrl   = 'https://prod-121.westus.logic.azure.com:443/workflows/33a3296f16a64204b5fe524c5a069d77/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ynqKDxuJmE0eWwcBeQ59IefV4vT-9uOPHI2_tl40tDY'
$StartFlowUrl = 'https://prod-33.westus.logic.azure.com:443/workflows/f764275c7b974d62ad48d6a89c078127/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=nignRkxQXZ5TndTCMOZU_Fv5mUEHjRtCxhV9uH1y2JA'

$regex = [regex]'^fm_payload_(\d{14})_(?<OP>.+?)_(?<WEEK>\d{4}-\d{2}-\d{2})\.json$'

# ---------------- Helper: completion notification ----------------
function Invoke-NotificationFlow {
    param(
        [string]$OperationCd,
        [string]$Status,
        [string]$Message,
        [string]$LogPath
    )

    # Encode log (optional)
    $b64 = ''
    $logName = ''
    if ($LogPath -and (Test-Path $LogPath)) {
        $bytes   = [System.IO.File]::ReadAllBytes($LogPath)
        $b64     = [Convert]::ToBase64String($bytes)
        $logName = [IO.Path]::GetFileName($LogPath)
    }

    # Remaining queue snapshot
    $queueFiles = Get-ChildItem -Path $InputDir -Filter 'fm_payload_*.json' -File | Sort-Object Name
    $nextOps = @()
    foreach ($f in $queueFiles) {
        if ($f.Name -match '^fm_payload_\d{14}_(?<op>.+?)_\d{4}-\d{2}-\d{2}\.json$') {
            $nextOps += $matches['op']
        }
    }

    $nextOpString = ''
    if ($nextOps.Count -gt 0) { $nextOpString = $nextOps[0] }

    $payload = @{
        OPERATION_CD         = $OperationCd
        Status               = $Status
        Message              = $Message
        LogFileName          = $logName
        LogFileContent       = $b64
        RemainingQueue       = $nextOps          # always array
        NextOperation        = $nextOpString     # always string
        RemainingQueueLength = [int]$nextOps.Count
    }

    try {
        Invoke-RestMethod -Uri $EndFlowUrl -Method Post `
                          -Body ($payload | ConvertTo-Json -Depth 5) `
                          -ContentType 'application/json'
        Write-Host "Completion notification sent for $OperationCd [$Status]"
    }
    catch {
        Write-Warning "ERROR posting completion flow for $OperationCd"
        if ($_.Exception.Response) {
            $reader = [IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
            Write-Warning "Flow response body: $($reader.ReadToEnd())"
        }
    }
}

# ---------------- Ensure local folders exist ----------------
foreach ($d in @($TasksRoot,$InputDir,$ArchiveDir,$FailedDir,$LogsDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null }
}

# ---------------- Weekly clean‑up (>14 days) ----------------
function Invoke-WeeklyCleanup {
    param(
        [string[]]$Dirs,
        [int]$MaxAgeDays = 14,
        [string]$StampFile
    )

    $runCleanup = $true
    if (Test-Path $StampFile) {
        $lastRun = (Get-Item $StampFile).LastWriteTimeUtc
        if ($lastRun -gt (Get-Date).ToUniversalTime().AddDays(-7)) {
            $runCleanup = $false
        }
    }

    if ($runCleanup) {
        $cutoff = (Get-Date).AddDays(-$MaxAgeDays)
        foreach ($dir in $Dirs) {
            if (-not (Test-Path $dir)) { continue }
            Get-ChildItem -Path $dir -File -Recurse |
                Where-Object { $_.LastWriteTime -lt $cutoff } |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }
        New-Item -Path $StampFile -ItemType File -Force | Out-Null
        Write-Host "Weekly clean‑up executed (files older than $MaxAgeDays days removed)."
    }
}

$CleanupStamp = Join-Path $LastCleanupDir '.last_cleanup'
Invoke-WeeklyCleanup -Dirs @($FailedDir,$ArchiveDir,$LogsDir) -StampFile $CleanupStamp

# ---------------- Copy OneDrive → Input ----------------
if (Test-Path $OneDrivePayloadDir) {
    Get-ChildItem $OneDrivePayloadDir -Filter 'fm_payload_*.json' -File | ForEach-Object {
        $src  = $_.FullName
        $dest = Join-Path $InputDir $_.Name
        try {
            Copy-Item $src $dest -Force -ErrorAction Stop
            if (-not (Test-Path $dest)) { throw "Copy verification failed for $($_.Name)" }
            try { Remove-Item $src -Force -ErrorAction Stop } catch {
                Write-Warning "Copied but could not delete original (locked?): $($_.Name)"
            }
            Write-Host "Copied payload '$($_.Name)' → '$InputDir'"
        } catch { Write-Warning "ERROR copying $($_.Name): $_" }
    }
} else {
    Write-Warning "OneDrive folder not found: $OneDrivePayloadDir"
}

# ---------------- Dedupe by OP + WEEK ----------------
$info = foreach ($f in Get-ChildItem $InputDir -Filter 'fm_payload_*.json' -File) {
    $m = $regex.Match($f.Name)
    if ($m.Success) {
        [pscustomobject]@{
            File      = $f
            TimeStamp = [datetime]::ParseExact($m.Groups[1].Value,'yyyyMMddHHmmss',$null)
            OP        = $m.Groups['OP'].Value
            Week      = $m.Groups['WEEK'].Value
        }
    }
}

$info | Group-Object OP,Week | ForEach-Object {
    $ordered = $_.Group | Sort-Object TimeStamp
    $ordered[1..($ordered.Count-1)] | ForEach-Object {
        Move-Item $_.File.FullName (Join-Path $ArchiveDir ('duplicate_'+$_.File.Name)) -Force
        Write-Host "Archived duplicate '$($_.File.Name)'"
    }
}

# ---------------- Startup flow notification ----------------
$queueFiles = Get-ChildItem $InputDir -Filter 'fm_payload_*.json' -File | Sort-Object Name
if ($StartFlowUrl -and $queueFiles.Count -gt 0) {
    $opsQueue = @()
    foreach ($f in $queueFiles) {
        if ($f.Name -match '^fm_payload_\d{14}_(?<op>.+?)_\d{4}-\d{2}-\d{2}\.json$') {
            $opsQueue += $matches['op']
        }
    }

    $body = @{
        run_id              = ([guid]::NewGuid().ToString().Substring(0,8))
        timestamp_utc       = (Get-Date).ToUniversalTime().ToString("o")
        operations_in_queue = $opsQueue
        current_operation   = $opsQueue[0]
        total               = $opsQueue.Count
    }

    try {
        Invoke-RestMethod -Uri $StartFlowUrl -Method Post `
                          -Body ($body | ConvertTo-Json -Depth 5) `
                          -ContentType 'application/json'
        Write-Host "Startup flow triggered (queue size = $($opsQueue.Count))"
    } catch { Write-Warning "Failed to trigger startup flow: $_" }
}

# ---------------- Main processing loop ----------------
while ((Get-ChildItem $InputDir -Filter 'fm_payload_*.json' -File).Count -gt 0) {

    foreach ($file in Get-ChildItem $InputDir -Filter 'fm_payload_*.json' -File | Sort-Object Name) {
        $inPath    = $file.FullName
        $base      = $file.BaseName
        $timestamp = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
        $m = $regex.Match($file.Name)
        $OperationCd = if ($m.Success) { $m.Groups['OP'].Value } else { 'Unknown' }

        Write-Host "`n▶ Processing payload: $base (OPERATION_CD=$OperationCd)"

        $stderrTemp = [IO.Path]::GetTempFileName()
        $jsonText   = & $PythonExe $Wrapper --input-file $inPath 2> $stderrTemp
        $exitCode   = $LASTEXITCODE

        $status  = ''
        $message = ''
        $logPath = ''

        if ($exitCode -ne 0) {
            $status  = 'Failure'
            $message = "Python exited with code $exitCode"
            $logPath = $stderrTemp
        } else {
            try   { $res = $jsonText | ConvertFrom-Json }
            catch { $status='Failure'; $message='Invalid JSON from Python' }

            if (-not $status) {
                $logPath = $res.Out_strLogPath
                if ($res.Out_boolWorkcompleted) {
                    $rfile = "result_${timestamp}_${base}.json"
                    $res | ConvertTo-Json | Out-File (Join-Path $ArchiveDir $rfile) -Encoding UTF8 -Force
                    Move-Item $inPath (Join-Path $ArchiveDir ("processed_${timestamp}_${base}.json"))
                    $status='Success'; $message='Completed successfully'
                } else {
                    $status='Failure'; $message=$res.Out_strWorkExceptionMessage
                }
            }
        }

        Invoke-NotificationFlow -OperationCd $OperationCd -Status $status -Message $message -LogPath $logPath

        if ($status -eq 'Failure') {
            $failName = "failed_${timestamp}_${base}.json"
            Move-Item $file.FullName (Join-Path $FailedDir $failName) -Force
        }

        Remove-Item $stderrTemp -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    }
}

# ---------------- END OF SCRIPT ----------------
# Note: The previous "Retry failed payloads" section has been intentionally
#       removed so failed files remain in $FailedDir and are *not* re‑queued.
