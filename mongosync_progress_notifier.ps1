# =============================================================================
# Polls mongosync /progress endpoint every X minutes (X shards)
# and notifies via any combination of: Outlook, SMTP, and/or local text file.
# 
# Some included features . . . . 
#   - SSH pre-flight check before each poll cycle to detect SDM tunnel loss
#   - Per-port retry logic (3 retries with 20-second gaps) before declaring error
#   - Consecutive failure counter: alert only after N failed cycles in a row
#   - Clearer error messaging distinguishing tunnel loss from mongosync failure
#
# NOTIFICATION METHOD PARAMS (mix and match freely):
#   -UseOutlook   $true/$false   Try Outlook COM object   (default: $false)
#   -UseSmtp      $true/$false   Try SMTP / Office 365    (default: $false)
#   -UseLocalFile $true/$false   Write timestamped file   (default: $true)
#
# USAGE EXAMPLES:
#   # Default: local file only
#   .\Poll-MongoSync.ps1
#
#   # SMTP + local file (skip Outlook)
#   PowerShell -ExecutionPolicy Bypass -File .\Poll-MongoSync.ps1 -UseSmtp $true -SmtpPassword "pw"
#
#   # Outlook + local file
#   PowerShell -ExecutionPolicy Bypass -File .\Poll-MongoSync.ps1 -UseOutlook $true
#
#   # Local file only (no email at all) - DEFAULT
#   PowerShell -ExecutionPolicy Bypass -File .\Poll-MongoSync.ps1
#
#   # Custom interval and retry settings
#   PowerShell -ExecutionPolicy Bypass -File .\Poll-MongoSync.ps1 -PollIntervalMinutes 30 -MaxRetries 5
# =============================================================================

param(
    # ------------------------------------------------------------------
    # SSH tunnel settings: tunnel to VM)
    # ------------------------------------------------------------------
    [string]$SshHost                = "127.0.0.1",
    [int]$SshPort                   = 10576,

    # ------------------------------------------------------------------
    # mongosync ports 
    # ------------------------------------------------------------------
    [int[]]$MongoSyncPorts          = @(27701, 27702, 27703),

    # ------------------------------------------------------------------
    # Polling interval
    # ------------------------------------------------------------------
    [int]$PollIntervalMinutes       = 15,

    # ------------------------------------------------------------------
    # Retry settings (per port, per poll cycle)
    # When a port returns an empty/error response, the script retries
    # this many times with a delay before declaring it a failure.
    # ------------------------------------------------------------------
    [int]$MaxRetries                = 3,
    [int]$RetryDelaySeconds         = 20,

    # ------------------------------------------------------------------
    # Consecutive failure alerting
    # The script suppresses error notifications until this many poll
    # cycles in a row have failed. This avoids false alarms from
    # transient SDM tunnel timeouts.
    # Set to 1 to alert on the very first failure (original behavior).
    # ------------------------------------------------------------------
    [int]$AlertAfterConsecutiveFails = 2,

    # ------------------------------------------------------------------
    # Notification method toggles  - set each to $true or $false
    # ------------------------------------------------------------------
    [bool]$UseOutlook               = $false,   # Try Outlook COM object
    [bool]$UseSmtp                  = $false,   # Try SMTP (Office 365)
    [bool]$UseLocalFile             = $true,    # Write timestamped text file

    # ------------------------------------------------------------------
    # Email settings (shared by Outlook and SMTP)
    # ------------------------------------------------------------------
    [string[]]$ToEmail              = @("_____@mongodb.com", "_____@mongodb.com"),
    [string]$FromEmail              = "_____@_____.com",

    # ------------------------------------------------------------------
    # SMTP settings (only used when -UseSmtp $true)
    # ------------------------------------------------------------------
    [string]$SmtpServer             = "smtp.office365.com",
    [int]$SmtpPort                  = 587,
    [string]$SmtpUser               = "_____@_____.com",
    [string]$SmtpPassword           = "",       # leave blank to be prompted securely

    # ------------------------------------------------------------------
    # Local file output folder (used when -UseLocalFile $true)
    # ------------------------------------------------------------------
    [string]$LocalOutputFolder      = "C:\temp\MongoSyncProgressMessages"
)

# Shard labels matching ports
$ShardLabels = @{
    27701 = "Coordinator  - Shard 0"
    27702 = "Follower 0   - Shard 1"
    27703 = "Follower 1   - Shard 2"
}

# Sentinel values per API docs:
#   Before collection copy phase, estimatedCopiedBytes=0 and estimatedTotalBytes=1
$PRE_COPY_SENTINEL_TOTAL = 1

# Consecutive failure tracking (across poll cycles)
$script:ConsecutiveFailCycles = 0

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

function Format-Bytes {
    param([object]$Bytes)
    if ($null -eq $Bytes) { return "N/A" }
    $b = [long]$Bytes
    if ($b -eq 0) { return "0 B" }
    if ($b -ge 1TB) { return "{0:N2} TB" -f ($b / 1TB) }
    if ($b -ge 1GB) { return "{0:N2} GB" -f ($b / 1GB) }
    if ($b -ge 1MB) { return "{0:N2} MB" -f ($b / 1MB) }
    if ($b -ge 1KB) { return "{0:N2} KB" -f ($b / 1KB) }
    return "$b B"
}

function Format-Seconds {
    param([object]$Seconds)
    if ($null -eq $Seconds) { return "N/A" }
    $s = [int]$Seconds
    if ($s -ge 3600) { return "{0}h {1}m {2}s" -f [math]::Floor($s/3600), [math]::Floor(($s%3600)/60), ($s%60) }
    if ($s -ge 60)   { return "{0}m {1}s"       -f [math]::Floor($s/60), ($s%60) }
    return "$s s"
}

function Format-Bool {
    param([object]$Val)
    if ($null -eq $Val) { return "N/A" }
    if ([bool]$Val) { return "YES" } else { return "NO" }
}

function Format-PingLatency {
    param([object]$Ms)
    if ($null -eq $Ms)  { return "N/A (not yet measured)" }
    if ($Ms -eq -1)     { return "FAILED (last ping attempt failed)" }
    if ($Ms -eq 0)      { return "< 1 ms" }
    return "$Ms ms"
}

# =============================================================================
# Test-SshTunnel
#   Runs a simple "echo ok" through SSH to verify the tunnel is alive.
#   Returns $true if reachable, $false if the tunnel is down.
#   This distinguishes "SDM session expired" from "mongosync crashed".
# =============================================================================

function Test-SshTunnel {
    try {
        $result = ssh -o "BatchMode=yes" -o "StrictHostKeyChecking=no" `
                      -o "ConnectTimeout=10" `
                      -o "ServerAliveInterval=60" `
                      -o "ServerAliveCountMax=120" `
                      -p $SshPort $SshHost "echo tunnel_ok" 2>$null
        return ($result -and $result.Trim() -eq "tunnel_ok")
    }
    catch {
        return $false
    }
}

# =============================================================================
# Get-ProgressForPort
#   Queries a single mongosync port via SSH with retry logic.
#   Retries up to $MaxRetries times with $RetryDelaySeconds between attempts
#   before declaring the port as failed.
#
#   Handles:
#     - SSH tunnel down (SDM session expired) - distinguished from mongosync down
#     - Empty / unreachable responses (with retries)
#     - API error responses (success=false, error, errorDescription)
#     - IDLE state (all fields except state/canCommit are null per API docs)
#     - PAUSED state (lagTimeSeconds is null per API docs)
#     - Pre-collection-copy sentinel (estimatedTotalBytes=1, estimatedCopiedBytes=0)
#     - estimatedCopiedBytes exceeding estimatedTotalBytes (retried ops)
# =============================================================================

function Get-ProgressForPort {
    param(
        [int]$Port,
        [bool]$TunnelIsAlive
    )
    $label = $ShardLabels[$Port]

    # --- If tunnel is already known to be down, skip SSH and report clearly ---
    if (-not $TunnelIsAlive) {
        return [PSCustomObject]@{
            Port       = $Port
            Label      = $label
            OK         = $false
            IsTunnelErr = $true
            Error      = "Tunnel is DOWN - re-authenticate SDM and reconnect the resource. mongosync itself may still be running on the VM."
            Raw        = ""
        }
    }

    # --- Retry loop ---
    $attempt  = 0
    $raw      = $null
    $lastErr  = ""

    while ($attempt -le $MaxRetries) {
        if ($attempt -gt 0) {
            Write-Host "    [RETRY $attempt/$MaxRetries]  Port $Port - waiting ${RetryDelaySeconds}s..." -ForegroundColor Yellow
            Start-Sleep -Seconds $RetryDelaySeconds
        }
        $attempt++

        try {
            $raw = ssh -o "BatchMode=yes" -o "StrictHostKeyChecking=no" `
                       -o "ConnectTimeout=10" `
                       -o "ServerAliveInterval=60" `
                       -o "ServerAliveCountMax=120" `
                       -p $SshPort $SshHost `
                       "curl -s --connect-timeout 5 http://localhost:$Port/api/v1/progress" 2>$null

            if ($raw) { break }   # Got a response - exit retry loop
            $lastErr = "Empty response from curl on port $Port"
        }
        catch {
            $lastErr = $_.Exception.Message
            $raw     = $null
        }
    }

    # --- All retries exhausted with no response ---
    if (-not $raw) {
        return [PSCustomObject]@{
            Port        = $Port
            Label       = $label
            OK          = $false
            IsTunnelErr = $false
            Error       = "No response after $MaxRetries retries. $lastErr. " +
                          "mongosync may not be running on port $Port, or the VM is unreachable. " +
                          "Check: ps aux | grep mongosync on the VM."
            Raw         = ""
        }
    }

    try {
        $j = $raw | ConvertFrom-Json

        # --- API-level error response (success=false) ---
        if (-not $j.success) {
            $errType = if ($j.error)            { $j.error }            else { "unknown" }
            $errDesc = if ($j.errorDescription) { $j.errorDescription } else { "(no description)" }
            return [PSCustomObject]@{
                Port        = $Port
                Label       = $label
                OK          = $false
                IsTunnelErr = $false
                Error       = "API Error Type: $errType | Description: $errDesc"
                Raw         = $raw
            }
        }

        $p = $j.progress

        # --- Detect special states per API Behavior section ---
        $state      = if ($null -ne $p.state) { $p.state } else { "N/A" }
        $isIdle     = ($state -eq "IDLE")
        $isPaused   = ($state -eq "PAUSED")
        $isCollCopy = ($p.info -eq "collection copy")

        # --- Byte progress ---
        $estTotal   = $null
        $estCopied  = $null
        $pct        = "N/A"
        $byteNote   = ""

        if ($null -ne $p.collectionCopy) {
            $estTotal  = $p.collectionCopy.estimatedTotalBytes
            $estCopied = $p.collectionCopy.estimatedCopiedBytes

            if ($isIdle) {
                $pct      = "N/A"
                $byteNote = "(IDLE state - byte fields are null per API docs)"
            }
            elseif ($null -ne $estTotal -and [long]$estTotal -eq $PRE_COPY_SENTINEL_TOTAL -and
                    $null -ne $estCopied -and [long]$estCopied -eq 0) {
                $pct      = "N/A"
                $byteNote = "(Pre-collection copy phase - byte counts are sentinel values, not real data)"
            }
            elseif ($isCollCopy -and $null -ne $estTotal -and [long]$estTotal -gt $PRE_COPY_SENTINEL_TOTAL -and
                    $null -ne $estCopied) {
                $rawPct = [double]$estCopied / [double]$estTotal * 100
                if ($rawPct -gt 100) {
                    $pct      = "~100% (copiedBytes exceeds totalBytes due to retried operations)"
                    $byteNote = "(Per API docs: estimatedCopiedBytes may exceed estimatedTotalBytes due to retries)"
                }
                else {
                    $pct = "{0:N1}%" -f $rawPct
                }
            }
            else {
                $pct = "N/A (not in collection copy phase)"
            }
        }

        # --- Ping latencies (v1.17+) ---
        $srcPing = $null
        $dstPing = $null
        if ($null -ne $p.source)      { $srcPing = $p.source.pingLatencyMs }
        if ($null -ne $p.destination) { $dstPing = $p.destination.pingLatencyMs }

        # --- Direction mapping ---
        $dirSrc = "N/A"
        $dirDst = "N/A"
        if ($null -ne $p.directionMapping) {
            $dirSrc = $p.directionMapping.Source
            $dirDst = $p.directionMapping.Destination
        }

        # --- lagTimeSeconds: null when PAUSED per API docs ---
        $lagNote = ""
        if ($isPaused -and $null -eq $p.lagTimeSeconds) {
            $lagNote = " (null - expected when PAUSED per API docs)"
        }

        return [PSCustomObject]@{
            Port                    = $Port
            Label                   = $label
            OK                      = $true
            IsTunnelErr             = $false
            Error                   = ""
            State                   = $state
            IsIdle                  = $isIdle
            IsPaused                = $isPaused
            Info                    = if ($null -ne $p.info)  { $p.info }  else { "" }
            CanCommit               = $p.canCommit
            CanWrite                = $p.canWrite
            LagTimeSeconds          = $p.lagTimeSeconds
            LagNote                 = $lagNote
            TotalEventsApplied      = $p.totalEventsApplied
            EstTotalBytes           = $estTotal
            EstCopiedBytes          = $estCopied
            EstTotalBytesHuman      = Format-Bytes $estTotal
            EstCopiedBytesHuman     = Format-Bytes $estCopied
            PctBytesCopied          = $pct
            ByteNote                = $byteNote
            IsCollectionCopy        = $isCollCopy
            EstSecsToCEACatchup     = $p.estimatedSecondsToCEACatchup
            MongosyncID             = $p.mongosyncID
            CoordinatorID           = $p.coordinatorID
            DirSource               = $dirSrc
            DirDestination          = $dirDst
            SrcPingLatencyMs        = $srcPing
            DstPingLatencyMs        = $dstPing
            Raw                     = $raw
        }
    }
    catch {
        return [PSCustomObject]@{
            Port        = $Port
            Label       = $label
            OK          = $false
            IsTunnelErr = $false
            Error       = "JSON parse error: $($_.Exception.Message). Raw: $raw"
            Raw         = $raw
        }
    }
}

# =============================================================================
# Build-ShardBlock
# =============================================================================

function Build-ShardBlock {
    param($r)
    $line = "=" * 60
    $dash = "-" * 60

    # --- Error block ---
    if (-not $r.OK) {
        $tunnelNote = if ($r.IsTunnelErr) {
            "`n  ACTION: Re-authenticate tunnel, then re-click the green icon`n          next to the resource to restore the tunnel."
        } else { "" }
        return @"
$line
PORT $($r.Port)  |  $($r.Label)
  *** ERROR: $($r.Error)$tunnelNote
"@
    }

    # --- IDLE state note ---
    $idleNote = ""
    if ($r.IsIdle) {
        $idleNote = @"

  *** IDLE STATE: Per API docs, all fields except State and
      CanCommit are null when mongosync is in IDLE state.
      This is expected before /start has been called.
"@
    }

    # --- PAUSED state note ---
    $pausedNote = ""
    if ($r.IsPaused) {
        $pausedNote = @"

  *** PAUSED STATE: Per API docs, lagTimeSeconds is null
      when mongosync is in the PAUSED state.
"@
    }

    # --- CEA catchup line (only when present) ---
    $ceaLine = ""
    if ($null -ne $r.EstSecsToCEACatchup) {
        $ceaLine = "  Est. CEA Catchup    : $(Format-Seconds $r.EstSecsToCEACatchup) (ignore after initial lag drop)`n"
    }

    # --- Byte progress section ---
    $byteNoteStr = if ($r.ByteNote) { "`n  NOTE: $($r.ByteNote)" } else { "" }

    if ($r.IsCollectionCopy -and -not $r.IsIdle) {
        $byteSection = @"

  -- Collection Copy Progress --
  Est. Total Bytes    : $($r.EstTotalBytesHuman)  (global across all shards)
  Est. Copied Bytes   : $($r.EstCopiedBytesHuman)  (this shard)
  Pct Bytes Copied    : $($r.PctBytesCopied)  (this shard copied / global total)$byteNoteStr
"@
    }
    else {
        $byteSection = @"

  -- Collection Copy Progress --
  Est. Total Bytes    : $($r.EstTotalBytesHuman)
  Est. Copied Bytes   : $($r.EstCopiedBytesHuman)
  Pct Bytes Copied    : $($r.PctBytesCopied)$byteNoteStr
"@
    }

    return @"
$line
PORT $($r.Port)  |  $($r.Label)
$dash$idleNote$pausedNote

  State               : $($r.State)
  Phase (info)        : $(if ($r.Info) { $r.Info } else { "(none)" })
  Can Commit          : $(Format-Bool $r.CanCommit)
  Can Write           : $(Format-Bool $r.CanWrite)
$dash
  Lag Time            : $(Format-Seconds $r.LagTimeSeconds)$($r.LagNote)
  Total Events Applied: $($r.TotalEventsApplied)
$($ceaLine)$byteSection

  -- Direction --
  Source              : $($r.DirSource)
  Destination         : $($r.DirDestination)

  -- Network Latency --
  Src Ping Latency    : $(Format-PingLatency $r.SrcPingLatencyMs)
  Dst Ping Latency    : $(Format-PingLatency $r.DstPingLatencyMs)

  -- Identity --
  mongosyncID         : $($r.MongosyncID)
  coordinatorID       : $($r.CoordinatorID)
"@
}

# =============================================================================
# Build-OverallByteSummary
# =============================================================================

function Build-OverallByteSummary {
    param([PSCustomObject[]]$Results)

    $copyResults = $Results | Where-Object {
        $_.OK -and
        $null -ne $_.EstCopiedBytes -and
        $null -ne $_.EstTotalBytes -and
        [long]$_.EstTotalBytes -gt $PRE_COPY_SENTINEL_TOTAL
    }
    if ($copyResults.Count -eq 0) { return "" }

    $totalEst    = ($copyResults | Select-Object -First 1).EstTotalBytes
    $totalCopied = 0
    $copyResults | ForEach-Object { $totalCopied += [long]$_.EstCopiedBytes }

    $overallPct = "N/A"
    if ($null -ne $totalEst -and [long]$totalEst -gt 0) {
        $rawPct = [double]$totalCopied / [double]$totalEst * 100
        if ($rawPct -gt 100) {
            $overallPct = "~100% (sum of copiedBytes exceeds totalBytes - retried operations)"
        }
        else {
            $overallPct = "{0:N1}%" -f $rawPct
        }
    }

    $line = "=" * 60
    return @"

$line
OVERALL BYTE PROGRESS  (all shards combined)
  Est. Total Bytes    : $(Format-Bytes $totalEst)  (global - same value reported by each shard)
  Est. Copied (sum)   : $(Format-Bytes $totalCopied)  (sum of estimatedCopiedBytes across all shards)
  Overall Pct Copied  : $overallPct
  NOTES:
    - estimatedTotalBytes is a global value (not per-shard).
    - Overall % = sum of all shards' estimatedCopiedBytes / estimatedTotalBytes x 100
    - Per API docs: estimatedTotalBytes is approximated before migration
      and does not update during sync. Not an accurate progress indicator.
    - estimatedCopiedBytes may exceed estimatedTotalBytes due to retries.
      Progress reaches 100% at end of collection copy phase.
"@
}

# =============================================================================
# Build-MessageBody
# =============================================================================

function Build-MessageBody {
    param([PSCustomObject[]]$Results, [string]$Timestamp, [bool]$TunnelWasAlive,
          [int]$ConsecFails)

    $methods = @()
    if ($UseOutlook)   { $methods += "Outlook" }
    if ($UseSmtp)      { $methods += "SMTP" }
    if ($UseLocalFile) { $methods += "Local File" }
    $methodsStr = if ($methods.Count -gt 0) { $methods -join " | " } else { "None" }

    $line = "=" * 60

    # Tunnel status block
    $tunnelStatus = if ($TunnelWasAlive) {
        "SSH Tunnel  : ALIVE (SSH connectivity verified)"
    } else {
        "SSH Tunnel  : *** DOWN - SDM session may have expired. Re-authenticate SDM and reconnect resource."
    }

    # Consecutive failure warning
    $failWarning = ""
    if ($ConsecFails -gt 0) {
        $failWarning = "`nWARNING: This is consecutive failed poll cycle #$ConsecFails.`n" +
                       "         mongosync processes on the VM are likely still running - this is probably an SDM tunnel issue.`n"
    }

    $header = @"
mongosync Progress Report  -  Environment: _____  (X Shards)
Polled   : $Timestamp
VM SSH   : $SshHost  port $SshPort  (SSH tunnel)
Ports    : 27701 (Coordinator/Shard 0)  |  27702 (Follower/Shard 1)  |  27703 (Follower/Shard 2)
Notify   : $methodsStr
$tunnelStatus$failWarning

"@

    $blocks  = ($Results | ForEach-Object { Build-ShardBlock $_ }) -join "`n`n"
    $overall = Build-OverallByteSummary -Results $Results

    $footer = @"

$line
Next poll in $PollIntervalMinutes minute(s).
(Script on Citrix desktop via SSH SDM tunnel - Ctrl+C to stop)
"@
    return $header + "`n`n" + $blocks + "`n`n" + $overall + "`n`n" + $footer + "`n`n"
}

# =============================================================================
# Get-SubjectLine
# =============================================================================

function Get-SubjectLine {
    param([PSCustomObject[]]$Results, [string]$Timestamp, [bool]$TunnelWasAlive)
    $states  = ($Results | Where-Object { $_.OK } | ForEach-Object { "$($_.Port):$($_.State)" }) -join "  "
    $infos   = ($Results | Where-Object { $_.OK -and $_.Info } | Select-Object -First 1).Info
    $anyErr  = ($Results | Where-Object { -not $_.OK }).Count
    $errTag  = if ($anyErr -gt 0) { " [!] $anyErr ERROR(s)" } else { "" }
    $tunnTag = if (-not $TunnelWasAlive) { " [SDM TUNNEL DOWN]" } else { "" }
    $infoTag = if ($infos) { " | $infos" } else { "" }
    return "[mongosync Status: ]$tunnTag$errTag $states$infoTag -- $Timestamp"
}

# =============================================================================
# Write-LocalFile
# =============================================================================

function Write-LocalFile {
    param([string]$Subject, [string]$Body, [string]$Timestamp)

    if (-not (Test-Path $LocalOutputFolder)) {
        New-Item -ItemType Directory -Path $LocalOutputFolder -Force | Out-Null
        Write-Host "  [FILE]  Created folder: $LocalOutputFolder" -ForegroundColor DarkGray
    }

    $safeTs   = $Timestamp -replace "[ :]", "-"
    $fileName = "mongosync_progress_$safeTs.txt"
    $filePath = Join-Path $LocalOutputFolder $fileName

    $fileContent = "Subject : $Subject`n$Body"
    try {
        $fileContent | Out-File -FilePath $filePath -Encoding UTF8 -Force
        Write-Host "  [FILE]  Saved: $filePath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  [FILE]  ERROR saving file: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# =============================================================================
# =============================================================================
# ConvertTo-HtmlBody
# =============================================================================

function ConvertTo-HtmlBody {
    param([string]$PlainText)
    # Escape HTML special characters
    $escaped = $PlainText `
        -replace '&',  '&amp;' `
        -replace '<',  '&lt;'  `
        -replace '>',  '&gt;'
    return @"
<html>
<body style="margin:0; padding:16px; background-color:#f4f6f8;">
  <pre style="font-family: Consolas, 'Courier New', monospace;
              font-size: 13px;
              color: #1a1a1a;
              background-color: #ffffff;
              border: 1px solid #d0d7de;
              border-radius: 6px;
              padding: 20px 24px;
              max-width: 900px;
              line-height: 1.6;
              white-space: pre-wrap;
              word-wrap: break-word;
              overflow-x: auto;">$escaped</pre>
</body>
</html>
"@
}

# Send-ViaOutlook
# =============================================================================

function Send-ViaOutlook {
    param([string]$Subject, [string]$Body)
    try {
        $ol   = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $mail = $ol.CreateItem(0)
        $mail.To       = ($ToEmail -join "; ")
        $mail.Subject  = $Subject
        $mail.HTMLBody = ConvertTo-HtmlBody -PlainText $Body
        $mail.Send()
        Write-Host "  [OUTLOOK]  Sent to $($ToEmail -join ", ")" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  [OUTLOOK]  Failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

# =============================================================================
# Send-ViaSmtp
# =============================================================================

function Send-ViaSmtp {
    param([string]$Subject, [string]$Body)
    try {
        $secPass = if ($SmtpPassword -ne "") {
            ConvertTo-SecureString $SmtpPassword -AsPlainText -Force
        }
        else {
            Write-Host "  [SMTP]  Enter password for $SmtpUser :" -NoNewline
            Read-Host -AsSecureString
        }
        $cred = New-Object System.Management.Automation.PSCredential($SmtpUser, $secPass)
        $htmlBody = ConvertTo-HtmlBody -PlainText $Body
        Send-MailMessage `
            -From       $FromEmail `
            -To         $ToEmail `
            -Subject    $Subject `
            -Body       $htmlBody `
            -BodyAsHtml `
            -SmtpServer $SmtpServer `
            -Port       $SmtpPort `
            -UseSsl     `
            -Credential $cred `
            -ErrorAction Stop
        Write-Host "  [SMTP]  Sent via $SmtpServer to $($ToEmail -join ", ")" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  [SMTP]  Failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# =============================================================================
# Dispatch-Notifications
# =============================================================================

function Dispatch-Notifications {
    param([string]$Subject, [string]$Body, [string]$Timestamp)

    $fileOk    = $false
    $outlookOk = $false
    $smtpOk    = $false

    if ($UseLocalFile) { $fileOk    = Write-LocalFile -Subject $Subject -Body $Body -Timestamp $Timestamp }
    if ($UseOutlook)   { $outlookOk = Send-ViaOutlook -Subject $Subject -Body $Body }
    if ($UseSmtp)      { $smtpOk    = Send-ViaSmtp    -Subject $Subject -Body $Body }

    $emailEnabled = $UseOutlook -or $UseSmtp
    $emailOk      = $outlookOk -or $smtpOk
    if ($emailEnabled -and -not $emailOk) {
        $fileNote = if ($UseLocalFile -and $fileOk) { " Message saved to local file." } else { "" }
        Write-Host "  [NOTIFY]  All enabled email method(s) failed.$fileNote" -ForegroundColor Yellow
    }
}

# =============================================================================
# STARTUP BANNER
# =============================================================================

$activeMethods = @()
if ($UseOutlook)   { $activeMethods += "Outlook" }
if ($UseSmtp)      { $activeMethods += "SMTP ($SmtpServer)" }
if ($UseLocalFile) { $activeMethods += "Local File ($LocalOutputFolder)" }
$methodBanner = if ($activeMethods.Count -gt 0) { $activeMethods -join "  +  " } else { "NONE - console output only" }

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  mongosync Progress Poller "          -ForegroundColor Cyan
Write-Host "  VM  : $SshHost  SSH port $SshPort  "         -ForegroundColor Cyan
Write-Host "  Poll: every $PollIntervalMinutes minute(s)"             -ForegroundColor Cyan
Write-Host "  Retries per port : $MaxRetries x ${RetryDelaySeconds}s" -ForegroundColor Cyan
Write-Host "  Alert after      : $AlertAfterConsecutiveFails consecutive failed cycles" -ForegroundColor Cyan
Write-Host "  To  : $($ToEmail -join ", ")"                            -ForegroundColor Cyan
Write-Host "  Notify via: $methodBanner"                              -ForegroundColor Cyan
Write-Host "  Press Ctrl+C to stop."                                  -ForegroundColor Cyan
Write-Host "========================================================"  -ForegroundColor Cyan
Write-Host ""

# =============================================================================
# MAIN LOOP
# =============================================================================

while ($true) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp]  Checking SSH tunnel..." -ForegroundColor White

    # ------------------------------------------------------------------
    # STEP 1: Pre-flight SSH tunnel check
    # Determines if SDM session is alive BEFORE attempting any curl calls.
    # ------------------------------------------------------------------
    $tunnelAlive = Test-SshTunnel

    if ($tunnelAlive) {
        Write-Host "  [TUNNEL]  SDM tunnel OK - polling $($MongoSyncPorts.Count) mongosync process(es)..." -ForegroundColor Green
    }
    else {
        Write-Host "  [TUNNEL]  *** SSH tunnel UNREACHABLE - SSH session may have expired." -ForegroundColor Red
        Write-Host "  [TUNNEL]      ACTION: Re-authenticate SSH, re-click the green icon for the resource." -ForegroundColor Red
        Write-Host "  [TUNNEL]      NOTE:   mongosync processes on the VM are likely still running." -ForegroundColor Yellow
    }

    # ------------------------------------------------------------------
    # STEP 2: Poll all ports (passing tunnel status so each port can
    #         return a meaningful error without redundant SSH attempts)
    # ------------------------------------------------------------------
    $results = $MongoSyncPorts | ForEach-Object { Get-ProgressForPort -Port $_ -TunnelIsAlive $tunnelAlive }

    # ------------------------------------------------------------------
    # STEP 3: Determine if this is a fully failed cycle
    # A cycle is "failed" if ALL ports returned errors.
    # (Partial failures - some ports OK, some not - always notify.)
    # ------------------------------------------------------------------
    $allFailed = ($results | Where-Object { $_.OK }).Count -eq 0

    if ($allFailed) {
        $script:ConsecutiveFailCycles++
        Write-Host "  [ALERT]   Consecutive failed cycles: $($script:ConsecutiveFailCycles) / $AlertAfterConsecutiveFails" -ForegroundColor Yellow
    }
    else {
        $script:ConsecutiveFailCycles = 0
    }

    # ------------------------------------------------------------------
    # STEP 4: Decide whether to notify
    #   - Always notify if at least one port is OK (partial or full success)
    #   - Only notify on all-failed if we've hit AlertAfterConsecutiveFails
    #   - This suppresses false alarms from transient SDM timeouts
    # ------------------------------------------------------------------
    $shouldNotify = (-not $allFailed) -or
                    ($allFailed -and $script:ConsecutiveFailCycles -ge $AlertAfterConsecutiveFails)

    # Console summary
    Write-Host ""
    $results | ForEach-Object {
        if ($_.OK) {
            $phaseStr = if ($_.Info) { $_.Info } else { "(no info)" }
            $stateTag = ""
            if ($_.IsIdle)   { $stateTag = "  [IDLE - fields are null]" }
            if ($_.IsPaused) { $stateTag = "  [PAUSED - lagTime is null]" }
            $pctStr = if ($_.IsCollectionCopy -and -not $_.IsIdle) { "  Pct: $($_.PctBytesCopied)" } else { "" }
            Write-Host "  Port $($_.Port)  |  $($_.Label)" -ForegroundColor Cyan
            Write-Host "    State: $($_.State)$stateTag  |  Phase: $phaseStr  |  Lag: $(Format-Seconds $_.LagTimeSeconds)$pctStr"
        }
        else {
            $tunnelTag = if ($_.IsTunnelErr) { " [SDM TUNNEL]" } else { " [MONGOSYNC?]" }
            Write-Host "  Port $($_.Port)  |  $($_.Label)  -$tunnelTag ERROR: $($_.Error)" -ForegroundColor Red
        }
    }
    Write-Host ""

    if ($shouldNotify) {
        $body    = Build-MessageBody -Results $results -Timestamp $timestamp `
                                     -TunnelWasAlive $tunnelAlive `
                                     -ConsecFails $script:ConsecutiveFailCycles
        $subject = Get-SubjectLine   -Results $results -Timestamp $timestamp `
                                     -TunnelWasAlive $tunnelAlive
        Dispatch-Notifications -Subject $subject -Body $body -Timestamp $timestamp
    }
    else {
        Write-Host "  [NOTIFY]  Suppressing notification (failed cycle $($script:ConsecutiveFailCycles)/$AlertAfterConsecutiveFails - waiting for $AlertAfterConsecutiveFails consecutive fails before alerting)." -ForegroundColor DarkYellow
        Write-Host "  [NOTIFY]  This is likely a transient SDM tunnel issue, not a mongosync failure." -ForegroundColor DarkYellow
    }

    Write-Host ""
    $nextPoll = (Get-Date).AddMinutes($PollIntervalMinutes).ToString('HH:mm:ss')
    Write-Host "  Next poll at: $nextPoll  (sleeping $PollIntervalMinutes min...)" -ForegroundColor DarkGray
    Write-Host ""

    Start-Sleep -Seconds ($PollIntervalMinutes * 60)
}
