# PowerShell-Graph-CalendarViewDeltaSyncWithLogging.ps1
# Initially generated with Copilot, then heavily edited by hand for better structure, error handling, and logging. 
# This will do a Graph calenadarviewdelta query and sync all item.  It can continue in suscessive runs.  
# Note the two files it creates - one is the log and the other to do follow-up syncs.
# Page size and the URL can be adjusted.
# I has switches to control logging of the token, Logging of the json body, beautification (foratting) of the jason content.
# 
# =====================================================================
# App-only (Client Credentials) CalendarView Delta Sync with Paging + Logging
# =====================================================================
# - Auth: OAuth2 client_credentials (application permissions)
# - API:  GET /users/{id}/calendarView/delta?startDateTime=...&endDateTime=...
# - Paging: Prefer: odata.maxpagesize=20; follow @odata.nextLink until @odata.deltaLink
# - Logging: Start-Transcript + explicit request/response logging to a text file.   
# - Adjust settings prior to running.  
# - Note the switches - they control how much is logged and formatting.
# 
# Docs:
# - CalendarView delta endpoint & Prefer header odata.maxpagesize [1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
# - Use full @odata.nextLink for paging [2](https://learn.microsoft.com/en-us/graph/paging)
# - Client credentials token request + scope .default [3](https://learn.microsoft.com/en-us/graph/auth-v2-service)

# -------------------------
# CONFIG (edit these)
# -------------------------
$TenantId       = "YOUR_TENANT_ID_GUID_OR_DOMAIN"
$ClientId       = "YOUR_APP_CLIENT_ID"
$ClientSecret   = "YOUR_CLIENT_SECRET"   # Consider using a secure vault in production

# The mailbox to sync (app-only must use /users/{id|UPN}, not /me) [1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
$UserIdOrUpn    = "user@contoso.com"

# Time range (ISO 8601). Graph expects ISO 8601 strings. [1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
$StartDateTime  = (Get-Date "2026-03-01T00:00:00-05:00").ToString("o")
$EndDateTime    = (Get-Date "2026-04-01T00:00:00-05:00").ToString("o")

# Page size target
$PageSize       = 20

# Log + state files
$LogPath        = "C:\test\calendar_delta_apponly.log"
$DeltaLinkPath  = "C:\test\calendar_deltaLink_apponly.txt"

# Swithes
$LogAuthToken = $false  # Only log if needed, as it contains sensitive info
$LogJasonResponses = $true # Set to $false to skip logging JSON responses (faster, but no visibility into actual data returned)
$BeautifyJson = $true # Set to $false to log raw JSON without reformatting (faster, but less readable)
$DebugWriteHostEnabled = $false # Set to $true to enable extra debug Write-Host lines  
 
# Ensure folder exists
$logDir = Split-Path -Parent $LogPath
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

# Start transcript (captures Write-Host/Write-Output and errors)
Start-Transcript -Path $LogPath -Append

try {
    Write-Host "=== START app-only calendarView delta sync ==="
    Write-Host "User: $UserIdOrUpn"
    Write-Host "Range: $StartDateTime -> $EndDateTime"
    Write-Host "PageSize (Prefer odata.maxpagesize): $PageSize"
    Write-Host "Log: $LogPath"
    Write-Host ""

    # -------------------------
    # Helper: URL-encode query string pieces safely
    # -------------------------
    Add-Type -AssemblyName System.Web
    $encStart = [System.Web.HttpUtility]::UrlEncode($StartDateTime)
    $encEnd   = [System.Web.HttpUtility]::UrlEncode($EndDateTime)

    # Initial delta URL (use /users/{id}/calendarView/delta?startDateTime&endDateTime) [1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
    $initialUrl = "https://graph.microsoft.com/v1.0/users/$UserIdOrUpn/calendarView/delta?startDateTime=$encStart&endDateTime=$encEnd"  # Default Details
    #$initialUrl = "https://graph.microsoft.com/v1.0/users/$UserIdOrUpn/calendarView/delta?startDateTime=$encStart&endDateTime=$encEnd&$select=id,subject,start,end"  # Select only specific fields (optional)
    #$initialUrl = "https://graph.microsoft.com/v1.0/users/$UserIdOrUpn/calendarView/delta?startDateTime=$encStart&endDateTime=$encEnd&`$select=id"
   
    # -------------------------
    # Helper: Write to Host if debugging
    # -------------------------
    function WriteDebug {
        param(
            [Parameter(Mandatory)][string]$Line 
        )
        if ($DebugWriteHostEnabled) {
            Write-Host "DEBUG: " + $Line 
        }
    }

    # -------------------------
    # Helper: Acquire app-only token (client credentials)
    # -------------------------
    function Get-AppOnlyToken {
        param(
            [Parameter(Mandatory)][string]$TenantId,
            [Parameter(Mandatory)][string]$ClientId,
            [Parameter(Mandatory)][string]$ClientSecret
        )

        # Token endpoint + client_credentials flow [3](https://learn.microsoft.com/en-us/graph/auth-v2-service)
        $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

        WriteDebug -Line "2"

        # For client credentials, request scope = https://graph.microsoft.com/.default [3](https://learn.microsoft.com/en-us/graph/auth-v2-service)
        $body = @{
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
            grant_type    = "client_credentials"
        }

        Write-Host "---- TOKEN REQUEST ----"
        Write-Host "POST $tokenUrl"
        Write-Host "Body: grant_type=client_credentials; scope=https://graph.microsoft.com/.default"
        Write-Host ""
        WriteDebug -Line  "3"
        # Use Invoke-WebRequest so we can log raw status/headers/content
        $resp = Invoke-WebRequest -Method POST -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop
        WriteDebug -Line  "3.1"
        Write-Host "Token HTTP Status: $($resp.StatusCode)"
 
        if ($LogAuthToken) {
            Write-Host "Token Response (raw):"
            Write-Host $resp.Content
        }
 
        Write-Host ""
        WriteDebug -Line  "4"
        $json = $resp.Content | ConvertFrom-Json
        if (-not $json.access_token) { throw "No access_token returned from token endpoint." }
        WriteDebug -Line "5"
        return $json.access_token
    }
 

    # -------------------------
    # Helper: Invoke Graph GET and log everything
    # -------------------------
    function Invoke-GraphGetLogged {
        param(
            [Parameter(Mandatory)][string]$Url,
            [Parameter(Mandatory)][string]$AccessToken,
            [int]$MaxPageSize = 20
        )

        $headers = @{
            "Authorization" = "Bearer $AccessToken"
            "Accept"        = "application/json"
            # calendarView/delta supports Prefer: odata.maxpagesize [1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
            "Prefer"        = "odata.maxpagesize=$MaxPageSize"
        }
        WriteDebug -Line  "6"
        Write-Host "---- GRAPH REQUEST ----"
        Write-Host "GET $Url"
        #Write-Host ("Headers: " + ($headers.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" } | Sort-Object | Out-String).Trim()) # this will provide the auth token in logs, so be cautious if you use it

        Write-Host ""

        $resp = Invoke-WebRequest -Method GET -Uri $Url -Headers $headers -ErrorAction Stop

        Write-Host "Graph HTTP Status: $($resp.StatusCode)"
        Write-Host "Graph Response Headers (selected):"
        Write-Host ("  client-request-id: " + ($resp.Headers["client-request-id"]))
        Write-Host ("  request-id: " + ($resp.Headers["request-id"]))
        Write-Host ("  X-AnchorMailbox: " + ($resp.Headers["x-anchormailbox"]))
        Write-Host ("  Date:       " + ($resp.Headers["Date"]))
        Write-Host ""

        WriteDebug -Line "7"
        
        if ($LogJasonResponses) {
            if ($BeautifyJson) {
                Write-Host "Graph Response Body (beautified):"
                $BeautifyJson = $resp.Content | ConvertFrom-Json | ConvertTo-Json -Depth 20
                Write-Host $BeautifyJson
            } else {
                Write-Host "Graph Response Body (raw):"
                Write-Host $resp.Content
            }
        }
        WriteDebug -Line  "8"
         
        Write-Host ""

        $json = $resp.Content | ConvertFrom-Json # Convert to object for easier access in main flow
        return $json
    }

    # -------------------------
    # MAIN FLOW
    # -------------------------
    $accessToken = Get-AppOnlyToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

    Write-Host "Access token acquired (length): $($accessToken.Length)"
    Write-Host ""

    $url = $initialUrl
    $page = 0
    $total = 0
    $deltaLink = $null

    # Optional: collect minimal event info in memory (so we can summarize)
    $events = New-Object System.Collections.Generic.List[object]

    while ($url) {
        $page++
        Write-Host "=== PAGE $page ==="
        WriteDebug -Line "9"
        $data = Invoke-GraphGetLogged -Url $url -AccessToken $accessToken -MaxPageSize $PageSize
        WriteDebug -Line "9.1"
        $countThisPage = 0  
        if ($data.value) {
            $countThisPage = @($data.value).Count
            $total += $countThisPage
            WriteDebug -Line "9.2"
            foreach ($ev in $data.value) {
                $isRemoved = $false
                $removedReason = $null
                if ($ev.PSObject.Properties.Name -contains "@removed") {
                    $isRemoved = $true
                    $removedReason = $ev.'@removed'.reason
                }

                $events.Add([pscustomobject]@{
                    id              = $ev.id
                    $subject        = $ev.PSObject.Properties['subject'] ? $ev.subject : "<MISSING: subject>" # Subject can be null/empty for removed items or if not in the URL's select fields
                    start           = $ev.PSObject.Properties['start.dateTime'] ? $ev.start.dateTime : "<MISSING: start>" # start can be missing depending upon URL used (e.g. if we didn't select it)
                    #start           = $ev.start.dateTime
                    end             = $ev.PSObject.Properties['end.dateTime'] ? $ev.end.dateTime : "<MISSING: end>" # end can be missing depending upon URL used (e.g. if we didn't select it)
                    #end             = $ev.end.dateTime
                    isRemoved       = $isRemoved
                    removedReason   = $removedReason
                })
            }
        }
        WriteDebug -Line "9.10"

        Write-Host ""
        Write-Host " -------------------------------------------------------"    
        Write-Host "Items this page: $countThisPage"
        Write-Host "Total items so far: $total"
        Write-Host " -------------------------------------------------------"   
        Write-Host ""
        WriteDebug -Line "10"
        # Paging rules:
        # - Follow @odata.nextLink until it disappears (use entire URL). [2](https://learn.microsoft.com/en-us/graph/paging)[1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
        if ($data.PSObject.Properties.Name -contains "@odata.nextLink") {
            $url = $data.'@odata.nextLink'
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            Write-Host "nextLink found -> continuing with: $url"
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            continue
        }
        if ($data.PSObject.Properties.Name -contains "@odata.skiptoken") {
            $skiptoken = $data.'@odata.skiptoken'
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            Write-Host "skiptoken found: $skiptoken"
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            continue
        }
        WriteDebug -Line "11"
        # When the round completes, Graph returns @odata.deltaLink. [1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
        # https://learn.microsoft.com/en-us/graph/delta-query-overview# - If you see @odata.deltaLink, it means you've reached the end of the 
        # current changes and can persist this deltaLink for the next incremental sync.
        if ($data.PSObject.Properties.Name -contains "@odata.deltaLink") {
            $deltaLink = $data.'@odata.deltaLink'
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            Write-Host "deltaLink found -> round complete"
            Write-Host "----------------------------------------------------------------------------------------------------------------"
        } else {
            Write-Warning "No @odata.nextLink or @odata.deltaLink returned; stopping defensively."
        }
        WriteDebug -Line "12"
        $url = $null
    }

    # -------------------------
    # Persist deltaLink for next incremental sync
    # -------------------------
    if ($deltaLink) {
        $deltaLink | Out-File -FilePath $DeltaLinkPath -Encoding utf8 -Force
        Write-Host "Saved deltaLink to: $DeltaLinkPath"
        Write-Host ""
    }

    # Summary
    Write-Host "=== SUMMARY ==="
    Write-Host "Pages fetched: $page"
    Write-Host "Total items returned (including @removed tombstones): $total"
    Write-Host "First 10 items (id/subject/start/end/isRemoved):"
     WriteDebug -Line "13"
    $events | Select-Object -First 10 | Format-Table -AutoSize
    WriteDebug -Line "14"
    Write-Host "=== END ==="
}
catch {
    Write-Error ("ERROR: " + $_.Exception.Message)
    Write-Error ("DETAILS: " + $_ | Out-String)
}
finally {
    Stop-Transcript | Out-Null
}
