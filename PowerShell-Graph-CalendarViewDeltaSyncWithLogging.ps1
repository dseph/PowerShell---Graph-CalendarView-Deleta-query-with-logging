# PowerShell-Graph-CalendarViewDeltaSyncWithLogging.ps1
# PowerShell - Graph CalendarView Deleta query with logging sample. 
# It can do a full sync run plus continue for another run.
# Initially generated with Copilot, then heavily edited by hand for better structure, error handling, and logging. 
# This will do a Graph calenadarviewdelta query and sync all item.  It can continue in successive runs.  
# Note the two files it creates - one is the log and the other to do follow-up syncs.
# Page size and the URL can be adjusted.
# It has switches to control logging of the token, logging of the json body, beautification (formatting) of the jason content.
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

# -----------------------------
# CONFIG - Start (edit these)
# -----------------------------
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
$LogJasonResponses = $false # Set to $false to skip logging JSON responses (faster, but no visibility into actual data returned)
$BeautifyJson = $true # Set to $false to log raw JSON without reformatting (faster, but less readable)
$DebugWriteHostEnabled = $false # Set to $true to enable extra debug Write-Host lines  
$EnableAdvancedPriorAndCurrentNextLinkComparison = $false # Set to $true to enable detailed comparison of prior and current nextLink URLs in logs (useful for debugging paging loops)  

# nextlink tracking - Used for checking for repeating nextlinks which could indicate a paging loop.  In production, you would
# typically rely on the script to follow nextLinks and deltaLinks correctly, but this can be helpful for testing and debugging 
# to confirm that the script is picking up where it left off if you are starting with a known nextLink URL.
$Prior_NextLinkUrl = ""  # For testing, you can set this to a known nextLink URL to start from there instead of the initial URL.  In production, this would typically be $null on the first run, and you would rely on the script to follow nextLinks and deltaLinks correctly.
$Prior_NextLinkPage = 0 # For logging/debugging, track the page number where the prior nextLink was found. This is just for visibility in logs to confirm we're picking up where we left off if using a prior nextLink.

# -----------------------------
# CONFIG - End
# -----------------------------

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
            # "X-AnchorMailbox" = $UserIdOrUpn # Optional but can help with performance and consistent routing in some cases [1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
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
                WriteDebug -Line "9.3"
                $isRemoved = $false
                $removedReason = "" # $null
                if ($ev.PSObject.Properties.Name -contains "@removed") {
                    $isRemoved = $true
                    $removedReason = $ev.'@removed'.reason
                }

                WriteDebug -Line "9.4"
                # Get the data ready - make it easeir to debug missing fields by checking if they exist before accessing (e.g. subject can be missing for removed items or if not in select, start/end can be missing if not in select).  This way we won't get errors when trying to access missing fields and we can see in logs when expected fields are missing.
                $use_id = $ev.id
                 WriteDebug -Line "9.4.1"
                $use_subject = $ev.PSObject.Properties['subject'] ? $ev.subject : "<MISSING: subject>" # Subject can be null/empty for removed items or if not in the URL's select fields
                WriteDebug -Line "9.4.2"
                $use_start = $ev.PSObject.Properties['start'] ? $ev.start.dateTime : "<MISSING: start>" # start can be missing depending upon URL used (e.g. if we didn't select it)
                WriteDebug -Line "9.4.3"
                $use_end = $ev.PSObject.Properties['end'] ? $ev.end.dateTime : "<MISSING: end>" # end can be missing depending upon URL used (e.g. if we didn't select it)
                WriteDebug -Line "9.4.4"
                $use_isremoved = $isRemoved
                WriteDebug -Line "9.4.5"
                $use_removedReason = $removedReason

                #$use_id 
                #$use_subject  
                #$use_start  
                #$use_end  
                #$use_isremoved  
                #use_removedReason  

                #https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-pscustomobject?view=powershell-7.5
    
                $myObject = [pscustomobject]@{
                    id     = $use_id 
                    subject     = $use_subject 
                    start = $use_start 
                    end = $use_end
                    isRemoved       = $use_isremoved 
                    removedReason   = $use_removedReason
                }

                $events.Add($myObject) 

                WriteDebug -Line "9.6"
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
        $NextLinkFound = $false
        $SkipTokenFound = $false
        $DeltaLinkFound = $false

        WriteDebug -Line "10.1"
        # - Follow @odata.nextLink until it disappears (use entire URL). [2](https://learn.microsoft.com/en-us/graph/paging)[1](https://learn.microsoft.com/en-us/graph/api/event-delta?view=graph-rest-1.0)
        if ($data.PSObject.Properties.Name -contains "@odata.nextLink") {
            $url = $data.'@odata.nextLink'
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            Write-Host "nextLink found -> continuing with: $url"
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            $NextLinkFound = $true

            WriteDebug -Line "10.2"
            if ($url -eq $Prior_NextLinkUrl) {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                Write-Warning "Warning:Prior and current nextLink URLs match !"
                Write-Warning "The prior and the current nextLink URLs are the same (found on prior page $Prior_NextLinkPage and current page $countThisPage). "
                Write-Warning "This could indicate a paging loop. Please investigate."
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                WriteDebug -Line "10.3"
            }
            else {
                Write-Host "----------------------------------------------------------------------------------------------------------------------------------------------------------"
                if ($Prior_NextLinkUrl -eq "") {
                    #Write-Host "The prior nextLink was found on page $Prior_NextLinkPage is blank, so we don't have a prior nextLink to compare to yet."
                } else {
                     
                    Write-Host "Good News: The current nextLink and the prior nextLink don't match, which is good."  
                    if ($EnableAdvancedPriorAndCurrentNextLinkComparison) {
                        Write-Host "Below is a detailed comparison of the two nextLink URLs to show where they differ."
                        Compare-StringDetailed -Left $Prior_NextLinkUrl -Right $url
                    }
                }   
 
                Write-Host "----------------------------------------------------------------------------------------------------------------------------------------------------------"
                WriteDebug -Line "10.4"
            }
        
             
            # Per documentation there should not be a deltaLink if there's a nextLink, but we'll log if we see both just in case.
            # See:  https://learn.microsoft.com/en-us/graph/delta-query-overview
            #    2.c: A page can't contain both @odata.deltaLink and @odata.nextLink.
            if ($data.PSObject.Properties.Name -contains "@odata.deltaLink") {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                Write-Warning "Both @odata.nextLink and @odata.deltaLink found in response. This is unexpected for Graph API responses. "   
                Write-Warning "Will observe only the deltaLink, which will end processing. However, please investigate the response content "
                Write-Warning "and consider reporting to Microsoft if you see this."
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                #$skiptoken = $data.'@odata.skiptoken'
                WriteDebug -Line "10.5"
                $DeltaLinkFound = $true
            }
            WriteDebug -Line "10.6"
            $Prior_NextLinkUrl  = $url 
            $Prior_NextLinkPage =  $countThisPage

            WriteDebug -Line "10.7"
            continue
        }
        if ($data.PSObject.Properties.Name -contains "@odata.skiptoken") {
            WriteDebug -Line "10.8"
            $skiptoken = $data.'@odata.skiptoken'
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            Write-Host "skiptoken found: $skiptoken"
            Write-Host "----------------------------------------------------------------------------------------------------------------"
            $SkipTokenFound = $true
            WriteDebug -Line "10.9"
            #continue
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
            $DeltaLinkFound = $true
        } else {
            Write-Warning "No @odata.nextLink or @odata.deltaLink returned; stopping defensively."
        }

        <# 
        if ($skiptokenFound -eq $false -and $NextLinkFound -eq $false -and $DeltaLinkFound -eq $false) {
            Write-Warning "No paging or delta links found in response; stopping defensively."
         }
        #>

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

 
<#

function Compare-StringDetailed

.SYNOPSIS
  Compare two strings and explain the differences (case, whole-string, and character-level).

.DESCRIPTION
  - Shows equality results (case-insensitive and case-sensitive).
  - Shows "diff-style" results via Compare-Object (with and without -CaseSensitive).
  - Produces a character-by-character diff (index + left/right chars) and a pointer line.

  Compare-Object behavior:
    <= means value exists only in ReferenceObject (left)
    => means value exists only in DifferenceObject (right)
    == means value exists in both when -IncludeEqual is used
#>
 
function Compare-StringDetailed {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Left,

        [Parameter(Mandatory)]
        [string]$Right
    )

    Write-Host "LEFT : [$Left]"
    Write-Host "RIGHT: [$Right]"
    Write-Host ""
    if ($Left -eq "") {
        Write-Host "Cannot compare because the left string is empty."
    }  
    if ($Right -eq "") {
        Write-Host "Cannot compare because the right string is empty."
    }  

    # 1) Quick equality checks
    $eqCaseInsensitive = ($Left -eq $Right)   # default is case-insensitive for -eq
    $eqCaseSensitive   = ($Left -ceq $Right)  # case-sensitive equals
    Write-Host "Equality:"
    Write-Host ("  -eq  (case-insensitive): {0}" -f $eqCaseInsensitive)
    Write-Host ("  -ceq (case-sensitive)  : {0}" -f $eqCaseSensitive)
    Write-Host ""

    # 2) "Diff-style" using Compare-Object
    # Compare-Object compares two sets of objects and indicates which side they appear on. [1](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/compare-object?view=powershell-7.5)[2](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/compare-object?view=powershell-7.5)
    Write-Host "Compare-Object (default: case-insensitive):"
    Compare-Object -ReferenceObject @($Left) -DifferenceObject @($Right) -IncludeEqual |
        Format-Table -AutoSize
    Write-Host ""

    Write-Host "Compare-Object (-CaseSensitive):"
    Compare-Object -ReferenceObject @($Left) -DifferenceObject @($Right) -IncludeEqual -CaseSensitive |
        Format-Table -AutoSize
    Write-Host ""

    # 3) Character-by-character diff (best for "where does it differ?")
    # Handles different lengths cleanly by comparing up to max length.
    $maxLen = [Math]::Max($Left.Length, $Right.Length)

    $diffs = for ($i = 0; $i -lt $maxLen; $i++) {
        $lChar = if ($i -lt $Left.Length)  { $Left[$i] }  else { $null }
        $rChar = if ($i -lt $Right.Length) { $Right[$i] } else { $null }

        if ($lChar -cne $rChar) {
            [pscustomobject]@{
                Index    = $i
                LeftChar = if ($null -eq $lChar) { "<END>" } else { [string]$lChar }
                RightChar= if ($null -eq $rChar) { "<END>" } else { [string]$rChar }
            }
        }
    }

    if (-not $diffs) {
        Write-Host "Character-level diff: No differences."
        return
    }

    Write-Host "Character-level differences (case-sensitive):"
    $diffs | Format-Table -AutoSize
    Write-Host ""

    # Create a visual pointer line showing where differences occur.
    # Example:
    # LEFT : PowerShell rocks
    # RIGHT: Powershell rocks
    #        ^            (caret under differing char)
    $pointer = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $maxLen; $i++) {
        $lChar = if ($i -lt $Left.Length)  { $Left[$i] }  else { $null }
        $rChar = if ($i -lt $Right.Length) { $Right[$i] } else { $null }
        [void]$pointer.Append( ($(if ($lChar -cne $rChar) { '^' } else { ' ' })) )
    }

    Write-Host "Diff pointer (^) positions:"
    Write-Host "LEFT : $Left"
    Write-Host "RIGHT: $Right"
    Write-Host ("      " + $pointer.ToString())
}

 
