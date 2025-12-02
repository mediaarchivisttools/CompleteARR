<#
    ############################################################
    # CompleteARR_SONARR_SeriesEngine.ps1
    #
    # 🎯 MAIN ENGINE FOR CompleteARR SONARR
    #
    # This is the core script that manages your series organization.
    #
    # 🔧 WHAT IT DOES:
    #   - Loads your CompleteARR SONARR settings
    #   - Connects to your Sonarr server
    #   - For each media category (Adult, Family, Anime, etc.):
    #       * Checks "Incomplete" series: Are they ready to be "Complete"?
    #       * Checks "Complete" series: Do they need to go back to "Incomplete"?
    #       * Manages episode monitoring (specials vs regular episodes)
    #       * Moves series between folders automatically
    #
    # 📊 PERFORMANCE NOTE:
    #   This script checks episodes for EVERY series in your "Complete" profiles.
    #   If you have a very large library (5,000+ series), this can be slow.
    #   Consider increasing the 'throttleMs' setting in your config to avoid
    #   overwhelming your Sonarr server with too many API calls at once.
    #
    # 🚀 USAGE:
    #   Usually run by CompleteARR_SONARR_Launcher.ps1, but you can also run it directly:
    #   pwsh .\CompleteARR_Scripts\CompleteARR_SONARR_SeriesEngine.ps1 -ConfigPath 'X:\Path\CompleteARR_SONARR_Settings.yml'
    #
    ############################################################
#>

[CmdletBinding()]
param(
    [string]$ConfigPath
)

# ------------------------------------------------------------
# INTERNAL GLOBALS
# ------------------------------------------------------------

# Determine CompleteARR root folder (one level above the Scripts folder this engine lives in).
$thisScriptPath = $MyInvocation.MyCommand.Path
$thisScriptDir  = Split-Path -Path $thisScriptPath -Parent
$Script:CompleteARR_ScriptRoot = Split-Path -Path $thisScriptDir -Parent

$Global:CompleteARR_Config      = $null
$Global:CompleteARR_LogFilePath = $null
$Global:CompleteARR_LogMinLevel = 'Info'
$Global:CompleteARR_LogToFile   = $true
$Global:CompleteARR_LogToConsole= $true
$Global:CompleteARR_UseColors   = $true
$Global:CompleteARR_ThrottleMs  = 200
$Global:CompleteARR_Behavior    = $null

$Global:CompleteARR_Summary = [PSCustomObject]@{
    SeriesChecked               = 0
    IncompleteSeriesSeen        = 0
    CompleteSeriesSeen          = 0
    Promotions                  = 0
    Demotions                   = 0
    AlreadyCorrect              = 0
    SkippedDueToErrors          = 0
    SkippedDueToFilter          = 0
    MonitoredFlagsChanged       = 0
    SpecialsUnmonitoredInInc    = 0
    SpecialsMonitoredInComplete = 0
    EpisodeMonitorChanges       = 0
    RootCorrections             = 0
}

# ------------------------------------------------------------
# Logging
# ------------------------------------------------------------

function Get-LogLevelRank {
    param([string]$Level)

    switch ($Level.ToUpperInvariant()) {
        'TRACE'   { return 0 }
        'DEBUG'   { return 1 }
        'INFO'    { return 2 }
        'WARNING' { return 3 }
        'ERROR'   { return 4 }
        'SUCCESS' { return 5 }
        'FILE'    { return 6 }
        'PROMOTION' { return 7 }
        'DEMOTION'  { return 7 }
        default   { return 2 }
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Level,
        [Parameter(Mandatory)][string]$Message,
        [string]$HighlightText = $null,
        [string]$HighlightColor = 'Green'
    )

    $minRank  = Get-LogLevelRank $Global:CompleteARR_LogMinLevel
    $thisRank = Get-LogLevelRank $Level
    if ($thisRank -lt $minRank) { return }

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line      = "[{0}] [{1}] {2}" -f $timestamp, $Level.ToUpperInvariant(), $Message

    # Always write to console (color-coded)
    if ($Global:CompleteARR_UseColors) {
        $origColor = $Host.UI.RawUI.ForegroundColor
        
        # Handle highlighted text if specified
        if ($HighlightText -and $Message -match [regex]::Escape($HighlightText)) {
            # Get the base color for this log level
            $baseColor = switch ($Level.ToUpperInvariant()) {
                'TRACE'   { 'DarkGray' }
                'DEBUG'   { 'Gray' }
                'INFO'    { 'Cyan' }
                'WARNING' { 'Yellow' }
                'ERROR'   { 'Red' }
                'SUCCESS' { 'Green' }
                'FILE'    { 'Magenta' }
                'PROMOTION' { 'Magenta' }
                'DEMOTION'  { 'DarkYellow' }
                default   { $origColor }
            }
            
            # Split the message and write with different colors
            $parts = [regex]::Split($Message, "($([regex]::Escape($HighlightText)))", "IgnoreCase")
            
            # Write timestamp and level prefix
            Write-Host "[$timestamp] [$($Level.ToUpperInvariant())] " -NoNewline -ForegroundColor $baseColor
            
            # Write message parts with appropriate colors
            foreach ($part in $parts) {
                if ($part -eq $HighlightText) {
                    Write-Host $part -NoNewline -ForegroundColor $HighlightColor
                } else {
                    Write-Host $part -NoNewline -ForegroundColor $baseColor
                }
            }
            Write-Host ""  # New line
        } else {
            # Standard single-color output
            switch ($Level.ToUpperInvariant()) {
                'TRACE'   { $Host.UI.RawUI.ForegroundColor = 'DarkGray' }
                'DEBUG'   { $Host.UI.RawUI.ForegroundColor = 'Gray' }
                'INFO'    { $Host.UI.RawUI.ForegroundColor = 'Cyan' }
                'WARNING' { $Host.UI.RawUI.ForegroundColor = 'Yellow' }
                'ERROR'   { $Host.UI.RawUI.ForegroundColor = 'Red' }
                'SUCCESS' { $Host.UI.RawUI.ForegroundColor = 'Green' }
                'FILE'    { $Host.UI.RawUI.ForegroundColor = 'Magenta' }
                'PROMOTION' { $Host.UI.RawUI.ForegroundColor = 'Magenta' }
                'DEMOTION'  { $Host.UI.RawUI.ForegroundColor = 'DarkYellow' }
                default   { $Host.UI.RawUI.ForegroundColor = $origColor }
            }
            Write-Host $line
            $Host.UI.RawUI.ForegroundColor = $origColor
        }
    }
    else {
        Write-Host $line
    }

    # File logging still respects logToFile + path (no colors in file)
    if ($Global:CompleteARR_LogToFile -and $Global:CompleteARR_LogFilePath) {
        Add-Content -Path $Global:CompleteARR_LogFilePath -Value $line
    }
}


function Import-CompleteARRConfig {
    param([string]$Path)

    if (-not $Path) {
        $Path = Join-Path $Script:CompleteARR_ScriptRoot 'CompleteARR_Settings' 'CompleteARR_SONARR_Settings.yml'
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log 'ERROR' "Config file not found: $Path"
        throw "Config file not found: $Path"
    }

    $raw = Get-Content -LiteralPath $Path -Raw
    $cfg = $raw | ConvertFrom-Yaml
    return ,@($cfg, $Path)
}

function Initialize-LoggingFromConfig {
    param(
        [pscustomobject]$Config,
        [string]$ConfigPath
    )

    $logging = $Config.logging

    $logFileName = $logging.logFileName
    if (-not $logFileName) { $logFileName = 'CompleteARR.log' }

    $timestamp = (Get-Date).ToString('yyyy-MM-dd_HHmm')
    $ext       = [System.IO.Path]::GetExtension($logFileName)
    if (-not $ext) { $ext = '.log' }

    $logsRoot = Join-Path $Script:CompleteARR_ScriptRoot 'CompleteARR_Logs'
    if (-not (Test-Path -LiteralPath $logsRoot)) {
        New-Item -Path $logsRoot -ItemType Directory -Force | Out-Null
    }

    # Use script-specific name for log file
    $scriptName = "CompleteARR_SONARR_SeriesEngine"
    $logFile = Join-Path $logsRoot ("{0}_{1}{2}" -f $scriptName, $timestamp, $ext)

    $Global:CompleteARR_LogFilePath  = $logFile
    $Global:CompleteARR_LogMinLevel  = $logging.minLevel    ?? 'Info'
    $Global:CompleteARR_LogToFile    = $logging.logToFile   ?? $true
    $Global:CompleteARR_LogToConsole = $logging.logToConsole?? $true
    $Global:CompleteARR_UseColors    = $logging.useColors   ?? $true
    $Global:CompleteARR_ThrottleMs   = $logging.throttleMs  ?? 200
    $Global:CompleteARR_Behavior     = $Config.behavior

    Write-Log 'FILE' ("Using configuration file: {0}" -f $ConfigPath)
    Write-Log 'FILE' ("Log file: {0}" -f $logFile)
}


# ------------------------------------------------------------
# SONARR API HELPERS
# ------------------------------------------------------------

function Invoke-SonarrApi {
    <#
        Generic Sonarr API wrapper.

        - $Path is a relative path like "system/status" or "episode"
          (we automatically prepend "api/v3/" unless you already did).
        - $Query is a hashtable of query string parameters.
        - $Body is a PSObject that will be converted to JSON for POST/PUT.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PUT','DELETE')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Path,

        [hashtable]$Query = $null,
        [object]$Body = $null,
        [string]$ErrorContext = ''
    )

    $baseUrl = $Global:CompleteARR_Config.sonarr.url
    if (-not $baseUrl) {
        throw "Sonarr base URL not configured (sonarr.url in YAML)."
    }
    $baseUrl = $baseUrl.TrimEnd('/')

    # If the path is a full URL, use it directly.
    if ($Path -match '^https?://') {
        $uri = $Path
    }
    else {
        $pathPart = $Path.TrimStart('/')

        if (-not $pathPart.ToLower().StartsWith('api/')) {
            $pathPart = "api/v3/$pathPart"
        }

        $uri = "$baseUrl/$pathPart"
    }

    # Append query string correctly: use "?" if none yet, otherwise "&"
    if ($Query -and $Query.Count -gt 0) {
        $qsPairs = @()
        foreach ($kv in $Query.GetEnumerator()) {
            $k = [uri]::EscapeDataString([string]$kv.Key)
            $v = [uri]::EscapeDataString([string]$kv.Value)
            $qsPairs += "$k=$v"
        }
        $qs  = $qsPairs -join '&'
        $sep = $uri.Contains('?') ? '&' : '?'
        $uri = "$uri$sep$qs"
    }

    Write-Log 'DEBUG' "SONARR $Method $uri"

    $headers = @{
        'X-Api-Key' = $Global:CompleteARR_Config.sonarr.apiKey
    }

    $invokeParams = @{
        Method      = $Method
        Uri         = $uri
        Headers     = $headers
        ErrorAction = 'Stop'
    }

    if ($Body) {
        $invokeParams['ContentType'] = 'application/json'
        $invokeParams['Body']        = ($Body | ConvertTo-Json -Depth 10)
    }

    try {
        if ($Global:CompleteARR_ThrottleMs -gt 0) {
            Start-Sleep -Milliseconds $Global:CompleteARR_ThrottleMs
        }
        return Invoke-RestMethod @invokeParams
    }
    catch {
        $ctx = if ($ErrorContext) { $ErrorContext } else { "$Method $Path" }
        Write-Log 'ERROR' ("Sonarr API error ({0}): {1}" -f $ctx, $_.Exception.Message)
        throw
    }
}

function Get-SonarrSystemStatus {
    return Invoke-SonarrApi -Method 'GET' -Path 'system/status' -ErrorContext 'system/status'
}

function Get-SonarrQualityProfiles {
    return Invoke-SonarrApi -Method 'GET' -Path 'qualityprofile' -ErrorContext 'qualityprofile'
}

function Get-SonarrSeries {
    return Invoke-SonarrApi -Method 'GET' -Path 'series' -ErrorContext 'series'
}

function Get-SonarrEpisodesForSeries {
    param(
        [int]$SeriesId
    )

    $result = Invoke-SonarrApi `
        -Method 'GET' `
        -Path 'episode' `
        -Query @{ seriesId = $SeriesId } `
        -ErrorContext ("episode?seriesId={0}" -f $SeriesId)

    if ($null -eq $result) {
        Write-Log 'DEBUG' ("  No episodes returned for seriesId={0}." -f $SeriesId)
        return @()
    }

    Write-Log 'DEBUG' ("  Found {0} episodes for seriesId={1}." -f $result.Count, $SeriesId)
    return $result
}

function Update-SonarrSeries {
    param(
        [pscustomobject]$Series,
        [string]$Context
    )

    if ($Global:CompleteARR_Behavior.dryRun) {
        Write-Log 'FILE' ("[DRY RUN] Would PUT series/{0} ({1})" -f $Series.id, $Context)
        return
    }

    $path  = "series/{0}" -f $Series.id
    $query = @{ moveFiles = 'true' }   # Always move files when path/root/profile change

    Write-Log 'DEBUG' ("SONARR PUT api/v3/{0}?moveFiles=true" -f $path)
    $null = Invoke-SonarrApi -Method 'PUT' -Path $path -Query $query -Body $Series -ErrorContext $Context
}

function Update-SonarrEpisode {
    param(
        [pscustomobject]$Episode,
        [string]$Context
    )

    if ($Global:CompleteARR_Behavior.dryRun) {
        Write-Log 'FILE' ("[DRY RUN] Would PUT episode/{0} ({1})" -f $Episode.id, $Context)
        return
    }

    $path = "episode/{0}" -f $Episode.id
    Write-Log 'DEBUG' ("SONARR PUT api/v3/{0}" -f $path)
    $null = Invoke-SonarrApi -Method 'PUT' -Path $path -Body $Episode -ErrorContext $Context
}

# ------------------------------------------------------------
# EPISODE / SERIES LOGIC
# ------------------------------------------------------------

function Get-SeriesEpisodeStats {
        <#
        📊 ANALYZES A SERIES' EPISODES TO DETERMINE COMPLETION STATUS
        
        This function examines all episodes in a series and figures out:
        - Which regular episodes (non-specials) have aired
        - Which aired episodes are missing files
        - Which missing episodes are old enough to trigger demotion
        
        🔍 HOW IT DECIDES IF AN EPISODE HAS "AIRED":
        1. First priority: Sonarr's built-in 'hasAired' flag (most reliable)
        2. Fallback: Compare air date to current time
        3. Special handling for episodes with unknown air dates
        
        ⚙️  CONFIGURATION NOTES:
        - 'graceDays' setting determines how long to wait before demoting a series
        - 'treatUnknownAirDateAsOld' handles episodes with missing air dates
        
        💡 TIP: The script only considers REGULAR episodes (not specials) when
                deciding if a series is complete, unless you change the settings.
    #>
    param(
        [pscustomobject]$Series,
        [object[]]$Episodes,
        [datetime]$NowUtc,
        [pscustomobject]$Behavior
    )

    $graceDays         = [int]$Behavior.graceDays
    $treatUnknownAsOld = [bool]$Behavior.treatUnknownAirDateAsOld

    # Use LINQ-style filtering for better performance
    $nonSpecial = [System.Collections.Generic.List[object]]::new()
    $specials   = [System.Collections.Generic.List[object]]::new()

    foreach ($ep in $Episodes) {
        if ($ep.seasonNumber -eq 0) {
            $specials.Add($ep)
        }
        else {
            $nonSpecial.Add($ep)
        }
    }

    $airedNonSpecial  = [System.Collections.Generic.List[object]]::new()
    $missingAired     = [System.Collections.Generic.List[object]]::new()
    $missingPastGrace = [System.Collections.Generic.List[object]]::new()

    foreach ($ep in $nonSpecial) {
        # Use modern PowerShell null-conditional operators
        $airDateUtc = $null
        if ($ep.airDateUtc) {
            try {
                $airDateUtc = [datetime]::Parse($ep.airDateUtc).ToUniversalTime()
            } catch { }
        }
        elseif ($ep.airDate) {
            try {
                $airDateUtc = [datetime]::Parse($ep.airDate).ToUniversalTime()
            } catch { }
        }

        # Use modern property access
        $hasAiredFlag = $ep.hasAired ?? $null

        $isAired        = $false
        $olderThanGrace = $false

        # 1) Prefer Sonarr's hasAired flag
        if ($null -ne $hasAiredFlag) {
            $isAired = $hasAiredFlag
        }
        # 2) Fall back to airDate comparison
        elseif ($airDateUtc) {
            if ($airDateUtc -le $NowUtc) {
                $isAired = $true
            }
        }
        # 3) No date and no hasAired -> treat as unreleased by default
        elseif ($treatUnknownAsOld) {
            $isAired        = $true
            $olderThanGrace = $true
        }

        # If we consider it aired and we have a date, check grace window
        if ($isAired -and $airDateUtc) {
            $deltaDays = ($NowUtc - $airDateUtc).TotalDays
            if ($deltaDays -ge $graceDays) {
                $olderThanGrace = $true
            }
        }

        if ($isAired) {
            $airedNonSpecial.Add($ep)

            if (-not $ep.hasFile) {
                $missingAired.Add($ep)
                if ($olderThanGrace) {
                    $missingPastGrace.Add($ep)
                }
            }
        }
    }

    # Calculate specials statistics
    $totalSpecials = $specials.Count
    $monitoredSpecials = ($specials | Where-Object { $_.monitored }).Count
    $completeSpecials = ($specials | Where-Object { $_.hasFile }).Count

    return [PSCustomObject]@{
        NonSpecialEpisodes                 = $nonSpecial
        SpecialEpisodes                    = $specials
        AiredNonSpecialEpisodes            = $airedNonSpecial
        MissingAiredNonSpecialEpisodes     = $missingAired
        MissingPastGraceNonSpecialEpisodes = $missingPastGrace
        TotalSpecials                      = $totalSpecials
        MonitoredSpecials                  = $monitoredSpecials
        CompleteSpecials                   = $completeSpecials
    }
}

function Update-MonitoringStateForSeries {
        <#
        🎯 MANAGES EPISODE MONITORING BASED ON COMPLETION STATUS
        
        This function ensures episodes are monitored/unmonitored correctly:
        
        For INCOMPLETE series:
        - Regular episodes: Always monitored (so they get downloaded)
        - Special episodes: Unmonitored (so they don't block completion)
        
        For COMPLETE series:
        - Regular episodes: Always monitored (for new seasons/episodes)
        - Special episodes: Monitored (so you can enjoy bonus content)
        
        🔧 WHY THIS MATTERS:
        - Prevents specials from stopping a series from being marked "Complete"
        - Ensures you get new episodes when they air
        - Lets you enjoy specials once the main series is complete
        
        💡 TIP: You can customize this behavior in the settings file!
    #>
    param(
        [pscustomobject]$Series,
        [object[]]$Episodes,
        [bool]$IsInCompleteProfile,
        [pscustomobject]$Behavior
    )

    $changes = 0

    $monitorNonSpecials              = [bool]$Behavior.monitorNonSpecials
    $unmonitorSpecialsWhenIncomplete = [bool]$Behavior.unmonitorSpecialsWhenIncomplete
    $monitorSpecialsWhenComplete     = [bool]$Behavior.monitorSpecialsWhenComplete

    # Batch episode updates to reduce API calls
    $episodesToUpdate = [System.Collections.Generic.List[object]]::new()

    foreach ($ep in $Episodes) {
        $isSpecial        = ($ep.seasonNumber -eq 0)
        $desiredMonitored = $ep.monitored

        if (-not $isSpecial) {
            if ($monitorNonSpecials) {
                $desiredMonitored = $true
            }
        }
        else {
            if ($IsInCompleteProfile) {
                if ($monitorSpecialsWhenComplete) {
                    $desiredMonitored = $true
                }
            }
            else {
                if ($unmonitorSpecialsWhenIncomplete) {
                    $desiredMonitored = $false
                }
            }
        }

        if ($ep.monitored -ne $desiredMonitored) {
            $season = $ep.seasonNumber
            $epNum  = $ep.episodeNumber

            $ep.monitored = $desiredMonitored
            $changes++

            $stateText    = if ($desiredMonitored) { 'monitored=True' } else { 'monitored=False' }
            $profileLabel = if ($IsInCompleteProfile) { 'Complete' } else { 'Incomplete' }
            $reason       = if ($isSpecial) {
                                if ($IsInCompleteProfile) { 'Complete: monitor specials now' }
                                else { 'Incomplete: ignore specials for now' }
                            }
                            else {
                                'Ensure non-specials stay monitored'
                            }

            Write-Log 'DEBUG' ("Setting S{0:D2}E{1:D2} {2} for series '{3}' (id={4}) [{5}: {6}]" -f $season, $epNum, $stateText, $Series.title, $Series.id, $profileLabel, $reason)

            $episodesToUpdate.Add($ep)
        }
    }

    # Batch update episodes to reduce API calls
    if ($episodesToUpdate.Count -gt 0) {
        Write-Log 'DEBUG' ("Batch updating {0} episodes for series '{1}' (id={2})" -f $episodesToUpdate.Count, $Series.title, $Series.id)
        
        foreach ($ep in $episodesToUpdate) {
            Update-SonarrEpisode -Episode $ep -Context ("batch update monitored for series {0}" -f $Series.id)
        }
    }

    $Global:CompleteARR_Summary.EpisodeMonitorChanges += $changes
}

function Join-SonarrRootAndLeaf {
    param(
        [string]$RootFolder,
        [string]$Leaf
    )

    $root = $RootFolder.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($Leaf)) {
        return $root
    }
    return "$root/$Leaf"
}

function Update-SeriesPathInRoot {
    <#
        🗂️  ENSURES SERIES ARE IN THE CORRECT FOLDERS
        
        This function checks if a series is in the right root folder and fixes it if not.
        
        Why this happens:
        - Sometimes series get moved manually or by other processes
        - File system changes might affect folder paths
        - This ensures consistency across your library
        
        What it does:
        - Checks if the series' current path matches the expected root
        - If not, moves it to the correct location
        - Keeps the same folder name (just changes the parent path)
        
        💡 TIP: This is a safety feature that keeps your library organized!
    #>
    param(
        [pscustomobject]$Series,
        [string]$ExpectedRoot
    )

    $expectedRoot = $ExpectedRoot.TrimEnd('/')
    $currentPath  = $Series.path

    if (-not $currentPath) {
        return $false
    }

    if ($currentPath.StartsWith("$expectedRoot/")) {
        return $false
    }

    $leaf    = Split-Path -Path $currentPath -Leaf
    $newPath = Join-SonarrRootAndLeaf -RootFolder $expectedRoot -Leaf $leaf

    Write-Log 'FILE' ("Target root folder for media set path correction is: {0}" -f $expectedRoot)
    Write-Log 'SUCCESS' ("Series '{0}' (id={1}) - ROOT CORRECTION: path '{2}' -> '{3}'" -f $Series.title, $Series.id, $currentPath, $newPath)

    $Series.path = $newPath
    $Global:CompleteARR_Summary.RootCorrections++
    return $true
}

# ------------------------------------------------------------
# MAIN SET PROCESSING
# ------------------------------------------------------------

function Invoke-MediaSetProcessing {
    <#
        🔄 PROCESSES A SINGLE MEDIA CATEGORY (Adult, Family, Anime, etc.)
        
        This function handles all the logic for one media type:
        - Processes "Incomplete" series: Promotes them to "Complete" when ready
        - Processes "Complete" series: Demotes them if missing episodes appear
        - Manages episode monitoring and folder locations
        
        📊 PERFORMANCE NOTE:
        This function makes API calls to check episodes for each series.
        For large libraries, this can take time. The 'throttleMs' setting
        in your config helps prevent overwhelming your Sonarr server with too many API calls at once.
    #>
    param(
        [pscustomobject]$ConfigSet,          # A single item from config.sets[]
        [object[]]$AllSeries,
        [hashtable]$QualityProfilesByName,
        [datetime]$NowUtc,
        [hashtable]$SeriesByProfileId = $null
    )

    $mediaType       = $ConfigSet.'Media Type'
    $incProfileName  = $ConfigSet.'Incomplete Profile Name'
    $incRootFolder   = $ConfigSet.'Incomplete Root Folder'
    $compProfileName = $ConfigSet.'Complete Profile Name'
    $compRootFolder  = $ConfigSet.'Complete Root Folder'

    Write-Log 'INFO' '------------------------------------------------------------'
    Write-Log 'INFO' ("Processing media set: {0}" -f $mediaType)

    if (-not $QualityProfilesByName.ContainsKey($incProfileName)) {
        Write-Log 'ERROR' ("  Incomplete profile '{0}' not found in Sonarr. Skipping this media set." -f $incProfileName)
        return
    }
    if (-not $QualityProfilesByName.ContainsKey($compProfileName)) {
        Write-Log 'ERROR' ("  Complete profile '{0}' not found in Sonarr. Skipping this media set." -f $compProfileName)
        return
    }

    $incProfileId  = $QualityProfilesByName[$incProfileName]
    $compProfileId = $QualityProfilesByName[$compProfileName]

    # Use pre-grouped series data for faster lookups (optimization)
    if ($SeriesByProfileId -and $SeriesByProfileId.ContainsKey($incProfileId)) {
        $incompleteSeries = $SeriesByProfileId[$incProfileId]
    } else {
        $incompleteSeries = $AllSeries | Where-Object { $_.qualityProfileId -eq $incProfileId }
    }
    
    if ($SeriesByProfileId -and $SeriesByProfileId.ContainsKey($compProfileId)) {
        $completeSeries = $SeriesByProfileId[$compProfileId]
    } else {
        $completeSeries = $AllSeries | Where-Object { $_.qualityProfileId -eq $compProfileId }
    }

    Write-Log 'INFO' ("  Found {0} series using Incomplete profile '{1}'." -f $incompleteSeries.Count, $incProfileName)
    Write-Log 'INFO' ("  Found {0} series using Complete profile '{1}'."   -f $completeSeries.Count,   $compProfileName)

    # -------------------------
    # Process INCOMPLETE series
    # -------------------------
    $incompleteCount = $incompleteSeries.Count
    $incompleteIndex = 0
    
    Write-Log 'INFO' ("Processing {0} INCOMPLETE series for media set '{1}'..." -f $incompleteCount, $mediaType)
    
    foreach ($series in $incompleteSeries) {
        $incompleteIndex++
        $Global:CompleteARR_Summary.SeriesChecked++
        $Global:CompleteARR_Summary.IncompleteSeriesSeen++

        # New improved logging format with highlighted series title
        Write-Log 'INFO' ("[{0}/{1}]" -f $incompleteIndex, $incompleteCount)
        Write-Log 'INFO' ("Processing series '{0}' (id={1})" -f $series.title, $series.id) -HighlightText $series.title
        Write-Log 'INFO' "Current Status: INCOMPLETE"

        $episodes = Get-SonarrEpisodesForSeries -SeriesId $series.id

        # Ensure monitoring state (non-specials monitored, specials unmonitored in incomplete)
        Update-MonitoringStateForSeries -Series $series -Episodes $episodes -IsInCompleteProfile:$false -Behavior $Global:CompleteARR_Behavior

        $stats        = Get-SeriesEpisodeStats -Series $series -Episodes $episodes -NowUtc $NowUtc -Behavior $Global:CompleteARR_Behavior
        $missingCount = $stats.MissingAiredNonSpecialEpisodes.Count
        $missingPastGraceCount = $stats.MissingPastGraceNonSpecialEpisodes.Count
        
        # Detailed episode breakdown
        Write-Log 'INFO' ("Found {0} total episodes" -f $episodes.Count)
        Write-Log 'INFO' ("Found {0} non-special episodes: out of {1} total aired, {2} missing past grace days" -f $stats.NonSpecialEpisodes.Count, $stats.AiredNonSpecialEpisodes.Count, $missingPastGraceCount)
        Write-Log 'INFO' ("Found {0} specials: {1} monitored, {2} with files" -f $stats.TotalSpecials, $stats.MonitoredSpecials, $stats.CompleteSpecials)

        if ($missingCount -eq 0) {
            # PROMOTE to COMPLETE
            $oldProfileId           = $series.qualityProfileId
            $series.qualityProfileId = $compProfileId

            $oldPath = $series.path
            $leaf    = Split-Path -Path $oldPath -Leaf
            $newPath = Join-SonarrRootAndLeaf -RootFolder $compRootFolder -Leaf $leaf

            Write-Log 'PROMOTION' ("🚀 Series '{0}' (id={1}) - PROMOTE TO COMPLETE: profileId {2} -> {3}; path '{4}' -> '{5}';" -f $series.title, $series.id, $oldProfileId, $compProfileId, $oldPath, $newPath)
            Write-Log 'FILE'    ("Target root folder for media set '{0}' is: {1}" -f $mediaType, $compRootFolder)

            $series.path = $newPath
            $Global:CompleteARR_Summary.Promotions++

            Update-SonarrSeries -Series $series -Context ("promote series {0}" -f $series.id)

            # After promotion, ensure specials are monitored
            $episodesAfter = Get-SonarrEpisodesForSeries -SeriesId $series.id
            Update-MonitoringStateForSeries -Series $series -Episodes $episodesAfter -IsInCompleteProfile:$true -Behavior $Global:CompleteARR_Behavior

            if ($Global:CompleteARR_Behavior.postMoveWaitSeconds -gt 0 -and -not $Global:CompleteARR_Behavior.dryRun) {
                Start-Sleep -Seconds $Global:CompleteARR_Behavior.postMoveWaitSeconds
            }
        }
        else {
            Write-Log 'INFO' ("Series '{0}' remains INCOMPLETE (missing {1} aired non-special episodes)." -f $series.title, $missingCount)

            # Even if we don't promote, ensure path is correct for incomplete root
            $changed = Update-SeriesPathInRoot -Series $series -ExpectedRoot $incRootFolder
            if ($changed) {
                Update-SonarrSeries -Series $series -Context ("root correction for incomplete series {0}" -f $series.id)
            }
        }
    }

    # -------------------------
    # Process COMPLETE series
    # -------------------------
    $completeCount = $completeSeries.Count
    $completeIndex = 0
    
    Write-Log 'INFO' ("Processing {0} COMPLETE series for media set '{1}'..." -f $completeCount, $mediaType)
    
    foreach ($series in $completeSeries) {
        $completeIndex++
        $Global:CompleteARR_Summary.SeriesChecked++
        $Global:CompleteARR_Summary.CompleteSeriesSeen++

        # New improved logging format with highlighted series title
        Write-Log 'INFO' ("[{0}/{1}]" -f $completeIndex, $completeCount)
        Write-Log 'INFO' ("Processing series '{0}' (id={1})" -f $series.title, $series.id) -HighlightText $series.title
        Write-Log 'INFO' "Current Status: COMPLETE"

        $episodes = Get-SonarrEpisodesForSeries -SeriesId $series.id

        # Ensure monitoring state (non-specials monitored, specials monitored in complete)
        Update-MonitoringStateForSeries -Series $series -Episodes $episodes -IsInCompleteProfile:$true -Behavior $Global:CompleteARR_Behavior

        $stats           = Get-SeriesEpisodeStats -Series $series -Episodes $episodes -NowUtc $NowUtc -Behavior $Global:CompleteARR_Behavior
        $missingPastGrace= $stats.MissingPastGraceNonSpecialEpisodes.Count
        
        # Detailed episode breakdown
        Write-Log 'INFO' ("Found {0} total episodes" -f $episodes.Count)
        Write-Log 'INFO' ("Found {0} non-special episodes: out of {1} total aired, {2} missing past grace days" -f $stats.NonSpecialEpisodes.Count, $stats.AiredNonSpecialEpisodes.Count, $missingPastGrace)
        Write-Log 'INFO' ("Found {0} specials: {1} monitored, {2} with files" -f $stats.TotalSpecials, $stats.MonitoredSpecials, $stats.CompleteSpecials)

        if ($missingPastGrace -gt 0) {
            # DEMOTE to INCOMPLETE
            $oldProfileId           = $series.qualityProfileId
            $series.qualityProfileId = $incProfileId

            $oldPath = $series.path
            $leaf    = Split-Path -Path $oldPath -Leaf
            $newPath = Join-SonarrRootAndLeaf -RootFolder $incRootFolder -Leaf $leaf

            Write-Log 'WARNING' ("Series '{0}' has {1} AIRED non-special episodes older than graceDays missing files. DEMOTING to INCOMPLETE." -f $series.title, $missingPastGrace)
            Write-Log 'DEMOTION' ("📉 Series '{0}' (id={1}) - DEMOTE TO INCOMPLETE: profileId {2} -> {3}; path '{4}' -> '{5}';" -f $series.title, $series.id, $oldProfileId, $incProfileId, $oldPath, $newPath)
            Write-Log 'FILE'    ("Target root folder for media set '{0}' is: {1}" -f $mediaType, $incRootFolder)

            $series.path = $newPath
            $Global:CompleteARR_Summary.Demotions++

            Update-SonarrSeries -Series $series -Context ("demote series {0}" -f $series.id)

            # After demotion, ensure specials are unmonitored again
            $episodesAfter = Get-SonarrEpisodesForSeries -SeriesId $series.id
            Update-MonitoringStateForSeries -Series $series -Episodes $episodesAfter -IsInCompleteProfile:$false -Behavior $Global:CompleteARR_Behavior

            if ($Global:CompleteARR_Behavior.postMoveWaitSeconds -gt 0 -and -not $Global:CompleteARR_Behavior.dryRun) {
                Start-Sleep -Seconds $Global:CompleteARR_Behavior.postMoveWaitSeconds
            }
        }
        else {
            Write-Log 'INFO' ("Series '{0}' remains COMPLETE (no qualifying missing episodes past graceDays)." -f $series.title)

            # Even if we don't demote, ensure path is correct for complete root
            $changed = Update-SeriesPathInRoot -Series $series -ExpectedRoot $compRootFolder
            if ($changed) {
                Update-SonarrSeries -Series $series -Context ("root correction for complete series {0}" -f $series.id)
            }
        }
    }
}

# ------------------------------------------------------------
# ENTRY POINT
# ------------------------------------------------------------

try {
    Write-Log 'INFO' ("CompleteARR SONARR engine starting. ScriptRoot = {0}" -f $Script:CompleteARR_ScriptRoot)

    $cfgAndPath = Import-CompleteARRConfig -Path $ConfigPath
    $config     = $cfgAndPath[0]
    $configPath = $cfgAndPath[1]
    $Global:CompleteARR_Config = $config

    Initialize-LoggingFromConfig -Config $config -ConfigPath $configPath

    $Global:CompleteARR_Behavior = $config.behavior

    $dryRun    = [bool]$Global:CompleteARR_Behavior.dryRun
    $graceDays = [int]$Global:CompleteARR_Behavior.graceDays
    $preflight = [int]($Global:CompleteARR_Behavior.preflightSeconds ?? 0)

    Write-Log 'DEBUG' ("Behavior: dryRun={0}, graceDays={1}, throttleMs={2}" -f $dryRun, $graceDays, $Global:CompleteARR_ThrottleMs)

    if ($preflight -gt 0) {
        Write-Log 'INFO' ("Preflight delay {0} seconds before starting work..." -f $preflight)
        Start-Sleep -Seconds $preflight
    }

    Write-Log 'INFO' "----- BEGIN CompleteARR SONARR run -----"

    # Connect to Sonarr
    $status = Get-SonarrSystemStatus
    Write-Log 'SUCCESS' ("Connected to Sonarr version {0} at {1}" -f $status.version, $config.sonarr.url)

    # Load quality profiles
    $profiles = Get-SonarrQualityProfiles
    Write-Log 'DEBUG' ("Loaded {0} quality profiles from Sonarr." -f $profiles.Count)

    $profilesByName = @{}
    foreach ($p in $profiles) {
        $profilesByName[$p.name] = $p.id
    }

    # Load all series
    $allSeries = Get-SonarrSeries
    Write-Log 'DEBUG' ("Loaded {0} series from Sonarr." -f $allSeries.Count)

    # Pre-group series by profile ID for faster lookups (optimization)
    $seriesByProfileId = @{}
    foreach ($series in $allSeries) {
        $profileId = $series.qualityProfileId
        if (-not $seriesByProfileId.ContainsKey($profileId)) {
            $seriesByProfileId[$profileId] = [System.Collections.Generic.List[object]]::new()
        }
        $seriesByProfileId[$profileId].Add($series)
    }

    $nowUtc = (Get-Date).ToUniversalTime()

    # Process media sets sequentially to ensure proper logging and summary updates
    Write-Log 'INFO' ("Processing {0} media sets sequentially..." -f $config.sets.Count)
    foreach ($set in $config.sets) {
        Invoke-MediaSetProcessing -ConfigSet $set -AllSeries $allSeries -QualityProfilesByName $profilesByName -NowUtc $nowUtc -SeriesByProfileId $seriesByProfileId
    }

    # SUMMARY
    Write-Log 'INFO' "----- COMPLETEARR SONARR SUMMARY -----"
    Write-Log 'INFO' ("  Series checked           : {0}" -f $Global:CompleteARR_Summary.SeriesChecked)
    Write-Log 'INFO' ("  Incomplete series seen   : {0}" -f $Global:CompleteARR_Summary.IncompleteSeriesSeen)
    Write-Log 'INFO' ("  Complete series seen     : {0}" -f $Global:CompleteARR_Summary.CompleteSeriesSeen)
    Write-Log 'INFO' ("  Promotions (-> complete) : {0}" -f $Global:CompleteARR_Summary.Promotions)
    Write-Log 'INFO' ("  Demotions (-> incomplete): {0}" -f $Global:CompleteARR_Summary.Demotions)
    Write-Log 'INFO' ("  Root corrections         : {0}" -f $Global:CompleteARR_Summary.RootCorrections)
    Write-Log 'INFO' ("  Episode monitor changes  : {0}" -f $Global:CompleteARR_Summary.EpisodeMonitorChanges)
    Write-Log 'INFO' "----- END CompleteARR SONARR run -----"
    
    # Return summary for master summary display
    return $Global:CompleteARR_Summary
}
catch {
    Write-Log 'ERROR' ("FATAL ERROR: {0}" -f $_.Exception.Message)
    Write-Log 'DEBUG' ("Stack trace:`n{0}" -f $_.ScriptStackTrace)
    throw
}
