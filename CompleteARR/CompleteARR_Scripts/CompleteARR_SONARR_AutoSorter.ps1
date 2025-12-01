#Requires -Version 7.0
<#
    CompleteARR_SONARR_AutoSorter.ps1
    ------------------------------------------------------------
    
    🎯 WHAT THIS SCRIPT DOES:
    - Runs BEFORE the main Series Engine
    - Automatically sorts NEW series into the right categories:
        • Family (kid-friendly content)
        • Adult (regular grown-up series)
        • Anime Family (kid-safe anime)
    - Uses content ratings, genres, and tags to make smart decisions
    - Moves series to the correct "Incomplete" folders
    - Creates detailed logs to help you understand its decisions
    🔍 HOW IT WORKS:
    1. Looks for series in your "source" quality profiles (like Kometa or User Requests)
    2. Examines each series' rating, genres, and tags
    4. Moves the series to the right "Incomplete" folder
    5. The main Series Engine will later promote complete series

        💡 TIP: For best results, make sure your request bot sets the correct
                series type when adding anime series!
    #>

    #>

[CmdletBinding()]
param(
    [string]$ConfigPathOverride,
    [string]$ScriptRootOverride
)

$ErrorActionPreference = 'Stop'

$Script:CompleteARR_ScriptRoot  = $null
$Global:CompleteARR_Config      = $null
$Global:CompleteARR_LogFilePath = $null
$Global:CompleteARR_LogMinLevel = 'Info'
$Global:CompleteARR_LogToFile   = $true
$Global:CompleteARR_LogToConsole= $true
$Global:CompleteARR_UseColors   = $true
$Global:CompleteARR_ThrottleMs  = 200
$Global:CompleteARR_Behavior    = $null

$Global:CompleteARR_Summary = [PSCustomObject]@{
    SeriesSeenTotal          = 0
    SeriesInSourceProfiles   = 0
    SeriesMoved              = 0
    SeriesAlreadyCorrect     = 0
    SeriesSkippedNoTarget    = 0
}

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
                default   { $Host.UI.RawUI.ForegroundColor = $origColor }
            }
            Write-Host $line
            $Host.UI.RawUI.ForegroundColor = $origColor
        }
    }
    else {
        Write-Host $line
    }

    # File logging still respects logToFile + path
    if ($Global:CompleteARR_LogToFile -and $Global:CompleteARR_LogFilePath) {
        Add-Content -Path $Global:CompleteARR_LogFilePath -Value $line
    }
}


function Import-CompleteARRConfig {
    param([string]$Path)

    if (-not $Path) {
        # Default to CompleteARR_SONARR_Settings.yml in the CompleteARR_Settings folder.
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
    $scriptName = "CompleteARR_SONARR_AutoSorter"
    $logFile = Join-Path $logsRoot ("{0}_{1}{2}" -f $scriptName, $timestamp, $ext)

    $Global:CompleteARR_LogFilePath  = $logFile
    $Global:CompleteARR_LogMinLevel  = if ($logging.minLevel) { $logging.minLevel } else { 'Info' }
    $Global:CompleteARR_LogToFile    = [bool]$logging.logToFile
    $Global:CompleteARR_LogToConsole = [bool]$logging.logToConsole
    $Global:CompleteARR_UseColors    = if ($null -ne $logging.useColors) { [bool]$logging.useColors } else { $true }
    $Global:CompleteARR_ThrottleMs   = if ($logging.throttleMs) { [int]$logging.throttleMs } else { 200 }

    $Global:CompleteARR_Config   = $Config
    $Global:CompleteARR_Behavior = $Config.behavior

    Write-Log 'FILE' ("Using configuration file: {0}" -f $ConfigPath)
    Write-Log 'FILE' ("AutoSorter log file: {0}" -f $logFile)
}

# ... [classification helpers, Invoke-SonarrApi, etc. from your current working AutoSorter] ...

# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

try {
    # 1) Resolve CompleteARR root folder.
    $thisScriptPath = $MyInvocation.MyCommand.Path
    $thisScriptDir  = Split-Path -Path $thisScriptPath -Parent

    if ($ScriptRootOverride) {
        $resolvedRoot = (Resolve-Path -LiteralPath $ScriptRootOverride).Path
    }
    else {
        # AutoSorter lives in CompleteARR_Scripts; root is its parent.
        $resolvedRoot = Split-Path -Path $thisScriptDir -Parent
    }

    $Script:CompleteARR_ScriptRoot = $resolvedRoot

    # 2) Load config and initialize logging.
    $configPathToUse = $null
    if ($ConfigPathOverride) {
        $configPathToUse = $ConfigPathOverride
    }

    $cfgPair = Import-CompleteARRConfig -Path $configPathToUse
    $cfg     = $cfgPair[0]
    $cfgPath = $cfgPair[1]

    Initialize-LoggingFromConfig -Config $cfg -ConfigPath $cfgPath

    Write-Log 'INFO' ("CompleteARR SONARR AutoSorter starting. ScriptRoot = {0}" -f $Script:CompleteARR_ScriptRoot)

    # ... [rest of your existing AutoSorter main loop unchanged] ...
}
catch {
    Write-Log 'ERROR' ("FATAL: AutoSorter failed: {0}" -f $_.Exception.Message)
    throw
}



# ------------------------------------------------------------
# SONARR API HELPER
# ------------------------------------------------------------

function Invoke-SonarrApi {
    <#
        Generic Sonarr API wrapper.

        - $Path is a relative path like "system/status" or "series"
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

        [hashtable]$Query,
        [object]$Body,
        [string]$ErrorContext
    )

    if (-not $Global:CompleteARR_Config -or -not $Global:CompleteARR_Config.sonarr) {
        throw "CompleteARR_Config.sonarr is not initialized."
    }

    $baseUrl = $Global:CompleteARR_Config.sonarr.url
    if (-not $baseUrl) {
        throw "sonarr.url is not set in CompleteARR_Settings.yml."
    }

    $baseUrl = $baseUrl.TrimEnd('/')

    # Prepend api/v3/ if the caller only provided a relative path.
    if ($Path -notmatch '^api/v[0-9]+/') {
        $Path = "api/v3/$Path"
    }

    $uri = "$baseUrl/$Path"

    if ($Query) {
        $qsPairs = @()
        foreach ($key in $Query.Keys) {
            $val = $Query[$key]
            if ($null -ne $val) {
                $qsPairs += ("{0}={1}" -f [System.Web.HttpUtility]::UrlEncode($key.ToString()),
                                           [System.Web.HttpUtility]::UrlEncode($val.ToString()))
            }
        }
        if ($qsPairs.Count -gt 0) {
            $qs  = $qsPairs -join '&'
            $sep = $uri.Contains('?') ? '&' : '?'
            $uri = "$uri$sep$qs"
        }
    }

    Write-Log 'DEBUG' ("SONARR {0} {1}" -f $Method, $uri)

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

# ------------------------------------------------------------
# CLASSIFICATION HELPERS (Anime, AgeGroup, Tags)
# ------------------------------------------------------------

function Get-RatingMapFromConfig {
    param([pscustomobject]$AutoCfg)

    $map = @()

    if ($AutoCfg.ratingToAgeGroup) {
        foreach ($entry in $AutoCfg.ratingToAgeGroup.GetEnumerator()) {
            $key = $entry.Key
            $val = $entry.Value

            if ($null -eq $key -or -not $val) { continue }

            $token    = $key.ToString().ToLowerInvariant()
            $ageGroup = $val.ToString().ToLowerInvariant()

            $map += [PSCustomObject]@{
                RatingToken = $token
                AgeGroup    = $ageGroup
            }
        }
    }

    return $map
}

function Get-SeriesAgeGroup {
    <#
        🎭 DETERMINES AGE GROUP FOR A SHOW
        
        This function looks at a show's information and decides which age group it belongs to:
        
        How it decides:
        1. First looks at the content rating (TV-Y, TV-PG, TV-MA, etc.)
        2. If no rating, looks at genres and networks for family-friendly hints
        3. Uses "strict family mode" to be extra careful about what goes to family libraries
        
        Age Groups:
        - under10 : Preschool / early kids (TV-Y, TV-Y7, G, PG)
        - preteen : 9-12 year olds (TV-PG, 12, 12A)  
        - teen    : 13-16 year olds (TV-14, PG-13, 15, 16)
        - adult   : 17+ (TV-MA, R, 18)
        - xplicit : Very explicit content
        - unrated : Unknown rating
    #>
    param(
        [pscustomobject]$Series,
        [object[]]$RatingMap,
        [string[]]$FamilyGenres,
        [string[]]$FamilyNetworks,
        [string[]]$NonFamilyGenres
    )

    # Use modern PowerShell property access with null-conditional operators
    $cert = $Series.certification ?? $Series.ratings?.value ?? $null
    
    $norm = if ($cert) { 
        $cert.ToString().ToLowerInvariant() 
    } else { 
        $null 
    }

    # Use hashtable for faster rating lookups instead of linear search
    $ratingLookup = @{}
    if ($RatingMap) {
        foreach ($entry in $RatingMap) {
            if ($entry.RatingToken) {
                $ratingLookup[$entry.RatingToken] = $entry.AgeGroup
            }
        }
    }

    $ageGroup = $null
    if ($norm -and $ratingLookup.Count -gt 0) {
        foreach ($token in $ratingLookup.Keys) {
            if ($norm -like "*$token*") {
                $ageGroup = $ratingLookup[$token]
                break
            }
        }
    }

    # Use HashSet for faster contains operations
    $genresLower = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($Series.genres) {
        foreach ($g in $Series.genres) {
            if ($g) { [void]$genresLower.Add($g.ToString()) }
        }
    }

    $networkLower = if ($Series.network) { $Series.network.ToString().ToLowerInvariant() } else { $null }
    
    # Convert arrays to HashSets for faster lookups
    $familyGenresSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($FamilyGenres) { 
        foreach ($g in $FamilyGenres) { 
            if ($g) { [void]$familyGenresSet.Add($g) } 
        } 
    }
    
    $familyNetworksSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($FamilyNetworks) { 
        foreach ($n in $FamilyNetworks) { 
            if ($n) { [void]$familyNetworksSet.Add($n) } 
        } 
    }
    
    $nonFamilyGenresSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($NonFamilyGenres) { 
        foreach ($g in $NonFamilyGenres) { 
            if ($g) { [void]$nonFamilyGenresSet.Add($g) } 
        } 
    }

    # Use HashSet intersections for faster checks
    $isFamilyHint = $false
    $hasNonFamilyGenre = $false

    if ($genresLower.Count -gt 0 -and $familyGenresSet.Count -gt 0) {
        $isFamilyHint = $genresLower.Overlaps($familyGenresSet)
    }

    if (-not $isFamilyHint -and $networkLower -and $familyNetworksSet.Count -gt 0) {
        foreach ($n in $familyNetworksSet) {
            if ($networkLower -like "*$n*") {
                $isFamilyHint = $true
                break
            }
        }
    }

    if ($genresLower.Count -gt 0 -and $nonFamilyGenresSet.Count -gt 0) {
        $hasNonFamilyGenre = $genresLower.Overlaps($nonFamilyGenresSet)
    }

    if (-not $ageGroup) {
        $ageGroup = if ($isFamilyHint) { 'under10' } else { 'unrated' }
    }

    return [PSCustomObject]@{
        AgeGroup          = $ageGroup
        IsFamilyHint      = $isFamilyHint
        HadRating         = [bool]$cert
        HasNonFamilyGenre = $hasNonFamilyGenre
    }
}

function Test-IsAnimeSeries {
    <#
        🎌 DETECTS IF A SHOW IS ANIME
        
        This function checks multiple clues to determine if a show is anime:
        
        How it detects anime:
        1. First checks if the series type is explicitly set to "anime" in Sonarr
        2. Looks for "anime" in the genres list
        3. Checks if the network name contains "anime"
        4. As a last resort, looks for "anime" in the show title
        
        💡 TIP: For best results, make sure your request bot sets the correct 
                series type when adding anime shows!
    #>
    param(
        [pscustomobject]$Series,
        [bool]$ForceAnime
    )

    if ($ForceAnime) { return $true }

    # 1) Check if Sonarr has this marked as an anime series
    $seriesTypeLower = if ($Series.seriesType) { $Series.seriesType.ToString().ToLowerInvariant() } else { $null }
    if ($seriesTypeLower -eq 'anime') {
        return $true
    }

    # 2) Look for "anime" in the genres using HashSet for faster lookup
    if ($Series.genres) {
        $genresLower = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($g in $Series.genres) {
            if ($g) { [void]$genresLower.Add($g.ToString()) }
        }
        if ($genresLower.Contains('anime')) {
            return $true
        }
    }

    # 3) Check if the network is anime-related
    $networkLower = if ($Series.network) { $Series.network.ToString().ToLowerInvariant() } else { $null }
    if ($networkLower -and $networkLower -like '*anime*') {
        return $true
    }

    # 4) Last resort: check the title for "anime"
    $titleLower = if ($Series.title) { $Series.title.ToString().ToLowerInvariant() } else { $null }
    if ($titleLower -and $titleLower -like '*anime*') {
        return $true
    }

    return $false
}

function Get-TagOverrideDecision {
    <#
        🏷️ HANDLES MANUAL OVERRIDES WITH TAGS
        
        This function lets you manually control where shows go using Sonarr tags.
        
        How it works:
        - Checks if a show has any of the override tags
        - Tags are checked in priority order (first match wins)
        - You can force a show to a specific category or just mark it as anime
        
        Available Override Tags:
        - libraryoverride-family    -> Forces to Family category
        - libraryoverride-adult     -> Forces to Adult category  
        - libraryoverride-xplicit   -> Forces to Adult category (explicit)
        - libraryoverride-anime     -> Forces anime detection (not a specific category)
        
        💡 TIP: Create these tags in Sonarr and assign them to shows that 
                don't get sorted correctly automatically!
    #>
    param(
        [pscustomobject]$Series,
        [hashtable]$TagsById,
        [string[]]$OverrideOrder
    )

    if (-not $Series.tags -or $Series.tags.Count -eq 0) {
        return $null
    }

    $labels = @()
    foreach ($tid in $Series.tags) {
        if ($TagsById.ContainsKey($tid)) {
            $labels += $TagsById[$tid]
        }
    }

    if ($labels.Count -eq 0) { return $null }

    $labelsLower = $labels | ForEach-Object {
        if ($_){ $_.ToString().ToLowerInvariant() }
    }

    foreach ($tagName in $OverrideOrder) {
        if (-not $tagName) { continue }
        $tagLower = $tagName.ToLowerInvariant()

        if ($labelsLower -contains $tagLower) {
            switch ($tagLower) {
                'libraryoverride-family' {
                    return [PSCustomObject]@{
                        Bucket     = 'Family'
                        ForceAnime = $false
                        Tag        = $tagName
                    }
                }
                'libraryoverride-adult' {
                    return [PSCustomObject]@{
                        Bucket     = 'Adult'
                        ForceAnime = $false
                        Tag        = $tagName
                    }
                }
                'libraryoverride-xplicit' {
                    return [PSCustomObject]@{
                        Bucket     = 'Adult'
                        ForceAnime = $false
                        Tag        = $tagName
                    }
                }
                'libraryoverride-anime' {
                    # Not a direct target set; just force anime routing.
                    return [PSCustomObject]@{
                        Bucket     = $null
                        ForceAnime = $true
                        Tag        = $tagName
                    }
                }
                default {
                    # Unknown override tag; just log and do nothing special.
                    return [PSCustomObject]@{
                        Bucket     = $null
                        ForceAnime = $false
                        Tag        = $tagName
                    }
                }
            }
        }
    }

    return $null
}

# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

try {
    # 1) Resolve CompleteARR root folder.
    $thisScriptPath = $MyInvocation.MyCommand.Path
    $thisScriptDir  = Split-Path -Path $thisScriptPath -Parent

    if ($ScriptRootOverride) {
        $resolvedRoot = (Resolve-Path -LiteralPath $ScriptRootOverride).Path
    }
    else {
        # AutoSorter lives in CompleteARR_Parts; root is its parent.
        $resolvedRoot = Split-Path -Path $thisScriptDir -Parent
    }

    $Script:CompleteARR_ScriptRoot = $resolvedRoot

    # 2) Load config and initialize logging.
    $configPathToUse = $null
    if ($ConfigPathOverride) {
        $configPathToUse = $ConfigPathOverride
    }

    $cfgPair = Import-CompleteARRConfig -Path $configPathToUse
    $cfg     = $cfgPair[0]
    $cfgPath = $cfgPair[1]

    Initialize-LoggingFromConfig -Config $cfg -ConfigPath $cfgPath

    Write-Log 'INFO' ("CompleteARR AutoSorter starting. ScriptRoot = {0}" -f $Script:CompleteARR_ScriptRoot)

    # 3) Check autoSorter section + enabled flag.
    if (-not $cfg.autoSorter) {
        Write-Log 'WARNING' "autoSorter section is missing in configuration; nothing to do."
        return
    }

    $autoCfg = $cfg.autoSorter

    if (-not $autoCfg.enabled) {
        Write-Log 'INFO' "autoSorter.enabled is false; AutoSorter will exit without changes."
        return
    }

    # 4) Optional preflight delay (reuse behavior.preflightSeconds if present).
    if ($cfg.behavior -and $cfg.behavior.preflightSeconds) {
        $preflight = [int]$cfg.behavior.preflightSeconds
        if ($preflight -gt 0) {
            Write-Log 'INFO' ("Preflight delay: {0} seconds before AutoSorter work begins." -f $preflight)
            Start-Sleep -Seconds $preflight
        }
    }

    # 5) Build classification helpers from config.
    $sourceProfiles     = $autoCfg.sourceProfiles
    $strictFamilyMode   = [bool]$autoCfg.strictFamilyMode
    $familyGenres       = @()
    $familyNetworks     = @()
    $nonFamilyGenres    = @()

    if ($autoCfg.familyGenres)    { $familyGenres    = [string[]]$autoCfg.familyGenres }
    if ($autoCfg.familyNetworks)  { $familyNetworks  = [string[]]$autoCfg.familyNetworks }
    if ($autoCfg.nonFamilyGenres) { $nonFamilyGenres = [string[]]$autoCfg.nonFamilyGenres }

    $sortTargets        = $autoCfg.sortTargets
    $tagsOverrideOrder  = $autoCfg.tagsOverrideOrder
    $routing            = $autoCfg.routing
    $routeNonAnime      = $routing.routeNonAnime
    $routeAnime         = $routing.routeAnime
    $ratingMap          = Get-RatingMapFromConfig -AutoCfg $autoCfg

    if (-not $sourceProfiles -or $sourceProfiles.Count -eq 0) {
        Write-Log 'WARNING' "autoSorter.sourceProfiles is empty; AutoSorter has nothing to process."
        return
    }

    if (-not $sortTargets -or $sortTargets.Keys.Count -eq 0) {
        Write-Log 'WARNING' "autoSorter.sortTargets is empty; AutoSorter has no target libraries defined."
        return
    }

    if (-not $routeNonAnime -and -not $routeAnime) {
        Write-Log 'WARNING' "autoSorter.routing is missing or incomplete; AutoSorter cannot classify series."
        return
    }

    $sourceProfileNamesLower = $sourceProfiles | ForEach-Object {
        $_.ToString().ToLowerInvariant()
    }

    Write-Log 'INFO' ("AutoSorter source profiles: {0}" -f ($sourceProfiles -join ', '))
    Write-Log 'INFO' ("AutoSorter sort targets: {0}" -f ($sortTargets.Keys -join ', '))

    # 6) Connect to Sonarr and pull data.
    $status = Invoke-SonarrApi -Method 'GET' -Path 'system/status' -ErrorContext 'GET system/status'
    Write-Log 'SUCCESS' ("Connected to Sonarr {0} at {1}" -f $status.version, $cfg.sonarr.url)

    $allProfiles = Invoke-SonarrApi -Method 'GET' -Path 'qualityprofile' -ErrorContext 'GET qualityprofile'
    $allTags     = Invoke-SonarrApi -Method 'GET' -Path 'tag'             -ErrorContext 'GET tag'
    $allSeries   = Invoke-SonarrApi -Method 'GET' -Path 'series'          -ErrorContext 'GET series'

    # Prepare maps for quick lookups.
    $profilesById   = @{}
    $profilesByName = @{}

    foreach ($p in $allProfiles) {
        if ($null -eq $p.id) { continue }
        $profilesById[$p.id] = $p
        if ($p.name) {
            $profilesByName[$p.name.ToString().ToLowerInvariant()] = $p
        }
    }

    $tagsById = @{}
    foreach ($t in $allTags) {
        if ($null -eq $t.id) { continue }
        $label = $t.label
        if (-not $label) { continue }
        $tagsById[$t.id] = $label
    }

    # 7) Initialize summary counters
    $Global:CompleteARR_Summary.SeriesSeenTotal        = 0
    $Global:CompleteARR_Summary.SeriesInSourceProfiles = 0
    $Global:CompleteARR_Summary.SeriesMoved            = 0
    $Global:CompleteARR_Summary.SeriesAlreadyCorrect   = 0
    $Global:CompleteARR_Summary.SeriesSkippedNoTarget  = 0

    # 8) Process each series using parallel processing for better performance
    $seriesToProcess = $allSeries | Where-Object { 
        $profileId = $_.qualityProfileId
        if (-not $profilesById.ContainsKey($profileId)) {
            Write-Log 'DEBUG' ("Series '{0}' (ID={1}) uses unknown qualityProfileId={2}; skipping." -f $_.title, $_.id, $profileId)
            return $false
        }

        $profileObj = $profilesById[$profileId]
        $profileName = $profileObj.name
        if (-not $profileName) {
            Write-Log 'DEBUG' ("Series '{0}' (ID={1}) has a profile with no name; skipping." -f $_.title, $_.id)
            return $false
        }

        $profileNameLower = $profileName.ToLowerInvariant()
        $sourceProfileNamesLower -contains $profileNameLower
    }

    Write-Log 'INFO' ("Found {0} series to process in source profiles." -f $seriesToProcess.Count)

    # Add progress tracking
    $seriesCount = $seriesToProcess.Count
    $seriesIndex = 0
    
    foreach ($series in $seriesToProcess) {
        $seriesIndex++
        $Global:CompleteARR_Summary.SeriesSeenTotal++
        $Global:CompleteARR_Summary.SeriesInSourceProfiles++

        Write-Log 'INFO' ("[{0}/{1}] Processing series '{2}' (ID={3})" -f $seriesIndex, $seriesCount, $series.title, $series.id) -HighlightText $series.title

        # Determine if tag overrides apply.
        $overrideDecision = $null
        if ($tagsOverrideOrder -and $tagsOverrideOrder.Count -gt 0) {
            $overrideDecision = Get-TagOverrideDecision -Series $series -TagsById $tagsById -OverrideOrder $tagsOverrideOrder
        }

        $forceAnime = $false
        $targetSet     = $null

        if ($overrideDecision) {
            if ($overrideDecision.Tag) {
                Write-Log 'DEBUG' ("Series '{0}' (ID={1}) has override tag '{2}'." -f $series.title, $series.id, $overrideDecision.Tag)
            }

            if ($overrideDecision.ForceAnime) {
                $forceAnime = $true
            }

            if ($overrideDecision.Bucket) {
                $targetSet = $overrideDecision.Bucket
            }
        }

        # Determine anime vs non-anime.
        $isAnime = Test-IsAnimeSeries -Series $series -ForceAnime:$forceAnime

        # Determine age group if target set not forced by override.
        $ageInfo  = $null
        $ageGroup = $null

        if (-not $targetSet) {
            $ageInfo  = Get-SeriesAgeGroup -Series $series `
                                           -RatingMap $ratingMap `
                                           -FamilyGenres $familyGenres `
                                           -FamilyNetworks $familyNetworks `
                                           -NonFamilyGenres $nonFamilyGenres
            $ageGroup = $ageInfo.AgeGroup

            # strictFamilyMode: if it looks like kids by rating (under10/preteen)
            # but either has NO family hint OR has a "nonFamily" genre, treat it as teen.
            if ($strictFamilyMode -and ($ageInfo.HasNonFamilyGenre -or -not $ageInfo.IsFamilyHint) -and $ageGroup -in @('under10','preteen')) {
                Write-Log 'DEBUG' ("Strict family mode: '{0}' (ID={1}) bumped from ageGroup '{2}' to 'teen' (no family hint or nonFamily genre)." -f $series.title, $series.id, $ageGroup)
                $ageGroup = 'teen'
            }
        }

        # If still no target set, route using anime/non-anime routing tables.
        if (-not $targetSet) {
            $routeTable = if ($isAnime) { $routeAnime } else { $routeNonAnime }

            if (-not $routeTable) {
                Write-Log 'WARNING' ("Series '{0}' (ID={1}) cannot be routed: routing table for {2} is missing." -f $series.title, $series.id, ($(if ($isAnime) { 'anime' } else { 'non-anime' })))
                $Global:CompleteARR_Summary.SeriesSkippedNoTarget++
                continue
            }

            if (-not $ageGroup) {
                # Should not happen, but be safe.
                $ageGroup = 'unrated'
            }

            if (-not $routeTable.ContainsKey($ageGroup)) {
                Write-Log 'WARNING' ("Series '{0}' (ID={1}) cannot be routed: no mapping for ageGroup '{2}'." -f $series.title, $series.id, $ageGroup)
                $Global:CompleteARR_Summary.SeriesSkippedNoTarget++
                continue
            }

            $targetSet = $routeTable[$ageGroup]
        }

        if (-not $targetSet) {
            Write-Log 'WARNING' ("Series '{0}' (ID={1}) could not be assigned a target set; skipping." -f $series.title, $series.id)
            $Global:CompleteARR_Summary.SeriesSkippedNoTarget++
            continue
        }

        if (-not $sortTargets.ContainsKey($targetSet)) {
            Write-Log 'WARNING' ("Series '{0}' (ID={1}) mapped to target set '{2}' but autoSorter.sortTargets has no entry for it; skipping." -f $series.title, $series.id, $targetSet)
            $Global:CompleteARR_Summary.SeriesSkippedNoTarget++
            continue
        }

        $target     = $sortTargets[$targetSet]
        $targetProf = $target.qualityProfile
        $targetRoot = $target.rootFolder

        if (-not $targetProf -or -not $targetRoot) {
            Write-Log 'WARNING' ("Series '{0}' (ID={1}) target set '{2}' has incomplete sortTargets entry; skipping." -f $series.title, $series.id, $targetSet)
            $Global:CompleteARR_Summary.SeriesSkippedNoTarget++
            continue
        }

        $targetProfLower = $targetProf.ToLowerInvariant()
        if (-not $profilesByName.ContainsKey($targetProfLower)) {
            Write-Log 'ERROR' ("Series '{0}' (ID={1}) target profile '{2}' (target set '{3}') does not exist in Sonarr; skipping." -f $series.title, $series.id, $targetProf, $targetSet)
            $Global:CompleteARR_Summary.SeriesSkippedNoTarget++
            continue
        }

        $targetProfileObj = $profilesByName[$targetProfLower]
        $targetProfileId  = $targetProfileObj.id

        # Determine current root and path.
        $currentPath = $series.path
        if ([string]::IsNullOrWhiteSpace($currentPath)) {
            Write-Log 'WARNING' ("Series '{0}' (ID={1}) has no path; skipping." -f $series.title, $series.id)
            $Global:CompleteARR_Summary.SeriesSkippedNoTarget++
            continue
        }

        # If series already matches target profile and root, nothing to do.
        $currentRoot = $null
        if ($series.PSObject.Properties.Match('rootFolderPath').Count -gt 0) {
            $currentRoot = $series.rootFolderPath
        }

        if ([string]::IsNullOrWhiteSpace($currentRoot)) {
            # Derive root by stripping the last segment from the path.
            $trimmed = $currentPath.TrimEnd('/','\')
            $parent  = Split-Path -Path $trimmed -Parent
            $currentRoot = $parent
        }

        $alreadyProfileMatch = ($series.qualityProfileId -eq $targetProfileId)
        $alreadyRootMatch    = ($currentRoot -eq $targetRoot)

        if ($alreadyProfileMatch -and $alreadyRootMatch) {
            $Global:CompleteARR_Summary.SeriesAlreadyCorrect++
            Write-Log 'INFO' ("Series '{0}' (ID={1}) already in correct target set '{2}' (Profile='{3}', Root='{4}')." -f $series.title, $series.id, $targetSet, $targetProf, $targetRoot)
            continue
        }

        # Build new path under the target root, keeping the existing folder name.
        $folderName = [System.IO.Path]::GetFileName($currentPath.TrimEnd('/','\'))
        if (-not $folderName) {
            # Fallback: use title as folder name.
            $folderName = $series.title
        }

        # Always use forward slashes for Sonarr's Linux paths.
        $newPath = ($targetRoot.TrimEnd('/')) + '/' + $folderName

        Write-Log 'INFO' ("Moving series '{0}' (ID={1}) to target set '{2}'." -f $series.title, $series.id, $targetSet)
        Write-Log 'INFO' ("  Profile: '{0}' -> '{1}'" -f $profileName, $targetProf)
        Write-Log 'INFO' ("  Root   : '{0}' -> '{1}'" -f $currentRoot, $targetRoot)
        Write-Log 'INFO' ("  Path   : '{0}' -> '{1}'" -f $currentPath, $newPath)

        # Apply changes to the series object.
        $series.qualityProfileId = $targetProfileId
        $series.rootFolderPath   = $targetRoot
        $series.path             = $newPath

        # IMPORTANT: AutoSorter is not a dry-run; always move files.
        $query = @{ moveFiles = 'true' }

        $null = Invoke-SonarrApi -Method 'PUT' `
                         -Path ("series/{0}" -f $series.id) `
                         -Query $query `
                         -Body $series `
                         -ErrorContext ("PUT series/{0} (AutoSorter)" -f $series.id)

        $Global:CompleteARR_Summary.SeriesMoved++
        Write-Log 'SUCCESS' ("Series '{0}' (ID={1}) successfully moved to target set '{2}'." -f $series.title, $series.id, $targetSet)

        # Optional small pause after move (reuse behavior.postMoveWaitSeconds if configured).
        if ($cfg.behavior -and $cfg.behavior.postMoveWaitSeconds) {
            $postMove = [int]$cfg.behavior.postMoveWaitSeconds
            if ($postMove -gt 0) {
                Start-Sleep -Seconds $postMove
            }
        }
    }

    # 9) Summary
    Write-Log 'INFO' "----- COMPLETEARR SONARR AUTOSORTER SUMMARY -----"
    Write-Log 'INFO' ("  Series seen total            : {0}" -f $Global:CompleteARR_Summary.SeriesSeenTotal)
    Write-Log 'INFO' ("  Series in source profiles    : {0}" -f $Global:CompleteARR_Summary.SeriesInSourceProfiles)
    Write-Log 'INFO' ("  Series moved                 : {0}" -f $Global:CompleteARR_Summary.SeriesMoved)
    Write-Log 'INFO' ("  Series already correct       : {0}" -f $Global:CompleteARR_Summary.SeriesAlreadyCorrect)
    Write-Log 'INFO' ("  Series skipped (no target)   : {0}" -f $Global:CompleteARR_Summary.SeriesSkippedNoTarget)
    Write-Log 'INFO' "----- END CompleteARR SONARR AutoSorter run -----"
    
    # Return summary for master summary display
    return $Global:CompleteARR_Summary
}
catch {
    Write-Log 'ERROR' ("FATAL: CompleteARR SONARR AutoSorter failed: {0}" -f $_.Exception.Message)
    throw
}
