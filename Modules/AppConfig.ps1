#Requires -Version 5.1
<#
.SYNOPSIS
    Mapy.com Routes & Map Generator — Configuration, DPAPI Security & Logging Subsystem.
.DESCRIPTION
    Manages DPAPI-encrypted API keys, application configuration persistence,
    multi-language localization catalogs, overlay banner templates, and Mapy.com API usage tracking.
.NOTES
    Encoding: UTF-8 with BOM
    Compatibility: Windows PowerShell 5.1 and PowerShell 7+
#>

$script:AppDataDir      = Join-Path $env:LOCALAPPDATA 'MapyComRoutes'
$script:ConfigFile      = Join-Path $script:AppDataDir 'config.json'
$script:LogFile         = Join-Path $script:AppDataDir 'MapyComRoutes.log'
$script:AppConfig       = $null
$script:LocCatalog      = $null
$script:CurrentLanguage = 'en'
$script:CurrentTheme    = 'Dark'

# Ensure application working directory exists in LocalAppData, migrating legacy configuration if present
if (-not (Test-Path $script:AppDataDir)) {
    New-Item -ItemType Directory -Path $script:AppDataDir -Force | Out-Null
    $legacyDir = Join-Path $env:LOCALAPPDATA 'GoogleMapsRoutes'
    $legacyCfg = Join-Path $legacyDir 'config.json'
    if (Test-Path $legacyCfg) {
        try { Copy-Item -LiteralPath $legacyCfg -Destination $script:ConfigFile -Force } catch { }
    }
}

#region 1. DPAPI Secret Protection & Key Management

<#
.SYNOPSIS
    Encrypts a plaintext string using Windows DPAPI (CurrentUser scope) or SecureString.
.DESCRIPTION
    Converts sensitive credentials (API keys) into a protected base64-encoded string bound
    to the current Windows user profile, preventing credential leakage in plaintext config files.
.PARAMETER PlainText
    The sensitive plaintext string to encrypt.
.OUTPUTS
    [string] Base64-encoded DPAPI encrypted string, or standard SecureString representation if DPAPI is unavailable.
.EXAMPLE
    $cipher = Protect-SecretString -PlainText "my_secret_api_key"
#>
function Protect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$PlainText)

    if ([string]::IsNullOrEmpty($PlainText)) { return $null }

    try {
        # Primary protection path: Windows DPAPI via ProtectedData API
        Add-Type -AssemblyName System.Security
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        $protected = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [Convert]::ToBase64String($protected)
    }
    catch {
        # Fallback to standard PowerShell SecureString serialization if ProtectedData fails
        try {
            $sec = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
            return (ConvertFrom-SecureString -SecureString $sec)
        }
        catch {
            return $null
        }
    }
}

<#
.SYNOPSIS
    Decrypts a DPAPI-encrypted or SecureString-encoded ciphertext back to plaintext.
.DESCRIPTION
    Reverses Protect-SecretString using Windows DPAPI (CurrentUser scope). If DPAPI unprotection
    fails, attempts SecureString BSTR pointer extraction with deterministic zero-free memory cleanup.
.PARAMETER EncryptedText
    The base64-encoded DPAPI ciphertext or SecureString representation.
.OUTPUTS
    [string] Decrypted plaintext string, or $null on decryption failure.
.EXAMPLE
    $plain = Unprotect-SecretString -EncryptedText $cipherText
#>
function Unprotect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$EncryptedText)

    if ([string]::IsNullOrWhiteSpace($EncryptedText)) { return $null }

    try {
        # Primary decryption path: Windows DPAPI via ProtectedData API
        Add-Type -AssemblyName System.Security
        $bytes = [Convert]::FromBase64String($EncryptedText)
        $unprotected = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($unprotected)
    }
    catch {
        # Fallback to SecureString marshalling
        try {
            $sec = ConvertTo-SecureString -String $EncryptedText
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            try {
                return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
            }
            finally {
                # Ensure plaintext memory pointer is promptly cleared
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            }
        }
        catch {
            return $null
        }
    }
}

<#
.SYNOPSIS
    Masks an API key string for safe display in logs and user interfaces.
.DESCRIPTION
    Redacts sensitive middle characters while retaining the first 4 and last 4 characters
    to allow user verification without exposing full credentials in diagnostic logs.
.PARAMETER Key
    The API key string to mask.
.OUTPUTS
    [string] Masked representation (e.g. 'ABCD...WXYZ') or '(none)'.
.EXAMPLE
    Get-MaskedKey -Key "AIzaSyD1234567890abcdef"
#>
function Get-MaskedKey {
    [CmdletBinding()]
    param([Parameter()][string]$Key)

    if ([string]::IsNullOrWhiteSpace($Key)) { return '(none)' }
    if ($Key.Length -le 8) { return '***' }
    return "$($Key.Substring(0, 4))...$($Key.Substring($Key.Length - 4))"
}

<#
.SYNOPSIS
    Retrieves the active Mapy.com / Google Maps API key from available sources.
.DESCRIPTION
    Probes configured sources in order of precedence:
    1. Active WPF Settings tab input controls (PasswordBox or Visible TextBox)
    2. Loaded in-memory application configuration ($script:AppConfig)
    3. Process environment variables ($env:MAPY_COM_API_KEY, $env:MAPY_API_KEY, $env:GOOGLE_MAPS_API_KEY).
.OUTPUTS
    [string] The resolved API key string, or empty string if not found.
.EXAMPLE
    $apiKey = Get-CurrentApiKey
#>
function Get-CurrentApiKey {
    [CmdletBinding()]
    param()

    # 1. Check UI inputs if window/controls are active
    if ($script:Controls) {
        if ($script:Controls.txtSettingsApiKey -and $script:Controls.txtSettingsApiKey.Visibility -eq [System.Windows.Visibility]::Visible) {
            $k = [string]$script:Controls.txtSettingsApiKey.Password
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
        if ($script:Controls.txtSettingsApiKeyVisible -and $script:Controls.txtSettingsApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
            $k = [string]$script:Controls.txtSettingsApiKeyVisible.Text
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
        if ($script:Controls.txtSettingsApiKey -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtSettingsApiKey.Password)) {
            return $script:Controls.txtSettingsApiKey.Password.Trim()
        }
        if ($script:Controls.txtSettingsApiKeyVisible -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtSettingsApiKeyVisible.Text)) {
            return $script:Controls.txtSettingsApiKeyVisible.Text.Trim()
        }
    }

    # 2. Check loaded AppConfig
    if ($script:AppConfig -and -not [string]::IsNullOrWhiteSpace($script:AppConfig.ApiKey)) {
        return $script:AppConfig.ApiKey.Trim()
    }

    # 3. Check environment variables
    if (-not [string]::IsNullOrWhiteSpace($env:MAPY_COM_API_KEY)) {
        return $env:MAPY_COM_API_KEY.Trim()
    }
    if (-not [string]::IsNullOrWhiteSpace($env:MAPY_API_KEY)) {
        return $env:MAPY_API_KEY.Trim()
    }
    if (-not [string]::IsNullOrWhiteSpace($env:GOOGLE_MAPS_API_KEY)) {
        return $env:GOOGLE_MAPS_API_KEY.Trim()
    }

    return ''
}

<#
.SYNOPSIS
    Sets the active routing API key across visible and masked WPF controls.
.DESCRIPTION
    Synchronizes the specified API key value into both txtSettingsApiKey (PasswordBox)
    and txtSettingsApiKeyVisible (TextBox) controls in the WPF settings tab.
.PARAMETER Key
    The API key string to assign.
.OUTPUTS
    None.
.EXAMPLE
    Set-CurrentApiKey -Key "my_new_key"
#>
function Set-CurrentApiKey {
    [CmdletBinding()]
    param([Parameter()][string]$Key)

    if ($script:Controls) {
        if ($script:Controls.txtSettingsApiKey) { $script:Controls.txtSettingsApiKey.Password = $Key }
        if ($script:Controls.txtSettingsApiKeyVisible) { $script:Controls.txtSettingsApiKeyVisible.Text = $Key }
    }
}

<#
.SYNOPSIS
    Retrieves the active CARTO basemap API key from controls, config, or environment.
.DESCRIPTION
    Checks UI input controls, in-memory AppConfig, and $env:CARTO_API_KEY to locate
    the authentication key required for CARTO tile layer requests.
.OUTPUTS
    [string] Resolved CARTO API key or empty string.
.EXAMPLE
    $cartoKey = Get-CurrentCartoApiKey
#>
function Get-CurrentCartoApiKey {
    [CmdletBinding()]
    param()

    # 1. Check UI inputs
    if ($script:Controls) {
        if ($script:Controls.txtSettingsCartoApiKey -and $script:Controls.txtSettingsCartoApiKey.Visibility -eq [System.Windows.Visibility]::Visible) {
            $k = [string]$script:Controls.txtSettingsCartoApiKey.Password
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
        if ($script:Controls.txtSettingsCartoApiKeyVisible -and $script:Controls.txtSettingsCartoApiKeyVisible.Visibility -eq [System.Windows.Visibility]::Visible) {
            $k = [string]$script:Controls.txtSettingsCartoApiKeyVisible.Text
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
        if ($script:Controls.txtSettingsCartoApiKey -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtSettingsCartoApiKey.Password)) {
            return $script:Controls.txtSettingsCartoApiKey.Password.Trim()
        }
        if ($script:Controls.txtSettingsCartoApiKeyVisible -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtSettingsCartoApiKeyVisible.Text)) {
            return $script:Controls.txtSettingsCartoApiKeyVisible.Text.Trim()
        }
    }

    # 2. Check loaded AppConfig
    if ($script:AppConfig -and -not [string]::IsNullOrWhiteSpace($script:AppConfig.CartoApiKey)) {
        return $script:AppConfig.CartoApiKey.Trim()
    }

    # 3. Check environment variable
    if (-not [string]::IsNullOrWhiteSpace($env:CARTO_API_KEY)) {
        return $env:CARTO_API_KEY.Trim()
    }

    return ''
}

<#
.SYNOPSIS
    Sets the active CARTO basemap API key in WPF settings controls.
.DESCRIPTION
    Assigns the CARTO key to txtSettingsCartoApiKey and txtSettingsCartoApiKeyVisible.
.PARAMETER Key
    The CARTO API key string to assign.
.OUTPUTS
    None.
.EXAMPLE
    Set-CurrentCartoApiKey -Key "carto_tile_key"
#>
function Set-CurrentCartoApiKey {
    [CmdletBinding()]
    param([Parameter()][string]$Key)

    if ($script:Controls) {
        if ($script:Controls.txtSettingsCartoApiKey) { $script:Controls.txtSettingsCartoApiKey.Password = $Key }
        if ($script:Controls.txtSettingsCartoApiKeyVisible) { $script:Controls.txtSettingsCartoApiKeyVisible.Text = $Key }
    }
}

#endregion 1. DPAPI Secret Protection & Key Management

#region 2. Application Logging & Activity Drawer

$script:AppLogEntries = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())

<#
.SYNOPSIS
    Refreshes the activity log drawer text display based on the selected log level filter.
.DESCRIPTION
    Filters buffered in-memory log entries against the active radio button filter (ALL, INFO,
    WARN, ERROR) and updates txtLogDrawer in the UI, scrolling automatically to the newest entry.
.OUTPUTS
    None.
.EXAMPLE
    Update-LogDrawerDisplay
#>
function Update-LogDrawerDisplay {
    [CmdletBinding()]
    param()

    if (-not $script:Controls -or -not $script:Controls.txtLogDrawer) { return }

    $filter = if ($script:Controls.rbLogInfo -and $script:Controls.rbLogInfo.IsChecked) { 'INFO' }
              elseif ($script:Controls.rbLogWarn -and $script:Controls.rbLogWarn.IsChecked) { 'WARN' }
              elseif ($script:Controls.rbLogError -and $script:Controls.rbLogError.IsChecked) { 'ERROR' }
              else { 'ALL' }

    $sb = [System.Text.StringBuilder]::new()
    $entries = @($script:AppLogEntries.ToArray())
    foreach ($entry in $entries) {
        if ($filter -eq 'ALL' -or $entry.Level -eq $filter -or ($filter -eq 'INFO' -and $entry.Level -eq 'OK')) {
            [void]$sb.AppendLine($entry.FullLine)
        }
    }
    $script:Controls.txtLogDrawer.Text = $sb.ToString()
    $script:Controls.txtLogDrawer.ScrollToEnd()
}

<#
.SYNOPSIS
    Clears all buffered application log entries and empties the log drawer UI.
.DESCRIPTION
    Resets the thread-safe in-memory log collection and empties the txtLogDrawer TextBox control.
.OUTPUTS
    None.
.EXAMPLE
    Clear-AppLogDrawer
#>
function Clear-AppLogDrawer {
    [CmdletBinding()]
    param()

    $script:AppLogEntries.Clear()
    if ($script:Controls -and $script:Controls.txtLogDrawer) {
        $script:Controls.txtLogDrawer.Clear()
    }
}

<#
.SYNOPSIS
    Writes an entry to the disk log file and updates the in-memory GUI log drawer.
.DESCRIPTION
    Appends a timestamped log line to MapyComRoutes.log, automatically redacting any
    contained API keys to prevent secret exposure. Appends the entry to the circular in-memory
    buffer (capped at 600 items) and dispatches UI updates to the WPF dispatcher thread.
.PARAMETER Message
    The log message text.
.PARAMETER Level
    Severity level ('INFO', 'OK', 'WARN', 'ERROR', 'DEBUG'). Defaults to 'INFO'.
.OUTPUTS
    None.
.EXAMPLE
    Write-AppLog -Message "Route calculated successfully." -Level OK
#>
function Write-AppLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter()][ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO'
    )

    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $safeMsg = $Message

    # Redact active API key from log output for security
    $cfg = Get-Variable -Scope Script -Name 'AppConfig' -ValueOnly -ErrorAction SilentlyContinue
    if ($cfg -and $cfg.ApiKey -and $cfg.ApiKey.Length -gt 8) {
        $masked = Get-MaskedKey $cfg.ApiKey
        $safeMsg = $safeMsg.Replace($cfg.ApiKey, $masked)
    }
    $line = "[$ts] [$Level] $safeMsg"

    # Append to persistent disk log file
    try {
        if ($script:LogFile) {
            [System.IO.File]::AppendAllText($script:LogFile, "$line`r`n", [System.Text.Encoding]::UTF8)
        }
    }
    catch { }

    # Buffer into circular in-memory array for Activity Log Drawer
    try {
        $entry = [PSCustomObject]@{
            Timestamp = $ts
            Level     = $Level
            Message   = $safeMsg
            FullLine  = $line
        }
        if ($script:AppLogEntries.Count -ge 600) {
            $script:AppLogEntries.RemoveAt(0)
        }
        [void]$script:AppLogEntries.Add($entry)

        # Dispatch updates to WPF UI thread safely
        $w = if ($script:Controls -and $script:Controls.Window) { 
            $script:Controls.Window 
        } elseif ($script:MainWindow) { 
            $script:MainWindow 
        } else { 
            [System.Windows.Application]::Current.MainWindow 
        }

        if ($script:Controls -and $script:Controls.txtLogDrawer -and $w -and $w.Dispatcher) {
            $w.Dispatcher.BeginInvoke([Action]{
                $filter = if ($script:Controls.rbLogInfo -and $script:Controls.rbLogInfo.IsChecked) { 'INFO' }
                          elseif ($script:Controls.rbLogWarn -and $script:Controls.rbLogWarn.IsChecked) { 'WARN' }
                          elseif ($script:Controls.rbLogError -and $script:Controls.rbLogError.IsChecked) { 'ERROR' }
                          else { 'ALL' }
                if ($filter -eq 'ALL' -or $Level -eq $filter -or ($filter -eq 'INFO' -and $Level -eq 'OK')) {
                    $script:Controls.txtLogDrawer.AppendText("$line`r`n")
                    $script:Controls.txtLogDrawer.ScrollToEnd()
                }
            }) | Out-Null
        }
    }
    catch { }
}

#endregion 2. Application Logging & Activity Drawer

#region 3. Localization Catalog & Translation Engine

$script:LocCatalog = $null
$script:CurrentLanguage = 'en'

<#
.SYNOPSIS
    Loads the multi-language localization catalog from localization.json.
.DESCRIPTION
    Probes script folders, parent directories, and LocalAppData for localization.json.
    Parses the JSON structure into an in-memory object and caches it into $script:LocCatalog.
    Falls back to embedded string resources if available.
.OUTPUTS
    [PSCustomObject] The loaded localization catalog object, or $null on failure.
.EXAMPLE
    $catalog = Load-LocalizationConfig
#>
function Load-LocalizationConfig {
    [CmdletBinding()]
    param()

    $locPath = $null
    $baseDir = if (-not [string]::IsNullOrWhiteSpace($script:AppDir)) {
        $script:AppDir
    } elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    } else {
        [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    }

    # Probing order for localization.json: app dir -> parent dir -> AppData
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($baseDir)) {
        $candidates.Add((Join-Path $baseDir 'localization.json'))
        $parentDir = Split-Path -Parent $baseDir
        if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
            $candidates.Add((Join-Path $parentDir 'localization.json'))
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($script:AppDataDir)) {
        $candidates.Add((Join-Path $script:AppDataDir 'localization.json'))
    }

    foreach ($cand in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path $cand)) {
            $locPath = (Resolve-Path $cand).Path
            break
        }
    }

    if ($locPath -and (Test-Path $locPath)) {
        try {
            $jsonRaw = [System.IO.File]::ReadAllText($locPath, [System.Text.Encoding]::UTF8)
            $parsed = $jsonRaw | ConvertFrom-Json
            if ($parsed.Languages) {
                $script:LocCatalog = $parsed
                Write-AppLog "Loaded localization catalog from: $locPath" "OK"
                return $parsed
            }
        }
        catch {
            Write-AppLog "Error parsing $($locPath): $($_.Exception.Message)" "WARN"
        }
    }

    # In-memory fallback if compiled with embedded JSON
    if ($script:EmbeddedLocalizationJson) {
        try {
            $script:LocCatalog = $script:EmbeddedLocalizationJson | ConvertFrom-Json
            return $script:LocCatalog
        }
        catch { }
    }
    return $null
}

<#
.SYNOPSIS
    Resolves a localized string from the loaded localization catalog.
.DESCRIPTION
    Looks up a string key for the active application language ($script:CurrentLanguage).
    If missing, falls back to the English ('en') translation, then to $DefaultText,
    and finally to the raw $Key name.
.PARAMETER Key
    The string key identifier defined in localization.json (e.g. 'btnCalculate').
.PARAMETER DefaultText
    Optional fallback text returned when key is not found.
.OUTPUTS
    [string] The localized string.
.EXAMPLE
    $label = Get-LocText -Key "lblOrigin" -DefaultText "Origin Address"
#>
function Get-LocText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter()][string]$DefaultText = ''
    )

    if (-not $script:LocCatalog -or -not $script:LocCatalog.Languages) {
        return $(if ($DefaultText) { $DefaultText } else { $Key })
    }

    $lang = if ($script:CurrentLanguage) { $script:CurrentLanguage } else { 'en' }
    $langObj = $script:LocCatalog.Languages.$lang
    if ($langObj -and $langObj.Strings -and ($langObj.Strings.PSObject.Properties.Name -contains $Key)) {
        $val = $langObj.Strings.$Key
        if (-not [string]::IsNullOrWhiteSpace($val)) { return [string]$val }
    }

    # Fallback to English (en)
    $enObj = $script:LocCatalog.Languages.en
    if ($enObj -and $enObj.Strings -and ($enObj.Strings.PSObject.Properties.Name -contains $Key)) {
        $val = $enObj.Strings.$Key
        if (-not [string]::IsNullOrWhiteSpace($val)) { return [string]$val }
    }

    return $(if ($DefaultText) { $DefaultText } else { $Key })
}

#endregion 3. Localization Catalog & Translation Engine

#region 4. Application Configuration Defaults & Persistence

<#
.SYNOPSIS
    Returns default layout configuration for static map overlay banners.
.DESCRIPTION
    Defines top and bottom banner visibility, field alignments, and ordering
    for route names, geocoded addresses, distance, duration, and timestamp badges.
.OUTPUTS
    [ordered] Ordered hashtable containing overlay configuration schema.
.EXAMPLE
    $overlayCfg = Get-DefaultOverlayConfig
#>
function Get-DefaultOverlayConfig {
    [CmdletBinding()]
    param()

    return [ordered]@{
        EnableTopOverlay    = $true
        EnableBottomOverlay = $true
        Properties          = [ordered]@{
            RouteName      = [ordered]@{ Enabled = $true;  Panel = 'Top';    Alignment = 'Left';   Order = 1 }
            RouteType      = [ordered]@{ Enabled = $true;  Panel = 'Top';    Alignment = 'Right';  Order = 1 }
            Timestamp      = [ordered]@{ Enabled = $false; Panel = 'Top';    Alignment = 'Right';  Order = 2 }
            StartGeocoded  = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 1 }
            EndGeocoded    = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 2 }
            Distance       = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 3 }
            Duration       = [ordered]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 3 }
            Waypoints      = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Alignment = 'Left';   Order = 4 }
            PointDistances = [ordered]@{ Enabled = $false; Panel = 'None';   Alignment = 'Left';   Order = 5 }
            StartRaw       = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Alignment = 'Left';   Order = 6 }
            EndRaw         = [ordered]@{ Enabled = $false; Panel = 'Bottom'; Alignment = 'Left';   Order = 7 }
        }
    }
}

<#
.SYNOPSIS
    Generates a default application configuration object with factory presets.
.DESCRIPTION
    Provides initial default settings including default output directory in Documents,
    Dark mode theme, English language, fastest routing mode, and blank API usage counters.
.OUTPUTS
    [ordered] Default application configuration hashtable.
.EXAMPLE
    $defaultCfg = Get-DefaultAppConfig
#>
function Get-DefaultAppConfig {
    [CmdletBinding()]
    param()

    $defaultOutput = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('MyDocuments'), 'TrasyMapyCom')
    $currentMonth = (Get-Date).ToString('yyyy-MM')

    return [ordered]@{
        ApiKey               = ''
        CartoApiKey          = ''
        RememberKey          = $true
        DefaultRouteType     = 'Fastest'
        DefaultEmissionType  = 'GASOLINE'
        MapWidth             = 900
        MapHeight            = 600
        OutputDirectory      = $defaultOutput
        Language             = 'en'
        Theme                = 'Dark'
        UseInteractiveMap    = $true
        AvoidTolls           = $false
        AvoidHighways        = $false
        AvoidFerries         = $false
        OverlayConfig        = (Get-DefaultOverlayConfig)
        ApiUsage             = [ordered]@{
            CurrentMonth          = $currentMonth
            MonthlyCallsGeocoding = 0
            MonthlyCallsRoutes    = 0
            MonthlyCallsStatic    = 0
            SessionCallsGeocoding = 0
            SessionCallsRoutes    = 0
            SessionCallsStatic    = 0
            BudgetLimitUsd        = 50.0
            PreferredCurrency     = 'USD'
        }
        RecentRoutes         = @()
    }
}

<#
.SYNOPSIS
    Loads persisted application settings from config.json and decrypts credentials.
.DESCRIPTION
    Reads config.json from %LOCALAPPDATA%\MapyComRoutes. Unprotects encrypted API keys via
    Unprotect-SecretString, merges existing values over factory defaults, handles monthly
    usage counter resets, and caches the result in $script:AppConfig.
.OUTPUTS
    [ordered] Populated application configuration object.
.EXAMPLE
    $cfg = Load-AppConfig
#>
function Load-AppConfig {
    [CmdletBinding()]
    param()

    $result = Get-DefaultAppConfig

    if (-not (Test-Path $script:ConfigFile)) {
        $script:AppConfig = $result
        $script:CurrentLanguage = $result.Language
        return $result
    }

    try {
        $raw = [System.IO.File]::ReadAllText($script:ConfigFile, [System.Text.Encoding]::UTF8)
        $cfg = $raw | ConvertFrom-Json

        # Decrypt stored DPAPI credentials
        if ($cfg.ApiKey) {
            $decrypted = Unprotect-SecretString -EncryptedText $cfg.ApiKey
            $result.ApiKey = if ($decrypted) { $decrypted } else { '' }
        }
        if ($cfg.CartoApiKey) {
            $decryptedCarto = Unprotect-SecretString -EncryptedText $cfg.CartoApiKey
            $result.CartoApiKey = if ($decryptedCarto) { $decryptedCarto } else { [string]$cfg.CartoApiKey }
        }
        if ($null -ne $cfg.RememberKey)         { $result.RememberKey = [bool]$cfg.RememberKey }
        if ($cfg.DefaultRouteType)             { $result.DefaultRouteType = [string]$cfg.DefaultRouteType }
        if ($cfg.DefaultEmissionType)          { $result.DefaultEmissionType = [string]$cfg.DefaultEmissionType }
        if ($cfg.MapWidth -gt 0)               { $result.MapWidth = [int]$cfg.MapWidth }
        if ($cfg.MapHeight -gt 0)              { $result.MapHeight = [int]$cfg.MapHeight }
        if ($cfg.OutputDirectory)              { $result.OutputDirectory = [string]$cfg.OutputDirectory }
        if ($cfg.Language)                     { $result.Language = [string]$cfg.Language }
        if ($cfg.Theme)                        { $result.Theme = [string]$cfg.Theme }
        if ($null -ne $cfg.UseInteractiveMap)  { $result.UseInteractiveMap = [bool]$cfg.UseInteractiveMap }
        if ($null -ne $cfg.AvoidTolls)         { $result.AvoidTolls = [bool]$cfg.AvoidTolls }
        if ($null -ne $cfg.AvoidHighways)      { $result.AvoidHighways = [bool]$cfg.AvoidHighways }
        if ($null -ne $cfg.AvoidFerries)       { $result.AvoidFerries = [bool]$cfg.AvoidFerries }

        if ($cfg.OverlayConfig) {
            $result.OverlayConfig = $cfg.OverlayConfig
        }

        # API Usage tracking and automatic monthly roll-over
        if ($cfg.ApiUsage) {
            $u = $cfg.ApiUsage
            $currentMonth = (Get-Date).ToString('yyyy-MM')
            if ($u.CurrentMonth -eq $currentMonth) {
                $result.ApiUsage.CurrentMonth          = $currentMonth
                $result.ApiUsage.MonthlyCallsGeocoding = [int]$u.MonthlyCallsGeocoding
                $result.ApiUsage.MonthlyCallsRoutes    = [int]$u.MonthlyCallsRoutes
                $result.ApiUsage.MonthlyCallsStatic    = [int]$u.MonthlyCallsStatic
            } else {
                # Reset monthly counters for new billing month
                $result.ApiUsage.CurrentMonth          = $currentMonth
                $result.ApiUsage.MonthlyCallsGeocoding = 0
                $result.ApiUsage.MonthlyCallsRoutes    = 0
                $result.ApiUsage.MonthlyCallsStatic    = 0
            }
            if ($u.BudgetLimitUsd)      { $result.ApiUsage.BudgetLimitUsd = [double]$u.BudgetLimitUsd }
            if ($u.PreferredCurrency)  { $result.ApiUsage.PreferredCurrency = [string]$u.PreferredCurrency }
        }

        if ($cfg.RecentRoutes) {
            $result.RecentRoutes = @($cfg.RecentRoutes)
        } else {
            $result.RecentRoutes = @()
        }

        $script:AppConfig = $result
        $script:CurrentLanguage = $result.Language
        return $result
    }
    catch {
        Write-AppLog "Error loading config: $($_.Exception.Message)" "ERROR"
        $script:AppConfig = Get-DefaultAppConfig
        $script:CurrentLanguage = $script:AppConfig.Language
        return $script:AppConfig
    }
}

<#
.SYNOPSIS
    Saves the provided configuration object to config.json with DPAPI key protection.
.DESCRIPTION
    Encrypts API keys via Protect-SecretString if RememberKey is enabled, serializes the
    configuration to JSON with UTF-8 BOM encoding, and updates script-level state caches.
.PARAMETER Config
    The configuration object to persist.
.OUTPUTS
    [bool] $true if save succeeded; otherwise $false.
.EXAMPLE
    Save-AppConfig -Config $script:AppConfig
#>
function Save-AppConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Config)

    try {
        $plainKey = if ($Config.ApiKey) { [string]$Config.ApiKey } else { '' }
        $plainCartoKey = if ($Config.CartoApiKey) { [string]$Config.CartoApiKey } else { '' }
        $remember = if ($null -ne $Config.RememberKey) { [bool]$Config.RememberKey } else { $true }

        # Encrypt API keys before writing to disk
        $encryptedKey = if ($remember -and -not [string]::IsNullOrWhiteSpace($plainKey)) {
            Protect-SecretString -PlainText $plainKey
        } else {
            ''
        }

        $encryptedCartoKey = if ($remember -and -not [string]::IsNullOrWhiteSpace($plainCartoKey)) {
            Protect-SecretString -PlainText $plainCartoKey
        } else {
            ''
        }

        $toSave = [ordered]@{
            ApiKey              = $encryptedKey
            CartoApiKey         = $encryptedCartoKey
            RememberKey         = $remember
            DefaultRouteType    = [string]$Config.DefaultRouteType
            DefaultEmissionType = [string]$Config.DefaultEmissionType
            MapWidth            = [int]$Config.MapWidth
            MapHeight           = [int]$Config.MapHeight
            OutputDirectory     = [string]$Config.OutputDirectory
            Language            = [string]$Config.Language
            Theme               = [string]$Config.Theme
            UseInteractiveMap   = [bool]$Config.UseInteractiveMap
            AvoidTolls          = [bool]$Config.AvoidTolls
            AvoidHighways       = [bool]$Config.AvoidHighways
            AvoidFerries        = [bool]$Config.AvoidFerries
            OverlayConfig       = $Config.OverlayConfig
            ApiUsage            = $Config.ApiUsage
            RecentRoutes        = if ($Config.RecentRoutes) { @($Config.RecentRoutes) } else { @() }
        }

        # Write formatted JSON using UTF-8 with BOM
        $json = $toSave | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($script:ConfigFile, $json, [System.Text.UTF8Encoding]::new($true))
        
        $script:AppConfig = $Config
        if ($Config.Language) { $script:CurrentLanguage = [string]$Config.Language }
        if ($Config.Theme)    { $script:CurrentTheme = [string]$Config.Theme }
        Write-AppLog "Configuration saved to: $script:ConfigFile" "OK"
        return $true
    }
    catch {
        Write-AppLog "Error saving configuration: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

<#
.SYNOPSIS
    Appends a calculated route to the recent route history list in application settings.
.DESCRIPTION
    Inserts the recent route record at the top of the history list, deduplicates matching
    start/end pairs, limits history length to 10 entries, and updates config.json.
.PARAMETER Start
    Origin address string.
.PARAMETER End
    Destination address string.
.PARAMETER Waypoints
    Array of intermediate stop address strings.
.PARAMETER RouteName
    Optional descriptive name for the route.
.PARAMETER RouteType
    Optimization mode ('Fastest', 'Shortest', 'Eco'). Defaults to 'Fastest'.
.OUTPUTS
    None.
.EXAMPLE
    Add-RecentRouteConfig -Start "Warsaw" -End "Krakow" -RouteName "Business Trip"
#>
function Add-RecentRouteConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Start,
        [Parameter(Mandatory = $true)][string]$End,
        [Parameter()][string[]]$Waypoints = @(),
        [Parameter()][string]$RouteName = '',
        [Parameter()][string]$RouteType = 'Fastest'
    )

    if ([string]::IsNullOrWhiteSpace($Start) -or [string]::IsNullOrWhiteSpace($End)) { return }
    if (-not $script:AppConfig) { $script:AppConfig = Load-AppConfig }
    if (-not $script:AppConfig.RecentRoutes) { $script:AppConfig.RecentRoutes = @() }

    $newItem = [ordered]@{
        Start     = $Start.Trim()
        End       = $End.Trim()
        Waypoints = @($Waypoints)
        Name      = if ($RouteName) { $RouteName.Trim() } else { "$($Start.Trim()) -> $($End.Trim())" }
        RouteType = $RouteType
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    }

    $list = [System.Collections.Generic.List[object]]::new()
    $list.Add($newItem)
    foreach ($r in $script:AppConfig.RecentRoutes) {
        if (-not ($r.Start -eq $newItem.Start -and $r.End -eq $newItem.End)) {
            $list.Add($r)
            if ($list.Count -ge 10) { break }
        }
    }
    $script:AppConfig.RecentRoutes = @($list.ToArray())
    Save-AppConfig -Config $script:AppConfig | Out-Null
}

<#
.SYNOPSIS
    Increments API call counters in application configuration and saves state.
.DESCRIPTION
    Updates session and monthly call counts for Geocoding, Routes, and Static Maps APIs.
    Preserves active in-memory credentials and persists updated totals to disk.
.PARAMETER GeocodingInc
    Number of geocoding API requests to add.
.PARAMETER RoutesInc
    Number of route calculation API requests to add.
.PARAMETER StaticMapsInc
    Number of static map image requests to add.
.OUTPUTS
    None.
.EXAMPLE
    Update-ApiUsageRecord -GeocodingInc 2 -RoutesInc 1 -StaticMapsInc 1
#>
function Update-ApiUsageRecord {
    [CmdletBinding()]
    param(
        [Parameter()][int]$GeocodingInc = 0,
        [Parameter()][int]$RoutesInc = 0,
        [Parameter()][int]$StaticMapsInc = 0
    )

    if (-not $script:AppConfig -or -not $script:AppConfig.ApiUsage) { return }

    # Preserve active key if in-memory copy was temporarily blanked
    if ([string]::IsNullOrWhiteSpace($script:AppConfig.ApiKey)) {
        $activeKey = Get-CurrentApiKey
        if (-not [string]::IsNullOrWhiteSpace($activeKey)) {
            $script:AppConfig.ApiKey = $activeKey
        }
    }

    $u = $script:AppConfig.ApiUsage
    $currentMonth = (Get-Date).ToString('yyyy-MM')
    if ($u.CurrentMonth -ne $currentMonth) {
        $u.CurrentMonth = $currentMonth
        $u.MonthlyCallsGeocoding = 0
        $u.MonthlyCallsRoutes = 0
        $u.MonthlyCallsStatic = 0
    }

    $u.SessionCallsGeocoding += $GeocodingInc
    $u.SessionCallsRoutes    += $RoutesInc
    $u.SessionCallsStatic    += $StaticMapsInc

    $u.MonthlyCallsGeocoding += $GeocodingInc
    $u.MonthlyCallsRoutes    += $RoutesInc
    $u.MonthlyCallsStatic    += $StaticMapsInc

    Save-AppConfig -Config $script:AppConfig | Out-Null
}

#endregion 4. Application Configuration Defaults & Persistence

#region 5. Toast Notifications Subsystem

$script:ToastTimer = $null

<#
.SYNOPSIS
    Displays an animated floating toast notification in the WPF application window.
.DESCRIPTION
    Configures and displays the floating notification border with status icon, title,
    message body, optional action buttons (Open File, Open Folder), and an automatic
    dismissal DispatcherTimer. Safely dispatches calls from background threads to the STA UI thread.
.PARAMETER Title
    Notification title heading.
.PARAMETER Message
    Notification message body text.
.PARAMETER Type
    Status style ('Success', 'Info', 'Warning', 'Error'). Defaults to 'Success'.
.PARAMETER ActionFile
    Optional file path linked to the 'Open File' action button.
.PARAMETER ActionFolder
    Optional directory path linked to the 'Open Folder' action button.
.PARAMETER DurationSec
    Duration in seconds before toast is automatically dismissed. Defaults to 6.
.OUTPUTS
    None.
.EXAMPLE
    Show-AppToastNotification -Title "Export Complete" -Message "PDF report generated." -Type Success `
        -ActionFile "C:\Reports\Route.pdf"
#>
function Show-AppToastNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter()][ValidateSet('Success', 'Info', 'Warning', 'Error')][string]$Type = 'Success',
        [Parameter()][string]$ActionFile = '',
        [Parameter()][string]$ActionFolder = '',
        [Parameter()][int]$DurationSec = 6
    )

    if (-not $script:Controls -or -not $script:Controls.pnlToastContainer) { return }

    # Ensure execution on STA UI thread via dispatcher invocation
    $w = if ($script:Controls -and $script:Controls.Window) { 
        $script:Controls.Window 
    } elseif ($script:MainWindow) { 
        $script:MainWindow 
    } else { 
        [System.Windows.Application]::Current.MainWindow 
    }

    if ($w -and $w.Dispatcher -and -not $w.Dispatcher.CheckAccess()) {
        $w.Dispatcher.BeginInvoke([Action]{
            Show-AppToastNotification -Title $Title -Message $Message -Type $Type -ActionFile $ActionFile -ActionFolder $ActionFolder -DurationSec $DurationSec
        }) | Out-Null
        return
    }

    $ctrl = $script:Controls
    $icon = switch ($Type) {
        'Success' { '✅' }
        'Info'    { 'ℹ️' }
        'Warning' { '⚠️' }
        'Error'   { '❌' }
    }
    $borderColor = switch ($Type) {
        'Success' { '#10B981' }
        'Info'    { '#38BDF8' }
        'Warning' { '#F59E0B' }
        'Error'   { '#EF4444' }
    }

    if ($ctrl.txtToastIcon) { $ctrl.txtToastIcon.Text = $icon }
    if ($ctrl.txtToastTitle) { $ctrl.txtToastTitle.Text = $Title }
    if ($ctrl.txtToastMessage) { $ctrl.txtToastMessage.Text = $Message }
    $ctrl.pnlToastContainer.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($borderColor)

    # Configure optional 'Open File' action button
    if ($ctrl.btnToastAction) {
        if (-not [string]::IsNullOrWhiteSpace($ActionFile) -and (Test-Path $ActionFile)) {
            $ctrl.btnToastAction.Visibility = [System.Windows.Visibility]::Visible
            $ctrl.btnToastAction.Tag = $ActionFile
        } else {
            $ctrl.btnToastAction.Visibility = [System.Windows.Visibility]::Collapsed
        }
    }

    # Configure optional 'Open Folder' action button
    if ($ctrl.btnToastActionFolder) {
        $folderToOpen = if (-not [string]::IsNullOrWhiteSpace($ActionFolder) -and (Test-Path $ActionFolder)) {
            $ActionFolder
        } elseif (-not [string]::IsNullOrWhiteSpace($ActionFile) -and (Test-Path $ActionFile)) {
            [System.IO.Path]::GetDirectoryName($ActionFile)
        } else { '' }

        if (-not [string]::IsNullOrWhiteSpace($folderToOpen) -and (Test-Path $folderToOpen)) {
            $ctrl.btnToastActionFolder.Visibility = [System.Windows.Visibility]::Visible
            $ctrl.btnToastActionFolder.Tag = $folderToOpen
        } else {
            $ctrl.btnToastActionFolder.Visibility = [System.Windows.Visibility]::Collapsed
        }
    }

    $ctrl.pnlToastContainer.Visibility = [System.Windows.Visibility]::Visible

    # Reset and launch auto-dismiss timer
    if ($script:ToastTimer) {
        try { $script:ToastTimer.Stop() } catch { }
    }
    $script:ToastTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:ToastTimer.Interval = [TimeSpan]::FromSeconds($DurationSec)
    $script:ToastTimer.Add_Tick({
        if ($script:Controls -and $script:Controls.pnlToastContainer) {
            $script:Controls.pnlToastContainer.Visibility = [System.Windows.Visibility]::Collapsed
        }
        if ($script:ToastTimer) { $script:ToastTimer.Stop() }
    })
    $script:ToastTimer.Start()
}

#endregion 5. Toast Notifications Subsystem

#region 6. Global Function Exports

# Export functions into global scope for caller scripts and GUI orchestrators
Set-Item -Path "function:global:Protect-SecretString" -Value (Get-Item "function:Protect-SecretString").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Unprotect-SecretString" -Value (Get-Item "function:Unprotect-SecretString").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Get-MaskedKey" -Value (Get-Item "function:Get-MaskedKey").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Get-CurrentApiKey" -Value (Get-Item "function:Get-CurrentApiKey").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Set-CurrentApiKey" -Value (Get-Item "function:Set-CurrentApiKey").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Get-CurrentCartoApiKey" -Value (Get-Item "function:Get-CurrentCartoApiKey").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Set-CurrentCartoApiKey" -Value (Get-Item "function:Set-CurrentCartoApiKey").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Update-LogDrawerDisplay" -Value (Get-Item "function:Update-LogDrawerDisplay").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Clear-AppLogDrawer" -Value (Get-Item "function:Clear-AppLogDrawer").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Write-AppLog" -Value (Get-Item "function:Write-AppLog").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Load-LocalizationConfig" -Value (Get-Item "function:Load-LocalizationConfig").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Get-LocText" -Value (Get-Item "function:Get-LocText").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Get-DefaultOverlayConfig" -Value (Get-Item "function:Get-DefaultOverlayConfig").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Get-DefaultAppConfig" -Value (Get-Item "function:Get-DefaultAppConfig").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Load-AppConfig" -Value (Get-Item "function:Load-AppConfig").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Save-AppConfig" -Value (Get-Item "function:Save-AppConfig").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Add-RecentRouteConfig" -Value (Get-Item "function:Add-RecentRouteConfig").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Update-ApiUsageRecord" -Value (Get-Item "function:Update-ApiUsageRecord").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Show-AppToastNotification" -Value (Get-Item "function:Show-AppToastNotification").ScriptBlock -ErrorAction SilentlyContinue

#endregion 6. Global Function Exports
