#Requires -Version 5.1
<#
.SYNOPSIS
    Shared functions for address geocoding, route calculation (Fastest, Shortest),
    static PNG map rendering, and multi-format data file processing (JSON, CSV, Excel).

.DESCRIPTION
    Function library utilizing Mapy.com REST API (api.mapy.cz/v1).
    Used by:
      - MapyComRoutes-GUI.ps1
      - Invoke-MapyComRoute.ps1
      - Process-SchoolTransportRoutes.ps1

    Core Functions:
      - Select-InputDataFile / Select-InputExcel : Interactive file selection dialogs
      - Protect-SecretString / Unprotect-SecretString : Secure DPAPI per-user API key encryption
      - Test-MapyApiKey : Non-blocking API key validity verification
      - Get-AddressCoordinates : Address geocoding via Mapy.com Geocoding API
      - Get-GeocodeStatusDescription : Detailed geocode accuracy and fallback status resolver
      - Get-CarRouteData : Route calculation via Mapy.com Routing API (Fastest, Shortest, multipoint)
      - Get-MapyComUrl : Navigation URL generator for Mapy.com in web browser
      - Save-RouteMapPng : Map image retrieval via Mapy.com Static Maps API with markers and overlay banners
      - Import-RouteDataFile : Universal route data loader for JSON, CSV, and Excel (XLSX, XLS)
      - Export-RouteResults : Universal route and waypoint exporter for Excel, CSV, and JSON

.NOTES
    Requires MAPY_COM_API_KEY environment variable or -ApiKey parameter.
    Encoding: UTF-8 with BOM
#>

# ══════════════════════════════════════════════════════════════════════════════
# 1. DIALOGS AND SECURITY (DPAPI)
# ══════════════════════════════════════════════════════════════════════════════

function Select-InputExcel {
    param([string]$InitialDirectory)
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $Dialog.Title = 'Select Excel file with addresses'
    $Dialog.Filter = 'Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*'
    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $Dialog.InitialDirectory = $InitialDirectory
    } else {
        $Dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    }
    $Dialog.RestoreDirectory = $true
    $Result = $Dialog.ShowDialog()
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) { return $Dialog.FileName }
    return $null
}

function Select-InputDataFile {
    param([string]$InitialDirectory)
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $Dialog.Title = 'Select route data file (JSON, CSV, Excel)'
    $Dialog.Filter = 'All Supported Files (*.xlsx;*.xls;*.csv;*.tsv;*.json)|*.xlsx;*.xls;*.csv;*.tsv;*.json|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV/TSV Files (*.csv;*.tsv)|*.csv;*.tsv|JSON Files (*.json)|*.json|All Files (*.*)|*.*'
    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $Dialog.InitialDirectory = $InitialDirectory
    } else {
        $Dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    }
    $Dialog.RestoreDirectory = $true
    $Result = $Dialog.ShowDialog()
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) { return $Dialog.FileName }
    return $null
}

function Protect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PlainText)
    if ([string]::IsNullOrEmpty($PlainText)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        $protected = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [Convert]::ToBase64String($protected)
    }
    catch {
        $sec = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
        return (ConvertFrom-SecureString -SecureString $sec)
    }
}

function Unprotect-SecretString {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$EncryptedText)
    if ([string]::IsNullOrWhiteSpace($EncryptedText)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $bytes = [Convert]::FromBase64String($EncryptedText)
        $unprotected = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($unprotected)
    }
    catch {
        try {
            $sec = ConvertTo-SecureString -String $EncryptedText
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            $str = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return $str
        }
        catch {
            return $null
        }
    }
}

function Invoke-MapyRestJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()][int]$TimeoutSec = 30
    )
    $req = [System.Net.HttpWebRequest]::Create($Uri)
    $req.Timeout = [math]::Max(1000, $TimeoutSec * 1000)
    $req.ReadWriteTimeout = [math]::Max(1000, $TimeoutSec * 1000)
    $req.Headers.Add('Accept-Charset', 'utf-8')
    $resp = $req.GetResponse()
    try {
        $stream = $resp.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        try {
            $jsonStr = $reader.ReadToEnd()
            return ($jsonStr | ConvertFrom-Json)
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $resp.Dispose()
    }
}
Set-Item -Path "function:global:Invoke-MapyRestJson" -Value (Get-Item "function:Invoke-MapyRestJson").ScriptBlock -ErrorAction SilentlyContinue

function Test-MapyApiKey {
    [CmdletBinding()]
    [Alias('Test-GoogleApiKey')]
    param(
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en'
    )
    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        return [PSCustomObject]@{ Valid = $false; Message = 'Klucz API jest pusty.' }
    }
    try {
        $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
        $Url = "https://api.mapy.cz/v1/geocode?query=Praha&lang=$lang&limit=1&apikey=$ApiKey"
        $Resp = Invoke-MapyRestJson -Uri $Url -TimeoutSec 15
        if ($Resp.items -and $Resp.items.Count -gt 0) {
            return [PSCustomObject]@{ Valid = $true; Message = 'Klucz Mapy.com API jest poprawny i aktywny.' }
        }
        else {
            return [PSCustomObject]@{ Valid = $true; Message = 'Klucz Mapy.com API jest poprawny (brak wyników dla testu, ale klucz zaakceptowany).' }
        }
    }
    catch {
        $ex = $_.Exception
        $webEx = if ($ex -is [System.Net.WebException]) { $ex } elseif ($ex.InnerException -is [System.Net.WebException]) { $ex.InnerException } else { $null }
        $statusCode = if ($webEx -and $webEx.Response) { [int]$webEx.Response.StatusCode } elseif ($ex.Response) { [int]$ex.Response.StatusCode } else { 0 }
        if ($statusCode -eq 401 -or $statusCode -eq 403) {
            return [PSCustomObject]@{ Valid = $false; Message = "Brak autoryzacji: klucz API odrzucony (HTTP $statusCode)." }
        }
        return [PSCustomObject]@{ Valid = $false; Message = "Błąd połączenia: $($ex.Message)" }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 2. ADDRESS GEOCODING (MAPY.COM GEOCODING API)
# ══════════════════════════════════════════════════════════════════════════════

function Get-AddressComponentValue {
    [CmdletBinding()]
    param(
        [Parameter()][object[]]$Components,
        [Parameter(Mandatory)][string[]]$Types
    )

    $Matches = @($Components) | Where-Object {
        $_ -and $_.PSObject.Properties.Name -contains 'type' -and
        ([string]$_.type -in $Types)
    } | Select-Object -First 1

    if (-not $Matches) { return $null }

    foreach ($FieldName in @('name', 'long_name', 'short_name')) {
        if ($Matches.PSObject.Properties.Name -contains $FieldName) {
            $Value = [string]$Matches.$FieldName
            if (-not [string]::IsNullOrWhiteSpace($Value)) {
                return $Value
            }
        }
    }

    return $null
}

function Get-AddressCoordinates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Address,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][switch]$RequireStreetNumber
    )
    if ([string]::IsNullOrWhiteSpace($Address)) { return $null }

    # If coordinates are provided directly in format "52.2297, 21.0122"
    if ($Address.Trim() -match '^\s*([+-]?\d+(?:\.\d+)?)\s*[,;\s]\s*([+-]?\d+(?:\.\d+)?)\s*$') {
        $lat = [double]$Matches[1]
        $lng = [double]$Matches[2]
        return [PSCustomObject]@{
            Latitude             = $lat
            Longitude            = $lng
            FormattedAddress     = "$lat, $lng"
            UlicaINumer          = $null
            KodPocztowy          = $null
            Miasto               = $null
            MatchType            = 'COORDINATES'
            PartialMatch         = $false
            Status               = 'OK'
            ErrorMessage         = $null
        }
    }

    $EncodedAddress = [System.Uri]::EscapeDataString($Address.Trim())
    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }
    $Url = "https://api.mapy.cz/v1/geocode?query=$EncodedAddress&lang=$lang&limit=1&type=regional&apikey=$ApiKey"
    try {
        Write-Verbose "Geokodowanie: '$Address'"
        $Response = Invoke-MapyRestJson -Uri $Url -TimeoutSec 30
        $Results = @($Response.items)
        if ($Results.Count -gt 0) {
            $ResultItem = $Results[0]
            $Position   = $ResultItem.position

            # Extract address components from regionalStructure
            $RegStructure = @($ResultItem.regionalStructure)
            $StreetName   = Get-AddressComponentValue -Components $RegStructure -Types @('regional.street')
            $AddressNum   = Get-AddressComponentValue -Components $RegStructure -Types @('regional.address')
            $City         = Get-AddressComponentValue -Components $RegStructure -Types @('regional.municipality')
            if ([string]::IsNullOrWhiteSpace($City)) {
                $City = Get-AddressComponentValue -Components $RegStructure -Types @('regional.municipality_part')
            }
            $PostalCode = if ($ResultItem.zip) { [string]$ResultItem.zip } else { $null }

            $StreetWithNumber = if ($StreetName -and $AddressNum) { "$StreetName $AddressNum" }
                                elseif ($StreetName) { $StreetName }
                                elseif ($AddressNum) { $AddressNum }
                                else { $null }

            # Build formatted address from name + location
            $FormattedAddress = if ($ResultItem.name -and $ResultItem.location) {
                "$($ResultItem.name), $($ResultItem.location)"
            } elseif ($ResultItem.name) {
                [string]$ResultItem.name
            } elseif ($ResultItem.location) {
                [string]$ResultItem.location
            } else {
                $Address
            }

            # Map Mapy.com entity type to Google-compatible match type
            $EntityType = [string]$ResultItem.type
            $LocationType = switch -Wildcard ($EntityType) {
                'regional.address'          { 'ROOFTOP' }
                'regional.street'           { 'RANGE_INTERPOLATED' }
                'regional.municipality_part' { 'GEOMETRIC_CENTER' }
                'regional.municipality'      { 'APPROXIMATE' }
                'regional.region'            { 'APPROXIMATE' }
                'regional.country'           { 'APPROXIMATE' }
                default                      { 'APPROXIMATE' }
            }

            return [PSCustomObject]@{
                Latitude             = [double]$Position.lat
                Longitude            = [double]$Position.lon
                FormattedAddress     = $FormattedAddress
                UlicaINumer          = $StreetWithNumber
                KodPocztowy          = $PostalCode
                Miasto               = $City
                MatchType            = $LocationType
                PartialMatch         = $false
                Status               = 'OK'
                ErrorMessage         = $null
            }
        }
        else {
            Write-Warning "Geokodowanie nieudane dla '$Address'. Brak wyników."
            return [PSCustomObject]@{
                Latitude             = $null; Longitude = $null; FormattedAddress = $null
                UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
                MatchType            = $null; PartialMatch = $null
                Status               = 'ZERO_RESULTS'
                ErrorMessage         = 'No geocoding results found.'
            }
        }
    }
    catch {
        $Message = $_.Exception.Message
        Write-Warning "Błąd geokodowania '$Address': $Message"
        return [PSCustomObject]@{
            Latitude             = $null; Longitude = $null; FormattedAddress = $null
            UlicaINumer          = $null; KodPocztowy = $null; Miasto = $null
            MatchType            = $null; PartialMatch = $null
            Status               = "EXCEPTION: $Message"
            ErrorMessage         = $Message
        }
    }
}
Set-Item -Path "function:global:Get-AddressCoordinates" -Value (Get-Item "function:Get-AddressCoordinates").ScriptBlock -ErrorAction SilentlyContinue

function Get-MapySuggest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Query,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][int]$Limit = 5
    )
    if ([string]::IsNullOrWhiteSpace($Query) -or $Query.Trim().Length -lt 2) { return @() }

    $encodedQuery = [System.Uri]::EscapeDataString($Query.Trim())
    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }

    # 1. Try Mapy.com suggest endpoint
    $suggestUrl = "https://api.mapy.cz/v1/suggest?query=$encodedQuery&lang=$lang&limit=$Limit&type=regional&apikey=$ApiKey"
    try {
        $resp = Invoke-MapyRestJson -Uri $suggestUrl -TimeoutSec 5
        if ($resp -and $resp.items -and @($resp.items).Count -gt 0) {
            $suggestions = [System.Collections.Generic.List[object]]::new()
            foreach ($it in $resp.items) {
                $name = if ($it.name) { [string]$it.name } else { '' }
                $label = if ($it.label) { [string]$it.label } else { '' }
                $full = if ($name -and $label -and $name -ne $label) { "$name, $label" } elseif ($name) { $name } else { $label }
                $suggestions.Add([PSCustomObject]@{
                    Name      = $name
                    Label     = $label
                    FullText  = $full
                    Latitude  = if ($it.position -and $it.position.lat) { [double]$it.position.lat } else { $null }
                    Longitude = if ($it.position -and $it.position.lon) { [double]$it.position.lon } else { $null }
                })
            }
            return @($suggestions.ToArray())
        }
    }
    catch { }

    # 2. Fallback to geocode endpoint with limit
    $geocodeUrl = "https://api.mapy.cz/v1/geocode?query=$encodedQuery&lang=$lang&limit=$Limit&type=regional&apikey=$ApiKey"
    try {
        $resp = Invoke-MapyRestJson -Uri $geocodeUrl -TimeoutSec 5
        if ($resp -and $resp.items -and @($resp.items).Count -gt 0) {
            $suggestions = [System.Collections.Generic.List[object]]::new()
            foreach ($it in $resp.items) {
                $name = if ($it.name) { [string]$it.name } else { '' }
                $label = if ($it.label) { [string]$it.label } else { '' }
                $loc = if ($it.location) { [string]$it.location } else { '' }
                $full = if ($name -and $loc) { "$name, $loc" } elseif ($name) { $name } else { $label }
                $suggestions.Add([PSCustomObject]@{
                    Name      = $name
                    Label     = $label
                    FullText  = $full
                    Latitude  = if ($it.position -and $it.position.lat) { [double]$it.position.lat } else { $null }
                    Longitude = if ($it.position -and $it.position.lon) { [double]$it.position.lon } else { $null }
                })
            }
            return @($suggestions.ToArray())
        }
    }
    catch { }

    return @()
}
Set-Item -Path "function:global:Get-MapySuggest" -Value (Get-Item "function:Get-MapySuggest").ScriptBlock -ErrorAction SilentlyContinue

function Get-GeocodeStatusDescription {
    [CmdletBinding()]
    param(
        [Parameter()][object]$Geo
    )
    if (-not $Geo) { return 'NOT_PROCESSED' }
    if ($Geo.Status -eq 'OK') {
        if ($Geo.PartialMatch -and $Geo.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER') {
            return "OK (Fallback: Approximate / Partial Match - $($Geo.MatchType))"
        }
        elseif ($Geo.PartialMatch) {
            return "OK (Fallback: Partial Match - $($Geo.MatchType))"
        }
        elseif ($Geo.MatchType -eq 'APPROXIMATE') {
            return 'OK (Fallback: Approximate)'
        }
        elseif ($Geo.MatchType -eq 'GEOMETRIC_CENTER') {
            return 'OK (Fallback: Geometric Center)'
        }
        elseif ($Geo.MatchType -eq 'RANGE_INTERPOLATED') {
            return 'OK (Interpolated)'
        }
        elseif ($Geo.MatchType -eq 'ROOFTOP') {
            return 'OK (Exact - ROOFTOP)'
        }
        elseif ($Geo.MatchType -eq 'COORDINATES') {
            return 'OK (Coordinates)'
        }
        else {
            return "OK ($($Geo.MatchType))"
        }
    }
    elseif ($Geo.Status -eq 'ZERO_RESULTS') {
        return 'ZERO_RESULTS (Address Not Found)'
    }
    else {
        return [string]$Geo.Status
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 3. ROUTE CALCULATION (MAPY.COM ROUTING API)
# ══════════════════════════════════════════════════════════════════════════════

function Get-CarRouteData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][object[]]$IntermediatePoints = @(),
        [Parameter()][ValidateSet('Fastest', 'Shortest')][string]$RouteType = 'Fastest',
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][string]$Units = 'METRIC',
        [Parameter()][switch]$TrafficAware,
        [Parameter()][switch]$AvoidTolls,
        [Parameter()][switch]$AvoidHighways
    )

    # Map route type to Mapy.com routeType enum
    $mapyRouteType = switch ($RouteType) {
        'Shortest' { 'car_short' }
        default {
            if ($TrafficAware) { 'car_fast_traffic' } else { 'car_fast' }
        }
    }

    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }

    # Build base URL — Mapy.com uses lon,lat coordinate order
    $RoutesUrl = "https://api.mapy.cz/v1/routing/route" +
        "?start=$OriginLng,$OriginLat" +
        "&end=$DestLng,$DestLat" +
        "&routeType=$mapyRouteType" +
        "&format=polyline" +
        "&lang=$lang" +
        "&apikey=$ApiKey"

    if ($AvoidTolls) { $RoutesUrl += '&avoidToll=true' }
    if ($AvoidHighways) { $RoutesUrl += '&avoidHighways=true' }

    # Add waypoints (up to 15)
    if ($null -ne $IntermediatePoints -and @($IntermediatePoints).Count -gt 0) {
        foreach ($pt in $IntermediatePoints) {
            if ($null -ne $pt -and $pt.Latitude -and $pt.Longitude) {
                $wpLng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", [double]$pt.Longitude)
                $wpLat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", [double]$pt.Latitude)
                $RoutesUrl += "&waypoints=$wpLng,$wpLat"
            }
        }
    }

    try {
        Write-Verbose "Routing API ($RouteType → $mapyRouteType): ($OriginLat,$OriginLng) -> ($DestLat,$DestLng), Waypoints: $(@($IntermediatePoints).Count)"
        $Response = Invoke-MapyRestJson -Uri $RoutesUrl -TimeoutSec 60

        if (-not $Response -or -not $Response.length) {
            Write-Warning "Routing API nie zwróciło żadnej trasy."
            return [PSCustomObject]@{
                OdlegloscKm     = $null
                CzasMin         = $null
                DurationSeconds = $null
                EncodedPolyline = $null
                RouteType       = $RouteType
                RouteLabels     = @()
                Legs            = @()
                Status          = 'NO_ROUTES'
                ErrorMessage    = 'Mapy.com Routing API did not return any routes.'
            }
        }

        # Parse response
        $DistanceMeters  = [int64]$Response.length
        $DistanceKm      = [math]::Round($DistanceMeters / 1000.0, 2)
        $DurationSec     = [double]$Response.duration
        $DurationMinutes = [math]::Round($DurationSec / 60.0, 0)

        # Geometry — when format=polyline, the geometry field is the encoded polyline string
        $Polyline = $null
        if ($Response.geometry -is [string]) {
            $Polyline = $Response.geometry
        }
        elseif ($Response.geometry -and $Response.geometry.geometry -and $Response.geometry.geometry.coordinates) {
            # GeoJSON fallback — would need conversion to polyline, but we request polyline format
            $Polyline = $null
        }

        # Parse route parts (legs between waypoints)
        $Legs = @()
        if ($Response.parts) {
            foreach ($part in $Response.parts) {
                $legDistMeters = if ($part.length) { [int64]$part.length } else { 0 }
                $legDistKm     = [math]::Round($legDistMeters / 1000.0, 2)
                $legDurSec     = if ($part.duration) { [double]$part.duration } else { 0 }
                $legDurMin     = [math]::Round($legDurSec / 60.0, 1)
                $Legs += [PSCustomObject]@{
                    DistanceMeters  = $legDistMeters
                    DistanceKm      = $legDistKm
                    DurationSeconds = $legDurSec
                    DurationMin     = $legDurMin
                }
            }
        }

        return [PSCustomObject]@{
            OdlegloscKm     = $DistanceKm
            CzasMin         = $DurationMinutes
            DurationSeconds = $DurationSec
            EncodedPolyline = $Polyline
            RouteType       = $RouteType
            RouteLabels     = @()
            Legs            = $Legs
            AvoidTolls      = [bool]$AvoidTolls
            AvoidHighways   = [bool]$AvoidHighways
            Status          = 'OK'
            ErrorMessage    = $null
        }
    }
    catch {
        $ErrorMsg = $_.Exception.Message
        Write-Error "Błąd Routing API ($RouteType): $ErrorMsg"
        return [PSCustomObject]@{
            OdlegloscKm     = $null
            CzasMin         = $null
            DurationSeconds = $null
            EncodedPolyline = $null
            RouteType       = $RouteType
            RouteLabels     = @()
            Legs            = @()
            AvoidTolls      = [bool]$AvoidTolls
            AvoidHighways   = [bool]$AvoidHighways
            Status          = "EXCEPTION: $ErrorMsg"
            ErrorMessage    = $ErrorMsg
        }
    }
}

function Get-MapyComUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Origin,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter()][object[]]$Waypoints = @(),
        [Parameter()][string]$TravelMode = 'driving'
    )
    $OriginEnc = [System.Uri]::EscapeDataString($Origin.Trim())
    $DestEnc = [System.Uri]::EscapeDataString($Destination.Trim())
    $Url = "https://mapy.com/fnc/v1/route?start=$OriginEnc&end=$DestEnc"

    if ($null -ne $Waypoints -and @($Waypoints).Count -gt 0) {
        $WpStrings = [System.Collections.Generic.List[string]]::new()
        foreach ($wp in $Waypoints) {
            if ($wp -is [string] -and -not [string]::IsNullOrWhiteSpace($wp)) {
                $WpStrings.Add($wp.Trim())
            }
            elseif ($wp.Latitude -and $wp.Longitude) {
                $WpStrings.Add("$($wp.Longitude),$($wp.Latitude)")
            }
            elseif ($wp.ZapytanieAdresowe) {
                $WpStrings.Add([string]$wp.ZapytanieAdresowe)
            }
            elseif ($wp.AdresGeokodowany) {
                $WpStrings.Add([string]$wp.AdresGeokodowany)
            }
        }
        if ($WpStrings.Count -gt 0) {
            foreach ($wpStr in $WpStrings) {
                $Url += '&waypoints=' + [System.Uri]::EscapeDataString($wpStr)
            }
        }
    }
    return $Url
}

# ══════════════════════════════════════════════════════════════════════════════
# 4. PNG MAP GENERATION AND RENDERING (MAPY.COM STATIC MAPS API + GDI+ OVERLAY)
# ══════════════════════════════════════════════════════════════════════════════

function Get-WrappedLines {
    param(
        [System.Drawing.Graphics]$G,
        [string]$Text,
        [System.Drawing.Font]$F,
        [float]$MaxW
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return [string[]]@('') }
    if ($G.MeasureString($Text, $F).Width -le $MaxW) { return [string[]]@($Text) }
    $Words = $Text -split '\s+'
    $L1 = ''; $L2 = ''; $On2 = $false
    foreach ($W in $Words) {
        if (-not $On2) {
            $T = if ($L1) { "$L1 $W" } else { $W }
            if ($G.MeasureString($T, $F).Width -le $MaxW) { $L1 = $T }
            else { $On2 = $true; $L2 = $W }
        }
        else {
            $T2 = if ($L2) { "$L2 $W" } else { $W }
            if ($G.MeasureString($T2, $F).Width -le $MaxW) { $L2 = $T2 }
            else {
                if ($L2.Length -gt 3) { $L2 = $L2.Substring(0, $L2.Length - 3) + '...' }
                break
            }
        }
    }
    if ($L2) { return [string[]]@($L1, $L2) } else { return [string[]]@($L1) }
}

function Save-RouteMapPng {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EncodedPolyline,
        [Parameter(Mandatory)][double]$OriginLat,
        [Parameter(Mandatory)][double]$OriginLng,
        [Parameter(Mandatory)][double]$DestLat,
        [Parameter(Mandatory)][double]$DestLng,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter()][int]$Width = 900,
        [Parameter()][int]$Height = 600,
        [Parameter()][object[]]$RoutePoints = @(),
        [Parameter()][Alias('TekstAdresA')][string]$AddressTextA = '',
        [Parameter()][Alias('TekstAdresB')][string]$AddressTextB = '',
        [Parameter()][Alias('TekstOdleglosc')][string]$DistanceText = '',
        [Parameter()][Alias('TekstCzas')][string]$DurationText = '',
        [Parameter()][Alias('TekstNaglowekLewy')][string]$HeaderLeftText = '',
        [Parameter()][Alias('TekstNaglowekPrawy')][string]$HeaderRightText = '',
        [Parameter()][Alias('TekstUmowa')][string]$ContractText = '',
        [Parameter()][Alias('TekstKierunek')][string]$DirectionText = '',
        [Parameter()][Alias('Opis')][string]$Description = '',
        [Parameter()][Alias('DataWygenerowania')][string]$GeneratedDate = '',
        [Parameter()][string]$LanguageCode = 'en',
        [Parameter()][string]$StartRaw = '',
        [Parameter()][string]$StartGeocoded = '',
        [Parameter()][string]$EndRaw = '',
        [Parameter()][string]$EndGeocoded = '',
        [Parameter()][object[]]$WaypointsList = @(),
        [Parameter()][string]$RouteName = '',
        [Parameter()][string]$RouteType = '',
        [Parameter()][object[]]$Legs = @(),
        [Parameter()][object]$OverlayConfig = $null
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls

    # Clamp dimensions to Mapy.com max of 1024
    $Width  = [math]::Min($Width, 1024)
    $Height = [math]::Min($Height, 1024)

    $lang = if ($LanguageCode) { ($LanguageCode -split '[-_]')[0].ToLower() } else { 'en' }

    # Build markers — Mapy.com uses semicolon-separated attributes and lon,lat order
    $originLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", $OriginLng)
    $originLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", $OriginLat)
    $destLngStr   = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", $DestLng)
    $destLatStr   = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", $DestLat)

    $MarkerParams = [System.Collections.Generic.List[string]]::new()

    # Start marker (green, label A)
    $MarkerParams.Add("&markers=" + [System.Uri]::EscapeDataString("color:green;size:large;label:A;$originLngStr,$originLatStr"))

    # Intermediate markers
    if ($null -ne $RoutePoints -and @($RoutePoints).Count -gt 0) {
        $IntermediatesOnly = @()
        if ($RoutePoints.Count -gt 2 -and
            [math]::Abs($RoutePoints[0].Latitude - $OriginLat) -lt 0.0001 -and
            [math]::Abs($RoutePoints[-1].Latitude - $DestLat) -lt 0.0001) {
            $IntermediatesOnly = @($RoutePoints[1..($RoutePoints.Count - 2)])
        }
        else {
            $IntermediatesOnly = @($RoutePoints | Where-Object {
                $null -ne $_.Latitude -and $null -ne $_.Longitude -and
                (-not ([math]::Abs($_.Latitude - $OriginLat) -lt 0.0001 -and [math]::Abs($_.Longitude - $OriginLng) -lt 0.0001)) -and
                (-not ([math]::Abs($_.Latitude - $DestLat) -lt 0.0001 -and [math]::Abs($_.Longitude - $DestLng) -lt 0.0001))
            })
        }

        $idx = 1
        foreach ($pt in $IntermediatesOnly) {
            if ($pt.Latitude -and $pt.Longitude) {
                $lbl = if ($idx -le 9) { [string]$idx }
                       elseif ($idx -le 35) { [string][char](55 + $idx) }
                       else { '' }
                $ptLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", [double]$pt.Longitude)
                $ptLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", [double]$pt.Latitude)
                $spec = if ($lbl) { "color:blue;size:normal;label:$lbl;$ptLngStr,$ptLatStr" }
                        else { "color:blue;size:normal;$ptLngStr,$ptLatStr" }
                $MarkerParams.Add("&markers=" + [System.Uri]::EscapeDataString($spec))
                $idx++
            }
        }
    }

    # End marker (red, label B)
    $MarkerEnd = [System.Uri]::EscapeDataString("color:red;size:large;label:B;$destLngStr,$destLatStr")
    $MarkerParams.Add("&markers=$MarkerEnd")

    # Shapes parameter for polyline
    $shapesParam = if ($EncodedPolyline) {
        "&shapes=" + [System.Uri]::EscapeDataString("color:#0066FF;width:4;path:[enc($EncodedPolyline)]")
    } else {
        ""
    }

    $StaticMapUrl = ("https://api.mapy.cz/v1/static/map" +
        "?width=$Width&height=$Height" +
        "&format=png" +
        "&mapset=basic" +
        "&lang=$lang" +
        $shapesParam +
        ($MarkerParams -join '') +
        "&apikey=$ApiKey")

    try {
        $TargetDir = Split-Path -Parent $OutputPath
        if (-not [string]::IsNullOrWhiteSpace($TargetDir) -and -not (Test-Path $TargetDir)) {
            New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        }

        if ($ApiKey -ne 'OFFLINE_TEST') {
            $wc = [System.Net.WebClient]::new()
            try {
                $wc.DownloadFile($StaticMapUrl, $OutputPath)
            }
            finally {
                $wc.Dispose()
            }
        }

        # Resolve overlay configuration
        if ($OverlayConfig -is [string] -and -not [string]::IsNullOrWhiteSpace($OverlayConfig)) {
            try { $OverlayConfig = $OverlayConfig | ConvertFrom-Json } catch { }
        }
        if (-not $OverlayConfig) {
            $OverlayConfig = [PSCustomObject]@{
                EnableTopOverlay    = $true
                EnableBottomOverlay = $true
                Properties          = [PSCustomObject]@{
                    RouteName     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Top';    Alignment = 'Left';   Order = 1 }
                    RouteType     = [PSCustomObject]@{ Enabled = $true;  Panel = 'Top';    Alignment = 'Right';  Order = 1 }
                    Timestamp     = [PSCustomObject]@{ Enabled = $false; Panel = 'Top';    Alignment = 'Right';  Order = 2 }
                    StartGeocoded = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 1 }
                    EndGeocoded   = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 2 }
                    Distance      = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Left';   Order = 3 }
                    Duration      = [PSCustomObject]@{ Enabled = $true;  Panel = 'Bottom'; Alignment = 'Center'; Order = 3 }
                    Waypoints     = [PSCustomObject]@{ Enabled = $false; Panel = 'Bottom'; Alignment = 'Left';   Order = 4 }
                    PointDistances= [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Alignment = 'Left';   Order = 5 }
                    StartRaw      = [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Alignment = 'Left';   Order = 1 }
                    EndRaw        = [PSCustomObject]@{ Enabled = $false; Panel = 'None';   Alignment = 'Left';   Order = 2 }
                }
            }
        }

        $enableTop = if ($null -ne $OverlayConfig.EnableTopOverlay) { [bool]$OverlayConfig.EnableTopOverlay } else { $true }
        $enableBtm = if ($null -ne $OverlayConfig.EnableBottomOverlay) { [bool]$OverlayConfig.EnableBottomOverlay } else { $true }

        # Check if PointDistances option is enabled
        $showPointDistances = $false
        if ($OverlayConfig) {
            $ovProps = if ($OverlayConfig.Properties) {
                $OverlayConfig.Properties
            } elseif ($OverlayConfig.Items) {
                $OverlayConfig.Items
            } else {
                $null
            }
            if ($ovProps) {
                $pdCfg = if ($ovProps -is [System.Collections.IDictionary]) {
                    $ovProps['PointDistances']
                } else {
                    $ovProps.PointDistances
                }
                if ($pdCfg -and $null -ne $pdCfg.Enabled) {
                    $showPointDistances = [bool]$pdCfg.Enabled
                }
            }
            if (-not $showPointDistances -and $null -ne $OverlayConfig.PointDistances) {
                $showPointDistances = [bool]$OverlayConfig.PointDistances
            }
        }

        # Resolve legs list
        $legParts = [System.Collections.Generic.List[string]]::new()
        $legsToUse = if ($Legs -and @($Legs).Count -gt 0) {
            @($Legs)
        } elseif ($RoutePoints -and @($RoutePoints).Count -gt 1) {
            $ptsWithLegs = @($RoutePoints[1..(@($RoutePoints).Count - 1)] | Where-Object { $null -ne $_.LegDistanceKm })
            if ($ptsWithLegs.Count -gt 0) {
                $ptsWithLegs | ForEach-Object {
                    [PSCustomObject]@{
                        DistanceKm  = $_.LegDistanceKm
                        DurationMin = if ($null -ne $_.LegDurationMin) { $_.LegDurationMin } else { 0 }
                    }
                }
            } else { @() }
        } else { @() }

        if ($legsToUse.Count -gt 0) {
            $totalLegs = $legsToUse.Count
            for ($k = 0; $k -lt $totalLegs; $k++) {
                $fromLbl = if ($k -eq 0) { 'A' } else { "$k" }
                $toLbl   = if ($k -eq ($totalLegs - 1)) { 'B' } else { "$($k + 1)" }
                $lObj    = $legsToUse[$k]
                $lDist   = if ($null -ne $lObj.DistanceKm) { $lObj.DistanceKm } else { 0 }
                $lDur    = if ($null -ne $lObj.DurationMin) { $lObj.DurationMin } else { 0 }
                $durPart = if ($lDur -gt 0) { " ($lDur min)" } else { "" }
                $legParts.Add("${fromLbl} → ${toLbl}: ${lDist} km${durPart}")
            }
        }
        $legsPrefix = switch ($lang) { 'de' { 'Etappen: ' } 'pl' { 'Odcinki: ' } default { 'Legs: ' } }
        $legsText   = ($legParts -join '  |  ')

        # Resolve data values
        $addrStartGeo = if ($StartGeocoded) { $StartGeocoded } elseif ($AddressTextA) { $AddressTextA } else { '' }
        $addrStartRaw = if ($StartRaw) { $StartRaw } else { '' }
        $addrEndGeo   = if ($EndGeocoded) { $EndGeocoded } elseif ($AddressTextB) { $AddressTextB } else { '' }
        $addrEndRaw   = if ($EndRaw) { $EndRaw } else { '' }

        # If PointDistances option is enabled, append distance in parentheses to Destination address
        if ($showPointDistances) {
            $endLegDist = $null
            if ($legsToUse -and $legsToUse.Count -gt 0 -and $null -ne $legsToUse[-1].DistanceKm) {
                $endLegDist = $legsToUse[-1].DistanceKm
            } elseif ($RoutePoints -and @($RoutePoints).Count -gt 1 -and $null -ne $RoutePoints[-1].LegDistanceKm) {
                $endLegDist = $RoutePoints[-1].LegDistanceKm
            }
            if ($null -ne $endLegDist -and $endLegDist -gt 0) {
                $endSuffix = " (+${endLegDist} km)"
                if ($addrEndGeo -and -not $addrEndGeo.EndsWith($endSuffix)) { $addrEndGeo += $endSuffix }
                if ($addrEndRaw -and -not $addrEndRaw.EndsWith($endSuffix)) { $addrEndRaw += $endSuffix }
            }
        }

        $nameVal = if ($RouteName) { $RouteName } elseif ($HeaderLeftText) { $HeaderLeftText } elseif ($Description) { $Description.Trim() } elseif ($ContractText) { $ContractText } else { '' }

        $typeVal = if ($RouteType) { $RouteType } elseif ($HeaderRightText) { $HeaderRightText } elseif ($DirectionText) { $DirectionText } else { '' }
        if ($typeVal -match '^(?:Type|Typ|Art):\s*(.+)$' -or $typeVal -match '^(Shortest|Fastest|Eco|Najkr[oó]tsza|Najszybsza|Eko|K[uü]rzeste|Schnellste)$') {
            $rawVal = if ($Matches[1]) { $Matches[1].Trim() } else { $Matches[0].Trim() }
            $normVal = if ($rawVal -match '(?i)short|kr[oó]t|k[uü]rz') { 'Shortest' }
                       elseif ($rawVal -match '(?i)eco|eko|fuel') { 'Eco' }
                       elseif ($rawVal -match '(?i)fast|szyb|schnell') { 'Fastest' }
                       else { $rawVal }
            $tPrefix = switch ($lang) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $tName = switch ($lang) {
                'de' { if ($normVal -eq 'Fastest') { 'Schnellste' } elseif ($normVal -eq 'Shortest') { 'Kürzeste' } elseif ($normVal -eq 'Eco') { 'Eco' } else { $normVal } }
                'pl' { if ($normVal -eq 'Fastest') { 'Najszybsza' } elseif ($normVal -eq 'Shortest') { 'Najkrótsza' } elseif ($normVal -eq 'Eco') { 'Eko' } else { $normVal } }
                default { $normVal }
            }
            $typeVal = "$tPrefix$tName"
        }

        $distPrefix = switch ($lang) { 'de' { 'Gesamt: ' } 'pl' { 'Razem: ' } default { 'Total: ' } }
        $distVal = if ($DistanceText) { $DistanceText } else { '' }

        $durVal = if ($DurationText) {
            if ($DurationText -match '^\(.*\)$') { $DurationText } else { "($DurationText)" }
        } else { '' }

        $dateVal = if ($GeneratedDate) { $GeneratedDate } else { (Get-Date -Format 'yyyy-MM-dd  HH:mm') }

        $wpItems = [System.Collections.Generic.List[PSCustomObject]]::new()
        $rawWpList = if ($WaypointsList -and @($WaypointsList).Count -gt 0) {
            $WaypointsList
        } elseif ($RoutePoints -and @($RoutePoints).Count -gt 2) {
            @($RoutePoints[1..($RoutePoints.Count - 2)])
        } else { @() }

        $wIdx = 1
        foreach ($w in $rawWpList) {
            $wText = if ($w -is [string]) { $w }
                     elseif ($w.FormattedAddress) { $w.FormattedAddress }
                     elseif ($w.Address) { $w.Address }
                     elseif ($w.GeocodedAddress) { $w.GeocodedAddress }
                     elseif ($w.OriginalAddress) { $w.OriginalAddress }
                     else { '' }
            if (-not [string]::IsNullOrWhiteSpace($wText)) {
                if ($showPointDistances) {
                    $wLegDist = $null
                    $legIdx = $wIdx - 1
                    if ($legsToUse -and $legIdx -lt $legsToUse.Count -and $null -ne $legsToUse[$legIdx].DistanceKm) {
                        $wLegDist = $legsToUse[$legIdx].DistanceKm
                    } elseif ($w -isnot [string] -and $null -ne $w.LegDistanceKm) {
                        $wLegDist = $w.LegDistanceKm
                    }
                    if ($null -ne $wLegDist -and $wLegDist -gt 0) {
                        $wSuffix = " (+${wLegDist} km)"
                        if (-not $wText.EndsWith($wSuffix)) {
                            $wText += $wSuffix
                        }
                    }
                }
                $wpItems.Add([PSCustomObject]@{
                    Index = $wIdx
                    Badge = "${wIdx}: "
                    Text  = $wText
                })
                $wIdx++
            }
        }

        # Build active property items map
        $propDataMap = @{
            'StartGeocoded'  = @{ Id='StartGeocoded';  Kind='address';        Badge='A: '; BadgeColor='Green'; Text=$addrStartGeo }
            'StartRaw'       = @{ Id='StartRaw';       Kind='address';        Badge='A: '; BadgeColor='Green'; Text=$addrStartRaw }
            'EndGeocoded'    = @{ Id='EndGeocoded';    Kind='address';        Badge='B: '; BadgeColor='Red';   Text=$addrEndGeo }
            'EndRaw'         = @{ Id='EndRaw';         Kind='address';        Badge='B: '; BadgeColor='Red';   Text=$addrEndRaw }
            'Distance'       = @{ Id='Distance';       Kind='stat';           Prefix=$distPrefix; Value=$distVal }
            'Duration'       = @{ Id='Duration';       Kind='stat';           Value=$durVal }
            'Timestamp'      = @{ Id='Timestamp';      Kind='date';           Text=$dateVal }
            'RouteName'      = @{ Id='RouteName';      Kind='title';          Text=$nameVal }
            'RouteType'      = @{ Id='RouteType';      Kind='type';           Text=$typeVal }
            'Waypoints'      = @{ Id='Waypoints';      Kind='waypoints';      Items=$wpItems }
            'PointDistances' = @{ Id='PointDistances'; Kind='pointdistances'; Badge=$legsPrefix; Text=$legsText }
        }

        $topItems = [System.Collections.Generic.List[PSCustomObject]]::new()
        $btmItems = [System.Collections.Generic.List[PSCustomObject]]::new()

        $overlayProps = if ($OverlayConfig.Properties) {
            $OverlayConfig.Properties
        } elseif ($OverlayConfig.Items) {
            $OverlayConfig.Items
        } else {
            $null
        }

        if ($overlayProps) {
            $propNames = if ($overlayProps -is [System.Collections.IDictionary]) {
                $overlayProps.Keys
            } else {
                $overlayProps.PSObject.Properties.Name
            }
            foreach ($pName in $propNames) {
                $iCfg = if ($overlayProps -is [System.Collections.IDictionary]) {
                    $overlayProps[$pName]
                } else {
                    $overlayProps.$pName
                }
                if (-not $iCfg) { continue }
                $pEnabled = if ($null -ne $iCfg.Enabled) { [bool]$iCfg.Enabled } else { $true }
                $pPanel   = if ($iCfg.Panel) { [string]$iCfg.Panel } else { 'None' }
                $pAlign   = if ($iCfg.Alignment) { [string]$iCfg.Alignment } elseif ($iCfg.Align) { [string]$iCfg.Align } else { 'Left' }
                $pOrder   = if ($iCfg.Order) { [int]$iCfg.Order } else { 1 }

                if (-not $pEnabled -or $pPanel -eq 'None') { continue }
                if (-not $propDataMap.ContainsKey($pName)) { continue }

                $pData = $propDataMap[$pName]
                $hasContent = $false
                if ($pData.Kind -eq 'waypoints') {
                    $hasContent = ($pData.Items -and $pData.Items.Count -gt 0)
                } elseif ($pData.Kind -eq 'stat') {
                    $hasContent = (-not [string]::IsNullOrWhiteSpace($pData.Value))
                } else {
                    $hasContent = (-not [string]::IsNullOrWhiteSpace($pData.Text))
                }
                if (-not $hasContent) { continue }

                $itemObj = [PSCustomObject]@{
                    Id         = $pName
                    Kind       = $pData.Kind
                    Badge      = $pData.Badge
                    BadgeColor = $pData.BadgeColor
                    Text       = $pData.Text
                    Prefix     = $pData.Prefix
                    Value      = $pData.Value
                    Items      = $pData.Items
                    Panel      = $pPanel
                    Align      = $pAlign
                    Order      = $pOrder
                }

                if ($pPanel -eq 'Top' -and $enableTop) {
                    $topItems.Add($itemObj)
                } elseif ($pPanel -eq 'Bottom' -and $enableBtm) {
                    $btmItems.Add($itemObj)
                }
            }
        }

        $MaTopOverlay = ($enableTop -and $topItems.Count -gt 0)
        $MaBottomOverlay = ($enableBtm -and $btmItems.Count -gt 0)

        if ($MaTopOverlay -or $MaBottomOverlay) {
            try {
                Add-Type -AssemblyName System.Drawing

                $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
                $MemStream = [System.IO.MemoryStream]::new($FileBytes)
                $BitmapSrc = [System.Drawing.Bitmap]::new($MemStream)

                $ActualW = $BitmapSrc.Width
                $ActualH = $BitmapSrc.Height

                # Fonts definition
                $FontTopTitle = [System.Drawing.Font]::new('Segoe UI', 10.0, [System.Drawing.FontStyle]::Bold)
                $FontTopType  = [System.Drawing.Font]::new('Segoe UI', 10.0, [System.Drawing.FontStyle]::Bold)
                $FontBadge    = [System.Drawing.Font]::new('Segoe UI', 9.0,  [System.Drawing.FontStyle]::Bold)
                $FontAddr     = [System.Drawing.Font]::new('Segoe UI', 9.5,  [System.Drawing.FontStyle]::Regular)
                $FontDistLbl  = [System.Drawing.Font]::new('Segoe UI', 9.0,  [System.Drawing.FontStyle]::Bold)
                $FontDist     = [System.Drawing.Font]::new('Segoe UI', 12.0, [System.Drawing.FontStyle]::Bold)
                $FontDate     = [System.Drawing.Font]::new('Segoe UI', 8.5,  [System.Drawing.FontStyle]::Regular)

                $PadX  = 14
                $LineH = 20

                # Pre-measurement Graphics
                $dummyBmp = [System.Drawing.Bitmap]::new(1, 1)
                $measGfx  = [System.Drawing.Graphics]::FromImage($dummyBmp)

                # Helper scriptblock to group items by Order
                $BuildRows = {
                    param($items)
                    $orders = @($items | Select-Object -ExpandProperty Order -Unique | Sort-Object)
                    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($ord in $orders) {
                        $rowItems = @($items | Where-Object { $_.Order -eq $ord })
                        $left   = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $center = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $right  = [System.Collections.Generic.List[PSCustomObject]]::new()
                        foreach ($it in $rowItems) {
                            if ($it.Align -eq 'Right') { $right.Add($it) }
                            elseif ($it.Align -eq 'Center') { $center.Add($it) }
                            else { $left.Add($it) }
                        }
                        $rows.Add([PSCustomObject]@{
                            Order  = $ord
                            Left   = $left
                            Center = $center
                            Right  = $right
                            Height = 20
                        })
                    }
                    return $rows.ToArray()
                }

                $topRows = @(if ($MaTopOverlay) { & $BuildRows $topItems } else { @() })
                $btmRows = @(if ($MaBottomOverlay) { & $BuildRows $btmItems } else { @() })

                # Measure row heights
                $MeasureRows = {
                    param($rows, $availWidth)
                    foreach ($row in @($rows)) {
                        $maxH = 20
                        $allItems = @($row.Left) + @($row.Center) + @($row.Right)
                        foreach ($it in $allItems) {
                            if ($it.Kind -eq 'address') {
                                $badgeSz = $measGfx.MeasureString($it.Badge, $FontBadge)
                                $addrW = [float]($availWidth - $badgeSz.Width)
                                $lines = @(Get-WrappedLines -G $measGfx -Text $it.Text -F $FontAddr -MaxW $addrW)
                                $it | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $lines -Force
                                $h = [math]::Max(1, $lines.Count) * $LineH
                                if ($h -gt $maxH) { $maxH = $h }
                            }
                            elseif ($it.Kind -eq 'waypoints') {
                                $totalWpH = 0
                                foreach ($wp in $it.Items) {
                                    $bSz = $measGfx.MeasureString($wp.Badge, $FontBadge)
                                    $wpMaxW = [float]($availWidth - $bSz.Width)
                                    $wpLines = @(Get-WrappedLines -G $measGfx -Text $wp.Text -F $FontAddr -MaxW $wpMaxW)
                                    $wp | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $wpLines -Force
                                    $totalWpH += [math]::Max(1, $wpLines.Count) * $LineH
                                }
                                if ($totalWpH -gt $maxH) { $maxH = $totalWpH }
                            }
                            elseif ($it.Kind -eq 'pointdistances') {
                                $badgeSz = $measGfx.MeasureString($it.Badge, $FontBadge)
                                $legMaxW = [float]($availWidth - $badgeSz.Width)
                                $lines   = @(Get-WrappedLines -G $measGfx -Text $it.Text -F $FontAddr -MaxW $legMaxW)
                                $it | Add-Member -NotePropertyName 'WrappedLines' -NotePropertyValue $lines -Force
                                $h = [math]::Max(1, $lines.Count) * $LineH
                                if ($h -gt $maxH) { $maxH = $h }
                            }
                            elseif ($it.Kind -eq 'stat') {
                                if ($maxH -lt 24) { $maxH = 24 }
                            }
                            elseif ($it.Kind -in @('title', 'type')) {
                                if ($maxH -lt 22) { $maxH = 22 }
                            }
                        }
                        $row.Height = $maxH
                    }
                }

                $availContentW = [float]($ActualW - ($PadX * 2))
                & $MeasureRows $topRows $availContentW
                & $MeasureRows $btmRows $availContentW

                $measGfx.Dispose()
                $dummyBmp.Dispose()

                # Calculate banner heights
                $TopPad = 8; $TopBotPad = 8; $TopRowSpacing = 4
                $TopBarH = 0
                if ($MaTopOverlay -and @($topRows).Count -gt 0) {
                    $sumTopH = (@($topRows) | Measure-Object -Property Height -Sum).Sum
                    if (-not $sumTopH) { $sumTopH = 20 }
                    $TopBarH = [int]($TopPad + $sumTopH + ((@($topRows).Count - 1) * $TopRowSpacing) + $TopBotPad)
                    if ($TopBarH -lt 38) { $TopBarH = 38 }
                }

                $BtmPadTop = 10; $BtmPadBot = 10; $BtmRowSpacing = 6
                $BtmBarH = 0
                if ($MaBottomOverlay -and @($btmRows).Count -gt 0) {
                    $sumBtmH = (@($btmRows) | Measure-Object -Property Height -Sum).Sum
                    if (-not $sumBtmH) { $sumBtmH = 20 }
                    $BtmBarH = [int]($BtmPadTop + $sumBtmH + ((@($btmRows).Count - 1) * $BtmRowSpacing) + $BtmPadBot)
                }

                $FinalW = $ActualW
                $FinalH = $ActualH + $TopBarH + $BtmBarH

                $Bitmap = [System.Drawing.Bitmap]::new($FinalW, $FinalH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
                $Graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

                # 1. Background fill
                $BrushBg = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 15, 23, 42))
                $Graphics.FillRectangle($BrushBg, 0, 0, $FinalW, $FinalH)

                # 2. Draw map image in the middle
                $Graphics.DrawImage($BitmapSrc, 0, $TopBarH, $ActualW, $ActualH)

                # 3. Brushes & Pens
                $PenSep      = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 51, 65, 85), 1.5)
                $BrushWhite  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 248, 250, 252))
                $BrushYellow = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 250, 204, 21))
                $BrushCyan   = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 56, 189, 248))
                $BrushGreen  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 16, 185, 129))
                $BrushRed    = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 239, 68, 68))
                $BrushMuted  = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 148, 163, 184))

                # Helper scriptblock to measure width of an item
                $MeasureItemWidth = {
                    param($it)
                    if ($it.Kind -eq 'address') {
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $tSz = $Graphics.MeasureString($it.Text, $FontAddr)
                        return ($bSz.Width + $tSz.Width)
                    }
                    elseif ($it.Kind -eq 'stat') {
                        $w = 0
                        if ($it.Prefix) { $w += $Graphics.MeasureString($it.Prefix, $FontDistLbl).Width }
                        if ($it.Value)  { $w += $Graphics.MeasureString($it.Value, $FontDist).Width }
                        return $w
                    }
                    elseif ($it.Kind -eq 'title') {
                        return $Graphics.MeasureString($it.Text, $FontTopTitle).Width
                    }
                    elseif ($it.Kind -eq 'type') {
                        return $Graphics.MeasureString($it.Text, $FontTopType).Width
                    }
                    elseif ($it.Kind -eq 'date') {
                        return $Graphics.MeasureString($it.Text, $FontDate).Width
                    }
                    elseif ($it.Kind -eq 'pointdistances') {
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $tSz = $Graphics.MeasureString($it.Text, $FontAddr)
                        return ($bSz.Width + $tSz.Width)
                    }
                    elseif ($it.Kind -eq 'waypoints') {
                        return 200
                    }
                    return 0
                }

                # Helper scriptblock to draw an item at specified coordinates
                $DrawItem = {
                    param($it, [float]$x, [float]$y)
                    if ($it.Kind -eq 'address') {
                        $badgeBrush = if ($it.BadgeColor -eq 'Red') { $BrushRed } else { $BrushGreen }
                        $Graphics.DrawString($it.Badge, $FontBadge, $badgeBrush, $x, $y)
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $curLineY = $y
                        $lines = if ($it.WrappedLines) { $it.WrappedLines } else { @($it.Text) }
                        foreach ($line in $lines) {
                            $Graphics.DrawString($line, $FontAddr, $BrushWhite, ($x + $bSz.Width), $curLineY)
                            $curLineY += [float]$LineH
                        }
                    }
                    elseif ($it.Kind -eq 'waypoints') {
                        $wpY = $y
                        foreach ($wp in $it.Items) {
                            $Graphics.DrawString($wp.Badge, $FontBadge, $BrushCyan, $x, $wpY)
                            $bSz = $Graphics.MeasureString($wp.Badge, $FontBadge)
                            $lines = if ($wp.WrappedLines) { $wp.WrappedLines } else { @($wp.Text) }
                            foreach ($line in $lines) {
                                $Graphics.DrawString($line, $FontAddr, $BrushWhite, ($x + $bSz.Width), $wpY)
                                $wpY += [float]$LineH
                            }
                        }
                    }
                    elseif ($it.Kind -eq 'pointdistances') {
                        $Graphics.DrawString($it.Badge, $FontBadge, $BrushCyan, $x, $y)
                        $bSz = $Graphics.MeasureString($it.Badge, $FontBadge)
                        $curLineY = $y
                        $lines = if ($it.WrappedLines) { $it.WrappedLines } else { @($it.Text) }
                        foreach ($line in $lines) {
                            $Graphics.DrawString($line, $FontAddr, $BrushYellow, ($x + $bSz.Width), $curLineY)
                            $curLineY += [float]$LineH
                        }
                    }
                    elseif ($it.Kind -eq 'stat') {
                        $statX = $x
                        if ($it.Prefix) {
                            $pSz = $Graphics.MeasureString($it.Prefix, $FontDistLbl)
                            $Graphics.DrawString($it.Prefix, $FontDistLbl, $BrushCyan, $statX, ($y + 2))
                            $statX += $pSz.Width
                        }
                        if ($it.Value) {
                            $Graphics.DrawString($it.Value, $FontDist, $BrushYellow, $statX, $y)
                        }
                    }
                    elseif ($it.Kind -eq 'title') {
                        $Graphics.DrawString($it.Text, $FontTopTitle, $BrushWhite, $x, $y)
                    }
                    elseif ($it.Kind -eq 'type') {
                        $Graphics.DrawString($it.Text, $FontTopType, $BrushYellow, $x, $y)
                    }
                    elseif ($it.Kind -eq 'date') {
                        $Graphics.DrawString($it.Text, $FontDate, $BrushMuted, $x, ($y + 3))
                    }
                }

                # Helper scriptblock to render a banner's rows
                $RenderBannerRows = {
                    param($rows, [float]$startY, [float]$spacing)
                    $curY = $startY
                    foreach ($row in $rows) {
                        $leftX = [float]$PadX

                        # 1. Left items
                        foreach ($it in $row.Left) {
                            & $DrawItem $it $leftX $curY
                            $w = & $MeasureItemWidth $it
                            $leftX += [float]($w + 14)
                        }

                        # 2. Right items
                        $totalRightW = 0
                        foreach ($it in $row.Right) {
                            $totalRightW += [float]((& $MeasureItemWidth $it) + 12)
                        }
                        $rightX = [float]($FinalW - $PadX - $totalRightW + 12)
                        foreach ($it in $row.Right) {
                            & $DrawItem $it $rightX $curY
                            $w = & $MeasureItemWidth $it
                            $rightX += [float]($w + 12)
                        }

                        # 3. Center items
                        $totalCenterW = 0
                        foreach ($it in $row.Center) {
                            $totalCenterW += [float]((& $MeasureItemWidth $it) + 12)
                        }
                        $centerX = [float][math]::Max($leftX + 10, ($FinalW - $totalCenterW + 12) / 2)
                        foreach ($it in $row.Center) {
                            & $DrawItem $it $centerX $curY
                            $w = & $MeasureItemWidth $it
                            $centerX += [float]($w + 12)
                        }

                        $curY += [float]($row.Height + $spacing)
                    }
                }

                # 4. Draw Top Header Banner
                if ($MaTopOverlay -and $TopBarH -gt 0 -and @($topRows).Count -gt 0) {
                    $Graphics.DrawLine($PenSep, 0, $TopBarH, $FinalW, $TopBarH)
                    $topStartY = [float]$TopPad
                    if (@($topRows).Count -eq 1) {
                        $topStartY = [float][math]::Max(6, ($TopBarH - $topRows[0].Height) / 2)
                    }
                    & $RenderBannerRows $topRows $topStartY $TopRowSpacing
                }

                # 5. Draw Bottom Footer Banner
                if ($MaBottomOverlay -and $BtmBarH -gt 0 -and @($btmRows).Count -gt 0) {
                    $BtmBarY = $TopBarH + $ActualH
                    $Graphics.DrawLine($PenSep, 0, $BtmBarY, $FinalW, $BtmBarY)
                    $btmStartY = [float]($BtmBarY + $BtmPadTop)
                    & $RenderBannerRows $btmRows $btmStartY $BtmRowSpacing
                }

                # Dispose GDI+ objects
                $PenSep.Dispose()
                $BrushBg.Dispose(); $BrushWhite.Dispose(); $BrushYellow.Dispose()
                $BrushCyan.Dispose(); $BrushGreen.Dispose(); $BrushRed.Dispose(); $BrushMuted.Dispose()
                $FontTopTitle.Dispose(); $FontTopType.Dispose()
                $FontBadge.Dispose(); $FontAddr.Dispose(); $FontDistLbl.Dispose(); $FontDist.Dispose(); $FontDate.Dispose()
                $Graphics.Dispose()

                $Bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $Bitmap.Dispose()
                $BitmapSrc.Dispose()
                $MemStream.Dispose()
            }
            catch {
                if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
                    Write-AppLog "GDI+ overlay rendering warning: $($_.Exception.Message)" "WARN"
                }
            }
        }
        return $true
    }
    catch {
        return $false
    }
}

function Find-MatchingPropertyName {
    param(
        [Parameter(Mandatory)][string[]]$AvailableProperties,
        [Parameter(Mandatory)][string[]]$Patterns
    )
    foreach ($pattern in $Patterns) {
        $found = $AvailableProperties | Where-Object {
            $null -ne $_ -and $_.Trim() -match $pattern
        } | Select-Object -First 1
        if ($found) { return $found }
    }
    return $null
}

function Import-RouteDataFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][string]$Delimiter = ''
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Plik wejściowy nie istnieje: $Path"
    }

    $Extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $RawRows = $null
    $Format = $null

    switch ($Extension) {
        { $_ -in '.xlsx', '.xls' } {
            $Format = 'Excel'
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                throw "Wymagany moduł ImportExcel nie jest zainstalowany. Zainstaluj go poleceniem: Install-Module -Name ImportExcel -Scope CurrentUser"
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            $RawRows = @(Import-Excel -Path $Path)
        }
        { $_ -in '.csv', '.tsv', '.txt' } {
            $Format = 'CSV'
            $FirstLine = Get-Content -LiteralPath $Path -TotalCount 1
            $UsedDelimiter = if (-not [string]::IsNullOrWhiteSpace($Delimiter)) {
                $Delimiter
            }
            elseif ($Extension -eq '.tsv' -or $FirstLine -match "`t") {
                "`t"
            }
            elseif ($FirstLine -match ';') {
                ';'
            }
            else {
                ','
            }
            $RawRows = @(Import-Csv -LiteralPath $Path -Delimiter $UsedDelimiter)
        }
        '.json' {
            $Format = 'JSON'
            $Content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
            $ParsedJson = $Content | ConvertFrom-Json
            if ($ParsedJson -is [System.Collections.IEnumerable] -and -not ($ParsedJson -is [string])) {
                $RawRows = @($ParsedJson)
            }
            elseif ($ParsedJson.PSObject.Properties.Name -contains 'Routes') {
                $RawRows = @($ParsedJson.Routes)
            }
            elseif ($ParsedJson.PSObject.Properties.Name -contains 'Stops') {
                $RawRows = @($ParsedJson.Stops)
            }
            else {
                $RawRows = @($ParsedJson)
            }
        }
        default {
            throw "Nieobsługiwany format pliku: $Extension. Obsługiwane rozszerzenia: .xlsx, .xls, .csv, .tsv, .json"
        }
    }

    if ($null -eq $RawRows -or $RawRows.Count -eq 0) {
        return [PSCustomObject]@{
            Mode       = 'Empty'
            Routes     = @()
            RawData    = @()
            FilePath   = $Path
            Format     = $Format
            TotalCount = 0
        }
    }

    # Analiza kolumn pierwszego wiersza
    $PropNames = @($RawRows[0].PSObject.Properties.Name)

    # Column matching patterns
    $ColRouteNamePatterns = @(
        '(?i)^(nazwa[\s_]*trasy|route[\s_]*name|routename|nazwatrasy)$',
        '(?i)(nazwa[\s_]*trasy|route[\s_]*name)',
        '(?i)^(name|nazwa)$',
        '(?i)^(umowa|contract|opis|description|tytul)$',
        '(?i)numer.*umowy',
        '(?i)^(id|nr)$'
    )

    # 1. Check if file represents a sequential waypoint sequence for single or multiple routes (SequentialStops)
    # e.g., columns: LP / Kolejnosc + Adres / Lokalizacja
    $ColSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(lp|l\.p\.|kolejnosc|stop|sequence|order|nr)$')
    $ColAddrSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(adres|address|lokalizacja|punkt|miejsce)$', 'lokalizacja.*(odbioru|dowozu)', 'adres.*(odbioru|dowozu)')
    $ColCitySeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(miejscowosc|miasto|city|town)$')
    $ColRouteNameSeq = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns $ColRouteNamePatterns

    $IsSequentialStops = ($ColSeq -and ($ColAddrSeq -or $ColCitySeq) -and -not (Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @('^(start|origin|adres.*a)$')))

    if ($IsSequentialStops) {
        # Group sequential stops into routes (supports single route or multiple grouped routes)
        $routeGroups = [System.Collections.Generic.List[System.Collections.Generic.List[object]]]::new()
        $currGroup = [System.Collections.Generic.List[object]]::new()
        $activeGrpName = ''

        foreach ($row in $RawRows) {
            $rSeq = if ($ColSeq) { [string]$row.$ColSeq } else { '' }
            $rName = if ($ColRouteNameSeq) { [string]$row.$ColRouteNameSeq } else { '' }
            $hasRowName = -not [string]::IsNullOrWhiteSpace($rName)

            $isNewRoute = $false
            if ($currGroup.Count -gt 0) {
                # New route if route name explicitly changes to another non-empty name
                if ($hasRowName -and -not [string]::IsNullOrWhiteSpace($activeGrpName) -and ($rName.Trim() -ne $activeGrpName)) {
                    $isNewRoute = $true
                }
                # Or if sequence index resets back to 1 after accumulating at least 2 stops
                elseif ($ColSeq -and ($rSeq.Trim() -eq '1') -and ($currGroup.Count -ge 2)) {
                    $isNewRoute = $true
                }
            }

            if ($isNewRoute) {
                $routeGroups.Add($currGroup)
                $currGroup = [System.Collections.Generic.List[object]]::new()
                $activeGrpName = ''
            }

            $currGroup.Add($row)
            if ($hasRowName -and [string]::IsNullOrWhiteSpace($activeGrpName)) {
                $activeGrpName = $rName.Trim()
            }
        }
        if ($currGroup.Count -gt 0) {
            $routeGroups.Add($currGroup)
        }

        $StopList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $RoutesList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $grpIdx = 1

        foreach ($grpRows in $routeGroups) {
            $orderedRows = @($grpRows)
            # Try sorting by sequence if numeric
            $canSort = $true
            if ($ColSeq) {
                foreach ($r in $orderedRows) {
                    $val = [string]$r.$ColSeq
                    if ($val -notmatch '^\d+$') { $canSort = $false; break }
                }
            } else { $canSort = $false }
            if ($canSort) {
                $orderedRows = @($orderedRows | Sort-Object { [int]($_.$ColSeq) })
            }

            # Find route name for this group: first non-empty name in group (first row is checked first)
            $grpRouteName = ''
            if ($ColRouteNameSeq) {
                foreach ($r in $orderedRows) {
                    $cand = [string]$r.$ColRouteNameSeq
                    if (-not [string]::IsNullOrWhiteSpace($cand)) {
                        $grpRouteName = $cand.Trim()
                        break
                    }
                }
            }
            if ([string]::IsNullOrWhiteSpace($grpRouteName)) {
                $grpRouteName = if ($routeGroups.Count -gt 1) { "Trasa $grpIdx" } else { "Multi-point Route ($($orderedRows.Count) stops)" }
            }

            $grpStops = [System.Collections.Generic.List[PSCustomObject]]::new()
            for ($s = 0; $s -lt $orderedRows.Count; $s++) {
                $st = $orderedRows[$s]
                $addr = if ($ColAddrSeq) { [string]$st.$ColAddrSeq } else { '' }
                $city = if ($ColCitySeq) { [string]$st.$ColCitySeq } else { '' }
                $fullAddr = if ($addr -and $city) { "$addr, $city" } elseif ($addr) { $addr } else { $city }

                # Route name visible in first row of the multipoint route
                $stopRouteName = if ($s -eq 0) {
                    $grpRouteName
                } elseif ($ColRouteNameSeq -and -not [string]::IsNullOrWhiteSpace($st.$ColRouteNameSeq)) {
                    ([string]$st.$ColRouteNameSeq).Trim()
                } else {
                    ''
                }

                $stopObj = [PSCustomObject]@{
                    Sequence  = if ($ColSeq) { [string]$st.$ColSeq } else { [string]($s + 1) }
                    RouteId   = [string]$grpIdx
                    RouteName = $stopRouteName
                    Address   = $fullAddr.Trim()
                    Raw       = $st
                }
                $grpStops.Add($stopObj)
                $StopList.Add($stopObj)
            }

            if ($grpStops.Count -ge 2) {
                $StartPoint = $grpStops[0].Address
                $EndPoint = $grpStops[$grpStops.Count - 1].Address
                $Waypoints = if ($grpStops.Count -gt 2) { @($grpStops[1..($grpStops.Count - 2)] | ForEach-Object { $_.Address }) } else { @() }
                $RouteObj = [PSCustomObject]@{
                    Id          = [string]$grpIdx
                    Name        = $grpRouteName
                    Start       = $StartPoint
                    End         = $EndPoint
                    Waypoints   = $Waypoints
                    RouteType   = 'Fastest'
                    OriginalRow = $orderedRows
                }
                $RoutesList.Add($RouteObj)
            }
            $grpIdx++
        }

        return [PSCustomObject]@{
            Mode       = 'SequentialStops'
            Stops      = $StopList
            Routes     = $RoutesList
            RawData    = $RawRows
            FilePath   = $Path
            Format     = $Format
            TotalCount = $StopList.Count
        }
    }

    # 2. RouteList Mode (each row is a separate route)
    $ColStart = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(start|origin|startpoint|poczat.*|poczatek|od|from|dom)$',
        '(?i)adres.*a|^a$',
        '(?i)punkt.*(poczat|start)'
    )
    $ColEnd = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(end|dest|destination|endpoint|koniec.*|konic.*|cel|meta|do|to|szkola)$',
        '(?i)adres.*b|^b$',
        '(?i)punkt.*(konic|koniec|docel|cel)'
    )
    $ColWaypoints = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(waypoints|waypoint|posredn.*|punkty.*posredn.*|przystank.*|via|stops|praca)$',
        '(?i)posrednie'
    )
    $ColName = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns $ColRouteNamePatterns
    $ColRouteType = Find-MatchingPropertyName -AvailableProperties $PropNames -Patterns @(
        '(?i)^(routetype|typ|typtrasy|tryb|mode|optimization)$'
    )

    $NormalizedRoutes = [System.Collections.Generic.List[PSCustomObject]]::new()
    $idx = 1

    foreach ($row in $RawRows) {
        $startVal = if ($ColStart) { [string]$row.$ColStart } else { '' }
        $endVal   = if ($ColEnd) { [string]$row.$ColEnd } else { '' }
        if ([string]::IsNullOrWhiteSpace($startVal) -or [string]::IsNullOrWhiteSpace($endVal)) {
            continue
        }

        $nameVal = if ($ColName -and -not [string]::IsNullOrWhiteSpace($row.$ColName)) {
            ([string]$row.$ColName).Trim()
        } else {
            "Trasa $idx"
        }
        $typeVal = if ($ColRouteType) { [string]$row.$ColRouteType } else { $null }

        # Normalizacja RouteType
        if ($typeVal -match '(?i)eco|fuel|paliw|eko') { $typeVal = 'Eco' }
        elseif ($typeVal -match '(?i)short|krot|krót') { $typeVal = 'Shortest' }
        elseif ($typeVal -match '(?i)fast|szyb') { $typeVal = 'Fastest' }
        else { $typeVal = $null }

        # Obsługa punktów pośrednich
        $waypointsList = [System.Collections.Generic.List[string]]::new()
        if ($ColWaypoints -and -not [string]::IsNullOrWhiteSpace($row.$ColWaypoints)) {
            $rawWp = $row.$ColWaypoints
            if ($rawWp -is [System.Collections.IEnumerable] -and -not ($rawWp -is [string])) {
                foreach ($item in $rawWp) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$item)) {
                        $waypointsList.Add(([string]$item).Trim())
                    }
                }
            }
            else {
                $splits = ([string]$rawWp) -split '(?<!\\)[|;]'
                foreach ($s in $splits) {
                    $cleaned = $s.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($cleaned)) {
                        $waypointsList.Add($cleaned)
                    }
                }
            }
        }

        $NormalizedRoutes.Add([PSCustomObject]@{
            Id          = [string]$idx
            Name        = $nameVal
            Start       = $startVal.Trim()
            End         = $endVal.Trim()
            Waypoints   = @($waypointsList)
            RouteType   = $typeVal
            OriginalRow = $row
        })
        $idx++
    }

    return [PSCustomObject]@{
        Mode       = 'RouteList'
        Routes     = $NormalizedRoutes
        RawData    = $RawRows
        FilePath   = $Path
        Format     = $Format
        TotalCount = $NormalizedRoutes.Count
        Columns    = [PSCustomObject]@{
            Start     = $ColStart
            End       = $ColEnd
            Waypoints = $ColWaypoints
            Name      = $ColName
            RouteType = $ColRouteType
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 6. UNIVERSAL RESULTS EXPORT (EXCEL, CSV, JSON)
# ══════════════════════════════════════════════════════════════════════════════

function Export-RouteResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter()][ValidateSet('Excel', 'CSV', 'JSON')][string]$Format = 'Excel'
    )

    $TargetDir = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($TargetDir) -and -not (Test-Path $TargetDir)) {
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    }

    # Extract flat summary rows (excluding nested Points array from main sheet/file)
    $RoutesFlat = [System.Collections.Generic.List[PSCustomObject]]::new()
    $PointsFlat = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($r in $Results) {
        $routeId   = if ($null -ne $r.Id) { [string]$r.Id } else { '' }
        $routeName = if ($r.Name) { [string]$r.Name } elseif ($r.Nazwa) { [string]$r.Nazwa } else { "Route $routeId" }
        $startOrig = if ($r.Start_Original) { [string]$r.Start_Original } elseif ($r.Start) { [string]$r.Start } elseif ($r.StartRaw) { [string]$r.StartRaw } else { '' }
        $startGeo  = if ($r.Start_Geocoded) { [string]$r.Start_Geocoded } elseif ($r.StartGeocoded) { [string]$r.StartGeocoded } elseif ($r.StartGeokodowany) { [string]$r.StartGeokodowany } else { '' }
        $startStat = if ($r.Start_Status) { [string]$r.Start_Status } elseif ($r.StartStatus) { [string]$r.StartStatus } else { '' }
        $endOrig   = if ($r.End_Original) { [string]$r.End_Original } elseif ($r.End) { [string]$r.End } elseif ($r.EndRaw) { [string]$r.EndRaw } elseif ($r.Koniec) { [string]$r.Koniec } else { '' }
        $endGeo    = if ($r.End_Geocoded) { [string]$r.End_Geocoded } elseif ($r.EndGeocoded) { [string]$r.EndGeocoded } elseif ($r.KoniecGeokodowany) { [string]$r.KoniecGeokodowany } else { '' }
        $endStat   = if ($r.End_Status) { [string]$r.End_Status } elseif ($r.EndStatus) { [string]$r.EndStatus } else { '' }
        $wpCount   = if ($null -ne $r.WaypointsCount) { [int]$r.WaypointsCount } elseif ($null -ne $r.LiczbaPrzystankow) { [int]$r.LiczbaPrzystankow } else { 0 }
        $rType     = if ($r.RouteType) { [string]$r.RouteType } elseif ($r.TypTrasy) { [string]$r.TypTrasy } else { '' }
        $dist      = if ($null -ne $r.DistanceKm) { $r.DistanceKm } elseif ($null -ne $r.OdlegloscKm) { $r.OdlegloscKm } else { $null }
        $dur       = if ($null -ne $r.DurationMin) { $r.DurationMin } elseif ($null -ne $r.CzasMin) { $r.CzasMin } else { $null }
        $status    = if ($r.Status) { [string]$r.Status } else { '' }
        $map       = if ($r.MapPath) { [string]$r.MapPath } elseif ($r.MapaPath) { [string]$r.MapaPath } else { '' }
        $url       = if ($r.MapyComUrl) { [string]$r.MapyComUrl } elseif ($r.GoogleMapsUrl) { [string]$r.GoogleMapsUrl } else { '' }

        # Build waypoints summary text
        $pts = if ($r.RoutePoints) { $r.RoutePoints } elseif ($r.Points) { $r.Points } else { $null }
        $wpSummaryList = [System.Collections.Generic.List[string]]::new()
        if ($pts -and ($pts -is [System.Collections.IEnumerable])) {
            foreach ($pt in $pts) {
                if ($pt.PointType -like 'Waypoint*') {
                    $ptSummary = "$($pt.PointType): '$($pt.OriginalAddress)'"
                    if ($pt.GeocodedAddress) { $ptSummary += " -> '$($pt.GeocodedAddress)'" }
                    if ($null -ne $pt.LegDistanceKm -and $pt.LegDistanceKm -gt 0) { $ptSummary += " ($($pt.LegDistanceKm) km)" }
                    if ($pt.GeocodeStatus) { $ptSummary += " [$($pt.GeocodeStatus)]" }
                    $wpSummaryList.Add($ptSummary)
                }

                $PointsFlat.Add([PSCustomObject]@{
                    RouteId         = $routeId
                    RouteName       = $routeName
                    PointOrder      = $pt.Order
                    PointType       = $pt.PointType
                    LegDistanceKm   = if ($null -ne $pt.LegDistanceKm) { $pt.LegDistanceKm } else { $null }
                    LegDurationMin  = if ($null -ne $pt.LegDurationMin) { $pt.LegDurationMin } else { $null }
                    OriginalAddress = $pt.OriginalAddress
                    GeocodedAddress = $pt.GeocodedAddress
                    GeocodeStatus   = $pt.GeocodeStatus
                    MatchType       = $pt.MatchType
                    IsFallback      = if ($null -ne $pt.IsFallback) { [bool]$pt.IsFallback } else { $false }
                    Latitude        = $pt.Latitude
                    Longitude       = $pt.Longitude
                })
            }
        }

        $wpSummaryText = $wpSummaryList -join ' | '

        $RoutesFlat.Add([PSCustomObject]@{
            Id               = $routeId
            Name             = $routeName
            Start_Original   = $startOrig
            Start_Geocoded   = $startGeo
            Start_Status     = $startStat
            End_Original     = $endOrig
            End_Geocoded     = $endGeo
            End_Status       = $endStat
            WaypointsCount   = $wpCount
            RouteType        = $rType
            DistanceKm       = $dist
            DurationMin      = $dur
            Status           = $status
            WaypointsSummary = $wpSummaryText
            MapPath          = $map
            MapyComUrl       = $url
            GoogleMapsUrl    = $url
        })
    }

    $csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8BOM' } else { 'UTF8' }

    switch ($Format) {
        'Excel' {
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                Write-Warning "Moduł ImportExcel nie jest zainstalowany. Eksportowanie do CSV zamiast Excel."
                $CsvPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
                $RoutesFlat | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
                if ($PointsFlat.Count -gt 0) {
                    $PtsCsv = [System.IO.Path]::Combine($TargetDir, "$([System.IO.Path]::GetFileNameWithoutExtension($CsvPath))_punkty.csv")
                    $PointsFlat | Export-Csv -LiteralPath $PtsCsv -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
                }
                return $CsvPath
            }
            Import-Module -Name ImportExcel -ErrorAction Stop
            if (Test-Path -LiteralPath $OutputPath) {
                Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
            }
            $RoutesFlat | Export-Excel -Path $OutputPath -WorksheetName 'Trasy' -TableName 'WynikiTras' -AutoSize -AutoFilter -FreezeTopRow
            if ($PointsFlat.Count -gt 0) {
                $PointsFlat | Export-Excel -Path $OutputPath -WorksheetName 'PunktyTrasy' -TableName 'PunktyTrasy' -AutoSize -AutoFilter -FreezeTopRow
            }
            return $OutputPath
        }
        'CSV' {
            $RoutesFlat | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
            if ($PointsFlat.Count -gt 0) {
                $PtsCsv = [System.IO.Path]::Combine($TargetDir, "$([System.IO.Path]::GetFileNameWithoutExtension($OutputPath))_punkty.csv")
                $PointsFlat | Export-Csv -LiteralPath $PtsCsv -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'
            }
            return $OutputPath
        }
        'JSON' {
            $JsonContent = $Results | ConvertTo-Json -Depth 6
            [System.IO.File]::WriteAllText($OutputPath, $JsonContent, [System.Text.UTF8Encoding]::new($true))
            return $OutputPath
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# 6. ENCODED POLYLINE DECODER & GPS EXPORTERS (GPX / KML)
# ══════════════════════════════════════════════════════════════════════════════

if (-not ([System.Management.Automation.PSTypeName]'GoogleMapsPolylineDecoder').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;

public class GoogleMapsPoint {
    public double Latitude { get; set; }
    public double Longitude { get; set; }
    public GoogleMapsPoint(double lat, double lng) {
        Latitude = lat;
        Longitude = lng;
    }
}

public static class GoogleMapsPolylineDecoder {
    public static List<GoogleMapsPoint> Decode(string encoded) {
        var points = new List<GoogleMapsPoint>();
        if (string.IsNullOrEmpty(encoded)) return points;

        int index = 0, len = encoded.Length;
        int lat = 0, lng = 0;

        while (index < len) {
            int b, shift = 0, result = 0;
            do {
                if (index >= len) return points;
                b = encoded[index++] - 63;
                result |= (b & 0x1f) << shift;
                shift += 5;
            } while (b >= 0x20);
            int dlat = ((result & 1) != 0 ? ~(result >> 1) : (result >> 1));
            lat += dlat;

            shift = 0;
            result = 0;
            do {
                if (index >= len) return points;
                b = encoded[index++] - 63;
                result |= (b & 0x1f) << shift;
                shift += 5;
            } while (b >= 0x20);
            int dlng = ((result & 1) != 0 ? ~(result >> 1) : (result >> 1));
            lng += dlng;

            points.Add(new GoogleMapsPoint(lat * 1e-5, lng * 1e-5));
        }
        return points;
    }
}
"@ -ErrorAction SilentlyContinue
}

function ConvertFrom-EncodedPolyline {
    [Alias('ConvertFrom-GoogleEncodedPolyline')]
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$EncodedPolyline)
    if ([string]::IsNullOrWhiteSpace($EncodedPolyline)) { return @() }
    try {
        return [GoogleMapsPolylineDecoder]::Decode($EncodedPolyline)
    }
    catch {
        return @()
    }
}

function Export-RouteGpx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $false)][string]$RouteName = 'Route',
        [Parameter(Mandatory = $true)][string]$EncodedPolyline,
        [Parameter()][object[]]$Waypoints = @(),
        [Parameter()][double]$DistanceKm = 0,
        [Parameter()][int]$DurationMin = 0
    )

    if ([string]::IsNullOrWhiteSpace($RouteName)) { $RouteName = 'Route' }
    $points = ConvertFrom-GoogleEncodedPolyline -EncodedPolyline $EncodedPolyline
    $safeName = [System.Security.SecurityElement]::Escape($RouteName)
    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    $null = $sb.AppendLine('<gpx version="1.1" creator="MapyComRoutes" xmlns="http://www.topografix.com/GPX/1/1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd">')
    $null = $sb.AppendLine("  <metadata>")
    $null = $sb.AppendLine("    <name>$safeName</name>")
    $null = $sb.AppendLine("    <desc>Distance: $DistanceKm km, Duration: $DurationMin min</desc>")
    $null = $sb.AppendLine("    <time>$timestamp</time>")
    $null = $sb.AppendLine("  </metadata>")

    # Export waypoints if provided
    if ($Waypoints -and @($Waypoints).Count -gt 0) {
        foreach ($wp in $Waypoints) {
            if ($wp.Latitude -and $wp.Longitude) {
                $lat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Latitude)
                $lng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Longitude)
                $wName = [System.Security.SecurityElement]::Escape($(if ($wp.Name) { [string]$wp.Name } elseif ($wp.Role) { [string]$wp.Role } else { 'Waypoint' }))
                $wDesc = [System.Security.SecurityElement]::Escape($(if ($wp.Address) { [string]$wp.Address } else { '' }))
                $null = $sb.AppendLine("  <wpt lat=`"$lat`" lon=`"$lng`">")
                $null = $sb.AppendLine("    <name>$wName</name>")
                if ($wDesc) { $null = $sb.AppendLine("    <desc>$wDesc</desc>") }
                $null = $sb.AppendLine("  </wpt>")
            }
        }
    }

    # Route track
    $null = $sb.AppendLine("  <trk>")
    $null = $sb.AppendLine("    <name>$safeName</name>")
    $null = $sb.AppendLine("    <trkseg>")
    foreach ($pt in $points) {
        $lat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$pt.Latitude)
        $lng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$pt.Longitude)
        $null = $sb.AppendLine("      <trkpt lat=`"$lat`" lon=`"$lng`" />")
    }
    $null = $sb.AppendLine("    </trkseg>")
    $null = $sb.AppendLine("  </trk>")
    $null = $sb.AppendLine("</gpx>")

    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($true))
    return $OutputPath
}

function Export-RouteKml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $false)][string]$RouteName = 'Route',
        [Parameter(Mandatory = $true)][string]$EncodedPolyline,
        [Parameter()][object[]]$Waypoints = @(),
        [Parameter()][double]$DistanceKm = 0,
        [Parameter()][int]$DurationMin = 0
    )

    if ([string]::IsNullOrWhiteSpace($RouteName)) { $RouteName = 'Route' }
    $points = ConvertFrom-GoogleEncodedPolyline -EncodedPolyline $EncodedPolyline
    $safeName = [System.Security.SecurityElement]::Escape($RouteName)

    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    $null = $sb.AppendLine('<kml xmlns="http://www.opengis.net/kml/2.2">')
    $null = $sb.AppendLine("  <Document>")
    $null = $sb.AppendLine("    <name>$safeName</name>")
    $null = $sb.AppendLine("    <description>Distance: $DistanceKm km, Duration: $DurationMin min</description>")
    $null = $sb.AppendLine('    <Style id="routeStyle">')
    $null = $sb.AppendLine('      <LineStyle>')
    $null = $sb.AppendLine('        <color>ff0066ff</color>')
    $null = $sb.AppendLine('        <width>4</width>')
    $null = $sb.AppendLine('      </LineStyle>')
    $null = $sb.AppendLine('    </Style>')

    # Add stop markers
    if ($Waypoints -and @($Waypoints).Count -gt 0) {
        foreach ($wp in $Waypoints) {
            if ($wp.Latitude -and $wp.Longitude) {
                $lat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Latitude)
                $lng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Longitude)
                $wName = [System.Security.SecurityElement]::Escape($(if ($wp.Name) { [string]$wp.Name } elseif ($wp.Role) { [string]$wp.Role } else { 'Stop' }))
                $wDesc = [System.Security.SecurityElement]::Escape($(if ($wp.Address) { [string]$wp.Address } else { '' }))
                $null = $sb.AppendLine("    <Placemark>")
                $null = $sb.AppendLine("      <name>$wName</name>")
                if ($wDesc) { $null = $sb.AppendLine("      <description>$wDesc</description>") }
                $null = $sb.AppendLine("      <Point><coordinates>$lng,$lat,0</coordinates></Point>")
                $null = $sb.AppendLine("    </Placemark>")
            }
        }
    }

    # Add linestring track
    $coordStrings = [System.Collections.Generic.List[string]]::new()
    foreach ($pt in $points) {
        $lat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$pt.Latitude)
        $lng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$pt.Longitude)
        $coordStrings.Add("$lng,$lat,0")
    }

    $null = $sb.AppendLine("    <Placemark>")
    $null = $sb.AppendLine("      <name>$safeName (Route Track)</name>")
    $null = $sb.AppendLine("      <styleUrl>#routeStyle</styleUrl>")
    $null = $sb.AppendLine("      <LineString>")
    $null = $sb.AppendLine("        <tessellate>1</tessellate>")
    $null = $sb.AppendLine("        <coordinates>$($coordStrings -join ' ')</coordinates>")
    $null = $sb.AppendLine("      </LineString>")
    $null = $sb.AppendLine("    </Placemark>")
    $null = $sb.AppendLine("  </Document>")
    $null = $sb.AppendLine("</kml>")

    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($true))
    return $OutputPath
}

# ══════════════════════════════════════════════════════════════════════════════
# 7. GOOGLE MAPS PLATFORM API USAGE & COST ESTIMATION HELPER
# ══════════════════════════════════════════════════════════════════════════════

function Get-EstimatedApiCost {
    [CmdletBinding()]
    param(
        [Parameter()][int]$GeocodingCalls = 0,
        [Parameter()][int]$RoutesBasicCalls = 0,
        [Parameter()][int]$RoutesAdvancedCalls = 0,
        [Parameter()][int]$StaticMapsCalls = 0,
        [Parameter()][ValidateSet('USD', 'EUR', 'PLN')][string]$Currency = 'USD',
        [Parameter()][double]$UsdToEur = 0.92,
        [Parameter()][double]$UsdToPln = 3.95,
        [Parameter()][double]$MonthlyFreeCreditUsd = 200.0
    )

    # Google Maps Platform Standard Tier Pricing:
    # - Geocoding API: $0.005 per call ($5.00 / 1000)
    # - Routes API (ComputeRoutes Basic): $0.005 per call ($5.00 / 1000)
    # - Routes API (ComputeRoutes Advanced/Traffic): $0.010 per call ($10.00 / 1000)
    # - Maps Static API: $0.002 per call ($2.00 / 1000)
    $costGeo     = $GeocodingCalls * 0.005
    $costBasic   = $RoutesBasicCalls * 0.005
    $costAdv     = $RoutesAdvancedCalls * 0.010
    $costStatic  = $StaticMapsCalls * 0.002

    $totalCostUsd = [math]::Round($costGeo + $costBasic + $costAdv + $costStatic, 4)
    $freeRemainingUsd = [math]::Max(0.0, [math]::Round($MonthlyFreeCreditUsd - $totalCostUsd, 2))

    $rate = switch ($Currency) {
        'EUR' { $UsdToEur }
        'PLN' { $UsdToPln }
        default { 1.0 }
    }
    $totalCostLocal = [math]::Round($totalCostUsd * $rate, 2)
    $symbol = switch ($Currency) {
        'EUR' { '€' }
        'PLN' { 'zł' }
        default { '$' }
    }

    return [PSCustomObject]@{
        TotalCalls          = ($GeocodingCalls + $RoutesBasicCalls + $RoutesAdvancedCalls + $StaticMapsCalls)
        GeocodingCalls      = $GeocodingCalls
        RoutesBasicCalls    = $RoutesBasicCalls
        RoutesAdvancedCalls = $RoutesAdvancedCalls
        StaticMapsCalls     = $StaticMapsCalls
        CostUsd             = $totalCostUsd
        CostLocal           = $totalCostLocal
        Currency            = $Currency
        CurrencySymbol      = $symbol
        FormattedCost       = "$totalCostLocal $symbol"
        FreeTierRemaining   = "$freeRemainingUsd $"
        IsWithinFreeCredit  = ($totalCostUsd -le $MonthlyFreeCreditUsd)
    }
}

