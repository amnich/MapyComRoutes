#Requires -Version 5.1
<#
.SYNOPSIS
    Mapy.com Routes & Map Generator — Background Runspace Workers Subsystem.
.DESCRIPTION
    Provides thread-safe background execution for manual route calculations,
    batch dataset processing, and pre-batch geocode validation using isolated MTA runspaces.
.NOTES
    Encoding: UTF-8 with BOM
    Compatibility: Windows PowerShell 5.1 and PowerShell 7+
#>

#region 1. Isolated Runspace Worker Factory

<#
.SYNOPSIS
    Creates and configures an isolated background PowerShell runspace worker.
.DESCRIPTION
    Builds an InitialSessionState pre-populated with core routing, geocoding, and crypto
    functions from the caller session. Configures a dedicated Multi-Threaded Apartment (MTA)
    runspace with UseNewThread policy so long-running API operations do not block the
    Single-Threaded Apartment (STA) WPF user interface.
.PARAMETER ScriptBlock
    The PowerShell ScriptBlock to execute inside the background runspace.
.OUTPUTS
    [System.Management.Automation.PowerShell] An open, configured PowerShell pipeline instance.
.EXAMPLE
    $worker = New-WorkerPowerShell -ScriptBlock $script:ManualCalcAsync
    $asyncHandle = $worker.BeginInvoke()
#>
function New-WorkerPowerShell {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    # Construct default InitialSessionState with core .NET and PowerShell built-ins
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
    # Export required business logic functions from host session into the isolated runspace state
    Get-ChildItem function: | Where-Object {
        $_.Name -in @(
            'Protect-SecretString', 'Unprotect-SecretString', 'Test-GoogleApiKey', 'Test-MapyApiKey',
            'Invoke-MapyRestJson',
            'Get-AddressComponentValue', 'Get-AddressCoordinates', 'Get-GeocodeStatusDescription',
            'Get-CarRouteData', 'Get-GoogleMapsUrl', 'Get-MapyComUrl', 'Get-WrappedLines', 'Save-RouteMapPng',
            'Find-MatchingPropertyName', 'Import-RouteDataFile', 'Export-RouteResults',
            'ConvertFrom-GoogleEncodedPolyline', 'ConvertFrom-EncodedPolyline', 'Export-RouteGpx', 'Export-RouteKml',
            'Get-EstimatedApiCost'
        )
    } | ForEach-Object {
        try {
            # Register function definition in runspace session state
            $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($_.Name, $_.Definition))
        }
        catch {
            # Ignore duplicate entries or unexportable built-ins
        }
    }

    # Allocate MTA runspace to allow concurrent background network I/O
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = [System.Threading.ApartmentState]::MTA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $rs.Open()

    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    
    # Discard pipeline return of AddScript to prevent unwanted objects polluting caller results
    $null = $ps.AddScript($ScriptBlock.ToString())
    return $ps
}

#endregion 1. Isolated Runspace Worker Factory

#region 2. Manual Route Calculation Async Worker

<#
.SYNOPSIS
    Asynchronous worker scriptblock executing single manual route calculation.
.DESCRIPTION
    Runs inside an MTA runspace. Performs geocoding of origin, destination, and intermediate waypoints,
    invokes the Mapy.com Routing REST API, renders the static route overview PNG map image,
    and returns a structured result payload with accurate API usage telemetry.
#>
$script:ManualCalcAsync = {
    param(
        $start, $end, $waypoints, $routeType, $emission, $trafficAware,
        $name, $apiKey, $outDir, $logFile, $languageCode = 'en',
        $overlayConfigJson = '',
        $avoidTolls = $false, $avoidHighways = $false, $avoidFerries = $false
    )

    # Enforce modern TLS security protocols inside isolated runspace thread
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    # API call counters for cost and usage tracking
    $geoCount = 0
    $routesCount = 0
    $staticCount = 0

    # Thread-safe log append helper targeting persistent log file
    $wlog = {
        param($msg, $lvl = 'INFO')
        if ($logFile) {
            $t = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
            try { 
                [System.IO.File]::AppendAllText($logFile, "[$t] [$lvl] [ManualWorker] $msg`r`n", [System.Text.UTF8Encoding]::new($true)) 
            } catch { }
        }
    }

    try {
        # 1. Geocode origin address
        & $wlog "Geocoding origin: '$start'..." "INFO"
        $geoStart = Get-AddressCoordinates -Address $start -ApiKey $apiKey -LanguageCode $languageCode
        $geoCount++
        if ($geoStart.Status -ne 'OK') {
            & $wlog "Origin geocoding error: $($geoStart.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Origin geocoding error: $($geoStart.Status)" }
        }
        & $wlog "Origin OK: $($geoStart.FormattedAddress) ($($geoStart.Latitude), $($geoStart.Longitude))" "INFO"

        # 2. Geocode destination address
        & $wlog "Geocoding destination: '$end'..." "INFO"
        $geoEnd = Get-AddressCoordinates -Address $end -ApiKey $apiKey -LanguageCode $languageCode
        $geoCount++
        if ($geoEnd.Status -ne 'OK') {
            & $wlog "Destination geocoding error: $($geoEnd.Status)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Destination geocoding error: $($geoEnd.Status)" }
        }
        & $wlog "Destination OK: $($geoEnd.FormattedAddress) ($($geoEnd.Latitude), $($geoEnd.Longitude))" "INFO"

        # 3. Geocode intermediate waypoints
        $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($waypoints) {
            foreach ($w in $waypoints) {
                if (-not [string]::IsNullOrWhiteSpace($w)) {
                    & $wlog "Geocoding waypoint: '$w'..." "INFO"
                    $g = Get-AddressCoordinates -Address $w -ApiKey $apiKey -LanguageCode $languageCode
                    $geoCount++
                    if ($g.Status -eq 'OK') {
                        $geoWp.Add($g)
                        & $wlog "Waypoint OK: $($g.FormattedAddress)" "INFO"
                    }
                    else {
                        & $wlog "Waypoint geocoding error '$w': $($g.Status)" "WARN"
                    }
                }
            }
        }

        # 4. Invoke Mapy.com Routing REST API
        & $wlog "Querying Mapy.com Routing API (Type: $routeType, Engine: $emission, AvoidTolls: $avoidTolls, AvoidHighways: $avoidHighways, AvoidFerries: $avoidFerries)..." "INFO"
        $trasa = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
            -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
            -IntermediatePoints $geoWp -RouteType $routeType  `
            -ApiKey $apiKey -LanguageCode $languageCode -TrafficAware:$trafficAware `
            -AvoidTolls:$avoidTolls -AvoidHighways:$avoidHighways 
        $routesCount++

        if ($trasa.Status -ne 'OK') {
            & $wlog "Routes API error: $($trasa.Status). $($trasa.ErrorMessage)" "WARN"
            return [PSCustomObject]@{ Success = $false; Error = "Routes API error: $($trasa.Status). $($trasa.ErrorMessage)" }
        }
        & $wlog "Routes API route found: $($trasa.OdlegloscKm) km, $($trasa.CzasMin) min" "INFO"

        # 5. Attach individual leg metrics to waypoints
        if ($trasa.Legs -and $trasa.Legs.Count -gt 0) {
            for ($w = 0; $w -lt $geoWp.Count; $w++) {
                if ($w -lt $trasa.Legs.Count) {
                    $geoWp[$w] | Add-Member -NotePropertyName 'LegDistanceKm' -NotePropertyValue $trasa.Legs[$w].DistanceKm -Force
                    $geoWp[$w] | Add-Member -NotePropertyName 'LegDurationMin' -NotePropertyValue $trasa.Legs[$w].DurationMin -Force
                }
            }
        }

        # 6. Generate interactive web URL
        $gUrl = Get-MapyComUrl -Origin "$($geoStart.Latitude),$($geoStart.Longitude)" `
            -Destination "$($geoEnd.Latitude),$($geoEnd.Longitude)" `
            -Waypoints $geoWp

        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $safeName = ($name -replace '[\\/:*?"<>|]', '_').Trim()
        $mapFileName = "${ts}_manual_route_${safeName}.png"
        $mapPath = Join-Path $outDir $mapFileName

        $allPts = [System.Collections.Generic.List[PSCustomObject]]::new()
        $allPts.Add($geoStart)
        foreach ($pt in $geoWp) { $allPts.Add($pt) }
        $allPts.Add($geoEnd)

        # 7. Render high-resolution static PNG map image
        & $wlog "Rendering static map image: $mapPath..." "INFO"
        $hdrTypePrefix = switch ($languageCode) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
        $hdrTypeName = switch ($languageCode) {
            'de' { if ($routeType -eq 'Fastest') { 'Schnellste' } elseif ($routeType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
            'pl' { if ($routeType -eq 'Fastest') { 'Najszybsza' } elseif ($routeType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
            default { $routeType }
        }
        $headerRightText = "$hdrTypePrefix$hdrTypeName"

        $saved = Save-RouteMapPng -EncodedPolyline $trasa.EncodedPolyline `
            -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
            -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
            -RoutePoints $allPts -OutputPath $mapPath -ApiKey $apiKey `
            -Width 900 -Height 600 `
            -AddressTextA $geoStart.FormattedAddress -AddressTextB $geoEnd.FormattedAddress `
            -DistanceText "$($trasa.OdlegloscKm) km" -DurationText "$($trasa.CzasMin) min" `
            -HeaderLeftText $name -HeaderRightText $headerRightText `
            -LanguageCode $languageCode `
            -StartRaw $start -StartGeocoded $geoStart.FormattedAddress `
            -EndRaw $end -EndGeocoded $geoEnd.FormattedAddress `
            -WaypointsList $geoWp -RouteName $name -RouteType $headerRightText `
            -Legs $trasa.Legs `
            -OverlayConfig $overlayConfigJson
        $staticCount++

        & $wlog "Map rendering complete. Saved: $saved" "INFO"
        $resolvedMapPath = $(if ($saved) { $mapPath } else { $null })

        # Return full calculation package to UI dispatcher
        return [PSCustomObject]@{
            Success         = $true
            DistanceKm      = $trasa.OdlegloscKm
            DurationMin     = $trasa.CzasMin
            RouteType       = $routeType
            EncodedPolyline = $trasa.EncodedPolyline
            MapyComUrl      = $gUrl
            GoogleMapsUrl   = $gUrl
            MapPath         = $resolvedMapPath
            OriginLat       = $geoStart.Latitude
            OriginLng       = $geoStart.Longitude
            OriginAddress   = $geoStart.FormattedAddress
            DestLat         = $geoEnd.Latitude
            DestLng         = $geoEnd.Longitude
            DestAddress     = $geoEnd.FormattedAddress
            Waypoints       = $geoWp
            AvoidTolls      = $avoidTolls
            AvoidHighways   = $avoidHighways
            AvoidFerries    = $avoidFerries
            ApiUsage        = [PSCustomObject]@{
                Geocoding  = $geoCount
                Routes     = $routesCount
                StaticMaps = $staticCount
            }
            Error           = $null
        }
    }
    catch {
        $errFull = $_.Exception.ToString()
        & $wlog "Worker thread exception: $errFull" "ERROR"
        return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message }
    }
}

#endregion 2. Manual Route Calculation Async Worker

#region 3. Batch Route Processing Async Worker

<#
.SYNOPSIS
    Asynchronous worker scriptblock executing batch route calculations.
.DESCRIPTION
    Iterates through a list of imported route records, geocoding start/end/waypoints,
    calling the routing engine, rendering map images, and reporting progress and status
    back to the main GUI thread via synchronized state and log queue.
#>
$script:BatchCalcAsync = {
    param(
        $routes, $apiKey, $outDir, $defaultRouteType, $syncState, $logFile,
        $languageCode = 'en', $overlayConfigJson = '',
        $defaultAvoidTolls = $false, $defaultAvoidHighways = $false, $defaultAvoidFerries = $false
    )

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $geoCount = 0
    $routesCount = 0
    $staticCount = 0

    # Thread-safe log dispatching into synchronized queue for UI display + disk file
    $wlog = {
        param($msg, $lvl = 'INFO')
        if ($syncState.LogQueue) {
            $syncState.LogQueue.Enqueue([PSCustomObject]@{ Level = $lvl; Message = $msg })
        }
        if ($logFile) {
            $t = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
            try { 
                [System.IO.File]::AppendAllText($logFile, "[$t] [$lvl] [BatchWorker] $msg`r`n", [System.Text.UTF8Encoding]::new($true)) 
            } catch { }
        }
    }

    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $total = $routes.Count

    for ($i = 0; $i -lt $total; $i++) {
        # Check cancellation flag before initiating next network request
        if ($syncState.CancelRequested) {
            & $wlog "Batch processing stopped by user at route $($i + 1)/$total." "WARN"
            break
        }
        $r = $routes[$i]
        $syncState.CurrentIndex = ($i + 1)

        # Resolve route optimization and avoid flags (record override > batch default)
        # Note: Mapy.com only supports 'Fastest' and 'Shortest' routing profiles (Eco is unsupported)
        $rType = if ($defaultRouteType -and $defaultRouteType -in 'Fastest', 'Shortest') { $defaultRouteType }
        elseif ($r.RouteType -and $r.RouteType -in 'Fastest', 'Shortest') { $r.RouteType }
        else { 'Fastest' }

        $avoidT = if ($null -ne $r.AvoidTolls) { [bool]$r.AvoidTolls } else { $defaultAvoidTolls }
        $avoidH = if ($null -ne $r.AvoidHighways) { [bool]$r.AvoidHighways } else { $defaultAvoidHighways }
        $avoidF = if ($null -ne $r.AvoidFerries) { [bool]$r.AvoidFerries } else { $defaultAvoidFerries }

        $routeName = if ($r.Name) { $r.Name } else { "Route $($i + 1)" }

        & $wlog "Route $($i + 1)/$($total): Processing '$($r.Start)' -> '$($r.End)' (Type: $rType)..." "INFO"

        try {
            # Geocode Start Address
            $geoStart = Get-AddressCoordinates -Address $r.Start -ApiKey $apiKey -LanguageCode $languageCode
            $geoCount++
            $startStatus = Get-GeocodeStatusDescription -Geo $geoStart
            $isStartFallback = if ($geoStart -and ($geoStart.PartialMatch -or $geoStart.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

            $routePointsList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $routePointsList.Add([PSCustomObject]@{
                Order           = 1
                PointType       = 'Start'
                LegDistanceKm   = 0
                LegDurationMin  = 0
                OriginalAddress = $r.Start
                GeocodedAddress = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                GeocodeStatus   = $startStatus
                MatchType       = if ($geoStart) { $geoStart.MatchType } else { 'NOT_FOUND' }
                PartialMatch    = if ($geoStart) { [bool]$geoStart.PartialMatch } else { $false }
                IsFallback      = $isStartFallback
                Latitude        = if ($geoStart) { $geoStart.Latitude } else { $null }
                Longitude       = if ($geoStart) { $geoStart.Longitude } else { $null }
            })

            # Geocode End Address
            $geoEnd = Get-AddressCoordinates -Address $r.End -ApiKey $apiKey -LanguageCode $languageCode
            $geoCount++
            $endStatus = Get-GeocodeStatusDescription -Geo $geoEnd
            $isEndFallback = if ($geoEnd -and ($geoEnd.PartialMatch -or $geoEnd.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }

            # Geocode Waypoints
            $geoWp = [System.Collections.Generic.List[PSCustomObject]]::new()
            $wpOrder = 2
            if ($r.Waypoints -and @($r.Waypoints).Count -gt 0) {
                foreach ($wpAddr in @($r.Waypoints)) {
                    if (-not [string]::IsNullOrWhiteSpace($wpAddr)) {
                        $gw = Get-AddressCoordinates -Address $wpAddr -ApiKey $apiKey -LanguageCode $languageCode
                        $geoCount++
                        $gwStatus = Get-GeocodeStatusDescription -Geo $gw
                        $isGwFallback = if ($gw -and ($gw.PartialMatch -or $gw.MatchType -in 'APPROXIMATE', 'GEOMETRIC_CENTER')) { $true } else { $false }
                        $routePointsList.Add([PSCustomObject]@{
                            Order           = $wpOrder
                            PointType       = "Waypoint $($wpOrder - 1)"
                            LegDistanceKm   = $null
                            LegDurationMin  = $null
                            OriginalAddress = $wpAddr
                            GeocodedAddress = if ($gw) { $gw.FormattedAddress } else { $null }
                            GeocodeStatus   = $gwStatus
                            MatchType       = if ($gw) { $gw.MatchType } else { 'NOT_FOUND' }
                            PartialMatch    = if ($gw) { [bool]$gw.PartialMatch } else { $false }
                            IsFallback      = $isGwFallback
                            Latitude        = if ($gw) { $gw.Latitude } else { $null }
                            Longitude       = if ($gw) { $gw.Longitude } else { $null }
                        })
                        if ($gw -and $gw.Status -eq 'OK') { $geoWp.Add($gw) }
                        $wpOrder++
                    }
                }
            }

            $routePointsList.Add([PSCustomObject]@{
                Order           = $wpOrder
                PointType       = 'End'
                LegDistanceKm   = $null
                LegDurationMin  = $null
                OriginalAddress = $r.End
                GeocodedAddress = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                GeocodeStatus   = $endStatus
                MatchType       = if ($geoEnd) { $geoEnd.MatchType } else { 'NOT_FOUND' }
                PartialMatch    = if ($geoEnd) { [bool]$geoEnd.PartialMatch } else { $false }
                IsFallback      = $isEndFallback
                Latitude        = if ($geoEnd) { $geoEnd.Latitude } else { $null }
                Longitude       = if ($geoEnd) { $geoEnd.Longitude } else { $null }
            })

            # Check if origin or destination geocoding failed
            if (-not $geoStart -or $geoStart.Status -ne 'OK' -or -not $geoEnd -or $geoEnd.Status -ne 'OK') {
                & $wlog "Route $($i + 1): Geocoding failed (Start: $($geoStart.Status), End: $($geoEnd.Status)). Skipping route." "WARN"
                $results.Add([PSCustomObject]@{
                    Id               = if ($r.Id) { $r.Id } else { ($i + 1) }
                    Name             = $routeName
                    Start_Original   = $r.Start
                    Start            = $r.Start
                    Start_Geocoded   = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    StartGeocoded    = if ($geoStart) { $geoStart.FormattedAddress } else { $null }
                    Start_Status     = $startStatus
                    StartStatus      = $startStatus
                    End_Original     = $r.End
                    End              = $r.End
                    End_Geocoded     = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    EndGeocoded      = if ($geoEnd) { $geoEnd.FormattedAddress } else { $null }
                    End_Status       = $endStatus
                    EndStatus        = $endStatus
                    WaypointsCount   = $geoWp.Count
                    RouteType        = $rType
                    DistanceKm       = $null
                    DurationMin      = $null
                    Status           = "Geocode Error"
                    MapPath          = $null
                    EncodedPolyline  = $null
                    GoogleMapsUrl    = $null
                    MapyComUrl       = $null
                    RoutePoints      = $routePointsList
                })
                continue
            }

            # Calculate route geometry and metrics via Mapy.com API
            $routeData = Get-CarRouteData -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -IntermediatePoints $geoWp -RouteType $rType `
                -ApiKey $apiKey -LanguageCode $languageCode `
                -AvoidTolls:$avoidT -AvoidHighways:$avoidH 
            $routesCount++

            if ($routeData.Status -ne 'OK') {
                & $wlog "Route $($i + 1): Route calculation failed ($($routeData.Status))." "WARN"
                $results.Add([PSCustomObject]@{
                    Id               = if ($r.Id) { $r.Id } else { ($i + 1) }
                    Name             = $routeName
                    Start_Original   = $r.Start
                    Start            = $r.Start
                    Start_Geocoded   = $geoStart.FormattedAddress
                    StartGeocoded    = $geoStart.FormattedAddress
                    Start_Status     = $startStatus
                    StartStatus      = $startStatus
                    End_Original     = $r.End
                    End              = $r.End
                    End_Geocoded     = $geoEnd.FormattedAddress
                    EndGeocoded      = $geoEnd.FormattedAddress
                    End_Status       = $endStatus
                    EndStatus        = $endStatus
                    WaypointsCount   = $geoWp.Count
                    RouteType        = $rType
                    DistanceKm       = $null
                    DurationMin      = $null
                    Status           = "Route Error ($($routeData.Status))"
                    MapPath          = $null
                    EncodedPolyline  = $null
                    GoogleMapsUrl    = $null
                    MapyComUrl       = $null
                    RoutePoints      = $routePointsList
                })
                continue
            }

            # Update route points and waypoints with leg distances/durations
            if ($routeData.Legs -and $routeData.Legs.Count -gt 0) {
                for ($p = 0; $p -lt $routePointsList.Count; $p++) {
                    if ($p -eq 0) {
                        $routePointsList[$p].LegDistanceKm = 0
                        $routePointsList[$p].LegDurationMin = 0
                    }
                    elseif (($p - 1) -lt $routeData.Legs.Count) {
                        $leg = $routeData.Legs[$p - 1]
                        $routePointsList[$p].LegDistanceKm = $leg.DistanceKm
                        $routePointsList[$p].LegDurationMin = $leg.DurationMin
                    }
                }
                for ($w = 0; $w -lt $geoWp.Count; $w++) {
                    if ($w -lt $routeData.Legs.Count) {
                        $geoWp[$w] | Add-Member -NotePropertyName 'LegDistanceKm' -NotePropertyValue $routeData.Legs[$w].DistanceKm -Force
                        $geoWp[$w] | Add-Member -NotePropertyName 'LegDurationMin' -NotePropertyValue $routeData.Legs[$w].DurationMin -Force
                    }
                }
            }

            # Render route overview static map image
            $cleanName = ($routeName -replace '[\\/:*?"<>|]', '_').Trim().Trim('.') -replace '\s+', ' '
            if ([string]::IsNullOrWhiteSpace($cleanName)) { $cleanName = "route_$($i + 1)" }
            $mapFileName = "${ts}_route_$($i + 1)_${cleanName}.png"
            $mapPath = Join-Path $outDir $mapFileName

            $allPts = [System.Collections.Generic.List[PSCustomObject]]::new()
            $allPts.Add($geoStart)
            foreach ($pt in $geoWp) { $allPts.Add($pt) }
            $allPts.Add($geoEnd)

            $hdrTypePrefix = switch ($languageCode) { 'de' { 'Typ: ' } 'pl' { 'Typ: ' } default { 'Type: ' } }
            $hdrTypeName = switch ($languageCode) {
                'de' { if ($rType -eq 'Fastest') { 'Schnellste' } elseif ($rType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                'pl' { if ($rType -eq 'Fastest') { 'Najszybsza' } elseif ($rType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                default { $rType }
            }
            $headerRightText = "$hdrTypePrefix$hdrTypeName"

            $saved = Save-RouteMapPng -EncodedPolyline $routeData.EncodedPolyline `
                -OriginLat $geoStart.Latitude -OriginLng $geoStart.Longitude `
                -DestLat $geoEnd.Latitude -DestLng $geoEnd.Longitude `
                -RoutePoints $allPts -OutputPath $mapPath -ApiKey $apiKey `
                -Width 900 -Height 600 `
                -AddressTextA $geoStart.FormattedAddress -AddressTextB $geoEnd.FormattedAddress `
                -DistanceText "$($routeData.OdlegloscKm) km" -DurationText "$($routeData.CzasMin) min" `
                -HeaderLeftText $routeName -HeaderRightText $headerRightText `
                -LanguageCode $languageCode `
                -StartRaw $r.Start -StartGeocoded $geoStart.FormattedAddress `
                -EndRaw $r.End -EndGeocoded $geoEnd.FormattedAddress `
                -WaypointsList $geoWp -RouteName $routeName -RouteType $headerRightText `
                -Legs $routeData.Legs `
                -OverlayConfig $overlayConfigJson
            $staticCount++

            & $wlog "Route $($i + 1)/$total OK: $($routeData.OdlegloscKm) km, $($routeData.CzasMin) min" "OK"

            $gUrl = Get-MapyComUrl -Origin "$($geoStart.Latitude),$($geoStart.Longitude)" `
                -Destination "$($geoEnd.Latitude),$($geoEnd.Longitude)" `
                -Waypoints $geoWp

            $results.Add([PSCustomObject]@{
                Id               = if ($r.Id) { $r.Id } else { ($i + 1) }
                Name             = $routeName
                Start_Original   = $r.Start
                Start            = $r.Start
                Start_Geocoded   = $geoStart.FormattedAddress
                StartGeocoded    = $geoStart.FormattedAddress
                Start_Status     = $startStatus
                StartStatus      = $startStatus
                End_Original     = $r.End
                End              = $r.End
                End_Geocoded     = $geoEnd.FormattedAddress
                EndGeocoded      = $geoEnd.FormattedAddress
                End_Status       = $endStatus
                EndStatus        = $endStatus
                WaypointsCount   = $geoWp.Count
                RouteType        = $rType
                DistanceKm       = $routeData.OdlegloscKm
                DurationMin      = $routeData.CzasMin
                Status           = 'OK'
                MapPath          = $(if ($saved) { $mapPath } else { $null })
                EncodedPolyline  = $routeData.EncodedPolyline
                GoogleMapsUrl    = $gUrl
                MapyComUrl       = $gUrl
                RoutePoints      = $routePointsList
                AvoidTolls       = $avoidT
                AvoidHighways    = $avoidH
                AvoidFerries     = $avoidF
            })
        }
        catch {
            & $wlog "Route $($i + 1) exception: $($_.Exception.Message)" "ERROR"
            $results.Add([PSCustomObject]@{
                Id               = if ($r.Id) { $r.Id } else { ($i + 1) }
                Name             = $routeName
                Start_Original   = $r.Start
                Start            = $r.Start
                Start_Geocoded   = $null
                StartGeocoded    = $null
                Start_Status     = "Exception"
                StartStatus      = "Exception"
                End_Original     = $r.End
                End              = $r.End
                End_Geocoded     = $null
                EndGeocoded      = $null
                End_Status       = "Exception"
                EndStatus        = "Exception"
                WaypointsCount   = 0
                RouteType        = $rType
                DistanceKm       = $null
                DurationMin      = $null
                Status           = "Exception: $($_.Exception.Message)"
                MapPath          = $null
                EncodedPolyline  = $null
                GoogleMapsUrl    = $null
                MapyComUrl       = $null
                RoutePoints      = @()
            })
        }
    }

    return [PSCustomObject]@{
        Results  = $results
        ApiUsage = [PSCustomObject]@{
            Geocoding  = $geoCount
            Routes     = $routesCount
            StaticMaps = $staticCount
        }
    }
}

#endregion 3. Batch Route Processing Async Worker

#region 4. Pre-Batch Geocode Validation Async Worker

<#
.SYNOPSIS
    Asynchronous worker scriptblock executing pre-batch geocoding validation.
.DESCRIPTION
    Iterates through deduplicated addresses before running full routing, verifying
    whether coordinates can be resolved, identifying approximate/fallback locations,
    and outputting structured precision diagnostics without incurring route calculation costs.
#>
$script:GeocodeValidationAsync = {
    param($addressItems, $apiKey, $languageCode = 'en', $syncState, $logFile)

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls

    $geoCount = 0
    $wlog = {
        param($msg, $lvl = 'INFO')
        if ($syncState.LogQueue) {
            $syncState.LogQueue.Enqueue([PSCustomObject]@{ Level = $lvl; Message = $msg })
        }
        if ($logFile) {
            $t = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
            try { 
                [System.IO.File]::AppendAllText($logFile, "[$t] [$lvl] [ValidationWorker] $msg`r`n", [System.Text.UTF8Encoding]::new($true)) 
            } catch { }
        }
    }

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $total = $addressItems.Count
    & $wlog "Starting pre-batch geocode validation ($total unique addresses)..." "INFO"

    for ($i = 0; $i -lt $total; $i++) {
        # Honor user cancellation request
        if ($syncState.CancelRequested) {
            & $wlog "Geocode validation cancelled by user at address $($i + 1)/$total." "WARN"
            break
        }
        $item = $addressItems[$i]
        $syncState.CurrentIndex = ($i + 1)
        $rawAddr = [string]$item.Address

        & $wlog "Validating [$($i + 1)/$total]: '$rawAddr'..." "INFO"

        # Query Geocoding API
        $geo = Get-AddressCoordinates -Address $rawAddr -ApiKey $apiKey -LanguageCode $languageCode
        $geoCount++

        $statusStr = if ($geo.Status -eq 'OK') { 'OK' } else { $geo.Status }
        $matchType = if ($geo.MatchType) { $geo.MatchType } else { 'NONE' }
        $lat = if ($geo.Latitude) { [double]$geo.Latitude } else { $null }
        $lng = if ($geo.Longitude) { [double]$geo.Longitude } else { $null }
        $formatted = if ($geo.FormattedAddress) { [string]$geo.FormattedAddress } else { '' }

        # Classify geocode accuracy precision
        $precision = switch ($matchType) {
            'ROOFTOP'            { 'ROOFTOP (Exact)' }
            'RANGE_INTERPOLATED' { 'RANGE (Interpolated)' }
            'GEOMETRIC_CENTER'   { 'CENTER (Geometric)' }
            'APPROXIMATE'        { 'APPROXIMATE (Low Precision)' }
            default              { $matchType }
        }

        $results.Add([PSCustomObject]@{
            Index            = ($i + 1)
            Address          = $rawAddr
            Role             = [string]$item.Role
            Status           = $statusStr
            Precision        = $precision
            FormattedAddress = $formatted
            Latitude         = $lat
            Longitude        = $lng
            IsOk             = ($geo.Status -eq 'OK')
            IsRooftop        = ($matchType -eq 'ROOFTOP')
        })
    }

    & $wlog "Geocode validation completed. Processed $geoCount addresses." "OK"
    return [PSCustomObject]@{
        Results  = $results
        ApiUsage = [PSCustomObject]@{
            Geocoding  = $geoCount
            Routes     = 0
            StaticMaps = 0
        }
    }
}

#endregion 4. Pre-Batch Geocode Validation Async Worker

#region 5. Global Function & ScriptBlock Exports

# Export helper function into global scope
Set-Item -Path "function:global:New-WorkerPowerShell" -Value (Get-Item "function:New-WorkerPowerShell").ScriptBlock -ErrorAction SilentlyContinue

#endregion 5. Global Function & ScriptBlock Exports
