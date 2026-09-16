#Requires -Version 5.1
<#
.SYNOPSIS
    Mapy.com Routes & Map Generator — Interactive WebView2 Map Subsystem.
.DESCRIPTION
    Initializes and manages the Microsoft Edge WebView2 Chromium control for WPF,
    generating dynamic vector route visualizations with Leaflet.js, markers, popups,
    and adaptive Dark/Light tile themes with seamless fallback to static PNG images.
.NOTES
    Encoding: UTF-8 with BOM
    Compatibility: Windows PowerShell 5.1 and PowerShell 7+
#>

$script:HasWebView2 = $false
$script:WebView2Control = $null
$script:CoreWebView2Env = $null
$script:FallbackWebBrowser = $null

#region 1. WebView2 Environment & Runtime Lifecycle

<#
.SYNOPSIS
    Initializes the WebView2 Chromium runtime environment and WPF host control.
.DESCRIPTION
    Scans known directory locations and bundled application folders for Microsoft.Web.WebView2.Wpf.dll
    and Microsoft.Web.WebView2.Core.dll. When located, configures a dedicated user data folder
    under %LOCALAPPDATA% to prevent E_ACCESSDENIED errors, initializes CoreWebView2Environment,
    and binds bidirectional WebMessageReceived event handlers to pass coordinate clicks from Leaflet.js
    back into the WPF UI dispatcher thread.
.OUTPUTS
    [bool] $true if WebView2 was successfully located, initialized, and bound; otherwise $false.
.EXAMPLE
    if (Initialize-WebView2Environment) {
        Write-Verbose "WebView2 Chromium runtime ready."
    }
#>
function Initialize-WebView2Environment {
    [CmdletBinding()]
    param()

    # Return cached true status if already successfully initialized
    if ($script:HasWebView2 -and $script:WebView2Control -and $script:CoreWebView2Env) { 
        return $true 
    }

    # Determine base directory across script execution and PS2EXE compiled standalone states
    $baseDir = if (-not [string]::IsNullOrWhiteSpace($script:AppDir)) {
        $script:AppDir
    } elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    } else {
        [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    }

    $parentDir = if (-not [string]::IsNullOrWhiteSpace($baseDir)) { Split-Path -Parent $baseDir } else { $null }

    # Candidates for Microsoft.Web.WebView2 assemblies (local lib folders + common system installs)
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($baseDir)) {
        $candidates.Add((Join-Path $baseDir 'lib\Microsoft.Web.WebView2.Wpf.dll'))
    }
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $candidates.Add((Join-Path $parentDir 'lib\Microsoft.Web.WebView2.Wpf.dll'))
    }
    $systemPaths = @(
        # Common software packages shipping with valid WebView2 WPF runtimes
        "C:\Program Files\Fortinet\FortiClient\Microsoft.Web.WebView2.Wpf.dll",
        "C:\Program Files\PowerToys\Microsoft.Web.WebView2.Wpf.dll",
        "C:\Program Files\Surfshark\Microsoft.Web.WebView2.Wpf.dll",
        "C:\Program Files\Microsoft SQL Server Management Studio 21\Release\Common7\IDE\PrivateAssemblies\Microsoft.Web.WebView2.Wpf.dll",
        "C:\Program Files\Microsoft Office\root\Office16\WritingAssistant\Microsoft.Web.WebView2.Wpf.dll"
    )
    foreach ($sp in $systemPaths) {
        $candidates.Add([string]$sp)
    }

    $loaded = $false
    foreach ($cand in $candidates) {
        if ($cand -and (Test-Path $cand)) {
            $dir = Split-Path -Parent $cand
            $core = Join-Path $dir 'Microsoft.Web.WebView2.Core.dll'
            if (Test-Path $core) {
                try {
                    # Load both Core and WPF assemblies into current AppDomain
                    Add-Type -Path $core -ErrorAction Stop
                    Add-Type -Path $cand -ErrorAction Stop
                    
                    # Verify instantiability with current CLR / runtime
                    $testControl = [Microsoft.Web.WebView2.Wpf.WebView2]::new()
                    $loaded = $true
                    if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
                        Write-AppLog "Loaded WebView2 WPF assemblies from: $dir" "OK"
                    }
                    break
                }
                catch {
                    # Continue checking next candidate if architecture/CLR mismatch occurs
                }
            }
        }
    }

    if (-not $loaded) {
        if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
            Write-AppLog "WebView2 WPF assembly not found or incompatible. Static PNG mode / fallback will be active." "INFO"
        }
        $script:HasWebView2 = $false
        return $false
    }

    # Crucial: Configure writable UserDataFolder in LocalAppData to avoid E_ACCESSDENIED (0x80070005)
    # which occurs when default working directory is C:\Windows\System32 or Program Files
    $userDataFolder = Join-Path $env:LOCALAPPDATA "MapyComRoutes\WebView2Data"
    if (-not (Test-Path $userDataFolder)) {
        try { [System.IO.Directory]::CreateDirectory($userDataFolder) | Out-Null } catch { }
    }

    try {
        $envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $userDataFolder, $null)
        $script:CoreWebView2Env = $envTask.GetAwaiter().GetResult()
    }
    catch {
        if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
            Write-AppLog "CoreWebView2Environment creation failed: $($_.Exception.Message)" "WARN"
        }
        $script:CoreWebView2Env = $null
    }

    try {
        $wv = [Microsoft.Web.WebView2.Wpf.WebView2]::new()
        
        # Ensure asynchronous CoreWebView2 initialization completes when control loads in visual tree
        $wv.Add_Loaded({
            if ($script:CoreWebView2Env -and (-not $script:WebView2Initialized)) {
                try {
                    $script:WebView2Control.EnsureCoreWebView2Async($script:CoreWebView2Env) | Out-Null
                    $script:WebView2Initialized = $true
                } catch { }
            }
        })

        # Register message listener to receive interactive clicks, popups, and pin actions from Leaflet.js
        $wv.Add_CoreWebView2InitializationCompleted({
            param($s, $e)
            if ($e.IsSuccess -and $wv.CoreWebView2) {
                $wv.CoreWebView2.add_WebMessageReceived({
                    param($sender, $args)
                    try {
                        $raw = $args.WebMessageAsJson
                        $data = $raw | ConvertFrom-Json
                        if ($data -and $data.action) {
                            $w = if ($script:Controls -and $script:Controls.Window) { 
                                $script:Controls.Window 
                            } elseif ($script:MainWindow) { 
                                $script:MainWindow 
                            } else { 
                                [System.Windows.Application]::Current.MainWindow 
                            }

                            # Dispatch UI updates to STA thread safely
                            if ($w -and $w.Dispatcher) {
                                $w.Dispatcher.BeginInvoke([Action]{
                                    $script:SuppressAutosuggest = $true
                                    try {
                                        if ($data.action -eq 'setStart' -and $script:Controls.txtManualStart) {
                                            $script:Controls.txtManualStart.Text = $data.coords
                                            if ($script:Controls.badgeStartGeocode) {
                                                $script:Controls.badgeStartGeocode.Text = '🟢 Exact'
                                                $script:Controls.badgeStartGeocode.Foreground = [System.Windows.Media.Brushes]::LimeGreen
                                                $script:Controls.badgeStartGeocode.ToolTip = "Coordinates: $($data.coords)"
                                            }
                                            if (Get-Command Show-AppToastNotification -ErrorAction SilentlyContinue) {
                                                Show-AppToastNotification -Title "Map Location Selected" -Message "Start set to $($data.coords)" -Type Info
                                            }
                                        }
                                        elseif ($data.action -eq 'setDest' -and $script:Controls.txtManualEnd) {
                                            $script:Controls.txtManualEnd.Text = $data.coords
                                            if ($script:Controls.badgeEndGeocode) {
                                                $script:Controls.badgeEndGeocode.Text = '🟢 Exact'
                                                $script:Controls.badgeEndGeocode.Foreground = [System.Windows.Media.Brushes]::LimeGreen
                                                $script:Controls.badgeEndGeocode.ToolTip = "Coordinates: $($data.coords)"
                                            }
                                            if (Get-Command Show-AppToastNotification -ErrorAction SilentlyContinue) {
                                                Show-AppToastNotification -Title "Map Location Selected" -Message "Destination set to $($data.coords)" -Type Info
                                            }
                                        }
                                        elseif ($data.action -eq 'addStop' -and $script:Controls.lstWaypoints) {
                                            $null = $script:Controls.lstWaypoints.Items.Add($data.coords)
                                            if (Get-Command Show-AppToastNotification -ErrorAction SilentlyContinue) {
                                                Show-AppToastNotification -Title "Map Location Selected" -Message "Waypoint added: $($data.coords)" -Type Info
                                            }
                                        }
                                        elseif ($data.action -eq 'copy') {
                                            [System.Windows.Clipboard]::SetText($data.coords)
                                            if (Get-Command Show-AppToastNotification -ErrorAction SilentlyContinue) {
                                                Show-AppToastNotification -Title "Coordinates Copied" -Message "Copied to clipboard: $($data.coords)" -Type Info
                                            }
                                        }
                                    }
                                    finally {
                                        $script:SuppressAutosuggest = $false
                                    }
                                }) | Out-Null
                            }
                        }
                    } catch { }
                })
            }
        })

        $wv.Add_NavigationCompleted({
            param($s, $e)
            if ($e.IsSuccess) {
                if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
                    Write-AppLog "Interactive Leaflet map rendered successfully." "INFO"
                }
            } else {
                if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
                    Write-AppLog "Interactive map navigation status: $($e.WebErrorStatus)" "WARN"
                }
            }
        })

        $script:WebView2Control = $wv
        $script:HasWebView2 = $true
        return $true
    }
    catch {
        if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
            Write-AppLog "WebView2 instantiation failed: $($_.Exception.Message)" "WARN"
        }
        $script:HasWebView2 = $false
        return $false
    }
}

#endregion 1. WebView2 Environment & Runtime Lifecycle

#region 2. Polyline Decoding Engine

<#
.SYNOPSIS
    Decodes a standard Google/Mapy.com Encoded Polyline string into latitude and longitude coordinates.
.DESCRIPTION
    Implements the standard lossy polyline compression decoding algorithm (5-bit chunking with
    ASCII offset of 63 and zigzag encoding). Produces a strongly typed list of latitude and
    longitude coordinate pairs accurate to 5-6 decimal places (WGS84).
.PARAMETER EncodedPolyline
    The compressed polyline string returned by routing APIs (e.g. Google Maps or Mapy.com).
.OUTPUTS
    [System.Collections.Generic.List[PSCustomObject]] Collection of objects with Latitude and Longitude properties.
.EXAMPLE
    $points = ConvertFrom-GoogleEncodedPolyline -EncodedPolyline "_p~iF~ps|U_ulLnnqC_mqNvxq`@"
    $points | ForEach-Object { "$($_.Latitude), $($_.Longitude)" }
#>
function ConvertFrom-GoogleEncodedPolyline {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$EncodedPolyline)

    $points = [System.Collections.Generic.List[PSCustomObject]]::new()
    if ([string]::IsNullOrWhiteSpace($EncodedPolyline)) { 
        return $points 
    }

    $len = $EncodedPolyline.Length
    $index = 0
    $lat = 0
    $lng = 0

    # Polyline decoding loop: sequentially unrolls latitude delta then longitude delta
    while ($index -lt $len) {
        $b = 0
        $shift = 0
        $result = 0
        do {
            $b = [int][char]$EncodedPolyline[$index++] - 63
            $result = $result -bor (($b -band 0x1f) -shl $shift)
            $shift += 5
        } while ($b -ge 0x20 -and $index -lt $len)
        $dlat = if (($result -band 1) -ne 0) { -bnot ($result -shr 1) } else { ($result -shr 1) }
        $lat += $dlat

        $shift = 0
        $result = 0
        do {
            $b = [int][char]$EncodedPolyline[$index++] - 63
            $result = $result -bor (($b -band 0x1f) -shl $shift)
            $shift += 5
        } while ($b -ge 0x20 -and $index -lt $len)
        $dlng = if (($result -band 1) -ne 0) { -bnot ($result -shr 1) } else { ($result -shr 1) }
        $lng += $dlng

        # Scaling factor: Google Encoded Polyline represents coordinate * 1e5
        $points.Add([PSCustomObject]@{
            Latitude  = [math]::Round($lat * 1e-5, 6)
            Longitude = [math]::Round($lng * 1e-5, 6)
        })
    }

    return $points
}

#endregion 2. Polyline Decoding Engine

#region 3. Dynamic HTML Leaflet Map Generation

<#
.SYNOPSIS
    Generates a standalone, interactive Leaflet.js HTML map file for a calculated route.
.DESCRIPTION
    Creates a responsive HTML5 page containing Leaflet.js map integration with multi-layer
    support (CARTO Voyager/Dark, OpenStreetMap, Esri World Imagery Satellite). Displays
    the decoded route polyline, origin marker (pin A), destination marker (pin B), intermediate
    waypoint markers, and an interactive context menu enabling clicks to set start, end, or stops.
.PARAMETER RouteName
    Title or identifier for the route. Defaults to 'Route'.
.PARAMETER EncodedPolyline
    Encoded polyline string representing the path geometry.
.PARAMETER OriginLat
    Latitude of the route start point.
.PARAMETER OriginLng
    Longitude of the route start point.
.PARAMETER OriginAddress
    Human-readable street address or label for the origin point.
.PARAMETER DestLat
    Latitude of the route destination point.
.PARAMETER DestLng
    Longitude of the route destination point.
.PARAMETER DestAddress
    Human-readable street address or label for the destination point.
.PARAMETER Waypoints
    Array of waypoint objects with Latitude, Longitude, and optional Address properties.
.PARAMETER DistanceKm
    Calculated route distance in kilometers.
.PARAMETER DurationMin
    Calculated route duration in minutes.
.PARAMETER RouteType
    Route optimization mode ('Fastest', 'Shortest', 'Eco'). Defaults to 'Fastest'.
.PARAMETER IsDarkMode
    Enables dark mode theme styling and CARTO Dark basemap tiles. Defaults to $true.
.PARAMETER OutputPath
    Target path for the generated HTML file. If omitted, creates a temporary file in %TEMP%.
.PARAMETER CartoApiKey
    Optional CARTO API key for high-volume or authenticated tile rendering.
.OUTPUTS
    [string] Full filesystem path to the generated HTML map file.
.EXAMPLE
    $htmlPath = New-RouteHtmlMap -RouteName "Fleet 1" -EncodedPolyline $polyline `
        -OriginLat 52.23 -OriginLng 21.01 -DestLat 50.06 -DestLng 19.94
#>
function New-RouteHtmlMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$RouteName = 'Route',
        [Parameter(Mandatory = $true)][string]$EncodedPolyline,
        [Parameter()][double]$OriginLat = 0,
        [Parameter()][double]$OriginLng = 0,
        [Parameter()][string]$OriginAddress = '',
        [Parameter()][double]$DestLat = 0,
        [Parameter()][double]$DestLng = 0,
        [Parameter()][string]$DestAddress = '',
        [Parameter()][object[]]$Waypoints = @(),
        [Parameter()][double]$DistanceKm = 0,
        [Parameter()][int]$DurationMin = 0,
        [Parameter()][string]$RouteType = 'Fastest',
        [Parameter()][bool]$IsDarkMode = $true,
        [Parameter()][string]$OutputPath = '',
        [Parameter()][string]$CartoApiKey = ''
    )

    if ([string]::IsNullOrWhiteSpace($RouteName)) {
        $RouteName = 'Route'
    }
    if (-not $OutputPath) {
        $ts = (Get-Date).ToString('yyyyMMdd_HHmmssfff')
        $OutputPath = Join-Path $env:TEMP "gmaps_interactive_route_${ts}.html"
    }

    # Clean up older temporary route map HTML files (> 30 mins old) to avoid disk clutter
    try {
        Get-ChildItem -Path $env:TEMP -Filter 'gmaps_interactive_route_*.html' -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddMinutes(-30) } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } catch { }

    # Resolve Carto API key priority: parameter > AppConfig > Environment variable
    $resolvedCartoKey = if (-not [string]::IsNullOrWhiteSpace($CartoApiKey)) {
        $CartoApiKey.Trim()
    } elseif ($script:AppConfig -and -not [string]::IsNullOrWhiteSpace($script:AppConfig.CartoApiKey)) {
        $script:AppConfig.CartoApiKey.Trim()
    } elseif (-not [string]::IsNullOrWhiteSpace($env:CARTO_API_KEY)) {
        $env:CARTO_API_KEY.Trim()
    } else {
        ''
    }

    # Decode path geometry into point coordinates
    $points = if ($EncodedPolyline) { ConvertFrom-GoogleEncodedPolyline -EncodedPolyline $EncodedPolyline } else { @() }
    
    # Format coordinates array for Leaflet: [[lat, lng], [lat, lng], ...]
    $coordJsonList = [System.Collections.Generic.List[string]]::new()
    foreach ($pt in $points) {
        $lat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$pt.Latitude)
        $lng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$pt.Longitude)
        $coordJsonList.Add("[$lat, $lng]")
    }
    $coordJson = "[$($coordJsonList -join ',')]"

    # Tile layer URL and attribution based on dark/light theme
    $tileUrl = if ($IsDarkMode) {
        'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png'
    } else {
        'https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png'
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedCartoKey)) {
        $escapedCarto = [System.Uri]::EscapeDataString($resolvedCartoKey)
        $tileUrl += "?key=$escapedCarto"
    }
    $bgStyle = if ($IsDarkMode) { 'background:#0f172a;color:#f8fafc;' } else { 'background:#f8fafc;color:#0f172a;' }
    $polyColor = if ($IsDarkMode) { '#38bdf8' } else { '#2563eb' }

    $safeRouteName = [System.Net.WebUtility]::HtmlEncode($RouteName)
    $safeOrigin = [System.Net.WebUtility]::HtmlEncode($OriginAddress)
    $safeDest = [System.Net.WebUtility]::HtmlEncode($DestAddress)

    # Format intermediate waypoints for Leaflet injection
    $wpJsonList = [System.Collections.Generic.List[string]]::new()
    if ($Waypoints -and @($Waypoints).Count -gt 0) {
        $idx = 1
        foreach ($wp in $Waypoints) {
            if ($wp.Latitude -and $wp.Longitude) {
                $lat = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Latitude)
                $lng = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Longitude)
                $addr = [System.Net.WebUtility]::HtmlEncode($(if ($wp.Address) { [string]$wp.Address } else { "Stop $idx" }))
                $wpJsonList.Add("{ lat: $lat, lng: $lng, label: '$idx', address: '$addr' }")
                $idx++
            }
        }
    }
    $wpJson = "[$($wpJsonList -join ',')]"

    $origLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $OriginLat)
    $origLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $OriginLng)
    $destLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $DestLat)
    $destLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $DestLng)

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>$safeRouteName</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>
    html, body, #map { width: 100%; height: 100%; margin: 0; padding: 0; overflow: hidden; $bgStyle font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
    .route-info-card {
      position: absolute;
      top: 12px;
      right: 12px;
      z-index: 1000;
      background: rgba(15, 23, 42, 0.88);
      color: #f8fafc;
      padding: 10px 14px;
      border-radius: 8px;
      backdrop-filter: blur(8px);
      box-shadow: 0 4px 16px rgba(0,0,0,0.3);
      font-size: 13px;
      line-height: 1.4;
      border: 1px solid rgba(255,255,255,0.1);
      max-width: 300px;
      pointer-events: none;
    }
    .route-info-card strong { color: #38bdf8; }
    .marker-pin {
      width: 30px;
      height: 30px;
      border-radius: 50% 50% 50% 0;
      background: #0284c7;
      position: absolute;
      transform: rotate(-45deg);
      left: 50%;
      top: 50%;
      margin: -15px 0 0 -15px;
      box-shadow: 0 2px 5px rgba(0,0,0,0.4);
    }
    .marker-pin::after {
      content: '';
      width: 14px;
      height: 14px;
      margin: 8px 0 0 8px;
      background: #fff;
      position: absolute;
      border-radius: 50%;
    }
    .marker-pin span {
      position: absolute;
      transform: rotate(45deg);
      top: 5px;
      left: 7px;
      font-size: 11px;
      font-weight: bold;
      color: #0f172a;
      z-index: 10;
    }
    .pin-start { background: #10b981; }
    .pin-dest { background: #ef4444; }
    .pin-wp { background: #3b82f6; }
    .leaflet-popup-content-wrapper { background: #1e293b; color: #f8fafc; border: 1px solid #334155; border-radius: 8px; }
    .leaflet-popup-tip { background: #1e293b; }
  </style>
</head>
<body>
  <div id="map"></div>
  <div class="route-info-card">
    <div><strong>$safeRouteName</strong></div>
    <div>Distance: $DistanceKm km | Time: $DurationMin min</div>
    <div style="font-size:11px;opacity:0.8;">Type: $RouteType</div>
  </div>

  <script>
    const coords = $coordJson;
    const waypoints = $wpJson;

    const map = L.map('map', { zoomControl: true });

    // Base Tile Layers: CARTO Standard, OpenStreetMap, Esri Satellite
    const baseStandard = L.tileLayer('$tileUrl', {
      maxZoom: 19,
      attribution: '&copy; CARTO &copy; OpenStreetMap'
    });

    const baseOsm = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap contributors'
    });

    const baseSatellite = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
      maxZoom: 18,
      attribution: '&copy; Esri, Earthstar Geographics'
    });

    const baseMaps = {
      "🗺️ Standard": baseStandard,
      "🌐 OpenStreetMap": baseOsm,
      "🛰️ Satellite": baseSatellite
    };

    baseStandard.addTo(map);
    L.control.layers(baseMaps, null, { position: 'topleft' }).addTo(map);

    // Send action from JavaScript runtime through WebView2 IPC channel to host WPF application
    function sendMapAction(action, lat, lng) {
      const coordStr = lat + ', ' + lng;
      try {
        if (window.chrome && window.chrome.webview) {
          window.chrome.webview.postMessage({ action: action, coords: coordStr, lat: lat, lng: lng });
        }
      } catch (err) {}
      try {
        navigator.clipboard.writeText(coordStr);
      } catch (err) {}
    }

    // Context menu / right-click handler for direct map coordinate capture
    map.on('contextmenu', function(e) {
      const lat = e.latlng.lat.toFixed(5);
      const lng = e.latlng.lng.toFixed(5);
      const coordStr = lat + ', ' + lng;
      const html = '<div style="font-family:sans-serif;font-size:12px;min-width:160px;line-height:1.4;">' +
        '<div style="font-weight:bold;margin-bottom:6px;color:#94a3b8;">📍 ' + coordStr + '</div>' +
        '<button style="width:100%;text-align:left;margin-bottom:4px;padding:4px 8px;cursor:pointer;background:#10b981;color:#fff;border:none;border-radius:4px;font-size:11px;" onclick="sendMapAction(\'setStart\',' + lat + ',' + lng + ');map.closePopup();">🟢 Set as Start</button>' +
        '<button style="width:100%;text-align:left;margin-bottom:4px;padding:4px 8px;cursor:pointer;background:#3b82f6;color:#fff;border:none;border-radius:4px;font-size:11px;" onclick="sendMapAction(\'addStop\',' + lat + ',' + lng + ');map.closePopup();">🔵 Add as Stop</button>' +
        '<button style="width:100%;text-align:left;margin-bottom:4px;padding:4px 8px;cursor:pointer;background:#ef4444;color:#fff;border:none;border-radius:4px;font-size:11px;" onclick="sendMapAction(\'setDest\',' + lat + ',' + lng + ');map.closePopup();">🔴 Set as Destination</button>' +
        '<button style="width:100%;text-align:left;padding:4px 8px;cursor:pointer;background:#475569;color:#fff;border:none;border-radius:4px;font-size:11px;" onclick="sendMapAction(\'copy\',' + lat + ',' + lng + ');map.closePopup();">📋 Copy Lat/Lng</button>' +
        '</div>';
      L.popup().setLatLng(e.latlng).setContent(html).openOn(map);
    });

    // Left-click shortcut popup
    map.on('click', function(e) {
      const lat = e.latlng.lat.toFixed(5);
      const lng = e.latlng.lng.toFixed(5);
      const coordStr = lat + ', ' + lng;
      const html = '<div style="font-family:sans-serif;font-size:12px;min-width:150px;">' +
        '<strong>📍 Location</strong><br/>' + coordStr + '<br/>' +
        '<div style="display:flex;gap:4px;margin-top:6px;">' +
        '<button style="flex:1;padding:3px 6px;cursor:pointer;background:#10b981;color:#fff;border:none;border-radius:4px;font-size:10px;" onclick="sendMapAction(\'setStart\',' + lat + ',' + lng + ');map.closePopup();">Start</button>' +
        '<button style="flex:1;padding:3px 6px;cursor:pointer;background:#ef4444;color:#fff;border:none;border-radius:4px;font-size:10px;" onclick="sendMapAction(\'setDest\',' + lat + ',' + lng + ');map.closePopup();">Dest</button>' +
        '<button style="flex:1;padding:3px 6px;cursor:pointer;background:#2563eb;color:#fff;border:none;border-radius:4px;font-size:10px;" onclick="sendMapAction(\'copy\',' + lat + ',' + lng + ');map.closePopup();">Copy</button>' +
        '</div></div>';
      L.popup().setLatLng(e.latlng).setContent(html).openOn(map);
    });

    function createIcon(label, cls) {
      return L.divIcon({
        className: 'custom-pin',
        html: '<div class="marker-pin ' + cls + '"><span>' + label + '</span></div>',
        iconSize: [30, 42],
        iconAnchor: [15, 38],
        popupAnchor: [0, -34]
      });
    }

    // Refit viewport on window resize or initial render
    window.addEventListener('resize', () => {
      map.invalidateSize();
    });

    setTimeout(() => {
      map.invalidateSize();
      if (coords && coords.length > 0 && typeof poly !== 'undefined') {
        map.fitBounds(poly.getBounds(), { padding: [40, 40] });
      }
    }, 200);

    let poly = null;
    if (coords && coords.length > 0) {
      poly = L.polyline(coords, {
        color: '$polyColor',
        weight: 5,
        opacity: 0.9,
        lineJoin: 'round'
      }).addTo(map);

      map.fitBounds(poly.getBounds(), { padding: [40, 40] });
    } else {
      map.setView([$origLatStr || 52.0, $origLngStr || 19.5], 10);
    }

    // Origin Marker (Pin A)
    if ($origLatStr && $origLngStr) {
      L.marker([$origLatStr, $origLngStr], { icon: createIcon('A', 'pin-start') })
        .addTo(map)
        .bindPopup('<strong>Origin (Start)</strong><br>$safeOrigin');
    }

    // Intermediate Waypoint Markers (Numbered Pins)
    waypoints.forEach(wp => {
      L.marker([wp.lat, wp.lng], { icon: createIcon(wp.label, 'pin-wp') })
        .addTo(map)
        .bindPopup('<strong>Waypoint ' + wp.label + '</strong><br>' + wp.address);
    });

    // Destination Marker (Pin B)
    if ($destLatStr && $destLngStr) {
      L.marker([$destLatStr, $destLngStr], { icon: createIcon('B', 'pin-dest') })
        .addTo(map)
        .bindPopup('<strong>Destination (End)</strong><br>$safeDest');
    }
  </script>
</body>
</html>
"@

    # Save HTML template using UTF-8 with BOM for cross-platform and accent compatibility
    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($true))
    return $OutputPath
}

#endregion 3. Dynamic HTML Leaflet Map Generation

#region 4. Browser Navigation & WPF Host Integration

<#
.SYNOPSIS
    Navigates the host WPF container to a rendered route map HTML document.
.DESCRIPTION
    Directs the active WebView2 control to the target HTML file URI. If WebView2 is
    unavailable, automatically mounts and navigates a standard WPF WebBrowser fallback
    control inside the provided HostPanel container.
.PARAMETER HtmlPath
    Filesystem path to the HTML map document to navigate.
.PARAMETER HostPanel
    WPF container element (e.g. Border or ContentControl) hosting the browser control.
.OUTPUTS
    [bool] $true if navigation was successfully executed; otherwise $false.
.EXAMPLE
    Navigate-RouteInteractiveMap -HtmlPath "C:\Temp\map.html" -HostPanel $Controls.pnlInteractiveMapHost
#>
function Navigate-RouteInteractiveMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$HtmlPath,
        [Parameter()][object]$HostPanel = $null
    )

    if (Initialize-WebView2Environment) {
        if ($HostPanel -and $script:WebView2Control) {
            if (-not $HostPanel.Child -or $HostPanel.Child -ne $script:WebView2Control) {
                $HostPanel.Child = $script:WebView2Control
            }
        }
        if ($script:WebView2Control) {
            try {
                if ($script:CoreWebView2Env -and (-not $script:WebView2Initialized)) {
                    try {
                        $script:WebView2Control.EnsureCoreWebView2Async($script:CoreWebView2Env) | Out-Null
                        $script:WebView2Initialized = $true
                    } catch {
                        # Outside active event loop; Loaded event handler will initialize CoreWebView2
                    }
                }

                $targetUri = [System.Uri]::new($HtmlPath)
                if ($script:WebView2Control.CoreWebView2) {
                    $script:WebView2Control.CoreWebView2.Navigate($targetUri.AbsoluteUri)
                }
                $script:WebView2Control.Source = $targetUri
                return $true
            }
            catch {
                if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
                    Write-AppLog "WebView2 navigation error: $($_.Exception.Message)" "WARN"
                }
            }
        }
    }

    # Fallback to WPF WebBrowser (Internet Explorer engine) if WebView2 is unavailable
    if ($HostPanel) {
        if (-not $script:FallbackWebBrowser) {
            $script:FallbackWebBrowser = [System.Windows.Controls.WebBrowser]::new()
        }
        if (-not $HostPanel.Child -or $HostPanel.Child -ne $script:FallbackWebBrowser) {
            $HostPanel.Child = $script:FallbackWebBrowser
        }
        try {
            $script:FallbackWebBrowser.Navigate([System.Uri]::new($HtmlPath))
            return $true
        }
        catch {
            if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
                Write-AppLog "WebBrowser fallback navigation failed: $($_.Exception.Message)" "WARN"
            }
        }
    }

    return $false
}

<#
.SYNOPSIS
    Refreshes the interactive route map based on the most recent calculation result.
.DESCRIPTION
    Extracts polyline coordinates, waypoints, addresses, and current application theme
    settings from the calculation result object or $script:LastManualResult, builds a new
    Leaflet HTML map file, and navigates the interactive map host panel to display it.
.PARAMETER CalcResult
    Optional route calculation result object. If omitted, uses $script:LastManualResult.
.OUTPUTS
    None.
.EXAMPLE
    Update-InteractiveRouteMap -CalcResult $routeResult
#>
function Update-InteractiveRouteMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object]$CalcResult = $null
    )

    $calc = if ($CalcResult) { $CalcResult } else { $script:LastManualResult }
    if (-not $calc -or -not $calc.EncodedPolyline) { 
        return 
    }

    $isDark = ($script:CurrentTheme -ne 'Light')
    $resolvedRouteName = if (-not [string]::IsNullOrWhiteSpace($script:ActiveManualRouteName)) {
        $script:ActiveManualRouteName
    } elseif ($script:Controls -and $script:Controls.txtManualName -and -not [string]::IsNullOrWhiteSpace($script:Controls.txtManualName.Text)) {
        $script:Controls.txtManualName.Text.Trim()
    } else {
        'Route'
    }
    if ([string]::IsNullOrWhiteSpace($resolvedRouteName)) { 
        $resolvedRouteName = 'Route' 
    }

    $cartoKey = if (Get-Command Get-CurrentCartoApiKey -ErrorAction SilentlyContinue) {
        Get-CurrentCartoApiKey
    } elseif ($script:AppConfig -and -not [string]::IsNullOrWhiteSpace($script:AppConfig.CartoApiKey)) {
        $script:AppConfig.CartoApiKey.Trim()
    } else {
        ''
    }

    $htmlMap = New-RouteHtmlMap -RouteName $resolvedRouteName `
        -EncodedPolyline $calc.EncodedPolyline `
        -OriginLat $calc.OriginLat -OriginLng $calc.OriginLng -OriginAddress $calc.OriginAddress `
        -DestLat $calc.DestLat -DestLng $calc.DestLng -DestAddress $calc.DestAddress `
        -Waypoints $calc.Waypoints -DistanceKm $calc.DistanceKm -DurationMin $calc.DurationMin `
        -RouteType $calc.RouteType -IsDarkMode $isDark -CartoApiKey $cartoKey

    $script:LastInteractiveMapPath = $htmlMap
    $mapHostPanel = if ($script:Controls) { $script:Controls.pnlInteractiveMapHost } else { $null }

    Navigate-RouteInteractiveMap -HtmlPath $htmlMap -HostPanel $mapHostPanel | Out-Null
}

#endregion 4. Browser Navigation & WPF Host Integration

#region 5. Global Function Exports

# Export functions into global scope so caller scripts and GUI orchestrators can invoke them
Set-Item -Path "function:global:Initialize-WebView2Environment" -Value (Get-Item "function:Initialize-WebView2Environment").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:ConvertFrom-GoogleEncodedPolyline" -Value (Get-Item "function:ConvertFrom-GoogleEncodedPolyline").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:New-RouteHtmlMap" -Value (Get-Item "function:New-RouteHtmlMap").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Navigate-RouteInteractiveMap" -Value (Get-Item "function:Navigate-RouteInteractiveMap").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Update-InteractiveRouteMap" -Value (Get-Item "function:Update-InteractiveRouteMap").ScriptBlock -ErrorAction SilentlyContinue

#endregion 5. Global Function Exports
