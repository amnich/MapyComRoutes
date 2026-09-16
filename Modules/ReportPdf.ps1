#Requires -Version 5.1
<#
.SYNOPSIS
    Mapy.com Routes & Map Generator — PDF Report Generation Subsystem.
.DESCRIPTION
    Generates publication-quality PDF reports for single routes and batch summaries
    using native headless Microsoft Edge, embedding base64 route maps and stop itineraries.
.NOTES
    Encoding: UTF-8 with BOM
    Compatibility: Windows PowerShell 5.1 and PowerShell 7+
#>

#region 1. Browser Engine Discovery

<#
.SYNOPSIS
    Locates the Microsoft Edge executable (msedge.exe) on the host system.
.DESCRIPTION
    Searches standard 32-bit and 64-bit Program Files installation paths for msedge.exe,
    falling back to PATH resolution via Get-Command. Returns the full executable path
    or $null if Edge is not installed.
.OUTPUTS
    [string] Full path to msedge.exe, or $null if not located.
.EXAMPLE
    $edgePath = Find-EdgeExecutable
    if (-not $edgePath) { Write-Error "Microsoft Edge is required for PDF generation." }
#>
function Find-EdgeExecutable {
    [CmdletBinding()]
    param()

    # Search common Program Files locations on 64-bit and 32-bit Windows
    $candidates = @(
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    )
    foreach ($cand in $candidates) {
        if (Test-Path $cand) { 
            # Found via fixed installation path
            return $cand 
        }
    }

    # Fall back to PATH environment search if installed in a non-standard location
    $cmd = Get-Command msedge.exe -ErrorAction SilentlyContinue
    if ($cmd) { 
        return $cmd.Source 
    }

    # Return $null if headless PDF rendering capability cannot be fulfilled
    return $null
}

#endregion 1. Browser Engine Discovery

#region 2. Single Route PDF Generation

<#
.SYNOPSIS
    Exports an individual route report to a PDF document using headless Microsoft Edge.
.DESCRIPTION
    Builds a styled, self-contained HTML document containing route KPIs (distance, duration,
    optimization mode, avoid options), an embedded base64 route map image, and a complete stop
    itinerary table. The HTML document is converted to PDF via headless Edge and saved to the
    specified output file.
.PARAMETER OutputPath
    Target filesystem path where the generated PDF document will be written.
.PARAMETER RouteName
    Descriptive label or name for the route (e.g., "Prague to Brno"). Defaults to 'Route'.
.PARAMETER DistanceKm
    Total calculated route driving distance in kilometers.
.PARAMETER DurationMin
    Total calculated route driving duration in integer minutes.
.PARAMETER RouteType
    Routing optimization strategy used ('Fastest', 'Shortest', 'Eco'). Defaults to 'Fastest'.
.PARAMETER EmissionType
    Optional vehicle emission class or powertrain profile.
.PARAMETER AvoidTolls
    Indicates whether toll roads were excluded from the route calculation.
.PARAMETER AvoidHighways
    Indicates whether highways/freeways were excluded from the route calculation.
.PARAMETER AvoidFerries
    Indicates whether ferry crossings were excluded from the route calculation.
.PARAMETER OriginAddress
    Human-readable street address or location name for the starting point.
.PARAMETER OriginLat
    Latitude coordinate (WGS84) for the origin point.
.PARAMETER OriginLng
    Longitude coordinate (WGS84) for the origin point.
.PARAMETER DestAddress
    Human-readable street address or location name for the destination point.
.PARAMETER DestLat
    Latitude coordinate (WGS84) for the destination point.
.PARAMETER DestLng
    Longitude coordinate (WGS84) for the destination point.
.PARAMETER Waypoints
    Array of intermediate waypoint objects with Latitude, Longitude, and optional Address properties.
.PARAMETER MapImagePath
    Local filesystem path to a rendered static route map image (PNG) to embed as base64 in the report.
.PARAMETER GoogleMapsUrl
    Interactive Mapy.com or Google Maps web navigation URL for the external link button.
.PARAMETER Language
    Two-letter ISO language code for report localization ('en', 'de', 'pl'). Defaults to 'en'.
.PARAMETER ReportTitle
    Custom title displayed at the top of the report. Defaults to 'Route Report'.
.OUTPUTS
    [string] The verified OutputPath of the created PDF document.
.EXAMPLE
    Export-RoutePdfReport -OutputPath "C:\Reports\Route1.pdf" -RouteName "HQ to Branch" `
        -DistanceKm 45.2 -DurationMin 38 -OriginAddress "Main St 1" -DestAddress "Park Ave 10"
#>
function Export-RoutePdfReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $false)][string]$RouteName = 'Route',
        [Parameter()][double]$DistanceKm = 0,
        [Parameter()][int]$DurationMin = 0,
        [Parameter()][string]$RouteType = 'Fastest',
        [Parameter()][string]$EmissionType = '',
        [Parameter()][bool]$AvoidTolls = $false,
        [Parameter()][bool]$AvoidHighways = $false,
        [Parameter()][bool]$AvoidFerries = $false,
        [Parameter()][string]$OriginAddress = '',
        [Parameter()][double]$OriginLat = 0,
        [Parameter()][double]$OriginLng = 0,
        [Parameter()][string]$DestAddress = '',
        [Parameter()][double]$DestLat = 0,
        [Parameter()][double]$DestLng = 0,
        [Parameter()][object[]]$Waypoints = @(),
        [Parameter()][string]$MapImagePath = '',
        [Parameter()][Alias('MapyComUrl')][string]$GoogleMapsUrl = '',
        [Parameter()][string]$Language = 'en',
        [Parameter()][string]$ReportTitle = 'Route Report'
    )

    # Sanitize route name fallback
    if ([string]::IsNullOrWhiteSpace($RouteName)) {
        $RouteName = 'Route'
    }

    # Locate Microsoft Edge for headless printing
    $edgePath = Find-EdgeExecutable
    if (-not $edgePath) {
        throw "Microsoft Edge executable (msedge.exe) not found. PDF export requires Edge."
    }

    # Ensure parent output directory exists before spawning PDF generation
    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    # Inline map image as base64 data URI to eliminate external asset dependencies during headless print
    $mapImgTag = ''
    if ($MapImagePath -and (Test-Path $MapImagePath)) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($MapImagePath)
            $b64 = [Convert]::ToBase64String($bytes)
            $mapImgTag = "<div class='map-container'><img src='data:image/png;base64,$b64' alt='Route Map' /></div>"
        }
        catch {
            # Debugging note: Log or ignore image read failure so report generation can continue
            if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
                Write-AppLog "Failed to embed map image '$MapImagePath': $($_.Exception.Message)" "WARN"
            }
        }
    }

    $genDate = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    $safeName = [System.Net.WebUtility]::HtmlEncode($RouteName)
    $safeOrigin = [System.Net.WebUtility]::HtmlEncode($OriginAddress)
    $safeDest = [System.Net.WebUtility]::HtmlEncode($DestAddress)

    # Render avoid options badges for routing diagnostics
    $avoidBadges = [System.Collections.Generic.List[string]]::new()
    if ($AvoidTolls)    { $avoidBadges.Add("<span class='badge badge-warn'>🚫 Avoid Tolls</span>") }
    if ($AvoidHighways) { $avoidBadges.Add("<span class='badge badge-warn'>🚫 Avoid Highways</span>") }
    if ($AvoidFerries)  { $avoidBadges.Add("<span class='badge badge-warn'>🚫 Avoid Ferries</span>") }
    $avoidHtml = if ($avoidBadges.Count -gt 0) { $avoidBadges -join ' ' } else { "<span class='badge badge-info'>None</span>" }

    # Build stop itinerary table: Origin -> Intermediate Waypoints -> Destination
    $rowsHtml = [System.Text.StringBuilder]::new()
    
    # 1. Origin row (Green Pin A)
    $oLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $OriginLat)
    $oLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $OriginLng)
    $null = $rowsHtml.AppendLine("<tr><td><span class='pin-badge pin-start'>A</span></td><td><strong>Origin (Start)</strong></td><td>$safeOrigin</td><td>$oLatStr, $oLngStr</td></tr>")

    # 2. Waypoints rows (Blue numbered Pins 1..N)
    if ($Waypoints -and @($Waypoints).Count -gt 0) {
        $wIdx = 1
        foreach ($wp in $Waypoints) {
            $wLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Latitude)
            $wLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", [double]$wp.Longitude)
            $wAddr = [System.Net.WebUtility]::HtmlEncode($(if ($wp.Address) { [string]$wp.Address } else { "Waypoint $wIdx" }))
            $null = $rowsHtml.AppendLine("<tr><td><span class='pin-badge pin-wp'>$wIdx</span></td><td>Waypoint $wIdx</td><td>$wAddr</td><td>$wLatStr, $wLngStr</td></tr>")
            $wIdx++
        }
    }

    # 3. Destination row (Red Pin B)
    $dLatStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $DestLat)
    $dLngStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F6}", $DestLng)
    $null = $rowsHtml.AppendLine("<tr><td><span class='pin-badge pin-dest'>B</span></td><td><strong>Destination (End)</strong></td><td>$safeDest</td><td>$dLatStr, $dLngStr</td></tr>")

    # Generate unique temporary HTML container file
    $tempHtml = Join-Path $env:TEMP "gmaps_report_$([System.Guid]::NewGuid().ToString('N')).html"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>$safeName - Route Report</title>
  <style>
    @page { size: A4 portrait; margin: 12mm 14mm; }
    * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
    body { color: #1e293b; background: #ffffff; font-size: 12px; line-height: 1.4; }
    .header { border-bottom: 2px solid #2563eb; padding-bottom: 10px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: flex-end; }
    .header h1 { font-size: 20px; color: #0f172a; margin-bottom: 2px; }
    .header .subtitle { font-size: 11px; color: #64748b; }
    .header .meta { text-align: right; font-size: 11px; color: #64748b; }
    .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 14px; }
    .kpi-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; }
    .kpi-label { font-size: 10px; text-transform: uppercase; color: #64748b; font-weight: 600; margin-bottom: 2px; }
    .kpi-value { font-size: 16px; font-weight: bold; color: #0f172a; }
    .kpi-value.blue { color: #2563eb; }
    .kpi-value.green { color: #059669; }
    .map-container { width: 100%; border: 1px solid #cbd5e1; border-radius: 6px; overflow: hidden; margin-bottom: 14px; page-break-inside: avoid; text-align: center; background: #0f172a; }
    .map-container img { width: 100%; height: auto; max-height: 480px; object-fit: contain; display: block; margin: 0 auto; }
    .section-title { font-size: 13px; font-weight: bold; color: #0f172a; border-bottom: 1px solid #e2e8f0; padding-bottom: 4px; margin-bottom: 8px; margin-top: 6px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 14px; font-size: 11px; }
    th { background: #f1f5f9; text-align: left; padding: 6px 8px; font-weight: 600; color: #475569; border-bottom: 1px solid #cbd5e1; }
    td { padding: 6px 8px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
    tr:nth-child(even) td { background: #fafafa; }
    .pin-badge { display: inline-block; width: 20px; height: 20px; border-radius: 50%; color: white; text-align: center; line-height: 20px; font-weight: bold; font-size: 10px; }
    .pin-start { background: #10b981; }
    .pin-dest { background: #ef4444; }
    .pin-wp { background: #3b82f6; }
    .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 600; }
    .badge-warn { background: #fef3c7; color: #92400e; border: 1px solid #fde68a; }
    .badge-info { background: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; }
    .footer { margin-top: 15px; border-top: 1px solid #e2e8f0; padding-top: 6px; font-size: 10px; color: #94a3b8; display: flex; justify-content: space-between; }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <h1>$safeName</h1>
      <div class="subtitle">Mapy.com Route & Navigation Report</div>
    </div>
    <div class="meta">
      Generated: <strong>$genDate</strong><br>
      Engine: <strong>Mapy.com Routing API</strong>
    </div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-label">Total Distance</div>
      <div class="kpi-value blue">$DistanceKm km</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Estimated Duration</div>
      <div class="kpi-value green">$DurationMin min</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Optimization</div>
      <div class="kpi-value">$RouteType</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Avoid Options</div>
      <div style="margin-top:2px;">$avoidHtml</div>
    </div>
  </div>

  $mapImgTag

  <div class="section-title">Route Itinerary & Stops</div>
  <table>
    <thead>
      <tr>
        <th style="width: 35px;">#</th>
        <th style="width: 140px;">Role / Point</th>
        <th>Address / Location</th>
        <th style="width: 160px;">Coordinates (Lat, Lng)</th>
      </tr>
    </thead>
    <tbody>
      $($rowsHtml.ToString())
    </tbody>
  </table>

  <div class="footer">
    <span>Mapy.com Route & Map Generator v2.1</span>
    <span>$(if ($GoogleMapsUrl) { "<a href='$GoogleMapsUrl' style='color:#2563eb;text-decoration:none;'>Open in Mapy.com &rarr;</a>" } else { '' })</span>
  </div>
</body>
</html>
"@

    # Save HTML template using UTF-8 with BOM to preserve Polish/German accents in street addresses
    [System.IO.File]::WriteAllText($tempHtml, $html, [System.Text.UTF8Encoding]::new($true))

    try {
        # Execute headless Chromium print with margins disabled to allow CSS @page rules to govern layout
        $p = Start-Process -FilePath $edgePath -ArgumentList "--headless", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf=`"$OutputPath`"", "`"$tempHtml`"" -Wait -PassThru
        if ($p.ExitCode -eq 0 -and (Test-Path $OutputPath)) {
            if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) { 
                Write-AppLog "Generated PDF report successfully: $OutputPath" "OK" 
            }
            return $OutputPath
        } else {
            throw "Edge process exited with code $($p.ExitCode) and PDF was not created."
        }
    }
    finally {
        # Clean up temporary HTML file to avoid accumulating temp disk artifacts
        if (Test-Path $tempHtml) {
            Remove-Item -Path $tempHtml -Force -ErrorAction SilentlyContinue
        }
    }
}

#endregion 2. Single Route PDF Generation

#region 3. Batch Summary PDF Generation

<#
.SYNOPSIS
    Exports a comprehensive multi-route batch execution summary report to PDF.
.DESCRIPTION
    Aggregates an array of processed route objects into an A4 landscape PDF summary.
    Calculates summary KPIs including total route count, successful vs failed routes,
    accumulated mileage (km), and cumulative driving time, presenting all route legs
    in a structured status table.
.PARAMETER OutputPath
    Target filesystem path where the generated batch summary PDF document will be written.
.PARAMETER Routes
    Array of route PSCustomObjects containing Id, Name, Start_Original, End_Original,
    RouteType, DistanceKm, DurationMin, and Status.
.PARAMETER BatchTitle
    Heading displayed on the summary report. Defaults to 'Batch Processing Summary Report'.
.PARAMETER SourceFileName
    Name of the imported batch dataset file (e.g. 'Routes.xlsx' or 'Fleet.csv') for audit tracking.
.OUTPUTS
    [string] The verified OutputPath of the created batch PDF document.
.EXAMPLE
    Export-BatchPdfReport -OutputPath "C:\Reports\Batch_Summary.pdf" -Routes $batchResults `
        -SourceFileName "TransportOrders.xlsx"
#>
function Export-BatchPdfReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][object[]]$Routes,
        [Parameter()][string]$BatchTitle = 'Batch Processing Summary Report',
        [Parameter()][string]$SourceFileName = ''
    )

    # Locate Microsoft Edge for headless printing
    $edgePath = Find-EdgeExecutable
    if (-not $edgePath) {
        throw "Microsoft Edge executable (msedge.exe) not found. PDF export requires Edge."
    }

    # Ensure parent output directory exists before spawning PDF generation
    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $genDate = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    $totalCount = @($Routes).Count
    $totalDist = 0.0
    $totalDur = 0
    $successCount = 0

    # Build route row elements and compute batch aggregates
    $rowsSb = [System.Text.StringBuilder]::new()
    foreach ($r in @($Routes)) {
        $dist = if ($r.DistanceKm) { [double]$r.DistanceKm } else { 0.0 }
        $dur  = if ($r.DurationMin) { [int]$r.DurationMin } else { 0 }
        $totalDist += $dist
        $totalDur  += $dur
        if ($r.Status -eq 'OK' -or $r.Status -eq 'Success') { $successCount++ }

        # Render status badge: Green for OK/Success, Red for GeocodeFailed/RouteFailed
        $statBadge = if ($r.Status -eq 'OK' -or $r.Status -eq 'Success') {
            "<span class='badge' style='background:#dcfce7;color:#166534;'>OK</span>"
        } else {
            "<span class='badge' style='background:#fee2e2;color:#991b1b;'>$([System.Net.WebUtility]::HtmlEncode([string]$r.Status))</span>"
        }

        $id = [System.Net.WebUtility]::HtmlEncode([string]$r.Id)
        $name = [System.Net.WebUtility]::HtmlEncode([string]$r.Name)
        $start = [System.Net.WebUtility]::HtmlEncode([string]$r.Start_Original)
        $end = [System.Net.WebUtility]::HtmlEncode([string]$r.End_Original)
        $type = [System.Net.WebUtility]::HtmlEncode([string]$r.RouteType)

        $null = $rowsSb.AppendLine("<tr><td>$id</td><td><strong>$name</strong></td><td>$start</td><td>$end</td><td>$type</td><td style='text-align:right;'><strong>$dist km</strong></td><td style='text-align:right;'>$dur min</td><td style='text-align:center;'>$statBadge</td></tr>")
    }
    $totalDistRound = [math]::Round($totalDist, 2)

    # Generate unique temporary HTML container file
    $tempHtml = Join-Path $env:TEMP "gmaps_batch_report_$([System.Guid]::NewGuid().ToString('N')).html"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>$BatchTitle</title>
  <style>
    @page { size: A4 landscape; margin: 12mm 12mm; }
    * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
    body { color: #1e293b; background: #ffffff; font-size: 11px; line-height: 1.4; }
    .header { border-bottom: 2px solid #2563eb; padding-bottom: 8px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: flex-end; }
    .header h1 { font-size: 18px; color: #0f172a; }
    .header .meta { text-align: right; font-size: 10px; color: #64748b; }
    .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 12px; }
    .kpi-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 6px 10px; }
    .kpi-label { font-size: 9px; text-transform: uppercase; color: #64748b; font-weight: 600; }
    .kpi-value { font-size: 15px; font-weight: bold; color: #0f172a; }
    table { width: 100%; border-collapse: collapse; font-size: 10px; }
    th { background: #f1f5f9; text-align: left; padding: 5px 6px; font-weight: 600; color: #475569; border-bottom: 1px solid #cbd5e1; }
    td { padding: 5px 6px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
    tr:nth-child(even) td { background: #fafafa; }
    .badge { display: inline-block; padding: 2px 5px; border-radius: 4px; font-size: 9px; font-weight: bold; }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <h1>$BatchTitle</h1>
      <div style="font-size: 10px; color: #64748b;">Source: $([System.Net.WebUtility]::HtmlEncode($SourceFileName))</div>
    </div>
    <div class="meta">Generated: <strong>$genDate</strong></div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-label">Total Routes</div><div class="kpi-value">$totalCount</div></div>
    <div class="kpi-card"><div class="kpi-label">Successful Routes</div><div class="kpi-value" style="color:#059669;">$successCount / $totalCount</div></div>
    <div class="kpi-card"><div class="kpi-label">Total Distance</div><div class="kpi-value" style="color:#2563eb;">$totalDistRound km</div></div>
    <div class="kpi-card"><div class="kpi-label">Total Travel Time</div><div class="kpi-value">$totalDur min ($([math]::Round($totalDur / 60.0, 1)) h)</div></div>
  </div>

  <table>
    <thead>
      <tr>
        <th style="width: 30px;">#</th>
        <th>Route Name</th>
        <th>Origin</th>
        <th>Destination</th>
        <th style="width: 70px;">Type</th>
        <th style="width: 75px; text-align:right;">Distance</th>
        <th style="width: 65px; text-align:right;">Duration</th>
        <th style="width: 60px; text-align:center;">Status</th>
      </tr>
    </thead>
    <tbody>
      $($rowsSb.ToString())
    </tbody>
  </table>
</body>
</html>
"@

    # Save HTML template using UTF-8 with BOM for headless printing
    [System.IO.File]::WriteAllText($tempHtml, $html, [System.Text.UTF8Encoding]::new($true))

    try {
        # Execute headless Chromium print with landscape A4 profile
        $p = Start-Process -FilePath $edgePath -ArgumentList "--headless", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf=`"$OutputPath`"", "`"$tempHtml`"" -Wait -PassThru
        if ($p.ExitCode -eq 0 -and (Test-Path $OutputPath)) {
            if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) { 
                Write-AppLog "Generated batch PDF report successfully: $OutputPath" "OK" 
            }
            return $OutputPath
        } else {
            throw "Edge process exited with code $($p.ExitCode) and PDF was not created."
        }
    }
    finally {
        # Clean up temporary HTML file
        if (Test-Path $tempHtml) {
            Remove-Item -Path $tempHtml -Force -ErrorAction SilentlyContinue
        }
    }
}

#endregion 3. Batch Summary PDF Generation

#region 4. Global Function Exports

# Export functions into global scope so caller scripts and GUI orchestrators can invoke them
Set-Item -Path "function:global:Find-EdgeExecutable" -Value (Get-Item "function:Find-EdgeExecutable").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Export-RoutePdfReport" -Value (Get-Item "function:Export-RoutePdfReport").ScriptBlock -ErrorAction SilentlyContinue
Set-Item -Path "function:global:Export-BatchPdfReport" -Value (Get-Item "function:Export-BatchPdfReport").ScriptBlock -ErrorAction SilentlyContinue

#endregion 4. Global Function Exports
