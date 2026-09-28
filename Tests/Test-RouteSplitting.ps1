#Requires -Version 5.1
<#
.SYNOPSIS
    Automated regression test suite for proportional route waypoint splitting and polyline codec.
.DESCRIPTION
    Validates:
      1. Split-RoutePoints algorithm with various point counts (2, 3, 17, 18, 24, 33, 34, 50, 100)
      2. Continuity invariant: Part[k].End == Part[k+1].Start
      3. Capacity invariant: MaxPointsPerPart <= 17 (Waypoints <= 15)
      4. Proportional invariant: max(legs) - min(legs) <= 1
      5. Import-RouteDataFile automatic splitting for SequentialStops and RouteList
      6. Polyline encode/decode roundtrip
      7. Trilingual localization catalog key parity (EN, DE, PL)
.NOTES
    Encoding: UTF-8 with BOM
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $PSScriptRoot
if (-not $ScriptRoot) { $ScriptRoot = 'd:\Skrypty\MapyComRoutes' }

. "$ScriptRoot\RouteMapFunctions.ps1"

$PassCount = 0
$FailCount = 0

function Assert-Condition {
    param(
        [bool]$Condition,
        [string]$TestName
    )
    if ($Condition) {
        Write-Host "  [PASS] $TestName" -ForegroundColor Green
        $script:PassCount++
    }
    else {
        Write-Host "  [FAIL] $TestName" -ForegroundColor Red
        $script:FailCount++
    }
}

Write-Host "`n=== 1. SPLIT-ROUTEPOINTS ALGORITHM TESTS ===" -ForegroundColor Cyan

# Test 1: 2 points (Start and End only)
$p2 = Split-RoutePoints -Points @('Start', 'End')
Assert-Condition ($p2.Count -eq 1) "2 points -> 1 part"
Assert-Condition ($p2[0].PointsCount -eq 2 -and $p2[0].LegsCount -eq 1) "2 points has 2 points and 1 leg"
Assert-Condition ($p2[0].Start -eq 'Start' -and $p2[0].End -eq 'End') "2 points start and end preserved"

# Test 2: 3 points with max 2 points/part (User exact example)
$p3 = Split-RoutePoints -Points @('address a', 'address b', 'address c') -MaxPointsPerPart 2
Assert-Condition ($p3.Count -eq 2) "3 points with max 2 -> 2 parts"
Assert-Condition ($p3[0].Start -eq 'address a' -and $p3[0].End -eq 'address b') "Part 1: address a -> address b"
Assert-Condition ($p3[1].Start -eq 'address b' -and $p3[1].End -eq 'address c') "Part 2: address b -> address c"
Assert-Condition ($p3[0].End -eq $p3[1].Start) "Endpoint of Part 1 is start point of Part 2"

# Test 3: 17 points (maximum single route: 1 start + 15 intermediate + 1 end)
$pts17 = 1..17 | ForEach-Object { "Stop_$_" }
$p17 = Split-RoutePoints -Points $pts17 -MaxPointsPerPart 17
Assert-Condition ($p17.Count -eq 1) "17 points -> 1 part (no split needed)"
Assert-Condition ($p17[0].WaypointsCount -eq 15) "17 points has exactly 15 intermediate waypoints"

# Test 4: 18 points (exceeds 17 by 1 -> must split proportionally into 2 parts)
$pts18 = 1..18 | ForEach-Object { "Stop_$_" }
$p18 = Split-RoutePoints -Points $pts18 -MaxPointsPerPart 17
Assert-Condition ($p18.Count -eq 2) "18 points -> 2 parts"
Assert-Condition ($p18[0].LegsCount -eq 9 -and $p18[1].LegsCount -eq 8) "18 points split proportionally: 9 legs and 8 legs"
Assert-Condition ($p18[0].End -eq $p18[1].Start) "18 points continuity: Part 1 End == Part 2 Start ($($p18[0].End))"

# Test 5: 24 points (User exact example: start to end 24 points so split in half etc.)
$pts24 = 1..24 | ForEach-Object { "Stop_$_" }
$p24 = Split-RoutePoints -Points $pts24 -MaxPointsPerPart 17
Assert-Condition ($p24.Count -eq 2) "24 points -> 2 parts"
Assert-Condition ($p24[0].LegsCount -eq 12 -and $p24[1].LegsCount -eq 11) "24 points split in half: 12 legs and 11 legs"
Assert-Condition ($p24[0].PointsCount -eq 13 -and $p24[1].PointsCount -eq 12) "24 points points count: 13 and 12 (both <= 17)"
Assert-Condition ($p24[0].WaypointsCount -eq 11 -and $p24[1].WaypointsCount -eq 10) "24 points waypoints: 11 and 10 (both <= 15)"
Assert-Condition ($p24[0].End -eq 'Stop_13' -and $p24[1].Start -eq 'Stop_13') "24 points continuity: split point is Stop_13"

# Test 6: 33 points (maximum for 2 parts: 2 * 16 = 32 legs)
$pts33 = 1..33 | ForEach-Object { "Stop_$_" }
$p33 = Split-RoutePoints -Points $pts33 -MaxPointsPerPart 17
Assert-Condition ($p33.Count -eq 2) "33 points -> 2 parts"
Assert-Condition ($p33[0].LegsCount -eq 16 -and $p33[1].LegsCount -eq 16) "33 points split evenly: 16 legs and 16 legs"
Assert-Condition ($p33[0].PointsCount -eq 17 -and $p33[1].PointsCount -eq 17) "33 points both parts have exactly 17 points"

# Test 7: 34 points (requires 3 parts)
$pts34 = 1..34 | ForEach-Object { "Stop_$_" }
$p34 = Split-RoutePoints -Points $pts34 -MaxPointsPerPart 17
Assert-Condition ($p34.Count -eq 3) "34 points -> 3 parts"
Assert-Condition ($p34[0].LegsCount -eq 11 -and $p34[1].LegsCount -eq 11 -and $p34[2].LegsCount -eq 11) "34 points split evenly: 11, 11, 11 legs"
Assert-Condition ($p34[0].End -eq $p34[1].Start -and $p34[1].End -eq $p34[2].Start) "34 points continuous chain through all 3 parts"

# Test 8: 100 points
$pts100 = 1..100 | ForEach-Object { "Stop_$_" }
$p100 = Split-RoutePoints -Points $pts100 -MaxPointsPerPart 17
Assert-Condition ($p100.Count -eq 7) "100 points -> 7 parts"
$allUnderLimit = $true
$continuous = $true
$maxDiffProportional = $true
$minLegs = 999
$maxLegs = 0
for ($i = 0; $i -lt $p100.Count; $i++) {
    if ($p100[$i].PointsCount -gt 17 -or $p100[$i].WaypointsCount -gt 15) { $allUnderLimit = $false }
    if ($p100[$i].LegsCount -lt $minLegs) { $minLegs = $p100[$i].LegsCount }
    if ($p100[$i].LegsCount -gt $maxLegs) { $maxLegs = $p100[$i].LegsCount }
    if ($i -gt 0 -and $p100[$i].Start -ne $p100[$i - 1].End) { $continuous = $false }
}
Assert-Condition $allUnderLimit "100 points: all 7 parts have <= 17 points and <= 15 waypoints"
Assert-Condition $continuous "100 points: continuous chaining across all 7 parts"
Assert-Condition (($maxLegs - $minLegs) -le 1) "100 points: strictly proportional (legs differ by at most 1: min $minLegs, max $maxLegs)"

# Test 9: Endpoints parameter set (-StartPoint, -EndPoint, -Waypoints)
$wp22 = 2..23 | ForEach-Object { "WP_$_" }
$pEndpoints = Split-RoutePoints -StartPoint "Origin_A" -EndPoint "Dest_B" -Waypoints $wp22 -MaxPointsPerPart 17
Assert-Condition ($pEndpoints.Count -eq 2) "Endpoints param set: 24 total points -> 2 parts"
Assert-Condition ($pEndpoints[0].Start -eq 'Origin_A' -and $pEndpoints[1].End -eq 'Dest_B') "Endpoints param set: preserved overall start and destination"
Assert-Condition ($pEndpoints[0].End -eq $pEndpoints[1].Start) "Endpoints param set: shared midpoint ($($pEndpoints[0].End))"

Write-Host "`n=== 2. IMPORT-ROUTEDATAFILE AUTO-SPLIT TESTS ===" -ForegroundColor Cyan

# Test SequentialStops auto-split
$tempCsv = Join-Path ([System.IO.Path]::GetTempPath()) "test_seq_stops_24.csv"
$csvRows = [System.Collections.Generic.List[string]]::new()
$csvRows.Add("Trasa;Kolejnosc;Adres")
for ($i = 1; $i -le 24; $i++) {
    $csvRows.Add("Trip1;$i;Address $i, City")
}
[System.IO.File]::WriteAllLines($tempCsv, $csvRows, [System.Text.UTF8Encoding]::new($true))

$imported = Import-RouteDataFile -Path $tempCsv
Assert-Condition ($imported.Mode -eq 'SequentialStops') "Imported mode is SequentialStops"
Assert-Condition ($imported.Routes.Count -eq 2) "24 sequential stops auto-split into 2 routes"
Assert-Condition ($imported.Routes[0].Name -match 'Part 1/2' -and $imported.Routes[1].Name -match 'Part 2/2') "Route names indicate Part 1/2 and Part 2/2"
Assert-Condition ($imported.Routes[0].End -eq $imported.Routes[1].Start) "SequentialStops split continuity: Route 1 End == Route 2 Start ($($imported.Routes[0].End))"
Assert-Condition ($imported.Routes[0].Waypoints.Count -eq 11 -and $imported.Routes[1].Waypoints.Count -eq 10) "SequentialStops waypoints: 11 and 10 (both <= 15)"

Remove-Item -LiteralPath $tempCsv -Force -ErrorAction SilentlyContinue

# Test RouteList auto-split
$tempCsvList = Join-Path ([System.IO.Path]::GetTempPath()) "test_route_list_24.csv"
$wpJoined = ($wp22 -join '|')
$csvListContent = @"
Nazwa;Start;Koniec;PunktyPosrednie
LongRoute;Start Address;End Address;$wpJoined
"@
[System.IO.File]::WriteAllText($tempCsvList, $csvListContent, [System.Text.UTF8Encoding]::new($true))

$importList = Import-RouteDataFile -Path $tempCsvList
Assert-Condition ($importList.Mode -eq 'RouteList') "Imported mode is RouteList"
Assert-Condition ($importList.Routes.Count -eq 2) "Route with 22 waypoints auto-split into 2 routes"
Assert-Condition ($importList.Routes[0].End -eq $importList.Routes[1].Start) "RouteList split continuity: Route 1 End == Route 2 Start"
Assert-Condition ($importList.Routes[0].Waypoints.Count -le 15 -and $importList.Routes[1].Waypoints.Count -le 15) "RouteList: both parts have <= 15 waypoints"

Remove-Item -LiteralPath $tempCsvList -Force -ErrorAction SilentlyContinue

Write-Host "`n=== 3. POLYLINE CODEC TESTS ===" -ForegroundColor Cyan

$testCoords = @(
    [GoogleMapsPoint]::new(52.2297, 21.0122),
    [GoogleMapsPoint]::new(51.4027, 21.1471),
    [GoogleMapsPoint]::new(50.0647, 19.9450)
)
$encoded = ConvertTo-EncodedPolyline -Points $testCoords
Assert-Condition (-not [string]::IsNullOrWhiteSpace($encoded)) "ConvertTo-EncodedPolyline generated polyline: $encoded"

$decoded = ConvertFrom-EncodedPolyline -EncodedPolyline $encoded
Assert-Condition ($decoded.Count -eq 3) "ConvertFrom-EncodedPolyline decoded 3 points"
$latDiff = [math]::Abs($decoded[0].Latitude - 52.2297)
$lngDiff = [math]::Abs($decoded[0].Longitude - 21.0122)
Assert-Condition ($latDiff -lt 0.0001 -and $lngDiff -lt 0.0001) "Decoded coordinates match original within precision"

Write-Host "`n=== 4. TRILINGUAL LOCALIZATION CATALOG TESTS ===" -ForegroundColor Cyan

$locPath = Join-Path $ScriptRoot 'localization.json'
$loc = Get-Content -LiteralPath $locPath -Raw -Encoding UTF8 | ConvertFrom-Json
$enKeys = @($loc.Languages.en.Strings.psobject.Properties.Name) | Sort-Object
$deKeys = @($loc.Languages.de.Strings.psobject.Properties.Name) | Sort-Object
$plKeys = @($loc.Languages.pl.Strings.psobject.Properties.Name) | Sort-Object

$diffDE = Compare-Object $enKeys $deKeys
$diffPL = Compare-Object $enKeys $plKeys
Assert-Condition ($null -eq $diffDE -or $diffDE.Count -eq 0) "100% key match between EN and DE ($($enKeys.Count) keys)"
Assert-Condition ($null -eq $diffPL -or $diffPL.Count -eq 0) "100% key match between EN and PL ($($enKeys.Count) keys)"

Write-Host "`n==============================================="
Write-Host "TEST RESULTS: $PassCount PASSED, $FailCount FAILED" -ForegroundColor $(if ($FailCount -eq 0) { 'Green' } else { 'Red' })
Write-Host "===============================================`n"

if ($FailCount -gt 0) {
    exit 1
}
