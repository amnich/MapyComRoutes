#Requires -Version 5.1
<#
.SYNOPSIS
    Mapy.com Routes & Map Generator — Manual Route Tab UI Handlers.
.DESCRIPTION
    Manages waypoint list operations, route optimization settings, avoid options,
    manual route calculation orchestration, WebView2/PNG map preview switching,
    and single-route exports (PDF, GPX, KML).
.NOTES
    Encoding: UTF-8 with BOM
    Compatibility: Windows PowerShell 5.1 and PowerShell 7+
#>

$script:LastManualResult = $null
$script:ActiveManualRouteName = ''

#region 1. Map View Mode Toggling

<#
.SYNOPSIS
    Toggles visibility between the interactive WebView2 map container and static PNG preview image.
.DESCRIPTION
    Inspects radio buttons rbViewInteractive and rbViewStatic, switching between
    pnlInteractiveMapHost (Chromium Leaflet.js) and imgMapPreview (GDI+ static PNG).
    Displays the placeholder label when no route calculation result is available.
.OUTPUTS
    None.
.EXAMPLE
    Update-MapViewMode
#>
function Update-MapViewMode {
    [CmdletBinding()]
    param()

    $rbInteractive  = if ($script:Controls) { $script:Controls.rbViewInteractive } else { $null }
    $imgPreview     = if ($script:Controls) { $script:Controls.imgMapPreview } else { $null }
    $pnlHost        = if ($script:Controls) { $script:Controls.pnlInteractiveMapHost } else { $null }
    $lblPlaceholder = if ($script:Controls) { $script:Controls.lblMapPlaceholder } else { $null }

    # If no route has been calculated yet, display initial placeholder prompt
    if (-not $script:LastManualResult) {
        if ($lblPlaceholder) { $lblPlaceholder.Visibility = [System.Windows.Visibility]::Visible }
        if ($imgPreview)     { $imgPreview.Visibility     = [System.Windows.Visibility]::Collapsed }
        if ($pnlHost)        { $pnlHost.Visibility        = [System.Windows.Visibility]::Collapsed }
        return
    }

    if ($lblPlaceholder) { $lblPlaceholder.Visibility = [System.Windows.Visibility]::Collapsed }

    # Toggle between Chromium vector interactive Leaflet map and static GDI+ PNG
    if ($rbInteractive -and $rbInteractive.IsChecked) {
        if ($imgPreview) { $imgPreview.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($pnlHost)    { $pnlHost.Visibility    = [System.Windows.Visibility]::Visible }
    } else {
        if ($pnlHost)    { $pnlHost.Visibility    = [System.Windows.Visibility]::Collapsed }
        if ($imgPreview) { $imgPreview.Visibility = [System.Windows.Visibility]::Visible }
    }
}
Set-Item -Path "function:global:Update-MapViewMode" -Value (Get-Item "function:Update-MapViewMode").ScriptBlock -ErrorAction SilentlyContinue

#endregion 1. Map View Mode Toggling

#region 2. Manual Tab Event Registration & Wiring

<#
.SYNOPSIS
    Registers and binds all user interface events and handlers for the Manual Route tab.
.DESCRIPTION
    Binds event handlers for waypoint controls (add, remove, clear, reorder up/down),
    endpoint swapping, route calculation invocation via background MTA runspace,
    autosuggest address completion popups, avoid toll/highway checkboxes, and
    single-route export actions (PDF reports, GPX, KML, static PNG, and complete zip package).
.PARAMETER Controls
    Hashtable containing mapped WPF UI controls instantiated from XAML.
.PARAMETER Window
    Optional reference to the main application Window.
.OUTPUTS
    None.
.EXAMPLE
    Register-UiManualTabEvents -Controls $Controls -Window $script:MainWindow
#>
function Register-UiManualTabEvents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Controls,
        [Parameter(Mandatory = $false)][System.Windows.Window]$Window = $null
    )

    $script:Controls = $Controls
    $script:SuppressAutosuggest = $false
    foreach ($k in $Controls.Keys) {
        Set-Variable -Name $k -Value $Controls[$k] -Scope Script -Force
        Set-Variable -Name "script:$k" -Value $Controls[$k] -Scope Script -Force
    }

    $aliases = [ordered]@{
        'btnClearStart'      = $Controls.btnClearManualStart
        'btnClearEnd'        = $Controls.btnClearManualEnd
        'txtStart'           = $Controls.txtManualStart
        'txtEnd'             = $Controls.txtManualEnd
        'txtNewWp'           = $Controls.txtNewWaypoint
        'btnAddWp'           = $Controls.btnAddWaypoint
        'lstWp'              = $Controls.lstWaypoints
        'btnWpUp'            = $Controls.btnWpUp
        'btnWpDown'          = $Controls.btnWpDown
        'btnWpRemove'        = $Controls.btnWpRemove
        'btnWpClear'         = $Controls.btnWpClear
        'txtName'            = $Controls.txtManualName
        'rbFastest'          = $Controls.rbTypeFastest
        'rbShortest'         = $Controls.rbTypeShortest
        'rbEco'              = $Controls.rbTypeEco
        'pnlEmission'        = $Controls.pnlEmission
        'cmbEmission'        = $Controls.cmbEmission
        'chkTrafficAware'    = $Controls.chkTrafficAware
        'chkAvoidTolls'      = $Controls.chkManualAvoidTolls
        'chkAvoidHighways'   = $Controls.chkManualAvoidHighways
        'chkAvoidFerries'    = $Controls.chkManualAvoidFerries
        'btnCalculate'       = $Controls.btnCalculateManual
        'lblDist'            = $Controls.lblManualDist
        'lblTime'            = $Controls.lblManualTime
        'lblType'            = $Controls.lblManualType
        'lblStatus'          = $Controls.lblManualStatus
        'lblPlaceholder'     = $Controls.lblMapPlaceholder
        'imgPreview'         = $Controls.imgMapPreview
        'pnlInteractiveHost' = $Controls.pnlInteractiveMapHost
        'rbViewInteractive'  = $Controls.rbViewInteractive
        'rbViewStatic'       = $Controls.rbViewStatic
        'lblUrlDisplay'      = $Controls.lblGoogleUrlDisplay
        'btnOpenMaps'        = $Controls.btnOpenGoogleMaps
        'btnCopyUrl'         = $Controls.btnCopyUrl
        'btnSaveMapAs'       = $Controls.btnSaveMapAs
        'btnExportPdf'           = $Controls.btnManualExportPdf
        'btnExportGpx'           = $Controls.btnManualExportGpx
        'btnExportKml'           = $Controls.btnManualExportKml
        'btnManualExportPackage' = $Controls.btnManualExportPackage
        'lblFooter'              = $Controls.lblFooterStatus
        'txtOutputDir'           = $Controls.txtSettingsOutputDir
        'tabMain'                = $Controls.tabMain
        'btnSwapEndpoints'       = $Controls.btnSwapEndpoints
        'btnResetManualForm'     = $Controls.btnResetManualForm
        'cmbRecentRoutes'        = $Controls.cmbRecentRoutes
        'lblManualRecentRoutes'  = $Controls.lblManualRecentRoutes
        'lblCopyFeedback'        = $Controls.lblCopyFeedback
        'btnToggleInputPanel'    = $Controls.btnToggleInputPanel
        'colManualInput'         = $Controls.colManualInput
        'badgeStartGeocode'      = $Controls.badgeStartGeocode
        'badgeEndGeocode'        = $Controls.badgeEndGeocode
        'popSuggestStart'        = $Controls.popSuggestStart
        'lstSuggestStart'        = $Controls.lstSuggestStart
        'popSuggestEnd'          = $Controls.popSuggestEnd
        'lstSuggestEnd'          = $Controls.lstSuggestEnd
        'popSuggestWp'           = $Controls.popSuggestWp
        'lstSuggestWp'           = $Controls.lstSuggestWp
        'drawerLog'              = $Controls.drawerLog
        'txtLogDrawer'           = $Controls.txtLogDrawer
        'btnToggleLogDrawer'     = $Controls.btnToggleLogDrawer
        'btnCloseLogDrawer'      = $Controls.btnCloseLogDrawer
        'btnClearLogDrawer'      = $Controls.btnClearLogDrawer
        'btnCopyLogDrawer'       = $Controls.btnCopyLogDrawer
        'rbLogAll'               = $Controls.rbLogAll
        'rbLogInfo'              = $Controls.rbLogInfo
        'rbLogWarn'              = $Controls.rbLogWarn
        'rbLogError'             = $Controls.rbLogError
    }

    foreach ($entry in $aliases.GetEnumerator()) {
        Set-Variable -Name $entry.Key -Value $entry.Value -Scope Script -Force
        Set-Variable -Name "script:$($entry.Key)" -Value $entry.Value -Scope Script -Force
    }

    $btnClearStart      = $aliases['btnClearStart']
    $btnClearEnd        = $aliases['btnClearEnd']
    $txtStart           = $aliases['txtStart']
    $txtEnd             = $aliases['txtEnd']
    $txtNewWp           = $aliases['txtNewWp']
    $btnAddWp           = $aliases['btnAddWp']
    $lstWp              = $aliases['lstWp']
    $btnWpUp            = $aliases['btnWpUp']
    $btnWpDown          = $aliases['btnWpDown']
    $btnWpRemove        = $aliases['btnWpRemove']
    $btnWpClear         = $aliases['btnWpClear']
    $txtName            = $aliases['txtName']
    $rbFastest          = $aliases['rbFastest']
    $rbShortest         = $aliases['rbShortest']
    $rbEco              = $aliases['rbEco']
    $pnlEmission        = $aliases['pnlEmission']
    $cmbEmission        = $aliases['cmbEmission']
    $chkTrafficAware    = $aliases['chkTrafficAware']
    $chkAvoidTolls      = $aliases['chkAvoidTolls']
    $chkAvoidHighways   = $aliases['chkAvoidHighways']
    $chkAvoidFerries    = $aliases['chkAvoidFerries']
    $btnCalculate       = $aliases['btnCalculate']
    $lblDist            = $aliases['lblDist']
    $lblTime            = $aliases['lblTime']
    $lblType            = $aliases['lblType']
    $lblStatus          = $aliases['lblStatus']
    $lblPlaceholder     = $aliases['lblPlaceholder']
    $imgPreview         = $aliases['imgPreview']
    $pnlInteractiveHost = $aliases['pnlInteractiveHost']
    $rbViewInteractive  = $aliases['rbViewInteractive']
    $rbViewStatic       = $aliases['rbViewStatic']
    $lblUrlDisplay      = $aliases['lblUrlDisplay']
    $btnOpenMaps        = $aliases['btnOpenMaps']
    $btnCopyUrl         = $aliases['btnCopyUrl']
    $btnSaveMapAs       = $aliases['btnSaveMapAs']
    $btnExportPdf       = $aliases['btnExportPdf']
    $btnExportGpx       = $aliases['btnExportGpx']
    $btnExportKml       = $aliases['btnExportKml']
    $lblFooter          = $aliases['lblFooter']
    $txtOutputDir       = $aliases['txtOutputDir']
    $tabMain            = $aliases['tabMain']

    # Clear address inputs
    $btnClearStart.Add_Click({ $txtStart.Clear() })
    $btnClearEnd.Add_Click({ $txtEnd.Clear() })

    # Waypoints list management
    $btnAddWp.Add_Click({
        $wp = $txtNewWp.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($wp)) {
            if ($lstWp.Items.Count -ge 25) {
                [System.Windows.MessageBox]::Show((Get-LocText 'MsgMaxWaypoints' 'Maximum 25 waypoints allowed.'), (Get-LocText 'MsgMaxWaypointsTitle' 'Waypoints Limit'), 'OK', 'Warning')
                return
            }
            $null = $lstWp.Items.Add($wp)
            $txtNewWp.Clear()
        }
    })

    $txtNewWp.Add_KeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
            $btnAddWp.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })

    $btnWpRemove.Add_Click({
        if ($lstWp.SelectedIndex -ge 0) { $lstWp.Items.RemoveAt($lstWp.SelectedIndex) }
    })

    $btnWpClear.Add_Click({ $lstWp.Items.Clear() })

    $btnWpUp.Add_Click({
        $idx = $lstWp.SelectedIndex
        if ($idx -gt 0) {
            $item = $lstWp.Items[$idx]
            $lstWp.Items.RemoveAt($idx)
            $lstWp.Items.Insert($idx - 1, $item)
            $lstWp.SelectedIndex = $idx - 1
        }
    })

    $btnWpDown.Add_Click({
        $idx = $lstWp.SelectedIndex
        if ($idx -ge 0 -and $idx -lt ($lstWp.Items.Count - 1)) {
            $item = $lstWp.Items[$idx]
            $lstWp.Items.RemoveAt($idx)
            $lstWp.Items.Insert($idx + 1, $item)
            $lstWp.SelectedIndex = $idx + 1
        }
    })

    # Optimization mode toggles
    $rbEco.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Visible })
    $rbFastest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })
    $rbShortest.Add_Checked({ $pnlEmission.Visibility = [System.Windows.Visibility]::Collapsed })

    # Interactive vs Static Map Mode Toggle (Feature 4.K)
    $rbViewInteractive.Add_Checked({ Update-MapViewMode })
    $rbViewStatic.Add_Checked({ Update-MapViewMode })

    # Calculate Route Handler
    $btnCalculate.Add_Click({
        $apiKey = if (Get-Command Get-CurrentApiKey -ErrorAction SilentlyContinue) { Get-CurrentApiKey } else { $script:AppConfig.ApiKey }
        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingApiKeyPrompt' 'Google Maps API Key is required.'), (Get-LocText 'MsgMissingApiKeyTitle' 'Missing API Key'), 'OK', 'Warning')
            if ($tabMain) { $tabMain.SelectedIndex = 2 }
            elseif ($script:Controls -and $script:Controls.tabMain) { $script:Controls.tabMain.SelectedIndex = 2 }
            return
        }

        $start = $txtStart.Text.Trim()
        $end = $txtEnd.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($start) -or [string]::IsNullOrWhiteSpace($end)) {
            [System.Windows.MessageBox]::Show((Get-LocText 'MsgMissingData' 'Please provide both Origin and Destination.'), (Get-LocText 'MsgMissingDataTitle' 'Missing Data'), 'OK', 'Warning')
            return
        }

        $waypoints = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $lstWp.Items) { $waypoints.Add([string]$item) }

        $routeType = if ($rbShortest.IsChecked) { 'Shortest' } elseif ($rbEco.IsChecked) { 'Eco' } else { 'Fastest' }
        $emission = if ($cmbEmission -and $cmbEmission.SelectedItem) { ($cmbEmission.SelectedItem.Tag -as [string]) } else { 'GASOLINE' }
        if ([string]::IsNullOrWhiteSpace($emission)) { $emission = 'GASOLINE' }
        $trafficAware = [bool]$chkTrafficAware.IsChecked
        $avoidTolls = [bool]$chkAvoidTolls.IsChecked
        $avoidHighways = [bool]$chkAvoidHighways.IsChecked
        $avoidFerries = [bool]$chkAvoidFerries.IsChecked

        $name = $txtName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { $name = "Route $start -> $end" }
        $script:ActiveManualRouteName = $name

        $outDir = $txtOutputDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'TrasyMapyCom' }
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

        $btnCalculate.IsEnabled = $false
        $btnCalculate.Content = '⏳ CALCULATING ROUTE...'
        $lblStatus.Text = 'Geocoding and calculating...'
        $lblStatus.Foreground = [System.Windows.Media.Brushes]::SkyBlue
        $lblFooter.Text = 'Calculating manual route...'

        Write-AppLog "Started manual route calculation: '$start' -> '$end' (Waypoints: $($waypoints.Count), Type: $routeType, Engine: $emission, AvoidTolls: $avoidTolls, AvoidHighways: $avoidHighways, AvoidFerries: $avoidFerries)..." "INFO"

        $psCmd = New-WorkerPowerShell -ScriptBlock $script:ManualCalcAsync
        $overlayCfgJson = ((Get-CurrentOverlayConfig -Controls $Controls) | ConvertTo-Json -Depth 6 -Compress)
        $langCode = if ($script:CurrentLanguage) { $script:CurrentLanguage } else { 'en' }

        $psCmd.AddArgument($start).AddArgument($end).AddArgument($waypoints).AddArgument($routeType).AddArgument($emission).AddArgument($trafficAware).AddArgument($name).AddArgument($apiKey).AddArgument($outDir).AddArgument($script:LogFile).AddArgument($langCode).AddArgument($overlayCfgJson).AddArgument($avoidTolls).AddArgument($avoidHighways).AddArgument($avoidFerries) | Out-Null

        try {
            $asyncHandle = $psCmd.BeginInvoke()
        }
        catch {
            Write-AppLog "CRITICAL: BeginInvoke() failed: $($_.Exception.Message)" "ERROR"
            $btnCalculate.IsEnabled = $true
            $btnCalculate.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
            $lblStatus.Text = '✕ Launch error'
            $lblStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
            $lblFooter.Text = "Error: $($_.Exception.Message)"
            return
        }

        $timer = [System.Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(150)
        $script:ActiveManualTimer = $timer
        $script:ActiveManualPs = $psCmd
        $script:ActiveManualAsyncHandle = $asyncHandle
        $script:ManualTimerTicks = 0

        $timer.Add_Tick({
            $localHandle = $script:ActiveManualAsyncHandle
            $localPs = $script:ActiveManualPs
            $script:ManualTimerTicks++

            if ($localHandle -and $localHandle.IsCompleted) {
                if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
                $btnCalculate.IsEnabled = $true
                $btnCalculate.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'

                foreach ($streamErr in $localPs.Streams.Error) {
                    Write-AppLog "[Stream.Error] $($streamErr.Exception.Message) @ $($streamErr.InvocationInfo.PositionMessage)" "ERROR"
                }

                try {
                    $res = $localPs.EndInvoke($localHandle)
                    $calc = $res[0]
                    if ($calc.Success) {
                        $script:LastManualResult = $calc
                        $lblDist.Text = "$($calc.DistanceKm) km"
                        $lblTime.Text = "$($calc.DurationMin) min"
                        $lblType.Text = switch ($script:CurrentLanguage) {
                            'de' { if ($calc.RouteType -eq 'Fastest') { 'Schnellste' } elseif ($calc.RouteType -eq 'Shortest') { 'Kürzeste' } else { 'Eco' } }
                            'pl' { if ($calc.RouteType -eq 'Fastest') { 'Najszybsza' } elseif ($calc.RouteType -eq 'Shortest') { 'Najkrótsza' } else { 'Eko' } }
                            default { [string]$calc.RouteType }
                        }
                        $lblStatus.Text = '✓ Success'
                        $lblStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                        $lblFooter.Text = "Route ready: $($calc.DistanceKm) km, $($calc.DurationMin) min"

                        if ($calc.ApiUsage) {
                            Update-ApiUsageRecord -GeocodingInc $calc.ApiUsage.Geocoding -RoutesInc $calc.ApiUsage.Routes -StaticMapsInc $calc.ApiUsage.StaticMaps
                            Update-ApiUsageBadgeText
                        }

                        $script:LastGoogleMapsUrl = if ($calc.MapyComUrl) { $calc.MapyComUrl } else { $calc.GoogleMapsUrl }
                        $lblUrlDisplay.Text = if ($script:LastGoogleMapsUrl) { $script:LastGoogleMapsUrl } else { 'Generated route ready' }
                        $btnOpenMaps.IsEnabled = $true
                        $btnCopyUrl.IsEnabled = $true
                        $btnExportPdf.IsEnabled = $true
                        $btnExportGpx.IsEnabled = $true
                        $btnExportKml.IsEnabled = $true

                        $lblPlaceholder.Visibility = [System.Windows.Visibility]::Collapsed

                        # Load Static Map PNG
                        if ($calc.MapPath -and (Test-Path $calc.MapPath)) {
                            $script:LastGeneratedMapPath = $calc.MapPath
                            $btnSaveMapAs.IsEnabled = $true

                            $imgBytes = [System.IO.File]::ReadAllBytes($calc.MapPath)
                            $ms = [System.IO.MemoryStream]::new($imgBytes)
                            $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
                            $bi.BeginInit()
                            $bi.StreamSource = $ms
                            $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                            $bi.EndInit()
                            $bi.Freeze()
                            $imgPreview.Source = $bi
                        }

                        # Load Interactive Map (Feature 4.K)
                        if ($calc.EncodedPolyline) {
                            Update-InteractiveRouteMap -CalcResult $calc
                        }

                        Update-MapViewMode

                        # Save to Recent Routes (Milestone 1)
                        try {
                            Add-RecentRouteConfig -Start $txtStart.Text -End $txtEnd.Text -Waypoints @($lstWp.Items) -RouteName $txtName.Text -RouteType $calc.RouteType
                            if ($PopulateRecentRoutes) { & $PopulateRecentRoutes }
                            if ($btnManualExportPackage) { $btnManualExportPackage.IsEnabled = $true }
                        } catch { }

                        # Update geocode badges
                        if ($Controls.badgeStartGeocode) {
                            $Controls.badgeStartGeocode.Text = '🟢 Exact'
                            $Controls.badgeStartGeocode.Foreground = [System.Windows.Media.Brushes]::LimeGreen
                        }
                        if ($Controls.badgeEndGeocode) {
                            $Controls.badgeEndGeocode.Text = '🟢 Exact'
                            $Controls.badgeEndGeocode.Foreground = [System.Windows.Media.Brushes]::LimeGreen
                        }
                    }
                    else {
                        $lblStatus.Text = '✕ Error'
                        $lblStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                        $lblFooter.Text = "Error: $($calc.Error)"
                        if ($Controls.badgeStartGeocode) {
                            $Controls.badgeStartGeocode.Text = '🔴 Error'
                            $Controls.badgeStartGeocode.Foreground = [System.Windows.Media.Brushes]::Salmon
                        }
                        if ($Controls.badgeEndGeocode) {
                            $Controls.badgeEndGeocode.Text = '🔴 Error'
                            $Controls.badgeEndGeocode.Foreground = [System.Windows.Media.Brushes]::Salmon
                        }
                        Show-AppToastNotification -Title "Route Error" -Message $calc.Error -Type Error
                    }
                }
                catch {
                    $lblStatus.Text = '✕ Error'
                    $lblStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                    $lblFooter.Text = "Exception: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error')
                }
                finally {
                    $localPs.Dispose()
                }
            }
            elseif ($script:ManualTimerTicks -ge 400) {
                if ($script:ActiveManualTimer) { try { $script:ActiveManualTimer.Stop() } catch { } }
                $btnCalculate.IsEnabled = $true
                $btnCalculate.Content = '🚀 CALCULATE ROUTE & DOWNLOAD MAP'
                $lblStatus.Text = '✕ Timeout (60s)'
                $lblStatus.Foreground = [System.Windows.Media.Brushes]::Salmon
                $lblFooter.Text = 'Route calculation timed out (60s).'
                try { $localPs.Stop(); $localPs.Dispose() } catch { }
            }
        })
        $timer.Start()
    })

    # Milestone 1: Return-trip Swap Origin and Destination
    if ($btnSwapEndpoints) {
        $btnSwapEndpoints.Add_Click({
            $script:SuppressAutosuggest = $true
            try {
                $temp = $txtStart.Text
                $txtStart.Text = $txtEnd.Text
                $txtEnd.Text = $temp

                if ($lstWp.Items.Count -gt 1) {
                    $items = @($lstWp.Items)
                    [array]::Reverse($items)
                    $lstWp.Items.Clear()
                    foreach ($it in $items) { $lstWp.Items.Add($it) | Out-Null }
                }

                if ($txtName.Text -match '^(?:Route|Trasa)\s+(.+)\s+-\s+(.+)$') {
                    $txtName.Text = "Route $($txtStart.Text) - $($txtEnd.Text)"
                }
                if ($Controls.badgeStartGeocode -and $Controls.badgeEndGeocode) {
                    $tempBadgeText = $Controls.badgeStartGeocode.Text
                    $tempBadgeFg   = $Controls.badgeStartGeocode.Foreground
                    $tempBadgeTip  = $Controls.badgeStartGeocode.ToolTip

                    $Controls.badgeStartGeocode.Text       = $Controls.badgeEndGeocode.Text
                    $Controls.badgeStartGeocode.Foreground = $Controls.badgeEndGeocode.Foreground
                    $Controls.badgeStartGeocode.ToolTip    = $Controls.badgeEndGeocode.ToolTip

                    $Controls.badgeEndGeocode.Text       = $tempBadgeText
                    $Controls.badgeEndGeocode.Foreground = $tempBadgeFg
                    $Controls.badgeEndGeocode.ToolTip    = $tempBadgeTip
                }
                $lblStatus.Text = "⇄ Endpoints swapped for return trip"
                $lblStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
            }
            finally {
                $script:SuppressAutosuggest = $false
            }
        })
    }

    # Milestone 1: Reset Form
    if ($btnResetManualForm) {
        $btnResetManualForm.Add_Click({
            $script:SuppressAutosuggest = $true
            try {
                $txtStart.Text = ''
                $txtEnd.Text = ''
                $txtNewWp.Text = ''
                $lstWp.Items.Clear()
                $txtName.Text = ''
                $lblDist.Text = '— km'
                $lblTime.Text = '— min'
                $lblStatus.Text = 'Form reset'
                $lblStatus.Foreground = [System.Windows.Media.Brushes]::Gray
                if ($Controls.badgeStartGeocode) { $Controls.badgeStartGeocode.Text = '' }
                if ($Controls.badgeEndGeocode) { $Controls.badgeEndGeocode.Text = '' }
                $script:LastManualResult = $null
                $script:LastGoogleMapsUrl = ''
                $lblUrlDisplay.Text = 'No generated link'
                $btnExportPdf.IsEnabled = $false
                $btnExportGpx.IsEnabled = $false
                $btnExportKml.IsEnabled = $false
                $btnOpenMaps.IsEnabled = $false
                $btnCopyUrl.IsEnabled = $false
                $btnSaveMapAs.IsEnabled = $false
                if ($btnManualExportPackage) { $btnManualExportPackage.IsEnabled = $false }
                Update-MapViewMode
            }
            finally {
                $script:SuppressAutosuggest = $false
            }
        })
    }

    # Milestone 1: Recent Routes Populator & Selection Handler
    $PopulateRecentRoutes = {
        if (-not $cmbRecentRoutes) { return }
        $cmbRecentRoutes.Items.Clear()
        $placeholder = [System.Windows.Controls.ComboBoxItem]::new()
        $placeholder.Content = (Get-LocText 'ManualRecentSelectPlaceholder' '-- Select Recent Route --')
        $placeholder.Tag = $null
        $cmbRecentRoutes.Items.Add($placeholder) | Out-Null
        $cmbRecentRoutes.SelectedIndex = 0

        if ($script:AppConfig -and $script:AppConfig.RecentRoutes) {
            foreach ($r in $script:AppConfig.RecentRoutes) {
                $item = [System.Windows.Controls.ComboBoxItem]::new()
                $display = if ($r.Name) { "$($r.Name) ($($r.Timestamp))" } else { "$($r.Start) -> $($r.End)" }
                $item.Content = $display
                $item.Tag = $r
                $cmbRecentRoutes.Items.Add($item) | Out-Null
            }
        }
    }
    & $PopulateRecentRoutes

    if ($cmbRecentRoutes) {
        $cmbRecentRoutes.Add_SelectionChanged({
            if ($cmbRecentRoutes.SelectedItem -and $cmbRecentRoutes.SelectedItem.Tag) {
                $r = $cmbRecentRoutes.SelectedItem.Tag
                $script:SuppressAutosuggest = $true
                try {
                    if ($r.Start) { $txtStart.Text = $r.Start }
                    if ($r.End) { $txtEnd.Text = $r.End }
                    if ($r.Name) { $txtName.Text = $r.Name }
                    $lstWp.Items.Clear()
                    if ($r.Waypoints) {
                        foreach ($w in $r.Waypoints) {
                            if (-not [string]::IsNullOrWhiteSpace($w)) { $lstWp.Items.Add($w) | Out-Null }
                        }
                    }
                    if ($Controls.badgeStartGeocode) {
                        $Controls.badgeStartGeocode.Text = '🟢 Exact'
                        $Controls.badgeStartGeocode.Foreground = [System.Windows.Media.Brushes]::LimeGreen
                        $Controls.badgeStartGeocode.ToolTip = "Restored from history"
                    }
                    if ($Controls.badgeEndGeocode) {
                        $Controls.badgeEndGeocode.Text = '🟢 Exact'
                        $Controls.badgeEndGeocode.Foreground = [System.Windows.Media.Brushes]::LimeGreen
                        $Controls.badgeEndGeocode.ToolTip = "Restored from history"
                    }
                }
                finally {
                    $script:SuppressAutosuggest = $false
                }
            }
        })
    }

    # Milestone 1: Click-to-copy handlers on stat cards
    $copyAction = {
        param($src)
        if ($src -and -not [string]::IsNullOrWhiteSpace($src.Text) -and $src.Text -ne '— km' -and $src.Text -ne '— min' -and $src.Text -ne 'No generated link') {
            try {
                [System.Windows.Clipboard]::SetText($src.Text)
                if ($lblCopyFeedback) {
                    $lblCopyFeedback.Visibility = [System.Windows.Visibility]::Visible
                    $script:FeedbackTimer = [System.Windows.Threading.DispatcherTimer]::new()
                    $script:FeedbackTimer.Interval = [TimeSpan]::FromSeconds(2)
                    $script:FeedbackTimer.Add_Tick({
                        if ($lblCopyFeedback) { $lblCopyFeedback.Visibility = [System.Windows.Visibility]::Collapsed }
                        if ($script:FeedbackTimer) { $script:FeedbackTimer.Stop() }
                    })
                    $script:FeedbackTimer.Start()
                }
            } catch { }
        }
    }
    if ($lblDist) { $lblDist.Add_PreviewMouseLeftButtonUp({ & $copyAction $this }) }
    if ($lblTime) { $lblTime.Add_PreviewMouseLeftButtonUp({ & $copyAction $this }) }
    if ($lblType) { $lblType.Add_PreviewMouseLeftButtonUp({ & $copyAction $this }) }
    if ($lblUrlDisplay) { $lblUrlDisplay.Add_PreviewMouseLeftButtonUp({ & $copyAction $this }) }

    # Milestone 5: Full Map toggle panel
    if ($btnToggleInputPanel -and $colManualInput) {
        $btnToggleInputPanel.Add_Click({
            if ($colManualInput.Width.Value -gt 0) {
                $script:PrevManualWidth = $colManualInput.Width.Value
                $colManualInput.Width = [System.Windows.GridLength]::new(0)
                $btnToggleInputPanel.Content = '▶ Show Panel'
            } else {
                $targetW = if ($script:PrevManualWidth -gt 50) { $script:PrevManualWidth } else { 430 }
                $colManualInput.Width = [System.Windows.GridLength]::new($targetW)
                $btnToggleInputPanel.Content = '◀ Full Map'
            }
        })
    }

    # Milestone 1: Keyboard Shortcuts on Window
    if ($Window) {
        $Window.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Enter -and ($e.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
                if ($btnCalculate -and $btnCalculate.IsEnabled) {
                    $btnCalculate.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                    $e.Handled = $true
                }
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::N -and ($e.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
                if ($btnResetManualForm) {
                    $btnResetManualForm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                    $e.Handled = $true
                }
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                if ($Controls.popSuggestStart) { $Controls.popSuggestStart.IsOpen = $false }
                if ($Controls.popSuggestEnd) { $Controls.popSuggestEnd.IsOpen = $false }
                if ($Controls.popSuggestWp) { $Controls.popSuggestWp.IsOpen = $false }
                if ($Controls.pnlToastContainer) { $Controls.pnlToastContainer.Visibility = [System.Windows.Visibility]::Collapsed }
                if ($Controls.drawerLog -and $Controls.drawerLog.Visibility -eq [System.Windows.Visibility]::Visible) {
                    $Controls.drawerLog.Visibility = [System.Windows.Visibility]::Collapsed
                }
                if ($btnCalculate -and -not $btnCalculate.IsEnabled) {
                    $lblStatus.Text = 'Calculation cancelled.'
                    $lblStatus.Foreground = [System.Windows.Media.Brushes]::OrangeRed
                    $btnCalculate.IsEnabled = $true
                    $btnCalculate.Content = (Get-LocText 'ManualBtnCalculate' '🚀 CALCULATE ROUTE & DOWNLOAD MAP')
                    $e.Handled = $true
                }
            }
        })
    }

    # External navigation links
    $btnOpenMaps.Add_Click({
        if ($script:LastGoogleMapsUrl) { Start-Process $script:LastGoogleMapsUrl }
    })

    $btnCopyUrl.Add_Click({
        if ($script:LastGoogleMapsUrl) {
            [System.Windows.Clipboard]::SetText($script:LastGoogleMapsUrl)
            Show-AppToastNotification -Title (Get-LocText 'MsgUrlCopiedTitle' 'Copied') -Message (Get-LocText 'MsgUrlCopied' 'Link copied to clipboard!') -Type Info
        }
    })

    $btnSaveMapAs.Add_Click({
        if ($script:LastGeneratedMapPath -and (Test-Path $script:LastGeneratedMapPath)) {
            $dlg = [System.Windows.Forms.SaveFileDialog]::new()
            $dlg.Title = 'Save PNG Map'
            $dlg.Filter = 'PNG Image (*.png)|*.png'
            $dlg.FileName = [System.IO.Path]::GetFileName($script:LastGeneratedMapPath)
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                Copy-Item -LiteralPath $script:LastGeneratedMapPath -Destination $dlg.FileName -Force
                Show-AppToastNotification -Title (Get-LocText 'MsgMapSavedTitle' 'Saved') -Message (((Get-LocText 'MsgMapSaved' 'Map saved to: {0}') -f $dlg.FileName)) -Type Success -ActionFile $dlg.FileName
            }
        }
    })

    # Export PDF Report
    $btnExportPdf.Add_Click({
        if (-not $script:LastManualResult) { return }
        $r = $script:LastManualResult
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title = 'Export Route PDF Report'
        $dlg.Filter = 'PDF Document (*.pdf)|*.pdf'
        $dlg.FileName = "Route_$($r.DistanceKm)km.pdf"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $rName = if ($txtName -and -not [string]::IsNullOrWhiteSpace($txtName.Text)) { $txtName.Text.Trim() } elseif ($script:ActiveManualRouteName) { $script:ActiveManualRouteName } else { 'Route' }
                Export-RoutePdfReport -OutputPath $dlg.FileName `
                    -RouteName $rName `
                    -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin `
                    -RouteType $r.RouteType `
                    -AvoidTolls $r.AvoidTolls -AvoidHighways $r.AvoidHighways -AvoidFerries $r.AvoidFerries `
                    -OriginAddress $r.OriginAddress -OriginLat $r.OriginLat -OriginLng $r.OriginLng `
                    -DestAddress $r.DestAddress -DestLat $r.DestLat -DestLng $r.DestLng `
                    -Waypoints $r.Waypoints -MapImagePath $r.MapPath -GoogleMapsUrl $r.GoogleMapsUrl

                Show-AppToastNotification -Title "PDF Exported" -Message "Report generated: $($dlg.FileName)" -Type Success -ActionFile $dlg.FileName
            }
            catch {
                Show-AppToastNotification -Title "Export Error" -Message "Failed to export PDF: $($_.Exception.Message)" -Type Error
            }
        }
    })

    # Export GPX
    $btnExportGpx.Add_Click({
        if (-not $script:LastManualResult -or -not $script:LastManualResult.EncodedPolyline) { return }
        $r = $script:LastManualResult
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title = 'Export Route GPX'
        $dlg.Filter = 'GPS Exchange Format (*.gpx)|*.gpx'
        $dlg.FileName = "Route_$($r.DistanceKm)km.gpx"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $rName = if ($txtName -and -not [string]::IsNullOrWhiteSpace($txtName.Text)) { $txtName.Text.Trim() } elseif ($script:ActiveManualRouteName) { $script:ActiveManualRouteName } else { 'Route' }
                Export-RouteGpx -OutputPath $dlg.FileName `
                    -RouteName $rName `
                    -EncodedPolyline $r.EncodedPolyline `
                    -Waypoints $r.Waypoints `
                    -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin
                Show-AppToastNotification -Title "GPX Exported" -Message "Track saved: $($dlg.FileName)" -Type Success -ActionFile $dlg.FileName
            }
            catch {
                Show-AppToastNotification -Title "Export Error" -Message "Failed to export GPX: $($_.Exception.Message)" -Type Error
            }
        }
    })

    # Export KML
    $btnExportKml.Add_Click({
        if (-not $script:LastManualResult -or -not $script:LastManualResult.EncodedPolyline) { return }
        $r = $script:LastManualResult
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title = 'Export Route KML'
        $dlg.Filter = 'Keyhole Markup Language (*.kml)|*.kml'
        $dlg.FileName = "Route_$($r.DistanceKm)km.kml"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $rName = if ($txtName -and -not [string]::IsNullOrWhiteSpace($txtName.Text)) { $txtName.Text.Trim() } elseif ($script:ActiveManualRouteName) { $script:ActiveManualRouteName } else { 'Route' }
                Export-RouteKml -OutputPath $dlg.FileName `
                    -RouteName $rName `
                    -EncodedPolyline $r.EncodedPolyline `
                    -Waypoints $r.Waypoints `
                    -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin
                Show-AppToastNotification -Title "KML Exported" -Message "Track saved: $($dlg.FileName)" -Type Success -ActionFile $dlg.FileName
            }
            catch {
                Show-AppToastNotification -Title "Export Error" -Message "Failed to export KML: $($_.Exception.Message)" -Type Error
            }
        }
    })

    # Milestone 5: One-Click All Formats Package Export (PNG, PDF, GPX)
    if ($btnManualExportPackage) {
        $btnManualExportPackage.Add_Click({
            if (-not $script:LastManualResult) { return }
            $r = $script:LastManualResult
            $baseDir = if ($txtOutputDir.Text -and (Test-Path $txtOutputDir.Text)) { $txtOutputDir.Text } else { [Environment]::GetFolderPath('MyDocuments') }
            $safeName = if ($txtName.Text) { ($txtName.Text -replace '[^\w\d_-]+', '_') } else { "Route_$($r.DistanceKm)km" }
            $pkgDir = Join-Path $baseDir "$($safeName)_$((Get-Date).ToString('yyyyMMdd_HHmmss'))"
            try {
                [System.IO.Directory]::CreateDirectory($pkgDir) | Out-Null

                # 1. Save Map PNG
                if ($r.MapPath -and (Test-Path $r.MapPath)) {
                    Copy-Item -LiteralPath $r.MapPath -Destination (Join-Path $pkgDir "Map_$($r.DistanceKm)km.png") -Force
                }

                # 2. Export PDF
                $pdfPath = Join-Path $pkgDir "Report_$($r.DistanceKm)km.pdf"
                Export-RoutePdfReport -OutputPath $pdfPath `
                    -RouteName $safeName `
                    -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin `
                    -RouteType $r.RouteType `
                    -AvoidTolls $r.AvoidTolls -AvoidHighways $r.AvoidHighways -AvoidFerries $r.AvoidFerries `
                    -OriginAddress $r.OriginAddress -OriginLat $r.OriginLat -OriginLng $r.OriginLng `
                    -DestAddress $r.DestAddress -DestLat $r.DestLat -DestLng $r.DestLng `
                    -Waypoints $r.Waypoints -MapImagePath $r.MapPath -GoogleMapsUrl $r.GoogleMapsUrl

                # 3. Export GPX
                if ($r.EncodedPolyline) {
                    $gpxPath = Join-Path $pkgDir "Track_$($r.DistanceKm)km.gpx"
                    Export-RouteGpx -OutputPath $gpxPath `
                        -RouteName $safeName `
                        -EncodedPolyline $r.EncodedPolyline `
                        -Waypoints $r.Waypoints `
                        -DistanceKm $r.DistanceKm -DurationMin $r.DurationMin
                }

                Show-AppToastNotification -Title "Package Export Complete" -Message "Map, PDF, and GPX saved to folder." -Type Success -ActionFolder $pkgDir
            }
            catch {
                Show-AppToastNotification -Title "Package Export Error" -Message $_.Exception.Message -Type Error
            }
        })
    }

    # Milestone 2: Debounced Address Autosuggest Setup
    $setupAutosuggest = {
        param($textBox, $popup, $listBox, $badge)
        if (-not $textBox -or -not $popup -or -not $listBox) { return }

        $isSelecting = $false

        $timer = [System.Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(350)
        $timer.Add_Tick({
            $timer.Stop()
            if ($script:SuppressAutosuggest -or $isSelecting) { return }
            $q = $textBox.Text.Trim()
            if ($q.Length -lt 2) {
                $popup.IsOpen = $false
                return
            }
            $apiKey = if (Get-Command Get-CurrentApiKey -ErrorAction SilentlyContinue) { Get-CurrentApiKey } else { $script:AppConfig.ApiKey }
            if ([string]::IsNullOrWhiteSpace($apiKey)) { return }

            try {
                $results = Get-MapySuggest -Query $q -ApiKey $apiKey -Limit 5 -LanguageCode $script:CurrentLanguage
                $listBox.Items.Clear()
                if ($results -and @($results).Count -gt 0) {
                    foreach ($r in $results) {
                        $item = [System.Windows.Controls.ListBoxItem]::new()
                        $item.Content = "📍 $($r.FullText)"
                        $item.Tag = $r
                        $listBox.Items.Add($item) | Out-Null
                    }
                    if ($popup.Child) {
                        $popup.Child.Width = [math]::Max(360, $textBox.ActualWidth)
                    }
                    $popup.PlacementTarget = $textBox
                    $popup.IsOpen = $true
                } else {
                    $popup.IsOpen = $false
                }
            }
            catch {
                $popup.IsOpen = $false
            }
        })

        $textBox.Add_TextChanged({
            if ($script:SuppressAutosuggest -or $isSelecting) { return }
            if ($badge) { $badge.Text = '' }
            $timer.Stop()
            $timer.Start()
        })

        $textBox.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Down -and $popup.IsOpen -and $listBox.Items.Count -gt 0) {
                $listBox.Focus()
                if ($listBox.SelectedIndex -lt 0) {
                    $listBox.SelectedIndex = 0
                }
                $e.Handled = $true
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Escape -and $popup.IsOpen) {
                $popup.IsOpen = $false
                $e.Handled = $true
            }
        })

        $commitSelection = {
            if ($listBox.SelectedItem) {
                $sel = $listBox.SelectedItem.Tag
                if ($sel) {
                    $isSelecting = $true
                    $timer.Stop()
                    $textBox.Text = $sel.FullText
                    if ($badge) {
                        $badge.Text = "🟢 Exact"
                        $badge.Foreground = [System.Windows.Media.Brushes]::LimeGreen
                        $badge.ToolTip = "Coordinates: $($sel.Latitude), $($sel.Longitude)"
                    }
                    $isSelecting = $false
                }
                $popup.IsOpen = $false
                $listBox.SelectedIndex = -1
                $textBox.Focus()
                $textBox.CaretIndex = $textBox.Text.Length
            }
        }

        $listBox.Add_PreviewMouseLeftButtonUp({
            & $commitSelection
        })

        $listBox.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Enter -or $e.Key -eq [System.Windows.Input.Key]::Return) {
                & $commitSelection
                $e.Handled = $true
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                $popup.IsOpen = $false
                $textBox.Focus()
                $e.Handled = $true
            }
        })
    }

    & $setupAutosuggest $txtStart $Controls.popSuggestStart $Controls.lstSuggestStart $Controls.badgeStartGeocode
    & $setupAutosuggest $txtEnd $Controls.popSuggestEnd $Controls.lstSuggestEnd $Controls.badgeEndGeocode
    & $setupAutosuggest $txtNewWp $Controls.popSuggestWp $Controls.lstSuggestWp $null

    # Milestone 4: Collapsible Activity Log Drawer Events
    if ($Controls.btnToggleLogDrawer) {
        $Controls.btnToggleLogDrawer.Add_Click({
            if ($Controls.drawerLog) {
                if ($Controls.drawerLog.Visibility -eq [System.Windows.Visibility]::Visible) {
                    $Controls.drawerLog.Visibility = [System.Windows.Visibility]::Collapsed
                } else {
                    $Controls.drawerLog.Visibility = [System.Windows.Visibility]::Visible
                    Update-LogDrawerDisplay
                }
            }
        })
    }

    if ($Controls.btnCloseLogDrawer) {
        $Controls.btnCloseLogDrawer.Add_Click({
            if ($Controls.drawerLog) { $Controls.drawerLog.Visibility = [System.Windows.Visibility]::Collapsed }
        })
    }

    if ($Controls.btnClearLogDrawer) {
        $Controls.btnClearLogDrawer.Add_Click({
            Clear-AppLogDrawer
        })
    }

    if ($Controls.btnCopyLogDrawer) {
        $Controls.btnCopyLogDrawer.Add_Click({
            if ($Controls.txtLogDrawer -and -not [string]::IsNullOrWhiteSpace($Controls.txtLogDrawer.Text)) {
                [System.Windows.Clipboard]::SetText($Controls.txtLogDrawer.Text)
                Show-AppToastNotification -Title "Log Copied" -Message "Activity log copied to clipboard." -Type Info
            }
        })
    }

    foreach ($rb in @($Controls.rbLogAll, $Controls.rbLogInfo, $Controls.rbLogWarn, $Controls.rbLogError)) {
        if ($rb) {
            $rb.Add_Checked({
                Update-LogDrawerDisplay
            })
        }
    }

    # Initialize map view mode based on configuration
    if ($script:AppConfig -and ($false -eq $script:AppConfig.UseInteractiveMap)) {
        if ($rbViewStatic) { $rbViewStatic.IsChecked = $true }
    } else {
        if ($rbViewInteractive) { $rbViewInteractive.IsChecked = $true }
    }
    Update-MapViewMode
}

#endregion 2. Manual Tab Event Registration & Wiring

#region 3. Global Function Exports

# Export function into global scope for GUI orchestrator
Set-Item -Path "function:global:Register-UiManualTabEvents" -Value (Get-Item "function:Register-UiManualTabEvents").ScriptBlock -ErrorAction SilentlyContinue

#endregion 3. Global Function Exports
