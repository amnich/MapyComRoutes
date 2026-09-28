# MapyComRoutes — User Experience (UX) Milestone Plan

## Objective
Elevate usability, workflow ergonomics, and speed for both single-route planning and large-scale batch route generation in MapyComRoutes.

---

## Milestone Breakdown & Implementation Status

### Milestone 1: Ergonomic Route Builder (Manual Tab) — [COMPLETE]
- [x] **Origin/Destination Swap**: Single click button (`btnSwapEndpoints`) to reverse start and end, with automatic waypoint reversal for instant return-trip calculations.
- [x] **Power Shortcuts**:
  - `Ctrl + Enter`: Trigger route calculation.
  - `Ctrl + N`: Clear form and start new route.
  - `Esc`: Cancel active calculation and close open popups/toasts.
- [x] **Click-to-Copy Data Pills**: Click on Distance, Duration, Route Type, or Map URL to copy text to clipboard with a brief visual feedback badge (`lblCopyFeedback`).
- [x] **Recent Routes History**: Dropdown (`cmbRecentRoutes`) remembering the last 10 calculated routes with one-click restore.

### Milestone 2: Mapy.com Autosuggest & Inline Validation — [COMPLETE]
- [x] **Debounced Address Autocomplete**: Integrated Mapy.com `/v1/suggest` (with fallback to `/v1/geocode?limit=5`) API to provide instant address dropdown suggestions (`popSuggestStart`, `popSuggestEnd`, `popSuggestWp`) as users type.
- [x] **Inline Geocode Quality Badges**: Visual indicator (`badgeStartGeocode`, `badgeEndGeocode`) showing `🟢 Exact` / `🟡 Approx` / `🔴 Error` next to input fields.
- [x] **Interactive Map Context Menu**: Right-click anywhere on the Leaflet map to "Set as Start", "Add as Stop", "Set as Destination", or "Copy Lat/Lng" with direct WebView2 `postMessage` synchronization and clipboard fallback.

### Milestone 3: Frictionless Batch Processing — [COMPLETE]
- [x] **File Drag-and-Drop**: Drop Excel (`.xlsx`), CSV, or JSON files directly onto the Batch tab card (`borderBatchInputCard`).
- [x] **Dynamic ETA & Throughput**: Live computation of remaining batch duration (`Route X of Y (Z%) • ~Xs remaining • Xs/route`).
- [x] **Results Filter & Search**: Instant text filtering (`txtBatchSearch`) and Status dropdown (`cmbBatchStatusFilter`: All / Success / Errors) above the Results DataGrid.
- [x] **Retry Failed Only**: Button (`btnRetryFailedBatch`) to re-run only rows that encountered network timeouts or geocode errors without restarting the entire batch.

### Milestone 4: Non-Intrusive Notifications & Log Drawer — [COMPLETE]
- [x] **WPF Toast Snackbars**: Replaced modal blocking message dialogs (`MessageBox.Show`) with non-intrusive bottom-right toast notifications (`pnlToastContainer`) for exports and saves with direct "Open File" and "Open Folder" action buttons.
- [x] **Collapsible Activity Log**: Sleek slide-out log drawer (`drawerLog`) with live streaming via `Write-AppLog`, severity toggles (All / Info / Warn / Error), clear, copy, and footer toggle button (`btnToggleLogDrawer`).

### Milestone 5: Enhanced Map Viewer & Multi-Export — [COMPLETE]
- [x] **Collapsible Control Panel**: Fullscreen / split-view toggle button (`btnToggleInputPanel`) allowing the map preview to take 100% of window width during route inspection.
- [x] **Map Layer Switcher**: Toggle between Mapy Standard, OpenStreetMap, and Satellite Orthophoto (`baseMaps`) in the interactive Leaflet viewer.
- [x] **One-Click Package Export**: Export Map PNG, Turn Directions PDF, and GPX track into a designated trip folder in a single click (`btnManualExportPackage`).

### Milestone 6: API Statistics & Clean Settings — [COMPLETE]
- [x] **Request Call Statistics**: Replaced legacy Google Maps dollar cost estimation with Mapy.com API request counters tracking Session vs. Monthly requests, with live breakdown across Geocoding/Suggest (`/v1/suggest`, `/v1/geocode`), Routes API (`/v1/routing`), and Map Imagery/Tiles.
- [x] **Billing Period Reset**: Displays active tracking month (`lblApiBillingPeriod`) and single-click month counter reset (`btnResetApiCounters`).

### Milestone 7: Automatic Proportional Waypoint Splitting & Multi-Map Output (v2.2) — [COMPLETE]
- [x] **Automatic Proportional Route Splitting (`Split-RoutePoints`)**: Automatically partitions routes exceeding Mapy.com's limit of 15 intermediate waypoints (17 total points: 1 start + 15 waypoints + 1 end) into $N = \lceil (P - 1) / 16 \rceil$ balanced parts.
  - **Proportional Leg Balancing**: Distributes legs evenly across parts so leg counts differ by at most 1 (e.g. 24 points split in half into 12 legs / 13 points and 11 legs / 12 points).
  - **Continuous Chaining Guarantee**: Invariant `Part[k].End == Part[k+1].Start` is strictly preserved across all sub-routes.
- [x] **Multi-Map Image Generation**:
  - Automatically renders $N$ individual sub-route PNG maps (`<name>_Part1ofN.png`, `<name>_Part2ofN.png`) with dedicated camera bounding boxes, markers, and distance badges.
  - Automatically renders 1 combined stitched overview PNG map (`<name>.png`) with full concatenated polyline and markers across the entire trip.
- [x] **High-Performance C# Polyline Codec (`GoogleMapsPolylineCodec`)**:
  - Integrated zero-dependency C# 5.0 compatible polyline encoder and decoder via `Add-Type`.
  - Enables microsecond polyline encoding, decoding, and lossless concatenation across partitioned sub-routes with coordinate deduplication at split junctions.
- [x] **Universal Sequential Data Ingestion (`Import-RouteDataFile`)**:
  - Automatically detects and splits tabular route sequences (`SequentialStops` and grouped `RouteList`) exceeding 17 points across Excel (`.xlsx`), CSV, and JSON formats into chained sub-routes with labeled part designations (`Trasa (Part 1/2)`).
- [x] **Expanded UI Route Builder**:
  - Increased manual waypoint list limit from 25 to 200 stops in `UiManualTab.ps1`.
  - Background calculation (`AsyncWorkers.ps1`) transparently renders and bundles all sub-route maps for instant inspection.
- [x] **Trilingual Localization Parity**:
  - Registered and synchronized all waypoint splitting and capacity keys (`ManualWaypoints`, `MsgMaxWaypoints`, `MsgLangReloadedTitle`) in `localization.json` with 100% key parity across English, German, and Polish.

---

## Technical Verification & Guardrails
- **Compatibility**: Verified 100% compliant and operational on Windows PowerShell 5.1 and PowerShell 7+.
- **Encoding**: Verified strict UTF-8 with BOM (`0xEF, 0xBB, 0xBF`) across all source code, XAML, JSON, and documentation files.
- **PS2EXE Standalone Executable**: Successfully bundled and compiled into `MapyComRoutes.exe` (828.5 KB), fully self-sufficient with zero external runtime file dependencies.
- **Multilingual**: All newly introduced UI keys registered and translated in `localization.json` for English, German, and Polish.