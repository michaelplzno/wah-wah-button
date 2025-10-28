# ============================================================================
# CONFIGURATION: Window titles to skip (widgets, background UI, etc.)
# Add window titles or patterns here to exclude them from being arranged
# ============================================================================
$script:skipWindowTitles = @(
    'Xbox'           # Xbox Game Bar widgets
)

Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace Win32 {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT { public int Left, Top, Right, Bottom; }

  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
  public class MONITORINFO {
    public int cbSize = Marshal.SizeOf(typeof(MONITORINFO));
    public RECT rcMonitor = new RECT();
    public RECT rcWork = new RECT();
    public uint dwFlags = 0;
  }

  public static class NativeMethods {
    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    public delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
    // Window style helpers for detecting topmost
    [DllImport("user32.dll", EntryPoint="GetWindowLongPtr", SetLastError=true)] public static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll", EntryPoint="GetWindowLong", SetLastError=true)] public static extern int GetWindowLong32(IntPtr hWnd, int nIndex);
    public static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex) { return IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : new IntPtr(GetWindowLong32(hWnd, nIndex)); }
    [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out bool pvAttribute, int cbAttribute);

    [DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern bool GetMonitorInfo(IntPtr hMonitor, MONITORINFO lpmi);
    
    public const int DWMWA_CLOAKED = 14;

    public const uint GA_ROOT = 2;
    public const uint MONITOR_DEFAULTTOPRIMARY = 1;
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_NOZORDER = 0x0004;
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_SHOWWINDOW = 0x0040;
        public static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
    public static readonly IntPtr HWND_TOP = new IntPtr(0);
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        public const int GWL_EXSTYLE = -20;
        public const int WS_EX_TOPMOST = 0x00000008;
    public const int GWL_STYLE = -16;
    public const int WS_VISIBLE = 0x10000000;
    public const int WS_MINIMIZE = 0x20000000;
  }
}
"@

# Fun: play SFX on start if available (helper process runs PlaySync so audio completes)
try {
    $wavPath = Join-Path $PSScriptRoot 'stooges.wav'
    if (Test-Path -LiteralPath $wavPath) {
        $tempDir = [System.IO.Path]::GetTempPath()
        $helperPath = Join-Path $tempDir ("PlayStooge_" + [Guid]::NewGuid().ToString('N') + ".ps1")

        $helperCode = @'
try {
    $p = New-Object System.Media.SoundPlayer -ArgumentList '<WAV_PATH>'
    $p.Load()
    $p.PlaySync()
} catch {
    try {
        "[$(Get-Date -Format o)] ERROR: $($_.Exception.Message)" | Out-File -FilePath "$env:TEMP\WahWahButton_SFX.log" -Append -Encoding UTF8
    } catch {}
}
'@
        # Inject the absolute path safely (escape single quotes for PowerShell literal string)
        $pathLiteral = $wavPath.Replace("'", "''")
        $helperCode = $helperCode.Replace('<WAV_PATH>', $pathLiteral)

        Set-Content -LiteralPath $helperPath -Value $helperCode -Encoding UTF8 -Force

        $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        Start-Process -FilePath $psExe -WindowStyle Hidden -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File', $helperPath) | Out-Null
    }
} catch { }

# First, enumerate all monitors
$allMonitors = [System.Collections.ArrayList]::new()
$monitorCallback = [Win32.NativeMethods+MonitorEnumProc]{
    param([IntPtr]$hMonitor, [IntPtr]$hdcMonitor, [ref]$lprcMonitor, [IntPtr]$dwData)
    $script:allMonitors.Add($hMonitor.ToInt64()) | Out-Null
    return $true
}
[Win32.NativeMethods]::EnumDisplayMonitors([IntPtr]::Zero, [IntPtr]::Zero, $monitorCallback, [IntPtr]::Zero) | Out-Null
Write-Host "System has $($allMonitors.Count) monitor(s) connected."

# Collect all top-level visible windows
$windows = [System.Collections.ArrayList]::new()

$callback = [Win32.NativeMethods+EnumWindowsProc]{
    param([IntPtr]$hWnd, [IntPtr]$lParam)

    # Skip invisible and child windows; only top-level visible windows
    if (-not [Win32.NativeMethods]::IsWindowVisible($hWnd)) { return $true }
    $root = [Win32.NativeMethods]::GetAncestor($hWnd, [Win32.NativeMethods]::GA_ROOT)
    if ($root -ne $hWnd) { return $true }

    # Skip taskbar and system UI windows
    $className = New-Object System.Text.StringBuilder 256
    [Win32.NativeMethods]::GetClassName($hWnd, $className, $className.Capacity) | Out-Null
    $class = $className.ToString()
    
    # Filter out taskbar, system tray, desktop, and shell staging elements (but allow UWP frames)
    $skipClasses = @('Shell_TrayWnd', 'Shell_SecondaryTrayWnd', 'Progman', 'WorkerW', 
                     'ForegroundStaging')
    if ($skipClasses -contains $class) { return $true }

    # Skip minimized windows; include maximized so we can control Z-order later
    if ([Win32.NativeMethods]::IsIconic($hWnd)) { return $true }
    
    # Additional check for minimized windows that don't report as iconic
    # Minimized windows often have extreme negative coordinates
    if ($rect.Left -le -30000 -or $rect.Top -le -30000) { return $true }

    # Check if window is cloaked (hidden UWP/modern apps)
    $isCloaked = $false
    try {
        $result = [Win32.NativeMethods]::DwmGetWindowAttribute($hWnd, [Win32.NativeMethods]::DWMWA_CLOAKED, [ref]$isCloaked, [System.Runtime.InteropServices.Marshal]::SizeOf([Type]::GetType("System.Boolean")))
        if ($isCloaked) { return $true }
    } catch {
        # DwmGetWindowAttribute might fail on some windows, continue anyway
    }

    # Get current window rect
    $rect = New-Object Win32.RECT
    if (-not [Win32.NativeMethods]::GetWindowRect($hWnd, [ref]$rect)) { return $true }
    $w = $rect.Right - $rect.Left
    $h = $rect.Bottom - $rect.Top
    if ($w -le 0 -or $h -le 0) { return $true }
    
    # Skip tiny windows (widgets, background UI elements, system tray elements)
    # Real user windows are typically at least 400x300
    if ($w -lt 400 -or $h -lt 300) { return $true }

    # Get the monitor this window is currently on
    $hMon = [Win32.NativeMethods]::MonitorFromWindow($hWnd, [Win32.NativeMethods]::MONITOR_DEFAULTTOPRIMARY)

    # Determine monitor work area and whether the window is effectively fullscreen on that monitor
    $mi = New-Object Win32.MONITORINFO
    [Win32.NativeMethods]::GetMonitorInfo($hMon, $mi) | Out-Null
    $work = $mi.rcWork
    $monWidth = $work.Right - $work.Left
    $monHeight = $work.Bottom - $work.Top
    $area = $w * $h
    $monArea = [Math]::Max(1, $monWidth * $monHeight)
    $coverage = [double]$area / [double]$monArea
    $alignedToWork = ([Math]::Abs($rect.Left - $work.Left) -le 8 -and 
                      [Math]::Abs($rect.Top - $work.Top) -le 8 -and 
                      [Math]::Abs($rect.Right - $work.Right) -le 8 -and 
                      [Math]::Abs($rect.Bottom - $work.Bottom) -le 8)
    # Skip truly fullscreen windows: maximized AND covering 98%+ of work area
    $isFullscreen = ([Win32.NativeMethods]::IsZoomed($hWnd)) -and ($coverage -ge 0.98) -and $alignedToWork

    # Detect always-on-top (topmost) extended style
    $exPtr = [Win32.NativeMethods]::GetWindowLongPtr($hWnd, [Win32.NativeMethods]::GWL_EXSTYLE)
    $exVal = [int]($exPtr.ToInt64() -band 0xFFFFFFFF)
    $isTopmost = (($exVal -band [Win32.NativeMethods]::WS_EX_TOPMOST) -ne 0)

    # Track raw enumeration order from EnumWindows (top-most to bottom)
    $rawZOrder = $script:enumIndex
    $script:enumIndex++

    # Get window title
    $titleBuilder = New-Object System.Text.StringBuilder 256
    [Win32.NativeMethods]::GetWindowText($hWnd, $titleBuilder, 256) | Out-Null
    $title = $titleBuilder.ToString()
    
    # Skip windows with empty titles (usually widgets or background UI)
    if ([string]::IsNullOrWhiteSpace($title)) { return $true }
    
    # Skip windows with system-like names (widgets, background UI, etc.)
    if ($title -match '^(Windows Input Experience|MainWindowView|wv_\d+)$') { return $true }
    
    # Skip windows in the user-configured skip list
    foreach ($skipTitle in $script:skipWindowTitles) {
        if ($title -eq $skipTitle) { return $true }
    }

    # Check if window is actually on-screen (visible area intersects with monitor work area)
    $isOnScreen = ($rect.Right -gt $work.Left -and $rect.Left -lt $work.Right -and 
                   $rect.Bottom -gt $work.Top -and $rect.Top -lt $work.Bottom)

    # Store window info with monitor handle as string for grouping
    $script:windows.Add([PSCustomObject]@{
        Handle    = $hWnd
        Width     = $w
        Height    = $h
        Area      = $w * $h
        Monitor   = $hMon.ToInt64()
        IsFullscreen = $isFullscreen
        IsTopmost = $isTopmost
        RawZOrder = $rawZOrder
        Title     = $title
        IsOnScreen = $isOnScreen
        Left      = $rect.Left
        Top       = $rect.Top
    }) | Out-Null

    return $true
}

# Run the enumeration to collect windows
$script:enumIndex = 0
[Win32.NativeMethods]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null

Write-Host "Collected $($windows.Count) non-maximized, non-system windows."

# Assign contiguous ZIndex starting from 1 based on the subset of windows we're actually arranging
# Sort by RawZOrder (EnumWindows order = top to bottom), then assign 1, 2, 3, etc.
$sortedByRawZ = $windows | Sort-Object -Property RawZOrder
for ($i = 0; $i -lt $sortedByRawZ.Count; $i++) {
    $sortedByRawZ[$i] | Add-Member -NotePropertyName ZIndex -NotePropertyValue ($i + 1) -Force
}

if ($windows.Count -eq 0) {
    Write-Host "No non-maximized windows found to arrange."
    exit 0
}

# Group windows by monitor using hashtable
$monitorGroups = @{}
foreach ($win in $windows) {
    $monKey = $win.Monitor.ToString()
    if (-not $monitorGroups.ContainsKey($monKey)) {
        $monitorGroups[$monKey] = [System.Collections.ArrayList]::new()
    }
    $monitorGroups[$monKey].Add($win) | Out-Null
}

Write-Host "Found $($windows.Count) windows across $($monitorGroups.Count) monitor(s) with windows."
Write-Host ""

# Process each monitor separately
foreach ($monKey in $monitorGroups.Keys) {
    $monitorWindows = $monitorGroups[$monKey]
    $hMon = [IntPtr]::new([Int64]$monKey)
    
    # Get this monitor's work area
    $mi = New-Object Win32.MONITORINFO
    [Win32.NativeMethods]::GetMonitorInfo($hMon, $mi) | Out-Null
    $work = $mi.rcWork
    
    # Partition windows into fullscreen and non-fullscreen; then split non-fullscreen by topmost
    $fullscreens = $monitorWindows | Where-Object { $_.IsFullscreen }
    $nonFullscreen = $monitorWindows | Where-Object { -not $_.IsFullscreen }
    
    # FILTER: Only arrange windows that are actually on-screen (visible area overlaps monitor)
    $onScreenWindows = $nonFullscreen | Where-Object { $_.IsOnScreen }
    
    $topmostWindows = $onScreenWindows | Where-Object { $_.IsTopmost }
    $normalWindows = $onScreenWindows | Where-Object { -not $_.IsTopmost }

    # IMPORTANT: Assign per-monitor local layer indices separately for normal and topmost windows
    # This ensures each group starts with layer 0, 1, 2... regardless of global Z values
    $normalWindowsSorted = $normalWindows | Sort-Object -Property ZIndex
    $topmostWindowsSorted = $topmostWindows | Sort-Object -Property ZIndex
    
    for ($i = 0; $i -lt $normalWindowsSorted.Count; $i++) {
        $normalWindowsSorted[$i] | Add-Member -NotePropertyName LocalLayer -NotePropertyValue $i -Force
    }
    
    for ($i = 0; $i -lt $topmostWindowsSorted.Count; $i++) {
        $topmostWindowsSorted[$i] | Add-Member -NotePropertyName LocalLayer -NotePropertyValue $i -Force
    }

    # Sort groups by area for display purposes
    $sortedWindows = $normalWindows | Sort-Object -Property Area -Descending
    $sortedTopmost = $topmostWindows | Sort-Object -Property Area -Descending
    
    $monWidth = $work.Right - $work.Left
    $monHeight = $work.Bottom - $work.Top
    $totalNonFS = $normalWindows.Count + $topmostWindows.Count
    Write-Host "  Monitor ${monKey}: ${monWidth}x${monHeight} - Arranging $totalNonFS windows ($($normalWindows.Count) normal, $($topmostWindows.Count) topmost, $($fullscreens.Count) fullscreen skipped) (work area: $($work.Left),$($work.Top) to $($work.Right),$($work.Bottom))"
    
    # Compute target sizes STARTING FROM BOTTOM (largest) working UP to smallest
    # Bottom layer should be nearly full monitor, top layer smallest
    $sizeMap = @{}
    if ($normalWindows.Count -gt 0) {
        # Use the sorted normal windows list with LocalLayer property (0 = topmost, N-1 = deepest)
        $ascendingNormalsForSize = $normalWindowsSorted | Sort-Object -Property LocalLayer
        
        # Calculate base margins and cascade offset
        $basePadX = 20
        $basePadY = 20
        $cascadeOffsetX = 35
        $cascadeOffsetY = 35
        $minWAllowed = 300
        $minHAllowed = 200
        
        $numLayers = $ascendingNormalsForSize.Count
        
        # Assign sizes layer by layer using LOCAL layer index (0 = top/smallest, N-1 = bottom/largest)
        for ($localLayer = 0; $localLayer -lt $numLayers; $localLayer++) {
            $wObj = $ascendingNormalsForSize[$localLayer]
            
            # Calculate margin: top layer (0) has max margin, bottom layer (N-1) has min margin
            $k = ($numLayers - 1 - $localLayer)  # Inverted: localLayer 0->max margin, (N-1)->min margin
            
            $maxLeftMargin  = [Math]::Max(0, [Math]::Floor(($monWidth  - $minWAllowed) / 2))
            $maxTopMargin   = [Math]::Max(0, [Math]::Floor(($monHeight - $minHAllowed) / 2))
            $marginX = [Math]::Min($k * $cascadeOffsetX, $maxLeftMargin)
            $marginY = [Math]::Min($k * $cascadeOffsetY, $maxTopMargin)
            
            # Calculate size: full monitor minus (base padding + layer margin) on each side
            $targetW = $monWidth - (2 * ($basePadX + $marginX))
            $targetH = $monHeight - (2 * ($basePadY + $marginY))
            $targetW = [Math]::Max($minWAllowed, [Math]::Min($monWidth, $targetW))
            $targetH = [Math]::Max($minHAllowed, [Math]::Min($monHeight, $targetH))
            
            $sizeMap[$wObj.Handle.ToInt64()] = @{ W = $targetW; H = $targetH; Layer = $localLayer }
        }
    }
    
    # Calculate sizes for topmost windows using their own layer indices
    if ($topmostWindows.Count -gt 0) {
        $ascendingTopmostForSize = $topmostWindowsSorted | Sort-Object -Property LocalLayer
        
        # Calculate base margins and cascade offset
        $basePadX = 20
        $basePadY = 20
        $cascadeOffsetX = 35
        $cascadeOffsetY = 35
        $minWAllowed = 300
        $minHAllowed = 200
        
        $numLayers = $ascendingTopmostForSize.Count
        
        # Assign sizes layer by layer using LOCAL layer index for topmost windows
        for ($localLayer = 0; $localLayer -lt $numLayers; $localLayer++) {
            $wObj = $ascendingTopmostForSize[$localLayer]
            
            # Calculate margin: top layer (0) has max margin, bottom layer (N-1) has min margin
            $k = ($numLayers - 1 - $localLayer)  # Inverted: localLayer 0->max margin, (N-1)->min margin
            
            $maxLeftMargin  = [Math]::Max(0, [Math]::Floor(($monWidth  - $minWAllowed) / 2))
            $maxTopMargin   = [Math]::Max(0, [Math]::Floor(($monHeight - $minHAllowed) / 2))
            $marginX = [Math]::Min($k * $cascadeOffsetX, $maxLeftMargin)
            $marginY = [Math]::Min($k * $cascadeOffsetY, $maxTopMargin)
            
            # Calculate size: full monitor minus (base padding + layer margin) on each side
            $targetW = $monWidth - (2 * ($basePadX + $marginX))
            $targetH = $monHeight - (2 * ($basePadY + $marginY))
            $targetW = [Math]::Max($minWAllowed, [Math]::Min($monWidth, $targetW))
            $targetH = [Math]::Max($minHAllowed, [Math]::Min($monHeight, $targetH))
            
            $sizeMap[$wObj.Handle.ToInt64()] = @{ W = $targetW; H = $targetH; Layer = $localLayer }
        }
    }
    
    # Annotate windows with resolved Layer for consistent positioning
    foreach ($w in $sortedWindows) {
        $key = $w.Handle.ToInt64()
        $layer = 0
        if ($sizeMap.ContainsKey($key)) { $layer = [int]$sizeMap[$key].Layer }
        $w | Add-Member -NotePropertyName Layer -NotePropertyValue $layer -Force
    }
    foreach ($w in $sortedTopmost) {
        $key = $w.Handle.ToInt64()
        $layer = 0
        if ($sizeMap.ContainsKey($key)) { $layer = [int]$sizeMap[$key].Layer }
        $w | Add-Member -NotePropertyName Layer -NotePropertyValue $layer -Force
    }

    # Cascade settings: fixed offset between each layer
    $cascadeOffsetX = 35
    $cascadeOffsetY = 35
    
    # Process normal windows on this monitor
    for ($i = 0; $i -lt $sortedWindows.Count; $i++) {
        $win = $sortedWindows[$i]
        $hWnd = $win.Handle
        # Determine target size from size map (ensures back windows are bigger than front ones)
        $key = $hWnd.ToInt64()
        if ($sizeMap.ContainsKey($key)) {
            $w = [int]$sizeMap[$key].W
            $h = [int]$sizeMap[$key].H
            $layer = $sizeMap[$key].Layer
        } else {
            # Fallback clamp
            $w = [Math]::Min($win.Width,  [Math]::Max(100, $monWidth))
            $h = [Math]::Min($win.Height, [Math]::Max(100, $monHeight))
            $layer = -1
            Write-Host "    Window (FALLBACK): ${w}x${h}"
        }
        
        # Pyramid positioning: symmetric margins based on layer so both top-left and bottom-right shift by layer*cascade
        $totalLayers = [Math]::Max(1, $sortedWindows.Count)
        if ($layer -lt 0) { $layer = 0 }
        # Smallest (layer 0) gets largest margins; deepest layer gets smallest margins
        $k = ($totalLayers - 1 - $layer)
        $basePadX = 20
        $basePadY = 20
        $minWAllowed = 300
        $minHAllowed = 200
        $maxLeftMargin  = [Math]::Max(0, [Math]::Floor(($monWidth  - $minWAllowed) / 2))
        $maxTopMargin   = [Math]::Max(0, [Math]::Floor(($monHeight - $minHAllowed) / 2))
        $marginX = [Math]::Min($k * $cascadeOffsetX, $maxLeftMargin)
        $marginY = [Math]::Min($k * $cascadeOffsetY, $maxTopMargin)

        $newX = $work.Left + $basePadX + $marginX
        $newY = $work.Top  + $basePadY + $marginY
        $right = $work.Right  - $basePadX - $marginX
        $bottom= $work.Bottom - $basePadY - $marginY

        $w = [Math]::Max($minWAllowed, $right - $newX)
        $h = [Math]::Max($minHAllowed, $bottom - $newY)

        # Move/resize window to new position (stays on same monitor)
        $result = [Win32.NativeMethods]::MoveWindow($hWnd, $newX, $newY, $w, $h, $true)
        
        # For stubborn UWP apps, also try SetWindowPos with size/move flags (preserve Z-order)
        if (-not $result) {
            [Win32.NativeMethods]::SetWindowPos($hWnd, [IntPtr]::Zero, $newX, $newY, $w, $h,
                ([Win32.NativeMethods]::SWP_NOACTIVATE -bor [Win32.NativeMethods]::SWP_NOZORDER -bor [Win32.NativeMethods]::SWP_SHOWWINDOW)) | Out-Null
        }
    }

    # Process topmost windows on this monitor using same pyramid sizing/position
    for ($i = 0; $i -lt $sortedTopmost.Count; $i++) {
        $win = $sortedTopmost[$i]
        $hWnd = $win.Handle
        $key = $hWnd.ToInt64()
        if ($sizeMap.ContainsKey($key)) {
            $w = [int]$sizeMap[$key].W
            $h = [int]$sizeMap[$key].H
            $layer = $sizeMap[$key].Layer
        } else {
            $w = [Math]::Min($win.Width,  [Math]::Max(100, $monWidth))
            $h = [Math]::Min($win.Height, [Math]::Max(100, $monHeight))
            $layer = -1
        }

        $totalLayers = [Math]::Max(1, ($normalWindows.Count + $topmostWindows.Count))
        if ($layer -lt 0) { $layer = 0 }
        $k = ($totalLayers - 1 - $layer)
        $basePadX = 20
        $basePadY = 20
        $minWAllowed = 300
        $minHAllowed = 200
        $maxLeftMargin  = [Math]::Max(0, [Math]::Floor(($monWidth  - $minWAllowed) / 2))
        $maxTopMargin   = [Math]::Max(0, [Math]::Floor(($monHeight - $minHAllowed) / 2))
        $marginX = [Math]::Min($k * $cascadeOffsetX, $maxLeftMargin)
        $marginY = [Math]::Min($k * $cascadeOffsetY, $maxTopMargin)

        $newX = $work.Left + $basePadX + $marginX
        $newY = $work.Top  + $basePadY + $marginY
        $right = $work.Right  - $basePadX - $marginX
        $bottom= $work.Bottom - $basePadY - $marginY

        $w = [Math]::Max($minWAllowed, $right - $newX)
        $h = [Math]::Max($minHAllowed, $bottom - $newY)

        $result = [Win32.NativeMethods]::MoveWindow($hWnd, $newX, $newY, $w, $h, $true)
        if (-not $result) {
            [Win32.NativeMethods]::SetWindowPos($hWnd, [IntPtr]::Zero, $newX, $newY, $w, $h,
                ([Win32.NativeMethods]::SWP_NOACTIVATE -bor [Win32.NativeMethods]::SWP_NOZORDER -bor [Win32.NativeMethods]::SWP_SHOWWINDOW)) | Out-Null
        }
    }

    # Keep existing Z-order unchanged per your request; do not restack windows.
}

Write-Host "Done: arranged windows with fullscreen at back, normal windows cascaded with progressive sizing (smallest on top, largest at back)."
