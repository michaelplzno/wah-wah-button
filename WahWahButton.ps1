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

    [DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern bool GetMonitorInfo(IntPtr hMonitor, MONITORINFO lpmi);

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

    # Get current window rect
    $rect = New-Object Win32.RECT
    if (-not [Win32.NativeMethods]::GetWindowRect($hWnd, [ref]$rect)) { return $true }
    $w = $rect.Right - $rect.Left
    $h = $rect.Bottom - $rect.Top
    if ($w -le 0 -or $h -le 0) { return $true }

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
    $alignedToWork = ([Math]::Abs($rect.Left - $work.Left) -le 1 -and 
                      [Math]::Abs($rect.Top - $work.Top) -le 1 -and 
                      [Math]::Abs($rect.Right - $work.Right) -le 1 -and 
                      [Math]::Abs($rect.Bottom - $work.Bottom) -le 1)
    $isFullscreen = ([Win32.NativeMethods]::IsZoomed($hWnd)) -or ($coverage -ge 0.95 -and $alignedToWork)

    # Detect always-on-top (topmost) extended style
    $exPtr = [Win32.NativeMethods]::GetWindowLongPtr($hWnd, [Win32.NativeMethods]::GWL_EXSTYLE)
    $exVal = [int]($exPtr.ToInt64() -band 0xFFFFFFFF)
    $isTopmost = (($exVal -band [Win32.NativeMethods]::WS_EX_TOPMOST) -ne 0)

    # Track Z-order index from enumeration (EnumWindows enumerates top-most to bottom)
    $zOrderIndex = $script:enumIndex
    $script:enumIndex++

    # Store window info with monitor handle as string for grouping
    $script:windows.Add([PSCustomObject]@{
        Handle    = $hWnd
        Width     = $w
        Height    = $h
        Area      = $w * $h
        Monitor   = $hMon.ToInt64()
        IsFullscreen = $isFullscreen
        IsTopmost = $isTopmost
        ZIndex    = $zOrderIndex
    }) | Out-Null

    return $true
}

# Run the enumeration to collect windows
$script:enumIndex = 0
[Win32.NativeMethods]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null

Write-Host "Collected $($windows.Count) non-maximized, non-system windows."

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
    $topmostWindows = $nonFullscreen | Where-Object { $_.IsTopmost }
    $normalWindows = $nonFullscreen | Where-Object { -not $_.IsTopmost }

    # Sort groups by area
    $sortedWindows = $normalWindows | Sort-Object -Property Area -Descending
    $sortedTopmost = $topmostWindows | Sort-Object -Property Area -Descending
    
    $monWidth = $work.Right - $work.Left
    $monHeight = $work.Bottom - $work.Top
    Write-Host "  Monitor ${monKey}: ${monWidth}x${monHeight} - Arranging $($sortedWindows.Count) windows (work area: $($work.Left),$($work.Top) to $($work.Right),$($work.Bottom))"
    
    # Compute target sizes so each layer is strictly larger in BOTH width AND height than all layers in front
    # Layer order is based on current Z (top-most gets smallest size), preserving existing Z ordering
    $n = ($normalWindows.Count + $topmostWindows.Count)
    $sizeMap = @{}
    if ($n -gt 0) {
        # Sort all non-fullscreen windows by current Z order (ascending = top to bottom)
        $ascendingNormalsForSize = $nonFullscreen | Sort-Object -Property ZIndex
        
        # Define base minimum increment per layer to ensure each is visibly larger
        $minIncrementW = [Math]::Max(50, [Math]::Floor($monWidth * 0.08))
        $minIncrementH = [Math]::Max(40, [Math]::Floor($monHeight * 0.08))
        
        # Start with smallest window: use 45% of monitor (but at least 400x300)
        $baseW = [Math]::Min($monWidth, [Math]::Max(400, [Math]::Floor($monWidth * 0.45)))
        $baseH = [Math]::Min($monHeight, [Math]::Max(300, [Math]::Floor($monHeight * 0.45)))
        
        # Assign sizes layer by layer, ensuring each is strictly larger in both dimensions
        for ($idx = 0; $idx -lt $ascendingNormalsForSize.Count; $idx++) {
            $wObj = $ascendingNormalsForSize[$idx]
            # Each successive layer grows by the increment in both dimensions
            $targetW = [Math]::Min($monWidth,  $baseW + ($idx * $minIncrementW))
            $targetH = [Math]::Min($monHeight, $baseH + ($idx * $minIncrementH))
            $sizeMap[$wObj.Handle.ToInt64()] = @{ W = $targetW; H = $targetH; Layer = $idx }
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
    }

    # Cascade settings: fixed offset between each layer
    $startX = $work.Left + 20
    $startY = $work.Top + 20
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
