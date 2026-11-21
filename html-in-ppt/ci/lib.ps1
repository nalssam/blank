# Shared helpers: window listing, screenshots, real mouse clicks.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class Win {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
    static string Describe(IntPtr h) {
        if (h == IntPtr.Zero) return "(none)";
        var t = new StringBuilder(256); GetWindowText(h, t, 256);
        var c = new StringBuilder(256); GetClassName(h, c, 256);
        uint pid; GetWindowThreadProcessId(h, out pid);
        string pn = "?"; try { pn = System.Diagnostics.Process.GetProcessById((int)pid).ProcessName; } catch {}
        return String.Format("{0}/{1} \"{2}\"", pn, c, t);
    }
    public static string At(int x, int y) {
        POINT p; p.X = x; p.Y = y;
        IntPtr h = WindowFromPoint(p);
        return Describe(h) + " in " + Describe(GetAncestor(h, 2));
    }
    public static string Foreground() { return Describe(GetForegroundWindow()); }
    // Modal Office dialogs (first-run notices and the like): "hwnd|L|T|R|B|title"
    public static List<string> Dialogs() {
        var res = new List<string>();
        EnumWindows((h, l) => {
            if (!IsWindowVisible(h)) return true;
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() != "NUIDialog") return true;
            var t = new StringBuilder(256); GetWindowText(h, t, 256);
            RECT r; GetWindowRect(h, out r);
            res.Add(String.Format("{0}|{1}|{2}|{3}|{4}|{5}", h.ToInt64(), r.L, r.T, r.R, r.B, t));
            return true;
        }, IntPtr.Zero);
        return res;
    }
    public static void Activate(IntPtr h) { ShowWindow(h, 5); BringWindowToTop(h); SetForegroundWindow(h); }
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    public static List<string> List() {
        var res = new List<string>();
        EnumWindows((h, l) => {
            if (!IsWindowVisible(h)) return true;
            var t = new StringBuilder(256); GetWindowText(h, t, 256);
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            uint pid; GetWindowThreadProcessId(h, out pid);
            RECT r; GetWindowRect(h, out r);
            if (t.Length == 0 && (r.R - r.L) < 50) return true;
            string pn = "?"; try { pn = System.Diagnostics.Process.GetProcessById((int)pid).ProcessName; } catch {}
            res.Add(String.Format("{0,-12} class={1,-28} rect=({2},{3},{4},{5}) title=\"{6}\"", pn, c, r.L, r.T, r.R, r.B, t));
            return true;
        }, IntPtr.Zero);
        return res;
    }
    public static void Click(int x, int y) {
        SetCursorPos(x, y); System.Threading.Thread.Sleep(150);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero); System.Threading.Thread.Sleep(60);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero); System.Threading.Thread.Sleep(150);
    }
}
"@
[Win]::SetProcessDPIAware() | Out-Null

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
"Screen: $($screen.Width)x$($screen.Height)   session: $([Diagnostics.Process]::GetCurrentProcess().SessionId)   user: $env:USERNAME"

function Shot([string]$name, [string]$dir = $OutDir) {
    try {
        $bmp = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
        $bmp.Save((Join-Path $dir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
        # A coarse 8x4 grid of colours, so the log alone shows whether something was drawn.
        $cells = @()
        foreach ($gy in 0..3) { foreach ($gx in 0..7) {
            $c = $bmp.GetPixel([int](($gx + 0.5) * $screen.Width / 8), [int](($gy + 0.5) * $screen.Height / 4))
            $cells += ('{0:x2}{1:x2}{2:x2}' -f $c.R, $c.G, $c.B)
        } }
        "SHOT $name grid=" + ($cells -join ' ')
        $g.Dispose(); $bmp.Dispose()
    } catch { "SHOT $name FAILED: $_" }
}
function Windows([string]$label) {
    "---- windows ($label)"
    [Win]::List() | ForEach-Object { "  $_" }
}
# Closes Office's own modal notices (for example "Your privacy option") that would block everything else.
function DismissDialogs([string]$label) {
    foreach ($try in 1..4) {
        $ds = [Win]::Dialogs()
        if ($ds.Count -eq 0) { return }
        foreach ($d in $ds) {
            $f = $d.Split('|')
            "DIALOG ($label) '$($f[5])' rect=($($f[1]),$($f[2]),$($f[3]),$($f[4])) -> closing"
            [Win]::Activate([IntPtr][int64]$f[0])
            Start-Sleep -Milliseconds 400
            [Win]::Click([int]$f[3] - 130, [int]$f[4] - 15)     # the Close / OK button sits bottom right
            Start-Sleep -Seconds 2
        }
    }
    if ([Win]::Dialogs().Count -gt 0) { "DIALOG ($label) still open" }
}
# Moves the pointer the way a hand does (a jump straight to the target is not what a presenter does).
function Hover([int]$x, [int]$y) {
    $c = [System.Windows.Forms.Cursor]::Position
    foreach ($k in 1..6) {
        [Win]::SetCursorPos([int]($c.X + ($x - $c.X) * $k / 6), [int]($c.Y + ($y - $c.Y) * $k / 6)) | Out-Null
        Start-Sleep -Milliseconds 70
    }
}
# A real click that also records which window was under the pointer and which had the keyboard.
function ClickAt([int]$x, [int]$y, [string]$what) {
    "CLICK $what at ($x,$y) over [$([Win]::At($x, $y))] foreground=[$([Win]::Foreground())]"
    [Win]::Click($x, $y)
}
# Stage coordinates (1920x1080) -> screen coordinates for a rectangle showing the slide.
function StagePoint([double]$sx, [double]$sy, $rect) {
    $k = [Math]::Min($rect.W / 1920.0, $rect.H / 1080.0)
    $ox = $rect.X + ($rect.W - 1920 * $k) / 2
    $oy = $rect.Y + ($rect.H - 1080 * $k) / 2
    return @([int]($ox + $sx * $k), [int]($oy + $sy * $k))
}

