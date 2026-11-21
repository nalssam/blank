# Drives real PowerPoint: registers the add-in, opens the deck, walks the slides in the
# editing view and in a slide show, clicks and types with real input events, takes screenshots.
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$Manifest,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [int]$Slides = 5,
    [int]$LoadWait = 12
)
$ErrorActionPreference = 'Continue'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$Pptx = (Resolve-Path $Pptx).Path
$Manifest = (Resolve-Path $Manifest).Path

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

function Shot([string]$name) {
    try {
        $bmp = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
        $bmp.Save((Join-Path $OutDir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
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
# Stage coordinates (1920x1080) -> screen coordinates for a rectangle showing the slide.
function StagePoint([double]$sx, [double]$sy, $rect) {
    $k = [Math]::Min($rect.W / 1920.0, $rect.H / 1080.0)
    $ox = $rect.X + ($rect.W - 1920 * $k) / 2
    $oy = $rect.Y + ($rect.H - 1080 * $k) / 2
    return @([int]($ox + $sx * $k), [int]($oy + $sy * $k))
}

# 1) Register the add-in for the current user and turn on Office's add-in runtime log.
$id = ([xml](Get-Content $Manifest -Raw)).OfficeApp.Id
$dev = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'
if (-not (Test-Path $dev)) { New-Item -Path $dev -Force | Out-Null }
Set-ItemProperty -Path $dev -Name $id -Value $Manifest
$rtlog = Join-Path $OutDir 'office-runtime.log'
if (-not (Test-Path "$dev\RuntimeLogging")) { New-Item -Path "$dev\RuntimeLogging" -Force | Out-Null }
Set-ItemProperty -Path "$dev\RuntimeLogging" -Name '(default)' -Value $rtlog
"Registered add-in $id -> $Manifest"

# 2) Start PowerPoint through COM.
$ppt = $null
try { $ppt = New-Object -ComObject PowerPoint.Application } catch { "COM create failed: $_" }
if (-not $ppt) { Shot '00-no-com'; Windows 'no com'; exit 1 }
try { $ppt.Visible = -1 } catch { "Visible failed: $_" }
"PowerPoint version: $($ppt.Version) build: $($ppt.Build)"
Start-Sleep -Seconds 12
Shot '00-started'; Windows 'started'

# 3) Open the deck (retry while a first-run dialog keeps the app busy).
$pres = $null
foreach ($try in 1..6) {
    try { $pres = $ppt.Presentations.Open($Pptx, 0, 0, -1); break }
    catch {
        "Open attempt $try failed: $($_.Exception.Message)"
        Shot "01-open-fail-$try"; Windows "open fail $try"
        [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
        Start-Sleep -Seconds 6
    }
}
if (-not $pres) { "Could not open the presentation"; try { $ppt.Quit() } catch {}; exit 1 }
"Opened: $($pres.FullName)  slides=$($pres.Slides.Count)  readOnly=$($pres.ReadOnly)"
try { $ppt.WindowState = 3 } catch { "maximize failed: $_" }   # ppWindowMaximized
Start-Sleep -Seconds ($LoadWait + 10)
Shot '02-opened'; Windows 'opened'

# What PowerPoint itself thinks is on each slide.
foreach ($s in $pres.Slides) {
    foreach ($sh in $s.Shapes) {
        "slide $($s.SlideIndex): shape '$($sh.Name)' type=$($sh.Type) $([int]$sh.Left),$([int]$sh.Top) $([int]$sh.Width)x$([int]$sh.Height)"
    }
}

# 4) Editing view: visit every slide.
foreach ($i in 1..$Slides) {
    try { $ppt.ActiveWindow.View.GotoSlide($i) } catch { "GotoSlide($i) failed: $_" }
    Start-Sleep -Seconds $LoadWait
    Shot ("10-edit-slide{0}" -f $i)
}
Windows 'after edit view'

# 4b) Editing view: click inside the slide (quiz option B on slide 3) with the real mouse.
try {
    $ppt.ActiveWindow.View.GotoSlide(3); Start-Sleep -Seconds $LoadWait
    $w = $ppt.ActiveWindow
    $x0 = $w.PointsToScreenPixelsX(0); $y0 = $w.PointsToScreenPixelsY(0)
    $x1 = $w.PointsToScreenPixelsX($pres.PageSetup.SlideWidth); $y1 = $w.PointsToScreenPixelsY($pres.PageSetup.SlideHeight)
    "edit view slide rect on screen: ($x0,$y0)-($x1,$y1)"
    $rect = @{ X = $x0; Y = $y0; W = ($x1 - $x0); H = ($y1 - $y0) }
    $p = StagePoint 1385 560 $rect
    "clicking edit view at $($p[0]),$($p[1])"
    [Win]::Click($p[0], $p[1]); Start-Sleep -Seconds 2
    [Win]::Click($p[0], $p[1]); Start-Sleep -Seconds 3
    Shot '11-edit-slide3-clicked'
} catch { "edit click failed: $_" }

# 5) Slide show: every slide, with real clicks and typing.
$full = @{ X = 0; Y = 0; W = $screen.Width; H = $screen.Height }
$show = $null
try { $show = $pres.SlideShowSettings.Run() } catch { "SlideShow.Run failed: $_" }
if ($show) {
    Start-Sleep -Seconds ($LoadWait + 4)
    Shot '20-show-slide1'; Windows 'slide show'
    # slide 1: click the text box and type
    $p = StagePoint 600 660 $full
    [Win]::Click($p[0], $p[1]); Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait('hello ppt'); Start-Sleep -Seconds 2
    Shot '21-show-slide1-typed'
    foreach ($i in 2..$Slides) {
        try { $show.View.GotoSlide($i) } catch { "show GotoSlide($i) failed: $_" }
        Start-Sleep -Seconds $LoadWait
        Shot ("22-show-slide{0}" -f $i)
        if ($i -eq 3) {
            $p = StagePoint 500 560 $full; [Win]::Click($p[0], $p[1]); Start-Sleep -Seconds 2    # option A (wrong)
            $p = StagePoint 1385 560 $full; [Win]::Click($p[0], $p[1]); Start-Sleep -Seconds 2   # option B (right)
            Shot '23-show-slide3-answered'
        }
        if ($i -eq 4) {
            $p = StagePoint 1385 800 $full; [Win]::Click($p[0], $p[1]); Start-Sleep -Seconds 3   # the draw button
            Shot '24-show-slide4-rolled'
        }
    }
    try { $show.View.Exit() } catch { "show exit failed: $_" }
}
Start-Sleep -Seconds 2
Shot '30-after-show'

# 6) Save a copy the way a teacher would, to see what PowerPoint writes back.
try {
    $copy = Join-Path $OutDir 'resaved-by-powerpoint.pptx'
    $pres.SaveCopyAs($copy)
    "Saved copy: $((Get-Item $copy).Length) bytes"
} catch { "SaveCopyAs failed: $($_.Exception.Message)" }

try { $pres.Saved = -1; $pres.Close() } catch { "close failed: $_" }
try { $ppt.Quit() } catch {}
Start-Sleep -Seconds 3
"---- office runtime log"
if (Test-Path $rtlog) { Get-Content $rtlog -Tail 80 } else { "(no runtime log written)" }
