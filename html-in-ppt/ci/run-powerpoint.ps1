# Drives real PowerPoint. Phase A opens a matrix of documents to find out which part of the
# chain works; phase B walks the real deck in the editing view and a slide show with real
# mouse and keyboard input. Screenshots and logs go to -OutDir.
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$Manifest,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [string]$Variants = '',
    [int]$Slides = 5,
    [int]$LoadWait = 10,
    [switch]$ShellLaunch
)
$ErrorActionPreference = 'Continue'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$Pptx = (Resolve-Path $Pptx).Path
$Manifest = (Resolve-Path $Manifest).Path

. (Join-Path $PSScriptRoot 'lib.ps1')

function Mark([string]$label) {
    try { Invoke-WebRequest "https://localhost:3000/mark?l=$label" -UseBasicParsing | Out-Null }
    catch { "MARK $label FAILED (server down?): $($_.Exception.Message)" }
}
function WebViews([string]$label) {
    $procs = @(Get-CimInstance Win32_Process -Filter "Name='msedgewebview2.exe'" -ErrorAction SilentlyContinue)
    "---- webview2 processes ($label): $($procs.Count)"
    $procs | Where-Object { $_.CommandLine -notmatch '--type=' } | ForEach-Object {
        $c = $_.CommandLine
        if ($c.Length -gt 600) { $c = $c.Substring(0, 600) }
        "  $c"
    }
}

# 1) The add-in is already registered (register_addin.py). Turn on Office's add-in runtime log.
$id = ([xml](Get-Content $Manifest -Raw)).OfficeApp.Id
$dev = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'
$rtlog = Join-Path $OutDir 'office-runtime.log'
if (-not (Test-Path "$dev\RuntimeLogging")) { New-Item -Path "$dev\RuntimeLogging" -Force | Out-Null }
Set-ItemProperty -Path "$dev\RuntimeLogging" -Name '(default)' -Value $rtlog
"---- developer registry"
Get-Item $dev | Select-Object -ExpandProperty Property | ForEach-Object { "  $_ = $((Get-ItemProperty $dev).$_)" }

# What Windows itself reports about connectivity (Office refuses to load add-ins when it thinks it is offline).
"---- network"
try {
    $nlm = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}'))
    "  NLM IsConnected=$($nlm.IsConnected) IsConnectedToInternet=$($nlm.IsConnectedToInternet) Connectivity=$($nlm.GetConnectivity())"
} catch { "  NLM failed: $($_.Exception.Message)" }
try { Get-NetConnectionProfile | ForEach-Object { "  profile '$($_.Name)' category=$($_.NetworkCategory) v4=$($_.IPv4Connectivity) v6=$($_.IPv6Connectivity)" } } catch { "  profiles failed: $_" }

# 2) Start PowerPoint: either purely through COM (-Embedding) or the way a person does, then attach.
Mark 'before-powerpoint'
$ppt = $null
if ($ShellLaunch) {
    "launch mode: shell"
    Start-Process 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
    Start-Sleep -Seconds 12
    Shot '00-shell-launched'; Windows 'shell launched'
} else { "launch mode: COM" }
try { $ppt = New-Object -ComObject PowerPoint.Application } catch { "COM create failed: $_" }
if (-not $ppt) { Shot '00-no-com'; Windows 'no com'; exit 1 }
try { $ppt.Visible = -1 } catch { "Visible failed: $_" }
"PowerPoint version: $($ppt.Version) build: $($ppt.Build)"
Start-Sleep -Seconds 10
try { $ppt.WindowState = 3 } catch {}
Shot '00-started'; Windows 'started'

function Open-Deck([string]$path) {
    foreach ($try in 1..5) {
        try { return $ppt.Presentations.Open($path, 0, 0, -1) }
        catch {
            "Open attempt $try failed: $($_.Exception.Message)"
            Shot "01-open-fail-$try"; Windows "open fail $try"
            [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
            Start-Sleep -Seconds 5
        }
    }
    return $null
}

# Phase A: which link of the chain works?
if ($Variants -and (Test-Path $Variants)) {
    foreach ($d in 'D4', 'D5', 'D2', 'D3', 'U1', 'N1', 'Z1', 'Z2', 'Z3', 'Z4', 'Z5') {
        $f = Join-Path $Variants "$d.pptx"
        if (-not (Test-Path $f)) { continue }
        Mark "open-$d"
        $p = Open-Deck $f
        if (-not $p) { "A/$d could not be opened"; continue }
        Start-Sleep -Seconds 16
        Shot "A-$d"
        if ($d -eq 'D4') { WebViews 'first add-in document open' }
        Mark "close-$d"
        try { $p.Saved = -1; $p.Close() } catch { "close $d failed: $_" }
        Start-Sleep -Seconds 2
    }
}

# Phase B: the real deck.
Mark 'open-D1'
$pres = Open-Deck $Pptx
if (-not $pres) { "Could not open the presentation"; try { $ppt.Quit() } catch {}; exit 1 }
"Opened: $($pres.FullName)  slides=$($pres.Slides.Count)  readOnly=$($pres.ReadOnly)"
Start-Sleep -Seconds ($LoadWait + 10)
Shot '02-opened'; Windows 'opened'
"---- CDP view of the add-in webview"
try {
    $targets = Invoke-RestMethod 'http://127.0.0.1:9222/json/list' -TimeoutSec 5
    $targets | ForEach-Object { "  target type=$($_.type) url=$(([string]$_.url).Substring(0, [Math]::Min(120, ([string]$_.url).Length)))" }
} catch { "  CDP not reachable: $($_.Exception.Message)" }
foreach ($s in $pres.Slides) {
    foreach ($sh in $s.Shapes) {
        "slide $($s.SlideIndex): shape '$($sh.Name)' type=$($sh.Type) $([int]$sh.Left),$([int]$sh.Top) $([int]$sh.Width)x$([int]$sh.Height)"
    }
}

# Editing view: visit every slide.
foreach ($i in 1..$Slides) {
    try { $ppt.ActiveWindow.View.GotoSlide($i) } catch { "GotoSlide($i) failed: $_" }
    Start-Sleep -Seconds $LoadWait
    Shot ("10-edit-slide{0}" -f $i)
}

# Editing view: click inside the slide (quiz option B on slide 3) with the real mouse.
try {
    Mark 'edit-click'
    $ppt.ActiveWindow.View.GotoSlide(3); Start-Sleep -Seconds $LoadWait
    $w = $ppt.ActiveWindow
    $x0 = $w.PointsToScreenPixelsX(0); $y0 = $w.PointsToScreenPixelsY(0)
    $x1 = $w.PointsToScreenPixelsX($pres.PageSetup.SlideWidth); $y1 = $w.PointsToScreenPixelsY($pres.PageSetup.SlideHeight)
    "edit view slide rect on screen: ($x0,$y0)-($x1,$y1)"
    $rect = @{ X = $x0; Y = $y0; W = ($x1 - $x0); H = ($y1 - $y0) }
    $p = StagePoint 1385 560 $rect
    ClickAt $p[0] $p[1] 'edit view, option B (1st)'; Start-Sleep -Seconds 2
    ClickAt $p[0] $p[1] 'edit view, option B (2nd)'; Start-Sleep -Seconds 3
    Shot '11-edit-slide3-clicked'
    # page-turn key pressed while the keyboard is inside the HTML (editing view)
    Mark 'edit-key-next'
    [System.Windows.Forms.SendKeys]::SendWait('{RIGHT}'); Start-Sleep -Seconds 3
    try { "after RIGHT in the editing view: selected slide is $($ppt.ActiveWindow.View.Slide.SlideIndex)" } catch { "edit position failed: $_" }
    # typing into the text box of slide 1 (editing view)
    $ppt.ActiveWindow.View.GotoSlide(1); Start-Sleep -Seconds $LoadWait
    $p = StagePoint 600 660 $rect
    ClickAt $p[0] $p[1] 'edit view, text box (1st)'; Start-Sleep -Seconds 1
    ClickAt $p[0] $p[1] 'edit view, text box (2nd)'; Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait('abc 123'); Start-Sleep -Seconds 2
    Shot '12-edit-slide1-typed'
} catch { "edit click failed: $_" }

# Slide show: every slide, with real clicks and typing.
$full = @{ X = 0; Y = 0; W = $screen.Width; H = $screen.Height }
$show = $null
Mark 'slideshow'
try { $show = $pres.SlideShowSettings.Run() } catch { "SlideShow.Run failed: $_" }
if ($show) {
    # A show started through COM is not the foreground window; a person pressing F5 gets that for free.
    Start-Sleep -Seconds 3
    try { $show.Activate() } catch { "show.Activate failed: $_" }
    try { [Win]::Activate([IntPtr]$show.HWND) } catch { "activate hwnd failed: $_" }
    "slide show window activated; foreground=[$([Win]::Foreground())]"
    foreach ($t in 2, 5, 9) { Start-Sleep -Seconds 3; Shot ("20-show-slide1-t{0}" -f $t) }
    Windows 'slide show'
    $p = StagePoint 600 660 $full                      # the text box on slide 1
    ClickAt $p[0] $p[1] 'show, text box (1st)'; Start-Sleep -Seconds 2
    Shot '21a-show-slide1-click1'
    ClickAt $p[0] $p[1] 'show, text box (2nd)'; Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait('hello ppt'); Start-Sleep -Seconds 2
    "after typing: slide show is on slide $($show.View.CurrentShowPosition); foreground=[$([Win]::Foreground())]"
    Shot '21b-show-slide1-typed'
    foreach ($i in 2..$Slides) {
        try { $show.View.GotoSlide($i) } catch { "show GotoSlide($i) failed: $_" }
        Start-Sleep -Seconds 2; Shot ("22-show-slide{0}-t2" -f $i)
        Start-Sleep -Seconds ($LoadWait - 2)
        Shot ("22-show-slide{0}" -f $i)
        if ($i -eq 2) {
            # drag the first slider (monthly growth) to the right
            $p = StagePoint 500 470 $full
            ClickAt $p[0] $p[1] 'show, slider'; Start-Sleep -Seconds 2
            Shot '22b-show-slide2-slider'
            "after slider click: slide show is on slide $($show.View.CurrentShowPosition)"
        }
        if ($i -eq 3) {
            $p = StagePoint 500 560 $full; ClickAt $p[0] $p[1] 'show, option A'; Start-Sleep -Seconds 2
            "after option A: slide show is on slide $($show.View.CurrentShowPosition)"
            $p = StagePoint 1385 560 $full; ClickAt $p[0] $p[1] 'show, option B'; Start-Sleep -Seconds 2
            "after option B: slide show is on slide $($show.View.CurrentShowPosition)"
            Shot '23-show-slide3-answered'
            # the page-turn key, pressed while the keyboard is inside the HTML, must reach PowerPoint
            Mark 'key-next'
            [System.Windows.Forms.SendKeys]::SendWait('{RIGHT}'); Start-Sleep -Seconds 4
            try { "after RIGHT inside the HTML: slide show is on slide $($show.View.CurrentShowPosition)" } catch { "position failed: $_" }
            Shot '23b-show-after-right'
        }
        if ($i -eq 4) {
            $p = StagePoint 1385 800 $full; ClickAt $p[0] $p[1] 'show, draw button'; Start-Sleep -Seconds 3
            "after draw: slide show is on slide $($show.View.CurrentShowPosition)"
            Shot '24-show-slide4-rolled'
        }
    }
    try { $show.View.Exit() } catch { "show exit failed: $_" }
}
Start-Sleep -Seconds 2
Shot '30-after-show'
WebViews 'after slide show'

# Save a copy the way a teacher would, to see what PowerPoint writes back.
try {
    $copy = Join-Path $OutDir 'resaved-by-powerpoint.pptx'
    $pres.SaveCopyAs($copy)
    "Saved copy: $((Get-Item $copy).Length) bytes"
} catch { "SaveCopyAs failed: $($_.Exception.Message)" }

try { $pres.Saved = -1; $pres.Close() } catch { "close failed: $_" }
try { $ppt.Quit() } catch {}
Start-Sleep -Seconds 3

"---- office runtime log"
if (Test-Path $rtlog) { Get-Content $rtlog -Tail 60 } else { "(no runtime log written)" }
"---- Office alerts (event log)"
try {
    Get-WinEvent -LogName 'OAlerts' -MaxEvents 12 -ErrorAction Stop | ForEach-Object {
        "  [$($_.TimeCreated.ToString('HH:mm:ss'))] " + (($_.Message -replace '\s+', ' ').Trim())
    }
} catch { "  (none: $($_.Exception.Message))" }
"---- Wef folder"
$wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
if (Test-Path $wef) {
    Get-ChildItem $wef -Recurse -ErrorAction SilentlyContinue | Select-Object -First 40 | ForEach-Object { "  $($_.FullName.Substring($wef.Length))  $($_.Length)" }
} else { "  (no Wef folder)" }
