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
DismissDialogs 'after start'
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
    foreach ($d in 'D3', 'U1', 'N1', 'Z4', 'Z5') {
        $f = Join-Path $Variants "$d.pptx"
        if (-not (Test-Path $f)) { continue }
        Mark "open-$d"
        $p = Open-Deck $f
        if (-not $p) { "A/$d could not be opened"; continue }
        Start-Sleep -Seconds 16
        Shot "A-$d"
        if ($d -eq 'D3') { WebViews 'first add-in document open' }
        DismissDialogs "after $d"
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
    $p = StagePoint 570 683 $rect
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
    $p = StagePoint 570 683 $full                      # the text box on slide 1
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
            $p = StagePoint 500 442 $full
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
            $p = StagePoint 1394 728 $full; ClickAt $p[0] $p[1] 'show, draw button'; Start-Sleep -Seconds 3
            "after draw: slide show is on slide $($show.View.CurrentShowPosition)"
            Shot '24-show-slide4-rolled'
        }
    }
    try { $show.View.Exit() } catch { "show exit failed: $_" }
}
Start-Sleep -Seconds 2
Shot '30-after-show'

# ---- A presenter's way: keyboard to turn pages, the pointer moved by hand, no COM navigation ----
function Keys([string]$k) { [System.Windows.Forms.SendKeys]::SendWait($k) }
function HumanShow($deck, [string]$tag) {
    Mark "human-$tag"
    $sh = $null
    try { $sh = $deck.SlideShowSettings.Run() } catch { "[$tag] SlideShow.Run failed: $_"; return }
    Start-Sleep -Seconds 3
    try { $sh.Activate() } catch {}
    Start-Sleep -Seconds 3
    "[$tag] show started on slide $($sh.View.CurrentShowPosition); foreground=[$([Win]::Foreground())]"
    $steps = @(
        @{ s = 1; x = 570; y = 683; what = 'text box' },
        @{ s = 2; x = 500; y = 442; what = 'slider' },
        @{ s = 3; x = 1387; y = 490; what = 'option B' },
        @{ s = 4; x = 1394; y = 728; what = 'draw button' }
    )
    foreach ($st in $steps) {
        $guard = 0
        while ($sh.View.CurrentShowPosition -lt $st.s -and $guard -lt 8) { Keys '{RIGHT}'; Start-Sleep -Seconds 2; $guard++ }
        while ($sh.View.CurrentShowPosition -gt $st.s -and $guard -lt 16) { Keys '{LEFT}'; Start-Sleep -Seconds 2; $guard++ }
        "[$tag] on slide $($sh.View.CurrentShowPosition) (wanted $($st.s)) after $guard key presses"
        Start-Sleep -Seconds 3
        $p = StagePoint $st.x $st.y $full
        foreach ($n in 1, 2) {
            Hover $p[0] $p[1]; Start-Sleep -Milliseconds 700
            ClickAt $p[0] $p[1] "[$tag] slide $($st.s) $($st.what) #$n"
            Start-Sleep -Seconds 2
            $pos = $sh.View.CurrentShowPosition
            "[$tag] slide $($st.s) $($st.what) click #${n}: show is on slide $pos"
            if ($pos -ne $st.s) { Keys '{LEFT}'; Start-Sleep -Seconds 3; "[$tag] went back with LEFT: slide $($sh.View.CurrentShowPosition)" }
        }
        Shot ("40-$tag-slide{0}" -f $st.s)
        if ($st.s -eq 1) {
            Keys 'zz'; Start-Sleep -Seconds 1; Shot "41-$tag-slide1-typed"
            # PageDown must turn the page even while the caret is in a text box
            Mark "human-$tag-pgdn"
            Keys '{PGDN}'; Start-Sleep -Seconds 3
            "[$tag] PageDown while typing in the text box: show is on slide $($sh.View.CurrentShowPosition)"
        }
        if ($st.s -eq 3) {
            Mark "human-$tag-key"
            Keys '{RIGHT}'; Start-Sleep -Seconds 3
            "[$tag] RIGHT after clicking inside slide 3: show is on slide $($sh.View.CurrentShowPosition)"
            # no click in between: does the keyboard still work right after a forwarded page turn?
            Keys '{LEFT}'; Start-Sleep -Seconds 3
            "[$tag] LEFT right after the forwarded page turn: show is on slide $($sh.View.CurrentShowPosition)"
        }
    }
    # the page-turn button drawn by the deck itself (calls liveSlide.next())
    $p = StagePoint 1864 1028 $full
    Hover $p[0] $p[1]; Start-Sleep -Milliseconds 700
    ClickAt $p[0] $p[1] "[$tag] in-deck next button"
    Start-Sleep -Seconds 3
    "[$tag] after the in-deck next button: show is on slide $($sh.View.CurrentShowPosition)"
    Shot "42-$tag-after-deck-next"
    # Esc with the keyboard inside the HTML: can the presenter still end the show?
    # (Not on the last slide of a deck whose clicks turn pages: that click would leave the show.)
    Mark "human-$tag-esc"
    try { if ($sh.View.CurrentShowPosition -ge $Slides) { Keys '{LEFT}'; Start-Sleep -Seconds 3 } } catch {}
    $p = StagePoint 960 620 $full
    Hover $p[0] $p[1]; Start-Sleep -Milliseconds 500
    ClickAt $p[0] $p[1] "[$tag] plain spot before Esc"
    Start-Sleep -Seconds 2
    foreach ($n in 1, 2) {
        Keys '{ESC}'; Start-Sleep -Seconds 3
        $left = -1; try { $left = $ppt.SlideShowWindows.Count } catch {}
        "[$tag] after Esc #${n}: slide show windows = $left"
        if ($left -eq 0) { break }
    }
    try { if ($ppt.SlideShowWindows.Count -gt 0) { $sh.View.Exit() } } catch {}
    Start-Sleep -Seconds 2
}

# The same show in a window ("browsed by an individual"): another hosting path for the add-in.
function WindowShow($deck, [string]$tag) {
    Mark "window-$tag"
    $sh = $null
    try { $deck.SlideShowSettings.ShowType = 2; $sh = $deck.SlideShowSettings.Run() } catch { "[$tag] window show failed: $_"; return }
    Start-Sleep -Seconds 3
    try { $sh.Activate() } catch {}
    try { [Win]::ShowWindow([IntPtr]$sh.HWND, 3) | Out-Null } catch {}     # maximised, so nothing of the editing window shows through
    try { [Win]::Activate([IntPtr]$sh.HWND) } catch {}
    Start-Sleep -Seconds 6
    Shot "45-$tag-window-slide1"
    "[$tag] window show on slide $($sh.View.CurrentShowPosition); client=$([Win]::ClientRect([IntPtr]$sh.HWND)); foreground=[$([Win]::Foreground())]"
    # where is the slide drawn? slide 1 has one flat background colour
    $box = [Win]::ColorBox(0xf4, 0xf1, 0xea, 2).Split('|')
    "[$tag] slide background box: ($($box[0]),$($box[1]))-($($box[2]),$($box[3])) pixels=$($box[4])"
    if ([int]$box[4] -gt 5000) {
        $rect = @{ X = [int]$box[0]; Y = [int]$box[1]; W = ([int]$box[2] - [int]$box[0]); H = ([int]$box[3] - [int]$box[1]) }
        $p = StagePoint 570 683 $rect
        foreach ($n in 1, 2) {
            Hover $p[0] $p[1]; Start-Sleep -Milliseconds 700
            ClickAt $p[0] $p[1] "[$tag] window show, text box #$n"; Start-Sleep -Seconds 2
            "[$tag] window show after click #${n}: slide $($sh.View.CurrentShowPosition)"
        }
        Keys 'ww'; Start-Sleep -Seconds 1
        Shot "46-$tag-window-typed"
        Keys '{PGDN}'; Start-Sleep -Seconds 4
        "[$tag] window show after PageDown: slide $($sh.View.CurrentShowPosition)"
        Shot "47-$tag-window-slide2"
    } else { "[$tag] window show: the slide was not found on screen" }
    try { $sh.View.Exit() } catch {}
    Start-Sleep -Seconds 2
    try { $deck.SlideShowSettings.ShowType = 1 } catch {}
}
HumanShow $pres 'clickadv'       # converter default: a click turns the page where the HTML is not running
WindowShow $pres 'clickadv'

if ($Variants -and (Test-Path $Variants)) {
    # the same deck built with --no-click-advance: clicks never turn the page
    $c1 = Join-Path $Variants 'C1.pptx'
    if (Test-Path $c1) {
        $pc = Open-Deck $c1
        if ($pc) {
            Start-Sleep -Seconds 12
            HumanShow $pc 'noclick'
            try { $pc.Saved = -1; $pc.Close() } catch {}
        }
    }
    # cold start: open and present at once, as a teacher who double-clicks the file and presses F5
    $cold = Join-Path $Variants 'D1.pptx'
    if (Test-Path $cold) {
        Mark 'cold'
        $pd = Open-Deck $cold
        if ($pd) {
            $sc = $null
            try { $sc = $pd.SlideShowSettings.Run() } catch { "cold Run failed: $_" }
            if ($sc) {
                try { $sc.Activate() } catch {}
                foreach ($t in 1, 2, 3, 5, 8) { Start-Sleep -Seconds $(if ($t -le 3) { 1 } elseif ($t -eq 5) { 2 } else { 3 }); Shot ("50-cold-slide1-t{0}" -f $t) }
                Keys '{RIGHT}'
                foreach ($t in 1, 2, 4, 7) { Start-Sleep -Seconds $(if ($t -le 2) { 1 } elseif ($t -eq 4) { 2 } else { 3 }); Shot ("51-cold-slide2-t{0}" -f $t) }
                "cold: show is on slide $($sc.View.CurrentShowPosition)"
                try { $sc.View.Exit() } catch {}
            }
            try { $pd.Saved = -1; $pd.Close() } catch {}
        }
    }
    # a PC without the add-in, in a slide show
    $u1 = Join-Path $Variants 'U1.pptx'
    if (Test-Path $u1) {
        $pu = Open-Deck $u1
        if ($pu) {
            Start-Sleep -Seconds 8
            $su = $null
            try { $su = $pu.SlideShowSettings.Run() } catch { "U1 Run failed: $_" }
            if ($su) { try { $su.Activate() } catch {}; Start-Sleep -Seconds 5; Shot '60-unregistered-show-slide1'; Keys '{RIGHT}'; Start-Sleep -Seconds 4; Shot '61-unregistered-show-slide2'; try { $su.View.Exit() } catch {} }
            try { $pu.Saved = -1; $pu.Close() } catch {}
        }
    }
}
# Last, because its outside link may bring a browser window to the front: the deck that uses
# localStorage, alert(), confirm(), history and links on purpose (results arrive through the probe).
if ($Variants -and (Test-Path (Join-Path $Variants 'B1.pptx'))) {
    Mark 'open-B1'
    $pb = Open-Deck (Join-Path $Variants 'B1.pptx')
    if ($pb) {
        Start-Sleep -Seconds 14; Shot 'A-B1'
        Start-Sleep -Seconds 4; Shot 'A-B1-late'; Windows 'after the self-test deck'
        try { $pb.Saved = -1; $pb.Close() } catch { "close B1 failed: $_" }
    }
}
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
