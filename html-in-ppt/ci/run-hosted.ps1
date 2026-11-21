# End-to-end check against the really hosted add-in page and the skill exactly as distributed:
# the deck was built from the skill zip's own template; PowerPoint loads app.html from the web.
# There is no probe here, so the evidence is the DOM read over CDP plus screenshots.
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [string]$UrlPart = '/live-html/app.html',
    [int]$LoadWait = 9
)
$ErrorActionPreference = 'Continue'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$Pptx = (Resolve-Path $Pptx).Path
. (Join-Path $PSScriptRoot 'lib.ps1')

function Dom([string]$label) { & node (Join-Path $PSScriptRoot 'cdp.js') $UrlPart $label 2>&1 | ForEach-Object { "$_" } }

$ppt = $null
try { $ppt = New-Object -ComObject PowerPoint.Application } catch { "COM create failed: $_" }
if (-not $ppt) { Shot '00-no-com'; exit 1 }
try { $ppt.Visible = -1 } catch {}
"PowerPoint version: $($ppt.Version) build: $($ppt.Build)"
Start-Sleep -Seconds 8
try { $ppt.WindowState = 3 } catch {}
$pres = $null
foreach ($try in 1..4) {
    try { $pres = $ppt.Presentations.Open($Pptx, 0, 0, -1); break }
    catch { "Open attempt $try failed: $($_.Exception.Message)"; Start-Sleep -Seconds 5 }
}
if (-not $pres) { "Could not open the presentation"; exit 1 }
"Opened: $($pres.FullName) slides=$($pres.Slides.Count)"
Start-Sleep -Seconds ($LoadWait + 8)
Shot 'H02-opened'; Windows 'opened'
Dom 'opened'

$w = $ppt.ActiveWindow
function EditRect {
    $x0 = $w.PointsToScreenPixelsX(0); $y0 = $w.PointsToScreenPixelsY(0)
    $x1 = $w.PointsToScreenPixelsX($pres.PageSetup.SlideWidth); $y1 = $w.PointsToScreenPixelsY($pres.PageSetup.SlideHeight)
    return @{ X = $x0; Y = $y0; W = ($x1 - $x0); H = ($y1 - $y0) }
}

# Editing view
$w.View.GotoSlide(2); Start-Sleep -Seconds $LoadWait
$p = StagePoint 383 611 (EditRect); ClickAt $p[0] $p[1] 'edit, flip card 1'; Start-Sleep -Seconds 2
Shot 'H10-edit-slide2-flip'; Dom 'edit slide 2 after one click on card 1'
$w.View.GotoSlide(4); Start-Sleep -Seconds $LoadWait
$p = StagePoint 1392 412 (EditRect); ClickAt $p[0] $p[1] 'edit, option B'; Start-Sleep -Seconds 2
Shot 'H11-edit-slide4-answer'; Dom 'edit slide 4 after click on option B'
$w.View.GotoSlide(3); Start-Sleep -Seconds $LoadWait
$p = StagePoint 600 347 (EditRect); ClickAt $p[0] $p[1] 'edit, denominator slider'; Start-Sleep -Seconds 2
Shot 'H12-edit-slide3-slider'; Dom 'edit slide 3 after click on the slider'

# Slide show
$full = @{ X = 0; Y = 0; W = $screen.Width; H = $screen.Height }
$show = $null
try { $show = $pres.SlideShowSettings.Run() } catch { "SlideShow.Run failed: $_" }
if ($show) {
    Start-Sleep -Seconds 3
    try { $show.Activate() } catch {}
    try { [Win]::Activate([IntPtr]$show.HWND) } catch {}
    "show activated; foreground=[$([Win]::Foreground())]"
    foreach ($t in 2, 6) { Start-Sleep -Seconds 3; Shot ("H20-show-slide1-t{0}" -f $t) }
    foreach ($i in 2..6) {
        try { $show.View.GotoSlide($i) } catch { "show GotoSlide($i) failed: $_" }
        Start-Sleep -Seconds 2; Shot ("H21-show-slide{0}-t2" -f $i)
        Start-Sleep -Seconds ($LoadWait - 2); Shot ("H21-show-slide{0}" -f $i)
        if ($i -eq 2) {
            $p = StagePoint 1537 611 $full; ClickAt $p[0] $p[1] 'show, flip card 3'; Start-Sleep -Seconds 2
            "position after click: $($show.View.CurrentShowPosition)"; Shot 'H22-show-slide2-flip'; Dom 'show slide 2 after click on card 3'
        }
        if ($i -eq 4) {
            $p = StagePoint 529 412 $full; ClickAt $p[0] $p[1] 'show, option A'; Start-Sleep -Seconds 2
            "position after click: $($show.View.CurrentShowPosition)"; Shot 'H23-show-slide4-wrong'; Dom 'show slide 4 after click on option A'
        }
        if ($i -eq 6) {
            foreach ($y in 417, 553, 689) { $p = StagePoint 960 $y $full; ClickAt $p[0] $p[1] "show, check at $y"; Start-Sleep -Seconds 1 }
            "position after clicks: $($show.View.CurrentShowPosition)"; Shot 'H24-show-slide6-checked'; Dom 'show slide 6 after three checks'
            [System.Windows.Forms.SendKeys]::SendWait('{LEFT}'); Start-Sleep -Seconds 4
            "after LEFT with the keyboard inside the HTML: slide show is on slide $($show.View.CurrentShowPosition)"
            Shot 'H25-show-after-left'
        }
    }
    try { $show.View.Exit() } catch {}
}
Start-Sleep -Seconds 2
Shot 'H30-after-show'
try { $pres.SaveCopyAs((Join-Path $OutDir 'hosted-resaved.pptx')) } catch { "SaveCopyAs failed: $($_.Exception.Message)" }
try { $pres.Saved = -1; $pres.Close() } catch {}
try { $ppt.Quit() } catch {}
Start-Sleep -Seconds 2
"---- Office alerts (event log)"
try {
    Get-WinEvent -LogName 'OAlerts' -MaxEvents 8 -ErrorAction Stop | ForEach-Object { "  [$($_.TimeCreated.ToString('HH:mm:ss'))] " + (($_.Message -replace '\s+', ' ').Trim()) }
} catch { "  (none: $($_.Exception.Message))" }
