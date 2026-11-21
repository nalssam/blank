# Installs PowerPoint only, unattended, with the Office Deployment Tool. No sign-in, no activation.
param(
    [string]$Product = 'O365ProPlusRetail',
    [string]$Channel = 'Current',
    [string]$Language = 'en-us'
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$dir = 'C:\odt'
New-Item -ItemType Directory -Force $dir | Out-Null

$exclude = 'Access', 'Bing', 'Excel', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'Publisher', 'Teams', 'Word'
$ex = ($exclude | ForEach-Object { "      <ExcludeApp ID=`"$_`" />" }) -join "`r`n"
$xml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$Channel">
    <Product ID="$Product">
      <Language ID="$Language" />
$ex
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Updates Enabled="FALSE" />
</Configuration>
"@
Set-Content -Path "$dir\config.xml" -Value $xml -Encoding ASCII
Get-Content "$dir\config.xml"

$sw = [Diagnostics.Stopwatch]::StartNew()
Invoke-WebRequest 'https://officecdn.microsoft.com/pr/wsus/setup.exe' -OutFile "$dir\setup.exe"
"ODT downloaded: $((Get-Item "$dir\setup.exe").Length) bytes"
$p = Start-Process "$dir\setup.exe" -ArgumentList '/configure', "$dir\config.xml" -Wait -PassThru
"setup.exe exit code: $($p.ExitCode) after $([int]$sw.Elapsed.TotalSeconds)s"

$exe = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
if (-not (Test-Path $exe)) {
    Get-ChildItem "$env:TEMP" -Filter '*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 3 |
        ForEach-Object { "---- $($_.FullName)"; Get-Content $_.FullName -Tail 40 }
    throw "PowerPoint was not installed"
}
"PowerPoint: $((Get-Item $exe).VersionInfo.FileVersion)"
Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' |
    Select-Object ProductReleaseIds, Platform, VersionToReport, CDNBaseUrl, AudienceId | Format-List

# First-run UI that would block automation.
$hkcu = 'HKCU:\Software\Microsoft\Office\16.0'
function Set-Dword($path, $name, $value) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name $name -Value $value -Type DWord
}
Set-Dword "$hkcu\Common\General" 'ShownFirstRunOptin' 1
Set-Dword "$hkcu\Common\General" 'FirstRunTime' 1
Set-Dword "$hkcu\Common\General" 'ShownFileFmtPrompt' 1
Set-Dword "$hkcu\FirstRun" 'BootedRTM' 1
Set-Dword "$hkcu\FirstRun" 'disablemovie' 1
Set-Dword "$hkcu\Registration" 'AcceptAllEulas' 1
Set-Dword "$hkcu\Common\Privacy\SettingsStore\Anonymous" 'OptionalConnectedExperiencesNoticeVersion' 2
Set-Dword "$hkcu\Common\Privacy\SettingsStore\Anonymous" 'ConnectedExperiencesNoticeVersion' 1
Set-Dword "$hkcu\Common\Privacy\SettingsStore\Anonymous" 'RequiredDiagnosticDataNoticeVersion' 1
Set-Dword "$hkcu\Common\Privacy\SettingsStore\Anonymous" 'FTEDismissed' 1
Set-Dword "$hkcu\PowerPoint\Options" 'DisableBootToOfficeStart' 1
Set-Dword "$hkcu\PowerPoint\Security\ProtectedView" 'DisableInternetFilesInPV' 1
Set-Dword "$hkcu\PowerPoint\Security\ProtectedView" 'DisableUnsafeLocationsInPV' 1
Set-Dword "$hkcu\PowerPoint\Security\ProtectedView" 'DisableAttachmentsInPV' 1
Set-Dword "$hkcu\Common\Licensing" 'DisableActivationUI' 1

"WebView2 runtime:"
Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}' -ErrorAction SilentlyContinue |
    Select-Object name, pv | Format-List
