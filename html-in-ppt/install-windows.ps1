# Live HTML 추가 기능을 이 PC의 PowerPoint에 등록한다 (현재 사용자만, 관리자 권한 불필요).
# 사용:  PowerShell에서   .\install-windows.ps1          (등록)
#                          .\install-windows.ps1 -Remove  (해제)
param([switch]$Remove)

$id = '6f2b8c1e-4d7a-4e55-9a3b-2c8f5d1e7a90'
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$manifest = Join-Path $PSScriptRoot 'addin\manifest.xml'

if ($Remove) {
    if (Test-Path $key) { Remove-ItemProperty -Path $key -Name $id -ErrorAction SilentlyContinue }
    Write-Host '등록을 해제했습니다. PowerPoint를 다시 시작하세요.'
    exit 0
}

if (-not (Test-Path $manifest)) { Write-Error "manifest.xml 이 없습니다: $manifest"; exit 1 }
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
Set-ItemProperty -Path $key -Name $id -Value $manifest
Write-Host "등록 완료: $manifest"
Write-Host 'PowerPoint를 완전히 종료했다가 다시 열어 주세요.'
