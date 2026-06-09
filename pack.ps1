# SYQ 배포 패키지 생성 스크립트
# 실행: PowerShell에서 .\pack.ps1

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataDir     = Join-Path $ProjectRoot "data"
$Date        = Get-Date -Format "yyMMdd"
$OutputZip   = Join-Path $ProjectRoot "SYQ_배포_$Date.zip"

if (Test-Path $OutputZip) { Remove-Item $OutputZip -Force }

Write-Host "패키징 중..."
Compress-Archive -Path "$DataDir\*" -DestinationPath $OutputZip

Write-Host ""
Write-Host "완료: $OutputZip"
Write-Host ""
Write-Host "배포 방법:"
Write-Host "  1. 위 zip 파일을 전달하세요"
Write-Host "  2. 상대방은 압축 해제 후 start.cmd 더블클릭"
Write-Host "  3. 브라우저에서 http://localhost:5273 접속"
