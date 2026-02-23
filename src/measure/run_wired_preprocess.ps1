# PowerShell - 한글 출력 UTF-8 설정
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

Set-Location $PSScriptRoot
chcp 65001 | Out-Null

Write-Host "[학교별 측정 값 현황] 유선망 전처리"
Write-Host "- 상단 RUN_REGION 값에 따라 충남/대전/둘 다 실행"
python wired_preprocess_v1.py

Read-Host "엔터를 눌러 종료"
