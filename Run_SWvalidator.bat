@echo off
:: GitHub 다운로드 시 발생하는 인코딩(BOM 유실) 및 권한(실행 정책) 문제 해결 래퍼 스크립트
chcp 65001 >nul
echo [안내] SWvalidator 실행을 준비합니다. (보안 권한 우회 및 인코딩 강제 교정 적용)
echo ==============================================================================
powershell -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content -Path 'SWvalidator.ps1' -Encoding UTF8 -Raw)"
