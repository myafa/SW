# =====================================================================
# SW 반입 전 사전 검증 스크립트 (엑셀 체크리스트 작성 지원용)
# 기능: 현재 폴더 및 하위 폴더의 모든 파일 해시(SHA-256)와 서명 추출
# =====================================================================

# PowerShell UTF-8 콘솔 출력 시 한글(CJK)이 두 번 중복 출력되는 버그 방지를 위해 
# 시스템 기본 인코딩(ANSI/CP949)으로 명시적 초기화합니다.
[Console]::OutputEncoding = [System.Text.Encoding]::Default

# 1. 실행 환경 설정
$currentPath = (Get-Item -Path ".\" -Verbose).FullName
$reportName  = "SW_Verification_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$reportPath  = Join-Path -Path $currentPath -ChildPath $reportName

# 스크립트 자신 파일명 확인 (검사 제외용)
$scriptName = $MyInvocation.MyCommand.Name
if ([string]::IsNullOrEmpty($scriptName)) { $scriptName = "SW_Check_Report.ps1" }

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "[SW 유해성 자가 검증 스크립트]" -ForegroundColor Cyan
Write-Host "==========================================`n"

# 2. 대상 파일 및 폴더 수집 (스크립트 자신과 검증기, 기존 txt 리포트 제외)
$files = Get-ChildItem -Path $currentPath -Recurse | Where-Object { 
    $_.Name -ne $scriptName -and $_.Name -ne "SWverifier.ps1" -and $_.Extension -ne ".txt" 
}

if ($null -eq $files -or $files.Count -eq 0) {
    Write-Host "[경고] 검사할 파일이 없습니다. (스크립트/txt 제외)" -ForegroundColor Yellow
    Read-Host "엔터를 누르면 종료됩니다..."
    exit
}

# 3. 폴더 포함 여부 사전 검사 (전체 폴더 리스트 출력용)
$folders = $files | Where-Object { $_.PSIsContainer }
if ($folders.Count -gt 0) {
    Write-Host "`n[오류] 다음 폴더들이 감지되었습니다:" -ForegroundColor Red
    foreach ($folder in $folders) {
        Write-Host " - $($folder.FullName)" -ForegroundColor Red
    }
    Write-Host "`n하위 폴더가 포함된 경우 zip으로 압축해서 다시 스크립트를 실행해 주세요.`n" -ForegroundColor Red
    Read-Host "엔터를 누르면 종료됩니다..."
    exit
}

Write-Host "총 $($files.Count)개의 파일 검사를 시작합니다...`n"

# 3. 엑셀 붙여넣기용 헤더 작성 (Tab 구분자로 엑셀 호환성 극대화)
$reportContent = @()
$reportContent += "파일명`t해시값(SHA-256)`t전자서명(Status)`t서명자(Signer)"
$reportContent += "-" * 100

# 4. 파일별 정보 추출 및 진행률 표시
$i = 1
foreach ($file in $files) {
    Write-Progress -Activity "파일 검사 중..." -Status "($i / $($files.Count)) $($file.Name)" -PercentComplete (($i / $files.Count) * 100)
    
    # 해시 추출
    try {
        $hash = (Get-FileHash -Path $file.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
    } catch {
        $hash = "해시 추출 실패 (권한/사용중)"
    }

        # 압축 파일의 경우 전자서명 추출 배제 (N/A 처리)
        if ($file.Extension -match "\.(zip|7z|rar|tar|gz|bz2)$") {
            $sigStatusMapped = "N/A(없음)"
            $signer = "N/A(없음)"
        } else {
            # 일반 파일 전자서명 검증
            try {
                $sig = Get-AuthenticodeSignature -FilePath $file.FullName -ErrorAction Stop
                
                # Excel 드롭다운 상태 매핑: [Valid(정상) / Invalid(경고) / N/A(없음)]
                if ($sig.Status -eq "Valid") {
                    $sigStatusMapped = "Valid(정상)"
                    $signer = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject -replace 'CN=', '' -split ',' | Select-Object -First 1 } else { "서명 없음" }
                } elseif ($sig.Status -eq "NotSigned" -or $sig.Status -eq "UnknownError") {
                    $sigStatusMapped = "N/A(없음)"
                    $signer = "서명 없음"
                } else {
                    $sigStatusMapped = "Invalid(경고)"
                    $signer = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject -replace 'CN=', '' -split ',' | Select-Object -First 1 } else { "서명 없음" }
                }

            } catch {
                $sigStatusMapped = "N/A(없음)"
                $signer = "확인 불가"
            }
        } # This brace closes the 'else' block.
    # 결과 조합 (Tab 구분자)
    $reportContent += "$($file.Name)`t$hash`t$sigStatusMapped`t$signer"
    $i++
}

# 5. 1차 리포트 파일 생성 (본문 저장)
$bodyText = $reportContent -join "`r`n"
$bodyText | Out-File -FilePath $reportPath -Encoding UTF8

# 6. 리포트 자체 무결성 검증을 위한 '봉인(Seal)' 생성
# 방금 만든 본문 txt 파일의 해시값을 계산
$reportHash = (Get-FileHash -Path $reportPath -Algorithm SHA256).Hash

# 봉인(Hash) 값을 리포트 맨 끝에 추가 (개행 문자 추가하여 분리)
$sealText = "`r`n========================================================`r`n"
$sealText += "[REPORT_INTEGRITY_SEAL] (보안팀 검증용)`r`n"
$sealText += "아래 해시값은 본문 내용(이 텍스트 위까지)의 SHA-256 해시입니다.`r`n"
$sealText += "내용이 1글자라도 수정되면 아래 해시값과 본문의 실제 해시값이 불일치하게 됩니다.`r`n"
$sealText += "SEAL_HASH: $reportHash`r`n"
$sealText += "========================================================"

Add-Content -Path $reportPath -Value $sealText -Encoding UTF8

Write-Host "`n[완료] 검사가 성공적으로 끝났습니다." -ForegroundColor Green
Write-Host "결과 파일이 같은 폴더에 생성되었습니다: $reportName"
Write-Host "생성된 txt 파일을 제출해 주세요.`n"


Read-Host "엔터를 누르면 종료됩니다..."
