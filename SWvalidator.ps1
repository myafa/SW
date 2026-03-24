# =====================================================================
# SW 반입 전 사전 검증 스크립트 (엑셀 체크리스트 작성 지원용)
# 기능: 현재 폴더 및 하위 폴더의 모든 파일 해시(SHA-256)와 서명 등 추출
# =====================================================================

# 윈도우 한글 환경(CP949)에서의 터미널 출력 호환성을 위해 기본 인코딩 사용
[Console]::OutputEncoding = [System.Text.Encoding]::Default
$utf8BOM = New-Object System.Text.UTF8Encoding $true

# 1. 실행 환경 설정
$currentPath = (Get-Item -Path ".\" -Verbose).FullName
$reportName  = "SW_Verification_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$reportPath  = Join-Path -Path $currentPath -ChildPath $reportName

# 스크립트 자신 파일명 확인 (검사 제외용)
$scriptName = $MyInvocation.MyCommand.Name
if ([string]::IsNullOrEmpty($scriptName)) { $scriptName = "SWvalidator.ps1" }

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "[SW 유해성 자가 검증 스크립트]" -ForegroundColor Cyan
Write-Host "==========================================`n"

# 2. 폴더 감지 시 자동 압축(ZIP)
$folders = Get-ChildItem -Path $currentPath -Directory | Where-Object { $_.Name -ne "Temp_Zip_Extract" }
if ($folders.Count -gt 0) {
    Write-Host "`n[알림] 폴더가 감지되어 자동으로 zip 형식으로 압축합니다." -ForegroundColor Cyan
    foreach ($folder in $folders) {
        $zipPath = Join-Path $currentPath "$($folder.Name).zip"
        Write-Host " >> 압축 중: $($folder.Name) -> $($folder.Name).zip"
        Compress-Archive -Path $folder.FullName -DestinationPath $zipPath -Force
        Write-Host " >> 완료. (원본 폴더는 그대로 유지됩니다.)" -ForegroundColor Green
    }
}

# 3. 대상 파일 수집 (스크립트 자신과 검증기, 기존 txt 리포트 제외)
$files = Get-ChildItem -Path $currentPath -File | Where-Object { 
    $_.Name -ne $scriptName -and $_.Name -ne "SWverifier.ps1" -and $_.Extension -ne ".txt" 
}

if ($null -eq $files -or $files.Count -eq 0) {
    Write-Host "[경고] 검사할 파일이 없습니다. (스크립트/txt 제외)" -ForegroundColor Yellow
    Read-Host "엔터를 누르면 종료됩니다..."
    exit
}

Write-Host "총 $($files.Count)개의 파일 검사를 시작합니다...`n"

# 3. 엑셀 붙여넣기용 헤더 작성 (Tab 구분자로 엑셀 호환성 극대화)
$reportContent = @()
$reportContent += "파일명`t해시값(SHA-256)`t전자서명(Status)`t서명자(Signer)`t제조사`t제품명`t버전"
$reportContent += "-" * 110

$advancedData = @()

# 4. 파일 처리할 목록 구성 (Zip 내부 파일 추출 기능 포함)
$tempExtractBase = Join-Path $currentPath "Temp_Zip_Extract"
if (-not (Test-Path $tempExtractBase)) { New-Item -ItemType Directory -Path $tempExtractBase | Out-Null }

$processList = @()
foreach ($f in $files) {
    $processList += [PSCustomObject]@{ FileInfo = $f; DisplayName = $f.Name; IsArchive = ($f.Extension -match "\.(zip|7z|rar|tar|gz|bz2)$") }

    if ($f.Extension -match "\.zip$") {
        $zipTempDir = Join-Path $tempExtractBase $f.Name
        try {
            Expand-Archive -Path $f.FullName -DestinationPath $zipTempDir -Force -ErrorAction Stop
            $innerPEs = Get-ChildItem -Path $zipTempDir -Recurse -File | Where-Object { $_.Extension -match "\.(exe|dll|sys|msi)$" }
            foreach ($innerPE in $innerPEs) {
                # zip 안의 상대경로 추출
                $relPath = $innerPE.FullName.Substring($zipTempDir.Length + 1)
                $innerDispName = "[ZIP내부] $($f.Name)\$relPath"
                $processList += [PSCustomObject]@{ FileInfo = $innerPE; DisplayName = $innerDispName; IsArchive = $false }
            }
        } catch {
            Write-Host "[$($f.Name)] 내부 추출 중 오류 발생 (암호 등)" -ForegroundColor Yellow
        }
    }
}

# 5. 파일별 정보 추출 및 진행률 표시
$i = 1
foreach ($item in $processList) {
    $file = $item.FileInfo
    $displayName = $item.DisplayName
    $isArchive = $item.IsArchive

    Write-Progress -Activity "파일 검사 중..." -Status "($i / $($processList.Count)) $displayName" -PercentComplete (($i / $processList.Count) * 100)
    
    # 해시 추출
    try {
        $hash = (Get-FileHash -Path $file.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
        $md5 = (Get-FileHash -Path $file.FullName -Algorithm MD5 -ErrorAction SilentlyContinue).Hash
        $sha1 = (Get-FileHash -Path $file.FullName -Algorithm SHA1 -ErrorAction SilentlyContinue).Hash
    } catch {
        $hash = "해시 추출 실패 (권한/사용중)"
        $md5 = "N/A"
        $sha1 = "N/A"
    }

    # 터미널(콘솔) 화면 출력 확인용
    Write-Host " [$i/$($processList.Count)] $displayName" -ForegroundColor Cyan
    Write-Host "   -> SHA-256 : $hash" -ForegroundColor Gray

    # PE 파일 정보 추출 (버전, 제조사, 제품명)
    try {
        $vi = (Get-Item -Path $file.FullName -ErrorAction SilentlyContinue).VersionInfo
        $company = if ([string]::IsNullOrWhiteSpace($vi.CompanyName)) { "N/A" } else { $vi.CompanyName.Trim() }
        $product = if ([string]::IsNullOrWhiteSpace($vi.ProductName)) { $file.Name } else { $vi.ProductName.Trim() }
        $version = if ([string]::IsNullOrWhiteSpace($vi.FileVersion)) { "N/A" } else { $vi.FileVersion.Trim() }
    } catch {
        $company = "N/A"
        $product = $file.Name
        $version = "N/A"
    }

    # 압축 파일의 경우 전자서명 추출 배제 (N/A 처리)
    if ($isArchive) {
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
    }

    # 결과 조합 (Tab 구분자) - 화면 및 리포트(엑셀)에 서명 및 S/W 정보 덧붙여 출력
    $reportContent += "$displayName`t$hash`t$sigStatusMapped`t$signer`t$company`t$product`t$version"
    
    Write-Host "   -> 전자서명(Status/Signer): $sigStatusMapped / $signer" -ForegroundColor Gray
    Write-Host "   -> S/W 정보(이름/버전/제조사): $product / $version / $company" -ForegroundColor Gray
    
    # 고급 정보 분류 (MD5/SHA1 등 해시 보조 데이터만 하단에 묶음)
    $advLine = "[ADV_INFO] Name: $displayName | MD5: $md5 | SHA1: $sha1"
    $advancedData += $advLine

    $i++
}

# 작업 완료 후 임시 폴더 삭제
if (Test-Path $tempExtractBase) {
    Remove-Item -Path $tempExtractBase -Recurse -Force | Out-Null
}

# 6. 고급 정보 취합 및 1차 리포트 파일 생성
if ($advancedData.Count -gt 0) {
    $reportContent += "`r`n[ADVANCED_DATA_SECTION] (분석기 전용 확장 데이터 - 수정 금지)"
    $reportContent += $advancedData
}

$bodyText = $reportContent -join "`r`n"
[System.IO.File]::WriteAllText($reportPath, $bodyText, $utf8BOM)

# 7. 리포트 자체 무결성 검증을 위한 '봉인(Seal)' 생성 (HMAC-SHA256 적용)
$secretKey = "SW_VALIDATOR_SECURE_KEY_2026!@"
$hmac = [System.Security.Cryptography.HMACSHA256]::new([System.Text.Encoding]::UTF8.GetBytes($secretKey))
$bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($bodyText) # 파일에 쓴 내용 그대로 해시
$hashBytes = $hmac.ComputeHash($bodyBytes)
$reportHash = ([System.BitConverter]::ToString($hashBytes) -replace '-').ToUpper()

# 봉인(Hash) 값을 리포트 맨 끝에 추가 (개행 문자 추가하여 분리)
$sealText = "`r`n========================================================`r`n"
$sealText += "[REPORT_INTEGRITY_SEAL] (보안팀 검증용)`r`n"
$sealText += "아래 해시값은 본문 내용(이 텍스트 위까지)의 HMAC-SHA256 변조방지 변환값입니다.`r`n"
$sealText += "내용이 수정되면 분석기가 위변조를 탐지하여 리포트를 거부합니다.`r`n"
$sealText += "SEAL_HASH: $reportHash`r`n"
$sealText += "========================================================"

[System.IO.File]::AppendAllText($reportPath, $sealText, $utf8BOM)

Write-Host "`n[완료] 검사가 성공적으로 끝났습니다." -ForegroundColor Green
Write-Host "결과 파일이 같은 폴더에 생성되었습니다: $reportName"
Write-Host "생성된 txt 파일의 내용을 엑셀 체크리스트에 복사하여 제출해 주세요.`n"

Read-Host "엔터를 누르면 종료됩니다..."
