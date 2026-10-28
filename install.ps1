# office-automation-kr 자동 설치
# 사용법(PowerShell 에 한 줄 붙여 넣기):
#   02장 기본 설치   : [Net.ServicePointManager]::SecurityProtocol='Tls12'; irm https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/setup-base.ps1 | iex
#   03장 프로그램 설치: [Net.ServicePointManager]::SecurityProtocol='Tls12'; irm https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/setup-program.ps1 | iex
#   한 번에 전부    : [Net.ServicePointManager]::SecurityProtocol='Tls12'; irm https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/install.ps1 | iex
# 순서: 1 Claude Desktop → 2 Node.js → 3 uv·uvx (기본) → 4 이 저장소 → 5 작업 폴더 → 6 기본 약속(전역 CLAUDE.md) (프로그램).
# 이미 있는 것은 건너뛴다. 설치 방법이 막히면 다음 방법으로 넘어간다. 전역 CLAUDE.md 가 이미 있으면 절대 건드리지 않는다.
param(
    [ValidateSet('all', 'base', 'program')]
    [string]$Part = 'all',  # base = 1~3단계(02장), program = 4~6단계(03장)
    [switch]$DryRun,  # 시험용: 1~3 단계 설치를 건너뜀
    [string]$RepoZip = 'https://github.com/dgkim3333-hash/office-automation-kr/archive/refs/heads/main.zip',
    [string]$InstallDir = 'C:\dev',
    [string]$WorkDir = 'C:\사무자동화',
    [string]$ClaudeHome = (Join-Path $env:USERPROFILE '.claude')  # 시험용으로 바꿀 수 있음
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'

# 화면의 단계 번호: 전부 = 1~6, 기본 = 1~3, 프로그램 = 4~6 을 1~3 으로
$StepFrom = 0
$StepTotal = 6
if ($Part -ne 'all') { $StepTotal = 3 }
if ($Part -eq 'program') { $StepFrom = 3 }
function Step($n, $text) { Write-Host ''; Write-Host "[$($n - $StepFrom)/$StepTotal] $text" -ForegroundColor Cyan }
function Done($text) { Write-Host "      $text" -ForegroundColor Green }

function Update-SessionPath {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machine;$user;$env:USERPROFILE\.local\bin"
}

function Add-UserPath($dir) {
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    if (-not $user) { $user = '' }
    if (-not (($user -split ';') -contains $dir)) {
        [Environment]::SetEnvironmentVariable('Path', (($user.TrimEnd(';') + ";$dir").TrimStart(';')), 'User')
    }
}

function Test-Cmd($name) { [bool](Get-Command $name -ErrorAction SilentlyContinue) }

function Get-Arch {
    if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { return 'arm64' }
    return 'x64'
}

function Get-File($from, $to) {
    if (Test-Path -LiteralPath $from) { Copy-Item -LiteralPath $from -Destination $to -Force; return }
    & curl.exe -fsSL -A $UserAgent -o $to $from
    if ($LASTEXITCODE -ne 0) { throw "내려받기 실패: $from" }
}

function Expand-ToTemp($zip) {
    $dir = Join-Path $env:TEMP ('oa_' + [Guid]::NewGuid().ToString('N'))
    Expand-Archive -LiteralPath $zip -DestinationPath $dir -Force
    return $dir
}

# 설치 방법들을 적힌 순서대로 시도한다. 한 방법이 오류를 내거나, 끝났는데도 도구가 안 보이면 다음 방법으로.
function Invoke-FirstWorking($what, $methods, $isDone) {
    foreach ($name in @($methods.Keys)) {
        Write-Host "      시도: $name"
        try { & $methods[$name] } catch { Write-Host "      막힘: $name ($($_.Exception.Message))" -ForegroundColor Yellow }
        Update-SessionPath
        if (& $isDone) { return $name }
    }
    throw "$what 설치를 확인하지 못했습니다."
}

function Invoke-Winget($id) {
    if (-not (Test-Cmd winget)) { throw 'winget 이 없습니다' }
    winget install --id $id -e --silent --accept-source-agreements --accept-package-agreements | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "winget 오류 $LASTEXITCODE" }
}

function Test-ClaudeDesktop {
    try { if (Get-AppxPackage -Name 'Claude' -ErrorAction SilentlyContinue) { return $true } } catch { }
    $keys = @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($key in $keys) {
        if (Get-ItemProperty $key -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq 'Claude' }) { return $true }
    }
    return $false
}

function Get-ClaudeMethods {
    [ordered]@{
        'claude-공식설치파일' = {
            $setup = Join-Path $env:TEMP 'ClaudeSetup.exe'
            Get-File "https://claude.ai/api/desktop/win32/$(Get-Arch)/setup/latest/redirect" $setup
            Start-Process -FilePath $setup -Wait
        }
        'claude-winget' = { Invoke-Winget 'Anthropic.Claude' }
    }
}

function Get-NodeLts {
    $index = (& curl.exe -fsSL -A $UserAgent 'https://nodejs.org/dist/index.json') -join "`n" | ConvertFrom-Json
    return ($index | Where-Object { $_.lts } | Select-Object -First 1).version
}

# 관리자 권한 없이 사용자 폴더에 풀어 쓰는 Node.js
function Install-NodeFromZip($dir, [switch]$NoPath) {
    $version = Get-NodeLts
    $name = "node-$version-win-$(Get-Arch)"
    $zip = Join-Path $env:TEMP "$name.zip"
    Get-File "https://nodejs.org/dist/$version/$name.zip" $zip
    $tmp = Expand-ToTemp $zip
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Copy-Item -Path (Join-Path (Join-Path $tmp $name) '*') -Destination $dir -Recurse -Force
    Remove-Item -LiteralPath $tmp, $zip -Recurse -Force -ErrorAction SilentlyContinue
    if (-not $NoPath) { Add-UserPath $dir }
}

function Get-NodeMethods {
    [ordered]@{
        'node-winget' = { Invoke-Winget 'OpenJS.NodeJS.LTS' }
        'node-공식설치파일' = {
            $version = Get-NodeLts
            $msi = Join-Path $env:TEMP "node-$version-$(Get-Arch).msi"
            Get-File "https://nodejs.org/dist/$version/node-$version-$(Get-Arch).msi" $msi
            $p = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/i `"$msi`" /passive" -Wait -PassThru
            if ($p.ExitCode -ne 0) { throw "msiexec 오류 $($p.ExitCode)" }
        }
        'node-압축본' = { Install-NodeFromZip (Join-Path $env:LOCALAPPDATA 'Programs\nodejs') }
    }
}

# GitHub 배포 파일(uv.exe·uvx.exe·uvw.exe)을 직접 풀어 넣는 uv
function Install-UvFromZip($binDir, [switch]$NoPath) {
    $arch = 'x86_64'
    if ((Get-Arch) -eq 'arm64') { $arch = 'aarch64' }
    $zip = Join-Path $env:TEMP "uv-$arch.zip"
    Get-File "https://github.com/astral-sh/uv/releases/latest/download/uv-$arch-pc-windows-msvc.zip" $zip
    $tmp = Expand-ToTemp $zip
    New-Item -ItemType Directory -Force -Path $binDir | Out-Null
    Copy-Item -Path (Join-Path $tmp 'uv*.exe') -Destination $binDir -Force
    Remove-Item -LiteralPath $tmp, $zip -Recurse -Force -ErrorAction SilentlyContinue
    if (-not $NoPath) { Add-UserPath $binDir }
}

# 업데이트 때 정리: 지난번에 설치한 파일 목록(.installed-files.txt)에 있었는데 새 판에 없는 파일만 지운다.
# 목록에 없는 파일(작업폴더.txt, .venv, 사용자가 넣은 파일)은 건드리지 않는다.
$ManifestName = '.installed-files.txt'
# 목록이 생기기 전(2026-10-28 장 번호 변경 전)에 설치한 PC 용: 이름이 바뀐 옛 장 파일만
$LegacyFiles = @(
    '02_설치.md', '03_첫_실행_10분.md', '04_내_양식_가르치기.md', '05_매일_쓰기.md',
    '06_프로젝트와_예정됨.md', '07_선택_Outlook_자동수집.md', '08_선택_엑셀PO변환.md'
)

function Remove-StaleFiles($repo, $newFiles) {
    $manifest = Join-Path $repo $ManifestName
    $old = $LegacyFiles
    if (Test-Path -LiteralPath $manifest) { $old = @(Get-Content -LiteralPath $manifest -Encoding UTF8) }
    $count = 0
    foreach ($rel in $old) {
        if (-not $rel -or ($newFiles -contains $rel)) { continue }
        if ($rel -match '\.\.' -or [IO.Path]::IsPathRooted($rel)) { continue }  # 이상한 줄은 건드리지 않음
        $path = Join-Path $repo $rel
        if (Test-Path -LiteralPath $path -PathType Leaf) { Remove-Item -LiteralPath $path -Force; $count++ }
    }
    return $count
}

function Get-UvMethods {
    [ordered]@{
        'uv-공식스크립트' = {
            & powershell.exe -NoProfile -ExecutionPolicy ByPass -Command 'irm https://astral.sh/uv/install.ps1 | iex' | Out-Host
            if ($LASTEXITCODE -ne 0) { throw "공식 스크립트 오류 $LASTEXITCODE" }
        }
        'uv-winget' = { Invoke-Winget 'astral-sh.uv' }
        'uv-GitHub압축본' = { Install-UvFromZip (Join-Path $env:USERPROFILE '.local\bin') }
    }
}

# 시험용: 함수만 불러오고 설치는 하지 않는다
if ($env:OA_INSTALL_FUNCTIONS_ONLY -eq '1') { return }

try {
    Write-Host '사무 자동화 설치를 시작합니다. 창을 닫지 마세요.' -ForegroundColor Yellow

    if ($Part -ne 'program') {
        Step 1 'Claude Desktop'
        if (Test-ClaudeDesktop) { Done '이미 설치돼 있습니다.' }
        elseif ($DryRun) { Done '(시험 실행) 건너뜀' }
        else {
            $used = Invoke-FirstWorking 'Claude Desktop' (Get-ClaudeMethods) { Test-ClaudeDesktop }
            Done "설치했습니다. ($used)"
        }

        Step 2 'Node.js'
        Update-SessionPath
        if (Test-Cmd node) { Done "이미 있습니다. $(node --version)" }
        elseif ($DryRun) { Done '(시험 실행) 건너뜀' }
        else {
            $used = Invoke-FirstWorking 'Node.js' (Get-NodeMethods) { Test-Cmd node }
            Done "설치했습니다. $(node --version) ($used)"
        }

        Step 3 'uv · uvx'
        if ((Test-Cmd uv) -and (Test-Cmd uvx)) { Done "이미 있습니다. $(uv --version)" }
        elseif ($DryRun) { Done '(시험 실행) 건너뜀' }
        else {
            $used = Invoke-FirstWorking 'uv' (Get-UvMethods) { (Test-Cmd uv) -and (Test-Cmd uvx) }
            Done "설치했습니다. $(uv --version) ($used)"
        }
        if (-not (Test-Cmd uvx)) { throw 'uvx 를 찾지 못했습니다.' }
    }

    if ($Part -ne 'base') {
        if (-not (Test-Cmd uv)) { throw '먼저 02장 「기본 설치」를 하세요 (uv 가 없습니다).' }
        Step 4 '사무 자동화 프로그램'
        # 다시 실행하면 프로그램 파일만 새 판으로 덮어쓴다. 작업폴더.txt·.venv·내가 넣은 파일은 그대로 둔다.
        $repo = Join-Path $InstallDir 'office-automation-kr-main'
        $hadBefore = Test-Path -LiteralPath (Join-Path $repo 'run.bat')
        New-Item -ItemType Directory -Force -Path $repo | Out-Null
        $zip = Join-Path $InstallDir 'office-automation-kr-main.zip'
        Get-File $RepoZip $zip
        $fresh = Expand-ToTemp $zip
        Remove-Item -LiteralPath $zip -Force
        $freshRoot = Join-Path $fresh 'office-automation-kr-main'
        $newFiles = @(Get-ChildItem -LiteralPath $freshRoot -Recurse -File -Force |
            ForEach-Object { $_.FullName.Substring($freshRoot.Length + 1) })
        $removed = Remove-StaleFiles $repo $newFiles
        Get-ChildItem -LiteralPath $freshRoot -Force | Copy-Item -Destination $repo -Recurse -Force
        # GitHub 압축은 파일 시각을 미국 시간으로 적어 '어제'처럼 보이므로, 받은 파일의 날짜를 설치한 시각으로 맞춘다
        $installedAt = Get-Date
        foreach ($rel in $newFiles) { (Get-Item -LiteralPath (Join-Path $repo $rel) -Force).LastWriteTime = $installedAt }
        [IO.File]::WriteAllLines((Join-Path $repo $ManifestName), [string[]]$newFiles, (New-Object Text.UTF8Encoding $false))
        Remove-Item -LiteralPath $fresh -Recurse -Force -ErrorAction SilentlyContinue
        $cleaned = ''
        if ($removed -gt 0) { $cleaned = ", 옛 파일 $($removed)개 정리" }
        if ($hadBefore) { Done "새 판으로 바꿨습니다. $repo (작업 폴더와 내 파일은 그대로$cleaned)" }
        else { Done "받았습니다. $repo" }

        Step 5 '작업 폴더 (처음 한 번은 1~3분 걸립니다)'
        Push-Location -LiteralPath $repo
        try { & uv run python -m app init $WorkDir; $code = $LASTEXITCODE } finally { Pop-Location }
        if ($code -ne 0) { throw '작업 폴더를 만들지 못했습니다.' }

        Step 6 '기본 약속 (전역 CLAUDE.md)'
        $globalFile = Join-Path $ClaudeHome 'CLAUDE.md'
        $template = Join-Path $repo 'claude\전역_CLAUDE.md'
        if ($DryRun -and -not $PSBoundParameters.ContainsKey('ClaudeHome')) { Done '(시험 실행) 건너뜀' }
        elseif (Test-Path -LiteralPath $globalFile) { Done "이미 있어서 건드리지 않았습니다. $globalFile" }
        elseif (-not (Test-Path -LiteralPath $template)) { Done '약속 파일이 없어 건너뜀' }
        else {
            New-Item -ItemType Directory -Force -Path $ClaudeHome | Out-Null
            Copy-Item -LiteralPath $template -Destination $globalFile
            Done "만들었습니다. $globalFile"
        }
    }

    Write-Host ''
    if ($Part -eq 'base') {
        Write-Host '기본 설치가 끝났습니다.' -ForegroundColor Green
        Write-Host '다음: Claude 열기 → 로그인 → 그다음 03장(프로그램 설치)을 하세요.'
    } else {
        Write-Host '끝났습니다.' -ForegroundColor Green
        Write-Host "다음: Claude 열기 → 로그인 → 「Code」 → 「폴더 열기」 → $repo"
        Write-Host "그다음 04장: $repo\체험.bat 을 두 번 클릭하세요."
        if (-not $DryRun) { Start-Process explorer.exe -ArgumentList "`"$repo`"" }
    }
} catch {
    Write-Host ''
    Write-Host "설치가 멈췄습니다: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host '이 창의 글자를 전부 복사해 두고, 99장 「자동 설치가 멈췄을 때」를 보세요.' -ForegroundColor Red
    throw
}
