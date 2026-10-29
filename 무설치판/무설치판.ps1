# 무설치판: Windows PowerShell 5.1 + 오피스만으로 PO 를 정리한다. 이 파일은 명령을 골라 lib\ 의 함수를 부르는 입구다.
#
# 쓰는 법: 무설치판.ps1 -Command <명령> [-Input <json>] [-Output <json>] [-Culture <문화권>]
#   엔진  시험용 입구. 입력 JSON 의 op 로 엔진 함수를 고르고 결과 JSON 을 -Output 에 쓴다.
#         오류면 {"error": "사유"} 를 쓰고 반환 코드 1.
#   실행  -Job NN [-Work <작업폴더>] [-Auto] [-Output <json>] [-NoNotify] [-Now yyyy-MM-ddTHH:mm:ss]
#         작업 하나를 실행한다(lib\Run.ps1 의 Invoke-RunCommand). -Work 를 빼면 이 스크립트 폴더의 위(묶음 폴더).
#         -Output 을 주면 보고(또는 {"error", "locked"})를 JSON 으로 쓴다. 끝은 늘 [Environment]::Exit(반환 코드):
#         0 정상 · 1 멈춤(시작 못함 · 처리기록을 못 씀) · 2 같은 작업이 실행 중 · 3 그 밖의 오류(작업 없음 등).
#         (시간 제한으로 버려 둔 오피스 runspace 가 남아 있어도 프로세스가 끝나게 Exit 로 끝낸다.)
#         시험용: -NoNotify(알림 끔) · -WordTimeoutSec <초>(워드 시간 제한) · -TestPdfWarningOff 0|1(PDF 안내 설정 확인 결과 대신)
#         -Auto(작업 스케줄러가 부름): 잠겨 있으면 기다렸다 다시 해 보고, 끝내 실패하면(2 · 3) 실행기록에 남기고 실패 알림.
#   준비  [-Bundle <묶음폴더>] [-Work <작업폴더>] [-NoDialog]  준비.bat 의 몸통(lib\Setup.ps1 의 Invoke-Setup). 묶음 = 이 스크립트 폴더의 위,
#         작업 폴더 = 묶음. 반환 코드 0 끝남 · 1 멈춤(또는 일부 못 함).
#   놓친실행  [-Work <작업폴더>] [-Now …] [-Output <json>] [-NoNotify]  시작프로그램 바로가기가 부름(오늘 시각이 지났는데 안 돈 작업).
#   점검  -File <PDF 또는 xlsx> -Rules <규칙.json>  대표 PC 에서 무설치판으로 읽은 결과 표(지금 판 check 와 같은 표)를 콘솔에 찍는다.
#   -TestHooks <json>: 시험용 끼우기(진짜 예약 작업 · 바탕화면 · 알림 · 폴더 고르기 창을 쓰지 않게). 열쇠:
#         schedule ok|fail · desktop · startup_dir · pick [폴더…] · self_test pass|pdf_fail|excel_fail · notice_file · lock_retry [간격, 한도]
# -Culture 를 주면 그 문화권(예: de-DE)으로 스레드를 바꾼 뒤 돈다(문화권에 흔들리지 않는지 시험할 때).
# ($Input 은 PowerShell 자동 변수라 이름을 InputPath · OutputPath 로 두고 -Input · -Output 은 별칭으로 받는다.)
param(
    [Parameter(Mandatory = $true)][string]$Command,
    [Alias('Input')][string]$InputPath,
    [Alias('Output')][string]$OutputPath,
    [string]$Culture,
    [int]$Job = 0,
    [string]$Work,
    [string]$Now,
    [switch]$Auto,
    [switch]$NoNotify,
    [int]$WordTimeoutSec = -1,
    [string]$TestPdfWarningOff,
    [string]$File,
    [string]$Rules,
    [string]$Bundle,
    [switch]$NoDialog,
    [string]$TestHooks
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

. (Join-Path $PSScriptRoot 'lib\Values.ps1')
. (Join-Path $PSScriptRoot 'lib\Rules.ps1')
. (Join-Path $PSScriptRoot 'lib\ExtractText.ps1')
. (Join-Path $PSScriptRoot 'lib\ExtractGrid.ps1')
. (Join-Path $PSScriptRoot 'lib\Verify.ps1')
. (Join-Path $PSScriptRoot 'lib\Markup.ps1')
. (Join-Path $PSScriptRoot 'lib\Company.ps1')
. (Join-Path $PSScriptRoot 'lib\Office.ps1')
. (Join-Path $PSScriptRoot 'lib\Writers.ps1')
. (Join-Path $PSScriptRoot 'lib\Notify.ps1')
. (Join-Path $PSScriptRoot 'lib\Run.ps1')
. (Join-Path $PSScriptRoot 'lib\Setup.ps1')

$script:Utf8 = New-Object Text.UTF8Encoding($false)

function Read-JsonFile([string]$Path) {
    return [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
}

function Write-JsonFile([string]$Path, [string]$Json) {
    [IO.File]::WriteAllText($Path, $Json, $script:Utf8)
}

# 엔진 op: 입력 JSON(op 를 뺀 나머지 키)을 받아 결과 JSON 글자를 돌려준다. 뒤 작업이 op 를 여기에 더한다.
$script:EngineOps = @{
    'extract_text' = {
        param($Payload)
        $rules = ConvertTo-PortableRules $Payload.rules
        ConvertTo-ResultJson (Get-TextExtract $rules ([string]$Payload.text))
    }
    'extract_grid' = {
        param($Payload)
        $rules = ConvertTo-PortableRules $Payload.rules
        ConvertTo-ResultJson (Get-GridExtract $rules $Payload.grid)
    }
    'verify' = {
        # 입력 {rules, result, source}: result 는 위 두 op 의 결과 객체, source 는 {kind: "text", text} 또는 {kind: "grid", grid}.
        param($Payload)
        $rules = ConvertTo-PortableRules $Payload.rules
        ConvertTo-ChecksJson (Get-Checks $Payload.result $Payload.source $rules)
    }
    'markup' = {
        # 입력 {action, ...}: 마크업표 읽기 · 값 · 없는 파트 사유 · 검산 · 지문 · 쓴 값(Markup.ps1 의 Invoke-MarkupOp 주석 참고).
        param($Payload)
        Invoke-MarkupOp $Payload
    }
    'company' = {
        # 입력 {action, ...}: 회사 양식 줄 · 파일 이름 채우기 · 이름 충돌(Company.ps1 의 Invoke-CompanyOp 주석 참고).
        param($Payload)
        Invoke-CompanyOp $Payload
    }
    'office' = {
        # 입력 {action, ...}: 오피스로 읽기 · 정규화 · 내 프로세스 고르기 · PDF 안내 설정(Office.ps1 의 Invoke-OfficeOp 주석 참고).
        param($Payload)
        Invoke-OfficeOp $Payload
    }
    'office_write' = {
        # 입력 {action, ...}: 엑셀로 회사 양식 · 장부 · 검산 파일 쓰기 · 다시 읽기 · 잠김 확인(Writers.ps1 의 Invoke-WriterOp 주석 참고).
        param($Payload)
        Invoke-WriterOp $Payload
    }
    'notify' = {
        # 입력 {action, ...}: 알림 글자 · 알림 XML(Notify.ps1 의 Invoke-NotifyOp 주석 참고). 알림은 띄우지 않는다.
        param($Payload)
        Invoke-NotifyOp $Payload
    }
    'setup' = {
        # 입력 {action, ...}: 예약(가짜 등록) · 바로가기 · 자체 시험(Setup.ps1 의 Invoke-SetupOp 주석 참고).
        param($Payload)
        Invoke-SetupOp $Payload
    }
}

function Get-HookValue([string]$Name) {
    if ($null -eq $script:Hooks) { return $null }
    $p = $script:Hooks.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Get-DefaultBundle {
    if ($Bundle) { return $Bundle }
    return (Split-Path -Parent $PSScriptRoot)
}

function Get-NowArgument {
    if ($Now) { return [datetime]::ParseExact($Now, 'yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) }
    return [datetime]::MinValue
}

# 시험용 끼우기(-TestHooks). 운영(준비.bat · 바로가기 · 작업 스케줄러)에서는 쓰지 않는다.
$script:Hooks = $null
if ($TestHooks) {
    $script:Hooks = Read-JsonFile $TestHooks
    if (Get-HookValue 'notice_file') { $script:NoticeFile = [string](Get-HookValue 'notice_file') }
    $retry = Get-HookValue 'lock_retry'
    if ($null -ne $retry) { $script:LockRetrySeconds = [int]$retry[0]; $script:LockRetryLimitSeconds = [int]$retry[1] }
}
if ($TestPdfWarningOff) { $script:RunPdfWarningOff = ($TestPdfWarningOff -eq '1') }

# 명령: 엔진 · 실행 · 놓친실행 · 준비 · 점검.
$script:Commands = @{
    '엔진' = {
        $payload = Read-JsonFile $InputPath
        $op = [string]$payload.op
        if (-not $script:EngineOps.ContainsKey($op)) { throw "엔진 op '$op' 는 없습니다." }
        Write-JsonFile $OutputPath (& $script:EngineOps[$op] $payload)
    }
    '실행' = {
        if ($Job -le 0) { throw '작업 번호(-Job NN)를 주세요.' }
        $workFolder = $Work
        if ([string]::IsNullOrEmpty($workFolder)) { $workFolder = Split-Path -Parent $PSScriptRoot }
        $when = Get-NowArgument
        if ($WordTimeoutSec -ge 0) { $script:RunWordTimeoutSec = $WordTimeoutSec }
        $label = if ($Auto) { '자동 실행' } else { '지금 실행' }
        $code = Invoke-RunCommand $workFolder $Job $when $label ([bool]$NoNotify) $OutputPath ([bool]$Auto)
        [Console]::Out.Flush()
        [Environment]::Exit($code)
    }
    '놓친실행' = {
        $workFolder = $Work
        if ([string]::IsNullOrEmpty($workFolder)) { $workFolder = Split-Path -Parent $PSScriptRoot }
        $code = Invoke-CatchupCommand $workFolder (Get-NowArgument) ([bool]$NoNotify) $OutputPath
        [Console]::Out.Flush()
        [Environment]::Exit($code)
    }
    '준비' = {
        $bundleFolder = Get-DefaultBundle
        $workFolder = $Work
        if ([string]::IsNullOrEmpty($workFolder)) { $workFolder = $bundleFolder }
        $options = @{ BundleDir = $bundleFolder; Work = $workFolder; NoDialog = [bool]$NoDialog }
        if ($null -ne $script:Hooks) {
            # 시험: 진짜 작업 스케줄러 · 바탕화면 · 시작프로그램 · 폴더 고르기 창을 쓰지 않는다(빠진 열쇠가 있으면 멈춤)
            foreach ($key in @('desktop', 'startup_dir')) { if (-not (Get-HookValue $key)) { throw "시험용 끼우기에 $key 가 없습니다." } }
            $options.Desktop = [string](Get-HookValue 'desktop')
            $options.StartupDir = [string](Get-HookValue 'startup_dir')
            if ([string](Get-HookValue 'schedule') -ceq 'ok') { $options.RegisterTask = { param($Text, $TaskPath, $TaskName) $true } }
            else { $options.RegisterTask = { param($Text, $TaskPath, $TaskName) throw '시험: 작업 스케줄러 막힘' } }
            $queue = New-Object 'System.Collections.Generic.Queue[string]'
            foreach ($folder in @(Get-HookValue 'pick')) { if ($null -ne $folder) { $queue.Enqueue([string]$folder) } }
            $options.FolderPicker = { param($Prompt) if ($queue.Count -gt 0) { $queue.Dequeue() } else { '' } }.GetNewClosure()
            $fixed = [string](Get-HookValue 'self_test')
            if ($fixed) {
                $result = [ordered]@{ ok = ($fixed -eq 'pass'); pdf_ok = ($fixed -ne 'pdf_fail'); excel_ok = ($fixed -ne 'excel_fail'); message = ''; detail = '' }
                if ($fixed -eq 'pdf_fail') { $result.message = $script:PdfDifferent }
                if ($fixed -eq 'excel_fail') { $result.message = $script:ExcelDifferent }
                $options.SelfTest = { param($BundleDir) $result }.GetNewClosure()
            }
        }
        $code = Invoke-Setup @options
        [Console]::Out.Flush()
        [Environment]::Exit($code)
    }
    '점검' = {
        if (-not $File -or -not $Rules) { throw '점검할 파일(-File)과 규칙.json(-Rules)을 주세요.' }
        $code = Invoke-CheckCommand $File $Rules
        [Console]::Out.Flush()
        [Environment]::Exit($code)
    }
}

try {
    if ($Culture) {
        $info = New-Object Globalization.CultureInfo($Culture)
        [Threading.Thread]::CurrentThread.CurrentCulture = $info
        [Threading.Thread]::CurrentThread.CurrentUICulture = $info
    }
    if (-not $script:Commands.ContainsKey($Command)) { throw "명령 '$Command' 는 없습니다." }
    & $script:Commands[$Command]
    exit 0
} catch {
    $message = $_.Exception.Message
    if ($OutputPath) { Write-JsonFile $OutputPath (ConvertTo-Json -InputObject @{ error = $message }) }
    [Console]::Error.WriteLine($message)
    exit 1
}
