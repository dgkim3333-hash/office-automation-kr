# 작업 실행(친구 PC 에서 매일 · 「지금 실행」): 입력 폴더의 PDF · 엑셀을 읽어 회사 양식 · 누적장부를 만든다.
# 지금 판 app\job_run.py 의 run_job 에서 승인(승인대기) · 고객마스터를 뺀 것과 같게 움직인다. 순서 · 문구 · 기록 파일 모양이 같다.
#   - 입력 폴더의 원본은 읽기만 한다(옮기기 · 이름 바꾸기 · 지우기 없음). 처리한 파일은 처리기록.csv 에 내용 지문(sha256)으로 남긴다.
#   - 처리기록은 줄을 더하기만 하며, 한 지문의 지금 상태는 그 지문의 마지막 줄이다.
# 점-소싱 순서: Values → Rules → ExtractText → ExtractGrid → Verify → Markup → Company → Office → Writers → Notify → Run.

$script:JobsDir = '작업'
$script:SettingsFile = '작업설정.json'
$script:RulesFile = '규칙.json'
$script:RecordFile = '처리기록.csv'
# 지금 판 RECORD_COLUMNS 그대로(지금 판이 이 파일을 읽을 수 있게). 양식지문(승인용)은 무설치판에서 늘 빈칸.
$script:RecordColumns = @('지문', '파일', '문서종류', '번호', '상태', '규칙지문', '날짜', '양식지문', '마크업지문')
$script:LogFile = '처리로그.csv'
$script:LogColumns = @('일시', '파일', '결과', '사유')
$script:LockFile = '실행중.lock'
$script:RunRecordFile = '실행기록.txt'
$script:ReviewDir = '확인필요'
$script:ChecksDir = '검산'
$script:LedgerName = '누적장부_{0}.xlsx'
$script:InputSuffixes = @('.pdf', '.xlsx')            # .xls 는 판별하지 못하므로 건너뛴다
$script:UnmatchedSuffixes = @('.pdf', '.xlsx', '.xls')
$script:SettleSeconds = 60                            # 이보다 최근에 바뀐 파일은 아직 저장 중일 수 있어 다음 실행으로 미룬다
$script:UnmatchedSeconds = 24 * 3600
$script:OutputModes = @('회사양식', '누적장부', '둘다')
$script:LedgerModes = @('누적장부', '둘다')
$script:CompanyModes = @('회사양식', '둘다')
$script:CountKeys = @('성공', '중복', '확인필요', '보류')
$script:NewsKeys = @('성공', '중복', '확인필요')
$script:SyncFolderNames = @('onedrive', 'dropbox', 'google drive', 'icloud')
$script:TitleForbidden = New-Object Text.RegularExpressions.Regex('[\\/:*?"<>|]')
$script:TitleMax = 30

# 지금 판 job_run.py 의 글자 그대로
$script:RetryNote = '규칙.yaml 을 고치면 다음 실행 때 다시 시도합니다. 규칙을 그대로 두어도 다음 날 한 번 더 시도합니다.'
$script:HeldReason = '다른 프로그램(PDF 뷰어 · 엑셀 등)에서 열려 있음. 닫으면 다음 실행 때 처리합니다'
$script:SettlingReason = '방금 바뀐 파일(1분 안). 저장이 끝난 뒤 다음 실행 때 처리합니다'
$script:LedgerLockedReason = '누적장부가 엑셀에서 열려 있음. 엑셀을 닫으면 다음 실행 때 처리합니다'
$script:Interrupted = '처리기록에 성공 기록이 없음. 지난 실행이 중간에 끊겼거나 되돌리기가 끝나지 못했을 수 있음'
$script:KeepHint = '입력 폴더에서 원본을 치우세요(그대로 두면 날마다 다시 확인필요가 됩니다).'
$script:RemakeHint = '그러면 내일 실행 때(또는 규칙.yaml 을 고친 뒤) 새로 만듭니다(처리기록은 고칠 필요 없음).'
$script:FailPrefix = '중요: {0} 작업을 실행하지 못했습니다(아무 파일도 처리하지 않음).'
# 설계 6-4: 워드의 「PDF 를 편집 가능한 문서로 변환합니다」 안내가 켜져 있으면 PDF 를 열지 않는다(설정은 바꾸지 않음)
$script:PdfWarningAlert = "워드의 PDF 변환 안내가 켜져 있어 PDF {0}개를 처리하지 않았습니다(입력 폴더에 그대로 두었고, 안내를 끈 뒤 다음 실행 때 처리합니다). 워드에서 PDF 를 한 번 열고 '이 메시지를 다시 표시 안 함'에 체크하세요."
# 준비(작업 10)의 자체 시험에서 PDF 가 정답과 달랐으면 작업 폴더의 자체시험결과.json 에 pdf_ok = false 가 남는다. 그동안 PDF 는 위 안내와 같은 길로 둔다
$script:SelfTestResultFile = '자체시험결과.json'
$script:SelfTestPdfAlert = '준비 때 자체 시험에서 PDF 를 통과하지 못해 PDF {0}개를 처리하지 않았습니다(입력 폴더에 그대로 두었습니다). 준비.bat 을 다시 두 번 클릭해도 그대로면 대표에게 알려 주세요. 자체 시험: '
# excel_ok = false 도 같은 길: 엑셀 PO 를 기록 없이 둔다(엑셀이 다르게 읽는 PC 에서 결과 · 장부에 들어가지 않게)
$script:SelfTestExcelAlert = '준비 때 자체 시험에서 엑셀을 통과하지 못해 엑셀 PO {0}개를 처리하지 않았습니다(입력 폴더에 그대로 두었습니다). 준비.bat 을 다시 두 번 클릭해도 그대로면 대표에게 알려 주세요. 자체 시험: '
# 자동 실행 · 놓친 실행이 잠긴 작업을 만나면 이 간격(초)으로 이 시간(초)까지 다시 해 본다(지금 판 schedule.retry_locked). 시험에서만 줄인다
$script:LockRetrySeconds = 30
$script:LockRetryLimitSeconds = 15 * 60
$script:AutoFixHint = '대표에게 알려 주세요(작업 설정이 깨졌으면 고친 묶음을 받아 준비.bat 을 다시 두 번 클릭합니다).'

# 시험용으로만 바꾸는 값(-WordTimeoutSec · -TestPdfWarningOff). $null 이면 보통 값.
$script:RunWordTimeoutSec = $null
$script:RunPdfWarningOff = $null

$script:Run = $null      # 지금 실행 하나의 상태(Invoke-PortableJob 마다 새로)
$script:RunUtf8 = New-Object Text.UTF8Encoding($false)
$script:RunUtf8Strict = New-Object Text.UTF8Encoding($false, $true)
$script:Bom = [byte[]](0xEF, 0xBB, 0xBF)

# ---- 작은 도우미 ---------------------------------------------------------------------------------

function New-StopRun([string]$Message) {
    # 처리기록을 쓰지 못해 실행을 멈춰야 할 때의 오류(지금 판 _StopRun). 한 건 오류 처리가 삼키지 않게 표시한다.
    $failure = New-Object System.Exception $Message
    $failure.Data['stop_run'] = $true
    return $failure
}

function Test-StopRun($ErrorRecord) {
    $e = $ErrorRecord.Exception
    return ($null -ne $e -and $null -ne $e.Data -and $e.Data.Contains('stop_run'))
}

function Get-TextSha256([byte[]]$Bytes) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash($Bytes) } finally { $sha.Dispose() }
    return ([BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant()
}

function Open-SharedRead([string]$Path) {
    # 다른 프로그램(엑셀 등)이 연 파일도 읽는다(파이썬 open 'rb' 처럼 읽기 · 쓰기 공유).
    return New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read,
                                    ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
}

function Get-FileFingerprint([string]$Path) {
    # 파일 내용의 sha256(소문자 16진수 64자). 이름이 바뀌어도 같은 파일이면 같은 값(지금 판 file_fingerprint).
    $stream = Open-SharedRead $Path
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash($stream) } finally { $sha.Dispose(); $stream.Dispose() }
    return ([BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant()
}

function Read-SharedBytes([string]$Path) {
    $stream = Open-SharedRead $Path
    try {
        $buffer = New-Object IO.MemoryStream
        $stream.CopyTo($buffer)
        return ,$buffer.ToArray()
    } finally { $stream.Dispose() }
}

function ConvertTo-WindowsText([string]$Text) {
    # 파이썬 텍스트 모드 쓰기(Windows)처럼 \n 을 모두 \r\n 으로.
    return $Text.Replace("`n", "`r`n")
}

function Add-TextFile([string]$Path, [string]$Text) {
    # 파이썬 open(path, 'a', encoding='utf-8') + write(text): BOM 없음, 줄바꿈 \r\n.
    $bytes = $script:RunUtf8.GetBytes((ConvertTo-WindowsText $Text))
    $stream = New-Object IO.FileStream($Path, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
}

function Get-FolderTitle([string]$Title) {
    # 지금 판 jobs._folder_title: 파일 이름에 못 쓰는 글자는 _, 앞뒤 공백 · 끝 점 빼고 30자까지.
    $text = $script:TitleForbidden.Replace($Title, '_').Trim().TrimEnd('.').Trim()
    if ($text.Length -gt $script:TitleMax) { $text = $text.Substring(0, $script:TitleMax) }
    return $text.TrimEnd().TrimEnd('.')
}

function Test-SameDir([string]$A, [string]$B) {
    $a1 = [IO.Path]::GetFullPath($A).TrimEnd('\')
    $b1 = [IO.Path]::GetFullPath($B).TrimEnd('\')
    return [string]::Equals($a1, $b1, [StringComparison]::OrdinalIgnoreCase)
}

function Assert-LocalFolder([string]$Path) {
    # 지금 판 local_guard.check_work_dir: 네트워크 위치 · 동기화 폴더(OneDrive 등) · 네트워크 드라이브는 쓰지 않는다.
    if ($Path.StartsWith('\\')) {
        throw "작업 폴더가 네트워크 위치입니다: $Path`n내 PC 의 로컬 디스크(예: C:\사무자동화)를 쓰세요."
    }
    foreach ($part in $Path.Split([char[]]@('\', '/'), [StringSplitOptions]::RemoveEmptyEntries)) {
        $lowered = $part.ToLowerInvariant()
        foreach ($name in $script:SyncFolderNames) {
            if ($lowered.StartsWith($name, [StringComparison]::Ordinal)) {
                throw ("작업 폴더가 동기화 폴더($part) 안에 있습니다: $Path`n" +
                       '장부와 PDF 가 클라우드로 자동 업로드되므로 쓸 수 없습니다. 예: C:\사무자동화 처럼 로컬 디스크를 쓰세요.')
            }
        }
    }
    $root = [IO.Path]::GetPathRoot($Path)
    if ($root -match '^[A-Za-z]:\\?$') {
        $drive = New-Object IO.DriveInfo($root.Substring(0, 1))
        if ($drive.DriveType -eq [IO.DriveType]::Network) {
            throw "작업 폴더가 네트워크 드라이브($($root.Substring(0, 2)))에 있습니다. 내 PC 의 로컬 디스크를 쓰세요."
        }
    }
}

# ---- CSV (지금 판 csv 모듈과 같은 모양: 새 파일만 BOM, 줄바꿈 \r\n, 필요한 칸만 따옴표) -----------------------------

function ConvertTo-CsvField($Value) {
    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    if ($text.IndexOfAny([char[]]@(',', '"', "`r", "`n")) -ge 0) { return '"' + $text.Replace('"', '""') + '"' }
    return $text
}

function Add-CsvRow([string]$Path, [string[]]$Columns, [object[]]$Values) {
    # 지금 판 _append_record · log.append_log: 새(또는 빈) 파일이면 utf-8-sig 로 머리글부터, 아니면 BOM 없이 줄만 덧붙인다.
    $new = -not [IO.File]::Exists($Path) -or (New-Object IO.FileInfo $Path).Length -eq 0
    $text = New-Object Text.StringBuilder
    if ($new) { [void]$text.Append([string]::Join(',', @($Columns | ForEach-Object { ConvertTo-CsvField $_ })) + "`r`n") }
    [void]$text.Append([string]::Join(',', @($Values | ForEach-Object { ConvertTo-CsvField $_ })) + "`r`n")
    $bytes = $script:RunUtf8.GetBytes($text.ToString())
    $stream = New-Object IO.FileStream($Path, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try {
        if ($new) { $stream.Write($script:Bom, 0, $script:Bom.Length) }
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
}

function Read-CsvRecords([string]$Path, [string[]]$Columns) {
    # 지금 판 load_records(csv.DictReader): 머리글 이름으로 읽고, 옛 머리글(열이 적음) 아래 넘친 값은 빠진 열에 순서대로 맞춘다.
    # 깨진 파일(글자 · 따옴표)이면 오류를 던진다. 없으면 빈 목록.
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not [IO.File]::Exists($Path)) { return ,$rows }
    $bytes = Read-SharedBytes $Path
    $text = $script:RunUtf8Strict.GetString($bytes)
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser(New-Object IO.StringReader $text)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        if ($parser.EndOfData) { return ,$rows }
        $header = $parser.ReadFields()
        $missing = @($Columns | Where-Object { $header -notcontains $_ })
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields()
            if ($null -eq $fields) { continue }
            $row = @{}
            for ($i = 0; $i -lt $header.Count; $i++) { if ($i -lt $fields.Count) { $row[$header[$i]] = $fields[$i] } }
            for ($i = $header.Count; $i -lt $fields.Count -and ($i - $header.Count) -lt $missing.Count; $i++) {
                $row[$missing[$i - $header.Count]] = $fields[$i]
            }
            $rows.Add($row)
        }
    } finally { $parser.Dispose() }
    return ,$rows
}

function Get-RecordValue($Row, [string]$Name) {
    if ($null -eq $Row -or -not $Row.ContainsKey($Name) -or $null -eq $Row[$Name]) { return '' }
    return [string]$Row[$Name]
}

# ---- 작업 찾기 · 설정 · 잠금 -----------------------------------------------------------------------

function Get-PortableJobFolders([string]$Work) {
    # 작업\NN_제목\ 폴더들(번호 순). 지운작업 같은 다른 폴더는 뺀다.
    $root = Join-Path $Work $script:JobsDir
    if (-not [IO.Directory]::Exists($root)) { return @() }
    return @([IO.Directory]::GetDirectories($root) | Where-Object { [IO.Path]::GetFileName($_) -match '^[0-9]{2,}_' } |
             Sort-Object { [IO.Path]::GetFileName($_).ToLowerInvariant() })
}

function Find-PortableJob([string]$Work, [int]$Number) {
    $prefix = '{0:D2}_' -f $Number
    $found = @(Get-PortableJobFolders $Work | Where-Object { ([IO.Path]::GetFileName($_)).StartsWith($prefix, [StringComparison]::Ordinal) })
    if ($found.Count -eq 0) { throw ("작업 {0} 이 없습니다: {1}" -f $Number, (Join-Path $Work $script:JobsDir)) }
    if ($found.Count -gt 1) { throw ("작업 {0} 폴더가 여럿입니다: {1}" -f $Number, [string]::Join(', ', $found)) }
    return $found[0]
}

function Read-JobSettings([string]$Folder) {
    # 작업설정.json → { name; folder; settings(원래 객체); title; input_dir; output_dir; output_mode }. 읽지 못하면 오류.
    $path = Join-Path $Folder $script:SettingsFile
    if (-not [IO.File]::Exists($path)) { throw "작업 설정 파일이 없습니다: $path" }
    try {
        $settings = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
    } catch {
        throw "작업 설정 파일을 읽지 못했습니다($path): $($_.Exception.Message)"
    }
    foreach ($key in @('title', 'input_dir', 'output_dir', 'output_mode')) {
        if ($null -eq $settings -or $null -eq $settings.PSObject.Properties[$key] -or [string]::IsNullOrWhiteSpace([string]$settings.$key)) {
            throw "작업 설정 파일에 '$key' 가 없습니다: $path"
        }
    }
    $mode = [string]$settings.output_mode
    if ($script:OutputModes -notcontains $mode) {
        throw ("작업 설정의 출력 방식 '{0}' 을 쓸 수 없습니다. 고를 수 있는 것: {1}" -f $mode, [string]::Join(', ', $script:OutputModes))
    }
    return @{ name = [IO.Path]::GetFileName($Folder); folder = $Folder; settings = $settings; title = [string]$settings.title
              input_dir = [string]$settings.input_dir; output_dir = [string]$settings.output_dir; output_mode = $mode }
}

function Save-JobLastRun($Job, [string]$Today) {
    # 작업설정.json 의 last_run 만 바꿔 저장한다(모르는 키는 그대로).
    Save-JobSettingsValues $Job ([ordered]@{ last_run = $Today })
}

function Save-JobSettingsValues($Job, $Values) {
    # 작업설정.json 의 주어진 키만 바꿔 저장한다(모르는 키는 그대로). 임시 파일에 쓴 뒤 바꾼다.
    $settings = $Job.settings
    foreach ($key in $Values.Keys) {
        if ($null -eq $settings.PSObject.Properties[$key]) { $settings | Add-Member -NotePropertyName $key -NotePropertyValue $Values[$key] }
        else { $settings.$key = $Values[$key] }
    }
    $path = Join-Path $Job.folder $script:SettingsFile
    $tmp = $path + '.tmp'
    [IO.File]::WriteAllText($tmp, (ConvertTo-Json -InputObject $settings -Depth 20), $script:RunUtf8)
    [IO.File]::Replace($tmp, $path, [NullString]::Value)
}

function Enter-JobLock([string]$Folder, [string]$Name) {
    # 작업 폴더의 실행중.lock 첫 바이트를 잠근다(지금 판 job_lock 의 msvcrt.locking 과 같은 1바이트 잠금 → 두 판이 서로 막는다).
    # 이미 누가 잡고 있으면 오류(.Data['job_locked'] = $true).
    $path = Join-Path $Folder $script:LockFile
    if (-not [IO.Directory]::Exists($Folder)) { throw "작업이 지워졌습니다($Name)." }
    $stream = New-Object IO.FileStream($path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite,
                                       ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {
        $stream.Lock(0, 1)
    } catch {
        $stream.Dispose()
        $failure = New-Object System.InvalidOperationException ("이 작업이 이미 실행 중입니다({0}). 끝난 뒤 다시 실행하세요." -f $Name)
        $failure.Data['job_locked'] = $true
        throw $failure
    }
    return $stream
}

function Exit-JobLock($Stream) {
    if ($null -eq $Stream) { return }
    try { $Stream.Unlock(0, 1) } catch { }   # 닫으면 잠금도 풀린다
    $Stream.Dispose()
}

# ---- 보고 --------------------------------------------------------------------------------------

function New-RunReport {
    $counts = [ordered]@{}
    foreach ($key in $script:CountKeys) { $counts[$key] = 0 }
    return [ordered]@{
        counts = $counts; unmatched_old = 0; unmatched_xls = 0; message = ''
        output_files = (New-Object System.Collections.Generic.List[string])
        ledger_locked = $false        # 누적장부가 엑셀에서 열려 있어 보류한 파일이 있음
        stopped = $false              # 시작하지 못했거나 처리기록을 쓰지 못해 중간에 멈춤(message 맨 앞에 「중요:」)
        alerts = (New-Object System.Collections.Generic.List[string])        # 알림에도 그대로 보일 경고
        office_notes = (New-Object System.Collections.Generic.List[string])  # 실행기록에만 남기는 오피스 정리 기록
        has_news = $false
    }
}

function Test-ReportNews($Report) {
    # 지금 판 has_news 에 경고(alerts)를 더했다: PDF 변환 안내 · 장부 .bak 멈춤은 세는 건이 없어도 사람이 알아야 한다.
    foreach ($key in $script:NewsKeys) { if ($Report.counts[$key] -gt 0) { return $true } }
    return ($Report.unmatched_old -gt 0 -or $Report.ledger_locked -or $Report.stopped -or $Report.alerts.Count -gt 0)
}

function New-StoppedReport([string]$Title, [string]$Reason) {
    # 시작 전에 멈춤. 매일 자동 실행이 조용히 건너뛰지 않도록 알림 대상(stopped)으로 돌려준다.
    $report = New-RunReport
    $report.message = ($script:FailPrefix -f $Title) + "`n" + $Reason
    $report.stopped = $true
    $report.has_news = $true
    return $report
}

# ---- 실행 하나 ----------------------------------------------------------------------------------

function Invoke-PortableJob([string]$Work, [int]$Number, [datetime]$Now) {
    # 작업 하나를 실행하고 보고 객체를 돌려준다(아래 New-RunReport 모양). 같은 작업이 이미 실행 중이면 오류(.Data['job_locked']).
    # $Now 를 주지 않으면(MinValue) 지금 시각. 초 아래는 버린다.
    if ($Now -eq [datetime]::MinValue) { $Now = [datetime]::Now }
    $Now = $Now.AddTicks(- ($Now.Ticks % [TimeSpan]::TicksPerSecond))
    $folder = Find-PortableJob $Work $Number
    $name = [IO.Path]::GetFileName($folder)
    $lock = Enter-JobLock $folder $name
    try {
        try {
            $job = Read-JobSettings $folder
        } catch {
            return New-StoppedReport $name $_.Exception.Message
        }
        return Invoke-JobRun $job $Work $Now
    } finally {
        Exit-JobLock $lock
        $script:Run = $null
    }
}

function Invoke-JobRun($Job, [string]$Work, [datetime]$Now) {
    $script:Run = @{
        job = $Job; work = $Work; now = $Now; today = $Now.ToString('yyyy-MM-dd', $script:Invariant)
        source = $Job.input_dir; report = (New-RunReport); warnings = (New-Object System.Collections.Generic.List[string])
        stop_reason = ''; ledger_locked = $false; ledger_held = 0; ledger_stop = $null
        markup_used = $false; markup_bytes = $null; markup_read_error = ''; markup_table = $null; markup_fp = ''
        pdf_skipped = 0; pdf_ok = $true; pdf_self_test = $null; excel_skipped = 0; excel_self_test = $null
    }
    $R = $script:Run
    $script:OfficeHandovers.Clear()
    $stop = Get-PreflightProblem $Job
    if ($null -ne $stop) { return New-StoppedReport $Job.title $stop }

    # 규칙과 규칙지문을 한 번 읽은 같은 바이트에서(실행 중 규칙.json 이 바뀌어도 섞이지 않게)
    $rulesPath = Join-Path $Job.folder $script:RulesFile
    try {
        if (-not [IO.File]::Exists($rulesPath)) { throw "규칙 파일이 없습니다: $rulesPath" }
        $raw = Read-SharedBytes $rulesPath
        $R.rules = ConvertTo-PortableRules ($script:RunUtf8Strict.GetString($raw).TrimStart([char]0xFEFF) | ConvertFrom-Json)
        if ($null -eq $R.rules -or $null -eq $R.rules.PSObject.Properties['doc_types']) { throw '문서종류(doc_types)가 없습니다' }
        $R.rules_fp = if ($null -ne $R.rules.PSObject.Properties['fingerprint'] -and $R.rules.fingerprint) { [string]$R.rules.fingerprint }
                      else { (Get-TextSha256 $raw).Substring(0, 16) }
    } catch {
        return New-StoppedReport $Job.title ("규칙.json 오류: {0}`n대표에게 알려 고친 규칙.json(새 묶음)을 받으세요." -f $_.Exception.Message)
    }
    $form = $null
    if ($null -ne $R.rules.PSObject.Properties['company_form']) { $form = $R.rules.company_form }
    $R.form = $form
    if ($script:CompanyModes -contains $Job.output_mode -and $null -eq $form) {
        return New-StoppedReport $Job.title ("출력 방식이 '{0}' 인데 규칙.json 에 회사양식이 없습니다. 대표에게 알려 회사양식을 적은 규칙.json 을 받거나, 출력 방식을 '누적장부' 로 바꾸세요." -f $Job.output_mode)
    }
    $R.markup_used = ($script:CompanyModes -contains $Job.output_mode) -and (Test-FormUsesMarkup $form)
    if ($R.markup_used) { Read-RunMarkupBytes }
    try {
        $R.state = @{}
        foreach ($row in (Read-CsvRecords (Join-Path $Job.folder $script:RecordFile) $script:RecordColumns)) {
            $R.state[(Get-RecordValue $row '지문')] = $row
        }
    } catch {
        $path = Join-Path $Job.folder $script:RecordFile
        return New-StoppedReport $Job.title ("{0} 를 읽지 못했습니다: {1}`n이 파일이 열려 있으면 닫고, 엑셀 등으로 고쳐 저장했다면 원래대로 되돌린 뒤 다시 실행하세요: {2}" -f $script:RecordFile, $_.Exception.Message, $path)
    }
    $R.success_keys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($row in $R.state.Values) {
        if ((Get-RecordValue $row '상태') -eq '성공') { [void]$R.success_keys.Add((Get-RecordValue $row '문서종류') + "`0" + (Get-RecordValue $row '번호')) }
    }
    $R.pdf_self_test = Get-SelfTestProblem $Work 'pdf_ok'
    $R.excel_self_test = Get-SelfTestProblem $Work 'excel_ok'
    if ($null -ne $R.pdf_self_test) { $R.pdf_ok = $false }
    elseif ($null -ne $script:RunPdfWarningOff) { $R.pdf_ok = [bool]$script:RunPdfWarningOff }
    else { $R.pdf_ok = Test-PdfWarningOff }

    $candidates = @()
    try {
        $candidates = Get-InputDocuments $R.source
    } catch {
        $R.warnings.Add(("경고: 입력 폴더를 읽지 못함 ({0}): {1}" -f $R.source, $_.Exception.Message))
    }
    try {
        foreach ($path in $candidates) { Invoke-OneDocument $path }
    } catch {
        if (-not (Test-StopRun $_)) { throw }
        $R.stop_reason = $_.Exception.Message
        $R.report.stopped = $true
    }
    if ($R.pdf_skipped -gt 0) {
        if ($null -ne $R.pdf_self_test) { $R.report.alerts.Add(($script:SelfTestPdfAlert -f $R.pdf_skipped) + $R.pdf_self_test) }
        else { $R.report.alerts.Add(($script:PdfWarningAlert -f $R.pdf_skipped)) }
    }
    if ($R.excel_skipped -gt 0) { $R.report.alerts.Add(($script:SelfTestExcelAlert -f $R.excel_skipped) + $R.excel_self_test) }

    try {
        Save-JobLastRun $Job $R.today
    } catch {
        $R.warnings.Add(("경고: 작업설정.json 에 마지막 실행 날짜를 쓰지 못함 ({0})" -f $_.Exception.Message))
    }
    try {
        $old = Get-UnmatchedFiles $Job $Work $Now
        $R.report.unmatched_old = $old.Count
        $R.report.unmatched_xls = @($old | Where-Object { [IO.Path]::GetExtension($_).ToLowerInvariant() -eq '.xls' }).Count
    } catch {
        $R.warnings.Add(("경고: 입력 폴더에 남은 파일을 세지 못함 ({0})" -f $_.Exception.Message))
    }
    $R.report.ledger_locked = $R.ledger_locked
    # 숨은 오피스에 친구가 연 문서가 있어 끝내지 않고 넘겨 준 일(Office.ps1 의 Exit-OfficeApp)은 실행기록에 남긴다
    foreach ($note in $script:OfficeHandovers) { $R.report.office_notes.Add($note) }
    $script:OfficeHandovers.Clear()
    $R.report.message = Get-RunMessage
    $R.report.has_news = Test-ReportNews $R.report
    return $R.report
}

function Get-SelfTestProblem([string]$Work, [string]$Key) {
    # 준비의 자체 시험 결과(작업 폴더의 자체시험결과.json)의 $Key(pdf_ok · excel_ok). 통과했거나 파일이 없으면(준비 전 · 대표 PC) $null,
    # 통과하지 못했으면 그때의 사유 글자. 파일을 읽지 못하면 막는 쪽으로(사유 = 읽지 못함).
    $path = Join-Path $Work $script:SelfTestResultFile
    if (-not [IO.File]::Exists($path)) { return $null }
    try {
        $saved = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $ok = $saved.PSObject.Properties[$Key]
        if ($null -ne $ok -and $ok.Value -is [bool] -and $ok.Value) { return $null }
        $message = $saved.PSObject.Properties['message']
        if ($null -ne $message -and -not [string]::IsNullOrWhiteSpace([string]$message.Value)) { return [string]$message.Value }
        return '자체 시험을 통과하지 못했습니다.'
    } catch {
        return "$($script:SelfTestResultFile) 를 읽지 못했습니다: $($_.Exception.Message)"
    }
}

function Get-PreflightProblem($Job) {
    # 지금 판 _preflight: 입력 · 출력 폴더(로컬 · 있음)와 처리기록이 다른 프로그램에서 열려 있지 않은지.
    foreach ($pair in @(@('입력', $Job.input_dir), @('출력', $Job.output_dir))) {
        $label, $path = $pair
        try { Assert-LocalFolder $path } catch { return "$label 폴더를 쓸 수 없습니다. $($_.Exception.Message)" }
        if (-not [IO.Directory]::Exists($path)) {
            return "$label 폴더가 없습니다: $path`n폴더를 다시 만들거나(외장 디스크면 연결) 대표에게 알려 폴더를 다시 고르세요."
        }
    }
    $record = Join-Path $Job.folder $script:RecordFile
    if (Test-FileLocked $record) {
        return "$($script:RecordFile) 가 다른 프로그램(엑셀 등)에서 열려 있습니다. 닫고 다시 실행하세요.`n파일 위치: $record"
    }
    return $null
}

function Get-InputDocuments([string]$Folder) {
    # 처리 대상(.pdf · .xlsx, 엑셀 임시 파일 ~$ 제외). 순서는 지금 판 sorted(Path) 와 같게 소문자 이름의 서수 순서.
    if (-not [IO.Directory]::Exists($Folder)) { return @() }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($path in [IO.Directory]::GetFiles($Folder)) {
        $name = [IO.Path]::GetFileName($path)
        if ($name.StartsWith('~$')) { continue }
        if ($script:InputSuffixes -notcontains [IO.Path]::GetExtension($name).ToLowerInvariant()) { continue }
        $names.Add($name)
    }
    $sorted = $names.ToArray()
    $keys = [string[]]@($sorted | ForEach-Object { $_.ToLowerInvariant() })
    [Array]::Sort($keys, $sorted, [StringComparer]::Ordinal)
    return @($sorted | ForEach-Object { [IO.Path]::Combine($Folder, $_) })
}

function Get-UnmatchedFiles($Job, [string]$Work, [datetime]$Now) {
    # 입력 폴더에 하루 넘게 남아 있고, 같은 입력 폴더를 쓰는 어느 작업의 처리기록에도 없는 파일(지금 판 unmatched_files).
    # 워드 PDF 안내 때문에 PDF 를 열지 않은 실행에서는 PDF 를 세지 않는다(따로 알림이 감).
    $folder = $Job.input_dir
    if (-not [IO.Directory]::Exists($folder)) { return ,@() }
    $known = New-Object 'System.Collections.Generic.HashSet[string]'
    $folders = New-Object System.Collections.Generic.List[string]
    foreach ($other in (Get-PortableJobFolders $Work)) {
        try { if (Test-SameDir (Read-JobSettings $other).input_dir $folder) { $folders.Add($other) } } catch { }
    }
    if (-not @($folders | Where-Object { Test-SameDir $_ $Job.folder }).Count) { $folders.Add($Job.folder) }
    foreach ($other in $folders) {
        try {
            foreach ($row in (Read-CsvRecords (Join-Path $other $script:RecordFile) $script:RecordColumns)) { [void]$known.Add((Get-RecordValue $row '지문')) }
        } catch { }
    }
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($path in (Get-InputNames $folder)) {
        $name = [IO.Path]::GetFileName($path)
        $suffix = [IO.Path]::GetExtension($name).ToLowerInvariant()
        if ($script:UnmatchedSuffixes -notcontains $suffix -or $name.StartsWith('~$')) { continue }
        if ($suffix -eq '.pdf' -and -not $script:Run.pdf_ok) { continue }
        if ($suffix -eq '.xlsx' -and $null -ne $script:Run.excel_self_test) { continue }
        try {
            if (($Now - [IO.File]::GetLastWriteTime($path)).TotalSeconds -le $script:UnmatchedSeconds) { continue }
            if (-not $known.Contains((Get-FileFingerprint $path))) { $found.Add($path) }
        } catch { continue }
    }
    return ,$found
}

function Get-InputNames([string]$Folder) {
    $all = [IO.Directory]::GetFiles($Folder)
    $keys = [string[]]@($all | ForEach-Object { $_.ToLowerInvariant() })
    [Array]::Sort($keys, $all, [StringComparer]::Ordinal)
    return $all
}

# ---- 마크업표 -----------------------------------------------------------------------------------

function Read-RunMarkupBytes {
    # 마크업표.xlsx 를 실행 시작 때 한 번 읽는다(사용자가 엑셀로 열어 두어도 읽힘). 마크업지문 = 그 바이트(또는 없음 · 못 읽음 표시)의 sha256 앞 16자.
    # 확인필요 문서는 마크업지문이 바뀌면 같은 날이라도 다시 시도한다(지금 판은 이 바이트를 규칙지문에 섞는다).
    $R = $script:Run
    $path = Join-Path $R.work $script:MarkupFile
    try {
        if (-not [IO.File]::Exists($path)) { throw (New-Object IO.FileNotFoundException) }
        $R.markup_bytes = Read-SharedBytes $path
        $marker = [byte[]]($script:RunUtf8.GetBytes("file`0") + $R.markup_bytes)
    } catch [IO.FileNotFoundException] {
        $R.markup_read_error = "$($script:MarkupFile) 가 없습니다: $path"
        $marker = $script:RunUtf8.GetBytes('missing')
    } catch {
        $R.markup_read_error = "$($script:MarkupFile) 를 읽지 못했습니다(엑셀에서 열려 있으면 닫으세요): $($_.Exception.Message)"
        $marker = $script:RunUtf8.GetBytes('unreadable')
    }
    $R.markup_fp = (Get-TextSha256 $marker).Substring(0, 16)
}

function Get-RunMarkup {
    # 이번 실행의 마크업표(처음 쓸 때 시작 때 읽은 바이트로 만든다). 쓸 수 없으면 오류(사유 글자). 오류도 기억해 다음 문서마다 다시 읽지 않는다.
    $R = $script:Run
    if ($null -eq $R.markup_bytes) { throw $R.markup_read_error }
    if ($null -ne $R.markup_table) { return ,$R.markup_table }
    if ($R.markup_read_error) { throw $R.markup_read_error }
    $copy = New-RunTempFile $script:MarkupFile $R.markup_bytes
    try {
        try {
            $raw = Invoke-OfficeWork $script:ExcelWork $copy.Path 'EXCEL.EXE' $script:OfficeTimeoutSec $script:ExcelTimeoutText '엑셀'
            $grid = ConvertTo-ExcelGrid $raw
        } catch {
            $R.markup_read_error = "$($script:MarkupFile) 를 엑셀로 읽지 못했습니다: $($_.Exception.Message)"
            throw $R.markup_read_error
        } finally {
            Add-OfficeNote $script:MarkupFile '읽기'
        }
        try {
            $R.markup_table = Read-MarkupGrid $grid
        } catch {
            $R.markup_read_error = $_.Exception.Message
            throw $R.markup_read_error
        }
    } finally {
        Remove-OfficeCopy $copy
    }
    return ,$R.markup_table
}

function New-RunTempFile([string]$Name, [byte[]]$Bytes) {
    # %TEMP%\무설치판_<임의>\<이름> 에 바이트를 쓴다(엑셀은 이 복사본을 연다). Remove-OfficeCopy 로 지운다.
    $folder = New-OfficeTempFolder
    $path = Join-Path $folder $Name
    try { [IO.File]::WriteAllBytes($path, $Bytes) } catch { Remove-OfficeCopy @{ Folder = $folder; Path = $path }; throw }
    return @{ Folder = $folder; Path = $path }
}

function Add-OfficeNote([string]$Name, [string]$What) {
    # 마지막 오피스 일에서 Quit 뒤 스스로 끝나지 않아 강제로 끝낸 오피스가 있으면 실행기록에 남긴다(친구 PC 에서 정리 상태를 보려고).
    if ($script:OfficeForced -gt 0) {
        $script:Run.report.office_notes.Add(("참고: {0} {1} 뒤 오피스 {2}개를 강제로 끝냄" -f $Name, $What, $script:OfficeForced))
        $script:OfficeForced = 0
    }
}

# ---- 한 건 ---------------------------------------------------------------------------------------

function Invoke-OneDocument([string]$Path) {
    $R = $script:Run
    $name = [IO.Path]::GetFileName($Path)
    try {
        $age = ($R.now - [IO.File]::GetLastWriteTime($Path)).TotalSeconds
        if (-not [IO.File]::Exists($Path)) { return }
    } catch { return }
    if ($age -lt $script:SettleSeconds) { Add-RunDone $Path '보류' $script:SettlingReason; return }
    try { $fp = Get-FileFingerprint $Path } catch { Add-RunDone $Path '보류' $script:HeldReason; return }
    # 시험용: 지문을 구한 바로 뒤 파일이 사라지는 경우를 만든다(환경 변수가 있을 때만)
    if (-not [string]::IsNullOrEmpty($env:NOINSTALL_TEST_VANISH) -and $name -eq $env:NOINSTALL_TEST_VANISH) { [IO.File]::Delete($Path) }
    if (Test-RunSkip $fp) { return }   # 이미 처리한 원본은 PDF 뷰어에 열려 있어도 보류로 세지 않는다
    # 지금 판 os.access(path, W_OK): 읽기 전용 · 그새 사라진 파일 · 속성을 못 읽는 파일은 거짓 → 잠김 검사를 건너뛰고,
    # 사라진 파일은 아래 읽기에서 그 파일만 확인필요(파일을 읽지 못함)가 된다. 다른 문서는 계속 처리한다.
    $writable = $false
    try { $writable = -not ([IO.File]::GetAttributes($Path) -band [IO.FileAttributes]::ReadOnly) } catch { $writable = $false }
    if ($writable -and (Test-FileLocked $Path)) {   # 읽기 전용은 잠김 검사가 늘 실패하므로 건너뜀
        Add-RunDone $Path '보류' $script:HeldReason
        return
    }
    $isPdf = [IO.Path]::GetExtension($name).ToLowerInvariant() -eq '.pdf'
    if ($isPdf -and -not $R.pdf_ok) { $R.pdf_skipped++; return }   # 설계 6-4: 기록하지 않고 그대로 둔다(다음 실행 때 다시 봄)
    if (-not $isPdf -and $null -ne $R.excel_self_test) { $R.excel_skipped++; return }   # 엑셀 자체 시험 실패: 같은 길

    $result = $null
    $source = $null
    try {
        if ($isPdf) {
            $timeout = $script:OfficeTimeoutSec
            if ($null -ne $script:RunWordTimeoutSec) { $timeout = [int]$script:RunWordTimeoutSec }
            try {
                $read = Read-WordText $Path $timeout
            } finally { Add-OfficeNote $name '워드 읽기' }
            if ($null -ne $read.close_error) { $R.report.office_notes.Add(("참고: {0} 워드 문서를 닫지 못함: {1}" -f $name, $read.close_error)) }
            $source = @{ kind = 'text'; text = [string]$read.text }
            $result = Get-TextExtract $R.rules ([string]$read.text)
        } else {
            try {
                $grid = Read-ExcelGrid $Path $script:OfficeTimeoutSec
            } finally { Add-OfficeNote $name '엑셀 읽기' }
            $source = @{ kind = 'grid'; grid = $grid }
            $result = Get-GridExtract $R.rules $grid
        }
    } catch [System.TimeoutException] {
        Add-RunReview $Path $fp $null $_.Exception.Message   # 그 문서만 확인필요(워드가 응답하지 않음), 다음 문서로
        return
    } catch {
        if (Test-StopRun $_) { throw }
        Add-RunReview $Path $fp $null ("파일을 읽지 못함: " + $_.Exception.Message)   # 깨진 파일 한 건이 나머지를 막지 않게
        return
    }
    if ($null -eq $result.doc_type -and $result.review_reason -ne $script:NoText) {
        return   # 판별 실패: 다른 작업 몫일 수 있으므로 기록하지 않고 그대로 둔다
    }
    if ($null -ne $result.review_reason) { Add-RunReview $Path $fp $result $result.review_reason; return }
    $tableProblem = Get-TableProblem $R.rules $result   # 표 줄이 어긋나 품목이 통째로 빠진 결과를 「성공」 으로 내보내지 않는다
    if ($null -ne $tableProblem) { Add-RunReview $Path $fp $result $tableProblem; return }
    $doc = Get-WriterDoc $R.rules ([string]$result.doc_type)
    $numberText = ConvertTo-FormatText $result.fields[[string]$doc.number_field]
    $key = [string]$doc.name + "`0" + $numberText
    if ($R.success_keys.Contains($key)) {
        Add-RunRecord $Path $fp $result '중복'
        Clear-RunNote $Path
        Add-RunDone $Path '중복' ("{0} {1} 이미 처리됨" -f $doc.name, $numberText)
        return
    }
    try {
        Invoke-RunProcess $Path $fp $result $source $doc $numberText
    } catch {
        if (Test-StopRun $_) { throw }
        Add-RunReview $Path $fp $result ("처리 중 오류: " + $_.Exception.Message)   # 한 건의 오류가 나머지를 막지 않게
    }
}

function Test-RunSkip([string]$Fp) {
    # 지금 판 _skip(승인 없이): 성공 · 중복은 건너뜀. 확인필요는 같은 날 · 같은 규칙지문 · 같은 마크업지문이면 건너뜀. 보류는 다시.
    $R = $script:Run
    if (-not $R.state.ContainsKey($Fp)) { return $false }
    $row = $R.state[$Fp]
    $status = Get-RecordValue $row '상태'
    if ($status -eq '성공' -or $status -eq '중복') { return $true }
    if ($status -eq '확인필요') {
        return ((Get-RecordValue $row '날짜') -eq $R.today -and (Get-RecordValue $row '규칙지문') -eq $R.rules_fp -and
                (Get-RecordValue $row '마크업지문') -eq $R.markup_fp)
    }
    return $false
}

function Format-FailedChecks($Checks, [string]$Source, [string]$Output) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($c in $Checks) {
        if ($null -ne $c -and $c['ok'] -eq $false) { $parts.Add(("{0} ({1} {2}, 결과 {3})" -f $c['item'], $Source, $c['source_value'], $c['result_value'])) }
    }
    return ,$parts
}

function Invoke-RunProcess([string]$Path, [string]$Fp, $Result, $Source, $Doc, [string]$NumberText) {
    # 지금 판 _process(승인 · 고객마스터 없이): 검산 → (마크업) 없는 파트 · 마크업 검산 → 결과 쓰기.
    $R = $script:Run
    $checks = Get-Checks $Result $Source $R.rules
    $bad = Format-FailedChecks $checks '원본'
    if ($bad.Count -gt 0) { Add-RunReview $Path $Fp $Result ("검산 불일치: " + [string]::Join('; ', $bad.ToArray())); return }
    $form = $R.form
    $markup = $null
    $all = New-Object System.Collections.Generic.List[object]
    foreach ($c in $checks) { $all.Add($c) }
    if ($R.markup_used) {
        try { $markup = Get-RunMarkup } catch { Add-RunReview $Path $Fp $Result $_.Exception.Message; return }
        $missing = Get-MissingReason $Result $form $markup
        if ($null -ne $missing) { Add-RunReview $Path $Fp $Result $missing; return }   # 0% 로 넘기지 않는다
        $rows = (Get-CompanyRows $Result $form $markup).rows
        $more = Get-MarkupChecks $Result (Get-WrittenMarkup $form $rows) $markup (Get-MarkupProp $form 'markup_columns')
        $bad = Format-FailedChecks $more '계산'
        if ($bad.Count -gt 0) { Add-RunReview $Path $Fp $Result ("마크업 검산 불일치: " + [string]::Join('; ', $bad.ToArray())); return }
        foreach ($c in $more) { $all.Add($c) }   # 검산 파일에 줄마다 마크업% 와 본사 단가가 보이게
    }
    $company = $null
    if ($script:CompanyModes -contains $R.job.output_mode) { $company = Get-CompanyRows $Result $form $markup }
    Invoke-RunWrite $Path $Fp $Result $Doc $NumberText $all $company
}

function Get-LedgerPath {
    $job = $script:Run.job
    return [IO.Path]::Combine($job.output_dir, ($script:LedgerName -f (Get-FolderTitle $job.title)))
}

function Add-RunHold([string]$Path, [string]$Fp, $Result, [string]$Reason) {
    # 누적장부를 쓸 수 없어 보류. 판별된 파일이라는 표시로 처리기록에 보류를 남긴다(「어느 작업에도 맞지 않는 파일」 로 세지 않게).
    # 보류 기록은 건너뛰지 않으므로 다음 실행 때 다시 처리한다.
    $script:Run.ledger_held++
    Add-RunRecord $Path $Fp $Result '보류'
    Add-RunDone $Path '보류' $Reason
}

function Undo-RunWrite($Added) {
    # 이번 건이 장부에 쓴 것을 되돌린다(지금 판 _rollback). 되돌리지 못하면 사유에 붙일 글자.
    if ($null -eq $Added -or -not $Added.saved) { return '' }
    try {
        Undo-LedgerEntry $Added
        return ''
    } catch {
        return ". 누적장부에서 이 문서 줄을 지우지 못함($($_.Exception.Message)). 누적장부를 열어 직접 지우세요"
    }
}

function Invoke-RunWrite([string]$Path, [string]$Fp, $Result, $Doc, [string]$NumberText, $Checks, $Company) {
    # 지금 판 _write: 누적장부 · 회사 양식. 저장한 뒤 다시 읽어 다르면 이번 건만 되돌리고 확인필요.
    # 여기까지 왔으면 처리기록에는 이 (문서종류, 번호)의 성공이 없다. 그런데 결과가 이미 있으면 지난 실행이 결과를 쓴 뒤
    # 기록 전에 끊긴 것일 수 있으므로, 다시 쓰지 않고 확인필요로 보낸다.
    $R = $script:Run
    $mode = $R.job.output_mode
    $form = $R.form
    $name = [IO.Path]::GetFileName($Path)
    $writeLedger = $script:LedgerModes -contains $mode
    $writeCompany = $script:CompanyModes -contains $mode
    if ($writeCompany) {
        $final = [IO.Path]::Combine($R.job.output_dir, (Get-CompanyFormName $Result $form).Replace('/', '\'))
        if ([IO.File]::Exists($final)) {
            $finalName = [IO.Path]::GetFileName($final)
            $also = if ($writeLedger) { ' 누적장부에 이 문서 줄이 있으면 그것도 지우세요.' } else { '' }
            Add-RunReview $Path $Fp $Result ("출력 폴더에 회사 양식 $finalName 이 이미 있는데 $($script:Interrupted). 두 번 입력하지 않도록 새로 " +
                "만들지 않았습니다. 그 파일이 맞으면 $($script:KeepHint) 다시 만들려면 $finalName 을 지우세요.$also $($script:RemakeHint)")
            return
        }
    }
    if ($writeLedger -and $R.ledger_locked) { Add-RunHold $Path $Fp $Result $script:LedgerLockedReason; return }
    if ($writeLedger -and $null -ne $R.ledger_stop) { Add-RunHold $Path $Fp $Result $R.ledger_stop; return }
    $ledgerPath = Get-LedgerPath
    if ($writeLedger) {
        try {
            Assert-LedgerNotLost $ledgerPath
        } catch {
            # 장부가 없는데 바로 전 판(.bak)이 있다: 새 장부를 만들면 이력을 덮으므로 이 실행의 장부 쓰기를 모두 멈추고 알린다
            $R.ledger_stop = $_.Exception.Message
            $R.report.alerts.Add($R.ledger_stop)
            Add-RunHold $Path $Fp $Result $R.ledger_stop
            return
        }
    }
    $added = $null
    $companyFile = $null
    $problems = New-Object System.Collections.Generic.List[string]
    try {
        if ($writeLedger) {
            try {
                $added = Add-LedgerEntry $ledgerPath $R.rules $Result $name $R.now $Path
            } finally { Add-OfficeNote $name '장부 쓰기' }
            if ($added.locked) {
                $R.ledger_locked = $true
                Add-RunHold $Path $Fp $Result $script:LedgerLockedReason
                return
            }
            if ($added.duplicate) {
                $also = if ($writeCompany) { ' 출력 폴더에 이 문서의 회사 양식이 있으면 그것도 지우세요.' } else { '' }
                Add-RunReview $Path $Fp $Result ("누적장부에는 $($Doc.name) $NumberText 이 있는데 $($script:Interrupted). 결과가 맞으면 $($script:KeepHint) " +
                    "다시 만들려면 누적장부에서 이 문서 줄(문서 시트와 품목 시트)을 지우세요.$also $($script:RemakeHint)")
                return
            }
            foreach ($p in $added.problems) { $problems.Add([string]$p) }
        }
        if ($writeCompany -and $problems.Count -eq 0) {
            try {
                $companyFile = Write-CompanyForm $R.job.output_dir $Result $Company $form -Exact
            } catch {
                $e = $_.Exception
                if ($null -ne $e.Data -and $e.Data.Contains('problems')) {
                    foreach ($p in $e.Data['problems']) { $problems.Add([string]$p) }
                } elseif ($null -ne $e.Data -and $e.Data.Contains('exists')) {
                    # 위에서 없던 같은 이름이 그 사이 생김(지금 판 _write_company 의 「저장하는 사이 … 생겼습니다」)
                    $problems.Add(('저장하는 사이 출력 폴더에 {0} 이 생겼습니다' -f [IO.Path]::GetFileName($final)))
                } elseif ($e.Message -like '저장하는 사이 출력 폴더에 *') {
                    $problems.Add($e.Message)
                } else { throw }
            } finally { Add-OfficeNote $name '회사 양식 쓰기' }
        }
    } catch {
        if (Test-StopRun $_) { throw }
        $message = $_.Exception.Message
        $undo = Undo-RunWrite $added
        Add-RunReview $Path $Fp $Result ("결과를 쓰는 중 오류: $message$undo")
        return
    }
    if ($problems.Count -gt 0) {
        $undo = Undo-RunWrite $added
        $shown = [string]::Join('; ', @($problems | Select-Object -First 5))
        if ($problems.Count -gt 5) { $shown += (' 외 {0}건' -f ($problems.Count - 5)) }
        Add-RunReview $Path $Fp $Result ("저장한 파일을 다시 읽어 보니 쓴 값과 다릅니다: $shown$undo")
        return
    }

    try {
        try {
            [void](Write-CheckFile (Join-Path $R.job.folder $script:ChecksDir) $Result $Checks $name $Company $R.rules)
        } finally { Add-OfficeNote $name '검산 파일 쓰기' }
    } catch {
        if (Test-StopRun $_) { throw }
        $R.warnings.Add(("경고: 검산 파일을 만들지 못함 ({0}): {1}" -f $name, $_.Exception.Message))   # 결과는 이미 나갔으므로 알리기만 한다
    }
    if ($null -ne $companyFile) { $R.report.output_files.Add([string]$companyFile) }
    if ($null -ne $added -and -not $R.report.output_files.Contains($ledgerPath)) { $R.report.output_files.Add($ledgerPath) }
    Add-RunRecord $Path $Fp $Result '성공'   # 못 쓰면 실행을 멈춘다
    [void]$R.success_keys.Add([string]$Doc.name + "`0" + $NumberText)
    Clear-RunNote $Path
    Add-RunDone $Path '성공' ([string]::Join('; ', @($Result.warnings)))
}

# ---- 기록 ----------------------------------------------------------------------------------------

function Add-RunRecord([string]$Path, [string]$Fp, $Result, [string]$Status) {
    $R = $script:Run
    $docName = ''
    $number = ''
    if ($null -ne $Result -and $null -ne $Result.doc_type) {
        $docName = [string]$Result.doc_type
        $doc = Get-WriterDoc $R.rules $docName
        $value = $Result.fields[[string]$doc.number_field]
        if ($null -ne $value) { $number = ConvertTo-FormatText $value }
    }
    $row = @{ '지문' = $Fp; '파일' = [IO.Path]::GetFileName($Path); '문서종류' = $docName; '번호' = $number; '상태' = $Status
              '규칙지문' = $R.rules_fp; '날짜' = $R.today; '양식지문' = ''; '마크업지문' = $R.markup_fp }
    try {
        Add-CsvRow (Join-Path $R.job.folder $script:RecordFile) $script:RecordColumns @($script:RecordColumns | ForEach-Object { $row[$_] })
    } catch {
        # 기록 없이 계속하면 다음 실행이 같은 문서를 또 내보낼 수 있다(주문 두 번 입력). 여기서 멈춘다
        throw (New-StopRun ("{0} 에 쓰지 못해 실행을 멈췄습니다({1}, {2}: {3}). 이 파일의 결과는 이미 만들어졌을 수 있으니 출력 폴더를 확인하세요. 남은 파일은 처리하지 않았습니다. {0} 를 연 프로그램을 닫고 다시 실행하세요." -f
            $script:RecordFile, [IO.Path]::GetFileName($Path), $Status, $_.Exception.Message))
    }
    $R.state[$Fp] = $row
}

function Get-RunNotePath([string]$Path) {
    return [IO.Path]::Combine($script:Run.job.folder, $script:ReviewDir, ([IO.Path]::GetFileName($Path) + '.txt'))
}

function Clear-RunNote([string]$Path) {
    # 이 파일이 지난번에 남긴 확인필요 사유 파일을 지운다(이제 처리됐으므로).
    try {
        $note = Get-RunNotePath $Path
        if ([IO.File]::Exists($note)) { [IO.File]::Delete($note) }
    } catch {
        $script:Run.warnings.Add(("경고: 지난 확인필요 사유 파일을 지우지 못함 ({0}): {1}" -f [IO.Path]::GetFileName($Path), $_.Exception.Message))
    }
}

function Add-RunReview([string]$Path, [string]$Fp, $Result, [string]$Reason) {
    # 지금 판 _review: 확인필요\<원본이름>.txt(원본 · 사유 · 다시 시도 안내) + 처리기록 확인필요 + 처리로그.
    $R = $script:Run
    $note = Get-RunNotePath $Path
    try {
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($note))
        $text = "원본: $Path`n사유: $Reason`n`n$($script:RetryNote)`n"
        [IO.File]::WriteAllText($note, (ConvertTo-WindowsText $text), $script:RunUtf8)
    } catch {
        $R.warnings.Add(("경고: 확인필요 사유 파일을 쓰지 못함 ({0}): {1}" -f [IO.Path]::GetFileName($Path), $_.Exception.Message))
    }
    Add-RunRecord $Path $Fp $Result '확인필요'
    Add-RunDone $Path '확인필요' $Reason
}

function Add-RunDone([string]$Path, [string]$Outcome, [string]$Reason) {
    $R = $script:Run
    $R.report.counts[$Outcome] = $R.report.counts[$Outcome] + 1
    try {
        Add-CsvRow (Join-Path $R.job.folder $script:LogFile) $script:LogColumns @(
            $R.now.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant), [IO.Path]::GetFileName($Path), $Outcome, $Reason)
    } catch {
        $R.warnings.Add(("경고: {0} 에 쓰지 못함 ({1})" -f $script:LogFile, $_.Exception.Message))
    }
}

function Get-RunMessage {
    # 지금 판 _message(승인대기 줄 없음).
    $R = $script:Run
    $counts = $R.report.counts
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($key in $script:NewsKeys) { $parts.Add("$key $($counts[$key])") }
    if ($counts['보류']) { $parts.Add("보류 $($counts['보류'])") }
    $lines = New-Object System.Collections.Generic.List[string]
    if ($R.stop_reason) { $lines.Add("중요: $($R.stop_reason)") }
    $lines.Add("$($R.job.title): " + [string]::Join(' · ', $parts.ToArray()))
    foreach ($a in $R.report.alerts) { $lines.Add("경고: $a") }
    if ($R.ledger_locked) {
        $lines.Add("누적장부가 엑셀에서 열려 있어 파일 $($R.ledger_held)개를 처리하지 못했습니다(보류). 엑셀을 닫으면 다음 실행 때 처리합니다: $(Get-LedgerPath)")
    }
    if ($counts['확인필요']) { $lines.Add("확인필요 폴더의 .txt 를 열어 사유를 보세요: $([IO.Path]::Combine($R.job.folder, $script:ReviewDir))") }
    if ($counts['보류']) { $lines.Add('보류 파일은 입력 폴더에 그대로 있고 다음 실행 때 처리합니다.') }
    if ($R.report.unmatched_old) {
        $lines.Add("어느 작업에도 맞지 않는 파일 $($R.report.unmatched_old)개가 입력 폴더에 하루 넘게 남아 있습니다: $($R.job.input_dir)")
    }
    if ($R.report.unmatched_xls) {
        $lines.Add("그중 옛 엑셀(.xls) $($R.report.unmatched_xls)개: 엑셀에서 열어 .xlsx 로 다시 저장하세요(「다른 이름으로 저장」 → Excel 통합 문서).")
    }
    foreach ($w in $R.warnings) { $lines.Add($w) }
    return [string]::Join("`n", $lines.ToArray())
}

# ---- 명령 「실행」 -------------------------------------------------------------------------------

function Add-RunRecordText([string]$Folder, [string]$Label, [string]$Message) {
    # 지금 판 __main__._append_record: 작업 폴더 실행기록.txt 에 「[일시] 이름」 + 내용 + 빈 줄.
    try {
        $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant)
        Add-TextFile (Join-Path $Folder $script:RunRecordFile) "[$stamp] $Label`n$Message`n`n"
    } catch {
        [Console]::Error.WriteLine("경고: $($script:RunRecordFile) 에 쓰지 못함 ($($_.Exception.Message))")
    }
}

function Test-JobLocked($ErrorRecord) {
    $e = $ErrorRecord.Exception
    return ($null -ne $e -and $null -ne $e.Data -and $e.Data.Contains('job_locked'))
}

function Invoke-WithLockRetry([scriptblock]$Call) {
    # 지금 판 schedule.retry_locked: $Call 이 잠김(job_locked)으로 실패하면 LockRetrySeconds 마다 LockRetryLimitSeconds 까지 다시 해 본다.
    # 그래도 잠겨 있으면 마지막 오류를 그대로 던진다(부른 쪽이 기록하고 알린다). 다른 오류는 바로 던진다.
    $waited = 0
    while ($true) {
        try {
            return (& $Call)
        } catch {
            if (-not (Test-JobLocked $_) -or $waited -ge $script:LockRetryLimitSeconds) { throw }
        }
        Start-Sleep -Seconds $script:LockRetrySeconds
        $waited += $script:LockRetrySeconds
    }
}

function Get-LockedText([string]$Folder) {
    # 지금 판 __main__._locked_text: 끝내 잠겨 있어 자동 실행 · 놓친 실행을 못 했을 때의 글자.
    $title = [IO.Path]::GetFileName($Folder)
    try { $title = (Read-JobSettings $Folder).title } catch { }
    return ("{0} 작업이 다른 일로 잠겨 있어 오늘 자동 실행을 못 했습니다. 지금 실행을 눌러 주세요(바탕화면 「지금 실행 - {0}」)." -f $title)
}

function Send-AutoFailure([string]$LogDir, [string]$What, [string]$Message, [string]$RecordLabel, [bool]$NoNotify) {
    # 지금 판 _auto_failed: 아무도 화면을 안 보는 자동 실행 · 놓친 실행의 실패를 실행기록.txt 에 남기고 알린다(알림 글에 기록 위치, 누르면 그 폴더).
    Add-RunRecordText $LogDir $RecordLabel $Message
    if (-not $NoNotify) { Send-FailureNotice $What ("$Message`n기록: " + (Join-Path $LogDir $script:RunRecordFile)) $LogDir }
}

function Invoke-RunCommand([string]$Work, [int]$Number, [datetime]$Now, [string]$Label, [bool]$NoNotify, [string]$OutputPath, [bool]$Auto = $false) {
    # -Command 실행: 작업 하나를 실행 → 화면 · 실행기록.txt 에 요약 → 알릴 것이 있으면 알림 → 반환 코드.
    #   반환 코드: 0 정상(확인필요 · 보류가 있어도) · 1 멈춤(시작 못함 · 처리기록을 쓰지 못함) · 2 같은 작업이 실행 중 · 3 그 밖의 오류(작업 없음 등)
    # $Auto(매일 자동 실행, 지금 판 run --자동): 잠겨 있으면 기다렸다 다시 해 보고(Invoke-WithLockRetry), 끝내 실패하면(2 · 3)
    #   실행기록에 남기고 실패 알림을 띄운다(아무도 화면을 안 보므로). 「지금 실행」 은 기다리지 않고 알리지도 않는다(화면에 보임).
    $what = '작업 {0:D2} 자동 실행' -f $Number
    $folder = $null
    try {
        $folder = Find-PortableJob $Work $Number
    } catch {
        $message = $_.Exception.Message
        [Console]::Out.WriteLine($message)   # 작업이 없음(지금 판 JobError 처럼 그 글자만)
        if ($Auto) { Send-AutoFailure $Work $what "$message`n$($script:AutoFixHint)" "$Label 실패" $NoNotify }
        if ($OutputPath) {
            [IO.File]::WriteAllText($OutputPath, (ConvertTo-Json -InputObject ([ordered]@{ error = $message; locked = $false }) -Depth 20), $script:RunUtf8)
        }
        return 3
    }
    try {
        if ($Auto) { $report = Invoke-WithLockRetry { Invoke-PortableJob $Work $Number $Now } }
        else { $report = Invoke-PortableJob $Work $Number $Now }
    } catch {
        $locked = Test-JobLocked $_
        $message = $_.Exception.Message
        if (-not $locked) { $message = "예상하지 못한 오류로 멈췄습니다: $message" }
        elseif ($Auto) { $message = Get-LockedText $folder }
        $code = if ($locked) { 2 } else { 3 }
        [Console]::Out.WriteLine($message)
        if ($Auto) { Send-AutoFailure $folder $what $message "$Label 실패" $NoNotify }
        else { Add-RunRecordText $folder "$Label 실패" $message }
        if ($OutputPath) {
            [IO.File]::WriteAllText($OutputPath, (ConvertTo-Json -InputObject ([ordered]@{ error = $message; locked = $locked }) -Depth 20), $script:RunUtf8)
        }
        return $code
    }
    $text = $report.message
    if ($report.office_notes.Count -gt 0) { $text += "`n" + [string]::Join("`n", $report.office_notes.ToArray()) }
    [Console]::Out.WriteLine($report.message)
    Add-RunRecordText $folder $Label $text
    if ($OutputPath) { [IO.File]::WriteAllText($OutputPath, (ConvertTo-Json -InputObject $report -Depth 20), $script:RunUtf8) }
    if ($report.has_news -and -not $NoNotify) {
        try {
            $job = Read-JobSettings $folder
            Send-ReportNotice $job $report
        } catch { }
    }
    if ($report.stopped) { return 1 }
    return 0
}
