# 준비 · 매일 실행 등록 · 놓친 실행 · 자체 시험 · 점검(친구 PC 의 준비.bat, 대표 PC 의 점검).
# 지금 판 app\init.py · app\schedule.py · app\shortcuts.py · app\__main__.py(catch-up · check)와 같은 순서 · 같은 문구를 따른다.
#   - 매일 실행: Windows 작업 스케줄러 → 막히면 시작프로그램 「놓친 실행」 바로가기 → 그것도 막히면 없음(바탕화면 「지금 실행」 만).
#   - 레지스트리 · 시스템 설정은 읽기만 한다. 등록하는 것은 현재 사용자의 예약 작업과 바로가기(.lnk) 뿐이다.
# 점-소싱 순서: … → Notify → Run → Setup.

$script:PortableDir = Split-Path -Parent $PSScriptRoot                  # 묶음\무설치판\
$script:DispatcherPath = Join-Path $script:PortableDir '무설치판.ps1'
$script:RunBatPath = Join-Path $script:PortableDir '지금실행.bat'
$script:PowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$script:TaskPath = '\사무자동화 무설치판\'          # 지금 판(\사무자동화\)과 겹치지 않게 따로
$script:CatchUpLink = '사무자동화 무설치판 놓친 작업 실행.lnk'
$script:RunNowLink = '지금 실행 - {0}.lnk'
$script:MethodTask = '작업스케줄러'
$script:MethodStartup = '시작프로그램'
$script:MethodNone = '없음'
$script:TaskTimeLimitHours = 2                     # 한 번 실행이 이보다 길면 작업 스케줄러가 멈춘다(다음 날을 막지 않게)
$script:FirstSlotMinutes = 16 * 60                 # 작업 n 의 매일 시각 = 16:00 + 20분 × (n − 1) (지금 판 slot_time)
$script:SlotStepMinutes = 20
$script:JobInfoFile = '작업.json'                  # 묶음의 작업마다 {number, title, output_mode} (작업 11 이 만든다)
$script:SelfTestDir = '자체시험'
$script:SelfTestAnswerFile = '정답.json'
$script:SelfTestRulesFile = '규칙.json'
$script:SetupTitle = '사무자동화 무설치판 준비'
$script:XlsReason = '옛 엑셀 형식(.xls)은 읽지 못합니다. 엑셀에서 열어 「다른 이름으로 저장」 → .xlsx 로 저장한 뒤 입력 폴더에 다시 넣으세요'
$script:PdfDifferent = '이 PC 의 워드는 PDF 를 다르게 읽습니다. PDF 작업은 켜지 않았습니다'
$script:ExcelDifferent = '이 PC 의 엑셀은 엑셀 PO 를 다르게 읽습니다'
$script:PdfNotTested = '워드의 PDF 변환 안내가 켜져 있어 PDF 자체 시험을 하지 못했습니다. PDF 작업은 켜지 않았습니다'
$script:PdfWarningGuide = "워드에서 PDF 를 한 번 열고 '이 메시지를 다시 표시 안 함'에 체크하세요. 그 뒤 준비.bat 을 다시 두 번 클릭하면 PDF 작업도 켜집니다."
$script:PickRefused = '이 폴더는 쓸 수 없습니다. 내 PC 의 로컬 디스크 폴더를 다시 고르세요.'
$script:LineCheck = New-Object Text.RegularExpressions.Regex('^\d+줄 금액$')
# 마크업표.xlsx 가짜 예시(지금 판 app\markup.py 의 EXAMPLE_ROWS 와 같음): 파트넘버 · 마크업% · 반올림자리
$script:MarkupHeader = @('파트넘버', '마크업%', '반올림자리')
$script:MarkupExampleRows = @(
    @('VTS-CDU-300', 3, $null), @('VTS-RDHX-42U', 3, $null), @('VTS-MFD-12', 2, $null), @('VTS-QD-DN25', 3, $null),
    @('VTS-SENS-LEAK', 0, $null), @('VTS-CHL-500RT', 2, $null), @('VTS-PMP-CDU-K', 3, $null), @('VTS-FLT-50', 2, 2)
)
$script:SetupLines = $null
$script:SetupNoDialog = $false

# ---- 작은 도우미 ---------------------------------------------------------------------------------

function Read-SetupJson([string]$Path) {
    return ([IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8).TrimStart([char]0xFEFF) | ConvertFrom-Json)
}

function ConvertTo-PsQuoted([string]$Text) {
    # PowerShell 작은따옴표 글자(지금 판 powershell.quote): ' 와 ‘ ’ ‚ ‛ 를 두 번 쓴다.
    return "'" + [regex]::Replace($Text, "['‘’‚‛]", { param($m) $m.Value + $m.Value }) + "'"
}

function Format-ArgPath([string]$Path) {
    # 명령줄 "…" 안에 넣을 경로. 끝의 \ 는 닫는 따옴표를 먹으므로 뺀다(드라이브 맨 위 C:\ 는 C:\.).
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.EndsWith('\')) { if ($full.Length -le 3) { return $full + '.' } return $full.TrimEnd('\') }
    return $full
}

function Get-FolderNumber([string]$Folder) {
    # 작업 폴더 이름 NN_제목 의 NN(Find-PortableJob 이 찾는 기준).
    return [int]([regex]::Match([IO.Path]::GetFileName($Folder), '^([0-9]+)_').Groups[1].Value)
}

function Get-SlotTime([int]$Number) {
    $minutes = ($script:FirstSlotMinutes + $script:SlotStepMinutes * ($Number - 1)) % (24 * 60)
    if ($minutes -lt 0) { $minutes += 24 * 60 }
    return '{0:D2}:{1:D2}' -f [int][Math]::Floor($minutes / 60), [int]($minutes % 60)
}

function Get-JobTime($Job) {
    # 작업설정.json 의 time(HH:mm)이 있으면 그것, 없으면 번호로 정한 시각.
    $p = $Job.settings.PSObject.Properties['time']
    if ($null -ne $p -and [string]$p.Value -match '^([01][0-9]|2[0-3]):[0-5][0-9]$') { return [string]$p.Value }
    return Get-SlotTime (Get-FolderNumber $Job.folder)
}

function Get-SetupJob([string]$Folder, [string]$Work) {
    # 예약 · 바로가기에 쓰는 작업: Read-JobSettings + number · time · work(작업 폴더 맨 위).
    $job = Read-JobSettings $Folder
    $job.number = Get-FolderNumber $Folder
    $job.time = Get-JobTime $job
    $job.work = $Work
    return $job
}

function Get-TaskName($Job) { return '{0:D2}_{1}' -f $Job.number, (Get-FolderTitle $Job.title) }

function Get-PortableArguments([string]$Command, [string]$Work) {
    # 작업 스케줄러 · 시작프로그램이 powershell.exe 에 넘기는 인수(창 없이).
    return ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Command {1} -Work "{2}"' -f
            $script:DispatcherPath, $Command, (Format-ArgPath $Work))
}

function New-Shortcut([string]$Path, [string]$Target, [string]$Arguments, [string]$WorkDir, [string]$Description) {
    # 바로가기 하나(있으면 덮어씀, 지금 판 shortcuts.create). WScript.Shell(Windows 기본)만 쓴다. 만든 뒤 파일이 생겼는지 본다.
    $shell = New-Object -ComObject WScript.Shell
    try {
        $link = $shell.CreateShortcut($Path)
        $link.TargetPath = $Target
        $link.Arguments = $Arguments
        $link.WorkingDirectory = $WorkDir
        $link.Description = $Description
        $link.Save()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($link)
    } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    }
    if (-not [IO.File]::Exists($Path)) { throw '바로가기 파일이 생기지 않았습니다' }
}

function Read-Shortcut([string]$Path) {
    $shell = New-Object -ComObject WScript.Shell
    try {
        $link = $shell.CreateShortcut($Path)
        $out = [ordered]@{ target = [string]$link.TargetPath; arguments = [string]$link.Arguments; workdir = [string]$link.WorkingDirectory }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($link)
        return $out
    } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    }
}

# ---- 매일 실행 등록 · 바로가기 ----------------------------------------------------------------------

function Get-ScheduleScript($Job) {
    # 작업 스케줄러에 등록하는 PowerShell 글자(지금 판 _register_script): 매일 시각, 현재 사용자 · 로그인했을 때만,
    # 놓친 실행은 PC 를 켜면 바로(StartWhenAvailable), 겹쳐 돌지 않게(IgnoreNew), 2시간 넘으면 멈춤.
    $user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $arguments = (Get-PortableArguments ('실행 -Job {0:D2} -Auto' -f $Job.number) $Job.work)
    $q = { param($Text) ConvertTo-PsQuoted ([string]$Text) }
    return [string]::Join("`n", @(
        "`$action = New-ScheduledTaskAction -Execute $(& $q $script:PowerShellExe) -Argument $(& $q $arguments) -WorkingDirectory $(& $q $script:PortableDir)",
        "`$trigger = New-ScheduledTaskTrigger -Daily -At $(& $q $Job.time)",
        ("`$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries " +
         "-MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours $($script:TaskTimeLimitHours))"),
        "`$principal = New-ScheduledTaskPrincipal -UserId $(& $q $user) -LogonType Interactive -RunLevel Limited",
        ("Register-ScheduledTask -TaskPath $(& $q $script:TaskPath) -TaskName $(& $q (Get-TaskName $Job)) -Action `$action " +
         "-Trigger `$trigger -Settings `$settings -Principal `$principal -Description $(& $q ('사무자동화 무설치판 매일 실행: ' + $Job.title)) -Force | Out-Null")
    ))
}

# 진짜 등록: 글자를 돌리고 다시 조회해서 있어야 참. 시험은 이것을 끼운 가짜로 바꾼다(진짜 예약 작업을 만들지 않음).
$script:RegisterTaskDefault = {
    param([string]$Text, [string]$TaskPath, [string]$TaskName)
    $ProgressPreference = 'SilentlyContinue'
    Invoke-Expression $Text
    return ($null -ne (Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue))
}

function New-CatchUpLink([string]$Work, [string]$StartupDir) {
    # 시작프로그램의 「놓친 실행」 바로가기 하나(모든 작업이 같이 씀): 로그인할 때 오늘 시각이 지났는데 아직 안 돈 작업을 실행.
    $path = Join-Path $StartupDir $script:CatchUpLink
    New-Shortcut $path $script:PowerShellExe (Get-PortableArguments '놓친실행' $Work) $script:PortableDir '사무자동화 무설치판: 오늘 시각이 지났는데 아직 안 돈 작업 실행'
    return $path
}

function Register-PortableSchedule($Job, [switch]$WhatIf, [scriptblock]$RegisterTask, [string]$StartupDir, $Notes) {
    # 매일 실행을 등록하고 쓰인 방법(작업스케줄러 · 시작프로그램 · 없음)을 돌려준다. 작업설정.json 에 number · time · method 를 남긴다.
    # -WhatIf: 아무것도 하지 않고 작업 스케줄러에 넘길 명령 글자만 돌려준다. $Notes(List[string])를 주면 막힌 이유를 더한다.
    $text = Get-ScheduleScript $Job
    if ($WhatIf) { return $text }
    if ($null -eq $Notes) { $Notes = New-Object System.Collections.Generic.List[string] }
    if ($null -eq $RegisterTask) { $RegisterTask = $script:RegisterTaskDefault }
    if ([string]::IsNullOrEmpty($StartupDir)) { $StartupDir = [Environment]::GetFolderPath('Startup') }
    $method = $script:MethodNone
    try {
        if (& $RegisterTask $text $script:TaskPath (Get-TaskName $Job)) { $method = $script:MethodTask }
        else { $Notes.Add('작업 스케줄러에 등록했다고 했지만 다시 조회하니 없음') }
    } catch {
        $Notes.Add("작업 스케줄러에 등록하지 못함($($_.Exception.Message))")
    }
    if ($method -eq $script:MethodNone) {
        try {
            [void](New-CatchUpLink $Job.work $StartupDir)
            $method = $script:MethodStartup
        } catch {
            $Notes.Add("시작프로그램 바로가기를 만들지 못함($($_.Exception.Message))")
        }
    }
    Save-JobSettingsValues $Job ([ordered]@{ number = $Job.number; time = $Job.time; method = $method })
    return $method
}

function New-RunShortcut($Job, [string]$Desktop) {
    # 바탕화면 「지금 실행 - 제목」 = 무설치판\지금실행.bat NN (작업 폴더가 묶음 폴더가 아니면 -Work "…" 를 붙임). 경로를 돌려준다.
    $path = Join-Path $Desktop ($script:RunNowLink -f (Get-FolderTitle $Job.title))
    $arguments = '{0:D2}' -f $Job.number
    if (-not (Test-SameDir $Job.work (Split-Path -Parent $script:PortableDir))) { $arguments += (' -Work "{0}"' -f (Format-ArgPath $Job.work)) }
    New-Shortcut $path $script:RunBatPath $arguments $script:PortableDir "사무자동화 무설치판 $($Job.title) 지금 실행"
    return $path
}

function Get-MethodLine($Job, [string]$Method) {
    # 지금 판 _method_line.
    if ($Method -eq $script:MethodTask) { return "자동 실행: Windows 작업 스케줄러 (매일 $($Job.time))" }
    if ($Method -eq $script:MethodStartup) {
        return ("자동 실행: 시작프로그램 바로가기 (작업 스케줄러가 막혀 있어 대신 씀. PC 에 로그인할 때 " +
                "매일 $($Job.time) 이 지났고 오늘 아직 안 돌았으면 실행합니다. 정각 실행은 안 됩니다)")
    }
    return "자동 실행: 없음 (작업 스케줄러 · 시작프로그램이 모두 막혀 있습니다). 바탕화면 「지금 실행 - $($Job.title)」을 눌러 실행하세요."
}

# ---- 놓친 실행 -------------------------------------------------------------------------------------

function Test-JobDue($Job, [string]$Today, [string]$Clock) {
    # 오늘 시각이 지났고(시각 ≤ 지금) 오늘 아직 안 돈(last_run ≠ 오늘) 작업(지금 판 due_jobs, 일시중지는 무설치판에 없음).
    $last = $Job.settings.PSObject.Properties['last_run']
    $lastRun = if ($null -ne $last -and $null -ne $last.Value) { [string]$last.Value } else { '' }
    return ($lastRun -ne $Today -and [string]::CompareOrdinal((Get-JobTime $Job), $Clock) -le 0)
}

function Invoke-Catchup([string]$Work, [datetime]$Now) {
    # 놓친 작업을 번호 순으로 하나씩 실행한다(지금 판 schedule.catch_up + __main__._catch_up 의 실행 부분).
    # 실행 직전에 작업설정.json 을 다시 읽어 그새 다른 실행(작업 스케줄러)이 끝냈으면 건너뛴다. 잠긴 작업은 기다렸다 다시 해 보고,
    # 끝내 잠겨 있으면 locked 에 더한다(부른 쪽이 기록하고 알림). 한 작업의 예상 못 한 오류는 기록하고 다음 작업으로.
    # 돌려줌: { ran(작업 이름들); locked(작업들); failed(글자들); finished(@{job; report}들) }
    if ($Now -eq [datetime]::MinValue) { $Now = [datetime]::Now }
    $today = $Now.ToString('yyyy-MM-dd', $script:Invariant)
    $clock = $Now.ToString('HH:mm', $script:Invariant)
    $out = [ordered]@{ ran = (New-Object System.Collections.Generic.List[string]); locked = (New-Object System.Collections.Generic.List[object])
                       failed = (New-Object System.Collections.Generic.List[string]); finished = (New-Object System.Collections.Generic.List[object]) }
    $due = New-Object System.Collections.Generic.List[string]
    foreach ($folder in (Get-PortableJobFolders $Work)) {
        try { if (Test-JobDue (Read-JobSettings $folder) $today $clock) { $due.Add($folder) } } catch { }   # 깨진 작업은 부른 쪽이 따로 알린다
    }
    foreach ($folder in @($due | Sort-Object { Get-FolderNumber $_ })) {
        $waited = 0
        while ($true) {
            try { $job = Read-JobSettings $folder } catch { break }   # 그새 지워짐 · 깨짐
            if (-not (Test-JobDue $job $today $clock)) { break }     # 그새 다른 실행이 끝냄
            try {
                $report = Invoke-PortableJob $Work (Get-FolderNumber $folder) $Now
            } catch {
                if (Test-JobLocked $_) {
                    if ($waited -ge $script:LockRetryLimitSeconds) { $out.locked.Add($job); break }
                    Start-Sleep -Seconds $script:LockRetrySeconds
                    $waited += $script:LockRetrySeconds
                    continue
                }
                # 한 작업의 예상 못 한 오류가 다음 작업과 앞 작업의 알림을 막지 않게
                $message = "예상하지 못한 오류로 멈췄습니다: $($_.Exception.Message)"
                [Console]::Out.WriteLine("$($job.name): $message")
                Add-RunRecordText $folder '놓친 실행 실패' $message
                $out.failed.Add(("{0}: {1} (기록: {2})" -f $job.name, $message, (Join-Path $folder $script:RunRecordFile)))
                break
            }
            $text = $report.message
            if ($report.office_notes.Count -gt 0) { $text += "`n" + [string]::Join("`n", $report.office_notes.ToArray()) }
            [Console]::Out.WriteLine($report.message)
            Add-RunRecordText $folder '놓친 실행' $text
            $out.ran.Add($job.name)
            $out.finished.Add(@{ job = $job; report = $report })
            break
        }
    }
    return $out
}

function Invoke-CatchupCommand([string]$Work, [datetime]$Now, [bool]$NoNotify, [string]$OutputPath) {
    # -Command 놓친실행(시작프로그램 바로가기): 놓친 작업 실행 → 결과 알림 → 실패는 모아 기록 · 알림(지금 판 _catch_up). 반환 코드 0 · 1.
    $what = '놓친 작업 실행'
    $problems = New-Object System.Collections.Generic.List[string]
    $ran = @()
    $lockedNames = @()
    $stopped = $false
    try {
        Assert-LocalFolder $Work
        $out = Invoke-Catchup $Work $Now
        $ran = @($out.ran)
        foreach ($job in $out.locked) {   # 끝내 잠겨 있던 작업: 조용히 넘어가지 않게 기록하고 아래에서 알린다
            $text = Get-LockedText $job.folder
            [Console]::Out.WriteLine($text)
            Add-RunRecordText $job.folder '놓친 실행 실패' $text
            $problems.Add("$($job.name): $text")
            $lockedNames += $job.name
        }
        foreach ($f in $out.failed) { $problems.Add($f) }
        foreach ($done in $out.finished) {
            if ($done.report.stopped) { $stopped = $true }
            if ($done.report.has_news -and -not $NoNotify) { try { Send-ReportNotice $done.job $done.report } catch { } }
        }
        if ($ran.Count -eq 0) { [Console]::Out.WriteLine('지금 실행할 놓친 작업이 없습니다.') }
        $broken = New-Object System.Collections.Generic.List[string]
        foreach ($folder in (Get-PortableJobFolders $Work)) {
            try { [void](Read-JobSettings $folder) } catch { $broken.Add("$([IO.Path]::GetFileName($folder)): $($_.Exception.Message)") }
        }
        if ($broken.Count -gt 0) {
            $problems.Add("설정이 깨진 작업이 있어 실행하지 못했습니다.`n$([string]::Join("`n", $broken.ToArray()))`n$($script:AutoFixHint)")
        }
    } catch {
        $problems.Add("예상하지 못한 오류로 멈췄습니다: $($_.Exception.Message)")
    }
    $code = 0
    if ($problems.Count -gt 0) {   # 결과 알림을 먼저 띄운 뒤, 실패는 한 번에 모아 알린다
        $message = [string]::Join("`n", $problems.ToArray())
        [Console]::Out.WriteLine($message)
        Send-AutoFailure $Work $what $message "$what 실패" $NoNotify
        $code = 1
    } elseif ($stopped) { $code = 1 }
    if ($OutputPath) {
        $json = ConvertTo-Json -InputObject ([ordered]@{ ran = [object[]]$ran; locked = [object[]]$lockedNames; problems = [object[]]$problems.ToArray() }) -Depth 20
        [IO.File]::WriteAllText($OutputPath, $json, $script:RunUtf8)
    }
    return $code
}

# ---- 자체 시험 -----------------------------------------------------------------------------------

function Find-JsonDifferences($Want, $Got, [string]$Path, $Found) {
    # ConvertFrom-Json 으로 읽은 두 값(정답 · 이 PC)을 견주어 다른 곳을 $Found(List[string])에 더한다. 이름 순서는 따지지 않는다.
    if ($null -eq $Want -or $null -eq $Got) {
        if (-not ($null -eq $Want -and $null -eq $Got)) { $Found.Add(("{0}: 정답 {1}, 이 PC {2}" -f $Path, (Format-JsonLeaf $Want), (Format-JsonLeaf $Got))) }
        return
    }
    if ($Want -is [System.Management.Automation.PSCustomObject]) {
        if (-not ($Got -is [System.Management.Automation.PSCustomObject])) { $Found.Add("${Path}: 모양이 다름"); return }
        $names = New-Object System.Collections.Generic.List[string]
        foreach ($p in $Want.PSObject.Properties) { $names.Add($p.Name) }
        foreach ($p in $Got.PSObject.Properties) { if (-not $names.Contains($p.Name)) { $names.Add($p.Name) } }
        foreach ($name in $names) {
            # 값은 대입으로만 꺼낸다($() · if 식은 빈 배열을 $null 로, 한 개짜리 배열을 낱개로 풀어 버린다)
            $wantValue = $null
            $gotValue = $null
            $w = $Want.PSObject.Properties[$name]
            $g = $Got.PSObject.Properties[$name]
            if ($null -ne $w) { $wantValue = $w.Value }
            if ($null -ne $g) { $gotValue = $g.Value }
            Find-JsonDifferences $wantValue $gotValue "$Path.$name" $Found
        }
        return
    }
    if ($Want -is [array]) {
        if (-not ($Got -is [array]) -or $Want.Count -ne $Got.Count) {
            $count = if ($Got -is [array]) { "$($Got.Count)개" } else { '목록 아님' }
            $Found.Add("${Path}: 정답 $($Want.Count)개, 이 PC $count")
            return
        }
        for ($i = 0; $i -lt $Want.Count; $i++) { Find-JsonDifferences $Want[$i] $Got[$i] "$Path[$i]" $Found }
        return
    }
    if ($Want.GetType() -ne $Got.GetType() -or -not ($Want -ceq $Got)) {
        $Found.Add(("{0}: 정답 {1}, 이 PC {2}" -f $Path, (Format-JsonLeaf $Want), (Format-JsonLeaf $Got)))
    }
}

function Format-JsonLeaf($Value) {
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [string]) { return "'$Value'" }
    return ConvertTo-Json -InputObject $Value -Depth 20 -Compress
}

function Test-SelfTestPart([string]$Label, $Want, [scriptblock]$Read) {
    # 자체 시험 한 가지(PDF · 엑셀): $Read 가 이 PC 의 엔진 결과(JSON 글자)를 돌려준다. → @{ ok; difference(글자) }
    $found = New-Object System.Collections.Generic.List[string]
    $got = (& $Read) | ConvertFrom-Json
    Find-JsonDifferences $Want $got $Label $found
    return @{ ok = ($found.Count -eq 0); difference = [string]::Join('; ', @($found | Select-Object -First 10)) }
}

function Invoke-SelfTest([string]$BundleDir) {
    # 묶음의 자체시험\(가짜 PDF 1 · 엑셀 1 · 규칙.json · 정답.json, app\portable.build_self_test)을 이 PC 의 워드 · 엑셀로 읽어
    # 엔진 결과가 정답(지금 판 엔진 결과)과 같은지 본다. → [ordered]@{ ok; pdf_ok; excel_ok; message; detail }
    # 워드의 PDF 변환 안내가 켜져 있으면 PDF 는 열지 않는다(숨은 확인 창에서 멈추므로): pdf_ok 거짓(안내를 끈 뒤 준비를 다시).
    $folder = Join-Path $BundleDir $script:SelfTestDir
    $out = [ordered]@{ ok = $false; pdf_ok = $false; excel_ok = $false; message = ''; detail = '' }
    $messages = New-Object System.Collections.Generic.List[string]
    $details = New-Object System.Collections.Generic.List[string]
    try {
        $answer = Read-SetupJson (Join-Path $folder $script:SelfTestAnswerFile)
        $rules = ConvertTo-PortableRules (Read-SetupJson (Join-Path $folder $script:SelfTestRulesFile))
        $pdf = Join-Path $folder ([string]$answer.pdf.file)
        $xlsx = Join-Path $folder ([string]$answer.excel.file)
    } catch {
        $out.message = "자체 시험 자료를 읽지 못했습니다($folder): $($_.Exception.Message)"
        return $out
    }
    $warningOff = if ($null -ne $script:RunPdfWarningOff) { [bool]$script:RunPdfWarningOff } else { Test-PdfWarningOff }
    if (-not $warningOff) {
        $messages.Add($script:PdfNotTested)
    } else {
        try {
            $part = Test-SelfTestPart 'PDF' $answer.pdf.result {
                ConvertTo-ResultJson (Get-TextExtract $rules ([string](Read-WordText $pdf $script:OfficeTimeoutSec '자체시험').text))
            }
            $out.pdf_ok = $part.ok
            if (-not $part.ok) { $messages.Add($script:PdfDifferent); $details.Add($part.difference) }
        } catch {
            $messages.Add("워드로 시험 PDF 를 읽지 못했습니다($($_.Exception.Message)). PDF 작업은 켜지 않았습니다")
        }
    }
    try {
        $part = Test-SelfTestPart '엑셀' $answer.excel.result {
            ConvertTo-ResultJson (Get-GridExtract $rules (Read-ExcelGrid $xlsx $script:OfficeTimeoutSec '자체시험'))
        }
        $out.excel_ok = $part.ok
        if (-not $part.ok) { $messages.Add($script:ExcelDifferent); $details.Add($part.difference) }
    } catch {
        $messages.Add("엑셀로 시험 엑셀을 읽지 못했습니다($($_.Exception.Message))")
    }
    $out.ok = $out.pdf_ok -and $out.excel_ok
    $out.message = [string]::Join("`n", $messages.ToArray())
    $out.detail = [string]::Join("`n", $details.ToArray())
    return $out
}

function Save-SelfTestResult([string]$Work, $Result) {
    # 작업 폴더의 자체시험결과.json: 실행(Run.ps1)이 pdf_ok 가 거짓이면 PDF 를 열지 않는다. 다시 준비하면 새 결과로 바뀐다.
    $data = [ordered]@{
        pdf_ok = [bool]$Result.pdf_ok; excel_ok = [bool]$Result.excel_ok
        date = (Get-Date).ToString('yyyy-MM-dd', $script:Invariant); office_version = (Get-OfficeVersion 'Word.Application')
        message = [string]$Result.message; detail = [string]$Result.detail
    }
    [IO.File]::WriteAllText((Join-Path $Work $script:SelfTestResultFile), (ConvertTo-Json -InputObject $data -Depth 20), $script:RunUtf8)
}

# ---- 준비 ---------------------------------------------------------------------------------------

function Write-SetupLine([string]$Text) {
    [Console]::Out.WriteLine($Text)
    if ($null -ne $script:SetupLines) { $script:SetupLines.Add($Text) }
}

function Show-SetupDialog([string]$Text, [int]$Icon) {
    if ($script:SetupNoDialog) { return }
    try { Show-MessageBox $script:SetupTitle $Text $Icon } catch { }
}

function Stop-Setup([string]$Message) {
    # 준비를 멈춘다: 화면 · 메시지 창에 사유. 반환 코드 1.
    Write-SetupLine $Message
    Show-SetupDialog $Message $script:MbIconWarning
    return 1
}

# 진짜 폴더 고르기 창(시험은 -FolderPicker 로 바꿔 끼운다). 취소하면 빈 글자.
$script:FolderPickerDefault = {
    param([string]$Prompt)
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $owner = New-Object System.Windows.Forms.Form
    try {
        $owner.TopMost = $true
        $dialog.Description = $Prompt
        $dialog.ShowNewFolderButton = $true
        if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.SelectedPath }
        return ''
    } finally {
        $owner.Dispose()
        $dialog.Dispose()
    }
}

function Select-LocalFolder([scriptblock]$Picker, [string]$Prompt) {
    # 폴더 고르기: 로컬 디스크만(지금 판 check_work_dir 와 같은 거절 문구). 쓸 수 없으면 다시 묻고, 취소하면 빈 글자.
    Write-SetupLine "폴더 고르기: $Prompt"
    while ($true) {
        $chosen = [string](@(& $Picker $Prompt) | Select-Object -Last 1)
        if ([string]::IsNullOrWhiteSpace($chosen)) { return '' }
        try {
            Assert-LocalFolder $chosen
        } catch {
            Write-SetupLine $_.Exception.Message
            Write-SetupLine $script:PickRefused
            Show-SetupDialog ($_.Exception.Message + "`n" + $script:PickRefused) $script:MbIconWarning
            continue
        }
        if (-not [IO.Directory]::Exists($chosen)) { Write-SetupLine "폴더가 없습니다: $chosen"; continue }
        return $chosen
    }
}

function Get-EnvironmentProblem {
    # 오피스 · PowerShell 버전 확인. 문제가 없으면 $null, 있으면 사유.
    if ($PSVersionTable.PSVersion.Major -lt 5) { return "Windows PowerShell 5.1 이 필요합니다(이 PC: $($PSVersionTable.PSVersion))." }
    foreach ($pair in @(@('워드', 'Word.Application'), @('엑셀', 'Excel.Application'))) {
        if ($null -eq (Get-OfficeVersion $pair[1])) { return "이 PC 에서 $($pair[0])를 찾지 못했습니다(오피스가 설치되어 있어야 합니다)." }
    }
    return $null
}

function Write-ExampleMarkup([string]$Path) {
    # 마크업표.xlsx 가짜 예시(지금 판 write_example_markup 와 같은 칸): 시트 마크업표, 머리글 굵게, 틀 고정 A2. 저장 뒤 다시 읽어 견줌.
    $header = [object[]]$script:MarkupHeader
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $script:MarkupExampleRows) { $rows.Add([object[]]@($r[0], $r[1], $r[2])) }
    $all = New-Object System.Collections.Generic.List[object]
    $all.Add($header)
    foreach ($r in $rows) { $all.Add($r) }
    [void](Save-NewBook $Path @((New-SheetPlan '마크업표' $all 1 2)) '마크업표' $header $rows)
}

function Copy-JobRules([string]$Source, [string]$Target) {
    # 묶음의 규칙.json 을 작업 폴더로(같은 폴더면 그대로 둔다). 임시 파일에 쓴 뒤 바꾼다.
    if (Test-SameDir $Source $Target) { return }
    $from = Join-Path $Source $script:RulesFile
    $to = Join-Path $Target $script:RulesFile
    $tmp = $to + '.tmp'
    [IO.File]::Copy($from, $tmp, $true)
    if ([IO.File]::Exists($to)) { [IO.File]::Replace($tmp, $to, [NullString]::Value) } else { [IO.File]::Move($tmp, $to) }
}

function Set-JobAutomation($Job, [scriptblock]$RegisterTask, [string]$StartupDir, [string]$Desktop) {
    # 매일 실행 등록 + 바탕화면 「지금 실행」 (지금 판 _install_automation).
    $notes = New-Object System.Collections.Generic.List[string]
    $method = Register-PortableSchedule $Job -RegisterTask $RegisterTask -StartupDir $StartupDir -Notes $notes
    Write-SetupLine "  $(Get-MethodLine $Job $method)"
    foreach ($note in $notes) { Write-SetupLine "    참고: $note" }
    try {
        $link = New-RunShortcut $Job $Desktop
        Write-SetupLine "  바탕화면 바로가기: $([IO.Path]::GetFileName($link))"
    } catch {
        Write-SetupLine ("  경고: 바탕화면 바로가기를 만들지 못함 ({0}). 대신 {1} 를 두 번 클릭하고 작업 번호 {2:D2} 로 실행하세요." -f
                         $_.Exception.Message, $script:RunBatPath, $Job.number)
    }
}

function Invoke-Setup([string]$BundleDir, [string]$Work, [scriptblock]$FolderPicker, [switch]$NoDialog,
                      [string]$Desktop, [string]$StartupDir, [scriptblock]$RegisterTask, [scriptblock]$SelfTest) {
    # 준비.bat(처음 한 번 · 새 묶음을 받았을 때): 오피스 · PowerShell 확인 → 자체 시험 → 작업마다(이미 있으면 규칙만 바꾸고
    # 폴더는 묻지 않음, 처리기록 · 장부는 그대로 / 새 작업이면 입력 · 출력 폴더 고르기) → 매일 실행 · 바로가기 →
    # 마크업표.xlsx 가 없으면 가짜 예시 → PDF 변환 안내가 켜져 있으면 끄라는 안내. 반환 코드 0(끝남) · 1(멈춤 · 일부 못 함).
    # 폴더 고르기 창 · 등록 · 자체 시험 · 바탕화면 · 시작프로그램 위치는 시험에서 바꿔 끼운다(-NoDialog 면 메시지 창 없음).
    $script:SetupLines = New-Object System.Collections.Generic.List[string]
    $script:SetupNoDialog = [bool]$NoDialog
    if ($null -eq $FolderPicker) { $FolderPicker = $script:FolderPickerDefault }
    if ($null -eq $SelfTest) { $SelfTest = { param($Bundle) Invoke-SelfTest $Bundle } }
    if ([string]::IsNullOrEmpty($Desktop)) { $Desktop = [Environment]::GetFolderPath('Desktop') }
    if ([string]::IsNullOrEmpty($StartupDir)) { $StartupDir = [Environment]::GetFolderPath('Startup') }

    Write-SetupLine "무설치판 준비: $BundleDir"
    $problem = Get-EnvironmentProblem
    if ($null -ne $problem) { return Stop-Setup "$problem`n준비를 멈췄습니다. 대표에게 알려 주세요." }
    Write-SetupLine ("확인: PowerShell {0} · 워드 {1} · 엑셀 {2}" -f $PSVersionTable.PSVersion, (Get-OfficeVersion 'Word.Application'), (Get-OfficeVersion 'Excel.Application'))
    try {
        Assert-LocalFolder $Work
    } catch {
        return Stop-Setup ("$($_.Exception.Message)`n묶음 폴더를 내 PC 의 로컬 디스크(예: C:\사무자동화)로 옮긴 뒤 준비.bat 을 다시 두 번 클릭하세요.")
    }
    $sources = @(Get-PortableJobFolders $BundleDir)
    if ($sources.Count -eq 0) { return Stop-Setup "묶음에 작업(작업\NN_제목\)이 없습니다: $BundleDir`n준비를 멈췄습니다. 대표에게 알려 주세요." }
    foreach ($source in $sources) {
        foreach ($name in @($script:RulesFile, $script:JobInfoFile)) {
            if (-not [IO.File]::Exists((Join-Path $source $name))) { return Stop-Setup "묶음의 작업 폴더에 $name 이 없습니다: $source`n준비를 멈췄습니다. 대표에게 알려 주세요." }
        }
    }

    Write-SetupLine '자체 시험: 이 PC 의 워드 · 엑셀로 시험 문서를 읽어 정답과 견줍니다(1분쯤 걸립니다).'
    $test = & $SelfTest $BundleDir
    try { Save-SelfTestResult $Work $test } catch { Write-SetupLine "경고: $($script:SelfTestResultFile) 를 쓰지 못함 ($($_.Exception.Message))" }
    if (-not $test.excel_ok) {
        return Stop-Setup ("$($test.message)`n준비를 멈췄습니다(매일 실행 · 바로가기를 만들지 않았습니다). 대표에게 알려 주세요.")
    }
    if ($test.pdf_ok) { Write-SetupLine '자체 시험: 통과(워드 PDF · 엑셀 읽기가 정답과 같습니다).' }
    else { Write-SetupLine "자체 시험: $($test.message)" }

    $code = 0
    foreach ($source in $sources) {
        $info = Read-SetupJson (Join-Path $source $script:JobInfoFile)
        $number = Get-FolderNumber $source
        $prefix = '{0:D2}_' -f $number
        $target = @(Get-PortableJobFolders $Work | Where-Object { ([IO.Path]::GetFileName($_)).StartsWith($prefix, [StringComparison]::Ordinal) }) | Select-Object -First 1
        if ($null -eq $target) { $target = Join-Path (Join-Path $Work $script:JobsDir) ([IO.Path]::GetFileName($source)) }
        $name = [IO.Path]::GetFileName($target)
        if ([IO.File]::Exists((Join-Path $target $script:SettingsFile))) {
            # 이미 있는 작업: 규칙만 새것으로. 입력 · 출력 폴더 · 처리기록 · 장부는 그대로(실행과 겹치지 않게 잠그고)
            try { $lock = Enter-JobLock $target $name } catch {
                Write-SetupLine "${name}: 지금 실행 중이라 바꾸지 못했습니다. 실행이 끝난 뒤 준비.bat 을 다시 두 번 클릭하세요."
                $code = 1
                continue
            }
            try {
                $job = Get-SetupJob $target $Work
                Copy-JobRules $source $target
                Write-SetupLine "${name}: 규칙을 새것으로 바꿨습니다(입력 · 출력 폴더 · 처리기록 · 장부는 그대로)."
                Set-JobAutomation $job $RegisterTask $StartupDir $Desktop
            } catch {
                Write-SetupLine "${name}: 바꾸지 못했습니다 ($($_.Exception.Message)). 대표에게 알려 주세요."
                $code = 1
            } finally {
                Exit-JobLock $lock
            }
            continue
        }
        $title = [string]$info.title
        Write-SetupLine "${name}: 새 작업입니다. 입력 폴더와 출력 폴더를 고르세요."
        $inputDir = Select-LocalFolder $FolderPicker "$title - 입력 폴더(PO 가 들어오는 폴더)를 고르세요"
        if (-not $inputDir) { return Stop-Setup '폴더를 고르지 않아 준비를 멈췄습니다. 준비.bat 을 다시 두 번 클릭하세요.' }
        $outputDir = Select-LocalFolder $FolderPicker "$title - 출력 폴더(회사 양식 · 누적장부를 둘 폴더)를 고르세요"
        if (-not $outputDir) { return Stop-Setup '폴더를 고르지 않아 준비를 멈췄습니다. 준비.bat 을 다시 두 번 클릭하세요.' }
        [void][IO.Directory]::CreateDirectory($target)
        Copy-JobRules $source $target
        $settings = [ordered]@{
            number = $number; title = $title; input_dir = $inputDir; output_dir = $outputDir
            output_mode = [string]$info.output_mode; time = (Get-SlotTime $number); method = $script:MethodNone; last_run = $null
        }
        [IO.File]::WriteAllText((Join-Path $target $script:SettingsFile), (ConvertTo-Json -InputObject $settings -Depth 20), $script:RunUtf8)
        Write-SetupLine "${name}: 입력 $inputDir · 출력 $outputDir"
        Set-JobAutomation (Get-SetupJob $target $Work) $RegisterTask $StartupDir $Desktop
    }

    $markup = Join-Path $Work $script:MarkupFile
    if (-not [IO.File]::Exists($markup)) {   # 있으면 친구가 채운 것이므로 건드리지 않는다
        try {
            Write-ExampleMarkup $markup
            Write-SetupLine "$($script:MarkupFile) 를 가짜 예시로 만들었습니다. 파트넘버 · 마크업% 를 직접 채우세요: $markup"
        } catch {
            Write-SetupLine "경고: $($script:MarkupFile) 를 만들지 못함 ($($_.Exception.Message))"
        }
    }
    $warningOff = if ($null -ne $script:RunPdfWarningOff) { [bool]$script:RunPdfWarningOff } else { Test-PdfWarningOff }
    if (-not $warningOff) { Write-SetupLine $script:PdfWarningGuide }
    if ($code -eq 0) { Write-SetupLine '준비가 끝났습니다.' } else { Write-SetupLine '준비가 끝났지만 못 한 일이 있습니다(위를 보세요).' }
    Show-SetupDialog ([string]::Join("`n", $script:SetupLines.ToArray())) $(if ($code -eq 0) { $script:MbIconInformation } else { $script:MbIconWarning })
    return $code
}

# ---- 점검(대표 PC) -------------------------------------------------------------------------------

function Format-CheckCell($Text) { return ([string]$Text).Replace('|', '\|') }

function Format-ShownValue($Value) {
    # 지금 판 _show: 빈 값 (비어 있음), 날짜 ISO, 나머지는 글자(금액은 자리 그대로).
    if ($null -eq $Value) { return '(비어 있음)' }
    return [string](ConvertTo-JsonValue $Value)
}

function Write-CheckTable($Result, $Source, $Rules) {
    # 지금 판 _print_checks: 검산 표(실행 때와 같은 검산). 불일치 건수를 돌려준다.
    try {
        $checks = Get-Checks $Result $Source $Rules
    } catch {
        [Console]::Out.WriteLine("`n확인필요: 검산을 하지 못함 ($($_.Exception.Message)). 이대로 실행하면 이 문서는 확인필요로 갑니다.")
        return 0
    }
    [Console]::Out.WriteLine("`n| 검산 | 원본 | 결과 | 일치 |`n|---|---|---|---|")
    $bad = 0
    $lines = 0
    $lineNone = 0
    foreach ($c in $checks) {
        [Console]::Out.WriteLine(("| {0} | {1} | {2} | {3} |" -f (Format-CheckCell $c['item']), (Format-CheckCell $c['source_value']),
                                  (Format-CheckCell $c['result_value']), (Get-CheckMark $c['ok'])))
        if ($c['ok'] -eq $false) { $bad++ }
        if ($script:LineCheck.IsMatch([string]$c['item'])) { $lines++; if ($null -eq $c['ok']) { $lineNone++ } }
    }
    $count = if ($null -ne $Result.table) { $Result.table.Count } else { 0 }
    if ($count -gt 0 -and $lines -gt 0 -and $lineNone -eq $lines) {
        [Console]::Out.WriteLine(("`n경고: 품목 {0}줄의 줄 검산(수량 x 단가 = 금액)이 모두 해당 없음입니다. " -f $count) +
            '단가가 없는 문서가 아니면 규칙.yaml 의 검산: 에 수량열 · 단가열 · 금액열 을 실제 열 이름으로 적으세요.')
    }
    return $bad
}

function Invoke-CheckCommand([string]$Path, [string]$RulesPath) {
    # -Command 점검(대표 PC): 무설치판으로 예시 파일 하나를 읽어 지금 판 check 와 같은 표(필드 · 품목 · 검산)를 콘솔에 찍는다.
    # 파일을 만들거나 옮기지 않는다(오피스는 임시 복사본을 연다). 마크업표 · 고객마스터는 보지 않는다.
    if (-not [IO.File]::Exists($Path)) { [Console]::Out.WriteLine("파일이 없습니다: $Path"); return 1 }
    $suffix = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($suffix -eq '.xls') { [Console]::Out.WriteLine($script:XlsReason); return 0 }
    $rules = Import-PortableRules $RulesPath
    $head = $null
    if ($suffix -eq '.xlsx') {
        $grid = Read-ExcelGrid $Path $script:OfficeTimeoutSec '점검'
        $result = Get-GridExtract $rules $grid
        $source = @{ kind = 'grid'; grid = $grid }
    } else {
        $text = [string](Read-WordText $Path $script:OfficeTimeoutSec '점검').text
        $result = Get-TextExtract $rules $text
        $source = @{ kind = 'text'; text = $text }
        $head = [string]::Join("`n", @(Split-TextLines $text | Select-Object -First 30))
    }
    $say = { param($Line) [Console]::Out.WriteLine($Line) }
    & $say "파일: $([IO.Path]::GetFileName($Path))"
    & $say ("문서종류: " + $(if ($null -ne $result.doc_type) { $result.doc_type } else { '(판별 실패)' }))
    if ($result.form) { & $say "엑셀양식: $($result.form)" }
    if ($null -ne $result.review_reason -and ($null -eq $result.doc_type -or $result.review_reason -eq $script:NoText)) {
        & $say "`n확인필요: $($result.review_reason)"
        if ($head) { & $say "`n--- 앞부분 글자(30줄) ---"; & $say $head }
        return 0
    }
    $doc = @($rules.doc_types) | Where-Object { $_.name -ceq $result.doc_type } | Select-Object -First 1
    & $say "번호필드: $($doc.number_field)"
    & $say "`n| 필드 | 값 | 비고 |`n|---|---|---|"
    foreach ($name in $result.fields.Keys) {
        $value = $result.fields[$name]
        $notes = New-Object System.Collections.Generic.List[string]
        if ($null -eq $value) { $notes.Add('못 찾음') }
        foreach ($w in $result.warnings) { if ($w.Contains("'$name' 값")) { $notes.Add($w) } }
        & $say ("| {0} | {1} | {2} |" -f (Format-CheckCell $name), (Format-CheckCell (Format-ShownValue $value)), (Format-CheckCell ([string]::Join(', ', $notes.ToArray()))))
    }
    $columns = @()
    if ($result.form) {
        $form = @($rules.excel_forms) | Where-Object { $_.name -ceq $result.form } | Select-Object -First 1
        if ($null -ne $form.items) { $columns = @($form.items.columns.PSObject.Properties | ForEach-Object { $_.Name }) }
    } elseif ($null -ne $doc.table) {
        $columns = @($doc.table.columns)
    }
    if ($columns.Count -gt 0) {
        if ($null -eq $result.table) {
            & $say "`n품목 표: 읽지 못함 (아래 경고 참고)"
        } else {
            & $say "`n품목 $($result.table.Count)줄"
            & $say ('| ' + [string]::Join(' | ', $columns) + ' |')
            & $say ('|' + ('---|' * $columns.Count))
            foreach ($row in $result.table) {
                $cells = foreach ($c in $columns) { if ($row.Contains($c)) { Format-CheckCell (Format-ShownValue $row[$c]) } else { '' } }
                & $say ('| ' + [string]::Join(' | ', @($cells)) + ' |')
            }
        }
    }
    foreach ($w in $result.warnings) { if (-not $w.Contains("' 못 찾음") -and -not $w.Contains("' 값 ")) { & $say "경고: $w" } }
    if ($null -ne $result.review_reason) { & $say "`n확인필요: $($result.review_reason)"; return 0 }
    $bad = Write-CheckTable $result $source $rules
    if ($bad) { & $say "`n확인필요: 검산 불일치 ${bad}건. 이대로 실행하면 이 문서는 결과를 내지 않고 확인필요로 갑니다." }
    $company = $null
    if ($null -ne $rules.PSObject.Properties['company_form']) { $company = $rules.company_form }
    if ($null -ne $company) {
        if (Test-FormUsesMarkup $company) {
            & $say "`n마크업: 이 점검은 마크업표를 보지 않습니다(친구 PC 에서 실행할 때 그 PC 의 마크업표.xlsx 로 계산하고 검산합니다)."
        }
        $count = if ($null -ne $result.table) { $result.table.Count } else { 0 }
        & $say "회사 양식: $(Get-CompanyFormName $result $company) (${count}줄)"
    }
    return 0
}

# ---- 엔진 op (시험용) -----------------------------------------------------------------------------

function Invoke-SetupOp($Payload) {
    # 엔진 op 'setup'(시험용 입구, 진짜 예약 작업 · 바탕화면을 건드리지 않는다). action:
    #   schedule      {folder, work, whatif?, task: ok|fail, startup_dir}  → {text} (whatif) 또는 {method, notes}
    #   shortcut      {folder, work, desktop}                              → {path}
    #   read_shortcut {path}                                               → {target, arguments, workdir}
    #   self_test     {bundle, row_tolerance?, pdf_warning_off?}           → Invoke-SelfTest 결과
    $get = { param($Name) $p = $Payload.PSObject.Properties[$Name]; if ($null -ne $p) { $p.Value } else { $null } }
    switch ([string]$Payload.action) {
        'schedule' {
            $job = Get-SetupJob ([string]$Payload.folder) ([string]$Payload.work)
            if (& $get 'whatif') { return ConvertTo-Json -InputObject @{ text = (Register-PortableSchedule $job -WhatIf) } -Depth 20 -Compress }
            if ([string](& $get 'task') -ceq 'ok') { $register = { param($Text, $TaskPath, $TaskName) $true } }
            elseif ([string](& $get 'task') -ceq 'fail') { $register = { param($Text, $TaskPath, $TaskName) throw '시험: 작업 스케줄러 막힘' } }
            else { throw '시험에서는 task 를 ok 또는 fail 로 주세요(진짜 작업 스케줄러에 등록하지 않음).' }
            $startup = [string](& $get 'startup_dir')
            if (-not $startup) { throw '시험에서는 startup_dir 을 주세요.' }
            $notes = New-Object System.Collections.Generic.List[string]
            $method = Register-PortableSchedule $job -RegisterTask $register -StartupDir $startup -Notes $notes
            return ConvertTo-Json -InputObject @{ method = $method; notes = [object[]]$notes.ToArray() } -Depth 20 -Compress
        }
        'shortcut' {
            $job = Get-SetupJob ([string]$Payload.folder) ([string]$Payload.work)
            return ConvertTo-Json -InputObject @{ path = (New-RunShortcut $job ([string]$Payload.desktop)) } -Depth 20 -Compress
        }
        'read_shortcut' { return ConvertTo-Json -InputObject (Read-Shortcut ([string]$Payload.path)) -Depth 20 -Compress }
        'self_test' {
            $tolerance = & $get 'row_tolerance'
            if ($null -ne $tolerance) { $script:WordRowTolerance = [double]$tolerance }
            $warningOff = & $get 'pdf_warning_off'
            if ($null -ne $warningOff) { $script:RunPdfWarningOff = [bool]$warningOff }
            return ConvertTo-Json -InputObject (Invoke-SelfTest ([string]$Payload.bundle)) -Depth 20 -Compress
        }
    }
    throw "setup action '$($Payload.action)' 는 없습니다."
}
