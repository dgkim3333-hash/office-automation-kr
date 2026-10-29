# 오피스로 읽기: 워드로 PDF 글자(Read-WordText) · 엑셀로 칸 격자(Read-ExcelGrid)를 읽는다.
# 멈춤 대비(설계 6절): 실행마다 새 이름의 임시 폴더에 복사본을 만들어 열고(원본은 그대로), 오피스 일은 따로 돌린 뒤
# 시간 제한($script:OfficeTimeoutSec)으로 묶는다. 넘으면 이 실행이 COM 으로 띄운 워드 · 엑셀(명령줄 /Automation)만 끝내고
# [System.TimeoutException](= OfficeTimeout, 문구 「워드가 응답하지 않음」 · 「엑셀이 응답하지 않음」)을 던진다.
# 사용자가 직접 연 워드 · 엑셀은 건드리지 않는다. 레지스트리는 읽기만 한다(Test-PdfWarningOff).

$script:OfficeTimeoutSec = 120        # 파일 하나에 주는 시간(초)
$script:OfficeExitWaitSec = 10        # 끝낸 오피스가 사라질 때까지 기다리는 시간(초)
$script:WordTimeoutText = '워드가 응답하지 않음'
$script:ExcelTimeoutText = '엑셀이 응답하지 않음'
$script:OfficeForced = 0             # 마지막 읽기에서 Quit 뒤 스스로 끝나지 않아 강제로 끝낸 오피스 수(0 이 정상)
$script:WordRowTolerance = 3.0        # 표 한 행 안에서 같은 줄로 보는 높이 차(포인트, pdfplumber 의 y_tolerance 와 같음)

# 엑셀 오류 칸(.Value2 가 정수 오류 번호)의 글자. openpyxl 은 이 칸을 오류 글자('#N/A' 등)로 읽는다.
$script:ExcelErrorText = @{
    -2146826281 = '#DIV/0!'; -2146826246 = '#N/A'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!'
    -2146826252 = '#NUM!'; -2146826265 = '#REF!'; -2146826273 = '#VALUE!'
}

# ---- 프로세스 · 설정 ----------------------------------------------------------------------------

function Get-OfficeProcesses([string]$Name, $Known = $null, [switch]$IdsOnly) {
    # 지금 떠 있는 그 이름(WINWORD.EXE · EXCEL.EXE)의 프로세스: {pid, commandline}. 결과는 언제나 배열.
    # 명령줄은 Win32_Process 로 읽는데 한 번에 몇 초 걸리므로, $Known(이전 목록)에 있는 PID 는 읽지 않고 $null 로 둔다.
    # -IdsOnly 면 명령줄을 하나도 읽지 않는다(실행 전 목록은 PID 만 있으면 된다).
    $skip = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($p in @($Known)) { if ($null -ne $p) { [void]$skip.Add([int]$p.pid) } }
    $ids = New-Object 'System.Collections.Generic.List[int]'
    foreach ($p in @(Get-Process -Name ([IO.Path]::GetFileNameWithoutExtension($Name)) -ErrorAction SilentlyContinue)) { $ids.Add([int]$p.Id) }
    $lines = @{}
    $ask = @($ids | Where-Object { -not $skip.Contains($_) })
    if ($ask.Count -gt 0 -and -not $IdsOnly) {
        $filter = [string]::Join(' OR ', @($ask | ForEach-Object { 'ProcessId=' + $_ }))
        foreach ($c in @(Get-CimInstance Win32_Process -Filter $filter)) { $lines[[int]$c.ProcessId] = $c.CommandLine }
    }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($id in $ids) { $list.Add([pscustomobject]@{ pid = $id; commandline = $lines[$id] }) }
    return ,$list.ToArray()
}

function Select-OwnOfficePids($Before, $After) {
    # 실행 전에 없던 PID 이고 명령줄에 /Automation 이 있는 것만(COM 이 띄운 오피스). 사용자가 연 오피스는 고르지 않는다.
    $known = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($p in @($Before)) { if ($null -ne $p) { [void]$known.Add([int]$p.pid) } }
    $own = New-Object 'System.Collections.Generic.List[int]'
    foreach ($p in @($After)) {
        if ($null -eq $p -or $known.Contains([int]$p.pid)) { continue }
        $line = $p.commandline
        if ($null -ne $line -and ([string]$line).IndexOf('/Automation', [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $own.Add([int]$p.pid)
        }
    }
    return ,$own.ToArray()
}

function Test-ProcessAlive([int]$ProcessId) {
    return $null -ne (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)
}

function Stop-OwnOffice($Before, [string]$Name, [int]$GraceSec = 0) {
    # 이 실행이 띄운 오피스만 끝낸다: 먼저 스스로 끝나기를 GraceSec 초까지 기다리고(Quit 뒤 잠깐 남는 일),
    # 남은 것은 강제로 끝낸 뒤 실제로 사라질 때까지 기다린다(최대 $script:OfficeExitWaitSec 초). 강제로 끝낸 PID 목록을 돌려준다.
    $own = Select-OwnOfficePids $Before (Get-OfficeProcesses $Name $Before)
    $deadline = [DateTime]::UtcNow.AddSeconds($GraceSec)
    $alive = @($own | Where-Object { Test-ProcessAlive $_ })
    while ($alive.Count -gt 0 -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 200
        $alive = @($alive | Where-Object { Test-ProcessAlive $_ })
    }
    foreach ($id in $alive) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
    $deadline = [DateTime]::UtcNow.AddSeconds($script:OfficeExitWaitSec)
    while ([DateTime]::UtcNow -lt $deadline -and @($alive | Where-Object { Test-ProcessAlive $_ }).Count -gt 0) {
        Start-Sleep -Milliseconds 200
    }
    return ,[int[]]$alive
}

function Get-OfficeVersion([string]$ProgId) {
    # HKCR\<ProgId>\CurVer = 'Word.Application.16' → '16.0'. 없으면 $null.
    $key = 'Registry::HKEY_CLASSES_ROOT\' + $ProgId + '\CurVer'
    $item = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
    if ($null -eq $item) { return $null }
    $m = [regex]::Match([string]$item.'(default)', '\.([0-9]+)$')
    if (-not $m.Success) { return $null }
    return $m.Groups[1].Value + '.0'
}

function Test-PdfWarningOff {
    # 워드의 「PDF 를 편집 가능한 문서로 변환합니다」 안내가 꺼져 있는지(DisableConvertPdfWarning = 1). 읽기만 한다.
    $version = Get-OfficeVersion 'Word.Application'
    if ($null -eq $version) { return $false }
    $item = Get-ItemProperty -LiteralPath ('HKCU:\Software\Microsoft\Office\' + $version + '\Word\Options') -ErrorAction SilentlyContinue
    if ($null -eq $item -or $null -eq $item.PSObject.Properties['DisableConvertPdfWarning']) { return $false }
    return ([int]$item.DisableConvertPdfWarning -eq 1)
}

# ---- 임시 복사본 --------------------------------------------------------------------------------

function New-OfficeCopy([string]$Path, [string]$Job) {
    # %TEMP%\무설치판_<작업>_<날짜시각>_<임의>\ 에 복사본을 만든다(이름은 그대로). 복사본의 인터넷 표시(Zone.Identifier)만 지운다.
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "파일이 없습니다: $Path" }
    $source = (Resolve-Path -LiteralPath $Path).ProviderPath   # .NET 복사는 PowerShell 의 현재 폴더를 모른다
    $base = $env:TEMP
    if ([string]::IsNullOrEmpty($base)) { $base = [IO.Path]::GetTempPath() }
    $safe = $Job
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $safe = $safe.Replace([string]$c, '_') }
    $random = [Guid]::NewGuid().ToString('N').Substring(0, 8)
    $folder = Join-Path $base ('무설치판_{0}_{1}_{2}' -f $safe, (Get-Date).ToString('yyyyMMdd-HHmmss', [Globalization.CultureInfo]::InvariantCulture), $random)
    [void][IO.Directory]::CreateDirectory($folder)
    $copy = Join-Path $folder ([IO.Path]::GetFileName($source))
    [IO.File]::Copy($source, $copy)
    Unblock-File -LiteralPath $copy
    return @{ Folder = $folder; Path = $copy }
}

function Remove-OfficeCopy($Copy) {
    # 임시 폴더를 지운다. 끝낸 오피스가 파일을 막 놓은 참이면 잠깐 기다렸다 다시 해 본다(최대 약 5초).
    if ($null -eq $Copy) { return }
    for ($i = 0; $i -lt 25; $i++) {
        if (-not (Test-Path -LiteralPath $Copy.Folder)) { return }
        try { Remove-Item -LiteralPath $Copy.Folder -Recurse -Force -ErrorAction Stop; return } catch { Start-Sleep -Milliseconds 200 }
    }
}

# ---- 따로 돌리기 · 시간 제한 ---------------------------------------------------------------------

function Invoke-OfficeWork([scriptblock]$Work, [string]$Path, [string]$ProcessName, [int]$TimeoutSec, [string]$TimeoutText, [string]$Label) {
    # 오피스 일($Work, 매개변수 하나 = 복사본 경로)을 새 runspace 에서 돌리고 TimeoutSec 초까지 기다린다.
    # 끝나면 그 결과를, 넘으면 이 실행이 띄운 오피스를 끝낸 뒤 TimeoutException 을 던진다. 어느 쪽이든 띄운 오피스는 남기지 않는다.
    $script:OfficeForced = 0
    $before = Get-OfficeProcesses $ProcessName -IdsOnly
    $runspace = [runspacefactory]::CreateRunspace()
    # MTA: STA 로 돌리면 끝날 때 COM 참조 정리(GC 종료자)가 그 스레드를 기다리며 수십 초 멈춘다(이 PC 엑셀에서 24초)
    $runspace.ApartmentState = [Threading.ApartmentState]::MTA
    $runspace.Open()
    $shell = [powershell]::Create()
    $shell.Runspace = $runspace
    [void]$shell.AddScript($Work.ToString()).AddArgument($Path)
    $handle = $shell.BeginInvoke()
    $done = $handle.AsyncWaitHandle.WaitOne([Math]::Max(0, $TimeoutSec) * 1000)
    if (-not $done) {
        # 멈춘 오피스를 끝내면 막혀 있던 COM 호출이 오류로 돌아와 일이 끝난다. 일이 끝날 때까지(최대 OfficeExitWaitSec) 새로 뜬 것도 끝낸다
        [void]$shell.BeginStop($null, $null)
        $deadline = [DateTime]::UtcNow.AddSeconds($script:OfficeExitWaitSec)
        do {
            [void](Stop-OwnOffice $before $ProcessName)
            if ($handle.IsCompleted) { break }
            Start-Sleep -Milliseconds 200
        } while ([DateTime]::UtcNow -lt $deadline)
        [void](Stop-OwnOffice $before $ProcessName)
        if ($handle.IsCompleted) { $shell.Dispose(); $runspace.Dispose() }   # 끝나지 않은 일은 버려 둔다(프로세스가 끝나면 함께 사라짐)
        throw (New-Object System.TimeoutException $TimeoutText)
    }
    try {
        $output = $shell.EndInvoke($handle)
        $failure = $null
        if ($shell.Streams.Error.Count -gt 0) { $failure = $shell.Streams.Error[0].ToString() }
    } catch {
        $inner = $_.Exception
        while ($null -ne $inner.InnerException) { $inner = $inner.InnerException }
        $failure = $inner.Message
    } finally {
        $shell.Dispose()
        $runspace.Dispose()
        # runspace 의 변수가 쥐고 있던 COM 참조를 놓아야 Quit 한 오피스(특히 엑셀)가 끝난다
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        # Quit 뒤 잠깐 남는 오피스는 스스로 끝나기를 기다리고, 그래도 남으면 끝낸다(그 수를 $script:OfficeForced 에 남김)
        $forced = Stop-OwnOffice $before $ProcessName $script:OfficeExitWaitSec
        $script:OfficeForced = $forced.Count
    }
    if ($null -ne $failure) { throw ('{0}로 읽지 못함: {1}' -f $Label, $failure) }
    if ($output.Count -eq 0) { throw ('{0}에서 아무것도 받지 못했습니다.' -f $Label) }
    return $output[0]
}

# ---- 워드 ---------------------------------------------------------------------------------------

# 새 runspace 에서 도는 워드 일(이 파일의 함수는 쓸 수 없다). 결과: {pages, blocks}.
#   blocks: 문서 순서대로 @{kind='p'; text; list} 또는 표 한 행 @{kind='row'; cells=@(@{text; list; page; y; x} …)}(행의 문단 전부).
$script:WordWork = {
    param([string]$Path)
    $ErrorActionPreference = 'Stop'
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0                 # wdAlertsNone
        $word.AutomationSecurity = 3            # msoAutomationSecurityForceDisable (매크로 끔, 이 워드에만)
        # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        $doc = $word.Documents.Open($Path, $false, $true, $false)
        $pages = [int]$doc.ComputeStatistics(2)   # wdStatisticPages
        # 표가 차지하는 구간과 그 행들. 칸 안 문단이 여러 개인 행만 문단마다 쪽 · 높이 · 왼쪽 위치를 문단 첫 글자 자리에서 읽는다
        # (위치 읽기는 느리고, 칸마다 문단이 하나인 행은 칸 순서대로 한 줄이면 된다. 위치를 안 읽은 문단은 y = -1).
        $tables = New-Object System.Collections.Generic.List[object]
        foreach ($table in $doc.Tables) {
            $rows = New-Object 'System.Collections.Generic.SortedDictionary[int,object]'
            foreach ($cell in $table.Range.Cells) {
                $index = [int]$cell.RowIndex
                if (-not $rows.ContainsKey($index)) { $rows[$index] = New-Object System.Collections.Generic.List[object] }
                $ranges = New-Object System.Collections.Generic.List[object]
                foreach ($p in $cell.Range.Paragraphs) { $ranges.Add($p.Range) }
                $rows[$index].Add($ranges)
            }
            $done = New-Object 'System.Collections.Generic.SortedDictionary[int,object]'
            foreach ($index in $rows.Keys) {
                $multi = @($rows[$index] | Where-Object { $_.Count -gt 1 }).Count -gt 0
                $cells = New-Object System.Collections.Generic.List[object]
                foreach ($ranges in $rows[$index]) {
                    foreach ($range in $ranges) {
                        $entry = @{ text = [string]$range.Text; list = [string]$range.ListFormat.ListString; page = 0; y = [double]-1; x = [double]-1 }
                        if ($multi) {
                            $start = $range.Duplicate
                            [void]$start.Collapse(1)            # wdCollapseStart
                            $entry.page = [int]$start.Information(3)      # wdActiveEndPageNumber
                            $entry.y = [double]$start.Information(6)      # wdVerticalPositionRelativeToPage
                            $entry.x = [double]$start.Information(5)      # wdHorizontalPositionRelativeToPage
                        }
                        $cells.Add($entry)
                    }
                }
                $done[$index] = $cells.ToArray()
            }
            $tables.Add(@{ Start = [int]$table.Range.Start; End = [int]$table.Range.End; Rows = $done })
        }
        $blocks = New-Object System.Collections.Generic.List[object]
        foreach ($p in $doc.Paragraphs) {
            $range = $p.Range
            $at = [int]$range.Start
            $inside = $null
            foreach ($t in $tables) { if ($at -ge $t.Start -and $at -lt $t.End) { $inside = $t; break } }
            if ($null -ne $inside) {
                if (-not $inside.ContainsKey('Done')) {
                    foreach ($row in $inside.Rows.Values) { $blocks.Add(@{ kind = 'row'; cells = $row }) }
                    $inside['Done'] = $true
                }
                continue
            }
            $blocks.Add(@{ kind = 'p'; text = [string]$range.Text; list = [string]$range.ListFormat.ListString })
        }
        [pscustomobject]@{ pages = $pages; blocks = $blocks.ToArray() }
    } finally {
        if ($null -ne $doc) { try { $doc.Close(0) } catch { } ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        if ($null -ne $word) { try { $word.Quit() } catch { } ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function ConvertTo-WordLine([string]$Text) {
    # 문단 글자 한 조각 → 한 줄: 칸 끝(\a) · 문단 끝(\r) 표시를 빼고, 워드 특수 하이픈(30 → '-', 31 은 지움),
    # 탭 · 연속 공백을 한 칸으로, 앞뒤 공백을 지운다.
    $line = $Text.Replace([string][char]7, '').Replace("`r", ' ').Replace([string][char]30, '-').Replace([string][char]31, '')
    return ([regex]::Replace($line, '[ \t]+', ' ')).Trim()
}

function Add-WordParagraph($Lines, [string]$Text, [string]$List) {
    # 표 밖 문단: 목록 번호가 있으면 앞에 붙이고 한 칸, 줄바꿈(\v) · 쪽 나눔(\f)에서 줄을 나눈다.
    $text = $Text
    if (-not [string]::IsNullOrEmpty($List)) { $text = $List + ' ' + $text }
    foreach ($piece in ($text -split "[`v`f`n]")) { $Lines.Add((ConvertTo-WordLine $piece)) }
}

function Add-WordRow($Lines, $Cells) {
    # 표 한 행: 칸 안 문단을 높이(쪽 → y)가 3 포인트 안인 것끼리 묶어 왼쪽(x)부터 한 칸 띄어 한 줄로(pdfplumber 의 줄 만들기와 같은 생각).
    # 위치를 못 읽은 문단(y 가 음수)이 있으면 행 전체를 칸 순서대로 한 줄로.
    $items = New-Object System.Collections.Generic.List[object]
    $positioned = $true
    $order = 0
    foreach ($cell in @($Cells)) {
        $text = [string]$cell.text
        if (-not [string]::IsNullOrEmpty([string]$cell.list)) { $text = [string]$cell.list + ' ' + $text }
        $text = ConvertTo-WordLine ($text -replace "[`v`f`n]", ' ')
        if ([double]$cell.y -lt 0) { $positioned = $false }
        if ($text -ne '') {
            $items.Add([pscustomobject]@{ text = $text; page = [int]$cell.page; y = [double]$cell.y; x = [double]$cell.x; order = $order })
        }
        $order++
    }
    if ($items.Count -eq 0) { return }
    if (-not $positioned) {
        $Lines.Add([string]::Join(' ', @($items | ForEach-Object { $_.text })))
        return
    }
    $sorted = @($items | Sort-Object page, y, order)
    $group = New-Object System.Collections.Generic.List[object]
    $last = $null
    foreach ($item in $sorted) {
        if ($null -ne $last -and ($item.page -ne $last.page -or ($item.y - $last.y) -gt $script:WordRowTolerance)) {
            $Lines.Add([string]::Join(' ', @($group | Sort-Object x, order | ForEach-Object { $_.text })))
            $group.Clear()
        }
        $group.Add($item)
        $last = $item
    }
    $Lines.Add([string]::Join(' ', @($group | Sort-Object x, order | ForEach-Object { $_.text })))
}

function ConvertTo-WordText($Blocks) {
    # 워드 글자 정규화(설계 4-3): 목록 번호 되살리기 · 탭과 연속 공백 → 한 칸 · 앞뒤 공백 지우기 · 가로줄 흔적('/' 한 글자 줄)과
    # 빈 줄 지우기 · 표는 행마다 한 줄(칸 안 문단이 여러 줄이면 높이가 같은 것끼리). 줄은 \n 으로 잇는다.
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($block in @($Blocks)) {
        if ($null -eq $block) { continue }
        if ([string]$block.kind -ceq 'row') { Add-WordRow $lines $block.cells }
        else { Add-WordParagraph $lines ([string]$block.text) ([string]$block.list) }
    }
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) { if ($line -ne '' -and $line -ne '/') { $kept.Add($line) } }
    return [string]::Join("`n", $kept.ToArray())
}

function Read-WordText([string]$Path, [int]$TimeoutSec = $script:OfficeTimeoutSec, [string]$Job = '워드', [scriptblock]$Work = $script:WordWork) {
    # 워드로 PDF(또는 워드 문서)를 읽어 정규화한 글자와 쪽 수: @{ text; pages }. 시간 제한을 넘으면 TimeoutException(워드가 응답하지 않음).
    # $Work 는 시험에서 멈춘 워드를 흉내 낼 때만 바꾼다.
    $copy = New-OfficeCopy $Path $Job
    try {
        $raw = Invoke-OfficeWork $Work $copy.Path 'WINWORD.EXE' $TimeoutSec $script:WordTimeoutText '워드'
    } finally {
        Remove-OfficeCopy $copy
    }
    return @{ text = (ConvertTo-WordText $raw.blocks); pages = [int]$raw.pages }
}

# ---- 엑셀 ---------------------------------------------------------------------------------------

# 새 runspace 에서 도는 엑셀 일. 결과: {active, sheets = @(@{name; row; col; rows; cols; value2 = object[,] 또는 값 하나; value; formats})}.
#   value2 · value 는 UsedRange 통째(.Value2 · .Value(10)), formats 는 값이 있는 칸의 '행,열' → .NumberFormat.
#   엑셀은 COM 참조가 하나라도 남으면 Quit 해도 끝나지 않으므로, 손에 쥔 COM 객체를 모두 $held 에 모아 끝에 놓는다.
$script:ExcelWork = {
    param([string]$Path)
    $ErrorActionPreference = 'Stop'
    $held = New-Object System.Collections.Generic.List[object]
    function Hold($Object) { if ($null -ne $Object) { $held.Add($Object) }; return ,$Object }   # , : 컬렉션(Range 등)이 풀리지 않게
    $excel = $null
    $book = $null
    try {
        $excel = Hold (New-Object -ComObject Excel.Application)
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3           # msoAutomationSecurityForceDisable (매크로 끔, 이 엑셀에만)
        $books = Hold $excel.Workbooks
        # Open(Filename, UpdateLinks=0, ReadOnly=$true)
        $book = Hold $books.Open($Path, 0, $true)
        $activeName = [string](Hold $book.ActiveSheet).Name
        $active = 0
        $sheets = New-Object System.Collections.Generic.List[object]
        $list = Hold $book.Worksheets
        for ($i = 1; $i -le [int]$list.Count; $i++) {
            $sheet = Hold $list.Item($i)
            if ([string]$sheet.Name -ceq $activeName) { $active = $sheets.Count }
            $used = Hold $sheet.UsedRange
            $rows = [int](Hold $used.Rows).Count
            $cols = [int](Hold $used.Columns).Count
            $cells = Hold $used.Cells
            $value2 = $used.Value2
            $value = $used.Value(10)                 # xlRangeValueDefault: 날짜 칸은 [datetime]
            $formats = @{}
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $v = if ($rows -eq 1 -and $cols -eq 1) { $value2 } else { $value2[$r, $c] }
                    if ($null -ne $v) { $formats["$r,$c"] = [string](Hold $cells.Item($r, $c)).NumberFormat }
                }
            }
            $sheets.Add(@{ name = [string]$sheet.Name; row = [int]$used.Row; col = [int]$used.Column; rows = $rows; cols = $cols
                           value2 = $value2; value = $value; formats = $formats })
        }
        [pscustomobject]@{ active = $active; sheets = $sheets.ToArray() }
    } finally {
        if ($null -ne $book) { try { $book.Close($false) } catch { } }
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        for ($i = $held.Count - 1; $i -ge 0; $i--) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($held[$i]) } catch { }
        }
        $held.Clear()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function ConvertTo-ExcelGridCell($Value2, $Value, [string]$Format) {
    # 엑셀 칸 하나 → 격자 칸 {t, v, f}(작업 3 모양, openpyxl 과 같게). $null 이면 빈 칸.
    #   날짜: .Value 가 [datetime] → d, v = 'YYYY-MM-DDTHH:MM:SS'(초 아래는 버림)
    #   숫자: .Value2 [double]. 정수 값이면 [long](openpyxl 이 정수로 읽는 칸과 같게, 엔진 글자가 '11.0' 이 아니라 '11')
    #   오류 칸(.Value2 가 [int]): openpyxl 처럼 오류 글자 s
    if ($null -eq $Value2) { return $null }
    if ($Value -is [datetime]) {
        return [pscustomobject]@{ t = 'd'; v = $Value.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture); f = $Format }
    }
    if ($Value2 -is [bool]) { return [pscustomobject]@{ t = 'b'; v = $Value2; f = $Format } }
    if ($Value2 -is [string]) { return [pscustomobject]@{ t = 's'; v = $Value2; f = $Format } }
    if ($Value2 -is [int]) {
        $text = $script:ExcelErrorText[[int]$Value2]
        if ($null -eq $text) { $text = '#ERROR!' }
        return [pscustomobject]@{ t = 's'; v = $text; f = $Format }
    }
    $number = [double]$Value2
    if ([Math]::Floor($number) -eq $number -and [Math]::Abs($number) -lt 1e15) {
        return [pscustomobject]@{ t = 'n'; v = [long]$number; f = $Format }
    }
    return [pscustomobject]@{ t = 'n'; v = $number; f = $Format }
}

function ConvertTo-ExcelGrid($Raw) {
    # 엑셀 일의 결과 → 격자 PSCustomObject {active, sheets = @({name, cells = @(@(행, 열, 칸) …)})}. 칸은 행 → 열 순서.
    $sheets = New-Object System.Collections.Generic.List[object]
    foreach ($s in @($Raw.sheets)) {
        $cells = New-Object System.Collections.Generic.List[object]
        $single = ($s.rows -eq 1 -and $s.cols -eq 1)
        for ($r = 1; $r -le $s.rows; $r++) {
            for ($c = 1; $c -le $s.cols; $c++) {
                if ($single) { $v2 = $s.value2; $v = $s.value } else { $v2 = $s.value2[$r, $c]; $v = $s.value[$r, $c] }
                $cell = ConvertTo-ExcelGridCell $v2 $v ([string]$s.formats["$r,$c"])
                if ($null -ne $cell) { $cells.Add([object[]]@(($s.row + $r - 1), ($s.col + $c - 1), $cell)) }
            }
        }
        $sheets.Add([pscustomobject]@{ name = $s.name; cells = $cells.ToArray() })
    }
    return [pscustomobject]@{ active = [int]$Raw.active; sheets = $sheets.ToArray() }
}

function Read-ExcelGrid([string]$Path, [int]$TimeoutSec = $script:OfficeTimeoutSec, [string]$Job = '엑셀', [scriptblock]$Work = $script:ExcelWork) {
    # 엑셀로 통합 문서를 읽어 격자(작업 3 모양, PSCustomObject)로. 시간 제한을 넘으면 TimeoutException(엑셀이 응답하지 않음).
    # $Work 는 시험에서 멈춘 엑셀을 흉내 낼 때만 바꾼다.
    $copy = New-OfficeCopy $Path $Job
    try {
        $raw = Invoke-OfficeWork $Work $copy.Path 'EXCEL.EXE' $TimeoutSec $script:ExcelTimeoutText '엑셀'
    } finally {
        Remove-OfficeCopy $copy
    }
    return ConvertTo-ExcelGrid $raw
}

# ---- 엔진 op ------------------------------------------------------------------------------------

# 시험용: 파일을 연 채로 멈춘 오피스. 멈춤 요청을 받으면 일은 끝나지만 문서를 닫거나 Quit 하지 않으므로, 오피스를 끝내는 길은 프로세스를 끝내는 것뿐이다.
$script:StuckWork = @{
    word = {
        param([string]$Path)
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($Path, $false, $true, $false)
        while ($true) { Start-Sleep -Milliseconds 200 }
    }
    excel = {
        param([string]$Path)
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $book = $excel.Workbooks.Open($Path, 0, $true)
        while ($true) { Start-Sleep -Milliseconds 200 }
    }
}

function Invoke-OfficeOp($Payload) {
    # 엔진 op 'office'(시험용 입구). action:
    #   select_pids     {before, after: [{pid, commandline}]}  → [pid, ...]
    #   normalize_word  {blocks}                               → {text}
    #   read_word       {path, timeout?}                       → {text, pages, forced}
    #   read_excel      {path, timeout?}                       → {grid, forced}   (forced = Quit 뒤 강제로 끝낸 오피스 수)
    #   pdf_warning_off {}                                     → {off}
    #   excel_markup    {path, timeout?}                       → 엑셀로 읽은 격자(메모리 그대로)로 마크업표 읽기(markup op 의 read 와 같은 출력)
    #   excel_extract   {path, rules, timeout?}                → 엑셀로 읽은 격자(메모리 그대로)로 extract_grid 결과
    #   new_copy        {path, job}                            → {folder, path}  (임시 복사본만 만들고 지우지 않음)
    #   stop_own        {name, before: [{pid}], grace}         → {forced: [pid, ...]}
    #   stuck           {app: word|excel, path, timeout}       → 언제나 시간 제한 오류(멈춘 오피스 흉내)
    #   work_path       {app: word|excel, path}                → {path}  (오피스 일이 받은 경로. 오피스는 띄우지 않음)
    $timeout = $script:OfficeTimeoutSec
    if ($null -ne $Payload.PSObject.Properties['timeout'] -and $null -ne $Payload.timeout) { $timeout = [int]$Payload.timeout }
    switch ([string]$Payload.action) {
        'select_pids' {
            $own = Select-OwnOfficePids @($Payload.before) @($Payload.after)
            return ConvertTo-Json -InputObject $own -Depth 20 -Compress
        }
        'normalize_word' { return ConvertTo-Json -InputObject @{ text = (ConvertTo-WordText $Payload.blocks) } -Depth 20 -Compress }
        'read_word' {
            $read = Read-WordText ([string]$Payload.path) $timeout
            return ConvertTo-Json -InputObject @{ text = $read.text; pages = $read.pages; forced = $script:OfficeForced } -Depth 20 -Compress
        }
        'read_excel' {
            $grid = Read-ExcelGrid ([string]$Payload.path) $timeout
            return ConvertTo-Json -InputObject @{ grid = $grid; forced = $script:OfficeForced } -Depth 20 -Compress
        }
        'pdf_warning_off' { return ConvertTo-Json -InputObject @{ off = (Test-PdfWarningOff) } -Depth 20 -Compress }
        'excel_markup' {
            $grid = Read-ExcelGrid ([string]$Payload.path) $timeout
            return Invoke-MarkupOp ([pscustomobject]@{ action = 'read'; grid = $grid })
        }
        'excel_extract' {
            $grid = Read-ExcelGrid ([string]$Payload.path) $timeout
            return ConvertTo-ResultJson (Get-GridExtract (ConvertTo-PortableRules $Payload.rules) $grid)
        }
        'new_copy' {
            $copy = New-OfficeCopy ([string]$Payload.path) ([string]$Payload.job)
            return ConvertTo-Json -InputObject @{ folder = $copy.Folder; path = $copy.Path } -Depth 20 -Compress
        }
        'stop_own' {
            $forced = Stop-OwnOffice @($Payload.before) ([string]$Payload.name) ([int]$Payload.grace)
            return ConvertTo-Json -InputObject @{ forced = $forced } -Depth 20 -Compress
        }
        'work_path' {
            # 오피스 대신 받은 경로만 돌려주는 일: 워드 결과는 글자 한 문단, 엑셀 결과는 시트 이름
            if ([string]$Payload.app -ceq 'excel') {
                $echo = { param([string]$Path) [pscustomobject]@{ active = 0; sheets = @(@{ name = $Path; row = 1; col = 1; rows = 0; cols = 0; value2 = $null; value = $null; formats = @{} }) } }
                $given = (Read-ExcelGrid ([string]$Payload.path) $timeout '엑셀' $echo).sheets[0].name
            } else {
                $echo = { param([string]$Path) [pscustomobject]@{ pages = 1; blocks = @(@{ kind = 'p'; text = $Path; list = '' }) } }
                $given = (Read-WordText ([string]$Payload.path) $timeout '워드' $echo).text
            }
            return ConvertTo-Json -InputObject @{ path = $given } -Depth 20 -Compress
        }
        'stuck' {
            if ([string]$Payload.app -ceq 'excel') { [void](Read-ExcelGrid ([string]$Payload.path) $timeout '엑셀' $script:StuckWork.excel) }
            else { [void](Read-WordText ([string]$Payload.path) $timeout '워드' $script:StuckWork.word) }
            throw '멈춘 오피스 흉내가 끝나 버렸습니다.'
        }
    }
    throw "office action '$($Payload.action)' 는 없습니다."
}
