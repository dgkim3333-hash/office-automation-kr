# 오피스로 쓰기: 엑셀 COM 으로 회사 양식(Write-CompanyForm) · 누적장부(Add-LedgerEntry · Add-LedgerLog) · 검산 파일(Write-CheckFile)을 쓰고
# 저장한 파일을 다시 읽어 쓴 값과 견준다(Test-ReadBack). 파일 모양은 지금 판 app\outputs.py · app\ledger.py 와 같다.
# 엑셀 일은 Office.ps1 의 Invoke-OfficeWork 로 따로 돌린다(시간 제한 · 이 실행이 띄운 엑셀만 끝냄 · COM 참조는 모두 놓음).
# 원본은 쓰지 않는다: 새 파일은 같은 폴더의 임시 이름(.<이름>.tmp.xlsx)에 저장한 뒤 다시 읽어 맞을 때만 최종 이름으로 옮기고,
# 장부는 열기만 하고(읽기 전용) 고친 판을 임시 이름에 저장 → 다시 읽기 → 맞으면 바로 전 판을 .bak.xlsx 로 남기고 바꾼다.

$script:LedgerItemsSheet = '품목'
$script:LedgerLogSheet = '처리로그'
$script:LedgerDocHead = @('처리일시', '원본파일')
$script:LedgerDocTail = @('정리된경로')
$script:LedgerItemHead = @('문서종류', '문서번호')
$script:LedgerLogColumns = @('일시', '파일', '결과', '사유')
# 날짜 형식의 '-' 는 따옴표로 감싼다: 감싸지 않으면 한국 Windows 엑셀이 '-' 를 지역 날짜 구분자로 바꿔 저장해
# (yyyy-mm-dd → 기본 형식 14, 일시는 yyyy/mm/dd) 다른 지역 PC 에서 10/14/2026 처럼 보인다. 보이는 모양은 지금 판 yyyy-mm-dd 와 같다.
# 빈칸도 따옴표 안에 둔다(그냥 빈칸은 엑셀이 '\ ' 로 저장). 역슬래시는 쓰지 않는다(한국 지역에서는 원화 기호로 보임).
$script:WriterDateFormat = 'yyyy"-"mm"-"dd'
$script:WriterDateTimeFormat = 'yyyy"-"mm"-"dd" "hh:mm:ss'
$script:WriterAmountFormat = '#,##0.00'
$script:CheckHeader = @('항목', '원본 값', '결과 값', '일치')
$script:CheckSheet = '검산'
$script:CheckFieldsSheet = '추출값'
$script:CheckItemsSheet = '품목'
$script:CheckCompanySheet = '회사양식'
$script:ExcelWriteFailText = '{0}로 쓰지 못함: {1}'
$script:LedgerOpened = $null   # 마지막 장부 쓰기에서 엑셀이 연 파일(시험용 확인, 새 장부면 $null)
$script:ReadBackShown = 5     # 다시 읽기 오류 글자에 보여 줄 문제 수(지금 판 job_run 과 같음)

# ---- 값 → 엑셀 칸 (부모 runspace) ---------------------------------------------------------------

function ConvertTo-ItemArray($Items) {
    # 목록(배열 · List · 낱개 · $null)을 object[] 로. New-Object 로 만든 List[object] 를 @() 로 펼치면 PowerShell 5.1 이
    # 'Argument types do not match'(PSToObjectArrayBinder)로 실패하므로 foreach 로 옮겨 담는다.
    $list = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Items) { return ,$list.ToArray() }
    if ($Items -is [System.Collections.IEnumerable] -and -not ($Items -is [string]) -and -not ($Items -is [System.Collections.IDictionary])) {
        foreach ($item in $Items) { $list.Add($item) }
    } else {
        $list.Add($Items)
    }
    return ,$list.ToArray()
}

function Remove-IllegalText([string]$Text) {
    # openpyxl ILLEGAL_CHARACTERS_RE 와 같은 제어문자(엑셀에 쓸 수 없음)를 지운다.
    return [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
}

function Get-AmountFormat([decimal]$Value) {
    # 지금 판 amount_format: 뒤의 0 을 뺀 소수 자리가 2 보다 많으면 그 자리까지(0.3745 → #,##0.0000), 아니면 #,##0.00.
    $places = (Remove-TrailingZeros $Value).ToString($script:Invariant)
    $dot = $places.IndexOf('.')
    $count = if ($dot -lt 0) { 0 } else { $places.Length - $dot - 1 }
    if ($count -le 2) { return $script:WriterAmountFormat }
    return '#,##0.' + ('0' * $count)
}

function Get-WriterFormat($Value) {
    # 칸 숫자 형식(지금 판 _append · _write_sheet): 일시 · 날짜(엔진 안에서는 같은 [datetime] → 시각이 0 이면 날짜) · 금액([decimal]). 그 밖은 $null(일반).
    if ($Value -is [datetime]) {
        if ($Value.TimeOfDay.Ticks -eq 0) { return $script:WriterDateFormat }
        return $script:WriterDateTimeFormat
    }
    if ($Value -is [decimal]) { return Get-AmountFormat $Value }
    return $null
}

function ConvertTo-ExcelValue($Value) {
    # 엔진 값 → 엑셀 Value2 에 넣을 값. 돈은 [decimal] 로 계산하고 쓸 때만 그 글자에 가장 가까운 double 로(openpyxl 이 글자로 쓰는 것과 같은 수).
    #   글자: 제어문자를 지우고 앞에 ' 를 붙인다(붙이지 않으면 엑셀이 '00123' · '1/2' 를 숫자 · 날짜로 바꿈). 빈 글자는 빈칸(openpyxl 도 빈칸으로 읽힘).
    #   날짜 · 일시: 엑셀 일련번호(형식은 Get-WriterFormat).
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) {
        $text = Remove-IllegalText $Value
        if ($text.Length -eq 0) { return $null }
        return "'" + $text
    }
    if ($Value -is [bool]) { return [bool]$Value }
    if ($Value -is [decimal]) { return [double]::Parse($Value.ToString($script:Invariant), [Globalization.NumberStyles]::Float, $script:Invariant) }
    if ($Value -is [double] -or $Value -is [single]) { return [double]$Value }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [int16] -or $Value -is [byte]) { return [double]$Value }
    if ($Value -is [datetime]) { return $Value.ToOADate() }
    throw "엑셀에 쓸 수 없는 값 종류입니다: $($Value.GetType().Name)"
}

# ---- runspace 와 함께 쓰는 함수 (본문은 param 으로 시작: runspace 에 글자로 옮겨 심는다) ----------

function Get-WriterWidth {
    # 엑셀에서 보이는 글자 폭(지금 판 _display_width): 일시 19 · 날짜 10, 숫자는 '#,##0.00' 글자 길이, 한글 등 넓은 글자는 2.
    # Loaded = 파일에서 읽은 칸(openpyxl 은 날짜 칸을 늘 일시로 읽으므로 19).
    param($Value, [bool]$Loaded)
    if ($null -eq $Value) { return 0 }
    if ($Value -is [datetime]) {
        if (-not $Loaded -and $Value.TimeOfDay.Ticks -eq 0) { return 10 }
        return 19
    }
    if ($Value -is [bool]) { return 4 }
    if ($Loaded -and $Value -is [int]) { return 7 }   # 엑셀 오류 칸(#VALUE! 등)
    if ($Value -is [double] -or $Value -is [single]) {
        if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) { return 3 }
        return ([double]$Value).ToString('#,##0.00', [Globalization.CultureInfo]::InvariantCulture).Length
    }
    if ($Value -is [decimal]) { return ([decimal]$Value).ToString('#,##0.00', [Globalization.CultureInfo]::InvariantCulture).Length }
    if ($Value -is [int] -or $Value -is [long]) { return ([long]$Value).ToString('#,##0.00', [Globalization.CultureInfo]::InvariantCulture).Length }
    $text = [string]$Value
    $width = 0
    for ($i = 0; $i -lt $text.Length; $i++) {
        $code = [int]$text[$i]
        if ([char]::IsHighSurrogate($text[$i]) -and $i + 1 -lt $text.Length -and [char]::IsLowSurrogate($text[$i + 1])) {
            $code = [char]::ConvertToUtf32($text[$i], $text[$i + 1])
            $i++
        }
        # 동아시아 넓은 글자(W · F)의 주요 범위: 한글 자모 · CJK · 한글 음절 · 호환 한자 · 전각 · 보조 평면 한자 · 그림 문자
        $wide = ($code -ge 0x1100 -and $code -le 0x115F) -or ($code -ge 0x2E80 -and $code -le 0x303E) -or
                ($code -ge 0x3041 -and $code -le 0x33FF) -or ($code -ge 0x3400 -and $code -le 0x4DBF) -or
                ($code -ge 0x4E00 -and $code -le 0x9FFF) -or ($code -ge 0xA000 -and $code -le 0xA4CF) -or
                ($code -ge 0xA960 -and $code -le 0xA97F) -or ($code -ge 0xAC00 -and $code -le 0xD7A3) -or
                ($code -ge 0xF900 -and $code -le 0xFAFF) -or ($code -ge 0xFE10 -and $code -le 0xFE19) -or
                ($code -ge 0xFE30 -and $code -le 0xFE6F) -or ($code -ge 0xFF01 -and $code -le 0xFF60) -or
                ($code -ge 0xFFE0 -and $code -le 0xFFE6) -or ($code -ge 0x1F300 -and $code -le 0x1F64F) -or
                ($code -ge 0x1F900 -and $code -le 0x1F9FF) -or ($code -ge 0x20000 -and $code -le 0x3FFFD)
        if ($wide) { $width += 2 } else { $width += 1 }
    }
    return $width
}

function ConvertTo-SavedText {
    # 파일에서 읽은 칸 값의 파이썬 str(값)(openpyxl 이 읽은 값 기준): 정수 같은 숫자는 정수(엑셀은 12345 로 저장 → openpyxl int),
    # 일시는 'YYYY-MM-DD HH:MM:SS'(날짜 칸도 일시로 읽힘). 장부의 번호 비교(지금 판 str(cell.value) == str(number))에 쓴다.
    param($Value)
    if ($null -eq $Value) { return 'None' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { if ($Value) { return 'True' } else { return 'False' } }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [double]) {
        if ([Math]::Floor($Value) -eq $Value -and [Math]::Abs($Value) -lt 1e15) { return ([long]$Value).ToString([Globalization.CultureInfo]::InvariantCulture) }
        return ([double]$Value).ToString('R', [Globalization.CultureInfo]::InvariantCulture).ToLowerInvariant()
    }
    return [string]$Value
}

$script:WriterShared = @('Get-WriterWidth', 'ConvertTo-SavedText')

# ---- runspace 에서 도는 엑셀 일 ---------------------------------------------------------------------
# 입력 $Plan = @{ steps = @(단계 …) }, 결과 [pscustomobject]@{ results = @(단계마다 결과 …) }. 단계(kind):
#   book    {tmp; sheets = @(시트 계획); read = 다시 읽을 시트 이름 | $null}   → {rows}            새 통합 문서를 tmp 에 저장(+ 다시 읽기)
#   ledger  {path; exists; tmp; layout; entry; logs}                         → {duplicate; doc; items}  장부 고친 판을 tmp 에 저장(+ 다시 읽기)
#   read    {path; sheet}                                                     → {found; rows}      시트 하나를 A1 부터 읽기(읽기 전용)
#   dup     {path; sheet; column; number}                                     → {duplicate}        장부에 그 번호가 있는지(읽기 전용)
# 읽은 칸: 날짜 서식 칸은 [datetime](.Value), 나머지는 .Value2(double · 글자 · 참거짓, 오류 칸은 오류 번호 [int]).
# 엑셀은 COM 참조가 하나라도 남으면 Quit 해도 끝나지 않으므로 손에 쥔 COM 객체를 모두 $held 에 모아 끝에 놓는다(Office.ps1 과 같음).
$script:WriterBody = {
    $ErrorActionPreference = 'Stop'
    $held = New-Object System.Collections.Generic.List[object]
    function Hold($Object) { if ($null -ne $Object) { $held.Add($Object) }; return ,$Object }   # , : 컬렉션(Range 등)이 풀리지 않게
    $missing = [Type]::Missing

    function Get-Cell($Sheet, [int]$Row, [int]$Col) { return ,(Hold $Sheet.Cells.Item($Row, $Col)) }
    function Get-Block($Sheet, [int]$Row, [int]$Col, [int]$Rows, [int]$Cols) {
        return ,(Hold $Sheet.Range((Get-Cell $Sheet $Row $Col), (Get-Cell $Sheet ($Row + $Rows - 1) ($Col + $Cols - 1))))
    }
    function Set-Values($Sheet, [int]$Row, [int]$Col, $Values) {
        # 값은 언제나 2차원 배열로 놓는다: 같은 Value2 에 메서드가 돌려준 낱값을 넣은 뒤 배열을 넣으면
        # PowerShell 5.1 의 COM 연결이 앞 형식으로 바꾸려다 'Specified cast is not valid' 로 실패한다(이 PC 에서 확인).
        $range = Get-Block $Sheet $Row $Col $Values.GetLength(0) $Values.GetLength(1)
        $range.Value2 = $Values
    }
    function New-Values([int]$Rows, [int]$Cols) { return ,([Array]::CreateInstance([object], $Rows, $Cols)) }
    function Read-Sheet($Sheet) {
        # A1 부터 쓴 범위 끝까지 행 목록(각 행 = object[], 길이 같음). 빈 시트는 빈칸 하나짜리 한 행.
        $used = Hold $Sheet.UsedRange
        $rows = [int]$used.Row + [int](Hold $used.Rows).Count - 1
        $cols = [int]$used.Column + [int](Hold $used.Columns).Count - 1
        $range = Get-Block $Sheet 1 1 $rows $cols
        $value2 = $range.Value2
        $value = $range.Value(10)                 # xlRangeValueDefault: 날짜 칸은 [datetime]
        $out = New-Object System.Collections.Generic.List[object]
        for ($r = 1; $r -le $rows; $r++) {
            $line = New-Object 'object[]' $cols
            for ($c = 1; $c -le $cols; $c++) {
                if ($rows -eq 1 -and $cols -eq 1) { $v2 = $value2; $v = $value } else { $v2 = $value2[$r, $c]; $v = $value[$r, $c] }
                if ($v -is [datetime]) { $line[$c - 1] = $v } else { $line[$c - 1] = $v2 }
            }
            $out.Add($line)
        }
        return ,$out
    }
    function Get-LastRow($Rows) {
        # 마지막으로 값이 있는 행 번호(없으면 1 = 머리글). 지금 판 _last_row.
        for ($r = $Rows.Count; $r -gt 1; $r--) {
            foreach ($v in $Rows[$r - 1]) { if ($null -ne $v) { return $r } }
        }
        return 1
    }
    function Find-Sheet($Book, [string]$Name) {
        # 이름이 꼭 같은(대소문자까지) 시트. 없으면 $null.
        $sheets = Hold $Book.Sheets
        for ($i = 1; $i -le [int]$sheets.Count; $i++) {
            $sheet = Hold $sheets.Item($i)
            if ([string]$sheet.Name -ceq $Name) { return ,$sheet }
        }
        return $null
    }
    function Find-Header($Header, [string]$Name) {
        # 머리글 목록에서 그 이름의 첫 자리(0부터). 없으면 -1. 파이썬 list.index 처럼 글자만 같은 것으로 본다.
        for ($i = 0; $i -lt $Header.Count; $i++) {
            if ($Header[$i] -is [string] -and [string]$Header[$i] -ceq $Name) { return $i }
        }
        return -1
    }
    function Set-Freeze($Book, $Sheet, [int]$Row) {
        # 틀 고정: Row 행부터 움직인다(A2 = 머리글 고정). 숨긴 시트는 고를 수 없으므로 건너뛴다.
        if ([int]$Sheet.Visible -ne -1) { return }
        $Sheet.Activate()
        $window = Hold (Hold $Book.Windows).Item(1)
        $window.FreezePanes = $false
        $window.ScrollRow = 1
        $window.ScrollColumn = 1
        $window.SplitColumn = 0
        $window.SplitRow = $Row - 1
        $window.FreezePanes = $true
    }
    function Save-Copy($Book, [string]$Tmp) {
        # tmp 에 저장. 통합 문서 자체는 원래 파일(또는 저장 안 한 새 문서)에 남으므로 tmp 를 열어 두지 않는다(~$ 소유자 파일이 생기지 않음).
        if ([int]$Book.FileFormat -eq 51) { $Book.SaveCopyAs($Tmp) } else { $Book.SaveAs($Tmp, 51) }   # 51 = xlOpenXMLWorkbook
    }
    function Read-Saved($Books, [string]$Path, $Names) {
        # 저장한 파일을 읽기 전용으로 다시 열어 그 시트들을 읽는다: 이름 → 행 목록($null = 시트 없음).
        $book = Hold $Books.Open($Path, 0, $true)
        try {
            $found = @{}
            foreach ($name in $Names) {
                $sheet = Find-Sheet $book $name
                if ($null -eq $sheet) { $found[$name] = $null } else { $found[$name] = Read-Sheet $sheet }
            }
            return $found
        } finally {
            $book.Close($false)
        }
    }
    function Set-FitWidths($Sheet, $Rows) {
        # 지금 판 fit_columns: 마지막 행까지 내용에 맞춰 넓힌다(최소 10, 최대 60). 지금 더 넓은 열은 줄이지 않는다.
        $last = Get-LastRow $Rows
        $cols = $Rows[0].Length
        for ($c = 0; $c -lt $cols; $c++) {
            $need = 0
            for ($r = 0; $r -lt $last; $r++) {
                $w = Get-WriterWidth $Rows[$r][$c] $true
                if ($w -gt $need) { $need = $w }
            }
            $need = [Math]::Min([Math]::Max($need + 4, 10), 60)
            $column = Hold (Hold $Sheet.Columns).Item($c + 1)
            if ($need -gt [double]$column.ColumnWidth) { $column.ColumnWidth = [double]$need }
        }
    }

    function Write-Book($Excel, $Step) {
        $books = Hold $Excel.Workbooks
        $book = Hold $books.Add(-4167)            # xlWBATWorksheet: 시트 한 장짜리 새 문서(사용자 설정의 시트 수와 무관)
        $mine.Add([string]$book.FullName)         # 우리가 만든 문서(끝낼 때 이것만 닫는다)
        try {
            $sheets = Hold $book.Worksheets
            $index = 0
            foreach ($plan in $Step.sheets) {
                if ($index -eq 0) { $sheet = Hold $sheets.Item(1) }
                else { $sheet = Hold $sheets.Add($missing, (Hold $sheets.Item([int]$sheets.Count))) }
                $sheet.Name = [string]$plan.name
                if ($plan.values.GetLength(0) -gt 0 -and $plan.values.GetLength(1) -gt 0) { Set-Values $sheet 1 1 $plan.values }
                foreach ($f in $plan.formats) { (Get-Cell $sheet $f[0] $f[1]).NumberFormat = [string]$f[2] }
                if ($plan.bold_cells -gt 0) { (Hold (Get-Block $sheet $plan.bold_row 1 1 $plan.bold_cells).Font).Bold = $true }
                for ($c = 0; $c -lt $plan.widths.Length; $c++) { (Hold (Hold $sheet.Columns).Item($c + 1)).ColumnWidth = [double]$plan.widths[$c] }
                Set-Freeze $book $sheet $plan.freeze
                $index++
            }
            (Hold $sheets.Item(1)).Activate()
            Save-Copy $book $Step.tmp
        } finally {
            $book.Close($false)
        }
        $rows = $null
        if ($null -ne $Step.read) { $rows = (Read-Saved $books $Step.tmp @($Step.read))[$Step.read] }
        return [pscustomobject]@{ rows = $rows }
    }

    function Get-LedgerModel($Sheet) {
        $rows = Read-Sheet $Sheet
        $header = New-Object System.Collections.Generic.List[object]
        foreach ($v in $rows[0]) { if ($null -ne $v) { $header.Add($v) } }
        return @{ sheet = $Sheet; header = $header; last = (Get-LastRow $rows); rows = $rows }
    }
    function Add-LedgerColumns($Book, $Model, $Columns) {
        # 지금 판 _ensure_columns: 없는 열만 오른쪽 끝에(머리글 수 + 1 번째 칸). 기존 열은 순서 · 이름을 바꾸지 않는다. 틀 고정 A2.
        foreach ($name in $Columns) {
            if ((Find-Header $Model.header $name) -ge 0) { continue }
            $Model.header.Add($name)
            $n = $Model.header.Count
            $one = New-Values 1 1
            $one[0, 0] = "'" + $name
            Set-Values $Model.sheet 1 $n $one
            (Hold (Hold $Model.sheet.Columns).Item($n)).ColumnWidth = [double][Math]::Max(12, $name.Length * 2 + 4)
        }
        Set-Freeze $Book $Model.sheet 2
    }
    function Add-LedgerRow($Model, $Cells) {
        # 지금 판 _append: 마지막 행 다음 줄에, 머리글에 있는 이름만, 값이 없는 칸은 건너뜀. 칸마다 {name; v(엑셀 값); f(형식)}.
        $n = $Model.header.Count
        if ($n -eq 0) { return }
        $row = $Model.last + 1
        $values = New-Values 1 $n
        $formats = New-Object System.Collections.Generic.List[object]
        $any = $false
        foreach ($cell in $Cells) {
            if ($null -eq $cell.v) { continue }
            $i = Find-Header $Model.header ([string]$cell.name)
            if ($i -lt 0) { continue }
            $values[0, $i] = $cell.v
            $any = $true
            if ($null -ne $cell.f) { $formats.Add(@(($i + 1), [string]$cell.f)) }
        }
        if (-not $any) { return }
        Set-Values $Model.sheet $row 1 $values
        foreach ($f in $formats) { (Get-Cell $Model.sheet $row $f[0]).NumberFormat = $f[1] }
        $Model.last = $row
    }
    function Test-LedgerHas($Model, [string]$Column, [string]$Number) {
        # 지금 판 is_duplicate: 번호 열(머리글 자리)의 칸 str 이 번호 str 과 같은 줄이 있는지.
        $i = Find-Header $Model.header $Column
        if ($i -lt 0) { return $false }
        for ($r = 2; $r -le $Model.last; $r++) {
            $line = $Model.rows[$r - 1]
            if ($i -lt $line.Length -and (ConvertTo-SavedText $line[$i]) -ceq $Number) { return $true }
        }
        return $false
    }
    function Move-Sheets($Book, $Order) {
        # 지금 판 _order_doc_sheets: 문서종류 시트가 있는 자리들만 시트순서대로 채운다(사용자 시트는 제자리).
        $sheets = Hold $Book.Sheets
        $names = New-Object System.Collections.Generic.List[string]
        for ($i = 1; $i -le [int]$sheets.Count; $i++) { $names.Add([string](Hold $sheets.Item($i)).Name) }
        $slots = New-Object System.Collections.Generic.List[int]
        for ($i = 0; $i -lt $names.Count; $i++) { if (@($Order) -ccontains $names[$i]) { $slots.Add($i) } }
        $want = New-Object System.Collections.Generic.List[string]
        foreach ($n in $names) { $want.Add($n) }
        for ($k = 0; $k -lt [Math]::Min($slots.Count, @($Order).Count); $k++) { $want[$slots[$k]] = [string]@($Order)[$k] }
        for ($i = 0; $i -lt $want.Count; $i++) {
            $sheet = Find-Sheet $Book $want[$i]
            if ([int]$sheet.Index -eq $i + 1) { continue }
            if ($i -eq 0) { $sheet.Move((Hold $sheets.Item(1))) } else { $sheet.Move($missing, (Hold $sheets.Item($i))) }
        }
    }
    function Update-LedgerTables($Book, $Managed) {
        # 지금 판 save: Ledger_* 표를 관리 시트마다 하나씩 Ledger_1 … 로(문서종류 → 품목 → 처리로그 순) 지금 범위(A1:끝열 마지막행, 데이터가 없어도 2행)로.
        # 관리 시트의 A1 표는 범위만 바꿔 다시 쓰고(사용자가 바꾼 표 모양 유지), 그 밖의 Ledger_* 표는 일반 범위로 되돌린다(지금 판은 지움).
        $keep = @{}
        $sheets = Hold $Book.Worksheets
        $serial = 0
        for ($s = 1; $s -le [int]$sheets.Count; $s++) {
            $sheet = Hold $sheets.Item($s)
            $managedSheet = $false
            foreach ($m in $Managed) { if ([string]$m.sheet.Name -ceq [string]$sheet.Name) { $managedSheet = $true } }
            $objects = Hold $sheet.ListObjects
            $found = New-Object System.Collections.Generic.List[object]
            for ($i = 1; $i -le [int]$objects.Count; $i++) { $found.Add((Hold $objects.Item($i))) }
            foreach ($table in $found) {
                if (-not ([string]$table.Name).StartsWith('Ledger_', [StringComparison]::Ordinal)) { continue }
                $at = Hold $table.Range
                if ($managedSheet -and -not $keep.ContainsKey([string]$sheet.Name) -and [int]$at.Row -eq 1 -and [int]$at.Column -eq 1) {
                    $serial++
                    $table.Name = 'LedgerMoving_' + $serial      # 이름이 겹치지 않게 잠시 다른 이름
                    $keep[[string]$sheet.Name] = $table
                } else {
                    $table.Unlist()
                }
            }
        }
        for ($i = 0; $i -lt $Managed.Count; $i++) {
            $m = $Managed[$i]
            $range = Get-Block $m.sheet 1 1 ([Math]::Max($m.last, 2)) $m.header.Count
            if ($keep.ContainsKey([string]$m.sheet.Name)) {
                $table = $keep[[string]$m.sheet.Name]
                $table.Resize($range)
            } else {
                $table = Hold (Hold $m.sheet.ListObjects).Add(1, $range, $missing, 1)   # xlSrcRange, 머리글 있음
                $table.TableStyle = 'TableStyleMedium2'
                $table.ShowTableStyleRowStripes = $true
            }
            $table.Name = 'Ledger_' + ($i + 1)
        }
    }

    function Write-Ledger($Excel, $Step) {
        $books = Hold $Excel.Workbooks
        $layout = $Step.layout
        if ($Step.exists) {
            $book = Hold $books.Open($Step.path, 0, $true)     # 장부의 임시 복사본을 읽기 전용으로(고친 판은 tmp 로)
        } else {
            $book = Hold $books.Add(-4167)
            $mine.Add([string]$book.FullName)
            $first = Hold (Hold $book.Worksheets).Item(1)
            $first.Name = $layout.items
            $second = Hold (Hold $book.Worksheets).Add($missing, $first)
            $second.Name = $layout.log
        }
        $opened = $null
        if ($Step.exists) { $opened = [string]$book.FullName }
        $result = [pscustomobject]@{ duplicate = $false; doc = $null; items = $null; opened = $opened }
        try {
            # 열 맞추기(지금 판 _sync_columns): 품목 · 처리로그가 없으면 끝에, 문서종류 시트가 없으면 품목 앞에
            foreach ($name in @($layout.items, $layout.log)) {
                if ($null -eq (Find-Sheet $book $name)) {
                    $sheets = Hold $book.Sheets
                    (Hold (Hold $book.Worksheets).Add($missing, (Hold $sheets.Item([int]$sheets.Count)))).Name = $name
                }
            }
            $models = @{}
            $managed = New-Object System.Collections.Generic.List[object]
            foreach ($doc in $layout.docs) {
                $sheet = Find-Sheet $book $doc.name
                if ($null -eq $sheet) {
                    $sheet = Hold (Hold $book.Worksheets).Add((Find-Sheet $book $layout.items))
                    $sheet.Name = $doc.name
                }
                $model = Get-LedgerModel $sheet
                Add-LedgerColumns $book $model $doc.columns
                $models[$doc.name] = $model
                $managed.Add($model)
            }
            foreach ($pair in @(@($layout.items, $layout.item_columns), @($layout.log, $layout.log_columns))) {
                $model = Get-LedgerModel (Find-Sheet $book $pair[0])
                Add-LedgerColumns $book $model $pair[1]
                $models[$pair[0]] = $model
                $managed.Add($model)
            }
            Move-Sheets $book $layout.order

            $entry = $Step.entry
            if ($null -ne $entry) {
                if ($entry.check_duplicate -and (Test-LedgerHas $models[$entry.doc] $entry.number_column $entry.number)) {
                    $result.duplicate = $true
                    return $result
                }
                Add-LedgerRow $models[$entry.doc] $entry.doc_cells
                foreach ($cells in $entry.item_rows) { Add-LedgerRow $models[$layout.items] $cells }
            }
            foreach ($cells in $Step.logs) { Add-LedgerRow $models[$layout.log] $cells }

            Update-LedgerTables $book $managed
            foreach ($m in $managed) { Set-FitWidths $m.sheet (Read-Sheet $m.sheet) }
            $front = Find-Sheet $book $layout.first     # 열면 맨 앞 문서종류 탭이 보이게(그 탭만 고름)
            if ($null -ne $front -and [int]$front.Visible -eq -1) { $front.Activate(); $front.Select($true) }
            Save-Copy $book $Step.tmp
        } finally {
            $book.Close($false)
        }
        if ($null -ne $entry) {
            $saved = Read-Saved $books $Step.tmp @($entry.doc, $layout.items)
            $result.doc = $saved[$entry.doc]
            $result.items = $saved[$layout.items]
        }
        return $result
    }

    function Read-OneSheet($Excel, $Step) {
        $saved = Read-Saved (Hold $Excel.Workbooks) $Step.path @($Step.sheet)
        $rows = $saved[$Step.sheet]
        return [pscustomobject]@{ found = ($null -ne $rows); rows = $rows }
    }

    function Test-Duplicate($Excel, $Step) {
        $book = Hold (Hold $Excel.Workbooks).Open($Step.path, 0, $true)
        try {
            $sheet = Find-Sheet $book $Step.sheet
            if ($null -eq $sheet) { return [pscustomobject]@{ duplicate = $false } }
            return [pscustomobject]@{ duplicate = [bool](Test-LedgerHas (Get-LedgerModel $sheet) $Step.column $Step.number) }
        } finally {
            $book.Close($false)
        }
    }

    $excel = $null
    # 우리가 여는 · 만드는 문서의 전체 경로(끝낼 때 이것만 닫고, 남의 문서가 남아 있으면 Quit 하지 않고 넘겨 준다. Office.ps1 의 Exit-OfficeApp)
    $mine = New-Object System.Collections.Generic.List[string]
    foreach ($step in $Plan.steps) {
        foreach ($key in @('path', 'tmp')) { if ($step.ContainsKey($key) -and $step[$key]) { $mine.Add([string]$step[$key]) } }
    }
    try {
        $excel = Hold (New-Object -ComObject Excel.Application)
        Assert-OfficeFresh $excel 'excel'       # 이미 통합 문서가 열린 엑셀이면 우리 것이 아니다: 손대지 않고 오류(finally 도 Quit 하지 않음)
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3           # msoAutomationSecurityForceDisable (매크로 끔, 이 엑셀에만)
        Invoke-OfficeTestExtra $excel 'excel'
        $results = New-Object System.Collections.Generic.List[object]
        foreach ($step in $Plan.steps) {
            switch ([string]$step.kind) {
                'book' { $results.Add((Write-Book $excel $step)) }
                'ledger' { $results.Add((Write-Ledger $excel $step)) }
                'read' { $results.Add((Read-OneSheet $excel $step)) }
                'dup' { $results.Add((Test-Duplicate $excel $step)) }
                default { throw "엑셀 쓰기 단계 '$($step.kind)' 는 없습니다." }
            }
        }
        [pscustomobject]@{ results = $results.ToArray() }
    } finally {
        Exit-OfficeApp $excel 'excel' $mine.ToArray()   # 남의 통합 문서가 남아 있으면 Quit 대신 보이게 넘겨 준다
        for ($i = $held.Count - 1; $i -ge 0; $i--) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($held[$i]) } catch { }
        }
        $held.Clear()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$script:WriterWorkCache = $null

function Get-WriterWork {
    # runspace 에 넘길 일: param($Plan) + 함께 쓰는 함수(글자로 옮겨 심음) + 본문.
    if ($null -eq $script:WriterWorkCache) {
        $text = New-Object System.Text.StringBuilder
        [void]$text.AppendLine('param($Plan)')
        foreach ($name in $script:WriterShared) {
            [void]$text.AppendLine('function ' + $name + ' {')
            [void]$text.AppendLine((Get-Item ('function:' + $name)).ScriptBlock.ToString())
            [void]$text.AppendLine('}')
        }
        [void]$text.AppendLine($script:WriterBody.ToString())
        $script:WriterWorkCache = [scriptblock]::Create($text.ToString())
    }
    return $script:WriterWorkCache
}

function Invoke-ExcelWrite($Steps) {
    # 단계들을 엑셀 하나로 차례로 돌린다(엑셀을 한 번만 띄움). 단계마다 결과 목록. 시간 제한을 넘으면 TimeoutException(엑셀이 응답하지 않음).
    $plan = @{ steps = (ConvertTo-ItemArray $Steps) }
    $out = Invoke-OfficeWork (Get-WriterWork) $plan 'EXCEL.EXE' $script:OfficeTimeoutSec $script:ExcelTimeoutText '엑셀' $script:ExcelWriteFailText
    return ,@($out.results)
}

# ---- 새 통합 문서(회사 양식 · 검산 파일) ------------------------------------------------------------

function New-SheetPlan([string]$Name, $Rows, [int]$BoldRow, [int]$FreezeRow) {
    # 시트 한 장의 쓸 거리: 값(2차원, 엑셀 값) · 형식(2행부터, 지금 판 _write_sheet) · 굵은 행 · 틀 고정 · 열 너비(지금 판 fit_columns).
    $lines = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Rows) { $lines.Add((ConvertTo-ItemArray $row)) }
    $cols = 0
    foreach ($line in $lines) { if ($line.Count -gt $cols) { $cols = $line.Count } }
    $values = [Array]::CreateInstance([object], $lines.Count, $cols)
    $formats = New-Object System.Collections.Generic.List[object]
    $widths = New-Object 'int[]' $cols
    for ($r = 0; $r -lt $lines.Count; $r++) {
        $line = $lines[$r]
        for ($c = 0; $c -lt $line.Count; $c++) {
            $value = $line[$c]
            $values[$r, $c] = ConvertTo-ExcelValue $value
            if ($r -ge 1) {
                $format = Get-WriterFormat $value
                if ($null -ne $format) { $formats.Add(@(($r + 1), ($c + 1), $format)) }
            }
            $shown = $value
            if ($shown -is [string]) { $shown = Remove-IllegalText $shown }
            $w = Get-WriterWidth $shown $false
            if ($w -gt $widths[$c]) { $widths[$c] = $w }
        }
    }
    $fit = New-Object 'double[]' $cols
    for ($c = 0; $c -lt $cols; $c++) { $fit[$c] = [Math]::Min([Math]::Max($widths[$c] + 4, 10), 60) }
    $bold = 0
    if ($BoldRow -ge 1 -and $BoldRow -le $lines.Count) { $bold = $lines[$BoldRow - 1].Count }
    return @{ name = $Name; values = $values; formats = $formats.ToArray(); bold_row = $BoldRow; bold_cells = $bold
              widths = $fit; freeze = $FreezeRow }
}

function Get-TmpPath([string]$Final) {
    # 지금 판 job_run: 같은 폴더의 .<이름>.tmp<확장자>
    $folder = [IO.Path]::GetDirectoryName($Final)
    return Join-Path $folder ('.' + [IO.Path]::GetFileNameWithoutExtension($Final) + '.tmp' + [IO.Path]::GetExtension($Final))
}

function New-ReadBackError($Problems) {
    # 다시 읽기에서 다른 칸이 있을 때 던질 오류(지금 판 job_run 의 확인필요 문구 모양). 문제 목록은 .Data['problems'].
    $list = ConvertTo-ItemArray $Problems
    $shown = [string]::Join('; ', @($list | Select-Object -First $script:ReadBackShown))
    if ($list.Count -gt $script:ReadBackShown) { $shown += (' 외 {0}건' -f ($list.Count - $script:ReadBackShown)) }
    $failure = New-Object System.Exception ('저장한 파일을 다시 읽어 보니 쓴 값과 다릅니다: ' + $shown)
    $failure.Data['problems'] = [string[]]$list
    return $failure
}

function Save-NewBook([string]$Final, $Sheets, [string]$ReadSheet, $Header, $Rows) {
    # 새 통합 문서를 tmp 에 저장 → (ReadSheet 를 주면) 다시 읽어 쓴 값과 견줌 → 맞을 때만 최종 이름으로 옮긴다(있던 파일은 덮지 않음).
    $folder = [IO.Path]::GetDirectoryName($Final)
    [void][IO.Directory]::CreateDirectory($folder)
    $tmp = Get-TmpPath $Final
    $read = $null
    if (-not [string]::IsNullOrEmpty($ReadSheet)) { $read = $ReadSheet }
    try {
        $raw = (Invoke-ExcelWrite @(@{ kind = 'book'; tmp = $tmp; sheets = (ConvertTo-ItemArray $Sheets); read = $read }))[0]
        if ($null -ne $read) {
            $problems = Compare-ReadBack $raw.rows $read $Header $Rows
            if ($problems.Count -gt 0) { throw (New-ReadBackError $problems) }
        }
        $appeared = '저장하는 사이 출력 폴더에 {0} 이 생겼습니다' -f [IO.Path]::GetFileName($Final)
        if (Test-Path -LiteralPath $Final) { throw $appeared }
        try { [IO.File]::Move($tmp, $Final) } catch { if (Test-Path -LiteralPath $Final) { throw $appeared }; throw }   # 옮기기는 있는 파일을 덮지 않는다
    } finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
    return $Final
}

function Get-RowList($Rows) {
    # 행 목록(List · 배열 · $null)을 행마다 object[] 로.
    $list = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Rows) { return ,$list }
    foreach ($row in $Rows) { $list.Add((ConvertTo-ItemArray $row)) }
    return ,$list
}

function Write-CompanyForm([string]$Folder, $Result, $Company, $Form, [switch]$Exact) {
    # 회사 양식 파일(지금 판 write_company_form): 시트 = 양식 시트, 머리글 굵게 · 틀 고정 A2, 이름은 양식 파일명(같은 이름이 있으면 _2 …).
    # -Exact: 같은 이름이 이미 있으면 _2 로 쓰지 않고 멈춘다(지금 판 job_run: 확인필요). 오류 .Data['exists'] = $true.
    #   마지막 옮기기(File.Move)는 있는 파일을 덮지 않으므로, 확인한 뒤 그 사이에 생겨도 쓰지 않는다.
    # $Company = Get-CompanyRows 결과 {header, rows}. 저장 뒤 다시 읽어 다르면 파일을 남기지 않고 오류(.Data['problems']).
    $wanted = (Join-Path $Folder (Get-CompanyFormName $Result $Form)).Replace('/', '\')
    if ($Exact) {
        if (Test-Path -LiteralPath $wanted) {
            $failure = New-Object System.IO.IOException ('출력 폴더에 {0} 이 이미 있습니다' -f [IO.Path]::GetFileName($wanted))
            $failure.Data['exists'] = $true
            throw $failure
        }
        $final = $wanted
    } else {
        $final = Get-FreeName $wanted $null
    }
    $header = ConvertTo-ItemArray $Company.header
    $rows = Get-RowList $Company.rows
    $all = New-Object System.Collections.Generic.List[object]
    $all.Add($header)
    foreach ($row in $rows) { $all.Add($row) }
    $sheet = New-SheetPlan ([string]$Form.sheet) $all 1 2
    return Save-NewBook $final @($sheet) ([string]$Form.sheet) $header $rows
}

function Get-CheckMark($Ok) {
    if ($null -eq $Ok) { return '해당 없음' }
    if ([bool]$Ok) { return '일치' }
    return '불일치'
}

function Write-CheckFile([string]$Folder, $Result, $Checks, [string]$SourceName, $Company, $Rules) {
    # 검산 파일(지금 판 write_check_file): 검산(1행 원본파일, 2행 항목 머리 굵게, 틀 고정 A3) · 추출값 · 품목(품목이 있을 때) ·
    # 회사양식($Company = {header, rows} 를 줄 때). 이름 검산_<번호필드>.xlsx(같은 이름이 있으면 _2 …). $Rules 는 번호필드를 찾는 데 쓴다.
    $result = ConvertTo-WriterResult $Rules $Result
    $doc = Get-WriterDoc $Rules $result.doc_type
    $sheets = New-Object System.Collections.Generic.List[object]

    $lines = New-Object System.Collections.Generic.List[object]
    $lines.Add([object[]]@('원본파일', $SourceName))
    $lines.Add([object[]]$script:CheckHeader)
    # (@() 로 감싸지 않는다: PowerShell 5.1 은 OrderedDictionary 가 든 List 를 @() 로 펼치면 'Argument types do not match')
    foreach ($check in $Checks) {
        if ($null -eq $check) { continue }
        $lines.Add([object[]]@((Get-MarkupProp $check 'item'), (Get-MarkupProp $check 'source_value'),
                                (Get-MarkupProp $check 'result_value'), (Get-CheckMark (Get-MarkupProp $check 'ok'))))
    }
    $sheets.Add((New-SheetPlan $script:CheckSheet $lines 2 3))

    $lines = New-Object System.Collections.Generic.List[object]
    $lines.Add([object[]]@('필드', '값'))
    foreach ($name in $result.fields.Keys) {
        $value = $result.fields[$name]
        if ($null -eq $value) { $value = '(못 찾음)' }
        $lines.Add([object[]]@([string]$name, $value))
    }
    $sheets.Add((New-SheetPlan $script:CheckFieldsSheet $lines 1 2))

    $table = $result.table
    if ($null -ne $table -and $table.Count -gt 0) {
        $columns = New-Object System.Collections.Generic.List[string]
        foreach ($row in $table) { foreach ($key in $row.Keys) { if (-not $columns.Contains([string]$key)) { $columns.Add([string]$key) } } }
        $lines = New-Object System.Collections.Generic.List[object]
        $lines.Add([object[]]$columns.ToArray())
        foreach ($row in $table) {
            $line = New-Object 'object[]' $columns.Count
            for ($c = 0; $c -lt $columns.Count; $c++) { if ($row.Contains($columns[$c])) { $line[$c] = $row[$columns[$c]] } }
            $lines.Add($line)
        }
        $sheets.Add((New-SheetPlan $script:CheckItemsSheet $lines 1 2))
    }
    if ($null -ne $Company) {
        $lines = New-Object System.Collections.Generic.List[object]
        $lines.Add((ConvertTo-ItemArray $Company.header))
        foreach ($row in (Get-RowList $Company.rows)) { $lines.Add($row) }
        $sheets.Add((New-SheetPlan $script:CheckCompanySheet $lines 1 2))
    }
    $name = Format-Template ('검산_{' + $doc.number_field + '}.xlsx') $result.fields
    $final = Get-FreeName (Join-Path $Folder $name) $null
    return Save-NewBook $final $sheets.ToArray() $null $null $null
}

# ---- 다시 읽기 ------------------------------------------------------------------------------------

function Test-Blank($Value) { return ($null -eq $Value) -or ($Value -is [string] -and $Value.Length -eq 0) }

function ConvertTo-SavedValue($Value) {
    # 엑셀에서 읽은 칸 → openpyxl(data_only)이 읽는 값: 오류 칸(오류 번호 [int])은 오류 글자.
    if ($Value -is [int]) {
        $text = $script:ExcelErrorText[[int]$Value]
        if ($null -eq $text) { $text = '#ERROR!' }
        return $text
    }
    return $Value
}

function ConvertTo-ReadBackDecimal($Value) {
    # 파이썬 Decimal(str(값)): 쓴 값 · 읽은 값 모두 글자를 거쳐 정확한 십진수로(double 은 가장 짧은 왕복 글자). 못 바꾸면 $null.
    try {
        if ($Value -is [decimal]) { return $Value }
        if ($Value -is [int] -or $Value -is [long]) { return [decimal]$Value }
        if ($Value -is [double]) {
            if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) { return $null }
            return [decimal]::Parse((ConvertTo-FloatText $Value), [Globalization.NumberStyles]::Float, $script:Invariant)
        }
    } catch {
        return $null
    }
    return $null
}

function Get-DecimalDigits([decimal]$Value) {
    # 파이썬 len(Decimal.as_tuple().digits): 계수의 자리 수(앞의 0 은 빼고, 뒤의 0 은 셈).
    $digits = [Math]::Abs($Value).ToString($script:Invariant).Replace('.', '').TrimStart('0')
    if ($digits.Length -eq 0) { return 1 }
    return $digits.Length
}

function Test-IsNumber($Value) {
    return ($Value -is [decimal] -or $Value -is [double] -or $Value -is [int] -or $Value -is [long]) -and -not ($Value -is [bool])
}

function Test-SameValue($Want, $Got) {
    # 지금 판 _same_value. 쓴 값은 엔진 값([decimal] · [long] · [double] · [datetime] · 글자 · 참거짓), 읽은 값은 엑셀 칸.
    # 돈([decimal])은 정확히: 읽은 double 을 가장 짧은 왕복 글자로 십진수로 바꿔 쓴 십진수와 견준다(double 끼리 대충 견주지 않음).
    if ((Test-Blank $Want) -or (Test-Blank $Got)) { return (Test-Blank $Want) -and (Test-Blank $Got) }
    if ($Want -is [string]) { return ($Got -is [string]) -and ((Remove-IllegalText $Want) -ceq $Got) }
    if ($Want -is [bool] -or $Got -is [bool]) { return ($Want -is [bool]) -and ($Got -is [bool]) -and ($Want -eq $Got) }
    if ($Want -is [datetime]) { return ($Got -is [datetime]) -and ($Want -eq $Got) }
    if ((Test-IsNumber $Want) -and (Test-IsNumber $Got)) {
        $number = ConvertTo-ReadBackDecimal $Want
        $back = ConvertTo-ReadBackDecimal $Got
        if ($null -eq $number -or $null -eq $back) { return $false }
        if ($number -eq $back) { return $true }
        # 15자리를 넘는 소수는 엑셀이 15자리로 저장하므로 그때만 double 로 견준다. 정수는 정확히 같아야 한다
        if ($number -eq [decimal]::Truncate($number) -or (Get-DecimalDigits $number) -le 15) { return $false }
        $wantDouble = [double]::Parse($number.ToString($script:Invariant), [Globalization.NumberStyles]::Float, $script:Invariant)
        $gotDouble = [double]::Parse($back.ToString($script:Invariant), [Globalization.NumberStyles]::Float, $script:Invariant)
        return $wantDouble -eq $gotDouble
    }
    return [object]::Equals($Want, $Got)
}

function Format-ReadBackText($Value, [switch]$Saved) {
    # 지금 판 _text: 빈칸 '(빈칸)', 금액은 자리 그대로, 날짜 'YYYY-MM-DD', 일시 'YYYY-MM-DD HH:MM:SS'(읽은 날짜 칸은 늘 일시).
    if (Test-Blank $Value) { return '(빈칸)' }
    if ($Value -is [decimal]) { return $Value.ToString($script:Invariant) }
    if ($Value -is [datetime]) {
        if (-not $Saved -and $Value.TimeOfDay.Ticks -eq 0) { return $Value.ToString('yyyy-MM-dd', $script:Invariant) }
        return $Value.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant)
    }
    if ($Value -is [bool]) { if ($Value) { return 'True' } else { return 'False' } }
    if ($Value -is [double]) {
        if ($Saved -and [Math]::Floor($Value) -eq $Value -and [Math]::Abs($Value) -lt 1e15) { return ([long]$Value).ToString($script:Invariant) }
        return ConvertTo-FloatText $Value
    }
    if ($Value -is [int] -or $Value -is [long]) { return ([long]$Value).ToString($script:Invariant) }
    return [string]$Value
}

function Get-UsedCount($Row) {
    $n = $Row.Count
    while ($n -gt 0 -and (Test-Blank $Row[$n - 1])) { $n-- }
    return $n
}

function Compare-ReadBack($SavedRows, [string]$Sheet, $Header, $Rows) {
    # 지금 판 read_back 의 비교: 읽은 행 목록($null = 시트 없음)과 쓴 머리글 · 줄. 다른 칸의 설명 목록(같으면 빈 목록).
    $problems = New-Object System.Collections.Generic.List[string]
    if ($null -eq $SavedRows) { $problems.Add("시트 '$Sheet' 가 없습니다"); return ,$problems }
    $saved = New-Object System.Collections.Generic.List[object]
    foreach ($line in $SavedRows) {
        $converted = New-Object 'object[]' $line.Length
        for ($c = 0; $c -lt $line.Length; $c++) { $converted[$c] = ConvertTo-SavedValue $line[$c] }
        $saved.Add($converted)
    }
    while ($saved.Count -gt 0 -and (Get-UsedCount $saved[$saved.Count - 1]) -eq 0) { $saved.RemoveAt($saved.Count - 1) }
    $want = ConvertTo-ItemArray $Header
    $gotHeader = [object[]]@()
    if ($saved.Count -gt 0) { $gotHeader = $saved[0] }
    for ($col = 0; $col -lt [Math]::Max($want.Count, (Get-UsedCount $gotHeader)); $col++) {
        $w = if ($col -lt $want.Count) { $want[$col] } else { $null }
        $g = if ($col -lt $gotHeader.Count) { $gotHeader[$col] } else { $null }
        if (-not (Test-SameValue $w $g)) {
            $problems.Add(('머리글 {0}열: 쓴 값 {1}, 읽은 값 {2}' -f ($col + 1), (Format-ReadBackText $w), (Format-ReadBackText $g -Saved)))
        }
    }
    $names = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $want.Count; $i++) { if ($null -ne $want[$i]) { $names.Add([string]$want[$i]) } else { $names.Add("$($i + 1)열") } }
    $lines = Get-RowList $Rows
    for ($index = 0; $index -lt $lines.Count; $index++) {
        $row = $lines[$index]
        $line = $index + 2
        $gotRow = $null
        if ($index + 1 -lt $saved.Count) { $gotRow = $saved[$index + 1] }
        if ($null -eq $gotRow) {
            if ((Get-UsedCount $row) -eq 0) { continue }    # 빈 줄은 파일 끝에서 사라질 수 있다
            $problems.Add("${line}행: 쓴 줄이 파일에 없습니다")
            continue
        }
        for ($col = 0; $col -lt [Math]::Max($row.Count, (Get-UsedCount $gotRow)); $col++) {
            $w = if ($col -lt $row.Count) { $row[$col] } else { $null }
            $g = if ($col -lt $gotRow.Count) { $gotRow[$col] } else { $null }
            if (-not (Test-SameValue $w $g)) {
                $name = if ($col -lt $names.Count) { $names[$col] } else { "$($col + 1)열" }
                $problems.Add(('{0}행 {1}: 쓴 값 {2}, 읽은 값 {3}' -f $line, $name, (Format-ReadBackText $w), (Format-ReadBackText $g -Saved)))
            }
        }
    }
    for ($index = $lines.Count + 1; $index -lt $saved.Count; $index++) { $problems.Add("$($index + 1)행: 쓰지 않은 줄이 파일에 있습니다") }
    return ,$problems
}

function Test-ReadBack([string]$Path, [string]$Sheet, $Header, $Rows) {
    # 저장된 파일(한 파일 = 한 문서)의 시트를 다시 열어 머리글과 각 칸이 쓴 값과 같은지 본다. 다른 칸의 설명 목록 List[string](같으면 빈 목록).
    # 파일은 임시 복사본을 연다(원본은 열지 않음).
    $copy = New-OfficeCopy $Path
    try {
        $raw = (Invoke-ExcelWrite @(@{ kind = 'read'; path = $copy.Path; sheet = $Sheet }))[0]
    } finally {
        Remove-OfficeCopy $copy
    }
    $saved = $null      # (PowerShell 변수 이름은 대소문자를 가리지 않으므로 $Rows 와 다른 이름)
    if ($raw.found) { $saved = $raw.rows }
    return Compare-ReadBack $saved $Sheet $Header $Rows
}

# ---- 누적장부 -------------------------------------------------------------------------------------

function Test-FileLocked([string]$Path) {
    # 다른 프로그램(엑셀 · PDF 뷰어 등)이 파일을 열고 있으면 $true(지금 판 is_locked: 혼자 읽고 쓰기로 열어 보고 바로 닫음, 파일은 바꾸지 않음).
    # 없는 파일은 $false(잡고 있는 프로그램이 없음). 읽기 전용 파일도 지금 판처럼 $true.
    if (-not [IO.File]::Exists($Path)) { return $false }
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $stream.Dispose()
        return $false
    } catch [IO.FileNotFoundException] {
        return $false
    } catch [IO.DirectoryNotFoundException] {
        return $false
    } catch {
        return $true
    }
}

function ConvertTo-WriterResult($Rules, $Result) {
    # 결과(메모리 [ordered] 또는 결과 JSON) → [ordered]@{ doc_type; fields; table }(값 종류가 산 것).
    $docName = Get-ResultMember $Result 'doc_type'
    if ($Result -is [System.Collections.IDictionary]) {
        return [ordered]@{ doc_type = [string]$docName; fields = $Result['fields']; table = $Result['table'] }
    }
    $typed = ConvertTo-CompanyResult $Rules $Result
    return [ordered]@{ doc_type = [string]$docName; fields = $typed.fields; table = $typed.table }
}

function Get-WriterDoc($Rules, [string]$Name) {
    $doc = @($Rules.doc_types) | Where-Object { [string]$_.name -ceq $Name } | Select-Object -First 1
    if ($null -eq $doc) { throw "문서종류 '$Name' 가 규칙에 없습니다." }
    return $doc
}

function Get-LedgerLayout($Rules) {
    # 장부 모양(지금 판 _sync_columns · _sheet_order): 문서종류 시트 = 처리일시 · 원본파일 + 필드 + 정리된경로, 품목 = 문서종류 · 문서번호 + 표 열.
    $docs = New-Object System.Collections.Generic.List[object]
    $itemColumns = New-Object System.Collections.Generic.List[string]
    foreach ($c in $script:LedgerItemHead) { $itemColumns.Add($c) }
    foreach ($doc in @($Rules.doc_types)) {
        $columns = New-Object System.Collections.Generic.List[string]
        foreach ($c in $script:LedgerDocHead) { $columns.Add($c) }
        foreach ($field in @($doc.fields)) { $columns.Add([string]$field.name) }
        foreach ($c in $script:LedgerDocTail) { $columns.Add($c) }
        $docs.Add(@{ name = [string]$doc.name; columns = $columns.ToArray() })
        if ($null -ne $doc.table) {
            foreach ($c in @($doc.table.columns)) { if (-not $itemColumns.Contains([string]$c)) { $itemColumns.Add([string]$c) } }
        }
    }
    $order = @($Rules.sheet_order | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ })
    if ($order.Count -eq 0) { $order = @($docs | ForEach-Object { $_.name }) }
    return @{ docs = $docs.ToArray(); items = $script:LedgerItemsSheet; log = $script:LedgerLogSheet
              item_columns = $itemColumns.ToArray(); log_columns = $script:LedgerLogColumns; order = [string[]]$order; first = [string]$order[0] }
}

function ConvertTo-LedgerCells($Values, [string[]]$DateTimeNames = @()) {
    # 이름 → 엔진 값(순서 있는 표) → 장부 칸 목록 {name; v(엑셀 값); f(형식)}.
    # DateTimeNames 의 열(처리일시 · 일시)은 0시여도 일시 형식(지금 판은 이 값이 늘 datetime 이라 늘 일시 형식).
    $cells = New-Object System.Collections.Generic.List[object]
    foreach ($name in $Values.Keys) {
        $value = $Values[$name]
        $format = Get-WriterFormat $value
        if ($value -is [datetime] -and $DateTimeNames -ccontains [string]$name) { $format = $script:WriterDateTimeFormat }
        $cells.Add(@{ name = [string]$name; v = (ConvertTo-ExcelValue $value); f = $format })
    }
    return ,$cells.ToArray()
}

function New-LedgerOutcome([string]$Path) {
    return [ordered]@{ locked = $false; duplicate = $false; saved = $false; problems = [string[]]@(); path = $Path; created = $false; backup = $null }
}

function Get-BackupPath([string]$Path) {
    # 지금 판: <이름>.bak<확장자>
    $folder = [IO.Path]::GetDirectoryName($Path)
    return Join-Path $folder ([IO.Path]::GetFileNameWithoutExtension($Path) + '.bak' + [IO.Path]::GetExtension($Path))
}

function Invoke-LedgerReplace([string]$Source, [string]$Destination, [string]$Backup) {
    # [IO.File]::Replace(Win32 ReplaceFile). Backup 이 비면 백업 없이.
    # 시험용 주입: 환경 변수 NOINSTALL_TEST_REPLACE_FAIL 이 있으면 ReplaceFile 이 중간에 실패한 상태를 만든다
    #   (ERROR_UNABLE_TO_MOVE_REPLACEMENT_2: 교체될 파일은 백업 이름 — 백업이 없으면 다른 이름 — 으로 옮겨졌고, 새 판은 제 이름 그대로).
    if (-not [string]::IsNullOrEmpty($env:NOINSTALL_TEST_REPLACE_FAIL)) {
        $moved = $Backup
        if ([string]::IsNullOrEmpty($moved)) { $moved = $Destination + '.replaced-' + [Guid]::NewGuid().ToString('N').Substring(0, 8) }
        if ([IO.File]::Exists($moved)) { [IO.File]::Delete($moved) }
        [IO.File]::Move($Destination, $moved)
        throw (New-Object System.IO.IOException '시험: 바꾸기 중간 실패(ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 흉내)')
    }
    if ([string]::IsNullOrEmpty($Backup)) { [IO.File]::Replace($Source, $Destination, [NullString]::Value) }
    else { [IO.File]::Replace($Source, $Destination, $Backup) }
}

function Complete-LedgerSave($Outcome, [string]$Tmp, [bool]$Existed) {
    # 다시 읽기를 통과한 tmp 를 장부로. 있던 장부는 ReplaceFile 한 번으로 바꾸며 바로 전 판을 .bak 으로 옮긴다(지금 판 save 의
    # 「.bak 복사 → 바꾸기」 와 같은 시점 · 내용). 백업 이름을 주므로 바꾸기가 중간에 실패해도 바로 전 판은 장부 또는 .bak 이름으로 남고,
    # 장부 이름에 파일이 없으면 .bak 을 장부 자리로 되살린다. 장부가 제자리에 없으면 tmp(새 판)는 지우지 않는다. 그 사이 누가 열었으면 locked.
    $path = $Outcome.path
    $backup = Get-BackupPath $path
    try {
        if ($Existed) {
            Invoke-LedgerReplace $Tmp $path $backup
            $Outcome.backup = $backup
        } else {
            [IO.File]::Move($Tmp, $path)
            $Outcome.created = $true
        }
        $Outcome.saved = $true
    } catch {
        $reason = $_.Exception.Message
        if ($Existed -and -not [IO.File]::Exists($path) -and [IO.File]::Exists($backup)) {
            try { [IO.File]::Copy($backup, $path) } catch { }
        }
        if (-not [IO.File]::Exists($path) -and [IO.File]::Exists($Tmp)) {
            throw ('누적장부를 저장하지 못함: {0} (새 판은 {1} 에 남겨 두었습니다)' -f $reason, $Tmp)
        }
        if ((Test-FileLocked $path) -or (Test-FileLocked $backup)) { $Outcome.locked = $true; return }
        throw ('누적장부를 저장하지 못함: ' + $reason)
    } finally {
        if ([IO.File]::Exists($path)) { Remove-LedgerTmp $Tmp }
    }
}

function Remove-LedgerTmp([string]$Tmp) {
    if (Test-Path -LiteralPath $Tmp) { Remove-Item -LiteralPath $Tmp -Force -ErrorAction SilentlyContinue }
}

function Assert-LedgerNotLost([string]$Path) {
    # 장부가 없는데 바로 전 판(.bak)이 있으면 멈춘다: 새 장부를 만들면 다음 저장 때 .bak(그때까지의 이력)을 덮는다.
    $backup = Get-BackupPath $Path
    if (-not [IO.File]::Exists($Path) -and [IO.File]::Exists($backup)) {
        $text = '누적장부 {0} 이 없는데 바로 전 판 {1} 이 있어 멈췄습니다. 새 장부를 만들면 다음 저장 때 {1} 의 기록을 덮습니다. ' +
                '{1} 을 {0} 으로 복사해 되살리거나, 필요 없으면 {1} 을 다른 곳으로 옮긴 뒤 다시 실행하세요.'
        throw ($text -f [IO.Path]::GetFileName($Path), [IO.Path]::GetFileName($backup))
    }
}

function Invoke-LedgerStep([string]$LedgerPath, $Rules, $Entry, $Logs) {
    # 장부 고친 판을 tmp 에 만든다(엑셀 한 번). 있던 장부는 원본이 아니라 임시 복사본을 연다. 결과 @{ raw; tmp; existed }.
    $existed = [IO.File]::Exists($LedgerPath)
    $tmp = Get-TmpPath $LedgerPath
    $copy = $null
    try {
        $source = $LedgerPath
        if ($existed) { $copy = New-OfficeCopy $LedgerPath; $source = $copy.Path }
        $raw = (Invoke-ExcelWrite @((New-LedgerStep $source $existed $tmp $Rules $Entry $Logs)))[0]
    } catch {
        Remove-LedgerTmp $tmp
        throw
    } finally {
        Remove-OfficeCopy $copy
    }
    $script:LedgerOpened = $raw.opened
    return @{ raw = $raw; tmp = $tmp; existed = $existed }
}

function Get-LedgerRowProblems([string]$Where, $Header, $Row, $Want) {
    # 지금 판 _row_problems: 쓴 값마다(장부에 열이 없는 값은 쓰지 않았으니 건너뜀) 읽은 칸과 견준다.
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($name in $Want.Keys) {
        $at = [array]::IndexOf([object[]]$Header, [object][string]$name)
        if ($at -lt 0) { continue }
        $got = if ($at -lt $Row.Count) { ConvertTo-SavedValue $Row[$at] } else { $null }
        $value = $Want[$name]
        if ($value -is [string]) { $value = Remove-IllegalText $value }
        if (-not (Test-SameValue $value $got)) {
            $problems.Add(('{0} {1}: 쓴 값 {2}, 읽은 값 {3}' -f $Where, $name, (Format-ReadBackText $value), (Format-ReadBackText $got -Saved)))
        }
    }
    return ,$problems
}

function Get-LedgerProblems($Raw, $Doc, $Number, $DocValues, $ItemValues) {
    # 지금 판 job_run _ledger_problems: 저장한 장부에서 이번 문서의 줄(문서종류 시트 한 줄 · 품목 시트 줄들)을 찾아 쓴 값과 견준다.
    $problems = New-Object System.Collections.Generic.List[string]
    $name = [string]$Doc.name
    if ($null -eq $Raw.doc -or $null -eq $Raw.items) {
        $problems.Add("누적장부에 '$name' 또는 '$($script:LedgerItemsSheet)' 시트가 없습니다"); return ,$problems
    }
    $numberText = ConvertTo-FormatText $Number
    $header = [object[]]$Raw.doc[0]
    $col = [array]::IndexOf($header, [object][string]$Doc.number_field)
    if ($col -lt 0) { $problems.Add("누적장부 $name 시트에 '$($Doc.number_field)' 열이 없습니다"); return ,$problems }
    $mine = @()
    for ($r = 1; $r -lt $Raw.doc.Count; $r++) {
        $row = $Raw.doc[$r]
        if ($col -lt $row.Count -and $null -ne $row[$col] -and (ConvertTo-SavedText $row[$col]) -ceq $numberText) { $mine += ,$row }
    }
    if ($mine.Count -ne 1) { $problems.Add("누적장부 $name 시트에 번호 $numberText 줄이 $($mine.Count)개입니다"); return ,$problems }
    foreach ($p in (Get-LedgerRowProblems "누적장부 $name" $header $mine[0] $DocValues)) { $problems.Add($p) }

    $header = [object[]]$Raw.items[0]
    $kind = [array]::IndexOf($header, [object]'문서종류')
    $num = [array]::IndexOf($header, [object]'문서번호')
    if ($kind -lt 0 -or $num -lt 0) { $problems.Add("누적장부 $($script:LedgerItemsSheet) 시트에 문서종류 · 문서번호 열이 없습니다"); return ,$problems }
    $mine = @()
    for ($r = 1; $r -lt $Raw.items.Count; $r++) {
        $row = $Raw.items[$r]
        if ([string]$row[$kind] -ceq $name -and $row[$kind] -is [string] -and $null -ne $row[$num] -and (ConvertTo-SavedText $row[$num]) -ceq $numberText) { $mine += ,$row }
    }
    if ($mine.Count -ne $ItemValues.Count) {
        $problems.Add("누적장부 품목 시트에 $numberText 줄이 $($mine.Count)개입니다(쓴 줄 $($ItemValues.Count)개)"); return ,$problems
    }
    for ($i = 0; $i -lt $mine.Count; $i++) {
        foreach ($p in (Get-LedgerRowProblems "누적장부 품목 $($i + 1)줄" $header $mine[$i] $ItemValues[$i])) { $problems.Add($p) }
    }
    return ,$problems
}

function New-LedgerStep([string]$Path, [bool]$Existed, [string]$Tmp, $Rules, $Entry, $Logs) {
    return @{ kind = 'ledger'; path = $Path; exists = $Existed; tmp = $Tmp; layout = (Get-LedgerLayout $Rules)
              entry = $Entry; logs = [object[]]$Logs }
}

function Add-LedgerEntry([string]$LedgerPath, $Rules, $Result, [string]$SourceName, [datetime]$When, [string]$OrganizedPath) {
    # 누적장부에 문서 한 건(지금 판 Ledger.open → is_duplicate → add → save → job_run _ledger_problems):
    #   문서종류 시트 한 줄(처리일시 · 원본파일 · 필드 · 정리된경로) + 품목 시트 줄들(문서종류 · 문서번호 + 표 열). 엑셀은 한 번만 띄운다.
    # 결과 [ordered]@{ locked; duplicate; saved; problems; path; created; backup }.
    #   locked    장부가 열려 있음(아무것도 바꾸지 않음)        duplicate 그 문서종류 · 번호가 이미 있음(바꾸지 않음)
    #   problems  저장한 판을 다시 읽어 보니 다른 칸(장부는 바꾸지 않음)
    #   saved     바꿈. created = 새로 만든 장부, backup = 바로 전 판(.bak.xlsx) 경로. 되돌리려면 Undo-LedgerEntry.
    $outcome = New-LedgerOutcome $LedgerPath
    if (Test-FileLocked $LedgerPath) { $outcome.locked = $true; return $outcome }
    Assert-LedgerNotLost $LedgerPath
    $result = ConvertTo-WriterResult $Rules $Result
    $doc = Get-WriterDoc $Rules $result.doc_type
    if ($When -eq [datetime]::MinValue) { $When = [datetime]::Now; $When = $When.AddTicks(- ($When.Ticks % [TimeSpan]::TicksPerSecond)) }
    $number = $result.fields[[string]$doc.number_field]
    $docValues = New-OrderedMap
    $docValues['처리일시'] = $When
    $docValues['원본파일'] = $SourceName
    if (-not [string]::IsNullOrEmpty($OrganizedPath)) { $docValues['정리된경로'] = $OrganizedPath }
    foreach ($key in $result.fields.Keys) { $docValues[[string]$key] = $result.fields[$key] }
    $itemValues = New-Object System.Collections.Generic.List[object]
    if ($null -ne $result.table) {
        foreach ($row in $result.table) {
            $values = New-OrderedMap
            $values['문서종류'] = [string]$doc.name
            $values['문서번호'] = $number
            foreach ($key in $row.Keys) { $values[[string]$key] = $row[$key] }
            $itemValues.Add($values)
        }
    }
    $itemRows = New-Object System.Collections.Generic.List[object]
    foreach ($values in $itemValues) { $itemRows.Add((ConvertTo-LedgerCells $values)) }
    $entry = @{ doc = [string]$doc.name; check_duplicate = $true; number_column = [string]$doc.number_field; number = (ConvertTo-FormatText $number)
                doc_cells = (ConvertTo-LedgerCells $docValues @('처리일시')); item_rows = $itemRows.ToArray() }
    $made = Invoke-LedgerStep $LedgerPath $Rules $entry @()
    $raw = $made.raw
    if ($raw.duplicate) { Remove-LedgerTmp $made.tmp; $outcome.duplicate = $true; return $outcome }
    $problems = Get-LedgerProblems $raw $doc $number $docValues $itemValues
    if ($problems.Count -gt 0) { Remove-LedgerTmp $made.tmp; $outcome.problems = [string[]]$problems.ToArray(); return $outcome }
    Complete-LedgerSave $outcome $made.tmp $made.existed
    return $outcome
}

function Add-LedgerLog([string]$LedgerPath, $Rules, $Entries) {
    # 처리로그 시트에 줄들(지금 판 Ledger.add_log: 일시 · 파일 · 결과 · 사유, 빈 사유는 빈칸)을 한 번에 넣고 저장(.bak 포함). 엑셀은 한 번만 띄운다.
    # $Entries = @(@{ when = [datetime]; file; outcome; reason } …). 결과는 Add-LedgerEntry 와 같은 모양(duplicate · problems 는 늘 비어 있음).
    $outcome = New-LedgerOutcome $LedgerPath
    if (Test-FileLocked $LedgerPath) { $outcome.locked = $true; return $outcome }
    Assert-LedgerNotLost $LedgerPath
    $logs = New-Object System.Collections.Generic.List[object]
    foreach ($e in (ConvertTo-ItemArray $Entries)) {
        if ($null -eq $e) { continue }
        $values = New-OrderedMap
        $values['일시'] = [datetime](Get-MarkupProp $e 'when')
        $values['파일'] = [string](Get-MarkupProp $e 'file')
        $values['결과'] = [string](Get-MarkupProp $e 'outcome')
        $reason = Get-MarkupProp $e 'reason'
        $values['사유'] = if ([string]::IsNullOrEmpty([string]$reason)) { $null } else { [string]$reason }
        $logs.Add((ConvertTo-LedgerCells $values @('일시')))
    }
    $made = Invoke-LedgerStep $LedgerPath $Rules $null $logs.ToArray()
    Complete-LedgerSave $outcome $made.tmp $made.existed
    return $outcome
}

function Undo-LedgerEntry($Added) {
    # Add-LedgerEntry 가 바꾼 것을 되돌린다(엑셀 없이): 새로 만든 장부는 지우고, 있던 장부는 바로 전 판(.bak)으로 되돌린다.
    # 지금 판 job_run _rollback(줄 지우고 백업 없이 저장)과 결과가 같다: 장부 = 이번 문서를 넣기 전, .bak = 그 판.
    if (-not [bool](Get-MarkupProp $Added 'saved')) { return }
    $path = [string](Get-MarkupProp $Added 'path')
    if (Test-FileLocked $path) { throw "누적장부가 열려 있어 되돌리지 못함: $path" }
    if ([bool](Get-MarkupProp $Added 'created')) {
        [IO.File]::Delete($path)
        return
    }
    $backup = [string](Get-MarkupProp $Added 'backup')
    if ([string]::IsNullOrEmpty($backup) -or -not [IO.File]::Exists($backup)) { throw "되돌릴 바로 전 판이 없습니다: $backup" }
    [IO.File]::Copy($backup, $path, $true)
}

function Test-LedgerDuplicate($LedgerPath, $DocType, $Number) {
    # 장부의 그 문서종류 시트에 같은 번호가 있는지(지금 판 is_duplicate). $DocType = 규칙의 문서종류 객체(name · number_field), $Number = 번호 값.
    # 장부 · 시트 · 번호 열이 없으면 $false. 장부는 임시 복사본을 읽는다(열려 있어도 됨).
    if (-not [IO.File]::Exists([string]$LedgerPath)) { return $false }
    $copy = New-OfficeCopy ([string]$LedgerPath)
    try {
        $raw = (Invoke-ExcelWrite @(@{ kind = 'dup'; path = $copy.Path; sheet = [string]$DocType.name; column = [string]$DocType.number_field
                                       number = (ConvertTo-FormatText $Number) }))[0]
    } finally {
        Remove-OfficeCopy $copy
    }
    return [bool]$raw.duplicate
}

# ---- 엔진 op --------------------------------------------------------------------------------------

function ConvertTo-OpDate([string]$Text) {
    return [datetime]::ParseExact($Text, 'yyyy-MM-ddTHH:mm:ss', $script:Invariant)
}

function ConvertTo-OpJson($Map) {
    # 결과 표 + forced(이 실행이 띄운 엑셀 가운데 Quit 뒤 강제로 끝낸 수, 정상 0) → JSON.
    $out = [ordered]@{}
    foreach ($key in $Map.Keys) { $out[$key] = $Map[$key] }
    $out['forced'] = $script:OfficeForced
    $out['left'] = $null      # 남의 문서가 있어 엑셀을 끝내지 않고 넘겨 준 기록(Office.ps1, 정상 $null)
    if ($script:OfficeHandovers.Count -gt 0) { $out['left'] = [string]::Join("`n", $script:OfficeHandovers.ToArray()) }
    return ConvertTo-Json -InputObject $out -Depth 20 -Compress
}

function Invoke-WriterOp($Payload) {
    # 엔진 op 'office_write'(시험용 입구). action:
    #   company_form {folder, rules, result, markup_grid?, exact?}          → {path, forced}
    #   check_file   {folder, rules, result, checks, source, company: bool, markup_grid?} → {path, forced}
    #   ledger_add   {ledger, rules, result, source, when, organized?}       → {locked, duplicate, saved, problems, path, created, backup, opened, forced}
    #   ledger_log   {ledger, rules, logs: [{when, file, outcome, reason}]}  → 같은 모양
    #   ledger_dup   {ledger, rules, doc_type, number}                       → {duplicate, forced}
    #   undo_ledger  {added: ledger_add 의 결과}                             → {}
    #   read_back    {path, sheet, header, rows: [[칸 JSON | null]]}         → {problems, forced, left}
    #   locked       {path}                                                  → {locked}
    # 결과 JSON 은 작업 2 · 3 의 결과, 칸 JSON 은 작업 6 의 {t, v}. 일시는 'YYYY-MM-DDTHH:MM:SS'.
    # 어느 action 이든 timeout(초)을 주면 엑셀 한 번에 주는 시간을 바꾼다(시험에서 시간 초과를 흉내 낼 때).
    $action = [string](Get-MarkupProp $Payload 'action')
    $timeout = Get-MarkupProp $Payload 'timeout'
    if ($null -ne $timeout) { $script:OfficeTimeoutSec = [int]$timeout }
    # 시험용(Office.ps1 의 남의 문서 보호): test_preopen · test_extra 는 처음 띄우는 엑셀에서 같은 인스턴스로 열 파일
    $hooks = @{}
    foreach ($key in @('preopen', 'extra')) { $value = Get-MarkupProp $Payload ('test_' + $key); if ($value) { $hooks[$key] = [string]$value } }
    if ($hooks.Count -gt 0) { $script:OfficeTestHooks = $hooks }
    $script:OfficeHandovers.Clear()
    if ($action -ceq 'locked') {
        return ConvertTo-Json -InputObject ([ordered]@{ locked = (Test-FileLocked ([string]$Payload.path)) }) -Depth 20 -Compress
    }
    if ($action -ceq 'undo_ledger') {
        Undo-LedgerEntry $Payload.added
        return '{}'
    }
    if ($action -ceq 'read_back') {
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($row in @($Payload.rows)) {
            $cells = @($row)
            $line = New-Object 'object[]' $cells.Count
            for ($c = 0; $c -lt $cells.Count; $c++) { $line[$c] = ConvertFrom-TypedCell $cells[$c] }
            $rows.Add($line)
        }
        $problems = Test-ReadBack ([string]$Payload.path) ([string]$Payload.sheet) ([object[]]@($Payload.header)) $rows
        return ConvertTo-OpJson ([ordered]@{ problems = [string[]]$problems.ToArray() })
    }
    $rules = ConvertTo-PortableRules (Get-MarkupProp $Payload 'rules')
    $markup = $null
    $grid = Get-MarkupProp $Payload 'markup_grid'
    if ($null -ne $grid) { $markup = Read-MarkupGrid $grid }
    switch ($action) {
        'company_form' {
            $form = $rules.company_form
            if ($null -eq $form) { throw '규칙에 회사양식이 없습니다.' }
            $result = ConvertTo-WriterResult $rules $Payload.result
            $path = Write-CompanyForm ([string]$Payload.folder) $result (Get-CompanyRows $result $form $markup) $form -Exact:([bool](Get-MarkupProp $Payload 'exact'))
            return ConvertTo-OpJson ([ordered]@{ path = $path })
        }
        'check_file' {
            $result = ConvertTo-WriterResult $rules $Payload.result
            $company = $null
            if ([bool]$Payload.company) { $company = Get-CompanyRows $result $rules.company_form $markup }
            $path = Write-CheckFile ([string]$Payload.folder) $result @($Payload.checks) ([string]$Payload.source) $company $rules
            return ConvertTo-OpJson ([ordered]@{ path = $path })
        }
        'ledger_add' {
            $organized = [string](Get-MarkupProp $Payload 'organized')
            $out = Add-LedgerEntry ([string]$Payload.ledger) $rules $Payload.result ([string]$Payload.source) (ConvertTo-OpDate ([string]$Payload.when)) $organized
            $out['opened'] = $script:LedgerOpened      # 시험용: 엑셀이 연 파일(새 장부면 $null)
            return ConvertTo-OpJson $out
        }
        'ledger_log' {
            $entries = @(@($Payload.logs) | ForEach-Object {
                @{ when = (ConvertTo-OpDate ([string]$_.when)); file = $_.file; outcome = $_.outcome; reason = $_.reason } })
            return ConvertTo-OpJson (Add-LedgerLog ([string]$Payload.ledger) $rules $entries)
        }
        'ledger_dup' {
            $doc = Get-WriterDoc $rules ([string]$Payload.doc_type)
            return ConvertTo-OpJson ([ordered]@{ duplicate = (Test-LedgerDuplicate ([string]$Payload.ledger) $doc ([string]$Payload.number)) })
        }
    }
    throw "office_write action '$action' 는 없습니다."
}
