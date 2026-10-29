# 검산: 결과 값이 원본에 그대로 있는지, 품목 줄 · 합계가 맞는지 확인한다. 동작 기준은 app\verify.py (항목 이름 · 값 글자 · 순서까지 같게).
# 틀린 숫자가 조용히 통과하지 않는 것이 우선이다. 애매하면 불일치로 본다.
# 점-소싱 순서: Values → Rules → ExtractText → ExtractGrid → Verify (칸 읽기 ConvertTo-GridCell 을 ExtractGrid 에서 쓴다).
#
# 값의 종류(지금 판 파이썬 값과 짝): [decimal]=Decimal(금액) · [long]=int · [double]=float · [datetime]=date(0시) 또는 datetime(시각 있음)
#   · [string]=str · [bool]=bool · $null=None.
# 결과 객체는 두 가지가 들어온다: PS 안에서 만든 것([ordered] 해시표, 값이 위 종류 그대로)과
#   JSON 에서 읽은 것(PSCustomObject, 금액 · 날짜는 글자라서 규칙의 형식으로 되살린다).
# 반환 목록의 항목: [ordered]@{ item; source_value; result_value; ok } — ok 는 $true / $false / $null(해당 없음).

# 원본 글자에서 숫자 · 날짜 토큰을 찾는 정규식(app\verify.py 의 NUMBER · DATE 와 같은 글자). DATE 는 겹쳐서 찾는다.
$script:NumberToken = New-Object Text.RegularExpressions.Regex('-?\d[\d,]*(?:\.\d+)?')
$script:DateToken = New-Object Text.RegularExpressions.Regex('(?=(?<!\d)(\d{1,4}[/\-]\d{1,2}[/\-]\d{1,4})(?!\d))')
$script:NotFound = '원본에 없음'
# 엑셀은 100.00 을 정수 100 으로 저장해 금액의 소수 자리를 알 수 없다. 줄 검산은 최소 이 자리까지 반올림해 비교한다
$script:ExcelMinPlaces = 2
# 번호 앞뒤에 이 글자가 붙어 있으면 더 긴 번호의 일부로 본다(INV-2026-1023 은 INV-2026-10231 의 일부)
$script:IdChars = 'A-Za-z0-9_\-/'
$script:IsoDate = New-Object Text.RegularExpressions.Regex('^[0-9]{4}-[0-9]{2}-[0-9]{2}$')
$script:IsoDateTime = New-Object Text.RegularExpressions.Regex('^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}$')
$script:FloatParts = New-Object Text.RegularExpressions.Regex('^(-?)([0-9]+)(?:\.([0-9]+))?(?:[Ee]([+-]?[0-9]+))?$')

# ---- 값 ------------------------------------------------------------------------------------------

function ConvertTo-CheckDecimal($Value) {
    # 지금 판 _decimal: 숫자로 볼 수 있는 값을 [decimal] 로. 실수(엑셀 칸)는 엑셀처럼 유효 15자리. 못 바꾸면 $null.
    if ($null -eq $Value -or $Value -is [bool]) { return $null }
    if ($Value -is [decimal]) { return $Value }
    if ($Value -is [int] -or $Value -is [long]) { return [decimal]$Value }
    if ($Value -is [double]) { return ConvertTo-ExcelNumber $Value }
    if ($Value -is [string]) {
        if ($Value.Trim() -eq '') { return $null }
        return ConvertTo-Amount $Value
    }
    return $null
}

function ConvertTo-FloatText([double]$Value) {
    # 파이썬 str(float) 와 같은 글자: 가장 짧은 왕복 글자, 정수 같은 값은 5.0, 지수 표기는 소수 자리 -4 이하 또는 16 자리 초과일 때 1e-05 · 1e+16.
    if ([double]::IsNaN($Value)) { return 'nan' }
    if ([double]::IsInfinity($Value)) { if ($Value -gt 0) { return 'inf' } else { return '-inf' } }
    if ($Value -eq 0) { if ([BitConverter]::DoubleToInt64Bits($Value) -lt 0) { return '-0.0' } else { return '0.0' } }
    # 가장 짧은 왕복 글자: 유효 자리 1~17 중 처음으로 같은 double 로 돌아오는 것(R 은 조금 안 짧거나 15자리로 끝나는 값이 있다)
    $text = $Value.ToString('G17', $script:Invariant)
    for ($digitsWanted = 1; $digitsWanted -le 17; $digitsWanted++) {
        $candidate = $Value.ToString("G$digitsWanted", $script:Invariant)
        try {
            if ([double]::Parse($candidate, $script:Invariant) -eq $Value) { $text = $candidate; break }
        } catch {
            continue   # 반올림하면 double 범위를 넘는 글자(1.797693134862316E+308)
        }
    }
    $m = $script:FloatParts.Match($text)
    $sign = $m.Groups[1].Value
    $integer = $m.Groups[2].Value
    $fraction = $m.Groups[3].Value
    $exponent = 0
    if ($m.Groups[4].Success) { $exponent = [int]$m.Groups[4].Value }
    $digits = $integer + $fraction
    $point = $integer.Length + $exponent        # 값 = 0.DIGITS x 10^point
    $lead = $digits.Length - $digits.TrimStart('0').Length
    $digits = $digits.Substring($lead)
    $point -= $lead
    $digits = $digits.TrimEnd('0')
    if ($point -le -4 -or $point -gt 16) {
        $mantissa = $digits.Substring(0, 1)
        if ($digits.Length -gt 1) { $mantissa += '.' + $digits.Substring(1) }
        $power = $point - 1
        $powerSign = if ($power -lt 0) { '-' } else { '+' }
        return $sign + $mantissa + 'e' + $powerSign + ([Math]::Abs($power)).ToString('00', $script:Invariant)
    }
    if ($point -le 0) { return $sign + '0.' + ('0' * (- $point)) + $digits }
    if ($digits.Length -le $point) { return $sign + $digits + ('0' * ($point - $digits.Length)) + '.0' }
    return $sign + $digits.Substring(0, $point) + '.' + $digits.Substring($point)
}

function Show-CheckValue($Value) {
    # 지금 판 _show.
    if ($null -eq $Value) { return '' }
    if ($Value -is [decimal]) { return $Value.ToString($script:Invariant) }
    if ($Value -is [datetime]) {
        if ($Value.TimeOfDay.Ticks -ne 0) { return $Value.ToString('yyyy-MM-ddTHH:mm:ss', $script:Invariant) }
        return $Value.ToString('yyyy-MM-dd', $script:Invariant)
    }
    if ($Value -is [double]) { return ConvertTo-FloatText $Value }
    if ($Value -is [bool]) { if ($Value) { return 'True' } else { return 'False' } }
    if ($Value -is [string]) { return $Value }
    return [string]::Format($script:Invariant, '{0}', $Value)
}

function Get-Places([decimal]$Number) {
    # 지금 판 _places: 소수 자리 수(1234.50 → 2).
    return [int](([decimal]::GetBits($Number)[3] -shr 16) -band 0xFF)
}

function Test-Comparable($Value) {
    # 지금 판 _comparable: 숫자 · 날짜(참거짓 제외).
    if ($null -eq $Value -or $Value -is [bool]) { return $false }
    return ($Value -is [decimal] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [datetime])
}

# ---- 결과 객체 → 값 종류가 정해진 모양 ----------------------------------------------------------------

function ConvertFrom-IsoDate([string]$Text) {
    # 서식 없는 엑셀 칸이 날짜였던 경우의 JSON 글자("YYYY-MM-DD" 또는 "YYYY-MM-DDTHH:MM:SS")를 [datetime] 으로. 아니면 $null.
    if ($script:IsoDate.IsMatch($Text)) { return ConvertTo-DateValue $Text }
    if ($script:IsoDateTime.IsMatch($Text)) {
        try { return [datetime]::ParseExact($Text, 'yyyy-MM-ddTHH:mm:ss', $script:Invariant) } catch { return $null }
    }
    return $null
}

function ConvertFrom-ResultValue($Value, $Format, [bool]$GuessDate) {
    # JSON 에서 읽은 값 하나를 되살린다. 금액 · 날짜 형식인 칸의 글자는 [decimal] · [datetime](자리 그대로),
    # JSON 숫자는 정수면 [long], 소수점이 있으면 실수([double]). 서식 없는 엑셀 칸의 날짜 글자는 날짜로.
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) {
        if ($Format -ceq '금액') {
            $amount = ConvertTo-Amount $Value
            if ($null -ne $amount) { return $amount }
        } elseif ($Format -ceq '날짜') {
            $day = ConvertTo-DateValue $Value
            if ($null -ne $day) { return $day }
        } elseif ($GuessDate) {
            $day = ConvertFrom-IsoDate $Value
            if ($null -ne $day) { return $day }
        }
        return $Value
    }
    if ($Value -is [bool] -or $Value -is [double] -or $Value -is [datetime]) { return $Value }
    if ($Value -is [int] -or $Value -is [long]) { return [long]$Value }
    if ($Value -is [decimal]) {
        # 5.1 은 소수점이 있는 JSON 숫자를 decimal 로 준다(5.0 은 자리 1). 자리가 0 이면 큰 정수다.
        if ((Get-Places $Value) -eq 0 -and [Math]::Abs($Value) -le [long]::MaxValue) { return [long]$Value }
        return [double]::Parse($Value.ToString($script:Invariant), $script:Invariant)
    }
    return $Value
}

function Get-MapEntries($Map) {
    # 해시표(PS 안) 또는 PSCustomObject(JSON)의 (이름, 값) 목록. 순서 유지.
    $entries = New-Object 'System.Collections.Generic.List[object]'
    if ($null -eq $Map) { return ,$entries }
    if ($Map -is [System.Collections.IDictionary]) {
        foreach ($key in $Map.Keys) { $entries.Add(@{ Name = [string]$key; Value = $Map[$key] }) }
    } else {
        foreach ($property in $Map.PSObject.Properties) { $entries.Add(@{ Name = [string]$property.Name; Value = $property.Value }) }
    }
    return ,$entries
}

function Get-ResultMember($Result, [string]$Name) {
    if ($Result -is [System.Collections.IDictionary]) {
        if ($Result.Contains($Name)) { return ,$Result[$Name] }
        return $null
    }
    $property = $Result.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return ,$property.Value   # 쉼표: 빈 배열(빈 표 [])이 $null 로 풀리지 않게
}

function New-NameSet($Names) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    foreach ($name in @($Names)) { if ($null -ne $name) { [void]$set.Add([string]$name) } }
    return ,$set
}

function ConvertTo-CheckModel($Rules, $Result, $Doc) {
    # 결과의 fields · table 을 값 종류가 정해진 Ordinal OrderedDictionary 들로. 반환: @{ Fields; Table($null 또는 List); IsExcel }.
    $memory = $Result -is [System.Collections.IDictionary]
    $isExcel = [bool](Get-ResultMember $Result 'is_excel')
    $formats = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
    $amountColumns = New-NameSet @()
    if (-not $memory) {
        # JSON 이면 형식(금액 · 날짜)과 금액열을 규칙에서 찾는다. 엑셀 PO 는 엑셀양식의 머리 · 품목, PDF 는 문서종류의 필드 · 표.
        $form = $null
        if ($isExcel) {
            $formName = Get-ResultMember $Result 'form'
            $form = @($Rules.excel_forms) | Where-Object { $_.name -ceq $formName } | Select-Object -First 1
        }
        if ($null -ne $form) {
            foreach ($head in @($form.head)) { if ($null -ne $head.format) { $formats[[string]$head.name] = [string]$head.format } }
            if ($null -ne $form.items) { $amountColumns = New-NameSet $form.items.amount_columns }
        } else {
            foreach ($field in @($Doc.fields)) { if ($null -ne $field.format) { $formats[[string]$field.name] = [string]$field.format } }
            if ($null -ne $Doc.table) { $amountColumns = New-NameSet $Doc.table.amount_columns }
        }
    }

    $fields = New-OrderedMap
    foreach ($entry in (Get-MapEntries (Get-ResultMember $Result 'fields'))) {
        $format = $null
        if ($formats.ContainsKey($entry.Name)) { $format = $formats[$entry.Name] }
        # 번호필드의 글자는 날짜 모양이어도 글자로 둔다(PO 번호 "2026-10-26" 이 날짜로 되살아나 칸 하나와의 완전 일치 대신 글자 속 날짜 대조를 하면 잘린 번호가 통과한다)
        $guessDate = $isExcel -and ($entry.Name -cne [string]$Doc.number_field)
        $fields[$entry.Name] = if ($memory) { $entry.Value } else { ConvertFrom-ResultValue $entry.Value $format $guessDate }
    }
    $table = $null
    $rawTable = Get-ResultMember $Result 'table'
    if ($null -ne $rawTable) {
        $table = New-Object 'System.Collections.Generic.List[object]'
        foreach ($rawRow in @($rawTable)) {
            $row = New-OrderedMap
            foreach ($entry in (Get-MapEntries $rawRow)) {
                $format = $null
                if ($amountColumns.Contains($entry.Name)) { $format = '금액' }
                $row[$entry.Name] = if ($memory) { $entry.Value } else { ConvertFrom-ResultValue $entry.Value $format $isExcel }
            }
            $table.Add($row)
        }
    }
    return @{ Fields = $fields; Table = $table; IsExcel = $isExcel }
}

# ---- 원본 다시 읽기 ---------------------------------------------------------------------------------

function Add-SourceToken($Bucket, $Key, [string]$Raw) {
    if (-not $Bucket.ContainsKey($Key)) { $Bucket[$Key] = New-Object 'System.Collections.Generic.List[string]' }
    if (-not $Bucket[$Key].Contains($Raw)) { $Bucket[$Key].Add($Raw) }
}

function Get-DateKey([datetime]$Value) {
    # 날짜(0시)와 일시(시각 있음)는 서로 다른 값이다(파이썬 date != datetime).
    if ($Value.TimeOfDay.Ticks -eq 0) { return 'D' + $Value.ToString('yyyy-MM-dd', $script:Invariant) }
    return 'T' + $Value.ToString('yyyy-MM-ddTHH:mm:ss', $script:Invariant)
}

function New-VerifySource([bool]$FromGrid) {
    $cells = $null
    if ($FromGrid) { $cells = New-NameSet @() }   # 엑셀: 글자 칸 값(앞뒤 공백 뺌). 번호필드는 칸 하나와 같아야 함
    return @{
        Parts     = New-Object 'System.Collections.Generic.List[string]'
        Numbers   = New-Object 'System.Collections.Generic.Dictionary[decimal,System.Collections.Generic.List[string]]'
        Dates     = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[string]]'
        HasPlaces = -not $FromGrid     # PDF 글자만 소수 자리 정보가 있다
        Cells     = $cells
    }
}

function Add-ScanText($Source, [string]$Text) {
    # 글자에서 숫자 · 날짜 토큰을 모은다. 숫자는 쉼표 · $ 를 뺀 [decimal](소수 자리 유지)로, 음수는 부호 없는 값도 함께.
    $Source.Parts.Add($Text)
    foreach ($match in $script:NumberToken.Matches($Text)) {
        $raw = $match.Value.TrimEnd(',')
        $number = ConvertTo-Amount $raw
        if ($null -eq $number) { continue }
        Add-SourceToken $Source.Numbers $number $raw
        if ($raw.StartsWith('-')) { Add-SourceToken $Source.Numbers (- $number) $raw.Substring(1) }
    }
    foreach ($match in $script:DateToken.Matches($Text)) {
        $raw = $match.Groups[1].Value
        $day = ConvertTo-DateValue $raw
        if ($null -ne $day) { Add-SourceToken $Source.Dates (Get-DateKey $day) $raw }
    }
}

function Add-GridCell($Source, $Raw) {
    # 지금 판 _add_cell. 격자 칸 {t, v}: 글자는 칸 값 + 글자 속 토큰, 날짜는 날짜 열쇠, 숫자는 엑셀 숫자. 참거짓은 건너뜀.
    $cell = ConvertTo-GridCell $Raw
    switch ($cell.t) {
        's' {
            $text = [string]$cell.v
            [void]$Source.Cells.Add($text.Trim())
            Add-ScanText $Source $text
        }
        'd' {
            # 0시면 날짜(date), 시각이 있으면 일시(datetime). 원본 글자는 지금 판 isoformat 과 같게.
            $day = [datetime]$cell.v
            $token = if ($day.TimeOfDay.Ticks -eq 0) { $day.ToString('yyyy-MM-dd', $script:Invariant) } else { $day.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant) }
            Add-SourceToken $Source.Dates (Get-DateKey $day) $token
        }
        'n' {
            $number = ConvertTo-ExcelNumber $cell.v
            if ($null -ne $number) { Add-SourceToken $Source.Numbers $number (Show-CheckValue $number) }
        }
    }
}

function ConvertTo-VerifySource($Source) {
    if ($Source.kind -ceq 'text') {
        $origin = New-VerifySource $false
        Add-ScanText $origin ([string]$Source.text)
        return $origin
    }
    if ($Source.kind -ceq 'grid') {
        $origin = New-VerifySource $true
        foreach ($sheet in @($Source.grid.sheets)) {
            if ($null -eq $sheet.cells) { continue }
            foreach ($item in $sheet.cells) { Add-GridCell $origin $item[2] }
        }
        return $origin
    }
    throw "원본 종류 '$($Source.kind)' 를 모릅니다(text 또는 grid)."
}

# ---- 1 · 2. 원문 대조와 소수 자리 -------------------------------------------------------------------

function New-Check([string]$Item, [string]$SourceValue, [string]$ResultValue, $Ok) {
    return [ordered]@{ item = $Item; source_value = $SourceValue; result_value = $ResultValue; ok = $Ok }
}

function Test-NumberText($Source, [string]$Value) {
    # 번호필드 글자가 원본에 그대로, 더 긴 번호의 일부가 아니게 있는가.
    if ($Value.Trim() -eq '') { return $false }
    if ($null -ne $Source.Cells) { return $Source.Cells.Contains($Value.Trim()) }   # 엑셀: 칸 하나의 값과 같아야 함
    # 앞: 번호 글자나 '숫자.' 뒤면 일부(1.2026). 'No.INV-1' 처럼 글자 뒤 마침표는 괜찮음. 뒤: 문장 끝 마침표는 괜찮음
    $pattern = '(?<![' + $script:IdChars + '])(?<!\d\.)' + [regex]::Escape($Value) + '(?![' + $script:IdChars + ']|\.[A-Za-z0-9])'
    return [regex]::IsMatch(($Source.Parts -join "`n"), $pattern, [Text.RegularExpressions.RegexOptions]::CultureInvariant)
}

function Get-CheckTargets($Doc, $Model) {
    # 대조할 (항목 이름, 값). 숫자 · 날짜 값과 번호필드.
    $targets = New-Object 'System.Collections.Generic.List[object]'
    foreach ($name in $Model.Fields.Keys) {
        $value = $Model.Fields[$name]
        if ($null -ne $value -and ($name -ceq [string]$Doc.number_field -or (Test-Comparable $value))) {
            $targets.Add(@{ Item = $name; Value = $value })
        }
    }
    if ($null -ne $Model.Table) {
        $n = 0
        foreach ($row in $Model.Table) {
            $n++
            foreach ($column in $row.Keys) {
                if (Test-Comparable $row[$column]) { $targets.Add(@{ Item = "품목 ${n}줄 $column"; Value = $row[$column] }) }
            }
        }
    }
    return ,$targets
}

function Get-SourceChecks($Doc, $Model, $Source, $Checks) {
    foreach ($target in (Get-CheckTargets $Doc $Model)) {
        $item = $target.Item
        $value = $target.Value
        if ($value -is [string]) {
            $ok = Test-NumberText $Source $value
            $shown = if ($ok) { $value } else { $script:NotFound }
            $Checks.Add((New-Check $item $shown $value $ok))
        } elseif ($value -is [datetime]) {
            $tokens = $null
            $key = Get-DateKey $value
            if ($Source.Dates.ContainsKey($key)) { $tokens = $Source.Dates[$key] }
            $found = $null -ne $tokens -and $tokens.Count -gt 0
            $Checks.Add((New-Check $item $(if ($found) { $tokens[0] } else { $script:NotFound }) (Show-CheckValue $value) $found))
        } else {
            $number = ConvertTo-CheckDecimal $value
            $tokens = $null
            if ($null -ne $number -and $Source.Numbers.ContainsKey($number)) { $tokens = $Source.Numbers[$number] }
            $found = $null -ne $tokens -and $tokens.Count -gt 0
            $Checks.Add((New-Check $item $(if ($found) { $tokens[0] } else { $script:NotFound }) (Show-CheckValue $value) $found))
            if ($found -and $Source.HasPlaces -and $value -is [decimal]) {
                $Checks.Add((Get-PlacesCheck $item $value $tokens))
            }
        }
    }
}

function Get-PlacesCheck([string]$Item, [decimal]$Value, $Tokens) {
    # 결과 소수 자리 수 = 같은 값을 준 원본 토큰의 소수 자리 수(그런 토큰이 하나라도 있으면 통과).
    $want = Get-Places $Value
    $same = @($Tokens | Where-Object { (Get-Places (ConvertTo-Amount $_)) -eq $want })
    $shown = if ($same.Count -gt 0) { $same[0] } else { $Tokens[0] }
    return New-Check "$Item 소수 자리" $shown (Show-CheckValue $Value) ($same.Count -gt 0)
}

# ---- 3. 줄 검산 ------------------------------------------------------------------------------------

function Test-RoundsTo([decimal]$Product, [decimal]$Amount, [int]$MinPlaces) {
    # product 를 amount 의 소수 자리(최소 MinPlaces)로 반올림(사사오입 · 은행원식 중 하나)하면 amount 인가.
    $places = [Math]::Max((Get-Places $Amount), $MinPlaces)
    try {
        if ([Math]::Round($Product, $places, [MidpointRounding]::AwayFromZero) -eq $Amount) { return $true }
        return ([Math]::Round($Product, $places, [MidpointRounding]::ToEven) -eq $Amount)
    } catch {
        return $false
    }
}

function Get-LineChecks($Model, $Columns, $Checks) {
    $n = 0
    foreach ($row in $Model.Table) {
        $n++
        $item = "${n}줄 금액"
        $quantity = ConvertTo-CheckDecimal $row[[string]$Columns.quantity_column]
        $price = ConvertTo-CheckDecimal $row[[string]$Columns.price_column]
        $amount = ConvertTo-CheckDecimal $row[[string]$Columns.amount_column]
        if ($null -eq $quantity -or $null -eq $price -or $null -eq $amount) {
            $Checks.Add((New-Check $item '' (Show-CheckValue $row[[string]$Columns.amount_column]) $null))
            continue
        }
        $minPlaces = if ($Model.IsExcel) { $script:ExcelMinPlaces } else { 0 }
        try {
            $product = $quantity * $price
        } catch {
            $Checks.Add((New-Check $item '' (Show-CheckValue $amount) $false))   # decimal 로 담을 수 없는 큰 곱(금액과 같을 수 없다)
            continue
        }
        $Checks.Add((New-Check $item "$(Show-CheckValue $quantity) x $(Show-CheckValue $price) = $(Show-CheckValue $product)" `
                (Show-CheckValue $amount) (Test-RoundsTo $product $amount $minPlaces)))
    }
}

# ---- 4. 합계 검산 ----------------------------------------------------------------------------------

function Get-ItemSum($Model, [string]$Column) {
    # 품목 금액 합. 표가 없거나 비었거나 숫자가 아닌 칸이 있으면 $null.
    if ($null -eq $Model.Table -or $Model.Table.Count -eq 0) { return $null }
    $sum = [decimal]0
    foreach ($row in $Model.Table) {
        $amount = ConvertTo-CheckDecimal $row[$Column]
        if ($null -eq $amount) { return $null }
        $sum += $amount
    }
    return $sum
}

function New-EqualCheck([string]$Item, $Expected, [string]$SourceText, $Stated) {
    if ($null -eq $Expected -or $null -eq $Stated) { return New-Check $Item $SourceText (Show-CheckValue $Stated) $null }
    return New-Check $Item $SourceText (Show-CheckValue $Stated) ($Expected -eq $Stated)
}

function Get-TotalChecks($Model, $Spec, $Checks) {
    $itemSum = Get-ItemSum $Model ([string]$Spec.amount_column)
    $hasSubtotal = -not [string]::IsNullOrEmpty([string]$Spec.subtotal)
    $subtotal = $null
    if ($hasSubtotal) {
        $subtotal = ConvertTo-CheckDecimal $Model.Fields[[string]$Spec.subtotal]
        $Checks.Add((New-EqualCheck '품목 합 = 소계' $itemSum (Show-CheckValue $itemSum) $subtotal))
    }
    if ([string]::IsNullOrEmpty([string]$Spec.total)) { return }

    $total = ConvertTo-CheckDecimal $Model.Fields[[string]$Spec.total]
    $extraNames = @($Spec.extras)
    $name = $(if ($hasSubtotal) { '소계' } else { '품목 합' }) + $(if ($extraNames.Count -gt 0) { ' + 부대비용' } else { '' }) + ' = 합계'
    # 소계 값이 비면 품목 합으로 대신 계산한다(합계를 아무것도 대조하지 않는 일이 없게).
    $base = if ($null -ne $subtotal) { $subtotal } else { $itemSum }
    # 부대비용 칸이 비면 0 으로 본다. 실제로 있었다면 합계가 안 맞아 불일치로 잡힌다.
    $extras = New-Object 'System.Collections.Generic.List[decimal]'
    foreach ($extraName in $extraNames) {
        $extra = ConvertTo-CheckDecimal $Model.Fields[[string]$extraName]
        $extras.Add($(if ($null -eq $extra) { [decimal]0 } else { $extra }))
    }
    $computed = $null
    $parts = ''
    if ($null -ne $base) {
        $computed = $base
        foreach ($extra in $extras) { $computed += $extra }
        $shownParts = New-Object 'System.Collections.Generic.List[string]'
        $shownParts.Add((Show-CheckValue $base))
        foreach ($extra in $extras) { $shownParts.Add((Show-CheckValue $extra)) }
        $parts = [string]::Join(' + ', $shownParts)
    }
    $sourceText = if ($parts -ne '') { "$parts = $(Show-CheckValue $computed)" } else { '' }
    $Checks.Add((New-EqualCheck $name $computed $sourceText $total))
}

# ---- 입구 --------------------------------------------------------------------------------------------

function Get-Checks($Result, $Source, $Rules) {
    # 결과 하나를 원본과 대조한 검산 목록(List). 순서: 원문 대조 · 소수 자리 → 줄 검산 → 합계 검산.
    #   $Result: Get-TextExtract / Get-GridExtract 가 만든 객체 또는 그 JSON(ConvertFrom-Json)
    #   $Source: @{ kind = 'text'; text = <PDF 글자> } 또는 @{ kind = 'grid'; grid = <격자> } (JSON 에서 읽은 PSCustomObject 도 됨)
    #   $Rules : ConvertTo-PortableRules 를 거친 규칙(문서종류의 검산 · 번호필드와 엑셀양식의 형식을 찾는 데 쓴다)
    $checks = New-Object 'System.Collections.Generic.List[object]'
    $docName = Get-ResultMember $Result 'doc_type'
    if ($null -eq $docName) { return ,$checks }
    $doc = @($Rules.doc_types) | Where-Object { $_.name -ceq [string]$docName } | Select-Object -First 1
    if ($null -eq $doc) { throw "문서종류 '$docName' 가 규칙에 없습니다." }

    $model = ConvertTo-CheckModel $Rules $Result $doc
    $origin = ConvertTo-VerifySource $Source
    $spec = $doc.check
    $columns = if ($null -ne $spec) { $spec } else { [pscustomobject]@{ quantity_column = '수량'; price_column = '단가'; amount_column = '금액' } }

    Get-SourceChecks $doc $model $origin $checks
    if ($null -ne $model.Table) { Get-LineChecks $model $columns $checks }
    if ($null -ne $spec) { Get-TotalChecks $model $spec $checks }
    return ,$checks
}

function ConvertTo-ChecksJson($Checks) {
    # 검산 목록을 JSON 배열 글자로(항목이 0개 · 1개여도 배열).
    return ConvertTo-Json -InputObject ([object[]]@($Checks.ToArray())) -Depth 20
}
