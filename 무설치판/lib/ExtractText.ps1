# PDF 글자 → 문서종류 판별 → 필드 · 표 추출. 동작 기준은 app\extract.py 의 extract_from_text(글자까지 같게).
# 결과 객체(엔진 공통): doc_type · form · is_excel · fields · table · warnings · review_reason

$script:NoText = '텍스트 없음 (스캔본, OCR 범위 밖)'
# 파이썬 str.splitlines() 가 줄을 나누는 글자와 같다
$script:LineBreak = [regex]"`r`n|[`n`r`v`f\x1c\x1d\x1e\x85\u2028\u2029]"

function Split-TextLines([string]$Text) {
    # splitlines 처럼: 빈 글자는 줄 0개, 끝의 줄바꿈 뒤에는 빈 줄을 만들지 않는다.
    $lines = New-Object 'System.Collections.Generic.List[string]'
    if ([string]::IsNullOrEmpty($Text)) { return ,$lines }
    $lines.AddRange($script:LineBreak.Split($Text))
    if ($lines[$lines.Count - 1] -eq '') { $lines.RemoveAt($lines.Count - 1) }
    return ,$lines
}

function New-OrderedMap {
    # 파이썬 dict 처럼 이름의 대소문자를 가린다([ordered]@{} 는 가리지 않음).
    return ,(New-Object System.Collections.Specialized.OrderedDictionary([StringComparer]::Ordinal))
}

function New-ExtractResult {
    return [ordered]@{
        doc_type      = $null
        form          = $null
        is_excel      = $false
        fields        = New-OrderedMap
        table         = $null
        warnings      = New-Object 'System.Collections.Generic.List[string]'
        review_reason = $null
    }
}

function Find-DocType($Rules, [string]$Text) {
    # 규칙에 적힌 순서대로 판별키워드를 찾는다(대소문자 무시). 없으면 $null.
    $lowered = $Text.ToLowerInvariant()
    foreach ($doc in @($Rules.doc_types)) {
        foreach ($keyword in @($doc.keywords)) {
            if ($lowered.Contains(([string]$keyword).ToLowerInvariant())) { return $doc }
        }
    }
    return $null
}

function Get-FieldValue($Field, [string]$Text, $Warnings) {
    $m = $Field.regex.Match($Text)
    if (-not $m.Success -or $m.Groups[1].Value.Trim() -eq '') {
        $Warnings.Add("필드 '$($Field.name)' 못 찾음")
        return $null
    }
    $raw = $m.Groups[1].Value.Trim()
    if ($Field.format -eq '날짜') {
        $value = ConvertTo-DateValue $raw
    } elseif ($Field.format -eq '금액') {
        $value = ConvertTo-Amount $raw
    } else {
        return $raw
    }
    if ($null -eq $value) { $Warnings.Add("필드 '$($Field.name)' 값 '$raw' 를 $($Field.format)로 바꾸지 못함") }
    return $value
}

function ConvertFrom-TableRow([string]$Line, $Table) {
    # 행패턴으로 한 줄을 나누고 금액열 · 수량열을 숫자로 바꾼다. 안 맞으면 $null.
    # 파이썬 re.match 와 같게 줄 처음에서 시작하는 일치만 인정한다(끝은 고정하지 않음).
    $m = $Table.row_regex.Match($Line)
    if (-not $m.Success -or $m.Index -ne 0) { return $null }
    $row = New-OrderedMap
    $columns = @($Table.columns)
    for ($i = 0; $i -lt $columns.Count; $i++) {
        $row[[string]$columns[$i]] = $m.Groups[$i + 1].Value.Trim()
    }
    foreach ($name in @($Table.amount_columns)) {
        $amount = ConvertTo-Amount $row[$name]
        if ($null -eq $amount) { return $null }
        $row[$name] = $amount
    }
    foreach ($name in @($Table.quantity_columns)) {
        $amount = ConvertTo-Amount $row[$name]
        if ($null -eq $amount -or $amount -ne [decimal]::Truncate($amount)) { return $null }
        $row[$name] = [long]$amount
    }
    return $row
}

function Get-TableRows($Table, [string]$Text, $Warnings) {
    # 시작키워드 줄 다음부터 끝키워드 줄 앞까지. 빈 줄은 건너뛰고, 행패턴과 안 맞는 줄이 하나라도 있으면 표 전체를 버린다.
    $lines = Split-TextLines $Text
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Contains($Table.start)) { $start = $i; break }
    }
    if ($start -lt 0) {
        $Warnings.Add("표 시작키워드 '$($Table.start)' 없음, 품목 건너뜀")
        return $null
    }
    $end = -1
    for ($i = $start + 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Contains($Table.end)) { $end = $i; break }
    }
    if ($end -lt 0) {
        $Warnings.Add("표 끝키워드 '$($Table.end)' 없음, 품목 건너뜀")
        return $null
    }
    $rows = New-Object 'System.Collections.Generic.List[object]'
    for ($i = $start + 1; $i -lt $end; $i++) {
        $line = $lines[$i].Trim()
        if ($line -eq '') { continue }
        $row = ConvertFrom-TableRow $line $Table
        if ($null -eq $row) {
            $Warnings.Add("표 행패턴과 맞지 않는 줄이 있어 품목 건너뜀: $line")
            return $null
        }
        $rows.Add($row)
    }
    return ,$rows
}

function Get-TextExtract($Rules, [string]$Text) {
    $result = New-ExtractResult
    if ([string]::IsNullOrWhiteSpace($Text)) {
        $result.review_reason = $script:NoText
        return $result
    }
    $doc = Find-DocType $Rules $Text
    if ($null -eq $doc) {
        $result.review_reason = '판별 실패 (판별키워드가 하나도 없음)'
        return $result
    }
    $result.doc_type = $doc.name
    foreach ($field in @($doc.fields)) {
        $result.fields[[string]$field.name] = Get-FieldValue $field $Text $result.warnings
    }
    if ($null -eq $result.fields[[string]$doc.number_field]) {
        $result.review_reason = "번호필드 미추출 ('$($doc.number_field)' 패턴이 맞지 않음)"
    }
    if ($null -ne $doc.table) {
        $result.table = Get-TableRows $doc.table $Text $result.warnings
    }
    return $result
}

function ConvertTo-ResultJson($Result) {
    # 결과 객체를 JSON 글자로. 금액 · 날짜는 글자로, 한 줄짜리 표 · 경고도 배열로 남게 직접 옮긴다.
    $fields = New-OrderedMap
    foreach ($key in $Result.fields.Keys) { $fields[$key] = ConvertTo-JsonValue $Result.fields[$key] }
    $table = $null
    if ($null -ne $Result.table) {
        $table = New-Object 'System.Collections.Generic.List[object]'
        foreach ($row in $Result.table) {
            $plain = New-OrderedMap
            foreach ($key in $row.Keys) { $plain[$key] = ConvertTo-JsonValue $row[$key] }
            $table.Add($plain)
        }
        $table = [object[]]$table.ToArray()
    }
    $out = [ordered]@{}
    foreach ($key in $Result.Keys) { $out[$key] = $Result[$key] }
    $out.fields = $fields
    $out.table = $table
    $out.warnings = [object[]]@($Result.warnings)
    return ConvertTo-Json -InputObject $out -Depth 20
}
