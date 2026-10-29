# 엑셀 격자 → 엑셀양식 판별 → 머리 · 품목 추출. 동작 기준은 app\excel_extract.py 의 read_excel(글자까지 같게).
# 격자: {"active": 활성 시트 번호(0부터, 없으면 0), "sheets": [{"name", "cells": [[행, 열, {"t", "v", "f"}], …]}]}
#   t = n(숫자: 정수 · 소수) · s(글자) · b(참거짓) · d(날짜, v 는 "YYYY-MM-DDTHH:MM:SS"). 빈 칸은 없다. f(표시 형식)는 읽지 않는다.
# 결과 객체는 ExtractText.ps1 의 New-ExtractResult 와 같다(form = 엑셀양식 이름, is_excel = $true).
# ConvertTo-ExcelNumber(엑셀 숫자)는 Values.ps1 에 있다.

$script:DetectRows = 20       # 판별키워드를 찾는 줄 수 (칸이 있는 줄만 센다)
$script:ScanRows = 30         # 판별 줄을 모으는 행 수 (지금 판 max(DETECT_ROWS, HEAD_LINES))
$script:RightOf = '오른쪽:'
$script:ColumnStride = [long]20000   # 칸 열쇠 = 행 x 20000 + 열 (열은 20000 보다 작다)
$script:NoForm = '판별 실패 (규칙.yaml 의 엑셀양식 중 맞는 고객이 없음. 05장대로 이 고객의 엑셀 양식을 가르치세요)'
$script:CellAddress = [regex]'^([A-Z]{1,3})([1-9][0-9]*)$'

# ---- 칸 값 -------------------------------------------------------------------------------------

function ConvertTo-NumberText($Value) {
    # 파이썬 str(숫자) 에 가깝게: 정수는 그대로, JSON 의 소수는 적힌 그대로(5.0 포함), double 은 왕복 글자(1E-07 → 1e-07).
    if ($Value -is [double]) { return $Value.ToString('R', $script:Invariant).Replace('E', 'e') }
    return $Value.ToString($script:Invariant)
}

function ConvertTo-CellText($Cell) {
    # 지금 판 _text: 빈 칸은 '', 날짜는 YYYY-MM-DD(시각이 있으면 'YYYY-MM-DD HH:MM:SS'), 나머지는 글자로 바꿔 앞뒤 공백을 지운다.
    if ($null -eq $Cell) { return '' }
    switch ($Cell.t) {
        'd' {
            if ($Cell.v.TimeOfDay.Ticks -eq 0) { return $Cell.v.ToString('yyyy-MM-dd', $script:Invariant) }
            return $Cell.v.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant)
        }
        'b' { if ($Cell.v) { return 'True' } else { return 'False' } }
        'n' { return ConvertTo-NumberText $Cell.v }
        default { return ([string]$Cell.v).Trim() }
    }
}

function ConvertTo-GridCell($Raw) {
    # 격자 JSON 의 {t, v} 를 안에서 쓰는 칸으로. 날짜 글자는 [datetime] 으로(PS 7 은 이미 바꿔 주기도 한다).
    $value = $Raw.v
    if ([string]$Raw.t -eq 'd' -and $value -isnot [datetime]) {
        $value = [datetime]::ParseExact([string]$value, 'yyyy-MM-ddTHH:mm:ss', $script:Invariant)
    }
    return @{ t = [string]$Raw.t; v = $value; row = [long]0; col = [long]0 }
}

function New-GridModel($Sheet) {
    # 칸을 행 순서 다음 열 순서로 늘어놓은 사전(SortedDictionary). 읽기 순서가 파이썬 iter_rows 와 같다.
    $cells = New-Object 'System.Collections.Generic.SortedDictionary[long,object]'
    if ($null -ne $Sheet.cells) {
        foreach ($item in $Sheet.cells) {
            $cell = ConvertTo-GridCell $item[2]
            $cell['row'] = [long]$item[0]
            $cell['col'] = [long]$item[1]
            $cells[$cell['row'] * $script:ColumnStride + $cell['col']] = $cell
        }
    }
    return @{ Cells = $cells }
}

function Get-GridCell($Model, [long]$Row, [long]$Col) {
    $key = $Row * $script:ColumnStride + $Col
    if ($Model.Cells.ContainsKey($key)) { return $Model.Cells[$key] }
    return $null
}

function Get-GridLines($Model, [int]$Rows) {
    # 1~Rows 행에서 칸이 있는 줄마다 칸 글자를 탭으로 이은 글자(지금 판 _sheet_lines).
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $parts = New-Object 'System.Collections.Generic.List[string]'
    $current = 0
    foreach ($cell in $Model.Cells.Values) {
        if ($cell['row'] -gt $Rows) { break }
        if ($cell['row'] -ne $current) {
            if ($parts.Count -gt 0) { $lines.Add([string]::Join("`t", $parts)); $parts.Clear() }
            $current = $cell['row']
        }
        $parts.Add((ConvertTo-CellText $cell))
    }
    if ($parts.Count -gt 0) { $lines.Add([string]::Join("`t", $parts)) }
    return ,$lines
}

function Find-GridCell($Model, [string]$Label) {
    # 글자가 정확히 같은 첫 칸(위에서 아래, 왼쪽에서 오른쪽). 없으면 $null.
    foreach ($cell in $Model.Cells.Values) {
        if ([string]::Equals((ConvertTo-CellText $cell), $Label, [StringComparison]::Ordinal)) { return $cell }
    }
    return $null
}

function ConvertTo-ColumnIndex([string]$Letters) {
    $index = 0
    foreach ($ch in $Letters.ToCharArray()) { $index = $index * 26 + ([int]$ch - 64) }
    return $index
}

function Get-GridHeadValue($Model, [string]$Where) {
    # 지금 판 _head_value: 머리 칸의 (칸, 칸을 찾았는지). 'B4' 면 그 칸, '오른쪽:글자' 면 그 글자가 쓰인 칸의 바로 오른쪽 칸.
    $address = $script:CellAddress.Match($Where)
    if ($address.Success) {
        return @{ Cell = Get-GridCell $Model ([long]$address.Groups[2].Value) (ConvertTo-ColumnIndex $address.Groups[1].Value); Found = $true }
    }
    $found = Find-GridCell $Model $Where.Substring($script:RightOf.Length).Trim()
    if ($null -eq $found) { return @{ Cell = $null; Found = $false } }
    return @{ Cell = Get-GridCell $Model $found['row'] ($found['col'] + 1); Found = $true }
}

function ConvertTo-GridDate($Cell) {
    # 지금 판 _as_date: 날짜 칸은 날짜만, 그 밖은 글자로 보고 %m/%d/%Y · %Y-%m-%d 로 읽는다.
    if ($Cell.t -eq 'd') { return $Cell.v.Date }
    return ConvertTo-DateValue (ConvertTo-CellText $Cell)
}

function ConvertTo-GridAmount($Cell) {
    # 지금 판 _as_amount: 빈 칸 · 참거짓 · 날짜는 $null, 숫자는 엑셀 숫자, 글자는 쉼표 · $ 를 뺀 금액.
    if ($null -eq $Cell) { return $null }
    if ($Cell.t -eq 'n') { return ConvertTo-ExcelNumber $Cell.v }
    if ($Cell.t -eq 's') { return ConvertTo-Amount ([string]$Cell.v) }
    return $null
}

function ConvertTo-PlainCellValue($Cell) {
    # 지금 판 _plain: 형식 지정이 없는 칸의 값. 글자는 앞뒤 공백을 지우고, 숫자는 숫자(JSON 숫자)로 남긴다.
    switch ($Cell.t) {
        'd' { return $Cell.v }
        'b' { return [bool]$Cell.v }
        'n' {
            if ($Cell.v -is [int] -or $Cell.v -is [long]) { return [long]$Cell.v }
            return [double]$Cell.v
        }
        default { return ([string]$Cell.v).Trim() }
    }
}

function ConvertTo-GridField([string]$Name, $Cell, $Format, $Warnings) {
    # 지금 판 _convert. 비었으면 $null(경고 없음), 형식(날짜 · 금액)을 못 맞추면 경고 하나.
    if ($null -eq $Cell -or (ConvertTo-CellText $Cell) -eq '') { return $null }
    if ($Format -eq '날짜') {
        $converted = ConvertTo-GridDate $Cell
    } elseif ($Format -eq '금액') {
        $converted = ConvertTo-GridAmount $Cell
    } else {
        return ConvertTo-PlainCellValue $Cell
    }
    if ($null -eq $converted) { $Warnings.Add("'$Name' 값 '$(ConvertTo-CellText $Cell)' 를 ${Format}로 바꾸지 못함") }
    return $converted
}

function Get-GridItems($Spec, $Model, $Warnings) {
    # 머리글 글자가 있는 첫 칸의 줄 바로 아래부터, 품번 칸(없으면 첫 열)이 빈 줄 앞까지. 숫자가 아닌 금액 · 수량이 있으면 표 전체를 버린다.
    $header = Find-GridCell $Model ([string]$Spec.header)
    if ($null -eq $header) {
        $Warnings.Add("품목 머리글 '$($Spec.header)' 를 못 찾음, 품목 건너뜀")
        return $null
    }
    $columns = @($Spec.columns.PSObject.Properties | ForEach-Object { @{ Name = $_.Name; Index = (ConvertTo-ColumnIndex ([string]$_.Value)) } })
    $stop = $columns | Where-Object { $_.Name -ceq '품번' } | Select-Object -First 1
    if ($null -eq $stop) { $stop = $columns[0] }
    $amountNames = @($Spec.amount_columns)
    $quantityNames = @($Spec.quantity_columns)

    $rows = New-Object 'System.Collections.Generic.List[object]'
    for ($r = $header['row'] + 1; ; $r++) {
        if ((ConvertTo-CellText (Get-GridCell $Model $r $stop.Index)) -eq '') { break }
        $row = New-OrderedMap
        foreach ($column in $columns) {
            $cell = Get-GridCell $Model $r $column.Index
            if ($amountNames -ccontains $column.Name) {
                $value = ConvertTo-GridAmount $cell
            } elseif ($quantityNames -ccontains $column.Name) {
                $amount = ConvertTo-GridAmount $cell
                $value = $null
                if ($null -ne $amount -and $amount -eq [decimal]::Truncate($amount)) { $value = [long]$amount }
            } elseif ($null -eq $cell) {
                $value = ''
            } else {
                $value = ConvertTo-PlainCellValue $cell
            }
            if ($null -eq $value) {
                $Warnings.Add("품목 ${r}행 '$($column.Name)' 값이 숫자가 아님, 품목 건너뜀")
                return $null
            }
            $row[[string]$column.Name] = $value
        }
        $rows.Add($row)
    }
    return ,$rows
}

function Get-GridExtract($Rules, $Grid) {
    # 활성 시트 하나를 읽는다: 칸이 있는 첫 20줄(1~30행)에서 판별키워드를 찾고, 맞는 엑셀양식의 머리 · 품목을 뽑는다.
    $result = New-ExtractResult
    $result.is_excel = $true
    $sheets = @($Grid.sheets)
    if ($sheets.Count -eq 0) { throw '격자에 시트가 없습니다.' }
    $index = 0
    if ($null -ne $Grid.PSObject.Properties['active']) { $index = [int]$Grid.active }
    $model = New-GridModel $sheets[$index]

    $lines = Get-GridLines $model $script:ScanRows
    if ($lines.Count -eq 0) {
        $result.review_reason = $script:NoText
        return $result
    }
    $detect = ([string]::Join("`n", $lines.GetRange(0, [Math]::Min($lines.Count, $script:DetectRows)))).ToLowerInvariant()
    $form = $null
    foreach ($candidate in @($Rules.excel_forms)) {
        foreach ($keyword in @($candidate.keywords)) {
            if ($detect.Contains(([string]$keyword).ToLowerInvariant())) { $form = $candidate; break }
        }
        if ($null -ne $form) { break }
    }
    if ($null -eq $form) {
        $result.review_reason = $script:NoForm
        return $result
    }

    $doc = @($Rules.doc_types) | Where-Object { $_.name -ceq $form.doc_type } | Select-Object -First 1
    if ($null -eq $doc) { throw "엑셀양식 '$($form.name)' 의 문서종류 '$($form.doc_type)' 가 규칙에 없습니다." }
    $result.doc_type = $doc.name
    $result.form = $form.name
    foreach ($head in @($form.head)) {
        $found = Get-GridHeadValue $model ([string]$head.where)
        if (-not $found.Found) { $result.warnings.Add("머리 '$($head.name)' 칸($($head.where))을 못 찾음") }
        $result.fields[[string]$head.name] = ConvertTo-GridField ([string]$head.name) $found.Cell $head.format $result.warnings
    }
    if ($null -eq $result.fields[[string]$doc.number_field]) {
        $result.review_reason = "번호필드 미추출 ('$($doc.number_field)' 패턴이 맞지 않음)"
    }
    if ($null -ne $form.items) {
        $result.table = Get-GridItems $form.items $model $result.warnings
    }
    return $result
}
