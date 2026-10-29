# 마크업표(파트넘버 → 마크업% · 반올림자리)를 읽고 본사 단가 · 금액을 계산 · 검산한다. 동작 기준은 app\markup.py 전부(오류 문구까지 같게).
# 점-소싱 순서: Values → Rules → ExtractText → ExtractGrid → Verify → Markup
# 마크업표는 마크업표.xlsx 첫 시트의 격자(ExtractGrid.ps1 의 격자 모양, 칸 서식 f 포함)에서 읽는다.
#
# 마크업표 객체(Read-MarkupGrid 의 결과): Ordinal OrderedDictionary, 열쇠 = 파트 열쇠(앞뒤 공백 뺀 소문자),
#   값 = @{ part = <표에 적힌 파트넘버>; rate = [decimal] (% 숫자, 3 = 3%); places = [int] 또는 $null (반올림 없음) }
# 계산은 정확한 십진수(BigInteger 계수 + 지수 = 파이썬 Decimal 과 같은 자리 규칙)로 하고, 값을 돌려줄 때만 [decimal] 로 바꾼다.
#   [decimal] 에 정확히 담을 수 없으면(28자리 넘음) 조용히 반올림하지 않고 멈춘다.

$script:MarkupFile = '마크업표.xlsx'
$script:PartColumn = '파트넘버'
$script:RateColumn = '마크업%'
$script:PlacesColumn = '반올림자리'
$script:MaxRateText = '10'
$script:MaxPlaces = 10
$script:RateSource = '마크업.율'
$script:PriceSource = '마크업.단가'
$script:AmountSource = '마크업.금액'
$script:BlankPart = '(빈칸)'
$script:MarkupPrefix = '마크업.'
$script:LiteralFormat = [regex]'"[^"]*"|\\.'
$script:DecimalTextPattern = [regex]'^([+-]?)([0-9]*)(?:\.([0-9]*))?(?:[eE]([+-]?[0-9]+))?\z'
$script:Pow32 = [System.Numerics.BigInteger]::Pow($script:BigTwo, 32)
$script:Pow64 = [System.Numerics.BigInteger]::Pow($script:BigTwo, 64)
$script:DecimalMax = [System.Numerics.BigInteger]::Parse('79228162514264337593543950335')
$script:WorkDigits = 60       # 지금 판 _exact 의 계산 자리(넘으면 Inexact)
$script:RoundDigits = 28      # 지금 판 _round(quantize) 는 기본 문맥 28자리

# ---- 정확한 십진수: @{ C = 계수(부호 있는 BigInteger); E = 지수 }  값 = C x 10^E -------------------------

function New-PyDec([System.Numerics.BigInteger]$Coef, [int]$Exp) {
    return @{ C = $Coef; E = $Exp }
}

function ConvertTo-PyDec([decimal]$Value) {
    # [decimal] 을 그대로(자리 수 유지: 100.00 → C=10000, E=-2).
    $bits = [decimal]::GetBits($Value)
    $low = [System.Numerics.BigInteger]([long]$bits[0] -band 4294967295)
    $mid = [System.Numerics.BigInteger]([long]$bits[1] -band 4294967295)
    $high = [System.Numerics.BigInteger]([long]$bits[2] -band 4294967295)
    $coef = $low + $mid * $script:Pow32 + $high * $script:Pow64
    if ($bits[3] -lt 0) { $coef = [System.Numerics.BigInteger]::Negate($coef) }
    return @{ C = $coef; E = - (($bits[3] -shr 16) -band 255) }
}

function ConvertFrom-PyDec($D) {
    # 정확히 같은 [decimal] 로. 자리 수 · 크기가 [decimal] 에 안 들어가면 throw(반올림하지 않음).
    $coef = $D.C
    $exp = [int]$D.E
    if ($exp -gt 0) {
        $coef = $coef * [System.Numerics.BigInteger]::Pow($script:BigTen, $exp)
        $exp = 0
    }
    while ($exp -lt -28 -and -not $coef.IsZero -and [System.Numerics.BigInteger]::Remainder($coef, $script:BigTen).IsZero) {
        $coef = [System.Numerics.BigInteger]::Divide($coef, $script:BigTen)
        $exp++
    }
    if ($coef.IsZero -and $exp -lt -28) { $exp = -28 }
    $magnitude = [System.Numerics.BigInteger]::Abs($coef)
    if ($exp -lt -28 -or $magnitude -gt $script:DecimalMax) {
        throw '마크업 계산 결과가 무설치판이 담을 수 있는 숫자 범위(28자리)를 넘어 정확히 담을 수 없습니다'
    }
    $value = [decimal]::Parse($magnitude.ToString() + 'E' + $exp, [Globalization.NumberStyles]::Float, $script:Invariant)
    if ($coef.Sign -lt 0) { return - $value }
    return $value
}

function Format-PyDec($D) {
    # 지금 판 _show(Decimal) = format(x, 'f'): 지수대로 자리를 채운다(103.0000).
    $digits = [System.Numerics.BigInteger]::Abs($D.C).ToString()
    $exp = [int]$D.E
    if ($exp -ge 0) {
        $text = $digits + ('0' * $exp)
    } else {
        $scale = - $exp
        if ($digits.Length -le $scale) { $digits = ('0' * ($scale - $digits.Length + 1)) + $digits }
        $text = $digits.Substring(0, $digits.Length - $scale) + '.' + $digits.Substring($digits.Length - $scale)
    }
    if ($D.C.Sign -lt 0) { return '-' + $text }
    return $text
}

function Compare-PyDec($A, $B) {
    $exp = [Math]::Min([int]$A.E, [int]$B.E)
    $left = $A.C * [System.Numerics.BigInteger]::Pow($script:BigTen, [int]$A.E - $exp)
    $right = $B.C * [System.Numerics.BigInteger]::Pow($script:BigTen, [int]$B.E - $exp)
    return $left.CompareTo($right)
}

function Assert-PyDecDigits($D, [int]$Limit) {
    if ([System.Numerics.BigInteger]::Abs($D.C).ToString().Length -gt $Limit) {
        throw "마크업 계산이 ${Limit}자리를 넘어 정확히 계산할 수 없습니다"
    }
    return $D
}

function Invoke-PyMultiply($A, $B) {
    return Assert-PyDecDigits (New-PyDec ($A.C * $B.C) ([int]$A.E + [int]$B.E)) $script:WorkDigits
}

function Invoke-PyAdd($A, $B) {
    $exp = [Math]::Min([int]$A.E, [int]$B.E)
    $left = $A.C * [System.Numerics.BigInteger]::Pow($script:BigTen, [int]$A.E - $exp)
    $right = $B.C * [System.Numerics.BigInteger]::Pow($script:BigTen, [int]$B.E - $exp)
    return Assert-PyDecDigits (New-PyDec ($left + $right) $exp) $script:WorkDigits
}

function Invoke-PyDivide100($D) {
    # 100 으로 나눈 정확한 값. 파이썬 Decimal 나눗셈처럼 몫의 뒤 0 은 피제수 자리(E)까지만 뺀다(10300.00 / 100 → 103.00, 10 / 100 → 0.1).
    $coef = $D.C
    $exp = [int]$D.E - 2
    $steps = 0
    while ($steps -lt 2 -and ($coef.IsZero -or [System.Numerics.BigInteger]::Remainder($coef, $script:BigTen).IsZero)) {
        $coef = [System.Numerics.BigInteger]::Divide($coef, $script:BigTen)
        $exp++
        $steps++
    }
    return New-PyDec $coef $exp
}

function Invoke-PyRound($D, $Places) {
    # 지금 판 _round: 반올림자리가 있을 때만 그 자리에서 사사오입(ROUND_HALF_UP, 0 에서 먼 쪽), 없으면 그대로.
    if ($null -eq $Places) { return $D }
    $target = - [int]$Places
    $coef = $D.C
    if ([int]$D.E -gt $target) {
        $coef = $coef * [System.Numerics.BigInteger]::Pow($script:BigTen, [int]$D.E - $target)
    } elseif ([int]$D.E -lt $target) {
        $divisor = [System.Numerics.BigInteger]::Pow($script:BigTen, $target - [int]$D.E)
        $magnitude = [System.Numerics.BigInteger]::Abs($coef)
        $quotient = [System.Numerics.BigInteger]::Divide($magnitude, $divisor)
        $rest = $magnitude - $quotient * $divisor
        if (($rest * $script:BigTwo).CompareTo($divisor) -ge 0) { $quotient = $quotient + [System.Numerics.BigInteger]::One }
        $coef = if ($coef.Sign -lt 0) { [System.Numerics.BigInteger]::Negate($quotient) } else { $quotient }
    }
    return Assert-PyDecDigits (New-PyDec $coef $target) $script:RoundDigits
}

function ConvertTo-PlainPyDec($D) {
    # 지금 판 _plain: 뒤의 0 을 빼고(3.00 → 3) 지수가 양수가 되면 정수로(300 이 3E+2 로 바뀌지 않게).
    if ($D.C.IsZero) { return New-PyDec $D.C 0 }
    $coef = $D.C
    $exp = [int]$D.E
    while ([System.Numerics.BigInteger]::Remainder($coef, $script:BigTen).IsZero) {
        $coef = [System.Numerics.BigInteger]::Divide($coef, $script:BigTen)
        $exp++
    }
    if ($exp -gt 0) {
        $coef = $coef * [System.Numerics.BigInteger]::Pow($script:BigTen, $exp)
        $exp = 0
    }
    return New-PyDec $coef $exp
}

function ConvertFrom-DecimalText([string]$Text) {
    # 파이썬 Decimal(글자) 가 받는 모양(부호 · 소수점 · 지수)만. 숫자가 아니거나 무한대 · NaN 이면 $null.
    $m = $script:DecimalTextPattern.Match($Text)
    if (-not $m.Success) { return $null }
    $integer = $m.Groups[2].Value
    $fraction = $m.Groups[3].Value
    if ($integer.Length + $fraction.Length -eq 0) { return $null }
    $exp = 0
    if ($m.Groups[4].Success) {
        if ($m.Groups[4].Value.Length -gt 9) { return $null }
        $exp = [int]$m.Groups[4].Value
    }
    $coef = [System.Numerics.BigInteger]::Parse($integer + $fraction)
    if ($m.Groups[1].Value -eq '-') { $coef = [System.Numerics.BigInteger]::Negate($coef) }
    return New-PyDec $coef ($exp - $fraction.Length)
}

# ---- 값 → 글자 · 파트 열쇠 ----------------------------------------------------------------------------

function ConvertTo-PyText($Value) {
    # 파이썬 str(값): 오류 문구에 들어가는 칸 값. 글자는 그대로, 참거짓 True/False, 날짜 'YYYY-MM-DD HH:MM:SS', 실수는 왕복 글자.
    if ($null -eq $Value) { return 'None' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { if ($Value) { return 'True' } else { return 'False' } }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant) }
    if ($Value -is [double]) { return ConvertTo-FloatText $Value }
    if ($Value -is [decimal]) { return $Value.ToString($script:Invariant) }
    return [string]::Format($script:Invariant, '{0}', $Value)
}

function ConvertTo-PartText($Value) {
    # 파트 칸 값의 글자: 지금 판 part_key · _label 처럼 정수 모양의 실수(12345.0)는 12345.
    if ($Value -is [decimal]) {
        if ([decimal]::Truncate($Value) -eq $Value) { return $Value.ToString('F0', $script:Invariant) }
        return (Remove-TrailingZeros $Value).ToString($script:Invariant)
    }
    if ($Value -is [double] -and -not [double]::IsNaN($Value) -and -not [double]::IsInfinity($Value) -and [Math]::Floor($Value) -eq $Value) {
        return ([System.Numerics.BigInteger]$Value).ToString()
    }
    return ConvertTo-PyText $Value
}

function Get-PartKey($Value) {
    # 지금 판 part_key: 앞뒤 공백을 빼고 대소문자를 무시한다(빈 칸 · 참거짓은 ''). 파이썬 casefold 의 ß → ss 만 따로 맞춘다.
    if ($null -eq $Value -or $Value -is [bool]) { return '' }
    return (ConvertTo-PartText $Value).Trim().ToLowerInvariant().Replace('ß', 'ss')
}

function Get-PartLabel($Value) {
    # 지금 판 _label: 사용자에게 보이는 파트넘버(빈 품번은 「(빈칸)」).
    if ((Get-PartKey $Value) -eq '') { return $script:BlankPart }
    return (ConvertTo-PartText $Value).Trim()
}

function Test-PercentFormat($NumberFormat) {
    # 지금 판 is_percent_format: 따옴표 안 · \ 다음 글자의 % 는 글자일 뿐 백분율 서식이 아니다.
    if ($null -eq $NumberFormat) { return $false }
    return $script:LiteralFormat.Replace([string]$NumberFormat, '').Contains('%')
}

# ---- 객체 · 칸 읽기 ------------------------------------------------------------------------------------

function Get-MarkupProp($Object, [string]$Name) {
    # 해시표(PS 안) 또는 PSCustomObject(JSON)의 이름 하나. 없으면 $null.
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function ConvertTo-MarkupList($Value) {
    # 배열 · List · 낱개 값을 List 로(값이 없으면 빈 List). 해시표가 든 List 에 @() 를 쓰면 5.1 이 "Argument types do not match" 로 멈춘다.
    $list = New-Object 'System.Collections.Generic.List[object]'
    if ($null -ne $Value) { foreach ($element in $Value) { $list.Add($element) } }
    return ,$list
}
function Get-MarkupColumn($Columns, [string]$Name, [string]$Default) {
    $value = Get-MarkupProp $Columns $Name
    if ($null -eq $value) { return $Default }
    return [string]$value
}

function Test-MarkupEmpty($Cell) {
    # 지금 판 _empty: 칸이 없거나 값이 없거나 공백뿐인 글자.
    if ($null -eq $Cell -or $null -eq $Cell.v) { return $true }
    return ($Cell.v -is [string] -and $Cell.v.Trim() -eq '')
}

function New-MarkupCells($Sheet) {
    # 첫 시트의 칸들을 (행 x 20000 + 열) 열쇠 사전으로. 칸 = @{ t; v; f(표시 형식); row; col }.
    $cells = New-Object 'System.Collections.Generic.Dictionary[long,object]'
    $maxRow = [long]0
    $maxCol = [long]0
    $list = Get-MarkupProp $Sheet 'cells'
    if ($null -ne $list) {
        foreach ($item in $list) {
            $cell = ConvertTo-GridCell $item[2]
            $format = Get-MarkupProp $item[2] 'f'
            $cell['f'] = if ($null -eq $format) { '' } else { [string]$format }
            $cell['row'] = [long]$item[0]
            $cell['col'] = [long]$item[1]
            $cells[$cell['row'] * $script:ColumnStride + $cell['col']] = $cell
            if ($cell['row'] -gt $maxRow) { $maxRow = $cell['row'] }
            if ($cell['col'] -gt $maxCol) { $maxCol = $cell['col'] }
        }
    }
    return @{ Cells = $cells; MaxRow = $maxRow; MaxCol = $maxCol }
}

function Get-MarkupCell($Model, [long]$Row, [long]$Col) {
    if ($Col -lt 1) { return $null }
    $key = $Row * $script:ColumnStride + $Col
    if ($Model.Cells.ContainsKey($key)) { return $Model.Cells[$key] }
    return $null
}

# ---- 마크업표 읽기 -------------------------------------------------------------------------------------

function ConvertTo-MarkupRate($Cell) {
    # 지금 판 _rate. 반환 @{ Error = 사유 또는 $null; Rate = [decimal] }.
    #   3 = 3%, 백분율 서식 칸 0.005 = 0.5%, 글자 "0.5%" = 0.5%. 서식 없는 1 미만(0.03)은 3% 를 잘못 적은 것일 수 있어 거절.
    $rateName = $script:RateColumn
    if (Test-MarkupEmpty $Cell) { return @{ Error = "$rateName 가 비어 있습니다(0% 면 0 을 적으세요)" } }
    $raw = ConvertTo-PyText $Cell.v
    if ($Cell.t -eq 'b') { return @{ Error = "$rateName '$raw' 가 숫자가 아닙니다" } }
    $explicit = $false
    $number = $null
    if ($Cell.t -eq 'n') {
        $excel = ConvertTo-ExcelNumber $Cell.v
        if ($null -ne $excel) {
            $number = ConvertTo-PyDec $excel
            $explicit = Test-PercentFormat $Cell.f
            if ($explicit) { $number = Invoke-PyMultiply $number (New-PyDec 100 0) }   # 백분율 서식 칸: 엑셀은 3% 를 0.03 으로 보관한다
        }
    } else {
        $text = $raw.Trim()
        $explicit = $text.EndsWith('%', [StringComparison]::Ordinal)
        if ($explicit) { $text = $text.Substring(0, $text.Length - 1).Trim() }
        $number = ConvertFrom-DecimalText $text.Replace(',', '')
    }
    if ($null -eq $number) { return @{ Error = "$rateName '$raw' 가 숫자가 아닙니다" } }
    $number = ConvertTo-PlainPyDec $number
    $zero = New-PyDec 0 0
    $one = New-PyDec 1 0
    if ((Compare-PyDec $number $zero) -lt 0 -or (Compare-PyDec $number (New-PyDec 10 0)) -gt 0) {
        return @{ Error = "$rateName $(Format-PyDec $number) 는 0 이상 $($script:MaxRateText) 이하여야 합니다" }
    }
    if ((Compare-PyDec $number $zero) -gt 0 -and (Compare-PyDec $number $one) -lt 0 -and -not $explicit) {
        # 서식 없는 0.03 은 0.03% 다. 3% 를 0.03 으로 잘못 적었으면 마크업이 1/100 이 되므로 받지 않는다
        $shown = Format-PyDec $number
        return @{ Error = "$shown 은 $shown% 로 읽힙니다. 3% 면 3 을, 1% 미만이면 0.5% 처럼 % 를 붙여 적으세요" }
    }
    try {
        return @{ Error = $null; Rate = (ConvertFrom-PyDec $number) }
    } catch {
        return @{ Error = "$rateName '$raw' 는 자리 수가 너무 많아 무설치판이 정확히 담을 수 없습니다" }
    }
}

function ConvertTo-MarkupPlaces($Cell) {
    # 지금 판 _places. 반환 @{ Error; Places = [int] 또는 $null(비어 있으면 반올림하지 않음) }.
    if (Test-MarkupEmpty $Cell) { return @{ Error = $null; Places = $null } }
    $number = $null
    if ($Cell.t -eq 'n') {
        $excel = ConvertTo-ExcelNumber $Cell.v
        if ($null -ne $excel) { $number = ConvertTo-PyDec $excel }
    } elseif ($Cell.t -eq 's') {
        $number = ConvertFrom-DecimalText ([string]$Cell.v).Trim()
    }
    if ($null -ne $number) {
        $number = ConvertTo-PlainPyDec $number
        # 정수(지수 0)이고 0 이상 MAX_PLACES 이하여야 한다
        if ([int]$number.E -eq 0 -and $number.C.Sign -ge 0 -and $number.C -le [System.Numerics.BigInteger]$script:MaxPlaces) {
            return @{ Error = $null; Places = [int]$number.C }
        }
    }
    return @{ Error = "$($script:PlacesColumn) '$(ConvertTo-PyText $Cell.v)' 는 0 부터 $($script:MaxPlaces) 까지의 정수여야 합니다(비우면 반올림하지 않음)"; Places = $null }
}

function Read-MarkupGrid($Grid) {
    # 마크업표 격자(첫 시트)를 표로 읽는다. 잘못된 줄이 하나라도 있으면 지금 판 MarkupError 와 같은 문구로 throw.
    $sheet = $null
    $sheets = Get-MarkupProp $Grid 'sheets'
    if ($null -ne $sheets) { $sheet = @($sheets)[0] }
    $model = New-MarkupCells $sheet

    $partAt = 0
    $rateAt = 0
    $placesAt = 0
    for ($col = 1; $col -le $model.MaxCol; $col++) {
        $cell = Get-MarkupCell $model 1 $col
        if ($null -eq $cell -or $cell.t -ne 's') { continue }
        $name = ([string]$cell.v).Trim()
        if ($partAt -eq 0 -and [string]::Equals($name, $script:PartColumn, [StringComparison]::Ordinal)) { $partAt = $col }
        if ($rateAt -eq 0 -and [string]::Equals($name, $script:RateColumn, [StringComparison]::Ordinal)) { $rateAt = $col }
        if ($placesAt -eq 0 -and [string]::Equals($name, $script:PlacesColumn, [StringComparison]::Ordinal)) { $placesAt = $col }
    }
    if ($partAt -eq 0 -or $rateAt -eq 0) {
        throw "$($script:MarkupFile) 첫 줄(머리글)에 '$($script:PartColumn)' 와 '$($script:RateColumn)' 열이 있어야 합니다."
    }

    $table = New-OrderedMap
    $firstLine = @{}
    $errors = New-Object 'System.Collections.Generic.List[string]'
    for ($line = 2; $line -le $model.MaxRow; $line++) {
        $part = Get-MarkupCell $model $line $partAt
        $rate = Get-MarkupCell $model $line $rateAt
        $places = Get-MarkupCell $model $line $placesAt
        if ((Test-MarkupEmpty $part) -and (Test-MarkupEmpty $rate) -and (Test-MarkupEmpty $places)) { continue }   # 빈 줄
        $where = "$($script:MarkupFile) ${line}번째 줄"
        $partValue = if ($null -ne $part) { $part.v } else { $null }
        $key = Get-PartKey $partValue
        if ($key -eq '') {
            $errors.Add("${where}: 파트넘버가 비어 있습니다")
            continue
        }
        $label = Get-PartLabel $partValue
        $rateResult = ConvertTo-MarkupRate $rate
        if ($null -ne $rateResult.Error) {
            $errors.Add("$where ($label): $($rateResult.Error)")
            continue
        }
        $placesResult = ConvertTo-MarkupPlaces $places
        if ($null -ne $placesResult.Error) {
            $errors.Add("$where ($label): $($placesResult.Error)")
            continue
        }
        if ($table.Contains($key)) {
            $errors.Add("${where}: 파트넘버 '$label' 가 $($firstLine[$key])번째 줄('$($table[$key].part)')에도 " +
                '있습니다(대소문자 · 앞뒤 공백은 같은 파트로 봄. 한 파트는 한 줄만)')
            continue
        }
        $table[$key] = @{ part = $label; rate = $rateResult.Rate; places = $placesResult.Places }
        $firstLine[$key] = $line
    }
    if ($errors.Count -gt 0) {
        $shown = [string]::Join('; ', @($errors | Select-Object -First 5))
        if ($errors.Count -gt 5) { $shown += " 외 $($errors.Count - 5)건" }
        throw "$($script:MarkupFile) 를 고쳐야 합니다: $shown"
    }
    return ,$table
}

# ---- 계산 ----------------------------------------------------------------------------------------------

function Get-MarkupUnit([decimal]$Price, $Row) {
    # 지금 판 apply_markup: 본사 단가 = 단가 x (1 + 율/100). 반올림자리가 있으면 그 자리에서 사사오입, 없으면 정확한 값 그대로. 정확한 십진수로 돌려준다.
    $factor = Invoke-PyAdd (New-PyDec 1 0) (Invoke-PyDivide100 (ConvertTo-PyDec $Row.rate))
    return Invoke-PyRound (Invoke-PyMultiply (ConvertTo-PyDec $Price) $factor) $Row.places
}

function Get-MarkupNumber($Item, [string]$Name) {
    # 지금 판 _number: 품목 칸 값을 [decimal] 로(실수는 엑셀처럼 유효 15자리, 글자는 쉼표 · $ 제거). 숫자가 아니면 $null.
    # JSON 에서 읽은 품목의 소수점 숫자는 엑셀 실수였던 값이라 유효 15자리로 맞추고, 메모리 품목의 [decimal] 은 금액이라 자리 그대로 둔다.
    $value = Get-MarkupProp $Item $Name
    if ($null -eq $value -or $value -is [bool]) { return $null }
    if ($value -is [decimal] -and $Item -isnot [System.Collections.IDictionary]) { return ConvertTo-ExcelNumber $value }
    return ConvertTo-CheckDecimal $value
}

function Find-MarkupRow($Item, $Columns, $Table) {
    # 지금 판 lookup_row: 품목의 파트넘버로 마크업표 줄 하나, 없으면 $null.
    if ($null -eq $Table) { return $null }
    $key = Get-PartKey (Get-MarkupProp $Item (Get-MarkupColumn $Columns 'part' '품번'))
    if ($Table.Contains($key)) { return $Table[$key] }
    return $null
}

function Get-MarkupValue([string]$Source, $Item, $Columns, $Table) {
    # 회사 양식 열 값 마크업.율 · 마크업.단가 · 마크업.금액([decimal]). 계산할 수 없으면 $null(검산이 확인필요로 보낸다).
    $row = Find-MarkupRow $Item $Columns $Table
    if ($null -eq $row) { return $null }
    if ([string]::Equals($Source, $script:RateSource, [StringComparison]::Ordinal)) { return $row.rate }
    $price = Get-MarkupNumber $Item (Get-MarkupColumn $Columns 'price' '단가')
    if ($null -eq $price) { return $null }
    $unit = Get-MarkupUnit $price $row
    if ([string]::Equals($Source, $script:PriceSource, [StringComparison]::Ordinal)) { return ConvertFrom-PyDec $unit }
    $quantity = Get-MarkupNumber $Item (Get-MarkupColumn $Columns 'quantity' '수량')
    if ($null -eq $quantity) { return $null }
    return ConvertFrom-PyDec (Invoke-PyMultiply (ConvertTo-PyDec $quantity) $unit)   # 금액은 반올림하지 않는다
}

# ---- 없는 파트 · 빈 품번 -------------------------------------------------------------------------------

function Test-FormUsesMarkup($Form) {
    # 지금 판 uses_markup: 회사 양식 열 값에 마크업.* 가 하나라도 있는가.
    if ($null -eq $Form) { return $false }
    foreach ($column in (ConvertTo-MarkupList (Get-MarkupProp $Form 'columns'))) {
        $value = Get-MarkupProp $column 'value'
        if ($null -ne $value -and ([string]$value).StartsWith($script:MarkupPrefix, [StringComparison]::Ordinal)) { return $true }
    }
    return $false
}

function Get-MarkupTableRows($Result) {
    # 결과의 품목 표를 List 로(표가 없으면 빈 List). 품목은 해시표(PS 안) 또는 PSCustomObject(JSON).
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $table = Get-ResultMember $Result 'table'
    if ($null -ne $table) { foreach ($item in $table) { $rows.Add($item) } }
    return ,$rows
}

function Get-MissingParts($Result, $Form, $Table) {
    # 마크업표에 없는 파트넘버(나온 순서, 한 번씩). 빈 품번은 넣지 않는다(Get-BlankParts).
    $missing = New-Object 'System.Collections.Generic.List[string]'
    if (-not (Test-FormUsesMarkup $Form)) { return ,$missing }
    $columns = Get-MarkupProp $Form 'markup_columns'
    $partName = Get-MarkupColumn $columns 'part' '품번'
    foreach ($item in (Get-MarkupTableRows $Result)) {
        $part = Get-MarkupProp $item $partName
        if ((Get-PartKey $part) -ne '' -and $null -eq (Find-MarkupRow $item $columns $Table)) {
            $label = Get-PartLabel $part
            if (-not $missing.Contains($label)) { $missing.Add($label) }
        }
    }
    return ,$missing
}

function Get-BlankParts($Result, $Form) {
    # 마크업을 쓰는데 품번이 비어 있는 품목 줄 수.
    if (-not (Test-FormUsesMarkup $Form)) { return 0 }
    $partName = Get-MarkupColumn (Get-MarkupProp $Form 'markup_columns') 'part' '품번'
    $count = 0
    foreach ($item in (Get-MarkupTableRows $Result)) {
        if ((Get-PartKey (Get-MarkupProp $item $partName)) -eq '') { $count++ }
    }
    return $count
}

function Get-MissingReason($Result, $Form, $Table) {
    # 마크업을 붙일 수 없는 품목이 있으면 확인필요 사유(빠진 파트 · 빈 품번), 없으면 $null. 마크업표에 없는 파트를 0% 로 넘기지 않는다.
    $missing = Get-MissingParts $Result $Form $Table
    $blanks = Get-BlankParts $Result $Form
    $lines = New-Object 'System.Collections.Generic.List[string]'
    if ($missing.Count -gt 0) {
        $lines.Add("마크업표에 없는 파트넘버: $([string]::Join(', ', $missing.ToArray())). $($script:MarkupFile) 에 이 파트 줄(파트넘버 · 마크업%)을 " +
            '더하면 다음 실행 때 다시 시도합니다(0% 면 0 을 적으세요)')
        foreach ($part in $missing) {
            if ($part.StartsWith('0', [StringComparison]::Ordinal)) {
                $lines.Add('앞자리가 0 인 품번은 마크업표 칸을 텍스트 서식으로 적어야 맞습니다')
                break
            }
        }
    }
    if ($blanks -gt 0) { $lines.Add("품번이 비어 있는 줄이 ${blanks}개 있습니다. 규칙.yaml 의 품번 열을 확인하세요") }
    if ($lines.Count -eq 0) { return $null }
    return [string]::Join("`n", $lines.ToArray())
}

function Get-WrittenMarkup($Form, $Rows) {
    # 회사 양식 행(값 목록)에서 마크업 열에 쓴 값만 줄마다 {값 출처: 값} 으로(지금 판 written_markup).
    $sources = @(foreach ($column in (ConvertTo-MarkupList (Get-MarkupProp $Form 'columns'))) { [string](Get-MarkupProp $column 'value') })
    $out = New-Object 'System.Collections.Generic.List[object]'
    if ($null -ne $Rows) {
        foreach ($row in $Rows) {
            $map = New-OrderedMap
            $cells = ConvertTo-MarkupList $row
            $count = [Math]::Min($sources.Count, $cells.Count)
            for ($i = 0; $i -lt $count; $i++) {
                $source = $sources[$i]
                if ($source -ceq $script:RateSource -or $source -ceq $script:PriceSource -or $source -ceq $script:AmountSource) { $map[$source] = $cells[$i] }
            }
            $out.Add($map)
        }
    }
    return ,$out
}

# ---- 검산 ----------------------------------------------------------------------------------------------

function Get-WrittenEntry($Entry, [string]$Key) {
    # 쓴 값 한 줄({값 출처: 값})에서 값 하나: @{ Has = 열이 있는가; Value = [decimal] 또는 $null }.
    $has = $false
    $value = $null
    if ($Entry -is [System.Collections.IDictionary]) {
        if ($Entry.Contains($Key)) { $has = $true; $value = $Entry[$Key] }
    } elseif ($null -ne $Entry) {
        $property = $Entry.PSObject.Properties[$Key]
        if ($null -ne $property) { $has = $true; $value = $property.Value }
    }
    if ($null -ne $value -and $value -isnot [decimal]) { $value = ConvertTo-CheckDecimal $value }
    return @{ Has = $has; Value = $value }
}

function Get-MarkupChecks($Result, $RowsWritten, $Table, $Columns) {
    # 줄마다 마크업% · 본사 단가(· 쓴 경우 본사 금액)를 따로 계산해 쓴 값과 대조한 검산 목록(List).
    #   본사 단가 = 고객 단가 x (100 + 마크업%) / 100 (반올림자리가 있으면 그 자리에서 사사오입), 본사 금액 = 수량 x 본사 단가(반올림 없음), 마크업% 는 0 이상 10 이하.
    #   $RowsWritten: 줄마다 {값 출처: 값} (Get-WrittenMarkup 의 결과). 목록이 짧으면 모자란 줄은 쓴 값이 없는 것으로 본다.
    $checks = New-Object 'System.Collections.Generic.List[object]'
    $partName = Get-MarkupColumn $Columns 'part' '품번'
    $priceName = Get-MarkupColumn $Columns 'price' '단가'
    $quantityName = Get-MarkupColumn $Columns 'quantity' '수량'
    $written = ConvertTo-MarkupList $RowsWritten
    $items = Get-MarkupTableRows $Result
    for ($n = 1; $n -le $items.Count; $n++) {
        $item = $items[$n - 1]
        $entry = if ($n - 1 -lt $written.Count) { $written[$n - 1] } else { $null }
        $label = Get-PartLabel (Get-MarkupProp $item $partName)
        $row = Find-MarkupRow $item $Columns $Table
        if ($null -eq $row) {
            $checks.Add((New-Check "${n}줄 마크업% ($label)" '마크업표에 없음' '' $false))
            continue
        }
        $writtenRate = Get-WrittenEntry $entry $script:RateSource
        $rate = if ($writtenRate.Has) { $writtenRate.Value } else { $row.rate }
        $rateOk = ($row.rate -ge 0 -and $row.rate -le 10 -and $null -ne $rate -and $rate -eq $row.rate)
        $checks.Add((New-Check "${n}줄 마크업% ($label)" "마크업표 $(Show-CheckValue $row.rate)%" (Show-CheckValue $rate) $rateOk))

        $writtenPrice = Get-WrittenEntry $entry $script:PriceSource
        $price = Get-MarkupNumber $item $priceName
        if ($null -eq $price) {
            $checks.Add((New-Check "${n}줄 본사 단가" "고객 단가($priceName) 없음" (Show-CheckValue $writtenPrice.Value) $false))
            continue
        }
        $exact = Invoke-PyDivide100 (Invoke-PyMultiply (ConvertTo-PyDec $price) (Invoke-PyAdd (New-PyDec 100 0) (ConvertTo-PyDec $row.rate)))
        $expected = Invoke-PyRound $exact $row.places
        $text = "$(Show-CheckValue $price) x (100 + $(Show-CheckValue $row.rate)) / 100 = $(Format-PyDec $exact)"
        if ($null -ne $row.places) { $text += ", 소수 $($row.places)자리 반올림 $(Format-PyDec $expected)" }
        if ($writtenPrice.Has) {
            $unitText = Show-CheckValue $writtenPrice.Value
            $unitOk = ($null -ne $writtenPrice.Value -and (Compare-PyDec (ConvertTo-PyDec $writtenPrice.Value) $expected) -eq 0)
        } else {
            $unit = Get-MarkupUnit $price $row
            $unitText = Format-PyDec $unit
            $unitOk = ((Compare-PyDec $unit $expected) -eq 0)
        }
        $checks.Add((New-Check "${n}줄 본사 단가" $text $unitText $unitOk))

        $writtenAmount = Get-WrittenEntry $entry $script:AmountSource
        if ($writtenAmount.Has) {
            $quantity = Get-MarkupNumber $item $quantityName
            if ($null -eq $quantity) {
                $checks.Add((New-Check "${n}줄 본사 금액" "수량($quantityName) 없음" (Show-CheckValue $writtenAmount.Value) $false))
                continue
            }
            $amount = Invoke-PyMultiply (ConvertTo-PyDec $quantity) $expected
            $amountOk = ($null -ne $writtenAmount.Value -and (Compare-PyDec (ConvertTo-PyDec $writtenAmount.Value) $amount) -eq 0)
            $checks.Add((New-Check "${n}줄 본사 금액" "$(Show-CheckValue $quantity) x $(Format-PyDec $expected) = $(Format-PyDec $amount)" (Show-CheckValue $writtenAmount.Value) $amountOk))
        }
    }
    return ,$checks
}

# ---- 지문 ----------------------------------------------------------------------------------------------

function Get-MarkupFingerprint($Table, $Parts) {
    # 이 문서가 쓰는 파트들의 마크업표 줄(율 · 반올림자리, 없으면 「-」)에서 만든 16자리 지문. 다른 파트를 더하거나 고쳐도 그대로이고,
    # 이 문서 파트의 마크업이 바뀔 때만 달라진다(지금 판 markup_stamp 와 같은 값).
    $used = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
    if ($null -ne $Parts) {
        foreach ($part in $Parts) {
            $key = Get-PartKey $part
            if ($null -ne $Table -and $Table.Contains($key)) {
                $row = $Table[$key]
                $places = if ($null -eq $row.places) { 'None' } else { [string]$row.places }
                $used[$key] = "$(Show-CheckValue $row.rate)|$places"
            } else {
                $used[$key] = '-'
            }
        }
    }
    $keys = [string[]]@($used.Keys)
    [Array]::Sort($keys, [StringComparer]::Ordinal)
    $lines = @(foreach ($key in $keys) { "$key`t$($used[$key])" })
    $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes([string]::Join("`n", $lines))
    $hash = [Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return ([BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant()).Substring(0, 16)
}

# ---- 엔진 op ---------------------------------------------------------------------------------------------

function ConvertTo-MarkupJson($Object) {
    return ConvertTo-Json -InputObject $Object -Depth 20
}

function Invoke-MarkupOp($Payload) {
    # 엔진 op `markup`. 입력의 action 으로 함수를 고른다(시험용 입구):
    #   read         {grid}                                         -> [{key, part, rate(글자), places}]
    #   value        {grid, items, sources, columns?}               -> {values: [[값 글자 또는 null, ...] 품목마다]}
    #   missing      {grid, result: {table}, form}                  -> {reason, missing, blanks}
    #   checks       {grid, result: {table}, rows_written, columns?} -> 검산 목록
    #   fingerprint  {grid, parts}                                  -> {fingerprint}
    #   written      {form, rows}                                   -> [{값 출처: 값}]
    # grid = 마크업표 첫 시트의 격자, form = 규칙 company_form(columns · markup_columns), columns = {part, price, quantity}.
    $action = [string](Get-MarkupProp $Payload 'action')
    if ($action -ceq 'written') {
        $list = Get-WrittenMarkup (Get-MarkupProp $Payload 'form') (Get-MarkupProp $Payload 'rows')
        return ConvertTo-MarkupJson ([object[]]$list.ToArray())
    }
    $table = Read-MarkupGrid (Get-MarkupProp $Payload 'grid')
    $columns = Get-MarkupProp $Payload 'columns'
    if ($action -ceq 'read') {
        $rows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($key in $table.Keys) {
            $row = $table[$key]
            $rows.Add([ordered]@{ key = $key; part = $row.part; rate = (ConvertTo-JsonValue $row.rate); places = $row.places })
        }
        return ConvertTo-MarkupJson ([object[]]$rows.ToArray())
    }
    if ($action -ceq 'value') {
        $sources = @(Get-MarkupProp $Payload 'sources')
        $rows = New-Object 'System.Collections.Generic.List[object]'
        $items = Get-MarkupProp $Payload 'items'
        if ($null -ne $items) {
            foreach ($item in $items) {
                $values = New-Object 'System.Collections.Generic.List[object]'
                foreach ($source in $sources) {
                    $value = Get-MarkupValue ([string]$source) $item $columns $table
                    if ($null -eq $value) { $values.Add($null) } else { $values.Add((ConvertTo-JsonValue $value)) }
                }
                $rows.Add([object[]]$values.ToArray())
            }
        }
        return ConvertTo-MarkupJson ([ordered]@{ values = [object[]]$rows.ToArray() })
    }
    if ($action -ceq 'missing') {
        $result = Get-MarkupProp $Payload 'result'
        $form = Get-MarkupProp $Payload 'form'
        $missing = Get-MissingParts $result $form $table
        return ConvertTo-MarkupJson ([ordered]@{
            reason = (Get-MissingReason $result $form $table)
            missing = [object[]]$missing.ToArray()
            blanks = (Get-BlankParts $result $form)
        })
    }
    if ($action -ceq 'checks') {
        $checks = Get-MarkupChecks (Get-MarkupProp $Payload 'result') (Get-MarkupProp $Payload 'rows_written') $table $columns
        return ConvertTo-ChecksJson $checks
    }
    if ($action -ceq 'fingerprint') {
        return ConvertTo-MarkupJson ([ordered]@{ fingerprint = (Get-MarkupFingerprint $table (Get-MarkupProp $Payload 'parts')) })
    }
    throw "마크업 action '$action' 는 없습니다."
}
