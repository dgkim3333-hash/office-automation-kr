# 회사 양식 줄과 파일 이름. 동작 기준은 app\outputs.py (_value · company_rows · company_form_name) 와 app\organize.py (render_text · free_name).
# 점-소싱 순서: Values → Rules → ExtractText → ExtractGrid → Verify → Markup → Company
#
# 회사 양식 줄: 품목 한 줄 = 한 행. 열마다 값 출처(머리.X · 품목.X · 줄번호 · 수량x단가 · 고정:글자 · 빈칸 · 마크업.단가/금액/율)로 칸 값을 만든다.
#   칸 값의 종류는 결과 값 그대로: [decimal] 금액 · [long] 정수(줄번호 10, 20 …) · [double] 실수 · [bool] · [datetime] 날짜(0시) 또는 일시 · [string] · $null(빈칸).
#   고객마스터.X 는 무설치판에 마스터가 없으므로 지금 판에 마스터 줄을 주지 않은 것과 같이 빈칸이다(규칙 내보내기가 이미 거절한다).
# 파일 이름: {필드} 자리를 채운다(Format-Template). 필드 값에서만 못 쓰는 글자를 _ 로 바꾼다(템플릿의 글자는 그대로).
#   파이썬 서식 문법 중 지원하는 것: {{ }} · {이름} · {이름:서식}. 서식은 날짜(strftime) · 글자(채움 · 정렬 · 너비 · 정밀도) · 숫자([부호][0][너비][,|_][.정밀도][d|f]).
#   변환(!r) · 속성(.x) · 첨자([0]) · 안쪽 {} · 그 밖의 서식은 다르게 채우지 않고 멈춘다.

$script:LineStep = 10                       # 줄번호: 10, 20, 30 …
$script:ForbiddenInName = New-Object Text.RegularExpressions.Regex('[\\/:*?"<>|]')
$script:MissingText = '미상'
$script:SpecPattern = New-Object Text.RegularExpressions.Regex('^(?:(?<fill>.)?(?<align>[<>=^]))?(?<sign>[+\- ])?(?<alt>#)?(?<zero>0)?(?<width>[0-9]+)?(?<group>[,_])?(?:\.(?<prec>[0-9]+))?(?<type>[A-Za-z%])?$',
    [Text.RegularExpressions.RegexOptions]::Singleline)
# 파이썬 str.strip() 이 지우는 빈칸 글자(.NET Trim 과 조금 달라 직접 적는다: 0x1C~0x1F 를 포함하고 U+200B 는 포함하지 않는다)
$script:PyWhitespace = [char[]](@(9, 10, 11, 12, 13, 28, 29, 30, 31, 32, 0x85, 0xA0, 0x1680, 0x2028, 0x2029, 0x202F, 0x205F, 0x3000) + (0x2000..0x200A))

# ---- 값 → 글자 (파이썬 format(값, '')) ------------------------------------------------------------------

function ConvertTo-DecimalStr([decimal]$Value) {
    # 파이썬 str(Decimal): 지수가 0 이하이고 맨 앞 자리 지수가 -6 이상이면 그대로, 아니면 과학 표기(1E-8).
    $d = ConvertTo-PyDec $Value
    $digits = [System.Numerics.BigInteger]::Abs($d.C).ToString()
    $adjusted = [int]$d.E + $digits.Length - 1
    if ([int]$d.E -le 0 -and $adjusted -ge -6) { return Format-PyDec $d }
    $text = $digits.Substring(0, 1)
    if ($digits.Length -gt 1) { $text += '.' + $digits.Substring(1) }
    $text += 'E' + $(if ($adjusted -lt 0) { '-' } else { '+' }) + ([Math]::Abs($adjusted)).ToString($script:Invariant)
    if ($d.C.Sign -lt 0) { return '-' + $text }
    return $text
}

function ConvertTo-FormatText($Value) {
    # 서식 없이 채울 때의 글자(파이썬 str(값)). 0시 일시는 날짜로 본다(지금 판 엑셀 읽기도 0시는 날짜로 읽는다).
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { if ($Value) { return 'True' } else { return 'False' } }
    if ($Value -is [decimal]) { return ConvertTo-DecimalStr $Value }
    if ($Value -is [double]) { return ConvertTo-FloatText $Value }
    if ($Value -is [datetime]) {
        if ($Value.TimeOfDay.Ticks -eq 0) { return $Value.ToString('yyyy-MM-dd', $script:Invariant) }
        return $Value.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant)
    }
    if ($Value -is [int] -or $Value -is [long]) { return ([long]$Value).ToString($script:Invariant) }
    throw "파일 이름에 넣을 수 없는 값 종류입니다: $($Value.GetType().Name)"
}

# ---- 서식 ------------------------------------------------------------------------------------------------

function Format-StrfTime([datetime]$Value, [string]$Format) {
    # 파이썬 date.__format__ = strftime. 영어 C 로케일 기준으로 흔한 지시자만 지원하고 나머지는 멈춘다.
    $out = New-Object Text.StringBuilder
    $i = 0
    while ($i -lt $Format.Length) {
        $c = [string]$Format[$i]
        $i++
        if ($c -cne '%') { [void]$out.Append($c); continue }
        if ($i -ge $Format.Length) { throw "날짜 서식 '$Format' 의 끝에 % 만 있습니다" }
        $d = [string]$Format[$i]
        $i++
        $yday = $Value.DayOfYear - 1
        $wday = [int]$Value.DayOfWeek
        switch -CaseSensitive ($d) {
            'Y' { [void]$out.Append($Value.ToString('yyyy', $script:Invariant)) }
            'y' { [void]$out.Append($Value.ToString('yy', $script:Invariant)) }
            'm' { [void]$out.Append($Value.ToString('MM', $script:Invariant)) }
            'd' { [void]$out.Append($Value.ToString('dd', $script:Invariant)) }
            'H' { [void]$out.Append($Value.ToString('HH', $script:Invariant)) }
            'M' { [void]$out.Append($Value.ToString('mm', $script:Invariant)) }
            'S' { [void]$out.Append($Value.ToString('ss', $script:Invariant)) }
            'f' { [void]$out.Append($Value.ToString('ffffff', $script:Invariant)) }
            'I' { [void]$out.Append($Value.ToString('hh', $script:Invariant)) }
            'p' { [void]$out.Append($Value.ToString('tt', $script:Invariant)) }
            'j' { [void]$out.Append($Value.DayOfYear.ToString('000', $script:Invariant)) }
            'B' { [void]$out.Append($Value.ToString('MMMM', $script:Invariant)) }
            'b' { [void]$out.Append($Value.ToString('MMM', $script:Invariant)) }
            'h' { [void]$out.Append($Value.ToString('MMM', $script:Invariant)) }
            'A' { [void]$out.Append($Value.ToString('dddd', $script:Invariant)) }
            'a' { [void]$out.Append($Value.ToString('ddd', $script:Invariant)) }
            'w' { [void]$out.Append($wday.ToString($script:Invariant)) }
            'u' { [void]$out.Append($(if ($wday -eq 0) { 7 } else { $wday })) }
            'U' { [void]$out.Append(([int][Math]::Floor(($yday + 7 - $wday) / 7)).ToString('00', $script:Invariant)) }
            'W' { [void]$out.Append(([int][Math]::Floor(($yday + 7 - (($wday + 6) % 7)) / 7)).ToString('00', $script:Invariant)) }
            'F' { [void]$out.Append($Value.ToString('yyyy-MM-dd', $script:Invariant)) }
            '%' { [void]$out.Append('%') }
            default { throw "날짜 서식 '%$d' 는 무설치판이 지원하지 않습니다" }
        }
    }
    return $out.ToString()
}

function Get-SpecParts([string]$Spec) {
    $m = $script:SpecPattern.Match($Spec)
    if (-not $m.Success) { throw "서식 '$Spec' 를 읽을 수 없습니다" }
    $parts = @{}
    foreach ($name in 'fill', 'align', 'sign', 'alt', 'zero', 'width', 'group', 'prec', 'type') {
        $parts[$name] = if ($m.Groups[$name].Success) { $m.Groups[$name].Value } else { $null }
    }
    return $parts
}

function Set-FieldWidth([string]$Prefix, [string]$Body, $Width, [string]$Fill, [string]$Align) {
    # 채움 글자로 너비를 맞춘다. '=' 는 부호(Prefix)와 숫자 사이를 채운다.
    $pad = 0
    if ($null -ne $Width) { $pad = [int]$Width - ($Prefix.Length + $Body.Length) }
    if ($pad -le 0) { return $Prefix + $Body }
    if ($Align -ceq '<') { return $Prefix + $Body + ($Fill * $pad) }
    if ($Align -ceq '^') {
        $left = [int][Math]::Floor($pad / 2)
        return ($Fill * $left) + $Prefix + $Body + ($Fill * ($pad - $left))
    }
    if ($Align -ceq '=') { return $Prefix + ($Fill * $pad) + $Body }
    return ($Fill * $pad) + $Prefix + $Body
}

function Format-StringSpec([string]$Value, [string]$Spec) {
    # 파이썬 글자 서식: [[채움]정렬][0][너비][.정밀도][s]. 부호 · # · 쉼표 · = 정렬 · 다른 형식은 파이썬도 오류다.
    $p = Get-SpecParts $Spec
    if ($null -ne $p.sign -or $null -ne $p.alt -or $null -ne $p.group -or $p.align -ceq '=') { throw "서식 '$Spec' 는 글자 값에 쓸 수 없습니다" }
    if ($null -ne $p.type -and $p.type -cne 's') { throw "서식 '$Spec' 는 글자 값에 쓸 수 없습니다" }
    $fill = if ($null -ne $p.fill) { $p.fill } else { ' ' }
    $align = if ($null -ne $p.align) { $p.align } else { '<' }
    if ($null -ne $p.zero) {
        if ($null -ne $p.align) { throw "서식 '$Spec' 의 0 과 정렬을 함께 쓰는 것은 무설치판이 지원하지 않습니다" }
        $fill = '0'
    }
    $text = $Value
    if ($null -ne $p.prec -and [int]$p.prec -lt $text.Length) { $text = $text.Substring(0, [int]$p.prec) }
    return Set-FieldWidth '' $text $p.width $fill $align
}

function Add-DigitGrouping([string]$Body, [string]$Separator) {
    # 정수 부분을 오른쪽부터 3자리씩 끊는다.
    if ($Body -match '[eEnNiI]') { throw '지수 · 무한대 표기에는 자리 구분(,)을 쓸 수 없습니다' }
    $point = $Body.IndexOf('.')
    $integer = if ($point -lt 0) { $Body } else { $Body.Substring(0, $point) }
    $rest = if ($point -lt 0) { '' } else { $Body.Substring($point) }
    $out = New-Object Text.StringBuilder
    for ($i = 0; $i -lt $integer.Length; $i++) {
        if ($i -gt 0 -and (($integer.Length - $i) % 3) -eq 0) { [void]$out.Append($Separator) }
        [void]$out.Append($integer[$i])
    }
    return $out.ToString() + $rest
}

function ConvertTo-FixedDigits([System.Numerics.BigInteger]$Scaled, [int]$Places) {
    # 소수 자리만큼 키운 정수(Scaled)를 소수점 있는 글자로.
    $digits = $Scaled.ToString()
    if ($Places -le 0) { return $digits }
    if ($digits.Length -le $Places) { $digits = ('0' * ($Places - $digits.Length + 1)) + $digits }
    return $digits.Substring(0, $digits.Length - $Places) + '.' + $digits.Substring($digits.Length - $Places)
}

function Get-HalfEvenQuotient([System.Numerics.BigInteger]$Numerator, [System.Numerics.BigInteger]$Denominator) {
    # 0 이상 Numerator/Denominator 를 정수로 반올림(꼭 반이면 짝수쪽, 파이썬 Decimal · float 서식과 같다).
    $quotient = [System.Numerics.BigInteger]::Divide($Numerator, $Denominator)
    $rest = $Numerator - $quotient * $Denominator
    $compare = ($rest * $script:BigTwo).CompareTo($Denominator)
    if ($compare -gt 0 -or ($compare -eq 0 -and -not $quotient.IsEven)) { $quotient = $quotient + [System.Numerics.BigInteger]::One }
    return $quotient
}

function Format-FixedBody($Value, [int]$Places) {
    # 부호를 뺀 크기를 소수 Places 자리로 반올림한 글자(정수 · 실수는 정확한 값으로, 금액은 자리 그대로의 정확한 십진수로).
    if ($Value -is [decimal]) {
        $d = ConvertTo-PyDec ([Math]::Abs($Value))
        $target = - $Places
        $coef = $d.C
        if ([int]$d.E -gt $target) {
            $coef = $coef * [System.Numerics.BigInteger]::Pow($script:BigTen, [int]$d.E - $target)
        } elseif ([int]$d.E -lt $target) {
            $coef = Get-HalfEvenQuotient $coef ([System.Numerics.BigInteger]::Pow($script:BigTen, $target - [int]$d.E))
        }
        return ConvertTo-FixedDigits $coef $Places
    }
    if ($Value -is [double]) {
        if ([double]::IsNaN($Value)) { return 'nan' }
        if ([double]::IsInfinity($Value)) { return 'inf' }
        if ($Value -eq 0) { return ConvertTo-FixedDigits ([System.Numerics.BigInteger]::Zero) $Places }
        $fraction = Get-DoubleFraction ([Math]::Abs($Value))
        $scaled = $fraction.Num * [System.Numerics.BigInteger]::Pow($script:BigTen, $Places)
        return ConvertTo-FixedDigits (Get-HalfEvenQuotient $scaled $fraction.Den) $Places
    }
    $whole = [System.Numerics.BigInteger]::Abs([System.Numerics.BigInteger]([long]$Value))
    return ConvertTo-FixedDigits ($whole * [System.Numerics.BigInteger]::Pow($script:BigTen, $Places)) $Places
}

function Format-NumberSpec($Value, [string]$Spec) {
    # 파이썬 숫자 서식의 일부: [[채움]정렬][부호][0][너비][,|_][.정밀도][d|f]. 그 밖(e · g · % · n · 정밀도만 · 0 과 정렬/쉼표 함께)은 멈춘다.
    $p = Get-SpecParts $Spec
    $kind = if ($Value -is [decimal]) { 'decimal' } elseif ($Value -is [double]) { 'float' } else { 'int' }
    $type = if ($null -ne $p.type) { $p.type } else { '' }
    if ($null -ne $p.alt -or ($type -cne '' -and $type -cne 'd' -and $type -cne 'f' -and $type -cne 'F')) { throw "서식 '$Spec' 는 무설치판이 지원하지 않습니다" }
    if ($type -ceq 'd' -and $kind -ne 'int') { throw "서식 '$Spec' 의 d 는 정수에만 쓸 수 있습니다" }
    if ($type -ceq '' -and $null -ne $p.prec) { throw "서식 '$Spec' 는 무설치판이 지원하지 않습니다(정밀도는 f 와 함께)" }
    if ($type -ceq 'd' -and $null -ne $p.prec) { throw "서식 '$Spec' 의 정수에는 정밀도를 쓸 수 없습니다" }
    if ($null -ne $p.zero -and ($null -ne $p.align -or $null -ne $p.group)) { throw "서식 '$Spec' 의 0 과 정렬 · 쉼표를 함께 쓰는 것은 무설치판이 지원하지 않습니다" }

    if ($kind -eq 'float') { $negative = (ConvertTo-FloatText $Value).StartsWith('-', [StringComparison]::Ordinal) } else { $negative = $Value -lt 0 }
    if ($type -ceq 'f' -or $type -ceq 'F') {
        if ($null -ne $p.prec) { $places = [int]$p.prec } elseif ($kind -eq 'decimal') { $places = $null } else { $places = 6 }
        if ($null -eq $places) {
            $body = Format-PyDec (ConvertTo-PyDec ([Math]::Abs($Value)))
        } else {
            $body = Format-FixedBody $Value $places
        }
    } elseif ($kind -eq 'decimal') {
        $body = ConvertTo-DecimalStr ([Math]::Abs($Value))
    } elseif ($kind -eq 'float') {
        $body = ConvertTo-FloatText ([Math]::Abs($Value))
    } else {
        $body = ([System.Numerics.BigInteger]::Abs([System.Numerics.BigInteger]([long]$Value))).ToString()
    }
    if ($null -ne $p.group) { $body = Add-DigitGrouping $body $p.group }

    $prefix = ''
    if ($negative) { $prefix = '-' } elseif ($p.sign -ceq '+') { $prefix = '+' } elseif ($p.sign -ceq ' ') { $prefix = ' ' }
    $fill = if ($null -ne $p.fill) { $p.fill } else { ' ' }
    $align = if ($null -ne $p.align) { $p.align } else { '>' }
    if ($null -ne $p.zero) { $fill = '0'; $align = '=' }
    return Set-FieldWidth $prefix $body $p.width $fill $align
}

function Format-Value($Value, [string]$Spec) {
    # 값 하나를 서식대로 글자로(파이썬 format(값, 서식)).
    if ($Spec -eq '') { return ConvertTo-FormatText $Value }
    if ($Value -is [datetime]) { return Format-StrfTime $Value $Spec }
    if ($Value -is [string]) { return Format-StringSpec $Value $Spec }
    if ($Value -is [bool]) { throw "서식 '$Spec' 는 참거짓 값에 쓸 수 없습니다" }
    if ($Value -is [decimal] -or $Value -is [double] -or $Value -is [int] -or $Value -is [long]) { return Format-NumberSpec $Value $Spec }
    throw "파일 이름에 넣을 수 없는 값 종류입니다: $($Value.GetType().Name)"
}

# ---- 파일 이름: 템플릿 채우기 ---------------------------------------------------------------------------------

function Get-TemplateField($Fields, [string]$Name) {
    # 필드 값 하나(없는 이름이면 멈춤). 값이 $null 이면 값을 못 찾은 필드.
    if ($Fields -is [System.Collections.IDictionary]) {
        if ($Fields.Contains($Name)) { return $Fields[$Name] }
    } elseif ($null -ne $Fields) {
        $property = $Fields.PSObject.Properties[$Name]
        if ($null -ne $property) { return $property.Value }
    }
    throw "파일 이름 서식의 '{$Name}' 에 해당하는 필드가 없습니다"
}

function Format-TemplateField([string]$Inner, $Fields) {
    # { } 안의 이름[:서식] 하나를 채운다. 못 쓰는 글자는 _, 앞뒤 공백과 끝의 점을 지우고, 비면 「미상」.
    $cut = $Inner.IndexOfAny([char[]]@(':', '!'))
    $name = if ($cut -lt 0) { $Inner } else { $Inner.Substring(0, $cut) }
    $rest = if ($cut -lt 0) { '' } else { $Inner.Substring($cut) }
    if ($rest.StartsWith('!', [StringComparison]::Ordinal)) { throw "파일 이름 서식의 변환(!)은 무설치판이 지원하지 않습니다: {$Inner}" }
    $spec = if ($rest.StartsWith(':', [StringComparison]::Ordinal)) { $rest.Substring(1) } else { '' }
    if ($spec.Contains('{') -or $spec.Contains('}')) { throw "파일 이름 서식 안의 {}는 무설치판이 지원하지 않습니다: {$Inner}" }
    if ($name -eq '' -or $name -match '^[0-9]+$') { throw "파일 이름 서식의 {} 에는 필드 이름이 있어야 합니다: {$Inner}" }
    if ($name.Contains('.') -or $name.Contains('[')) { throw "파일 이름 서식의 속성 · 첨자는 무설치판이 지원하지 않습니다: {$Inner}" }
    $value = Get-TemplateField $Fields $name
    $text = if ($null -eq $value) { $script:MissingText } else { Format-Value $value $spec }
    $text = $script:ForbiddenInName.Replace($text, '_').Trim($script:PyWhitespace).TrimEnd('.')
    if ($text -eq '') { return $script:MissingText }
    return $text
}

function Format-Template([string]$Template, $Fields) {
    # 지금 판 render_text: {필드} · {필드:%Y-%m} 자리를 채운다. 템플릿의 글자는 그대로, 필드 값에서만 금지 글자 \ / : * ? " < > | 를 _ 로 바꾼다.
    $out = New-Object Text.StringBuilder
    $length = $Template.Length
    $i = 0
    while ($i -lt $length) {
        $c = $Template[$i]
        if ($c -ceq '{') {
            if ($i + 1 -lt $length -and $Template[$i + 1] -ceq '{') { [void]$out.Append('{'); $i += 2; continue }
            $depth = 1
            $j = $i + 1
            while ($j -lt $length -and $depth -gt 0) {
                if ($Template[$j] -ceq '{') { $depth++ } elseif ($Template[$j] -ceq '}') { $depth-- }
                $j++
            }
            if ($depth -gt 0) { throw "파일 이름 서식에 닫히지 않은 '{' 가 있습니다: $Template" }
            [void]$out.Append((Format-TemplateField $Template.Substring($i + 1, $j - $i - 2) $Fields))
            $i = $j
        } elseif ($c -ceq '}') {
            if ($i + 1 -lt $length -and $Template[$i + 1] -ceq '}') { [void]$out.Append('}'); $i += 2; continue }
            throw "파일 이름 서식에 짝 없는 '}' 가 있습니다: $Template"
        } else {
            [void]$out.Append($c)
            $i++
        }
    }
    return $out.ToString()
}

# ---- 파일 이름: 이름 충돌 -------------------------------------------------------------------------------------

function Get-FreeName([string]$Path, $Reserved) {
    # 지금 판 free_name: 같은 이름이 있거나(Reserved 포함) 쓰기로 한 이름이면 확장자 앞에 _2, _3 … 을 붙인다.
    #   확장자는 파이썬 pathlib 규칙(맨 앞 점 · 끝의 점은 확장자가 아님). Reserved 는 대소문자 · / \ 표기를 가리지 않고 비교한다.
    $taken = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if ($null -ne $Reserved) { foreach ($item in $Reserved) { [void]$taken.Add(([string]$item).Replace('/', '\')) } }
    $full = $Path.Replace('/', '\')
    if (-not ((Test-Path -LiteralPath $full) -or $taken.Contains($full))) { return $full }
    $cut = $full.LastIndexOf('\')
    $folder = if ($cut -lt 0) { '' } else { $full.Substring(0, $cut + 1) }
    $name = $full.Substring($cut + 1)
    $dot = $name.LastIndexOf('.')
    if ($dot -gt 0 -and $dot -lt $name.Length - 1) { $stem = $name.Substring(0, $dot); $suffix = $name.Substring($dot) } else { $stem = $name; $suffix = '' }
    $n = 2
    while ($true) {
        $candidate = "$folder${stem}_$n$suffix"
        if (-not ((Test-Path -LiteralPath $candidate) -or $taken.Contains($candidate))) { return $candidate }
        $n++
    }
}

# ---- 회사 양식 줄 ---------------------------------------------------------------------------------------------

function ConvertTo-CompanyNumber($Value) {
    # 지금 판 _number: 수량 · 단가 칸 값을 정확한 십진수로(실수는 엑셀처럼 유효 15자리, 글자는 파이썬 Decimal 이 읽는 모양). 숫자가 아니면 멈춘다.
    if ($Value -is [decimal]) { return ConvertTo-PyDec $Value }
    if ($Value -is [int] -or $Value -is [long]) { return ConvertTo-PyDec ([decimal]([long]$Value)) }
    if ($Value -is [double]) {
        $number = ConvertTo-ExcelNumber $Value
        if ($null -ne $number) { return ConvertTo-PyDec $number }
    } elseif ($Value -is [string]) {
        $parsed = ConvertFrom-DecimalText $Value.Trim()
        if ($null -ne $parsed) { return $parsed }
    }
    throw "수량x단가: 숫자가 아닌 값 '$(ConvertTo-PyText $Value)' 입니다"
}

function Get-QuantityTimesPrice($Item) {
    # 지금 판 수량x단가: 수량과 단가가 둘 다 있을 때 곱. 파이썬 Decimal 곱은 28자리로 반올림하므로, 그보다 길면 다르게 담지 않고 멈춘다.
    $quantity = Get-MarkupProp $Item '수량'
    $price = Get-MarkupProp $Item '단가'
    if ($null -eq $quantity -or $null -eq $price) { return $null }
    $product = Invoke-PyMultiply (ConvertTo-CompanyNumber $quantity) (ConvertTo-CompanyNumber $price)
    if ([System.Numerics.BigInteger]::Abs($product.C).ToString().Length -gt $script:RoundDigits) {
        throw "수량x단가 계산이 ${script:RoundDigits}자리를 넘어 파이썬과 같게 담을 수 없습니다"
    }
    return ConvertFrom-PyDec $product
}

function Get-CompanyValue([string]$Source, $Fields, $Item, [int]$Index, $Form, $MarkupTable) {
    # 지금 판 _value: 값 출처 하나로 칸 값 하나.
    if ($Source -ceq '줄번호') { return [long](($Index + 1) * $script:LineStep) }
    if ($Source -ceq '빈칸') { return $null }
    if ($Source -ceq '수량x단가') { return Get-QuantityTimesPrice $Item }
    if ($Source.StartsWith('고정:', [StringComparison]::Ordinal)) { return $Source.Substring(3) }
    $dot = $Source.IndexOf('.')
    $group = if ($dot -lt 0) { $Source } else { $Source.Substring(0, $dot) }
    $name = if ($dot -lt 0) { '' } else { $Source.Substring($dot + 1) }
    if ($group -ceq '머리') { return Get-MarkupProp $Fields $name }
    if ($group -ceq '품목') { return Get-MarkupProp $Item $name }
    if ($group -ceq '마크업') { return Get-MarkupValue $Source $Item (Get-MarkupProp $Form 'markup_columns') $MarkupTable }
    return $null   # 고객마스터.X 와 알 수 없는 출처: 마스터 줄이 없으면 빈칸
}

function Get-CompanyRows($Result, $Form, $MarkupTable) {
    # 지금 판 company_rows: 회사 양식의 머리글과 행(품목 한 줄 = 한 행).
    #   $Result    : fields · table 이 값 종류 그대로인 결과(IDictionary). JSON 에서 읽은 결과는 ConvertTo-CompanyResult 로 먼저 되살린다.
    #   $Form      : 규칙 company_form(columns[].name · value, markup_columns, filename)
    #   $MarkupTable: Read-MarkupGrid 의 결과 또는 $null(마크업표 없음 → 마크업 칸은 빈칸)
    # 돌려주는 것: [ordered]@{ header = List[string]; rows = List[object](각 행 = List[object] 칸 값) }
    $fields = Get-ResultMember $Result 'fields'
    $table = Get-ResultMember $Result 'table'
    $columns = ConvertTo-MarkupList (Get-MarkupProp $Form 'columns')
    $header = New-Object 'System.Collections.Generic.List[string]'
    foreach ($column in $columns) { $header.Add([string](Get-MarkupProp $column 'name')) }
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $index = 0
    if ($null -ne $table) {
        foreach ($item in $table) {
            $cells = New-Object 'System.Collections.Generic.List[object]'
            foreach ($column in $columns) {
                $cells.Add((Get-CompanyValue ([string](Get-MarkupProp $column 'value')) $fields $item $index $Form $MarkupTable))
            }
            $rows.Add($cells)
            $index++
        }
    }
    return [ordered]@{ header = $header; rows = $rows }
}

function Get-CompanyFormName($Result, $Form) {
    # 지금 판 company_form_name: 회사 양식 파일 이름(번호필드 등을 채운 글자).
    return Format-Template ([string](Get-MarkupProp $Form 'filename')) (Get-ResultMember $Result 'fields')
}

function ConvertTo-CompanyResult($Rules, $Result) {
    # 결과 JSON(ConvertFrom-Json)을 값 종류가 살아 있는 메모리 결과로: 규칙의 형식으로 금액 · 날짜를 되살린다(Verify.ps1 의 ConvertTo-CheckModel). 메모리 결과는 그대로.
    if ($Result -is [System.Collections.IDictionary]) { return $Result }
    $docName = Get-ResultMember $Result 'doc_type'
    $doc = @($Rules.doc_types) | Where-Object { $_.name -ceq [string]$docName } | Select-Object -First 1
    if ($null -eq $doc) { throw "문서종류 '$docName' 가 규칙에 없습니다." }
    $model = ConvertTo-CheckModel $Rules $Result $doc
    return [ordered]@{ fields = $model.Fields; table = $model.Table }
}

# ---- 엔진 op (시험용 입구) ----------------------------------------------------------------------------------

function ConvertTo-TypedCell($Value) {
    # 칸 값을 종류가 보이는 JSON 모양 {t, v} 으로(금액은 자리 그대로의 글자, 실수는 파이썬 str(float), 날짜는 YYYY-MM-DD).
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return [ordered]@{ t = 'bool'; v = $Value } }
    if ($Value -is [decimal]) { return [ordered]@{ t = 'decimal'; v = $Value.ToString($script:Invariant) } }
    if ($Value -is [int] -or $Value -is [long]) { return [ordered]@{ t = 'int'; v = [long]$Value } }
    if ($Value -is [double]) { return [ordered]@{ t = 'float'; v = (ConvertTo-FloatText $Value) } }
    if ($Value -is [datetime]) {
        if ($Value.TimeOfDay.Ticks -eq 0) { return [ordered]@{ t = 'date'; v = $Value.ToString('yyyy-MM-dd', $script:Invariant) } }
        return [ordered]@{ t = 'datetime'; v = $Value.ToString('yyyy-MM-ddTHH:mm:ss', $script:Invariant) }
    }
    if ($Value -is [string]) { return [ordered]@{ t = 'text'; v = $Value } }
    throw "칸 값 종류를 알 수 없습니다: $($Value.GetType().Name)"
}

function ConvertFrom-TypedCell($Cell) {
    if ($null -eq $Cell) { return $null }
    $value = $Cell.v
    switch ([string]$Cell.t) {
        'bool' { return [bool]$value }
        'decimal' { return [decimal]::Parse([string]$value, [Globalization.NumberStyles]::Float, $script:Invariant) }
        'int' { return [long]$value }
        'float' { return [double]::Parse([string]$value, [Globalization.NumberStyles]::Float, $script:Invariant) }
        'date' { return [datetime]::ParseExact([string]$value, 'yyyy-MM-dd', $script:Invariant) }
        'datetime' { return [datetime]::ParseExact([string]$value, 'yyyy-MM-ddTHH:mm:ss', $script:Invariant) }
        'text' { return [string]$value }
    }
    throw "칸 종류 '$($Cell.t)' 를 알 수 없습니다"
}

function Invoke-CompanyOp($Payload) {
    # 엔진 op `company`. 입력의 action 으로 고른다(시험용 입구):
    #   rows       {rules, result, markup_grid?}            -> {header, rows: [[칸 또는 null, ...] 품목마다], filename}
    #   template   {template, fields: {이름: 칸 또는 null}}   -> {text}
    #   free_name  {path, reserved: [경로, ...]}            -> {path}
    $action = [string](Get-MarkupProp $Payload 'action')
    if ($action -ceq 'template') {
        $fields = New-OrderedMap
        foreach ($property in (Get-MarkupProp $Payload 'fields').PSObject.Properties) { $fields[[string]$property.Name] = ConvertFrom-TypedCell $property.Value }
        return ConvertTo-Json -InputObject ([ordered]@{ text = (Format-Template ([string](Get-MarkupProp $Payload 'template')) $fields) }) -Depth 20
    }
    if ($action -ceq 'free_name') {
        $reserved = @(Get-MarkupProp $Payload 'reserved')
        return ConvertTo-Json -InputObject ([ordered]@{ path = (Get-FreeName ([string](Get-MarkupProp $Payload 'path')) $reserved) }) -Depth 20
    }
    if ($action -ceq 'rows') {
        $rules = ConvertTo-PortableRules (Get-MarkupProp $Payload 'rules')
        $form = Get-MarkupProp $rules 'company_form'
        if ($null -eq $form) { throw '규칙에 회사양식이 없습니다.' }
        $table = $null
        $grid = Get-MarkupProp $Payload 'markup_grid'
        if ($null -ne $grid) { $table = Read-MarkupGrid $grid }
        $result = ConvertTo-CompanyResult $rules (Get-MarkupProp $Payload 'result')
        $built = Get-CompanyRows $result $form $table
        $outRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($row in $built.rows) {
            $cells = New-Object 'System.Collections.Generic.List[object]'
            foreach ($value in $row) { $cells.Add((ConvertTo-TypedCell $value)) }
            $outRows.Add([object[]]$cells.ToArray())
        }
        return ConvertTo-Json -Depth 20 -InputObject ([ordered]@{
            header = [object[]]$built.header.ToArray()
            rows = [object[]]$outRows.ToArray()
            filename = (Get-CompanyFormName $result $form)
        })
    }
    throw "회사 양식 action '$action' 는 없습니다."
}
