# 값 바꾸기: 글자 → 금액([decimal]) · 날짜([datetime]), 엑셀 칸 숫자 → 유효 15자리 금액(ConvertTo-ExcelNumber), 그리고 결과 JSON 표기.
# 동작 기준은 app\extract.py 의 to_amount · to_date. 문화권(소수점 쉼표 등)에 흔들리지 않게 InvariantCulture 만 쓴다.

$script:Invariant = [Globalization.CultureInfo]::InvariantCulture
# 파이썬 Decimal 처럼 부호 · 소수점 · 지수(1e3)를 받는다(지수는 값만 같고 글자 표기는 1000 으로 다름)
$script:AmountStyle = [Globalization.NumberStyles]::AllowLeadingSign -bor [Globalization.NumberStyles]::AllowDecimalPoint -bor [Globalization.NumberStyles]::AllowExponent

# strptime 의 %m · %d · %Y 가 받는 글자(파이썬 _strptime 과 같음: 한 자리 · 두 자리, %d 는 " 5" 도 받음)
$script:Month = '(1[0-2]|0[1-9]|[1-9])'
$script:Day = '(3[01]|[12][0-9]|0[1-9]|[1-9]| [1-9])'
$script:DateForms = @(
    # %m/%d/%Y
    @{ Regex = [regex]('^' + $script:Month + '/' + $script:Day + '/([0-9]{4})$'); Year = 3; Month = 1; Day = 2 },
    # %Y-%m-%d
    @{ Regex = [regex]('^([0-9]{4})-' + $script:Month + '-' + $script:Day + '$'); Year = 1; Month = 2; Day = 3 }
)

function ConvertTo-Amount([string]$Value) {
    # 쉼표 · $ 를 빼고 앞뒤 공백을 지운 뒤 [decimal] 로. 못 바꾸면 $null. 소수 자리(1234.50)는 그대로 남는다.
    if ($null -eq $Value) { return $null }
    $clean = $Value.Replace(',', '').Replace('$', '').Trim()
    $amount = [decimal]0
    if ([decimal]::TryParse($clean, $script:AmountStyle, $script:Invariant, [ref]$amount)) { return $amount }
    return $null
}

function ConvertTo-DateValue([string]$Value) {
    # 형식은 %m/%d/%Y · %Y-%m-%d 두 가지만. 없는 날짜(2/30)는 $null.
    if ($null -eq $Value) { return $null }
    $clean = $Value.Trim()
    foreach ($form in $script:DateForms) {
        $m = $form.Regex.Match($clean)
        if (-not $m.Success) { continue }
        $year = [int]$m.Groups[$form.Year].Value
        $month = [int]$m.Groups[$form.Month].Value
        $day = [int]$m.Groups[$form.Day].Value.Trim()
        if ($year -ge 1 -and $day -le [DateTime]::DaysInMonth($year, $month)) {
            return New-Object DateTime($year, $month, $day)
        }
    }
    return $null
}

function ConvertTo-JsonValue($Value) {
    # 결과 JSON 표기: 금액은 자리 그대로의 글자("1234.50"), 날짜는 "YYYY-MM-DD", 나머지는 그대로.
    if ($Value -is [decimal]) { return $Value.ToString($script:Invariant) }
    if ($Value -is [DateTime]) {
        if ($Value.TimeOfDay.Ticks -ne 0) { return $Value.ToString('yyyy-MM-ddTHH:mm:ss', $script:Invariant) }   # 시각이 있는 칸(엑셀 PO 의 서식 없는 칸)
        return $Value.ToString('yyyy-MM-dd', $script:Invariant)
    }
    return $Value
}

Add-Type -AssemblyName System.Numerics

$script:ExcelDigits = 15      # 엑셀이 숫자를 보관 · 계산하는 유효 자리 수
$script:BigTwo = [System.Numerics.BigInteger]2
$script:BigTen = [System.Numerics.BigInteger]10

# ---- 숫자: 엑셀처럼 유효 15자리 ------------------------------------------------------------------

function Get-DoubleFraction([double]$Value) {
    # 0 보다 큰 double 을 정확한 분수 Num/Den 으로(Den 은 2 의 거듭제곱). 글자 변환의 어림 없이 비트에서 바로 뽑는다.
    $bits = [BitConverter]::DoubleToInt64Bits($Value)
    $exponent = [int](($bits -shr 52) -band 0x7FF)
    $fraction = $bits -band [long]4503599627370495
    if ($exponent -eq 0) {
        $mantissa = $fraction
        $power = -1074
    } else {
        $mantissa = $fraction + [long]4503599627370496
        $power = $exponent - 1075
    }
    $num = [System.Numerics.BigInteger]$mantissa
    if ($power -ge 0) {
        return @{ Num = $num * [System.Numerics.BigInteger]::Pow($script:BigTwo, $power); Den = [System.Numerics.BigInteger]::One }
    }
    return @{ Num = $num; Den = [System.Numerics.BigInteger]::Pow($script:BigTwo, - $power) }
}

function Get-NearestDouble([System.Numerics.BigInteger]$Num, [System.Numerics.BigInteger]$Den) {
    # 양수 Num/Den 에 가장 가까운 double(가운데면 짝수쪽)의 정확한 분수. 15자리 다음이 꼭 5 인 글자의 올림 · 내림은 double 의 실제 값으로 가려야 한다.
    $low = [System.Numerics.BigInteger]::Pow($script:BigTwo, 52)
    $high = [System.Numerics.BigInteger]::Pow($script:BigTwo, 53)
    $shift = 52 - [int][Math]::Floor([System.Numerics.BigInteger]::Log($Num, 2) - [System.Numerics.BigInteger]::Log($Den, 2))
    while ($true) {
        if ($shift -ge 0) {
            $n = $Num * [System.Numerics.BigInteger]::Pow($script:BigTwo, $shift)
            $d = $Den
        } else {
            $n = $Num
            $d = $Den * [System.Numerics.BigInteger]::Pow($script:BigTwo, - $shift)
        }
        $m = [System.Numerics.BigInteger]::Divide($n, $d)
        if ($m -lt $low) { $shift++ } elseif ($m -ge $high) { $shift-- } else { break }
    }
    $compare = (($n - $m * $d) * $script:BigTwo).CompareTo($d)
    if ($compare -gt 0 -or ($compare -eq 0 -and -not $m.IsEven)) { $m = $m + [System.Numerics.BigInteger]::One }
    if ($shift -ge 0) {
        return @{ Num = $m; Den = [System.Numerics.BigInteger]::Pow($script:BigTwo, $shift) }
    }
    return @{ Num = $m * [System.Numerics.BigInteger]::Pow($script:BigTwo, - $shift); Den = [System.Numerics.BigInteger]::One }
}

function Get-Digits15([System.Numerics.BigInteger]$Num, [System.Numerics.BigInteger]$Den) {
    # 양수 Num/Den 을 유효 15자리로 반올림(꼭 반이면 짝수쪽)한 [decimal]. 지금 판 format(x, '.15g') 와 같다. decimal 에 못 담으면 $null.
    $low = [System.Numerics.BigInteger]::Pow($script:BigTen, $script:ExcelDigits - 1)
    $high = [System.Numerics.BigInteger]::Pow($script:BigTen, $script:ExcelDigits)
    $k = ($script:ExcelDigits - 1) - ($Num.ToString().Length - $Den.ToString().Length)
    while ($true) {
        if ($k -ge 0) {
            $n = $Num * [System.Numerics.BigInteger]::Pow($script:BigTen, $k)
            $d = $Den
        } else {
            $n = $Num
            $d = $Den * [System.Numerics.BigInteger]::Pow($script:BigTen, - $k)
        }
        $q = [System.Numerics.BigInteger]::Divide($n, $d)
        if ($q -lt $low) { $k++ } elseif ($q -ge $high) { $k-- } else { break }
    }
    $compare = (($n - $q * $d) * $script:BigTwo).CompareTo($d)
    if ($compare -gt 0 -or ($compare -eq 0 -and -not $q.IsEven)) { $q = $q + [System.Numerics.BigInteger]::One }
    if ($k -gt 28) { return $null }   # decimal 은 소수 28자리까지
    try {
        return [decimal]::Parse("${q}E$(- $k)", [Globalization.NumberStyles]::Float, $script:Invariant)
    } catch {
        return $null
    }
}

function Remove-TrailingZeros([decimal]$Value) {
    # 소수 뒤의 0 을 뺀다(1.50 → 1.5, 100.0 → 100). 파이썬 '.15g' 글자와 같은 모양.
    $text = $Value.ToString($script:Invariant)
    if (-not $text.Contains('.')) { return $Value }
    return [decimal]::Parse($text.TrimEnd('0').TrimEnd('.'), $script:Invariant)
}

function ConvertTo-ExcelNumber($Value) {
    # 엑셀 칸의 숫자를 [decimal] 로(지금 판 excel_number). 정수는 그대로, 소수는 유효 15자리로 맞추고 뒤의 0 은 뺀다.
    #   [double]  : 엑셀 Value2 그대로. 비트에서 정확한 분수를 뽑아 반올림한다.
    #   [decimal] : JSON 에서 읽은 소수(5.1 의 ConvertFrom-Json 은 소수점이 있는 숫자를 decimal 로 줌). 적힌 글자가 파이썬 float 의
    #               가장 짧은 글자이므로 15자리 이하면 그대로, 더 길면 가장 가까운 double 의 실제 값으로 반올림한다.
    # 숫자가 아니거나 참거짓이면 $null.
    if ($null -eq $Value -or $Value -is [bool]) { return $null }
    if ($Value -is [int] -or $Value -is [long]) { return [decimal]$Value }
    $negative = $false
    if ($Value -is [double]) {
        if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) { return $null }
        if ($Value -eq 0) { return [decimal]0 }
        $negative = $Value -lt 0
        $fraction = Get-DoubleFraction ([Math]::Abs($Value))
    } elseif ($Value -is [decimal]) {
        $negative = $Value -lt 0
        $unsigned = [Math]::Abs($Value).ToString($script:Invariant)
        $digits = $unsigned.Replace('.', '')
        if ($digits.Trim('0').Length -le $script:ExcelDigits) { return Remove-TrailingZeros $Value }
        $point = $unsigned.IndexOf('.')
        $places = if ($point -lt 0) { 0 } else { $unsigned.Length - $point - 1 }
        $fraction = Get-NearestDouble ([System.Numerics.BigInteger]::Parse($digits)) ([System.Numerics.BigInteger]::Pow($script:BigTen, $places))
    } else {
        return $null
    }
    $rounded = Get-Digits15 $fraction.Num $fraction.Den
    if ($null -eq $rounded) { return $null }
    $rounded = Remove-TrailingZeros $rounded
    if ($negative) { return - $rounded }
    return $rounded
}
