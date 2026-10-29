# 값 바꾸기: 글자 → 금액([decimal]) · 날짜([datetime]), 그리고 결과 JSON 표기.
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
    if ($Value -is [DateTime]) { return $Value.ToString('yyyy-MM-dd', $script:Invariant) }
    return $Value
}
