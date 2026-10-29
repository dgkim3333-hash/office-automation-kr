# 규칙 읽기: 대표 PC 에서 export-portable 로 만든 규칙.json(app\portable.py) 을 읽고 정규식을 미리 컴파일해 붙인다.
# 필드에는 .regex, 표에는 .row_regex 가 붙는다. 컴파일하지 못하면 「규칙 정규식 오류: <이름>」 으로 멈춘다.

# 파이썬 re 처럼 옵션 없이(대소문자 무시는 패턴 안의 (?i)). 대소문자 비교만 문화권에 흔들리지 않게 한다.
$script:RegexOptions = [Text.RegularExpressions.RegexOptions]::CultureInvariant

function New-RuleRegex([string]$Pattern, [string]$Name) {
    try {
        return New-Object Text.RegularExpressions.Regex($Pattern, $script:RegexOptions)
    } catch {
        throw "규칙 정규식 오류: $Name"
    }
}

function ConvertTo-PortableRules($Data) {
    # ConvertFrom-Json 으로 읽은 규칙 객체에 컴파일한 정규식을 붙여 그대로 돌려준다.
    foreach ($doc in @($Data.doc_types)) {
        foreach ($field in @($doc.fields)) {
            $regex = New-RuleRegex $field.pattern "[$($doc.name)] $($field.name)"
            $field | Add-Member -NotePropertyName regex -NotePropertyValue $regex -Force
        }
        if ($null -ne $doc.table) {
            $regex = New-RuleRegex $doc.table.row_pattern "[$($doc.name)] 표 행패턴"
            $doc.table | Add-Member -NotePropertyName row_regex -NotePropertyValue $regex -Force
        }
    }
    return $Data
}

function Import-PortableRules([string]$Path) {
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    return ConvertTo-PortableRules ($text | ConvertFrom-Json)
}
