# 결과 알림(지금 판 app\notify.py 와 같은 글자 · 같은 방법). PC 안에서만 뜬다.
#   - 기본: Windows 알림(WinRT ToastNotificationManager, 외부 모듈 없음). 누르면 출력 폴더(결과가 0건이면 작업 폴더)가 열린다.
#   - 알림이 꺼져 있거나 실패하면 메시지 창(user32 MessageBoxW)으로 대신한다. 둘 다 실패해도 오류를 밖으로 내지 않는다.

$script:NoticeTitle = '사무자동화'
$script:NoticeFailureTitle = '사무자동화: 자동 실행 실패'
# Windows PowerShell 의 앱 ID. 따로 앱을 등록하지 않아도 이 이름으로 알림을 띄울 수 있다
$script:ToastAppId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
$script:NoticeLabels = @(@('성공', '결과'), @('중복', '중복'), @('확인필요', '확인필요'))   # 지금 판 LABELS 에서 승인대기를 뺌
$script:XlsHint = '옛 엑셀(.xls) {0}개: 엑셀에서 열어 .xlsx 로 다시 저장하세요'
$script:MbIconInformation = 0x40
$script:MbIconWarning = 0x30
$script:MbTopmost = 0x40000

function Get-NoticeValue($Object, [string]$Name, $Default) {
    # 해시표 · PSCustomObject(JSON) 어느 쪽이든 값 하나. 없으면 기본값.
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $Default
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $Default }
    return $p.Value
}

function Get-ReportText([string]$Title, $Report) {
    # 지금 판 report_text: 예 「일일 PO 정리: 결과 3건, 확인필요 1건」. 0인 항목은 뺀다. 경고는 줄마다 「경고: 」.
    $parts = New-Object System.Collections.Generic.List[string]
    if ([bool](Get-NoticeValue $Report 'stopped' $false)) { $parts.Add('실행이 중간에 멈췄습니다(작업 폴더의 실행기록.txt 를 보세요)') }
    $counts = Get-NoticeValue $Report 'counts' $null
    foreach ($pair in $script:NoticeLabels) {
        $n = [int](Get-NoticeValue $counts $pair[0] 0)
        if ($n) { $parts.Add("$($pair[1]) ${n}건") }
    }
    if ([bool](Get-NoticeValue $Report 'ledger_locked' $false)) { $parts.Add('누적장부가 엑셀에서 열려 있어 보류(엑셀을 닫으세요)') }
    $xls = [int](Get-NoticeValue $Report 'unmatched_xls' 0)
    $others = [int](Get-NoticeValue $Report 'unmatched_old' 0) - $xls
    if ($others -gt 0) { $parts.Add("어느 작업에도 맞지 않는 파일 ${others}개") }
    if ($xls -gt 0) { $parts.Add(($script:XlsHint -f $xls)) }
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("${Title}: " + [string]::Join(', ', $parts.ToArray()))
    foreach ($a in @(Get-NoticeValue $Report 'alerts' @())) { if ($null -ne $a) { $lines.Add("경고: $a") } }
    return [string]::Join("`n", $lines.ToArray())
}

function ConvertTo-XmlEscaped([string]$Text, [switch]$Attribute) {
    # 파이썬 xml.sax.saxutils.escape(& < >, 속성이면 " 도).
    $text = $Text.Replace('&', '&amp;').Replace('>', '&gt;').Replace('<', '&lt;')
    if ($Attribute) { $text = $text.Replace('"', '&quot;') }
    return $text
}

function ConvertTo-FileUri([string]$Folder) {
    # 파이썬 Path(folder).as_uri(): file:///C:/… , 드라이브 뒤는 UTF-8 바이트를 %XX 로(영숫자 · _.-~ · / 는 그대로).
    $full = [IO.Path]::GetFullPath($Folder)
    if ($full.Length -gt 3) { $full = $full.TrimEnd('\') }
    $drive = $full.Substring(0, 2)
    $rest = $full.Substring(2).Replace('\', '/')
    $text = New-Object Text.StringBuilder
    foreach ($b in [Text.Encoding]::UTF8.GetBytes($rest)) {
        $c = [char]$b
        if (($b -lt 128) -and ([char]::IsLetterOrDigit($c) -or '_.-~/'.IndexOf($c) -ge 0)) { [void]$text.Append($c) }
        else { [void]$text.Append('%' + $b.ToString('X2')) }
    }
    return 'file:///' + $drive + $text.ToString()
}

function Get-ToastXml([string]$Title, [string]$Body, [string]$Folder) {
    # 지금 판 toast_script 가 LoadXml 에 넣는 XML 과 같은 글자.
    $launch = ''
    if (-not [string]::IsNullOrEmpty($Folder)) {
        $launch = ' activationType="protocol" launch="' + (ConvertTo-XmlEscaped (ConvertTo-FileUri $Folder) -Attribute) + '"'
    }
    return "<toast$launch><visual><binding template=`"ToastGeneric`"><text>$(ConvertTo-XmlEscaped $Title)</text><text>$(ConvertTo-XmlEscaped $Body)</text></binding></visual></toast>"
}

function Show-Toast([string]$Title, [string]$Body, [string]$Folder) {
    [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $xml.LoadXml((Get-ToastXml $Title $Body $Folder))
    $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
    $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($script:ToastAppId)
    # 사용자나 회사 정책이 알림을 꺼 두면 Show 는 성공해도 아무것도 안 보인다. 그러면 실패로 돌려 메시지 창으로
    if ($notifier.Setting -ne 'Enabled') { throw "Windows 알림이 꺼져 있음($($notifier.Setting))" }
    $notifier.Show($toast)
}

function Show-MessageBox([string]$Title, [string]$Body, [int]$Icon) {
    if (-not ('NoInstall.User32' -as [type])) {
        Add-Type -Namespace NoInstall -Name User32 -MemberDefinition @'
[DllImport("user32.dll", CharSet = CharSet.Unicode)]
public static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);
'@
    }
    [void][NoInstall.User32]::MessageBoxW([IntPtr]::Zero, $Body, $Title, [uint32]($Icon -bor $script:MbTopmost))
}

function Show-Notice([string]$Title, [string]$Body, [string]$Folder, [int]$Icon = $script:MbIconInformation) {
    # 알림을 띄운다. 실패하면 메시지 창. 둘 다 실패해도 오류를 내지 않는다(실행 결과는 이미 기록됨).
    try { Show-Toast $Title $Body $Folder; return } catch { }
    try { Show-MessageBox $Title $Body $Icon } catch { }
}

function Send-ReportNotice($Job, $Report) {
    # 실행 결과 알림. 알릴 것이 없으면(has_news 거짓) 아무것도 하지 않는다. 실행 잠금을 푼 뒤에 부른다.
    if (-not [bool](Get-NoticeValue $Report 'has_news' $false)) { return }
    $counts = Get-NoticeValue $Report 'counts' $null
    $folder = if ([int](Get-NoticeValue $counts '성공' 0) -gt 0) { $Job.output_dir } else { $Job.folder }
    Show-Notice $script:NoticeTitle (Get-ReportText $Job.title $Report) $folder
}

function Send-FailureNotice([string]$What, [string]$Message, [string]$Folder) {
    # 자동 실행이 작업을 찾기도 전에 실패했을 때(지금 판 notify_failure).
    Show-Notice $script:NoticeFailureTitle "${What}: $Message" $Folder $script:MbIconWarning
}

function Invoke-NotifyOp($Payload) {
    # 엔진 op notify(시험용, 알림은 띄우지 않는다):
    #   text      {title, report}        → {text}   (Get-ReportText)
    #   toast_xml {title, body, folder?} → {xml}    (Get-ToastXml)
    switch ([string]$Payload.action) {
        'text' { return ConvertTo-Json -InputObject @{ text = (Get-ReportText ([string]$Payload.title) $Payload.report) } -Depth 20 -Compress }
        'toast_xml' {
            $folder = if ($null -ne $Payload.PSObject.Properties['folder']) { [string]$Payload.folder } else { '' }
            return ConvertTo-Json -InputObject @{ xml = (Get-ToastXml ([string]$Payload.title) ([string]$Payload.body) $folder) } -Depth 20 -Compress
        }
        default { throw "notify action '$($Payload.action)' 는 없습니다." }
    }
}
