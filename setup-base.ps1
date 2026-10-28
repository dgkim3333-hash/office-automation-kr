# 02장 기본 설치: Claude Desktop · Node.js · uv·uvx
# 사용법: [Net.ServicePointManager]::SecurityProtocol='Tls12'; irm https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/setup-base.ps1 | iex
[Net.ServicePointManager]::SecurityProtocol = 'Tls12'
$installScript = Invoke-RestMethod 'https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/install.ps1'
& ([scriptblock]::Create($installScript)) -Part base
