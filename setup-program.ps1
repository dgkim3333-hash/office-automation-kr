# 03장 프로그램 설치: 사무 자동화 프로그램 · 작업 폴더 · 기본 약속(전역 CLAUDE.md). 다시 실행하면 업데이트.
# 사용법: [Net.ServicePointManager]::SecurityProtocol='Tls12'; irm https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/setup-program.ps1 | iex
[Net.ServicePointManager]::SecurityProtocol = 'Tls12'
$installScript = Invoke-RestMethod 'https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/install.ps1'
& ([scriptblock]::Create($installScript)) -Part program
