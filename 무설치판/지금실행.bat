@echo off
chcp 65001 > nul
rem 지금 실행: 작업 번호(예: 01) 하나를 받아 그 작업을 한 번 실행한다. 바탕화면 「지금 실행 - 제목」 바로가기가 부른다.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0무설치판.ps1" -Command 실행 -Job %1
pause
