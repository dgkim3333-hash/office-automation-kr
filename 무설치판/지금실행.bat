@echo off
chcp 65001 > nul
rem 지금 실행: 작업 번호(예: 01)를 받아 그 작업을 한 번 실행한다. 바탕화면 「지금 실행 - 제목」 바로가기가 부른다.
rem 작업 폴더가 묶음 폴더가 아니면 바로가기가 뒤에 -Work "작업 폴더" 를 붙인다(%2 %3).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0무설치판.ps1" -Command 실행 -Job %1 %2 %3
pause
