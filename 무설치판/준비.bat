@echo off
chcp 65001 > nul
rem 준비: 처음 한 번, 그리고 대표에게서 새 묶음을 받아 덮어쓴 뒤에 두 번 클릭한다.
rem 자체 시험 → 작업마다 입력 · 출력 폴더 고르기(처음만) → 매일 실행 등록 · 바탕화면 「지금 실행」 을 만든다.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0무설치판.ps1" -Command 준비
pause
