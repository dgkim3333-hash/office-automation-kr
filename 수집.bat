@echo off
chcp 65001 > nul
cd /d "%~dp0"
uv run --extra outlook python -m app outlook
pause
