@echo off
setlocal
set "PROJECT_ROOT=%~dp0"
set "PYTHONPATH=%PROJECT_ROOT%src;%PYTHONPATH%"
set "PYTHONUTF8=1"
if exist "%PROJECT_ROOT%.venv\Scripts\python.exe" (
  "%PROJECT_ROOT%.venv\Scripts\python.exe" -m thesis_format_checker %*
) else (
  python -m thesis_format_checker %*
)
exit /b %ERRORLEVEL%
