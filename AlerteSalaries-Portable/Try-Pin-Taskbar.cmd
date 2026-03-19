@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PIN_VBS=%SCRIPT_DIR%Try-Pin-Taskbar.vbs"

if not exist "%PIN_VBS%" (
  echo Script d'epinglage introuvable: "%PIN_VBS%"
  exit /b 1
)

wscript.exe "%PIN_VBS%"
exit /b %ERRORLEVEL%
