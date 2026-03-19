@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "APP_EXE=%SCRIPT_DIR%AlerteSalaries.exe"
set "SHORTCUT_VBS=%SCRIPT_DIR%Create-Shortcut.vbs"
set "DESKTOP_DIR=%USERPROFILE%\Desktop"

if not exist "%APP_EXE%" (
  echo Executable introuvable: "%APP_EXE%"
  exit /b 1
)

if not exist "%SHORTCUT_VBS%" (
  echo Helper raccourci introuvable: "%SHORTCUT_VBS%"
  exit /b 2
)

wscript.exe "%SHORTCUT_VBS%" "%DESKTOP_DIR%\Alerte Salaries.lnk" "%APP_EXE%" "" "%APP_EXE%" "Declencher une alerte securite discrete."
wscript.exe "%SHORTCUT_VBS%" "%DESKTOP_DIR%\Configurer Alerte Salaries.lnk" "%APP_EXE%" "/configure" "%APP_EXE%" "Modifier la configuration de l'application d'alerte."

echo Raccourcis crees sur le Bureau: %DESKTOP_DIR%
exit /b 0
