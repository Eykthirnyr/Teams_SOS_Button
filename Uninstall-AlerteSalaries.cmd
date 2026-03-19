@echo off
setlocal EnableExtensions

set "APP_NAME=Alerte Salaries"
set "INSTALL_DIR_X86=%ProgramFiles(x86)%\AlerteSalaries"
set "INSTALL_DIR_X64=%ProgramFiles%\AlerteSalaries"
set "PROGRAMDATA_DIR=%ProgramData%\AlerteSalaries"
set "PUBLIC_DESKTOP=%Public%\Desktop"
set "COMMON_STARTMENU_DIR=%ProgramData%\Microsoft\Windows\Start Menu\Programs\Alerte Salaries"

echo Desinstallation de %APP_NAME%...

for /f "tokens=2*" %%A in ('reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "Alerte Salaries" 2^>nul ^| findstr /i "UninstallString"') do (
    set "UNINSTALL_CMD=%%B"
    goto :found_uninstall
)

for /f "tokens=2*" %%A in ('reg query "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "Alerte Salaries" 2^>nul ^| findstr /i "UninstallString"') do (
    set "UNINSTALL_CMD=%%B"
    goto :found_uninstall
)

goto :remove_files

:found_uninstall
echo MSI detecte. Lancement de la desinstallation...
call :normalize_uninstall
if defined UNINSTALL_CMD (
    %UNINSTALL_CMD% /qn /norestart
)

:remove_files
echo Suppression des raccourcis...
del /f /q "%PUBLIC_DESKTOP%\Alerte Salaries.lnk" 2>nul
del /f /q "%PUBLIC_DESKTOP%\Configurer Alerte Salaries.lnk" 2>nul
del /f /q "%COMMON_STARTMENU_DIR%\Alerte Salaries.lnk" 2>nul
rmdir /s /q "%COMMON_STARTMENU_DIR%" 2>nul

echo Suppression des fichiers applicatifs...
rmdir /s /q "%INSTALL_DIR_X86%" 2>nul
if /i not "%INSTALL_DIR_X64%"=="%INSTALL_DIR_X86%" (
    rmdir /s /q "%INSTALL_DIR_X64%" 2>nul
)
rmdir /s /q "%PROGRAMDATA_DIR%" 2>nul

echo Suppression des preferences utilisateur...
for /d %%U in ("C:\Users\*") do (
    if exist "%%~fU\AppData\Roaming\AlerteSalaries" (
        rmdir /s /q "%%~fU\AppData\Roaming\AlerteSalaries" 2>nul
    )
)

echo Nettoyage termine.
exit /b 0

:normalize_uninstall
setlocal EnableDelayedExpansion
set "CMD=!UNINSTALL_CMD!"
set "CMD=!CMD:/I=/x!"
set "CMD=!CMD:/X=/x!"
if /i "!CMD:~0,7!"=="MsiExec" (
    endlocal & set "UNINSTALL_CMD=%CMD%" & goto :eof
)
endlocal
goto :eof
