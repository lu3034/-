@echo off
:: �������ԱȨ��
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo �������ԱȨ��...
    set "params=%*"
    cd /d "%~dp0" && cd ..
    mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe", "/c ""%~0"" %params%", "", "runas", 1)(window.close)&&exit
)

:: ������ȷ��ִ�в��Բ����нű�
set "scriptPath=%~dp0FileBackupMonitor.ps1"
echo ��������ִ�в���...
PowerShell.exe -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force" >nul
echo �����ļ����ݼ����...
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%scriptPath%"
pause