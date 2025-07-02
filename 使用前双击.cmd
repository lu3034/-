@echo off
:: 请求管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo 请求管理员权限...
    set "params=%*"
    cd /d "%~dp0" && cd ..
    mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe", "/c ""%~0"" %params%", "", "runas", 1)(window.close)&&exit
)

:: 设置正确的执行策略并运行脚本
set "scriptPath=%~dp0FileBackupMonitor.ps1"
echo 正在设置执行策略...
PowerShell.exe -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force" >nul
echo 启动文件备份监控器...
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%scriptPath%"
pause