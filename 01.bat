@echo off
setlocal enabledelayedexpansion

:: ====== 权限检查 ======
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo 错误：需要管理员权限！
    echo 请右键点击此脚本，选择"以管理员身份运行"
    timeout /t 5
    exit /b 1
)

:: ====== 配置目标文件夹 ======
set "target_folder=C:\Users\LU\Desktop\KM"

:: 检查路径有效性
if not exist "%target_folder%" (
    echo 正在创建目标文件夹：%target_folder%
    mkdir "%target_folder%"
    if errorlevel 1 (
        echo 错误：无法创建文件夹！请检查路径权限
        pause
        exit /b 1
    )
)

:: ====== 准备PowerShell环境 ======
powershell -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force" >nul 2>&1
if errorlevel 1 (
    echo 错误：无法设置PowerShell执行策略
    pause
    exit /b 1
)

:: 创建上一次剪贴板的临时文件
set "last_clip=%temp%\last_clip.txt"
echo. > "%last_clip%"

:: ====== 主程序 ======
echo.
echo [文件备份监控器 已启动]
echo 目标目录: %target_folder%
echo 操作说明: 在文件资源管理器复制文件后自动粘贴到目标文件夹
echo 退出方法: 按 Ctrl+C 然后按 Y
echo.

:loop
:: 获取剪贴板内容
set "new_data="
powershell -noprofile -command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetFileDropList()" > "%temp%\clip.txt" 2>nul

:: 检查文件是否有新内容
if exist "%temp%\clip.txt" (
    fc "%temp%\clip.txt" "%last_clip%" >nul 2>&1
    if errorlevel 1 set "new_data=1"
)

:: 有新内容则处理
if defined new_data (
    set "first=1"
    for /f "usebackq delims=" %%f in ("%temp%\clip.txt") do (
        if defined first (
            echo [%time%] 检测到新文件：
            set "first="
        )
        echo   正在复制: %%~nxf
        
        :: 判断路径类型并复制
        if exist "%%f\" (
            robocopy "%%f" "%target_folder%\%%~nxf" /e /is /it /np /njh /njs >nul 2>&1
        ) else (
            copy /y "%%f" "%target_folder%\" >nul
        )
    )
    echo 完成！文件已保存到目标文件夹
    echo.
    
    :: 保存当前剪贴板状态
    copy "%temp%\clip.txt" "%last_clip%" >nul 2>&1
)

del "%temp%\clip.txt" >nul 2>&1
timeout /t 1 /nobreak >nul
goto loop