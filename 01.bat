@echo off
setlocal enabledelayedexpansion

:: ====== Ȩ�޼�� ======
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo ������Ҫ����ԱȨ�ޣ�
    echo ���Ҽ�����˽ű���ѡ��"�Թ���Ա�������"
    timeout /t 5
    exit /b 1
)

:: ====== ����Ŀ���ļ��� ======
set "target_folder=C:\Users\LU\Desktop\KM"

:: ���·����Ч��
if not exist "%target_folder%" (
    echo ���ڴ���Ŀ���ļ��У�%target_folder%
    mkdir "%target_folder%"
    if errorlevel 1 (
        echo �����޷������ļ��У�����·��Ȩ��
        pause
        exit /b 1
    )
)

:: ====== ׼��PowerShell���� ======
powershell -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force" >nul 2>&1
if errorlevel 1 (
    echo �����޷�����PowerShellִ�в���
    pause
    exit /b 1
)

:: ������һ�μ��������ʱ�ļ�
set "last_clip=%temp%\last_clip.txt"
echo. > "%last_clip%"

:: ====== ������ ======
echo.
echo [�ļ����ݼ���� ������]
echo Ŀ��Ŀ¼: %target_folder%
echo ����˵��: ���ļ���Դ�����������ļ����Զ�ճ����Ŀ���ļ���
echo �˳�����: �� Ctrl+C Ȼ�� Y
echo.

:loop
:: ��ȡ����������
set "new_data="
powershell -noprofile -command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetFileDropList()" > "%temp%\clip.txt" 2>nul

:: ����ļ��Ƿ���������
if exist "%temp%\clip.txt" (
    fc "%temp%\clip.txt" "%last_clip%" >nul 2>&1
    if errorlevel 1 set "new_data=1"
)

:: ������������
if defined new_data (
    set "first=1"
    for /f "usebackq delims=" %%f in ("%temp%\clip.txt") do (
        if defined first (
            echo [%time%] ��⵽���ļ���
            set "first="
        )
        echo   ���ڸ���: %%~nxf
        
        :: �ж�·�����Ͳ�����
        if exist "%%f\" (
            robocopy "%%f" "%target_folder%\%%~nxf" /e /is /it /np /njh /njs >nul 2>&1
        ) else (
            copy /y "%%f" "%target_folder%\" >nul
        )
    )
    echo ��ɣ��ļ��ѱ��浽Ŀ���ļ���
    echo.
    
    :: ���浱ǰ������״̬
    copy "%temp%\clip.txt" "%last_clip%" >nul 2>&1
)

del "%temp%\clip.txt" >nul 2>&1
timeout /t 1 /nobreak >nul
goto loop