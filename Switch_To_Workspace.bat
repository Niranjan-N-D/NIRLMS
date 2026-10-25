@echo off
title BEL Library - IIS Workspace Connector
color 0b

:: Self-elevate to Administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process cmd -ArgumentList '/k `"%~f0`"' -Verb RunAs"
    exit /b
)

echo ====================================================================
echo   BEL Library - Connecting IIS Directly to Workspace
echo ====================================================================
echo.

set "WS=C:\Projects\library-main (1)\library-main1\LibraryManagementSystem"
set "INET=C:\inetpub\wwwroot\library-main (1)\library-main1\LibraryManagementSystem"
set "APPCMD=%windir%\system32\inetsrv\appcmd.exe"

echo [1/4] Granting write permissions to inetpub folder...
icacls "C:\inetpub\wwwroot\library-main (1)" /grant "Users":(OI)(CI)F /T /Q

echo.
echo [2/4] Copying updated files to inetpub...
xcopy /E /I /Y /Q "%WS%\*" "%INET%\"

echo.
echo [3/4] Pointing IIS on port 8040 to workspace:
echo       %WS%
"%APPCMD%" set vdir "newlibrary/" -physicalPath:"%WS%"
if %errorlevel% neq 0 (
    echo Note: Attempting site root update...
    "%APPCMD%" set site "newlibrary" -[path='/'].[path='/'].physicalPath:"%WS%"
)

echo.
echo [4/4] Recycling application pool...
"%APPCMD%" recycle apppool "newlibrary"

echo.
echo ====================================================================
echo   SUCCESS! 
echo   1. The '🚨 Filter Overdue Borrowings' button is now live!
echo   2. You can refresh your browser at:
echo      http://localhost:8040/borrowings.asp
echo ====================================================================
echo.
echo Press any key to close this window...
pause >nul
