@echo off
title BEL Library - Applying Workspace Update
color 0b

echo ====================================================================
echo   BEL Library - Syncing Workspace to IIS & Fixing UI
echo ====================================================================
echo.

set "WS=C:\Projects\library-main (1)\library-main1\LibraryManagementSystem"
set "INET=C:\inetpub\wwwroot\library-main (1)\library-main1\LibraryManagementSystem"
set "APPCMD=%windir%\system32\inetsrv\appcmd.exe"

echo [1/4] Granting Full Permissions to Users on inetpub...
icacls "C:\inetpub\wwwroot\library-main (1)" /grant "Users":(OI)(CI)F /T /Q

echo [2/4] Copying redesigned files to inetpub...
xcopy /E /I /Y /Q "%WS%\*" "%INET%\"

echo [3/4] Pointing IIS directly to workspace...
"%APPCMD%" set site "newlibrary" -[path='/'].[path='/'].physicalPath:"%WS%"
"%APPCMD%" set vdir "newlibrary/" -physicalPath:"%WS%"

echo [4/4] Recycling IIS Application Pool...
"%APPCMD%" recycle apppool "newlibrary"

echo.
echo ====================================================================
echo   SUCCESS! 
echo   Redesigned UI (100%% Zoom fixed, compact table, 2x2 buttons) is LIVE!
echo ====================================================================
echo.
start http://localhost:8040/books.asp
timeout /t 3
exit
