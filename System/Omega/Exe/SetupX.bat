@echo off
echo.
echo.
echo Installing....
 

@echo off

ver | find "XP" > nul
if %ERRORLEVEL% == 0 goto ver_xp

if not exist %SystemRoot%\system32\systeminfo.exe goto warnthenexit

systeminfo | find "OS Name" > %TEMP%\osname.txt
FOR /F "usebackq delims=: tokens=2" %%i IN (%TEMP%\osname.txt) DO set vers=%%i

echo %vers% | find "Windows 10" > nul
if %ERRORLEVEL% == 0 goto ver_10

echo %vers% | find "Windows 8" > nul
if %ERRORLEVEL% == 0 goto ver_8

echo %vers% | find "Windows 7" > nul
if %ERRORLEVEL% == 0 goto ver_7

echo %vers% | find "Windows Vista" > nul
if %ERRORLEVEL% == 0 goto ver_vista

goto warnthenexit

:ver_10
:Run Windows 10 specific commands here.
echo. 
echo.
echo Windows 10
goto check_version

:ver_8
:Run Windows 8 specific commands here.
echo. 
echo.
echo Windows 8
goto check_version

:ver_7
:Run Windows 7 specific commands here.
echo. 
echo.
echo Windows 7
goto check_version

:ver_vista
:Run Windows Vista specific commands here.
echo. 
echo.
echo Windows Vista
goto check_version

:ver_xp
:Run Windows XP specific commands here.
echo. 
echo.
echo Windows XP
goto check_version

:warnthenexit
echo Machine undetermined.
goto pause1


:check_version

@echo off
 
Set RegQry=HKLM\Hardware\Description\System\CentralProcessor\0
 
REG.exe Query %RegQry% > %TEMP%\checkOS.txt
 
Find /i "x86" < %TEMP%\CheckOS.txt > %TEMP%\StringCheck.txt
 
If %ERRORLEVEL% == 0 (
    	Echo.32 Bit Operating system
	goto 32_bit
) ELSE (
    	Echo.64 Bit Operating System
   	goto 64_bit    	
)


:32_bit
(copy \\192.168.1.2\Dll\*.* c:\windows\system32\
regsvr32.exe c:\windows\system32\comdlg32.ocx /s
regsvr32.exe c:\windows\system32\mscomctl.ocx /s
regsvr32.exe c:\windows\system32\mscomm32.ocx /s
regsvr32.exe c:\windows\system32\msflxgrd.ocx /s
regsvr32.exe c:\windows\system32\msmask32.ocx /s
regsvr32.exe c:\windows\system32\richtx32.ocx /s
regsvr32.exe c:\windows\system32\prjXTab.ocx /s
regsvr32.exe c:\windows\system32\lvButton.ocx /s
regsvr32.exe c:\windows\system32\HoverButton.ocx /s)>NUL
goto installprogram

:64_bit
(copy \\192.168.1.2\Dll\*.* c:\windows\SysWOW64\
regsvr32.exe c:\windows\SysWOW64\comdlg32.ocx /s
regsvr32.exe c:\windows\SysWOW64\mscomctl.ocx /s
regsvr32.exe c:\windows\SysWOW64\mscomm32.ocx /s
regsvr32.exe c:\windows\SysWOW64\msflxgrd.ocx /s
regsvr32.exe c:\windows\SysWOW64\msmask32.ocx /s
regsvr32.exe c:\windows\SysWOW64\richtx32.ocx /s
regsvr32.exe c:\windows\SysWOW64\prjXTab.ocx /s
regsvr32.exe c:\windows\SysWOW64\lvButton.ocx /s
regsvr32.exe c:\windows\SysWOW64\HoverButton.ocx /s)>NUL
goto installprogram


:installprogram
if NOT exist "D:\Omega" (md D:\Omega)
if NOT exist "D:\Omega\Reports" (md D:\Omega\Reports)
(
copy \\192.168.1.2\Omega\Omega.exe D:\Omega\
copy \\192.168.1.2\Omega\Reports\*.* D:\Omega\Reports\
copy \\192.168.1.2\Omega\Update.bat D:\Omega\
copy \\192.168.1.2\Omega\Setup.bat D:\Omega\
)>NUL

REM install Crystal Report
Echo.
Echo.
echo.Check Crystal Report

Reg Query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /S> %Temp%\Installedsoftware.txt
Find /i "Crystal Reports" < %TEMP%\Installedsoftware.txt > %TEMP%\CheckInstalledsoftware.txt
If %ERRORLEVEL% == 0 (
	goto CreateShortcut
) ELSE (
	goto InstallCrystal
)

:InstallCrystal
echo.
echo.
echo.Installing Crystal Report 8.5
\\192.168.1.2\Installer\CrystalReport8.5\scrdev.msi /qr PIDKEY=A6A50-8900008-ZE1007S INSTALLLEVEL=3



:CreateShortcut
@Echo off
echo.
echo.
echo.Create Desktop Shortcut
rem Window Style
REM 1 = Normal, 3 Maximized, 7 = Minimized

rem Choose "Desktop" or "AllUsersDesktop"
set Location="AllUsersDesktop"

set DisplayName="Omega"
set filename="D:\Omega\Omega.exe"

REM Set icon to an icon from an exe or "something.ico"
set icon="D:\Omega\Omega.exe, 0"

set WorkingDir="D:\Omega"

set Arguments=""

REM Make temporary VBS file to create shortcut
REM Then execute and delete it

(echo Dim DisplayName,Location,Path,shell,link
echo Set shell = CreateObject^("WScript.shell"^)
echo path = shell.SpecialFolders^(%Location%^)
echo Set link = shell.CreateShortcut^(path ^& "\" ^& %DisplayName% ^& ".lnk"^)

echo link.Description = %DisplayName%
echo link.TargetPath = %filename%
echo link.Arguments = %arguments%
echo link.HotKey = "ALT+CTRL+O"

echo link.WindowStyle = 1
echo link.IconLocation = %icon%

echo link.WorkingDirectory = %WorkingDir%
echo link.Save

)> "%temp%\makelink.vbs"
cscript //nologo "%temp%\makelink.vbs"
del "%temp%\makelink.vbs" 2>NUL

@Echo off
rem Window Style
REM 1 = Normal, 3 Maximized, 7 = Minimized

rem Choose "Desktop" or "AllUsersDesktop"
set Location="AllUsersDesktop"

set DisplayName="Update Omega"
set filename="D:\Omega\Update.bat"

REM Set icon to an icon from an exe or "something.ico"
set icon="D:\Omega\Update.bat, 0"

set WorkingDir="D:\Omega"

set Arguments=""

REM Make temporary VBS file to create shortcut
REM Then execute and delete it

(echo Dim DisplayName,Location,Path,shell,link
echo Set shell = CreateObject^("WScript.shell"^)
echo path = shell.SpecialFolders^(%Location%^)
echo Set link = shell.CreateShortcut^(path ^& "\" ^& %DisplayName% ^& ".lnk"^)

echo link.Description = %DisplayName%
echo link.TargetPath = %filename%
echo link.Arguments = %arguments%

echo link.WindowStyle = 1
echo link.IconLocation = %icon%

echo link.WorkingDirectory = %WorkingDir%
echo link.Save

)> "%temp%\makelink.vbs"
cscript //nologo "%temp%\makelink.vbs"
del "%temp%\makelink.vbs" 2>NUL


:pause
echo.
echo.
echo All Done !!!! 
echo.
echo.
pause
goto exit


:pause1
pause

:exit
