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
echo. 
echo.
echo Installing/Registering DLL . . . 

(
copy \\192.168.1.2\Omega\Dll\*.* %windir%\system32\
if errorlevel 1 goto mapdrive32
if NOT exist "C:\Omega" (md C:\Omega)
if NOT exist "C:\Omega\Reports" (md C:\Omega\Reports)
copy \\192.168.1.2\Omega\Omega.exe C:\Omega\
copy \\192.168.1.2\Omega\Reports\*.* C:\Omega\Reports\
copy \\192.168.1.2\Omega\Config C:\Omega\
copy \\192.168.1.2\Omega\Update.bat C:\Omega\
copy \\192.168.1.2\Omega\Setup.bat C:\Omega\
)>NUL

echo.
echo.
echo.Checking Crystal Report 8.5
SET ProgFilesRoot=%ProgramFiles%
IF EXIST "%ProgFilesRoot%\Seagate Software\Crystal Reports\crw32.exe" goto regdll32

echo.
echo.
echo.Installing Crystal Report 8.5
(
\\192.168.1.2\Omega\CrystalReport8_5\scrdev.msi /qr PIDKEY=A6A50-8900008-ZE1007S INSTALLLEVEL=3
)>NUL

goto regdll32

:mapdrive32
(
net use b: \\192.168.1.2\Omega 123456 /user:update
copy b:\Dll\*.* %windir%\system32\
if NOT exist "C:\Omega" (md C:\Omega)
if NOT exist "C:\Omega\Reports" (md C:\Omega\Reports)
copy b:\Omega.exe C:\Omega\
copy b:\Reports\*.* C:\Omega\Reports\
copy b:\Config C:\Omega\
copy b:\Update.bat C:\Omega\
copy b:\Setup.bat C:\Omega\
)>NUL

echo.
echo.
echo.Checking Crystal Report 8.5
SET ProgFilesRoot=%ProgramFiles%
IF EXIST "%ProgFilesRoot%\Seagate Software\Crystal Reports\crw32.exe" goto regdll32

echo.
echo.
echo.Installing Crystal Report 8.5
(
b:\CrystalReport8_5\scrdev.msi /qr PIDKEY=A6A50-8900008-ZE1007S INSTALLLEVEL=3
net use b: /delete /yes
)>NUL

:regdll32
(
regsvr32.exe %windir%\system32\crviewer.dll /s
regsvr32.exe %windir%\system32\xqviewer.dll /s
regsvr32.exe %windir%\system32\Crystl32.OCX /s
regsvr32.exe %windir%\system32\sviewhlp.dll /s
regsvr32.exe %windir%\system32\swebrs.dll /s
regsvr32.exe %windir%\system32\craxdrt.dll /s
regsvr32.exe %windir%\system32\craxddrt.dll /s
regsvr32.exe %windir%\system32\p2sodbc.dll /s
regsvr32.exe %windir%\system32\pdsodbc.dll /s
regsvr32.exe %windir%\system32\Comdlg32.ocx /s
regsvr32.exe %windir%\system32\mscomctl.ocx /s
regsvr32.exe %windir%\system32\MSCOMM32.OCX /s
regsvr32.exe %windir%\system32\MSFLXGRD.ocx /s
regsvr32.exe %windir%\system32\MSMASK32.ocx /s
regsvr32.exe %windir%\system32\RICHTX32.ocx /s
regsvr32.exe %windir%\system32\prjXTab.ocx /s
regsvr32.exe %windir%\system32\lvButton.ocx /s
regsvr32.exe %windir%\system32\HoverButton.ocx /s
regsvr32.exe %windir%\system32\MSWINSCK.OCX /s
)>NUL
goto CreateShortcut

:64_bit

echo. 
echo.
echo Installing/Registering DLL . . . 

(copy \\192.168.1.2\Omega\Dll\*.* %windir%\SysWOW64\
if errorlevel 1 goto mapdrive64
if NOT exist "C:\Omega" (md C:\Omega)
if NOT exist "C:\Omega\Reports" (md C:\Omega\Reports)
copy \\192.168.1.2\Omega\Omega.exe C:\Omega\
copy \\192.168.1.2\Omega\Reports\*.* C:\Omega\Reports\
copy \\192.168.1.2\Omega\Config C:\Omega\
copy \\192.168.1.2\Omega\Update.bat C:\Omega\
copy \\192.168.1.2\Omega\Setup.bat C:\Omega\
)>NUL

echo.
echo.
echo.Checking Crystal Report 8.5
SET ProgFiles86Root=%ProgramFiles(x86)%
IF EXIST "%ProgFiles86Root%\Seagate Software\Crystal Reports\crw32.exe" goto regdll64

echo.
echo.
echo.Installing Crystal Report 8.5
(
\\192.168.1.2\Omega\CrystalReport8_5\scrdev.msi /qr PIDKEY=A6A50-8900008-ZE1007S INSTALLLEVEL=3
)>NUL

goto regdll64

:mapdrive64
(
net use b: \\192.168.1.2\Omega 123456 /user:update
copy b:\Dll\*.* %windir%\SysWOW64\
if NOT exist "C:\Omega" (md C:\Omega)
if NOT exist "C:\Omega\Reports" (md C:\Omega\Reports)
copy b:\Omega.exe C:\Omega\
copy b:\Reports\*.* C:\Omega\Reports\
copy b:\Config C:\Omega\
copy b:\Update.bat C:\Omega\
copy b:\Setup.bat C:\Omega\
)>NUL

echo.
echo.
echo.Checking Crystal Report 8.5
SET ProgFiles86Root=%ProgramFiles(x86)%
IF EXIST "%ProgFiles86Root%\Seagate Software\Crystal Reports\crw32.exe" goto regdll64

echo.
echo.
echo.Installing Crystal Report 8.5
(
b:\CrystalReport8_5\scrdev.msi /qr PIDKEY=A6A50-8900008-ZE1007S INSTALLLEVEL=3
net use b: /delete /yes
)>NUL


:regdll64
(
regsvr32.exe %windir%\SysWOW64\crviewer.dll /s
regsvr32.exe %windir%\SysWOW64\xqviewer.dll /s
regsvr32.exe %windir%\SysWOW64\Crystl32.OCX /s
regsvr32.exe %windir%\SysWOW64\sviewhlp.dll /s
regsvr32.exe %windir%\SysWOW64\swebrs.dll /s
regsvr32.exe %windir%\SysWOW64\craxdrt.dll /s
regsvr32.exe %windir%\SysWOW64\craxddrt.dll /s
regsvr32.exe %windir%\SysWOW64\p2sodbc.dll /s
regsvr32.exe %windir%\SysWOW64\pdsodbc.dll /s
regsvr32.exe %windir%\SysWOW64\Comdlg32.ocx /s
regsvr32.exe %windir%\SysWOW64\mscomctl.ocx /s
regsvr32.exe %windir%\SysWOW64\MSCOMM32.OCX /s
regsvr32.exe %windir%\SysWOW64\MSFLXGRD.ocx /s
regsvr32.exe %windir%\SysWOW64\MSMASK32.ocx /s
regsvr32.exe %windir%\SysWOW64\RICHTX32.ocx /s
regsvr32.exe %windir%\SysWOW64\prjXTab.ocx /s
regsvr32.exe %windir%\SysWOW64\lvButton.ocx /s
regsvr32.exe %windir%\SysWOW64\HoverButton.ocx /s
regsvr32.exe %windir%\SysWOW64\MSWINSCK.OCX /s
)>NUL

goto CreateShortcut

:CreateShortcut
@Echo off
echo.
echo.
echo.Create Desktop Shortcut
rem Window Style
REM 1 = Normal, 3 Maximized, 7 = Minimized

rem Choose "Desktop" or "AllUsersDesktop"
set Location="AllUsersDesktop"

set DisplayName="O m e g a"
set filename="C:\Omega\Omega.exe"

REM Set icon to an icon from an exe or "something.ico"
set icon="C:\Omega\Omega.exe, 0"

set WorkingDir="C:\Omega"

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
set filename="C:\Omega\Update.bat"

REM Set icon to an icon from an exe or "something.ico"
set icon="C:\Omega\Update.bat, 0"

set WorkingDir="C:\Omega"

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
