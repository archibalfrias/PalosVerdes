
@echo off
echo.
echo.
echo Updating . . . 

(copy \\192.168.1.2\omega\omega.exe C:\omega
if errorlevel 1 goto mapdrive
copy \\192.168.1.2\omega\update.bat C:\omega
copy \\192.168.1.2\omega\Setup.bat C:\omega
copy \\192.168.1.2\omega\Reports\*.* C:\omega\Reports\)>NUL
goto pause

:mapdrive
(net use b: \\192.168.1.2\omega 123456 /user:update
copy b:\Omega.exe C:\Omega
copy b:\Update.bat C:\Omega
copy b:\Setup.bat C:\Omega
copy b:\Reports\*.* C:\omega\Reports\
net use b: /delete /yes)>NUL

:pause
echo.
echo.
echo Successfully update . . .
echo.
echo.
pause