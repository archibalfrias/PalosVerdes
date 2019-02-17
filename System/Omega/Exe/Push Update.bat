@Echo Off
echo.
echo.
echo Pushing Updated Files . . . 

(copy omega.exe \\192.168.1.2\Omega\
if errorlevel 1 goto mapdrive
cd ..
copy Reports\*.* \\192.168.1.2\Omega\Reports\)>NUL
goto pause

:mapdrive
(net use b: \\192.168.1.2\Omega albert /user:administrator
copy omega.exe b:\
cd ..
copy Reports\*.* b:\Reports\
net use b: /delete /yes)>NUL

:pause

echo.
echo.
echo Successfully Push Updated Files . . .
echo.
echo.
pause