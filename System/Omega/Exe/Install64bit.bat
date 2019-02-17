@echo Installing ....
@echo off
copy \\192.168.1.2\Omega\Dll\comdlg32.ocx c:\windows\SysWOW64\
if errorlevel 1 goto mapdrive
copy \\192.168.1.2\Omega\Dll\HoverButton.oca c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\HoverButton.ocx c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\lvButton.oca c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\lvButton.ocx c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\mscomctl.ocx c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\mscomm32.ocx c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\msflxgrd.ocx c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\msmask32.ocx c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\prjXTab.oca c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\prjXTab.ocx c:\windows\SysWOW64\
copy \\192.168.1.2\Omega\Dll\richtx32.ocx c:\windows\SysWOW64\

md D:\Omega
md D:\Omega\Images

copy \\192.168.1.2\Omega\Images\*.* D:\Omega\Images\
copy \\192.168.1.2\Omega\Omega.exe D:\Omega\
copy \\192.168.1.2\Omega\Update.bat D:\Omega\
copy \\192.168.1.2\Omega\Install64bit.bat D:\Omega\

goto RegOCX

:mapdrive
net use b: \\192.168.1.2\omega 123456 /user:update

copy b:\Dll\comdlg32.ocx c:\windows\SysWOW64\
copy b:\Dll\HoverButton.oca c:\windows\SysWOW64\
copy b:\Dll\HoverButton.ocx c:\windows\SysWOW64\
copy b:\Dll\lvButton.oca c:\windows\SysWOW64\
copy b:\Dll\lvButton.ocx c:\windows\SysWOW64\
copy b:\Dll\mscomctl.ocx c:\windows\SysWOW64\
copy b:\Dll\mscomm32.ocx c:\windows\SysWOW64\
copy b:\Dll\msflxgrd.ocx c:\windows\SysWOW64\
copy b:\Dll\msmask32.ocx c:\windows\SysWOW64\
copy b:\Dll\prjXTab.oca c:\windows\SysWOW64\
copy b:\Dll\prjXTab.ocx c:\windows\SysWOW64\
copy b:\Dll\richtx32.ocx c:\windows\SysWOW64\

md D:\Omega
md D:\Omega\Images

copy b:\Images\*.* D:\Omega\Images\
copy b:\Omega.exe D:\Omega\
copy b:\Update.bat D:\Omega\
copy b:\Install64bit.bat D:\Omega\

net use b: /delete /yes

:RegOCX
regsvr32.exe c:\windows\SysWOW64\comdlg32.ocx /s
regsvr32.exe c:\windows\SysWOW64\mscomctl.ocx /s
regsvr32.exe c:\windows\SysWOW64\mscomm32.ocx /s
regsvr32.exe c:\windows\SysWOW64\msflxgrd.ocx /s
regsvr32.exe c:\windows\SysWOW64\msmask32.ocx /s
regsvr32.exe c:\windows\SysWOW64\richtx32.ocx /s
regsvr32.exe c:\windows\SysWOW64\prjXTab.ocx /s
regsvr32.exe c:\windows\SysWOW64\lvButton.ocx /s
regsvr32.exe c:\windows\SysWOW64\HoverButton.ocx /s


:pause
pause