@echo off

set TimeUsed=""
set scriptdir=%~dp0
set /p Timer="Choose time conversion (0 = seconds, 1 = minute, 2 = hour): "

if %Timer%==0 (set TimeUsed=seconds)
if %Timer%==1 (set TimeUsed=minutes)
if %Timer%==2 (set TimeUsed=hour)
if not %Timer%==0 if not %Timer%==1 if not %Timer%==2 exit

set /p UserInput="Set the new timer by %TimeUsed% (this will reset the timer): "
set /a newTimer=0
if %Timer%==0 (set /a newTimer=%UserInput%*1000)
if %Timer%==1 (set /a newTimer=%UserInput%*60000)
if %Timer%==2 (set /a newTimer=%UserInput%*3600000)

set /p previousTimer=<config.txt
powershell -Command "(Get-Content config.txt) | ForEach-Object { $_ -replace "%previousTimer%", "%newTimer%" } | Set-Content config.txt" 
TASKKILL /F /IM wscript.exe
start %scriptdir%shutdown-script.vbs
