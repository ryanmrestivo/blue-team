@echo off
title C2 Hunter Lite
echo C2 Hunter Lite
echo.


goto check_Permissions

:check_Permissions
    echo Administrative permissions required. Detecting permissions...

    net session >nul 2>&1
    if %errorLevel% == 0 (
        echo Success: Administrative permissions confirmed.
	echo Hit "enter" to continue.  Output will be output to the location of this file as "output.txt."
	echo note: output.txt is not currently activated.
    ) else (
        echo Failure: Current permissions inadequate.
    )

    pause >nul

::tasklist
echo tasklist
tasklist
echo.

::netstat
echo netstat -o
netstat -o
echo.

::todo
::combine port firewall rules with taskkill / task-tree-kills 
::todo
::output to output.txt
::notes
::https://docs.microsoft.com/en-us/troubleshoot/windows-server/networking/netsh-advfirewall-firewall-control-firewall-behavior

:: echo netsh advfirewall firewall delete rule ?
:: echo taskkill /PID #### /F

cmd /k
