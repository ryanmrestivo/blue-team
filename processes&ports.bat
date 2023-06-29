@echo off
goto check_Permissions

:check_Permissions
    echo Administrative permissions required. Detecting permissions...

    net session >nul 2>&1
    if %errorLevel% == 0 (
        echo Success: Administrative permissions confirmed.
    ) else (
        echo Failure: Current permissions inadequate.
        echo Please run this script as an administrator.
        pause >nul
        exit /b
    )

:: Redirect output to "output.txt" in the script's directory
> "%~dp0output.txt" (

    ::tasklist with command line
    echo tasklist /v
    tasklist /v
    echo.

    ::netstat
    echo netstat -o
    netstat -o
    echo.
)

echo Results have been saved to "%~dp0output.txt".
cmd /k
