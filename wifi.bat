@echo off
mode 60,30
setlocal enabledelayedexpansion

:: This script retrieves and displays the passwords of saved Wi-Fi networks on a Windows machine.
:: It asks the user if they want to save the results to a file named 'wifi.txt'.

:: Store the output in a variable
set "output="
set "output=!output!WiFi Passwords:`n"

:: Get all the profiles
FOR /F "usebackq tokens=2 delims=:" %%a in (
    `netsh wlan show profiles ^| findstr /C:"All User Profile"`) DO (
    
    :: Trim the profile name
    set "profile=%%a"
    set "profile=!profile:~1!"
    
    :: Get the WiFi key of the profile
    FOR /F "usebackq tokens=2 delims=:" %%b in (
        `netsh wlan show profile name^="!profile!" key^=clear ^| findstr /C:"Key Content"`) DO (
        
        :: Trim the key
        set "key=%%b"
        set "key=!key:~1!"
        
        :: Display the password in the terminal
        echo WiFi: [!profile!] Password: [!key!]
        
        :: Append the password to the output variable
        set "output=!output!WiFi: [!profile!] Password: [!key!]`n"
    )
)

:: Ask the user if they want to save the output to a text file
echo.
set /p "saveToFile=Do you want to save the output to a text file (Y/N)? "
if /i "%saveToFile%" equ "Y" (
    echo !output! | findstr /r /c:".*" > "wifi.txt"
    echo Results exported to wifi.txt
)
