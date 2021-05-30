@echo off

:: Number of scanning threads
set NUMTH=8

set ipfile=t%random%m%random%p
set scanfx=s%random%c%random%n
set /A thinc=256/%NUMTH% >nul

ping -4 -n 1 -w 1 %computername% |find "statistics" >%ipfile%
for /F "tokens=4,5,6,7 delims=:. " %%i in (%ipfile%) do (
  if not "%%i"=="" (
    echo Local IP address: %%i.%%j.%%k.%%l
    echo Scanning range  : %%i.%%j.%%k.0/24

    echo for /L %%%%a in ^(1,1,32^) do @ping -n 1 -w 1  %%i.%%j.%%k.%%%%a ^|find "time=" ^>^>%%0.log >%scanfx%1.bat
    echo for /L %%%%a in ^(33,1,64^) do @ping -n 1 -w 1  %%i.%%j.%%k.%%%%a ^|find "time=" ^>^>%%0.log >%scanfx%2.bat
    echo for /L %%%%a in ^(65,1,96^) do @ping -n 1 -w 1  %%i.%%j.%%k.%%%%a ^|find "time=" ^>^>%%0.log >%scanfx%3.bat
    echo for /L %%%%a in ^(97,1,128^) do @ping -n 1 -w 1  %%i.%%j.%%k.%%%%a ^|find "time=" ^>^>%%0.log >%scanfx%4.bat
    echo for /L %%%%a in ^(129,1,160^) do @ping -n 1 -w 1  %%i.%%j.%%k.%%%%a ^|find "time=" ^>^>%%0.log >%scanfx%5.bat
    echo for /L %%%%a in ^(161,1,192^) do @ping -n 1 -w 1  %%i.%%j.%%k.%%%%a ^|find "time=" ^>^>%%0.log >%scanfx%6.bat
    echo for /L %%%%a in ^(193,1,224^) do @ping -n 1 -w 1  %%i.%%j.%%k.%%%%a ^|find "time=" ^>^>%%0.log >%scanfx%7.bat
    echo for /L %%%%a in ^(225,1,254^) do @ping -n 1 -w 1  %%i.%%j.%%k.%%%%a ^|find "time=" ^>^>%%0.log >%scanfx%8.bat


  )
)

for /L %%i in (1,1,%NUMTH%) do echo @echo 1 ^>%%0.txt ^&exit >>%scanfx%%%i.bat

for /L %%i in (1,1,%NUMTH%) do start /MIN %scanfx%%%i.bat
echo Scanning...

:: Wait for all threads to finish
:waitthread
ping -n 2 127.0.0.1 >nul
for /L %%i in (1,1,%NUMTH%) do if not exist %scanfx%%%i.bat.txt goto waitthread

:: Copy the scan logs to a single file
copy %scanfx%*.bat.log %scanfx%.scan.log >nul

:: Clean up, delete temp files
del /F /Q %ipfile%
for /L %%i in (1,1,%NUMTH%) do @del /F /Q %scanfx%%%i.bat
for /L %%i in (1,1,%NUMTH%) do @del /F /Q %scanfx%%%i.bat.txt
for /L %%i in (1,1,%NUMTH%) do @del /F /Q %scanfx%%%i.bat.log

start "" %windir%\notepad.exe %scanfx%.scan.log
::type %scanfx%.scan.log |more

:: Wait for notepad to open before killing the file
:waitnotepad
ping -n 2 127.0.0.1 >nul
tasklist /FI "WINDOWTITLE eq %scanfx%.scan.log*" |find "notepad.exe" >nul
if not "%errorlevel%"=="0" goto waitnotepad

del /F /Q %scanfx%.scan.log