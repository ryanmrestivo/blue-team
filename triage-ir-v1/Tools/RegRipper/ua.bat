@echo off
REM Batch file to automate running of userassist.pl plugin file
REM in order to enumerate the contents of the UserAssist key from
REM an NTUSER.DAT Registry hive file.
REM 
REM Usage: ua <path_to_NTUSER.DAT_hive_file>
REM
REM copyright 2008 H. Carvey, keydet89@yahoo.com

REM rip.exe -r %1 -p userassist
rip.exe -r %1 -p userassist2