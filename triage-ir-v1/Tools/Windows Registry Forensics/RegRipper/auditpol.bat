@echo off
REM Batch file to automate running of auditpol.pl plugin file
REM in order to enumerate the audit policy from the Security hive
REM file.
REM 
REM Usage: auditpol <path_to_Security_hive_file>
REM
REM copyright 2008 H. Carvey, keydet89@yahoo.com

rip.exe -r %1 -p auditpol