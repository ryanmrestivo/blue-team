# wifi.bat
Displays cleartext wifi information.  Doesn't do anything with it.

# password.bat
Generates passwords & "random" strings.

# WindowsEnum

# PSRecon
See readme in directory.

# scri%00pt
see readme in directory.

# windows-enumeration.bat
No PowerShell?  No Problem?

## Usage

Run as administrator if you can, otherwise open where you want the output.txt file.

# windows-enumeration.ps1
A Powershell Privilege Escalation Enumeration Script.

## Usage

To run the quick standard checks.

```powershell
.\WindowsEnum.ps1
```

Directly from CMD (you may need to change directory)

```
powershell -nologo -executionpolicy bypass -file WindowsEnum.ps1
```

Extended checks will search for config files, various interesting files, and passwords in files and the registry, etc. It will take some time so be patient.

```powershell
.\WindowsEnum.ps1 extended
```

```
powershell -nologo -executionpolicy bypass -file WindowsEnum.ps1 extended
```
