;InnoSetupVersion=4.2.6

[Setup]
AppName=SysAnalyzer
AppVerName=SysAnalyzer 2.x
DefaultDirName=c:\iDEFENSE\SysAnalyzer\
DefaultGroupName=SysAnalyzer
OutputBaseFilename=SysAnalyzer_Setup
OutputDir=./

[Files]
Source: ./dependancy\spSubclass2.dll; DestDir: {sys}; Flags: uninsneveruninstall regserver
Source: ./dependancy\TABCTL32.OCX; DestDir: {sys}; Flags: uninsneveruninstall regserver
Source: ./dependancy\vbDevKit.dll; DestDir: {sys}; Flags: uninsneveruninstall regserver
Source: ./dependancy\mscomctl.ocx; DestDir: {sys}; Flags: uninsneveruninstall regserver
Source: ./dependancy\MSWINSCK.OCX; DestDir: {sys}; Flags: uninsneveruninstall regserver
Source: ./dependancy\hexed.ocx; DestDir: {sys}; Flags: uninsneveruninstall regserver
Source: ./dependancy\richtx32.ocx; DestDir: {sys}; Flags: uninsneveruninstall regserver
Source: ./dependancy\libVT.dll; DestDir: {sys}; Flags: uninsneveruninstall regserver
Source: ./dependancy\zlib.dll; DestDir: {app}
Source: ./dependancy\MSINET.OCX; DestDir: {sys}; Flags: uninsneveruninstall regserver promptifolder

;
Source: strDump.exe; DestDir: {app}; Flags: ignoreversion
Source: virustotal.exe; DestDir: {app}; Flags: ignoreversion
Source: api_log.dll; DestDir: {app}; Flags: ignoreversion
Source: enumMutex.dll; DestDir: {app}; Flags: ignoreversion
;Source: api_log.x64.dll; DestDir: {app}; Flags: ignoreversion
Source: api_logger.exe; DestDir: {app}; Flags: ignoreversion
Source: dir_watch.dll; DestDir: {app}; Flags: ignoreversion
Source: exploit_sigs.txt; DestDir: {app}
Source: proc_watch.exe; DestDir: {app}; Flags: ignoreversion
;Source: safe_test1.exe; DestDir: {app}
Source: sniff_hit.exe; DestDir: {app}; Flags: ignoreversion
Source: sysAnalyzer.exe; DestDir: {app}; Flags: ignoreversion
Source: SysAnalyzer_help.chm; DestDir: {app}
;Source: SysAnalyzer.pdb; DestDir: {app}
Source: known_files.mdb; DestDir: {app}; Flags: uninsneveruninstall onlyifdoesntexist
Source: dirwatch_ui.exe; DestDir: {app}; Flags: ignoreversion
Source: shellext.external.txt; DestDir: {app}; Flags: ignoreversion
Source: win_dump.exe; DestDir: {app}
Source: WinPcap_4_1_2.exe; DestDir: {app}
Source: loadlib.exe; DestDir: {app}
Source: ShellExt.exe; DestDir: {app}
Source: x64Helper.exe; DestDir: {app}
Source: goat.html; DestDir: {app}

[Dirs]

[Icons]
Name: {group}\SysAnalyzer; Filename: {app}\sysAnalyzer.exe
Name: {group}\ApiLogger; Filename: {app}\api_logger.exe
Name: {group}\Sniff_Hit; Filename: {app}\sniff_hit.exe
Name: {group}\Uninstall; Filename: {app}\unins000.exe; WorkingDir: {app}
Name: {userdesktop}\SysAnalyzer; Filename: {app}\sysAnalyzer.exe; WorkingDir: {app}; IconIndex: 0
;Name: {group}\Test Binary; Filename: {app}\sysAnalyzer.exe; Parameters: safe_test1.exe; WorkingDir: {app}; IconFilename: {app}\safe_test1.exe
Name: {group}\Help File; Filename: {app}\SysAnalyzer_help.chm; WorkingDir: {app}
Name: {userdesktop}\DirWatch; Filename: {app}\dirwatch_ui.exe; IconIndex: 0
Name: {userdesktop}\Sniffhit; Filename: {app}\sniff_hit.exe; IconIndex: 0
Name: {userdesktop}\ApiLogger; Filename: {app}\api_logger.exe; IconIndex: 0
Name: {group}\ProcWatch; Filename: {app}\proc_watch.exe
Name: {userdesktop}\ProcWatch; Filename: {app}\proc_watch.exe

[CustomMessages]
NameAndVersion=%1 version %2
AdditionalIcons=Additional icons:
CreateDesktopIcon=Create a &desktop icon
CreateQuickLaunchIcon=Create a &Quick Launch icon
ProgramOnTheWeb=%1 on the Web
UninstallProgram=Uninstall %1
LaunchProgram=Launch %1
AssocFileExtension=&Associate %1 with the %2 file extension
AssocingFileExtension=Associating %1 with the %2 file extension...

[Run]
Filename: {app}\WinPcap_4_1_2.exe; StatusMsg: Installing WinPcap Packet Sniffer Driver; Flags: postinstall runascurrentuser unchecked
