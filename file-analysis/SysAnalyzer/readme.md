
![screenshot](https://raw.githubusercontent.com/dzzie/SysAnalyzer/master/chm_src/img/main_ui_screens.gif)

<pre>

Copyright (C) 2005 iDefense, a Verisign Company 
Author: David Zimmer  dzzie@yahoo.com 

Videos:
-------------------------------
updates:          https://www.youtube.com/watch?v=4twR8xtVWPk
apiLogger:        https://www.youtube.com/watch?v=SqdGjihhDoU
original trainer: https://www.youtube.com/watch?v=OPXwKChdO4c
-------------------------------

Installer:  http://sandsprite.com/tools.php?id=13
Help File:  http://sandsprite.com/iDef/SysAnalyzer/

SysAnalyzer is an application that was designed to give malcode analysts an 
automated tool to quickly collect, compare, and report on the actions a 
binary took while running on the system. 

The main components of SysAnalyzer work off of comparing snapshots of the 
system over a user specified time interval. The reason a snapshot mechanism 
was used compared to a live logging implementation is to reduce the amount 
of data that analysts must wade through when conducting their analysis. By 
using a snapshot system, we can effectively present viewers with only the 
persistent changes found on the system since the application was first run. 

While this mechanism does help to eliminate allot of the possible noise 
caused by other applications, or inconsequential runtime nuances, it also 
opens up the possibility for missing key data. Because of this SysAnalyzer 
also gives the analyst the option to include several forms of live logging 
into the analysis procedure. 

Note: SysAnalyzer is not a sandboxing utility. Target executables are run 
in a fully live test on the system. If you are testing malicious code, you 
must realize you will be infecting your test system. 

SysAnalyzer's is designed to take snapshots of the following system 
attributes: 

* Running processes 
* Open ports and associated process 
* Dlls loaded into explorer.exe and Internet Explorer 
* System Drivers loaded into the kernel 
* Snapshots of certain registry keys 

For more information see the chm help file or videos.

</pre>
