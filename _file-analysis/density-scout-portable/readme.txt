Quickstart: Run command prompt // Drag win 32 or 64 into cmd to start program

Syntax: densityscout [options] file_or_directory

options: -a .............. Show errors and empties, too
         -d .............. Just output data (Format: density|path)
         -l density ...... Just files with density lower than the given value
         -g density ...... Just files with density greater than the given value
         -n number ....... Maximum number of lines to print
         -m mode ......... Mode ABS (default) or CHI (for filesize > 100 Kb)
         -o file ......... File to write output to
         -p density ...... Immediately print if lower than the given density
         -P density ...... Immediately print if greater than the given density
         -r .............. Walk recursively
         -s suffix(es) ... Filetype(s) (i.e.: dll or dll,exe,...)
         -S suffix(es) ... Filetype(s) to ignore (i.e.: dll or dll,exe)
         -pe ............. Include all portable executables by magic number
         -PE ............. Ignore all portable executables by magic number

Note:    Packed and/or encrypted data usually has a much higher density than
         normal data (like text or executable binaries).

Modes:   ABS ... Computes the average distance from the ideal quantity for each
                 byte-state according to the overall byte-quantity of the
                 evaluated file.
                 Typical ABS-density for a packed file: < 0.1
                 Typical ABS-density for a normal file: > 0.9

         CHI ... Just the same as ABS but actually squaring each distance.
                 Typical CHI-density for a packed file: < 100.0
                 Typical CHI-density for a normal file: > 1000.0


Here is one of the fastest ways to get a quick glance of if there's anything "suspicious" of a specific Microsoft Windows installation:

densityscout -s cpl,exe,dll,ocx,sys,scr -p 0.1 -o results.txt c:\Windows\System32


On a healthy Windows 7 Professional installation during the run-time of DensityScout you should see something similar to the following:

Calculating density for file ...
(0.03763) | c:\Windows\System32\bootres.dll
(0.05963) | c:\Windows\System32\VAIO S Series - Summer 2011.scr
(0.05214) | c:\Windows\System32\WdfCoinstaller01009.dll
This promptly reveals that Sony has put some strange screensaver on my notebook.

The first 20 lines of the final result list should look like this:

(0.03763) | c:\Windows\System32\bootres.dll
(0.05214) | c:\Windows\System32\WdfCoinstaller01009.dll
(0.05963) | c:\Windows\System32\VAIO S Series - Summer 2011.scr
(0.11521) | c:\Windows\System32\LkmdfCoInst.dll
(0.12726) | c:\Windows\System32\mcupdate_GenuineIntel.dll
(0.20664) | c:\Windows\System32\iglhsip64.dll
(0.27113) | c:\Windows\System32\pegibbfc.rs
(0.27516) | c:\Windows\System32\usk.rs
(0.27633) | c:\Windows\System32\cero.rs
(0.28895) | c:\Windows\System32\pegi.rs
(0.30524) | c:\Windows\System32\AuthFWGP.dll
(0.30681) | c:\Windows\System32\iscsicpl.exe
(0.32147) | c:\Windows\System32\msshavmsg.dll
(0.32388) | c:\Windows\System32\SrpUxNativeSnapIn.dll
(0.32859) | c:\Windows\System32\qedwipes.dll
(0.34056) | c:\Windows\System32\imagesp1.dll
(0.34697) | c:\Windows\System32\oflc.rs
(0.36592) | c:\Windows\System32\auditpolmsg.dll
(0.36870) | c:\Windows\System32\onexui.dll
(0.38369) | c:\Windows\System32\resmon.exe
As you can see you won't find a lot packed less than 0.1 portable executables on a healthy Microsoft Windows installation.