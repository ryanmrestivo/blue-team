This is the readme file for rescan

This directory contains RegScan, which is a command line tool meant to be 
run on a live system.  This tool is described in chapter 2 of the book, 
and can be copied from the DVD and run from another location. 

Running regscan from a command prompt will extract information from the
system you're currently logged into; to save the information, use a 
command similar to the following:

C:\tools>regscan > regscan.txt

..or..

C:\tools>regscan > %COMPUTERNAME%_regscan.txt

As with the other tools, be sure that you keep the p2x588.dll file in the
same directory as the EXE file.


This Perl code is provided AS-IS, with no warantees or guarantees as to its
functionality.  Please feel free to view, use, and modify the code as you like,
but please provide proper credit and attribution.