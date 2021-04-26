# triage-ir-v1
Triage Incident Response.  Must have administrative privileges on host workstation.  If this is not feasible, check out the repository windows-enumeration

##### Incident Response Triage:
Scripted collection of system information valuable to a Forensic Analyst. 
IRTriage will automatically "Run As ADMINISTRATOR" in all Windows versions except WinXP.

The original source was [Triage-ir v0.851](https://code.google.com/p/triage-ir/) an Autoit script written by Michael Ahrendt.
Unfortunately Michael's last changes were [posted](http://mikeahrendt.blogspot.ca/2012/01/automated-triage-utility.html) on 9th November 2012

I let Michael [know](http://mikeahrendt.blogspot.com/2012/01/automated-triage-utility.html?showComment=1455628200788#c6111030418808145121) that I have forked his project:
I am pleased to anounce that he gave me his blessing to fork his source code, long live Open Source!)

###### What if having a full disk image is not an option during an incident?
Imagine that you are investigating a dozen or more possibly infected or compromised systems.
Can you spend 2-8 hours making a forensic copy of the hard drives on those computers?
In such situation fast forensics\"Triage" is the solution for such a situation.
Instead of copying everything, collecting some key files can solve this issue.

IRTriage will collect:
- system information
- network information
- registry hives
- disk information, and
- dump memory.

One of the powerful capabilities of IRTriage is collecting information from "Volume Shadow Copy" which can defeat many anti-forensics techniques.

The IRTriage is itself just an autoit script that depend on other tools such as:
- Win32|64dd (free from Moonsols) or FDpro *(HBGary's commercial product)
- Sysinternals Suite
- The Sleuth Kit
- Regripper
- NirSoft => MFTDump and WinPrefetchView
- md5deep and sha1deep
- CSVFileView
- 7zip
- and some windows built-in commands.

In case of an incident, you want to make minimal changes to the "evidence machine", therefore I would suggest you copy IRTriage to a USB drive, the only issue here is if you are planning to dump the memory, the USB drive must be larger than the physical ram installed in the computer.

Once you launch the GUI application you can select what information you would like to collect.
Each category is in a separate tab.
All the collected information will be dumped into a new folder labled with [hostname-date-time].
