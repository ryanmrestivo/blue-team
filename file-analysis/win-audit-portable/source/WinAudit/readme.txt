BUILDING WINAUDIT
=================

You do not need to build WinAudit, an executable is available for download at 
www.parmavex.co.uk/winaudit.html . 

Prequisites:

1. Visual Studio 2010 or newer

2. Microsoft SDK 7.1 or newer 

3. Microsoft DirectX SDK (June 2010) or newer

WinAudit is free software built with free software. If for some reason you wish to
build WinAudit yourself, a Visual Studio C++ 2010 Express solution is provided.
There are two projects, a static library called PxsBase which is a collection of classes 
that provide basic functionality. The second is the WinAudit exectuable project which
depends on PxsBase. No third party libraries are required. The resultant 32-bit executable
will run on a fresh install of XP or newer.

1. Start Visual Studio

2. Verify the VC++ directories for both PxsBase and WinAudit point to where
the Microsoft SDK is installed on your system. Additionally, the VC++
include directory for WinAudit must point to the DirectX SDK include directory.
Typically this is "$(DXSDK_DIR)include" without the quotes.  

3. If the WinAudit project is not in bold then "Set as StartUp Project".

4. Verify in "Project Dependencies" that WinAudit depends on PxsBase.

5. Select a "Release" build.

6. Finally, "Build Solution"

The solution should build with no warnings or errors. The output is at 
WinAudit\Win32\Release\WinAudit.exe. 

Compilation is set to the level /W4 with additional warnings enabled as listed 
in pxsbase.h. Compilers are always improving, if you see any compilation 
messages please inform us. At present you cannot build WinAudit with Mingw32/64 
as some operating system functionality used by WinAudit, such as Task Manager 2, 
has not yet been included with the Mingw project. Before distributing the 
executable to others it is recommended that:

1. If you have a professonal version of Visual Studio, use the /analyze switch.

2. Use Microsoft's Application Verifier to test its runtime behaviour.

3. Test the exectutable on all versions of Windows on which you intend it to be
run. Use the logger to verify there are no unexpected errors.

4. Make a note of the exectuable's MD5 so you can verify it has not been altered.


Parmavex
Birmingham
England

 
  