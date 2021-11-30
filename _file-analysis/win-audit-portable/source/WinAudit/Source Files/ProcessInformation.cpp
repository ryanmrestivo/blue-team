///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Process Information Class Implementation
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Copyright 1987-2017 PARMAVEX SERVICES
//
// Licensed under the European Union Public Licence (EUPL), Version 1.1 or -
// as soon they will be approved by the European Commission - subsequent
// versions of the EUPL (the "Licence"). You may not use this work except in
// compliance with the Licence. You may obtain a copy of the Licence at:
//
// http://ec.europa.eu/idabc/eupl
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the Licence is distributed on an "AS IS" basis,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the Licence for the specific language governing permissions and
// limitations under the Licence. This source code is free software. It
// must not be sold, leased, rented, sub-licensed or used for any form of
// monetary recompense whatsoever. This notice must not be removed or altered
// from this source distribution.
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/ProcessInformation.h"

// 2. C System Files
#include <Psapi.h>
#include <TlHelp32.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AutoCloseHandle.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemException.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ProcessInformation::ProcessInformation()
                   :m_Processes()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
ProcessInformation::~ProcessInformation()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Enable/disable a token privilege
//
//  Parameters:
//      pszPrivilege - name of privilege to enable/disable
//      enable       - true to enable, false to disable
//
//  Returns:
//       void
//===============================================================================================//
void ProcessInformation::ChangePrivilege( LPCWSTR pszPrivilege, bool enable )
{
    LUID   luid;
    DWORD  lastError = 0;
    HANDLE hToken    = nullptr;
    TOKEN_PRIVILEGES tp;

    if ( ( pszPrivilege == nullptr ) || ( *pszPrivilege == PXS_CHAR_NULL ) )
    {
        throw ParameterException( L"pszPrivilege", __FUNCTION__ );
    }

    // Open this thread's token
    if ( OpenThreadToken( GetCurrentThread(),
                          TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, FALSE, &hToken ) == 0 )
    {
        // If error is no token, impersonate self and try again
        lastError = GetLastError();
        if ( lastError != ERROR_NO_TOKEN )
        {
            throw SystemException( lastError, L"OpenThreadToken", __FUNCTION__ );
        }

        if ( ImpersonateSelf( SecurityImpersonation ) == 0 )
        {
            throw SystemException( GetLastError(), L"ImpersonateSelf", __FUNCTION__ );
        }

        if ( OpenThreadToken( GetCurrentThread(),
                              TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, FALSE, &hToken ) == 0 )
        {
            throw SystemException( GetLastError(), L"OpenThreadToken", __FUNCTION__ );
        }
    }
    AutoCloseHandle CloseToken( hToken );

    // Get the locally unique id for the privilege
    memset( &luid, 0, sizeof ( luid ) );
    if ( LookupPrivilegeValue( nullptr, pszPrivilege, &luid ) == 0 )
    {
        throw SystemException( GetLastError(), L"LookupPrivilegeValue", __FUNCTION__ );
    }

    // Set the token privileges
    tp.PrivilegeCount = 1;
    tp.Privileges[ 0 ].Luid = luid;     // 64-bit value
    if ( enable )
    {
        tp.Privileges[ 0 ].Attributes = SE_PRIVILEGE_ENABLED;
    }
    else
    {
        tp.Privileges[ 0 ].Attributes = 0;
    }

    if ( AdjustTokenPrivileges( hToken, FALSE, &tp, sizeof ( tp ), nullptr, nullptr ) == 0 )
    {
        throw SystemException( GetLastError(), L"AdjustTokenPrivileges", __FUNCTION__ );
    }

    // Besides the return result also need to check the error code as
    // AdjustTokenPrivileges can say success but not all privileges have
    // were assigned, e.g. 1300 == not all privileges assigned
    lastError = GetLastError();
    if ( lastError != ERROR_SUCCESS )
    {
        throw SystemException( lastError, L"AdjustTokenPrivileges", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Fill a class scope array with process information
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void ProcessInformation::FillProcessesInfo()
{
    bool      privilegeChanged = false;
    HANDLE    hSnapProcess, hSnapModule;
    String    Message, Insert1;
    Formatter Format;
    MODULEENTRY32     me32;
    PROCESSENTRY32    pe32;
    TYPE_PROCESS_INFO ProcessInfo;

    m_Processes.RemoveAll();

    // PSAPI usually requires SeDebugPrivilege
    try
    {
        if ( IsPrivilegeEnabled( SE_DEBUG_NAME ) == false )
        {
            ChangePrivilege( SE_DEBUG_NAME, true );
            privilegeChanged = true;
        }
    }
    catch ( const Exception& e )
    {
        // Log it but continue
        PXSLogException( L"Failed to determine/change a privilege.", e, __FUNCTION__ );
    }

    // Snap
    hSnapProcess = CreateToolhelp32Snapshot( TH32CS_SNAPPROCESS, 0 );
    if ( hSnapProcess == INVALID_HANDLE_VALUE )
    {
        throw SystemException( GetLastError(), L"TH32CS_SNAPPROCESS", __FUNCTION__ );
    }
    AutoCloseHandle CloseSnapProcess( hSnapProcess );

    // Get first process, Process32First changes dwSize to the number
    // of bytes written to the structure
    memset( &pe32, 0, sizeof ( pe32 ) );
    pe32.dwSize = sizeof ( pe32 );
    if ( Process32First( hSnapProcess, &pe32 ) == 0 )
    {
        throw SystemException( GetLastError(), L"Process32First", __FUNCTION__ );
    }
    pe32.szExeFile[ ARRAYSIZE( pe32.szExeFile ) - 1 ] = PXS_CHAR_NULL;

    // Enumerate
    do
    {
        memset( &ProcessInfo, 0, sizeof ( ProcessInfo ) );
        ProcessInfo.processID = pe32.th32ProcessID;
        StringCchCopy( ProcessInfo.szExeName,
                       ARRAYSIZE( ProcessInfo.szExeName ), pe32.szExeFile );

        // Get the exe path via the first module, ignore system processes
        if ( pe32.th32ProcessID > 4 )
        {
            hSnapModule = CreateToolhelp32Snapshot( TH32CS_SNAPMODULE,
                                                    pe32.th32ProcessID );
            if ( hSnapModule != INVALID_HANDLE_VALUE )
            {
                AutoCloseHandle CloseSnapModule( hSnapModule );

                memset( &me32, 0, sizeof ( me32 ) );
                me32.dwSize = sizeof ( me32 );
                if ( Module32First( hSnapModule, &me32 ) )
                {
                    me32.szExePath[ ARRAYSIZE( me32.szExePath ) - 1 ] = 0x00;
                    StringCchCopy( ProcessInfo.szExePath,
                                   ARRAYSIZE( ProcessInfo.szExePath ), me32.szExePath );
                }
                else
                {
                    Insert1 = pe32.szExeFile;
                    PXSLogSysWarn1( GetLastError(), L"Module32First failed for '%%1'.", Insert1 );
                }
            }
            else
            {
                Insert1 = pe32.szExeFile;
                PXSLogSysInfo1( GetLastError(), L"TH32CS_SNAPMODULE failed for '%%1'.", Insert1 );
            }

            // Working Set
            try
            {
                ProcessInfo.workingSet = GetWorkingSetSize( ProcessInfo.processID );
            }
            catch ( const Exception& e )
            {
                // Log and continue to the next process
                Insert1 = ProcessInfo.szExePath;
                Message = Format.String1( L"Failed to get the working set for '%%1'.", Insert1 );
                PXSLogException( Message.c_str(), e, __FUNCTION__ );
            }
        }
        m_Processes.Add( ProcessInfo );
    }
    while ( Process32Next( hSnapProcess, &pe32 ) );

    // Reset privilege if it was changed
    if ( privilegeChanged )
    {
        ChangePrivilege( SE_DEBUG_NAME, false );
    }
}

//===============================================================================================//
//  Description:
//      Get the full path to an exe from its process id
//
//  Parameters:
//      processID - process id
//      pExePath  - string object to receive the exe file path
//
//  Remarks:
//      Must first have called FillProcessesInfo
//
//  Returns:
//      void, if the process is not found ExePath will be ""
//===============================================================================================//
void ProcessInformation::GetExePath( DWORD processID, String* pExePath )
{
    size_t i = 0, numProcesses = 0;
    TYPE_PROCESS_INFO ProcessInfo;

    if ( pExePath == nullptr )
    {
        throw ParameterException( L"pExePath", __FUNCTION__ );
    }
    *pExePath = PXS_STRING_EMPTY;

    numProcesses = m_Processes.GetSize();
    for ( i = 0; i < numProcesses; i++ )
    {
        ProcessInfo = m_Processes.Get( i );
        if ( processID == ProcessInfo.processID )
        {
            ProcessInfo.szExePath[ ARRAYSIZE(ProcessInfo.szExePath ) - 1 ] = 0x00;
            *pExePath = ProcessInfo.szExePath;
            break;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the loaded modules as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void ProcessInformation::GetLoadedModuleRecords( TArray<AuditRecord>* pRecords)
{
    DWORD     i = 0;
    HANDLE    hSnapProcess;
    HANDLE    hSnapModule  = INVALID_HANDLE_VALUE;
    String    Insert1;
    FILETIME  fileTime;
    AuditRecord    Record;
    StringArray    ModulePaths;
    MODULEENTRY32  me32;
    PROCESSENTRY32 pe32;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    hSnapProcess = CreateToolhelp32Snapshot( TH32CS_SNAPPROCESS, 0 );
    if ( hSnapProcess == INVALID_HANDLE_VALUE )
    {
       throw SystemException( GetLastError(),
                             L"CreateToolhelp32Snapshot/TH32CS_SNAPPROCESS", __FUNCTION__);
    }
    AutoCloseHandle CloseSnapProcess( hSnapProcess );

    // Get first process,
    memset( &pe32, 0, sizeof ( pe32 ) );
    pe32.dwSize = sizeof ( pe32 );
    if ( Process32First( hSnapProcess, &pe32 ) == 0 )
    {
        throw SystemException( GetLastError(), L"Process32First", __FUNCTION__ );
    }

    do
    {
        if ( pe32.th32ProcessID > 4 )       // <=4 seem to be system processes
        {
            hSnapModule = CreateToolhelp32Snapshot( TH32CS_SNAPMODULE,
                                                    pe32.th32ProcessID );
            if ( hSnapModule != INVALID_HANDLE_VALUE )
            {
                AutoCloseHandle CloseSnapModule( hSnapModule );

                // Get the first module
                memset( &me32, 0, sizeof ( me32 ) );
                me32.dwSize = sizeof ( me32 );
                if ( Module32First( hSnapModule, &me32 ) )
                {
                    do
                    {
                        // Ignore dupilicates
                        me32.szExePath[ ARRAYSIZE(me32.szExePath) - 1 ] = 0x00;
                        ModulePaths.AddUniqueI( me32.szExePath );
                    } while ( Module32Next( hSnapModule, &me32 ) );
                }
                else
                {
                    Insert1 = pe32.szExeFile;
                    PXSLogSysWarn1( GetLastError(), L"Module32First failed for '%%1'.", Insert1 );
                }
            }
            else
            {
               Insert1 = pe32.szExeFile;
               PXSLogSysInfo1( GetLastError(), L"TH32CS_SNAPMODULE failed for '%%1'.", Insert1 );
            }
        }
    }
    while ( Process32Next( hSnapProcess, &pe32 ) );
    ModulePaths.Sort( true );

    ///////////////////////////////////////////////////////////////////////////
    // Process module paths

    File        ModuleFile;
    size_t      numPaths = 0;
    String      ModulePath, ManufacturerName, VersionString, Description;
    String      Drive, Dir, FName, Ext, LastWrite;
    Formatter   Format;
    Directory   ModuleDir;
    FileVersion FileVer;

    // For each module, get its file name, version etc.
    numPaths = ModulePaths.GetSize();
    for ( i = 0; i < numPaths; i++ )
    {
        // Lengthy task, see if got stop signal
        if ( g_pApplication && g_pApplication->GetStopBackgroundTasks() )
        {
            return;
        }

        ModulePath = ModulePaths.Get( i );
        if ( ModulePath.GetLength() > 0 )
        {
            // Get the name, want name.ext
            Drive = PXS_STRING_EMPTY;
            Dir   = PXS_STRING_EMPTY;
            FName = PXS_STRING_EMPTY;
            Ext   = PXS_STRING_EMPTY;
            ModuleDir.SplitPath( ModulePath, &Drive, &Dir, &FName, &Ext );
            FName += Ext;

            ManufacturerName = PXS_STRING_EMPTY;
            VersionString    = PXS_STRING_EMPTY;
            Description      = PXS_STRING_EMPTY;
            LastWrite        = PXS_STRING_EMPTY;
            try
            {
                FileVer.GetVersion( ModulePath, &ManufacturerName, &VersionString, &Description );
                memset( &fileTime, 0, sizeof ( fileTime ) );
                ModuleFile.GetTimes( ModulePath, nullptr, nullptr, &fileTime );
                LastWrite = Format.FileTimeToLocalTimeIso( fileTime );
            }
            catch ( const Exception& e )
            {
                PXSLogException( e, __FUNCTION__ );
            }

            Record.Reset( PXS_CATEGORY_LOADED_MODULES );
            Record.Add( PXS_LOADED_MODULES_NAME        , FName );
            Record.Add( PXS_LOADED_MODULES_VERSION     , VersionString );
            Record.Add( PXS_LOADED_MODULES_MODIFIED    , LastWrite );
            Record.Add( PXS_LOADED_MODULES_MANUFACTURER, ManufacturerName);
            pRecords->Add( Record );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_LOADED_MODULES_NAME );
}

//===============================================================================================//
//  Description:
//      Get the processes that are running as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void ProcessInformation::GetRunningProcesseRecords( TArray< AuditRecord >* pRecords )
{
    size_t      i = 0, memoryUsage = 0, numProcesses = 0;
    File        FileObject;
    String      Value, ManufacturerName, VersionString, Description, ExePath;
    String      LocaleKB;
    Formatter   Format;
    FileVersion FileVer;
    AuditRecord Record;
    TYPE_PROCESS_INFO ProcessInfo;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    PXSGetResourceString( PXS_IDS_135_KB, &LocaleKB );
    ManufacturerName.Allocate( 256 );
    VersionString.Allocate( 256 );
    Description.Allocate( 256 );

    FillProcessesInfo();
    numProcesses = m_Processes.GetSize();
    for ( i = 0; i < numProcesses; i++ )
    {
        ProcessInfo = m_Processes.Get( i );
        ProcessInfo.szExePath[ARRAYSIZE( ProcessInfo.szExePath ) - 1 ] = 0x00;

        // Catch exceptions so can progress to the next process
        try
        {
            // Ignore process 0
            if ( ProcessInfo.processID )
            {
                Record.Reset( PXS_CATEGORY_RUNNING_PROCS );

                // Process Name
                ProcessInfo.szExeName[ ARRAYSIZE( ProcessInfo.szExeName ) - 1 ] = PXS_CHAR_NULL;
                Record.Add( PXS_RUNNING_PROCS_NAME, ProcessInfo.szExeName );

                // Process ID
                Value = Format.UInt32( ProcessInfo.processID );
                Record.Add( PXS_RUNNING_PROCS_PROCESS_ID, Value );

                // Memory usage in KB, leave empty if have no value
                Value = PXS_STRING_EMPTY;
                memoryUsage = GetMemoryUsage( ProcessInfo.processID );
                if ( memoryUsage > 0 )
                {
                    Value  = Format.SizeT( memoryUsage / 1024 );
                    Value += LocaleKB;
                }
                Record.Add( PXS_RUNNING_PROCS_MEM_USAGE_KB, Value );

                // Process Description, make sure have a path
                ManufacturerName = PXS_STRING_EMPTY;
                VersionString    = PXS_STRING_EMPTY;
                Description      = PXS_STRING_EMPTY;
                ExePath          = ProcessInfo.szExePath;
                if ( FileObject.Exists( ExePath ) )
                {
                    FileVer.GetVersion( ExePath,
                                        &ManufacturerName, &VersionString, &Description );
                }
                Record.Add( PXS_RUNNING_PROCS_DESCRIPTION, Description );
                pRecords->Add( Record );
            }
        }
        catch ( const Exception& eProcess )
        {
            // Typically cannot get data for special system processes
            if ( ProcessInfo.processID > 4 )
            {
                PXSLogException( ProcessInfo.szExePath, eProcess, __FUNCTION__);
            }
        }
    }
    PXSSortAuditRecords( pRecords, PXS_RUNNING_PROCS_NAME );
}

//===============================================================================================//
//  Description:
//      Determine if the current process has a given token privilege
//
//  Parameters:
//      pszPrivilege - name of privilege to test for
//
//  Remarks:
//      Will log errors and return false rather that throw because the caller
//      is usually testing if need to set a privilege.
//
//  Returns:
//      true if has privilege
//===============================================================================================//
bool ProcessInformation::IsPrivilegeEnabled( LPCWSTR pszPrivilege )
{
    bool    enabled = false, match = false;
    DWORD   i = 0, ReturnLength = 0, lastError = 0, cchName = 0;
    DWORD   TokenInformationLength = 0;
    wchar_t szPrivilegeName[ 256 ] = { 0 };     // Enough for any privilege name
    HANDLE  ProcessHandle = nullptr, TokenHandle = nullptr;
    String  Insert1;
    AllocateBytes     AllocBytes;
    TOKEN_PRIVILEGES* pTokenPrivileges = nullptr;

    if ( pszPrivilege == nullptr )
    {
       return false;    // Nothing to do or error
    }

    // Get a pseudo-handle to the process, there is no need to close it
    ProcessHandle = GetCurrentProcess();
    if ( OpenProcessToken( ProcessHandle, TOKEN_READ, &TokenHandle ) == 0 )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysWarn1( GetLastError(), L"OpenProcessToken failed in '%%1'.", Insert1 );
        return false;
    }
    AutoCloseHandle CloseTokenHandle( TokenHandle );

    // Get the security token information using 'TokenPrivileges' the first
    // call is to get the required buffer size
    if ( GetTokenInformation( TokenHandle,
                              TokenPrivileges,
                              nullptr,        // i.e. null buffer
                              0,              // i.e. zero length buffer
                              &ReturnLength  ) )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo1( L"GetTokenInformation said no data '%%1'.", Insert1 );
        return false;
    }

    // Test for error other than insufficient buffer
    lastError = GetLastError();
    if ( lastError != ERROR_INSUFFICIENT_BUFFER )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysError1( lastError,
                         L"GetTokenInformation(1) failed in '%%1'.", Insert1 );
        return false;
    }

    // Second call, allocate extra bytes
    ReturnLength     = PXSMultiplyUInt32( ReturnLength, 2 );
    pTokenPrivileges = reinterpret_cast<TOKEN_PRIVILEGES*>( AllocBytes.New( ReturnLength ) );
    TokenInformationLength = ReturnLength;
    ReturnLength = 0;
    if ( GetTokenInformation( TokenHandle,
                              TokenPrivileges,
                              pTokenPrivileges, TokenInformationLength, &ReturnLength ) == 0 )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysError1( GetLastError(), L"GetTokenInformation(2) failed in '%%1'.", Insert1 );
        return false;
    }

    for ( i = 0; i < pTokenPrivileges->PrivilegeCount; i++ )
    {
        // Get the privilege name
        memset( szPrivilegeName, 0, sizeof ( szPrivilegeName ) );
        cchName = ARRAYSIZE( szPrivilegeName );
        if ( LookupPrivilegeName( nullptr,  // Local computer
                                  &pTokenPrivileges->Privileges[ i ].Luid,
                                  szPrivilegeName,
                                  &cchName ) )    // In characters
        {
            szPrivilegeName[ARRAYSIZE( szPrivilegeName ) - 1 ] = PXS_CHAR_NULL;
            if ( lstrcmpi( pszPrivilege, szPrivilegeName ) == 0 )
            {
                match = true;
                if ( SE_PRIVILEGE_ENABLED & pTokenPrivileges->Privileges[ i ].Attributes )
                {
                    enabled = true;
                }
            }
        }

        if ( match )
        {
            break;
        }
    }

    return enabled;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the memory usage of the specified process id.
//
//  Parameters:
//      processID - process id
//
//  Remarks:
//      This is the cached value obtained when FillProcessesInfo was called
//
//  Returns:
//      size_t of memory, zero if process id not recognised
//===============================================================================================//
size_t ProcessInformation::GetMemoryUsage( DWORD processID )
{
    size_t  i = 0, workingSet = 0, numProcesses = 0;
    TYPE_PROCESS_INFO ProcessInfo;

    numProcesses = m_Processes.GetSize();
    for ( i = 0; i < numProcesses; i++ )
    {
        ProcessInfo = m_Processes.Get( i );
        if ( processID == ProcessInfo.processID )
        {
            workingSet = ProcessInfo.workingSet;
            break;
        }
    }

    return workingSet;
}

//===============================================================================================//
//  Description:
//      Get the memory working set of the specified process identfier
//
//  Parameters:
//      processID - the process ID
//
//  Returns:
//      void
//===============================================================================================//
size_t ProcessInformation::GetWorkingSetSize( DWORD processID )
{
    HANDLE  hProcess;
    PROCESS_MEMORY_COUNTERS pmc;

    hProcess = OpenProcess( PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processID );
    if ( hProcess == nullptr )
    {
        throw SystemException( GetLastError(), L"OpenProcess failed.", __FUNCTION__ );
    }
    AutoCloseHandle CloseProcess( hProcess );

    memset( &pmc, 0, sizeof ( pmc ) );
    pmc.cb = sizeof ( pmc );
    if ( GetProcessMemoryInfo( hProcess, &pmc, sizeof ( pmc ) ) == 0 )
    {
        throw SystemException( GetLastError(), L"GeprocessInfoMemoryInfo failed.", __FUNCTION__ );
    }

    return pmc.WorkingSetSize;
}
