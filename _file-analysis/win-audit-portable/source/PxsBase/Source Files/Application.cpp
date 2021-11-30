///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Application Class Implementation
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
#include "PxsBase/Header Files/Application.h"

// 2. C System Files
#include <stdint.h>
#include <ShlObj.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AutoUnlockMutex.h"
#include "PxsBase/Header Files/ByteArray.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/Mutex.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Resources.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/SystemInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Global POD Variables
///////////////////////////////////////////////////////////////////////////////////////////////////

Application* g_pApplication = nullptr;  // Pointer to application object

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

//
// Default constructor, an empty application object
// Do not throw any exceptions
//
Application::Application()
            :m_bStopBackgroundTasks( false ),
             m_isAppLogging( FALSE ),
             m_hWndMainFrame( nullptr ),
             m_uLogLevel( PXS_LOG_LEVEL_NONE ),
             m_crGradientOne( CLR_INVALID ),
             m_crGradientTwo( CLR_INVALID ),
             m_Mutex(),
             m_Logger(),
             m_ApplicationName(),
             m_AppTempFolderName(),
             m_CopyrightNotice(),
             m_LastAccessedDirectory(),
             m_LogDirectoryPath(),
             m_ManufacturerAddress(),
             m_ManufacturerName(),
             m_ProductID(),
             m_RegisteredOrgName(),
             m_RegisteredOwnerName(),
             m_SupportEmail(),
             m_WebSiteURL(),
             m_StringTable(),
             m_AppResourceLanguageIDs()
{
}

// Copy constructor - not allowed so no implementation

// Destructor - do not throw any exceptions
Application::~Application()
{
    try
    {
        StopLogger();
    }
    catch ( const Exception& e )
    {
       PXSLogException( e, __FUNCTION__ );
    }
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
//      Add a resource identifier for a string table for the specified language
//
//  Parameters:
//      resourceID      - the resource id
//      primaryLanguage - the primary language id
//      subLanguage     - the sub language id
//
//  Returns:
//        void
//===============================================================================================//
void Application::AddAppResourceLanguageID( WORD resourceID,
                                            WORD primaryLanguage,
                                            WORD subLanguage )
{
    RESOURCE_LANGIDS ResourceLangIDs;

    if ( resourceID == 0 )
    {
        throw ParameterException( L"resourceID", __FUNCTION__ );
    }

    ResourceLangIDs.resourceID      = resourceID;
    ResourceLangIDs.primaryLanguage = primaryLanguage;
    ResourceLangIDs.subLanguage     = subLanguage;
    m_AppResourceLanguageIDs.Add( ResourceLangIDs );
}

//===============================================================================================//
//  Description:
//      Get the application's name
//
//  Parameters:
//      pApplicationName - receives the application name
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetApplicationName( String* pApplicationName )
{
    if ( pApplicationName == nullptr )
    {
        throw ParameterException( L"pApplicationName", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pApplicationName = m_ApplicationName;
}

//===============================================================================================//
//  Description:
//      Get the temporary directory specific to this application
//
//  Parameters:
//      create               - create it if it does not exist
//      pApplicationTempPath - string object to receive the path
//
//  Remarks:
//      Application temporary directory, is formed from the user's
//      temporary directory and an application specific folder.
//
//      Creation occurs if flag is set and an application specific
//      folder has been named.
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetApplicationTempPath( bool create,
                                          String* pApplicationTempPath )
{
    String    SubDirPath;
    Directory DirObject;

    if ( pApplicationTempPath == nullptr )
    {
        throw ParameterException( L"pApplicationTempPath", __FUNCTION__ );
    }
    *pApplicationTempPath = PXS_STRING_EMPTY;

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );

    // Default to the user's temp directory
    DirObject.GetTempDirectory( pApplicationTempPath );

    // Build the name of sub-folder that will used by the application for
    // temporary storage.
    if ( m_AppTempFolderName.GetLength() )
    {
        SubDirPath  = *pApplicationTempPath;
        SubDirPath += m_AppTempFolderName;
        if ( SubDirPath.EndsWithCharacter( PXS_PATH_SEPARATOR ) == false )
        {
            SubDirPath += PXS_PATH_SEPARATOR;
        }

        if ( create )
        {
            if ( DirObject.Exists( SubDirPath ) == false )
            {
                DirObject.CreateNew( SubDirPath );
            }
        }
        *pApplicationTempPath = SubDirPath;
    }
}

//===============================================================================================//
//  Description:
//      Get the copyright notice
//
//  Parameters:
//      pCopyrightNotice - receives the copyright notice
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetCopyrightNotice( String* pCopyrightNotice )
{
    if ( pCopyrightNotice == nullptr )
    {
        throw ParameterException( L"pCopyrightNotice", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pCopyrightNotice = m_CopyrightNotice;
}

//===============================================================================================//
//  Description:
//      Get the details about the environment
//
//  Parameters:
//      extraInfo      - indicates want extra information
//      pDiagnosticData - string array to receive data
//
//  Remarks:
//      Strings are formatted for displaying/logging.
//      Does not throw exceptions as this method is used to obtain
//      exception information
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetDiagnosticData( bool extraInfo, StringArray* pDiagnosticData )
{
    File    ExeFile;
    BYTE    productType = 0;
    UINT    errorMode   = 0, driveType = 0;
    WORD    suiteMask   = 0;
    BYTE*   pExecutable = nullptr;
    DWORD   infoType    = 0, executableSize = 0;
    UINT64  fileSize    = 0;
    wchar_t szRootPathName[ 8 ] = { 0 };  // drive letter + colon + separator
    String  DataString, ExePath, ManufacturerName, Version, Md5Hash;
    String  ExeDirectory, Description, ValueString, TranslatedDriveType;
    String  NetFunctionVersion, WmiWin32Version, Kernel32Version, RegCurrentVersion;
    Formatter   Format;
    SYSTEM_INFO si;
    FileVersion FileVer;
    AllocateBytes      AllocBytes;
    SystemInformation  SystemInfo;

    // Do not show errors
    if ( pDiagnosticData == nullptr )
    {
        return;
    }
    errorMode = SetErrorMode( SEM_FAILCRITICALERRORS );

    // As this is a diagnostic method will trap exceptions
    try
    {
        pDiagnosticData->RemoveAll();
        DataString  = L"Timestamp                : ";
        DataString += Format.LocalTimeInIsoFormat();
        pDiagnosticData->Add( DataString );

        DataString  = L"Application name         : ";
        DataString += m_ApplicationName;
        pDiagnosticData->Add( DataString );

        DataString = L"Application version      : ";
        PXSGetExePath( &ExePath );
        FileVer.GetVersion( ExePath,
                            &ManufacturerName, &Version, &Description );
        DataString += Version;
        pDiagnosticData->Add( DataString );

        DataString  = L"Executable Path          : ";
        DataString += ExePath;
        pDiagnosticData->Add( DataString );

        DataString  = L"Executable on drive type : ";
        PXSGetExeDirectory( &ExeDirectory );
        if ( ExeDirectory.GetLength() )
        {
            szRootPathName[0] = ExeDirectory.CharAt( 0 );
            szRootPathName[1] = PXS_CHAR_COLON;
            szRootPathName[2] = PXS_PATH_SEPARATOR;
            szRootPathName[3] = PXS_CHAR_NULL;
            driveType = GetDriveType( szRootPathName );
            PXSTranslateDriveType( driveType, &TranslatedDriveType );
        }
        DataString += TranslatedDriveType;
        pDiagnosticData->Add( DataString );

        DataString  = L"User default language ID : ";
        DataString += Format.UInt16Hex( GetUserDefaultLangID() );
        pDiagnosticData->Add( DataString );

        DataString  = L"Administrator            : ";
        DataString += Format.Int32YesNo( SystemInfo.IsAdministrator() );
        pDiagnosticData->Add( DataString );

        DataString  = L"Operating system version : ";
        ValueString = PXS_STRING_EMPTY;
        SystemInfo.GetMajorMinorBuild( &ValueString );
        DataString += ValueString;
        pDiagnosticData->Add( DataString );

        DataString  = L"Service pack version     : ";
        ValueString = PXS_STRING_EMPTY;
        SystemInfo.GetServicePackMajorDotMinor( &ValueString );
        DataString += ValueString;
        pDiagnosticData->Add( DataString );

        if ( extraInfo )
        {
            DataString  = L"Operating system edition : ";
            ValueString = PXS_STRING_EMPTY;
            SystemInfo.GetEdition( &ValueString );
            DataString += ValueString;
            pDiagnosticData->Add( DataString );

            DataString   = L"Product suite            : ";
            suiteMask = SystemInfo.GetVersionSuiteMask();
            DataString  += Format.UInt16Hex( suiteMask );
            pDiagnosticData->Add( DataString );

            DataString     = L"Product type             : ";
            productType = SystemInfo.GetVersionProductType();
            DataString     += Format.UInt8Hex( productType, true );
            pDiagnosticData->Add( DataString );

            // Product info, NT6 and above
            if ( SystemInfo.GetMajorVersion() >= 6 )
            {
                DataString  = L"Product type  (Vista+)   : ";
                infoType = SystemInfo.GetProductInfoType();
                DataString += Format.UInt32Hex( infoType, true );
                pDiagnosticData->Add( DataString );
            }

            // MD5 of the exe, continue on error
            DataString = L"Executable's MD5         : ";
            try
            {
                // Digital signature, it is slow so limit file size will check
                PXSGetExePath( &ExePath );
                ExeFile.Open( ExePath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, 1, false );
                fileSize       = ExeFile.GetSize();
                executableSize = PXSCastUInt64ToUInt32( fileSize );
                pExecutable    = AllocBytes.New( executableSize );
                ExeFile.Read( pExecutable, executableSize );
                PXSGetMd5Hash( pExecutable, executableSize, &Md5Hash );
                DataString += Md5Hash;
                ExeFile.Close();
            }
            catch ( const Exception& )    // Ignore
            { }
            pDiagnosticData->Add( DataString );

            DataString  = L"ANSI code-page identifier: ";
            DataString += Format.UInt32( GetACP() );
            pDiagnosticData->Add( DataString );

            memset( &si, 0, sizeof ( si ) );
            GetSystemInfo( &si );

            DataString = Format.StringUInt32( L"Processors (logical)     : %%1",
                                             si.dwNumberOfProcessors );
            pDiagnosticData->Add( DataString );

            DataString  = L"Processor architecture   : ";
            DataString += Format.UInt16( si.wProcessorArchitecture );
            pDiagnosticData->Add( DataString );

            DataString  = L"Active processor mask    : ";
            DataString += Format.BinaryUInt64( si.dwActiveProcessorMask );
            pDiagnosticData->Add( DataString );

            DataString  = L"64-bit processor         : ";
            DataString += Format.Int32YesNo( SystemInfo.Is64BitProcessor() );
            pDiagnosticData->Add( DataString );

            DataString  = L"Emulating 32-bit (WOW64) : ";
            DataString += Format.Int32YesNo( SystemInfo.IsWOW64() );
            pDiagnosticData->Add( DataString );

            DataString  = L"Pre-install environment  : ";
            DataString += Format.Int32YesNo( SystemInfo.IsPreInstall() );
            pDiagnosticData->Add( DataString );

            DataString = L"Terminal services client : ";
            DataString += Format.Int32YesNo( SystemInfo.IsRemoteSession() );
            pDiagnosticData->Add( DataString );

            DataString = L"Remotely controlled      : ";
            DataString += Format.Int32YesNo( SystemInfo.IsRemoteControl() );
            pDiagnosticData->Add( DataString );

            DataString  = L"Session name             : ";
            ValueString = PXS_STRING_EMPTY;
            SystemInfo.GetSessionName( &ValueString );
            DataString += ValueString;
            pDiagnosticData->Add( DataString );

            // Version strings
            SystemInfo.GetVersionStrings( &NetFunctionVersion,
                                          &WmiWin32Version, &Kernel32Version, &RegCurrentVersion );
            DataString  = L"WKSTA_INFO_100           : ";
            DataString += NetFunctionVersion;
            pDiagnosticData->Add( DataString );

            DataString  = L"Win32_OperatingSystem    : ";
            DataString += WmiWin32Version;
            pDiagnosticData->Add( DataString );

            DataString  = L"Kernel32.dll             : ";
            DataString += Kernel32Version;
            pDiagnosticData->Add( DataString );

            DataString  = L"Registry Current Version : ";
            DataString += RegCurrentVersion;
            pDiagnosticData->Add( DataString );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
    SetErrorMode( errorMode );
}

//===============================================================================================//
//  Description:
//      Get diagnostic data about the application and system as a string
//
//  Parameters:
//      pDiagnostics - receives the string data
//
//  Returns:
//       void
//===============================================================================================//
void Application::GetDiagnostics( String* pDiagnostics )
{
    size_t i = 0, numElements = 0;
    StringArray DiagnosticData;

    if ( pDiagnostics == nullptr )
    {
        throw ParameterException( L"pDiagnostics", __FUNCTION__ );
    }
    *pDiagnostics = PXS_STRING_EMPTY;

    try
    {
        // Application and Operating System Details
        *pDiagnostics  = L"Application & Operating System\r\n";
        *pDiagnostics += L"==============================\r\n\r\n";

        GetDiagnosticData( true, &DiagnosticData );
        numElements = DiagnosticData.GetSize();
        for ( i = 0; i < numElements; i++ )
        {
            *pDiagnostics += DiagnosticData.Get( i );
            *pDiagnostics += PXS_STRING_CRLF;
        }
    }
    catch ( const Exception& e )
    {
        *pDiagnostics += L"Error getting application/system data.";
        *pDiagnostics += PXS_STRING_CRLF;
        *pDiagnostics += e.GetMessage();
    }
    *pDiagnostics += PXS_STRING_CRLF;
}

//===============================================================================================//
//  Description:
//      Get the instance handle of the runtime process.
//
//  Parameters:
//      hInstance - instance handle to receive the handle
//
//  Returns:
//      HINSTANCE
//===============================================================================================//
HINSTANCE Application::GetExeInstanceHandle()
{
    return GetModuleHandle( nullptr );
}

//===============================================================================================//
//  Description:
//      Get the application's gradient colours
//
//  Parameters:
//      pGradientOne - optional, receives the first colour
//      pGradientTwo - optional, receives the second colour
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetGradientColours( COLORREF* pGradientOne, COLORREF* pGradientTwo )
{
    if ( ( pGradientOne == nullptr ) && ( pGradientTwo == nullptr ) )
    {
        return;     // Nothing to do
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    if ( pGradientOne )
    {
        *pGradientOne = m_crGradientOne;
    }

    if ( pGradientTwo )
    {
        *pGradientTwo = m_crGradientTwo;
    }
}

//===============================================================================================//
//  Description:
//      Get the handle of the application's main window.
//
//  Parameters:
//      void
//
//  Returns:
//      HWND - handle to application's main window
//===============================================================================================//
HWND Application::GetHwndMainFrame() const
{
    return m_hWndMainFrame;
}

//===============================================================================================//
//  Description:
//      Get the directory path last accessed by the user using OPENFILENAME
//
//  Parameters:
//      pLastAccessedDirectory - receives the directory path
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetLastAccessedDirectory( String* pLastAccessedDirectory )
{
    if ( pLastAccessedDirectory == nullptr )
    {
        throw ParameterException( L"pLastAccessedDirectory", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pLastAccessedDirectory = m_LastAccessedDirectory;
}

//===============================================================================================//
//  Description:
//      Get the logging level
//
//  Parameters:
//      None
//
//  Returns:
//      DWORD of defined logging level constant
//===============================================================================================//
DWORD Application::GetLogLevel() const
{
    return m_uLogLevel;
}

//===============================================================================================//
//  Description:
//      Get the log messages stored in the string array
//
//  Parameters:
//      purge        - if true, purges any existing messages in the logger
//                     string array after first copying them to the input array
//      pLogMessages - receives the log messages
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetLogMessages( bool purge, StringArray* pLogMessages )
{
    if ( pLogMessages == nullptr )
    {
        throw ParameterException( L"pLogMessages", __FUNCTION__ );
    }
    pLogMessages->RemoveAll();

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_Logger.GetLogMessages( pLogMessages );
    if ( purge )
    {
        m_Logger.PurgeLogMessages();
    }
}

//===============================================================================================//
//  Description:
//      Get the manufacturer's address
//
//  Parameters:
//      pManufacturerAddress - receives the address
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetManufacturerAddress( String* pManufacturerAddress )
{
    if ( pManufacturerAddress == nullptr )
    {
        throw ParameterException( L"pManufacturerAddress", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pManufacturerAddress = m_ManufacturerAddress;
}

//===============================================================================================//
//  Description:
//      Get the information about the software manufacturer
//
//  Parameters:
//      pCopyrightNotice     - optional, receives the copyright notice
//      pManufacturerAddress - optional, receives the manufacturer's address
//      pManufacturerName    - optional, receives the manufacturer's name
//      pRegisteredOrgName   - optional, receives the registered org's name
//      pRegisteredOwnerName - optional, receives the software's owner name
//      pProductID           - optional, receives the software's product ID
//      pSupportEmail        - optional, receives the support e-mail address
//      pWebSiteURL          - optional, receives the manufactures web URL
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetManufacturerInfo( String* pCopyrightNotice,
                                       String* pManufacturerAddress,
                                       String* pManufacturerName,
                                       String* pRegisteredOrgName,
                                       String* pRegisteredOwnerName,
                                       String* pProductID,
                                       String* pSupportEmail, String* pWebSiteURL )
{
    if ( pCopyrightNotice )
    {
        *pCopyrightNotice = m_CopyrightNotice;
    }

    if ( pManufacturerAddress )
    {
        *pManufacturerAddress = m_ManufacturerAddress;
    }

    if ( pManufacturerName )
    {
        *pManufacturerName = m_ManufacturerName;
    }

    if ( pRegisteredOrgName )
    {
        *pRegisteredOrgName = m_RegisteredOrgName;
    }

    if ( pRegisteredOwnerName )
    {
        *pRegisteredOwnerName = m_RegisteredOwnerName;
    }

    if ( pProductID )
    {
        *pProductID = m_ProductID;
    }

    if ( pSupportEmail )
    {
        *pSupportEmail = m_SupportEmail;
    }

    if ( pWebSiteURL )
    {
        *pWebSiteURL = m_WebSiteURL;
    }
}

//===============================================================================================//
//  Description:
//      Get the manufacturer name property
//
//  Parameters:
//      pManufacturerName - receives the manufacture name
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetManufacturerName( String* pManufacturerName )
{
    if ( pManufacturerName == nullptr )
    {
        throw ParameterException( L"pManufacturerName", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pManufacturerName = m_ManufacturerName;
}

//===============================================================================================//
//  Description:
//      Get the product identifier property
//
//  Parameters:
//      pProductID - receives the product identifier
//
//  Returns:
//     void
//===============================================================================================//
void Application::GetProductID( String* pProductID )
{
    if ( pProductID == nullptr )
    {
        throw ParameterException( L"pProductID", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pProductID = m_ProductID;
}

//===============================================================================================//
//  Description:
//      Get a string based on its ID
//
//  Parameters:
//      resourceID     - unique ID for the string
//      pResourceString- receives the resource string
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetResourceString( DWORD resourceID, String* pResourceString )
{
    String Insert1, Insert2, Insert3;

    GetResourceString3( resourceID,
                        Insert1, Insert2, Insert3, pResourceString );
}

//===============================================================================================//
//  Description:
//      Get a string based on its ID
//
//  Parameters:
//      resourceID     - unique ID for the string
//      Insert1        - insert string for placeholder %%1
//      pResourceString- receives the resource string
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetResourceString1( DWORD resourceID,
                                      const String& Insert1, String* pResourceString )
{
    String Insert2, Insert3;

    GetResourceString3( resourceID,
                        Insert1, Insert2, Insert3, pResourceString );
}

//===============================================================================================//
//  Description:
//      Get a string based on its ID
//
//  Parameters:
//      resourceID     - unique ID for the string
//      Insert1        - insert string for place holder %%1
//      Insert2        - insert string for place holder %%2
//      pResourceString- receives the resource string
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetResourceString2( DWORD resourceID,
                                      const String& Insert1,
                                      const String& Insert2, String* pResourceString )
{
    String Insert3;

    GetResourceString3( resourceID, Insert1, Insert2, Insert3, pResourceString );
}

//===============================================================================================//
//  Description:
//      Get a string based on its ID
//
//  Parameters:
//      resourceID     - unique ID for the string
//      Insert1        - insert string for placeholder %%1
//      Insert2        - insert string for placeholder %%2
//      Insert3        - insert string for placeholder %%3
//      pResourceString- receives the resource string
//
//  Remarks:
//      Defaults to English. If a non-English resource or file has been
//      loaded will use the foreign string
//
//  Returns:
//      void. On error, "" is put into output ResourceString
//===============================================================================================//
void Application::GetResourceString3( DWORD resourceID,
                                      const String& Insert1,
                                      const String& Insert2,
                                      const String& Insert3, String* pResourceString )
{
    bool      bFound   = false;
    ID_TEXT*  pElement = nullptr;
    Formatter Format;

    if ( pResourceString == nullptr )
    {
        return;
    }
    *pResourceString = PXS_STRING_EMPTY;

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    if ( m_StringTable.IsEmpty() )
    {
        return;
    }

    // Find the matching string id
    m_StringTable.Rewind();
    do
    {
        pElement = m_StringTable.GetPointer();
        if ( pElement && ( resourceID == pElement->id ) )
        {
            *pResourceString = Format.String3( pElement->Text.c_str(), Insert1, Insert2, Insert3 );
        }
    } while ( ( bFound == false ) && ( m_StringTable.Advance() ) );
}

//===============================================================================================//
//  Description:
//      Get the registered organization name
//
//  Parameters:
//      pRegisteredOrgName - receives the registered organisation
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetRegisteredOrgName( String* pRegisteredOrgName )
{
    if ( pRegisteredOrgName == nullptr )
    {
        throw ParameterException( PXS_STRING_EMPTY, __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pRegisteredOrgName = m_RegisteredOrgName;
}

//===============================================================================================//
//  Description:
//      Get the registered owner name
//
//  Parameters:
//      pRegisteredOwner - receives the registered owner
//
//  Returns:
//      String reference to the registered owner name
//===============================================================================================//
void Application::GetRegisteredOwnerName( String* pRegisteredOwnerName )
{
    if ( pRegisteredOwnerName == nullptr )
    {
        throw ParameterException( L"pRegisteredOwnerName", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pRegisteredOwnerName = m_RegisteredOwnerName;
}

//===============================================================================================//
//  Description:
//      Get if want to any running background/worker tasks to stop
//
//  Parameters:
//      None
//
//  Remarks:
//      A worker thread should poll at regular intervals to see if
//      it needs to stop.
//
//  Returns:
//      true if want to stop task, else false
//===============================================================================================//
bool Application::GetStopBackgroundTasks() const
{
    return m_bStopBackgroundTasks;
}

//===============================================================================================//
//  Description:
//      Get the support email property
//
//  Parameters:
//      pSupportEmail - receives the support e-mail
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetSupportEmail( String* pSupportEmail )
{
    if ( pSupportEmail == nullptr )
    {
        throw ParameterException( L"pSupportEmail", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pSupportEmail = m_SupportEmail;
}

//===============================================================================================//
//  Description:
//      Get the web site URL property
//
//  Parameters:
//      pWebSiteURL - receives the website URL property
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetWebSiteURL( String* pWebSiteURL )
{
    if ( pWebSiteURL == nullptr )
    {
        throw ParameterException( L"pWebSiteURL", __FUNCTION__ );
    }
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    *pWebSiteURL = m_WebSiteURL;
}

//===============================================================================================//
//  Description:
//      Determine whether or not this application is being run from a
//      removable drive
//
//  Parameters:
//      None
//
//  Returns:
//      true if application is running from a removable drive
//===============================================================================================//
bool Application::IsOnRemoveableDrive()
{
    bool    removeable = false;
    String  ExeDirectory;
    wchar_t szRootPathName[ 8 ] = { 0 };  // letter + colon + separator

    PXSGetExeDirectory( &ExeDirectory );
    ExeDirectory.Trim();
    if ( ExeDirectory.GetLength() )
    {
        szRootPathName[ 0 ] = ExeDirectory.CharAt( 0 );
        szRootPathName[ 1 ] = PXS_CHAR_COLON;
        szRootPathName[ 2 ] = PXS_PATH_SEPARATOR;
        szRootPathName[ 3 ] = PXS_CHAR_NULL;
        if ( DRIVE_REMOVABLE == GetDriveType( szRootPathName ) )
        {
            removeable = true;
        }
    }
    else
    {
        throw SystemException( ERROR_INVALID_NAME, L"ExeDirectory = 0", __FUNCTION__ );
    }

    return removeable;
}

//===============================================================================================//
//  Description:
//      Says whether or not logging is taking place
//
//  Parameters:
//      None
//
//  Remarks:
//      Does not throw exceptions
//
//  Returns:
//      true if application is logging, else false
//===============================================================================================//
bool Application::IsLogging()
{
    if ( m_isAppLogging )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Read the specified exe resource into a byte array
//
//  Parameters:
//      resourceID  - resource id of text data file
//      pBytes      - receives the data bytes
//
//  Returns:
//      void
//===============================================================================================//
void Application::LoadByteDataResource( DWORD resourceID, ByteArray* pBytes ) const
{
    size_t  resSize;
    void*   pResource;
    HRSRC   hrsrcFile;
    HGLOBAL hResource;
    String    ErrorMessage;
    Formatter Format;

    if ( resourceID == 0 )
    {
        return;   // Nothing to do
    }

    if ( pBytes == nullptr )
    {
        throw ParameterException( L"pBytes", __FUNCTION__ );
    }
    pBytes->Zero();

    // Load the resource
    hrsrcFile = FindResource( nullptr, MAKEINTRESOURCE( resourceID ), RT_RCDATA );
    if ( hrsrcFile == nullptr )
    {
        ErrorMessage = Format.StringUInt32( L"resourceID=%%1", resourceID );
        throw SystemException( GetLastError(), ErrorMessage.c_str(), "FindResource" );
    }

    // It is not necessary to destroy the resource
    hResource = LoadResource( nullptr, hrsrcFile );
    if ( hResource == nullptr )
    {
        ErrorMessage = Format.StringUInt32( L"resourceID=%%1", resourceID );
        throw SystemException( GetLastError(), ErrorMessage.c_str(), "LoadResource" );
    }

    // It is not necessary to unlock the resource
    pResource = LockResource( hResource );
    if ( pResource == nullptr )
    {
        ErrorMessage = Format.StringUInt32( L"resourceID=%%1", resourceID );
        throw SystemException( GetLastError(), ErrorMessage.c_str(), "LockResource" );
    }
    resSize = SizeofResource( nullptr, hrsrcFile );
    pBytes->Append( static_cast< const BYTE*>( pResource ), resSize );
}

//===============================================================================================//
//  Description:
//      Load this application's strings from a resource file
//
//  Parameters:
//      primaryLanguage - the primary language id
//      subLanguage     - the sublanguage id
//
//  Returns:
//      void
//===============================================================================================//
void Application::LoadStringFile( WORD primaryLanguage, WORD subLanguage )
{
    WORD    pxsBaseLocaleID = 0, appEnglishID = 0, appLocaleID = 0;
    LANGID  langID = 0;
    ID_TEXT Element;
    TList< ID_TEXT > TempTable;
    TList< ID_TEXT > StringTable;

    // These combinations have a special meaning, see MSDN MAKELANGID.
    if ( primaryLanguage == LANG_NEUTRAL )
    {
        if ( subLanguage == SUBLANG_DEFAULT )
        {
            langID = GetUserDefaultLangID();
            primaryLanguage = static_cast<WORD>( PRIMARYLANGID( langID ) );
            subLanguage     = static_cast<WORD>( SUBLANGID( langID ) );
        }
        else if ( subLanguage == SUBLANG_SYS_DEFAULT )
        {
            langID = GetSystemDefaultLangID();
            primaryLanguage = static_cast<WORD>( PRIMARYLANGID( langID ) );
            subLanguage     = static_cast<WORD>( SUBLANGID( langID ) );
        }
        else
        {
            // For our purposes, LANG_NEUTRAL means English (UK)
            primaryLanguage = LANG_ENGLISH;
            subLanguage     = SUBLANG_ENGLISH_UK;
        }
    }
    GetLanguageResourceIDs( primaryLanguage,
                            subLanguage, &pxsBaseLocaleID, &appEnglishID, &appLocaleID );

    // Make the string table by first populating with English then
    // override with any non-English strings. This is done in case there is
    // no translation for a string so can fall back to English. Start with
    // the PxsBase library strings then add those specific to the application
    // so that the string identifiers are in numerical order.
    GetStringTable( IDR_PXSBASE_STRINGS_EN, &StringTable );
    if ( pxsBaseLocaleID != IDR_PXSBASE_STRINGS_EN )
    {
        GetStringTable( pxsBaseLocaleID, &TempTable );
        MergeStringTables( &TempTable, &StringTable );
    }

    // Application strings
    TempTable.RemoveAll();
    GetStringTable( appEnglishID, &TempTable );
    MergeStringTables( &TempTable, &StringTable );
    if ( appLocaleID != appEnglishID )
    {
        TempTable.RemoveAll();
        GetStringTable( appLocaleID, &TempTable );
        MergeStringTables( &TempTable, &StringTable );
    }

    // Replace at class scope. Note, this is not exception safe, sorry!
    if ( StringTable.IsEmpty() == false )
    {
        m_StringTable.RemoveAll();
        StringTable.Rewind();
        do
        {
            Element = StringTable.Get();
            m_StringTable.Append( Element );
        } while ( StringTable.Advance() );
    }
}

//===============================================================================================//
//  Description:
//      Read the lines of a Unicode text file resource in an array
//
//  Parameters:
//      resourceID  - resource id of text data file
//      pStringData - receives the text
//
//  Returns:
//      void
//===============================================================================================//
void Application::LoadTextDataResource( DWORD resourceID, String* pStringData ) const
{
    size_t  resSize   = 0, numTextBytes = 0;
    void*   pResource = nullptr;
    BYTE*   pText     = nullptr;
    HRSRC   hrsrcFile = nullptr;
    String  ErrorMessage;
    HGLOBAL hResource = nullptr;
    Formatter     Format;
    AllocateBytes AllocBytes;

    if ( resourceID == 0 )
    {
        return;   // Nothing to do
    }

    if ( pStringData == nullptr )
    {
        throw ParameterException( L"pStringData", __FUNCTION__ );
    }
    *pStringData = PXS_STRING_EMPTY;

    // Load the resource
    hrsrcFile = FindResource( nullptr, MAKEINTRESOURCE( resourceID ), RT_RCDATA );
    if ( hrsrcFile == nullptr )
    {
        ErrorMessage = Format.StringUInt32( L"resourceID=%%1", resourceID );
        throw SystemException( GetLastError(), ErrorMessage.c_str(), "FindResource" );
    }

    // It is not necessary to destroy the resource
    hResource = LoadResource( nullptr, hrsrcFile );
    if ( hResource == nullptr )
    {
        ErrorMessage = Format.StringUInt32( L"resourceID=%%1", resourceID );
        throw SystemException( GetLastError(), ErrorMessage.c_str(), "LoadResource" );
    }

    // It is not necessary to unlock the resource
    pResource = LockResource( hResource );
    if ( pResource == nullptr )
    {
        ErrorMessage = Format.StringUInt32( L"resourceID=%%1", resourceID );
        throw SystemException( GetLastError(), ErrorMessage.c_str(), "LockResource" );
    }

    // Allocate an array with an additional character for a null
    // terminator. Use wchar_t in the event UNICODE is defined
    resSize      = SizeofResource( nullptr, hrsrcFile );
    numTextBytes = PXSAddSizeT( resSize, sizeof ( wchar_t ) );
    pText        = AllocBytes.New( numTextBytes );

    // Zero the ending bytes to act as a null terminator
    memcpy( pText, pResource, resSize );
    memset( pText + resSize, 0, sizeof ( wchar_t ) );

    // Test for BOM
    if ( numTextBytes > 2 )
    {
        if ( ( pText[ 0 ] == 0xFF ) && ( pText[ 1 ] == 0xFE )  )
        {
            *pStringData = reinterpret_cast<LPCWSTR>( pText + 2 );
        }
        else
        {
            pStringData->SetAnsi( reinterpret_cast<const char*>( pText ) );
        }
    }
    else
    {
        pStringData->SetAnsi( reinterpret_cast<const char*>( pText ) );
    }
}

//===============================================================================================//
//  Description:
//      Flush the application's log file
//
//  Parameters:
//      void
//
//  Returns:
//      void
//===============================================================================================//
void Application::LogFlush()
{
    try
    {
        // Access shared data, only wait for a short while
        m_Mutex.Lock();
        AutoUnlockMutex AutoUnlock( &m_Mutex );
        if ( IsLogging() )
        {
            m_Logger.Flush();
        }
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Register a handler for when the new operator fails
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Application::RegisterNewHandler()
{
    // Set the new handler
    #ifdef _MSC_VER

       _set_new_handler( PXSNewHandler );

    #elif __GNUC__

        std::set_new_handler( PXSNewHandler );

    #else

        #error Unsupported compiler

    #endif
}

//===============================================================================================//
//  Description:
//      Set the application's name
//
//  Parameters:
//      pszApplicationName - pointer to the application name
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetApplicationName( LPCWSTR pszApplicationName )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_ApplicationName = pszApplicationName;
}

//===============================================================================================//
//  Description:
//      Set the copyright notice
//
//  Parameters:
//      pszCopyrightNotice - pointer to the copyright notice
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetCopyrightNotice( LPCWSTR pszCopyrightNotice )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_CopyrightNotice = pszCopyrightNotice;
}

//===============================================================================================//
//  Description:
//      Set the gradient colours
//
//  Parameters:
//      gradientOne - the first gradient colour
//      gradientTwo - the second gradient colour
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetGradientColours( COLORREF gradientOne, COLORREF gradientTwo )
{
    m_crGradientOne = gradientOne;
    m_crGradientTwo = gradientTwo;
}

//===============================================================================================//
//  Description:
//      Set the handle of the application's main window.
//
//  Parameters:
//      hWndMainFrame - handle to application's main window
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetHwndMainFrame( HWND hWndMainFrame )
{
    m_hWndMainFrame = hWndMainFrame;
}

//===============================================================================================//
//  Description:
//      Set the directory path last accessed by the user using OPENFILENAME
//
//  Parameters:
//      LastAccessedDirectory - the directory path
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetLastAccessedDirectory(const String& LastAccessedDirectory)
{
    m_LastAccessedDirectory = LastAccessedDirectory;
}

//===============================================================================================//
//  Description:
//      Set the logging level
//
//  Parameters:
//      logLevel - defined logging level constant
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetLogLevel( DWORD logLevel )
{
    m_uLogLevel = logLevel;
}

//===============================================================================================//
//  Description:
//      Set the manufacturer's address
//
//  Parameters:
//      pszManufacturerAddress - the manufacturer's address
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetManufacturerAddress( LPCWSTR pszManufacturerAddress )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_ManufacturerAddress = pszManufacturerAddress;
}

//===============================================================================================//
//  Description:
//      Set the manufacturer name property
//
//  Parameters:
//      pszManufacturerName - pointer to the manufacturer name
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetManufacturerName( LPCWSTR pszManufacturerName )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_ManufacturerName = pszManufacturerName;
}

//===============================================================================================//
//  Description:
//      Set the product identifier property
//
//  Parameters:
//      pszProductID - pointer to the Product ID
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetProductID( LPCWSTR pszProductID )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_ProductID = pszProductID;
}

//===============================================================================================//
//  Description:
//      Set the registered organization name
//
//  Parameters:
//      pszRegisteredOrgName - the registered organization name
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetRegisteredOrgName( LPCWSTR pszRegisteredOrgName )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_RegisteredOrgName = pszRegisteredOrgName;
}

//===============================================================================================//
//  Description:
//      Set the registered owner name
//
//  Parameters:
//      pszRegisteredOwnerName - pointer to the name
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetRegisteredOwnerName( LPCWSTR pszRegisteredOwnerName )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_RegisteredOwnerName = pszRegisteredOwnerName;
}

//===============================================================================================//
//  Description:
//      Set the flag to indicate want any background tasks to stop
//
//  Parameters:
//      stopTask - if true indicates want worker tasks to stop
//
//  Remarks:
//      A worker thread can poll this property to determine if it needs
//      to stop.
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetStopBackgroundTasks( bool stopBackgroundTasks )
{
    m_bStopBackgroundTasks = stopBackgroundTasks;
}

//===============================================================================================//
//  Description:
//      Set the support email property
//
//  Parameters:
//      pszSupportEmail - pointer to the support email
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetSupportEmail( LPCWSTR pszSupportEmail )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_SupportEmail = pszSupportEmail;
}

//===============================================================================================//
//  Description:
//      Set the web site URL property
//
//  Parameters:
//      pszWebSiteURL - pointer to the web site URL property
//
//  Returns:
//      void
//===============================================================================================//
void Application::SetWebSiteURL( LPCWSTR pszWebSiteURL )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_WebSiteURL = pszWebSiteURL;
}

//===============================================================================================//
//  Description:
//      Start the logging to the specified file
//
//  Parameters:
//      LogFilePath      - pointer to string of the log path
//      createNew        - flag to set if want new log file
//      wantComputerName - flag to indicate want the Computer column
//
//  Returns:
//      void
//===============================================================================================//
void Application::StartLogger( const String& LogFilePath, bool createNew, bool wantComputerName )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_Logger.SetWantComputerName( wantComputerName );
    m_Logger.Start( LogFilePath, createNew );

    // If the caller has not set a level, use verbose
    if ( m_uLogLevel == PXS_LOG_LEVEL_NONE )
    {
        SetLogLevel( PXS_LOG_LEVEL_VERBOSE );
    }
    m_isAppLogging = TRUE;
}

//===============================================================================================//
//  Description:
//      Start the logger with the log messages going to a string array
//
//  Parameters:
//      wantComputerName - flag to indicate want the Computer column
//      wantExtraInfo    - indicates wants extra diagnostic messages
//
//  Returns:
//      void
//===============================================================================================//
void Application::StartLogger( bool wantComputerName, bool wantExtraInfo )
{
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_Logger.SetWantComputerName( wantComputerName );
    m_Logger.Start( wantExtraInfo );

    // If the caller has not set a level, use verbose
    if ( m_uLogLevel == PXS_LOG_LEVEL_NONE )
    {
        SetLogLevel( PXS_LOG_LEVEL_VERBOSE );
    }
    m_isAppLogging = TRUE;
}

//===============================================================================================//
//  Description:
//      Stop the logger
//
//  Parameters:
//      void
//
//  Remarks:
//      This method does not throw exceptions or show message boxes
//
//  Returns:
//      void
//===============================================================================================//
void Application::StopLogger()
{
    try
    {
        m_Mutex.Lock();
        AutoUnlockMutex AutoUnlock( &m_Mutex );
        m_Logger.Stop();
    }
    catch ( const Exception& )           // Ignore
    { }
    m_isAppLogging = FALSE;
}

//===============================================================================================//
//  Description:
//      Write a line to the application's log
//
//  Parameters:
//      severity      - defined constant of severity e.g. info, warning etc.
//      errorType     - type of error, e.g. system, network etc.
//      errorCode     - a numerical code of the error
//      translateError- flag to indicate want to translate the error code
//      pszMessage    - string message with optional substitution parameters
//      Insert1       - string insert parameter, any %%1 are replaced
//      Insert2       - string insert parameter, any %%2 are replaced
//
//  Remarks:
//      Does not throw.
//
//  Returns:
//      void
//===============================================================================================//
void Application::WriteToAppLog( DWORD severity,
                                 DWORD errorType,
                                 DWORD errorCode,
                                 bool translateError,
                                 LPCWSTR pszMessage, const String& Insert1, const String& Insert2 )
{
    bool entryWritten = false;

    if ( m_isAppLogging == FALSE )
    {
        return;     // Not logging
    }

    if ( severity > m_uLogLevel )
    {
        return;     // Nothing to do
    }

    // Catch all errors, do not want recursion in case this method throws
    try
    {
        m_Mutex.Lock();
        AutoUnlockMutex AutoUnlock( &m_Mutex );
        if ( m_Logger.IsStarted() )
        {
            m_Logger.WriteEntry( severity,
                                 errorType,
                                 errorCode,
                                 translateError, pszMessage, Insert1, Insert2 );
            entryWritten = true;
        }
    }
    catch ( const Exception& )
    { }     // Ignore

    // Tell the GUI a message has been logged. Note, want to post the message
    // after unlocking the mutex otherwise observed response delay at the UI
    // even though PostMessage is expected to return immediately
    if ( m_hWndMainFrame && entryWritten )
    {
        PostMessage( m_hWndMainFrame, PXS_APP_MSG_LOGGER_UPDATE, 0, 0 );
    }
}


///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the resource identifiers of the string tables corresponding to
//      the specified primary and sub language identifiers
//
//  Parameters:
//      primaryLanguage - the primary language id
//      subLanguage     - the sug language id
//      pPxsBaseLocaleID    - receives the resource id for the language
//                            strings of the PxsBase library corresponding
//                            to the primary and sub iddentifiers
//      pAppEnglishID       - receives the resource id for the English
//                            language strings of the application
//      pAppLocaleID        - receives the resource id for the language
//                            strings of the application corresponding
//                            to the primaray and sub ids
//  Remarks:
//      Must have already called SetAppResourceLanguageIDs
//      If there is no matching language resource then the output is 0
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetLanguageResourceIDs( WORD  primaryLanguage,
                                          WORD  subLanguage,
                                          WORD* pPxsBaseLocaleID,
                                          WORD* pAppEnglishID, WORD* pAppLocaleID )
{
    size_t i = 0, size;
    RESOURCE_LANGIDS ResourceLangIDs;

    if ( ( pPxsBaseLocaleID == nullptr ) ||
         ( pAppEnglishID    == nullptr ) ||
         ( pAppLocaleID     == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pPxsBaseLocaleID = 0;
    *pAppEnglishID    = 0;
    *pAppLocaleID     = 0;

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );

    ///////////////////////////////////////////////////////////////////////////
    // Language resource ID for the PXS base library

    switch( primaryLanguage )
    {
        default:
            break;

        case LANG_CHINESE:
            // Traditional Chinese script is used in Taiwan, Honk Kong,
            // Macau and Guangzhou. Simplified Chinese script is used in
            // mainland China and Singapore.
            if ( ( subLanguage == SUBLANG_CHINESE_TRADITIONAL ) ||
                 ( subLanguage == SUBLANG_CHINESE_HONGKONG    ) ||
                 ( subLanguage == SUBLANG_CHINESE_MACAU       )  )
            {
                *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_ZH_TW;
            }
            break;

        case LANG_CZECH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_CS;
            break;

        case LANG_DANISH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_DA;
            break;

        case LANG_DUTCH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_NL;
            break;

        case LANG_ENGLISH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_EN;
            break;

        case LANG_FINNISH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_FI;
            break;

        case LANG_FRENCH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_FR_FR;
            if ( subLanguage == SUBLANG_FRENCH_BELGIAN )
            {
                *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_FR_BE;
            }
            break;

        case LANG_GERMAN:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_DE;
            break;

        case LANG_GREEK:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_EL;
            break;

        case LANG_HEBREW:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_HE;
            break;

        case LANG_HUNGARIAN:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_HU;
            break;

        case LANG_INDONESIAN:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_ID;
            break;

        case LANG_ITALIAN:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_IT;
            break;

        case LANG_JAPANESE:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_JP;
            break;

        case LANG_KOREAN:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_KO;
            break;

        case LANG_POLISH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_PL;
            break;

        case LANG_PORTUGUESE:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_PT_PT;
            if ( subLanguage == SUBLANG_PORTUGUESE_BRAZILIAN )
            {
                *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_PT_BR;
            }
            break;

        case LANG_RUSSIAN:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_RU;
            break;

        case LANG_SERBIAN:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_SR;
            break;

        case LANG_SLOVAK:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_SK;
            break;

        case LANG_SPANISH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_ES;
            break;

        case LANG_THAI:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_TH;
            break;

        case LANG_TURKISH:
            *pPxsBaseLocaleID = IDR_PXSBASE_STRINGS_TR;
            break;
    }

    ///////////////////////////////////////////////////////////////////////////
    // The application's English language resource ID

    size = m_AppResourceLanguageIDs.GetSize();
    for ( i = 0; i < size; i++ )
    {
        ResourceLangIDs = m_AppResourceLanguageIDs.Get( i );
        if ( ResourceLangIDs.primaryLanguage == LANG_ENGLISH )
        {
            *pAppEnglishID = ResourceLangIDs.resourceID;
            break;
        }
    }

    ///////////////////////////////////////////////////////////////////////////
    // Language resource ID for the application

    // First attempt to match on primary and sub language
    size = m_AppResourceLanguageIDs.GetSize();
    for ( i = 0; i < size; i++ )
    {
        ResourceLangIDs = m_AppResourceLanguageIDs.Get( i );
        if ( ( ResourceLangIDs.primaryLanguage == primaryLanguage ) &&
             ( ResourceLangIDs.subLanguage     == subLanguage     )  )
        {
            *pAppLocaleID = ResourceLangIDs.resourceID;
            break;
        }
    }

    // If not found, attempt to match on primary alone
    if( *pAppLocaleID == 0 )
    {
        for ( i = 0; i < size; i++ )
        {
            ResourceLangIDs = m_AppResourceLanguageIDs.Get( i );
            if ( ResourceLangIDs.primaryLanguage == primaryLanguage )
            {
                *pAppLocaleID = ResourceLangIDs.resourceID;
                break;
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the string table corresponding to the specified resource id
//
//  Parameters:
//      resourceID   - resource id of the string table
//      pStringTable - receives the string table
//
//  Returns:
//      void
//===============================================================================================//
void Application::GetStringTable( WORD resourceID, TList< ID_TEXT >* pStringTable )
{
    size_t  i, numLines;
    String  StringData, Line;
    ID_TEXT Element;
    StringArray Lines;

    if ( pStringTable == nullptr )
    {
        throw ParameterException( L"pStringTable", __FUNCTION__ );
    }
    pStringTable->RemoveAll();

    if ( resourceID == 0 )
    {
        return;   // Nothing to do
    }

    LoadTextDataResource( resourceID, &StringData );
    StringData.ToArray( PXS_CHAR_CR, &Lines );
    numLines = Lines.GetSize();
    for ( i = 0; i < numLines; i++ )
    {
        Line = Lines.Get( i );
        if ( StringTableLineToIdString( Line, &Element ) )
        {
            pStringTable->Append( Element );
        }
    }
}

//===============================================================================================//
//  Description:
//      Merge the specified string tables with the output placed in Table2
//
//  Parameters:
//      pTable1 - String table
//      pTable2 - String table
//
//  Remarks:
//      Matches the strings on identifier, if a match is found, the string from
//      table1 replaces that in table2. If there is no match, then the string
//      from table1 is appended to table2.
//      The tables must be sorted by string identifer otherwise the resultant
//      list is unpredictable.
//
//  Returns:
//      void
//===============================================================================================//
void Application::MergeStringTables( TList< ID_TEXT >* pTable1, TList< ID_TEXT >* pTable2 )
{
    bool    bMatch = false;
    ID_TEXT Element;

    if ( ( pTable1 == nullptr ) || ( pTable2 == nullptr ) )
    {
        throw ParameterException( L"pTable1/Table2", __FUNCTION__ );
    }

    if ( pTable1->IsEmpty() )
    {
        return;     // Nothing to do
    }

    pTable1->Rewind();
    do
    {
        Element = pTable1->Get();
        if ( Element.id && Element.Text.c_str() )
        {
            bMatch = false;
            if ( pTable2->IsEmpty() == false )
            {
                pTable2->Rewind();
                do
                {
                    if ( Element.id == pTable2->Get().id )
                    {
                        bMatch = true;
                        break;
                    }
                } while ( pTable2->Advance() );
            }

            if ( bMatch )
            {
                pTable2->Set( Element );
            }
            else
            {
                pTable2->Append( Element );
            }
        }
    } while ( pTable1->Advance() );
}

//===============================================================================================//
//  Description:
//      Get the id and string from the specified string table and put into
//      the ID_TEXT structure
//
//  Parameters:
//      Line    - Line of text
//      pIdText - receives the data
//
//  Remarks:
//      Format is: id|Text
//
//  Returns:
//      true if the ID_TEXT structure has valid data otherwise false
//===============================================================================================//
bool Application::StringTableLineToIdString( const String& Line, ID_TEXT* pIdText )
{
    size_t idxPipe;
    String ID;

    if ( pIdText == nullptr )
    {
        throw ParameterException( L"pIdText", __FUNCTION__ );
    }

    if ( Line.IsEmpty() )
    {
        return false;
    }

    // Ignore comments
    if ( Line.StartsWith( L"!", true ) )
    {
        return false;
    }
    idxPipe = Line.IndexOf( '|', 0 );
    if ( idxPipe == PXS_MINUS_ONE )
    {
        return false;
    }
    Line.SubString( 0, idxPipe, &ID );
    Line.SubString( idxPipe + 1, PXS_MINUS_ONE, &pIdText->Text );
    if ( ID.IsEmpty() || pIdText->Text.IsEmpty() )
    {
        return false;
    }
    pIdText->id = wcstoul( ID.c_str(), nullptr, 10 );
    if ( pIdText->id == 0 )
    {
        return false;
    }

    return true;
}
