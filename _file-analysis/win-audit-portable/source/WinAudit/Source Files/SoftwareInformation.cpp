///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Software Information Implementation
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
#include "WinAudit/Header Files/SoftwareInformation.h"

// 2. C System Files
#include <shappmgr.h>
#include <ShlObj.h>
#include <wuapi.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoIUnknownRelease.h"
#include "PxsBase/Header Files/AutoSysFreeString.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"
#include "PxsBase/Header Files/TList.h"
#include "PxsBase/Header Files/Wmi.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
SoftwareInformation::SoftwareInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
SoftwareInformation::~SoftwareInformation()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

//
// Assignment operator - not allowed so no implementation
//

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get Active Setup registry settings
//
//  Parameters:
//      pRecords - array to receive the audit records
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetActiveSetupRecords( TArray< AuditRecord >* pRecords )
{
    DWORD     isInstalled = 0;
    size_t    i = 0, numKeys = 0;
    String    SubKey, Key, Name, VersionString, Installed, LocaleNo, LocaleYes;
    Registry  RegObject;
    Formatter Format;
    StringArray Keys;
    AuditRecord Record;
    LPCWSTR     STR_ACTIVE_SETUP = L"Software\\Microsoft\\Active Setup\\Installed Components";

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();
    PXSGetResourceString( PXS_IDS_1260_NO , &LocaleNo );
    PXSGetResourceString( PXS_IDS_1261_YES, &LocaleYes );

    // Scan registry for installed components
    RegObject.Connect( HKEY_LOCAL_MACHINE );
    RegObject.GetSubKeyList( STR_ACTIVE_SETUP, &Keys );
    numKeys = Keys.GetSize();
    for ( i = 0; i < numKeys; i++ )
    {
        try
        {
            // For each subkey the the values of items
            SubKey = Keys.Get( i );
            if ( SubKey.GetLength() > 0 )
            {
                // Get the subkeys for this service pack
                Key  = STR_ACTIVE_SETUP;
                Key += L"\\";
                Key += SubKey;

                // Get the display name first, if none try for the
                // default value
                RegObject.GetStringValue( Key.c_str(), L"DisplayName", &Name );
                if ( Name.IsEmpty() )
                {
                    RegObject.GetStringValue( Key.c_str(), nullptr, &Name );
                    Name.Trim();
                }

                // Make sure have a name
                if ( Name.GetLength() )
                {
                    // Make the record
                    Record.Reset( PXS_CATEGORY_ACTIVE_SETUP );
                    Record.Add( PXS_ACTIVE_SETUP_NAME, Name );

                    RegObject.GetStringValue( Key.c_str(), L"Version", &VersionString );
                    Record.Add( PXS_ACTIVE_SETUP_VERSION, VersionString );

                    // Sometimes the IsInstalled key is missing so will use an
                    // empty string. Normally it is stored as a DWORD but
                    // sometimes its binary data
                    Installed   = PXS_STRING_EMPTY;
                    isInstalled = DWORD_MAX;    // -1 signifies unknown
                    RegObject.GetDoubleWordValue( Key.c_str(),
                                                  L"IsInstalled", DWORD_MAX, &isInstalled );

                    // If got nothing, look for binary data
                    if ( isInstalled == DWORD_MAX )
                    {
                        RegObject.GetBinaryData(
                                        Key.c_str(),
                                        L"IsInstalled",
                                        reinterpret_cast<BYTE*>( &isInstalled ),
                                        sizeof ( isInstalled ) );
                    }

                    // Non-zero means not installed
                    if ( isInstalled && ( isInstalled != DWORD_MAX ) )
                    {
                        Installed = LocaleYes;
                    }
                    else
                    {
                        Installed = LocaleNo;
                    }
                    Record.Add( PXS_ACTIVE_SETUP_INSTALLED, Installed );
                    pRecords->Add( Record );
                }
            }
        }
        catch ( const Exception& e )
        {
            PXSLogException( L"Error getting ActiveX data.", e, __FUNCTION__ );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_ACTIVE_SETUP_NAME );
}

//===============================================================================================//
//  Description:
//      Get diagnostic data concerning software on the machine
//
//  Parameters:
//      pDiagnostics - string object to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetDiagnostics( String* pDiagnostics )
{
    String      Title, RegKey, ApplicationName, DataString, Value;
    Formatter   Format;
    const TYPE_UPDATE_DATA*    pUpdate;
    const TYPE_INSTALLED_DATA* pInstalled;

    if ( pDiagnostics == nullptr )
    {
        throw ParameterException( L"pDiagnostics", __FUNCTION__ );
    }
    *pDiagnostics = PXS_STRING_EMPTY;

    DataString.Allocate( 96 * 1024 );
    PXSGetApplicationName( &ApplicationName );
    Title  = L"Software Data by ";
    Title += ApplicationName;
    DataString += Title;
    DataString += PXS_STRING_CRLF;
    DataString.AppendChar( '=', Title.GetLength() );
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    //////////////////////////////
    // Installed from the registry

    TList<TYPE_INSTALLED_DATA> RegInstalled;

    DataString += PXS_STRING_CRLF;
    DataString += L"-----------------------------\r\n";
    DataString += L"Installed Programs - Registry\r\n";
    DataString += L"-----------------------------\r\n";
    DataString += PXS_STRING_CRLF;

    try
    {
        GetInstalledFromRegistry( &RegInstalled );
        RegInstalled.Rewind();
        if ( RegInstalled.IsEmpty() == false )
        {
            do
            {
                pInstalled = RegInstalled.GetPointer();
                DataString += L"Name            : ";
                DataString += pInstalled->szName;
                DataString += PXS_STRING_CRLF;

                DataString += L"Software Key    : ";
                DataString += pInstalled->szSoftwareKey;
                DataString += PXS_STRING_CRLF;

                DataString += L"Exe Path        : ";
                DataString += pInstalled->szExePath;
                DataString += PXS_STRING_CRLF;

                DataString += L"Install Date    : ";
                DataString += pInstalled->szInstallDate;
                DataString += PXS_STRING_CRLF;

                DataString += L"Install Location: ";
                DataString += pInstalled->szInstallLocation;
                DataString += PXS_STRING_CRLF;

                DataString += L"Install Source  : ";
                DataString += pInstalled->szInstallSource;
                DataString += PXS_STRING_CRLF;

                DataString += L"Language        : ";
                DataString += pInstalled->szLanguage;
                DataString += PXS_STRING_CRLF;

                DataString += L"Last Used       : ";
                DataString += pInstalled->szLastUsed;
                DataString += PXS_STRING_CRLF;

                DataString += L"Product ID      : ";
                DataString += pInstalled->szProductID;
                DataString += PXS_STRING_CRLF;

                DataString += L"Vendor          : ";
                DataString += pInstalled->szVendor;
                DataString += PXS_STRING_CRLF;

                Value = pInstalled->szVersion;
                DataString += L"Version         : ";
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
            } while ( RegInstalled.Advance() );
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }

    ///////////////////////////////////////
    // Installed from the Windows Installer

    TList<TYPE_INSTALLED_DATA> MsiInstalled;

    try
    {
        DataString += PXS_STRING_CRLF;
        DataString += L"--------------------------------------\r\n";
        DataString += L"Installed Programs - Windows Installer\r\n";
        DataString += L"--------------------------------------\r\n";
        DataString += PXS_STRING_CRLF;

        GetInstalledFromMsi( &MsiInstalled );
        MsiInstalled.Rewind();
        if ( MsiInstalled.IsEmpty() == false )
        {
            do
            {
                pInstalled = MsiInstalled.GetPointer();
                DataString += L"Name            : ";
                DataString += pInstalled->szName;
                DataString += PXS_STRING_CRLF;

                DataString += L"Software Key    : ";
                DataString += pInstalled->szSoftwareKey;
                DataString += PXS_STRING_CRLF;

                DataString += L"Assignment Type : ";
                DataString += pInstalled->szAssignmentType;
                DataString += PXS_STRING_CRLF;

                DataString += L"Install Date    : ";
                DataString += pInstalled->szInstallDate;
                DataString += PXS_STRING_CRLF;

                DataString += L"Install Location: ";
                DataString += pInstalled->szInstallLocation;
                DataString += PXS_STRING_CRLF;

                DataString += L"Install Source  : ";
                DataString += pInstalled->szInstallSource;
                DataString += PXS_STRING_CRLF;

                DataString += L"Install State   : ";
                DataString += pInstalled->szInstallState;
                DataString += PXS_STRING_CRLF;

                DataString += L"Language        : ";
                DataString += pInstalled->szLanguage;
                DataString += PXS_STRING_CRLF;

                DataString += L"Local Package   : ";
                DataString += pInstalled->szLocalPackage;
                DataString += PXS_STRING_CRLF;

                DataString += L"Package Code    : ";
                DataString += pInstalled->szPackageCode;
                DataString += PXS_STRING_CRLF;

                DataString += L"Package Name    : ";
                DataString += pInstalled->szPackageName;
                DataString += PXS_STRING_CRLF;

                DataString += L"Product ID      : ";
                DataString += pInstalled->szProductID;
                DataString += PXS_STRING_CRLF;

                DataString += L"Reg Company     : ";
                DataString += pInstalled->szRegCompany;
                DataString += PXS_STRING_CRLF;

                DataString += L"Reg Owner       : ";
                DataString += pInstalled->szRegOwner;
                DataString += PXS_STRING_CRLF;

                DataString += L"Vendor          : ";
                DataString += pInstalled->szVendor;
                DataString += PXS_STRING_CRLF;

                DataString += L"Version         : ";
                DataString += pInstalled->szVersion;
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
            } while ( MsiInstalled.Advance() );
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }

    ////////////////////////////
    // Updates from the registry

    TList<TYPE_UPDATE_DATA> RegUpdates;

    try
    {
        DataString += PXS_STRING_CRLF;
        DataString += L"------------------\r\n";
        DataString += L"Updates - Registry\r\n";
        DataString += L"------------------\r\n";
        DataString += PXS_STRING_CRLF;

        GetUpdatesFromRegistry( &RegUpdates );
        RegUpdates.Rewind();
        if ( RegUpdates.IsEmpty() == false )
        {
            do
            {
                pUpdate = RegUpdates.GetPointer();
                DataString += L"Update ID       : ";
                DataString += pUpdate->szUpdateID;
                DataString += PXS_STRING_CRLF;

                DataString += L"Installed On    : ";
                DataString += pUpdate->szInstalledOn;
                DataString += PXS_STRING_CRLF;

                DataString += L"Description     : ";
                DataString += pUpdate->szDescription;
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
            } while ( RegUpdates.Advance() );
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }

    //////////////////////////////////////////////////
    // Updates from Windows Management Instrumentation

    TList<TYPE_UPDATE_DATA> WmiUpdates;

    try
    {
        DataString += PXS_STRING_CRLF;
        DataString += L"-----------------------------------------------\r\n";
        DataString += L"Updates - Windows(R) Management Instrumentation\r\n";
        DataString += L"-----------------------------------------------\r\n";
        DataString += PXS_STRING_CRLF;

        GetUpdatesFromWmi( &WmiUpdates );
        WmiUpdates.Rewind();
        if ( WmiUpdates.IsEmpty() == false )
        {
            do
            {
                pUpdate = WmiUpdates.GetPointer();
                DataString += L"Update ID       : ";
                DataString += pUpdate->szUpdateID;
                DataString += PXS_STRING_CRLF;

                DataString += L"Installed On    : ";
                DataString += pUpdate->szInstalledOn;
                DataString += PXS_STRING_CRLF;

                DataString += L"Description     : ";
                DataString += pUpdate->szDescription;
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
            } while ( WmiUpdates.Advance() );
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }

    /////////////////////////////////
    // Updates from Windows Installer

    TList<TYPE_UPDATE_DATA> MsiUpdates;

    try
    {
        DataString += PXS_STRING_CRLF;
        DataString += L"------------------------------\r\n";
        DataString += L"Updates - Windows(R) Installer\r\n";
        DataString += L"------------------------------\r\n";
        DataString += PXS_STRING_CRLF;

        GetUpdatesFromMsi( &MsiUpdates );
        MsiUpdates.Rewind();
        if ( MsiUpdates.IsEmpty() == false )
        {
            do
            {
                pUpdate = MsiUpdates.GetPointer();
                DataString += L"Update ID       : ";
                DataString += pUpdate->szUpdateID;
                DataString += PXS_STRING_CRLF;

                DataString += L"Installed On    : ";
                DataString += pUpdate->szInstalledOn;
                DataString += PXS_STRING_CRLF;

                DataString += L"Description     : ";
                DataString += pUpdate->szDescription;
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
            } while ( MsiUpdates.Advance() );
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }

    ////////////////////////////////////
    // Updates from Windows Update Agent

    TList<TYPE_UPDATE_DATA> WuaUpdates;

    try
    {
        DataString += PXS_STRING_CRLF;
        DataString += L"---------------------------------\r\n";
        DataString += L"Updates - Windows(R) Update Agent\r\n";
        DataString += L"---------------------------------\r\n";
        DataString += PXS_STRING_CRLF;

        GetUpdatesFromWua( &WuaUpdates );
        WuaUpdates.Rewind();
        if ( WuaUpdates.IsEmpty() == false )
        {
            do
            {
                pUpdate = WuaUpdates.GetPointer();
                DataString += L"Update ID       : ";
                DataString += pUpdate->szUpdateID;
                DataString += PXS_STRING_CRLF;

                DataString += L"Installed On    : ";
                DataString += pUpdate->szInstalledOn;
                DataString += PXS_STRING_CRLF;

                DataString += L"Description     : ";
                DataString += pUpdate->szDescription;
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
            } while ( WuaUpdates.Advance() );
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }
    *pDiagnostics = DataString;
}

//===============================================================================================//
//  Description:
//      Get the software installed on the machine
//
//  Parameters:
//      pRecords - array to receive the audit records
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetInstalledSoftwareRecords( TArray< AuditRecord >* pRecords )
{
    String      Value;
    Formatter   Format;
    AuditRecord Record;
    TList<TYPE_INSTALLED_DATA> Installed;
    TList<TYPE_INSTALLED_DATA> InstalledReg;
    TList<TYPE_INSTALLED_DATA> InstalledMso;
    const TYPE_INSTALLED_DATA* pData = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetInstalledFromMsi( &Installed );
    GetInstalledFromRegistry( &InstalledReg );
    GetInstalledMsoFromRegistry( &InstalledMso );
    MergeInstalledSoftwareLists( &InstalledReg, &Installed );
    MergeInstalledSoftwareLists( &InstalledMso, &Installed );
    if ( Installed.IsEmpty() )
    {
        return;     // Nothing else to do
    }

    Installed.Rewind();
    do
    {
        pData = Installed.GetPointer();
        if ( pData && lstrlen( pData->szName ) )
        {
            Record.Reset( PXS_CATEGORY_INSTALLED_PROGS );

            Record.Add( PXS_INSTAL_PROGS_NAME         , pData->szName );
            Record.Add( PXS_INSTAL_PROGS_VENDOR       , pData->szVendor );
            Record.Add( PXS_INSTAL_PROGS_VERSION      , pData->szVersion );
            Record.Add( PXS_INSTAL_PROGS_PRODUCT_LANG , pData->szLanguage );
            Record.Add( PXS_INSTAL_PROGS_INSTALL_DATE , pData->szInstallDate);
            Record.Add( PXS_INSTAL_PROGS_INSTALL_LOC  , pData->szInstallLocation);
            Record.Add( PXS_INSTAL_PROGS_INSTALL_SRC  , pData->szInstallSource);
            Record.Add( PXS_INSTAL_PROGS_INSTALL_STATE, pData->szInstallState );
            Record.Add( PXS_INSTAL_PROGS_ASSIGN_TYPE  , pData->szAssignmentType );
            Record.Add( PXS_INSTAL_PROGS_PACKAGE_CODE , pData->szPackageCode );
            Record.Add( PXS_INSTAL_PROGS_PACKAGE_NAME , pData->szPackageName );
            Record.Add( PXS_INSTAL_PROGS_LOCAL_PACKAGE, pData->szLocalPackage );
            Record.Add( PXS_INSTAL_PROGS_PRODUCT_ID   , pData->szProductID );
            Record.Add( PXS_INSTAL_PROGS_REG_COMPANY  , pData->szRegCompany );
            Record.Add( PXS_INSTAL_PROGS_REG_OWNER    , pData->szRegOwner );

            // Add the number of times used only if have a last used
            // date or an exe path.
            Value = PXS_STRING_EMPTY;
            if ( lstrlen( pData->szLastUsed ) || lstrlen(pData->szExePath) )
            {
                Value = Format.Int32( pData->timesUsed );
            }
            Record.Add( PXS_INSTAL_PROGS_TIMES_USED , Value );
            Record.Add( PXS_INSTAL_PROGS_LAST_USED  , pData->szLastUsed );
            Record.Add( PXS_INSTAL_PROGS_EXE_PATH   , pData->szExePath );
            Record.Add( PXS_INSTAL_PROGS_EXE_VERSION, pData->szExeVersion );
            Record.Add( PXS_INSTAL_PROGS_EXE_DESCRIP, pData->szExeDescription );
            Record.Add( PXS_INSTAL_PROGS_SOFTWARE_ID, pData->szSoftwareKey);
            pRecords->Add( Record );
        }
    } while ( Installed.Advance() );
    PXSSortAuditRecords( pRecords, PXS_INSTAL_PROGS_NAME );
}

//===============================================================================================//
//  Description:
//      Get the software updates on the computer
//
//  Parameters:
//      pRecords - array to receive the audit records
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetSoftwareUpdateRecords( TArray< AuditRecord >* pRecords )
{
    AuditRecord Record;
    TList<TYPE_UPDATE_DATA> Updates, MsiUpdates, WmiUpdates, RegUpdates;
    const TYPE_UPDATE_DATA* pUpdate = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetUpdatesFromWua( &Updates );
    GetUpdatesFromMsi( &MsiUpdates );
    GetUpdatesFromWmi( &WmiUpdates );
    GetUpdatesFromRegistry( &RegUpdates );

    // Merge the lists, the update agent list as the base
    MergeSoftwareUpdatesLists( &MsiUpdates, &Updates );
    MergeSoftwareUpdatesLists( &WmiUpdates, &Updates );
    MergeSoftwareUpdatesLists( &RegUpdates, &Updates );
    if ( Updates.IsEmpty() )
    {
        return;     // Nothing else to do
    }

    // Make the records
    Updates.Rewind();
    do
    {
        pUpdate = Updates.GetPointer();
        if ( pUpdate && wcslen( pUpdate->szUpdateID ) )
        {
          Record.Reset( PXS_CATEGORY_SOFTWARE_UPDATES );
          Record.Add( PXS_SOFTWARE_UPDATES_UPDATE_ID  , pUpdate->szUpdateID );
          Record.Add( PXS_SOFTWARE_UPDATES_INSTALL_ON , pUpdate->szInstalledOn);
          Record.Add( PXS_SOFTWARE_UPDATES_DESCRIPTION, pUpdate->szDescription);
          pRecords->Add( Record );
        }
    } while ( Updates.Advance() );
    PXSSortAuditRecords( pRecords, PXS_SOFTWARE_UPDATES_UPDATE_ID );
}

//===============================================================================================//
//  Description:
//      Get programmes that run on startup
//
//  Parameters:
//      pRecords - array to receive the audit records
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetStartupProgramRecords( TArray< AuditRecord >* pRecords )
{
    size_t      i = 0, j = 0, numElements = 0, numFileNames = 0;
    String      Name, Value, DirectoryPath, FullPath, DotStar;
    Registry    RegObject;
    Directory   DirObject;
    Formatter   Format;
    StringArray FileNames;
    AuditRecord Record;
    TArray< NameValue > NameValues;

    struct _KEYS
    {
        HKEY    hKey;
        LPCWSTR pszKey;
        LPCWSTR pszSubKey;
    } Keys[] =
        { { HKEY_LOCAL_MACHINE,
            L"HKLM",
            L"Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce"    },
          { HKEY_LOCAL_MACHINE,
            L"HKLM",
            L"Software\\Microsoft\\Windows\\CurrentVersion\\RunOnceEx" },
          { HKEY_LOCAL_MACHINE,
            L"HKLM",
            L"Software\\Microsoft\\Windows\\CurrentVersion\\Run"      },
          { HKEY_CURRENT_USER,
            L"HKCU",
            L"Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce"  },
          { HKEY_CURRENT_USER,
            L"HKCU",
            L"Software\\Microsoft\\Windows\\CurrentVersion\\Run"      } };

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    /////////////////////////////
    // Enumerate registry entries

    for ( i = 0; i < ARRAYSIZE( Keys ); i++ )
    {
        try
        {
            RegObject.Connect( Keys[ i ].hKey );
            NameValues.RemoveAll();
            RegObject.GetNameValueList( Keys[ i ].pszSubKey, &NameValues );
            numElements = NameValues.GetSize();
            for ( j = 0; j < numElements; j++ )
            {
                Name  = NameValues.Get( j ).GetName();
                Value = NameValues.Get( j ).GetValue();

                Name.Trim();
                if ( Name.GetLength() )
                {
                    // Make the full registry path
                    FullPath  = Keys[ i ].pszKey;
                    FullPath += PXS_PATH_SEPARATOR;
                    FullPath += Keys[ i ].pszSubKey;

                    Record.Reset( PXS_CATEGORY_STARTUP );
                    Record.Add( PXS_STARTUP_PROGRAM_NAME, Name );
                    Record.Add( PXS_STARTUP_SETTINGS_FOLDER, FullPath );
                    Record.Add( PXS_STARTUP_COMMAND, Value );
                    pRecords->Add( Record );
                }
            }
        }
        catch ( const Exception& e )
        {
            PXSLogException( L"Error getting registry start up data.", e, __FUNCTION__ );
        }
    }

    ///////////////////
    // Startup folders

    int nFolder[ 2 ] = { CSIDL_COMMON_STARTUP,  // Common start up for all users
                         CSIDL_STARTUP   };     // User's start up folder
    try
    {
        DotStar = L".*";
        for ( i = 0; i < ARRAYSIZE( nFolder ); i++ )
        {
            // Get list of files in directory
            DirectoryPath = PXS_STRING_EMPTY;
            DirObject.GetSpecialDirectory( nFolder[ i ], &DirectoryPath );
            if ( DirectoryPath.GetLength() )
            {
                DirObject.ListFiles( DirectoryPath, DotStar, &FileNames );
                numFileNames = FileNames.GetSize();
                for ( j = 0; j < numFileNames; j++ )
                {
                    Name = FileNames.Get( j );
                    Name.Trim();
                    if ( Name.GetLength() )
                    {
                        Record.Reset( PXS_CATEGORY_STARTUP );
                        Record.Add( PXS_STARTUP_PROGRAM_NAME, Name );
                        Record.Add( PXS_STARTUP_SETTINGS_FOLDER, DirectoryPath);
                        Record.Add( PXS_STARTUP_COMMAND, PXS_STRING_EMPTY );
                        pRecords->Add( Record );
                    }
                }
            }
            else
            {
                PXSLogSysError( ERROR_BAD_PATHNAME, L"Special directory not found." );
            }
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error start up programs.",
                         e, __FUNCTION__ );
    }
    PXSSortAuditRecords( pRecords, PXS_STARTUP_PROGRAM_NAME );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add information about the software specified by the registry key
//      name to the list of installed software
//
//  Parameters:
//      pRegObject      - registry object, must be connected
//      pszUnistallPath - the un-install path in the registry
//      KeyName         - the registry key name
//      pInstalled      - list to recieve installed software data
//
//  Returns:
//      true if found the software key, else false
//===============================================================================================//
bool SoftwareInformation::AddSoftwareInfoFromRegistry( Registry* pRegObject,
                                                       LPCWSTR pszUnistallPath,
                                                       const String& KeyName,
                                                       TList<TYPE_INSTALLED_DATA>* pInstalled )
{
    File   FileObject;
    DWORD  language = DWORD_MAX;      // -1 = not found
    String RegKey, Value, Name, LastUsed, ExePath;
    String ExeManufacturer, ExeVersion, ExeDescription;
    Formatter   Format;
    FileVersion FileVer;
    TYPE_INSTALLED_DATA Data;

    if ( pRegObject == nullptr )
    {
        throw ParameterException( L"pRegObject", __FUNCTION__ );
    }

    if ( pszUnistallPath == nullptr )
    {
        return false;
    }

    if ( pInstalled == nullptr )
    {
        throw ParameterException( L"pInstalled", __FUNCTION__ );
    }
    memset( &Data, 0, sizeof ( Data ) );

    // Ignore KB numbers
    if ( IsQorKbNumber( KeyName.c_str() ) )
    {
        return false;
    }

    // Build the key for this entry
    RegKey  = pszUnistallPath;
    RegKey += KeyName;

    // Try DisplayName then QuietDisplayName
    pRegObject->GetStringValue( RegKey.c_str(), L"DisplayName", &Name );
    Name.Trim();
    if ( Name.IsEmpty() )
    {
        pRegObject->GetStringValue( RegKey.c_str(), L"QuietDisplayName", &Name );
    }
    if ( Name.IsEmpty() )
    {
       return false;
    }

    // Sometimes items are duplicated in this section of the
    // registry. A side-effect of this check may be to filter
    // out different versions of the same software.
    if ( FindInInstalledData( Name, pInstalled ) )
    {
        return false;
    }

    // ID Number, use the registry Key
    StringCchCopy( Data.szSoftwareKey,
                   ARRAYSIZE( Data.szSoftwareKey ), KeyName.c_str() );

    // Name
    StringCchCopy( Data.szName, ARRAYSIZE( Data.szName ), Name.c_str() );

    Value = PXS_STRING_EMPTY;
    pRegObject->GetStringValue( RegKey.c_str(), L"Publisher", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( Data.szVendor, ARRAYSIZE( Data.szVendor ), Value.c_str() );
    }

    Value = PXS_STRING_EMPTY;
    pRegObject->GetStringValue( RegKey.c_str(), L"DisplayVersion", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( Data.szVersion, ARRAYSIZE( Data.szVersion ), Value.c_str() );
    }

    Value = PXS_STRING_EMPTY;
    pRegObject->GetDoubleWordValue( RegKey.c_str(), L"Language", DWORD_MAX, &language );
    if ( language != DWORD_MAX )
    {
        Value = Format.LanguageIdToName( LOWORD( language ) );
    }
    else
    {
        // Sometimes stored as as string
        pRegObject->GetStringValue( RegKey.c_str(), L"Language", &Value );
    }
    if ( Value.c_str() )
    {
        StringCchCopy( Data.szLanguage, ARRAYSIZE( Data.szLanguage ), Value.c_str() );
    }

    Value = PXS_STRING_EMPTY;
    pRegObject->GetStringValue( RegKey.c_str(), L"InstallDate", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( Data.szInstallDate, ARRAYSIZE( Data.szInstallDate ), Value.c_str() );
    }

    Value = PXS_STRING_EMPTY;
    pRegObject->GetStringValue( RegKey.c_str(), L"InstallLocation", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( Data.szInstallLocation, ARRAYSIZE(Data.szInstallLocation), Value.c_str() );
    }

    Value = PXS_STRING_EMPTY;
    pRegObject->GetStringValue( RegKey.c_str(), L"InstallSource", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( Data.szInstallSource, ARRAYSIZE( Data.szInstallSource ), Value.c_str() );
    }

    Value = PXS_STRING_EMPTY;
    pRegObject->GetStringValue( RegKey.c_str(), L"ProductID", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( Data.szProductID, ARRAYSIZE( Data.szProductID ), Value.c_str() );
    }

    // Get Add/Remove Programs and version data
    GetArpCacheData( KeyName, &Data.timesUsed, &LastUsed, &ExePath );
    if ( LastUsed.c_str() )
    {
        StringCchCopy( Data.szLastUsed, ARRAYSIZE( Data.szLastUsed ), LastUsed.c_str() );
    }

    if ( ExePath.c_str() )
    {
        StringCchCopy( Data.szExePath, ARRAYSIZE( Data.szExePath ), ExePath.c_str() );
    }

    // Get version information
    if ( FileObject.Exists( ExePath ) )
    {
        FileVer.GetVersion( ExePath, &ExeManufacturer, &ExeVersion, &ExeDescription );
        if ( ExeVersion.c_str() )
        {
            StringCchCopy( Data.szExeVersion, ARRAYSIZE( Data.szExeVersion ), ExeVersion.c_str() );
        }

        if ( ExeDescription.c_str() )
        {
            StringCchCopy( Data.szExeDescription,
                           ARRAYSIZE( Data.szExeDescription ),  ExeDescription.c_str() );
        }
    }

    pInstalled->Append( Data );

    return true;
}

//===============================================================================================//
//  Description:
//      Copy data that is present in Update1 to empty fields of Update2
//
//  Parameters:
//      pUpdate1 - pointer to an update structure that may have additional data
//      pUpdate2 - pointer to an update structure to backfill
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::BackFillMissingData( const TYPE_UPDATE_DATA* pUpdate1,
                                               TYPE_UPDATE_DATA* pUpdate2 )
{
    if ( ( pUpdate1 == nullptr ) || ( pUpdate2 == nullptr ) )
    {
        return;     // Nothing to do
    }

    if ( pUpdate2->szInstalledOn[ 0 ] == PXS_CHAR_NULL )
    {
        StringCchCopy( pUpdate2->szInstalledOn,
                       ARRAYSIZE( pUpdate2->szInstalledOn ), pUpdate1->szInstalledOn );
    }

    if ( pUpdate2->szDescription[ 0 ] == PXS_CHAR_NULL )
    {
        StringCchCopy( pUpdate2->szDescription,
                       ARRAYSIZE( pUpdate2->szDescription ), pUpdate1->szDescription );
    }
}

//===============================================================================================//
//  Description:
//      Find the specified software name in the list, case insensitive
//
//  Parameters:
//      Name       - the name to find
//      pInstalled - the list to search
//
//  Returns:
//      true if found, otherwise false
//===============================================================================================//
bool SoftwareInformation::FindInInstalledData( const String& Name,
                                               TList<TYPE_INSTALLED_DATA>* pInstalled )
{
    bool match = false;
    const TYPE_INSTALLED_DATA* pData = nullptr;

    if ( ( Name.IsEmpty()        ) ||
         ( pInstalled == nullptr ) ||
         ( pInstalled->IsEmpty() )  )
    {
        return false;
    }

    pInstalled->Rewind();
    do
    {
        pData = pInstalled->GetPointer();
        if ( Name.CompareI( pData->szName ) == 0 )
        {
            match = true;
        }
    } while ( ( match == false ) && ( pInstalled->Advance() ) );

    return match;
}

//===============================================================================================//
//  Description:
//      Find the specified data item in the specified list. If found the list
//      is placed on that element otherwise on return, it is positioned on the
//      tail.
//
//  Parameters:
//      pData      - the data item to search for
//      pInstalled - the list to search in
//      pElement   - receives the element if found
//
//  Returns:
//      true if the element was found otherwise false
//===============================================================================================//
bool SoftwareInformation::FindInstalledSoftwareElement( const TYPE_INSTALLED_DATA* pData,
                                                        TList<TYPE_INSTALLED_DATA>* pInstalled,
                                                        TYPE_INSTALLED_DATA* pElement )
{
    bool found = false;
    const TYPE_INSTALLED_DATA* pSoftware = nullptr;

    if ( pElement == nullptr )
    {
        throw ParameterException( L"pElement", __FUNCTION__ );
    }

    if ( ( pData == nullptr      ) ||
         ( pInstalled == nullptr ) ||
         ( pInstalled->IsEmpty() )  )
    {
        return false;
    }

    pInstalled->Rewind();
    do
    {
        // Match on key then try name
        pSoftware = pInstalled->GetPointer();
        if ( ( lstrlen(  pData->szSoftwareKey ) != 0 ) &&
             ( lstrcmpi( pData->szSoftwareKey, pSoftware->szSoftwareKey) ) == 0 )
        {
            found = true;
        }
        else if ( ( lstrlen(  pData->szName ) != 0 ) &&
                  ( lstrcmpi( pData->szName, pSoftware->szName ) == 0 ) )
        {
            found = true;
        }
        else if ( ( lstrlen(  pData->szSoftwareKey ) != 0 ) &&
                  ( lstrcmpi( pData->szSoftwareKey, pSoftware->szName ) == 0 ) )
        {
            found = true;
        }

        if ( found )
        {
            memcpy( pElement, pSoftware, sizeof ( TYPE_INSTALLED_DATA ) );
        }
    }
    while ( ( found == false ) && pInstalled->Advance() );

    return found;
}

//===============================================================================================//
//  Description:
//      Get the Add/Remove cached data for the specified registry key
//
//  Parameters:
//      KeyName    - software registry key name to the ARP cache
//      pTimesUsed - receives the number of times the software has been used
//      pLastUsed  - the last time the software was used, ISO format in the
//                   machine's local time zone
//      pExePath   - receives the path to the executable
//
//  Remarks:
//      The data in the Add/Remove programs cache seems to be of a low
//      quality. Use cautiously.
//      As data is often missing will return an error rather than throw
//
//  Returns:
//      System error code
//===============================================================================================//
DWORD SoftwareInformation::GetArpCacheData( const String& KeyName,
                                            int* pTimesUsed, String* pLastUsed, String* pExePath )
{
    DWORD     errorCode = 0;
    String    ArpRegKey;
    Registry  RegObject;
    Formatter Format;
    TYPE_SLOWINFOCACHE SlowInfoCache;
    WindowsInformation WindowsInfo;
    LPCWSTR ARP_CACHE_KEY = L"Software\\Microsoft\\Windows\\CurrentVersion"
                            L"\\App Management\\ARPCache\\";

    if ( ( pTimesUsed == nullptr ) ||
         ( pLastUsed  == nullptr ) ||
         ( pExePath   == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pTimesUsed = 0;
    *pLastUsed  = PXS_STRING_EMPTY;
    *pLastUsed  = PXS_STRING_EMPTY;

    if ( KeyName.IsEmpty() )
    {
         return ERROR_INVALID_PARAMETER;
    }

    // Connect to the registry for the Add/Remove Programs information. On NT6
    // the data is at the user level, prior to that it was at the machine level
    if ( WindowsInfo.GetMajorVersion() >= 6 )
    {
        RegObject.Connect( HKEY_CURRENT_USER );
    }
    else
    {
        RegObject.Connect( HKEY_LOCAL_MACHINE );
    }

    // Read the cache directly into the structure
    ArpRegKey  = ARP_CACHE_KEY;
    ArpRegKey += KeyName;
    memset( &SlowInfoCache, 0, sizeof ( SlowInfoCache ) );
    errorCode = RegObject.GetBinaryData( ArpRegKey.c_str(),
                                         L"SlowInfoCache",
                                         reinterpret_cast<BYTE*>( &SlowInfoCache ),
                                         sizeof ( SlowInfoCache ) );
    if ( errorCode != ERROR_SUCCESS )
    {
        return errorCode;
    }

    // Times used, sometimes iTimesUsed = -1
    if ( SlowInfoCache.iTimesUsed >= 0 )
    {
        *pTimesUsed = SlowInfoCache.iTimesUsed;
    }

    // Get the date on which the programme was last used. No documentation if
    // this is a UTC or local time but time zone testing indicates its local.
    if ( ( SlowInfoCache.ftLastUsed.dwHighDateTime > 0 ) &&
         ( SlowInfoCache.ftLastUsed.dwLowDateTime  > 0 ) &&
         ( SlowInfoCache.ftLastUsed.dwHighDateTime < DWORD_MAX ) &&
         ( SlowInfoCache.ftLastUsed.dwLowDateTime  < DWORD_MAX )  )
    {
       *pLastUsed = Format.FileTimeToLocalTimeIso( SlowInfoCache.ftLastUsed );
    }

    // The exe path, NB wImage is not terminated
    SlowInfoCache.wImage[ ARRAYSIZE( SlowInfoCache.wImage ) - 1 ] = 0x00;
    *pExePath = SlowInfoCache.wImage;
    pExePath->Trim();

    return errorCode;
}

//===============================================================================================//
//  Description:
//      Get the software that was installed by the Windows Installer
//
//  Parameters:
//      pInstalled - a list to populate with the discovered software
//
//  Remarks:
//      Windows Installer data is held at:
//      HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer
//
//      This is the same information provided by WMI's Win32_Product class
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetInstalledFromMsi( TList<TYPE_INSTALLED_DATA>* pInstalled )
{
    int     temp   = 0;
    WORD    langID = 0;
    UINT    resultProducts = 0, result = 0;
    DWORD   cchValueBuf  = 0;
    wchar_t szProduct [ 64 ]  = { 0 };       // Enough for a GUID
    wchar_t szValueBuf[ 256 ] = { 0 };       // Data buffer
    String Value, Insert1;
    Formatter    Format;
    INSTALLSTATE state;
    TYPE_INSTALLED_DATA  Data;

    if ( pInstalled == nullptr )
    {
        throw ParameterException( L"pInstalled", __FUNCTION__ );
    }
    pInstalled->RemoveAll();

    DWORD iProductIndex = 0;     // Must be zero on first call
    resultProducts = MsiEnumProducts( iProductIndex, szProduct );
    while ( resultProducts == ERROR_SUCCESS )
    {
        szProduct[ ARRAYSIZE( szProduct ) - 1 ] = PXS_CHAR_NULL;

        // INSTALLPROPERTY_INSTALLEDPRODUCTNAME, product name. This is a
        // required property
        memset( szValueBuf, 0, sizeof ( szValueBuf ) );
        cchValueBuf = ARRAYSIZE( szValueBuf );
        result    = MsiGetProductInfo( szProduct,
                                       L"InstalledProductName", szValueBuf, &cchValueBuf );
        szValueBuf[ ARRAYSIZE( szValueBuf ) - 1 ] = PXS_CHAR_NULL;
        if ( result == ERROR_SUCCESS )
        {
            // Add a new item to the list
            memset( &Data, 0, sizeof ( Data ) );

            // Name
            StringCchCopy( Data.szName, ARRAYSIZE( Data.szName ), szValueBuf );

            // ID Number
            StringCchCopy( Data.szSoftwareKey, ARRAYSIZE( Data.szSoftwareKey ), szProduct );

            // INSTALLPROPERTY_PUBLISHER, Publisher
            cchValueBuf = ARRAYSIZE( Data.szVendor );
            MsiGetProductInfo( szProduct, L"Publisher", Data.szVendor, &cchValueBuf );
            Data.szVendor[ ARRAYSIZE( Data.szVendor ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_VERSIONSTRING, Version
            cchValueBuf = ARRAYSIZE( Data.szVersion );
            MsiGetProductInfo( szProduct, L"VersionString", Data.szVersion, &cchValueBuf );
             Data.szVersion[ ARRAYSIZE( Data.szVersion ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_LANGUAGE, Language
            memset( szValueBuf, 0, sizeof ( szValueBuf ) );
            cchValueBuf = ARRAYSIZE( szValueBuf );
            if ( ERROR_SUCCESS == MsiGetProductInfo( szProduct,
                                                     L"Language", szValueBuf, &cchValueBuf ) )
            {
                // Convert to the language id to a name, 0 = LANG_NEUTRAL
                szValueBuf[ ARRAYSIZE( szValueBuf ) - 1 ] = PXS_CHAR_NULL;
                Value  = PXS_STRING_EMPTY;
                langID = 0;
                temp   = _ttoi( szValueBuf );
                if ( temp > 0 )
                {
                    langID = PXSCastInt32ToUInt16( temp );
                }
                Value = Format.LanguageIdToName( langID );
                StringCchCopy( Data.szLanguage, ARRAYSIZE( Data.szLanguage ), Value.c_str() );
            }

            // INSTALLPROPERTY_INSTALLDATE, InstallDate
            cchValueBuf = ARRAYSIZE( Data.szInstallDate );
            MsiGetProductInfo( szProduct, L"InstallDate", Data.szInstallDate, &cchValueBuf );
            Data.szInstallDate[ ARRAYSIZE( Data.szInstallDate ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_INSTALLLOCATION, InstallLocation
            cchValueBuf = ARRAYSIZE( Data.szInstallLocation );
            MsiGetProductInfo( szProduct,
                               L"InstallLocation", Data.szInstallLocation, &cchValueBuf );
            Data.szInstallLocation[ ARRAYSIZE( Data.szInstallLocation ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_INSTALLSOURCE, InstallSource
            cchValueBuf = ARRAYSIZE( Data.szInstallSource );
            MsiGetProductInfo( szProduct, L"InstallSource", Data.szInstallSource, &cchValueBuf );
            Data.szInstallSource[ ARRAYSIZE( Data.szInstallSource ) - 1 ] = PXS_CHAR_NULL;

            // InstallState
            Value = PXS_STRING_EMPTY;
            state = MsiQueryProductState( szProduct );
            TranslateMsiInstallState( state, &Value );
            StringCchCopy( Data.szInstallState, ARRAYSIZE( Data.szInstallState ), Value.c_str());

            // INSTALLPROPERTY_ASSIGNMENTTYPE, AssignmentType
            memset( szValueBuf, 0, sizeof ( szValueBuf ) );
            cchValueBuf = ARRAYSIZE( szValueBuf );
            if ( ERROR_SUCCESS == MsiGetProductInfo( szProduct,
                                                     L"AssignmentType", szValueBuf, &cchValueBuf ) )
            {
                szValueBuf[ ARRAYSIZE( szValueBuf ) - 1 ] = PXS_CHAR_NULL;
                if ( '0' == szValueBuf[ 0 ] )
                {
                    StringCchCopy( Data.szAssignmentType,
                                   ARRAYSIZE( Data.szAssignmentType ), L"Per User" );
                }
                else if ( '1' == szValueBuf[ 0 ] )
                {
                    StringCchCopy( Data.szAssignmentType,
                                   ARRAYSIZE( Data.szAssignmentType ), L"Per Machine" );
                }
            }

            // INSTALLPROPERTY_PACKAGECODE, PackageCode
            cchValueBuf = ARRAYSIZE( Data.szPackageCode);
            MsiGetProductInfo( szProduct, L"PackageCode", Data.szPackageCode, &cchValueBuf );
            Data.szPackageCode[ ARRAYSIZE( Data.szPackageCode ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_PACKAGENAME, PackageName
            cchValueBuf = ARRAYSIZE( Data.szPackageName );
            MsiGetProductInfo( szProduct, L"PackageName", Data.szPackageName, &cchValueBuf );
            Data.szPackageName[ ARRAYSIZE( Data.szPackageName ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_LOCALPACKAGE, PackageCache
            cchValueBuf = ARRAYSIZE( Data.szLocalPackage );
            MsiGetProductInfo( szProduct, L"LocalPackage", Data.szLocalPackage, &cchValueBuf );
            Data.szLocalPackage[ ARRAYSIZE( Data.szLocalPackage ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_PRODUCTID, ProductID
            cchValueBuf = ARRAYSIZE( Data.szProductID );
            MsiGetProductInfo( szProduct, L"ProductID", Data.szProductID, &cchValueBuf );
            Data.szProductID[ ARRAYSIZE( Data.szProductID ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_REGCOMPANY, RegCompany
            cchValueBuf = ARRAYSIZE( Data.szRegCompany );
            MsiGetProductInfo( szProduct, L"RegCompany", Data.szRegCompany, &cchValueBuf );
            Data.szRegCompany[ ARRAYSIZE( Data.szRegCompany ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_REGOWNER, RegOwner
            cchValueBuf = ARRAYSIZE( Data.szRegOwner );
            MsiGetProductInfo( szProduct, L"RegOwner", Data.szRegOwner, &cchValueBuf );
            Data.szRegOwner[ ARRAYSIZE( Data.szRegOwner ) - 1 ] = PXS_CHAR_NULL;

            pInstalled->Append( Data );
        }
        else
        {
            Insert1 = szProduct;
            PXSLogSysError1( result, L"MsiGetProductInfo failed for GUID='%%1'.", Insert1 );
        }

        // Next product
        iProductIndex++;
        memset( szProduct, 0, sizeof ( szProduct ) );
        resultProducts = MsiEnumProducts( iProductIndex, szProduct );
        szProduct[ ARRAYSIZE( szProduct ) - 1 ] = PXS_CHAR_NULL;
    };
    pInstalled->Rewind();
}

//===============================================================================================//
//  Description:
//      Get the installed software as found in the registry
//
//  Parameters:
//      pInstalled - a list to populate with the discovered software
//
//  Remarks:
//      Want to enumerate both the 32- and 64-bit views. However, the software uninstall key
//      is not a "shared key" so is not handled by the registry re-director/reflector. See
//      Article ID: KB896459. Further, MSDN article "Accessing an Alternate Registry View"
//      says applications should not use Wow6432Node key directly. As documented, will
//      use the KEY_WOW64_64KEY and KEY_WOW64_32KEY access rights. This means that on 32-bit
//      systems will read the same information twice but as this method filters duplicates
//      that will be handled.
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetInstalledFromRegistry( TList<TYPE_INSTALLED_DATA>* pInstalled )
{
    struct _REG_ACCESS
    {
        HKEY    hKey;
        DWORD   samAccess;
        LPCWSTR lpwzView;
    } REG_ACCESS[] = { { HKEY_LOCAL_MACHINE, KEY_READ | KEY_WOW64_64KEY, L"WOW64_64:HKLM\\" },
                       { HKEY_LOCAL_MACHINE, KEY_READ | KEY_WOW64_32KEY, L"WOW64_32:HKLM\\" },
                       { HKEY_CURRENT_USER , KEY_READ | KEY_WOW64_64KEY, L"WOW64_64:HKCU\\" },
                       { HKEY_CURRENT_USER , KEY_READ | KEY_WOW64_32KEY, L"WOW64_32:HKCU\\" } };

    size_t      i = 0, j = 0, numKeys = 0;
    String      KeyName, UninstallKey;
    Registry    RegObject;
    StringArray KeyNames;
    LPCWSTR UNINSTALL_KEY = L"Software\\Microsoft\\Windows\\CurrentVersion\\UnInstall\\";

    if ( pInstalled == nullptr )
    {
        throw ParameterException( L"pInstalled", __FUNCTION__ );
    }
    pInstalled->RemoveAll();

    // One pass for the machine then another for the user installed software
    for ( i = 0; i < ARRAYSIZE( REG_ACCESS ); i++ )
    {
         RegObject.Connect2( REG_ACCESS[ i ].hKey, REG_ACCESS[ i ].samAccess );
         UninstallKey  = REG_ACCESS[ i ].lpwzView;
         UninstallKey += UNINSTALL_KEY;

        // Trap individual errors so can continue
        try
        {
            RegObject.GetSubKeyList( UNINSTALL_KEY, &KeyNames );
            numKeys = KeyNames.GetSize();
            for ( j = 0; j < numKeys; j++ )
            {
                KeyName = KeyNames.Get( j );
                PXSLogAppInfo2( L"Found registry software key '%%1%%2'.", UninstallKey, KeyName );
                AddSoftwareInfoFromRegistry( &RegObject,
                                             UNINSTALL_KEY, KeyName, pInstalled );
            }
        }
        catch ( const Exception& e )
        {
            PXSLogException( L"Error getting registry software data.", e, __FUNCTION__ );
        }
        RegObject.Disconnect();
    }
    pInstalled->Rewind();
}

//===============================================================================================//
//  Description:
//      Get Office 2013 (v15.0) as found in the registry
//
//  Parameters:
//      pInstalled - a list to populate with the discovered software
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetInstalledMsoFromRegistry( TList<TYPE_INSTALLED_DATA>* pInstalled )
{
    struct _REG_ACCESS
    {
        HKEY    hKey;
        DWORD   samAccess;
        LPCWSTR lpwzView;
    } REG_ACCESS[] = { { HKEY_LOCAL_MACHINE, KEY_READ | KEY_WOW64_64KEY, L"WOW64_64:HKLM\\" },
                       { HKEY_LOCAL_MACHINE, KEY_READ | KEY_WOW64_32KEY, L"WOW64_32:HKLM\\" },
                       { HKEY_CURRENT_USER , KEY_READ | KEY_WOW64_64KEY, L"WOW64_64:HKCU\\" },
                       { HKEY_CURRENT_USER , KEY_READ | KEY_WOW64_32KEY, L"WOW64_32:HKCU\\" } };

    size_t      i = 0, j = 0;
    String      RegKey, Value;
    Registry    RegObject;
    TYPE_INSTALLED_DATA Data;
    LPCWSTR MSO365_CONFIG    = L"SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\Configuration";
    LPCWSTR MSO365_CULTURE   = L"SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\ProductReleaseIDs\\"
                               L"Active\\culture";

    if ( pInstalled == nullptr )
    {
        throw ParameterException( L"pInstalled", __FUNCTION__ );
    }
    pInstalled->RemoveAll();

    //////////////////////////////////////
    // Look for Office 365 ( 2013 = 15.0)

    for ( i = 0; i < ARRAYSIZE( REG_ACCESS ); i++ )
    {
        // Trap individual errors so can continue
        try
        {
            RegObject.Connect2( REG_ACCESS[ i ].hKey, REG_ACCESS[ i ].samAccess );
            //RegKey  = REG_ACCESS[ i ].lpwzView;
            RegKey  = MSO365_CONFIG;
            Value   = PXS_STRING_EMPTY;
            RegObject.GetStringValue( RegKey.c_str(), L"ProductReleaseIds", &Value );
            if ( Value.StartsWith( L"O365", false ))
            {
                memset( &Data, 0, sizeof ( Data ) );

                // Name
                StringCchCopy( Data.szName, ARRAYSIZE( Data.szName ), L"Microsoft Office 365" );

                // Publisher
                StringCchCopy( Data.szVendor, ARRAYSIZE( Data.szVendor ), L"Microsoft Corporation" );

                // Version
                Value = PXS_STRING_EMPTY;
                RegObject.GetStringValue( MSO365_CULTURE, L"x-none", &Value );
                if ( Value.c_str() )
                {
                    StringCchCopy( Data.szVersion, ARRAYSIZE( Data.szVersion ), Value.c_str() );
                }

                // Product Language
                Value = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"ClientCulture", &Value );
                if ( Value.c_str() )
                {
                    StringCchCopy( Data.szLanguage, ARRAYSIZE( Data.szLanguage ), Value.c_str() );
                }

                // Install Location
                Value = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"InstallationPath", &Value );
                if ( Value.c_str() )
                {
                    StringCchCopy( Data.szInstallLocation,
                                   ARRAYSIZE(Data.szInstallLocation), Value.c_str() );
                }

                // Install Source
                Value = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"CDNBaseUrl", &Value );
                if ( Value.c_str() )
                {
                    StringCchCopy( Data.szInstallSource,
                                   ARRAYSIZE( Data.szInstallSource ), Value.c_str() );
                }
                pInstalled->Append( Data );
            }
            RegObject.Disconnect();
        }
        catch ( const Exception& e )
        {
            RegObject.Disconnect();
            PXSLogException( L"Error getting registry software data.", e, __FUNCTION__ );
        }
    }


    ///////////////////////////////////////
    // Office 2013

    StringArray Keys;
    LPCWSTR MSO2013_PACKAGES = L"SOFTWARE\\Microsoft\\Office\\15.0\\Common\\InstalledPackages\\";

    i = 0;
    while( ( i < ARRAYSIZE( REG_ACCESS ) ) && ( Keys.GetSize() == 0 ) )
    {
        // Trap individual errors so can continue
        try
        {
            RegObject.Connect2( REG_ACCESS[ i ].hKey, REG_ACCESS[ i ].samAccess );
            RegObject.GetSubKeyList( MSO2013_PACKAGES, &Keys );
            for ( j = 0; j < Keys.GetSize(); j++ )
            {
                RegKey  = MSO2013_PACKAGES;
                RegKey += Keys.Get( j );
                RegObject.GetStringValue( RegKey.c_str(), L"", &Value );
                if ( Value.StartsWith( L"Microsoft", false ))
                {
                    memset( &Data, 0, sizeof ( Data ) );

                    // Name
                    StringCchCopy( Data.szName, ARRAYSIZE( Data.szName ), Value.c_str() );

                    // Publisher
                    StringCchCopy( Data.szVendor, ARRAYSIZE( Data.szVendor ), L"Microsoft Corporation" );

                    // Version
                    Value = PXS_STRING_EMPTY;
                    RegObject.GetStringValue( RegKey.c_str(), L"ProductVersion", &Value );
                    if ( Value.c_str() )
                    {
                        StringCchCopy( Data.szVersion, ARRAYSIZE( Data.szVersion ), Value.c_str() );
                    }
                    pInstalled->Append( Data );
                }
            }
        }
        catch ( const Exception& e )
        {
            PXSLogException( L"Error getting registry software data.", e, __FUNCTION__ );
        }
        RegObject.Disconnect();
        i++;
    }
    pInstalled->Rewind();
}

//===============================================================================================//
//  Description:
//      Get the updates stored in the registry
//
//  Parameters:
//      pUpdates - a list to populate with the updates
//
//  Remarks:
//      Updates are stored for 95-2003, from Vista onwards this has stopped
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetUpdatesFromRegistry( TList<TYPE_UPDATE_DATA>* pUpdates )
{
    size_t   i = 0, j = 0, numKeys = 0, numUpdateKeys = 0;
    String   RegKey, SubKey, Update, OSKey, UpdateKey, InstalledDate;
    String   Description, ServicePackKey;
    Registry RegObject;
    StringArray Keys, UpdateKeys;
    TYPE_UPDATE_DATA   Data;
    WindowsInformation WindowsInfo;

    if ( pUpdates == nullptr )
    {
        throw ParameterException( L"pUpdates", __FUNCTION__ );
    }
    pUpdates->RemoveAll();

    // XP/2003 only, operating system updates are in the form
    // SOFTWARE\Microsoft\Updates\[OS NAME]\[ServicePack]\Qxxxxxx
    // SOFTWARE\Microsoft\Updates\[OS NAME]\[ServicePack]\KBxxxxxx
    // Usage of Qxxxxxx is being phased out. eg:
    RegObject.Connect( HKEY_LOCAL_MACHINE );
    OSKey = L"SOFTWARE\\Microsoft\\Updates\\Windows ";
    if ( WindowsInfo.IsWinXP() )
    {
        OSKey += L"XP";
    }
    else if ( WindowsInfo.IsWin2003() )
    {
        OSKey += L"Server 2003";
    }
    else
    {
        PXSLogAppInfo1( L"Windows Updates(R) not at registry path '%%1'", OSKey );
        return;
    }

    // Enumerate
    RegObject.GetSubKeyList( OSKey.c_str(), &Keys );
    numKeys = Keys.GetSize();
    for ( i = 0; i < numKeys; i++ )
    {
        SubKey = Keys.Get( i );

        // Make sure key is at least 3 characters long and begins with SP
        if ( ( SubKey.GetLength() >= 3 ) &&
             ( SubKey.IndexOfI( L"SP" ) == 0 ) )
        {
            // Get the subkeys for this service pack
            ServicePackKey  = OSKey;
            ServicePackKey += L"\\";
            ServicePackKey += SubKey;
            RegObject.GetSubKeyList( ServicePackKey.c_str(), &UpdateKeys );
            numUpdateKeys = UpdateKeys.GetSize();
            for ( j = 0; j < numUpdateKeys; j++ )
            {
                Update = UpdateKeys.Get( j );
                Update.Trim();
                if ( IsQorKbNumber( Update.c_str() ) )
                {
                    // Get its description
                    RegKey  = OSKey;
                    RegKey += L"\\";
                    RegKey += SubKey;
                    RegKey += L"\\";
                    RegKey += Update;

                    InstalledDate = PXS_STRING_EMPTY;
                    RegObject.GetStringValue( RegKey.c_str(), L"InstalledDate", &InstalledDate );
                    InstalledDate.Trim();

                    Description = PXS_STRING_EMPTY;
                    RegObject.GetStringValue( RegKey.c_str(), L"Description", &Description );
                    Description.Trim();

                    // Add to the list
                    memset( &Data, 0, sizeof ( Data ) );
                    StringCchCopy( Data.szUpdateID, ARRAYSIZE(Data.szUpdateID), Update.c_str() );
                    StringCchCopy( Data.szInstalledOn,
                                   ARRAYSIZE( Data.szInstalledOn ), InstalledDate.c_str() );
                    StringCchCopy( Data.szDescription,
                                   ARRAYSIZE( Data.szDescription ), Description.c_str() );
                    pUpdates->Append( Data );
                }
            }
        }
    }
    pUpdates->Rewind();
}

//===============================================================================================//
//  Description:
//      Get the updates installed by the Windows Installer
//
//  Parameters:
//      pUpdates - a list to populate with the updates
//
//  Remarks:
//      These are the patches for applications/services installed by the
//      Windows Installer
//
//      Windows Installer data is held at:
//      HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetUpdatesFromMsi( TList<TYPE_UPDATE_DATA>* pUpdates )
{
    UINT    resultPatches    = 0, resultProducts = 0, resultInfoEx = 0;
    DWORD   iProductIndex    = 0, iPatchIndex = 0;
    DWORD   cchTransformsBuf = 0, cchValue = 0;
    wchar_t szPatchBuf[ 64 ]    = { 0 };         // Enough for a GUID
    wchar_t szProductBuf[ 64 ]  = { 0 };         // Enough for a GUID
    wchar_t szInstallDate[ 256 ]  = { 0 };
    wchar_t szDisplayName[ 256 ]  = { 0 };
    wchar_t szTransformsBuf[ 256 ] = { 0 };
    TYPE_UPDATE_DATA Data;

    if ( pUpdates == nullptr )
    {
        throw ParameterException( L"pUpdates", __FUNCTION__ );
    }
    pUpdates->RemoveAll();

    // Enumerate the products to get the product GUIDs
    iProductIndex  = 0;     // Must be zero on first call
    resultProducts = MsiEnumProducts( iProductIndex, szProductBuf );
    while ( resultProducts == ERROR_SUCCESS )
    {
        szProductBuf[ ARRAYSIZE( szProductBuf ) - 1 ] = PXS_CHAR_NULL;

        // Enumerate the patches for this product
        iPatchIndex = 0;
        memset( szPatchBuf     , 0, sizeof ( szPatchBuf ) );
        memset( szTransformsBuf, 0, sizeof ( szTransformsBuf ) );
        cchTransformsBuf = ARRAYSIZE( szTransformsBuf );
        resultPatches = MsiEnumPatches( szProductBuf,
                                        iPatchIndex,
                                        szPatchBuf, szTransformsBuf, &cchTransformsBuf );
        while ( resultPatches == ERROR_SUCCESS )
        {
            szPatchBuf[ ARRAYSIZE( szPatchBuf ) - 1 ] = PXS_CHAR_NULL;
            szTransformsBuf[ ARRAYSIZE(szTransformsBuf) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_INSTALLDATE, install date
            memset( szInstallDate, 0, sizeof ( szInstallDate ) );
            cchValue = ARRAYSIZE( szInstallDate );
            resultInfoEx = MsiGetPatchInfoEx( szPatchBuf,
                                              szProductBuf,
                                              nullptr,
                                              MSIINSTALLCONTEXT_MACHINE,
                                              L"InstallDate", szInstallDate, &cchValue );
            if ( resultInfoEx != ERROR_SUCCESS )
            {
                memset( szInstallDate, 0, sizeof ( szInstallDate ) );
                PXSLogSysError( resultInfoEx, L"INSTALLPROPERTY_DISPLAYNAME failed." );
            }
            szInstallDate[ ARRAYSIZE( szInstallDate ) - 1 ] = PXS_CHAR_NULL;

            // INSTALLPROPERTY_DISPLAYNAME, DisplayName
            memset( szDisplayName, 0, sizeof ( szDisplayName ) );
            cchValue = ARRAYSIZE( szDisplayName );
            resultInfoEx = MsiGetPatchInfoEx( szPatchBuf,
                                              szProductBuf,
                                              nullptr,
                                              MSIINSTALLCONTEXT_MACHINE,
                                              L"DisplayName", szDisplayName, &cchValue );
            if ( resultInfoEx != ERROR_SUCCESS )
            {
                memset( szDisplayName, 0, sizeof ( szDisplayName ) );
                PXSLogSysError( resultInfoEx, L"INSTALLPROPERTY_DISPLAYNAME failed." );
            }
            szDisplayName[ ARRAYSIZE( szDisplayName ) - 1 ] = PXS_CHAR_NULL;

            // Add to the list
            memset( &Data, 0, sizeof ( Data ) );
            StringCchCopy( Data.szUpdateID, ARRAYSIZE( Data.szUpdateID ), szPatchBuf );

            StringCchCopy( Data.szInstalledOn, ARRAYSIZE( Data.szInstalledOn ), szInstallDate );

            StringCchCopy( Data.szDescription, ARRAYSIZE( Data.szDescription ), szDisplayName );
            pUpdates->Append( Data );

            // Next patch
            iPatchIndex++;
            memset( szPatchBuf     , 0, sizeof ( szPatchBuf ) );
            memset( szTransformsBuf, 0, sizeof ( szTransformsBuf ) );
            cchTransformsBuf = ARRAYSIZE( szTransformsBuf );
            resultPatches = MsiEnumPatches( szProductBuf,
                                            iPatchIndex,
                                            szPatchBuf, szTransformsBuf, &cchTransformsBuf );
        };

        // Log patch error
        if ( ( resultPatches != ERROR_SUCCESS ) &&
             ( resultPatches != ERROR_NO_MORE_ITEMS )  )
        {
            PXSLogSysError( resultProducts, L"MsiEnumPatches failed." );
        }

        // Next product
        iProductIndex++;
        memset( szProductBuf, 0, sizeof ( szProductBuf ) );
        resultProducts = MsiEnumProducts( iProductIndex, szProductBuf );
        szProductBuf[ ARRAYSIZE( szProductBuf ) - 1 ] = PXS_CHAR_NULL;
    };

    // Log product error
    if ( ( resultProducts != ERROR_SUCCESS ) &&
         ( resultProducts != ERROR_NO_MORE_ITEMS )  )
    {
        PXSLogSysError( resultProducts, L"MsiEnumProducts failed." );
    }
    pUpdates->Rewind();
}

//===============================================================================================//
//  Description:
//      Get the updates installed using Windows Management Instrumentation
//
//  Parameters:
//      pUpdates - a list to populate with the updates
//
//  Remarks:
//      On Vista and newer, this does not show those updates installed by
//      the Windows installer or from the Microsoft update site.
//
//      Win32_Win32_QuickFixEngineering, only applies to the Operating system
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetUpdatesFromWmi( TList<TYPE_UPDATE_DATA>* pUpdates )
{
    Wmi    WMI;
    String HotFixID, InstalledOn, Description;
    TYPE_UPDATE_DATA Data;

    if ( pUpdates == nullptr )
    {
        throw ParameterException( L"pUpdates", __FUNCTION__ );
    }
    pUpdates->RemoveAll();

    WMI.Connect( L"root\\cimv2" );
    WMI.ExecQuery( L"Select * from Win32_QuickFixEngineering" );
    while ( WMI.Next() )
    {
        // A HotFixID should be a Q/KB number but have observed unexpected
        // values so will filter out those.
        HotFixID = PXS_STRING_EMPTY;
        WMI.Get( L"HotFixID", &HotFixID );
        if ( IsQorKbNumber( HotFixID.c_str() ) )
        {
            InstalledOn = PXS_STRING_EMPTY;
            Description = PXS_STRING_EMPTY;
            WMI.Get( L"InstalledOn", &InstalledOn );
            WMI.Get( L"Description", &Description );

            // Add to the list
            memset( &Data, 0, sizeof ( Data ) );
            StringCchCopy( Data.szUpdateID,
                           ARRAYSIZE( Data.szUpdateID ), HotFixID.c_str() );
            if ( InstalledOn.c_str() )
            {
                StringCchCopy( Data.szInstalledOn,
                               ARRAYSIZE( Data.szInstalledOn ), InstalledOn.c_str() );
                if ( IsValidWmiUpdateInstalledOn( InstalledOn ) == false )
                {
                    PXSLogAppWarn2( L"Win32_QuickFixEngineering::InstalledOn value '%%1' is in an "
                                    L" unexpected date format, HotFixID = %%2", InstalledOn, HotFixID );
                }
            }

            if ( Description.c_str() )
            {
                StringCchCopy( Data.szDescription,
                               ARRAYSIZE( Data.szDescription ), Description.c_str() );
            }
            pUpdates->Append( Data );
        }
    }
    pUpdates->Rewind();
}

//===============================================================================================//
//  Description:
//      Get the updates installed using Windows Update Agent, i.e. from the
//      Windows update website
//
//  Parameters:
//      pUpdates - a list to populate with the updates
//
//  Remarks:
//      GetTotalHistoryCount reported as causing an access violation
//      in ConfirmDecline of wuapi.dll. On the two occasions this was reported
//      it was with XP Service pack 2 beta v5.4.3790.2055
//      (xpsp_sp2_beta1.031215-1745).
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::GetUpdatesFromWua( TList<TYPE_UPDATE_DATA>* pUpdates )
{
    const DWORD MAX_TIME_MILLI_SEC = 60000;   // Time out in milli-seconds
    LONG      count      = 0, i = 0;
    DATE      date     = 0.0;        // Its a double
    BSTR      bstrTitle   = nullptr, bstrUpdateID = nullptr;
    DWORD     tickStart = 0;
    String    InstalledOn, Title, UpdateID;
    HRESULT   hResult  = 0;
    Formatter Format;
    TYPE_UPDATE_DATA Data;
    UpdateOperation      updateOperation;
    OperationResultCode  operationResultCode;
    AutoIUnknownRelease  ReleaseUpdateSearcher;
    AutoIUnknownRelease  ReleaseEntryCollection;
    AutoIUnknownRelease  ReleaseUpdateIdentity;
    IUpdateSearcher*     pUpdateSearcher = nullptr;
    IUpdateIdentity*     pUpdateIdentity = nullptr;
    IUpdateHistoryEntry* pEntry          = nullptr;
    IUpdateHistoryEntryCollection* pEntryCollection = nullptr;

    if ( pUpdates == nullptr )
    {
        throw ParameterException( L"pUpdates", __FUNCTION__ );
    }
    pUpdates->RemoveAll();

    hResult = CoCreateInstance( CLSID_UpdateSearcher,
                                NULL,
                                CLSCTX_INPROC_SERVER,
                                IID_IUpdateSearcher,
                                reinterpret_cast<void**>( &pUpdateSearcher ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult,  L"CoCreateInstance", __FUNCTION__ );
    }

    if ( pUpdateSearcher == nullptr )
    {
        throw NullException( L"pUpdateSearcher", __FUNCTION__ );
    }
    ReleaseUpdateSearcher.Set( pUpdateSearcher );    // Set an auto delete

    // Get the count of updates
    hResult = pUpdateSearcher->GetTotalHistoryCount( &count );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IUpdateSearcher::GetTotalHistoryCount", __FUNCTION__ );
    }
    PXSLogAppInfo1( L"IUpdateSearcher::GetTotalHistoryCount = %%1.", Format.Int32( count ) );

    // Check that there are entries otherwise QueryHistory returns an error
    if ( count <= 0 )
    {
        return;     // Nothing to do
    }

    hResult = pUpdateSearcher->QueryHistory( 0, count, &pEntryCollection );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IUpdateSearcher::QueryHistory", __FUNCTION__ );
    }

    if ( pEntryCollection == nullptr )
    {
        throw NullException( L"pEntryCollection", __FUNCTION__ );
    }
    ReleaseEntryCollection.Set( pEntryCollection );

    // Get the count, again
    count = 0;
    hResult = pEntryCollection->get_Count( &count );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IUpdateHistoryEntryCollection::get_Count", __FUNCTION__ );
    }
    PXSLogAppInfo1( L"IUpdateHistoryEntryCollection::get_Count = %%1.", Format.Int32( count ) );

    // This enumeration can be lengthy set a timeout and task cancel
    tickStart = GetTickCount();
    for ( i = 0; i < count; i++ )
    {
        AutoIUnknownRelease ReleaseHistoryEntry;

        if ( ( GetTickCount() - tickStart ) > MAX_TIME_MILLI_SEC )
        {
            PXSLogAppWarn( L"Enumerating Update Agent timed out." );
            break;
        }

        if ( g_pApplication && g_pApplication->GetStopBackgroundTasks() )
        {
            return;
        }

        pEntry = nullptr;
        hResult = pEntryCollection->get_Item( i, &pEntry );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult,
                                L"IUpdateHistoryEntryCollection::get_Item", __FUNCTION__ );
        }

        if ( pEntry == nullptr )
        {
            throw NullException( L"pHistoryEntry", __FUNCTION__ );
        }
        ReleaseHistoryEntry.Set( pEntry );

        // Only want Installation + Succeeded
        if ( (SUCCEEDED( pEntry->get_Operation( &updateOperation     ) ) ) &&
             (SUCCEEDED( pEntry->get_ResultCode( &operationResultCode) ) )  )
        {
            if ( ( updateOperation     == uoInstallation ) &&
                 ( operationResultCode == orcSucceeded   )  )
            {
                // Get the update id
                UpdateID = PXS_STRING_EMPTY;
                pUpdateIdentity = nullptr;
                hResult = pEntry->get_UpdateIdentity( &pUpdateIdentity );
                if ( SUCCEEDED( hResult ) && pUpdateIdentity )
                {
                    ReleaseUpdateIdentity.Set( pUpdateIdentity );
                    bstrUpdateID = nullptr;
                    if ( SUCCEEDED( pUpdateIdentity->get_UpdateID( &bstrUpdateID ) ) )
                    {
                        AutoSysFreeString SysFreeUpdateID( bstrUpdateID );
                        UpdateID = bstrUpdateID;
                        UpdateID.Trim();
                    }
                    else
                    {
                        PXSLogComWarn( hResult, L"IUpdateIdentity::get_UpdateID failed." );
                    }
                }
                else
                {
                    PXSLogComWarn( hResult, L"IUpdateHistoryEntry::get_UpdateIdentity failed." );
                }

                // Add the date, according to the docs its in UTC
                InstalledOn = PXS_STRING_EMPTY;
                date    = 0.0;
                hResult = pEntry->get_Date( &date );
                if ( SUCCEEDED( hResult ) && ( date > 1.0 ) )
                {
                    // Convert date, which in an OLE date, throw away the time
                    InstalledOn = Format.OleTimeInIsoFormat( date );
                    InstalledOn.ReplaceChar( PXS_CHAR_SPACE, PXS_CHAR_NULL );
                    InstalledOn.Trim();
                }
                else
                {
                    PXSLogComWarn( hResult, L"IUpdateHistoryEntry::get_Date failed." );
                }

                // Add title
                bstrTitle = nullptr;  // Set to NULL
                Title  = PXS_STRING_EMPTY;
                if ( SUCCEEDED( pEntry->get_Title( &bstrTitle ) ) )
                {
                    AutoSysFreeString SysFreeTitle( bstrTitle );
                    Title =bstrTitle;
                    Title.Trim();
                }
                else
                {
                   PXSLogComWarn( hResult, L"IUpdateHistoryEntry::get_Title failed" );
                }

                // Add to the list
                memset( &Data, 0, sizeof ( Data ) );
                StringCchCopy( Data.szUpdateID, ARRAYSIZE( Data.szUpdateID ), UpdateID.c_str() );

                StringCchCopy( Data.szInstalledOn,
                               ARRAYSIZE( Data.szInstalledOn ), InstalledOn.c_str() );

                StringCchCopy( Data.szDescription,
                               ARRAYSIZE( Data.szDescription ), Title.c_str() );
                pUpdates->Append( Data );
            }
        }
    }
    pUpdates->Rewind();
}

//===============================================================================================//
//  Description:
//      Test if the InstalledOn member of the Win32_QuickFixEngineering WMI class appears to
//      be a vild date. 
//
//  Parameters:
//      InstalledOn - the String
//
//  Remarks:
//      User reported 16 character hex string appearing in the updates retrieved
//      by WMI for them member InstalledOn of class Win32_QuickFixEngineering . Expected formats
//      are of type yyyymmdd, dd/mm/yyyy, d/m/yyyy. However, The InstalledOn field value
//      ultimately comes from the registy and is of type REG_SZ so can be any value. Hence,
//      this is a very simple test to filter out what does not appear valid. Unfortunately,
//      the InstallDate member is often missing so won't fallback to that.
//
//  Returns:
//      true if the input seems to be a date
//===============================================================================================//
bool SoftwareInformation::IsValidWmiUpdateInstalledOn( const String& InstalledOn ) const
{
    bool   valid  = true;
    size_t i      = 0;
    size_t length = InstalledOn.GetLength();
    wchar_t wch   = 0;
    const wchar_t* VALID_CHARS = L"0123456789-/";

    // Expecting at least 8 characters, e.g. d/m/yyyy
    if ( length < 8 )
    {
        return false;
    }

    while ( ( i < length ) && ( valid == true ) )
    {
        wch = InstalledOn.CharAt( i );
        if ( wcschr( VALID_CHARS, wch ) == nullptr )
        {
            valid = false;
        }
        i++;
    }

    return valid;
}

//===============================================================================================//
//  Description:
//      Determine if the two specified updates are the same
//
//  Parameters:
//      pUpdate1 - pointer to an update structure to compare
//      pUpdate2 - pointer to an update structure to compare
//
//  Returns:
//      true if they are the same update
//===============================================================================================//
bool SoftwareInformation::IsSameUpdate( const TYPE_UPDATE_DATA* pUpdate1,
                                        const TYPE_UPDATE_DATA* pUpdate2 )
{
    if ( ( pUpdate1 == nullptr ) || ( pUpdate2 == nullptr ) )
    {
        return false;
    }

    if ( pUpdate1 == pUpdate2 )
    {
        return true;
    }

    // KB number
    if ( ( pUpdate1->szUpdateID[ 0 ] != PXS_CHAR_NULL ) &&
         ( lstrcmpi( pUpdate1->szUpdateID, pUpdate2->szUpdateID ) == 0 ) )
    {
        return true;
    }

    // See if the Update ID is in the description.
    if ( ( IsQorKbNumber( pUpdate1->szUpdateID ) ) &&
         ( _tcsstr( pUpdate2->szDescription, pUpdate1->szUpdateID ) ) )
    {
        return true;
    }

    // Description
    if ( ( pUpdate1->szDescription[ 0 ] != PXS_CHAR_NULL ) &&
         ( lstrcmpi( pUpdate1->szDescription, pUpdate2->szDescription ) == 0 ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Merge the installed software lists. Puts the merged result in the
//      MsiInstalled list.
//
//  Parameters:
//      pRegInstalled - list of software installed from the registry
//      pMsiInstalled - list of software installed by MSI
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::MergeInstalledSoftwareLists( TList<TYPE_INSTALLED_DATA>* pRegInstalled,
                                                       TList<TYPE_INSTALLED_DATA>* pMsiInstalled )
{
    TYPE_INSTALLED_DATA Element;
    const TYPE_INSTALLED_DATA* pReg = nullptr;

    if ( ( pRegInstalled == nullptr ) || ( pRegInstalled->IsEmpty() ) )
    {
        return;     // Nothing to do
    }

    if ( pMsiInstalled == nullptr )
    {
        throw ParameterException( L"pMsiInstalled", __FUNCTION__ );
    }

    // Merge the lists by matching on ID.
    pRegInstalled->Rewind();
    do
    {
        pReg = pRegInstalled->GetPointer();
        memset( &Element, 0, sizeof ( Element ) );
        if ( FindInstalledSoftwareElement( pReg, pMsiInstalled, &Element ) )
        {
            // Update the element
            // Copy the data obtained from the registry into the MSI
            // list if its missing that data. Some fields for MSI are
            // required so no need to test if those are missing.
            if ( lstrlen( Element.szInstallDate ) == 0 )
            {
                StringCchCopy( Element.szInstallDate,
                               ARRAYSIZE( Element.szInstallDate ), pReg->szInstallDate );
            }

            if ( lstrlen( Element.szInstallLocation ) == 0 )
            {
                StringCchCopy( Element.szInstallLocation,
                               ARRAYSIZE( Element.szInstallLocation ), pReg->szInstallLocation );
            }

            if ( lstrlen( Element.szInstallSource ) == 0 )
            {
                StringCchCopy( Element.szInstallSource,
                               ARRAYSIZE( Element.szInstallSource ), pReg->szInstallSource );
            }

            if ( lstrlen( Element.szProductID ) == 0 )
            {
                StringCchCopy( Element.szProductID,
                               ARRAYSIZE( Element.szProductID ), pReg->szProductID );
            }

            // These are not available using the Windows Installer
            Element.timesUsed = pReg->timesUsed;
            StringCchCopy( Element.szLastUsed, ARRAYSIZE( Element.szLastUsed ), pReg->szLastUsed );

            StringCchCopy( Element.szExePath, ARRAYSIZE( Element.szExePath ), pReg->szExePath );

            StringCchCopy( Element.szExeVersion,
                           ARRAYSIZE( Element.szExeVersion ), pReg->szExeVersion);

            StringCchCopy( Element.szExeDescription,
                           ARRAYSIZE( Element.szExeDescription ), pReg->szExeDescription );
            pMsiInstalled->Set( Element );
        }
        else
        {
            // Add the registry element to the list
            pMsiInstalled->Append( *pReg );
        }
    } while ( pRegInstalled->Advance() );
}

//===============================================================================================//
//  Description:
//      Merge the software updates lists. Puts the merged result pUpdatesList
//
//  Parameters:
//      pList        - list of additional updates to merge into pUpdatesList
//      pUpdatesList - resultant list of software updates
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::MergeSoftwareUpdatesLists( TList<TYPE_UPDATE_DATA>* pList,
                                                     TList<TYPE_UPDATE_DATA>* pUpdatesList )
{
    bool match = false;
    const TYPE_UPDATE_DATA* pData   = nullptr;
    const TYPE_UPDATE_DATA* pUpdate = nullptr;

    if ( ( pList == nullptr ) || ( pList->IsEmpty() ) )
    {
        return;     // Nothing to do
    }

    if ( pUpdatesList == nullptr )
    {
        throw ParameterException( L"pUpdatesList", __FUNCTION__ );
    }

    pList->Rewind();
    do
    {
        pData = pList->GetPointer();
        match = false;
        if ( pUpdatesList->IsEmpty() == false )
        {
            pUpdatesList->Rewind();
            do
            {
                pUpdate = pUpdatesList->GetPointer();
                if ( IsSameUpdate( pUpdate, pData ) )
                {
                    match = true;
                }
            } while ( ( match == false ) && ( pUpdatesList->Advance() ) );
        }

        if ( match == false )
        {
            pUpdatesList->Append( *pData );
        }
    } while ( pList->Advance() );
}

//===============================================================================================//
//  Description:
//      Determine if the specified string is a Knowledge Base number
//
//  Parameters:
//      pszString - the string to test
//
//  Remarks:
//      Current format starts with Q or KB and has at least 6 digits. If the
//      format changes, this function breaks.
//
//  Returns:
//      true if Q or KB number, else false
//===============================================================================================//
bool SoftwareInformation::IsQorKbNumber( LPCWSTR pszString )
{
    size_t i = 0, idxStart = 0, idxEnd = 0, charLength;

    if ( pszString == nullptr )
    {
        return false;
    }
    charLength = wcslen( pszString );

    // Must Start with either "Q" or "KB"
    if ( charLength < 2 )
    {
        return false;
    }

    if ( pszString[ 0 ] == 'Q' )
    {
        if ( charLength < 7 )
        {
            return false;
        }
        idxStart = 1;
        idxEnd   = 7;
    }
    else if ( ( pszString[ 0 ] == 'K' ) && ( pszString[ 1 ] == 'B' ) )
    {
        if ( charLength < 8 )
        {
            return false;
        }
        idxStart = 2;
        idxEnd   = 8;
    }
    else
    {
        // Does not start with "Q" or "KB"
        return false;
    }

    // These must be digits
    for ( i = idxStart; i < idxEnd; i++ )
    {
        if ( ( pszString[ i ] < '0' ) || ( pszString[ i ] > '9' ) )
        {
            return false;
        }
    }

    return true;    // Get here so it looks like a Q or KB number
}

//===============================================================================================//
//  Description:
//      Translate an MSI enumeration install state
//
//  Parameters:
//      state - the install state
//      pInstallState - receives the install state
//
//  Returns:
//      void
//===============================================================================================//
void SoftwareInformation::TranslateMsiInstallState( INSTALLSTATE state,
                                                    String* pInstallState )
{
    Formatter Format;

    if ( pInstallState == nullptr )
    {
        throw ParameterException( L"pInstallState", __FUNCTION__ );
    }

    switch( state )
    {
        default:
            *pInstallState = Format.UInt32( state );
            break;

        case INSTALLSTATE_NOTUSED:
            *pInstallState = L"Component disabled.";
            break;

        case INSTALLSTATE_BADCONFIG:
            *pInstallState = L"Configuration data corrupt.";
            break;

        case INSTALLSTATE_INCOMPLETE:
            *pInstallState = L"Installation suspended or in progress.";
            break;

        case INSTALLSTATE_SOURCEABSENT:
            *pInstallState = L"Run from source, source is unavailable.";
            break;

        case INSTALLSTATE_MOREDATA:
            *pInstallState = L"Return buffer overflow.";
            break;

        case INSTALLSTATE_INVALIDARG:
            *pInstallState = L"An invalid parameter was passed to the function.";
            break;

        case INSTALLSTATE_UNKNOWN:
            *pInstallState = L"The product is neither advertised or installed.";
            break;

        case INSTALLSTATE_BROKEN:
            *pInstallState = L"Broken.";
            break;

        case INSTALLSTATE_ADVERTISED:
            *pInstallState = L"The product is advertised but not installed.";
            break;

     // INSTALLSTATE_REMOVED is same value as INSTALLSTATE_ADVERTISED.
     // case INSTALLSTATE_REMOVED:
     //    *pInstallState =
     //             L"Component being removed (action state, not settable).";
     //    break;

        case INSTALLSTATE_ABSENT:
            *pInstallState = L"The product is installed for a different user.";
            break;

        case INSTALLSTATE_LOCAL:
            *pInstallState = L"Installed on local drive.";
            break;

        case INSTALLSTATE_SOURCE:
            *pInstallState = L"Run from source, CD or net.";
            break;

        case INSTALLSTATE_DEFAULT:
            *pInstallState = L"The product is installed for the current user.";
            break;
    }
}
