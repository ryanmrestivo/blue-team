///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Windows Information Class Implementation
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
#include "WinAudit/Header Files/WindowsInformation.h"

// 2. C System Files
#include <dxdiag.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AutoIUnknownRelease.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/Wmi.h"

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
WindowsInformation::WindowsInformation()
{
}

// Copy constructor
WindowsInformation::WindowsInformation( const WindowsInformation& oWindows )
{
    WindowsInformation();
    *this = oWindows;
}


// Destructor
WindowsInformation::~WindowsInformation()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
WindowsInformation& WindowsInformation::operator= ( const WindowsInformation& oWindows )
{
    if ( this == &oWindows ) return *this;

    // Nowt!

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the Corrective Service Disk Version for the operating system
//
//  Parameters:
//      pCSDVersion - receives the CSD version
//
//  Returns:
//      void
//===============================================================================================//
void WindowsInformation::GetCSDVersion( String* pCSDVersion ) const
{
    OSVERSIONINFOEX VersionInfoEx;

    if ( pCSDVersion == nullptr )
    {
        throw ParameterException( L"pCSDVersion", __FUNCTION__ );
    }

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );
    VersionInfoEx.szCSDVersion[ ARRAYSIZE( VersionInfoEx.szCSDVersion ) - 1 ] = PXS_CHAR_NULL;
    *pCSDVersion = VersionInfoEx.szCSDVersion;
    pCSDVersion->Trim();
}

//===============================================================================================//
//  Description:
//      Get the version of DirectX.
//
//  Parameters:
//      pDirectXVersion - string object to receive the version
//
//  Remarks:
//      As ever, particulary on Vista and newer, the registry is not reliable but this method
//      is fast.
//
//  Returns:
//      void, zero length "" string if not found or error
//===============================================================================================//
void WindowsInformation::GetDirectXVersion( String* pDirectXVersion ) const
{
    DWORD     i = 0, major = 0, minor = 0;
    String    RegValue, SubKey;
    Registry  RegObject;
    Formatter Format;

    if ( pDirectXVersion == nullptr )
    {
        throw ParameterException( L"pDirectXVersion", __FUNCTION__ );
    }
    *pDirectXVersion = PXS_STRING_EMPTY;

    ////////////////////////////////////////////////////////////////
    // Try 1, Check version key in registry against published values

    // XP->Windows 10/2012
    struct _VERSION
    {
        LPCWSTR pszString;
        LPCWSTR pszVersion;
    } Versions[] = { { L"4.08.00.0400"     , L"8.0"  },       // 8.0 or 8.0a
                     { L"4.08.01.0810"     , L"8.1"  },
                     { L"4.08.01.0881"     , L"8.1"  },
                     { L"4.08.01.0901"     , L"8.1"  },       // 8.1a or 8.1b
                     { L"4.08.02.0134"     , L"8.2"  },
                     { L"4.09.00.0900"     , L"9.0"  },
                     { L"4.09.00.0901"     , L"9.0a" },
                     { L"4.09.00.0902"     , L"9.0b" },
                     { L"4.09.00.0903"     , L"9.0c" },
                     { L"4.09.00.0904"     , L"9.0c" },
                     { L"6.00.6000.16386"  , L"10"   },
                     { L"6.00.6001.18000"  , L"10.1" },
                     { L"6.00.6002.18005"  , L"10.1" },
                     { L"6.01.7600.16385"  , L"11"   },
                     { L"6.00.6002.18107"  , L"11"   },
                     { L"6.01.7601.17514"  , L"11"   },
                     { L"6.02.9200.16384"  , L"11.1" },
                     { L"6.03.9600.16384"  , L"11.2" },
                     { L"10.00.10240.16384", L"12.0" } };

    try
    {
        SubKey = L"SOFTWARE\\Microsoft\\DirectX";
        RegObject.Connect( HKEY_LOCAL_MACHINE );
        RegObject.GetStringValue( SubKey.c_str(), L"Version", &RegValue );
        RegValue.Trim();
        for ( i = 0; i < ARRAYSIZE( Versions ); i++ )
        {
            if ( RegValue.CompareI( Versions[ i ].pszString ) == 0 )
            {
                *pDirectXVersion = Versions[ i ].pszVersion;
                break;
            }
        }
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( e, __FUNCTION__ );
    }

    if ( pDirectXVersion->GetLength() )
    {
        return;
    }

    //////////////////////////////////////////////////////////////////////
    // Try 2: Look in the registry for InstalledVersion key it seems often
    // to be an eight byte block of 2 DWORDS

    // Binary version in registry
    struct _INSTALLED_VERSION
    {
        DWORD high;
        DWORD low;
    } installedVersion = { 0, 0 };

    try
    {
        if ( ERROR_SUCCESS == RegObject.GetBinaryData( SubKey.c_str(),
                                                       L"InstalledVersion",
                                                       reinterpret_cast<BYTE*>(&installedVersion),
                                                       sizeof ( installedVersion ) ) )
        {
            major = HIBYTE( HIWORD ( installedVersion.high ) );
            minor = HIBYTE( HIWORD ( installedVersion.low  ) );
            if ( major )
            {
                *pDirectXVersion  = Format.UInt32( major );
                *pDirectXVersion += PXS_CHAR_DOT;
                *pDirectXVersion += Format.UInt32( minor );
            }
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }


    // If still no version number use whateever was found in the registry
    if ( pDirectXVersion->IsEmpty() )
    {
        *pDirectXVersion = RegValue;
    }
}

//===============================================================================================//
//  Description:
//      Get the version of DirectX on XP and newer
//
//  Parameters:
//      pDirectXVersion - string object to receive the version
//
//  Remarks:
//      See DirectX SDK sample GetDXVer.cpp, DxDiagProvider is slow, but there do not seem to
//      be a more eficient option
//
//      Will avoid DirectXSetupGetVersion in dsetup.dll as present by defualt on OS and even if
//      present it is not usually in the LoadLibrary search path.
//
//  Returns:
//      void, zero length "" string if not found or error
//===============================================================================================//
void WindowsInformation::GetDirectXVersionXP( String* pDirectXVersion ) const
{
    CLSID     clsid, iid;
    HRESULT   hResult;
    VARIANT   varProp;
    Formatter Format;
    DXDIAG_INIT_PARAMS  dxDiagInitParams;
    IDxDiagProvider*    pIDxDiagProvider = nullptr;
    IDxDiagContainer*   pIDxChildContainer = nullptr;
    IDxDiagContainer*   pIDxRootContainer  = nullptr;
    AutoIUnknownRelease AutoReleaseIDxDiagProvider;
    AutoIUnknownRelease AutoReleaseIDxChildContainer;
    AutoIUnknownRelease AutoReleaseIDxRootContainer;

    if ( pDirectXVersion == nullptr )
    {
        throw ParameterException( L"pDirectXVersion", __FUNCTION__ );
    }
    *pDirectXVersion = PXS_STRING_EMPTY;

    // CLSID_DxDiagProvider
    memset( &clsid, 0, sizeof ( clsid ) );
    hResult = CLSIDFromString( L"{A65B8071-3BFE-4213-9A5B-491DA4461CA7}" , &clsid );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CLSIDFromString, CLSID_DxDiagProvider", __FUNCTION__ );
    }

    // IID_IDxDiagProvider
    memset( &iid, 0, sizeof ( iid ) );
    hResult = CLSIDFromString( L"{9C6B4CB0-23F8-49CC-A3ED-45A55000A6D2}" , &iid );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CLSIDFromString, CLSID_DxDiagProvider", __FUNCTION__ );
    }

    // Get an instance of the DirectX Diagnostic Tool 
    hResult = CoCreateInstance( clsid,
                                nullptr,
                                CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                                iid,
                                reinterpret_cast<void**>( &pIDxDiagProvider ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoCreateInstance", __FUNCTION__ );
    }

    if ( pIDxDiagProvider == nullptr )
    {
        throw NullException( L"pIDxDiagProvider", __FUNCTION__ );
    }
    AutoReleaseIDxDiagProvider.Set( pIDxDiagProvider );

    // Initialise
    memset( &dxDiagInitParams, 0, sizeof( dxDiagInitParams ) );
    dxDiagInitParams.dwSize                  = sizeof( dxDiagInitParams );
    dxDiagInitParams.dwDxDiagHeaderVersion   = DXDIAG_DX9_SDK_VERSION;
    hResult = pIDxDiagProvider->Initialize( &dxDiagInitParams );

    // Get the containers
    hResult = pIDxDiagProvider->GetRootContainer( &pIDxRootContainer );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"GetRootContainer", __FUNCTION__ );
    }

    if ( pIDxRootContainer == nullptr )
    {
        throw NullException( L"pIDxRootContainer", __FUNCTION__ );
    }
    AutoReleaseIDxRootContainer.Set( pIDxRootContainer );

    hResult = pIDxRootContainer->GetChildContainer( L"DxDiag_SystemInfo", &pIDxChildContainer );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"GetChildContainer", __FUNCTION__ );
    }

    if ( pIDxChildContainer == nullptr )
    {
        throw NullException( L"pIDxChildContainer", __FUNCTION__ );
    }
    AutoReleaseIDxChildContainer.Set( pIDxChildContainer );

    // For some reason, property names are not in the SDK documentation, but can be seen using
    // the IDXDiagContainer::EnumPropNames method. The values returned are the user-friendly
    // "Display" ones, not the version numbers typically found in the registry or returned by
    // DirectXSetupGetVersion.
    VariantInit( &varProp );
    try
    {
        // Major version
        hResult = pIDxChildContainer->GetProp( L"dwDirectXVersionMajor", &varProp );
        if( SUCCEEDED( hResult ) )
        {
            if ( varProp.vt == VT_UI4 )
            {
                *pDirectXVersion += Format.UInt32( varProp.ulVal );
            }
            VariantClear( &varProp );
        }

        // Minor version
        VariantInit( &varProp );
        hResult = pIDxChildContainer->GetProp( L"dwDirectXVersionMinor", &varProp );
        if( SUCCEEDED( hResult ) )
        {
            if ( varProp.vt == VT_UI4 )
            {
                if ( pDirectXVersion->GetLength() )
                {
                    *pDirectXVersion += L".";
                }
                *pDirectXVersion += Format.UInt32( varProp.ulVal );
            }
            VariantClear( &varProp );
        }

        // Letter version
        VariantInit( &varProp );
        hResult = pIDxChildContainer->GetProp( L"szDirectXVersionLetter", &varProp );
        if( SUCCEEDED( hResult ) )
        {
            if ( varProp.vt == VT_BSTR && varProp.bstrVal != nullptr )
            {
                *pDirectXVersion += varProp.bstrVal[ 0 ];
            }
            VariantClear( &varProp );
        }
    }
    catch ( Exception& e )
    {
        if ( varProp.vt != VT_EMPTY )
        {
            VariantClear( &varProp );
        }
        throw e;
    }
}

//===============================================================================================//
//  Description:
//      Get the system's firmware data
//
//  Parameters:
//      pFirmwareDataNT6 - receives the data as a string
//
//  Remarks:
//      Requires NT6.
//
//  Returns:
//      void
//===============================================================================================//
void WindowsInformation::GetFirmwareDataNT6( String* pFirmwareDataNT6 ) const
{
    UINT    k = 0, bytesNeeded = 0;
    BYTE*   pBuffer = nullptr;
    DWORD   FirmwareTableBuffer[ 32 ] = { 0 };
    DWORD   j = 0, numTables = 0, ACPI, RSMB, FIRM;
    size_t  i = 0;
    String  ApplicationName, Title, DataString;
    HMODULE hKernel32 = nullptr;
    Formatter     Format;
    AllocateBytes FirmwareBytes;
    LPFN_GET_SYSTEM_FIRMWARE_TABLE   pfnGet  = nullptr;
    LPFN_ENUM_SYSTEM_FIRMWARE_TABLES pfnEnum = nullptr;

    // Convert character constants to DWORD
    ACPI = ('A' << 24) + ('C' << 16) + ('P' << 8)  + 'I';
    RSMB = ('R' << 24) + ('S' << 16) + ('M' << 8)  + 'B';
    FIRM = ('F' << 24) + ('I' << 16) + ('R' << 8)  + 'M';
    struct _SIGNATURE
    {
        DWORD   signature;
        LPCWSTR pszSignature;
    } Signatures[] = { { ACPI, L"ACPI" },
                       { RSMB, L"RSMB" },
                       { FIRM, L"FIRM" } };

    if ( pFirmwareDataNT6 == nullptr )
    {
        throw ParameterException( L"pFirmwareDataNT6", __FUNCTION__ );
    }
    *pFirmwareDataNT6 = PXS_STRING_EMPTY;

    hKernel32 = GetModuleHandle( L"kernel32.dll" );
    if ( hKernel32 == nullptr )
    {
        throw SystemException( GetLastError(),
                               L"kernel32.dll", "GetModuleHandle" );
    }

    // Get the firmware function pointers, disable C4191
    #ifdef _MSC_VER
        #pragma warning( push )
        #pragma warning ( disable : 4191 )
    #endif

        pfnGet  = (LPFN_GET_SYSTEM_FIRMWARE_TABLE)GetProcAddress( hKernel32,
                                                                  "GetSystemFirmwareTable" );
        pfnEnum = (LPFN_ENUM_SYSTEM_FIRMWARE_TABLES)GetProcAddress( hKernel32,
                                                                    "EnumSystemFirmwareTables" );
    #ifdef _MSC_VER
        #pragma warning( pop )
    #endif

    if ( ( pfnGet == nullptr ) || ( pfnEnum == nullptr ) )
    {
        throw SystemException( ERROR_PROC_NOT_FOUND, L"GetSystemFirmwareTable", __FUNCTION__);
    }

    // Title
    DataString.Allocate( 256 * 1024 );       // 256KB
    PXSGetApplicationName( &ApplicationName );
    Title  = L"Firmware Data by ";
    Title += ApplicationName;
    DataString += Title;
    DataString += PXS_STRING_CRLF;
    DataString.AppendChar( '=', Title.GetLength() );
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    // Get the Firmware Tables
    for ( i = 0; i < ARRAYSIZE( Signatures ); i++ )
    {
        memset( FirmwareTableBuffer, 0, sizeof ( FirmwareTableBuffer ) );
        bytesNeeded = pfnEnum( Signatures[ i ].signature,
                               FirmwareTableBuffer, sizeof ( FirmwareTableBuffer ) );

        // Buffer check
        if ( bytesNeeded > sizeof ( FirmwareTableBuffer ) )
        {
            throw SystemException( ERROR_INSUFFICIENT_BUFFER,
                                   L"EnumSystemFirmwareTables", __FUNCTION__ );
        }

        // Cycle through each table, according to documentation:
        //      ACPI not stated,
        //      FIRM returns 8 bytes (i.e. 2 tables) and
        //      RSMB returns 4 bytes (i.e. 1 table);
        numTables = bytesNeeded / sizeof ( FirmwareTableBuffer[ 0 ] );
        for ( j = 0; j < numTables; j++ )
        {
            // First call determines the buffer size. Observed
            // GetSystemFirmwareTable can return zero on success
            bytesNeeded = pfnGet( Signatures[ i ].signature,
                                  FirmwareTableBuffer[ j ], nullptr, 0 );

            FirmwareBytes.Delete();
            pBuffer = FirmwareBytes.New( bytesNeeded );
            pfnGet( Signatures[ i ].signature, FirmwareTableBuffer[ j ], pBuffer, bytesNeeded );

            DataString += L"\r\n--------------\r\n";
            DataString += L"FIRMWARE TABLE";
            DataString += L"\r\n--------------\r\n";

            DataString += L"Firmware Type : ";
            DataString += Signatures[ i ].pszSignature;
            DataString += PXS_STRING_CRLF;

            DataString += L"Table Number  : ";
            DataString += Format.UInt32( j );
            DataString += PXS_STRING_CRLF;

            DataString += L"Table ID      : ";
            DataString += Format.UInt32Hex(FirmwareTableBuffer[ j ], true);
            DataString += PXS_STRING_CRLF;

            DataString += L"Size (bytes)  : ";
            DataString += Format.UInt32( bytesNeeded );
            DataString += PXS_STRING_CRLF;

            DataString += L"Data          : ";
            for ( k = 0; k < bytesNeeded; k++ )
            {
                if ( ( k % 16 ) == 0 )
                {
                    DataString += PXS_STRING_CRLF;
                }
                DataString += Format.UInt8Hex( pBuffer[ k ], false );
                DataString += PXS_CHAR_SPACE;
            }
            DataString += PXS_STRING_CRLF;
        }
    }
    *pFirmwareDataNT6 = DataString;
}

//===============================================================================================//
//  Description:
//      Get the system's firmware data
//
//  Parameters:
//      pFirmwareDataXP - receives the firmware data as a formatted string
//
//  Returns:
//      void
//===============================================================================================//
void WindowsInformation::GetFirmwareDataXP( String* pFirmwareDataXP ) const
{
    const  DWORD MAX_SMBIOS_DATA_LEN = 65536;       // 64KB
    Wmi    WMI;
    BYTE   majorVersion = 0, minorVersion = 0, dmiRevision = 0;
    BYTE*  pBuffer = nullptr;
    DWORD  i = 0, size = 0;
    String ApplicationName, Title, Details, Data;
    Formatter     Format;
    AllocateBytes SmBiosBytes;

    if ( pFirmwareDataXP == nullptr )
    {
        throw ParameterException( L"pFirmwareDataXP", __FUNCTION__ );
    }
    *pFirmwareDataXP = PXS_STRING_EMPTY;

    pBuffer = SmBiosBytes.New( MAX_SMBIOS_DATA_LEN );
    WMI.Connect( L"root\\wmi" );
    WMI.ExecQuery( L"SELECT * FROM MSSmBios_RawSMBiosTables WHERE InstanceName=\"SMBiosData\"" );
    if ( WMI.Next() )      // Only expecting one record
    {
        // MOF says the data type is CIM_UINT32 but query returns INT32
        size = 0;
        WMI.GetUInt32( L"Size", &size );
        if ( size == 0 )
        {
            throw SystemException( ERROR_INVALID_DATA, L"size=0", __FUNCTION__ );
        }

        if ( size > MAX_SMBIOS_DATA_LEN )
        {
            Details = Format.StringUInt32( L"size=%%1", size );
            throw SystemException( ERROR_INSUFFICIENT_BUFFER, Details.c_str(), __FUNCTION__ );
        }

        // Fetch
        if ( size != WMI.GetUInt8Array( L"SmBiosData", pBuffer, MAX_SMBIOS_DATA_LEN ) )
        {
            throw SystemException( ERROR_INVALID_DATA, L"size" , __FUNCTION__ );
        }
        WMI.GetUInt8( L"SmbiosMajorVersion", &majorVersion );
        WMI.GetUInt8( L"SmbiosMinorVersion", &minorVersion );
        WMI.GetUInt8( L"DmiRevision"       , &dmiRevision );
    }
    PXSGetApplicationName( &ApplicationName );

    // Title
    Data.Allocate( 8092 );
    Title  = L"SMBIOS Firmware Data by ";
    Title += ApplicationName;
    Data += Title;
    Data += PXS_STRING_CRLF;
    Data.AppendChar( '=', Title.GetLength() );
    Data += PXS_STRING_CRLF;
    Data += PXS_STRING_CRLF;

    // Firmware Data
    Data += Format.StringUInt8( L"Major version: %%1\r\n", majorVersion );
    Data += Format.StringUInt8( L"Minor version: %%1\r\n", minorVersion );
    Data += Format.StringUInt8( L"DMI Revision : %%1\r\n", dmiRevision  );
    Data += Format.StringUInt32(L"Size (bytes) : %%1\r\n", size );
    Data += L"Data         :";
    for ( i = 0; i < size; i++ )
    {
        if ( ( i % 16 ) == 0 )
        {
            Data += PXS_STRING_CRLF;
        }
        Data += Format.UInt8Hex( pBuffer[ i ], false );
        Data += PXS_CHAR_SPACE;
    }

    // End of Data
    if ( Data.EndsWithStringI( PXS_STRING_CRLF ) == false )
    {
        Data += PXS_STRING_CRLF;
    }
    Data += L"!EOD\r\n";
    *pFirmwareDataXP = Data;
}

//===============================================================================================//
//  Description:
//      Get the language of the installed operating system
//
//  Parameters:
//      pLanguage - string to receive the language
//
//  Returns:
//      void
//===============================================================================================//
void WindowsInformation::GetLanguage( String* pLanguage ) const
{
    WORD      langID = 0;
    DWORD     dword  = 0;
    String    Locale;
    Registry  RegObject;
    Formatter Format;
    LPCWSTR   REG_LANGUAGE_KEY = L".DEFAULT\\Control Panel\\International\\";

    if ( pLanguage == nullptr )
    {
        throw ParameterException( L"pLanguage", __FUNCTION__ );
    }
    *pLanguage = PXS_STRING_EMPTY;

    // Get 'locale', its stored as a string in hex format
    RegObject.Connect( HKEY_USERS );
    RegObject.GetStringValue( REG_LANGUAGE_KEY, L"locale", &Locale );
    if ( Locale.GetLength() )
    {
        dword  = Format.HexStringToNumber( Locale );
        langID = PXSCastUInt32ToUInt16( 0xFFFF & dword );
        LanguageIDToName( langID, pLanguage );
    }
    else
    {
        PXSLogAppWarn( L"Did not find the 'locale' value in the registry." );
    }
}

//===============================================================================================//
//  Description:
//      Get the name of the operating system
//
//  Parameters:
//      pName - receives the name
//
//  Remarks:
//      Major.Minor    Name
//      5.1            XP
//      5.2            XP 64-bit
//      5.2            2003
//      6.0            Vista
//      6.0            2008
//      6.1            7
//      6.1            2008R2
//      6.2            8
//      6.2            2012
//      6.3            8.1
//      6.3            2012R2
//     10.0            Windows 10
//     10.0            Windows Server Technical Preview
//
//  Returns:
//      void
//===============================================================================================//
void WindowsInformation::GetName( String* pName ) const
{
    DWORD version;
    Formatter Format;
    OSVERSIONINFOEX VersionInfoEx;

    if ( pName == nullptr )
    {
        throw ParameterException( L"pName", __FUNCTION__ );
    }
    *pName = PXS_STRING_EMPTY;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );
    version = (10*VersionInfoEx.dwMajorVersion) + VersionInfoEx.dwMinorVersion;
    switch ( version )
    {
        default:
            *pName  = Format.UInt32( VersionInfoEx.dwMajorVersion );
            *pName += PXS_CHAR_DOT;
            *pName += Format.UInt32( VersionInfoEx.dwMinorVersion );
            break;

        case 51:
            *pName = L"XP";
            break;

        case 52:
            if ( VersionInfoEx.wProductType == VER_NT_WORKSTATION )
            {
                *pName = L"XP";
            }
            else
            {
                *pName = L"2003";
                if ( GetSystemMetrics( SM_SERVERR2 ) )
                {
                    *pName += L"R2";
                }
            }
            break;

        case 60:
            if ( VersionInfoEx.wProductType == VER_NT_WORKSTATION )
            {
                *pName = L"Vista";
            }
            else
            {
                *pName = L"2008";
            }
            break;

        case 61:
            if ( VersionInfoEx.wProductType == VER_NT_WORKSTATION )
            {
                *pName = L"7";
            }
            else
            {
                *pName = L"2008R2";
            }
            break;

        case 62:
            if ( VersionInfoEx.wProductType == VER_NT_WORKSTATION )
            {
                *pName = L"8";
            }
            else
            {
                *pName = L"2012";
            }
            break;

        case 63:
            if ( VersionInfoEx.wProductType == VER_NT_WORKSTATION )
            {
                *pName = L"8.1";
            }
            else
            {
                *pName = L"2012 R2";
            }
            break;

        case 100:
            if ( VersionInfoEx.wProductType == VER_NT_WORKSTATION )
            {
                *pName = L"10";
            }
            else
            {
                *pName = L"10 Server";
            }
            break;
    }
}

//===============================================================================================//
//  Description:
//      Get the name and edition of the operating system
//
//  Parameters:
//      pNameAndEdition - string object to receive the data
//
//  Remarks:
//      Form is: Microsoft Windows [Name] [Edition]
//      e.g. Microsoft Windows 95 OEM
//
//  Returns:
//      void
//===============================================================================================//
void WindowsInformation::GetNameAndEdition( String* pNameAndEdition ) const
{
    String Name, Edition;

    if ( pNameAndEdition == nullptr )
    {
        throw ParameterException( L"pNameAndEdition", __FUNCTION__ );
    }
    GetName( &Name );
    GetEdition( &Edition );

    // Some edition names contain "Windows" so avoid
    Edition.ReplaceI( L"Windows", L"" );
    Edition.Trim();

    // Start with Microsoft, some edition names contain "Windows" so avoid
    // duplication
    *pNameAndEdition  = L"Microsoft Windows ";
    *pNameAndEdition += Name;
    *pNameAndEdition += PXS_CHAR_SPACE;
    *pNameAndEdition += Edition;
}

//===============================================================================================//
//  Description:
//      Get the registration details for the operating system
//
//  Parameters:
//      pRegisteredOwner   - string object to receive registered owner
//      pRegisteredOrgName - string object to receive registered organization
//      pProductID         - string object to receive product ID
//      pPlusVersionNumber - string object to receive plus version number
//      pInstallDate       - receives installation date in UNIX time format
//
//  Returns:
//      void
//===============================================================================================//
void WindowsInformation::GetRegistrationDetails( String* pRegisteredOwner,
                                                 String* pRegisteredOrgName,
                                                 String* pProductID,
                                                 String* pPlusVersionNumber,
                                                 String* pInstallDate ) const
{
    Wmi       WMI;
    DWORD     value = 0;
    time_t    timeT = 0;
    String    SubKey, DateTime;
    Registry  RegObject;
    Formatter Format;
    SystemInformation SystemInfo;

    if ( ( pRegisteredOwner   == nullptr ) ||
         ( pRegisteredOrgName == nullptr ) ||
         ( pProductID         == nullptr ) ||
         ( pPlusVersionNumber == nullptr ) ||
         ( pInstallDate       == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pRegisteredOwner   = PXS_STRING_EMPTY;
    *pRegisteredOrgName = PXS_STRING_EMPTY;
    *pProductID         = PXS_STRING_EMPTY;
    *pPlusVersionNumber = PXS_STRING_EMPTY;
    *pInstallDate       = PXS_STRING_EMPTY;

    if ( SystemInfo.GetNumberOperatingSystemBits() == 32 )
    {
        // Look in the registry
        SubKey = L"Software\\Microsoft\\Windows NT\\CurrentVersion";
        RegObject.Connect( HKEY_LOCAL_MACHINE );
        RegObject.GetStringValue( SubKey.c_str(), L"RegisteredOwner", pRegisteredOwner );
        RegObject.GetStringValue( SubKey.c_str(), L"RegisteredOrganization", pRegisteredOrgName );
        RegObject.GetStringValue( SubKey.c_str(), L"ProductID", pProductID );
        RegObject.GetStringValue( SubKey.c_str(), L"Plus! VersionNumber", pPlusVersionNumber );

        // InstallDate is in UNIX time_t format
        RegObject.GetDoubleWordValue( SubKey.c_str(), L"InstallDate", 0, &value );
        timeT = PXSCastUInt32ToTimeT( value );
        *pInstallDate = Format.TimeTToLocalTimeInIso( timeT );
    }
    else
    {
        // 64-bit
        WMI.Connect( L"root\\cimv2" );
        WMI.ExecQuery( L"Select * from Win32_OperatingSystem" );
        if ( WMI.Next() )      // Only expecting one record
        {
            WMI.Get( L"RegisteredUser"   , pRegisteredOwner );
            WMI.Get( L"Organization"     , pRegisteredOrgName );
            WMI.Get( L"SerialNumber"     , pProductID );
            WMI.Get( L"PlusVersionNumber", pPlusVersionNumber );
            WMI.Get( L"InstallDate"      , &DateTime );

            // CIM_DATETIME format is yyyymmddHHMMSS.mmmmmmsUUU
            //                        0123456789012345678901234
            if ( DateTime.GetLength() == 25 )
            {
                // Express as yyyy-MM-dd hh:mm:ss. ignore the timezone information as
                // want a system data (i.e. UTC/GMT)
                *pInstallDate  = DateTime.CharAt( 0 );
                *pInstallDate += DateTime.CharAt( 1 );
                *pInstallDate += DateTime.CharAt( 2 );
                *pInstallDate += DateTime.CharAt( 3 );
                *pInstallDate += '-';
                *pInstallDate += DateTime.CharAt( 4 );
                *pInstallDate += DateTime.CharAt( 5 );
                *pInstallDate += '-';
                *pInstallDate += DateTime.CharAt( 6 );
                *pInstallDate += DateTime.CharAt( 7 );
                *pInstallDate += PXS_CHAR_SPACE;
                *pInstallDate += DateTime.CharAt( 8 );
                *pInstallDate += DateTime.CharAt( 9 );
                *pInstallDate += ':';
                *pInstallDate += DateTime.CharAt( 10 );
                *pInstallDate += DateTime.CharAt( 11 );
                *pInstallDate += ':';
                *pInstallDate += DateTime.CharAt( 12 );
                *pInstallDate += DateTime.CharAt( 13 );
            }
            else
            {
                // Shouldn't get here if CIM_DATETIME format is as exprected
                *pInstallDate = DateTime;
            }
        }
        WMI.Disconnect();
    }
}

/*
void WindowsInformation::GetRegistrationDetails( String* pRegisteredOwner,
                                                 String* pRegisteredOrgName,
                                                 String* pProductID,
                                                 String* pPlusVersionNumber,
                                                 String* pInstallDate ) const
{
    Wmi       WMI;
    DWORD     value = 0;
    time_t    timeT = 0;
    String    SubKey;
    Registry  RegObject;
    Formatter Format;
    SystemInformation SystemInfo;

    if ( ( pRegisteredOwner   == nullptr ) ||
         ( pRegisteredOrgName == nullptr ) ||
         ( pProductID         == nullptr ) ||
         ( pPlusVersionNumber == nullptr ) ||
         ( pInstallDate       == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pRegisteredOwner   = PXS_STRING_EMPTY;
    *pRegisteredOrgName = PXS_STRING_EMPTY;
    *pProductID         = PXS_STRING_EMPTY;
    *pPlusVersionNumber = PXS_STRING_EMPTY;
    *pInstallDate       = PXS_STRING_EMPTY;

    if ( SystemInfo.GetNumberOperatingSystemBits() == 32 )
    {
        // Look in the registry
        SubKey = L"Software\\Microsoft\\Windows NT\\CurrentVersion";
        RegObject.Connect( HKEY_LOCAL_MACHINE );
        RegObject.GetStringValue( SubKey.c_str(), L"RegisteredOwner", pRegisteredOwner );
        RegObject.GetStringValue( SubKey.c_str(), L"RegisteredOrganization", pRegisteredOrgName );
        RegObject.GetStringValue( SubKey.c_str(), L"ProductID", pProductID );
        RegObject.GetStringValue( SubKey.c_str(), L"Plus! VersionNumber", pPlusVersionNumber );

        // InstallDate is in UNIX time_t format
        RegObject.GetDoubleWordValue( SubKey.c_str(), L"InstallDate", 0, &value );
        timeT = PXSCastUInt32ToTimeT( value );
        *pInstallDate = Format.TimeTToLocalTimeInIso( timeT );
    }
    else
    {
        // 64-bit
        WMI.Connect( L"root\\cimv2" );
        WMI.ExecQuery( L"Select * from Win32_OperatingSystem" );
        if ( WMI.Next() )      // Only expecting one record
        {
            WMI.Get( L"RegisteredUser"   , pRegisteredOwner );
            WMI.Get( L"Organization"     , pRegisteredOrgName );
            WMI.Get( L"SerialNumber"     , pProductID );
            WMI.Get( L"PlusVersionNumber", pPlusVersionNumber );
            WMI.Get( L"InstallDate"      , pInstallDate );
        }
        WMI.Disconnect();
    }
}
*/

//===============================================================================================//
//  Description:
//      Determine if the OS Windows 2003
//
//  Parameters:
//      None
//
//  Returns:
//      true if Windows 2003, otherwise false
//===============================================================================================//
bool WindowsInformation::IsWin2003() const
{
    return IsMajorMinor( 5, 2 );
}

//===============================================================================================//
//  Description:
//      Determine if the OS Windows 10
//
//  Parameters:
//      None
//
//  Returns:
//      true if Windows 10, otherwise false
//===============================================================================================//
bool WindowsInformation::IsWindows10() const
{
    OSVERSIONINFOEX VersionInfoEx;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );
    if ( VersionInfoEx.dwMajorVersion == 10 )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//        Determine if the OS Windows XP
//
//  Parameters:
//      None
//
//  Returns:
//      true if Windows XP, otherwise false
//===============================================================================================//
bool WindowsInformation::IsWinXP() const
{
    return IsMajorMinor( 5, 1 );
}

//===============================================================================================//
//  Description:
//      Determine if the OS is Windows Vista or newer
//
//  Parameters:
//      None
//
//  Returns:
//      true if Windows Vista or newer othewise false
//===============================================================================================//
bool WindowsInformation::IsWinVistaOrNewer() const
{
    if ( GetMajorVersion() >= 6 )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the OS is Windows XP or newer
//
//  Parameters:
//      None
//
//  Returns:
//      true if Windows XP or newer oterhwise false
//===============================================================================================//
bool WindowsInformation::IsWinXPorNewer() const
{
    OSVERSIONINFOEX VersionInfoEx;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );
    if ( ( VersionInfoEx.dwMajorVersion >= 6 ) )
    {
        return true;
    }

    if ( ( VersionInfoEx.dwMajorVersion == 5 ) &&
         ( VersionInfoEx.dwMinorVersion >= 1 )  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Translate a language identifier to an English name
//
//  Parameters:
//      langID    - the language identifier
//      pLanguage - receives the language name
//
//  Returns:
//      void
//===============================================================================================//
void WindowsInformation::LanguageIDToName( WORD langID,
                                           String* pLanguage ) const
{
    LCID    lcid = 0;
    wchar_t LCData[ 256 ] = { 0 };     // Enough for a language name

    if ( pLanguage == nullptr )
    {
        throw ParameterException( L"pLanguage", __FUNCTION__ );
    }
    *pLanguage = PXS_STRING_EMPTY;

    lcid = MAKELCID( langID, 0 );
    if ( GetLocaleInfo( lcid,
                        LOCALE_SENGLANGUAGE,       // English language
                        LCData, ARRAYSIZE( LCData ) ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetLocaleInfo", __FUNCTION__ );
    }
    LCData[ ARRAYSIZE( LCData ) - 1 ] = PXS_CHAR_NULL;
    *pLanguage = LCData;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Determine if the operating system is the specified Major.Minor version
//
//  Parameters:
//      major - the major version number
//      minor - the minor version number
//
//  Returns:
//      true if the OS is Major.Minor, otherwise false
//===============================================================================================//
bool WindowsInformation::IsMajorMinor( DWORD major, DWORD minor ) const
{
    OSVERSIONINFOEX VersionInfoEx;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );
    if ( VersionInfoEx.dwMajorVersion == major &&
         VersionInfoEx.dwMinorVersion == minor  )
    {
        return true;
    }

    return false;
}
