///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Disk Information Class Implementation
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
#include "WinAudit/Header Files/DiskInformation.h"

// 2. C System Files
#include <devguid.h>
#include <ntddscsi.h>
#include <SetupAPI.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AllocateChars.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoCloseHandle.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/TArray.h"
#include "PxsBase/Header Files/Wmi.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Constants
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
DiskInformation::DiskInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
DiskInformation::~DiskInformation()
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
//      Get a information about the disk devices
//
//  Parameters:
//      pDiagnostics - string object to receive the formatted text
//
//  Remarks:
//      This data is diagnostic in nature so will add any errors or
//      exceptions to the input string
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetDiagnostics( String* pDiagnostics )
{
    size_t  i = 0;
    String  ApplicationName, Title, TempString, WmiDiskDriveData;
    String  ConfigManagerDiskData, DisksAndDevicesData, RawDiskData;
    String  DataString, LibraryVersion;
    SystemInformation  SystemInfo;

    if ( pDiagnostics == nullptr )
    {
        throw ParameterException( L"pDiagnostics", __FUNCTION__ );
    }
    *pDiagnostics = PXS_STRING_EMPTY;
    pDiagnostics->Allocate( 65536 );

    ////////////////////////////////////////////////////////////////////////////
    // Title

    PXSGetApplicationName( &ApplicationName );
    Title  = L"Disk, Adapter and Driver Data by ";
    Title += ApplicationName;
    DataString += Title;
    DataString += PXS_STRING_CRLF;
    DataString.AppendChar( '=', Title.GetLength() );
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    ////////////////////////////////////////////////////////////////////////////
    // Disks and Devices

    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;
    DataString += L"=================================================\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"|               DISKS & DEVICES                 |\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"=================================================\r\n";
    DataString += PXS_STRING_CRLF;

    try
    {
        GetDisksAndDevicesData( &DisksAndDevicesData );
        DataString += DisksAndDevicesData;
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
    }
    DataString += PXS_STRING_CRLF;


    ////////////////////////////////////////////////////////////////////////////
    // Data Blocks

    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;
    DataString += L"=================================================\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"|             RAW DISK DATA BLOCKS              |\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"=================================================";
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    try
    {
        GetRawDiskData( &RawDiskData );
        DataString += RawDiskData;
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
    }
    DataString += PXS_STRING_CRLF;

    /////////////////////////////////////////////////////////
    // WMI Enumeration

    DataString += PXS_STRING_CRLF;
    DataString += L"=================================================\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"|      Windows Management Disk Enumeration      |\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"=================================================";
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    try
    {
        GetWmiDiskDriveData( &WmiDiskDriveData );
        DataString += WmiDiskDriveData;
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
    }
    DataString += PXS_STRING_CRLF;

    ////////////////////////////////////////////////////////////////////////////
    // Configuration Manager Data

    DataString += PXS_STRING_CRLF;
    DataString += L"=================================================\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"|     Configuration Manager Disk Enumeration    |\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"=================================================";
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    try
    {
        GetConfigManagerDiskData( &ConfigManagerDiskData );
        DataString += ConfigManagerDiskData;
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
    }
    DataString += PXS_STRING_CRLF;

    ////////////////////////////////////////////////////////////////////////////
    // List some system drivers

    DataString += PXS_STRING_CRLF;
    DataString += L"=================================================\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"|                 SYSTEM FILES                  |\r\n";
    DataString += L"|                                               |\r\n";
    DataString += L"=================================================";
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    // File versions
    // Log system files used by the programme
    LPCWSTR Files[] = { L"drivers\\atapi.sys"   ,
                        L"drivers\\cdfs.sys"    ,
                        L"drivers\\cdrom.sys"   ,
                        L"drivers\\crcdisk.sys" ,
                        L"drivers\\classpnp.sys",
                        L"drivers\\disk.sys"    ,
                        L"drivers\\dmboot.sys"  ,
                        L"drivers\\dmio.sys"    ,
                        L"drivers\\dmload.sys"  ,
                        L"drivers\\ftdisk.sys"  ,
                        L"drivers\\imapi.sys"   ,
                        L"drivers\\intelide.sys",
                        L"drivers\\ntfs.sys"    ,
                        L"drivers\\partmgr.sys" ,
                        L"drivers\\pciidex.sys" ,
                        L"drivers\\redbook.sys" ,
                        L"drivers\\viaide.sys"  ,
                        L"drivers\\scsiport.sys",
                        L"drivers\\sfloppy.sys" ,
                        L"drivers\\storport.sys",
                        L"drivers\\tape.sys"    ,
                        L"updspapi.dll"         };

    DataString += L"-------------------------------------------------\r\n";
    DataString += L"Name                               | Version    |\r\n";
    DataString += L"-------------------------------------------------\r\n";

    try
    {
        for ( i = 0; i < ARRAYSIZE( Files ); i++ )
        {
            TempString = Files[ i ];
            TempString.FixedWidth( 36, PXS_CHAR_SPACE );
            DataString += TempString;

            LibraryVersion = PXS_STRING_EMPTY;
            SystemInfo.GetSystemLibraryVersion( Files[ i ], &LibraryVersion );
            DataString += LibraryVersion;
            DataString += PXS_STRING_CRLF;
        }
        DataString += PXS_STRING_CRLF;
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
    }
    DataString += PXS_STRING_CRLF;
    *pDiagnostics = DataString;
}

//===============================================================================================//
//  Description:
//      Get physical disk information
//
//  Parameters:
//      pRecords - array to receive the disk data
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetAuditRecords( TArray< AuditRecord >* pRecords )
{
    const  BYTE MAX_DISKS = 8;
    bool   haveWmiData = false, haveIddData = false, smartSupported = false;
    bool   smartEnabled = false;
    BYTE   i = 0;
    WORD   bufferKB = 0;
    DWORD  sectorsPerTrack = 0, totalHeads = 0;
    String MediaType, Model, SerialHex, FirmwareRevision, SerialHexIoctl;
    String FirmwareRevisionIoctl, SerialNumber, SerialHexIdd, DiskBufferKB;
    String FirmwareRevisionIdd, ManufacturerName, Rank, MasterSlave;
    String SmartTestResult, LocaleKB, LocaleMB, LocaleSmartSupported;
    String LocaleSmartEnabled, DiskSizeMB, LocaleNo, LocaleYes, Value;
    Formatter   Format;
    AuditRecord Record;
    unsigned __int64 diskSizeBytes = 0, totalCylinders = 0;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    PXSGetResourceString( PXS_IDS_136_MB  , &LocaleMB  );
    PXSGetResourceString( PXS_IDS_135_KB  , &LocaleKB  );
    PXSGetResourceString( PXS_IDS_1260_NO , &LocaleNo  );
    PXSGetResourceString( PXS_IDS_1261_YES, &LocaleYes );

    // Enumerate, i is the operating system drive numbering e.g. PhysicalDisk0
    for ( i = 0; i < MAX_DISKS; i++ )
    {
        // WMI
        try
        {
            haveWmiData = GetDataWMI( i,
                                      &diskSizeBytes,
                                      &totalCylinders,
                                      &totalHeads,
                                      &sectorsPerTrack,
                                      &MediaType,
                                      &ManufacturerName, &Model, &SerialHex, &FirmwareRevision );
        }
        catch ( const Exception& e )
        {
            haveWmiData = false;
            PXSLogException( L"Error getting WMI disk data",
                             e, __FUNCTION__ );
        }

        // Continue if it was found using WMI
        if ( haveWmiData )
        {
            // IOCTL
            try
            {
                GetDataIoctl( i, &SerialHexIoctl, &FirmwareRevisionIoctl );
            }
            catch ( const Exception& e )
            {
                PXSLogException( L"Error getting IOCTL disk data", e, __FUNCTION__ );
            }

            // IDD
            try
            {
                haveIddData = GetDataIdd( i,
                                          &bufferKB,
                                          &SerialHexIdd,
                                          &FirmwareRevisionIdd, &smartSupported, &smartEnabled );
            }
            catch ( const Exception& e )
            {
                PXSLogException( L"Error getting IDD disk data", e, __FUNCTION__ );
            }

            // SMART
            try
            {
                GetSmartTestResult( i, &SmartTestResult );
            }
            catch ( const Exception& e )
            {
                // Log and continue
                PXSLogException( L"Error performing a SMART self test.", e, __FUNCTION__ );
            }

            // Controller position, only applicable to internal fixed disks
            try
            {
                if ( MediaType.IndexOfI( L"fixed" ) != PXS_MINUS_ONE )
                {
                    GetFixedDiskControllerPosition( i, &Rank, &MasterSlave );
                }
            }
            catch ( const Exception& e )
            {
                // Log and continue
                PXSLogException( L"Error getting controller postion.", e, __FUNCTION__ );
            }


            // Serial Number
            if ( SerialHex.IsEmpty() )
            {
                SerialHex = SerialHexIoctl;
                if ( SerialHex.IsEmpty() && haveIddData )
                {
                    SerialHex = SerialHexIdd;
                }
            }
            DiskSerialHexToSerialNumber( SerialHex, &SerialNumber );

            // Firmware revision
            if ( FirmwareRevision.IsEmpty() && haveIddData )
            {
                FirmwareRevision = FirmwareRevisionIoctl;
                if ( FirmwareRevision.IsEmpty() )
                {
                    FirmwareRevision = FirmwareRevisionIdd;
                }
            }

            // Manufacturer, WMI sometimes reports (Standard Disk Drives)
            // for the manufacturer so will ignore this field.
            if ( ManufacturerName.CompareI( L"(Standard Disk Drives)" ) == 0 )
            {
                ManufacturerName = PXS_STRING_EMPTY;
            }

            // Weaker test for localised string, e.g.
            // (Lecteurs de disque standard)
            if ( ( ManufacturerName.IndexOf( '(', 0 ) == 0 )  &&
                 ( ManufacturerName.EndsWithCharacter( ')' ) ) )
            {
                ManufacturerName = PXS_STRING_EMPTY;
            }

            if ( ManufacturerName.IsEmpty() )
            {
                ResolveManufacturer( Model, &ManufacturerName );
            }

            DiskSizeMB    = Format.UInt64( diskSizeBytes / ( 1024 * 1024 ) );
            DiskSizeMB   += LocaleMB;

            DiskBufferKB = PXS_STRING_EMPTY;
            if ( bufferKB && haveIddData )
            {
                DiskBufferKB  = Format.UInt16( bufferKB );
                DiskBufferKB += LocaleKB;
            }

            if ( haveIddData )
            {
                if ( smartSupported )
                {
                    LocaleSmartSupported = LocaleYes;
                }
                else
                {
                    LocaleSmartSupported = LocaleNo;
                }

                if ( smartEnabled )
                {
                        LocaleSmartEnabled = LocaleYes;
                }
                else
                {
                    LocaleSmartEnabled = LocaleNo;
                }
            }

            Record.Reset( PXS_CATEGORY_PHYS_DISKS );
            Record.Add( PXS_PHYS_DISKS_ITEM_NUMBER  , Format.Int32( i + 1 ) );
            Record.Add( PXS_PHYS_DISKS_CAPACITY_MB  , DiskSizeMB );
            Record.Add( PXS_PHYS_DISKS_TYPE         , MediaType );
            Record.Add( PXS_PHYS_DISKS_MANUFACTURER , ManufacturerName );
            Record.Add( PXS_PHYS_DISKS_MODEL        , Model );
            Record.Add( PXS_PHYS_DISKS_SERIAL_NUMBER, SerialNumber );
            Record.Add( PXS_PHYS_DISKS_FIRMWARE_REV , FirmwareRevision );
            Record.Add( PXS_PHYS_DISKS_CONTROL_RANK , Rank );
            Record.Add( PXS_PHYS_DISKS_MASTER_SLAVE , MasterSlave );

            Value = Format.UInt64( totalCylinders );
            Record.Add( PXS_PHYS_DISKS_TOTAL_CYLINDERS, Value );

            Value = Format.UInt32( totalHeads );
            Record.Add( PXS_PHYS_DISKS_TOTAL_HEADS, Value );

            Value = Format.UInt32( sectorsPerTrack );
            Record.Add( PXS_PHYS_DISKS_SECTORS_PER_TRK, Value );

            Record.Add( PXS_PHYS_DISKS_BUFFER_KB      , DiskBufferKB );
            Record.Add( PXS_PHYS_DISKS_SMART_SUPPORTED, LocaleSmartSupported );
            Record.Add( PXS_PHYS_DISKS_SMART_ENABLED  , LocaleSmartEnabled );
            Record.Add( PXS_PHYS_DISKS_SMART_SELF_TEST, SmartTestResult );
            pRecords->Add( Record );
        }
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
//      Convert a hex format (raw) disk serial number to a display string
//
//  Parameters:
//      SerialHex     - hex string of the serial number
//      pSerialNumber - receives the diplay string
//
//  Remarks:
//      Normallu the serial number is a hex represention of character swapped
//      ASCII e.g.
//      30534538314a414f323132353738202020202020
//      0 S E 8 1 J A O 2 1 2 5 7 8
//      S08EJ1OA125287
//      However, if this does not seem to be the case will return the raw data
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::DiskSerialHexToSerialNumber( const String& SerialHex, String* pSerialNumber )
{
    bool      swap = true;   // Assume will swap characters
    size_t    i  = 0, lenChars = 0;
    wchar_t   ch = 0;
    DWORD     number = 0;
    String    HexByte;
    Formatter Format;

    if ( pSerialNumber == nullptr )
    {
        throw ParameterException( L"pSerialNumber", __FUNCTION__ );
    }
    *pSerialNumber = PXS_STRING_EMPTY;

    lenChars = SerialHex.GetLength();
    if ( lenChars == 0 )
    {
        return;     // Nothing to do
    }

    // Must be an divisble 4 for character swapping
    lenChars = SerialHex.GetLength();
    if ( lenChars % 4 )
    {
        swap = false;
        PXSLogAppWarn( L"HDD serial number length not multiple of 4." );
    }

    // Scan the string for unexpected characters
    for ( i = 0; i < lenChars; i++ )
    {
        ch = SerialHex.CharAt( i );
        if ( PXSIsHexitW( ch ) == false )
        {
            swap = false;
            PXSLogAppWarn( L"HDD serial number has non-hexadecimal chars." );
            break;
        }
    }

    // Sometimes the hex string has non-printable (< 0x20) characters. Perhaps
    // the disk's manufacturer intended the serial number to be shown
    // in hex format.
    for ( i = 0; i < lenChars; i+=2 )  // NB lenChars is even
    {
        // Test the first hexit of each 2-hexit byte
        ch = SerialHex.CharAt( i );
        if ( ch < '1' )
        {
            swap = false;
            PXSLogAppWarn( L"HDD serial number has non-printable chars." );
            break;
        }
    }

    // If no reversal, return the input string
    if ( swap == false )
    {
        *pSerialNumber = SerialHex;
        return;
    }

    // Swap the hex pairs
    for ( i = 0; i < (lenChars - 3); i += 4 )  // NB lenChars is > 3 and even
    {
        HexByte  = SerialHex.CharAt( i + 2 );
        HexByte += SerialHex.CharAt( i + 3 );
        number   = Format.HexStringToNumber( HexByte );
        pSerialNumber->AppendChar( 0xFF & number, 1 );

        HexByte  = SerialHex.CharAt( i );
        HexByte += SerialHex.CharAt( i + 1 );
        number   = Format.HexStringToNumber( HexByte );
        pSerialNumber->AppendChar( 0xFF & number, 1 );
    }
    pSerialNumber->Trim();      // Often the data is padded with spaces.
}

//===============================================================================================//
//  Description:
//      Get a string from an IDENTIFY_DEVICE_DATA structure
//
//  Parameters:
//      pbData      - byte buffer holding the string data
//      numBytes    - size of the buffer
//      pDataString - string object to receive the data
//
//  Remarks:
//      String data is in bytes that need reversing
//
//  Returns:
//      true if got data, else false
//===============================================================================================//
bool DiskInformation::ExtractIddStringValue( BYTE* pbData, size_t numBytes, String* pDataString )
{
    DWORD   i  = 0;
    wchar_t ch = 0;

    if ( pDataString == nullptr )
    {
        throw ParameterException( L"pDataString", __FUNCTION__ );
    }
    pDataString->Allocate( 128 );        // Only expecting short strings
    *pDataString = PXS_STRING_EMPTY;

    if ( ( pbData == nullptr ) || ( numBytes == 0 ) )
    {
        return false;
    }

    // Add bytes to string, make sure only going up to an even number of bytes
    numBytes = ( numBytes / 2 ) * 2;
    for ( i = 0; i < numBytes; i++ )
    {
        // Reverse bytes
        if ( ( i % 2 ) == 0 )
        {
            ch = pbData[ i + 1 ];
        }
        else
        {
            ch = pbData[ i - 1 ];
        }

        // Replace white spaces with spaces
        if ( PXSIsWhiteSpace( ch ) )
        {
            ch = PXS_CHAR_SPACE;
        }
        *pDataString += ch;
    }
    pDataString->Trim();

    return true;
}

//===============================================================================================//
//  Description:
//      Format the IDENTIFY_DEVICE_DATA structure as a string
//
//  Parameters:
//      pIDD       - pointer to the IDENTIFY_DEVICE_DATA structure
//      pIddString - receives the formatted data
//
//  Remarks:
//      ulBufferSize should normally be 512
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::FormatDiskFirmwareDataBlock( const PXSDDK::IDENTIFY_DEVICE_DATA* pIDD,
                                                   String* pIddString )
{
    ULONG     i = 0;
    String    FormattedIdd;
    Formatter Format;
    const BYTE* pBytes = nullptr;

    if ( pIddString == nullptr )
    {
        throw ParameterException( L"pIddString", __FUNCTION__ );
    }
    pIddString->Allocate( 4096 );
    *pIddString = PXS_STRING_EMPTY;

    if ( pIDD == nullptr )
    {
        return;     // Nothing to do
    }

    // Write out as an line of bytes
    pBytes     = reinterpret_cast<const BYTE*>( pIDD );
    *pIddString += L"0000:";
    for ( i = 0; i < sizeof ( PXSDDK::IDENTIFY_DEVICE_DATA ); i++ )
    {
        // Make the line start
        if ( ( i != 0 ) && ( ( i % 16 ) == 0 ) )
        {
            *pIddString += PXS_STRING_CRLF;
            if ( i < 99 )
            {
                // e.g. 0032:
                *pIddString += Format.StringUInt32( L"00%%1:", i );
            }
            else
            {
                // >= 100, // e.g. 0128:
                *pIddString += Format.StringUInt32( L"0%%1:", i );
            }
        }

        // Write the bytes
        *pIddString += Format.UInt8Hex( pBytes[ i ], false );
        *pIddString += PXS_CHAR_SPACE;
    }
    *pIddString += PXS_STRING_CRLF;

    // Translation the data structure.
    FormatIddStructureAsString( pIDD, &FormattedIdd );
    *pIddString += PXS_STRING_CRLF;
    *pIddString += L"Data Block Translation:\r\n";
    *pIddString += FormattedIdd;
    *pIddString += PXS_STRING_CRLF;
    *pIddString += PXS_STRING_CRLF;
}

//===============================================================================================//
//  Description:
//      Interpret an IDD structure
//
//  Parameters:
//      pIDD           - pointer to an IDENTIFY_DEVICE_DATA with the data
//      pFormattedIdd - string object to receive the formatted string
//
//  Remarks:
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::FormatIddStructureAsString( const PXSDDK::IDENTIFY_DEVICE_DATA* pIDD,
                                                  String* pFormattedIdd )
{
    DWORD     i = 0;
    String    Data;
    Formatter Format;

    if ( pFormattedIdd == nullptr )
    {
        throw ParameterException( L"pFormattedIdd", __FUNCTION__ );
    }
    *pFormattedIdd = PXS_STRING_EMPTY;

    if ( pIDD == nullptr )
    {
        return;     // Nothing to do
    }
    Data.Allocate( 2048 );

    // Add notice
    Data += L"Where appropriate 0 = false/no and 1 = true/yes.";
    Data += PXS_STRING_CRLF;
    Data += L"Some information requires character reversal, for";
    Data += PXS_STRING_CRLF;
    Data += L"example, aMtxro = Maxtor.\r\n";
    Data += PXS_STRING_CRLF;

    /////////////////////////////////
    // Add the data

    Data += L"GeneralConfiguration:\r\n";
    Data += Format.StringUInt32( L"\tReserved1            \t= %%1\r\n",
                                 pIDD->GeneralConfiguration.Reserved1 );
    Data += Format.StringUInt32( L"\tRetired3             \t= %%1\r\n",
                                 pIDD->GeneralConfiguration.Retired3 );
    Data += Format.StringUInt32( L"\tResponseIncomplete   \t= %%1\r\n",
                                 pIDD->GeneralConfiguration.ResponseIncomplete);
    Data += Format.StringUInt32( L"\tRetired2             \t= %%1\r\n",
                                 pIDD->GeneralConfiguration.Retired2 );
    Data += Format.StringUInt32( L"\tFixedDevice          \t= %%1\r\n",
                                 pIDD->GeneralConfiguration.FixedDevice );
    Data += Format.StringUInt32( L"\tRemovableMedia       \t= %%1\r\n",
                                pIDD->GeneralConfiguration.RemovableMedia );
    Data += Format.StringUInt32( L"\tRetired1             \t= %%1\r\n",
                                 pIDD->GeneralConfiguration.Retired1 );
    Data += Format.StringUInt32( L"\tDeviceType           \t= %%1\r\n",
                                 pIDD->GeneralConfiguration.DeviceType );

    Data += Format.StringUInt32( L"NumCylinders            \t= %%1\r\n",
                                 pIDD->NumCylinders );
    Data += Format.StringUInt32( L"ReservedWord2           \t= %%1\r\n",
                                 pIDD->ReservedWord2 );
    Data += Format.StringUInt32( L"NumHeads                \t= %%1\r\n",
                                 pIDD->NumHeads );

    Data += L"Retired1[2]:\r\n";
    Data += Format.StringUInt32( L"\tRetired1[0]          \t= %%1\r\n",
                                 pIDD->Retired1[0] );
    Data += Format.StringUInt32( L"\tRetired1[1]          \t= %%1\r\n",
                                 pIDD->Retired1[1] );

    Data += Format.StringUInt32( L"NumSectorsPerTrack      \t= %%1\r\n",
                                 pIDD->NumSectorsPerTrack );

    Data += L"VendorUnique1[3]:\r\n";
    Data += Format.StringUInt32( L"\tVendorUnique1[0]     \t= %%1\r\n",
                                 pIDD->VendorUnique1[0] );
    Data += Format.StringUInt32( L"\tVendorUnique1[1]     \t= %%1\r\n",
                                 pIDD->VendorUnique1[1] );
    Data += Format.StringUInt32( L"\tVendorUnique1[2]     \t= %%1\r\n",
                                 pIDD->VendorUnique1[2] );

    Data += L"SerialNumber[20]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->SerialNumber ); i++ )
    {
        Data += Format.StringUInt32( L"\tSerialNumber[%%1]   \t= ", i );
        Data += pIDD->SerialNumber[ i ];
        Data += PXS_STRING_CRLF;
    }

    Data += L"Retired2[2]:\r\n";
    Data += Format.StringUInt32( L"\tRetired2[0]          \t= %%1\r\n",
                                 pIDD->Retired2[0] );
    Data += Format.StringUInt32( L"\tRetired2[1]          \t= %%1\r\n",
                                 pIDD->Retired2[1] );

    Data += Format.StringUInt32( L"Obsolete1              \t= %%1\r\n",
                                 pIDD->Obsolete1 );

    Data += L"FirmwareRevision[8]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->FirmwareRevision ); i++ )
    {
        Data += Format.StringUInt32( L"\tFirmwareRevision[%%1]\t= ", i );
        Data += pIDD->FirmwareRevision[ i ];
        Data += PXS_STRING_CRLF;
    }

    Data += L"ModelNumber[40]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->ModelNumber ); i++ )
    {
        Data += Format.StringUInt32( L"\tModelNumber[%%1]      \t= ", i );
        Data += pIDD->ModelNumber[ i ];
        Data += PXS_STRING_CRLF;
    }

    Data += Format.StringUInt32( L"MaximumBlockTransfer       \t= %%1\r\n",
                                 pIDD->MaximumBlockTransfer );
    Data += Format.StringUInt32( L"VendorUnique2              \t= %%1\r\n",
                                pIDD->VendorUnique2 );
    Data += Format.StringUInt32( L"ReservedWord48             \t= %%1\r\n",
                                 pIDD->ReservedWord48 );

    Data += L"Capabilities:\r\n";
    Data += Format.StringUInt32( L"\tReservedByte49       \t= %%1\r\n",
                                pIDD->Capabilities.ReservedByte49 );
    Data += Format.StringUInt32( L"\tDmaSupported         \t= %%1\r\n",
                                 pIDD->Capabilities.DmaSupported );
    Data += Format.StringUInt32( L"\tLbaSupported         \t= %%1\r\n",
                                pIDD->Capabilities.LbaSupported );
    Data += Format.StringUInt32( L"\tIordyDisable         \t= %%1\r\n",
                                 pIDD->Capabilities.IordyDisable );
    Data += Format.StringUInt32( L"\tIordySupported       \t= %%1\r\n",
                                 pIDD->Capabilities.IordySupported );
    Data += Format.StringUInt32( L"\tReserved1            \t= %%1\r\n",
                                 pIDD->Capabilities.Reserved1 );
    Data += Format.StringUInt32( L"\tStandybyTimerSupport \t= %%1\r\n",
                                 pIDD->Capabilities.StandybyTimerSupport );
    Data += Format.StringUInt32( L"\tReserved2            \t= %%1\r\n",
                                 pIDD->Capabilities.Reserved2 );
    Data += Format.StringUInt32( L"\tReservedWord50       \t= %%1\r\n",
                                 pIDD->Capabilities.ReservedWord50 );

    Data += L"ObsoleteWords51[2]:\r\n";
    Data += Format.StringUInt32( L"\tObsoleteWords51[0]   \t= %%1\r\n",
                                 pIDD->ObsoleteWords51[0] );
    Data += Format.StringUInt32( L"\tObsoleteWords51[1]   \t= %%1\r\n",
                                 pIDD->ObsoleteWords51[1] );

    Data += Format.StringUInt32( L"TranslationFieldsValid:3   \t= %%1\r\n",
                                 pIDD->TranslationFieldsValid );
    Data += Format.StringUInt32( L"Reserved3:13               \t= %%1\r\n",
                                pIDD->Reserved3 );
    Data += Format.StringUInt32( L"NumberOfCurrentCylinders   \t= %%1\r\n",
                                 pIDD->NumberOfCurrentCylinders );
    Data += Format.StringUInt32( L"NumberOfCurrentHeads       \t= %%1\r\n",
                                 pIDD->NumberOfCurrentHeads );
    Data += Format.StringUInt32( L"CurrentSectorsPerTrack     \t= %%1\r\n",
                                 pIDD->CurrentSectorsPerTrack );
    Data += Format.StringUInt32( L"CurrentSectorCapacity      \t= %%1\r\n",
                                 pIDD->CurrentSectorCapacity );
    Data += Format.StringUInt32( L"CurrentMultiSectorSetting  \t= %%1\r\n",
                                 pIDD->CurrentMultiSectorSetting );
    Data += Format.StringUInt32( L"MultiSectorSettingValid : 1\t= %%1\r\n",
                                 pIDD->MultiSectorSettingValid );
    Data += Format.StringUInt32( L"ReservedByte59: 7          \t= %%1\r\n",
                                 pIDD->ReservedByte59 );
    Data += Format.StringUInt32( L"UserAddressableSectors     \t= %%1\r\n",
                                 pIDD->UserAddressableSectors );
    Data += Format.StringUInt32( L"ObsoleteWord62             \t= %%1\r\n",
                                 pIDD->ObsoleteWord62 );
    Data += Format.StringUInt32( L"MultiWordDMASupport : 8    \t= %%1\r\n",
                                 pIDD->MultiWordDMASupport );
    Data += Format.StringUInt32( L"MultiWordDMAActive : 8     \t= %%1\r\n",
                                 pIDD->MultiWordDMAActive );
    Data += Format.StringUInt32( L"AdvancedPIOModes : 8       \t= %%1\r\n",
                                 pIDD->AdvancedPIOModes );
    Data += Format.StringUInt32( L"ReservedByte64 : 8         \t= %%1\r\n",
                                 pIDD->ReservedByte64 );
    Data += Format.StringUInt32( L"MinimumMWXferCycleTime     \t= %%1\r\n",
                                 pIDD->MinimumMWXferCycleTime );
    Data += Format.StringUInt32( L"MinimumPIOCycleTime        \t= %%1\r\n",
                                 pIDD->MinimumPIOCycleTime );
    Data += Format.StringUInt32( L"MinimumPIOCycleTimeIORDY   \t= %%1\r\n",
                                 pIDD->MinimumPIOCycleTimeIORDY );

    Data += L"ReservedWords69[6]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->ReservedWords76 ); i++ )
    {
       Data += Format.StringUInt32_2( L"\tReservedWords69[%%1]   \t= %%2\r\n",
                                      i, pIDD->ReservedWords69[0] );
    }

    Data += Format.StringUInt32( L"QueueDepth : 5             \t= %%1\r\n",
                                 pIDD->QueueDepth );
    Data += Format.StringUInt32( L"ReservedWord75 : 11        \t= %%1\r\n",
                                 pIDD->ReservedWord75 );

    Data += L"ReservedWords76[4]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->ReservedWords76 ); i++ )
    {
        Data += Format.StringUInt32_2(
                                L"\tReservedWords76[%%1]   \t= %%2\r\n",
                                i, pIDD->ReservedWords76[ i ] );
    }

    Data += Format.StringUInt32( L"MajorRevision              \t= %%1\r\n",
                                 pIDD->MajorRevision );
    Data += Format.StringUInt32( L"MinorRevision              \t= %%1\r\n",
                                 pIDD->MinorRevision );

    Data += L"CommandSetSupport:\r\n";
    Data += Format.StringUInt32( L"\tSmartCommands        \t= %%1\r\n",
                                 pIDD->CommandSetSupport.SmartCommands );
    Data += Format.StringUInt32( L"\tSecurityMode         \t= %%1\r\n",
                                 pIDD->CommandSetSupport.SecurityMode );
    Data += Format.StringUInt32( L"\tRemovableMedia       \t= %%1\r\n",
                                 pIDD->CommandSetSupport.RemovableMedia );
    Data += Format.StringUInt32( L"\tPowerManagement      \t= %%1\r\n",
                                 pIDD->CommandSetSupport.PowerManagement );
    Data += Format.StringUInt32( L"\tReserved1            \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Reserved1 );
    Data += Format.StringUInt32( L"\tWriteCache           \t= %%1\r\n",
                                 pIDD->CommandSetSupport.WriteCache );
    Data += Format.StringUInt32( L"\tLookAhead            \t= %%1\r\n",
                                 pIDD->CommandSetSupport.LookAhead );
    Data += Format.StringUInt32( L"\tReleaseInterrupt     \t= %%1\r\n",
                                 pIDD->CommandSetSupport.ReleaseInterrupt );
    Data += Format.StringUInt32( L"\tServiceInterrupt     \t= %%1\r\n",
                                 pIDD->CommandSetSupport.ServiceInterrupt );
    Data += Format.StringUInt32( L"\tDeviceReset          \t= %%1\r\n",
                                 pIDD->CommandSetSupport.DeviceReset );
    Data += Format.StringUInt32( L"\tHostProtectedArea    \t= %%1\r\n",
                                 pIDD->CommandSetSupport.HostProtectedArea);
    Data += Format.StringUInt32( L"\tObsolete1            \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Obsolete1 );
    Data += Format.StringUInt32( L"\tWriteBuffer          \t= %%1\r\n",
                                 pIDD->CommandSetSupport.WriteBuffer );
    Data += Format.StringUInt32( L"\tReadBuffer           \t= %%1\r\n",
                                 pIDD->CommandSetSupport.ReadBuffer );
    Data += Format.StringUInt32( L"\tNop                  \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Nop );
    Data += Format.StringUInt32( L"\tObsolete2            \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Obsolete2 );
    Data += Format.StringUInt32( L"\tDownloadMicrocode    \t= %%1\r\n",
                                 pIDD->CommandSetSupport.DownloadMicrocode );
    Data += Format.StringUInt32( L"\tDmaQueued            \t= %%1\r\n",
                                 pIDD->CommandSetSupport.DmaQueued );
    Data += Format.StringUInt32( L"\tCfa                  \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Cfa );
    Data += Format.StringUInt32( L"\tAdvancedPm           \t= %%1\r\n",
                                 pIDD->CommandSetSupport.AdvancedPm );
    Data += Format.StringUInt32( L"\tMsn                  \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Msn );
    Data += Format.StringUInt32( L"\tPowerUpInStandby     \t= %%1\r\n",
                                 pIDD->CommandSetSupport.PowerUpInStandby );
    Data += Format.StringUInt32( L"\tManualPowerUp        \t= %%1\r\n",
                                 pIDD->CommandSetSupport.ManualPowerUp );
    Data += Format.StringUInt32( L"\tReserved2            \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Reserved2 );
    Data += Format.StringUInt32( L"\tSetMax               \t= %%1\r\n",
                                 pIDD->CommandSetSupport.SetMax );
    Data += Format.StringUInt32( L"\tAcoustics            \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Acoustics );
    Data += Format.StringUInt32( L"\tBigLba               \t= %%1\r\n",
                                pIDD->CommandSetSupport.BigLba );
    Data += Format.StringUInt32( L"\tResrved3             \t= %%1\r\n",
                                 pIDD->CommandSetSupport.Resrved3 );

    Data += Format.StringUInt32( L"ReservedWord84             \t= %%1\r\n",
                                 pIDD->ReservedWord84 );

    Data += L"CommandSetActive:\r\n";
    Data += Format.StringUInt32( L"\tSmartCommands        \t= %%1\r\n",
                                 pIDD->CommandSetActive.SmartCommands );
    Data += Format.StringUInt32( L"\tSecurityMode         \t= %%1\r\n",
                                 pIDD->CommandSetActive.SecurityMode );
    Data += Format.StringUInt32( L"\tRemovableMedia       \t= %%1\r\n",
                                 pIDD->CommandSetActive.RemovableMedia );
    Data += Format.StringUInt32( L"\tPowerManagement      \t= %%1\r\n",
                                 pIDD->CommandSetActive.PowerManagement );
    Data += Format.StringUInt32( L"\tReserved1            \t= %%1\r\n",
                                 pIDD->CommandSetActive.Reserved1 );
    Data += Format.StringUInt32( L"\tWriteCache           \t= %%1\r\n",
                                 pIDD->CommandSetActive.WriteCache );
    Data += Format.StringUInt32( L"\tLookAhead            \t= %%1\r\n",
                                 pIDD->CommandSetActive.LookAhead );
    Data += Format.StringUInt32( L"\tReleaseInterrupt     \t= %%1\r\n",
                                 pIDD->CommandSetActive.ReleaseInterrupt );
    Data += Format.StringUInt32( L"\tServiceInterrupt     \t= %%1\r\n",
                                 pIDD->CommandSetActive.ServiceInterrupt );
    Data += Format.StringUInt32( L"\tDeviceReset          \t= %%1\r\n",
                                 pIDD->CommandSetActive.DeviceReset );
    Data += Format.StringUInt32( L"\tHostProtectedArea    \t= %%1\r\n",
                                 pIDD->CommandSetActive.HostProtectedArea );
    Data += Format.StringUInt32( L"\tObsolete1            \t= %%1\r\n",
                                 pIDD->CommandSetActive.Obsolete1 );
    Data += Format.StringUInt32( L"\tWriteBuffer          \t= %%1\r\n",
                                 pIDD->CommandSetActive.WriteBuffer );
    Data += Format.StringUInt32( L"\tReadBuffer           \t= %%1\r\n",
                                 pIDD->CommandSetActive.ReadBuffer );
    Data += Format.StringUInt32( L"\tNop                  \t= %%1\r\n",
                                 pIDD->CommandSetActive.Nop );
    Data += Format.StringUInt32( L"\tObsolete2            \t= %%1\r\n",
                                 pIDD->CommandSetActive.Obsolete2 );
    Data += Format.StringUInt32( L"\tDownloadMicrocode    \t= %%1\r\n",
                                 pIDD->CommandSetActive.DownloadMicrocode );
    Data += Format.StringUInt32( L"\tDmaQueued            \t= %%1\r\n",
                                 pIDD->CommandSetActive.DmaQueued );
    Data += Format.StringUInt32( L"\tCfa                  \t= %%1\r\n",
                                 pIDD->CommandSetActive.Cfa );
    Data += Format.StringUInt32( L"\tAdvancedPm           \t= %%1\r\n",
                                 pIDD->CommandSetActive.AdvancedPm );
    Data += Format.StringUInt32( L"\tMsn                  \t= %%1\r\n",
                                 pIDD->CommandSetActive.Msn );
    Data += Format.StringUInt32( L"\tPowerUpInStandby     \t= %%1\r\n",
                                 pIDD->CommandSetActive.PowerUpInStandby );
    Data += Format.StringUInt32( L"\tManualPowerUp        \t= %%1\r\n",
                                 pIDD->CommandSetActive.ManualPowerUp );
    Data += Format.StringUInt32( L"\tReserved2            \t= %%1\r\n",
                                 pIDD->CommandSetActive.Reserved2 );
    Data += Format.StringUInt32( L"\tSetMax               \t= %%1\r\n",
                                 pIDD->CommandSetActive.SetMax );
    Data += Format.StringUInt32( L"\tAcoustics            \t= %%1\r\n",
                                 pIDD->CommandSetActive.Acoustics );
    Data += Format.StringUInt32( L"\tBigLba               \t= %%1\r\n",
                                 pIDD->CommandSetActive.BigLba );
    Data += Format.StringUInt32( L"\tResrved3             \t= %%1\r\n",
                                 pIDD->CommandSetActive.Resrved3 );

    Data += Format.StringUInt32( L"ReservedWord87             \t= %%1\r\n",
                                 pIDD->ReservedWord87 );
    Data += Format.StringUInt32( L"UltraDMASupport : 8        \t= %%1\r\n",
                                 pIDD->UltraDMASupport );
    Data += Format.StringUInt32( L"UltraDMAActive  : 8        \t= %%1\r\n",
                                 pIDD->UltraDMAActive );

    Data += L"ReservedWord89[4]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->ReservedWord89 ); i++ )
    {
       Data += Format.StringUInt32_2( L"\tReservedWord89[%%1]    \t= %%2\r\n",
                                      i, pIDD->ReservedWord89[ i ] );
    }

    Data += Format.StringUInt32( L"HardwareResetResult        \t= %%1\r\n",
                                 pIDD->HardwareResetResult );
    Data += Format.StringUInt32( L"CurrentAcousticValue : 8   \t= %%1\r\n",
                                 pIDD->CurrentAcousticValue );
    Data += Format.StringUInt32( L"RecommendedAcousticValue:8 \t= %%1\r\n",
                                 pIDD->RecommendedAcousticValue );

    Data += L"ReservedWord95[5]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->ReservedWord95 ); i++ )
    {
       Data += Format.StringUInt32_2( L"\tReservedWord95[%%1]    \t= %%2\r\n",
                                      i, pIDD->ReservedWord95[ i ] );
    }

    Data += L"Max48BitLBA[2]:\r\n";
    Data += Format.StringUInt32( L"\tMax48BitLBA[0]       \t= %%1\r\n",
                                 pIDD->Max48BitLBA[0] );
    Data += Format.StringUInt32( L"\tMax48BitLBA[1]       \t= %%1\r\n",
                                 pIDD->Max48BitLBA[1] );

    Data += L"ReservedWord104[23]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->ReservedWord104 ); i++ )
    {
        Data += Format.StringUInt32_2( L"\tReservedWord104[%%1] \t= %%2\r\n",
                                       i, pIDD->ReservedWord104[i] );
    }

    Data += Format.StringUInt32( L"MsnSupport : 2             \t= %%1\r\n",
                                 pIDD->MsnSupport );
    Data += Format.StringUInt32( L"ReservedWord127 : 14       \t= %%1\r\n",
                                 pIDD->ReservedWord127 );
    Data += Format.StringUInt32( L"SecurityStatus : 14        \t= %%1\r\n",
                                 pIDD->SecurityStatus );

    Data += L"ReservedWord129[126]:\r\n";
    for ( i = 0; i < ARRAYSIZE( pIDD->ReservedWord129 ); i++ )
    {
        Data += Format.StringUInt32_2( L"\tReservedWord129[%%1] \t= %%2\r\n",
                                       i, pIDD->ReservedWord129[i] );
    }

    Data += Format.StringUInt32( L"Signature : 8              \t= %%1\r\n",
                                 pIDD->Signature );
    Data += Format.StringUInt32( L"CheckSum : 8               \t= %%1\r\n",
                                 pIDD->CheckSum );

    *pFormattedIdd = Data;
}

//===============================================================================================//
//  Description:
//      Get formatted string data about disks from the Configuration
//      Manager (Device Manager)
//
//  Parameters:
//      ConfigManagerDiskData - receives the string data
//
//  Remarks:
//      This is diagnostic data so will add any error messages to the output
//      string
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetConfigManagerDiskData( String* pConfigManagerDiskData )
{
    GUID    guid;
    DWORD   MemberIndex = 0;
    wchar_t PropertyBuffer[ 512 ]   = { 0 };  // Enough for device description
    wchar_t DeviceInstanceId[ 256 ] = { 0 };  // Enough for device id
    String  DataString;
    HDEVINFO  DeviceInfoSet = nullptr;
    Formatter Format;
    SP_DEVINFO_DATA DeviceInfoData;

    if ( pConfigManagerDiskData == nullptr )
    {
        throw ParameterException( L"pConfigManagerDiskData", __FUNCTION__ );
    }
    DataString.Allocate( 4096 );

    // Get the hard disk controller devices
    guid = GUID_DEVCLASS_HDC;
    DeviceInfoSet = SetupDiGetClassDevs( &guid,
                                         nullptr, nullptr, DIGCF_PRESENT );
    if ( DeviceInfoSet == nullptr )
    {
        DataString += Format.SystemError( GetLastError() );
        DataString += L" SetupDiGetClassDevs, GUID_DEVCLASS_HDC failed.";
        DataString += PXS_STRING_CRLF;
        *pConfigManagerDiskData = DataString;
        return;
    }

    // Catch exceptions so can free the handle
    try
    {
        DataString += PXS_STRING_CRLF;
        DataString += L"IDE Disk Controllers\r\n";
        DataString += L"--------------------\r\n";
        DataString += PXS_STRING_CRLF;
        DataString += PXS_STRING_CRLF;

        memset( &DeviceInfoData, 0, sizeof ( DeviceInfoData ) );
        DeviceInfoData.cbSize = sizeof ( DeviceInfoData );
        while ( SetupDiEnumDeviceInfo( DeviceInfoSet, MemberIndex, &DeviceInfoData ) )
        {
            MemberIndex++;
            DataString += Format.StringUInt32( L"Index      : %%1\r\n", MemberIndex );

            // Get the controller's description
            DataString += L"Description: ";
            memset( PropertyBuffer, 0, sizeof ( PropertyBuffer ) );
            if ( SetupDiGetDeviceRegistryProperty( DeviceInfoSet,
                                                  &DeviceInfoData,
                                                  SPDRP_DEVICEDESC,
                                                  nullptr,
                                                  reinterpret_cast<BYTE*>( &PropertyBuffer ),
                                                  ARRAYSIZE( PropertyBuffer ),
                                                  nullptr ) )
            {
                PropertyBuffer[ ARRAYSIZE( PropertyBuffer ) - 1 ] = 0x00;
                DataString += PropertyBuffer;
            }
            else
            {
                DataString += Format.SystemError( GetLastError() );
                DataString += L" SPDRP_DEVICEDESC failed.\r\n";
            }
            DataString += PXS_STRING_CRLF;

            // Controller location
            DataString += L"Location   : ";
            memset( PropertyBuffer, 0, sizeof ( PropertyBuffer ) );
            if ( SetupDiGetDeviceRegistryProperty( DeviceInfoSet,
                                                   &DeviceInfoData,
                                                   SPDRP_LOCATION_INFORMATION,
                                                   nullptr,
                                                   reinterpret_cast<BYTE*>( &PropertyBuffer ),
                                                   ARRAYSIZE( PropertyBuffer ),
                                                   nullptr ) )
            {
                PropertyBuffer[ ARRAYSIZE( PropertyBuffer ) - 1 ] = 0x00;
                DataString += PropertyBuffer;
            }
            else
            {
                DataString += Format.SystemError( GetLastError() );
                DataString += L"SPDRP_LOCATION_INFORMATION failed.\r\n";
            }
            DataString += PXS_STRING_CRLF;

            // Get the controller's device id
            DataString += L"Hardware ID: ";
            memset( DeviceInstanceId, 0, sizeof ( DeviceInstanceId ) );
            if ( SetupDiGetDeviceInstanceId( DeviceInfoSet,
                                             &DeviceInfoData,
                                             DeviceInstanceId,
                                             ARRAYSIZE( DeviceInstanceId ),
                                             nullptr ) )
            {
                DeviceInstanceId[ ARRAYSIZE( DeviceInstanceId ) - 1 ] = 0x00;
                DataString += DeviceInstanceId;
            }
            else
            {
                DataString += Format.SystemError( GetLastError() );
                DataString += L" SetupDiGetDeviceInstanceId failed.\r\n";
            }
            DataString += PXS_STRING_CRLF;
            DataString += PXS_STRING_CRLF;
        }
    }
    catch ( const Exception& e )
    {
        // Continue
        DataString += PXS_STRING_CRLF;
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }
    SetupDiDestroyDeviceInfoList( DeviceInfoSet );
    *pConfigManagerDiskData = DataString;
}

//===============================================================================================//
//  Description:
//      Get the physical disk information using IDD
//
//  Parameters:
//      pPhysicalDrive    - the physical drive number as us by the
//                          operating system, e.g. PhysicalDrive0
//      pBufferKB         - receives the disks buffer size in KB
//      pSerialNumber     - receives the "raw" serial number, not byte swapped
//      pFirmwareRevision - receives the firmware revision
//      pSmartSupported   - receives if SMART is supported
//      pSmartEnabled     - receives if SMART is enabled
//
//  Returns:
//        void
//===============================================================================================//
bool DiskInformation::GetDataIdd( BYTE    physicalDrive,
                                  WORD*   pBufferKB,
                                  String* pSerialNumber,
                                  String* pFirmwareRevision,
                                  bool*   pSmartSupported, bool*   pSmartEnabled )
{
    bool      haveIDD = false;
    Formatter Format;
    SystemInformation  SystemInfo;
    WindowsInformation WindowsInfo;
    PXSDDK::IDENTIFY_DEVICE_DATA idd;

    if ( ( pBufferKB         == nullptr ) ||
         ( pSerialNumber     == nullptr ) ||
         ( pFirmwareRevision == nullptr ) ||
         ( pSmartSupported   == nullptr ) ||
         ( pSmartEnabled     == nullptr )  )
    {
        throw ParameterException( L"", __FUNCTION__ );
    }
    *pBufferKB         = 0;
    *pSerialNumber     = PXS_STRING_EMPTY;
    *pFirmwareRevision = PXS_STRING_EMPTY;
    *pSmartSupported   = false;
    *pSmartEnabled     = false;

    // Try to get the IDENTIFY_DEVICE_DATA structure
    memset( &idd, 0, sizeof ( idd ) );
    if ( SystemInfo.IsAdministrator() )
    {
        haveIDD = GetIDDSmart( physicalDrive, &idd );
    }
    else
    {
        // Usually ordinary users can get disk data using IOCTL_SCSI_MINIPORT
        // on older versions of NT
        if ( WindowsInfo.GetMajorVersion() <= 5 )
        {
            haveIDD = GetIDDScsiMiniPort( physicalDrive, &idd );
        }
    }

    if ( haveIDD == false )
    {
        PXSLogAppError1( L"Did not get IDD data for physical drive %%1. Try "
                         L"'Run As' then supply administrator credentials.",
                         Format.UInt8( physicalDrive ) );
        return false;
    }

    // This usually requires byte reversing but will return unmodified
    ExtractIddStringValue( idd.SerialNumber, ARRAYSIZE( idd.SerialNumber ), pSerialNumber );

    ExtractIddStringValue( idd.FirmwareRevision,
                           ARRAYSIZE( idd.FirmwareRevision ), pFirmwareRevision );

    // Buffer
    if ( idd.Retired2[ 1 ] > 0 )
    {
        *pBufferKB = idd.Retired2[ 1 ];
    }

    // SMART
    if ( idd.CommandSetSupport.SmartCommands )
    {
        *pSmartSupported = true;
    }
    if ( idd.CommandSetSupport.SmartCommands )
    {
        *pSmartEnabled = true;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the physical disk information using WMI
//
//  Parameters:
//      physicalDrive     - the physical drive number as us by the
//                          operating system, e.g. PhysicalDrive0
//      pSerialNumber     - receives the "raw" serial number
//      pFirmwareRevision - receives the firmware revision
//
//  Returns:
//        void
//===============================================================================================//
void DiskInformation::GetDataIoctl( BYTE physicalDrive,
                                    String* pSerialNumber, String* pFirmwareRevision )
{
    HANDLE    hDisk;
    String    FileName, DeviceType, DeviceTypeModifier;
    String    RemovableMedia, VendorId, ProductId, BusType, ErrorMessage;
    Formatter Format;

    if ( ( pSerialNumber == nullptr ) || ( pFirmwareRevision == nullptr ) )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pSerialNumber     = PXS_STRING_EMPTY;
    *pFirmwareRevision = PXS_STRING_EMPTY;

    // Open the disk
    FileName  = L"\\\\.\\PhysicalDrive";
    FileName += Format.UInt8( physicalDrive );
    hDisk = CreateFile( FileName.c_str(),
                        GENERIC_READ | GENERIC_WRITE,
                        FILE_SHARE_READ | FILE_SHARE_WRITE,
                        nullptr,
                        OPEN_EXISTING,        // Mandatory flag
                        0,
                        nullptr );
    if ( hDisk == INVALID_HANDLE_VALUE )
    {
        // Usually permission problem or drive not found
        PXSLogSysError1( GetLastError(),
                        L"CreateFile failed to open drive '%%1'.", FileName );
        return;
    }
    AutoCloseHandle CloseDisk( hDisk );

    // Use IOCTL_STORAGE_QUERY_PROPERTY
    if ( false == GetStorageDeviceProperties( hDisk,
                                              &DeviceType,
                                              &DeviceTypeModifier,
                                              &RemovableMedia,
                                              &VendorId,
                                              &ProductId,
                                              pFirmwareRevision,  // = Product revision
                                              pSerialNumber, &BusType, &ErrorMessage ) )
                {
        PXSLogAppError2( L"GetStorageDeviceProperty failed for '%%1'. %%2",
                         FileName, ErrorMessage );
    }
}

//===============================================================================================//
//  Description:
//      Get the physical disk information using WMI
//
//  Parameters:
//      physicalDrive    - the physical drive number as us by the
//                         operating system, e.g. PhysicalDrive0
//      diskSizeBytes    - receives the disk size in bytes
//      totalCylinders   - receives the total cylinders
//      totalHeads       - receives the total number of heads
//      sectorsPerTrack  - sectors per track
//      MediaType        - receives the media type, e.g disk, tape etc.
//      ManufacturerName - receives the manufacturer name. Not used
//      Model            - receives the model number
//      SerialNumber     - receives the "raw" serial number, not byte swapped
//      FirmwareRevision - receives the firmware revision
//
//  Remarks:
//      The amount of data that can be obtained depends on privilege level and
//      operating system. SerialNumber and FirmwareRevision require
//      Vista and newer.
//
//      Sometimes the caller does not know if the drive exists so will
//      return a success flag rather than throw.
//
//  Returns:
//      true if obtained data otherwise false;
//===============================================================================================//
bool DiskInformation::GetDataWMI( BYTE physicalDrive,
                                  UINT64* pDiskSizeBytes,
                                  UINT64* pTotalCylinders,
                                  DWORD*  pTotalHeads,
                                  DWORD*  pSectorsPerTrack,
                                  String* pMediaType,
                                  String* pManufacturerName,
                                  String* pModel,
                                  String* pSerialNumber,
                                  String* pFirmwareRevision )
{
    Wmi       WMI;
    bool      success = false;
    String    Query;
    Formatter Format;
    SystemInformation  SystemInfo;
    WindowsInformation WindowsInfo;

    if ( ( pDiskSizeBytes    == nullptr ) ||
         ( pTotalCylinders   == nullptr ) ||
         ( pTotalHeads       == nullptr ) ||
         ( pSectorsPerTrack  == nullptr ) ||
         ( pMediaType        == nullptr ) ||
         ( pManufacturerName == nullptr ) ||
         ( pModel            == nullptr ) ||
         ( pSerialNumber     == nullptr ) ||
         ( pFirmwareRevision == nullptr )  )
    {
        throw ParameterException( L"pDiskSizeBytes...", __FUNCTION__ );
    }

    *pDiskSizeBytes    = 0;
    *pTotalCylinders   = 0,
    *pTotalHeads       = 0,
    *pSectorsPerTrack  = 0,
    *pMediaType        = PXS_STRING_EMPTY;
    *pManufacturerName = PXS_STRING_EMPTY;
    *pModel            = PXS_STRING_EMPTY;
    *pSerialNumber     = PXS_STRING_EMPTY;
    *pFirmwareRevision = PXS_STRING_EMPTY;

    // Get Win32_DiskDrive data
    WMI.Connect( L"root\\cimv2" );
    Query  = L"Select * From Win32_DiskDrive WHERE DeviceID='";
    Query += L"\\\\\\\\.\\\\PHYSICALDRIVE";
    Query += Format.UInt8( physicalDrive );
    Query += L"'";
    WMI.ExecQuery( Query.c_str() );
    if ( WMI.Next() )
    {
        // The device id is recognised so drive exists, call that success
        success = true;
        WMI.GetUInt64( L"Size"           , pDiskSizeBytes );
        WMI.GetUInt64( L"TotalCylinders" , pTotalCylinders );
        WMI.GetUInt32( L"TotalHeads"     , pTotalHeads );
        WMI.GetUInt32( L"SectorsPerTrack", pSectorsPerTrack );
        WMI.Get( L"MediaType"   , pMediaType );
        WMI.Get( L"Manufacturer", pManufacturerName );
        WMI.Get( L"Model"       , pModel );

        // These are avialable from Vista onwards
        if ( WindowsInfo.GetMajorVersion() >= 6 )
        {
            WMI.Get( L"SerialNumber"    , pSerialNumber );
            WMI.Get( L"FirmwareRevision", pFirmwareRevision );
        }
    }
    WMI.Disconnect();

    // Win32_PhysicalMedia can get the serial number, needs admin privileges
    if ( success )
    {
        pSerialNumber->Trim();
        if ( pSerialNumber->IsEmpty() && SystemInfo.IsAdministrator() )
        {
            WMI.Connect( L"root\\cimv2" );
            Query  = L"Select SerialNumber From Win32_PhysicalMedia ";
            Query += L"WHERE Tag='\\\\\\\\.\\\\PHYSICALDRIVE";
            Query += Format.UInt8( physicalDrive );
            Query += L"'";
            WMI.ExecQuery( Query.c_str() );
            if ( WMI.Next() )
            {
                WMI.Get( L"SerialNumber", pSerialNumber );
            }
            WMI.Disconnect();
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Get partition information as a formatted string for the specified disk.
//
//  Parameters:
//      hDisk            - handle to the disk
//      pDiskPartionInfo - receives the partition information
//
//  Remarks:
//      This is diagnostic data so will add any error messages to the output
//      string
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetDiskPartionInfo( HANDLE hDisk, String* pDiskPartionInfo )
{
    const  size_t MAX_PARTITIONS_PER_DISK = 16;
    const  size_t DRIVE_LAYOUT_LEN = sizeof ( DRIVE_LAYOUT_INFORMATION ) +
                   ( MAX_PARTITIONS_PER_DISK * sizeof (PARTITION_INFORMATION) );
    size_t    i = 0;
    BYTE      OutBuffer[ DRIVE_LAYOUT_LEN ] = { 0 };
    DWORD     BytesReturned = 0;
    String    Insert1, PartitionType, DataString;
    Formatter Format;
    PARTITION_INFORMATION     PI;
    GET_LENGTH_INFORMATION    GLI;
    DRIVE_LAYOUT_INFORMATION* pDLI = nullptr;

    if ( pDiskPartionInfo == nullptr )
    {
        throw ParameterException( L"pDiskPartionInfo", __FUNCTION__ );
    }
    DataString.Allocate( 4096 );

    if ( ( hDisk == nullptr ) || ( hDisk == INVALID_HANDLE_VALUE ) )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        DataString = Format.String1(L"Invalid handle in %%1.\r\n", Insert1);
        *pDiskPartionInfo = DataString;
        return;
    }

    // Partition zero, usually or maybe always the size of the disk
    // IOCTL_DISK_GET_PARTITION_INFO is reported to show a maximum of
    // 137,438,952,960 bytes = 1FFFFFFE00 = 128GB
    DataString += L"Partition 'Zero'       : ";
    memset( &PI, 0, sizeof ( PI ) );
    if ( DeviceIoControl( hDisk,
                          IOCTL_DISK_GET_PARTITION_INFO,
                          nullptr, 0, &PI, sizeof ( PI ), &BytesReturned, nullptr ) )
    {
        DataString += Format.Int64( PI.PartitionLength.QuadPart );
        DataString += L" Bytes";
    }
    else
    {
        DataString += L"IOCTL_DISK_GET_PARTITION_INFO failed. ";
        DataString += Format.SystemError( GetLastError() );
    }
    DataString += PXS_STRING_CRLF;

    // Try IOCTL_DISK_GET_LENGTH_INFO, available on XP
    DataString += L"Disk Size              : ";
    BytesReturned = 0;
    memset( &GLI, 0, sizeof ( GLI ) );
    if ( DeviceIoControl( hDisk,
                          IOCTL_DISK_GET_LENGTH_INFO,
                          nullptr, 0, &GLI, sizeof ( GLI ), &BytesReturned, nullptr ) )
    {
        DataString += Format.Int64( GLI.Length.QuadPart );
        DataString += L" Bytes";
    }
    else
    {
        DataString += L"IOCTL_DISK_GET_LENGTH_INFO failed. ";
        DataString += Format.SystemError( GetLastError() );
    }
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    ////////////////////////////////////////////////////////////////////////////
    // Get the partition layout for the disk

    BytesReturned = 0;
    memset( OutBuffer, 0, sizeof ( OutBuffer ) );
    if ( DeviceIoControl( hDisk,
                          IOCTL_DISK_GET_DRIVE_LAYOUT,
                          nullptr,
                          0,
                          OutBuffer, sizeof ( OutBuffer ), &BytesReturned, nullptr ) == 0 )
    {
        DataString += L"IOCTL_DISK_GET_DRIVE_LAYOUT failed. ";
        DataString += Format.SystemError( GetLastError() );
        return;
    }

    // Do not show partition count as DRIVE_LAYOUT_INFORMATION always
    // returns a multiple of 4 for the PartitionCount on MBR partitions.
    pDLI = reinterpret_cast<DRIVE_LAYOUT_INFORMATION*>( OutBuffer );
    for ( i = 0; i < MAX_PARTITIONS_PER_DISK; i++ )
    {
        // Documentation says partition numbering starts at 1, so filter
        // out zeros.
        if ( ( pDLI->PartitionEntry[ i ].PartitionNumber > 0 ) &&
             ( pDLI->PartitionEntry[ i ].PartitionLength.QuadPart > 0 ) )
        {
            DataString += L"Partition Number       : ";
            DataString += Format.UInt32( pDLI->PartitionEntry[ i ].PartitionNumber );
            DataString += PXS_STRING_CRLF;

            DataString += L"Start Offset           : ";
            DataString += Format.Int64( pDLI->PartitionEntry[ i ].StartingOffset.QuadPart );
            DataString += PXS_STRING_CRLF;

            DataString += L"Partition Length[Bytes]: ";
            DataString += Format.Int64( pDLI->PartitionEntry[ i ].PartitionLength.QuadPart );
            DataString += PXS_STRING_CRLF;

            DataString += L"Hidden Sectors         : ";
            DataString += Format.UInt32( pDLI->PartitionEntry[ i ].HiddenSectors );
            DataString += PXS_STRING_CRLF;

            TranslatePartitionType( pDLI->PartitionEntry[ i ].PartitionType, &PartitionType );
            DataString += L"Partition Type         : ";
            DataString += PartitionType;
            DataString += PXS_STRING_CRLF;

            DataString += L"Bootable               : ";
            DataString += Format.Int32YesNo( pDLI->PartitionEntry[ i ].BootIndicator );
            DataString += PXS_STRING_CRLF;

            DataString += L"Recognized Partition   : ";
            DataString += Format.Int32YesNo( pDLI->PartitionEntry[ i ].RecognizedPartition );
            DataString += PXS_STRING_CRLF;

            DataString += L"Rewrite Partition      : ";
            DataString += Format.Int32YesNo( pDLI->PartitionEntry[ i ].RewritePartition );
            DataString += PXS_STRING_CRLF;
        }
    }
    DataString += PXS_STRING_CRLF;
    *pDiskPartionInfo = DataString;
}

//===============================================================================================//
//  Description:
//      Get diagnostic data about the systems disks and devices
//
//  Parameters:
//      pDisksAndDevicesData - receives formatted string data
//
//  Remarks:
//      Diagnostic so any errors are added to the input/output string
//
//  Returns:
//      true if got the data otherwise false
//===============================================================================================//
void DiskInformation::GetDisksAndDevicesData( String* pDisksAndDevicesData )
{
    const BYTE MAX_DEVICES = 8;  // Max of 8 adapters/devices per device type
    size_t i = 0;
    DWORD  deviceNum = 0, deviceCounter = 0;
    HANDLE hDisk = INVALID_HANDLE_VALUE;
    String FileName, Title, TempString, DeviceType, VendorId, ProductId;
    String ProductRevision, SerialNumber, ErrorMessage, DeviceTypeModifier;
    String RemovableMedia, AdapterProperty, ScsiAddressData, ScsiCapabilities;
    String ScsiInquiryData, DiskPartionInfo, BusType, DataString;
    Formatter Format;

    // Device Types
    LPCWSTR DeviceTypes[] = { L"\\\\.\\PhysicalDrive",
                              L"\\\\.\\Cdrom"        ,
                              L"\\\\.\\Tape"         ,
                              L"\\\\.\\Scanner"      ,
                              L"\\\\.\\Scsi"         };

    if ( pDisksAndDevicesData == nullptr )
    {
        throw ParameterException( L"pDisksAndDevicesData", __FUNCTION__ );
    }
    *pDisksAndDevicesData = PXS_STRING_EMPTY;

    try
    {
        for ( i = 0; i < ARRAYSIZE( DeviceTypes ); i++ )
        {
            // Enumerate the each device for the current device type
            for ( deviceNum = 0; deviceNum < MAX_DEVICES; deviceNum++ )
            {
                // Make the device name and see it it exists
                // e.g. \\.\PhysicalDrive0
                FileName  = DeviceTypes[ i ];
                FileName += Format.UInt32( deviceNum );
                hDisk = CreateFile( FileName.c_str(),
                                    GENERIC_READ | GENERIC_WRITE,
                                    FILE_SHARE_READ | FILE_SHARE_WRITE,
                                    nullptr, OPEN_EXISTING, 0, nullptr );
                if ( hDisk != INVALID_HANDLE_VALUE )
                {
                    AutoCloseHandle CloseDiskHandle( hDisk );

                    deviceCounter++;

                    // Make title string for this device
                    Title  = Format.StringUInt32( L"%%1) ADAPTER/DEVICE ",
                                                  deviceCounter );
                    Title += FileName;
                    TempString = PXS_STRING_EMPTY;
                    TempString.FixedWidth( Title.GetLength(), '=' );

                    DataString += L"\r\n\r\n";
                    DataString += Title;
                    DataString += PXS_STRING_CRLF;
                    DataString += TempString;
                    DataString += L"\r\n\r\n";

                    ///////////////////////////////////////////////////////////
                    // Get basic storage property data. Do twice, once for
                    // device and then for adapter. Only one will work;

                    DataString += L"------------------------\r\n";
                    DataString += L"a. Storage Property Data\r\n";
                    DataString += L"------------------------\r\n\r\n";

                    // Try to get the device property data. Note, the file name
                    // might point to an adapter
                    DataString += L"Device Data:\r\n";
                    DataString += L"------------\r\n\r\n";
                    GetStorageDeviceProperties( hDisk,
                                                &DeviceType,
                                                &DeviceTypeModifier,
                                                &RemovableMedia,
                                                &VendorId,
                                                &ProductId,
                                                &ProductRevision,
                                                &SerialNumber, &BusType, &ErrorMessage );
                    DataString += ErrorMessage;

                    DataString += L"Device Type            : ";
                    DataString += DeviceType;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Device Type Modifier   : ";
                    DataString += DeviceTypeModifier;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Removable Media        : ";
                    DataString += RemovableMedia;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Vendor Id              : ";
                    DataString += VendorId;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Product Id             : ";
                    DataString += ProductId;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Product Revision       : ";
                    DataString += ProductRevision;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Serial Number (Raw)    : ";
                    DataString += SerialNumber;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Bus Type               : ";
                    DataString += BusType;
                    DataString += PXS_STRING_CRLF;

                    // Try to get the adapter property data. Note, the file name
                    // might point to a device
                    DataString += L"Adapter Data:\r\n";
                    DataString += L"-------------\r\n\r\n";
                    GetStorageAdapterProperty( hDisk, &AdapterProperty );
                    DataString += AdapterProperty;

                    DataString += PXS_STRING_CRLF;
                    DataString += L"---------------\r\n";
                    DataString += L"b. Address Data\r\n";
                    DataString += L"---------------\r\n\r\n";
                    GetScsiAddressData( hDisk, &ScsiAddressData );
                    DataString += ScsiAddressData;

                    DataString += L"----------------------------\r\n";
                    DataString += L"c. Input/Output Capabilities\r\n";
                    DataString += L"----------------------------\r\n\r\n";
                    GetScsiCapabilities( hDisk, &ScsiCapabilities );
                    DataString += ScsiCapabilities;

                    DataString += L"-------------------\r\n";
                    DataString += L"d. Adapter Bus Data\r\n";
                    DataString += L"-------------------\r\n\r\n";
                    GetScsiInquiryData( hDisk, &ScsiInquiryData );
                    DataString += ScsiInquiryData;

                    DataString += L"-----------------\r\n";
                    DataString += L"e. Partition Data\r\n";
                    DataString += L"-----------------\r\n\r\n";
                    GetDiskPartionInfo( hDisk, &DiskPartionInfo );
                    DataString += DiskPartionInfo;
                }
            }
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }
    *pDisksAndDevicesData = DataString;
}

//===============================================================================================//
//  Description:
//      Get rank and master/slave postion of the specified disk
//
//  Parameters:
//      physicalDrive - the physical drive number as used by the operating
//                      system, e.g. PhysicalDrive0
//      pRank         - receives the controller rank
//      pMasterSlave  - receives if master or slave
//
//  Remarks:
//      Only applicable to fixed internal disks
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetFixedDiskControllerPosition( BYTE physicalDrive,
                                                      String* pRank, String* pMasterSlave )
{
    BYTE diskNumber = 0;
    Formatter Format;

    if ( ( pRank == nullptr ) || ( pMasterSlave == nullptr ) )
    {
        throw ParameterException( L"pRank/pMasterSlave", __FUNCTION__ );
    }
    *pRank        = PXS_STRING_EMPTY;
    *pMasterSlave = PXS_STRING_EMPTY;

    if ( PhysDriveNumToDiskNum( physicalDrive, &diskNumber ) == false )
    {
        return;
    }

    switch( diskNumber / 2 )
    {
        default:
            *pRank = Format.Int32( diskNumber / 2 );
            PXSLogAppWarn1( L"Unrecognised conroller rank '%%1;.", *pRank );
            break;

        case 0:
            *pRank = L"Primary";
            break;

        case 1:
            *pRank = L"Secondary";
            break;

        case 2:
            *pRank = L"Tertiary";
            break;

        case 3:
            *pRank = L"Quaternary";
            break;
    }

    // Master/slave is based on controller connection order
    if ( ( diskNumber % 2 ) == 0 )
    {
        *pMasterSlave = L"Master";
    }
    else
    {
        *pMasterSlave = L"Slave";
    }
}

//===============================================================================================//
//  Description:
//      Get the IDENTIFY_DEVICE_DATA structure for the specified disk
//
//  Parameters:
//      physicalDrive - the physical drive number as used by the operating
//                      system, e.g. PhysicalDrive0
//      pIDD          - receives the data
//
//  Remarks:
//      Uses the SCSI mini-port method. IOCTL_SCSI_MINIPORT_IDENTIFY is
//      largely defunct and often does not retrieve all the desired
//      information. However, does work on XP without administrator
//      privileges.
//
//      Rather than throw as the caller may be trying to see
//      if there is data to read, on error will log and return false false.
//
//  Returns:
//      true if got the data otherwise false
//===============================================================================================//
bool DiskInformation::GetIDDScsiMiniPort( BYTE physicalDrive,
                                          PXSDDK::IDENTIFY_DEVICE_DATA* pIDD )
{
    BYTE      diskNumber = 0;           // Controller numbering
    DWORD     BytesReturned = 0;
    HANDLE    hHandle = nullptr;
    String    FileName, Insert1;
    Formatter Format;

    // DeviceIoControl In-Out buffers
    BYTE InBuffer[  sizeof (SRB_IO_CONTROL) + sizeof (SENDCMDINPARAMS ) + 512 ];
    BYTE OutBuffer[ sizeof (SRB_IO_CONTROL) + sizeof (SENDCMDOUTPARAMS) + 512 ];
    SRB_IO_CONTROL*   pSIC  = reinterpret_cast<SRB_IO_CONTROL*>( InBuffer );
    SENDCMDINPARAMS*  pSCIP = reinterpret_cast<SENDCMDINPARAMS*>(
                                                          InBuffer + sizeof ( SRB_IO_CONTROL ) );
    SENDCMDOUTPARAMS* pSCOP = reinterpret_cast<SENDCMDOUTPARAMS*>(
                                                          OutBuffer + sizeof ( SRB_IO_CONTROL ) );

    // Check input
    if ( pIDD == nullptr )
    {
        return false;
    }
    memset( pIDD, 0, sizeof ( PXSDDK::IDENTIFY_DEVICE_DATA ) );

    // Open the disk, in this case the number used is the controller number.
    if ( PhysDriveNumToDiskNum( physicalDrive, &diskNumber ) == false )
    {
        // Log and return
        Insert1 = Format.UInt8( physicalDrive );
        PXSLogAppInfo1( L"Did not map physical drive %%1 to a disk number.",
                         Insert1 );
        return false;
    }

    FileName = Format.StringInt32( L"\\\\.\\Scsi%%1:", diskNumber/2 );
    hHandle  = CreateFile( FileName.c_str(),
                           GENERIC_READ | GENERIC_WRITE,
                           FILE_SHARE_READ | FILE_SHARE_WRITE,
                           nullptr, OPEN_EXISTING, 0, nullptr );
    if ( hHandle == INVALID_HANDLE_VALUE )
    {
        // Log and return rather than throw as the caller may be trying to see
        // if the disk exists
        PXSLogSysWarn1( GetLastError(), L"CreateFile failed for '%%1'.", FileName );
        return false;
    }
    AutoCloseHandle CloseFileHandle( hHandle );

    // Set up for DeviceIoControl
    memset( InBuffer , 0, sizeof ( InBuffer  ) );
    memset( OutBuffer, 0, sizeof ( OutBuffer ) );
    pSIC  = reinterpret_cast<SRB_IO_CONTROL*>( InBuffer );
    pSCIP = reinterpret_cast<SENDCMDINPARAMS*>( InBuffer + sizeof ( SRB_IO_CONTROL ) );

    // Length member is the Start of data after SRB_IO_CONTROL in InBuffer
    pSIC->HeaderLength = sizeof ( SRB_IO_CONTROL );
    pSIC->Timeout      = 5;                                 // Seconds
    pSIC->Length       = sizeof ( SENDCMDINPARAMS  ) + 512;
    pSIC->ControlCode  = IOCTL_SCSI_MINIPORT_IDENTIFY;
    memcpy( pSIC->Signature, "SCSIDISK", 8 );               // ASCII signature

    pSCIP->irDriveRegs.bCommandReg = ID_CMD;                // ATA identify
    if ( ( diskNumber % 2 ) == 0 )                          // Controller number
    {                                                       // either 0 or 1.
        pSCIP->bDriveNumber = 0;
    }
    else
    {
        pSCIP->bDriveNumber = 1;
    }

    // Issue command
    if ( DeviceIoControl( hHandle,
                          IOCTL_SCSI_MINIPORT,
                          InBuffer,
                          sizeof ( InBuffer ),
                          OutBuffer, sizeof ( OutBuffer ), &BytesReturned, nullptr ) == 0 )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysError1( GetLastError(), L"IOCTL_SCSI_MINIPORT failed in '%%1'.", Insert1 );
        return false;
    }

    // Test for returned data
    if ( BytesReturned == 0 )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogAppError1( L"BytesReturned = 0 in '%%1'.", Insert1 );
        return false;
    }

    // Verify that the returned buffer is at least as large as
    // IDENTIFY_DEVICE_DATA
    if ( pSCOP->cBufferSize < sizeof ( PXSDDK::IDENTIFY_DEVICE_DATA ) )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysError1( ERROR_INVALID_DATA,
                         L"IOCTL_SCSI_MINIPORT output is too small in '%%1'.", Insert1 );
        return false;
    }
    memcpy( pIDD, pSCOP->bBuffer, sizeof ( PXSDDK::IDENTIFY_DEVICE_DATA ) );


    return true;
}

//===============================================================================================//
//  Description:
//      Get the IDENTIFY_DEVICE_DATA structure for the specified disk
//
//  Parameters:
//      physicalDrive - the physical drive number as used by the
//                       operating system, e.g. PhysicalDrive0
//      pIDD           - receives the data
//
//  Remarks:
//      Uses the SCSI mini-port method. IOCTL_SCSI_MINIPORT_IDENTIFY is
//      largely defunct and often does not retrieve all the desired
//      information. However, does work on XP without administrator privileges.
//
//      See SMART_RCV_DRIVE_DATA documentation for buffer size.
//
//  Returns:
//      true if got the data otherwise false
//===============================================================================================//
bool DiskInformation::GetIDDSmart( BYTE physicalDrive, PXSDDK::IDENTIFY_DEVICE_DATA* pIDD )
{
    bool      result = false;
    BYTE      diskNumber = 0;       // Controller numbering
    BYTE      OutBuffer[ sizeof ( SENDCMDOUTPARAMS ) - 1 + 512 ] = { 0 };
    UCHAR     commandReg = 0;
    DWORD     BytesReturned = 0;
    HANDLE    hHandle = nullptr;
    String    FileName, Insert1;
    Formatter Format;
    SENDCMDINPARAMS    SCIP;
    GETVERSIONINPARAMS GVIP;
    SENDCMDOUTPARAMS*  pSCOP = reinterpret_cast<SENDCMDOUTPARAMS*>( OutBuffer );

    // Check input
    if ( pIDD == nullptr )
    {
        return false;
    }
    memset( pIDD, 0, sizeof ( PXSDDK::IDENTIFY_DEVICE_DATA ) );

    // Map from OS physical disk to controller disk number
    if ( PhysDriveNumToDiskNum( physicalDrive, &diskNumber ) == false )
    {
        Insert1 = Format.UInt8( physicalDrive );
        PXSLogAppInfo1( L"Did not map physical drive %%1 to a disk number.", Insert1 );
        return false;
    }


    // Create file
    FileName = Format.StringUInt32( L"\\\\.\\PhysicalDrive%%1",
                                    physicalDrive );
    hHandle  = CreateFile( FileName.c_str(),
                           GENERIC_READ | GENERIC_WRITE,
                           FILE_SHARE_READ | FILE_SHARE_WRITE,
                           nullptr, OPEN_EXISTING, 0, nullptr );
    if ( hHandle == INVALID_HANDLE_VALUE )
    {
        // Log and return
        PXSLogSysWarn1( GetLastError(),
                        L"CreateFile failed for '%%1'.", FileName );
        return false;
    }
    AutoCloseHandle CloseFileHandle( hHandle );

    // Get the version, etc. of PhysicalDrive IOCTL
    memset( &GVIP, 0, sizeof ( GVIP ) );
    BytesReturned = 0;
    if ( DeviceIoControl( hHandle,
                          SMART_GET_VERSION,
                          nullptr, 0, &GVIP, sizeof ( GVIP ), &BytesReturned, nullptr ) == 0 )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysError1( GetLastError(), L"SMART_GET_VERSION failed in %%1.", Insert1 );
        return false;
    }

    // Catch exceptions so can clean up
    try
    {
        // Show the IDE Device map value. Format as a binary with up to 8
        // digits, 1 for each possible disk
        Insert1 = Format.BinaryUInt8( GVIP.bIDEDeviceMap );
        Insert1.Reverse();
        Insert1.Truncate( 8 );
        Insert1.Reverse();
        PXSLogAppInfo1( L"IDE device map (primary master is the LSB): '%%1'", Insert1 );

        // The member bCommandReg takes as its command either disk or cd-rom
        // (ATAPI) identify. Bits 0-3 signify IDE drive, 4-7 for ATAPI drive.
        commandReg = ID_CMD;
        if ( ( GVIP.bIDEDeviceMap / 16 ) & (1 << diskNumber) )
        {
            commandReg = ATAPI_ID_CMD;  // CR-ROM drive
        }

        // Set up data structures for IDENTIFY command.
        memset( &SCIP, 0, sizeof ( SCIP ) );
        SCIP.cBufferSize                  = sizeof ( OutBuffer );
        SCIP.irDriveRegs.bFeaturesReg     = 0;
        SCIP.irDriveRegs.bSectorCountReg  = 1;
        SCIP.irDriveRegs.bSectorNumberReg = 1;
        SCIP.irDriveRegs.bCylLowReg       = 0;
        SCIP.irDriveRegs.bCylHighReg      = 0;
        SCIP.irDriveRegs.bCommandReg      = commandReg;
        SCIP.bDriveNumber                 = diskNumber;  // Controller numbering

        // 0xA0 for master, 0xB0 for slave
        SCIP.irDriveRegs.bDriveHeadReg = 0xA0;
        if ( 1 == ( diskNumber % 2 ) )
        {
            SCIP.irDriveRegs.bDriveHeadReg = 0xB0;
        }

        // Get the IDENTIFY_DEVICE_DATA data for the disk. It is in the bBuffer
        // member of the output buffer bSCOP.
        BytesReturned = 0;
        memset( OutBuffer, 0, sizeof ( OutBuffer ) );
        if ( DeviceIoControl( hHandle,
                              SMART_RCV_DRIVE_DATA,
                              &SCIP,
                              sizeof ( SCIP ),
                              OutBuffer, sizeof ( OutBuffer ), &BytesReturned, nullptr ) )
        {
            // Verify that the returned buffer is at least as large as
            // IDENTIFY_DEVICE_DATA
            PXSLogAppInfo1( L"SENDCMDOUTPARAMS.cBufferSize = %%1 bytes.",
                           Format.UInt32( pSCOP->cBufferSize ) );

            if ( pSCOP->cBufferSize >= sizeof ( PXSDDK::IDENTIFY_DEVICE_DATA ) )
            {
                memcpy( pIDD, pSCOP->bBuffer, sizeof ( PXSDDK::IDENTIFY_DEVICE_DATA ) );
                result = true;
            }
        }
        else
        {
            Insert1.SetAnsi( __FUNCTION__ );
            PXSLogSysError1( GetLastError(), L"SMART_RCV_DRIVE_DATA failed in %%1.", Insert1 );
        }
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"IDENTIFY_DEVICE_DATA using SMART.", e, __FUNCTION__ );
    }

    return result;
}

//===============================================================================================//
//  Description:
//      Get the IDENTIFY_DEVICE_DATA structure for the specified disk
//
//  Parameters:
//      physicalDrive - the physical drive number as used by the
//                      operating system, e.g. PhysicalDrive0
//      pIDD          - pointer to IDENTIFY_DEVICE_DATA
//
//  Returns:
//      true if got the data otherwise false
//===============================================================================================//
bool DiskInformation::GetIdentifyDeviceData( BYTE physicalDrive,
                                             PXSDDK::IDENTIFY_DEVICE_DATA* pIDD )
{
    bool result = false;
    SystemInformation  SystemInfo;
    WindowsInformation WindowsInfo;

    try
    {
        // Administrators only
        if ( SystemInfo.IsAdministrator() )
        {
            result = GetIDDSmart( physicalDrive, pIDD );
        }
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"GetIDDSmart failed.", e, __FUNCTION__ );
    }

    // If got nothing, try again for NT 5 and older
    if ( ( result == false ) && ( WindowsInfo.GetMajorVersion() <= 5 ) )
    {
        result = GetIDDScsiMiniPort( physicalDrive, pIDD );
    }
    else
    {
        PXSLogAppInfo( L"Administrator privileges are required to read "
                       L"IDENTIFY_DEVICE_DATA information. If logged on as "
                       L"Adminstrator on Vista or newer, try 'Run As' "
                       L"and supply adminstrator credentials." );
    }

    return result;
}

//===============================================================================================//
//  Description:
//      Get the raw data blocks for the disks on the system
//
//  Parameters:
//      pRawDiskData - receives the formatted string data
//
//  Remarks:
//      As this method returns diagnostic information, any errors will
//      be added to the output string.
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetRawDiskData( String* pRawDiskData )
{
    const BYTE MAX_DEVICES = 8;
    BYTE       i = 0;
    String     DataBlock;
    Formatter  Format;
    PXSDDK::IDENTIFY_DEVICE_DATA IDD;

    if ( pRawDiskData == nullptr )
    {
        throw ParameterException( L"pRawDiskData", __FUNCTION__ );
    }
    *pRawDiskData = PXS_STRING_EMPTY;

    try
    {
        for ( i = 0; i < MAX_DEVICES; i++ )
        {
            *pRawDiskData += L"---------------------\r\n";
            *pRawDiskData += L"Raw Data For Device ";
            *pRawDiskData += Format.UInt8( i );
            *pRawDiskData += L"\r\n---------------------\r\n\r\n";

            memset( &IDD, 0, sizeof ( IDD ) );
            if ( GetIdentifyDeviceData( i, &IDD ) )
            {
                DataBlock = PXS_STRING_EMPTY;
                FormatDiskFirmwareDataBlock( &IDD, &DataBlock );
            }
            else
            {
                DataBlock  = L"No information. Try 'Run As' administrator";
            }
            *pRawDiskData += DataBlock;
            *pRawDiskData += PXS_STRING_CRLF;
            *pRawDiskData += PXS_STRING_CRLF;
        }
        *pRawDiskData += PXS_STRING_CRLF;
    }
    catch ( const Exception& e )
    {
        *pRawDiskData += e.GetMessage();
        *pRawDiskData += PXS_STRING_CRLF;
    }
}

//===============================================================================================//
//  Description:
//      Get the specified disk's address data as a formatted string
//
//  Parameters:
//      hDisk       - handle to the disk
//      AddressData - receives the formatted string data
//
//  Remarks:
//      As this method returns diagnostic information, any errors will
//      be added to the output string.
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetScsiAddressData( HANDLE hDisk, String* pAddressData )
{
    DWORD     BytesReturned = 0;
    String    Insert1;
    Formatter Format;
    SCSI_ADDRESS scsiAddress;

    if ( pAddressData == nullptr )
    {
        throw ParameterException( L"pAddressData", __FUNCTION__ );
    }
    pAddressData->Allocate( 4096 );
    *pAddressData = PXS_STRING_EMPTY;

    if ( ( hDisk == nullptr ) || ( hDisk == INVALID_HANDLE_VALUE ) )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        *pAddressData = Format.String1( L"Invalid handle in '%%1'.\r\n", Insert1 );
        return;
    }

    BytesReturned = 0;
    memset( &scsiAddress, 0, sizeof ( scsiAddress ) );
    scsiAddress.Length = sizeof ( scsiAddress );
    if ( DeviceIoControl ( hDisk,
                           IOCTL_SCSI_GET_ADDRESS,
                           nullptr,
                           0,
                           &scsiAddress, sizeof ( scsiAddress ), &BytesReturned, nullptr ) == 0 )
    {
        *pAddressData += L"IOCTL_SCSI_GET_ADDRESS failed. ";
        *pAddressData += Format.SystemError( GetLastError() );
        return;
    }

    if ( BytesReturned == 0 )
    {
        *pAddressData += L"IOCTL_SCSI_GET_ADDRESS returned no data.\r\n";
        return;
    }

    *pAddressData += L"Port Number            : ";
    *pAddressData += Format.UInt32( scsiAddress.PortNumber );
    *pAddressData += PXS_STRING_CRLF;

    *pAddressData += L"Path Id                : ";
    *pAddressData += Format.UInt32( scsiAddress.PathId );
    *pAddressData += PXS_STRING_CRLF;

    *pAddressData += L"Target Id              : ";
    *pAddressData += Format.UInt32( scsiAddress.TargetId );
    *pAddressData += PXS_STRING_CRLF;

    *pAddressData += L"LUN                    : ";
    *pAddressData += Format.UInt32( scsiAddress.Lun );
    *pAddressData += PXS_STRING_CRLF;
}

//===============================================================================================//
//  Description:
//      Get the specified disk's input-output capabilities as a formatted
//      string
//
//  Parameters:
//      hDisk - handle to the disk
//      pScsiCapabilities - receives the formatted string data
//
//  Remarks:
//      As this method returns diagnostic information, any errors will
//      be added to the output string.
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetScsiCapabilities( HANDLE hDisk, String* pScsiCapabilities )
{
    DWORD     BytesReturned = 0;
    String    Insert1;
    Formatter Format;
    IO_SCSI_CAPABILITIES iosc;

    if ( pScsiCapabilities == nullptr )
    {
        throw ParameterException( L"pScsiCapabilities", __FUNCTION__ );
    }
    pScsiCapabilities->Allocate( 4096 );
    *pScsiCapabilities = PXS_STRING_EMPTY;

    if ( ( hDisk == nullptr ) || ( hDisk == INVALID_HANDLE_VALUE ) )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        *pScsiCapabilities = Format.String1( L"Invalid handle in '%%1'.\r\n", Insert1 );
        return;
    }

    memset( &iosc, 0, sizeof ( iosc ) );
    if ( DeviceIoControl( hDisk,
                          IOCTL_SCSI_GET_CAPABILITIES,
                          nullptr, 0, &iosc, sizeof ( iosc ), &BytesReturned, nullptr ) == 0 )
    {
        *pScsiCapabilities += L"IOCTL_SCSI_GET_CAPABILITIES failed. ";
        *pScsiCapabilities += Format.SystemError( GetLastError() );
        return;
    }

    if ( BytesReturned == 0 )
    {
        *pScsiCapabilities += L"IOCTL_SCSI_GET_CAPABILITIES bytes = 0.";
        *pScsiCapabilities += PXS_STRING_CRLF;
        return;
    }

    *pScsiCapabilities += Format.StringUInt32( L"Maximum Transfer Length: %%1KB\r\n",
                                               iosc.MaximumTransferLength / 1024 );
    *pScsiCapabilities += Format.StringUInt32( L"Maximum Physical Pages : %%1KB\r\n",
                                               iosc.MaximumPhysicalPages );
    *pScsiCapabilities += PXS_STRING_CRLF;
}

//===============================================================================================//
//  Description:
//      Get the specified disk's inquiry bus data as a formatted string
//
//  Parameters:
//      hDisk            - handle to the disk
//      pScsiInquiryData - receives the formatted string data
//
//  Remarks:
//      As this method returns diagnostic information, any errors will
//      be added to the output string.
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetScsiInquiryData( HANDLE hDisk, String* pScsiInquiryData )
{
    const  DWORD BYTE_BUF_LEN = 16384;        // 16KB for inquiry data
    char   szAnsi[ 256 ] = { 0 };             // Enough for a vendor string
    BYTE*  pOutBuffer = nullptr;
    UCHAR  i = 0;
    DWORD  BytesReturned = 0, counter = 0;
    String Value, Insert1, DataString;
    Formatter     Format;
    AllocateBytes AllocBytes;
    PXSDDK::INQUIRYDATA*   pID   = nullptr;
    SCSI_INQUIRY_DATA*     pSCI  = nullptr;
    SCSI_ADAPTER_BUS_INFO* pSABI = nullptr;

    if ( pScsiInquiryData == nullptr )
    {
        throw ParameterException( L"pScsiInquiryData", __FUNCTION__ );
    }
    DataString.Allocate( 4096 );
    DataString = PXS_STRING_EMPTY;

    if ( ( hDisk == nullptr ) || ( hDisk == INVALID_HANDLE_VALUE ) )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        DataString = Format.String1( L"Invalid handle in '%%1'.\r\n", Insert1 );
        *pScsiInquiryData = DataString;
        return;
    }

    // Create a buffer for output data
    pOutBuffer = AllocBytes.New( BYTE_BUF_LEN );
    memset( pOutBuffer, 0, BYTE_BUF_LEN );
    if ( DeviceIoControl( hDisk,
                          IOCTL_SCSI_GET_INQUIRY_DATA,
                          nullptr, 0, pOutBuffer, BYTE_BUF_LEN, &BytesReturned, nullptr ) == 0 )
    {
        DataString += L"IOCTL_SCSI_GET_INQUIRY_DATA failed. ";
        DataString += Format.SystemError( GetLastError() );
        return;
    }
    pSABI = reinterpret_cast<SCSI_ADAPTER_BUS_INFO*>( pOutBuffer );

    DataString += L"Number of Buses        : ",
    DataString += Format.UInt8( pSABI->NumberOfBuses );
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    // Cycle through buses
    for ( i = 0; i < pSABI->NumberOfBuses; i++ )
    {
        DataString += Format.StringUInt8( L"Bus Number %%1\r\n", i );
        DataString += L"------------\r\n\r\n";

        DataString += L"Logical Units          : ",
        DataString += Format.UInt8( pSABI->BusData[ i ].NumberOfLogicalUnits );
        DataString += PXS_STRING_CRLF;

        DataString += L"Initiator Bus Id       : ",
        DataString += Format.UInt8( pSABI->BusData[ i ].InitiatorBusId );
        DataString += PXS_STRING_CRLF;

        // Print the Inquiry data
        counter = 0;
        pSCI = reinterpret_cast<SCSI_INQUIRY_DATA*>( pOutBuffer +
                                                     pSABI->BusData[ i ].InquiryDataOffset );
        while ( pSCI && ( pSCI->InquiryDataLength != 0 ) )
        {
            DataString += L"Device Number          : ";
            DataString += Format.UInt32( counter );
            DataString += PXS_STRING_CRLF;

            DataString += L"Path Id                : ";
            DataString += Format.UInt8( pSCI->PathId );
            DataString += PXS_STRING_CRLF;

            DataString += L"Target Id              : ";
            DataString += Format.UInt8( pSCI->TargetId );
            DataString += PXS_STRING_CRLF;

            DataString += L"LUN                    : ";
            DataString += Format.UInt8( pSCI->Lun );
            DataString += PXS_STRING_CRLF;

            DataString += L"Device Claimed         : ";
            DataString += Format.UInt8( pSCI->DeviceClaimed );
            DataString += PXS_STRING_CRLF;

            // Inquiry data
            pID = (PXSDDK::INQUIRYDATA*)pSCI->InquiryData;

            Value = PXS_STRING_EMPTY;
            TranslateIdeDeviceType( pID->DeviceType, &Value );
            DataString += L"Device Type            : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            Value = PXS_STRING_EMPTY;
            TranslateDeviceTypeQualifier( pID->DeviceTypeQualifier, &Value );
            DataString += L"Device Type Qualifier  : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            // DeviceTypeModifier
            DataString += L"Device Type Modifier   : ";
            if ( pID->DeviceTypeModifier == DEVICE_CONNECTED )
            {
                DataString += L"Connected";
            }
            else
            {
                DataString += Format.UInt8( pID->DeviceTypeModifier );
            }
            DataString += PXS_STRING_CRLF;

            DataString += L"Removable Media        : ";
            DataString += Format.Int32YesNo( pID->RemovableMedia );
            DataString += PXS_STRING_CRLF;

            StringCchCopyNA( szAnsi,
                             sizeof ( szAnsi ),
                             reinterpret_cast<char*>( pID->VendorId ), sizeof ( pID->VendorId ) );
            Value.SetAnsi( szAnsi );
            DataString += L"Vendor String          : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            // ProductId
            StringCchCopyNA( szAnsi,
                             sizeof ( szAnsi ),
                             reinterpret_cast<char*>( pID->ProductId ), sizeof (pID->ProductId) );
            Value.SetAnsi( szAnsi );
            DataString += L"Product Id             : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            // ProductRevisionLevel
            StringCchCopyNA( szAnsi,
                             sizeof ( szAnsi ),
                             reinterpret_cast<char*>( pID->ProductRevisionLevel ),
                             sizeof ( pID->ProductRevisionLevel ) );
            Value.SetAnsi( szAnsi );
            DataString += L"Product Rev. Level     : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            // VendorSpecific
            Value = PXS_STRING_EMPTY;
            for ( i = 0; i < sizeof ( pID->VendorSpecific ); i++ )
            {
                Value += Format.UInt8Hex( pID->VendorSpecific[ i ], true );
                Value += PXS_CHAR_SPACE;
            }
            DataString += L"Vendor Specific        : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            // Set pointer for next pass
            if ( pSCI->NextInquiryDataOffset )
            {
                pSCI = reinterpret_cast<SCSI_INQUIRY_DATA*>( pOutBuffer +
                                                             pSCI->NextInquiryDataOffset );
            }
            else
            {
                pSCI = nullptr;
            }
            counter++;
        }
    }
    DataString += PXS_STRING_CRLF;
    *pScsiInquiryData = DataString;
}

//===============================================================================================//
//  Description:
//      Do SMART self-assessment test
//
//  Parameters:
//      diskNumber       - byte representing the OS drive number, e.g.
//                         PhysicalDrive0
//      pSmartTestResult - string to receive SMART test result
//
//  Returns:
//      true if got SMART info, else false
//===============================================================================================//
bool DiskInformation::GetSmartTestResult( BYTE physicalDrive, String* pSmartTestResult )
{
    BYTE      diskNumber = 0;
    DWORD     BytesReturned = 0;
    HANDLE    hDisk = INVALID_HANDLE_VALUE;
    String    FileName, Insert1;
    Formatter Format;
    STORAGE_PREDICT_FAILURE SPF;

    if ( pSmartTestResult == nullptr )
    {
        throw ParameterException( L"pSmartTestResult", __FUNCTION__ );
    }
    *pSmartTestResult = PXS_STRING_EMPTY;

    // Map the physical driver number to a controller disk number
    if ( PhysDriveNumToDiskNum( physicalDrive, &diskNumber ) == false )
    {
        Insert1 = Format.UInt8( physicalDrive );
        PXSLogAppInfo1( L"Did not map physical drive %%1 to a disk number.",
                        Insert1 );
        return false;
    }

    // Get a handle to the disk
    FileName  = L"\\\\.\\PhysicalDrive";
    FileName += Format.UInt32( physicalDrive );
    hDisk = CreateFile( FileName.c_str(),
                        GENERIC_READ | GENERIC_WRITE,
                        FILE_SHARE_READ | FILE_SHARE_WRITE,
                        nullptr, OPEN_EXISTING, 0, nullptr );
    if ( hDisk == INVALID_HANDLE_VALUE )
    {
        // Log and return
        PXSLogSysWarn1( GetLastError(), L"CreateFile failed for '%%1'.", FileName );
        return false;
    }
    AutoCloseHandle CloseFileHandle( hDisk );

    if ( DeviceIoControl( hDisk,
                          IOCTL_STORAGE_PREDICT_FAILURE,
                          nullptr, 0, &SPF, sizeof ( SPF ), &BytesReturned, nullptr ) == 0 )
    {
        PXSLogSysWarn1( GetLastError(),
                        L"IOCTL_STORAGE_PREDICT_FAILURE failed for '%%1'.", FileName );
        return false;
    }

    if ( SPF.PredictFailure == 0 )
    {
        *pSmartTestResult = L"OK";
    }
    else
    {
        *pSmartTestResult = Format.StringUInt32( L"Failed with code %%1", SPF.PredictFailure );
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the specified disk's adapter properties as a formatted string
//
//  Parameters:
//      hDisk            - handle to the disk
//      pAdapterProperty - receives the formatted string data
//
//  Remarks:
//      As this method returns diagnostic information, any errors will
//      be added to the output string.
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetStorageAdapterProperty( HANDLE hDisk, String* pAdapterProperty )
{
    const  DWORD BYTE_BUF_LEN = 16384;  // DeviceIoControl output buffer
    BYTE*  pbOutBuffer   = nullptr;
    DWORD  BytesReturned = 0;
    String BusType, Insert1, DataString;
    Formatter     Format;
    AllocateBytes AllocBytes;
    STORAGE_PROPERTY_QUERY spq;
    STORAGE_ADAPTER_DESCRIPTOR* pSAD = nullptr;

    if ( pAdapterProperty == nullptr )
    {
        throw ParameterException( L"pAdapterProperty", __FUNCTION__ );
    }
    DataString.Allocate( 4096 );
    DataString = PXS_STRING_EMPTY;

    if ( ( hDisk == nullptr ) || ( hDisk == INVALID_HANDLE_VALUE ) )
    {
       Insert1.SetAnsi( __FUNCTION__ );
       DataString += Format.String1(L"Invalid handle in %%1.\r\n", Insert1 );
       *pAdapterProperty = DataString;
        return;
    }

    // Set up for DeviceIoControl
    memset( &spq, 0, sizeof ( spq ) );
    spq.PropertyId = StorageAdapterProperty;
    spq.QueryType  = PropertyStandardQuery;
    pbOutBuffer    = AllocBytes.New( BYTE_BUF_LEN );
    if ( DeviceIoControl( hDisk,
                          IOCTL_STORAGE_QUERY_PROPERTY,
                          &spq,
                          sizeof ( spq ),
                          pbOutBuffer, BYTE_BUF_LEN, &BytesReturned, nullptr ) == 0 )
    {
        DataString += Format.SystemError( GetLastError() );
        DataString += L" IOCTL_STORAGE_QUERY_PROPERTY failed.\r\n";
        return;
    }

    if ( BytesReturned == 0 )
    {
        DataString += L"IOCTL_STORAGE_QUERY_PROPERTY returned no data.\r\n";
        return;
    }
    pSAD = reinterpret_cast<STORAGE_ADAPTER_DESCRIPTOR*>( pbOutBuffer );

    DataString += L"Maximum Transfer Length: ";
    DataString += Format.UInt32( pSAD->MaximumTransferLength / 1024 );
    DataString += L"KB";
    DataString += PXS_STRING_CRLF;

    DataString += L"Maximum Physical Pages : ";
    DataString += Format.UInt32( pSAD->MaximumPhysicalPages );
    DataString += PXS_STRING_CRLF;

    DataString += L"Transfer Alignment Mask: ";
    DataString += Format.UInt32Hex( pSAD->AlignmentMask, true );
    DataString += PXS_STRING_CRLF;

    DataString += L"Adapter Uses PIO       : ";
    DataString += Format.Int32YesNo( pSAD->AdapterUsesPio );
    DataString += PXS_STRING_CRLF;

    DataString += L"Adapter Scans Down     : ";
    DataString += Format.Int32YesNo( pSAD->AdapterScansDown );
    DataString += PXS_STRING_CRLF;

    DataString += L"Command Queueing       : ";
    DataString += Format.Int32YesNo( pSAD->CommandQueueing );
    DataString += PXS_STRING_CRLF;

    DataString += L"Accelerated Transfer   : ";
    DataString += Format.Int32YesNo( pSAD->AcceleratedTransfer );
    DataString += PXS_STRING_CRLF;

    DataString += L"Bus Type               : ";
    TranslateStorageBusType( pSAD->BusType, &BusType );
    DataString += BusType;
    DataString += PXS_STRING_CRLF;

    DataString += L"Bus Major Version      : ";
    DataString += Format.UInt32( pSAD->BusMajorVersion );
    DataString += PXS_STRING_CRLF;

    DataString += L"Bus Minor Version      : ";
    DataString += Format.UInt32( pSAD->BusMinorVersion );
    DataString += PXS_STRING_CRLF;

    *pAdapterProperty = DataString;
}

//===============================================================================================//
//  Description:
//      Get the specified disk's properties as a formatted string
//
//  Parameters:
//      hDisk             - handle to the disk
//      pDeviceType        - receives the device type
//      pDeviceTypeModifier- receives the device type modifier
//      pRemovableMedia    - receives if the disk is removable
//      pVendorId          - receives the vendor id string
//      pProductId         - receives the product id string
//      pProductRevision   - receives the product/firmware revision
//      pSerialNumber      - receives the "raw" serial number
//      pBusType           - receives the bus type
//      pErrorMessage      - output error message
//
//  Remarks:
//      Used by diagnostic methods so will avoid unecessary throws and
//      append error messages to the output string
//
//  Returns:
//      true on success, otherwise false and sets an message
//===============================================================================================//
bool DiskInformation::GetStorageDeviceProperties( HANDLE  hDisk,
                                                  String* pDeviceType,
                                                  String* pDeviceTypeModifier,
                                                  String* pRemovableMedia,
                                                  String* pVendorId,
                                                  String* pProductId,
                                                  String* pProductRevision,
                                                  String* pSerialNumber,
                                                  String* pBusType,
                                                  String* pErrorMessage )
{
    const  DWORD BYTE_BUF_LEN = 16384;   // 16KB for DeviceIoControl output
    char*  pOutBuffer    = nullptr;
    DWORD  BytesReturned = 0;
    String Insert1;
    Formatter     Format;
    AllocateChars AllocChar;
    STORAGE_PROPERTY_QUERY spq;
    STORAGE_DEVICE_DESCRIPTOR* pSDD = nullptr;

    if ( ( pDeviceType         == nullptr ) ||
         ( pDeviceTypeModifier == nullptr ) ||
         ( pRemovableMedia     == nullptr ) ||
         ( pVendorId           == nullptr ) ||
         ( pProductId          == nullptr ) ||
         ( pProductRevision    == nullptr ) ||
         ( pSerialNumber       == nullptr ) ||
         ( pBusType            == nullptr ) ||
         ( pErrorMessage       == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pDeviceType         = PXS_STRING_EMPTY;
    *pDeviceTypeModifier = PXS_STRING_EMPTY;
    *pRemovableMedia     = PXS_STRING_EMPTY;
    *pVendorId           = PXS_STRING_EMPTY;
    *pProductId          = PXS_STRING_EMPTY;
    *pProductRevision    = PXS_STRING_EMPTY;
    *pSerialNumber       = PXS_STRING_EMPTY;
    *pBusType            = PXS_STRING_EMPTY;
    *pErrorMessage       = PXS_STRING_EMPTY;

    if ( ( hDisk == nullptr ) || ( hDisk == INVALID_HANDLE_VALUE ) )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        *pErrorMessage = Format.String1( L"Invalid handle in %%1.\r\n",
                                         Insert1 );
        return false;
    }

    // Set up for DeviceIoControl
    memset( &spq, 0, sizeof ( spq ) );
    spq.PropertyId = StorageDeviceProperty;
    spq.QueryType  = PropertyStandardQuery;
    pOutBuffer     = AllocChar.New( BYTE_BUF_LEN );
    if ( DeviceIoControl( hDisk,
                          IOCTL_STORAGE_QUERY_PROPERTY,
                          &spq,
                          sizeof ( spq ),
                          pOutBuffer, BYTE_BUF_LEN,  &BytesReturned, nullptr ) == 0 )
    {
        *pErrorMessage += Format.SystemError( GetLastError() );
        *pErrorMessage += L" IOCTL_STORAGE_QUERY_PROPERTY failed.\r\n";
        return false;
    }

    if ( BytesReturned == 0 )
    {
        *pErrorMessage += L"IOCTL_STORAGE_QUERY_PROPERTY bytes = 0.\r\n";
        return false;
    }
    pSDD = reinterpret_cast<STORAGE_DEVICE_DESCRIPTOR*>( pOutBuffer );

    // Device Type
    TranslateIdeDeviceType( pSDD->DeviceType, pDeviceType );

    // DeviceTypeModifier
    *pDeviceTypeModifier = Format.UInt8( pSDD->DeviceTypeModifier );
    if ( pSDD->DeviceTypeModifier == DEVICE_CONNECTED )
    {
        *pDeviceTypeModifier = L"Connected";
    }

    // Removable Media
    *pRemovableMedia = Format.Boolean( pSDD->RemovableMedia );

    // Vendor Id, its ASCII
    if ( pSDD->VendorIdOffset )
    {
        pVendorId->SetAnsi( pOutBuffer + pSDD->VendorIdOffset );
    }

    // Product Id, its ASCII
    if ( pSDD->ProductIdOffset )
    {
        pProductId->SetAnsi( pOutBuffer + pSDD->ProductIdOffset );
    }

    // Product Revision
    if ( pSDD->ProductRevisionOffset )
    {
        pProductRevision->SetAnsi( pOutBuffer + pSDD->ProductRevisionOffset );
    }

    // Serial Number
    if ( ( pSDD->SerialNumberOffset ) &&
         ( pSDD->SerialNumberOffset != ULONG_MAX ) )  //  -1 means no data
    {
        pSerialNumber->SetAnsi( pOutBuffer + pSDD->SerialNumberOffset );
    }

    // To be consistent with STORAGE_ADAPTER_DESCRIPTOR will treat BusType
    // as a BYTE
    TranslateStorageBusType( static_cast<BYTE>(0xFF & pSDD->BusType), pBusType);

    return true;
}

//===============================================================================================//
//  Description:
//      Use WMI to get an enumeration of the disks/devices as a formatted string
//
//  Parameters:
//      pWmiDiskDriveData - receives the formatted string
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::GetWmiDiskDriveData( String* pWmiDiskDriveData )
{
    size_t i = 0, numElements = 0;
    Wmi    WMI;
    String Line, DataString;
    TArray< NameValue > NameValues;

    if ( pWmiDiskDriveData == nullptr )
    {
        throw ParameterException( L"pWmiDiskDriveData", __FUNCTION__ );
    }
    DataString.Allocate( 4096 );
    DataString = PXS_STRING_EMPTY;

    // Get Win32_DiskDrive
    DataString += PXS_STRING_CRLF;
    DataString += L"----------------------\r\n";
    DataString += L"Disk Drive Enumeration\r\n";
    DataString += L"----------------------\r\n";
    WMI.Connect( L"root\\cimv2" );
    WMI.ExecQuery( L"Select * from Win32_DiskDrive" );
    while ( WMI.Next() )
    {
        DataString += PXS_STRING_CRLF;
        DataString += PXS_STRING_CRLF;
        DataString += L"---------------\r\n";
        DataString += L"Disk Properties\r\n";
        DataString += L"---------------\r\n";

        NameValues.RemoveAll();
        WMI.GetPropertyValues( &NameValues );
        numElements = NameValues.GetSize();
        for ( i = 0; i < numElements; i++ )
        {
            Line = NameValues.Get( i ).GetName();
            Line.FixedWidth( 48, PXS_CHAR_SPACE );
            Line += NameValues.Get( i ).GetValue();
            Line += PXS_STRING_CRLF;
            DataString += Line;
        }
        DataString += PXS_STRING_CRLF;
    }
    WMI.Disconnect();
    DataString += PXS_STRING_CRLF;

    // Win32_PhysicalMedia
    DataString += PXS_STRING_CRLF;
    DataString += L"--------------------------\r\n";
    DataString += L"Physical Media Enumeration\r\n";
    DataString += L"--------------------------\r\n";
    WMI.Connect( L"root\\cimv2" );
    WMI.ExecQuery( L"Select * from Win32_PhysicalMedia" );
    while ( WMI.Next() )
    {
        DataString += PXS_STRING_CRLF;
        DataString += PXS_STRING_CRLF;
        DataString += L"----------------\r\n";
        DataString += L"Media Properties\r\n";
        DataString += L"----------------\r\n";

        NameValues.RemoveAll();
        WMI.GetPropertyValues( &NameValues );
        numElements = NameValues.GetSize();
        for ( i = 0; i < numElements; i++ )
        {
            Line = NameValues.Get( i ).GetName();
            Line.FixedWidth( 48, PXS_CHAR_SPACE );
            Line += NameValues.Get( i ).GetValue();
            Line += PXS_STRING_CRLF;
            DataString += Line;
        }
    }
    WMI.Disconnect();
    DataString += PXS_STRING_CRLF;

    *pWmiDiskDriveData = DataString;
}

//===============================================================================================//
//  Description:
//      Get the disk number as used by the controller from the physical
//      drive number
//
//  Parameters:
//      physicalDrive - the OS drive number, e.g. PhysicalDrive#
//      pDiskNumber   - receives the disk number used by the controller
//
//  Returns:
//      true on success, otherwise false
//===============================================================================================//
bool DiskInformation::PhysDriveNumToDiskNum( BYTE physicalDrive, BYTE* pDiskNumber)
{
    bool result = false;
    int  scsiPort = 0, scsiTargetId = 0, number = 0;
    Wmi  WMI;
    String    Query;
    Formatter Format;

    if ( pDiskNumber == nullptr )
    {
        throw ParameterException( L"pDiskNumber", __FUNCTION__ );
    }
    *pDiskNumber = UINT8_MAX;    // = -1

    // WMI Win32_DiskDrive
    Query  = L"Select * From Win32_DiskDrive WHERE DeviceID='";
    Query += L"\\\\\\\\.\\\\PHYSICALDRIVE";
    Query += Format.UInt8( physicalDrive );
    Query += L"'";
    WMI.Connect( L"root\\cimv2" );
    WMI.ExecQuery( Query.c_str() );
    if ( WMI.Next() )
    {
        scsiPort     = 0;
        scsiTargetId = 0;
        if ( ( WMI.GetInt32( L"SCSIPort"    , &scsiPort     ) ) &&
             ( WMI.GetInt32( L"SCSITargetId", &scsiTargetId ) )  )
        {
            // Two disks per port with the primary at targetid=0 and the
            // secondary at tatgetid=1
            number = PXSMultiplyInt32( 2, scsiPort );
            number = PXSAddInt32( number, scsiTargetId );
            if ( number < 256 )
            {
                *pDiskNumber = PXSCastInt32ToUInt8( number );
                result = true;
            }
        }
    }
    WMI.Disconnect();

    return result;
}

//===============================================================================================//
//  Description:
//      Attempt to identify the manufacturer of the physical disk
//
//  Parameters:
//      Model - the model number
//      pManufacturerName - receives the name of the manufacturer
//
//  Remarks:
//      Be conservative, don't want to show an incorrect name.
//      Use the model number to identify manufacturer
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::ResolveManufacturer( const String& Model, String* pManufacturerName)
{
    if ( pManufacturerName == nullptr )
    {
        throw ParameterException( L"pManufacturerName", __FUNCTION__ );
    }
    *pManufacturerName = PXS_STRING_EMPTY;

    // Check for the manufacturer's name in the model or serial numbers
    if ( Model.IndexOfI( PXS_HDD_MANUFACTURER_FUJITSU ) != PXS_MINUS_ONE )
    {
        *pManufacturerName = PXS_HDD_MANUFACTURER_FUJITSU;
    }
    else if ( Model.IndexOfI( PXS_HDD_MANUFACTURER_HITACHI ) != PXS_MINUS_ONE )
    {
        *pManufacturerName = PXS_HDD_MANUFACTURER_HITACHI;
    }
    else if ( Model.IndexOfI( PXS_HDD_MANUFACTURER_MAXTOR ) != PXS_MINUS_ONE )
    {
        *pManufacturerName = PXS_HDD_MANUFACTURER_MAXTOR;
    }
    else if ( Model.IndexOfI( PXS_HDD_MANUFACTURER_SAMSUNG ) != PXS_MINUS_ONE )
    {
        *pManufacturerName = PXS_HDD_MANUFACTURER_SAMSUNG;
    }
    else if ( Model.IndexOfI( PXS_HDD_MANUFACTURER_SEAGATE ) != PXS_MINUS_ONE )
    {
        *pManufacturerName = PXS_HDD_MANUFACTURER_SEAGATE;
    }
    else if ( Model.IndexOfI( PXS_HDD_MANUFACTURER_WESTERN ) != PXS_MINUS_ONE )
    {
        *pManufacturerName = PXS_HDD_MANUFACTURER_WESTERN;
    }
    else if ( Model.StartsWith( L"ST", false ) )
    {
        // Seagate model numbers usually begin with "ST"
        *pManufacturerName = PXS_HDD_MANUFACTURER_SEAGATE;
    }
    else if ( Model.StartsWith( L"WDC ", false ) )
    {
        // Western digital model numbers usually begin with "WDC "
        *pManufacturerName = PXS_HDD_MANUFACTURER_WESTERN;
    }
}

//===============================================================================================//
//  Description:
//      Translate an INQUIRYDATA DeviceTypeQualifier value
//
//  Parameters:
//      deviceTypeQualifier - the qualifier
//      pTranslation        - receives the translated value
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::TranslateDeviceTypeQualifier( BYTE deviceTypeQualifier,
                                                    String* pTranslation )
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }

    switch( deviceTypeQualifier )   // scsi.h
    {
        default:
            *pTranslation = Format.StringUInt8( L"%%1", deviceTypeQualifier);
            break;

        case DEVICE_QUALIFIER_ACTIVE:
            *pTranslation = L"Active";
            break;

        case DEVICE_QUALIFIER_NOT_ACTIVE:
            *pTranslation = L"Not Active";
            break;

        case DEVICE_QUALIFIER_NOT_SUPPORTED:
            *pTranslation = L"Not Supported";
            break;
    }
}

//===============================================================================================//
//  Description:
//      Translate and IDE device code to a string
//
//  Parameters:
//      DeviceType   - the code
//      pTranslation - receives the translated device type
//
//  Remarks:
//      See "Identifiers for IDE Devices" and STORAGE_DEVICE_DESCRIPTOR
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::TranslateIdeDeviceType( UCHAR DeviceType, String* pTranslation )
{
    size_t    i = 0;
    Formatter Format;

    struct _TYPES
    {
        UCHAR   DeviceType;
        LPCWSTR pszType;
    } Types[] =
        { { DIRECT_ACCESS_DEVICE           , L"Disk"      },
          { SEQUENTIAL_ACCESS_DEVICE       , L"Tape"      },
          { PRINTER_DEVICE                 , L"Printer"   },
          { PROCESSOR_DEVICE               , L"Processor" },
          { WRITE_ONCE_READ_MULTIPLE_DEVICE, L"Worm"      },
          { READ_ONLY_DIRECT_ACCESS_DEVICE , L"CdRom"     },
          { SCANNER_DEVICE                 , L"Scanner"   },
          { OPTICAL_DEVICE                 , L"Optical"   },
          { MEDIUM_CHANGER                 , L"Changer"   },
          { COMMUNICATION_DEVICE           , L"Net"       },
          { 10                             , L"Other"     }
        };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( DeviceType == Types[ i ].DeviceType )
        {
            *pTranslation = Types[ i ].pszType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
       *pTranslation = Format.UInt8( DeviceType );
       PXSLogAppWarn1( L"Unrecognised device type '%%1'.", *pTranslation);
    }
}

//===============================================================================================//
//  Description:
//      Translate a partition type to a string
//
//  Parameters:
//      PartitionType - the partition type
//      pTranslation  - receives the translation
//
//  Remarks:
//      See PARTITION_INFORMATION_EX and MSDN "Disk Partition Types"
//
//      PARTITION_IFS = Installable File System = NTFS and HPFS.
//      HPFS is high performace file system is used in OS/2 which is now
//      non-existant and to be discontinued on 31/12/2006.
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::TranslatePartitionType( BYTE PartitionType, String* pTranslation )
{
    size_t   i = 0;
    Formatter Format;

    struct _TYPES
    {
        BYTE    PartitionType;
        LPCWSTR pszType;
    } Types[] =
        { { PARTITION_ENTRY_UNUSED, L"An unused entry partition."        },
          { PARTITION_EXTENDED    , L"An extended partition."            },
          { PARTITION_FAT_12      , L"A FAT12 file system partition."    },
          { PARTITION_FAT_16      , L"A FAT16 file system partition."    },
          { PARTITION_FAT32       , L"A FAT32 file system partition."    },
          { PARTITION_IFS         , L"An NTFS partition (code=0x07)."    },
          { PARTITION_LDM         , L"A logical disk manager partition." },
          { PARTITION_NTFT        , L"An NTFT partition."                },
          { VALID_NTFT            , L"A valid NTFT partition."           }
        };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( PartitionType == Types[ i ].PartitionType )
        {
            *pTranslation = Types[ i ].pszType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( PartitionType );
        PXSLogAppWarn1( L"Unrecognised partition type %%1.", *pTranslation);
    }
}

//===============================================================================================//
//  Description:
//      Translate a storage bus type code to a string
//
//  Parameters:
//      BusType      - the bus type enumeration code
//      pTranslation - receives the translateed bus type
//
//  Remarks:
//      See STORAGE_BUS_TYPE
//
//      In STORAGE_ADAPTER_DESCRIPTOR, BusType is declared as a BYTE
//      whereas In STORAGE_DEVICE_DESCRIPTOR BusType it is the enumeration
//      type STORAGE_BUS_TYPE. Will treat as a BYTE
//
//  Returns:
//      void
//===============================================================================================//
void DiskInformation::TranslateStorageBusType( BYTE BusType, String* pTranslation )
{
    size_t i = 0;
    Formatter Format;

    // 0x10 = BusTypeSpaces
    struct _TYPE
    {
       BYTE    BusType;
       LPCWSTR pszType;
    } Types[] =
        { { BusTypeUnknown          , L"Unknown"             },
          { BusTypeScsi             , L"SCSI"                },
          { BusTypeAtapi            , L"ATAPI"               },
          { BusTypeAta              , L"ATA"                 },
          { BusType1394             , L"IEEE 1394"           },
          { BusTypeSsa              , L"SSA"                 },
          { BusTypeFibre            , L"Fibre Channel"       },
          { BusTypeUsb              , L"USB"                 },
          { BusTypeRAID             , L"RAID"                },
          { BusTypeiScsi            , L"iSCSI"               },
          { BusTypeSas              , L"Serial-Attached SCSI"},
          { BusTypeSata             , L"SATA"                },
          { BusTypeSd               , L"Secure digital"      },
          { BusTypeMmc              , L"Multimedia card"     },
          { BusTypeVirtual          , L"Virtual"             },
          { BusTypeFileBackedVirtual, L"File backed virtual" },
          { 0x10                    , L"Storage spaces"      } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( BusType == Types[ i ].BusType )
        {
            *pTranslation = Types[ i ].pszType;
            break;
        }
    }

    // Test if unrecognised, enumeration maximum is BusTypeMax
    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( BusType );
        PXSLogAppWarn1( L" Unrecognised bus type '%%'1.", *pTranslation );
    }
}
