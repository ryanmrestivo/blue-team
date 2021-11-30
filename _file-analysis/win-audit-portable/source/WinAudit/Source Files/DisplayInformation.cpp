///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Display/Monitor Information Class Implementation
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
#include "WinAudit/Header Files/DisplayInformation.h"

// 2. C System Files
#include <math.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
DisplayInformation::DisplayInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
DisplayInformation::~DisplayInformation()
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
//      Get information about the display adapters on the system
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetAdapterRecords( TArray< AuditRecord >* pRecords )
{
    int      bitsPixel = 0, logPixelsY = 0, vertRefresh = 0;
    HDC      hdcTop = nullptr;
    DWORD    iDevNum = 0, memoryMB = 0;
    size_t   idxSystem = 0;
    String   Value, RegPath, DeviceKey, AdapterRamMB, LocaleMB, LocaleHz;
    String   LocaleBit, LocaleDPI, ChipType, DacType, AdapterString, AdapterName;
    String   BiosString, ColourDepth, VerticalResolution, Insert2, DeviceNumber, StateFlags;
    Registry RegObject;
    Formatter       Format;
    AuditRecord     Record;
    DISPLAY_DEVICE DisplayDevice;
    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Locale strings
    PXSGetResourceString( PXS_IDS_136_MB  , &LocaleMB  );
    PXSGetResourceString( PXS_IDS_1262_HZ , &LocaleHz  );
    PXSGetResourceString( PXS_IDS_1264_BIT, &LocaleBit );
    PXSGetResourceString( PXS_IDS_1267_DPI, &LocaleDPI );

    RegObject.Connect( HKEY_LOCAL_MACHINE );
    memset( &DisplayDevice, 0, sizeof ( DisplayDevice ) );
    DisplayDevice.cb = sizeof ( DisplayDevice );
    while ( EnumDisplayDevices( nullptr, iDevNum, &DisplayDevice, EDD_GET_DEVICE_INTERFACE_NAME ) )
    {
        // Only want those devices attached to the desk top
        AdapterName = PXS_STRING_EMPTY;
        AdapterName.AppendChars( DisplayDevice.DeviceString, ARRAYSIZE( DisplayDevice.DeviceString ) );
        if ( DISPLAY_DEVICE_ACTIVE & DisplayDevice.StateFlags )
        {
            // Reset values
            RegPath      = PXS_STRING_EMPTY;
            memoryMB     = 0;
            ChipType     = PXS_STRING_EMPTY;
            DacType      = PXS_STRING_EMPTY;
            AdapterString= PXS_STRING_EMPTY;
            BiosString   = PXS_STRING_EMPTY;
            hdcTop       = nullptr;
            bitsPixel    = 0;
            logPixelsY   = 0;
            vertRefresh  = 0;

            // Get the device key, i.e. the location in the registry. The
            // data is in binary format as described in "Setting Hardware
            // Information in the Registry'
            DeviceKey = PXS_STRING_EMPTY;
            DeviceKey.AppendChars( DisplayDevice.DeviceKey, ARRAYSIZE( DisplayDevice.DeviceKey ) );
            idxSystem = DeviceKey.IndexOfI( L"System\\" );
            if ( idxSystem != PXS_MINUS_ONE )
            {
                DeviceKey.SubString( idxSystem, DeviceKey.GetLength() - idxSystem, &RegPath );
                // Memory size, its a ULONG in MB according to documentation,
                // but always seems to be in bytes so if get back a value of
                // at least 1MB assume its bytes
                RegObject.GetBinaryData( RegPath.c_str(),
                                         L"HardwareInformation.MemorySize",
                                         reinterpret_cast<BYTE*>( &memoryMB ),
                                         sizeof ( memoryMB ) );
                if ( memoryMB >= (1024 * 1024) )
                {
                    memoryMB = memoryMB / ( 1024 * 1024 );
                }

                // Chip type
                RegObject.GetValueAsString( RegPath.c_str(),
                                            L"HardwareInformation.ChipType", &ChipType );

                // DAC type
                RegObject.GetValueAsString( RegPath.c_str(),
                                            L"HardwareInformation.DacType", &DacType );
                // Adapter ID
                RegObject.GetValueAsString( RegPath.c_str(),
                                            L"HardwareInformation.AdapterString", &AdapterString );
                // BIOS
                RegObject.GetValueAsString( RegPath.c_str(),
                                            L"HardwareInformation.BiosString", &BiosString );
            }
            else
            {
                // Unexpected format of device key
                Insert2.SetAnsi( __FUNCTION__ );
                PXSLogAppWarn2( L"Unexpected device key '%%1' in '%%2'.", DeviceKey, Insert2 );
            }

            // Get the Desk top device and show some of its settings
            hdcTop = GetDC( nullptr );
            if ( hdcTop )
            {
                bitsPixel  = GetDeviceCaps( hdcTop, BITSPIXEL );
                logPixelsY = GetDeviceCaps( hdcTop, LOGPIXELSY );
                vertRefresh = GetDeviceCaps( hdcTop, VREFRESH );
                ReleaseDC( nullptr, hdcTop );
                hdcTop = nullptr;       // Reset
            }

            // Locale
            AdapterRamMB  = Format.UInt32( memoryMB );
            AdapterRamMB += LocaleMB;
            // Refresh rate, docs 0 or 1 are used to indicate the default rate
            Value = PXS_STRING_EMPTY;
            if ( vertRefresh > 1 )
            {
                Value  = Format.Int32( vertRefresh );
                Value += LocaleHz;
            }

            if ( bitsPixel > 0 )
            {
                ColourDepth  = Format.Int32( bitsPixel );
                ColourDepth += LocaleBit;
            }

            if ( logPixelsY > 0 )
            {
                VerticalResolution  = Format.Int32( logPixelsY );
                VerticalResolution += LocaleDPI;
            }

            // Make the record
            DeviceNumber = Format.UInt32( iDevNum + 1 );
            Record.Reset( PXS_CATEGORY_DISPLAY_ADAPTERS );
            Record.Add( PXS_DISP_ADAPT_ITEM_NUMBER, DeviceNumber );
            Record.Add( PXS_DISP_ADAPT_NAME       , DisplayDevice.DeviceString);
            Record.Add( PXS_DISP_ADAPT_ADAPTER_RAM_MB  , AdapterRamMB );
            Record.Add( PXS_DISP_ADAPT_COLOR_DEPTH_BIT , ColourDepth );
            Record.Add( PXS_DISP_ADAPT_LOGPIXELSY_DPI  , VerticalResolution );
            Record.Add( PXS_DISP_ADAPT_CURR_REFRESH_HZ , Value );
            Record.Add( PXS_DISP_ADAPT_VIDEO_PROCESSOR , ChipType );
            Record.Add( PXS_DISP_ADAPT_ADAPTER_DAC_TYPE, DacType );
            Record.Add( PXS_DISP_ADAPT_ADAPTER_ID      , AdapterString );
            Record.Add( PXS_DISP_ADAPT_BIOS            , BiosString );
            pRecords->Add( Record );
        }
        else
        {
            StateFlags = Format.UInt32( DisplayDevice.StateFlags );
            PXSLogAppInfo2( L"Ignoring display device '%%1' because it is not attached "
                            L"to the desktop. StateFlags = %%2", AdapterName, StateFlags );
        }

        // Reset for next pass
        memset( &DisplayDevice, 0, sizeof ( DisplayDevice ) );
        DisplayDevice.cb = sizeof ( DisplayDevice );
        iDevNum++;
    }
    PXSSortAuditRecords( pRecords, PXS_DISP_ADAPT_NAME );
}

//===============================================================================================//
//  Description:
//      Get a description of the primary display
//
//  Parameters:
//      pDescriptionString - string object to receive the data
//
//  Remarks:
//      Look for product id and physical size of the primary display. If none
//      will return a resolution type string.
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetDescriptionString( String* pDescriptionString )
{
    if ( pDescriptionString == nullptr )
    {
        throw ParameterException( L"pDescriptionString", __FUNCTION__ );
    }
    *pDescriptionString = PXS_STRING_EMPTY;

    try
    {
        GetProductIdSizeString( pDescriptionString );
        if ( pDescriptionString->GetLength() == 0 )
        {
            GetResolutionString( pDescriptionString );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting display product ID data.", e, __FUNCTION__ );
        GetResolutionString( pDescriptionString );
    }
}

//===============================================================================================//
//  Description:
//      Get diagnostic data about the display for writing to a file
//
//  Parameters:
//      pDiagnostics - string to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetDiagnostics( String* pDiagnostics )
{
    size_t    idxSystem = 0, i = 0, j = 0, charLength = 0;
    DWORD     adapterNumber = 0, monitorNumber = 0, memoryMB = 0;
    wchar_t   szAdapterName[ 32 ] = { 0 };  // Max is 15 chars
    String    Title, RegPath, ChipType, DacType, DeviceKey, DataString;
    String    RegDisplay, DisplayKey, RegKey, RegValue, ApplicationName;
    Registry  RegObject;
    Formatter Format;
    StringArray     DeviceKeys, DisplayKeys;
    DISPLAY_DEVICE DisplayAdapter;
    DISPLAY_DEVICE DisplayMonitor;

    if ( pDiagnostics == nullptr )
    {
        throw ParameterException( L"pDiagnostics", __FUNCTION__ );
    }

    DataString.Allocate( 4096 );

    // Start
    PXSGetApplicationName( &ApplicationName );
    Title  = L"Disk, Adapter and Driver Data by ";
    Title += ApplicationName;
    DataString += Title;
    DataString += PXS_STRING_CRLF;
    DataString.AppendChar( '=', Title.GetLength() );
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;


    DataString += L"\r\nDEVICE ENUMERATION\r\n";
    DataString += L"==================\r\n\r\n";

    try
    {
        memset( &DisplayAdapter, 0, sizeof ( DisplayAdapter ) );
        DisplayAdapter.cb = sizeof ( DisplayAdapter );
        adapterNumber = 0;
        while ( EnumDisplayDevices( nullptr,
                                    adapterNumber, &DisplayAdapter, 0 ) )
        {
            DataString += L"Adapter at Index #";
            DataString += Format.UInt32( adapterNumber );
            DataString += PXS_STRING_CRLF;
            DataString += L"-------------------\r\n";
            DataString += PXS_STRING_CRLF;

            DataString += L"Device Name        : ";
            DataString += DisplayAdapter.DeviceName;
            DataString += PXS_STRING_CRLF;

            DataString += L"Device String      : ";
            DataString += DisplayAdapter.DeviceString;
            DataString += PXS_STRING_CRLF;

            DataString += L"Device ID          : ";
            DataString += DisplayAdapter.DeviceID;
            DataString += PXS_STRING_CRLF;

            DataString += L"Device Key         : ";
            DataString += DisplayAdapter.DeviceKey;
            DataString += PXS_STRING_CRLF;

            DataString += L"State Flags        : ";
            DataString += Format.UInt32( DisplayAdapter.StateFlags );
            DataString += PXS_STRING_CRLF;

            DataString += L"Attached To Desktop: ";
            DataString += Format.UInt32YesNo( DISPLAY_DEVICE_ACTIVE
                                                  & DisplayAdapter.StateFlags );
            DataString += PXS_STRING_CRLF;

            DataString += L"Primary Device     : ";
            DataString += Format.UInt32YesNo( DISPLAY_DEVICE_PRIMARY_DEVICE
                                                  & DisplayAdapter.StateFlags );
            DataString += PXS_STRING_CRLF;

            DataString += L"Mirroring Driver   : ";
            DataString += Format.UInt32YesNo( DISPLAY_DEVICE_MIRRORING_DRIVER
                                                  & DisplayAdapter.StateFlags );
            DataString += PXS_STRING_CRLF;

            DataString += L"Modes Pruned       : ";
            DataString += Format.UInt32YesNo( DISPLAY_DEVICE_MODESPRUNED
                                                  & DisplayAdapter.StateFlags );
            DataString += PXS_STRING_CRLF;

            DataString += L"VGA Compatible     : ";
            DataString += Format.UInt32YesNo( DISPLAY_DEVICE_VGA_COMPATIBLE
                                                  & DisplayAdapter.StateFlags );
            DataString += PXS_STRING_CRLF;

            // DISPLAY_DEVICE_REMOVABLE is for Win2000+
            DataString += L"Removable          : ";
            DataString += Format.UInt32YesNo( DISPLAY_DEVICE_REMOVABLE
                                                  & DisplayAdapter.StateFlags );
            DataString += PXS_STRING_CRLF;

            DataString += PXS_STRING_CRLF;

            ////////////////////////////////////////////////////////////////////
            // Add registry data for the adapter

            try
            {
                // Connect
                RegObject.Connect( HKEY_LOCAL_MACHINE );

                // Reset values
                RegPath   = PXS_STRING_EMPTY;
                memoryMB   = 0;
                ChipType  = PXS_STRING_EMPTY;
                DacType   = PXS_STRING_EMPTY;

                // Get the location in the registry, the device key usually
                // begins with Registry\Machine\System\.....
                // See MSDN 'Setting Hardware Information in the Registry'
                DeviceKey  = DisplayAdapter.DeviceKey;
                charLength = DeviceKey.GetLength();
                idxSystem = DeviceKey.IndexOfI( L"System\\" );
                if ( idxSystem < charLength )
                {
                    DeviceKey.SubString( idxSystem,
                                         charLength - idxSystem, &RegPath );
                    // Memory size
                    memoryMB = 0;
                    RegObject.GetBinaryData( RegPath.c_str(),
                                             L"HardwareInformation.MemorySize",
                                             reinterpret_cast<BYTE*>(&memoryMB),
                                             sizeof ( memoryMB ) );
                    // Chip type
                    ChipType = PXS_STRING_EMPTY;
                    RegObject.GetValueAsString( RegPath.c_str(),
                                                L"HardwareInformation.ChipType", &ChipType );
                    // DAC type
                    DacType = PXS_STRING_EMPTY;
                    RegObject.GetValueAsString( RegPath.c_str(),
                                                L"HardwareInformation.DacType", &DacType );
                }
            }
            catch ( const Exception& eRegistry )
            {
                DataString += L"Error getting display adapter data. ";
                DataString += eRegistry.GetMessage();
                DataString += PXS_STRING_CRLF;
            }

            DataString += L"Device Reg. Path   : ";
            DataString += RegPath;
            DataString += PXS_STRING_CRLF;

            // Data is registry seems to be in bytes, not MB as the
            // documentation indicates
            DataString += L"Memory Size        : ";
            DataString += Format.UInt32( memoryMB );
            DataString += PXS_STRING_CRLF;

            DataString += L"Chip Type          : ";
            DataString += ChipType;
            DataString += PXS_STRING_CRLF;

            DataString += L"DAC Type           : ";
            DataString += DacType;
            DataString += PXS_STRING_CRLF;
            DataString += PXS_STRING_CRLF;

            // Enumerate monitors for each adapter
            StringCchCopy( szAdapterName,
                           ARRAYSIZE( szAdapterName ),
                           DisplayAdapter.DeviceName);

            memset( &DisplayMonitor, 0, sizeof ( DisplayMonitor ) );
            DisplayMonitor.cb = sizeof ( DisplayMonitor );
            monitorNumber = 0;
            while ( EnumDisplayDevices( szAdapterName, monitorNumber, &DisplayMonitor, 0 ) )
            {
                DataString += L"\tMonitor at Index #";
                DataString += Format.UInt32( monitorNumber );
                DataString += L" on Adapter ";
                DataString += szAdapterName;
                DataString += PXS_STRING_CRLF;
                DataString += L"\t---------------------------------------";
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;

                DataString += L"\tDevice Name        : ";
                DataString += DisplayMonitor.DeviceName;
                DataString += PXS_STRING_CRLF;

                DataString += L"\tDevice String      : ";
                DataString += DisplayMonitor.DeviceString;
                DataString += PXS_STRING_CRLF;

                DataString += L"\tDevice ID          : ";
                DataString += DisplayMonitor.DeviceID;
                DataString += PXS_STRING_CRLF;

                DataString += L"\tDevice Key         : ";
                DataString += DisplayMonitor.DeviceKey;
                DataString += PXS_STRING_CRLF;

                DataString += L"\tState Flags        : ";
                DataString += Format.BinaryUInt32( DisplayMonitor.StateFlags );
                DataString += PXS_STRING_CRLF;

                DataString += L"\tAttached To Desktop: ";
                DataString += Format.UInt32YesNo( DISPLAY_DEVICE_ACTIVE
                                                  & DisplayMonitor.StateFlags );
                DataString += PXS_STRING_CRLF;

                DataString += L"\tPrimary Device     : ";
                DataString += Format.UInt32YesNo( DISPLAY_DEVICE_PRIMARY_DEVICE
                                                  & DisplayMonitor.StateFlags );
                DataString += PXS_STRING_CRLF;

                DataString += L"\tMirroring Driver   : ";
                DataString += Format.UInt32YesNo(DISPLAY_DEVICE_MIRRORING_DRIVER
                                                  & DisplayMonitor.StateFlags );
                DataString += PXS_STRING_CRLF;

                DataString += L"\tModes Pruned       : ";
                DataString += Format.UInt32YesNo( DISPLAY_DEVICE_MODESPRUNED
                                                  & DisplayMonitor.StateFlags );
                DataString += PXS_STRING_CRLF;

                DataString += L"\tremovable          : ";
                DataString += Format.UInt32YesNo( DISPLAY_DEVICE_REMOVABLE
                                                  & DisplayMonitor.StateFlags );
                DataString += PXS_STRING_CRLF;

                DataString += L"\tVGA Compatible     : ";
                DataString += Format.UInt32YesNo( DISPLAY_DEVICE_VGA_COMPATIBLE
                                                  & DisplayMonitor.StateFlags );
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;

                // Next monitor
                memset( &DisplayMonitor, 0, sizeof ( DisplayMonitor ) );
                DisplayMonitor.cb    = sizeof ( DisplayMonitor );
                monitorNumber++;
            }
            // Next adapter
            memset( &DisplayAdapter, 0, sizeof ( DisplayAdapter ) );
            DisplayAdapter.cb = sizeof ( DisplayAdapter );
            adapterNumber++;
        }
    }
    catch ( const Exception& e )
    {
        DataString += PXS_STRING_CRLF;
        DataString += L"Error enumerating display devices.";
        DataString += PXS_STRING_CRLF;
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }

    ////////////////////////////////////////////////////////////////////////////
    // Get the EDID data for Desktop monitors

    bool   validHeader =false, validCheckSum = false;
    BYTE   edidBlock[ 256 ] = { 0 };      // EDID version 2 is 256 bytes;
    DWORD  errorCode = 0;
    size_t numDisplayKeys = 0, numDeviceKeys = 0;
    String VersionRev, ManufacturerID, ProductID, SerialNumber;
    String ManufacDate, DigDisplay, DisplaySize,  Gamma;
    String DisplayType, MonitorSerial, TextData, MonitorName;
    String Features, EstTimings,  StdTimings, HexData;

    try
    {
        RegDisplay = L"System\\CurrentControlSet\\Enum\\Display\\";

        DataString += L"\r\nEXTENDED DISPLAY IDENTIFICATION DATA\r\n";
        DataString += L"====================================\r\n\r\n";
        DataString += L"\r\nThis section also scans your system for\r\n";
        DataString += L"historic data so may list monitors that \r\n";
        DataString += L"are no longer in use.\r\n\r\n\r\n";

        // Connect to the registry
        RegObject.Connect( HKEY_LOCAL_MACHINE );
        RegObject.GetSubKeyList( RegDisplay.c_str(), &DisplayKeys );

        // Cycle through display keys
        numDisplayKeys = DisplayKeys.GetSize();
        for ( i = 0; i < numDisplayKeys; i++ )
        {
            DisplayKey  = RegDisplay;
            DisplayKey += DisplayKeys.Get( i );
            RegObject.GetSubKeyList( DisplayKey.c_str(), &DeviceKeys );

            // Cycle through the sub-keys
            numDeviceKeys = DeviceKeys.GetSize();
            for ( j = 0; j < numDeviceKeys; j++ )
            {
                RegKey  = DisplayKey;
                RegKey += L"\\";
                RegKey += DeviceKeys.Get( j );

                // Add section for EDID data
                DataString += L"Monitor Data\r\n";
                DataString += L"------------\r\n";
                DataString += PXS_STRING_CRLF;

                // Device Description
                RegValue = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"DeviceDesc", &RegValue );
                DataString += L"Description: ";
                DataString += RegValue;
                DataString += PXS_STRING_CRLF;

                // Hardware ID
                RegValue = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"HardwareID", &RegValue );
                DataString += L"HardwareID : ";
                DataString += RegValue;
                DataString += PXS_STRING_CRLF;

                // Driver
                RegValue = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"Driver", &RegValue );
                DataString += L"Driver     : ";
                DataString += RegValue;
                DataString += PXS_STRING_CRLF;

                // Add the registry key
                DataString += L"Registry   : ";
                DataString += RegKey;
                DataString += PXS_STRING_CRLF;

                // Look for the EDID data block, its location is different
                // if on NT
                RegKey += L"\\";
                RegKey += L"Device Parameters";

                memset( edidBlock, 0, sizeof ( edidBlock ) );
                errorCode = RegObject.GetBinaryData( RegKey.c_str(),
                                                     L"EDID", edidBlock, sizeof ( edidBlock ) );
                if ( ( ERROR_SUCCESS   == errorCode ) ||
                     ( ERROR_MORE_DATA == errorCode )  )
                {
                    DataString += L"EDID Data  : ";
                    DataString += PXS_STRING_CRLF;
                    DataString += PXS_STRING_CRLF;

                    // Interpret EDID data
                    TranslateEdidData( edidBlock   , sizeof ( edidBlock ),
                                       &validHeader, &validCheckSum,
                                       &VersionRev , &ManufacturerID,
                                       &ProductID  , &SerialNumber,
                                       &ManufacDate, &DigDisplay,
                                       &DisplaySize, &Gamma,
                                       &DisplayType, &MonitorSerial,
                                       &TextData   , &MonitorName,
                                       &Features   , &EstTimings,
                                       &StdTimings , &HexData );

                    // Check header bytes, failure does not signify bad data
                    DataString += L"Valid Header         : ";
                    DataString += Format.Int32YesNo( validHeader );
                    DataString += PXS_STRING_CRLF;

                    // Check the checksum, failure does not signify bad data
                    DataString += L"Passed Checksum Test : ";
                    DataString += Format.Int32YesNo( validCheckSum );
                    DataString += PXS_STRING_CRLF;

                    DataString += L"EDID Version         : ";
                    DataString += VersionRev;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"ID Manufacturer Name : ";
                    DataString += ManufacturerID;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"ID Product Code      : ";
                    DataString += ProductID;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"ID Serial Number     : ";
                    DataString += SerialNumber;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Date of Manufacture  : ";
                    DataString += ManufacDate;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Digital Display      : ";
                    DataString += DigDisplay;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Display Size         : ";
                    DataString += DisplaySize;
                    DataString += PXS_STRING_CRLF;

                    // Gamma - stored as = (gamma x 100 ) - 100
                    DataString += L"Display Gamma        : ";
                    DataString += Gamma;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Display type         : ";
                    DataString += DisplayType;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Monitor Serial Num.  : ";
                    DataString += MonitorSerial;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Text Data            : ";
                    DataString += TextData;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Monitor Name         : ";
                    DataString += MonitorName;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Supported Features   : ";
                    DataString += Features;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Established Timing(s): \r\n";
                    DataString += EstTimings;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"Standard Timing(s)   :\r\n";
                    DataString += StdTimings;
                    DataString += PXS_STRING_CRLF;

                    DataString += L"EDID hex data        :";
                    DataString += HexData;
                    DataString += PXS_STRING_CRLF;
                }
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
            }
            DataString += PXS_STRING_CRLF;
            DataString += PXS_STRING_CRLF;
        }   // DisplayKeys Loop
    }
    catch ( const Exception& e )
    {
        DataString += PXS_STRING_CRLF;
        DataString += L"Error getting EDID information.";
        DataString += PXS_STRING_CRLF;
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }
    *pDiagnostics = DataString;
}

//===============================================================================================//
//  Description:
//      Get the display identification data for the monitors attached to
//      the desktop
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetIdentificationRecords( TArray< AuditRecord >* pRecords )
{
    size_t i = 0, numMonitors = 0;
    String MonitorID, DeviceName;
    StringArray MonitorIDs, DeviceNames;
    AuditRecord Record;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetDesktopMonitorIDs( &MonitorIDs, &DeviceNames );
    numMonitors = MonitorIDs.GetSize();
    for ( i = 0; i < numMonitors; i++ )
    {
        MonitorID  = MonitorIDs.Get( i );
        DeviceName = DeviceNames.Get( i );
        Record.Reset( PXS_CATEGORY_DISPLAY_EDID );
        GetDisplayRecord( i + 1, MonitorID, DeviceName, &Record );
        if ( Record.GetNumberOfValues() )
        {
            pRecords->Add( Record );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_DISPLAY_EDID_DISPLAY_NAME );
}

//===============================================================================================//
//  Description:
//      Debug/Testing routine to read the hex data stored in a file
//
//  Parameters:
//      FilePath    - path to EDID data
//      pEdid       - pointer to receive byte data
//      bufferSize  - size of buffer
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::ReadEdidHexFile( const String& FilePath, BYTE* pEdid, size_t bufferSize )
{
    size_t    i = 0, maxBytes = 0;
    DWORD     number = 0;
    File      InputFile;
    String    Line, EdidData, HexByte;
    Formatter Format;

    if ( FilePath.IsEmpty() )
    {
        throw ParameterException( L"FilePath", __FUNCTION__ );
    }

    if ( ( pEdid == nullptr ) || ( bufferSize == 0 ) )
    {
        throw ParameterException( L"pEdid, bufferSize ", __FUNCTION__ );
    }
    memset( pEdid, 0, bufferSize );

    // Read the file
    EdidData.Allocate( 1024 );
    InputFile.OpenText( FilePath );
    while ( InputFile.ReadLine( &Line ) )
    {
        Line.Trim();
        EdidData += Line;
    }
    EdidData.Trim();
    EdidData.ToUpperCase();

    // Hex data is in 2-letter form so there are only half as many bytes
    maxBytes = EdidData.GetLength() / 2;
    if ( maxBytes > bufferSize )
    {
        maxBytes = bufferSize;
    }

    for ( i = 0; i < maxBytes; i++ )
    {
        HexByte = PXS_STRING_EMPTY;
        EdidData.SubString( ( 2 * i ), 2, &HexByte );
        number = Format.HexStringToNumber( HexByte);
        pEdid[ i ] = static_cast<BYTE>( 0xFF & number );
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
//      Get the drawing capabilities of the display as an array of formatted
///     strings
//
//  Parameters:
//      monitorNumber - the number of the monitor, counting starts at one
//      pRecords      - string array to receive the data
//
//  Remarks:
//      Just getting the desktop's capabilities, normally it is the same for all
//      monitors, the. The monitor number is used for reporting purposes.
//
//  Returns:
//      void
//
//===============================================================================================//
void DisplayInformation::GetDesktopCapabilityRecords( DWORD monitorNumber,
                                                      TArray< AuditRecord >* pRecords )
{
    int     deviceCaps = 0;
    HDC     hdc     = nullptr;
    size_t  i       = 0;
    LPCWSTR TEXT    = L"Text";
    LPCWSTR LINE    = L"Line";
    LPCWSTR CURVE   = L"Curve";
    LPCWSTR RASTER  = L"Raster";
    LPCWSTR POLYGON = L"Polygon";
    String  Value;
    Formatter   Format;
    AuditRecord Record;

    // Initialize data display capabilities
    struct  _DISPLAY_CAPS
    {
        LPCWSTR pszType;
        int     index;
        int     flag;
        LPCWSTR pszDesc;
    }  Capabilities[] = {
        { CURVE  , CURVECAPS    , CC_CHORD       ,
          L"Draw chord arcs"                           },
        { CURVE  , CURVECAPS    , CC_CIRCLES     ,
          L"Draw circles"                              },
        { CURVE  , CURVECAPS    , CC_ELLIPSES    ,
          L"Draw chord ellipses"                       },
        { CURVE  , CURVECAPS    , CC_INTERIORS   ,
          L"Draw chord interiors"                      },
        { CURVE  , CURVECAPS    , CC_PIE         ,
          L"Draw pie wedges"                           },
        { CURVE  , CURVECAPS    , CC_ROUNDRECT   ,
          L"Draw rounded rectangles"                   },
        { CURVE  , CURVECAPS    , CC_STYLED      ,
          L"Draw styled borders"                       },
        { CURVE  , CURVECAPS    , CC_WIDE        ,
          L"Draw wide borders"                         },
        { CURVE  , CURVECAPS    , CC_WIDESTYLED  ,
          L"Draw wide styled borders"                  },
        { LINE   , LINECAPS     , LC_POLYLINE    ,
          L"Draw a polyline"                           },
        { LINE   , LINECAPS     , LC_STYLED      ,
          L"Draw styled lines"                         },
        { LINE   , LINECAPS     , LC_WIDE        ,
          L"Draw wide lines"                           },
        { LINE   , LINECAPS     , LC_WIDESTYLED  ,
          L"Draw wide styled lines"                    },
        { LINE   , LINECAPS     , LC_INTERIORS   ,
          L"Draw interiors"                            },
        { LINE   , LINECAPS     , LC_MARKER      ,
          L"Draw a marker"                             },
        { LINE   , LINECAPS     , LC_POLYMARKER  ,
          L"Draw multiple markers"                     },
        { POLYGON, POLYGONALCAPS, PC_INTERIORS   ,
          L"Draw interior polygon"                     },
        { POLYGON, POLYGONALCAPS, PC_POLYGON     ,
          L"Draw alternate-fill polygons"              },
        { POLYGON, POLYGONALCAPS, PC_RECTANGLE   ,
          L"Draw rectangles"                           },
        { POLYGON, POLYGONALCAPS, PC_SCANLINE    ,
          L"Draw a single scanline"                    },
        { POLYGON, POLYGONALCAPS, PC_STYLED      ,
          L"Draw polygon styled border"                },
        { POLYGON, POLYGONALCAPS, PC_WIDE        ,
          L"Draw polygon wide border"                  },
        { POLYGON, POLYGONALCAPS, PC_WIDESTYLED  ,
          L"Draw polygon wide styled border"           },
        { POLYGON, POLYGONALCAPS, PC_WINDPOLYGON ,
          L"Draw winding-fill polygons"                },
        { RASTER , RASTERCAPS   , RC_BANDING     ,
          L"Requires banding support"                  },
        { RASTER , RASTERCAPS   , RC_BITBLT      ,
          L"Can transfer bitmaps"                      },
        { RASTER , RASTERCAPS   , RC_BITMAP64    ,
          L"Support bitmaps over 64 KB"                },
        { RASTER , RASTERCAPS   , RC_DI_BITMAP   ,
          L"Set/get bits in DI bitmap"                 },
        { RASTER , RASTERCAPS   , RC_DIBTODEV    ,
          L"Set DI bits to device"                     },
        { RASTER , RASTERCAPS   , RC_FLOODFILL   ,
          L"Can perform flood fills"                   },
        { RASTER , RASTERCAPS   , RC_PALETTE     ,
          L"Specifies a palette-based device"          },
        { RASTER , RASTERCAPS   , RC_SCALING     ,
          L"Capable of scaling"                        },
        { RASTER , RASTERCAPS   , RC_STRETCHBLT  ,
          L"Can copy and stretch bitmap"               },
        { RASTER , RASTERCAPS   , RC_STRETCHDIB  ,
          L"Can copy and stretch DI bitmap"            },
        { TEXT   , TEXTCAPS     , TC_OP_CHARACTER,
          L"Capable of character output precision"     },
        { TEXT   , TEXTCAPS     , TC_OP_STROKE   ,
          L"Capable of stroke output precision"        },
        { TEXT   , TEXTCAPS     , TC_CP_STROKE   ,
          L"Capable of stroke clip precision"          },
        { TEXT   , TEXTCAPS     , TC_CR_90       ,
          L"Capable of 90-degree character rotation"   },
        { TEXT   , TEXTCAPS     , TC_CR_ANY      ,
          L"Capable of any character rotation"         },
        { TEXT   , TEXTCAPS     , TC_SF_X_YINDEP ,
          L"Scale independently x and y directions"    },
        { TEXT   , TEXTCAPS     , TC_SA_DOUBLE   ,
          L"Capable of doubled character for scaling"  },
        { TEXT   , TEXTCAPS     , TC_SA_INTEGER  ,
          L"Integer multiples character scaling"       },
        { TEXT   , TEXTCAPS     , TC_SA_CONTIN   ,
          L"Any multiples for exact character scaling" },
        { TEXT   , TEXTCAPS     , TC_EA_DOUBLE   ,
          L"Draw double-weight characters"             },
        { TEXT   , TEXTCAPS     , TC_IA_ABLE     ,
          L"Able to italicize"                         },
        { TEXT   , TEXTCAPS     , TC_UA_ABLE     ,
          L"Able to underline"                         },
        { TEXT   , TEXTCAPS     , TC_SO_ABLE     ,
          L"Able to strikeouts"                        },
        { TEXT   , TEXTCAPS     , TC_RA_ABLE     ,
          L"Able to draw raster fonts"                 },
        { TEXT   , TEXTCAPS     , TC_VA_ABLE     ,
          L"Able to draw vector fonts"                 },
        { TEXT   , TEXTCAPS     , TC_SCROLLBLT   ,
          L"Scroll using a bit-block transfer"         } };

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get the device context
    hdc = GetDC( nullptr );
    if ( hdc == nullptr )
    {
        throw SystemException( GetLastError(), L"GetDC", __FUNCTION__ );
    }

    // Catch all exception because need to release the DC
    try
    {
        for ( i = 0; i < ARRAYSIZE( Capabilities ); i++ )
        {
            deviceCaps = GetDeviceCaps( hdc, Capabilities[ i ].index );
            Record.Reset( PXS_CATEGORY_DISPLAY_CAPS );

            Value = Format.UInt32( monitorNumber );
            Record.Add( PXS_DISPLAY_CAPS_DISPLAY_NUMBER, Value );
            Record.Add( PXS_DISPLAY_CAPS_TYPE,  Capabilities[ i ].pszType );
            Record.Add( PXS_DISPLAY_CAPS_DESCRIPTION, Capabilities[i].pszDesc);

            Value = Format.Int32Bool( deviceCaps & Capabilities[ i ].flag );
            Record.Add( PXS_DISPLAY_CAPS_CAPABLE,  Value );

            pRecords->Add( Record );
        }
    }
    catch ( const Exception& )
    {
        ReleaseDC( nullptr, hdc );
        throw;
    }
    ReleaseDC( nullptr, hdc );
}

//===============================================================================================//
//  Description:
//      Get the identifiers of monitors attached to the desktop
//
//  Parameters:
//      pMonitorIDs  - string array to receive the monitor identifiers
//      pDeviceNames - string array to receive the device adaptor names
//
//  Remarks:
//      The returned arrays MonitorIDs and DeviceNames are of the same
//      size.
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetDesktopMonitorIDs( StringArray* pMonitorIDs,
                                               StringArray* pDeviceNames )
{
    DWORD     adapterNumber = 0, monitorNumber = 0;
    String    MonitorIdentifier;
    Formatter Format;
    DISPLAY_DEVICE DisplayAdapter;
    DISPLAY_DEVICE DisplayMonitor;

    if ( ( pMonitorIDs == nullptr ) || ( pDeviceNames == nullptr ) )
    {
        throw ParameterException( L"pMonitorIDs/pDeviceNames", __FUNCTION__);
    }
    pMonitorIDs->RemoveAll();
    pMonitorIDs->RemoveAll();

    // Do not want EDD_GET_DEVICE_INTERFACE_NAME
    memset( &DisplayAdapter, 0, sizeof ( DisplayAdapter ) );
    DisplayAdapter.cb = sizeof ( DisplayAdapter );
    adapterNumber = 0;
    while ( EnumDisplayDevices( nullptr, adapterNumber, &DisplayAdapter, 0 ) )
    {
        // Enumerate monitors for each adapter
        memset( &DisplayMonitor, 0, sizeof ( DisplayMonitor ) );
        DisplayMonitor.cb = sizeof ( DisplayMonitor );
        monitorNumber = 0;
        while ( EnumDisplayDevices( DisplayAdapter.DeviceName, monitorNumber, &DisplayMonitor, 0) )
        {
            // If its a real monitor attached to the desktop add it to the array
            // DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = DISPLAY_DEVICE_ACTIVE = "On"
           if ( (DISPLAY_DEVICE_ACTIVE           & DisplayAdapter.StateFlags) &&
                (DISPLAY_DEVICE_ACTIVE           & DisplayMonitor.StateFlags) &&
               !(DISPLAY_DEVICE_MIRRORING_DRIVER & DisplayAdapter.StateFlags) &&
               !(DISPLAY_DEVICE_MIRRORING_DRIVER & DisplayMonitor.StateFlags)  )
           {
                // Ensure Terminated
                DisplayMonitor.DeviceID[ ARRAYSIZE( DisplayMonitor.DeviceID ) - 1 ] = 0x00;
                DisplayAdapter.DeviceName[ ARRAYSIZE( DisplayAdapter.DeviceName ) - 1 ] = 0x00;
                MonitorIdentifier = DisplayMonitor.DeviceID;
                MonitorIdentifier.Trim();
                if ( MonitorIdentifier.GetLength() )
                {
                    pMonitorIDs->Add( MonitorIdentifier );
                    pDeviceNames->Add( DisplayAdapter.DeviceName );
                }
            }

            // Reset for next monitor loop
            memset( &DisplayMonitor, 0, sizeof ( DisplayMonitor ) );
            DisplayMonitor.cb = sizeof ( DisplayMonitor );
            monitorNumber++;
        }

        // Reset for next adapter loop
        memset( &DisplayAdapter, 0, sizeof ( DisplayAdapter ) );
        DisplayAdapter.cb = sizeof ( DisplayAdapter );
        adapterNumber++;
    }

    // Informational message
    PXSLogAppInfo1( L"Found %%1 monitor(s) attached to the desktop.",
                    Format.SizeT( pMonitorIDs->GetSize() ) );
}

//===============================================================================================//
//  Description:
//      Get the EDID data for a display
//
//  Parameters:
//      monitorNumber - the monitor number, counting starts at 1
//      MonitorID     - the device id as reported by EnumDisplayDevices
//      DeviceName    - the name of the device to which the monitor is attached,
//                      should be of the form \\.\DisplayX, see
//                      EnumDisplaySettings
//      pRecord       - receives the data
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetDisplayRecord( size_t monitorNumber,
                                           const String& MonitorID,
                                           const String& DeviceName, AuditRecord* pRecord )
{
    bool      validHeader = false, validCheckSum = false;
    BYTE      edidBlock[ 512 ] = { 0 };     // EDID v2.0 is 256 bytes
    DWORD     lastError = 0;
    String    VersionRev, ManufacturerID, ProductID, SerialNumber;
    String    ManufacDate, DigDisplay, DisplaySize, Gamma;
    String    DisplayType, MonitorSerial, TextData, MonitorName;
    String    DeviceDesc, Features, EstTimings, StdTimings;
    String    HexData, ManufacturerName, DisplayResolution;
    DEVMODE   DevMode;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_DISPLAY_EDID );

    if ( false == ReadEdidDataBlock( MonitorID, edidBlock, sizeof ( edidBlock ), &DeviceDesc ) )
    {
        PXSLogAppInfo1( L"EDID data not found for monitor '%%1'.",
                       MonitorID );
        return;
    }

    TranslateEdidData( edidBlock   , sizeof ( edidBlock ),
                       &validHeader, &validCheckSum,
                       &VersionRev , &ManufacturerID,
                       &ProductID  , &SerialNumber,
                       &ManufacDate, &DigDisplay,
                       &DisplaySize, &Gamma,
                       &DisplayType, &MonitorSerial,
                       &TextData   , &MonitorName,
                       &Features   , &EstTimings,
                       &StdTimings , &HexData );

    // Check header bytes, failure does not signify bad data
    if ( validHeader == false )
    {
        // Hmm... not the expected header
        PXSLogAppWarn1( L"Invalid EDID header for '%%1.", MonitorID );
    }

    // Check the checksum failure does not signify bad data
    if ( validCheckSum == false )
    {
        PXSLogAppWarn1( L"Invalid EDID checksum for '%%1.", MonitorID );
    }

    // Make the audit record
    pRecord->Add( PXS_DISPLAY_EDID_ITEM_NUMBER, Format.SizeT( monitorNumber ) );

    // Monitor Name
    if ( MonitorName.IsEmpty() )
    {
        MonitorName = DeviceDesc;   // If none use the description
    }
    pRecord->Add( PXS_DISPLAY_EDID_DISPLAY_NAME, MonitorName );

    TranslateIdManufacturerCode( ManufacturerID, &ManufacturerName );
    if ( ManufacturerName.GetLength() )
    {
        pRecord->Add( PXS_DISPLAY_EDID_MANUFACTURER, ManufacturerName );
    }
    else
    {
        pRecord->Add( PXS_DISPLAY_EDID_MANUFACTURER, ManufacturerID );
    }
    pRecord->Add( PXS_DISPLAY_EDID_MANUFAC_DATE, ManufacDate );

    // Serial Number, preference is for the monitor serial number
    if ( MonitorSerial.GetLength() )
    {
        pRecord->Add( PXS_DISPLAY_EDID_SERIAL_NUMBER, MonitorSerial );
    }
    else
    {
        pRecord->Add( PXS_DISPLAY_EDID_SERIAL_NUMBER, SerialNumber );
    }

    pRecord->Add( PXS_DISPLAY_EDID_PRODUCT_ID  , ProductID   );
    pRecord->Add( PXS_DISPLAY_EDID_DISPLAY_SIZE, DisplaySize );
    pRecord->Add( PXS_DISPLAY_EDID_DISPLAY_TYPE, DisplayType );
    pRecord->Add( PXS_DISPLAY_EDID_FEATURES    , Features    );

    // Display size in pixels, a device adaptor could have more than one
    // attached monitor, will use whatever is set into DEVMOD for the
    // the current settings
    DisplayResolution = PXS_STRING_EMPTY;
    if ( DeviceName.GetLength() )
    {
        memset( &DevMode, 0, sizeof ( DevMode ) );
        DevMode.dmSize = sizeof ( DevMode );
        if ( EnumDisplaySettings( DeviceName.c_str(),
                                  ENUM_CURRENT_SETTINGS, &DevMode ) )
        {
            DisplayResolution  = Format.UInt32( DevMode.dmPelsWidth );
            DisplayResolution += L" x ";
            DisplayResolution += Format.UInt32( DevMode.dmPelsHeight );
            DisplayResolution += L" pixels";
        }
        else
        {
            lastError = GetLastError();
            PXSLogSysError1( lastError, L"EnumDisplaySettings failed for '%%1'.", DeviceName );
        }
    }
    pRecord->Add( PXS_DISPLAY_RESOLUTION, DisplayResolution );
}

//===============================================================================================//
//  Description:
//      Get the EDID data of the primary display
//
//  Parameters:
//      pRecord - receives the data
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetPrimaryMonitorRecord( AuditRecord* pRecord )
{
    DWORD   adapterNumber = 0, monitorNumber = 0;
    String  MonitorIdentifier, DeviceName;
    DISPLAY_DEVICE DisplayAdapter;
    DISPLAY_DEVICE DisplayMonitor;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_DISPLAY_EDID );

    // Scan Adapters. Do not want EDD_GET_DEVICE_INTERFACE_NAME
    adapterNumber = 0;
    memset( &DisplayAdapter, 0, sizeof ( DisplayAdapter ) );
    DisplayAdapter.cb = sizeof ( DisplayAdapter );
    while ( MonitorIdentifier.IsEmpty() &&
            EnumDisplayDevices( nullptr, adapterNumber, &DisplayAdapter, 0 ) )
    {
        // Verify the adapter is attached to the desktop (aka Active). Also that
        // it is the primary device and a real one.
        if ( ( DISPLAY_DEVICE_ACTIVE           & DisplayAdapter.StateFlags ) &&
             ( DISPLAY_DEVICE_PRIMARY_DEVICE   & DisplayAdapter.StateFlags ) &&
            !( DISPLAY_DEVICE_MIRRORING_DRIVER & DisplayAdapter.StateFlags )  )
        {
            // Scan this adapter for monitors. Do not want
            // EDD_GET_DEVICE_INTERFACE_NAME
            monitorNumber = 0;
            memset( &DisplayMonitor, 0, sizeof ( DisplayMonitor ) );
            DisplayMonitor.cb = sizeof ( DisplayMonitor );
            while ( EnumDisplayDevices( DisplayAdapter.DeviceName,
                                        monitorNumber, &DisplayMonitor, 0 ) )
            {
                // Test if attached to the desktop (aka Active) and ignore
                // pseudo/mirrored/invisible monitors. Note, it is not necessary
                // to test the primary device bit.
                if ( ( DISPLAY_DEVICE_ACTIVE & DisplayMonitor.StateFlags ) &&
                    !( DISPLAY_DEVICE_MIRRORING_DRIVER & DisplayMonitor.StateFlags ) )
                {
                    DisplayMonitor.DeviceID[ ARRAYSIZE( DisplayMonitor.DeviceID ) - 1 ] = 0x00;
                    MonitorIdentifier = DisplayMonitor.DeviceID;
                    MonitorIdentifier.Trim();
                }

                // Set for the next monitor pass
                memset( &DisplayMonitor, 0, sizeof ( DisplayMonitor ) );
                DisplayMonitor.cb = sizeof ( DisplayMonitor );
                monitorNumber++;
            }
        }

        // Set for the adapter next pass
        memset( &DisplayAdapter, 0, sizeof ( DisplayAdapter ) );
        DisplayAdapter.cb    = sizeof ( DisplayAdapter );
        adapterNumber++;
    }

    // Get the display's data to fill the EDID information
    PXSLogAppInfo1( L"Primary monitor id: '%%1'", MonitorIdentifier );
    if ( MonitorIdentifier.GetLength() )
    {
        DeviceName = PXS_STRING_EMPTY;
        GetDisplayRecord( monitorNumber, MonitorIdentifier, DeviceName, pRecord );
    }
}

//===============================================================================================//
//  Description:
//      Get a combined product id and size of the primary display
//
//  Parameters:
//      pProductIdSizeString - String object to receive the data
//
//  Remarks:
//      Applies to the primary display only, make a string of the
//      form: DEL1234 32cm x 24cm (15.7in);
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetProductIdSizeString( String* pProductIdSizeString )
{
    String ProductId, SizeString, MonitorName, ManufacturerName;
    AuditRecord  Record;

    if ( pProductIdSizeString == nullptr )
    {
        throw ParameterException( L"pProductIdSizeString", __FUNCTION__ );
    }
    GetPrimaryMonitorRecord( &Record );
    if ( Record.GetNumberOfValues() == 0 )
    {
        PXSLogAppInfo( L"Primary monitor data record is empty" );
        return;
    }
    Record.GetItemValue( PXS_DISPLAY_EDID_DISPLAY_NAME, &MonitorName );
    Record.GetItemValue( PXS_DISPLAY_EDID_MANUFACTURER, &ManufacturerName );
    Record.GetItemValue( PXS_DISPLAY_EDID_PRODUCT_ID  , &ProductId );
    Record.GetItemValue( PXS_DISPLAY_EDID_DISPLAY_SIZE, &SizeString );

    // Use name, if none, use the manufacture name but avoid using the
    // 3-digit EISA code
    *pProductIdSizeString = MonitorName;
    if ( pProductIdSizeString->IsEmpty() )
    {
        if ( ManufacturerName.GetLength() > 3 )
        {
            *pProductIdSizeString = ManufacturerName;
        }
    }

    // If none use the product id
    if ( pProductIdSizeString->IsEmpty() )
    {
        *pProductIdSizeString = ProductId;
    }

    // Add to the size description
    if ( SizeString.GetLength() )
    {
        if ( pProductIdSizeString->GetLength() )
        {
            *pProductIdSizeString += L", ";
        }
        *pProductIdSizeString += SizeString;
    }
}

//===============================================================================================//
//  Description:
//      Get the resolution string e.g. 800x600x256colours
//
//  Parameters:
//      ResolutionString - String object to receive the resolution
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::GetResolutionString( String* pResolutionString )
{
    int       planes = 0, bitsPixel = 0;
    HDC       hdc = nullptr;
    String    SizeString, Insert1;
    Formatter Format;

    if ( pResolutionString == nullptr )
    {
        throw ParameterException( L"pResolutionString", __FUNCTION__ );
    }
    *pResolutionString = PXS_STRING_EMPTY;

    hdc = GetDC( nullptr );
    if ( hdc == nullptr )
    {
        throw SystemException( GetLastError(), L"GetDC", __FUNCTION__ );
    }

    // Need to release the handle
    try
    {
        *pResolutionString  = Format.Int32( GetDeviceCaps(  hdc, HORZRES ) );
        *pResolutionString += L" x ";
        *pResolutionString += Format.Int32( GetDeviceCaps(  hdc, VERTRES ) );
        *pResolutionString += L" pixels";

        // Number of colours is = 2^(PLANES * BITSPIXEL)
        bitsPixel = GetDeviceCaps( hdc, BITSPIXEL );
        planes    = GetDeviceCaps( hdc, PLANES    );
        if ( bitsPixel > 0 && planes > 0 )
        {
            // Above 24-bit colour resolution, say "true colour"
            if ( bitsPixel <= ( 24 / planes ) )
            {
                *pResolutionString += L", ";
                *pResolutionString += Format.Int32( 1 << (bitsPixel * planes));
                *pResolutionString += L" colours";
            }
            else
            {
                *pResolutionString += L", true colour";
            }
        }
        else
        {
            Insert1.SetAnsi( __FUNCTION__ );
            PXSLogAppWarn1( L"PLANES and/or BITSPIXEL non-positive in %%1.", Insert1 );
        }
    }
    catch ( const Exception& )
    {
        ReleaseDC( nullptr, hdc );
        throw;
    }
    ReleaseDC( nullptr, hdc );
}

//===============================================================================================//
//  Description:
//      Determine if the specified EDID data block has a valid header
//
//  Parameters:
//      pEdid      - buffer to receive the data
//      bufferSize - sizeof the buffer in bytes
//
//  Returns:
//      true if valid, otherwise false
//===============================================================================================//
bool DisplayInformation::IsValidEdidHeader(const BYTE* pEdid, size_t bufferSize)
{
    bool valid    = false;
    BYTE Header[] = { 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00 };

    if ( ( pEdid == nullptr ) || ( ARRAYSIZE( Header ) > bufferSize ) )
    {
        return false;
    }

    if ( memcmp( pEdid, Header, ARRAYSIZE( Header ) ) == 0 )
    {
        valid = true;
    }

    return valid;
}

//===============================================================================================//
//  Description:
//      Read the EDID data block in the registry
//
//  Parameters:
//      MonitorID    - the monitor identifier
//      pEdid        - buffer to receive the data
//      bufferSize   - sizeof the buffer in bytes
//      pDeviceDesc  - receives the "DeviceDesc" value of the monitor
//
//  Returns:
//      true if read the data, otherwise false
//===============================================================================================//
bool DisplayInformation::ReadEdidDataBlock( const String& MonitorID,
                                            BYTE* pEdid, size_t bufferSize, String* pDeviceDesc )
{
    bool        success = false;
    size_t      i = 0, j = 0, numDisplayKeys = 0, numDeviceKeys = 0;
    DWORD       errorCode = 0;
    String      RegDisplay, InstanceKey, DisplayKey;
    String      DriverValue, ParametersKey;
    Registry    RegObject;
    StringArray DisplayKeys, DeviceKeys;

    if ( MonitorID.IsEmpty() )
    {
        throw ParameterException( L"MonitorID", __FUNCTION__ );
    }

    if ( pDeviceDesc == nullptr )
    {
        throw ParameterException( L"pDeviceDesc", __FUNCTION__ );
    }
    *pDeviceDesc = PXS_STRING_EMPTY;

    // EDID v2 is 256 bytes long
    if ( ( pEdid == nullptr ) || ( bufferSize < 256 ) )
    {
        throw ParameterException( L"pEdid/bufferSize", __FUNCTION__ );
    }
    memset( pEdid, 0, bufferSize );

    // EDID is in the registry
    RegDisplay = L"System\\CurrentControlSet\\Enum\\Display\\";
    RegObject.Connect( HKEY_LOCAL_MACHINE );
    RegObject.GetSubKeyList( RegDisplay.c_str(), &DisplayKeys );
    numDisplayKeys = DisplayKeys.GetSize();
    for ( i = 0; i < numDisplayKeys; i++ )
    {
        // Find a match on driver value
        DisplayKey  = RegDisplay;
        DisplayKey += DisplayKeys.Get( i );
        RegObject.GetSubKeyList( DisplayKey.c_str(), &DeviceKeys );
        numDeviceKeys = DeviceKeys.GetSize();
        for ( j = 0; j < numDeviceKeys; j++ )
        {
            InstanceKey  = DisplayKey;
            InstanceKey += L"\\";
            InstanceKey += DeviceKeys.Get( j );
            DriverValue = PXS_STRING_EMPTY;
            RegObject.GetStringValue( InstanceKey.c_str(), L"Driver", &DriverValue );
            DriverValue.Trim();
            if ( MonitorID.EndsWithStringI( DriverValue.c_str() ) )
            {
                ParametersKey  = InstanceKey;
                ParametersKey += L"\\";
                ParametersKey += L"Device Parameters";
                errorCode = RegObject.GetBinaryData( ParametersKey.c_str(),
                                                     L"EDID", pEdid, bufferSize );
                if ( ( errorCode == ERROR_SUCCESS   ) ||
                     ( errorCode == ERROR_MORE_DATA ) )
                {
                    // Success, get the description as sometimes
                    // the monitor name is not in the EDID block
                    RegObject.GetStringValue( InstanceKey.c_str(), L"DeviceDesc", pDeviceDesc );
                    success = true;
                    break;
                }
            }
        }

        // Exit on success
        if ( success )
        {
            break;
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Translate an EDID data block, the first 128 bytes which corresponds
//      to EDID version 1
//
//  Parameters:
//      pEdid         - pointer to EDID data block
//      bufferSize    - size of the data block, must be at least 128 bytes
//      validHeader   - receives if the header data is invalid
//      validCheckSum - receives if the check sum is valid
//      Strings       - these receive the values
//
//  Remarks:
//      The data block for EDID version 1 is 128 bytes long, for version
//      it is 256 bytes long
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::TranslateEdidData( const BYTE* pEdid   , size_t  bufferSize,
                                            bool*   pValidHeader, bool*   pValidCheckSum,
                                            String* pVersionRev , String* pManufacturerID,
                                            String* pProductID  , String* pSerialNumber,
                                            String* pManufacDate, String* pDigDisplay,
                                            String* pDisplaySize, String* pGamma,
                                            String* pDisplayType, String* pMonitorSerial,
                                            String* pTextData   , String* pMonitorName,
                                            String* pFeatures   , String* pEstTimings,
                                            String* pStdTimings , String* pHexData )
{
    int       code = 0;
    size_t    i    = 0;
    BYTE      xDim = 0, yDim = 0, b1 = 0, b2 = 0;
    double    size = 0.0;
    DWORD     checkSum  = 0, serial = 0, value = 0;
    wchar_t   szID[ 8 ] = { 0 };             // 3-letter manufacturer id
    Formatter Format;

    *pValidHeader    = false;
    *pValidCheckSum  = false;
    *pVersionRev     = PXS_STRING_EMPTY;
    *pManufacturerID = PXS_STRING_EMPTY;
    *pProductID      = PXS_STRING_EMPTY;
    *pSerialNumber   = PXS_STRING_EMPTY;
    *pManufacDate    = PXS_STRING_EMPTY;
    *pDigDisplay     = PXS_STRING_EMPTY;
    *pDisplaySize    = PXS_STRING_EMPTY;
    *pGamma          = PXS_STRING_EMPTY;
    *pDisplayType    = PXS_STRING_EMPTY;
    *pMonitorSerial  = PXS_STRING_EMPTY;
    *pTextData       = PXS_STRING_EMPTY;
    *pMonitorName    = PXS_STRING_EMPTY;
    *pFeatures       = PXS_STRING_EMPTY;
    *pEstTimings     = PXS_STRING_EMPTY;
    *pStdTimings     = PXS_STRING_EMPTY;
    *pHexData        = PXS_STRING_EMPTY;

    if ( pEdid == nullptr )
    {
        return;     // Nothing to do
    }

    // Must have at 128 bytes, will only examine up to 128 bytes
    if ( bufferSize < 128 )
    {
        throw ParameterException( L"bufferSize", __FUNCTION__ );
    }
    bufferSize = 128;        // Limit

    // Check header bytes, failure does not signify bad data
    // Checksum, byte sum should be zero. Failure does not always signify
    // bad data
    *pValidHeader = IsValidEdidHeader( pEdid, bufferSize );
    for ( i = 0; i < bufferSize; i++ )
    {
        checkSum = PXSAddUInt32( checkSum, pEdid[ i ] );
    }
    if ( ( checkSum % 256 ) == 0 )
    {
        *pValidCheckSum = true;
    }

    // EDID Version.Revision
    *pVersionRev  = Format.UInt32( pEdid[ 0x12 ] );
    *pVersionRev += PXS_STRING_DOT;
    *pVersionRev += Format.UInt32( pEdid[ 0x13 ] );

    // Manufacturer Code - three letters, each of 5-bits encoded into two bytes
    // with 1 = A (or 0 = @ = 64). This is the manufacturer's Plug and Play ID.
    // First  letter: bits 14-10
    // Second letter: bits  9-5
    // Third  letter: bits  4-0
    code = ( pEdid[ 0x09 ] + ( pEdid[ 0x08 ] * 0x0100 ) );
    szID[ 0 ] = static_cast<wchar_t>( (0x1F & ( code >> 10 ) ) + 64 );
    szID[ 1 ] = static_cast<wchar_t>( (0x1F & ( code >>  5 ) ) + 64 );
    szID[ 2 ] = static_cast<wchar_t>( (0x1F & ( code >>  0 ) ) + 64 );
    szID[ 3 ] = PXS_CHAR_NULL;
    *pManufacturerID = szID;

    // Product Code
    *pProductID  = Format.UInt8Hex( pEdid[ 0x0B ], false );
    *pProductID += Format.UInt8Hex( pEdid[ 0x0A ], false );

    // Serial Number - express as decimal, zero if not used
    serial = PXSMultiplyUInt32( pEdid[ 0x0C ], 0x00000001 ) +
             PXSMultiplyUInt32( pEdid[ 0x0D ], 0x00000100 ) +
             PXSMultiplyUInt32( pEdid[ 0x0E ], 0x00010000 ) +
             PXSMultiplyUInt32( pEdid[ 0x0F ], 0x01000000 );
    if ( serial )
    {
        *pSerialNumber = Format.UInt32( serial );
    }

    // Week of Manufacture - range is 1 - 54 inclusive
    if ( ( pEdid[ 0x10 ] >= 1 ) && ( pEdid[ 0x10 ] <= 54 ) )
    {
        *pManufacDate  = L"Week ";
        *pManufacDate += Format.UInt8( pEdid[ 0x10 ] );
    }

    // Year of Manufacture - minimum value is 3
    if ( pEdid[ 0x11 ] > 3 )
    {
        if ( pManufacDate->GetLength() > 0 )
        {
            *pManufacDate += L", ";
        }
        *pManufacDate += L"Year ";
        *pManufacDate += Format.Int32( 1990 + pEdid[ 0x11 ] );
    }

    // Digital Display
    *pDigDisplay = PXS_STRING_NO;
    if ( pEdid[ 0x14 ] & ( 1 << 7 ) )
    {
        *pDigDisplay = PXS_STRING_YES;
    }

    // Display size, make sure it looks reasonable, have seen 255*255
    // so test for that
    xDim = pEdid[ 0x15 ];
    yDim = pEdid[ 0x16 ];
    if ( ( xDim > 0    ) &&
         ( yDim > 0    ) &&
         ( xDim < 0xff ) &&
         ( yDim < 0xff )  )
    {
        size = ( xDim * xDim ) + ( yDim * yDim );
        size = pow( size, 0.5 );
        size = ( size / 2.54 );

        *pDisplaySize  = Format.Double( size, 1 );  // 1 decimal place
        *pDisplaySize += L"\" (";
        *pDisplaySize += Format.UInt32( xDim );
        *pDisplaySize += L"cm x ";
        *pDisplaySize += Format.UInt32( yDim );
        *pDisplaySize += L"cm)";
    }

    // Gamma - stored as = (gamma x 100 ) - 100
    *pGamma = Format.Double( ( pEdid[ 0x17 ] / 100.0 ) + 1.0 );

    // Display type, combination of analog, colour mode. When pertaining to an
    // electronic attribute American spelling is correct.
    value = 0;
    if ( pEdid[ 0x18 ] & (1 << 4) )
    {
        value += 2;
    }
    if ( pEdid[ 0x18 ] & (1 << 3) )
    {
        value += 1;
    }

    if ( value == 0 )
    {
        *pDisplayType = L"Monochrome";
    }
    else if ( 1 == value )
    {
        *pDisplayType = L"RGB Color";
    }
    else if ( 2 == value )
    {
        *pDisplayType = L"Non-RGB multicolor";
    }
    else
    {
        *pDisplayType = L"Undefined";
    }

    TranslateEdidDescriptorBlock( pEdid, bufferSize, pMonitorSerial, pTextData, pMonitorName );

    // Power Management
    if ( pEdid[ 0x18 ] & ( 1 << 7 ) )
    {
        *pFeatures += L"Standby";
    }

    if ( pEdid[ 0x18 ] & ( 1 << 6 ) )
    {
        if ( pFeatures->GetLength() )
        {
            *pFeatures += L", ";
        }
        *pFeatures += L"Suspend";
    }

    if ( pEdid[ 0x18 ] & ( 1 << 5 ) )
    {
        if ( pFeatures->GetLength() )
        {
            *pFeatures += L", ";
        }
        *pFeatures += L"Active-Off";
    }

    // Established timings I and II
    LPCWSTR Timings[] = { L" 800 x 600  @ 60Hz",
                          L" 800 x 600  @ 56Hz",
                          L" 640 x 480  @ 75Hz",
                          L" 640 x 480  @ 72Hz",
                          L" 640 x 480  @ 67Hz",
                          L" 640 x 480  @ 60Hz",
                          L" 720 x 400  @ 88Hz",
                          L" 720 x 400  @ 70Hz",
                          L"1280 x 1024 @ 75Hz",
                          L"1024 x 768  @ 75Hz",
                          L"1024 x 768  @ 70Hz",
                          L"1024 x 768  @ 60Hz",
                          L"1024 x 768  @ 87Hz (Interlaced)",
                          L" 832 x 624  @ 75Hz",
                          L" 800 x 600  @ 75Hz",
                          L" 800 x 600  @ 72Hz" };

    for ( i = 0; i < ARRAYSIZE( Timings ); i++ )
    {
        // Bit is set if used
        if ( pEdid[ 0x23 + ( i / 8 ) ] & ( 1 << ( i % 8 ) ) )
        {
            *pEstTimings += PXS_CHAR_TAB;
            *pEstTimings += Timings[ i ];
            *pEstTimings += PXS_STRING_CRLF;
        }
    }

    // Standard Timing Identification
    for ( i = 0x26; i <= 0x34; i+=2 )
    {
        b1 = pEdid[ i ];
        b2 = pEdid[ i + 1 ];

        // 0x0101 means no data
        if ( ( b1 != 0x01) && ( b2 != 0x01 ) )
        {
            // Resolution
            *pStdTimings += L"\tHorizontal Res.: ";
            *pStdTimings += Format.Int32( (b1 * 8) + 248 );

            // Aspect ratio
            value = 0;
            if ( b2 & (1 << 7) )
            {
                value += 2;
            }
            if ( b2 & (1 << 6) )
            {
                value += 1;
            }

            *pStdTimings += L", Aspect Ratio: ";
            if ( value == 0 )
            {
                *pStdTimings += L"16:10";
            }
            else if ( 1 == value )
            {
                *pStdTimings += L"4:3";
            }
            else if ( 2 == value )
            {
                *pStdTimings += L"5:4";
            }
            else
            {
                *pStdTimings += L"16:9";
            }

            // Vertical frequency
            // 0x3F = 00111111
            *pStdTimings += L", Vertical Freq.: ";
            *pStdTimings += Format.Int32( ( b2 & 0x3F ) + 60 );
            *pStdTimings += L"Hz";
            *pStdTimings += PXS_STRING_CRLF;
        }
    }

    // EDID block in hex format
    pHexData->Allocate( 1024 );
    for ( i = 0; i < bufferSize; i++ )
    {
        if ( (i % 8) == 0 )
        {
            *pHexData += PXS_STRING_CRLF;
        }
        *pHexData += Format.UInt8Hex( pEdid[ i ], false );
    }
}

//===============================================================================================//
//  Description:
//      Translate an EDID data block, the first 128 bytes which corresponds
//      to EDID version 1
//
//  Parameters:
//      pEdid         - pointer to EDID data block
//      bufferSize    - size of the data block, must be at least 128 bytes
//      pMonitorSerial- receives the monitor's serial number
//      pTextData     - receives any associated text data
//      pMonitorName  - receives the monitor's name
//
//  Remarks:
//      The descriptor data starts at offset 0x36. There are 4 blocks, each
//      of 18 bytes.
//
//  Returns:
//      void
//===============================================================================================//
void DisplayInformation::TranslateEdidDescriptorBlock( const BYTE* pEdid,
                                                       size_t  bufferSize,
                                                       String* pMonitorSerial,
                                                       String* pTextData,
                                                       String* pMonitorName )
{
    char   szAnsi[ 32 ] = { 0 };    // Big enough for 13 chars + terminator
    BYTE   bTag = 0;
    size_t i = 0, idxStart = 0;
    String DataString;

    if ( pEdid == nullptr )
    {
        return;     // Nothing to do
    }

    // Must have at 126 bytes of data, 0x36 + ( 4 * 18 )
    if ( bufferSize < 126 )
    {
        throw ParameterException( L"bufferSize", __FUNCTION__ );
    }

    if ( ( pMonitorSerial == nullptr ) ||
         ( pTextData      == nullptr ) ||
         ( pMonitorName   == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pMonitorSerial = PXS_STRING_EMPTY;
    *pTextData      = PXS_STRING_EMPTY;
    *pMonitorName   = PXS_STRING_EMPTY;

    // Scan for the blocks
    for ( i = 0; i < 4; i++ )
    {
        // Set start in array
        idxStart = 0x36 + ( i * 18 );  // Each descriptor block is 18 bytes

        // FFh = Monitor Serial Number
        // FEh = ASCII string
        // FDh = Monitor Range Limits,
        // FCh = Monitor name
        // FBh = Colour Point Data
        // FAh = Standard Timing Data
        // F9h = Currently undefined
        // F8h = Defined by manufacturer
        bTag = pEdid[ idxStart + 3 ];

        // Look for tag format 000n0
        memset( szAnsi, 0, sizeof ( szAnsi ) );
        if ( ( pEdid[ idxStart + 0 ] == 0 ) &&
             ( pEdid[ idxStart + 1 ] == 0 ) &&
             ( pEdid[ idxStart + 2 ] == 0 ) &&
             ( pEdid[ idxStart + 4 ] == 0 )  )
        {
            // pEdid has ASCII data comprising of line-feed separated fields.
            // The maximum length is 13 characters as its starts at offset 5
            // and a descriptor block is 18 characters long
            memcpy( szAnsi, pEdid + idxStart + 5, 13 );
            szAnsi[ sizeof ( szAnsi ) - 1 ] = 0x00;
            DataString.SetAnsi( szAnsi );
            DataString.ReplaceChar( 0x0A, PXS_CHAR_NULL );
            if ( bTag == 0xFF )
            {
                *pMonitorSerial = DataString;
            }
            else if ( bTag == 0xFE )
            {
                // Concat to existing data
                *pTextData += DataString;
            }
            else if ( bTag == 0xFC )
            {
                // Concat to existing data
                *pMonitorName += DataString;
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a 3 letter ID to a Manufacturers Name
//
//  Parameters:
//      PnPID             - the Plug and Play ID
//      pManufacturerName - receives the manufacturer's name
//
//  Remarks:
//      The 3 character serial number in the EDID data block is the display's
//      manufacturers Plug and Play Identifier.
//
//      Plug and Play ID's are maintained by Microsoft, see "Plug and Play
//      ID - PNPID Request". A list is available for download:
//      http://download.microsoft.com/download/7/E/7/7E7662CF-CBEA-470B-A97E-CE7CE0D98DC2/ISA_PNPID_List.xlsx
//      The last update 3/Feb/2013 which contains 2249 companies
//
//  Returns:
//      void, zero length string "" if not identified
//===============================================================================================//
void DisplayInformation::TranslateIdManufacturerCode( const String& PnPID,
                                                      String* pManufacturerName)
{
    size_t i = 0;

    // Common manufacturer names
    struct _NAMES
    {
        LPCWSTR pszPnPID;
        LPCWSTR pszManufacturer;
    } Names[] =
        { { L"TCM", L"3Com Corp"                    },
          { L"TRU", L"Aashima Technology B.V."      },
          { L"ADV", L"AMD"                          },
          { L"ACC", L"Accton Technology Corp"       },
          { L"AMP", L"Amptron"                      },
          { L"ACR", L"Acer Corporation"             },
          { L"AXS", L"Analog Devices"               },
          { L"ANX", L"Acer Netxus Inc."             },
          { L"AOC", L"AOC International (USA) Ltd." },
          { L"ACS", L"ACS"                          },
          { L"APP", L"Apple Computer"               },
          { L"ACT", L"Actebis"                      },
          { L"ART", L"Artmedia"                     },
          { L"ADP", L"Adaptec"                      },
          { L"AST", L"AST"                          },
          { L"ADD", L"Addtron"                      },
          { L"ATI", L"ATI"                          },
          { L"ADI", L"ADI Systems Inc"              },
          { L"ATK", L"ATKK"                         },
          { L"ALS", L"Advance Logic (ADL) Inc"      },
          { L"AZT", L"Aztech"                       },
          { L"ASC", L"AdvanSys"                     },
          { L"AZU", L"Azura"                        },
          { L"AMB", L"Ambient Technologies Inc"     },
          { L"API", L"Benq Corp"                    },
          { L"BRG", L"Bridge"                       },
          { L"BDP", L"BestData"                     },
          { L"BUS", L"Buslogic"                     },
          { L"BRI", L"Boca"                         },
          { L"CMI", L"C-Media Electronics Inc"      },
          { L"CRN", L"Cornerstone Imaging"          },
          { L"CSI", L"Cabletron"                    },
          { L"CPI", L"CPI"                          },
          { L"CAC", L"Cardinal"                     },
          { L"CTL", L"Creative"                     },
          { L"CTX", L"Chuntex Electronic"           },
          { L"CSC", L"Crystal"                      },
          { L"CIR", L"Cirrus Logic"                 },
          { L"CYB", L"CyberVision"                  },
          { L"CPL", L"Compal Electronics"           },
          { L"CPQ", L"Compaq"                       },
          { L"CRX", L"Cyrix Corporation"            },
          { L"WCI", L"Conexant Systems"             },
          { L"DWE", L"Daewoo"                       },
          { L"SUP", L"Diamond Multimedia"           },
          { L"DBK", L"Databook"                     },
          { L"DMB", L"Digicom System"               },
          { L"OEC", L"Daytek"                       },
          { L"DRT", L"Digital Research"             },
          { L"DVC", L"DecaView"                     },
          { L"DPT", L"DPT"                          },
          { L"DEL", L"Dell Computer Corp"           },
          { L"DPC", L"Delta Electronics"            },
          { L"EIZ", L"Eizo"                         },
          { L"EPI", L"EnVision Inc"                 },
          { L"CCP", L"Epson"                        },
          { L"ECS", L"ELITEGROUP Computer Systems"  },
          { L"ESS", L"ESS Technology Inc"           },
          { L"ELS", L"ELSA GmbH"                    },
          { L"EPI", L"EPI"                          },
          { L"FAR", L"Farallon"                     },
          { L"FPA", L"Fujutsu"                      },
          { L"FUJ", L"Fujitsu"                      },
          { L"FCM", L"Funai"                        },
          { L"ICL", L"Fujitsu ICL"                  },
          { L"FDC", L"Future Domain"                },
          { L"FUS", L"Fujitsu Siemens"              },
          { L"GVT", L"G-Vision"                     },
          { L"GWY", L"Gateway 2000"                 },
          { L"HSL", L"Hansol Electronics"           },
          { L"HIT", L"Hitachi"                      },
          { L"HAY", L"Hayes Microcomputer Products" },
          { L"HEI", L"Hyundai Electronics"          },
          { L"HCM", L"HCL Peripherals"              },
          { L"CPQ", L"Hewlett Packard"              },
          { L"IOD", L"I-O Data"                     },
          { L"INT", L"Intel Corporation"            },
          { L"IBM", L"IBM"                          },
          { L"IXD", L"Intertex"                     },
          { L"IVM", L"Idek Iiyama North America"    },
          { L"ISA", L"Iomega"                       },
          { L"IQT", L"ImageQuest"                   },
          { L"KTR", L"IMRI"                         },
          { L"JEN", L"Jean"                         },
          { L"KFC", L"KFC Computek"                 },
          { L"KOR", L"KXPro"                        },
          { L"KTC", L"Kingston Technology Corp"     },
          { L"KYE", L"KYE Systems Corp"             },
          { L"DWT", L"Korea Data Systems"           },
          { L"LEN", L"Lenovo"                       },
          { L"LEO", L"LEO Systems"                  },
          { L"LTN", L"Lite-on Technology Corp."     },
          { L"GSM", L"LG Electronics Inc."          },
          { L"SLI", L"LSI Logic"                    },
          { L"LKM", L"Likom"                        },
          { L"LNK", L"LINK Technologie"             },
          { L"MDG", L"Madge"                        },
          { L"MIR", L"Miro Computer Products AG"    },
          { L"PMV", L"MAG Technology Co"            },
          { L"MTC", L"Mitac"                        },
          { L"MAX", L"Maxdata Computer GmbH"        },
          { L"MEL", L"Mitsubishi Electronics"       },
          { L"MED", L"Medion"                       },
          { L"MDD", L"Modis"                        },
          { L"MDY", L"Microdyne"                    },
          { L"MOT", L"Motorola"                     },
          { L"MS_", L"Microsoft Corp"               },
          { L"NAN", L"Nanao"                        },
          { L"NLM", L"Newcomm"                      },
          { L"NCD", L"NCD"                          },
          { L"NOK", L"Nokia"                        },
          { L"ACU", L"NCR"                          },
          { L"NVL", L"Novell/Anthem"                },
          { L"NMX", L"Neomagic Corp"                },
          { L"NCL", L"Netcomm"                      },
          { L"OKI", L"OKI"                          },
          { L"OPT", L"OPTi Inc"                     },
          { L"OLC", L"Olicom"                       },
          { L"OQI", L"Optiquest"                    },
          { L"OLI", L"Olivett"                      },
          { L"PMC", L"Pace"                         },
          { L"PGS", L"Princeton Graphic Systems"    },
          { L"NCI", L"Packard Bell"                 },
          { L"PRO", L"Proteon"                      },
          { L"MEI", L"Panasonic"                    },
          { L"PEI", L"Proton"                       },
          { L"PEA", L"Peacock"                      },
          { L"BMM", L"Proview"                      },
          { L"PLB", L"Philips Consumer Electronics" },
          { L"EMC", L"ProView (EMC)"                },
          { L"PPI", L"Practical Peripherals"        },
          { L"RII", L"Racal"                        },
          { L"TRL", L"Royal Information Company"    },
          { L"RDS", L"Radius (KDS)"                 },
          { L"RPT", L"RPTI"                         },
          { L"RTL", L"Realtek Semiconductor Corp"   },
          { L"REL", L"Relisys"                      },
          { L"SAM", L"Samsung"                      },
          { L"XOC", L"Sky Wide Technology"          },
          { L"STN", L"Samtron"                      },
          { L"SMC", L"SMC"                          },
          { L"SIB", L"Sanyo"                        },
          { L"SML", L"Smile Technologies"           },
          { L"SPT", L"Sceptre"                      },
          { L"SNY", L"Sony Corporation"             },
          { L"SCM", L"SCM"                          },
          { L"STA", L"Stesa"                        },
          { L"SRC", L"Shamrock"                     },
          { L"SUN", L"Sun"                          },
          { L"SHP", L"Sharp Corp"                   },
          { L"SVE", L"SVEC"                         },
          { L"SHT", L"SHINHO"                       },
          { L"SYL", L"Sylvania"                     },
          { L"SIE", L"Siemens"                      },
          { L"SYN", L"Synaptics"                    },
          { L"SNI", L"Siemens Nixdorf"              },
          { L"SKD", L"SysKonnect"                   },
          { L"SSC", L"Sierra Semiconductor"         },
          { L"SGX", L"Silicon Graphics"             },
          { L"TAT", L"Tatung"                       },
          { L"TVP", L"Top Victory Electronics"      },
          { L"TAX", L"Taxan"                        },
          { L"TOS", L"Toshiba"                      },
          { L"TEA", L"Teac"                         },
          { L"TTK", L"Totoku / TeleVideo"           },
          { L"TEI", L"TECO"                         },
          { L"TVD", L"Trans Video Deutschland"      },
          { L"TEO", L"Teco Information Systems"     },
          { L"TTX", L"TTX"                          },
          { L"TVI", L"TeleVideo"                    },
          { L"TCI", L"Tulip"                        },
          { L"TER", L"TerraTec Electronic GmbH"     },
          { L"TVM", L"TVM"                          },
          { L"TCO", L"Thomas-Conrad Corp"           },
          { L"USC", L"UltraStor"                    },
          { L"UTB", L"Utobia"                       },
          { L"UMC", L"UMC Electronics Co Ltd"       },
          { L"UNM", L"Unisys Corporation"           },
          { L"AMT", L"V7 Video Seven"               },
          { L"VSC", L"ViewSonic Corporation"        },
          { L"VDM", L"Vadem"                        },
          { L"VOB", L"VOBIS"                        },
          { L"VES", L"Vestel"                       },
          { L"VIA", L"VIA Technologies Inc"         },
          { L"WTC", L"Wen Technology"               },
          { L"WEC", L"Winbond Electronics Corp"     },
          { L"DNV", L"Xenon"                        },
          { L"YMH", L"YAMAHA Corp"                  },
          { L"XDM", L"Yi Ray Electronic Co."        },
          { L"ZCM", L"Zenith Data Systems"          },
          { L"OZO", L"Zoom Telephonics Inc"         },
          { L"ZDS", L"Zeos"                         },
          { L"ZIN", L"Zhi-Ying Display Equipment"   }
        };

    if ( pManufacturerName == nullptr )
    {
        throw ParameterException( L"pManufacturerName", __FUNCTION__ );
    }
    *pManufacturerName = PXS_STRING_EMPTY;

    // List is unsorted so will do a linear search
    for ( i = 0; i < ARRAYSIZE( Names ); i++ )
    {
        if ( PnPID.CompareI( Names[ i ].pszPnPID ) == 0 )
        {
            *pManufacturerName = Names[ i ].pszManufacturer;
            break;
        }
    }

    if ( pManufacturerName->IsEmpty() )
    {
        PXSLogAppWarn1( L"Unrecognised PnP ID '%%1'.", PnPID );
    }
}
