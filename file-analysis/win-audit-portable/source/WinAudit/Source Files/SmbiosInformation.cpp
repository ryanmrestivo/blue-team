///////////////////////////////////////////////////////////////////////////////////////////////////
//
// System Management BIOS Information Class Implementation
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
#include "WinAudit/Header Files/SmbiosInformation.h"

// 2. C System Files
#include <math.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"
#include "PxsBase/Header Files/Wmi.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/Ddk.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
SmbiosInformation::SmbiosInformation()
                  :m_SmBiosData()
{
    // Initialize class scope data members
    memset( &m_SmBiosData, 0, sizeof ( m_SmBiosData ) );
}

// Copy constructor - not allowed so no implementation

// Destructor
SmbiosInformation::~SmbiosInformation()
{
    // Clean up
    if ( m_SmBiosData.pDataTable )
    {
        delete[] m_SmBiosData.pDataTable;
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
//      Get the Chassis Asset Tag
//
//  Parameters:
//      pRecords - array to receive the records
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetAuditRecords( BYTE structureType, TArray< AuditRecord >* pRecords )
{
    WORD structureNumber = 1;   // Counting starts at zero
    AuditRecord Record;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    if ( m_SmBiosData.validData == false )
    {
        ReadSmbiosData();
    }

    // Read each SMBIOS entry
    while ( GetSmbiosRecord( structureType, structureNumber, &Record ) )
    {
        pRecords->Add( Record );
        structureNumber = PXSAddUInt16( structureNumber, 1 );
    }
}

//===============================================================================================//
//  Description:
//      Get the Chassis Asset Tag
//
//  Parameters:
//      pChassisAssetTag - string object to receive data
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetChassisAssetTag( String* pChassisAssetTag ) const
{
    GetStringValue( PXS_SMBIOS_CHASSIS_ASSET_TAG_NUM, 1, pChassisAssetTag );
}

//===============================================================================================//
//  Description:
//      Get the name of an SMBIOS Item
//
//  Parameters:
//      itemID   - defined constant identifying the item
//      pItemName- string object to receive the name
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetItemName( DWORD itemID, String* pItemName ) const
{
    PXS_TYPE_SMBIOS_SPECIFICATION SmbiosSpec;

    if ( pItemName == nullptr )
    {
        throw ParameterException( L"pItemName", __FUNCTION__ );
    }
    memset( &SmbiosSpec, 0, sizeof ( SmbiosSpec ) );
    GetItemSpecification( itemID, &SmbiosSpec );
    *pItemName = SmbiosSpec.szName;
    pItemName->Trim();
}

//===============================================================================================//
//  Description:
//      Get the Product Name in the SMBIOS
//
//  Parameters:
//      pProductName - string object to receive data
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetProductName( String* pProductName ) const
{
    GetStringValue( PXS_SMBIOS_SYSINFO_PRODUCT_NAME, 1, pProductName );
}

//===============================================================================================//
//  Description:
//      Get an SMBIOS type as an audit record
//
//  Parameters:
//      structureType   - named constant identifying the SMBIOS
//                        structure type
//      structureNumber - the structure number, one-based
//      pRecord         - receives the audit record
//
//  Returns:
//      true if particular instance of SMBIOS type was found else false
//===============================================================================================//
bool SmbiosInformation::GetSmbiosRecord( BYTE structureType,
                                         WORD structureNumber, AuditRecord* pRecord ) const
{
    bool success = false;

    // Must have read the SMBIOS data
    if ( m_SmBiosData.validData == false )
    {
        throw FunctionException( L"m_SmBiosData.validData", __FUNCTION__ );
    }

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_UKNOWN );

    switch ( structureType )
    {
        default:
            throw ParameterException( L"structureType", __FUNCTION__ );

        case PXS_SMBIOS_TYPE_0_BIOS:

            // BIOS Information (Type 0) only has one structure
            if ( structureNumber == 1 )
            {
                success = GetType0_BiosRecord( pRecord );
            }
            break;

        case PXS_SMBIOS_TYPE_1_SYSTEM:

            // System Information (Type 1) only has one structure
            if ( structureNumber == 1 )
            {
                success = GetType1_SystemRecord( pRecord );
            }
            break;

        case PXS_SMBIOS_TYPE_2_BASE_BOARD:
            success = GetType2_BaseBoardRecord( structureNumber, pRecord );
            break;

        case PXS_SMBIOS_TYPE_3_CHASSIS:

            // System Enclosure or Chassis (Type 3) only has one structure
            if ( structureNumber == 1 )
            {
                success = GetType3_ChassisRecord( pRecord );
            }
            break;

        case PXS_SMBIOS_TYPE_4_PROCESSOR:
            success = GetType4_ProcessorRecord( structureNumber, pRecord );
            break;

        case PXS_SMBIOS_TYPE_5_MEM_CONTROL:
            success = GetType5_MemoryControllerRecord( structureNumber, pRecord );
            break;

        case PXS_SMBIOS_TYPE_6_MEMORY_MODULE:
            success = GetType6_MemoryModuleRecord( structureNumber, pRecord );
            break;

        case PXS_SMBIOS_TYPE_7_CPU_CACHE:
            success = GetType7_CpuCacheRecord( structureNumber, pRecord );
            break;

        case PXS_SMBIOS_TYPE_8_PORT_CONN:
            success = GetType8_PortConnectorRecord( structureNumber, pRecord );
            break;

        case PXS_SMBIOS_TYPE_9_SYSTEM_SLOT:
            success = GetType9_SystemSlotRecord( structureNumber, pRecord );
            break;

        case PXS_SMBIOS_TYPE_16_MEM_ARRAY:
            success = GetType16_MemoryArrayRecord( structureNumber, pRecord );
            break;

        case PXS_SMBIOS_TYPE_17_MEMORY_DEVICE:
            success = GetType17_MemoryDeviceRecord( structureNumber, pRecord );
            break;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Get system management BIOS version number
//
//  Parameters:
//      pMajorDotMinor - string to receive major.minor version
//
//  Remarks:
//      Must have read the SMBIOS data first
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetSmbiosMajorDotMinor( String* pMajorDotMinor ) const
{
    Formatter Format;

    if ( pMajorDotMinor == nullptr )
    {
        throw ParameterException( L"pMajorDotMinor", __FUNCTION__ );
    }

    // Must have read the SMBIOS data
    if ( m_SmBiosData.validData == false )
    {
        throw FunctionException( L"m_SmBiosData.validData", __FUNCTION__ );
    }
    *pMajorDotMinor  = Format.UInt8( m_SmBiosData.majorVersion );
    *pMajorDotMinor += PXS_CHAR_DOT;
    *pMajorDotMinor += Format.UInt8( m_SmBiosData.minorVersion );
}

//===============================================================================================//
//  Description:
//      Get the offset in the SMBIOS data of the specified structure type
//      and structure number
//
//  Parameters:
//      structureType   - a byte denoting the BIOS type number
//      structureNumber - one-based occurrence number of this BIOS type
//      pStructureOffset- receives the array offset if the structure
//
//  Returns:
//      true if found structure otherwise false
//===============================================================================================//
bool SmbiosInformation::GetStructureOffset( BYTE structureType,
                                            WORD structureNumber, WORD* pStructureOffset ) const
{
    bool found  = false;
    WORD offset = 0, counter = 0, length = 0;

    if ( pStructureOffset == nullptr )
    {
        throw ParameterException( L"pStructureOffset", __FUNCTION__ );
    }
    *pStructureOffset = 0;

    // Must have read the SMBIOS data
    if ( m_SmBiosData.validData == false )
    {
        throw FunctionException( L"m_SmBiosData.validData", __FUNCTION__ );
    }

    // Scan for the specified occurrence of the SMBIOS type. Some BIOS data
    // tables start with OEM types (i.e above 127).
    while ( ( found == false ) &&
            ( PXSAddUInt16( offset, 2 ) < m_SmBiosData.tableLength ) )
    {
        // Keep track of the number of times have found the structure type
        if ( structureType == m_SmBiosData.pDataTable[ offset ] )
        {
            counter++;
            if ( counter == structureNumber )
            {
                found = true;
                *pStructureOffset = offset;
            }
        }

        // Advance to end of formatted part of the structure
        length = m_SmBiosData.pDataTable[ offset + 1 ];
        if ( ( length == 0 ) ||
             ( PXSAddUInt16( offset, length ) > m_SmBiosData.tableLength ) )
        {
            break;  // Data is bad.
        }
        else
        {
            offset = PXSAddUInt16( offset, length );
        }

        // Advance past any strings at the end of the formatted structure
        // to the start of the next structure
        while ( PXSAddUInt16( offset, 2 ) < m_SmBiosData.tableLength )
        {
            if ( ( m_SmBiosData.pDataTable[ offset     ] == 0x00 ) &&
                 ( m_SmBiosData.pDataTable[ offset + 1 ] == 0x00 )  )
            {
                offset++;
                offset++;
                break;
            }
            else
            {
                offset++;
            }
        }
    }

    return found;
}

//===============================================================================================//
//  Description:
//      Get the System Manufacturer
//
//  Parameters:
//      pSystemManufacturer - string object to receive data
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetSystemManufacturer( String* pSystemManufacturer ) const
{
    if ( pSystemManufacturer == nullptr )
    {
        throw ParameterException( L"pSystemManufacturer", __FUNCTION__ );
    }
    *pSystemManufacturer = PXS_STRING_EMPTY;

    // Try in System structure (Type 1)
    GetStringValue( PXS_SMBIOS_SYSINFO_MANUFACTURER, 1, pSystemManufacturer );
    pSystemManufacturer->Trim();

    // Try in Chassis structure (Type 3)
    if ( pSystemManufacturer->IsEmpty() )
    {
        GetStringValue( PXS_SMBIOS_CHASSIS_MANUFACTURER,
                        1, pSystemManufacturer );
        pSystemManufacturer->Trim();
    }
}

//===============================================================================================//
//  Description:
//      Get the System Serial Number Tag
//
//  Parameters:
//      pSystemSerialNumber - string object to receive data
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetSystemSerialNumber( String* pSystemSerialNumber ) const
{
    if ( pSystemSerialNumber == nullptr )
    {
        throw ParameterException( L"pSystemSerialNumber", __FUNCTION__ );
    }
    *pSystemSerialNumber = PXS_STRING_EMPTY;

    // Try in System structure (Type 1)
    GetStringValue( PXS_SMBIOS_SYSINFO_SERIAL_NUMBER, 1, pSystemSerialNumber );
    pSystemSerialNumber->Trim();

    // Try in Chassis structure (Type 3)
    if ( pSystemSerialNumber->IsEmpty() )
    {
        GetStringValue( PXS_SMBIOS_CHASSIS_SERIAL_NUMBER, 1, pSystemSerialNumber );
        pSystemSerialNumber->Trim();
    }
}

//===============================================================================================//
//  Description:
//      Get the System's Universal Unique Identifier Serial Number Tag
//
//  Parameters:
//      pSystemUUID - string object to receive data
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetSystemUUID( String* pSystemUUID ) const
{
    if ( pSystemUUID == nullptr )
    {
        throw ParameterException( L"pSystemUUID", __FUNCTION__ );
    }

    // System structure (Type 1)
    GetStringValue( PXS_SMBIOS_SYSINFO_UUID, 1, pSystemUUID );
    pSystemUUID->Trim();
}

//===============================================================================================//
//  Description:
//      Get the total RAM in MB
//
//  Parameters:
//      None
//
//  Returns:
//      Total of the memory devices in MB
//===============================================================================================//
DWORD SmbiosInformation::GetTotalRamMB() const
{
    bool  success = true;   // Assume success
    WORD  deviceSize = 0, deviceNumber = 0;
    DWORD totalRamMB = 0;

    // Cycle through the memory devices and sum up the total
    while ( success )
    {
        // Counting starts at 1
        deviceNumber++;
        deviceSize = 0;
        success    = GetWordValue( PXS_SMBIOS_MEM_DEV_SIZE,
                                   deviceNumber, &deviceSize );
        if ( ( success == true ) && ( deviceSize != 0xFFFF ) )
        {
            // If bit 15 is set then in value is in KB else MB
            if ( 0x8000 & deviceSize )
            {
                deviceSize = static_cast<WORD>( deviceSize / 1024 );
            }
            totalRamMB = PXSAddUInt32( totalRamMB, deviceSize );
        }
    }

    return totalRamMB;
}

//===============================================================================================//
//  Description:
//      Read the SMBIOS data
//
//  Parameters:
//      None
//
//  Remarks:
//      First try to read the firmware tables then WMI.
//
//      The registry also caches mssmbios.sys data at SYSTEM\CurrentControlSet
//      \Services\mssmbios\Data. However, for privacy reasons, identifiers are
//      stripped from this data.
//
//  Returns:
//        void
//===============================================================================================//
void SmbiosInformation::ReadSmbiosData()
{
    const     DWORD MAX_SMBIOS_DATA_LEN = 65536;       // 64KB
    Wmi       WMI;
    DWORD     RSMB = 0, size = 0;
    HMODULE   hKernel32 = nullptr;
    Formatter Format;
    AllocateBytes AllocBytes;
    PXSDDK::RAW_SMBIOS_DATA* pRawSmBiosData = nullptr;
    LPFN_GET_SYSTEM_FIRMWARE_TABLE lpfnGSFT = nullptr;

    // Test if already read in the data
    if ( m_SmBiosData.validData )
    {
        return;
    }

    // Allocate memory for the data table if have not already done so
    m_SmBiosData.validData = false;
    if ( m_SmBiosData.pDataTable == nullptr )
    {
        m_SmBiosData.pDataTable = new BYTE[ MAX_SMBIOS_DATA_LEN ];
        if ( m_SmBiosData.pDataTable == nullptr )
        {
            throw MemoryException( __FUNCTION__ );
        }
    }
    memset( m_SmBiosData.pDataTable, 0, MAX_SMBIOS_DATA_LEN );

    ////////////////////////////////////////////////////////////////////////////
    // GetSystemFirmwareTable

    // Available on 2003+SP1 and 64-bit XP
    pRawSmBiosData = reinterpret_cast<PXSDDK::RAW_SMBIOS_DATA*>(
                                                          AllocBytes.New( MAX_SMBIOS_DATA_LEN ) );
    hKernel32      = GetModuleHandle( L"kernel32.dll" );  // do not free
    if ( hKernel32 )
    {
        // Disable C4191
        #ifdef _MSC_VER
            #pragma warning( push )
            #pragma warning ( disable : 4191 )
        #endif
            lpfnGSFT = (LPFN_GET_SYSTEM_FIRMWARE_TABLE)GetProcAddress( hKernel32,
                                                                       "GetSystemFirmwareTable" );
        #ifdef _MSC_VER
            #pragma warning( pop )
        #endif

        if ( lpfnGSFT )
        {
            // The raw SMBIOS firmware table provider is 'RSMB'
            RSMB = ('R' << 24) + ('S' << 16) + ('M' << 8)  + ('B' << 0);

            // Ignore errors as observed GetSystemFirmwareTable to return 0
            // even though data is usable
            lpfnGSFT( RSMB, 0, pRawSmBiosData, MAX_SMBIOS_DATA_LEN );
            if ( ( pRawSmBiosData->Length >  0 ) &&
                 ( pRawSmBiosData->Length <= UINT16_MAX ) )
            {
                // Store at class scope
                m_SmBiosData.tableLength  = PXSCastUInt32ToUInt16( pRawSmBiosData->Length );
                m_SmBiosData.majorVersion = pRawSmBiosData->SMBIOSMajorVersion;
                m_SmBiosData.minorVersion = pRawSmBiosData->SMBIOSMinorVersion;
                m_SmBiosData.BCDRevision  = pRawSmBiosData->DmiRevision;
                memcpy(m_SmBiosData.pDataTable,
                       pRawSmBiosData->SMBIOSTableData, pRawSmBiosData->Length);

                m_SmBiosData.validData = true;
            }
            else
            {
                PXSLogAppWarn1( L"Invalid SMBIOS buffer length: %%1.",
                                Format.UInt32( pRawSmBiosData->Length ) );
            }
        }
        else
        {
            PXSLogAppInfo( L"GetSystemFirmwareTable not available." );
        }
    }
    else
    {
        PXSLogSysError( GetLastError(), L"GetModuleHandle( kernel32.dll ) failed." );
    }

    ////////////////////////////////////////////////////////////////////////////
    // MSSmBios_RawSMBiosTables

    if ( m_SmBiosData.validData == false )
    {
        WMI.Connect( L"root\\wmi" );
        WMI.ExecQuery( L"SELECT * FROM MSSmBios_RawSMBiosTables WHERE "
                       L"InstanceName=\"SMBiosData\"" );
        if ( WMI.Next() )      // Only expecting one record
        {
            // MOF says the data type is CIM_UINT32 but query returns INT32
            size = 0;
            WMI.GetUInt32( L"Size", &size );
            if ( size == 0 )
            {
                throw SystemException( ERROR_INVALID_DATA, L"size=0", __FUNCTION__ );
            }

            if ( size != WMI.GetUInt8Array( L"SmBiosData",
                                            m_SmBiosData.pDataTable, MAX_SMBIOS_DATA_LEN ) )
            {
                throw SystemException( ERROR_INVALID_DATA, L"size", __FUNCTION__ );
            }

            if ( size > UINT16_MAX )
            {
                throw BoundsException( L"size", __FUNCTION__ );
            }
            m_SmBiosData.tableLength  = PXSCastUInt32ToUInt16( size );
            WMI.GetUInt8( L"SmbiosMajorVersion", &m_SmBiosData.majorVersion );
            WMI.GetUInt8( L"SmbiosMinorVersion", &m_SmBiosData.minorVersion );
            WMI.GetUInt8( L"DmiRevision"       , &m_SmBiosData.BCDRevision );
            m_SmBiosData.validData = true;
        }
    }

    // Log some data, the BCD revision is the same as major.minor, e.g. 35 = 2.3
    PXSLogAppInfo1( L"SMBIOS table length    : %%1", Format.UInt16( m_SmBiosData.tableLength ) );
    PXSLogAppInfo1( L"SMBIOS major version   : %%1", Format.UInt8( m_SmBiosData.majorVersion ) );
    PXSLogAppInfo1( L"SMBIOS minor version   : %%1", Format.UInt8( m_SmBiosData.minorVersion ) );
    PXSLogAppInfo1( L"SMBIOS BCD/DMI Revision: %%1", Format.UInt8( m_SmBiosData.BCDRevision  ) );
}

//===============================================================================================//
//  Description:
//      Read the SMBIOS data from a file
//
//  Parameters:
//      pszPath - file path
//
//  Remarks:
//      Used for debug only, the source file must be called
//      firmware_information.txt and be in the same directory as the
//      executable. Assumes the file was created using data obtained with
//      WMI. To debug, need to call ReadSmbiosDataFromFile at the start of
//      ReadSmbiosData. Do forget to remove the call when finished debugging.
//
//  Returns:
//        void
//===============================================================================================//
void SmbiosInformation::ReadSmbiosDataFromFile()
{
    const  DWORD MAX_SMBIOS_DATA_LEN = 65536;       // 64KB
    bool   isInData = 0;
    int    n32  = 0;
    DWORD  un32 = 0;
    size_t i    = 0, j = 0, numLines = 0, numValues = 0, idx = 0;
    size_t dataSizeBytes = 0;
    File   DataFile;
    String FilePath, Value, Line, ErrorMessage;
    Formatter   Format;
    StringArray Lines, HexBytes;

    // Allocate memory for the data table if have not already done so
    m_SmBiosData.validData = false;
    if ( m_SmBiosData.pDataTable == nullptr )
    {
        m_SmBiosData.pDataTable = new BYTE[ MAX_SMBIOS_DATA_LEN ];
        if ( m_SmBiosData.pDataTable == nullptr )
        {
            throw MemoryException( __FUNCTION__ );
        }
    }
    memset( m_SmBiosData.pDataTable, 0, MAX_SMBIOS_DATA_LEN );

    PXSGetExeDirectory( &FilePath );
    FilePath += L"firmware_information.txt";
    DataFile.ReadLineArray( FilePath, 1000, &Lines );
    numLines = Lines.GetSize();
    for ( i = 0; i < numLines; i++ )
    {
        Line = Lines.Get( i );
        Line.Trim();

        // End of Data
        if ( Line.StartsWith( L"!EOD", false ) )
        {
            if ( idx == dataSizeBytes )
            {
                m_SmBiosData.tableLength = PXSCastSizeTToUInt16( idx );
                m_SmBiosData.validData   = true;
            }
            else
            {
                ErrorMessage = L"idx != dataSizeBytes";
                throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
            }
            break;
        }

        // If already have the Major Version then reading data
        if ( isInData )
        {
            if ( idx >= dataSizeBytes )
            {
                throw BoundsException( L"idx >= dataSizeBytes", __FUNCTION__ );
            }
            HexBytes.RemoveAll();
            Line.ToArray( PXS_CHAR_SPACE, &HexBytes );
            numValues = HexBytes.GetSize();
            for ( j = 0; j < numValues; j++ )
            {
                Value = HexBytes.Get( j );
                Value.Trim();
                un32 = Format.HexStringToNumber( Value );
                m_SmBiosData.pDataTable[ idx ] = PXSCastUInt32ToUInt8( un32 );
                idx++;
                if ( idx >= MAX_SMBIOS_DATA_LEN )
                {
                    throw BoundsException( L"idx", __FUNCTION__ );
                }
            }
        }
        else if ( Line.StartsWith( L"Major version:", false ) )
        {
            Line.SubString( 14, PXS_MINUS_ONE, &Value );
            Value.Trim();
            n32 = Format.StringToInt32( Value );
            m_SmBiosData.majorVersion = PXSCastInt32ToUInt8( n32 );
        }
        else if ( Line.StartsWith( L"Minor version:", false ) )
        {
            Line.SubString( 14, PXS_MINUS_ONE, &Value );
            Value.Trim();
            n32 = Format.StringToInt32( Value );
            m_SmBiosData.minorVersion = PXSCastInt32ToUInt8( n32 );
        }
        else if ( Line.StartsWith( L"DMI Revision :", false ) )
        {
            Line.SubString( 14, PXS_MINUS_ONE, &Value );
            Value.Trim();
            n32 = Format.StringToInt32( Value );
            m_SmBiosData.BCDRevision = PXSCastInt32ToUInt8( n32 );
        }
        else if ( Line.StartsWith( L"Size (bytes) :", false ) )
        {
            Line.SubString( 14, PXS_MINUS_ONE, &Value );
            Value.Trim();
            n32 = Format.StringToInt32( Value );
            dataSizeBytes = PXSCastInt32ToUInt32( n32 );
        }
        else if ( Line.StartsWith( L"Data         :", false ) )
        {
            isInData = true;
        }
    }

    // Log some data, the BCD revision is the same as major.minor, e.g. 35 = 2.3
    PXSLogAppInfo1( L"Read SMDIOS data file  : %%1", FilePath );
    PXSLogAppInfo1( L"SMBIOS table length    : %%1", Format.UInt16( m_SmBiosData.tableLength ) );
    PXSLogAppInfo1( L"SMBIOS major version   : %%1", Format.UInt8( m_SmBiosData.majorVersion ) );
    PXSLogAppInfo1( L"SMBIOS minor version   : %%1", Format.UInt8( m_SmBiosData.minorVersion ) );
    PXSLogAppInfo1( L"SMBIOS BCD/DMI Revision: %%1", Format.UInt8( m_SmBiosData.BCDRevision  ) );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Format the specified data as a UUID
//
//  Parameters:
//      pData       - pointer to the data
//      dataLen     - length of the data, must be 16 bytes
//      pszUUID     - ASCII buffer to receive the UUID
//      charLenUUID - length of the UUID buffer in chars
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::FormatUUID( const BYTE* pData,
                                    DWORD dataLen, char* pszUUID, size_t charLenUUID ) const
{
    if ( ( pData == nullptr ) || ( dataLen != 16 ) )
    {
        throw ParameterException( L"pData/dataLen", __FUNCTION__ );
    }

    // UUID is 38 characters
    if ( ( pszUUID == nullptr ) || ( charLenUUID < 38 ) )
    {
        throw ParameterException( L"pszUUID/charLenUUID", __FUNCTION__ );
    }

    // The non-byte fields of _uuid_t require reversal for correction string
    // representation on Windows
    // typedef struct _uuid_t {
    //    unsigned32          time_low;
    //    unsigned16          time_mid;
    //    unsigned16          time_hi_and_version;
    //    unsigned8           clock_seq_hi_and_reserved;
    //    unsigned8           clock_seq_low;
    //    byte                node[6];
    // } uuid_t;
    StringCchPrintfA(
       pszUUID,
       charLenUUID,
       "{%02X%02X%02X%02X-%02X%02X-%02X%02X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
       pData[  3 ], pData[ 2 ], pData[ 1 ], pData[ 0 ],
       pData[  5 ], pData[ 4 ],
       pData[  7 ], pData[ 6 ],
       pData[  8 ],
       pData[  9 ],
       pData[ 10 ],
       pData[ 11 ],
       pData[ 12 ],
       pData[ 13 ],
       pData[ 14 ],
       pData[ 15 ] );
}

//===============================================================================================//
//  Description:
//      Get a bit/boolean item in the data array
//
//  Parameters:
//      itemID          - a named constant identifying the item
//      structureNumber - structure number of the BIOS type
//      pValue          - receives the value
//
//  Returns:
//      true if found item, else false
//===============================================================================================//
bool SmbiosInformation::GetBooleanValue( DWORD itemID, WORD structureNumber, bool* pValue ) const
{
    bool success;
    TYPE_SMBIOS_VALUE SmbiosValue;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = false;

    memset( &SmbiosValue, 0, sizeof ( SmbiosValue ) );
    success = GetSmbiosValue( itemID, structureNumber, &SmbiosValue );
    if ( success )
    {
        *pValue = SmbiosValue.boolean;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Get a byte item in the data array
//
//  Parameters:
//      itemID          - a named constant identifying the item
//      structureNumber - structure number of the BIOS type
//      pValue          - receives the value
//
//  Returns:
//      true if found item, else false
//===============================================================================================//
bool SmbiosInformation::GetByteValue( DWORD itemID, WORD structureNumber, BYTE* pValue) const
{
    bool success;
    TYPE_SMBIOS_VALUE SmbiosValue;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    memset( &SmbiosValue, 0, sizeof ( SmbiosValue ) );
    success = GetSmbiosValue( itemID, structureNumber, &SmbiosValue );
    if ( success )
    {
        *pValue = SmbiosValue.byte;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Get a DWORD item in the data array
//
//  Parameters:
//      itemID          - a named constant identifying the item
//      structureNumber - structure number of the BIOS type
//      pValue          - receives the value
//
//  Returns:
//      true if found item, else false
//===============================================================================================//
bool SmbiosInformation::GetDWordValue ( DWORD itemID, WORD structureNumber, DWORD* pValue ) const
{
    bool success;
    TYPE_SMBIOS_VALUE SmbiosValue;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    memset( &SmbiosValue, 0, sizeof ( SmbiosValue ) );
    success = GetSmbiosValue( itemID, structureNumber, &SmbiosValue );
    if ( success )
    {
        *pValue = SmbiosValue.dword;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Get a QWORD item in the data array
//
//  Parameters:
//      itemID          - a named constant identifying the item
//      structureNumber - structure number of the BIOS type
//      pValue          - receives the value
//
//  Returns:
//      true if found item, else false
//===============================================================================================//
bool SmbiosInformation::GetQWordValue ( DWORD itemID, WORD structureNumber, QWORD* pValue ) const
{
    bool success;
    TYPE_SMBIOS_VALUE SmbiosValue;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    memset( &SmbiosValue, 0, sizeof ( SmbiosValue ) );
    success = GetSmbiosValue( itemID, structureNumber, &SmbiosValue );
    if ( success )
    {
        *pValue = SmbiosValue.qword;
    }

    return success;
}


//===============================================================================================//
//  Description:
//      Get a string item in the data array
//
//  Parameters:
//      itemID          - a name constant identifying the item
//      structureNumber - structure number of the BIOS type
//      pValue          - receives the data
//
//  Returns:
//      true if found item, else false
//===============================================================================================//
bool SmbiosInformation::GetStringValue( DWORD itemID, WORD structureNumber, String* pValue ) const
{
    bool      success;
    Formatter Format;
    TYPE_SMBIOS_VALUE SmbiosValue;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = PXS_STRING_EMPTY;

    memset( &SmbiosValue, 0, sizeof ( SmbiosValue ) );
    success = GetSmbiosValue( itemID, structureNumber, &SmbiosValue );
    if ( success )
    {
        SmbiosValue.szString[ ARRAYSIZE( SmbiosValue.szString ) - 1 ] = 0x00;
        pValue->SetAnsi( SmbiosValue.szString );  // SMBIOS data is in ANSI
    }
    pValue->Trim();   // Sometimes the spaces are used when there is no data

    return success;
}


//===============================================================================================//
//  Description:
//      Get a WORD item in the data array
//
//  Parameters:
//      itemID          - a name constant identifying the item
//      structureNumber - structure number of the BIOS type
//      pValue           - receives the value
//
//  Returns:
//      true if found item, else false
//===============================================================================================//
bool SmbiosInformation::GetWordValue( DWORD itemID, WORD structureNumber, WORD* pValue ) const
{
    bool success;
    TYPE_SMBIOS_VALUE SmbiosValue;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    memset( &SmbiosValue, 0, sizeof ( SmbiosValue ) );
    success = GetSmbiosValue( itemID, structureNumber, &SmbiosValue );
    if ( success )
    {
        *pValue = SmbiosValue.word;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Get the specification of an SMBIOS item
//
//  Parameters:
//      itemID    - named constant identifying the item
//      SmbiosSpec - structure to receive the SMBIOS specification
//
//  Remarks:
//      Specification taken from Distributed Management Task Force (DTMF)
//      document DSP0134.
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetItemSpecification( DWORD itemID,
                                              PXS_TYPE_SMBIOS_SPECIFICATION* pSmbiosSpec ) const
{
    bool      success = false;
    size_t    i = 0;
    String    Error;
    Formatter Format;

    if ( pSmbiosSpec == nullptr )
    {
        throw ParameterException( L"pSmbiosSpec", __FUNCTION__ );
    }
    memset( pSmbiosSpec, 0, sizeof ( PXS_TYPE_SMBIOS_SPECIFICATION ) );

    // Identify the item in the SMBIOS specification array
    for ( i = 0; i < ARRAYSIZE( SMBIOS_SPECIFICATION ); i++ )
    {
        if ( itemID == SMBIOS_SPECIFICATION[ i ].itemID )
        {
            success = true;
            memcpy( pSmbiosSpec,
                    SMBIOS_SPECIFICATION + i, sizeof ( PXS_TYPE_SMBIOS_SPECIFICATION ) );
            break;
        }
    }

    if ( success == false )
    {
        Error = Format.StringUInt32( L"itemID = %%1", itemID );
        throw SystemException( ERROR_INVALID_DATA, Error.c_str(), __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the specified occurrence of a data item
//
//  Parameters:
//      itemID          - a named constant identifying the item
//      structureNumber - 0ne-based structure number of the BIOS type
//      pSmbiosValue    - receives the value
//
//  Returns:
//      true if found item otherwise false
//===============================================================================================//
bool SmbiosInformation::GetSmbiosValue( DWORD itemID,
                                        WORD structureNumber,
                                        TYPE_SMBIOS_VALUE* pSmbiosValue ) const
{
    WORD      structureOffset = 0;
    DWORD     bytePos = 0, idx = 0;
    String    Value, Name, Error;
    Formatter Format, Format2;
    PXS_TYPE_SMBIOS_SPECIFICATION SmbiosSpec;

    if ( pSmbiosValue == nullptr )
    {
        throw ParameterException( L"pSmbiosValue", __FUNCTION__ );
    }
    memset( pSmbiosValue, 0, sizeof ( TYPE_SMBIOS_VALUE ) );

    if ( m_SmBiosData.validData == false )
    {
        throw FunctionException( L"m_SmBiosData.validData", __FUNCTION__ );
    }

    // Get the specification for the item, return if not found
    memset( &SmbiosSpec, 0, sizeof ( SmbiosSpec ) );
    GetItemSpecification( itemID, &SmbiosSpec );

    // Verify the item is present in the SMBIOS version, for readability
    // will test major and minor version separatedly
    if ( SmbiosSpec.majorVersion > m_SmBiosData.majorVersion )
    {
        return false;
    }
    if ( ( SmbiosSpec.majorVersion == m_SmBiosData.majorVersion ) &&
         ( SmbiosSpec.minorVersion >  m_SmBiosData.minorVersion )  )
    {
        return false;
    }

    // Locate SMBIOS type in data array
    if ( false == GetStructureOffset( SmbiosSpec.smbiosType, structureNumber, &structureOffset ) )
    {
        PXSLogAppInfo2(L"SMBIOS type %%1, structure number %%2 not found.",
                       Format.UInt8( SmbiosSpec.smbiosType ), Format2.UInt16( structureNumber ) );
        return false;
    }

    // Check bounds
    if ( ( structureOffset + SmbiosSpec.offset + SmbiosSpec.length ) >= m_SmBiosData.tableLength )
    {
        throw BoundsException( L"structureOffset + tItem.offset", __FUNCTION__ );
    }

    // Data extraction depends on the data type
    idx = PXSAddUInt32( structureOffset, SmbiosSpec.offset );
    if ( SmbiosSpec.dataType == PXS_SMBIOS_DATA_TYPE_BOOLEAN )
    {
        // Based on bit, get the byte's position
        bytePos = idx + ( SmbiosSpec.bitNumber / 8 );
        pSmbiosValue->byte = m_SmBiosData.pDataTable[ bytePos ];
        if ( pSmbiosValue->byte & ( 1 << ( SmbiosSpec.bitNumber % 8 ) ) )
        {
            pSmbiosValue->boolean = true;
        }
    }
    else if ( SmbiosSpec.dataType == PXS_SMBIOS_DATA_TYPE_UINT8 )
    {
        pSmbiosValue->byte = m_SmBiosData.pDataTable[ idx ];
    }
    else if ( SmbiosSpec.dataType == PXS_SMBIOS_DATA_TYPE_UINT16 )
    {
       pSmbiosValue->word = MAKEWORD( m_SmBiosData.pDataTable[ idx ] ,
                                      m_SmBiosData.pDataTable[ idx + 1 ] );
    }
    else if ( SmbiosSpec.dataType == PXS_SMBIOS_DATA_TYPE_UINT32 )
    {
        pSmbiosValue->dword  = (DWORD)m_SmBiosData.pDataTable[ idx ];
        pSmbiosValue->dword += (DWORD)m_SmBiosData.pDataTable[ idx + 1 ] << 8;
        pSmbiosValue->dword += (DWORD)m_SmBiosData.pDataTable[ idx + 2 ] << 16;
        pSmbiosValue->dword += (DWORD)m_SmBiosData.pDataTable[ idx + 3 ] << 24;
    }
    else if ( SmbiosSpec.dataType == PXS_SMBIOS_DATA_TYPE_UINT64 )
    {
        pSmbiosValue->qword  = (QWORD)m_SmBiosData.pDataTable[ idx ];
        pSmbiosValue->qword += (QWORD)m_SmBiosData.pDataTable[ idx+1 ] << 8;
        pSmbiosValue->qword += (QWORD)m_SmBiosData.pDataTable[ idx+2 ] << 16;
        pSmbiosValue->qword += (QWORD)m_SmBiosData.pDataTable[ idx+3 ] << 24;
        pSmbiosValue->qword += (QWORD)m_SmBiosData.pDataTable[ idx+4 ] << 32;
        pSmbiosValue->qword += (QWORD)m_SmBiosData.pDataTable[ idx+5 ] << 40;
        pSmbiosValue->qword += (QWORD)m_SmBiosData.pDataTable[ idx+6 ] << 48;
        pSmbiosValue->qword += (QWORD)m_SmBiosData.pDataTable[ idx+7 ] << 56;
    }
    else if ( SmbiosSpec.dataType == PXS_SMBIOS_DATA_TYPE_STRING )
    {
        if ( itemID == PXS_SMBIOS_SYSINFO_UUID )
        {
            FormatUUID( m_SmBiosData.pDataTable + idx,
                        SmbiosSpec.length,
                        pSmbiosValue->szString, ARRAYSIZE( pSmbiosValue->szString ) );
        }
        else
        {
            GetString( m_SmBiosData.pDataTable[ idx ],
                       m_SmBiosData.pDataTable + structureOffset,
                       static_cast<WORD>(m_SmBiosData.tableLength - structureOffset),
                       pSmbiosValue->szString, ARRAYSIZE( pSmbiosValue->szString ) );
        }
    }
    else
    {
        Error = Format.StringUInt32( L"Item ID=%%1.", itemID );
        throw SystemException( ERROR_INVALID_DATA, Error.c_str(), __FUNCTION__ );
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the a string from the specified SMBIOS structure
//
//  Parameters:
//      stringNumber - the string number
//      pStructure   - pointer to the start of the structure
//      bufferLen    - length of the data buffer in chars
//      pszBuffer    - ASCII buffer to receive the string
//      bufferChars  - size if the string buffer
//
//  Remarks:
//      bufferLen should the number of bytes from he start of the
//      structure to the end of the SMBIOS data
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::GetString( BYTE stringNumber,
                                   const BYTE* pStructure,
                                   WORD  bufferLen, char* pszBuffer, size_t bufferChars ) const
{
    bool found   = false, atEnd = false;
    BYTE counter = 0;
    WORD offset  = 0, start = 0;

    if ( pStructure == nullptr )
    {
        return;     // No data
    }

    if ( ( pszBuffer == nullptr ) || ( bufferChars == 0 ) )
    {
        throw ParameterException( L"pszBuffer", __FUNCTION__ );
    }
    memset( pszBuffer, 0, bufferChars );    // ASCII buffer


    // Skip over the formatted area, its length is at index 1
    offset = pStructure[ 1 ];
    start  = offset;
    while ( ( atEnd == false ) &&
            ( found == false ) &&
            ( PXSAddUInt16( offset, 2 ) < bufferLen ) )
    {
        // Test for end of string
        if ( pStructure[ offset ] == 0x00 )
        {
            counter++;
            if ( stringNumber == counter )
            {
                StringCchCopyA(
                            pszBuffer,
                            bufferChars, reinterpret_cast<const char*>( pStructure + start ) );
                found = true;
            }

            // Test for end of structure
            if ( pStructure[ offset + 1 ] == 0x00 )
            {
                atEnd = true;
            }
            else
            {
                // Set the new start of position of next string
                start = offset;
                start++;
            }
        }
        offset++;
    }
}

//===============================================================================================//
//  Description:
//      Get the BIOS information (Type 0) as an audit record
//
//  Parameters:
//      pRecord - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType0_BiosRecord( AuditRecord* pRecord ) const
{
    bool       supported = false;
    BYTE       byteValue = 0;
    WORD       offset    = 0, word = 0;
    QWORD      qword     = 0;
    size_t     i = 0;
    String     Value, LocaleKB;
    Formatter Format;

    // Bit fields for PXS_SMBIOS_BIOS_CHARACTERISTICS and
    // PXS_SMBIOS_BIOS_CHARAC_EXT_1
    DWORD Items[] = { PXS_SMBIOS_BIOS_C1_SUPPORTED,
                      PXS_SMBIOS_BIOS_C1_ISA,
                      PXS_SMBIOS_BIOS_C1_MCA,
                      PXS_SMBIOS_BIOS_C1_EISA,
                      PXS_SMBIOS_BIOS_C1_PCI,
                      PXS_SMBIOS_BIOS_C1_PCMCIA,
                      PXS_SMBIOS_BIOS_C1_PNP,
                      PXS_SMBIOS_BIOS_C1_APM,
                      PXS_SMBIOS_BIOS_C1_FLASH,
                      PXS_SMBIOS_BIOS_C1_SHADOW,
                      PXS_SMBIOS_BIOS_C1_VL_VESA,
                      PXS_SMBIOS_BIOS_C1_SUPPORT_ESCD,
                      PXS_SMBIOS_BIOS_C1_BOOT_CD,
                      PXS_SMBIOS_BIOS_C1_SELECT_BOOT,
                      PXS_SMBIOS_BIOS_C1_ROM_SOCKETED,
                      PXS_SMBIOS_BIOS_C1_BOOT_PCMCIA,
                      PXS_SMBIOS_BIOS_C1_EDD,
                      PXS_SMBIOS_BIOS_C1_FDD_NEC_1_2MB,
                      PXS_SMBIOS_BIOS_C1_FDD_TOS_1_2MB,
                      PXS_SMBIOS_BIOS_C1_FDD_360KB,
                      PXS_SMBIOS_BIOS_C1_FDD_1_2MB,
                      PXS_SMBIOS_BIOS_C1_FDD_720KB,
                      PXS_SMBIOS_BIOS_C1_FDD_2_88MB,
                      PXS_SMBIOS_BIOS_C1_PRINT_SCREEN,
                      PXS_SMBIOS_BIOS_C1_KEYBOARD_8042,
                      PXS_SMBIOS_BIOS_C1_SERIAL,
                      PXS_SMBIOS_BIOS_C1_PRINTER,
                      PXS_SMBIOS_BIOS_C1_CGA_MONO,
                      PXS_SMBIOS_BIOS_C1_NEC_PC98,
                      PXS_SMBIOS_BIOS_CX1_SUPPORT_ACPI,
                      PXS_SMBIOS_BIOS_CX1_LEGACY_USB,
                      PXS_SMBIOS_BIOS_CX1_SUPPORT_AGP,
                      PXS_SMBIOS_BIOS_CX1_BOOT_I2O,
                      PXS_SMBIOS_BIOS_CX1_BOOT_LS120,
                      PXS_SMBIOS_BIOS_CX1_BOOT_ATAPI_Z,
                      PXS_SMBIOS_BIOS_CX1_BOOT_1394,
                      PXS_SMBIOS_BIOS_CX1_SMART_BAT };

     if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_INFO );

    // Verify the structure exists in the data
    if ( GetStructureOffset( PXS_SMBIOS_TYPE_0_BIOS, 1, &offset ) == false )
    {
        PXSLogAppWarn( L"SMBIOS structure (type 0) not found." );
        return false;
    }

    // Locale strings
    PXSGetResourceString( PXS_IDS_135_KB, &LocaleKB );

    // Add the system management version here, not specified in
    // type_0 structure but this seems the most logical place
    GetSmbiosMajorDotMinor( &Value );
    pRecord->Add( PXS_SMBIOS_BIOS_SMBIOS_VERSION, Value );

    // Vendor
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_BIOS_SMBIOS_VENDOR, 1, &Value );
    pRecord->Add( PXS_SMBIOS_BIOS_SMBIOS_VENDOR, Value );

    // BIOS Version
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_BIOS_VERSION, 1, &Value );
    pRecord->Add( PXS_SMBIOS_BIOS_VERSION, Value );

    // BIOS Starting Address Segment
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_BIOS_START_ADDRESS, 1, &word ) )
    {
        Value = Format.UInt16Hex( word );
    }
    pRecord->Add( PXS_SMBIOS_BIOS_START_ADDRESS, Value );

    // BIOS Release Date
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_BIOS_RELEASE_DATE, 1, &Value );
    pRecord->Add( PXS_SMBIOS_BIOS_RELEASE_DATE,  Value );

    // BIOS ROM Size
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_BIOS_ROM_SIZE_KB, 1, &byteValue ) )
    {
        Value  = Format.Int32( (byteValue + 1) * 64 );
        Value += LocaleKB;
    }
    pRecord->Add( PXS_SMBIOS_BIOS_ROM_SIZE_KB, Value );

    // BIOS Characteristics
    Value = PXS_STRING_EMPTY;
    if ( GetQWordValue( PXS_SMBIOS_BIOS_CHARACTERISTICS, 1, &qword ) )
    {
        Value = Format.UInt64Hex( qword, true );
    }
    pRecord->Add( PXS_SMBIOS_BIOS_CHARACTERISTICS, Value );

    // BIOS Characteristics Extension Bytes #1
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_BIOS_CHAR_EXT_BYTE_1, 1, &byteValue ) )
    {
        Value = Format.UInt8Hex( byteValue, true );
    }
    pRecord->Add( PXS_SMBIOS_BIOS_CHAR_EXT_BYTE_1, Value );

    // BIOS Characteristics Extension Bytes #2
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_BIOS_CHAR_EXT_BYTE_2, 1, &byteValue ) )
    {
        Value = Format.UInt8Hex( byteValue, true );
    }
    pRecord->Add( PXS_SMBIOS_BIOS_CHAR_EXT_BYTE_2, Value );

    // System BIOS Major Release
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue(PXS_SMBIOS_BIOS_SYS_BIOS_MAJOR, 1, &byteValue))
    {
        if ( byteValue != 0xFF )     // System does not support this field
        {
            Value = Format.UInt8( byteValue );
        }
    }
    pRecord->Add( PXS_SMBIOS_BIOS_SYS_BIOS_MAJOR, Value  );

    // System BIOS Minor Release
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_BIOS_SYS_BIOS_MINOR, 1, &byteValue))
    {
        if ( byteValue != 0xFF )     // System does not support this field
        {
            Value = Format.UInt8( byteValue );
        }
    }
    pRecord->Add( PXS_SMBIOS_BIOS_SYS_BIOS_MINOR,  Value );

    // Embedded Controller Firmware Major Release
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_BIOS_FIRMWARE_MAJOR, 1, &byteValue ) )
    {
        if ( byteValue != 0xFF )     // System does not support this field
        {
            Value = Format.UInt8( byteValue );
        }
    }
    pRecord->Add( PXS_SMBIOS_BIOS_FIRMWARE_MAJOR, Value );

    // Embedded Controller Firmware Minor Release
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_BIOS_FIRMWARE_MINOR, 1, &byteValue ) )
    {
        if ( byteValue != 0xFF )     // System does not support this field
        {
            Value = Format.UInt8( byteValue );
        }
    }
    pRecord->Add( PXS_SMBIOS_BIOS_FIRMWARE_MINOR, Value );

    // Bit fields
    for ( i = 0; i < ARRAYSIZE( Items ); i++ )
    {
        Value = PXS_STRING_EMPTY;
        if ( GetBooleanValue( Items[ i ], 1, &supported ) )
        {
            if ( supported )
            {
                Value = L"1";
            }
            else
            {
                Value = L"0";
            }
        }
        else
        {
            Value = L"0";
        }
        pRecord->Add( Items[ i ], Value );
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the System information (Type 1) as an audit record
//
//  Parameters:
//      pRecord - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType1_SystemRecord( AuditRecord* pRecord ) const
{
    BYTE   byteValue = 0;
    WORD   offset    = 0;
    String Value;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_SYSINFO );

    // Verify the specified structure number exists in the data
    if ( GetStructureOffset( PXS_SMBIOS_TYPE_1_SYSTEM, 1, &offset ) == false )
    {
        PXSLogAppWarn( L"SMBIOS system (type 1) table not found." );
        return false;
    }

    // Manufacturer
    GetStringValue( PXS_SMBIOS_SYSINFO_MANUFACTURER, 1, &Value );
    pRecord->Add( PXS_SMBIOS_SYSINFO_MANUFACTURER, Value );

    // Product Name
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_SYSINFO_PRODUCT_NAME, 1, &Value );
    pRecord->Add( PXS_SMBIOS_SYSINFO_PRODUCT_NAME, Value );

    // Version
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_SYSINFO_VERSION, 1, &Value );
    pRecord->Add( PXS_SMBIOS_SYSINFO_VERSION, Value );

    // Serial Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_SYSINFO_SERIAL_NUMBER, 1, &Value );
    pRecord->Add( PXS_SMBIOS_SYSINFO_SERIAL_NUMBER, Value );

    // UUID, 16 bytes. Treat as a string
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_SYSINFO_UUID, 1, &Value );
    pRecord->Add( PXS_SMBIOS_SYSINFO_UUID, Value );

    // Wake-up Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_SYSINFO_WAKEUP_TYPE, 1, &byteValue ) )
    {
        TranslateWakeUpType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_SYSINFO_WAKEUP_TYPE, Value );

    // SKU Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_SYSINFO_SKU_NUMBER, 1, &Value );
    pRecord->Add( PXS_SMBIOS_SYSINFO_SKU_NUMBER, Value );

    // Family
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_SYSINFO_FAMILY, 1, &Value );
    pRecord->Add( PXS_SMBIOS_SYSINFO_FAMILY, Value );

    return true;
}


//===============================================================================================//
//  Description:
//      Get the Baseboard information (Type 2) as an audit record
//
//  Parameters:
//      boardNumber - number of board/structure in SMBIOS data
//      pRecord     - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType2_BaseBoardRecord( WORD boardNumber, AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      offset    = 0;
    String    Value;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_BOARD );

    // Verify the specified structure number exists in the data
    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_2_BASE_BOARD, boardNumber, &offset ) )
    {
        return false;
    }

    Value = Format.UInt16( boardNumber );
    pRecord->Add( PXS_SMBIOS_BOARD_ITEM_NUMBER, Value  );

    // Manufacturer
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_BOARD_MANUFACTURER,  boardNumber, &Value );
    pRecord->Add( PXS_SMBIOS_BOARD_MANUFACTURER, Value );

    // Product
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_BOARD_PRODUCT, boardNumber, &Value );
    pRecord->Add( PXS_SMBIOS_BOARD_PRODUCT, Value );

    // Version
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_BOARD_VERSION, boardNumber, &Value );
    pRecord->Add( PXS_SMBIOS_BOARD_VERSION, Value );

    // Serial Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_BOARD_SERIAL_NUMBER, boardNumber, &Value );
    pRecord->Add( PXS_SMBIOS_BOARD_SERIAL_NUMBER, Value );

    // Asset Tag
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_BOARD_ASSET_TAG, boardNumber, &Value );
    pRecord->Add( PXS_SMBIOS_BOARD_ASSET_TAG, Value );

    // Feature Flags
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue(PXS_SMBIOS_BOARD_FEATURES_FLAG, boardNumber, &byteValue))
    {
        TranslateBoardFeaturesFlags( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_BOARD_FEATURES_FLAG, Value );

    // Board Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_BOARD_BOARD_TYPE, boardNumber, &byteValue ) )
    {
        TranslateBaseBoardType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_BOARD_BOARD_TYPE, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the Chassis information (Type 3) as an audit record
//
//  Parameters:
//      pRecord - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType3_ChassisRecord( AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      offset    = 0;
    DWORD     dwordValue= 0;
    String    Value;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_CHASSIS );

    if ( GetStructureOffset(PXS_SMBIOS_TYPE_3_CHASSIS, 1, &offset ) == false )
    {
        return false;
    }

    // Manufacturer
    GetStringValue( PXS_SMBIOS_CHASSIS_MANUFACTURER,  1, &Value );
    pRecord->Add( PXS_SMBIOS_CHASSIS_MANUFACTURER, Value );

    // Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CHASSIS_TYPE, 1, &byteValue ) )
    {
        TranslateChassisType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CHASSIS_TYPE, Value );

    // Version
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_CHASSIS_VERSION, 1, &Value );
    pRecord->Add( PXS_SMBIOS_CHASSIS_VERSION, Value );

    // Serial Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_CHASSIS_SERIAL_NUMBER, 1, &Value );
    pRecord->Add( PXS_SMBIOS_CHASSIS_SERIAL_NUMBER, Value );

    // Asset Tag Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_CHASSIS_ASSET_TAG_NUM, 1, &Value );
    pRecord->Add( PXS_SMBIOS_CHASSIS_ASSET_TAG_NUM, Value );

    // Boot-up State
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CHASSIS_BOOT_UP_STATE, 1, &byteValue ) )
    {
        TranslateChassisState( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CHASSIS_BOOT_UP_STATE, Value );

    // Power Supply State
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CHASSIS_POWER_SUPPLY, 1, &byteValue ) )
    {
        TranslateChassisState( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CHASSIS_POWER_SUPPLY, Value );

    // Thermal State
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CHASSIS_THERMAL_STATE, 1, &byteValue ) )
    {
        TranslateChassisState( byteValue, &Value);
    }
    pRecord->Add( PXS_SMBIOS_CHASSIS_THERMAL_STATE, Value );

    // Security Status
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CHASSIS_SECUR_STATUS, 1, &byteValue ) )
    {
        TranslateChassisSecurityStatus( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CHASSIS_SECUR_STATUS, Value );

    // OEM-defined
    Value = PXS_STRING_EMPTY;
    if ( GetDWordValue( PXS_SMBIOS_CHASSIS_OEM_DEFINED, 1, &dwordValue ) )
    {
        Value = Format.UInt32( dwordValue );
    }
    pRecord->Add( PXS_SMBIOS_CHASSIS_OEM_DEFINED, Value );

    // Height
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CHASSIS_HEIGHT_U, 1, &byteValue ) )
    {
        // 00h indicates that the enclosure height is unspecified, also will
        // filter out the space character, possibly meant to indicate
        // not specified.
        if ( ( byteValue > 0 ) && ( byteValue != 0x20 ) )
        {
            Value = Format.Double( 1.75 * byteValue, 2 );
        }
    }
    pRecord->Add( PXS_SMBIOS_CHASSIS_HEIGHT_U, Value );

    // Number of Power Cords
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CHASSIS_POWER_CORDS, 1, &byteValue ) )
    {
        // 00h indicates that the number is unspecified, also will filter out
        // the space character, possibly meant to indicate not specified.
        if ( ( byteValue > 0 ) && ( byteValue != 0x20 ) )
        {
            Value = Format.UInt8( byteValue );
        }
    }
    pRecord->Add( PXS_SMBIOS_CHASSIS_POWER_CORDS, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the Processor information (Type 4) as an audit record
//
//  Parameters:
//      processorNumber - number of processor/structure in SMBIOS data
//      pRecord         - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType4_ProcessorRecord( WORD processorNumber,
                                                  AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0, family = 0;
    WORD      word  = 0, offset = 0;
    QWORD     qword = 0;
    String    Value, LocaleMHz;
    Formatter Format;

     if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_PROC );

    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_4_PROCESSOR, processorNumber, &offset ) )
    {
        return false;
    }

    // Locale string
    PXSGetResourceString( PXS_IDS_1263_MHZ, &LocaleMHz );

    Value = Format.UInt16( processorNumber );
    pRecord->Add( PXS_SMBIOS_PROCESSOR_ITEM_NUMBER, Value );

    // Socket Designation
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_PROCESSOR_SOCKET, processorNumber, &Value );
    pRecord->Add( PXS_SMBIOS_PROCESSOR_SOCKET,  Value );

    // Processor Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_PROCESSOR_TYPE, processorNumber, &byteValue ))
    {
        TranslateProcessorType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_TYPE, Value );

    // Processor Family
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_PROCESSOR_FAMILY, processorNumber, &family ) )
    {
        TranslateProcessorFamily( family, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_FAMILY, Value );

    // Processor Manufacturer
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_PROCESSOR_MANUFAC, processorNumber, &Value );
    pRecord->Add( PXS_SMBIOS_PROCESSOR_MANUFAC, Value );

    // Processor ID
    Value = PXS_STRING_EMPTY;
    if ( GetQWordValue( PXS_SMBIOS_PROCESSOR_ID, processorNumber, &qword ) )
    {
        // Format as two hex values, EAX-EDX
        Value  = Format.UInt32Hex( static_cast<DWORD>( qword >> 32 ), true );
        Value += L"-";
        Value += Format.UInt32Hex( qword & 0xFFFFFFFF, true );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_ID, Value );

    // Processor Version
    Value = PXS_STRING_EMPTY;
    GetStringValue(PXS_SMBIOS_PROCESSOR_VERSION, processorNumber, &Value);
    pRecord->Add( PXS_SMBIOS_PROCESSOR_VERSION, Value );

    // Voltage
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_PROCESSOR_VOLTAGE, processorNumber, &byteValue ) )
    {
        TranslateProcessorVoltage( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_VOLTAGE, Value );

    // External Clock
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_PROCESSOR_EXT_CLK_MHZ, processorNumber, &word ) )
    {
        Value  = Format.UInt16( word );
        Value += LocaleMHz;
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_EXT_CLK_MHZ, Value );

    // Max Speed
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_PROCESSOR_MAXIMUM_MHZ, processorNumber, &word ) )
    {
        Value  = Format.UInt16( word );
        Value += LocaleMHz;
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_MAXIMUM_MHZ, Value );

    // Current Speed
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_PROCESSOR_CURRENT_MHZ, processorNumber, &word ) )
    {
        Value  = Format.UInt16( word );
        Value += LocaleMHz;
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_CURRENT_MHZ, Value );

    // Status
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_PROCESSOR_STATUS, processorNumber, &byteValue ) )
    {
        TranslateProcessorStatus( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_STATUS, Value );

    // Processor Upgrade
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_PROCESSOR_UPGRADE, processorNumber, &byteValue ) )
    {
        TranslateProcessorUpgrade( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_UPGRADE, Value );

    // Serial Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_PROCESSOR_SERIAL_NUM, processorNumber, &Value );
    pRecord->Add( PXS_SMBIOS_PROCESSOR_SERIAL_NUM, Value );

    // Asset Tag
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_PROCESSOR_ASSET_TAG, processorNumber, &Value );
    pRecord->Add( PXS_SMBIOS_PROCESSOR_ASSET_TAG, Value );

    // Part Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_PROCESSOR_PART_NUM, processorNumber, &Value );
    pRecord->Add( PXS_SMBIOS_PROCESSOR_PART_NUM, Value );

    // Core Count
    byteValue  = 0;       // Signifies Unknown
    Value = PXS_STRING_EMPTY;
    GetByteValue(PXS_SMBIOS_PROCESSOR_CORE_COUNT, processorNumber, &byteValue);
    if ( byteValue )
    {
        Value = Format.UInt8( byteValue );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_CORE_COUNT, Value );

    // Core Enabled
    byteValue  = 0;       // Signifies Unknown
    Value = PXS_STRING_EMPTY;
    GetByteValue( PXS_SMBIOS_PROCESSOR_CORE_ENABLE, processorNumber, &byteValue );
    if ( byteValue )
    {
        Value = Format.UInt8( byteValue );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_CORE_ENABLE, Value );

    // Thread Count
    byteValue  = 0;       // Signifies Unknown
    Value = PXS_STRING_EMPTY;
    GetByteValue( PXS_SMBIOS_PROCESSOR_THREADS,
                  processorNumber, &byteValue );
    if ( byteValue )
    {
        Value = Format.UInt8( byteValue );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_THREADS, PXS_STRING_EMPTY );

    // Processor Characteristics
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_PROCESSOR_CHARACTERIS, processorNumber, &word ) )
    {
        TranslateProcessorCharacteristics( word, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_CHARACTERIS, Value );

    // Processor family 2, only used if processor family value is 0xFE
    Value = PXS_STRING_EMPTY;
    if ( family == 0xFE )
    {
        if ( GetWordValue( PXS_SMBIOS_PROCESSOR_FAMILY_2, processorNumber, &word ) )
        {
            TranslateProcessorFamily2( word, &Value );
        }
    }
    pRecord->Add( PXS_SMBIOS_PROCESSOR_FAMILY_2, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the Memory Controller information (Type 5) as an audit record
//
//  Parameters:
//      controllerNumber - number of controller/structure in SMBIOS data
//      pRecord          - receives the record
//
//  Remarks:
//      Although this information is obsolete with SMBIOS v2.1, have seen newer
//      SMBIOS versions that do not have Type 16 and Type 17 data.
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType5_MemoryControllerRecord( WORD controllerNumber,
                                                         AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      word  = 0, offset = 0;
    DWORD     dword = 0;
    double    moduleMB = 0.0;
    String    Value, LocaleMB;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_MEMCTRL );

    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_5_MEM_CONTROL, controllerNumber, &offset ) )
    {
        return false;
    }

    // Locale string
    PXSGetResourceString( PXS_IDS_136_MB, &LocaleMB );

    Value = Format.UInt16( controllerNumber );
    pRecord->Add( PXS_SMBIOS_MEMCTRL_ITEM_NUMBER, Value );

    // Error Detecting Method
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMCTRL_ERR_DETECT, controllerNumber, &byteValue ) )
    {
        TranslateMemoryControllerErrCorrection( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMCTRL_ERR_DETECT, Value );

    // Error Correcting Capability
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMCTRL_ERR_CORRECT, controllerNumber, &byteValue ) )
    {
        TranslateMemoryControllerErrCorrection( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMCTRL_ERR_CORRECT, Value );

    // Supported Interleave
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMCTRL_SUPPORT_INTER, controllerNumber, &byteValue ) )
    {
        TranslateMemoryControllerInterleave( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMCTRL_SUPPORT_INTER, Value );

    // Current Interleave
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMCTRL_CURRENT_INTER, controllerNumber, &byteValue ) )
    {
        TranslateMemoryControllerInterleave( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMCTRL_CURRENT_INTER, Value );

    // Maximum Memory Module Size
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMCTRL_MAX_MODULE_MB, controllerNumber, &byteValue ) )
    {
        moduleMB = pow( 2.0, byteValue );
        dword    = PXSCastDoubleToUInt32( moduleMB );
        Value    = Format.UInt32( dword );
    }
    Value += LocaleMB;
    pRecord->Add( PXS_SMBIOS_MEMCTRL_MAX_MODULE_MB, Value );

    // Supported Speeds
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_MEMCTRL_SPEEDS, controllerNumber, &word ) )
    {
        TranslateMemoryControllerSpeed( word, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMCTRL_SPEEDS, Value );

    // Supported Memory Types
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_MEMCTRL_MEM_TYPES, controllerNumber, &word ) )
    {
        TranslateMemoryModuleType( word, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMCTRL_MEM_TYPES, Value );

    // Memory Module Voltage
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMCTRL_VOLTAGE, controllerNumber, &byteValue ) )
    {
        TranslateMemoryControllerVoltage( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMCTRL_VOLTAGE, Value );

    // Number of Associated Memory Slots
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMCTRL_NUM_SLOTS, controllerNumber, &byteValue ) )
    {
        Value = Format.UInt8( byteValue );
    }
    pRecord->Add( PXS_SMBIOS_MEMCTRL_NUM_SLOTS, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the Memory Module information (Type 6) as an audit record
//
//  Parameters:
//      moduleNumber - number of module/structure in SMBIOS data
//      pRecord      - receives the data
//
//  Remarks:
//      Although this information is obsolete with SMBIOS v2.1, have seen newer
//      SMBIOS versions that do not have Type 16 and Type 17 data.
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType6_MemoryModuleRecord( WORD moduleNumber,
                                                     AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      word = 0, offset = 0;
    String    Value, LocaleNS;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_MEMMODULE );

    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_6_MEMORY_MODULE, moduleNumber, &offset ) )
    {
        return false;
    }

    // Locale string
    PXSGetResourceString( PXS_IDS_1266_NS, &LocaleNS );

    Value = Format.UInt16( moduleNumber );
    pRecord->Add( PXS_SMBIOS_MEMMODULE_ITEM_NUMBER, Value );

    // Socket Designation
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_MEMMODULE_SOCKET, moduleNumber, &Value );
    pRecord->Add( PXS_SMBIOS_MEMMODULE_SOCKET, Value );

    // Bank Connections
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMMODULE_BANK_CONN, moduleNumber, &byteValue ) )
    {
        TranslateMemoryModuleBank( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMMODULE_BANK_CONN, Value );

    // Current Speed
    byteValue  = 0;       // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetByteValue( PXS_SMBIOS_MEMMODULE_SPEED_NS, moduleNumber, &byteValue );
    if ( byteValue )
    {
        Value  = Format.UInt8( byteValue );
        Value += LocaleNS;
    }
    pRecord->Add( PXS_SMBIOS_MEMMODULE_SPEED_NS, Value );

    // Current Memory Type
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_MEMMODULE_MEM_TYPE, moduleNumber, &word ) )
    {
        TranslateMemoryModuleType( word, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMMODULE_MEM_TYPE, Value );

    // Installed Size
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMMODULE_INSTAL_SIZE, moduleNumber, &byteValue ) )
    {
        TranslateMemoryModuleSize( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMMODULE_INSTAL_SIZE, Value );

    // Enabled Size
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMMODULE_ENABLE_SIZE, moduleNumber, &byteValue ) )
    {
        TranslateMemoryModuleSize( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMMODULE_ENABLE_SIZE, Value );

    // Error Status
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMMODULE_ERR_STATUS, moduleNumber, &byteValue ) )
    {
        TranslateMemoryModuleErrorStatus( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMMODULE_ERR_STATUS, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the CPU Cache information (Type 7) as an audit record
//
//  Parameters:
//      cacheNumber - number of cache/structure in SMBIOS data
//      pRecord     - receives the data
//
//  Remarks:
//
//  Returns:
//      true on success, else false.
//
//===============================================================================================//
bool SmbiosInformation::GetType7_CpuCacheRecord( WORD cacheNumber, AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      word = 0, offset = 0;
    String    Value, LocaleKB, LocaleNS;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_CPUCACHE );

    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_7_CPU_CACHE, cacheNumber, &offset ) )
    {
         return false;
    }

    // Locale string
    PXSGetResourceString( PXS_IDS_135_KB , &LocaleKB );
    PXSGetResourceString( PXS_IDS_1266_NS, &LocaleNS );

    Value = Format.UInt16( cacheNumber );
    pRecord->Add( PXS_SMBIOS_CPUCACHE_ITEM_NUMBER, Value );

    // Socket Designation
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_CPUCACHE_SOCKET, cacheNumber, &Value );
    pRecord->Add( PXS_SMBIOS_CPUCACHE_SOCKET, Value );

    // Cache Configuration
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_CPUCACHE_CONFIG, cacheNumber, &word ) )
    {
        TranslateCpuCacheConfiguration( word, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_CONFIG, Value );

    // Maximum Cache Size
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_CPUCACHE_MAX_SIZE_KB, cacheNumber, &word ) )
    {
        // Bit 15 Granularity, 1 - 64K granularity or 0 - 1K granularity
        if ( 0x8000 & word )
        {
            Value = Format.Int32( 64 * word );
        }
        else
        {
            Value = Format.UInt16( word );
        }
        Value += LocaleKB;
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_MAX_SIZE_KB, Value );

    // Installed Size
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_CPUCACHE_INST_SIZE_KB, cacheNumber, &word ) )
    {
        // Bit 15 Granularity, 1 - 64K granularity or 0 - 1K granularity
        if ( 0x8000 & word )
        {
            Value = Format.Int32( 64 * word );
        }
        else
        {
            Value = Format.UInt16( word );
        }
        Value += LocaleKB;
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_INST_SIZE_KB, Value );

    // Supported SRAM Type
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_CPUCACHE_SUPPORT_SRAM, cacheNumber, &word ) )
    {
        TranslateCpuCacheSramType( word, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_SUPPORT_SRAM, Value );

    // Current SRAM Type
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_CPUCACHE_CURRENT_SRAM, cacheNumber, &word ) )
    {
        TranslateCpuCacheSramType( word, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_CURRENT_SRAM, Value );

    // Cache Speed
    byteValue  = 0;       // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetByteValue( PXS_SMBIOS_CPUCACHE_SPEED_NS, cacheNumber, &byteValue );
    if ( byteValue )
    {
        Value  = Format.UInt8( byteValue );
        Value += LocaleNS;
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_SPEED_NS, Value );

    // Error Correction Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CPUCACHE_ERR_CORRECT, cacheNumber, &byteValue ) )
    {
        TranslateCpuCacheErrCorrection( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_ERR_CORRECT, Value );

    // System Cache Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CPUCACHE_SYSTEM_TYPE, cacheNumber, &byteValue ) )
    {
        TranslateCpuCacheType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_SYSTEM_TYPE, Value );

    // Associativity
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_CPUCACHE_ASSOCIATVITY, cacheNumber, &byteValue ) )
    {
        TranslateCpuCacheAssociativity( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_CPUCACHE_ASSOCIATVITY, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the Port Connector Information (Type 8) as an audit record
//
//  Parameters:
//      portNumber - number of port/structure in SMBIOS data
//      pRecord    - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType8_PortConnectorRecord( WORD portNumber,
                                                      AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      offset = 0;
    String    Value;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_PORTCONN );

    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_8_PORT_CONN, portNumber, &offset ) )
    {
        return false;
    }

    Value = Format.UInt16( portNumber );
    pRecord->Add( PXS_SMBIOS_PORTCONN_ITEM_NUMBER, Value );

    // Internal Reference Designator
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_PORTCONN_INT_REF, portNumber, &Value );
    pRecord->Add( PXS_SMBIOS_PORTCONN_INT_REF,  Value );

    // Internal Connector Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_PORTCONN_INT_CONN, portNumber, &byteValue ) )
    {
        TranslatePortConnectionType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PORTCONN_INT_CONN, Value );

    // External Reference Designator
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_PORTCONN_EXT_REF, portNumber, &Value );
    pRecord->Add( PXS_SMBIOS_PORTCONN_EXT_REF, Value );

    // External Connector Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_PORTCONN_EXT_CONN, portNumber, &byteValue ) )
    {
        TranslatePortConnectionType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PORTCONN_EXT_CONN, Value );

    // Port Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_PORTCONN_PORT_TYPE, portNumber, &byteValue ) )
    {
        TranslatePortType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_PORTCONN_PORT_TYPE, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the System Slot information (Type 9) as an audit record
//
//  Parameters:
//      slotNumber - number of slot/structure in SMBIOS data
//      pRecord    - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType9_SystemSlotRecord( WORD slotNumber, AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      word = 0, offset = 0;
    String    Value;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_SYSSLOT );

    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_9_SYSTEM_SLOT, slotNumber, &offset ) )
    {
        return false;
    }

    Value = Format.UInt16( slotNumber );
    pRecord->Add( PXS_SMBIOS_SYSSLOT_ITEM_NUMBER, Value );

    // Slot Designation
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_SYSSLOT_DESIGNATION, slotNumber, &Value );
    pRecord->Add( PXS_SMBIOS_SYSSLOT_DESIGNATION, Value );

    // Slot Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_SYSSLOT_SLOT_TYPE, slotNumber, &byteValue ) )
    {
        TranslateSystemSlotType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_SLOT_TYPE, Value );

    // Slot Data Bus Width
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_SYSSLOT_BUS_WIDTH, slotNumber, &byteValue ) )
    {
        TranslateSystemSlotWidth( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_BUS_WIDTH, Value );

    // Current Usage
    Value = PXS_STRING_EMPTY;
    if (GetByteValue( PXS_SMBIOS_SYSSLOT_CURRENT_USAGE, slotNumber, &byteValue))
    {
        TranslateSystemSlotUsage( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_CURRENT_USAGE, Value );

    // Slot Length
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_SYSSLOT_SLOT_LENGTH, slotNumber, &byteValue ))
    {
        TranslateSystemSlotLength( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_SLOT_LENGTH, Value );

    // Slot ID
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_SYSSLOT_SLOT_ID, slotNumber, &word ) )
    {
        Value = Format.UInt16Hex( word );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_SLOT_ID, Value );

    // Characteristics 1
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_SYSSLOT_C_1, slotNumber, &byteValue ) )
    {
        TranslateSystemSlotCharac1( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_C_1, Value );

    // Characteristics 2
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_SYSSLOT_C_2, slotNumber, &byteValue ) )
    {
        TranslateSystemSlotCharac2( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_C_2, Value );

    // Segment Group Number
    word  = 0x00FF;      // Indicates not used
    Value = PXS_STRING_EMPTY;
    GetWordValue( PXS_SMBIOS_SYSSLOT_SEGMENT_NUM, slotNumber, &word );
    if ( word != 0x00FF )
    {
        Value = Format.UInt16Hex( word );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_SEGMENT_NUM, Value );

    // Bus Number
    word  = 0x00FF;      // Indicates not used
    Value = PXS_STRING_EMPTY;
    GetByteValue( PXS_SMBIOS_SYSSLOT_BUS_NUM, slotNumber, &byteValue );
    if ( word != 0x00FF )
    {
        Value = Format.UInt16Hex( word );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_BUS_NUM, Value );

    // Device/Function Number
    Value = PXS_STRING_EMPTY;
    GetByteValue( PXS_SMBIOS_SYSSLOT_DEVICE_NUM, slotNumber, &byteValue );
    if ( byteValue != 0xFF )
    {
        // Bits 7:3 - device number
        Value  = L"Device ";
        Value += Format.Int32( byteValue >> 3 );

        // Bits 2:0 - function number
        Value  = L", Function ";
        Value += Format.Int32( 0x7 & byteValue );
    }
    pRecord->Add( PXS_SMBIOS_SYSSLOT_DEVICE_NUM, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the Physical Memory Array information (Type 16) as an audit record
//
//  Parameters:
//      arrayNumber - number of memory array/structure in SMBIOS data
//      pRecord     - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType16_MemoryArrayRecord( WORD arrayNumber,
                                                     AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      word  = 0, offset = 0;
    DWORD     dword = 0, maximumCapacity = 0;
    QWORD     qwordValue = 0;
    String    Value, LocaleKB;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_MEMARRAY );

    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_16_MEM_ARRAY,
                                      arrayNumber, &offset ) )
    {
        return false;
    }

    // Locale string
    PXSGetResourceString( PXS_IDS_135_KB, &LocaleKB );

    Value = Format.UInt16( arrayNumber );
    pRecord->Add( PXS_SMBIOS_MEMARRAY_ITEM_NUMBER, Value );

    // Location
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMARRAY_LOCATION, arrayNumber, &byteValue) )
    {
        TranslateMemoryArrayLocation( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMARRAY_LOCATION, Value );

    // Use
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMARRAY_USE, arrayNumber, &byteValue ) )
    {
        TranslateMemoryArrayUse( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMARRAY_USE, Value );

    // Memory Error Correction
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEMARRAY_ERR_CORRECT, arrayNumber, &byteValue ) )
    {
        TranslateMemoryArrayErrCorrection( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEMARRAY_ERR_CORRECT, Value );

    // Maximum Capacity
    maximumCapacity = 0x80000000;     // Indicates not used
    Value = PXS_STRING_EMPTY;
    GetDWordValue ( PXS_SMBIOS_MEMARRAY_MAX_CAP_KB, arrayNumber, &maximumCapacity );
    if ( maximumCapacity != 0x80000000 )
    {
        Value  = Format.UInt32( dword );
        Value += LocaleKB;
    }
    pRecord->Add( PXS_SMBIOS_MEMARRAY_MAX_CAP_KB, PXS_STRING_EMPTY );

    // Memory Error Information Handle
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_MEMARRAY_ERROR_HANDLE, arrayNumber, &word ) )
    {
        Value = Format.UInt16( word );
    }
    pRecord->Add( PXS_SMBIOS_MEMARRAY_ERROR_HANDLE, Value );

    // Number of Memory Devices
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_MEMARRAY_NUM_DEVICES, arrayNumber, &word ) )
    {
        Value = Format.UInt16( word );
    }
    pRecord->Add( PXS_SMBIOS_MEMARRAY_NUM_DEVICES, Value );

    // Extended Maximum Capacity, only used if maximum capacity is 0x80000000
    Value = PXS_STRING_EMPTY;
    if ( maximumCapacity == 0x80000000 )
    {
        GetQWordValue(PXS_SMBIOS_MEMARRAY_EX_MAX_CAP, arrayNumber, &qwordValue);
        Value = Format.UInt64( qwordValue );
    }
    pRecord->Add( PXS_SMBIOS_MEMARRAY_EX_MAX_CAP, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the Memory Device information (Type 17) as an audit record
//
//  Parameters:
//      deviceNumber - number of device/structure in SMBIOS data
//      pRecord      - receives the data
//
//  Returns:
//      true on success, else false.
//===============================================================================================//
bool SmbiosInformation::GetType17_MemoryDeviceRecord( WORD deviceNumber,
                                                      AuditRecord* pRecord ) const
{
    BYTE      byteValue = 0;
    WORD      word  = 0, offset = 0, deviceSize = 0;
    DWORD     dwordValue = 0;
    String    Value, LocaleKB, LocaleMB, LocaleBit, LocaleMHz;
    Formatter Format;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SMBIOS_MEMDEV );

    if ( false == GetStructureOffset( PXS_SMBIOS_TYPE_17_MEMORY_DEVICE, deviceNumber, &offset ) )
    {
       return false;
    }

    // Locale strings
    PXSGetResourceString( PXS_IDS_136_MB  , &LocaleMB );
    PXSGetResourceString( PXS_IDS_135_KB  , &LocaleKB );
    PXSGetResourceString( PXS_IDS_1263_MHZ, &LocaleMHz );
    PXSGetResourceString( PXS_IDS_1264_BIT, &LocaleBit );

    Value = Format.UInt16( deviceNumber );
    pRecord->Add( PXS_SMBIOS_MEM_DEV_ITEM_NUMBER, Value );

    // Memory Error Information Handle
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_MEM_DEV_ERROR_HANDLE,
                       deviceNumber, &word ) )
    {
        Value = Format.UInt16Hex( word );
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_ERROR_HANDLE, Value );

    // Total Width
    word  = 0xFFFF;      // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetWordValue(PXS_SMBIOS_MEM_DEV_TOT_WIDTH_BIT, deviceNumber, &word );
    if ( word != 0xFFFF )
    {
        Value  = Format.UInt16( word );
        Value += LocaleBit;
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_TOT_WIDTH_BIT, Value );

    // Data Width
    word  =  0xFFFF;     // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetWordValue( PXS_SMBIOS_MEM_DEV_DAT_WIDTH_BIT, deviceNumber, &word );
    if ( word != 0xFFFF )
    {
        Value = Format.UInt16( word );
        Value += LocaleBit;
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_DAT_WIDTH_BIT, Value );

    // Size, store this value for later use
    deviceSize = 0xFFFF;      // Signifies unknown
    Value      = PXS_STRING_EMPTY;
    GetWordValue( PXS_SMBIOS_MEM_DEV_SIZE, deviceNumber, &deviceSize );
    if ( deviceSize != 0xFFFF )
    {
        // If bit 15 is 1 then in KB else in MB
        if ( 0x8000 & deviceSize )
        {
            Value  = Format.Int32( 0x7FFF & deviceSize );
            Value += LocaleKB;
        }
        else
        {
            Value  = Format.Int32( deviceSize );
            Value += LocaleMB;
        }
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_SIZE, Value );

    // Form Factor
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEM_DEV_FORM_FACTOR,
                       deviceNumber, &byteValue ) )
    {
        TranslateMemoryDeviceFormFactor( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_FORM_FACTOR, Value );

    // Device Set
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEM_DEV_DEVICE_SET, deviceNumber, &byteValue ) )
    {
        TranslateMemoryDeviceSet( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_DEVICE_SET, Value );

    // Device Locator
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_MEM_DEV_DEVICE_LOCAT, deviceNumber, &Value );
    pRecord->Add( PXS_SMBIOS_MEM_DEV_DEVICE_LOCAT, Value );

    // Bank Locator
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_MEM_DEV_BANK_LOCATOR, deviceNumber, &Value );
    pRecord->Add( PXS_SMBIOS_MEM_DEV_BANK_LOCATOR, Value );

    // Memory Type
    Value = PXS_STRING_EMPTY;
    if ( GetByteValue( PXS_SMBIOS_MEM_DEV_MEMORY_TYPE, deviceNumber, &byteValue ) )
    {
        TranslateMemoryDeviceType( byteValue, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_MEMORY_TYPE, Value );

    // Type Detail
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_MEM_DEV_TYPE_DETAIL, deviceNumber, &word ) )
    {
        TranslateMemoryDeviceTypeDetail( word, &Value );
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_TYPE_DETAIL, Value );

    // Speed
    Value = PXS_STRING_EMPTY;
    if ( GetWordValue( PXS_SMBIOS_MEM_DEV_SPEED_MHZ, deviceNumber, &word ) )
    {
        Value = Format.UInt16( word );
        Value += LocaleMHz;
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_SPEED_MHZ, Value );

    // Manufacturer
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_MEM_DEV_MANUFACTURER, deviceNumber, &Value );
    pRecord->Add( PXS_SMBIOS_MEM_DEV_MANUFACTURER, Value );

    // Serial Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_MEM_DEV_SERIAL_NUMBER, deviceNumber, &Value );
    pRecord->Add( PXS_SMBIOS_MEM_DEV_SERIAL_NUMBER, Value );

    // Asset Tag
    Value = PXS_STRING_EMPTY;
    GetStringValue(PXS_SMBIOS_MEM_DEV_ASSET_TAG, deviceNumber, &Value );
    pRecord->Add( PXS_SMBIOS_MEM_DEV_ASSET_TAG, Value );

    // Part Number
    Value = PXS_STRING_EMPTY;
    GetStringValue( PXS_SMBIOS_MEM_DEV_PART_NUMBER, deviceNumber, &Value );
    pRecord->Add( PXS_SMBIOS_MEM_DEV_PART_NUMBER, Value );

    // Attributes
    byteValue  = 0;       // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetByteValue( PXS_SMBIOS_MEM_DEV_ATTRIBUTES, deviceNumber, &byteValue );
    if ( byteValue )
    {
        Value = Format.Int32( 0x0F & byteValue );
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_ATTRIBUTES, Value );

    // Extended Size, only applicable if Size is 0x7FFFF
    Value = PXS_STRING_EMPTY;
    if ( deviceSize == 0x7FFF )
    {
        if ( GetDWordValue( PXS_SMBIOS_MEM_DEV_EXT_SIZE_GB, deviceNumber, &dwordValue ) )
        {
            // Bits 30:0 represent the size of the memory device in megabytes.
            // Convert to GB as the specification shows
            Value = Format.UInt32( ( 0x7FFFFFFF & dwordValue ) / 1024 );
        }
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_EXT_SIZE_GB, Value );

    // Configured Memory Clock Speed
    word  = 0;       // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetWordValue( PXS_SMBIOS_MEM_DEV_CLOCK_MHZ, deviceNumber, &word );
    if ( word )
    {
        Value = Format.UInt16( word );
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_CLOCK_MHZ, Value );

    // Minimum voltage
    word  = 0;       // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetWordValue( PXS_SMBIOS_MEM_DEV_MIN_VOLTAGE, deviceNumber, &word );
    if ( word )
    {
        Value = Format.UInt16( word );
        Value += L"mV";
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_MIN_VOLTAGE, Value );

    // Maximum voltage
    word  = 0;       // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetWordValue( PXS_SMBIOS_MEM_DEV_MAX_VOLTAGE, deviceNumber, &word );
    if ( word )
    {
        Value = Format.UInt16( word );
        Value += L"mV";
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_MAX_VOLTAGE, Value );

    // Configured voltage
    word  = 0;       // Signifies unknown
    Value = PXS_STRING_EMPTY;
    GetWordValue( PXS_SMBIOS_MEM_DEV_CONF_VOLTAGE, deviceNumber, &word );
    if ( word )
    {
        Value = Format.UInt16( word );
        Value += L"mV";
    }
    pRecord->Add( PXS_SMBIOS_MEM_DEV_CONF_VOLTAGE, Value );

    return true;
}

//===============================================================================================//
//  Description:
//      Translate a base-board type value
//
//  Parameters:
//      type         - the base-board type
//      pTranslation - receives the translation
//
//  Remarks:
//      Baseboard (or Module) Information (Type 2)
//      0Dh, Board Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateBaseBoardType( BYTE type,
                                                String* pTranslation ) const
{
    size_t   i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    type;
        LPCWSTR pszBoardType;
    } Types[] = { { 0x00, L""                        },
                  { 0x01, L"Other"                   },
                  { 0x02, L"Unknown"                 },
                  { 0x03, L"Server Blade"            },
                  { 0x04, L"Connectivity Switch"     },
                  { 0x05, L"System Management Module"},
                  { 0x06, L"Process Module"          },
                  { 0x07, L"I/O Module"              },
                  { 0x08, L"Memory Module"           },
                  { 0x09, L"Daughter Board"          },
                  { 0x0A, L"Mother Board"            },
                  { 0x0B, L"Processor/Memory Module" },
                  { 0x0C, L"Processor/IO Module"    },
                  { 0x0D, L"Interconnect Board"     } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszBoardType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS baseboard type: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate the base board features flag value
//
//  Parameters:
//      flags        - the features
//      pTranslation - receives the translation
//
//  Remarks:
//      Base Board (or Module) Information (Type 2)
//      09h, Feature Flags, BYTE, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateBoardFeaturesFlags(BYTE flags, String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Features[] = { L"Motherboard",
                           L"Requires Daughterboard",
                           L"Removable",
                           L"Replacable",
                           L"Hot Swapable" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Features ); i++ )
    {
        if ( flags & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Features[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a chassis security status value
//
//  Parameters:
//      status       - the security status
//      pTranslation - receives the translation
//
//  Remarks:
//      System Enclosure or Chassis (Type 3)
//      0Ch, 2.1+, Security Status, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateChassisSecurityStatus( BYTE status,
                                                        String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _STATUS
    {
        BYTE    status;
        LPCWSTR pszSecurityStatus;
    } Status[] = { { 0x01, L"Other"                        },
                   { 0x02, L"Unknown"                      },
                   { 0x03, L"None"                         },
                   { 0x04, L"External interface locked out"},
                   { 0x05, L"External interface enabled"   } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Status ); i++ )
    {
        if ( status == Status[ i ].status )
        {
            *pTranslation = Status[ i ].pszSecurityStatus;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( status );
        PXSLogAppWarn1( L"Unrecognised SMBIOS chassis security status: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a chassis state value
//
//  Parameters:
//      state        - the state
//      pTranslation - receives the translation
//
//  Remarks:
//      System Enclosure or Chassis (Type 3)
//      09h, 2.1+, Boot-up State     , BYTE, ENUM
//      0Ah, 2.1+, Power Supply State, BYTE, ENUM
//      0Bh, 2.1+, Thermal State     , BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateChassisState( BYTE state, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _STATE
    {
        BYTE    state;
        LPCWSTR pszChassisState;
    } States[] = { { 0x00, L""                },
                   { 0x01, L"Other"           },
                   { 0x02, L"Unknown"         },
                   { 0x03, L"Safe"            },
                   { 0x04, L"Warning"         },
                   { 0x05, L"Critical"        },
                   { 0x06, L"Non-Recoverable" } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( States ); i++ )
    {
        if ( state == States[ i ].state )
        {
            *pTranslation = States[ i ].pszChassisState;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( state );
        PXSLogAppWarn1( L"Unrecognised SMBIOS chassis state: %%1.",
                        *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a chassis type value
//
//  Parameters:
//      type         - the chassis type
//      pTranslation - receives the translation
//
//  Remarks:
//      System Enclosure or Chassis (Type 3)
//      05h, 2.0+, Type, BYTE, Varies
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateChassisType( BYTE type, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    type;
        LPCWSTR pszChassisType;
    } Types[] = { { 0x00, L""                     },
                  { 0x01, L"Other"                },
                  { 0x02, L"Unknown"              },
                  { 0x03, L"Desktop"              },
                  { 0x04, L"Low Profile Desktop"  },
                  { 0x05, L"Pizza Box"            },
                  { 0x06, L"Mini Tower"           },
                  { 0x07, L"Tower"                },
                  { 0x08, L"Portable"             },
                  { 0x09, L"Laptop"               },
                  { 0x0A, L"Notebook"             },
                  { 0x0B, L"Hand Held"            },
                  { 0x0C, L"Docking Station"      },
                  { 0x0D, L"All in One"           },
                  { 0x0E, L"Sub Notebook"         },
                  { 0x0F, L"Space Saving"         },
                  { 0x10, L"Lunch Box"            },
                  { 0x11, L"Main Server Chassis"  },
                  { 0x12, L"Expansion Chassis"    },
                  { 0x13, L"Sub Chassis"          },
                  { 0x14, L"Bus Expansion Chassis"},
                  { 0x15, L"Peripheral Chassis"   },
                  { 0x16, L"RAID Chassis"         },
                  { 0x17, L"Rack Mount Chassis"   },
                  { 0x18, L"Sealed Case PC"       },
                  { 0x19, L"Multi-System Chassis" },
                  { 0x1A, L"CompactPCI"           },
                  { 0x1B, L"AdvancedTCA"          },
                  { 0x1C, L"Blade"                },
                  { 0x1D, L"Blade Enclosure"      } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszChassisType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS chassis type: %%1.", *pTranslation );
    }

    // Bit 7 represents the presence of a chassis lock
    if ( 0x80 & type )
    {
        *pTranslation += L", Chassis lock present";
    }
}

//===============================================================================================//
//  Description:
//      Translate a CPU cache associativity value
//
//  Parameters:
//      associativity - the CPU cache associativity
//      pTranslation  - receives the translation
//
//  Remarks:
//      Cache Information (Type 7)
//      12h, 2.1+, Associativity, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateCpuCacheAssociativity( BYTE associativity,
                                                        String* pTranslation ) const
{
    size_t   i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    associativity;
        LPCWSTR pszCacheAssociativity;
    } Associativities[] = { { 0x01, L"Other"                  },
                            { 0x02, L"Unknown"                },
                            { 0x03, L"Direct Mapped"          },
                            { 0x04, L"2-way Set-Associative"  },
                            { 0x05, L"4-way Set-Associative"  },
                            { 0x06, L"Fully Associative"      },
                            { 0x07, L"8-way Set-Associative"  },
                            { 0x08, L"16-way Set-Associative" },
                            { 0x09, L"12-way Set-Associative" },
                            { 0x0A, L"24-way Set-Associative" },
                            { 0x0B, L"32-way Set-Associative" },
                            { 0x0C, L"48-way Set-Associative" },
                            { 0x0D, L"64-way Set-Associative" },
                            { 0x0E, L"20-way Set-Associative" } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Associativities ); i++ )
    {
        if ( associativity == Associativities[ i ].associativity )
        {
            *pTranslation = Associativities[ i ].pszCacheAssociativity;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( associativity );
        PXSLogAppWarn1( L"Unrecognised SMBIOS CPU cache associativity: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a CPU cache configuration value
//
//  Parameters:
//      configuration - the security status
//      pTranslation  - receives the translation
//
//  Remarks:
//      Cache Information (Type 7)
//      05h, 2.0+, Cache Configuration, WORD, Varies
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateCpuCacheConfiguration( WORD configuration,
                                                        String* pTranslation ) const
{
    int       bits = 0;
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }

    // Bits 2:0 Cache Level
    *pTranslation  = L"Level ";
    *pTranslation += Format.Int32( 1 + ( configuration & 0x7 ) );

    // Bit 3 Cache Socketed
    if ( 0x8 & configuration )
    {
        *pTranslation += L", Socketed";
    }
    else
    {
        *pTranslation += L", Not Socketed";
    }

    // Bits 6:5 Location, relative to the CPU module:
    bits = 0x3 & ( configuration >> 5 );
    if ( bits == 0 )
    {
        *pTranslation += L", Internal";
    }
    else if ( bits == 1 )
    {
        *pTranslation += L", External";
    }
    else if ( bits == 3 )
    {
        *pTranslation += L", Unknown";
    }

    // Bit 7 Enabled/Disabled (at boot time)
    if ( 0x80 & configuration )
    {
        *pTranslation += L", Enabled";
    }
    else
    {
        *pTranslation += L", Disabled";
    }

    // Bits 9:8 Operational Mode
    bits = 0x3 & ( configuration >> 8 );
    if ( bits == 0 )
    {
        *pTranslation += L", Write Through";
    }
    else if ( bits == 1 )
    {
        *pTranslation += L", Write Back";
    }
    else if ( bits == 2 )
    {
        *pTranslation += L", Varies with Memory Address";
    }
    else if ( bits == 3 )
    {
        *pTranslation += L", Unknown";
    }
}

//===============================================================================================//
//  Description:
//      Translate a CPU cache error correction value
//
//  Parameters:
//      correction   - the error detection value
//      pTranslation - receives the translation
//
//  Remarks:
//      Cache Information (Type 7)
//      10h, 2.1+, Error Correction Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateCpuCacheErrCorrection( BYTE correction,
                                                        String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _CORRECTION
    {
        BYTE    correction;
        LPCWSTR pszErrCorrection;
    } Corrections[] = { { 0x01, L"Other"         },
                        { 0x02, L"Unknown"       },
                        { 0x03, L"None"          },
                        { 0x04, L"Parity"        },
                        { 0x05, L"Single-bit ECC" },
                        { 0x06, L"Multi-bit ECC" } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Corrections ); i++ )
    {
        if ( correction == Corrections[ i ].correction )
        {
            *pTranslation = Corrections[ i ].pszErrCorrection;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
       *pTranslation = Format.UInt8( correction );
       PXSLogAppWarn1( L"Unrecognised SMBIOS CPU error correction: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate CPU cache SRAM memory type
//
//  Parameters:
//      type         - the SRAM memory type
//      pTranslation - receives the translation
//
//  Remarks:
//      Cache Information (Type 7)
//      0Bh, 2.0+, Supported SRAM Type, WORD, Bit Field
//      0Dh, 2.0+, Current SRAM Type  , WORD, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateCpuCacheSramType( WORD type, String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Types[] = { L"Other",
                        L"Unknown",
                        L"Non-Burst",
                        L"Burst",
                        L"Pipeline Burst",
                        L"Synchronous",
                        L"Asynchronous" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Types[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a CPU cache type value
//
//  Parameters:
//      type         - the error detection value
//      pTranslation - receives the translation
//
//  Remarks:
//      Cache Information (Type 7)
//      11h, 2.1+, System Cache Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateCpuCacheType( BYTE type, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    type;
        LPCWSTR pszCacheType;
    } Types[] = { { 0x01, L"Other"       },
                  { 0x02, L"Unknown"     },
                  { 0x03, L"Instruction" },
                  { 0x04, L"Data"        },
                  { 0x05, L"Unified"     } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszCacheType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS CPU cache type: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory array error correction value
//
//  Parameters:
//      correction   - the error detection value
//      pTranslation - receives the translation
//
//  Remarks:
//      Physical Memory Array (Type 16)
//      06h, 2.1+, Memory Error Correction, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryArrayErrCorrection( BYTE correction,
                                                           String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _CORRECTION
    {
        BYTE    correction;
        LPCWSTR pszErrCorrection;
    } Corrections[] = { { 0x01, L"Other"         },
                        { 0x02, L"Unknown"       },
                        { 0x03, L"None"          },
                        { 0x04, L"Parity"        },
                        { 0x05, L"Single-bit ECC"},
                        { 0x06, L"Multi-bit ECC" },
                        { 0x07, L"CRC"           } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Corrections ); i++ )
    {
        if ( correction == Corrections[ i ].correction )
        {
            *pTranslation = Corrections[ i ].pszErrCorrection;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( correction );
        PXSLogAppWarn1( L"Unrecognised SMBIOS memory array correction: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory array location value
//
//  Parameters:
//      location     - the location value
//      pTranslation - receives the translation
//
//  Remarks:
//      Physical Memory Array (Type 16)
//      04h, 2.1+, Location, BYTE, ENUM,
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryArrayLocation( BYTE location,
                                                      String* pTranslation ) const
{
    size_t   i = 0;
    Formatter Format;

    struct _LOCATION
    {
        BYTE    location;
        LPCWSTR pszArrayLocation;
    } Locations[] = { { 0x01, L"Other"                       },
                      { 0x02, L"Unknown"                     },
                      { 0x03, L"System board or motherboard" },
                      { 0x04, L"ISA add-on card"             },
                      { 0x05, L"EISA add-on card"            },
                      { 0x06, L"PCI add-on card"             },
                      { 0x07, L"MCA add-on card"             },
                      { 0x08, L"PCMCIA add-on card"          },
                      { 0x09, L"Proprietary add-on card"     },
                      { 0x0A, L"NuBus"                       },
                      { 0xA0, L"PC-98/C20 add-on card"       },
                      { 0xA1, L"PC-98/C24 add-on card"       },
                      { 0xA2, L"PC-98/E add-on card"         },
                      { 0xA3, L"PC-98/Local bus add-on card" } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Locations ); i++ )
    {
        if ( location == Locations[ i ].location )
        {
            *pTranslation = Locations[ i ].pszArrayLocation;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( location );
        PXSLogAppWarn1( L"Unrecognised SMBIOS memory array location: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory array use value
//
//  Parameters:
//      use          - the use value
//      pTranslation - receives the translation
//
//  Remarks:
//      Physical Memory Array (Type 16)
//      05h, 2.1+, Use, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryArrayUse( BYTE use, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _USE
    {
        BYTE    use;
        LPCWSTR pszArrayUse;
    } Uses[] = { { 0x01, L"Other"            },
                 { 0x02, L"Unknown"          },
                 { 0x03, L"System memory"    },
                 { 0x04, L"Video memory"     },
                 { 0x05, L"Flash memory"     },
                 { 0x06, L"Non-volatile RAM" },
                 { 0x07, L"Cache memory"     } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Uses ); i++ )
    {
        if ( use == Uses[ i ].use )
        {
            *pTranslation = Uses[ i ].pszArrayUse;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( use );
        PXSLogAppWarn1( L"Unrecognised SMBIOS memory array use: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory controller error correction value
//
//  Parameters:
//      correction   - the error detection value
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Controller Information (Type 5, Obsolete)
//      05h, 2.0+, Error Correcting Capability, BYTE, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryControllerErrCorrection( BYTE correction,
                                                                String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Corrections[] = { L"Other",
                              L"Unknown",
                              L"None",
                              L"Single-Bit Error Correcting",
                              L"Double-Bit Error Correcting",
                              L"Error Scrubbing" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Corrections ); i++ )
    {
        if ( correction & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation = Corrections[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory controller error detection value
//
//  Parameters:
//      detection    - the error detection value
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Controller Information (Type 5, Obsolete)
//      04h, 2.0+, Error Detecting Method, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryControllerErrDetect( BYTE detection,
                                                            String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _DETECT
    {
        BYTE    detection;
        LPCWSTR pszErrDetection;
    } Detections[] = { { 0x01, L"Other"        },
                       { 0x02, L"Unknown"      },
                       { 0x03, L"None"         },
                       { 0x04, L"8-bit Parity" },
                       { 0x05, L"32-bit ECC"   },
                       { 0x06, L"64-bit ECC"   },
                       { 0x07, L"128-bit ECC"  },
                       { 0x08, L"CRC"          } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Detections ); i++ )
    {
        if ( detection == Detections[ i ].detection )
        {
            *pTranslation = Detections[ i ].pszErrDetection;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( detection );
        PXSLogAppWarn1( L"Unrecognised SMBIOS controller error detect: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory controller interleave method
//
//  Parameters:
//      interleave   - the interleave method
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Controller Information (Type 5, Obsolete)
//      06h, 2.0+, Supported Interleave, BYTE, ENUM
//      07h, 2.0+, Current Interleave  , BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryControllerInterleave( BYTE interleave,
                                                             String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _INTERLEAVE
    {
        BYTE    interleave;
        LPCWSTR pszControllerInterleave;
    } Interleaves[] = { { 0x01, L"Other"                 },
                        { 0x02, L"Unknown"               },
                        { 0x03, L"One-Way Interleave"    },
                        { 0x04, L"Two-Way Interleave"    },
                        { 0x05, L"Four-Way Interleave"   },
                        { 0x06, L"Eight-Way Interleave"  },
                        { 0x07, L"Sixteen-Way Interleave"} };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Interleaves ); i++ )
    {
        if ( interleave == Interleaves[ i ].interleave )
        {
            *pTranslation = Interleaves[ i ].pszControllerInterleave;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( interleave );
        PXSLogAppWarn1( L"Unrecognised SMBIOS controller interleave: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory controller speed
//
//  Parameters:
//      speed        - the memory controller speed
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Controller Information (Type 5, Obsolete)
//      09h, 2.0+, Supported Speeds, WORD, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryControllerSpeed( WORD speed,
                                                        String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Speeds[] = { L"Other",
                         L"Unknown",
                         L"70ns",
                         L"60ns",
                         L"50ns" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Speeds ); i++ )
    {
        if ( speed & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Speeds[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory controller voltage value
//
//  Parameters:
//      voltage      - the memory controller voltage
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Controller Information (Type 5, Obsolete)
//      0Dh, 2.0+, Memory Module Voltage, BYTE, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryControllerVoltage( BYTE voltage,
                                                          String* pTranslation ) const
{
    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    if ( 0x01 & voltage ) *pTranslation += L"5.0V, ";
    if ( 0x02 & voltage ) *pTranslation += L"3.3V, ";
    if ( 0x04 & voltage ) *pTranslation += L"2.9V";

    // Clean up
    pTranslation->Trim();
    if ( pTranslation->EndsWithCharacter( PXS_CHAR_COMMA ) )
    {
        pTranslation->Truncate( pTranslation->GetLength() - 1 );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory device form factor
//
//  Parameters:
//      form         - the form factor
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Device (Type 17)
//      0Eh, 2.1+, Form Factor, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryDeviceFormFactor( BYTE form,
                                                         String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _FORM
    {
        BYTE    form;
        LPCWSTR pszFormFactor;
    } Forms[] = { { 0x01, L"Other"            },
                  { 0x02, L"Unknown"          },
                  { 0x03, L"SIMM"             },
                  { 0x04, L"SIP"              },
                  { 0x05, L"Chip"             },
                  { 0x06, L"DIP"              },
                  { 0x07, L"ZIP"              },
                  { 0x08, L"Proprietary Card" },
                  { 0x09, L"DIMM"             },
                  { 0x0A, L"TSOP"             },
                  { 0x0B, L"Row of chips"     },
                  { 0x0C, L"RIMM"             },
                  { 0x0D, L"SODIMM"           },
                  { 0x0E, L"SRIMM"            },
                  { 0x0F, L"FB-DIMM"          } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Forms ); i++ )
    {
        if ( form == Forms[ i ].form )
        {
            *pTranslation = Forms[ i ].pszFormFactor;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
      *pTranslation = Format.UInt8( form );
      PXSLogAppWarn1( L"Unrecognised SMBIOS memory device form factor: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory device set value
//
//  Parameters:
//      set          - the set value
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Device (Type 17)
//      0Fh, 2.1+, Device Set, BYTE, Varies
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryDeviceSet( BYTE set, String* pTranslation ) const
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }

    if ( set == 0 )
    {
        *pTranslation = L"Unknown";
    }
    else if ( set == 0xFF )
    {
        *pTranslation = L"Not part of a set";
    }
    else
    {
        *pTranslation = Format.UInt8( set );
    }
}


//===============================================================================================//
//  Description:
//      Translate a memory device type value
//
//  Parameters:
//      type         - the device type
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Device (Type 17)
//      12h, 2.1+, Memory Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryDeviceType( BYTE type, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    type;
        LPCWSTR pszDeviceType;
    } Types[] = { { 0x01, L"Other"        },
                  { 0x02, L"Unknown"      },
                  { 0x03, L"DRAM"         },
                  { 0x04, L"EDRAM"        },
                  { 0x05, L"VRAM"         },
                  { 0x06, L"SRAM"         },
                  { 0x07, L"RAM"          },
                  { 0x08, L"ROM"          },
                  { 0x09, L"FLASH"        },
                  { 0x0A, L"EEPROM"       },
                  { 0x0B, L"FEPROM"       },
                  { 0x0C, L"EPROM"        },
                  { 0x0D, L"CDRAM"        },
                  { 0x0E, L"3DRAM"        },
                  { 0x0F, L"SDRAM"        },
                  { 0x10, L"SGRAM"        },
                  { 0x11, L"RDRAM"        },
                  { 0x12, L"DDR"          },
                  { 0x13, L"DDR2"         },
                  { 0x14, L"DDR2 FB-DIMM" },
                  { 0x18, L"DDR3"         },
                  { 0x19, L"FBD2"         } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszDeviceType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS memory device type: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory device type detail
//
//  Parameters:
//      detail       - the type detail
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Device (Type 17)
//      13h, 2.1+, Type Detail, WORD, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryDeviceTypeDetail( WORD detail,
                                                         String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Details[] = { L"",
                          L"Other",
                          L"Unknown",
                          L"Fast-paged",
                          L"Static column",
                          L"Pseudo-static",
                          L"RAMBUS",
                          L"Synchronous" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Details ); i++ )
    {
        if ( detail & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Details[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory module bank connection value
//
//  Parameters:
//      bank         - the memory module bank connection value
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Module Information (Type 6, Obsolete)
//      05h, Bank Connections, BYTE, Varies
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryModuleBank( BYTE bank, String* pTranslation ) const
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    // Low nibble
    if ( ( bank & 0xF ) != 0xF )
    {
        *pTranslation  = L"RAS-";
        *pTranslation += Format.Int32( bank & 0xF );
    }

    // High nibble
    if ( ( bank >> 4 ) != 0xF )
    {
        if ( pTranslation->GetLength() )
        {
            *pTranslation += L", ";
        }
        *pTranslation  = L"RAS-";
        *pTranslation += Format.Int32( bank >> 4 );
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory module error status
//
//  Parameters:
//      status       - the memory module bank connection value
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Module Information (Type 6, Obsolete)
//      0Bh, Error Status, BYTE, Varies
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryModuleErrorStatus( BYTE status,
                                                          String* pTranslation ) const
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    if ( 0x04 & status )
    {
        *pTranslation = L"Error status information should be obtained "
                        L"from the SMBIOS event log";
    }
    else
    {
        // The value is described as "Varies" so will treat as an enumeration
        // rather than a bit field
        if ( 0x02 & status )
        {
            *pTranslation = L"Correctable errors received for the module";
        }
        else if ( 0x01 & status )
        {
            *pTranslation = L"Uncorrectable errors received for the module";
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a memory module bank connection value
//
//  Parameters:
//      size         - the memory module's size
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Module Information (Type 6, Obsolete)
//      09h, Installed Size, BYTE, Varies
//      0Ah, Enabled Size  , BYTE, Varies
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryModuleSize( BYTE size, String* pTranslation ) const
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    // Bits 0:6
    if ( ( size & 0x7F ) == 0x7D )
    {
        *pTranslation = L"Not determinable (Installed Size only)";
    }
    else if ( ( size & 0x7F ) == 0x7E )
    {
        *pTranslation = L"Module is installed, but no memory has been enabled";
    }
    else if ( ( size & 0x7F ) == 0x7F )
    {
        *pTranslation = L"Not installed";
    }
    else
    {
        // Size is 2**n in MB
        if (  size < 32 )
        {
            *pTranslation  = Format.Int32( 1 << size );
            *pTranslation += L"MB";
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate the a memory controller or module memory type
//
//  Parameters:
//      type         - the memory type
//      pTranslation - receives the translation
//
//  Remarks:
//      Memory Controller Information (Type 5, Obsolete)
//      0Bh, 2.0+, Supported Memory Types, WORD, Bit Field
//
//      Memory Module Information (Type 6, Obsolete)
//      07h, Current Memory Type, WORD, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateMemoryModuleType( WORD type, String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Types[] = { L"Other",
                        L"Unknown",
                        L"Standard",
                        L"Fast Page Mode",
                        L"EDO",
                        L"Parity",
                        L"ECC",
                        L"SIMM",
                        L"DIMM",
                        L"Burst EDO" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Types[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a port connection type value
//
//  Parameters:
//      bType        - the port connection type
//      pTranslation - receives the translation
//
//  Remarks:
//      Processor Information (Type 4)
//      05h, Internal Connector Type, BYTE, ENUM
//      07h, External Connector Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslatePortConnectionType( BYTE type,
                                                     String* pTranslation ) const
{
    size_t   i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    type;
        LPCWSTR pszConnectionType;
    } Types[] = { { 0x00, L"None"                            },
                  { 0x01, L"Centronics"                      },
                  { 0x02, L"Mini Centronics"                 },
                  { 0x03, L"Proprietary"                     },
                  { 0x04, L"DB-25 pin male"                  },
                  { 0x05, L"DB-25 pin female"                },
                  { 0x06, L"DB-15 pin male"                  },
                  { 0x07, L"DB-15 pin female"                },
                  { 0x08, L"DB-9 pin male"                   },
                  { 0x09, L"DB-9 pin female"                 },
                  { 0x0A, L"RJ-11"                           },
                  { 0x0B, L"RJ-45"                           },
                  { 0x0C, L"50-pin MiniSCSI"                 },
                  { 0x0D, L"Mini-DIN"                        },
                  { 0x0E, L"Micro-DIN"                       },
                  { 0x0F, L"PS/2"                            },
                  { 0x10, L"Infrared"                        },
                  { 0x11, L"HP-HIL"                          },
                  { 0x12, L"Access Bus (USB)"                },
                  { 0x13, L"SSA SCSI"                        },
                  { 0x14, L"Circular DIN-8 male"             },
                  { 0x15, L"Circular DIN-8 female"           },
                  { 0x16, L"On Board IDE"                    },
                  { 0x17, L"On Board Floppy"                 },
                  { 0x18, L"9-pin Dual Inline (pin 10 cut)"  },
                  { 0x19, L"25-pin Dual Inline (pin 26 cut)" },
                  { 0x1A, L"50-pin Dual Inline"              },
                  { 0x1B, L"68-pin Dual Inline"              },
                  { 0x1C, L"On Board Sound Input from CD-ROM"},
                  { 0x1D, L"Mini-Centronics Type-14"         },
                  { 0x1E, L"Mini-Centronics Type-26"         },
                  { 0x1F, L"Mini-jack (headphones)"          },
                  { 0x20, L"BNC"                             },
                  { 0x21, L"1394"                            },
                  { 0x22, L"SAS/SATA Plug Receptacle"        },
                  { 0xA0, L"PC-98"                           },
                  { 0xA1, L"PC-98Hireso"                     },
                  { 0xA2, L"PC-H98"                          },
                  { 0xA3, L"PC-98Note"                       },
                  { 0xA4, L"PC-98Full"                       },
                  { 0xFF, L"Other"                           } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszConnectionType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS port connector type: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a port type value
//
//  Parameters:
//      type         - the port type
//      pTranslation - receives the translation
//
//  Remarks:
//      Processor Information (Type 4)
//      08h, Port Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslatePortType( BYTE type, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    type;
        LPCWSTR pszPortType;
    } Types[] = { { 0x00, L"None" },
                  { 0x01, L"Parallel Port XT/AT Compatible"},
                  { 0x02, L"Parallel Port PS/2"            },
                  { 0x03, L"Parallel Port ECP"             },
                  { 0x04, L"Parallel Port EPP"             },
                  { 0x05, L"Parallel Port ECP/EPP"         },
                  { 0x06, L"Serial Port XT/AT Compatible"  },
                  { 0x07, L"Serial Port 16450 Compatible"  },
                  { 0x08, L"Serial Port 16550 Compatible"  },
                  { 0x09, L"Serial Port 16550A Compatible "},
                  { 0x0A, L"SCSI Port"                     },
                  { 0x0B, L"MIDI Port"                     },
                  { 0x0C, L"Joy Stick Port"                },
                  { 0x0D, L"Keyboard Port"                 },
                  { 0x0E, L"Mouse Port"                    },
                  { 0x0F, L"SSA SCSI"                      },
                  { 0x10, L"USB"                           },
                  { 0x11, L"FireWire (IEEE P1394)"         },
                  { 0x12, L"PCMCIA Type I2"                },
                  { 0x13, L"PCMCIA Type II"                },
                  { 0x14, L"PCMCIA Type III"               },
                  { 0x15, L"Cardbus"                       },
                  { 0x16, L"Access Bus Port"               },
                  { 0x17, L"SCSI II"                       },
                  { 0x18, L"SCSI Wide"                     },
                  { 0x19, L"PC-98"                         },
                  { 0x1A, L"PC-98-Hireso"                  },
                  { 0x1B, L"PC-H98"                        },
                  { 0x1C, L"Video Port"                    },
                  { 0x1D, L"Audio Port"                    },
                  { 0x1E, L"Modem Port"                    },
                  { 0x1F, L"Network Port"                  },
                  { 0x20, L"SATA"                          },
                  { 0x21, L"SAS"                           },
                  { 0xA0, L"8251 Compatible"               },
                  { 0xA1, L"8251 FIFO Compatible"          },
                  { 0xFF, L"Other"                         } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszPortType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS port type: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a processor characteristics value
//
//  Parameters:
//      word         - the processor characteristics value
//      pTranslation - receives the translation
//
//  Remarks:
//      System Slots (Type 9)
//      26h, 2.5+, Processor Characteristics, WORD, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateProcessorCharacteristics( WORD characteristics,
                                                           String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Values[] = { L"",
                         L"Unknown",
                         L"64-bit Capable",
                         L"Multi-Core",
                         L"Hardware Thread",
                         L"Execute Protection",
                         L"Enhanced Virtualization",
                         L"Power/Performance Control" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Values ); i++ )
    {
        if ( characteristics & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Values[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a processor family value
//
//  Parameters:
//      family       - the processor family value
//      pTranslation - receives the translation
//
//  Remarks:
//      Processor Information (Type 4)
//      06h, 2.0+, Processor Family, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateProcessorFamily( BYTE family, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _FAMILY
    {
        BYTE    family;
        LPCWSTR pszProcessorFamily;
    } Families[] = {
        { 0x00, L""                                                     },
        { 0x01, L"Other"                                                },
        { 0x02, L"Unknown"                                              },
        { 0x03, L"8086"                                                 },
        { 0x04, L"80286"                                                },
        { 0x05, L"Intel386(TM) processor"                               },
        { 0x06, L"Intel486(TM) processor"                               },
        { 0x07, L"8087"                                                 },
        { 0x08, L"80287"                                                },
        { 0x09, L"80387"                                                },
        { 0x0A, L"80487"                                                },
        { 0x0B, L"Pentium(R) processor Family"                          },
        { 0x0C, L"Pentium(R) Pro processor"                             },
        { 0x0D, L"Pentium(R) II processor"                              },
        { 0x0E, L"Pentium(R) processor with MMX(TM) technology"         },
        { 0x0F, L"Intel(R) Celeron(R) processor"                        },
        { 0x10, L"Pentium(R) II Xeon(TM) processor"                     },
        { 0x11, L"Pentium(R) III processor"                             },
        { 0x12, L"M1 Family"                                            },
        { 0x13, L"M2 Family"                                            },
        { 0x14, L"Intel(R) Celeron(R) M processor"                      },
        { 0x15, L"Intel(R) Pentium(R) 4 HT processor"                   },
        { 0x18, L"AMD Duron(TM) Processor Family"                       },
        { 0x19, L"K5 Family"                                            },
        { 0x1A, L"K6 Family"                                            },
        { 0x1B, L"K6-2"                                                 },
        { 0x1C, L"K6-3"                                                 },
        { 0x1D, L"AMD Athlon(TM) Processor Family"                      },
        { 0x1E, L"AMD29000 Family"                                      },
        { 0x1F, L"K6-2+"                                                },
        { 0x20, L"Power PC Family"                                      },
        { 0x21, L"Power PC 601"                                         },
        { 0x22, L"Power PC 603"                                         },
        { 0x23, L"Power PC 603+"                                        },
        { 0x24, L"Power PC 604"                                         },
        { 0x25, L"Power PC 620"                                         },
        { 0x26, L"Power PC x704"                                        },
        { 0x27, L"Power PC 750"                                         },
        { 0x28, L"Intel(R) Core(TM) Duo processor"                      },
        { 0x29, L"Intel(R) Core(TM) Duo mobile processor"               },
        { 0x2A, L"Intel(R) Core(TM) Solo mobile processor"              },
        { 0x2B, L"Intel(R) Atom(TM) processor"                          },
        { 0x30, L"Alpha Family3"                                        },
        { 0x31, L"Alpha 21064"                                          },
        { 0x32, L"Alpha 21066"                                          },
        { 0x33, L"Alpha 21164"                                          },
        { 0x34, L"Alpha 21164PC"                                        },
        { 0x35, L"Alpha 21164a"                                         },
        { 0x36, L"Alpha 21264"                                          },
        { 0x37, L"Alpha 21364"                                          },
        { 0x38, L"AMD Turion(TM) II Ultra Dual-Core Mobile M Processor" },
        { 0x39, L"AMD Turion(TM) II Dual-Core Mobile M Processor Family"},
        { 0x3A, L"AMD Athlon(TM) II Dual-Core M Processor Family"       },
        { 0x3B, L"AMD Opteron(TM) 6100 Series Processor"                },
        { 0x3C, L"AMD Opteron(TM) 4100 Series Processor"                },
        { 0x3D, L"AMD Opteron(TM) 6200 Series Processor"                },
        { 0x3E, L"AMD Opteron(TM) 4200 Series Processor"                },
        { 0x3F, L"AMD FX(TM) Series Processor"                          },
        { 0x40, L"MIPS Family"                                          },
        { 0x41, L"MIPS R4000"                                           },
        { 0x42, L"MIPS R4200"                                           },
        { 0x43, L"MIPS R4400"                                           },
        { 0x44, L"MIPS R4600"                                           },
        { 0x45, L"MIPS R10000"                                          },
        { 0x46, L"AMD C-Series Processor"                               },
        { 0x47, L"AMD E-Series Processor"                               },
        { 0x48, L"AMD A-Series Processor"                               },
        { 0x49, L"AMD G-Series Processor"                               },
        { 0x4A, L"AMD Z-Series Processor"                               },
        { 0x4B, L"AMD R-Series Processor"                               },
        { 0x4C, L"AMD Opteron(TM) 4300 Series Processor"                },
        { 0x4D, L"AMD Opteron(TM) 6300 Series Processor"                },
        { 0x4E, L"AMD Opteron(TM) 3300 Series Processor"                },
        { 0x4F, L"AMD FirePro(TM) Series Processor"                     },
        { 0x50, L"SPARC Family"                                         },
        { 0x51, L"SuperSPARC"                                           },
        { 0x52, L"microSPARC II"                                        },
        { 0x53, L"microSPARC IIep"                                      },
        { 0x54, L"UltraSPARC"                                           },
        { 0x55, L"UltraSPARC II"                                        },
        { 0x56, L"UltraSPARC IIi"                                       },
        { 0x57, L"UltraSPARC III"                                       },
        { 0x58, L"UltraSPARC IIIi"                                      },
        { 0x60, L"68040 Family"                                         },
        { 0x61, L"68xxx"                                                },
        { 0x62, L"68000"                                                },
        { 0x63, L"68010"                                                },
        { 0x64, L"68020"                                                },
        { 0x65, L"68030"                                                },
        { 0x70, L"Hobbit Family"                                        },
        { 0x78, L"Crusoe(TM) TM5000 Family"                             },
        { 0x79, L"Crusoe(TM) TM3000 Family"                             },
        { 0x7A, L"Efficeon(TM) TM8000 Family"                           },
        { 0x80, L"Weitek"                                               },
        { 0x82, L"Itanium(TM) processor"                                },
        { 0x83, L"AMD Athlon(TM) 64 Processor Family"                   },
        { 0x84, L"AMD Opteron(TM) Processor Family"                     },
        { 0x85, L"AMD Sempron(TM) Processor Family"                     },
        { 0x86, L"AMD Turion(TM) Mobile Technology"                     },
        { 0x87, L"Dual-Core AMD Opteron(TM) Processor Family"           },
        { 0x88, L"AMD Athlon(TM) 64 X2 Dual-Core Processor Family"      },
        { 0x89, L"AMD Turion(TM) 64 X2 Mobile Technology"               },
        { 0x8A, L"Quad-Core AMD Opteron(TM) Processor Family"           },
        { 0x8B, L"Third-Generation AMD Opteron(TM) Processor Family"    },
        { 0x8C, L"AMD Phenom(TM) FX Quad-Core Processor Family"         },
        { 0x8D, L"AMD Phenom(TM) X4 Quad-Core Processor Family"         },
        { 0x8E, L"AMD Phenom(TM) X2 Dual-Core Processor Family"         },
        { 0x8F, L"AMD Athlon(TM) X2 Dual-Core Processor Family"         },
        { 0x90, L"PA-RISC Family"                                       },
        { 0x91, L"PA-RISC 8500"                                         },
        { 0x92, L"PA-RISC 8000"                                         },
        { 0x93, L"PA-RISC 7300LC"                                       },
        { 0x94, L"PA-RISC 7200"                                         },
        { 0x95, L"PA-RISC 7100LC"                                       },
        { 0x96, L"PA-RISC 7100"                                         },
        { 0xA0, L"V30 Family"                                           },
        { 0xA1, L"Quad-Core Intel(R) Xeon(R) processor 3200 Series"     },
        { 0xA2, L"Dual-Core Intel(R) Xeon(R) processor 3000 Series"     },
        { 0xA3, L"Quad-Core Intel(R) Xeon(R) processor 5300 Series"     },
        { 0xA4, L"Dual-Core Intel(R) Xeon(R) processor 5100 Series"     },
        { 0xA5, L"Dual-Core Intel(R) Xeon(R) processor 5000 Series"     },
        { 0xA6, L"Dual-Core Intel(R) Xeon(R) processor LV"              },
        { 0xA7, L"Dual-Core Intel(R) Xeon(R) processor ULV"             },
        { 0xA8, L"Dual-Core Intel(R) Xeon(R) processor 7100 Series"     },
        { 0xA9, L"Quad-Core Intel(R) Xeon(R) processor 5400 Series"     },
        { 0xAA, L"Quad-Core Intel(R) Xeon(R) processor"                 },
        { 0xAB, L"Dual-Core Intel(R) Xeon(R) processor 5200 Series"     },
        { 0xAC, L"Dual-Core Intel(R) Xeon(R) processor 7200 Series"     },
        { 0xAD, L"Quad-Core Intel(R) Xeon(R) processor 7300 Series"     },
        { 0xAE, L"Quad-Core Intel(R) Xeon(R) processor 7400 Series"     },
        { 0xAF, L"Multi-Core Intel(R) Xeon(R) processor 7400 Series"    },
        { 0xB0, L"Pentium(R) III Xeon(TM) processor"                    },
        { 0xB1, L"Pentium(R) III Processor with Intel(R) SpeedStep(TM)" },
        { 0xB2, L"Pentium(R) 4 Processor"                               },
        { 0xB3, L"Intel(R) Xeon(TM)"                                    },
        { 0xB4, L"AS400 Family"                                         },
        { 0xB5, L"Intel(R) Xeon(TM) processor MP"                       },
        { 0xB6, L"AMD Athlon(TM) XP Processor Family"                   },
        { 0xB7, L"AMD Athlon(TM) MP Processor Family"                   },
        { 0xB8, L"Intel(R) Itanium(R) 2 processor"                      },
        { 0xB9, L"Intel(R) Pentium(R) M processor"                      },
        { 0xBA, L"Intel(R) Celeron(R) D processor"                      },
        { 0xBB, L"Intel(R) Pentium(R) D processor"                      },
        { 0xBC, L"Intel(R) Pentium(R) Processor Extreme Edition"        },
        { 0xBD, L"Intel(R) Core(TM) Solo Processor"                     },
        { 0xBF, L"Intel(R) Core(TM) 2 Duo Processor"                    },
        { 0xC0, L"Intel(R) Core(TM) 2 Solo processor"                   },
        { 0xC1, L"Intel(R) Core(TM) 2 Extreme processor"                },
        { 0xC2, L"Intel(R) Core(TM) 2 Quad processor"                   },
        { 0xC3, L"Intel(R) Core(TM) 2 Extreme mobile processor"         },
        { 0xC4, L"Intel(R) Core(TM) 2 Duo mobile processor"             },
        { 0xC5, L"Intel(R) Core(TM) 2 Solo mobile processor"            },
        { 0xC6, L"Intel(R) Core(TM) i7 processor"                       },
        { 0xC7, L"Dual-Core Intel(R) Celeron(R) processor"              },
        { 0xC8, L"IBM390 Family"                                        },
        { 0xC9, L"G4"                                                   },
        { 0xCA, L"G5"                                                   },
        { 0xCB, L"ESA/390 G6"                                           },
        { 0xCC, L"z/Architecture base"                                  },
        { 0xCD, L"Intel(R) Core(TM) i5 processor"                       },
        { 0xCE, L"Intel(R) Core(TM) i3 processor"                       },
        { 0xD2, L"VIA C7(TM)-M Processor Family"                        },
        { 0xD3, L"VIA C7(TM)-D Processor Family"                        },
        { 0xD4, L"VIA C7(TM) Processor Family"                          },
        { 0xD5, L"VIA Eden(TM) Processor Family"                        },
        { 0xD6, L"Multi-Core Intel(R) Xeon(R) processor"                },
        { 0xD7, L"Dual-Core Intel(R) Xeon(R) processor 3xxx Series"     },
        { 0xD8, L"Quad-Core Intel(R) Xeon(R) processor 3xxx Series"     },
        { 0xD9, L"VIA Nano(TM) Processor Family"                        },
        { 0xDA, L"Dual-Core Intel(R) Xeon(R) processor 5xxx Series"     },
        { 0xDB, L"Quad-Core Intel(R) Xeon(R) processor 5xxx Series"     },
        { 0xDD, L"Dual-Core Intel(R) Xeon(R) processor 7xxx Series"     },
        { 0xDE, L"Quad-Core Intel(R) Xeon(R) processor 7xxx Series"     },
        { 0xDF, L"Multi-Core Intel(R) Xeon(R) processor 7xxx Series"    },
        { 0xE0, L"Multi-Core Intel(R) Xeon(R) processor 3400 Series"    },
        { 0xE4, L"AMD Opteron(TM) 3000 Series Processor"                },
        { 0xE5, L"AMD Sempron(TM) II Processor"                         },
        { 0xE6, L"Embedded AMD Opteron(TM) Quad-Core Processor Family"  },
        { 0xE7, L"AMD Phenom(TM) Triple-Core Processor Family"          },
        { 0xE8, L"AMD Turion(TM) Ultra Dual-Core Mobile Processor"      },
        { 0xE9, L"AMD Turion(TM) Dual-Core Mobile Processor Family"     },
        { 0xEA, L"AMD Athlon(TM) Dual-Core Processor Family"            },
        { 0xEB, L"AMD Sempron(TM) SI Processor Family"                  },
        { 0xEC, L"AMD Phenom(TM) II Processor Family"                   },
        { 0xED, L"AMD Athlon(TM) II Processor Family"                   },
        { 0xEE, L"Six-Core AMD Opteron(TM) Processor Family"            },
        { 0xEF, L"AMD Sempron(TM) M Processor Family"                   },
        { 0xFA, L"i860"                                                 },
        { 0xFB, L"i960"                                                 }};

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Families ); i++ )
    {
        if ( family == Families[ i ].family )
        {
            *pTranslation = Families[ i ].pszProcessorFamily;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( family );
        PXSLogAppWarn1( L"Unrecognised SMBIOS family type: %%1.", *pTranslation);
    }
}

//===============================================================================================//
//  Description:
//      Translate a processor family 2 value
//
//  Parameters:
//      family2      - the processor family value
//      pTranslation - receives the translation
//
//  Remarks:
//      System Slots (Type 9)
//      28h, 2.6+, Processor Family 2, WORD, Enum
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateProcessorFamily2( WORD family2, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _FAMILY2
    {
        WORD    family2;
        LPCWSTR pszProcessorFamily2;
    } Families2[] = { { 0x104, L"SH-3"            },
                      { 0x105, L"SH-4"            },
                      { 0x118, L"ARM"             },
                      { 0x119, L"StrongARM"       },
                      { 0x12C, L"6x86"            },
                      { 0x12D, L"MediaGX"         },
                      { 0x12E, L"MII"             },
                      { 0x140, L"WinChip"         },
                      { 0x15E, L"DSP"             },
                      { 0x1F4, L"Video Processor" }, };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Families2 ); i++ )
    {
        if ( family2 == Families2[ i ].family2 )
        {
            *pTranslation = Families2[ i ].pszProcessorFamily2;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt16( family2 );
        PXSLogAppWarn1( L"Unrecognised SMBIOS family 2: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a processor status value
//
//  Parameters:
//      status       - the processor status value
//      pTranslation - receives the translation
//
//  Remarks:
//      Processor Information (Type 4)
//      18h, 2.0+, Status, BYTE, Varies
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateProcessorStatus( BYTE status, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _STATUS
    {
        BYTE    status;
        LPCWSTR pszProcessorStatus;
    } aStatus[] = { { 0x00, L"Unknown"                                },
                    { 0x01, L"CPU Enabled"                            },
                    { 0x02, L"CPU Disabled by User through BIOS Setup"},
                    { 0x03, L"CPU Disabled By BIOS (POST Error)"      },
                    { 0x04, L"CPU is Idle, waiting to be enabled."    },
                    { 0x05, L""                                       },
                    { 0x06, L""                                       },
                    { 0x07, L"Other"                                  } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }

    // Test bit 6
    if ( status & 0x40 )
    {
        *pTranslation = L"CPU Socket Populated";
    }
    else
    {
        *pTranslation = L"CPU Socket Unpopulated";
    }

    for ( i = 0; i < ARRAYSIZE( aStatus ); i++ )
    {
        // Comparison is on bits 2:0
        if ( ( 0x7 & status ) == aStatus[ i ].status )
        {
            if ( aStatus[ i ].pszProcessorStatus[ 0 ] != PXS_CHAR_NULL )
            {
                *pTranslation += L", ";
            }
            *pTranslation += aStatus[ i ].pszProcessorStatus;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( status );
        PXSLogAppWarn1( L"Unrecognised SMBIOS processor status: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a processor type value
//
//  Parameters:
//      type         - processor type
//      pTranslation - receives the translation
//
//  Remarks:
//      Processor Information (Type 4)
//      05h, 2.0+, Processor Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateProcessorType( BYTE type, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    type;
        LPCWSTR pszProcessorType;
    } Types[] = { { 0x01, L"Other"             },
                  { 0x02, L"Unknown"           },
                  { 0x03, L"Central Processor" },
                  { 0x04, L"Math Processor"    },
                  { 0x05, L"DSP Processor"     },
                  { 0x06, L"Video Processor"   } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszProcessorType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS processor type: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a processor upgrade value
//
//  Parameters:
//      upgrade      - the upgrade value
//      pTranslation - receives the translation
//
//  Remarks:
//      Processor Information (Type 4)
//      19h, 2.0+, Processor Upgrade, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateProcessorUpgrade( BYTE upgrade, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    upgrade;
        LPCWSTR pszProcessorUpgrade;
    } Upgrades[] = { { 0x00, L""                       },
                     { 0x01, L"Other"                  },
                     { 0x02, L"Unknown"                },
                     { 0x03, L"Daughter Board"         },
                     { 0x04, L"ZIF Socket"             },
                     { 0x05, L"Replaceable Piggy Back" },
                     { 0x06, L"None"                   },
                     { 0x07, L"LIF Socket"             },
                     { 0x08, L"Slot 1"                 },
                     { 0x09, L"Slot 2"                 },
                     { 0x0A, L"370-pin socket"         },
                     { 0x0B, L"Slot A"                 },
                     { 0x0C, L"Slot M"                 },
                     { 0x0D, L"Socket 423"             },
                     { 0x0E, L"Socket A (Socket 462)"  },
                     { 0x0F, L"Socket 478"             },
                     { 0x10, L"Socket 754"             },
                     { 0x11, L"Socket 940"             },
                     { 0x12, L"Socket 939"             },
                     { 0x13, L"Socket mPGA604"         },
                     { 0x14, L"Socket LGA771"          },
                     { 0x15, L"Socket LGA775"          },
                     { 0x16, L"Socket S1"              },
                     { 0x17, L"Socket AM2"             },
                     { 0x18, L"Socket F (1207)"        },
                     { 0x19, L"Socket LGA1366"         },
                     { 0x1A, L"Socket G34"             },
                     { 0x1B, L"Socket AM3"             },
                     { 0x1C, L"Socket C32"             },
                     { 0x1D, L"Socket LGA1156"         },
                     { 0x1E, L"Socket LGA1567"         },
                     { 0x1F, L"Socket PGA988A"         },
                     { 0x20, L"Socket BGA1288"         },
                     { 0x21, L"Socket rPGA988B"        },
                     { 0x22, L"Socket BGA1023"         },
                     { 0x23, L"Socket BGA1224"         },
                     { 0x24, L"Socket LGA1155"         },
                     { 0x25, L"Socket LGA1356"         },
                     { 0x26, L"Socket LGA2011"         },
                     { 0x27, L"Socket FS1"             },
                     { 0x28, L"Socket FS2"             },
                     { 0x29, L"Socket FM1"             },
                     { 0x2A, L"Socket FM2"             },
                     { 0x2B, L"Socket LGA2011-3"       },
                     { 0x2C, L"Socket LGA1356-3"       } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Upgrades ); i++ )
    {
        if ( upgrade == Upgrades[ i ].upgrade )
        {
            *pTranslation = Upgrades[ i ].pszProcessorUpgrade;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( upgrade );
        PXSLogAppWarn1( L"Unrecognised SMBIOS processor upgrade: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a processor voltage value
//
//  Parameters:
//      voltage      - the processor voltage value
//      pTranslation - receives the translation
//
//  Remarks:
//      Processor Information (Type 4)
//      11h, 2.0+, Voltage, BYTE, Varies
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateProcessorVoltage( BYTE voltage, String* pTranslation ) const
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    // Test for legacy mode, if so bit 7 is zero
    if ( voltage < 128 )
    {
        if ( 0x01 & voltage ) *pTranslation += L"5.0V, ";
        if ( 0x02 & voltage ) *pTranslation += L"3.3V, ";
        if ( 0x04 & voltage ) *pTranslation += L"2.9V";

        pTranslation->Trim();
        if ( pTranslation->EndsWithCharacter( PXS_CHAR_COMMA ) )
        {
            pTranslation->Truncate( pTranslation->GetLength() - 1 );
        }
    }
    else
    {
        *pTranslation  = Format.Double( (voltage - 128) / 10.0 );
        *pTranslation += L"V";
    }
}

//===============================================================================================//
//  Description:
//      Translate slot characteristics 1 into a string
//
//  Parameters:
//      characteristics1 - the slot characteristics
//      pTranslation     - receives the translation
//
//  Remarks:
//      System Slots (Type 9)
//      0Bh, 2.0+, Slot Characteristics 1, BYTE, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateSystemSlotCharac1( BYTE characteristics1,
                                                    String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Values[] = { L"Unknown",
                         L"5.0 Volts",
                         L"3.3 Volts",
                         L"Shared Slot",
                         L"Supports PC Card-162",
                         L"Supports CardBus",
                         L"Supports Zoom Video",
                         L"Supports Modem Ring Resume" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Values ); i++ )
    {
        if ( characteristics1 & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Values[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate slot characteristics 2 into a string
//
//  Parameters:
//      characteristics2 - the slot characteristics
//      pTranslation     - receives the translation
//
//  Remarks:
//      System Slots (Type 9)
//      0Bh, 2.1+, Slot Characteristics 2, BYTE, Bit Field
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateSystemSlotCharac2( BYTE characteristics1,
                                                    String* pTranslation ) const
{
    size_t  i = 0;
    LPCWSTR Values[] = { L"Supports PME Signal",
                         L"Supports Hot Plug Devices",
                         L"Supports SMBus Signal" };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Values ); i++ )
    {
        if ( characteristics1 & ( 1 << i ) )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Values[ i ];
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a system slot length
//
//  Parameters:
//      length        - the length
//      pTranslation - receives the translation
//
//  Remarks:
//      System Slots (Type 9)
//      08h, 2.0+, Slot Length, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateSystemSlotLength( BYTE length, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _LENGTH
    {
        BYTE    length;
        LPCWSTR pszSlotLength;
    } Lengths[] = { { 0x01, L"Other"        },
                    { 0x02, L"Unknown"      },
                    { 0x03, L"Short Length" },
                    { 0x04, L"Long Length"  } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Lengths ); i++ )
    {
        if ( length == Lengths[ i ].length )
        {
            *pTranslation = Lengths[ i ].pszSlotLength;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( length );
        PXSLogAppWarn1( L"Unrecognised SMBIOS system slot length: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a system slot type
//
//  Parameters:
//      type         - the slot type
//      pTranslation - receives the translation
//
//  Remarks:
//      System Slots (Type 9)
//      05h, 2.0+, Slot Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateSystemSlotType( BYTE type, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _LENGTH
    {
        BYTE    type;
        LPCWSTR pszSlotType;
    } Types[] = { { 0x01, L"Other"                       },
                  { 0x02, L"Unknown"                     },
                  { 0x03, L"ISA"                         },
                  { 0x04, L"MCA"                         },
                  { 0x05, L"EISA"                        },
                  { 0x06, L"PCI"                         },
                  { 0x07, L"PC Card (PCMCIA)"            },
                  { 0x08, L"VL-VESA"                     },
                  { 0x09, L"Proprietary"                 },
                  { 0x0A, L"Processor Card Slot"         },
                  { 0x0B, L"Proprietary Memory Card Slot"},
                  { 0x0C, L"I/O Riser Card Slot"         },
                  { 0x0D, L"NuBus"                       },
                  { 0x0E, L"PCI - 66MHz Capable"         },
                  { 0x0F, L"AGP"                         },
                  { 0x10, L"AGP 2X"                      },
                  { 0x11, L"AGP 4X"                      },
                  { 0x12, L"PCI-X"                       },
                  { 0x13, L"AGP 8X"                      },
                  { 0xA0, L"PC-98/C20"                   },
                  { 0xA1, L"PC-98/C24"                   },
                  { 0xA2, L"PC-98/E"                     },
                  { 0xA3, L"PC-98/Local Bus"             },
                  { 0xA4, L"PC-98/Card"                  },
                  { 0xA5, L"PCI Express"                 },
                  { 0xA6, L"PCI Express x1"              },
                  { 0xA7, L"PCI Express x2"              },
                  { 0xA8, L"PCI Express x4"              },
                  { 0xA9, L"PCI Express x8"              },
                  { 0xAA, L"PCI Express x16"             },
                  { 0xAB, L"PCI Express Gen 2"           },
                  { 0xAC, L"PCI Express Gen 2 x1"        },
                  { 0xAD, L"PCI Express Gen 2 x2"        },
                  { 0xAE, L"PCI Express Gen 2 x4"        },
                  { 0xAF, L"PCI Express Gen 2 x8"        },
                  { 0xB0, L"PCI Express Gen 2 x16"       },
                  { 0xB1, L"PCI Express Gen 3"           },
                  { 0xB2, L"PCI Express Gen 3 x1"        },
                  { 0xB3, L"PCI Express Gen 3 x2"        },
                  { 0xB4, L"PCI Express Gen 3 x4"        },
                  { 0xB5, L"PCI Express Gen 3 x8"        },
                  { 0xB6, L"PCI Express Gen 3 x16"       } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszSlotType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS system slot type: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a system slot current usage value
//
//  Parameters:
//      Usage        - the slot usage value
//      pTranslation - receives the translation
//
//  Remarks:
//      System Slots (Type 9)
//      2.0+, Current Usage, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateSystemSlotUsage( BYTE Usage, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _USAGE
    {
        BYTE    Usage;
        LPCWSTR pszSlotUsage;
    } Usages[] = { { 0x01, L"Other"     },
                   { 0x02, L"Unknown"   },
                   { 0x03, L"Available" },
                   { 0x04, L"In use"  } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Usages ); i++ )
    {
        if ( Usage == Usages[ i ].Usage )
        {
            *pTranslation = Usages[ i ].pszSlotUsage;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( Usage );
        PXSLogAppWarn1( L"Unrecognised SMBIOS system current usage: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a system slot current bus width value
//
//  Parameters:
//      width        - the slot bus width value
//      pTranslation - receives the translation
//
//  Remarks:
//      System Slots (Type 9)
//      06h, 2.0+, Slot Data Bus Width, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateSystemSlotWidth( BYTE width, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _WIDTH
    {
        BYTE    width;
        LPCWSTR pszSlotWidth;
    } Widths[] = { { 0x01, L"Other"      },
                   { 0x02, L"Unknown"    },
                   { 0x03, L"8 bit"      },
                   { 0x04, L"16 bit"     },
                   { 0x05, L"32 bit"     },
                   { 0x06, L"64 bit"     },
                   { 0x07, L"128 bit"    },
                   { 0x08, L"1x or x1"   },
                   { 0x09, L"2x or x2"   },
                   { 0x0A, L"4x or x4"   },
                   { 0x0B, L"8x or x8"   },
                   { 0x0C, L"12x or x12" },
                   { 0x0D, L"16x or x16" },
                   { 0x0E, L"32x or x32" } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Widths ); i++ )
    {
        if ( width == Widths[ i ].width )
        {
            *pTranslation = Widths[ i ].pszSlotWidth;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( width );
        PXSLogAppWarn1( L"Unrecognised SMBIOS system slot bus width: %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a wake-up type value
//
//  Parameters:
//      type         - the wake-up type
//      pTranslation - receives the translation
//
//  Remarks:
//      System Information (Type 1)
//      18h, 2.1+, Wake-up Type, BYTE, ENUM
//
//  Returns:
//      void
//===============================================================================================//
void SmbiosInformation::TranslateWakeUpType( BYTE type, String* pTranslation ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _TYPE
    {
        BYTE    type;
        LPCWSTR pszWakeUpType;
    } Types[] = { { 0x00, L""                  },
                  { 0x01, L"Other"             },
                  { 0x02, L"Unknown"           },
                  { 0x03, L"APM Timer"         },
                  { 0x04, L"Modem Ring"        },
                  { 0x05, L"LAN Remote"        },
                  { 0x06, L"Power Switch"      },
                  { 0x07, L"PCI PME#"          },
                  { 0x08, L"AC Power resorted" } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pTranslation = Types[ i ].pszWakeUpType;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt8( type );
        PXSLogAppWarn1( L"Unrecognised SMBIOS wake-up type: %%1.", *pTranslation );
    }
}
