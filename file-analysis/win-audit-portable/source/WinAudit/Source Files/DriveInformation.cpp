///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Logical Drive Information Class Implementation
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
#include "WinAudit/Header Files/DriveInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"
#include "PxsBase/Header Files/Wmi.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
DriveInformation::DriveInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
DriveInformation::~DriveInformation()
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
//      Get the total fixed local hard drive capacity
//
//  Parameters:
//      none
//
//  Remarks:
//      Will use Win32_DiskDrive rather than enumerate local hdd volumes in the
//      event "subst" has been used to define a virtual volume
//
//  Returns:
//      64-bit unsigned signed integer
//===============================================================================================//
UINT64 DriveInformation::GetTotalHDDCapacity()
{
    UINT    errorMode;
    UINT64  totalBytes = 0, size = 0;
    Wmi     WMI;
    String  MediaType;

    // Suppress errors
    errorMode = SetErrorMode( SEM_FAILCRITICALERRORS );
    try
    {
        WMI.Connect( L"root\\cimv2" );
        WMI.ExecQuery( L"Select * from Win32_DiskDrive" );
        while ( WMI.Next() )
        {
            MediaType = PXS_STRING_EMPTY;
            WMI.Get( L"MediaType", &MediaType );
            if ( MediaType.IndexOfI( L"hard" ) != PXS_MINUS_ONE )
            {
                size = 0;
                WMI.GetUInt64( L"Size", &size );
                totalBytes += size;
            }
        }
        WMI.Disconnect();
    }
    catch ( const Exception& )
    {
        SetErrorMode( errorMode );
        throw;
    }
    SetErrorMode( errorMode );

    return totalBytes;
}

//===============================================================================================//
//  Description:
//      Get the audit records about the volumes/drives
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//      void
//===============================================================================================//
void DriveInformation::GetAuditRecords( TArray< AuditRecord >* pRecords )
{
    DWORD   logicalDrives;
    wchar_t driveLetter;
    AuditRecord Record;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    logicalDrives = GetLogicalDrives();
    if ( logicalDrives == 0 )
    {
        throw SystemException( GetLastError(),
                               L"GetLogicalDrives", __FUNCTION__ );
    }

    driveLetter = 'A';
    while ( logicalDrives )
    {
        if ( logicalDrives & 0x01 )
        {
            GetVolumneRecord( driveLetter, &Record );
            pRecords->Add( Record );
        }
        driveLetter++;
        logicalDrives = logicalDrives >> 1;
    }
}

//===============================================================================================//
//  Description:
//      Translate a defined constant representing the drive type to a string
//
//  Parameters:
//      driveType    - the drive type
//      pTranslation - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void DriveInformation::TranslateDriveType( UINT driveType, String* pTranslation ) const
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }

    switch ( driveType )
    {
        default:
            *pTranslation = Format.UInt32( driveType );
            PXSLogAppWarn1( L"Unrecognised drive type: %%1.", *pTranslation );
            break;

            case DRIVE_UNKNOWN:
            *pTranslation = L"Unnown";
            break;

        case DRIVE_REMOVABLE:
            *pTranslation = L"Removable Drive";
            break;

        case DRIVE_FIXED:
            *pTranslation = L"Fixed Drive";
            break;

        case DRIVE_REMOTE:
            *pTranslation = L"Network Drive";
            break;

        case DRIVE_CDROM:
            *pTranslation = L"CD-ROM";
            break;

        case DRIVE_RAMDISK:
            *pTranslation = L"RAM Disk";
            break;
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
//      Get volume information as an audit record
//
//  Parameters:
//      driveLetter - character representing the drive, from A to Z
//      pRecord     - receives the audit record
//
//  Returns:
//      true if volume exists, otherwise false
//===============================================================================================//
bool DriveInformation::GetVolumneRecord( wchar_t driveLetter, AuditRecord* pRecord )
{
    UINT    errorMode = 0, driveType = 0;
    DWORD   BytesPerSector = 0, NumberOfFreeClusters = 0, lastError = 0;
    DWORD   nVolumeNameSize  = 0, SectorsPerCluster  = 0;
    DWORD   TotalNumberOfClusters  = 0;
    UINT64  percentUsed = 0;
    wchar_t szSerialNumber[ 32 ] = { 0 };             // E.g. 1234-5678
    wchar_t szVolumeNameBuffer[ MAX_PATH + 1 ] = { 0 };
    wchar_t szFileSystemNameBuffer[ MAX_PATH + 1 ] = { 0 };
    String  RootPathName, Value;
    Formatter Format;
    ULARGE_INTEGER   TotalNumberOfBytes, TotalNumberOfFreeBytes;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_DRIVES );

    errorMode = SetErrorMode( SEM_FAILCRITICALERRORS );     // Suppress errors
    RootPathName  = driveLetter;
    RootPathName += PXS_CHAR_COLON;
    RootPathName += PXS_PATH_SEPARATOR;
    driveType = GetDriveType( RootPathName.c_str() );
    if ( ( driveType == DRIVE_UNKNOWN ) || ( driveType == DRIVE_NO_ROOT_DIR ) )
    {
        SetErrorMode( errorMode );
        return false;
    }

    // Volume
    if ( GetVolumeInformation( RootPathName.c_str(),
                               szVolumeNameBuffer,
                               ARRAYSIZE( szVolumeNameBuffer ),
                               &nVolumeNameSize ,
                               nullptr,
                               nullptr,
                               szFileSystemNameBuffer, ARRAYSIZE( szFileSystemNameBuffer ) ) == 0 )
    {
        // Usually drive not ready so reset, log and return
        SetErrorMode( errorMode );
        lastError = GetLastError();
        if ( lastError == ERROR_NOT_READY )
        {
            PXSLogAppInfo1( L"Drive '%%1' is not ready.", RootPathName );
        }
        else
        {
            PXSLogSysError1( lastError , L"GetVolumeInformation for drive '%%1'.", RootPathName );
        }
        Value = driveLetter;
        pRecord->Add( PXS_DRIVES_LETTER, Value );
        Value = PXS_STRING_EMPTY;
        TranslateDriveType( driveType, &Value );
        pRecord->Add( PXS_DRIVES_TYPE, Value );
        return true;        // Nothing more to do
    }
    szVolumeNameBuffer[ ARRAYSIZE( szVolumeNameBuffer ) - 1 ] = PXS_CHAR_NULL;
    szFileSystemNameBuffer[ ARRAYSIZE( szFileSystemNameBuffer ) - 1 ] = 0x00;

    // Space
    memset( &TotalNumberOfBytes, 0, sizeof ( TotalNumberOfBytes ) );
    memset( &TotalNumberOfFreeBytes  , 0, sizeof ( TotalNumberOfFreeBytes ) );
    if ( GetDiskFreeSpaceEx( RootPathName.c_str(),
                             nullptr,
                             &TotalNumberOfBytes, &TotalNumberOfFreeBytes ) == 0 )
    {
        SetErrorMode( errorMode );
        throw SystemException( GetLastError(), RootPathName.c_str(), "GetDiskFreeSpaceEx" );
    }

    // Geometry
    if ( GetDiskFreeSpace( RootPathName.c_str(),
                           &SectorsPerCluster,
                           &BytesPerSector,
                           &NumberOfFreeClusters, &TotalNumberOfClusters ) == 0 )
    {
        SetErrorMode( errorMode );
        throw SystemException( GetLastError(), RootPathName.c_str(), "GetDiskFreeSpace" );
    }

    // Catch errors to reset the error mode
    try
    {
        Value = driveLetter;
        pRecord->Add( PXS_DRIVES_LETTER, Value );

        Value = PXS_STRING_EMPTY;
        TranslateDriveType( driveType, &Value );
        pRecord->Add( PXS_DRIVES_TYPE, Value );

        Value = PXS_STRING_EMPTY;
        if ( TotalNumberOfBytes.QuadPart > 0 )
        {
            percentUsed = 100 - ( ( 100 * TotalNumberOfFreeBytes.QuadPart ) /
                                                  TotalNumberOfBytes.QuadPart );
            Value  = Format.UInt64( percentUsed );
            Value += L"%";
        }
        pRecord->Add( PXS_DRIVES_USED_SPACE_PERCENT, Value );

        Value = Format.StorageBytes( TotalNumberOfBytes.QuadPart -
                                     TotalNumberOfFreeBytes.QuadPart );
        pRecord->Add( PXS_DRIVES_USED_SPACE, Value );

        Value = Format.StorageBytes( TotalNumberOfFreeBytes.QuadPart );
        pRecord->Add( PXS_DRIVES_FREE_SPACE, Value );

        // Total space
        Value = Format.StorageBytes( TotalNumberOfBytes.QuadPart );
        pRecord->Add( PXS_DRIVES_TOTAL_SPACE, Value );

        pRecord->Add( PXS_DRIVES_VOLUME_NAME, szVolumeNameBuffer );
        pRecord->Add( PXS_DRIVES_FILE_SYSTEM, szFileSystemNameBuffer );

        memset( szSerialNumber, 0, sizeof ( szSerialNumber ) );
        StringCchPrintf( szSerialNumber,
                         ARRAYSIZE( szSerialNumber ),
                         L"%04lX-%04lX",
                         ( nVolumeNameSize >> 16 ),
                         ( nVolumeNameSize & 0xFFFF ) );
        pRecord->Add( PXS_DRIVES_VOLUME_SERIAL_NUMBER, szSerialNumber );

        Value = Format.UInt32( SectorsPerCluster );
        pRecord->Add( PXS_DRIVES_SECTORS_PER_CLUSTER, Value );

        Value = Format.UInt32( BytesPerSector );
        pRecord->Add( PXS_DRIVES_BYTES_PER_SECTOR, Value );

        Value = Format.UInt32( NumberOfFreeClusters );
        pRecord->Add( PXS_DRIVES_FREE_CLUSTERS, Value );

        Value = Format.UInt32( TotalNumberOfClusters  );
        pRecord->Add( PXS_DRIVES_TOTAL_CLUSTERS, Value );
    }
    catch ( const Exception& )
    {
        SetErrorMode( errorMode );
        throw;
    }
    SetErrorMode( errorMode );

    return true;
}
