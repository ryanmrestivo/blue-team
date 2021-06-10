///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Windows Network Information Class Implementation
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
#include "WinAudit/Header Files/WindowsNetworkInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoNetApiBufferFree.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
WindowsNetworkInformation::WindowsNetworkInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
WindowsNetworkInformation::~WindowsNetworkInformation()
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
//      Get the open files on the computer as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the records
//
//  Remarks:
//      Only members of the Administrators or Account Operators
//      local group can successfully execute NetFileEnum.
//
//  Returns:
//      void
//===============================================================================================//
void WindowsNetworkInformation::GetOpenFileRecords( TArray< AuditRecord >* pRecords )
{
    DWORD       i = 0, totalEntries = 0, entriesRead = 0;
    String      Permissions, Value;
    Formatter   Format;
    AuditRecord Record;
    FILE_INFO_3*   pFileInfo3 = nullptr;
    NET_API_STATUS status     = 0;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get all the data in one go at level 3
    status = NetFileEnum( nullptr,                     // Local computer
                          nullptr,
                          nullptr,
                          3,
                          reinterpret_cast<LPBYTE*>( &pFileInfo3 ),
                          MAX_PREFERRED_LENGTH,        // Get all the data
                          &entriesRead, &totalEntries, nullptr );
    AutoNetApiBufferFree AutoFreeFileInfo3( pFileInfo3 );

    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK,
                         status, L"NetFileEnum", __FUNCTION__ );
    }

    // Is there any data?
    if ( pFileInfo3 == nullptr )
    {
        return;     // Nothing to do
    }

    for ( i = 0; i < entriesRead; i++ )
    {
        if ( pFileInfo3[ i ].fi3_pathname )
        {
            // Make the audit record
            Record.Reset( PXS_CATEGORY_WIN_NET_FILES );
            Record.Add( PXS_WIN_NET_FILES_PATHNAME, pFileInfo3[i].fi3_pathname);
            Record.Add( PXS_WIN_NET_FILES_USERNAME, pFileInfo3[i].fi3_username);

            Value = Format.UInt32( pFileInfo3[ i ].fi3_num_locks );
            Record.Add( PXS_WIN_NET_FILES_LOCKS, Value );

            Permissions = PXS_STRING_EMPTY;
            TranslateOpenFilePermissions( pFileInfo3[ i ].fi3_permissions, &Permissions );
            Record.Add( PXS_WIN_NET_FILES_PERMISSIONS, Permissions );
            pRecords->Add( Record );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_WIN_NET_FILES_PATHNAME );
}

//===============================================================================================//
//  Description:
//      Get the sessions, i.e. connected users as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the records
//
//  Returns:
//      void
//===============================================================================================//
void WindowsNetworkInformation::GetSessionRecords( TArray< AuditRecord >* pRecords )
{
    DWORD       i = 0, entriesRead = 0, totalEntries = 0;
    String      Value, LocaleMinutes;
    Formatter   Format;
    AuditRecord Record;
    NET_API_STATUS   status = 0;
    SESSION_INFO_10* pSessionInfo10 = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get all the data in one go, no special permissions at level 10
    status = NetSessionEnum( nullptr,        // Local computer
                             nullptr,        // All sessions
                             nullptr,        // All users
                             10,
                             reinterpret_cast<LPBYTE*>( &pSessionInfo10 ),
                             MAX_PREFERRED_LENGTH,
                             &entriesRead, &totalEntries, nullptr );
    AutoNetApiBufferFree AutoFreeSessionInfo10( pSessionInfo10 );

    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetSessionEnum", __FUNCTION__ );
    }

    // Is there any data?
    if ( pSessionInfo10 == nullptr )
    {
        return;     // Nothing to do
    }

    // Locale string
    PXSGetResourceString( PXS_IDS_140_MINUTES, &LocaleMinutes );

    for ( i = 0; i < entriesRead; i++ )
    {
        // Make the audit record
        Record.Reset( PXS_CATEGORY_WIN_NET_SESSIONS );

        Value = pSessionInfo10[ i ].sesi10_cname;
        Record.Add( PXS_WIN_NET_SESS_COMPUTERNAME, Value );

        Value = pSessionInfo10[ i ].sesi10_username;
        Record.Add( PXS_WIN_NET_SESS_USERNAME, Value );

        Value  = Format.UInt32( pSessionInfo10[ i ].sesi10_time / 60 );
        Value += PXS_CHAR_SPACE;
        Value += LocaleMinutes;
        Record.Add( PXS_WIN_NET_SESS_CONN_TIME_MIN, Value );

        Value  = Format.UInt32( pSessionInfo10[ i ].sesi10_idle_time / 60 );
        Value += PXS_CHAR_SPACE;
        Value += LocaleMinutes;
        Record.Add( PXS_WIN_NET_SESS_IDLE_TIME_MIN, Value );

        pRecords->Add( Record );
    }
    PXSSortAuditRecords( pRecords, PXS_WIN_NET_SESS_COMPUTERNAME );
}

//===============================================================================================//
//  Description:
//      Get the network shares as a string array
//
//  Parameters:
//      pShareNames - string array to receive the names
//
//  Remarks:
//      Admin or comm, print, or server operator group membership
//      is required to successfully execute NetShareEnum at level 2.
//      No special group membership is required for level 0 or
//      level 1 calls.
//
//  Returns:
//      void
//===============================================================================================//
void WindowsNetworkInformation::GetShareNames( StringArray* pShareNames )
{
    DWORD     i = 0, entriesRead = 0, totalEntries = 0;
    String    ShareName;
    Formatter Format;
    SHARE_INFO_1*   pShareInfo1 = nullptr;
    NET_API_STATUS  status      = 0;

    if ( pShareNames == nullptr )
    {
        throw ParameterException( L"pShareNames", __FUNCTION__ );
    }
    pShareNames->RemoveAll();

    // Get all the data in one go at level 1
    status = NetShareEnum( nullptr,                // Local computer
                           1,
                           reinterpret_cast<LPBYTE*>( &pShareInfo1 ),
                           MAX_PREFERRED_LENGTH,
                           &entriesRead, &totalEntries, nullptr );
    AutoNetApiBufferFree AutoFreeShareInfo1( pShareInfo1 );

    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK,
                         status, L"NetShareEnum", __FUNCTION__ );
    }

    // Is there any data?
    if ( pShareInfo1 == nullptr )
    {
        return;     // Nothing to do
    }

    for ( i = 0; i < entriesRead; i++ )
    {
        ShareName = pShareInfo1[ i ].shi1_netname;
        ShareName.Trim();
        if ( ShareName.GetLength() )
        {
            pShareNames->Add( ShareName );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the network shares as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the records
//
//  Remarks:
//      Admin or comm, print, or server operator group membership
//      is required to successfully execute NetShareEnum at level 2.
//      No special group membership is required for level 0 or
//      level 1 calls.
//
//  Returns:
//      void
//===============================================================================================//
void WindowsNetworkInformation::GetShareRecords( TArray<AuditRecord>* pRecords)
{
    DWORD       i = 0, entriesRead = 0, totalEntries = 0;
    String      ShareName, Connections, ShareType, SharePath, Insert1, Insert2;
    Formatter   Format;
    AuditRecord Record;
    SHARE_INFO_1*  pShareInfo1 = nullptr;
    SHARE_INFO_2*  pShareInfo2 = nullptr;
    NET_API_STATUS status      = 0;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get all the data in one go at level 1
    status = NetShareEnum( nullptr,                // Local computer
                           1,
                           reinterpret_cast<LPBYTE*>( &pShareInfo1 ),
                           MAX_PREFERRED_LENGTH,
                           &entriesRead, &totalEntries, nullptr );
    AutoNetApiBufferFree AutoFreeShareInfo1( pShareInfo1 );

    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetShareEnum", __FUNCTION__ );
    }
    Insert1 = Format.UInt32( entriesRead );
    Insert2 = Format.UInt32( totalEntries );
    PXSLogAppInfo2( L"NetShareEnum result: entriesRead='%%1', totalEntries='%%2'.",
                    Insert1, Insert2 );

    // Is there any data?
    if ( pShareInfo1 == nullptr )
    {
        PXSLogAppInfo( L"NetShareEnum found no network shares (pShareInfo1 = NULL)." );
        return;     // Nothing to do
    }

    for ( i = 0; i < entriesRead; i++ )
    {
        ShareName = pShareInfo1[ i ].shi1_netname;
        ShareName.Trim();
        if ( ShareName.GetLength() )
        {
            // Make the audit record
            Record.Reset( PXS_CATEGORY_WIN_NET_SHARES );
            Record.Add( PXS_WIN_NET_SHARES_NETNAME, ShareName );

            ShareType = PXS_STRING_EMPTY;
            TranslateShareType( pShareInfo1[ i ].shi1_type, &ShareType );
            Record.Add( PXS_WIN_NET_SHARES_TYPE, ShareType );

            // Try to get level 2 information, optional information so continue
            // on error
            Connections = PXS_STRING_EMPTY;
            SharePath   = PXS_STRING_EMPTY;
            pShareInfo2 = nullptr;
            status = NetShareGetInfo( nullptr,
                                      pShareInfo1[ i ].shi1_netname,
                                      2,
                                      reinterpret_cast<LPBYTE*>(&pShareInfo2) );
            if ( status == NERR_Success )
            {
                AutoNetApiBufferFree AutoFreeShareInfo2( pShareInfo2 );

                Connections = Format.UInt32( pShareInfo2->shi2_current_uses );
                SharePath   = pShareInfo2->shi2_path;
            }
            else
            {
                Insert1 = Format.UInt32( status );
                PXSLogAppError2(
                   L"NetShareGetInfo error: share='%%1', NET_API_STATUS='%%2'.",
                   ShareName, Insert1 );
            }
            Record.Add( PXS_WIN_NET_SHARES_CONNECTIONS, Connections );
            Record.Add( PXS_WIN_NET_SHARES_PATH, SharePath );
            pRecords->Add( Record );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_WIN_NET_SHARES_NETNAME );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Translate the specified permissions to a string
//
//  Parameters:
//      permissions  - the permissions
//      pTranslation - string object to receive the translation
//
//  Remarks:
//      Other permissions exist in dwPermissions, but named ones are:
//      PERM_FILE_READ, PERM_FILE_WRITE and PERM_FILE_CREATE
//
//      MSDN
//      from lmaccess.h:
//          ACCESS_EXEC       = 0x08  Execute Permission (X)
//          ACCESS_DELETE     = 0x10  Delete Permission (D)
//          ACCESS_ATRIB      = 0x20  Change Attribute Permission (A)
//          ACCESS_PERM       = 0x40  Change ACL Permission (P)
//
//  Returns:
//      void
//===============================================================================================//
void WindowsNetworkInformation::TranslateOpenFilePermissions( DWORD permissions,
                                                              String* pTranslation )
{
    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    if ( PERM_FILE_READ & permissions )
    {
        *pTranslation += L"Read, Execute";
    }

    if ( PERM_FILE_WRITE & permissions )
    {
        if ( pTranslation->GetLength() )
        {
            *pTranslation += L", ";
        }
        *pTranslation += L"Write";
    }

    if ( PERM_FILE_CREATE & permissions )
    {
        if ( pTranslation->GetLength() )
        {
            *pTranslation += L", ";
        }
        *pTranslation += L"Create";
    }
}

//===============================================================================================//
//  Description:
//      Translate a share type into a description
//
//  Parameters:
//      type         - defined constant of share type
//      pTranslation - string object to receive the translation
//
//  Remarks:
//      See SHARE_INFO_2 Structure
//
//  Returns:
//      void
//===============================================================================================//
void WindowsNetworkInformation::TranslateShareType( DWORD type,
                                                    String* pTranslation )
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }

    // See SHARE_INFO_2 Structure
    switch ( type & 0xff )
    {
        default:
            *pTranslation = Format.UInt32( type );
            break;

        case STYPE_DISKTREE:
            *pTranslation = L"Disk Drive";
            break;

        case STYPE_PRINTQ:
            *pTranslation = L"Print Queue";
            break;

        case STYPE_DEVICE:
            *pTranslation = L"Communication Device";
            break;

        case STYPE_IPC:
             *pTranslation = L"Interprocess Communication";
             break;
    }

    // Bitmask for temporary shares
    if ( type & STYPE_TEMPORARY )
    {
        *pTranslation += L" Temporary";
    }

    // Bitmask for special shares, say Admin to be consistent with Win32_Share
    if ( type & STYPE_SPECIAL )
    {
        *pTranslation += L" Admin";
    }
}
