///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Group and User Information Class Implementation
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
#include "WinAudit/Header Files/GroupUserInformation.h"

// 2. C System Files
#include <NTSecAPI.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoNetApiBufferFree.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/Ddk.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
GroupUserInformation::GroupUserInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
GroupUserInformation::~GroupUserInformation()
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
//      Get information about the group accounts
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetGroupRecords( TArray< AuditRecord >* pRecords )
{
    size_t      i = 0, numGroups = 0;
    String      GroupName, Comment;
    AuditRecord Record, MemberRecord;
    StringArray GroupNames;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Local Groups
    GetLocalGroupNames( &GroupNames );
    numGroups = GroupNames.GetSize();
    for ( i = 0; i < numGroups; i++ )
    {
        GroupName = GroupNames.Get( i );
        Comment   = PXS_STRING_EMPTY;
        GetLocalGroupComment( GroupName, &Comment );
        Record.Reset( PXS_CATEGORY_GROUPS );
        Record.Add( PXS_GROUPS_TYPE, L"Local" );
        Record.Add( PXS_GROUPS_NAME , GroupName );
        Record.Add( PXS_GROUPS_COMMENT, Comment );
        pRecords->Add( Record );
    }

    // Global Groups
    GroupNames.RemoveAll();
    GetGlobalGroupNames( &GroupNames );
    numGroups = GroupNames.GetSize();
    for ( i = 0; i < numGroups; i++ )
    {
        GroupName = GroupNames.Get( i );
        Comment   = PXS_STRING_EMPTY;
        GetGlobalGroupComment( GroupName, &Comment );
        Record.Reset( PXS_CATEGORY_GROUPS );
        Record.Add( PXS_GROUPS_TYPE, L"Global" );
        Record.Add( PXS_GROUPS_NAME , GroupName );
        Record.Add( PXS_GROUPS_COMMENT, Comment );
        pRecords->Add( Record );
    }
    PXSSortAuditRecords( pRecords, PXS_GROUPS_NAME );
}

//===============================================================================================//
//  Description:
//      Get the members of all the groups
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetGroupMemberRecords(
                                              TArray< AuditRecord >* pRecords )
{
    size_t      i = 0, j = 0, numGroups = 0, numMembers = 0;
    String      GroupName, GroupType, MemberName;
    StringArray Members;
    AuditRecord Record, MemberRecord;
    TArray< AuditRecord > GroupRecords;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetGroupRecords( &GroupRecords );
    numGroups = GroupRecords.GetSize();
    for ( i = 0; i < numGroups; i++ )
    {
        Record    = GroupRecords.Get( i );
        GroupName = PXS_STRING_EMPTY;
        GroupType = PXS_STRING_EMPTY;
        Record.GetItemValue( PXS_GROUPS_NAME, &GroupName );
        Record.GetItemValue( PXS_GROUPS_TYPE, &GroupType );
        Members.RemoveAll();
        if ( GroupType.CompareI( L"Local" ) == 0 )
        {
            GetLocalGroupMembers3( GroupName, &Members );
        }
        else
        {
            GetGlobalGroupMembers( GroupName, &Members );
        }
        numMembers = Members.GetSize();
        for ( j = 0; j < numMembers; j++ )
        {
            MemberName = Members.Get( j );
            MemberRecord.Reset( PXS_CATEGORY_GROUPMEMBERS );
            MemberRecord.Add( PXS_GROUPMEMBERS_GROUP_NAME , GroupName );
            MemberRecord.Add( PXS_GROUPMEMBERS_MEMBER_NAME, MemberName );
            pRecords->Add( MemberRecord );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_GROUPMEMBERS_GROUP_NAME );
}

//===============================================================================================//
//  Description:
//      Get the privileges of all the groups
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetGroupPolicyRecords( TArray< AuditRecord >* pRecords )
{
    size_t      i = 0, j = 0, numGroups = 0, numPrivileges = 0;
    String      GroupName, PrivilegeName;
    StringArray GroupNames, GlobalGroupNames, Privileges;
    AuditRecord Record;
    TArray< AuditRecord > GroupRecords;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetLocalGroupNames( &GroupNames );
    GetGlobalGroupNames( &GlobalGroupNames );
    GroupNames.AddArray( GlobalGroupNames );
    GroupNames.Sort( true );

    numGroups = GroupNames.GetSize();
    for ( i = 0; i < numGroups; i++ )
    {
        GroupName = GroupNames.Get( i );
        GetGroupPrivileges( GroupName, &Privileges );
        numPrivileges = Privileges.GetSize();
        for ( j = 0; j < numPrivileges; j++ )
        {
            PrivilegeName = Privileges.Get( j );
            Record.Reset( PXS_CATEGORY_GROUPPOLICY );
            Record.Add( PXS_GROUPPOLICY_GROUP_NAME    , GroupName );
            Record.Add( PXS_GROUPPOLICY_PRIVILEGE_NAME, PrivilegeName );
            pRecords->Add( Record );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_GROUPPOLICY_GROUP_NAME );
}

//===============================================================================================//
//  Description:
//        Get the user names as an array of audit records
//
//  Parameters:
//        pRecords - array to receive the data
//
//  Remarks:
//      Normally, NetUserGetInfo does not require admin/account operator
//      privileges.
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetUserRecords( TArray< AuditRecord >* pRecords )
{
    bool      passwordExpired = false;
    time_t    timeT = 0;
    DWORD     index = 0;   // First call to NetQueryDisplayInformation
    DWORD     i = 0, returnedEntryCount = 0, lastIndex = 0;
    DWORD     passwordAge  = 0, lastLogon = 0, lastLogoff = 0;
    DWORD     numberLogons = 0, numBadLogons = 0, accountExpires = 0;
    String    LocalGroups, GlobalGroups, Value, UserName;
    String    AccountStatus, FullName, Comment, LocaleNo, LocaleYes;
    String    LocaleDays, LocaleDay, Temp;
    Formatter Format;
    AuditRecord       Record;
    NET_API_STATUS    status = 0;
    NET_DISPLAY_USER* pDisplayUsers = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Locale Strings
    PXSGetResourceString( PXS_IDS_1260_NO , &LocaleNo );
    PXSGetResourceString( PXS_IDS_1261_YES, &LocaleYes );
    PXSGetResourceString( PXS_IDS_144_DAYS, &LocaleDays );
    PXSGetResourceString( PXS_IDS_143_DAY , &LocaleDay  );

    do
    {
        // On Windows 2000 and newer the limit is 100 entries
        lastIndex = index;
        returnedEntryCount = 0;
        pDisplayUsers = nullptr;
        status = NetQueryDisplayInformation( nullptr,     // Local computer
                                             1,
                                             index,
                                             100,
                                             MAX_PREFERRED_LENGTH,
                                             &returnedEntryCount,
                                             reinterpret_cast<PVOID*>( &pDisplayUsers ) );
        Temp = Format.StringUInt32_3( L"NetQueryDisplayInformation result: status = %%1, "
                                      L"index=%%2, returnedEntryCount=%%3.",
                                      status, index, returnedEntryCount );
        PXSLogAppInfo( Temp.c_str() );
        if ( ( status != NERR_Success ) && ( status != ERROR_MORE_DATA ) )
        {
            throw Exception( PXS_ERROR_TYPE_NETWORK,
                             status,
                             L"NetQueryDisplayInformation", __FUNCTION__ );
        }

        if ( pDisplayUsers )
        {
            AutoNetApiBufferFree AutoFreeDisplayUsers( pDisplayUsers );

            for ( i = 0; i < returnedEntryCount; i++ )
            {
                UserName = pDisplayUsers[ i ].usri1_name;
                if ( UserName.GetLength() )
                {
                    TranslateUserAccountFlags( pDisplayUsers[ i ].usri1_flags,
                                               &AccountStatus );

                    // Get more information with NetUserGetInfo at level 11
                    passwordAge  = 0;          // Zero signifies unknown
                    lastLogon    = 0;          // Zero signifies unknown
                    lastLogoff   = 0;          // Zero signifies unknown
                    numberLogons = DWORD_MAX;  // -1 signifies unknown
                    numBadLogons = DWORD_MAX;  // -1 signifies unknown
                    GetUserInfo_11( UserName,
                                    &passwordAge,
                                    &lastLogon, &lastLogoff, &numberLogons, &numBadLogons );

                    // Get level 4 information
                    passwordExpired = false;
                    accountExpires  = TIMEQ_FOREVER;   // Implies no expiry
                    GetUserInfo_4( UserName, &passwordExpired, &accountExpires );

                    // Get Local Group membership as comma separated list
                    LocalGroups = PXS_STRING_EMPTY;
                    GetUserLocalGroups( UserName, &LocalGroups );

                    // Get Global Group membership as comma separated list
                    GlobalGroups = PXS_STRING_EMPTY;
                    GetUserGlobalGroups( UserName, &GlobalGroups );

                    // Make the record
                    FullName = pDisplayUsers[ i ].usri1_full_name;
                    Comment  = pDisplayUsers[ i ].usri1_comment;
                    Record.Reset( PXS_CATEGORY_USERS );
                    Record.Add( PXS_USERS_USERNAME      , UserName );
                    Record.Add( PXS_USERS_FULL_NAME     , FullName );
                    Record.Add( PXS_USERS_DESCRIPTION   , Comment );
                    Record.Add( PXS_USERS_ACCOUNT_STATUS, AccountStatus );
                    Record.Add( PXS_USERS_LOCAL_GROUPS  , LocalGroups );
                    Record.Add( PXS_USERS_GLOBAL_GROUPS , GlobalGroups );

                    // Last log on
                    Value = PXS_STRING_EMPTY;
                    if ( lastLogon )
                    {
                        // Value is in seconds from 1/1/70 00:00:00
                        timeT = PXSCastUInt32ToTimeT( lastLogon );
                        Value = Format.TimeTToLocalTimeInIso( timeT );
                    }
                    Record.Add( PXS_USERS_LAST_LOGON, Value );

                    // Last Log off, if the value is unknown use an empty string
                    Value = PXS_STRING_EMPTY;
                    if ( lastLogoff )
                    {
                        timeT = PXSCastUInt32ToTimeT( lastLogoff );
                        Value = Format.TimeTToLocalTimeInIso( timeT );
                    }
                    Record.Add( PXS_USERS_LAST_LOGOFF, Value );

                    // Number of log ons, -1 indicates unknown
                    Value = PXS_STRING_EMPTY;
                    if ( numberLogons != DWORD_MAX )
                    {
                        Value = Format.UInt32( numberLogons );
                    }
                    Record.Add( PXS_USERS_NUMBER_OF_LOGONS, Value );

                    // Number of bad log ons, -1 indicates unknown
                    Value = PXS_STRING_EMPTY;
                    if ( numBadLogons != DWORD_MAX )
                    {
                        Value = Format.UInt32( numBadLogons );
                    }
                    Record.Add( PXS_USERS_BAD_PASSWORD_COUNT, Value );

                    // Password age
                    Value  = Format.UInt32( passwordAge );
                    Value += PXS_CHAR_SPACE;
                    if ( passwordAge == 1 )
                    {
                        Value += LocaleDay;
                    }
                    else
                    {
                        Value += LocaleDays;
                    }
                    Record.Add( PXS_USERS_PASSWORD_AGE_DAY, Value );

                    // Add if the password has expired
                    Value = LocaleNo;
                    if ( passwordExpired )
                    {
                        Value = LocaleYes;
                    }
                    Record.Add( PXS_USERS_PASSWORD_EXPIRED, Value );

                    // Account expires
                    Value = PXS_STRING_EMPTY;
                    if ( ( accountExpires != 0 ) &&
                         ( accountExpires != TIMEQ_FOREVER ) )
                    {
                        timeT = PXSCastUInt32ToTimeT( accountExpires );
                        Value = Format.TimeTToLocalTimeInIso( timeT );
                    }
                    Record.Add( PXS_USERS_ACCOUNT_EXPIRES, Value );

                    pRecords->Add( Record );
                }
                index = pDisplayUsers[ i ].usri1_next_index;    // Next index
            }
        }  // Loop for more users
    } while ( ( index > lastIndex         ) &&  // Enumeration has advanced
              ( returnedEntryCount != 0   ) &&  // At least 1 entry on this pass
              ( status == ERROR_MORE_DATA )  );

    PXSSortAuditRecords( pRecords, PXS_USERS_USERNAME );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the comment associates with the specified global group
//
//  Parameters:
//      GroupName - the global group name
//      pComment  - receives the comment
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetGlobalGroupComment( const String& GroupName, String* pComment )
{
    GROUP_INFO_1*  pGroupInfo = nullptr;
    NET_API_STATUS status = NERR_Success;

    if ( pComment == nullptr )
    {
        throw ParameterException( L"pComment", __FUNCTION__ );
    }
    *pComment = PXS_STRING_EMPTY;

    if ( GroupName.IsEmpty() )
    {
        return;     // Nothing to do
    }
    // Query at level 1, requires Administrators or Account Operators
    // permissions
    status = NetGroupGetInfo( nullptr,
                              GroupName.c_str(), 1, reinterpret_cast<LPBYTE*>( &pGroupInfo ) );
    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetGroupGetInfo", __FUNCTION__ );
    }

    if ( pGroupInfo == nullptr )
    {
        throw NullException( L"pGroupInfo", __FUNCTION__ );
    }
    AutoNetApiBufferFree AutoFreeGroupInfo( pGroupInfo );
    *pComment = pGroupInfo->grpi1_comment;
}

//===============================================================================================//
//  Description:
//      Get the members of the specified global group
//
//  Parameters:
//      GroupName - name of global group
//      pMembers  - receives the member names
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetGlobalGroupMembers( const String& GroupName, StringArray* pMembers )
{
    DWORD       i = 0, entriesRead = 0, totalEntries = 0;
    DWORD_PTR   ResumeHandle = 0;
    AuditRecord Record;
    NET_API_STATUS      status = NERR_Success;
    GROUP_USERS_INFO_0* pGroupUsers = nullptr;

    if ( pMembers == nullptr )
    {
        throw ParameterException( L"pMembers", __FUNCTION__ );
    }
    pMembers->RemoveAll();

    if ( GroupName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Get data, requires Administrators or Account Operators permissions
    ResumeHandle = 0;    // Must set to zero for first call
    do
    {
        // Get global groups
        pGroupUsers = nullptr;
        entriesRead = 0;
        status = NetGroupGetUsers( nullptr,
                                   GroupName.c_str(),
                                   0,
                                   reinterpret_cast<LPBYTE*>( &pGroupUsers ),
                                   MAX_PREFERRED_LENGTH,
                                   &entriesRead, &totalEntries, &ResumeHandle );

        if ( ( status != NERR_Success ) && ( status != ERROR_MORE_DATA ) )
        {
            throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetGroupGetUsers", __FUNCTION__ );
        }

        if ( pGroupUsers )
        {
            AutoNetApiBufferFree AutoFreeGroupUsers( pGroupUsers );

            for ( i = 0; i < entriesRead; i++ )
            {
                pMembers->Add( pGroupUsers[ i ].grui0_name );
            }
        }
    }  while ( ( entriesRead  != 0 ) &&    // At least one entry was obtained
               ( ResumeHandle != 0 ) &&    // The handle should not be zero
               ( status == ERROR_MORE_DATA ) );
}

//===============================================================================================//
//  Description:
//      Get the names of the global groups
//
//  Parameters:
//      pNames - array to receive the global group names
//
//  Remarks:
//      No special permissions required to enumerate global groups
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetGlobalGroupNames( StringArray* pNames )
{
    DWORD     index = 0;  // Zero for first call to NetQueryDisplayInformation
    DWORD     i = 0, returnedEntryCount = 0, lastIndex = 0;
    NET_API_STATUS     status;
    NET_DISPLAY_GROUP* pGlobalGroup = nullptr;

    if ( pNames == nullptr )
    {
        throw ParameterException( L"pNames", __FUNCTION__ );
    }
    pNames->RemoveAll();

    do
    {
        // Maximum number of entries, on Windows 2000 and newer the limit
        // is 100 entries
        lastIndex = index;
        returnedEntryCount = 0;
        pGlobalGroup = nullptr;
        status = NetQueryDisplayInformation(
                                    nullptr,
                                    3,
                                    index,
                                    100,
                                    MAX_PREFERRED_LENGTH,
                                    &returnedEntryCount,
                                    reinterpret_cast<PVOID*>( &pGlobalGroup ) );
        if ( ( status != NERR_Success ) && ( status != ERROR_MORE_DATA ) )
        {
            throw Exception( PXS_ERROR_TYPE_NETWORK,
                             status,
                             L"NetQueryDisplayInformation", __FUNCTION__ );
        }

        if ( pGlobalGroup )
        {
            AutoNetApiBufferFree AutoFreeGlobalGroup( pGlobalGroup );

            for ( i = 0; i < returnedEntryCount; i++ )
            {
                pNames->Add( pGlobalGroup[ i ].grpi3_name );

                // Need to store as the index for NetQueryDisplayInformation
                if ( i == ( returnedEntryCount - 1 ) )
                {
                    index = pGlobalGroup->grpi3_next_index;
                }
            }
        }
    } while ( ( index > lastIndex         ) &&  // Enumeration has advanced
              ( returnedEntryCount != 0   ) &&  // Ensure read at least 1 entry
              ( status == ERROR_MORE_DATA )  );

    pNames->Sort( true );
}

//===============================================================================================//
//  Description:
//        Fill an array with the policy rights of the specified local group
//
//  Parameters:
//      GroupName   - group name
//      pPrivileges - array to receive the policy rights
//
//  Remarks:
//      Names of users, local groups and global groups are unique.
//      The domain prefix is not required to get the rights for a
//      a global groups
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetGroupPrivileges( const String& GroupName, StringArray* pPrivileges )
{
    NTSTATUS  status = 0;
    ULONG     i = 0, countOfRights = 0;
    wchar_t   ReferencedDomain [ 256 + 1 ];  // Enough for a domain name
    wchar_t   szWide[ 256 + 1 ];             // Enough for a privilege descrip.
    DWORD     cbSid = 0, cchReferencedDomain = 0, lastError = 0;
    String    Insert2, PiviliegeName;
    Formatter Format;
    AllocateBytes AllocBytes;
    PSID         Sid = nullptr;
    LSA_HANDLE   PolicyHandle = nullptr;
    SID_NAME_USE eUse;
    LSA_UNICODE_STRING*   UserRights = nullptr;
    LSA_OBJECT_ATTRIBUTES ObjectAttributes;

    if ( pPrivileges == nullptr )
    {
        throw ParameterException( L"pPrivileges", __FUNCTION__ );
    }
    pPrivileges->RemoveAll();

    if ( GroupName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Get the SID by looking up the account name. Do a first call
    // to get the required buffer size.
    cchReferencedDomain = ARRAYSIZE ( ReferencedDomain );
    if ( LookupAccountName( nullptr,
                            GroupName.c_str(),
                            nullptr,
                            &cbSid,
                            ReferencedDomain,
                            &cchReferencedDomain,  // In chars
                            &eUse ) == 0 )
    {
        lastError = GetLastError();
        if ( lastError != ERROR_INSUFFICIENT_BUFFER )
        {
            // Some other error
            throw SystemException( lastError,
                                   L"LookupAccountName", __FUNCTION__ );
        }
    }

    // Make there is data to be had, otherwise nothing to do
    if ( cbSid == 0 )
    {
        PXSLogAppInfo1( L"There is no data for group '%%1'.", GroupName );
        return;
    }
    Sid = reinterpret_cast<PSID>( AllocBytes.New( cbSid ) );

    // Second call
    cchReferencedDomain = ARRAYSIZE ( ReferencedDomain );
    if ( LookupAccountName( nullptr,
                            GroupName.c_str(),
                            Sid,
                            &cbSid, ReferencedDomain, &cchReferencedDomain, &eUse ) == 0 )
    {
        throw SystemException( GetLastError(), L"LookupAccountName", __FUNCTION__ );
    }

    // Open the policy with lookup access
    memset( &ObjectAttributes, 0, sizeof ( ObjectAttributes ) );
    status = LsaOpenPolicy( nullptr,               // Local computer
                            &ObjectAttributes, POLICY_LOOKUP_NAMES, &PolicyHandle );
    if ( status )
    {
        throw SystemException( LsaNtStatusToWinError( status ), L"LsaOpenPolicy", __FUNCTION__ );
    }

    if ( PolicyHandle == nullptr )
    {
        throw NullException( L"PolicyHandle", __FUNCTION__ );
    }

    status = LsaEnumerateAccountRights( PolicyHandle, Sid, &UserRights, &countOfRights );
    // First test for STATUS_OBJECT_NAME_NOT_FOUND because this is returned
    // if there are no rights for the account.
    if ( status == STATUS_OBJECT_NAME_NOT_FOUND )
    {
        // Nothing to do
        LsaClose( PolicyHandle );
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo2( L"No privileges found for '%%1' in '%%2'.", GroupName, Insert2 );
        return;
    }

    if ( status )
    {
        LsaClose( PolicyHandle );
        throw SystemException( LsaNtStatusToWinError( status ),
                               L"LsaEnumerateAccountRights", __FUNCTION__);
    }

    if ( UserRights == nullptr )
    {
        // Nothing to do
        LsaClose( PolicyHandle );
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo2( L"pUserRights = 0 for '%%1' in '%%2'", GroupName, Insert2 );
        return;
    }

    // Catch exceptions so can clean up
    try
    {
        for ( i = 0; i < countOfRights; i++ )
        {
            if ( UserRights[ i ].Buffer )
            {
                memset( szWide, 0, sizeof ( szWide ) );
                StringCchCopyN( szWide,
                                ARRAYSIZE( szWide ),
                                UserRights[ i ].Buffer,
                                UserRights[ i ].Length / sizeof ( wchar_t ) );
                PiviliegeName = szWide;
                pPrivileges->Add( szWide );
            }
        }
    }
    catch ( const Exception& )
    {
        LsaFreeMemory( UserRights );
        LsaClose( PolicyHandle );
        throw;
    }
    LsaFreeMemory( UserRights );
    LsaClose( PolicyHandle );
}

//===============================================================================================//
//  Description:
//      Get the comment associates with the specified local group
//
//  Parameters:
//      GroupName - the groups's name
//      pComment  - receives the comment
//
//  Remarks:
//      Will log errors rather than throw as this method is used to get
//      optional information
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetLocalGroupComment( const String& GroupName, String* pComment )
{
    NET_API_STATUS     status = 0;
    LOCALGROUP_INFO_1* pLocalGroup = nullptr;

    if ( pComment == nullptr )
    {
        throw ParameterException( L"pComment", __FUNCTION__ );
    }
    *pComment = PXS_STRING_EMPTY;

    if ( GroupName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Get the comment at Level 1
    status = NetLocalGroupGetInfo( nullptr,         // Local server
                                   GroupName.c_str(),
                                   1,               // Only Level 1 allowed
                                   reinterpret_cast<LPBYTE*>( &pLocalGroup ) );
    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetLocalGroupGetInfo", __FUNCTION__ );
    }

    if ( pLocalGroup == nullptr )
    {
        throw NullException( L"pLocalGroup", __FUNCTION__ );
    }
    AutoNetApiBufferFree AutoFreeLocalGroup( pLocalGroup );
    *pComment = pLocalGroup->lgrpi1_comment;
}

//===============================================================================================//
//  Description:
//      Get the members of the specified local group at information level 1
//
//  Parameters:
//      GroupName - name of the local group
//      pMembers  - array to to receive the member names
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetLocalGroupMembers1( const String& GroupName, StringArray* pMembers )
{
    DWORD     i = 0, entriesRead = 0, totalEntries = 0;
    DWORD_PTR resumeHandle = 0;  // Must set to zero for first call
    NET_API_STATUS  status = 0;
    LOCALGROUP_MEMBERS_INFO_1* pInfo1 = nullptr;

    if ( pMembers == nullptr )
    {
        throw ParameterException( L"pMembers", __FUNCTION__ );
    }
    pMembers->RemoveAll();

    if ( GroupName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    do
    {
        pInfo1      = nullptr;
        entriesRead = 0;
        status = NetLocalGroupGetMembers( nullptr,   // Local computer
                                          GroupName.c_str(),
                                          1,
                                          reinterpret_cast<LPBYTE*>( &pInfo1 ),
                                          MAX_PREFERRED_LENGTH,
                                          &entriesRead, &totalEntries, &resumeHandle );
        if ( ( status != NERR_Success ) && ( status != ERROR_MORE_DATA ) )
        {
            throw Exception( PXS_ERROR_TYPE_NETWORK,
                             status, L"NetLocalGroupGetMembers", __FUNCTION__ );
        }

        if ( pInfo1 )
        {
            AutoNetApiBufferFree AutoFreeMembers( pInfo1 );
            for ( i = 0; i < entriesRead; i++ )
            {
                pMembers->Add( pInfo1[ i ].lgrmi1_name );
            }
        }
    }  while ( ( entriesRead  != 0 ) &&    // At least one entry was obtained
               ( resumeHandle != 0 ) &&    // The handle should not be zero
               ( status == ERROR_MORE_DATA ) );
}


//===============================================================================================//
//  Description:
//      Get the members of the specified local group at information level 3
//
//  Parameters:
//      GroupName - name of the local group
//      pMembers  - array to to receive the member names
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetLocalGroupMembers3( const String& GroupName, StringArray* pMembers )
{
    DWORD     i = 0, entriesRead = 0, totalEntries = 0;
    DWORD_PTR resumeHandle = 0;  // Must set to zero for first call
    NET_API_STATUS  status = 0;
    LOCALGROUP_MEMBERS_INFO_3* pInfo3 = nullptr;

    if ( pMembers == nullptr )
    {
        throw ParameterException( L"pMembers", __FUNCTION__ );
    }
    pMembers->RemoveAll();

    if ( GroupName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    do
    {
        pInfo3      = nullptr;
        entriesRead = 0;
        status = NetLocalGroupGetMembers( nullptr,   // Local computer
                                          GroupName.c_str(),
                                          3,
                                          reinterpret_cast<LPBYTE*>( &pInfo3 ),
                                          MAX_PREFERRED_LENGTH,
                                          &entriesRead, &totalEntries, &resumeHandle );
        if ( ( status != NERR_Success ) && ( status != ERROR_MORE_DATA ) )
        {
            throw Exception( PXS_ERROR_TYPE_NETWORK,
                             status, L"NetLocalGroupGetMembers", __FUNCTION__ );
        }

        if ( pInfo3 )
        {
            AutoNetApiBufferFree AutoFreeMembers( pInfo3 );
            for ( i = 0; i < entriesRead; i++ )
            {
                pMembers->Add( pInfo3[ i ].lgrmi3_domainandname );
            }
        }
    }  while ( ( entriesRead  != 0 ) &&    // At least one entry was obtained
               ( resumeHandle != 0 ) &&    // The handle should not be zero
               ( status == ERROR_MORE_DATA ) );
}

//===============================================================================================//
//  Description:
//      Get the names of the local groups
//
//  Parameters:
//      pGroupNames - array to receive the local group names
//
//  Remarks:
//      Only members of the Administrators or Account Operators local group
//      can successfully execute NetLocalGroupEnum.
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetLocalGroupNames( StringArray* pGroupNames )
{
    DWORD     i = 0, entriesRead = 0, totalEntries = 0;
    DWORD_PTR      resumeHandle  = 0;  // Must set to zero for first call
    NET_API_STATUS status = 0;
    LOCALGROUP_INFO_1* pLocalGroup = nullptr;

    if ( pGroupNames == nullptr )
    {
        throw ParameterException( L"pGroupNames", __FUNCTION__ );
    }
    pGroupNames->RemoveAll();

    do
    {
        pLocalGroup = nullptr;
        entriesRead = 0;
        status = NetLocalGroupEnum( nullptr,  // Local computer
                                    1,
                                    reinterpret_cast<LPBYTE*>( &pLocalGroup ),
                                    MAX_PREFERRED_LENGTH,
                                    &entriesRead, &totalEntries, &resumeHandle );
        if ( ( status != NERR_Success ) && ( status != ERROR_MORE_DATA ) )
        {
            throw Exception( PXS_ERROR_TYPE_NETWORK,
                             status,
                             L"NetLocalGroupEnum", __FUNCTION__ );
        }

        // Process the returned data
        if ( pLocalGroup )
        {
            AutoNetApiBufferFree AutoFreeLocalGroup( pLocalGroup );

            for ( i = 0; i < entriesRead; i++ )
            {
                pGroupNames->Add( pLocalGroup[ i ].lgrpi1_name );
            }
        }
    }  while ( ( entriesRead != 0 ) &&    // Got at least one entry was
               ( resumeHandle!= 0 ) &&    // The handle should not be zero
               ( status == ERROR_MORE_DATA ) );

    pGroupNames->Sort( true );
}

//===============================================================================================//
//  Description:
//      Get the global groups of which the specified user is a member
//
//  Parameters:
//      UserName      - the user's name
//      pGlobalGroups - receives the a comma separated list of group names
//
//  Remarks:
//      Will log errors rather than throw as this method is used to get
//      optional information
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetUserGlobalGroups( const String& UserName, String* pGlobalGroups )
{
    DWORD  i = 0, entriesRead = 0, totalEntries = 0;
    String GroupName, Insert2;
    NET_API_STATUS      status = 0;
    GROUP_USERS_INFO_0* pGroupUsers = nullptr;

    if ( pGlobalGroups == nullptr )
    {
        throw ParameterException( L"pGlobalGroups", __FUNCTION__ );
    }
    *pGlobalGroups = PXS_STRING_EMPTY;

    if ( UserName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Get a much data as possible in 1 go using MAX_PREFERRED_LENGTH
    status = NetUserGetGroups( nullptr,        // Local machine
                               UserName.c_str(),
                               0,              // Level 0
                               reinterpret_cast<LPBYTE*>( &pGroupUsers ),
                               MAX_PREFERRED_LENGTH, &entriesRead, &totalEntries );
    if ( status != NERR_Success &&
         status != ERROR_MORE_DATA )
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogNetError2( status,
                         L"NetUserGetGroups failed for user '%%1' in '%%2'.", UserName, Insert2 );
        return;
    }

    if ( pGroupUsers == nullptr )
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppError2( L"No data for user '%%1' in '%%2'.", UserName, Insert2 );
        return;
    }
    AutoNetApiBufferFree AutoFreeGroupUsers( pGroupUsers );

    for ( i = 0; i < entriesRead; i++ )
    {
        GroupName = pGroupUsers->grui0_name;
        if ( GroupName.GetLength() )
        {
            if ( pGlobalGroups->GetLength() )
            {
                *pGlobalGroups += L", ";
            }
            *pGlobalGroups += GroupName;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get log on information about the specified user at level 4
//
//  Parameters:
//      UserName         - pointer to the user's name
//      pPasswordExpired - receives if the password has expired
//      pPasswordExpired - receives when the account expires
//
//  Remarks:
//      See USER_INFO_4 Structure
//
//      Will log errors rather than throw as this method is used to get
//      optional information
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetUserInfo_4( const String& UserName,
                                          bool*   pPasswordExpired, DWORD* pAccountExpires )
{
    String Insert2;
    USER_INFO_4*   pUserInfo = nullptr;
    NET_API_STATUS status = 0;

    if ( ( pPasswordExpired == nullptr ) || ( pAccountExpires == nullptr ) )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pPasswordExpired = false;
    *pAccountExpires  = TIMEQ_FOREVER;

    if ( UserName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    status = NetUserGetInfo( nullptr,
                             UserName.c_str(),
                             4,            // Level 4
                             reinterpret_cast<LPBYTE*>( &pUserInfo ) );
    if ( status != NERR_Success )
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogNetError2( status,
                         L"NetUserGetInfo(4) failed for user '%%1' in '%%2'.", UserName, Insert2 );
        return;
    }

    if ( pUserInfo == nullptr )
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppError2( L"No data for user '%%1' in '%%2'.", UserName, Insert2 );
        return;
    }
    AutoNetApiBufferFree AutoFreeUserInfo( pUserInfo );

    if ( pUserInfo->usri4_password_expired )
    {
        *pPasswordExpired = true;
    }
    *pAccountExpires = pUserInfo->usri4_acct_expires;
}

//===============================================================================================//
//  Description:
//      Get logon information about the specified user at level 11
//
//  Parameters:
//      UserName      - pointer to the user's name
//      pPasswordAge  - receives the password age in days
//      pLastLogon    - receives the last logon time
//      pLastLogoff   - receives the last logoff time
//      pNumberLogons - receives the number of logons
//      pNumBadLogons - receives the number of bad logon attempts
//
//  Remarks:
//      See USER_INFO_11 Structure
//
//      Will log errors rather than throw as this method is used to get
//      optional information
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetUserInfo_11( const String& UserName,
                                           DWORD* pPasswordAge,
                                           DWORD* pLastLogon,
                                           DWORD* pLastLogoff,
                                           DWORD* pNumberLogons, DWORD* pNumBadLogons )
{
    String Insert2;
    USER_INFO_11*  pUserInfo = nullptr;
    NET_API_STATUS status    = 0;

    if ( ( pPasswordAge  == nullptr ) ||
         ( pLastLogon    == nullptr ) ||
         ( pLastLogoff   == nullptr ) ||
         ( pNumberLogons == nullptr ) ||
         ( pNumBadLogons == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pPasswordAge  = 0;          // = unknown
    *pLastLogon    = 0;          // = unknown
    *pLastLogoff   = 0;          // = unknown
    *pNumberLogons = DWORD_MAX;  // = unknown
    *pNumBadLogons = DWORD_MAX;  // = unknown

    if ( UserName.IsEmpty() )
    {
        return;     // Nothing to do
    }
    status = NetUserGetInfo( nullptr,
                             UserName.c_str(),
                             11,            // Level 11
                             reinterpret_cast<LPBYTE*>( &pUserInfo ) );
    if ( status != NERR_Success )
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogNetError2( status,
                         L"NetUserGetInfo failed for user '%%1' in '%%2'.", UserName, Insert2 );
        return;
    }

    if ( pUserInfo == nullptr )
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppError2( L"No data for user '%%1' in '%%2'.", UserName, Insert2 );
        return;
    }
    AutoNetApiBufferFree AutoFreeUserInfo( pUserInfo );

    *pPasswordAge  = pUserInfo->usri11_password_age /( 24 * 3600 );
    *pLastLogon    = pUserInfo->usri11_last_logon;
    *pLastLogoff   = pUserInfo->usri11_last_logoff;
    *pNumberLogons = pUserInfo->usri11_num_logons;
    *pNumBadLogons = pUserInfo->usri11_bad_pw_count;
}

//===============================================================================================//
//  Description:
//      Get the local groups to which the user belongs
//
//  Parameters:
//      UserName     - pointer to the user's name
//      pLocalGroups - receives the a comma separated list of group names
//
//  Remarks:
//      Will log errors rather than throw as this method is used to get
//      optional information
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::GetUserLocalGroups( const String& UserName, String* pLocalGroups )
{
    DWORD   i = 0, entriesRead = 0, totalEntries = 0;
    String  Insert2, GroupName;
    NET_API_STATUS status = 0;
    LOCALGROUP_USERS_INFO_0* pLocalGroup = nullptr;

    if ( pLocalGroups == nullptr )
    {
        throw ParameterException( L"pLocalGroups", __FUNCTION__ );
    }
    *pLocalGroups = PXS_STRING_EMPTY;

    if ( UserName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Get a much data as possible in 1 go using MAX_PREFERRED_LENGTH
    status = NetUserGetLocalGroups( nullptr,       // Local machine
                                    UserName.c_str(),
                                    0,             // Only level 0 is valid
                                    0,
                                    reinterpret_cast<LPBYTE*>( &pLocalGroup ),
                                    MAX_PREFERRED_LENGTH, &entriesRead, &totalEntries );
    if ( ( status != NERR_Success ) && ( status != ERROR_MORE_DATA ) )
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogNetError2( status,
                         L"NetUserGetLocalGroups failed for '%%1' in '%%2'.", UserName, Insert2 );
       return;
    }

    if ( pLocalGroup == nullptr )
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo2( L"No groups found for user '%%1' in '%%2'.", UserName, Insert2 );
        return;
    }
    AutoNetApiBufferFree AutoFreeLocalGroup( pLocalGroup );

    for ( i = 0; i < entriesRead; i++ )
    {
        // Make sure name looks sensible
        GroupName = pLocalGroup[ i ].lgrui0_name;
        if ( GroupName.GetLength() )
        {
            if ( pLocalGroups->GetLength() )
            {
                *pLocalGroups += L", ";
            }
            *pLocalGroups += GroupName;
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a user's account flags
//
//  Parameters:
//      flags        - bit mask of the account flags
//      pTranslation - receives the translation
//
//  Remarks:
//      See NET_DISPLAY_USER::usri1_flags
//
//  Returns:
//      void
//===============================================================================================//
void GroupUserInformation::TranslateUserAccountFlags( DWORD flags, String* pTranslation )
{
    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    if ( UF_ACCOUNTDISABLE & flags )
    {
        *pTranslation += L"Disabled, ";
    }
    else
    {
        *pTranslation += L"Enabled, ";
    }

    // Account Locked out
    if ( UF_LOCKOUT & flags )
    {
        *pTranslation += L"Locked";
    }
    else
    {
        *pTranslation += L"Not Locked";
    }

    // Password required
    if ( UF_PASSWD_NOTREQD & flags )
    {
        *pTranslation += L"Password not required, ";
    }
    else
    {
        *pTranslation += L"Password is required, ";
    }

    // Password can be changed
    if ( UF_PASSWD_CANT_CHANGE & flags )
    {
        *pTranslation += L"Cannot change password, ";
    }
    else
    {
        *pTranslation += L"Can change password, ";
    }

    // Password has expire date
    if ( UF_DONT_EXPIRE_PASSWD & flags )
    {
        *pTranslation += L"Password never expires, ";
    }
    else
    {
        *pTranslation += L"Password can expire, ";
    }

    // Password has expired
    if ( UF_PASSWORD_EXPIRED & flags )
    {
        *pTranslation += L"Password has expired";
    }
    else
    {
        *pTranslation += L"Password has not expired";
    }
}
