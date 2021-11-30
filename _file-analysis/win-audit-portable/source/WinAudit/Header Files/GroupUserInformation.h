///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Group and User Information Class Header
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

#ifndef WINAUDIT_GROUP_USER_INFORMATION_H_
#define WINAUDIT_GROUP_USER_INFORMATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards
class AuditRecord;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class GroupUserInformation
{
    public:
        // Default constructor
        GroupUserInformation();

        // Destructor
        ~GroupUserInformation();

        // Methods
        void    GetGroupRecords( TArray< AuditRecord >* pRecords );
        void    GetGroupMemberRecords( TArray< AuditRecord >* pRecords );
        void    GetGroupPolicyRecords( TArray< AuditRecord >* pRecords );
        void    GetUserRecords( TArray< AuditRecord >* pRecords );

    protected:
        // Methods

       // Data members

    private:
        // Copy constructor - not allowed
        GroupUserInformation(
                          const GroupUserInformation& oGroupUserInformation );

        // Assignment operator - not allowed
        GroupUserInformation& operator= (
                           const GroupUserInformation& oGroupUserInformation );

        // Methods
        void    GetGlobalGroupComment( const String& GroupName, String* pComment );
        void    GetGlobalGroupMembers( const String& GroupName, StringArray* pMembers );
        void    GetGlobalGroupNames( StringArray* pNames );
        void    GetGroupPrivileges( const String& GroupName, StringArray* pPrivileges );
        void    GetLocalGroupComment( const String& GroupName, String* pComment );
        void    GetLocalGroupMembers1( const String& GroupName, StringArray* pMembers );
        void    GetLocalGroupMembers3( const String& GroupName, StringArray* pMembers );
        void    GetLocalGroupNames( StringArray* pGroupNames );
        void    GetUserGlobalGroups( const String& UserName, String* pGlobalGroups );
        void    GetUserInfo_4( const String& UserName, bool* pPasswordExpired,
                               DWORD* pAccountExpires );
        void    GetUserInfo_11( const String& UserName,
                                DWORD* pPasswordAge,
                                DWORD* pLastLogon,
                                DWORD* pLastLogoff,
                                DWORD* pNumberLogons, DWORD* pNumBadLogons );
        void    GetUserLocalGroups( const String& UserName, String* pLocalGroups );
        void    TranslateUserAccountFlags( DWORD flags, String* pTranslation );
        // Data members
};

#endif  // WINAUDIT_GROUP_USER_INFORMATION_H_
