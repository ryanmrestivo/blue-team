///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Security Information Class Header
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

#ifndef WINAUDIT_SECURITY_INFORMATION_H_
#define WINAUDIT_SECURITY_INFORMATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files
#include <NTSecAPI.h>
#include <wuapi.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards
class AuditRecord;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class SecurityInformation
{
    public:
        // Default constructor
        SecurityInformation();

        // Destructor
        ~SecurityInformation();

        // Methods
        void GetLoggedOnUserPrivilegePrecords( TArray<AuditRecord>* pRecords );
        void GetNetworkTimeProtocolRecords( TArray< AuditRecord >* pRecords );
        void GetRegistrySecurityValueRecods( TArray< AuditRecord >* pRecords );
        void GetRsopHtmlResult( String* pRsopHtmlResult );
        void GetRSOPSecuritySetting( LPCWSTR pszKeyName, VARTYPE varType, String* pSetting );
        void GetSecuritySettingsRecords( TArray< AuditRecord >* pRecords );
        void GetSystemRestorePointRecords( TArray< AuditRecord >* pRecords );
        void GetUserRightsAssignmentRecords( TArray< AuditRecord >* pRecords );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        SecurityInformation( const SecurityInformation& oSecurity );

        // Assignment operator - not allowed
        SecurityInformation& operator= ( const SecurityInformation& oSecurity );

        // Methods
        void    EnumerationInformationToAccounts( const LSA_ENUMERATION_INFORMATION* pEnumeration,
                                                  ULONG count, String* pAccounts );
        void    MakeAccountLockingRecords( TArray< AuditRecord >* pRecords );
        void    MakeAccountSettingsRecords( TArray< AuditRecord >* pRecords );
        void    MakeAuditPolicyRecords( TArray< AuditRecord >* pRecords );
        void    MakeAutoLogonRecord( AuditRecord* pRecord );
        void    MakeAutomaticUpdateRecords( TArray< AuditRecord >* pRecords );
        void    MakeExecuteDataRecord( TArray< AuditRecord >* pRecords );
        void    MakeInternetExplorerRecords( TArray< AuditRecord >* pRecords );
        void    MakeNetworkAccessRecord( AuditRecord* pRecord );
        void    MakeNetworkSecurityRecords( TArray< AuditRecord >* pRecords );
        void    MakePasswordPolicyRecords( TArray< AuditRecord >* pRecords );
        void    MakeScreenSaverRecords( TArray< AuditRecord >* pRecords );
        void    TranslateAutomaticUpdateDay( AutomaticUpdatesScheduledInstallationDay day,
                                             String* pUpdateDay );
        void    TranslateAutomaticUpdateStatus( AutomaticUpdatesNotificationLevel level,
                                                String* pUpdateStatus );
        void    TranslateInternetActionPolicy( DWORD policy, String* pActionPolicy );
        void    TranslatePolicyAuditEventType( POLICY_AUDIT_EVENT_TYPE type,
                                               String* pAuditEventType );
        // Data members
};

#endif  // WINAUDIT_SECURITY_INFORMATION_H_
