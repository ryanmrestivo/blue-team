///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Audit Data Class Header
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

#ifndef WINAUDIT_AUDIT_DATA_H_
#define WINAUDIT_AUDIT_DATA_H_

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
#include "WinAudit/Header Files/SmbiosInformation.h"

// 6. Forwards
class AuditRecord;
class String;
class StringArray;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class AuditData
{
    public:
        // Default constructor
        AuditData();

        // Destructor
        ~AuditData();

        // Methods
        void    GetBIOSIdentificationRecord( AuditRecord* pRecord );
        void    GetCategoryName( DWORD categoryID, String* pCategoryName );
        void    GetCategoryRecords( DWORD categoryID,
                                    const String& LocalTime,
                                    TArray< AuditRecord >* pRecords );
        void    GetEnvironmentVarsRecords( TArray< AuditRecord >* pRecords );
        DWORD   GetMemoryInformationRecord( AuditRecord* pRecord );
        void    GetOleDbProviderRecords( TArray< AuditRecord >* pRecords );
        void    GetOleDbProviders( TArray< NameValue >* pOleDbProviders );
        void    GetOperatingSystemRecord( AuditRecord* pRecord );
        void    GetOtherIdentifier( String* pValue );
        void    GetRegionalSettingsRecords( TArray< AuditRecord >* pRecords );
        void    GetSystemFilesRecords( LPCWSTR pszFilter,
                                       TArray< AuditRecord >* pRecords );
        void    GetSystemOverviewRecord( const String& LocalTime,
                                   AuditRecord* pRecord );
        void    MakeAuditMasterRecord( AuditRecord* pAuditMaster );
        void    MakeComputerMasterRecord( AuditRecord* pComputerMaster );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        AuditData( const AuditData& oAuditData );

        // Assignment operator - not allowed
        AuditData& operator= ( const AuditData& oAuditData );

        // Methods
        void    TranslateCalenderType( LPCWSTR pszCalendar, String* pMeaning );

        // Data members
        SmbiosInformation   m_SmbiosInfo;
};

#endif  // WINAUDIT_AUDIT_DATA_H_
