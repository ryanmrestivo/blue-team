///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Audit Database Class Header
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

#ifndef WINAUDIT_AUDIT_DATABASE_H_
#define WINAUDIT_AUDIT_DATABASE_H_

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
#include "WinAudit/Header Files/OdbcDatabase.h"

// 6. Forwards
class AuditRecord;
class SmbiosInformation;
class String;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class AuditDatabase : public OdbcDatabase
{
    public:
        // Default constructor
        AuditDatabase();

        // Destructor
        ~AuditDatabase();

        // Methods
        SQLINTEGER IdentifyComputerID( const AuditRecord& Record );
        SQLINTEGER InsertAuditMaster(  SQLINTEGER computerID, const AuditRecord& AuditMaster );
        SQLINTEGER InsertComputerMaster( const AuditRecord& ComputerMaster );
        void UpdateComputerMaster( const AuditRecord& ComputerMaster,
                                   SQLINTEGER computerID, SQLINTEGER auditID );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        AuditDatabase( const AuditDatabase& oAuditDatabase );

        // Assignment operator - not allowed
        AuditDatabase& operator= ( const AuditDatabase& oAuditDatabase );

        // Methods
        void ComputerMasterRecordToValues( const AuditRecord &ComputerMaster,
                                           String* pMacAddress,
                                           String* pSmbiosUUID,
                                           String* pAssetTag,
                                           String* pFQDN,
                                           String* pSiteName,
                                           String* pDomainName,
                                           String* pComputerName,
                                           String* pOsProductID,
                                           String* pOtherIdentifier, String* pWinAuditGUID);
        // Data members
};

#endif  // WINAUDIT_AUDIT_DATABASE_H_
