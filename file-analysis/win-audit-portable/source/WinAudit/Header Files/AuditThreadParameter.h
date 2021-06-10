///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Audit Thread Parameter Class Header
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

#ifndef WINAUDIT_AUDIT_THREAD_PARAMETER_H_
#define WINAUDIT_AUDIT_THREAD_PARAMETER_H_

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
#include "PxsBase/Header Files/TArray.h"

// 5. This Project

// 6. Forwards
class WinAuditFrame;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class AuditThreadParameter
{
    public:
        // Default constructor
        AuditThreadParameter();

        // Copy constructor
        AuditThreadParameter( const AuditThreadParameter& oParameter );

        // Destructor
        ~AuditThreadParameter();

        // Assignment operator
        AuditThreadParameter& operator= ( const AuditThreadParameter& oParameter );

        // Methods

        // Data members
        time_t          timeoutAt;       // When the thread times out
        String          LocalTime;       // The start time of the audit
        WinAuditFrame*  pWinAuditFrame;  // The application's main frame
        TArray< DWORD > Categories;      // The data categories to get

    protected:
        // Methods

        // Data members

    private:
        // Methods

        // Data members
};

#endif  // WINAUDIT_AUDIT_THREAD_PARAMETER_H_
