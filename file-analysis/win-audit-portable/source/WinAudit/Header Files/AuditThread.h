///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Audit Thread Class Header
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

#ifndef WINAUDIT_AUDIT_THREAD_H_
#define WINAUDIT_AUDIT_THREAD_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Thread.h"
#include "PxsBase/Header Files/Mutex.h"

// 5. This Project
#include "WinAudit/Header Files/AuditThreadParameter.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class AuditThread : public Thread
{
    public:
        // Default constructor
        AuditThread();

        // Destructor
        ~AuditThread();

        // Methods
        void SetAuditThreadParameter( AuditThreadParameter* pParameter );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        AuditThread( const AuditThread& oAuditThread );

        // Assignment operator - not allowed
        AuditThread& operator= ( const AuditThread& oAuditThread );

        // Methods
        DWORD   DoAuditLocked( AuditThreadParameter* pParameter );
        DWORD   RunWorkerThread();

        // Data members
        Mutex   m_Mutex;
        AuditThreadParameter*   m_pAuditThreadParameterMT;  // Shared variable
};

#endif  // WINAUDIT_AUDIT_THREAD_H_
