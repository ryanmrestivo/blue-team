///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Peripheral Information Class Header
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

#ifndef WINAUDIT_PERIPHERAL_INFORMATION_H_
#define WINAUDIT_PERIPHERAL_INFORMATION_H_

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

class PeripheralInformation
{
    public:
        // Default constructor
        PeripheralInformation();

        // Destructor
        ~PeripheralInformation();

        // Methods
        void   GetAuditRecords( TArray< AuditRecord >* pRecords );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        PeripheralInformation( const PeripheralInformation& oPeripherals );

        // Assignment operator - not allowed
        PeripheralInformation& operator= ( const PeripheralInformation& oPeripherals );

        // Methods
        void GetDisplayRecord( AuditRecord* pDisplayRecord );
        void GetMappedDriveRecords( TArray<AuditRecord>* pMappedDriveRecords );
        void GetKeyboardRecord( AuditRecord* pKeyboardRecord );
        void GetMouseRecord( AuditRecord* pMouseRecord );
        void GetNetworkRecord( AuditRecord* pNetworkRecord );
        void GetPrinterRecords( TArray< AuditRecord >* pPrinterRecords );

        // Data members
};

#endif  // WINAUDIT_PERIPHERAL_INFORMATION_H_
