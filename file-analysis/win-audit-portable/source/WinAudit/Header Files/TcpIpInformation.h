///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Network TCP/IP Information Class Header
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

#ifndef WINAUDIT_TCP_IP_INFORMATION_H_
#define WINAUDIT_TCP_IP_INFORMATION_H_

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

class TcpIpInformation
{
    public:
        // Default constructor
        TcpIpInformation();

        // Destructor
        ~TcpIpInformation();

        // Methods
        void   GetAdaptersRecords( TArray< AuditRecord >* pRecords ) const;
        void   GetMacAddress( String* pMacAddress ) const;
        void   GetRoutingTableRecords( TArray< AuditRecord >* pRecords ) const;

    protected:
        // Methods

        // Data members

    private:
        // See MIB_IPADDRROW Structure
        typedef struct _ADDRESS_INDEX
        {
            DWORD dwAddr;   // Interface IPv4 address
            DWORD dwIndex;  // Index of the interface
        } ADDRESS_INDEX;

        // Copy constructor - not allowed
        TcpIpInformation( const TcpIpInformation& oTcpIpInfo );

        // Assignment operator - not allowed
        TcpIpInformation& operator= ( const TcpIpInformation& oTcpIpInfo );

        // Methods
        void GetAdaptersRecords6( TArray< AuditRecord >* pRecords ) const;
        void GetAdaptersRecordsLegacy( TArray< AuditRecord >* pRecords ) const;
        void GetInterfaces( TArray< ADDRESS_INDEX >* pInterfaces ) const;
        void GetWmiAdapterData( const String& ServiceName,
                                DWORD*  pConfigManError,
                                String* pAdapterStatus,
                                String* pAdapterType,
                                String* pMacAddress,
                                String* pConnectionStatus, DWORD*  pLinkSpeedMbps ) const;
        void TranslateConfigManError( DWORD configManError, String* pError ) const;
        void TranslateProtocol( DWORD protocol, String* pTranslation ) const;
        void TranslateNetConnectionStatus( WORD status, String* pTranslation ) const;
        void TranslateRouteType( DWORD routeType, String* pTranslation ) const;

        // Data members
};

#endif  // WINAUDIT_TCP_IP_INFORMATION_H_
