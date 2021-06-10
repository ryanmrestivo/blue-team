///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Open Network Port Information Class Header
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

#ifndef WINAUDIT_OPEN_NETWORK_PORT_INFORMATION_H_
#define WINAUDIT_OPEN_NETWORK_PORT_INFORMATION_H_

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

class OpenNetworkPortInformation
{
    public:
        // Default constructor
        OpenNetworkPortInformation();

        // Destructor
        ~OpenNetworkPortInformation();

        // Methods
        void    GetAuditRecords( TArray< AuditRecord >* pRecords );

    protected:
        // Methods

        // Data members

    private:
        // Structure to hold port information
        typedef struct _TYPE_PORT_INFO
        {
            DWORD   owningPid;                  // Process ID
            DWORD   state;                      // MIB_ status code
            DWORD   localPort;                  // Port Number
            DWORD   remotePort;                 // Port Number
            wchar_t szProtocol[ 4 ];            // TCP or UDP
            wchar_t szLocalAddress[ 64 ];       // e.g. 3FFE:FFFF:7654:FEDA:1245:BA98:3210:4562
            wchar_t szRemoteAddress[ 64 ];      //
        } TYPE_PORT_INFO;

        // Copy constructor - not allowed
        OpenNetworkPortInformation( const OpenNetworkPortInformation& oInformation );

        // Assignment operator - not allowed
        OpenNetworkPortInformation& operator= ( const OpenNetworkPortInformation& oInformation );

        // Methods
        void    GetPortsIphlp( TArray< TYPE_PORT_INFO >* pPorts );
        void    GetPortsXPSP2( TArray< TYPE_PORT_INFO >* pPorts );
        void    GetPortsXPSP2v6( TArray< TYPE_PORT_INFO >* pPorts );
        DWORD   ReversePortNumber( DWORD portNumber );
        void    TranslatePortState( DWORD portState, String* pState );

    // Data members
};

#endif  // WINAUDIT_OPEN_NETWORK_PORT_INFORMATION_H_
