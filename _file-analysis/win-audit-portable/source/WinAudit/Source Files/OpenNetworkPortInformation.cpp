///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Open Network Ports Information Implementation
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
#include "WinAudit/Header Files/OpenNetworkPortInformation.h"

// 2. C System Files
#include <ws2ipdef.h>
#include <IPHlpApi.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/ProcessInformation.h"
#include "WinAudit/Header Files/WellKnowPortsMap.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
OpenNetworkPortInformation::OpenNetworkPortInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
OpenNetworkPortInformation::~OpenNetworkPortInformation()
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
//      Get open network ports as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//     void
//===============================================================================================//
void OpenNetworkPortInformation::GetAuditRecords(TArray<AuditRecord>* pRecords)
{
    void*     pResult = nullptr;
    size_t    i = 0, numPorts = 0;
    String    Caption, ServiceName, State, ProcessName;
    String    ManufacturerName, Description;
    String    VersionString, LocalPort, ProcessID;
    Formatter Format;
    FileVersion    FileVer;
    AuditRecord    Record;
    TYPE_PORT_INFO PortInfo;
    ProcessInformation     ProcessInfo;
    PXS_TYPE_PORT_SERVICE* pPortService = nullptr;
    TArray< TYPE_PORT_INFO > Ports;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Preference is to get ports along with their process identifiers
    try
    {
        GetPortsXPSP2v6( &Ports );
        ProcessInfo.FillProcessesInfo();
    }
    catch ( const Exception& e )
    {
        // Log it but continue
        PXSLogException( L"Failed to get open network ports.", e, __FUNCTION__ );
    }

    // If unsuccessful, fall back of Internet Helper Protocol, works on Win98
    // and above but does not have the process id information
    if ( Ports.GetSize() == 0 )
    {
        GetPortsIphlp( &Ports );
    }

    // Make records
    numPorts = Ports.GetSize();
    for ( i = 0; i < numPorts; i++ )
    {
        PortInfo = Ports.Get( i );
        if ( lstrlen( PortInfo.szLocalAddress ) )
        {
           LocalPort = Format.UInt32( PortInfo.localPort );
           Record.Reset( PXS_CATEGORY_OPEN_PORTS );
           Record.Add( PXS_OPEN_PORTS_PROTOCOL, PortInfo.szProtocol );
           Record.Add( PXS_OPEN_PORTS_LOCAL_ADDRESS, PortInfo.szLocalAddress);
           Record.Add( PXS_OPEN_PORTS_LOCAL_PORT, LocalPort );

            // Need a caption string
            Caption  = PortInfo.szProtocol;
            Caption += PXS_CHAR_SPACE;
            Caption += PortInfo.szLocalAddress;
            Caption += PXS_CHAR_COLON;
            Caption += LocalPort;
            Record.Add( PXS_OPEN_PORTS_CAPTION, Caption );

            // Service name
            ServiceName = PXS_STRING_EMPTY;
            pResult = bsearch( &PortInfo.localPort,
                               PXS_WELL_KNOWN_PORT_SERVICES,
                               ARRAYSIZE( PXS_WELL_KNOWN_PORT_SERVICES ),
                               sizeof ( PXS_WELL_KNOWN_PORT_SERVICES[ 0 ] ),
                               PXSBSearchComparePortServices );
            pPortService = reinterpret_cast<PXS_TYPE_PORT_SERVICE*>( pResult );
            if ( pPortService )
            {
                ServiceName = pPortService->pszServiceName;
            }
            Record.Add( PXS_OPEN_PORTS_SERVICE_NAME, ServiceName );

            // Only TCP has remote port information
            if ( CSTR_EQUAL == CompareString( LOCALE_INVARIANT,  // XP or newer
                                             NORM_IGNORECASE,
                                             PortInfo.szProtocol,
                                             -1,
                                             L"TCP",
                                              -1 ) )
            {
                Record.Add( PXS_OPEN_PORTS_REMOTE_ADDRESS, PortInfo.szRemoteAddress );
                Record.Add( PXS_OPEN_PORTS_REMOTE_PORT, Format.UInt32( PortInfo.remotePort ) );

                // Port state
                TranslatePortState( PortInfo.state, &State );
                Record.Add( PXS_OPEN_PORTS_CONNECTION_STATE, State );
            }

            ProcessName = PXS_STRING_EMPTY;
            ProcessInfo.GetExePath( PortInfo.owningPid, &ProcessName );
            Record.Add( PXS_OPEN_PORTS_PROCESS_NAME, ProcessName );

            ProcessID = Format.UInt32( PortInfo.owningPid );
            Record.Add( PXS_OPEN_PORTS_PROCESS_ID, ProcessID );

            ManufacturerName = PXS_STRING_EMPTY;
            Description  = PXS_STRING_EMPTY;
            try
            {
                if ( ProcessName.GetLength() )
                {
                    FileVer.GetVersion( ProcessName,
                                        &ManufacturerName, &VersionString, &Description );
                }
            }
            catch ( const Exception& eVersion )
            {
               // Log and continue to next port
               PXSLogException( L"Error getting a file version.", eVersion, __FUNCTION__ );
            }
            Record.Add( PXS_OPEN_PORTS_PROCESS_DESC, Description );
            Record.Add( PXS_OPEN_PORTS_PROCESS_MANUFAC, ManufacturerName );

            pRecords->Add( Record );
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get open network ports using the Internet Helper
//
//  Parameters:
//      pPorts - array to receive the data strings
//
//  Remarks:
//      Iphlpapi.dll is available on Windows 98 and newer
//
//  Returns:
//     void
//===============================================================================================//
void OpenNetworkPortInformation::GetPortsIphlp(TArray< TYPE_PORT_INFO >* pPorts)
{
    DWORD     i = 0, errorCode = 0, dwSize = 0;
    String    Address;
    Formatter Format;
    AllocateBytes  AllocTcpTable, AllocUdpTable;
    PMIB_TCPTABLE  pTcpTable = nullptr;
    PMIB_UDPTABLE  pUdpTable = nullptr;
    TYPE_PORT_INFO PortInfo;

    if ( pPorts == nullptr )
    {
        throw ParameterException( L"pPorts", __FUNCTION__ );
    }
    pPorts->RemoveAll();

    ////////////////////////////////////////////////////////////////////////////
    // TCP ports

    // First call to GetTcpTable to get the memory requirement
    errorCode = GetTcpTable( nullptr, &dwSize, TRUE );
    if ( errorCode == ERROR_INSUFFICIENT_BUFFER )
    {
        // As ports open and close regularly will allocate double the amount
        dwSize    = PXSMultiplyUInt32( dwSize, 2 );
        pTcpTable = reinterpret_cast<MIB_TCPTABLE*>(
                                                  AllocTcpTable.New( dwSize ) );
        errorCode = GetTcpTable( pTcpTable, &dwSize, TRUE );
        if ( errorCode == NO_ERROR )
        {
            PXSLogAppInfo1( L"GetTcpTable found %%1 entries",
                            Format.UInt32( pTcpTable->dwNumEntries ) );
            for ( i = 0; i < pTcpTable->dwNumEntries; i++ )
            {
                memset( &PortInfo, 0, sizeof ( PortInfo ) );
                PortInfo.state      = pTcpTable->table[ i ].dwState;
                PortInfo.localPort  = ReversePortNumber( pTcpTable->table[ i ].dwLocalPort );
                PortInfo.remotePort = ReversePortNumber( pTcpTable->table[ i ].dwRemotePort );
                StringCchCopy( PortInfo.szProtocol, ARRAYSIZE( PortInfo.szProtocol ), L"TCP" );

                // Local Address
                Address = Format.AddressToIPv4( pTcpTable->table[i].dwLocalAddr );
                StringCchCopy( PortInfo.szLocalAddress,
                               ARRAYSIZE( PortInfo.szLocalAddress ), Address.c_str() );

                // Remote Address
                Address = Format.AddressToIPv4( pTcpTable->table[ i ].dwRemoteAddr );
                StringCchCopy( PortInfo.szRemoteAddress,
                               ARRAYSIZE( PortInfo.szRemoteAddress ), Address.c_str() );

                pPorts->Add( PortInfo );
            }
        }
        else
        {
           // Log and continue
           PXSLogSysError( errorCode, L"GetTcpTable failed on second call." );
        }
    }
    else
    {
        // Unexpected but necessarily an error as there may be no ports
        PXSLogSysError( errorCode, L"GetTcpTable result on first call." );
    }

    ////////////////////////////////////////////////////////////////////////////
    // UDP ports

    // Get the memory needed to hold ports table
    dwSize = 0;
    errorCode = GetUdpTable( nullptr, &dwSize, TRUE );
    if ( errorCode == ERROR_INSUFFICIENT_BUFFER )
    {
        // As ports open and close regularly will allocate double the amount
        dwSize    = PXSMultiplyUInt32( dwSize, 2 );
        pUdpTable = reinterpret_cast<MIB_UDPTABLE*>( AllocUdpTable.New( dwSize ) );
        errorCode = GetUdpTable( pUdpTable, &dwSize, TRUE );
        if ( errorCode == NO_ERROR )
        {
            PXSLogAppInfo1( L"GetUdpTable found %%1 entries",
                            Format.UInt32( pUdpTable->dwNumEntries ) );
            for ( i = 0; i < pUdpTable->dwNumEntries; i++ )
            {
                memset( &PortInfo, 0, sizeof ( PortInfo ) );
                PortInfo.localPort = ReversePortNumber( pUdpTable->table[ i ].dwLocalPort );
                StringCchCopy( PortInfo.szProtocol, ARRAYSIZE( PortInfo.szProtocol ), L"UDP" );

                // Local Address
                Address = Format.AddressToIPv4( pUdpTable->table[ i ].dwLocalAddr );
                StringCchCopy( PortInfo.szLocalAddress,
                               ARRAYSIZE( PortInfo.szLocalAddress ),
                               Address.c_str() );

                pPorts->Add( PortInfo );
            }
        }
        else
        {
           // Log and continue
           PXSLogSysError( errorCode, L"GetUdpTable failed on second call." );
        }
    }
    else
    {
        PXSLogSysError( errorCode, L"GetUdpTable result on first call." );
    }
}

//===============================================================================================//
//  Description:
//      Get open network IPV6 ports using API available from XP + SP2
//
//  Parameters:
//      pPorts - array to receive port information
//
//  Remarks:
//      This method provides both the port's status and the process id of the
//      owning the port
//
//      The module that owns the port cannot be easily determined because the
//      member pOwningModuleInfo of MIB_TCPROW_OWNER_MODULE does not seem to
//      be documented. It is an array of DWORDS.
//
//  Returns:
//     void
//
//===============================================================================================//
void OpenNetworkPortInformation::GetPortsXPSP2v6( TArray< TYPE_PORT_INFO >* pPorts )
{
    DWORD     i = 0, dwSize = 0, errorCode;
    String    Address;
    Formatter Format;
    AllocateBytes  AllocTcp, AllocUdp;
    TYPE_PORT_INFO PortInfo;
    MIB_TCP6TABLE_OWNER_PID* pTcp6Table = nullptr;
    MIB_UDP6TABLE_OWNER_PID* pUdp6Table = nullptr;

    if ( pPorts == nullptr )
    {
        throw ParameterException( L"pPorts", __FUNCTION__ );
    }
    pPorts->RemoveAll();

    ////////////////////////////////////////////////////////////////////////////
    // TCP ports

    // Get the memory needed to hold the TCP ports table
    errorCode = GetExtendedTcpTable( nullptr,
                                     &dwSize, TRUE, AF_INET6, TCP_TABLE_OWNER_PID_ALL, 0 );

    if ( errorCode == ERROR_INSUFFICIENT_BUFFER )
    {
        // As ports open and close regularly will allocate double the amount
        dwSize     = PXSMultiplyUInt32( dwSize, 2 );
        pTcp6Table = reinterpret_cast<MIB_TCP6TABLE_OWNER_PID*>( AllocTcp.New( dwSize ) );
        errorCode  = GetExtendedTcpTable( pTcp6Table,
                                          &dwSize, TRUE, AF_INET6, TCP_TABLE_OWNER_PID_ALL, 0 );
        if ( errorCode == NO_ERROR )
        {
            PXSLogAppInfo1( L"GetExtendedTcpTable found %%1 entries",
                            Format.UInt32( pTcp6Table->dwNumEntries ) );
            for ( i = 0; i < pTcp6Table->dwNumEntries; i++ )
            {
                memset( &PortInfo, 0, sizeof ( PortInfo ) );
                StringCchCopy( PortInfo.szProtocol, ARRAYSIZE( PortInfo.szProtocol ), L"TCP" );
                PortInfo.owningPid = pTcp6Table->table[ i ].dwOwningPid;
                PortInfo.state     = pTcp6Table->table[ i ].dwState;
                PortInfo.localPort = ReversePortNumber( pTcp6Table->table[ i ].dwLocalPort );

                // Only sensible if have a port and not listening
                if ( pTcp6Table->table[ i ].dwState != MIB_TCP_STATE_LISTEN )
                {
                    PortInfo.remotePort = ReversePortNumber( pTcp6Table->table[ i ].dwRemotePort );
                }

                // Local Address
                Address = Format.AddressToIPv6( pTcp6Table->table[ i ].ucLocalAddr,
                                                ARRAYSIZE( pTcp6Table->table[ i ].ucLocalAddr ) );
                StringCchCopy( PortInfo.szLocalAddress,
                               ARRAYSIZE( PortInfo.szLocalAddress ), Address.c_str() );

                // Remote Address
                Address = Format.AddressToIPv6( pTcp6Table->table[ i ].ucRemoteAddr,
                                                ARRAYSIZE( pTcp6Table->table[ i ].ucRemoteAddr ) );
                StringCchCopy( PortInfo.szRemoteAddress,
                               ARRAYSIZE( PortInfo.szRemoteAddress ), Address.c_str() );

                pPorts->Add( PortInfo );
            }
        }
        else
        {
           PXSLogSysError( errorCode, L"GetExtendedTcpTable failed on second call." );
        }
    }
    else
    {
        PXSLogSysInfo( errorCode, L"GetExtendedTcpTable failed on first call." );
    }

    ////////////////////////////////////////////////////////////////////////////
    // UDP ports

    // Get the memory needed to hold the TCP ports table
    dwSize  = 0;
    errorCode = GetExtendedUdpTable( nullptr, &dwSize, TRUE, AF_INET6, UDP_TABLE_OWNER_PID, 0 );

    if ( errorCode == ERROR_INSUFFICIENT_BUFFER )
    {
        // As ports open and close regularly will allocate double the amount
        dwSize     = PXSMultiplyUInt32( dwSize, 2 );
        pUdp6Table = reinterpret_cast<MIB_UDP6TABLE_OWNER_PID*>(
                                                       AllocUdp.New( dwSize ) );
        errorCode = GetExtendedUdpTable( pUdp6Table,
                                         &dwSize, TRUE, AF_INET6, UDP_TABLE_OWNER_PID, 0 );
        if ( errorCode == NO_ERROR )
        {
            PXSLogAppInfo1( L"GetExtendedUdpTable found %%1 entries",
                            Format.UInt32( pUdp6Table->dwNumEntries ) );
            for ( i = 0; i < pUdp6Table->dwNumEntries; i++ )
            {
                memset( &PortInfo, 0, sizeof ( PortInfo ) );
                StringCchCopy( PortInfo.szProtocol, ARRAYSIZE( PortInfo.szProtocol ), L"UDP" );
                PortInfo.owningPid = pUdp6Table->table[ i ].dwOwningPid;
                PortInfo.localPort = ReversePortNumber( pUdp6Table->table[ i ].dwLocalPort );

                // Local Address
                Address = Format.AddressToIPv6( pUdp6Table->table[ i ].ucLocalAddr,
                                                ARRAYSIZE( pUdp6Table->table[ i ].ucLocalAddr ) );
                StringCchCopy( PortInfo.szLocalAddress,
                               ARRAYSIZE( PortInfo.szLocalAddress ), Address.c_str() );
                pPorts->Add( PortInfo );
            }
        }
        else
        {
            PXSLogSysError( errorCode, L"GetExtendedUdpTable failed on second call." );
        }
    }
    else
    {
        PXSLogSysInfo( errorCode, L"GetExtendedUdpTable result on first call." );
    }
}

//===============================================================================================//
//  Description:
//      Get open network IPV4 ports using API available from XP + SP2
//
//  Parameters:
//      pPorts - array to receive port information
//
//  Remarks:
//      This method provides both the port's status and the process id of the
//      owning the port
//
//      The module that owns the port cannot be easily determined because the
//      member pOwningModuleInfo of MIB_TCPROW_OWNER_MODULE does not seem to
//      be documented. It is an array of DWORDS.
//
//  Returns:
//     void
//===============================================================================================//
void OpenNetworkPortInformation::GetPortsXPSP2( TArray< TYPE_PORT_INFO >* pPorts )
{
    DWORD     i = 0, dwSize = 0, errorCode;
    String    Address;
    Formatter Format;
    AllocateBytes  AllocTcp, AllocUdp;
    TYPE_PORT_INFO PortInfo;
    MIB_TCPTABLE_OWNER_PID* pTcpTable = nullptr;
    MIB_UDPTABLE_OWNER_PID* pUdpTable = nullptr;

    if ( pPorts == nullptr )
    {
        throw ParameterException( L"pPorts", __FUNCTION__ );
    }
    pPorts->RemoveAll();

    ////////////////////////////////////////////////////////////////////////////
    // TCP ports

    // Get the memory needed to hold the TCP ports table
    errorCode = GetExtendedTcpTable( nullptr,
                                     &dwSize, TRUE, AF_INET, TCP_TABLE_OWNER_PID_ALL, 0 );

    if ( errorCode == ERROR_INSUFFICIENT_BUFFER )
    {
        // As ports open and close regularly will allocate double the amount
        dwSize    = PXSMultiplyUInt32( dwSize, 2 );
        pTcpTable = reinterpret_cast<MIB_TCPTABLE_OWNER_PID*>( AllocTcp.New( dwSize ) );
        errorCode = GetExtendedTcpTable( pTcpTable,
                                         &dwSize, TRUE, AF_INET, TCP_TABLE_OWNER_PID_ALL, 0 );
        if ( errorCode == NO_ERROR )
        {
            PXSLogAppInfo1( L"GetExtendedTcpTable found %%1 entries",
                            Format.UInt32( pTcpTable->dwNumEntries ) );
            for ( i = 0; i < pTcpTable->dwNumEntries; i++ )
            {
                memset( &PortInfo, 0, sizeof ( PortInfo ) );
                StringCchCopy( PortInfo.szProtocol, ARRAYSIZE( PortInfo.szProtocol ), L"TCP" );
                PortInfo.owningPid = pTcpTable->table[ i ].dwOwningPid;
                PortInfo.state     = pTcpTable->table[ i ].dwState;
                PortInfo.localPort = ReversePortNumber( pTcpTable->table[ i ].dwLocalPort );

                // Only sensible if have a port and not listening
                if ( ( pTcpTable->table[ i ].dwRemoteAddr ) &&
                     ( pTcpTable->table[ i ].dwState != MIB_TCP_STATE_LISTEN ) )
                {
                    PortInfo.remotePort = ReversePortNumber( pTcpTable->table[ i ].dwRemotePort );
                }

                // Local Address
                Address = Format.AddressToIPv4( pTcpTable->table[ i ].dwLocalAddr );
                StringCchCopy( PortInfo.szLocalAddress,
                               ARRAYSIZE( PortInfo.szLocalAddress ), Address.c_str() );

                // Remote Address
                Address = Format.AddressToIPv4( pTcpTable->table[ i ].dwRemoteAddr );
                StringCchCopy( PortInfo.szRemoteAddress,
                               ARRAYSIZE( PortInfo.szRemoteAddress ), Address.c_str() );

                pPorts->Add( PortInfo );
            }
        }
        else
        {
           PXSLogSysError( errorCode, L"GetExtendedTcpTable failed on second call." );
        }
    }
    else
    {
        PXSLogSysInfo( errorCode, L"GetExtendedTcpTable failed on first call." );
    }

    ////////////////////////////////////////////////////////////////////////////
    // UDP ports

    // Get the memory needed to hold the TCP ports table
    dwSize  = 0;
    errorCode = GetExtendedUdpTable( nullptr, &dwSize, TRUE, AF_INET, UDP_TABLE_OWNER_PID, 0 );

    if ( errorCode == ERROR_INSUFFICIENT_BUFFER )
    {
        // As ports open and close regularly will allocate double the amount
        dwSize    = PXSMultiplyUInt32( dwSize, 2 );
        pUdpTable = reinterpret_cast<MIB_UDPTABLE_OWNER_PID*>(
                                                       AllocUdp.New( dwSize ) );
        errorCode = GetExtendedUdpTable( pUdpTable,
                                         &dwSize, TRUE, AF_INET, UDP_TABLE_OWNER_PID, 0 );
        if ( errorCode == NO_ERROR )
        {
            PXSLogAppInfo1( L"GetExtendedUdpTable found %%1 entries",
                            Format.UInt32( pUdpTable->dwNumEntries ) );
            for ( i = 0; i < pUdpTable->dwNumEntries; i++ )
            {
                memset( &PortInfo, 0, sizeof ( PortInfo ) );
                StringCchCopy( PortInfo.szProtocol, ARRAYSIZE( PortInfo.szProtocol ), L"UDP" );
                PortInfo.owningPid = pUdpTable->table[ i ].dwOwningPid;
                PortInfo.localPort = ReversePortNumber( pUdpTable->table[ i ].dwLocalPort );

                // Local Address
                Address = Format.AddressToIPv4( pUdpTable->table[ i ].dwLocalAddr );
                StringCchCopy( PortInfo.szLocalAddress,
                               ARRAYSIZE( PortInfo.szLocalAddress ), Address.c_str() );
                pPorts->Add( PortInfo );
            }
        }
        else
        {
            PXSLogSysError( errorCode, L"GetExtendedUdpTable failed on second call." );
        }
    }
    else
    {
        PXSLogSysInfo( errorCode, L"GetExtendedUdpTable result on first call." );
    }
}

//===============================================================================================//
//  Description:
//      Reverse the bytes of a port number
//
//  Parameters:
//      portNumber - The value
//
//  Remarks:
//      Port numbers are network byte order
//
//  Returns:
//        DWORD
//===============================================================================================//
DWORD OpenNetworkPortInformation::ReversePortNumber( DWORD portNumber )
{
    // Only need to reverse the lo-word
    return ( ( portNumber & 0xFF00 ) >> 8 ) +
           ( ( portNumber & 0x00FF ) << 8 );
}

//===============================================================================================//
//  Description:
//      Translate a port state to a string
//
//  Parameters:
//      portState - the port state
//      pState    - string to receive state description
//
//  Returns:
//      void
//===============================================================================================//
void OpenNetworkPortInformation::TranslatePortState( DWORD portState, String* pState )
{
    size_t    i = 0;
    Formatter Format;

    struct _PORT_STATES
    {
        DWORD   portState;
        LPCWSTR pszState;
    } States[] =
      { { MIB_TCP_STATE_CLOSED   ,  L"Closed (CLOSED)"                     },
        { MIB_TCP_STATE_LISTEN   ,  L"Listening (LISTEN)"                  },
        { MIB_TCP_STATE_SYN_SENT ,  L"Connecting (SYN-SENT)"               },
        { MIB_TCP_STATE_SYN_RCVD ,  L"Connecting (SYN-RECEIVED)"           },
        { MIB_TCP_STATE_ESTAB    ,  L"Connection established (ESTABLISHED)"},
        { MIB_TCP_STATE_FIN_WAIT1,  L"Finalizing (FIN-WAIT-1)"             },
        { MIB_TCP_STATE_FIN_WAIT2,  L"Finalizing (FIN-WAIT-2)"             },
        { MIB_TCP_STATE_CLOSE_WAIT, L"Closing (CLOSE-WAIT)",               },
        { MIB_TCP_STATE_CLOSING  ,  L"Closing (CLOSING)",                  },
        { MIB_TCP_STATE_LAST_ACK ,  L"Closing (LAST-ACK)",                 },
        { MIB_TCP_STATE_TIME_WAIT,  L"Closing (TIME-WAIT)"                 },
        { MIB_TCP_STATE_DELETE_TCB, L"Deleted TCB"                         }
      };

    if ( pState == nullptr )
    {
        throw ParameterException( L"pState", __FUNCTION__ );
    }
    *pState = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( States ); i++ )
    {
        if ( portState == States[ i ].portState )
        {
            *pState = States[ i ].pszState;
            break;
        }
    }

    // If got nothing use the numerical value
    if ( pState->IsEmpty() )
    {
        *pState = Format.UInt32( portState );
        PXSLogAppWarn1( L"Unrecognised port state = %%1.", *pState );
    }
}
