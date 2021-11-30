///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Network TCP/IP Information Class Implementation
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
#include "WinAudit/Header Files/TcpIpInformation.h"

// 2. C System Files
#include <IPHlpApi.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/Wmi.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
TcpIpInformation::TcpIpInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
TcpIpInformation::~TcpIpInformation()
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
//     Get data about network adapters as an array of audit records
//
//  Parameters:
//      pRecords - array to receive records
//
//  Remarks:
//      For Windows Vista+ will use WMI only as can get stale data from the registry.
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::GetAdaptersRecords( TArray< AuditRecord >* pRecords ) const
{
    WindowsInformation WinInfo;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    if ( WinInfo.IsWinVistaOrNewer() )
    {
        GetAdaptersRecords6( pRecords );
    }
    else
    {
        GetAdaptersRecordsLegacy( pRecords );
    }
}

//===============================================================================================//
//  Description:
//      Get the MAC address of the machine
//
//  Parameters:
//      pMacAddress - string object to receive the MAC address
//
//  Remarks:
//      On machines with multiple adapters, the MAC address of the
//      first one is returned
//
//  Returns:
//      void, "" if none found
//===============================================================================================//
void TcpIpInformation::GetMacAddress( String* pMacAddress ) const
{
    DWORD   errorCode   = 0;
    ULONG   sizePointer = 0;
    wchar_t szMacAddress[ 32 ] = { 0 };   // Big enough for a MAC address
    AllocateBytes         AllocBytes;
    IP_ADAPTER_ADDRESSES* pAdapter = nullptr;
    IP_ADAPTER_ADDRESSES* pAdapterAddesses = nullptr;

    if ( pMacAddress == nullptr )
    {
        throw ParameterException( L"pMacAddress", __FUNCTION__ );
    }
    pMacAddress->Allocate( 32 );
    *pMacAddress = PXS_STRING_EMPTY;

    // First call to get the buffer size required
    errorCode = GetAdaptersAddresses( AF_UNSPEC,
                                      0, nullptr, nullptr, &sizePointer );
    if ( errorCode == ERROR_NO_DATA )
    {
        PXSLogAppInfo( L"GetAdaptersAddresses said no network adapters." );
        return;     // Nothing to do
    }

    // Test for expected error
    if ( errorCode != ERROR_BUFFER_OVERFLOW )
    {
        throw SystemException( errorCode,
                                 L"GetAdaptersAddresses(1)", __FUNCTION__ );
    }

    // Second call, allocate extra in case the number of adapers has changed
    sizePointer = PXSMultiplyUInt32( 2, sizePointer );
    pAdapterAddesses = reinterpret_cast<IP_ADAPTER_ADDRESSES*>(
                                                AllocBytes.New( sizePointer ) );
    errorCode = GetAdaptersAddresses( AF_UNSPEC,
                                      0,
                                      nullptr,
                                      pAdapterAddesses, &sizePointer );
    if ( errorCode != ERROR_SUCCESS )
    {
        throw SystemException( errorCode, L"GetAdaptersAddresses(2)", __FUNCTION__ );
    }

    pAdapter = pAdapterAddesses;
    while ( pAdapter && pMacAddress->IsEmpty() )
    {
        if ( pAdapter->PhysicalAddressLength == 6 )
        {
            // Will return lower case
            StringCchPrintf( szMacAddress,
                             ARRAYSIZE( szMacAddress ),
                             L"%02x:%02x:%02x:%02x:%02x:%02x",
                             pAdapter->PhysicalAddress[ 0 ],
                             pAdapter->PhysicalAddress[ 1 ],
                             pAdapter->PhysicalAddress[ 2 ],
                             pAdapter->PhysicalAddress[ 3 ],
                             pAdapter->PhysicalAddress[ 4 ],
                             pAdapter->PhysicalAddress[ 5 ] );

            *pMacAddress = szMacAddress;
        }
        pAdapter = pAdapter->Next;
    }
}

//===============================================================================================//
//  Description:
//      Get the IPv4 routing table
//
//  Parameters:
//      pRecords - array to receive records
//
//  Remarks:
//      Descriptions taken from MSDN
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::GetRoutingTableRecords( TArray< AuditRecord >* pRecords ) const
{
    ULONG  dwSize = 0;
    size_t i = 0, j = 0, numInterfaces = 0;
    time_t forwardAge = 0;
    String Value;
    Formatter     Format;
    AuditRecord   Record;
    AllocateBytes AllocBytes;
    MIB_IPFORWARDROW*   pIpForwardRow   = nullptr;
    MIB_IPFORWARDTABLE* pIpForwardTable = nullptr;
    TArray< ADDRESS_INDEX > Interfaces;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // MIB_IPFORWARDROW stores the interface index but want the address
    GetInterfaces( &Interfaces );
    numInterfaces = Interfaces.GetSize();

    // First call to get the required array size
    if ( GetIpForwardTable( nullptr, &dwSize, FALSE ) == 0 )
    {
        // Success with no buffer
        PXSLogAppInfo( L"GetIpForwardTable returned no data [NO_ERROR]." );
        return;
    }

    if ( dwSize == 0 )
    {
        PXSLogAppInfo( L"GetIpForwardTable returned no data [dwSize]." );
        return;
    }

    // Allocate a bit extra and call for a second time
    dwSize = PXSMultiplyUInt32( dwSize, 2 );
    pIpForwardTable = reinterpret_cast<MIB_IPFORWARDTABLE*>( AllocBytes.New( dwSize ) );
    if ( GetIpForwardTable( pIpForwardTable, &dwSize, TRUE ) )
    {
        throw SystemException( GetLastError(), L"GetIpForwardTable", __FUNCTION__ );
    }
    PXSLogAppInfo1( L"Found %%1 entries in the IPv4 routing table.",
                    Format.UInt32( pIpForwardTable->dwNumEntries ) );

    for ( i = 0; i < pIpForwardTable->dwNumEntries; i++ )
    {
        pIpForwardRow = pIpForwardTable->table + i;

        Record.Reset( PXS_CATEGORY_ROUTING_TABLE );
        Value = Format.AddressToIPv4( pIpForwardRow->dwForwardDest );
        Record.Add( PXS_ROUTING_TABLE_DESTINATION, Value );

        Value = Format.AddressToIPv4( pIpForwardRow->dwForwardMask );
        Record.Add( PXS_ROUTING_TABLE_NETMASK, Value );

        Value = Format.AddressToIPv4( pIpForwardRow->dwForwardNextHop );
        Record.Add( PXS_ROUTING_TABLE_NEXT_HOP, Value );

        // Match the interface by index
        Value = PXS_STRING_EMPTY;
        for ( j = 0; j < numInterfaces; j++ )
        {
            if ( pIpForwardRow->dwForwardIfIndex == Interfaces.Get( j ).dwIndex)
            {
                Value = Format.AddressToIPv4( Interfaces.Get( j ).dwAddr );
                break;
            }
        }
        Record.Add( PXS_ROUTING_TABLE_INTERFACE, Value );

        Value = PXS_STRING_EMPTY;
        TranslateRouteType( pIpForwardRow->dwForwardType, &Value );
        Record.Add( PXS_ROUTING_TABLE_ROUTE_TYPE, Value );

        Value = PXS_STRING_EMPTY;
        TranslateProtocol( pIpForwardRow->dwForwardProto, &Value );
        Record.Add( PXS_ROUTING_TABLE_PROTOCOL, Value );

        forwardAge = PXSCastUInt32ToTimeT( pIpForwardRow->dwForwardAge );
        Value = Format.SecondsToDDHHMM( forwardAge );
        Record.Add( PXS_ROUTING_TABLE_AGE, Value );

        Value = PXS_STRING_EMPTY;
        if ( pIpForwardRow->dwForwardMetric1 != (DWORD)-1 )
        {
            Value = Format.UInt32( pIpForwardRow->dwForwardMetric1 );
        }
        Record.Add( PXS_ROUTING_TABLE_METIC, Value );
        pRecords->Add( Record );
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
//     Get data about network adapters as an array of audit records on Windows Vista or newer.
//
//  Parameters:
//      pRecords - array to receive records
//
//  Remarks:
//      Using WMI only. The GUID member of Win32_NetworkAdapter is only available on Vista
//      and newer.
//
//      Docs say on XP and newer should use GetAdaptersAddresses, however
//      it does not provide all the data so will use WMI. The registry may have
//      stale/obsolete data so will avoid. Note, many values such require the adaper to be
//      connected otherwise WMI return s NULL values, e.g. static IP address. The IPEnabled
//      member of Win32_NetworkAdapterConfiguration can be used to determine if the adapter
//      is connected.
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::GetAdaptersRecords6( TArray< AuditRecord >* pRecords ) const
{
    bool        dchpEnabled   = false;
    WORD        netConnStatus = 0;
    DWORD       adapterNumber = 0;
    DWORD       linkSpeedMbps = 0, configManError = 0;
    Wmi         WMIAdapter, WMIConfig, WMILinkSpeed;
    String      Value, LocaleMbs, AdapterStatus, AdapterGuid, Query, Name;
    Formatter   Format;
    AuditRecord Record;
    LPCWSTR     STR_WMI_NULL = L"<null>";

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Locale strings
    PXSGetResourceString( PXS_IDS_1265_MBS, &LocaleMbs );

    WMIAdapter.Connect( L"root\\cimv2" );
    WMIAdapter.ExecQuery( L"SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID IS NOT NULL" );
    while ( WMIAdapter.Next() )
    {
        Record.Reset( PXS_CATEGORY_NET_ADAPTERS );
        adapterNumber++;
        Value = Format.UInt32( adapterNumber );
        Record.Add( PXS_NET_ADAPT_ITEM_NUMBER, Value );

        // Adapter name
        Value.Zero();
        WMIAdapter.Get( L"Description", &Value );
        Record.Add( PXS_NET_ADAPT_ADAPTER_NAME, Value );

        // Get the configuration. Match is on SettingID = GUID
        AdapterGuid.Zero();
        WMIAdapter.Get( L"GUID", &AdapterGuid );
        WMIConfig.Connect( L"root\\cimv2" );
        Query   = L"Select * from Win32_NetworkAdapterConfiguration WHERE SettingID=\"";
        Query  += AdapterGuid;
        Query  += L"\"";
        WMIConfig.ExecQuery( Query.c_str() );
        if ( WMIConfig.Next() )
        {
            // HostName
            Value.Zero();
            WMIConfig.Get( L"DNSHostName", &Value );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_DNS_HOSTNAME, Value );

            // NameServer
            Value.Zero();
            WMIConfig.Get( L"DNSDomain", &Value );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_DNS_SERVERS, Value );

            // IP Address
            Value.Zero();
            WMIConfig.Get( L"IPAddress", &Value );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_IP_ADDRESS, Value );

            // Subnet mask
            Value.Zero();
            WMIConfig.Get( L"IPSubnet", &Value );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_IP_SUBNET, Value );

            // Default Gateway
            Value.Zero();
            WMIConfig.Get( L"DefaultIPGateway", &Value );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_DEFAULT_GATEWAY, Value );

            // DHCP enabled
            dchpEnabled = false;
            WMIConfig.GetBool( L"DHCPEnabled", &dchpEnabled );
            if ( dchpEnabled )
            {
                Record.Add( PXS_NET_ADAPT_DHCP_ENABLED, PXS_STRING_ONE );
            }
            else
            {
                Record.Add( PXS_NET_ADAPT_DHCP_ENABLED, PXS_STRING_ZERO );
            }

            // DHCP Server - WMI returns an IP address for this, i.e. DhcpIPAddress
            Value.Zero();
            WMIConfig.Get( L"DhcpServer", &Value );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_DHCP_SERVER, Value );

            // DHCP Server IP Address, same as above. For registry TCP/IP configuration
            // this is DhcpIPAddress
            Value.Zero();
            WMIConfig.Get( L"DhcpServer", &Value );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_DHCP_IPADDRESS, Value );

            // Time Lease obtained, WMI datetime is stored as a BSTR
            Value.Zero();
            if ( dchpEnabled )
            {
                WMIConfig.Get( L"DHCPLeaseObtained", &Value );
            }
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_DHCP_LEASE_OBTAIN, Value );

            // Lease Expires, WMI datetime is stored as a BSTR
            Value.Zero();
            if ( dchpEnabled )
            {
                WMIConfig.Get( L"DHCPLeaseExpires", &Value );
            }
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_DHCP_LEASE_EXPIRES, Value );

            // Status Code and adapter status
            Value.Zero();
            AdapterStatus.Zero();
            configManError = 0;
            if ( WMIAdapter.GetUInt32( L"ConfigManagerErrorCode", &configManError ) )
            {
                Value = Format.UInt32( configManError );
                TranslateConfigManError( configManError, &AdapterStatus );
            }
            Record.Add( PXS_NET_ADAPT_STATUS_CODE   , Value );
            Record.Add( PXS_NET_ADAPT_ADAPTER_STATUS, AdapterStatus );

            // Adapter Type
            Value.Zero();
            WMIAdapter.Get( L"AdapterType", &Value );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_ADAPTER_TYPE, Value );

            // MAC Address
            Value.Zero();
            WMIAdapter.Get( L"MACAddress", &Value  );
            if ( Value.CompareI( STR_WMI_NULL ) == 0 )
            {
                Value.Zero();
            }
            Record.Add( PXS_NET_ADAPT_MAC_ADDRESS, Value );

            // Connection Status
            Value.Zero();
            netConnStatus = 0;
            if ( WMIAdapter.GetUInt16( L"NetConnectionStatus", &netConnStatus ) )
            {
                TranslateNetConnectionStatus( netConnStatus, &Value );
            }
            Record.Add( PXS_NET_ADAPT_CONN_STATUS, Value );

            // Connection Speed, only valid if connected = 0x2. Match is on Name = InstanceName
            Value.Zero();
            if ( netConnStatus == 0x2 )
            {
                Name.Zero();
                WMIAdapter.Get( L"Name", &Name );

                WMILinkSpeed.Connect( L"root\\wmi" );
                Query  = L"Select * from MSNdis_LinkSpeed WHERE InstanceName=\"";
                Query += Name;
                Query += L"\"";
                WMILinkSpeed.ExecQuery( Query.c_str() );
                if ( WMILinkSpeed.Next() )
                {
                    // Units for NdisLinkSpeed are 100bps
                    WMILinkSpeed.GetUInt32( L"NdisLinkSpeed", &linkSpeedMbps );
                    Value  = Format.UInt32( linkSpeedMbps / 10000 );
                    Value += LocaleMbs;
                }
                WMILinkSpeed.Disconnect();
            }
            Record.Add( PXS_NET_ADAPT_CONNECT_MBPS, Value );

            pRecords->Add( Record );
        }
        else
        {
            PXSLogAppInfo1( L"No WMI network adapter configuration found for GUID='%%1'.",
                            AdapterGuid );
        }
        WMIConfig.CloseQuery();
    }
    WMIAdapter.CloseQuery();

    PXSSortAuditRecords( pRecords, PXS_NET_ADAPT_ADAPTER_NAME );
}

//===============================================================================================//
//  Description:
//     Get data about network adapters as an array of audit records on Windows Vista or newer.
//
//  Parameters:
//      pRecords - array to receive records
//
//  Remarks:
//      This method might detect obsolete data in the registry.
//
//      Docs say on XP and newer should use GetAdaptersAddresses, however
//      it does not provide all the data so will use the registry anf WMI.
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::GetAdaptersRecordsLegacy( TArray< AuditRecord >* pRecords ) const
{
    DWORD       adapterNumber = 0, enableDHCP = 0, dword = 0;
    DWORD       linkSpeedMbps = 0, configManError = 0;
    size_t      i = 0, numKeys = 0;
    time_t      timeT = 0;
    String      Key, RegKey, Value, ServiceName, Description, LocaleMbs;
    String      AdapterStatus, AdapterType, MacAddress, ConnectionStatus;
    Registry    RegObject;
    Formatter   Format;
    StringArray Keys;
    AuditRecord Record;
    LPCWSTR NETWORK_CARDS_KEY = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\NetworkCards";
    LPCWSTR PARAMETERS_KEY    = L"SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters";
    LPCWSTR INTERFACES_KEY    = L"SYSTEM\\CurrentControlSet\\Services\\"
                                L"Tcpip\\Parameters\\Interfaces\\";

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Locale strings
    PXSGetResourceString( PXS_IDS_1265_MBS, &LocaleMbs );

    RegObject.Connect( HKEY_LOCAL_MACHINE );
    RegObject.GetSubKeyList( NETWORK_CARDS_KEY, &Keys );
    numKeys = Keys.GetSize();
    for ( i = 0; i < numKeys; i++ )
    {
        // Catch any errors and proceed to next adapter
        try
        {
            // Get the service name
            Key  = NETWORK_CARDS_KEY;
            Key += L"\\";
            Key += Keys.Get( i );
            RegObject.GetStringValue( Key.c_str(),
                                      L"ServiceName", &ServiceName );
            RegObject.GetStringValue( Key.c_str(),
                L"Description", &Description );
            if ( ServiceName.GetLength() > 0 )
            {
                Record.Reset( PXS_CATEGORY_NET_ADAPTERS );
                adapterNumber++;
                Value = Format.UInt32( adapterNumber );
                Record.Add( PXS_NET_ADAPT_ITEM_NUMBER, Value );

                // Adaptor name
                Value = Description;
                if ( Description.IsEmpty() )
                {
                    Value = ServiceName;
                }
                Record.Add( PXS_NET_ADAPT_ADAPTER_NAME, Value );

                // HostName
                RegKey = PARAMETERS_KEY;
                Value  = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"HostName", &Value);
                Record.Add( PXS_NET_ADAPT_DNS_HOSTNAME, Value );

                RegKey  = INTERFACES_KEY;
                RegKey += ServiceName;

                // NameServer
                Value = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"NameServer", &Value );
                Record.Add( PXS_NET_ADAPT_DNS_SERVERS, Value );

                // IP Address
                Value  = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"IPAddress", &Value );
                Record.Add( PXS_NET_ADAPT_IP_ADDRESS, Value );

                // Subnet mask, if DHCP is in use need to use DhcpSubnetMask.
                Value      = PXS_STRING_EMPTY;
                enableDHCP = 0;
                RegObject.GetDoubleWordValue( RegKey.c_str(), L"EnableDHCP", 0, &enableDHCP );
                if ( enableDHCP )
                {
                    RegObject.GetStringValue( RegKey.c_str(), L"DhcpSubnetMask", &Value );
                }
                else
                {
                    RegObject.GetStringValue( RegKey.c_str(), L"SubnetMask", &Value );
                }
                Record.Add( PXS_NET_ADAPT_IP_SUBNET, Value );

                // Default Gateway
                Value = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"DefaultGateway", &Value );
                Record.Add( PXS_NET_ADAPT_DEFAULT_GATEWAY, Value );

                // DHCP enabled
                if ( enableDHCP == 0 )
                {
                    Record.Add( PXS_NET_ADAPT_DHCP_ENABLED, PXS_STRING_ZERO );
                }
                else
                {
                    Record.Add( PXS_NET_ADAPT_DHCP_ENABLED, PXS_STRING_ONE );
                }


                // Dhcp Server
                Value = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"DhcpServer", &Value );
                Record.Add( PXS_NET_ADAPT_DHCP_SERVER, Value );

                // Dhcp Server IP Address
                Value = PXS_STRING_EMPTY;
                RegObject.GetStringValue( RegKey.c_str(), L"DhcpIPAddress", &Value );
                Record.Add( PXS_NET_ADAPT_DHCP_IPADDRESS, Value );

                // Time Lease obtained, Unix time_t format
                Value = PXS_STRING_EMPTY;
                if ( enableDHCP )
                {
                    dword  = 0;
                    RegObject.GetDoubleWordValue( RegKey.c_str(),
                                                  L"LeaseObtainedTime", 0, &dword );
                    if ( dword > 0 )
                    {
                        timeT = PXSCastUInt32ToTimeT( dword );
                        Value = Format.TimeTToLocalTimeInIso( timeT );
                    }
                }
                Record.Add( PXS_NET_ADAPT_DHCP_LEASE_OBTAIN, Value );

                // Lease Expires, Unix time_t format
                Value = PXS_STRING_EMPTY;
                if ( enableDHCP )
                {
                    dword = 0;
                    RegObject.GetDoubleWordValue( RegKey.c_str(),
                                                  L"LeaseTerminatesTime", 0, &dword );
                    if ( dword > 0 )
                    {
                        timeT = PXSCastUInt32ToTimeT( dword );
                        Value = Format.TimeTToLocalTimeInIso( timeT );
                    }
                }
                Record.Add( PXS_NET_ADAPT_DHCP_LEASE_EXPIRES, Value );

                // Get WMI Stuff, continue on error
                AdapterStatus = PXS_STRING_EMPTY;
                AdapterType   = PXS_STRING_EMPTY;
                MacAddress    = PXS_STRING_EMPTY;
                ConnectionStatus = PXS_STRING_EMPTY;
                linkSpeedMbps  = 0;
                try
                {
                    GetWmiAdapterData( ServiceName,
                                       &configManError,
                                       &AdapterStatus,
                                       &AdapterType,
                                       &MacAddress, &ConnectionStatus, &linkSpeedMbps );
                }
                catch ( const Exception& eWmi )
                {
                  // Log it but continue
                  PXSLogException( L"Error getting WMI adapter data.", eWmi, __FUNCTION__ );
                }

                // Status code 0xFFFFFFFF means unknown
                Value = PXS_STRING_EMPTY;
                if ( configManError != 0xFFFFFFFF )
                {
                    Value = Format.UInt32( configManError );
                }
                Record.Add( PXS_NET_ADAPT_STATUS_CODE   , Value );
                Record.Add( PXS_NET_ADAPT_ADAPTER_STATUS, AdapterStatus);
                Record.Add( PXS_NET_ADAPT_ADAPTER_TYPE  , AdapterType );
                Record.Add( PXS_NET_ADAPT_MAC_ADDRESS   , MacAddress );
                Record.Add( PXS_NET_ADAPT_CONN_STATUS   , ConnectionStatus );

                Value  = Format.UInt32( linkSpeedMbps );
                Value += LocaleMbs;
                Record.Add( PXS_NET_ADAPT_CONNECT_MBPS, Value );

                pRecords->Add( Record );
            }
            else
            {
               PXSLogAppWarn1(L"ServiceName key not found for '%%1'.", Key);
            }
        }
        catch ( const Exception& e )
        {
            PXSLogException( L"Error while getting TCP/IP data.", e, __FUNCTION__ );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_NET_ADAPT_ADAPTER_NAME );
}

//===============================================================================================//
//  Description:
//      Get network interfaces, see MIB_IPADDRROW Structure
//
//  Parameters:
//      pInterfaces - receives the address/index pairs
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::GetInterfaces(TArray< ADDRESS_INDEX >* pInterfaces) const
{
    ULONG     i = 0, dwSize = 0;
    Formatter Format;
    AllocateBytes    AllocBytes;
    ADDRESS_INDEX    AddressIndex = { 0, 0 };
    MIB_IPADDRROW*   pIpAddrRow   = nullptr;
    MIB_IPADDRTABLE* pIpAddrTable = nullptr;

    if ( pInterfaces == nullptr )
    {
        throw ParameterException( L"pInterfaces", __FUNCTION__ );
    }
    pInterfaces->RemoveAll();

    // First call to get the required array size
    if ( GetIpAddrTable( nullptr, &dwSize, FALSE ) == 0 )
    {
        // Success with no buffer
        PXSLogAppInfo( L"GetIpAddrTable returned no data [NO_ERROR]." );
        return;
    }

    if ( dwSize == 0 )
    {
        PXSLogAppInfo( L"GetIpAddrTable returned no data [dwSize]." );
        return;
    }

    // Allocate a bit extra and call for a second time
    dwSize = PXSMultiplyUInt32( dwSize, 2 );
    pIpAddrTable = (PMIB_IPADDRTABLE)AllocBytes.New( dwSize );
    if ( GetIpAddrTable( pIpAddrTable, &dwSize, TRUE ) )
    {
        throw SystemException( GetLastError(), L"pIpAddrTable", __FUNCTION__ );
    }
    PXSLogAppInfo1( L"GetIpAddrTable found %%1 entries",
                    Format.UInt32( pIpAddrTable->dwNumEntries ) );

    for ( i = 0; i < pIpAddrTable->dwNumEntries; i++ )
    {
        pIpAddrRow = pIpAddrTable->table + i;
        AddressIndex.dwAddr  = pIpAddrRow->dwAddr;
        AddressIndex.dwIndex = pIpAddrRow->dwIndex;
        pInterfaces->Add( AddressIndex );
    }
}

//===============================================================================================//
//  Description:
//      Get Network Adapter data from its service name
//
//  Parameters:
//      ServiceName       - string of service name
//      pConfigManError   - receives the configuration manager error
//      pAdapterStatus    - receives the adapter's status
//      pAdapterType      - receives the adapter type
//      pMacAddress       - receives the MAC Address
//      pConnectionStatus - receives the connection status
//      pLinkSpeedMbps    - receives the connection speed, units of Mbps
//
//  Remarks:
//      Noticed connection speed says 10.0Mbps if cable not connected
//      Connection Status is available from XP and above
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::GetWmiAdapterData( const String& ServiceName,
                                          DWORD*  pConfigManError,
                                          String* pAdapterStatus,
                                          String* pAdapterType,
                                          String* pMacAddress,
                                          String* pConnectionStatus, DWORD*  pLinkSpeedMbps ) const
{
    Wmi    WMI;
    int    index   = 0;
    WORD   status  = 0;
    String Query, Name;
    Formatter Format;

    *pConfigManError   = 0xFFFFFFFF;      // 0 means the device is OK
    *pAdapterStatus    = PXS_STRING_EMPTY;
    *pAdapterType      = PXS_STRING_EMPTY;
    *pMacAddress       = PXS_STRING_EMPTY;
    *pConnectionStatus = PXS_STRING_EMPTY;
    *pLinkSpeedMbps    = 0;
    if ( ServiceName.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Get the index of this adapter from Win32_NetworkAdapterConfiguration
    WMI.Connect( L"root\\cimv2" );
    Query   = L"Select * from Win32_NetworkAdapterConfiguration "
              L"WHERE SettingID=\"";
    Query  += ServiceName;
    Query  += L"\"";
    WMI.ExecQuery( Query.c_str() );
    if ( WMI.Next() )
    {
        if ( WMI.GetInt32( L"Index", &index ) == false )
        {
            PXSLogAppInfo1( L"Index for adapter '%%1' is NULL or EMPTY.", ServiceName );
            return;
        }
    }
    WMI.CloseQuery();

    // Get the adapter instance from Win32_NetworkAdapter
    Query   = L"Select * from Win32_NetworkAdapter WHERE Index=";
    Query  += Format.Int32( index );
    WMI.ExecQuery( Query.c_str() );
    if ( WMI.Next() )
    {
        if ( WMI.GetUInt32( L"ConfigManagerErrorCode", pConfigManError ) )
        {
            TranslateConfigManError( *pConfigManError, pAdapterStatus );
        }
        WMI.Get( L"AdapterType", pAdapterType );
        WMI.Get( L"MACAddress" , pMacAddress  );

        if ( WMI.GetUInt16( L"NetConnectionStatus", &status ) )
        {
            TranslateNetConnectionStatus( status, pConnectionStatus );
        }
        WMI.Get( L"Name", &Name  );
    }
    WMI.CloseQuery();

    // Get the linkspeed, only valid if connected = 0x2
    if ( status == 0x2 )
    {
        WMI.Connect( L"root\\wmi" );
        Query  = L"Select * from MSNdis_LinkSpeed WHERE InstanceName=\"";
        Query += Name;
        Query += L"\"";
        WMI.ExecQuery( Query.c_str() );
        if ( WMI.Next() )
        {
            // Units for NdisLinkSpeed are 100bps
            *pLinkSpeedMbps = ( *pLinkSpeedMbps / 10000 );
        }
        WMI.Disconnect();
    }
}

//===============================================================================================//
//  Description:
//      Translate a configuration management error code
//
//  Parameters:
//      configManError - the configuration manager error
//      pError         - receives the translation
//
//  Remarks:
//      See Win32_NetworkAdapter Class
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::TranslateConfigManError( DWORD configManError, String* pError ) const
{
    size_t    i = 0;
    Formatter Format;

    struct _ERROR
    {
        DWORD   errorCode;
        LPCWSTR pszError;
    } Errors[] = {
        {  0, L"This device is working properly."                        },
        {  1, L"This device is not configured correctly."                },
        {  2, L"Windows cannot load the driver for this device."         },
        {  3, L"The driver for this device might be corrupted, or your "
              L"system may be running low on memory or other resources." },
        {  4, L"This device is not working properly. One of its "
              L"drivers or your registry might be corrupted."            },
        {  5, L"The driver for this device needs a resource that "
              L"Windows cannot manage."                                  },
        {  6, L"The boot configuration for this device conflicts with "
              L"other devices."                                          },
        {  7, L"Cannot filter."                                          },
        {  8, L"The driver loader for the device is missing."            },
        {  9, L"This device is not working properly because the "
              L"controlling firmware is reporting the resources for "
              L"the device incorrectly."                                 },
        { 10, L"This device cannot start."                               },
        { 11, L"This device failed."                                     },
        { 12, L"This device cannot find enough free resources that "
              L"it can use."                                             },
        { 13, L"Windows cannot verify this device's resources."          },
        { 14, L"This device cannot work properly until you restart "
              L"your computer."                                          },
        { 15, L"This device is not working properly because there is "
              L"probably a re-enumeration problem."                      },
        { 16, L"Windows cannot identify all the resources this "
              L"device uses."                                            },
        { 17, L"This device is asking for an unknown resource type."     },
        { 18, L"Reinstall the drivers for this device."                  },
        { 19, L"Failure using the VXD loader."                           },
        { 20, L"Your registry might be corrupted."                       },
        { 21, L"System failure: Try changing the driver for this "
              L"device. If that does not work, see your hardware "
              L"documentation. Windows is removing this device."         },
        { 22, L"This device is disabled." },
        { 23, L"System failure: Try changing the driver for this device. "
              L"If that doesn't work, see your hardware documentation."  },
        { 24, L"This device is not present, is not working properly, or "
              L"does not have all its drivers installed."                },
        { 25, L"Windows is still setting up this device."                },
        { 26, L"Windows is still setting up this device."                },
        { 27, L"This device does not have valid log configuration."      },
        { 28, L"The drivers for this device are not installed."          },
        { 29, L"This device is disabled because the firmware of the "
              L"device did not give it the required resources."          },
        { 30, L"This device is using an Interrupt Request (IRQ) "
              L"resource that another device is using."                  },
        { 31, L"This device is not working properly because Windows "
              L"cannot load the drivers required for this device"        }
    };

    if ( pError == nullptr )
    {
        throw ParameterException( L"pError", __FUNCTION__ );
    }
    *pError = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Errors ); i++ )
    {
        if ( configManError == Errors[ i ] .errorCode )
        {
            *pError = Errors[ i ].pszError;
            break;
        }
    }

    if ( pError->IsEmpty() )
    {
        *pError = Format.UInt32( configManError );
        PXSLogAppWarn1( L"Unrecognised ConfigMan error: '%%1'.", *pError );
    }
}

//===============================================================================================//
//  Description:
//      Translate an IPv4 Protocol
//
//  Parameters:
//      protocol     - the protocol code
//      pTranslation - receives the translation
//
//  Remarks:
//      See MIB_IPFORWARDROW Structure
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::TranslateProtocol( DWORD protocol, String* pTranslation ) const
{
    size_t i = 0;
    Formatter Format;

    struct _PROTOCOLS
    {
        DWORD   protocol;
        LPCWSTR pszProtocol;
    } Protcols [] = {
        { MIB_IPPROTO_OTHER            , L"Other"                      },
        { MIB_IPPROTO_LOCAL            , L"Local"                      },
        { MIB_IPPROTO_NETMGMT          , L"Static"                     },
        { MIB_IPPROTO_ICMP             , L"ICMP"                       },
        { MIB_IPPROTO_EGP              , L"EGP"                        },
        { MIB_IPPROTO_GGP              , L"GGP"                        },
        { MIB_IPPROTO_HELLO            , L"Hellospeak"                 },
        { MIB_IPPROTO_RIP              , L"RIP or RIP-II"              },
        { MIB_IPPROTO_IS_IS            , L"IS-IS"                      },
        { MIB_IPPROTO_ES_IS            , L"ES-IS"                      },
        { MIB_IPPROTO_CISCO            , L"IGRP"                       },
        { MIB_IPPROTO_BBN              , L"BBN-IGP"                    },
        { MIB_IPPROTO_OSPF             , L"OSPF"                       },
        { MIB_IPPROTO_BGP              , L"BGP"                        },
        { MIB_IPPROTO_NT_AUTOSTATIC    , L"Windows Autostatic"         },
        { MIB_IPPROTO_NT_STATIC        , L"Windows Static"             },
        { MIB_IPPROTO_NT_STATIC_NON_DOD, L"Windows Static without DOD" } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Protcols ); i++ )
    {
        if ( protocol == Protcols[ i ].protocol )
        {
            *pTranslation = Protcols[ i ].pszProtocol;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt32( protocol );
        PXSLogAppWarn1( L"Unrecognised IPv4Protocal %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a configuration management error code
//
//  Parameters:
//      configManError - the configuration manager error
//      pTranslation   - receives the translation
//
//  Remarks:
//      See Win32_NetworkAdapter Class
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::TranslateNetConnectionStatus( WORD status, String* pTranslation ) const
{
    size_t   i = 0;
    Formatter Format;

    struct _STATUS
    {
        WORD    wStatus;
        LPCWSTR pszStatus;
    } aStatus[] =
        { {  0, L"Disconnected"             },
          {  1, L"Connecting"               },
          {  2, L"Connected"                },
          {  3, L"Disconnecting"            },
          {  4, L"Hardware not present"     },
          {  5, L"Hardware disabled"        },
          {  6, L"Hardware malfunction"     },
          {  7, L"Media disconnected"       },
          {  8, L"Authenticating"           },
          {  9, L"Authentication succeeded" },
          { 10, L"Authentication failed"    },
          { 11, L"Invalid address"          },
          { 12, L"Credentials required"     }
        };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( aStatus ); i++ )
    {
        if ( status == aStatus[ i ] .wStatus )
        {
            *pTranslation = aStatus[ i ].pszStatus;
            break;
        }
    }

    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.UInt16( status );
        PXSLogAppWarn1( L"Unrecognised connection status: '%%1'.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate a route type code
//
//  Parameters:
//      routeType    - the route type code
//      pTranslation - receives the translation
//
//  Remarks:
//      See MIB_IPFORWARDROW Structure
//
//  Returns:
//      void
//===============================================================================================//
void TcpIpInformation::TranslateRouteType( DWORD routeType, String* pTranslation ) const
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }

    switch ( routeType )
    {
        default:
            *pTranslation = Format.UInt32( routeType );
            PXSLogAppWarn1( L"Unrecognised routing table row type = %%1.", *pTranslation );
            break;

        case MIB_IPROUTE_TYPE_OTHER:
            // "Some other type not specified in RFC 1354."
            *pTranslation = L"Other";
            break;

        case MIB_IPROUTE_TYPE_INVALID:
            // An invalid route.
            *pTranslation = L"Invalid.";
            break;

        case MIB_IPROUTE_TYPE_DIRECT:
            // A local route where the next hop is the final destination
            *pTranslation = L"Local";
            break;

        case MIB_IPROUTE_TYPE_INDIRECT:
            // The remote route where the next hop is not the final destination.
            *pTranslation = L"Remote";
            break;
    }
}
