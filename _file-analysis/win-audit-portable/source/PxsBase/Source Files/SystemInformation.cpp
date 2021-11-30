///////////////////////////////////////////////////////////////////////////////////////////////////
//
// System Information Class Implementation
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
#include "PxsBase/Header Files/SystemInformation.h"

// 2. C System Files
#include <DsGetDC.h>
#include <malloc.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/AutoCloseHandle.h"
#include "PxsBase/Header Files/AutoNetApiBufferFree.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/Library.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/Wmi.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// WinNT.h constants after SDK 7.1
///////////////////////////////////////////////////////////////////////////////////////////////////

#ifndef PRODUCT_CORE
    #define PRODUCT_CORE                                0x00000065 // Windows 10 Home
#endif

#ifndef PRODUCT_CORE_COUNTRYSPECIFIC
    #define PRODUCT_CORE_COUNTRYSPECIFIC                0x00000063 // Windows 10 Home China
#endif

#ifndef PRODUCT_CORE_N
    #define PRODUCT_CORE_N                              0x00000062 // Windows 10 Home N
#endif

#ifndef PRODUCT_CORE_SINGLELANGUAGE
    #define PRODUCT_CORE_SINGLELANGUAGE                 0x00000064 // Windows 10 Home Single
#endif                                                             // Language

#ifndef PRODUCT_DATACENTER_EVALUATION_SERVER
    #define PRODUCT_DATACENTER_EVALUATION_SERVER        0x00000050 // Server Datacenter
#endif                                                             // (evaluation installation)

#ifndef PRODUCT_ENTERPRISE_EVALUATION
    #define PRODUCT_ENTERPRISE_EVALUATION               0x00000048 // Windows 10 Enterprise
#endif                                                             // Evaluation

#ifndef PRODUCT_ENTERPRISE_N_EVALUATION
    #define PRODUCT_ENTERPRISE_N_EVALUATION             0x00000054 // Windows 10 Enterprise N
#endif                                                             // Evaluation

#ifndef PRODUCT_MOBILE_CORE
    #define PRODUCT_MOBILE_CORE                         0x00000068 // Windows 10 Mobile
#endif

#ifndef PRODUCT_MULTIPOINT_PREMIUM_SERVER
    #define PRODUCT_MULTIPOINT_PREMIUM_SERVER           0x0000004D // Windows MultiPoint Server
#endif                                                             // Premium (full installation)

#ifndef PRODUCT_MULTIPOINT_STANDARD_SERVER
    #define PRODUCT_MULTIPOINT_STANDARD_SERVER          0x0000004C // Windows MultiPoint Server
#endif                                                             // Standard (full installation)

#ifndef PRODUCT_PROFESSIONAL_WMC
    #define PRODUCT_PROFESSIONAL_WMC                    0x00000067 // Professional with Media
#endif                                                             // Center

#ifndef PRODUCT_STANDARD_EVALUATION_SERVER
    #define PRODUCT_STANDARD_EVALUATION_SERVER          0x0000004F // Server Standard
#endif                                                             // (evaluation installation)

#ifndef PRODUCT_STORAGE_STANDARD_EVALUATION_SERVER
    #define PRODUCT_STORAGE_STANDARD_EVALUATION_SERVER  0x00000060 // Storage Server Standard
#endif                                                             // (evaluation installation)

#ifndef PRODUCT_STORAGE_WORKGROUP_EVALUATION_SERVER
    #define PRODUCT_STORAGE_WORKGROUP_EVALUATION_SERVER 0x0000005F // Storage Server Workgroup
#endif                                                             // (evaluation installation)

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
SystemInformation::SystemInformation()
{
}

// Copy constructor
SystemInformation::SystemInformation( const SystemInformation& oSystemInfo )
{
    *this = oSystemInfo;
}

// Destructor
SystemInformation::~SystemInformation()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
SystemInformation& SystemInformation::operator= ( const SystemInformation& oSystemInfo )
{
    if ( this == &oSystemInfo ) return *this;

    // Nowt!

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Do a heap check
//
//  Parameters:
//      pHeapCheck - string to receive a description of the heap check
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::DoHeapCheck( String* pHeapCheck )
{
    int       heapStatus = 0;
    Formatter Format;

    if ( pHeapCheck == nullptr )
    {
        throw ParameterException( L"pHeapCheck", __FUNCTION__ );
    }
    *pHeapCheck = PXS_STRING_EMPTY;

    heapStatus = _heapchk();
    switch( heapStatus )
    {
        default:
            *pHeapCheck  = L"Code ";
            *pHeapCheck += Format.Int32( heapStatus );
            break;

        case _HEAPBADBEGIN:
            *pHeapCheck = L"Initial header information is bad.";
            break;

        case _HEAPBADNODE:
           *pHeapCheck = L"Bad node has been found or heap is damaged.";
           break;

        case _HEAPBADPTR:
            *pHeapCheck = L"Pointer into heap is not valid.";
            break;

        case _HEAPEMPTY:
            *pHeapCheck = L"Heap has not been initialized.";
            break;

        case _HEAPOK:
            *pHeapCheck = L"Heap appears to be consistent.";
            break;
    }
}

//===============================================================================================//
//  Description:
//      Get the build number of the operating system
//
//  Parameters:
//      None
//
//  Returns:
//      The build number or 0 if unknown or failure
//===============================================================================================//
DWORD SystemInformation::GetBuildNumber() const
{
    OSVERSIONINFOEX VersionInfoEx;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );

    return VersionInfoEx.dwBuildNumber;
}

//===============================================================================================//
//  Description:
//      Get the account name of specified relative identifier for the
//      local computer
//
//  Parameters:
//      RID          - the RID, e.g. DOMAIN_USER_RID_ADMIN
//      pAccountName - receives the account name
//
//  Remarks:
//      See MSDN "The Guts of Security" by Ruediger R. Asche
//      See See Q157234
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetBuiltInAccountNameFromRID( DWORD RID, String* pAccountName ) const
{
    BYTE    i = 0, nDomainSubAuthority = 0, nSubAuthorityCount = 0;
    BYTE    bSid[ SECURITY_MAX_SID_SIZE ] = { 0 };
    DWORD   cchName = 0, cchDomainName = 0;
    PSID    pDomainSid    = nullptr;
    PSID    pAccountSid   = nullptr;
    wchar_t szName[ 256 ] = { 0 };          // Enough, only need UNLEN
    wchar_t szDomainName[ 256 ]  = { 0 };   // Enough, only need DNLEN
    UCHAR*  pUChar        = nullptr;
    DWORD*  pDomainSubAuthority  = nullptr;
    DWORD*  pAccountSubAuthority = nullptr;
    SID_NAME_USE    eUse;
    NET_API_STATUS  status   = 0;
    USER_MODALS_INFO_2* pUserModalsInfo2 = nullptr;

    if ( pAccountName == nullptr )
    {
        throw ParameterException( L"pAccountName", __FUNCTION__ );
    }
    *pAccountName = PXS_STRING_EMPTY;

    // Get the domain ID, level 2
    status = NetUserModalsGet( nullptr, 2, reinterpret_cast<LPBYTE*>( &pUserModalsInfo2 ) );
    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetUserModalsGet", __FUNCTION__ );
    }

    // Verify have a buffer, regardless of status
    if ( pUserModalsInfo2 == nullptr )
    {
        throw NullException( L"pUserModalsInfo2", __FUNCTION__ );
    }
    AutoNetApiBufferFree AutoFreeUserModalsInfo2( pUserModalsInfo2 );

    // Want a non-null SID for IsValidSid
    pDomainSid = pUserModalsInfo2->usrmod2_domain_id;
    if ( pDomainSid == nullptr )
    {
        throw NullException( L"pDomainSid", __FUNCTION__ );
    }

    if ( IsValidSid( pDomainSid ) == 0 )
    {
        throw SystemException( ERROR_INVALID_SID, L"IsValidSid",  __FUNCTION__ );
    }

    // Get the subauthority count
    pUChar = GetSidSubAuthorityCount( pDomainSid );
    if ( pUChar == nullptr )
    {
        throw SystemException( GetLastError(), L"GetSidSubAuthorityCount", __FUNCTION__ );
    }
    nDomainSubAuthority = *pUChar;

    nSubAuthorityCount = (BYTE)( 0xff & ( nDomainSubAuthority + 1 ) );
    pAccountSid = (PSID)bSid;
    if ( InitializeSid( pAccountSid,
                        GetSidIdentifierAuthority( pDomainSid ), nSubAuthorityCount ) == 0 )
    {
        throw SystemException( GetLastError(), L"InitializeSid", __FUNCTION__ );
    }

    // Copy the subauthorities from the domain SID to the account SID
    // then append the RID
    for ( i = 0; i < nDomainSubAuthority; i++)
    {
        pDomainSubAuthority   = GetSidSubAuthority( pDomainSid, i );
        pAccountSubAuthority  = GetSidSubAuthority( pAccountSid, i );
        *pAccountSubAuthority = *pDomainSubAuthority;
    }
    *GetSidSubAuthority( pAccountSid, nDomainSubAuthority ) = RID;

    // Get the account name
    cchName = ARRAYSIZE( szName );
    cchDomainName  = ARRAYSIZE( szDomainName );
    if ( LookupAccountSid( nullptr,     // Local computer
                           pAccountSid,
                           szName, &cchName, szDomainName, &cchDomainName, &eUse ) == 0 )
    {
        throw SystemException( GetLastError(), L"LookupAccountSid", __FUNCTION__ );
    }
    szName[ ARRAYSIZE( szName ) - 1 ] = PXS_CHAR_NULL;
    *pAccountName = szName;
}

//===============================================================================================//
//  Description:
//      Get the description of the local computer
//
//  Parameters:
//      pDescription - string to receive the description
//
//  Remarks
//      Not all computers have descriptions so "" is valid
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetComputerDescription( String* pDescription ) const
{
    NET_API_STATUS   status = 0;
    SERVER_INFO_101* pServerInfo101 = nullptr;

    if ( pDescription == nullptr )
    {
        throw ParameterException( L"pDescription", __FUNCTION__ );
    }
    *pDescription = PXS_STRING_EMPTY;

    status = NetServerGetInfo( nullptr, 101, reinterpret_cast<LPBYTE*>( &pServerInfo101 ) );
    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetServerGetInfo", __FUNCTION__ );
    }

    if ( pServerInfo101 == nullptr )
    {
        throw NullException( L"pServerInfo101", __FUNCTION__ );
    }
    AutoNetApiBufferFree AutoFreeServerInfo101( pServerInfo101 );
    *pDescription = pServerInfo101->sv101_comment;
    pDescription->Trim();
}

//===============================================================================================//
//  Description:
//      Get the name of the local computer
//
//  Parameters:
//      pNetBiosName - string to receive name
//
//    Remarks
//      Not all computers have names so "" is valid
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetComputerNetBiosName( String* pNetBiosName ) const
{
    DWORD   nSize;
    wchar_t szBuffer[ MAX_COMPUTERNAME_LENGTH + 1 ] = { 0 };

    if ( pNetBiosName == nullptr )
    {
        throw ParameterException( L"pNetBiosName", __FUNCTION__ );
    }
    *pNetBiosName = PXS_STRING_EMPTY;

    nSize = ARRAYSIZE( szBuffer );
    if ( GetComputerName( szBuffer, &nSize ) )
    {
        szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = PXS_CHAR_NULL;
        *pNetBiosName = szBuffer;
        pNetBiosName->Trim();
    }
    else
    {
        PXSLogSysWarn( GetLastError(), L"GetComputerName failed." );
    }
}

//===============================================================================================//
//  Description:
//      Get this computers role
//
//  Parameters:
//      pComputerRoles - string to receive the roles
//
//  Remarks:
//      On NT no special permissions are required at at level 101.
//      Will ignore most roles and of little interest
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetComputerRoles( String* pComputerRoles ) const
{
    size_t  i = 0;
    NET_API_STATUS   status  = 0;
    SERVER_INFO_101* pServerInfo101  = nullptr;

    struct _ROLE
    {
        DWORD   role;
        LPCWSTR pszRole;
    } Roles[] =
        { { SV_TYPE_WORKSTATION      , L"Workstation"                   },
          { SV_TYPE_SERVER           , L"Server"                        },
          { SV_TYPE_SQLSERVER        , L"SQL Server"                    },
          { SV_TYPE_DOMAIN_CTRL      , L"Primary Domain Controller"     },
          { SV_TYPE_DOMAIN_BAKCTRL   , L"Backup Domain Controller"      },
          { SV_TYPE_NOVELL           , L"Novell Server"                 },
          { SV_TYPE_SERVER_NT        , L"Non-Domain Controller Server"  },
          { SV_TYPE_POTENTIAL_BROWSER, L"Potential Browser"             },
          { SV_TYPE_BACKUP_BROWSER   , L"Backup Browser"                },
          { SV_TYPE_MASTER_BROWSER   , L"Master Browser"                },
          { SV_TYPE_DOMAIN_MASTER    , L"Domain Master Browser"         },
          { SV_TYPE_TERMINALSERVER   , L"Terminal Server"               } };

    if ( pComputerRoles == nullptr )
    {
        throw ParameterException( L"pComputerRoles", __FUNCTION__ );
    }
    *pComputerRoles = PXS_STRING_EMPTY;

    status = NetServerGetInfo( nullptr, 101, reinterpret_cast<LPBYTE*>( &pServerInfo101 ) );
    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetServerGetInfo", __FUNCTION__ );
    }

    if ( pServerInfo101 == nullptr )
    {
        throw NullException( L"pServerInfo101", __FUNCTION__ );
    }
    AutoNetApiBufferFree AutoFreeServerInfo101( pServerInfo101 );

    for ( i = 0; i < ARRAYSIZE( Roles ); i++ )
    {
        if ( pServerInfo101->sv101_type & Roles[ i ].role )
        {
            if ( pComputerRoles->GetLength() )
            {
                *pComputerRoles += L", ";
            }
            *pComputerRoles += Roles[ i ].pszRole;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get site name for the local computer
//
//  Parameters:
//      pComputerSiteName - string to receive the site name
//
//  Remarks:
//      Windows 2000 and newer, uses active directory
//
//      GetSiteName can take a long time, up to 45s after a successful first
//      call the second one is instant. According the the documentation:
//      If the domain that is being queried by a computer is the same as the
//      domain to which the computer is joined, the site in which the computer
//      resides (as reported by a domain controller) is stored in the computer
//      registry. The client stores this site name in the DynamicSiteName
//      registry entry in HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services
//      \Netlogon\Parameters.
//
//      To override the dynamic site name, add the SiteName entry with the
//      REG_SZ data type in HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet
//      \Services\Netlogon\Parameters. When a value is present for the
//      SiteName entry, the DynamicSiteName entry is not used.
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetComputerSiteName( String* pComputerSiteName ) const
{
    DWORD     errorCode   = NO_ERROR;
    wchar_t*  pszSiteName = nullptr;
    Registry  RegObject;
    Formatter Format;
    LPCWSTR   PARAMETERS_KEY = L"SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters";

    if ( pComputerSiteName == nullptr )
    {
        throw ParameterException( L"pComputerSiteName", __FUNCTION__ );
    }
    *pComputerSiteName = PXS_STRING_EMPTY;

    // Because DsGetSiteName can take a long time, first check for cached
    // values in the registry.
    try
    {
        // Connect to the machine registry then look for SiteName
        RegObject.Connect( HKEY_LOCAL_MACHINE );
        RegObject.GetStringValue( PARAMETERS_KEY, L"SiteName", pComputerSiteName );
        pComputerSiteName->Trim();

        // If got nothing then look for DynamicSiteName
        if ( pComputerSiteName->IsEmpty() )
        {
            RegObject.GetStringValue( PARAMETERS_KEY, L"DynamicSiteName", pComputerSiteName );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
        *pComputerSiteName = PXS_STRING_EMPTY;     // Reset for second pass
    }
    pComputerSiteName->Trim();
    if ( pComputerSiteName->GetLength() )
    {
        return;     // Success
    }

    // DsGetSiteName
    errorCode = DsGetSiteName( nullptr, &pszSiteName );
    if ( ( errorCode != NO_ERROR ) && pszSiteName )
    {
        AutoNetApiBufferFree AutoFreeSiteName( pszSiteName );
        *pComputerSiteName = pszSiteName;
        pComputerSiteName->Trim();
    }
    else
    {
        // Probably ERROR_NO_SITENAME = the computer does not belong to a site
        PXSLogAppWarn1( L"DsGetSiteName failed code = %%1.", Format.UInt32( errorCode ) );
    }
}

//===============================================================================================//
//  Description:
//      Get the user name
//
//  Parameters:
//      pCurrentUserName - string to receive name
//
//  Remarks:
//      In Lmcons.h #define UNLEN  256  // Maximum user name length
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetCurrentUserName( String* pCurrentUserName ) const
{
    DWORD   nSize;
    wchar_t Buffer[ UNLEN + 1 ] = { 0 };

    if ( pCurrentUserName == nullptr )
    {
        throw ParameterException( L"pCurrentUserName", __FUNCTION__ );
    }
    *pCurrentUserName = PXS_STRING_EMPTY;

    nSize = ARRAYSIZE( Buffer );
    if ( GetUserName( Buffer, &nSize ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetUserName", __FUNCTION__ );
    }
    Buffer[ ARRAYSIZE( Buffer ) - 1 ] = PXS_CHAR_NULL;
    *pCurrentUserName = Buffer;
    pCurrentUserName->Trim();
}

//===============================================================================================//
//  Description:
//      Get the NetBIOS domain name
//
//  Parameters:
//      pDomainNetBiosName - string to receive the domain name
//
//  Returns:
//       void
//===============================================================================================//
void SystemInformation::GetDomainNetBiosName( String* pDomainNetBiosName ) const
{
    NET_API_STATUS  status = 0;
    WKSTA_INFO_100* pWkstaInfo100 = nullptr;

    // wki100_langroup is defined as LMSTR which is only Unicode if one of the
    // following are defined _WIN32_WINNT WINNT, __midl, FORCE_UNICODE
    #if !defined( _WIN32_WINNT )
        #error Must define _WIN32_WINNT to use WKSTA_INFO_100::wki100_langroup
    #endif

    if ( pDomainNetBiosName == nullptr )
    {
        throw ParameterException( L"pDomainNetBiosName", __FUNCTION__ );
    }
    *pDomainNetBiosName = PXS_STRING_EMPTY;

    // Get workstation info at level 100, requires no special permissions
    status = NetWkstaGetInfo( nullptr, 100, reinterpret_cast<LPBYTE*>( &pWkstaInfo100 ) );
    if ( ( NERR_Success != status ) || ( pWkstaInfo100 == nullptr ) )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK,
                         static_cast<DWORD>( status ), L"NetWkstaGetInfo", __FUNCTION__ );
    }
    AutoNetApiBufferFree AutoFreeWkstaInfo100( pWkstaInfo100 );
    *pDomainNetBiosName = pWkstaInfo100->wki100_langroup;
    pDomainNetBiosName->Trim();
}

//===============================================================================================//
//  Description:
//      Get the edition of the operating system
//
//  Parameters:
//      pEdition - string object to receive the edition description
//
//  Remarks:
//      The operating system number (e.g. 2008, 10 etc.) have been removed from
//      the descriptions as Windows 7 Professional reports a product type of
//      PRODUCT_PROFESSIONAL (0x00000030) but the description is
//      "Windows 10 Pro". As the operating system number is obtained in method
//      GetName(), no inormation is lost by reming the numbers from these
//      descriptions
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetEdition( String* pEdition ) const
{
    size_t i   = 0;
    DWORD  type = 0;
    OSVERSIONINFOEX  VersionInfoEx;
    SystemInformation SystemInfo;

    // See GetProductInfo
    struct _TYPE
    {
        DWORD   type;
        LPCWSTR pszType;
    } Types[] = {
        { PRODUCT_BUSINESS,                            L"Business"},
        { PRODUCT_BUSINESS_N,                          L"Business N"},
        { PRODUCT_CLUSTER_SERVER,                      L"HPC Edition"},
        { PRODUCT_CLUSTER_SERVER_V,                    L"Server Hyper Core V"},
        { PRODUCT_CORE,                                L"Windows Home"},
        { PRODUCT_CORE_COUNTRYSPECIFIC,                L"Windows Home China"},
        { PRODUCT_CORE_N,                              L"Windows Home N"},
        { PRODUCT_CORE_SINGLELANGUAGE,                 L"Windows Home Single Language"},
        { PRODUCT_DATACENTER_EVALUATION_SERVER,        L"Server Datacenter (evaluation "
                                                       L"installation)"},
        { PRODUCT_DATACENTER_SERVER,                   L"Server Datacenter (full installation)"},
        { PRODUCT_DATACENTER_SERVER_CORE,              L"Server Datacenter (core installation)"},
        { PRODUCT_DATACENTER_SERVER_CORE_V,            L"Server Datacenter without Hyper-V "
                                                       L"(core installation)"},
        { PRODUCT_DATACENTER_SERVER_V,                 L"Server Datacenter without Hyper-V (full "
                                                       L"installation)"},
        { PRODUCT_EDUCATION,                           L"Windows Education"},
        { PRODUCT_EDUCATION_N,                         L"Windows Education N"},
        { PRODUCT_ENTERPRISE,                          L"Windows Enterprise"},
        { PRODUCT_ENTERPRISE_E,                        L"Windows Enterprise E"},
        { PRODUCT_ENTERPRISE_EVALUATION,               L"Windows Enterprise Evaluation"},
        { PRODUCT_ENTERPRISE_N,                        L"Windows Enterprise N"},
        { PRODUCT_ENTERPRISE_N_EVALUATION,             L"Windows Enterprise N Evaluation"},
        { PRODUCT_ENTERPRISE_S,                        L"Windows Enterprise 2015 LTSB"},
        { PRODUCT_ENTERPRISE_S_EVALUATION,             L"Windows Enterprise 2015 LTSB Evaluation"},
        { PRODUCT_ENTERPRISE_S_N,                      L"Windows Enterprise 2015 LTSB N"},
        { PRODUCT_ENTERPRISE_S_N_EVALUATION,           L"Windows Enterprise 2015 LTSB N "
                                                       L"Evaluation"},
        { PRODUCT_ENTERPRISE_SERVER,                   L"Server Enterprise (full installation)"},
        { PRODUCT_ENTERPRISE_SERVER_CORE,              L"Server Enterprise (core installation)"},
        { PRODUCT_ENTERPRISE_SERVER_CORE_V,            L"Server Enterprise without Hyper-V (core "
                                                       L"installation)"},
        { PRODUCT_ENTERPRISE_SERVER_IA64,              L"Server Enterprise for Itanium-based "
                                                       L"Systems"},
        { PRODUCT_ENTERPRISE_SERVER_V,                 L"Server Enterprise without Hyper-V (full "
                                                       L"installation)"},
        { PRODUCT_ESSENTIALBUSINESS_SERVER_ADDL,       L"Windows Essential Server Solution "
                                                       L"Additional"},
        { PRODUCT_ESSENTIALBUSINESS_SERVER_ADDLSVC,    L"Windows Essential Server Solution "
                                                       L"Additional SVC"},
        { PRODUCT_ESSENTIALBUSINESS_SERVER_MGMT,       L"Windows Essential Server Solution "
                                                       L"Management"},
        { PRODUCT_ESSENTIALBUSINESS_SERVER_MGMTSVC,    L"Windows Essential Server Solution "
                                                       L"Management SVC"},
        { PRODUCT_HOME_BASIC,                          L"Home Basic"},
        { PRODUCT_HOME_BASIC_E,                        L"Not supported"},
        { PRODUCT_HOME_BASIC_N,                        L"Home Basic N"},
        { PRODUCT_HOME_PREMIUM,                        L"Home Premium"},
        { PRODUCT_HOME_PREMIUM_E,                      L"Not supported"},
        { PRODUCT_HOME_PREMIUM_N,                      L"Home Premium N"},
        { PRODUCT_HOME_PREMIUM_SERVER,                 L"Windows Home Server"},
        { PRODUCT_HOME_SERVER,                         L"Windows Storage Server Essentials"},
        { PRODUCT_HYPERV,                              L"Microsoft Hyper-V Server"},
        { PRODUCT_IOTUAP,                              L"Windows IoT Core"},
        { PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT,    L"Windows Essential Business Server "
                                                       L"Management Server"},
        { PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING,     L"Windows Essential Business Server "
                                                       L"Messaging Server"},
        { PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY,      L"Windows Essential Business Server "
                                                       L"Security Server"},
        { PRODUCT_MOBILE_CORE,                         L"Windows Mobile"},
        { PRODUCT_MOBILE_ENTERPRISE,                   L"Windows Mobile Enterprise"},
        { PRODUCT_MULTIPOINT_PREMIUM_SERVER,           L"Windows MultiPoint Server Premium (full "
                                                       L"installation)"},
        { PRODUCT_MULTIPOINT_STANDARD_SERVER,          L"Windows MultiPoint Server Standard (full "
                                                       L"installation)"},
        { PRODUCT_PROFESSIONAL,                        L"Windows Pro"},
        { PRODUCT_PROFESSIONAL_E,                      L"Not supported"},
        { PRODUCT_PROFESSIONAL_N,                      L"Windows Pro N"},
        { PRODUCT_PROFESSIONAL_WMC,                    L"Professional with Media Center"},
        { PRODUCT_SB_SOLUTION_SERVER,                  L"Windows Small Business Server 2011 "
                                                       L"Essentials"},
        { PRODUCT_SB_SOLUTION_SERVER_EM,               L"Server For SB Solutions EM"},
        { PRODUCT_SERVER_FOR_SB_SOLUTIONS,             L"Server For SB Solutions"},
        { PRODUCT_SERVER_FOR_SB_SOLUTIONS_EM,          L"Server For SB Solutions EM"},
        { PRODUCT_SERVER_FOR_SMALLBUSINESS,            L"Windows Server for Windows Essential "
                                                       L"Server Solutions"},
        { PRODUCT_SERVER_FOR_SMALLBUSINESS_V,          L"Windows Server without Hyper-V for "
                                                       L"Windows Essential Server Solutions"},
        { PRODUCT_SERVER_FOUNDATION,                   L"Server Foundation"},
        { PRODUCT_SMALLBUSINESS_SERVER,                L"Windows Small Business Server"},
        { PRODUCT_SMALLBUSINESS_SERVER_PREMIUM,        L"Small Business Server Premium"},
        { PRODUCT_SMALLBUSINESS_SERVER_PREMIUM_CORE,   L"Small Business Server Premium (core "
                                                       L"installation)"},
        { PRODUCT_SOLUTION_EMBEDDEDSERVER,             L"Windows MultiPoint Server"},
        { PRODUCT_STANDARD_EVALUATION_SERVER,          L"Server Standard (evaluation "
                                                       L"installation)"},
        { PRODUCT_STANDARD_SERVER,                     L"Server Standard"},
        { PRODUCT_STANDARD_SERVER_CORE ,               L"Server Standard (core installation)"},
        { PRODUCT_STANDARD_SERVER_CORE_V,              L"Server Standard without Hyper-V (core "
                                                       L"installation)"},
        { PRODUCT_STANDARD_SERVER_V,                   L"Server Standard without Hyper-V"},
        { PRODUCT_STANDARD_SERVER_SOLUTIONS,           L"Server Solutions Premium"},
        { PRODUCT_STANDARD_SERVER_SOLUTIONS_CORE,      L"Server Solutions Premium (core "
                                                       L"installation)"},
        { PRODUCT_STARTER,                             L"Starter"},
        { PRODUCT_STARTER_E,                           L"Not supported"},
        { PRODUCT_STARTER_N,                           L"Starter N"},
        { PRODUCT_STORAGE_ENTERPRISE_SERVER,           L"Storage Server Enterprise"},
        { PRODUCT_STORAGE_ENTERPRISE_SERVER_CORE,      L"Storage Server Enterprise (core "
                                                       L"installation)"},
        { PRODUCT_STORAGE_EXPRESS_SERVER,              L"Storage Server Express"},
        { PRODUCT_STORAGE_EXPRESS_SERVER_CORE,         L"Storage Server Express (core "
                                                       L"installation)"},
        { PRODUCT_STORAGE_STANDARD_EVALUATION_SERVER,  L"Storage Server Standard (evaluation "
                                                       L"installation)"},
        { PRODUCT_STORAGE_STANDARD_SERVER,             L"Storage Server Standard"},
        { PRODUCT_STORAGE_STANDARD_SERVER_CORE,        L"Storage Server Standard (core "
                                                       L"installation)"},
        { PRODUCT_STORAGE_WORKGROUP_EVALUATION_SERVER, L"Storage Server Workgroup (evaluation "
                                                       L"installation)"},
        { PRODUCT_STORAGE_WORKGROUP_SERVER,            L"Storage Server Workgroup"},
        { PRODUCT_STORAGE_WORKGROUP_SERVER_CORE,       L"Storage Server Workgroup (core "
                                                       L"installation)"},
        { PRODUCT_ULTIMATE,                            L"Ultimate"},
        { PRODUCT_ULTIMATE_E,                          L"Not supported"},
        { PRODUCT_ULTIMATE_N,                          L"Ultimate N"},
        { PRODUCT_UNDEFINED,                           L"An unknown product"},
        { PRODUCT_WEB_SERVER,                          L"Web Server (full installation)"},
        { PRODUCT_WEB_SERVER_CORE,                     L"Web Server (core installation)"}
    };

    if ( pEdition == nullptr )
    {
        throw ParameterException( L"pEdition", __FUNCTION__ );
    }
    *pEdition = PXS_STRING_EMPTY;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );

    // Vista onwards
    if ( VersionInfoEx.dwMajorVersion >= 6 )
    {
        type = GetProductInfoType();
        for ( i = 0; i < ARRAYSIZE( Types ); i++ )
        {
            if ( type == Types[ i ].type )
            {
                *pEdition = Types[ i ].pszType;
                break;
            }
        }
    }
    else
    {
        // XP or 2003
        if ( VersionInfoEx.wProductType == VER_NT_WORKSTATION )
        {
            TranslateXPEdition( VersionInfoEx.wSuiteMask, pEdition );
        }
        else if ( VersionInfoEx.wProductType == VER_NT_SERVER )
        {
            // 2003, VER_SUITE_SMALLBUSINESS is unreliable
            if ( VersionInfoEx.wSuiteMask & VER_SUITE_BLADE )
            {
                *pEdition = L"Web Edition";
            }
            else if ( VersionInfoEx.wSuiteMask & VER_SUITE_ENTERPRISE )
            {
                *pEdition = L"Enterprise Edition";
            }
            else if ( VersionInfoEx.wSuiteMask & VER_SUITE_WH_SERVER )
            {
                *pEdition = L"Home Server";    // Variant of 2008
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the full qualified domain name of the computer
//
//  Parameters:
//      pFQDN - string to receive name
//
//    Remarks
//      Not all computers have names so "" is valid
//      Win2000 and newer
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetFullyQualifiedDomainName( String* pFQDN ) const
{
    DWORD    lastError = 0, nSize = 0, nBufSize = 0;
    wchar_t* pBuffer   = nullptr;
    AllocateWChars AllocWChars;

    if ( pFQDN == nullptr )
    {
        throw ParameterException( L"pFQDN", __FUNCTION__ );
    }
    *pFQDN = PXS_STRING_EMPTY;

    // First call to get required length
    GetComputerNameEx( ComputerNameDnsFullyQualified, nullptr, &nSize );
    lastError = GetLastError();
    if ( lastError != ERROR_MORE_DATA )
    {
        // Some other error than buffer too small
        throw SystemException( lastError, L"GetComputerNameEx", __FUNCTION__ );
    }

    // Allocate then do the second call
    nBufSize = PXSAddUInt32( nSize, 1 );    // NULL terminator
    pBuffer  = AllocWChars.New( nBufSize );
    nSize    = nBufSize;                    // note, nSize will be overwritten
                                            // GetComputerNameEx
    if ( GetComputerNameEx( ComputerNameDnsFullyQualified, pBuffer, &nSize ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetComputerNameEx", __FUNCTION__ );
    }

    pBuffer[ nBufSize - 1 ] = PXS_CHAR_NULL;    // note, using original size
    *pFQDN = pBuffer;
    pFQDN->Trim();
}

//===============================================================================================//
//  Description:
//      Get the major, minor, build and service pack of the operating system
//
//  Parameters:
//      pMajorMinorBuild - string to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetMajorMinorBuild( String* pMajorMinorBuild ) const
{
    Formatter Format;

    if ( pMajorMinorBuild == nullptr )
    {
        throw ParameterException( L"pMajorMinorBuild", __FUNCTION__ );
    }
    *pMajorMinorBuild  = Format.UInt32( GetMajorVersion() );
    *pMajorMinorBuild += PXS_STRING_DOT;
    *pMajorMinorBuild += Format.UInt32( GetMinorVersion() );
    *pMajorMinorBuild += PXS_STRING_DOT;
    *pMajorMinorBuild += Format.UInt32( GetBuildNumber() );
}

//===============================================================================================//
//  Description:
//      Get the major version number of the operating system
//
//  Parameters:
//      None
//
//  Returns:
//      The major version number
//===============================================================================================//
DWORD SystemInformation::GetMajorVersion() const
{
    OSVERSIONINFOEX VersionInfoEx;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );

    return VersionInfoEx.dwMajorVersion;
}

//===============================================================================================//
//  Description:
//      Get the minor version number of the operating system
//
//  Parameters:
//      None
//
//  Returns:
//      The minor version number
//===============================================================================================//
DWORD SystemInformation::GetMinorVersion() const
{
    OSVERSIONINFOEX VersionInfoEx;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );

    return VersionInfoEx.dwMinorVersion;
}

//===============================================================================================//
//  Description:
//      Get the operating system bit size
//
//  Parameters:
//      None
//
//  Remarks:
//      Only 32-bit or 64-bit are supported
//
//  Returns:
//      the number of bits
//===============================================================================================//
DWORD SystemInformation::GetNumberOperatingSystemBits() const
{
    DWORD numBits = 0;

    // Use the compiler constants
    #if defined ( _WIN64 )

        numBits = 64;

    #elif defined( _WIN32 )     // Always defined for MSVC

        // Test for WOW64
        if ( IsWOW64() )
        {
            numBits = 64;
        }
        else
        {
            numBits = 32;
        }

    #else

        #error _WIN32 or _WIN64 must be defined

    #endif

    return numBits;
}

//===============================================================================================//
//  Description:
//      Get various OS version numbers as reported by different sub-systems
//
//  Parameters:
//      pNetFunction   - receives the version from NetWkstaGetInfo
//      pWmiWin32      - receives the version from Win32_OperatingSystem
//      pKernel32      - receives the version from kernel32.dll
//      pRegistryVer   - receives the version from registry
//
//  Remarks:
//      GetVersionEx is deprecated with Window8.1
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetVersionStrings( String* pNetFunction,
                                           String* pWmiWin32,
                                           String* pKernel32, String* pRegistryVer ) const
{
    Wmi       WMI;
    File      FileObject;
    String    Insert1, FullPath, ManufacturerName, Description, RegValue;
    Registry  RegObject;
    Formatter Format;
    FileVersion     FileVer;
    NET_API_STATUS  status = 0;
    WKSTA_INFO_100* pWkstaInfo100 = nullptr;
    LPCWSTR REG_KEY = L"Software\\Microsoft\\Windows NT\\CurrentVersion";

    if ( ( pNetFunction   == nullptr ) ||
         ( pWmiWin32      == nullptr ) ||
         ( pKernel32      == nullptr ) ||
         ( pRegistryVer   == nullptr )  )
    {
        throw NullException( L"String*", __FUNCTION__ );
    }
    *pNetFunction   = PXS_STRING_EMPTY;
    *pWmiWin32      = PXS_STRING_EMPTY;
    *pKernel32      = PXS_STRING_EMPTY;
    *pRegistryVer   = PXS_STRING_EMPTY;

    // Net Functions
    try
    {
        // NetWkstaGetInfo at level 100, requires no special permissions
        status = NetWkstaGetInfo( nullptr, 100, reinterpret_cast<LPBYTE*>(&pWkstaInfo100) );
        if ( ( status == NERR_Success ) && ( pWkstaInfo100 != nullptr ) )
        {
            AutoNetApiBufferFree AutoFreeWkstaInfo100( pWkstaInfo100 );
            *pNetFunction  = L"wki100_ver_major = ";
            *pNetFunction += Format.UInt32( pWkstaInfo100->wki100_ver_major );
            *pNetFunction += L", wki100_ver_minor = ";
            *pNetFunction += Format.UInt32( pWkstaInfo100->wki100_ver_minor );
        }
        else
        {
           Insert1.SetAnsi( __FUNCTION__ );
           PXSLogNetError1( status, L"NetWkstaGetInfo failed in '%%1'.", Insert1 );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"NetWkstaGetInfo error.", e, __FUNCTION__ );
    }

    // WMI
    try
    {
        WMI.Connect( L"root\\cimv2" );
        WMI.ExecQuery( L"Select * from Win32_OperatingSystem" );
        WMI.Next();     // There should be only one record
        WMI.Get( L"Version", pWmiWin32 );
        WMI.Disconnect();
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"WMI error.", e, __FUNCTION__ );
    }

    // Kernel32 Version
    try
    {
        // Make the full path
        GetSystemDirectoryPath( &FullPath );
        FullPath += L"Kernel32.dll";
        if ( FileObject.Exists( FullPath ) )
        {
            FileVer.GetVersion( FullPath, &ManufacturerName, pKernel32, &Description );
        }
        else
        {
            PXSLogAppError1( L"File '%%1' not found.", FullPath );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Kernel32.dll version error.", e, __FUNCTION__ );
    }

    // Registry
    try
    {
        RegObject.Connect( HKEY_LOCAL_MACHINE );

        RegValue       = PXS_STRING_EMPTY;
        RegObject.GetStringValue( REG_KEY, L"CurrentVersion", &RegValue );
        *pRegistryVer  = L"CurrentVersion = ";
        *pRegistryVer += RegValue;

        RegValue       = PXS_STRING_EMPTY;
        RegObject.GetStringValue( REG_KEY, L"CurrentBuildNumber", &RegValue );
        *pRegistryVer += L", CurrentBuildNumber = ";
        *pRegistryVer += RegValue;
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Registry error.", e, __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get product type
//
//  Parameters:
//      none
//
//  Returns:
//      DWORD of product info type, 0 on error
//===============================================================================================//
DWORD SystemInformation::GetProductInfoType() const
{
    DWORD    dwReturnedProductType = 0;
    HMODULE  hKernel32;
    OSVERSIONINFOEX      VersionInfoEx;
    LPFN_GET_PRODUCT_INFO lpfnGetProductInfo = nullptr;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );
    hKernel32 = GetModuleHandle( L"kernel32.dll" );
    if ( hKernel32 == nullptr )
    {
        throw SystemException( GetLastError(), L"kernel32.dll", "GetModuleHandle" );
    }

    // For MSVC, disable function pointer type cast warning
    #ifdef _MSC_VER
        #pragma warning( push )
        #pragma warning ( disable : 4191 )
    #endif

    lpfnGetProductInfo = (LPFN_GET_PRODUCT_INFO)GetProcAddress( hKernel32, "GetProductInfo" );

    #ifdef _MSC_VER
        #pragma warning( pop )  // Re-instate
    #endif

    if ( lpfnGetProductInfo == nullptr )
    {
        PXSLogSysError( ERROR_PROC_NOT_FOUND, L"GetProductInfo" );
        return 0;
    }

    if ( lpfnGetProductInfo( VersionInfoEx.dwMajorVersion,
                             VersionInfoEx.dwMinorVersion,
                             VersionInfoEx.wServicePackMajor,
                             VersionInfoEx.wServicePackMinor,
                             &dwReturnedProductType ) == 0 )
    {
        throw SystemException( ERROR_INVALID_DATA, L"GetProductInfo", __FUNCTION__ );
    }

    return dwReturnedProductType;
}

//===============================================================================================//
//  Description:
//      Get the service pack version in major.minor form
//
//  Parameters:
//      pMajorDotMinor - string object to receive the data
//
//  Remarks:
//      Only available on NT4 SP6 and NT5+
//      0.0 means no service pack
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetServicePackMajorDotMinor( String* pMajorDotMinor ) const
{
    Formatter Format;
    OSVERSIONINFOEX VersionInfoEx;

    if ( pMajorDotMinor == nullptr )
    {
        throw ParameterException( L"pMajorDotMinor", __FUNCTION__ );
    }
    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );
    *pMajorDotMinor  = Format.UInt32( VersionInfoEx.wServicePackMajor );
    *pMajorDotMinor += PXS_CHAR_DOT;
    *pMajorDotMinor += Format.UInt32( VersionInfoEx.wServicePackMinor );
}

//===============================================================================================//
//  Description:
//      Get the SESSIONNAME environmental variable
//
//  Parameters:
//      pSessionName - string to receive data
//
//  Remarks:
//      Normally is is 'Console' but can be something else for remote sessions
//
//      GetEnvironmentVariable() fails if buffer is too small,
//      in which case it will return the number of bytes required
//      Zero bytes is a valid result i.e. SESSIONNAME = ""
//
//  Returns:
//      true if got session name
//===============================================================================================//
void SystemInformation::GetSessionName( String* pSessionName ) const
{
    LPCWSTR  SESSION_NAME = L"SESSIONNAME";
    DWORD    nSize   = 0;
    wchar_t* pBuffer = nullptr;
    AllocateWChars AllocWChars;

    if ( pSessionName == nullptr )
    {
        throw ParameterException( L"pSessionName", __FUNCTION__ );
    }
    *pSessionName = PXS_STRING_EMPTY;

    // First call to get the required buffer size
    nSize = GetEnvironmentVariable( SESSION_NAME, nullptr, 0 );
    if ( nSize == 0 )
    {
        // Unexpected so log it
        PXSLogSysInfo( GetLastError(), L"GetEnvironmentVariable for SESSIONNAME" );
        return;
    }

    // Second call
    nSize   = PXSAddUInt32( nSize, 1 );     // NULL terminator
    pBuffer = AllocWChars.New( nSize );
    if ( GetEnvironmentVariable( SESSION_NAME, pBuffer, nSize ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetEnvironmentVariable", __FUNCTION__ );
    }
    pBuffer[ nSize - 1 ] = PXS_CHAR_NULL;
    *pSessionName = pBuffer;
}

//===============================================================================================//
//  Description:
//      Get the system directory
//
//  Parameters:
//      pDirectoryPath - string to receive directory path
//
//  Remarks:
//      If obtained the system path, ensures it has a trailing separator
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetSystemDirectoryPath( String* pDirectoryPath ) const
{
    DWORD   numChars = 0;
    wchar_t Buffer[ MAX_PATH + 1 ] = { 0 };

    if ( pDirectoryPath == nullptr )
    {
        throw ParameterException( L"pDirectoryPath", __FUNCTION__ );
    }
    *pDirectoryPath = PXS_STRING_EMPTY;

    numChars = GetSystemDirectory( Buffer, ARRAYSIZE( Buffer ) );
    if ( numChars == 0 )
    {
        throw SystemException( GetLastError(), L"GetSystemDirectory", __FUNCTION__ );
    }

    if ( numChars > ARRAYSIZE( Buffer ) )
    {
        throw SystemException( ERROR_INSUFFICIENT_BUFFER, L"GetSystemDirectory", __FUNCTION__ );
    }
    Buffer[ ARRAYSIZE( Buffer ) - 1 ] = PXS_CHAR_NULL;
    *pDirectoryPath = Buffer;
    pDirectoryPath->Trim();

    // Add a path separator if required
    if ( ( pDirectoryPath->GetLength() ) &&
         ( pDirectoryPath->EndsWithCharacter( PXS_PATH_SEPARATOR ) == false ) )
    {
        *pDirectoryPath += PXS_PATH_SEPARATOR;
    }
}

//===============================================================================================//
//  Description:
//      Get the version of a Window's system library
//
//  Parameters:
//      pszLibName      - pointer to string of library name. NB name only
//      pLibraryVersion - receives the version
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetSystemLibraryVersion( LPCWSTR pszLibName,
                                                 String* pLibraryVersion ) const
{
    File        FileObject;
    String      FullPath, ManufacturerName, Description;
    FileVersion FileVer;

    if ( pLibraryVersion == nullptr )
    {
        throw ParameterException( L"pLibraryVersion", __FUNCTION__ );
    }
    *pLibraryVersion = PXS_STRING_EMPTY;

    if ( ( pszLibName == nullptr ) || ( *pszLibName == PXS_CHAR_NULL ) )
    {
        return;     // Nothing to do or error
    }

    // Make the full path
    GetSystemDirectoryPath( &FullPath );
    FullPath += pszLibName;
    if ( FileObject.Exists( FullPath ) )
    {
        FileVer.GetVersion( FullPath, &ManufacturerName, pLibraryVersion, &Description );
    }
}

//===============================================================================================//
//  Description:
//      Get the system's uptime
//
//  Parameters:
//      none
//
//  Remarks
//      HKEY_PERFORMANCE_DATA is for NT only however, on NT4.0 when call
//      RegCloseKey after using HKEY_PERFORMANCE_DATA have observed
//      exceptions raised in other dll's and a memory leak. Each time
//      it is called the process handle count increments by 1 so only
//      use this method on NT5 and up.
//
//  Returns:
//      seconds up, zero on error
//===============================================================================================//
time_t SystemInformation::GetUpTime() const
{
    const DWORD BYTE_INCREMENT  = 4096;
    const DWORD MAX_BUFFER_SIZE = 1024 * 1024;

    LONG    result = 0;
    DWORD   i = 0, bufferSize = 0, cbData = 0;
    BYTE*   pData  = nullptr;
    size_t  offset = 0, counterOffset = 0;
    wchar_t szTwo[ 2 ] = { '2', PXS_CHAR_NULL };    // 2 == System category
    __int64 systemUpTime = 0;
    AllocateBytes     AllocBytes;
    PERF_DATA_BLOCK*  pDataBlock      = nullptr;
    PERF_OBJECT_TYPE* pObjectType     = nullptr;
    PERF_COUNTER_DEFINITION* pCounter = nullptr;

    // Need to determine the required buffer size, stop if too big.
    // The documentation says keep a separate variable for the buffer size
    // as the value placed in cbData is unpredictable
    do
    {
        AllocBytes.Delete();
        bufferSize = PXSAddUInt32( bufferSize, BYTE_INCREMENT );
        cbData     = bufferSize;
        pData      = AllocBytes.New( bufferSize );
        result     = RegQueryValueEx( HKEY_PERFORMANCE_DATA,
                                      szTwo, nullptr, nullptr, pData, &cbData);
    } while ( ( result    == ERROR_MORE_DATA ) &&
              ( bufferSize < MAX_BUFFER_SIZE )  );
    RegCloseKey( HKEY_PERFORMANCE_DATA );

    // If got a result, then try to get the uptime. The Signature is WCHAR. It
    // is the first member of PERF_DATA_BLOCK. See MSDN "How the Performance
    // Data Is Structured"
    if ( ( pData  != nullptr ) &&
         ( result == ERROR_SUCCESS ) &&
         ( memcmp( pData, L"PERF", 4 ) == 0 ) )  // Signature[4] is first member
    {
        pDataBlock    = reinterpret_cast<PERF_DATA_BLOCK*>( pData );
        pObjectType   = reinterpret_cast<PERF_OBJECT_TYPE*>( pData + pDataBlock->HeaderLength );
        counterOffset = pDataBlock->HeaderLength + pObjectType->HeaderLength;
        while ( ( i < pObjectType->NumCounters ) && ( systemUpTime == 0 ) )
        {
            // System Up Time is number 674
            pCounter  = reinterpret_cast<PERF_COUNTER_DEFINITION*>( pData + counterOffset );
            if ( pCounter->CounterNameTitleIndex == 674 )
            {
                offset = pDataBlock->HeaderLength      +
                         pObjectType->DefinitionLength +
                         pCounter->CounterOffset;
                memcpy( &systemUpTime, pData + offset, sizeof (systemUpTime) );

                // Need to adjust by time of snap shot then apply the frequency
                if ( pObjectType->PerfFreq.QuadPart > 0 )
                {
                    systemUpTime -= pObjectType->PerfTime.QuadPart;
                    systemUpTime /= pObjectType->PerfFreq.QuadPart;
                }
            }
            counterOffset = PXSAddSizeT( counterOffset, pCounter->ByteLength );
            i++;
        }
    }
    else
    {
        PXSLogAppWarn( L"Failed to get registry performance data." );
    }

    // If got nothing use the GetTickCount
    if ( systemUpTime <= 0 )
    {
        systemUpTime = GetTickCount() / 1000;
    }

    return PXSCastInt64ToTimeT( systemUpTime );
}

//===============================================================================================//
//  Description:
//      Get product type as in the OSVERSIONINFOEX structure
//
//  Parameters:
//      none
//
//  Returns:
//      BYTE representing the product type
//===============================================================================================//
BYTE SystemInformation::GetVersionProductType() const
{
    OSVERSIONINFOEX VersionInfoEx;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );

    return VersionInfoEx.wProductType;
}

//===============================================================================================//
//  Description:
//      Get product suite mask
//
//  Parameters:
//      none
//
//  Returns:
//      WORD representing the product suit mask
//===============================================================================================//
WORD SystemInformation::GetVersionSuiteMask() const
{
    OSVERSIONINFOEX VersionInfoEx;

    memset( &VersionInfoEx, 0, sizeof ( VersionInfoEx ) );
    FillVersionInfoEx( &VersionInfoEx );

    return VersionInfoEx.wSuiteMask;
}

//===============================================================================================//
//  Description:
//      Get the windows directory path
//
//  Parameters:
//      pDirectoryPath - string to receive the directory path
//
//  Remarks:
//      On success appends ensures the path has a trailing separator
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::GetWindowsDirectoryPath( String* pDirectoryPath ) const
{
    DWORD   numChars;
    wchar_t Buffer[ MAX_PATH + 1 ] = { 0 };

    if ( pDirectoryPath == nullptr )
    {
        throw ParameterException( L"pDirectoryPath", __FUNCTION__ );
    }
    *pDirectoryPath = PXS_STRING_EMPTY;

    numChars = GetWindowsDirectory( Buffer, ARRAYSIZE( Buffer ) );
    if ( numChars == 0 )
    {
        throw SystemException( GetLastError(), L"GetWindowsDirectory", __FUNCTION__ );
    }

    if ( numChars > ARRAYSIZE( Buffer ) )
    {
        throw SystemException( ERROR_INSUFFICIENT_BUFFER, L"GetWindowsDirectory", __FUNCTION__ );
    }
    Buffer[ ARRAYSIZE( Buffer ) - 1 ] = PXS_CHAR_NULL;
    *pDirectoryPath = Buffer;
    pDirectoryPath->Trim();

    // Add separator if required
    if ( ( pDirectoryPath->GetLength() ) &&
         ( pDirectoryPath->EndsWithCharacter( PXS_PATH_SEPARATOR ) == false ) )
    {
        *pDirectoryPath += PXS_PATH_SEPARATOR;
    }
}

//===============================================================================================//
//  Description:
//      Get if the system is running on a 64-bit processor
//
//  Parameters:
//      None
//
//  Remarks:
//      XP and newer

//  Returns:
//      true if running on a 64-bit processor, otherwise false
//===============================================================================================//
bool SystemInformation::Is64BitProcessor() const
{
    SYSTEM_INFO si;

    memset( &si, 0, sizeof ( si ) );
    GetNativeSystemInfo( &si );
    if ( (si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64        ) ||
         (si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64         ) ||
         (si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA32_ON_WIN64)  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the specified account on the local machine is enabled
//
//  Parameters:
//      UserName - the account name
//
//  Remarks:
//      NT only
//
//  Returns:
//      true if is account is enabled, otherwise false
//===============================================================================================//
bool SystemInformation::IsAccountEnabled( const String& UserName ) const
{
    wchar_t   wzUserName[ 256 ] = { 0 };    // Enough for a user name
    size_t    length = 0;
    Formatter Format;
    USER_INFO_1*   pUserInfo1 = nullptr;
    NET_API_STATUS status     = 0;

    // Check input
    length = UserName.GetLength();
    if ( length == 0 )
    {
        return false;       // Or error
    }

    if ( length >= ARRAYSIZE( wzUserName ) )
    {
        throw SystemException( ERROR_INSUFFICIENT_BUFFER, L"wzUserName", __FUNCTION__ );
    }
    Format.StringToWide( UserName, wzUserName, ARRAYSIZE( wzUserName ) );

    // Get USER_INFO_1 info
    status = NetUserGetInfo( nullptr, wzUserName, 1, reinterpret_cast<LPBYTE*>( &pUserInfo1 ) );
    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, UserName.c_str(), "NetUserGetInfo" );
    }

    // Verify have a buffer was allocated
    if ( pUserInfo1 == nullptr )
    {
        throw NullException( L"pUserInfo1", __FUNCTION__ );
    }
    AutoNetApiBufferFree AutoFreeUserInfo1( pUserInfo1 );

    if ( pUserInfo1->usri1_flags & UF_ACCOUNTDISABLE )
    {
        return false;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Determine if the current user is in the Admin group
//
//  Parameters:
//      void
//
//  Remarks:
//      Only applies to NT
//
//  Returns:
//      true if current user is Admin, otherwise false
//===============================================================================================//
bool SystemInformation::IsAdministrator() const
{
    bool   isAdmin   = false;
    PSID   pSidAdmin = nullptr;
    DWORD  lastError = 0, i = 0;
    DWORD  TokenInformationLength = 0, ReturnLength = 0;
    HANDLE TokenHandle    = nullptr;
    String Insert1;
    AllocateBytes AllocBytes;
    TOKEN_GROUPS* pTokenInformation = nullptr;

    // SID_IDENTIFIER_AUTHORITY.Value[] is a 6 byte array
    SID_IDENTIFIER_AUTHORITY IdentifierAuthority = { SECURITY_NT_AUTHORITY };

    // Get an access token
    if ( OpenThreadToken( GetCurrentThread(), TOKEN_QUERY, TRUE, &TokenHandle ) == 0 )
    {
       // Retry against process token
       if ( OpenProcessToken( GetCurrentProcess(), TOKEN_QUERY, &TokenHandle ) == 0 )
       {
            throw SystemException( GetLastError(), L"OpenProcessToken", __FUNCTION__ );
       }
    }
    AutoCloseHandle CloseTokenHandle( TokenHandle );

    // First call to get the required buffer size
    ReturnLength = 0;
    if ( GetTokenInformation( TokenHandle, TokenGroups, nullptr, 0, &ReturnLength ) )
    {
        // Unexpected success but no data
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo1( L"GetTokenInformation returned no data in '%%1'.", Insert1 );
        return false;
    }

    lastError = GetLastError();
    if ( lastError != ERROR_INSUFFICIENT_BUFFER )
    {
        throw SystemException( lastError, L"GetTokenInformation, first call", __FUNCTION__ );
    }

    // Second call
    TokenInformationLength = PXSMultiplyUInt32( ReturnLength, 2 );
    pTokenInformation  = reinterpret_cast<TOKEN_GROUPS*>(
                                                        AllocBytes.New( TokenInformationLength ) );
    ReturnLength       = 0;     // Reset
    if ( GetTokenInformation( TokenHandle,
                              TokenGroups,
                              pTokenInformation, TokenInformationLength, &ReturnLength )  == 0 )
    {
        throw SystemException( GetLastError(), L"GetTokenInformation, call 2", __FUNCTION__ );
    }

    if ( AllocateAndInitializeSid( &IdentifierAuthority,
                                   2,
                                   SECURITY_BUILTIN_DOMAIN_RID,
                                   DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &pSidAdmin ) == 0 )
    {
        throw SystemException( GetLastError(), L"AllocateAndInitializeSid", __FUNCTION__ );
    }

    if ( pSidAdmin == nullptr )
    {
        throw NullException( L"pSidAdmin", __FUNCTION__ );
    }

    for ( i = 0; i < pTokenInformation->GroupCount; i++ )
    {
        if ( EqualSid( pSidAdmin, pTokenInformation->Groups[ i ].Sid ) )
        {
            isAdmin = true;
            break;
        }
    }
    FreeSid( pSidAdmin );

    return isAdmin;
}

//===============================================================================================//
//  Description:
//      See if the process is a remotely controlled session
//
//  Parameters:
//      void
//
//  Remarks
//      XP and newer
//
//      GetSystemMetrics returns 0 on error and does not
//      provide error information via GetLastError.
//
//  Returns:
//      true if running as a remotely controlled session
//===============================================================================================//
bool SystemInformation::IsRemoteControl() const
{
    if ( GetSystemMetrics( SM_REMOTECONTROL ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      See if the process is a remote session
//
//  Parameters:
//      void
//
//  Remarks
//      NT4+SP3 and newer. Not on Win9X
//
//      Shows if connected via Remote Desktop, either exclusively as when
//      connecting to a workstation or when logged in as a remote user onto
//      a Terminal Server.
//
//  Returns:
//      true if running as a remotely controlled session
//===============================================================================================//
bool SystemInformation::IsRemoteSession() const
{
    if ( GetSystemMetrics( SM_REMOTESESSION ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the OS is running in the pre-install environment
//
//  Parameters:
//      None
//
//  Remarks:
//      According to BartPE documentation, test for existence
//      of registry key: HKLM\system\currentcontrolset\control\minint
//
//      Note, BartPE can be detected by looking for:
//      HKLM\system\currentcontrolset\control\PE Builder
//
//  Returns:
//      true if in Pre-Install environment otherwise false
//===============================================================================================//
bool SystemInformation::IsPreInstall() const
{
    bool pe     = false;
    LONG result = 0;
    HKEY hKey   = nullptr;

    result = RegOpenKeyEx( HKEY_LOCAL_MACHINE,
                           L"SYSTEM\\CurrentControlSet\\Control\\MiniNT", 0, KEY_READ, &hKey );
    if ( ( result == ERROR_SUCCESS ) && hKey )
    {
        pe = true;
        RegCloseKey( hKey );
    }

    return pe;
}

//===============================================================================================//
//  Description:
//      Get if the system emulating 32-bit windows to run this programme.
//
//  Parameters:
//      None
//
//  Remarks:
//      Uses IsWow64Process which requires XP + SP2
//
//  Returns:
//      true if is emulating 32-bit windows otherwise false
//===============================================================================================//
bool SystemInformation::IsWOW64() const
{
    BOOL wow64 = FALSE;

    if ( IsWow64Process( GetCurrentProcess(), &wow64 ) == 0 )
    {
        throw SystemException( GetLastError(), L"IsWow64Process", __FUNCTION__ );
    }

    if ( wow64 )
    {
        return true;
    }

    return false;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Fill the OSVERSIONINFOEX version structure
//
//  Parameters:
//      pVersionInfoEx - OSVERSIONINFOEX structure
//
//  Remarks:
//      Requires Windows NT Workstation 4.0 SP6 and later.
//      the structure OSVERSIONINFOEX has changed definition in Winbase.h
//      since release of NT5, but the size is the same
//
//      For Windows 8.1, GetVersionEx is deprecated and will return incorrect
//      version and build numbers unless something is manifested.
//      Unfortunately, GetVersionEx returns with success so will test for OS
//      version and use WMI which returns the correct value.
//
//  Returns:
//     void
//===============================================================================================//
void SystemInformation::FillVersionInfoEx( OSVERSIONINFOEX* pVersionInfoEx ) const
{
    Wmi         WMI;
    String      Token, Version, CSDVersion;
    Formatter   Format;
    StringArray VersionTokens;
    OSVERSIONINFO* pOSI;

    if ( pVersionInfoEx == nullptr )
    {
        throw ParameterException( L"pVersionInfoEx", __FUNCTION__ );
    }
    memset( pVersionInfoEx, 0, sizeof ( OSVERSIONINFOEX ) );

    pVersionInfoEx->dwOSVersionInfoSize = sizeof ( OSVERSIONINFOEX );
    pOSI = reinterpret_cast<OSVERSIONINFO*>( pVersionInfoEx );
    if ( GetVersionEx( pOSI ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetVersionEx", __FUNCTION__ );
    }

    // For version Windows 8 and newer will use WMI. According to MSDN
    // GetVersionEx will report 6.2 on Windows 10
    if ( ( pOSI->dwMajorVersion < 6 ) ||
         ( ( pOSI->dwMajorVersion == 6 ) && ( pOSI->dwMinorVersion <= 1 ) ) )
    {
        return;
    }

    WMI.Connect( L"root\\cimv2" );
    WMI.ExecQuery( L"Select * from Win32_OperatingSystem" );
    WMI.Next();     // There should be only one record
    WMI.Get( L"Version", &Version );
    WMI.Disconnect();
    Version.ToArray( PXS_CHAR_DOT, &VersionTokens );
    if ( VersionTokens.GetSize() != 3 )
    {
        PXSLogAppWarn1( L"WMI failed. Unexpected OS Version '%%1'.", Version );
        return;
    }
    // Major version
    Token = VersionTokens.Get( 0 );
    pOSI->dwMajorVersion = Format.StringToUInt32( Token );

    // Minor version
    Token = VersionTokens.Get( 1 );
    pOSI->dwMinorVersion = Format.StringToUInt32( Token );

    // Build Number
    Token = VersionTokens.Get( 2 );
    pOSI->dwBuildNumber = Format.StringToUInt32( Token );
}

//===============================================================================================//
//  Description:
//      Translate the specified product type to an XP edition
//
//  Parameters:
//      suiteMask - the suite mask
//      pEdition  - receives the edition
//
//  Returns:
//      void
//===============================================================================================//
void SystemInformation::TranslateXPEdition( WORD suiteMask, String* pEdition ) const
{
    SystemInformation SystemInfo;

    if ( pEdition == nullptr )
    {
        throw ParameterException( L"pEdition", __FUNCTION__ );
    }
    *pEdition = PXS_STRING_EMPTY;

    if ( suiteMask & VER_SUITE_PERSONAL )
    {
        *pEdition = L"Home";
    }
    else if ( suiteMask & VER_SUITE_EMBEDDEDNT )
    {
        *pEdition = L"Embedded";
    }
    else
    {
        *pEdition = L"Professional";
    }

    // 64-bit
    if ( SystemInfo.GetNumberOperatingSystemBits() == 64 )
    {
        *pEdition += L" x64";    // XP has no IA64
    }

    // Override with System metrics
    if ( GetSystemMetrics( SM_MEDIACENTER ) )
    {
        *pEdition = L"Media Center";
    }
    else if ( GetSystemMetrics( SM_TABLETPC ) )
    {
        *pEdition = L"Tablet PC";
    }
    else if ( GetSystemMetrics( SM_STARTER ) )
    {
        *pEdition = L"Starter";
    }
}
