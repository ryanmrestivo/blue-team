///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Windows Firewall Information Class Implementation
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
#include "WinAudit/Header Files/WindowsFirewallInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AutoIUnknownRelease.h"
#include "PxsBase/Header Files/AutoSysFreeString.h"
#include "PxsBase/Header Files/AutoVariantClear.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
WindowsFirewallInformation::WindowsFirewallInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
WindowsFirewallInformation::~WindowsFirewallInformation()
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
//      Get windows firewall settings as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the records
//
//  Remarks:
//      Available on XP with service pack 2
//
//  Returns:
//      void
//===============================================================================================//
void WindowsFirewallInformation::GetWindowsFirewallRecords( TArray< AuditRecord >* pRecords )
{
    HRESULT        hResult = 0;
    AuditRecord    Record;
    INetFwMgr*     pINetFwMgr     = nullptr;
    INetFwPolicy*  pINetFwPolicy  = nullptr;
    INetFwProfile* pINetFwProfile = nullptr;
    AutoIUnknownRelease ReleaseINetFwMgr, ReleaseINetFwPolicy;
    AutoIUnknownRelease ReleaseINetFwProfile;
    TArray< AuditRecord > AuthorisedApps;
    TArray< AuditRecord > AuthorisedSvcs;
    TArray< AuditRecord > AuthorisedPorts;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    hResult = CoCreateInstance( CLSID_NetFwMgr,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_INetFwMgr,
                                reinterpret_cast<void**>( &pINetFwMgr ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoCreateInstance/INetFwMgr", __FUNCTION__ );
    }

    if ( pINetFwMgr == nullptr )
    {
        throw NullException( L"pINetFwMgr", __FUNCTION__ );
    }
    ReleaseINetFwMgr.Set( pINetFwMgr );

    // Firewall Local Policy
    hResult = pINetFwMgr->get_LocalPolicy( &pINetFwPolicy );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwMgr::get_LocalPolicy", __FUNCTION__ );
    }

    if ( pINetFwPolicy == nullptr )
    {
        throw NullException( L"pINetFwPolicy", __FUNCTION__ );
    }
    ReleaseINetFwPolicy.Set( pINetFwPolicy );

    // Firewall Current Profile, this will fail with EPT_S_NOT_REGISTERED if
    // the firewall service is not running
    hResult = pINetFwPolicy->get_CurrentProfile( &pINetFwProfile );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwPolicy::get_CurrentProfile", __FUNCTION__ );
    }

    if ( pINetFwProfile == nullptr )
    {
        throw NullException( L"pINetFwProfile", __FUNCTION__ );
    }
    ReleaseINetFwProfile.Set( pINetFwProfile );

    Record.Reset( PXS_CATEGORY_WINDOWS_FIREWALL );
    MakeFwStatusRecord( pINetFwProfile, &Record );
    pRecords->Add( Record );

    MakeFwAuthorisedAppsRecords( pINetFwProfile, &AuthorisedApps );
    pRecords->Append( AuthorisedApps );

    MakeFwAuthorisedSvcsRecords( pINetFwProfile, &AuthorisedSvcs );
    pRecords->Append( AuthorisedSvcs );

    MakeFwAuthorisedPortsRecords( pINetFwProfile, &AuthorisedPorts );
    pRecords->Append( AuthorisedPorts );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the next Windows Firewall Authorised application using the
//      specified enumerator
//
//  Parameters:
//      pIEnumVariant  - the enumeration interface
//      pAuthorisedApp - receives the authorised application name
//
//  Returns:
//      True if fetched an item otherwise false because at end
//      of the enumeration
//===============================================================================================//
bool WindowsFirewallInformation::GetNextFwAuthorisedApp( IEnumVARIANT* pIEnumVariant,
                                                         String* pAuthorisedApp )
{
    BSTR      name        = nullptr;
    ULONG     CeltFetched = 0;
    HRESULT   hResult     = 0;
    String    VarType;
    VARIANT   variant;
    Formatter Format;
    AutoIUnknownRelease ReleaseApplication;
    INetFwAuthorizedApplication* pApplication = nullptr;

    if ( pIEnumVariant == nullptr )
    {
        throw ParameterException( L"pIEnumVariant", __FUNCTION__ );
    }

    if ( pAuthorisedApp == nullptr )
    {
        throw ParameterException( L"pAuthorisedApp", __FUNCTION__ );
    }
    *pAuthorisedApp = PXS_STRING_EMPTY;

    // Get the next application
    VariantInit( &variant );
    hResult = pIEnumVariant->Next( 1, &variant, &CeltFetched );
    if ( hResult == S_FALSE )       // S_FALSE means did not get the requested
    {                               // number of elements, i.e one
        return false;
    }
    AutoVariantClear VariantClearVar( &variant );

    // The returned variant should point to a dispatch interface
    if ( ( variant.pdispVal == nullptr ) || ( variant.vt != VT_DISPATCH ) )
    {
        VarType = Format.StringInt32( L"VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, VarType.c_str(), __FUNCTION__ );
    }

    // Get the application pointed to by the dispatch interface
    hResult = variant.pdispVal->QueryInterface( IID_INetFwAuthorizedApplication,
                                                reinterpret_cast<void**>( &pApplication ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IDispatch::QueryInterface", __FUNCTION__ );
    }

    if ( pApplication == nullptr )
    {
        throw NullException( L"pApplication", __FUNCTION__ );
    }
    ReleaseApplication.Set( pApplication );

    // Get the application's name
    hResult = pApplication->get_Name( &name );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwAuthorizedApplication::get_Name", __FUNCTION__ );
    }
    AutoSysFreeString AutoSysFreeName( name );

    *pAuthorisedApp = name;
    pAuthorisedApp->Trim();

    return true;
}

//===============================================================================================//
//  Description:
//      Get the next Windows Firewall Authorised network port using the
//      specified enumerator
//
//  Parameters:
//      pIEnumVariant   - the enumeration interface
//      pAuthorisedPort - receives the authorised port
//
//  Returns:
//      True if fetched an item otherwise false because at end
//      of the enumeration
//===============================================================================================//
bool WindowsFirewallInformation::GetNextFwAuthorisedPort( IEnumVARIANT* pIEnumVariant,
                                                          String* pAuthorisedPort )
{
    ULONG     CeltFetched = 0;
    HRESULT   hResult     = 0;
    VARIANT   variant;
    String    VarType;
    Formatter Format;
    INetFwOpenPort*     pINetFwOpenPort = nullptr;
    AutoIUnknownRelease ReleaseINetFwOpenPort;


    if ( pIEnumVariant == nullptr )
    {
        throw ParameterException( L"pIEnumVariant", __FUNCTION__ );
    }

    if ( pAuthorisedPort == nullptr )
    {
        throw ParameterException( L"pAuthorisedPort", __FUNCTION__ );
    }
    *pAuthorisedPort = PXS_STRING_EMPTY;

    // Get the next port
    VariantInit( &variant );
    hResult = pIEnumVariant->Next( 1, &variant, &CeltFetched );
    if ( hResult == S_FALSE )       // S_FALSE means did not get the requested
    {                               // number of elements, i.e one
        return false;
    }
    AutoVariantClear VariantClearVar( &variant );

    // The returned variant should point to a dispatch interface
    if ( ( variant.pdispVal == nullptr ) || ( variant.vt != VT_DISPATCH ) )
    {
        VarType = Format.StringInt32( L"VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, VarType.c_str(), __FUNCTION__ );
    }

    // Get the port pointed to by the dispatch interface
    hResult = variant.pdispVal->QueryInterface( IID_INetFwOpenPort,
                                                reinterpret_cast<void**>( &pINetFwOpenPort ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IDispatch::QueryInterface", __FUNCTION__ );
    }

    if ( pINetFwOpenPort == nullptr )
    {
        throw NullException( L"pINetFwOpenPort", __FUNCTION__ );
    }
    ReleaseINetFwOpenPort.Set( pINetFwOpenPort );
    MakeOpenPortName( pINetFwOpenPort, pAuthorisedPort );

    return true;
}

//===============================================================================================//
//  Description:
//      Get the next Windows Firewall Authorised service using the
//      specified enumerator
//
//  Parameters:
//      pIEnumVariant  - the enumeration interface
//      pAuthorisedSvc - receives the authorised application name
//
//  Returns:
//      True if fetched an item otherwise false because at end
//      of the enumeration
//===============================================================================================//
bool WindowsFirewallInformation::GetNextFwAuthorisedSvc( IEnumVARIANT* pIEnumVariant,
                                                         String* pAuthorisedSvc )
{
    BSTR      name        = nullptr;
    ULONG     CeltFetched = 0;
    HRESULT   hResult     = 0;
    VARIANT   variant;
    String    VarType;
    Formatter Format;
    INetFwService*      pINetFwService = nullptr;
    AutoIUnknownRelease ReleaseINetFwService;

    if ( pIEnumVariant == nullptr )
    {
        throw ParameterException( L"pIEnumVariant", __FUNCTION__ );
    }

    if ( pAuthorisedSvc == nullptr )
    {
        throw ParameterException( L"pAuthorisedSvc", __FUNCTION__ );
    }
    *pAuthorisedSvc = PXS_STRING_EMPTY;

    // Get the next application
    VariantInit( &variant );
    hResult = pIEnumVariant->Next( 1, &variant, &CeltFetched );
    if ( hResult == S_FALSE )       // S_FALSE means did not get the requested
    {                               // number of elements, i.e one
        return false;
    }
    AutoVariantClear VariantClearVar( &variant );

    // The returned variant should point to a dispatch interface
    if ( ( variant.pdispVal == nullptr ) || ( variant.vt != VT_DISPATCH ) )
    {
        VarType = Format.StringInt32( L"VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, VarType.c_str(), __FUNCTION__ );
    }

    // Get the application pointed to by the dispatch interface
    hResult = variant.pdispVal->QueryInterface( IID_INetFwService,
                                                reinterpret_cast<void**>(&pINetFwService) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IDispatch::QueryInterface", __FUNCTION__ );
    }

    // Ensure have the pointer
    if ( pINetFwService == nullptr )
    {
        throw NullException( L"pINetFwService", __FUNCTION__ );
    }
    ReleaseINetFwService.Set( pINetFwService );

    // Get the service's name
    hResult = pINetFwService->get_Name( &name );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwService::get_Name", __FUNCTION__ );
    }
    AutoSysFreeString AutoSysFreeName( name );
    *pAuthorisedSvc = name;
    pAuthorisedSvc->Trim();

    return true;
}

//===============================================================================================//
//  Description:
//      Make the data records the applications authorised by the
//      Windows Firewall
//
//  Parameters:
//      pINetFwProfile - pointer to the firewall policy interface
//      pRecords       - receives the records
//
//  Returns:
//     void
//===============================================================================================//
void WindowsFirewallInformation::MakeFwAuthorisedAppsRecords( INetFwProfile* pINetFwProfile,
                                                              TArray< AuditRecord >* pRecords )
{
    String     AuthorisedApp;
    HRESULT    hResult  = 0;
    AuditRecord Record;
    IUnknown*    pNewEnum      = nullptr;
    IEnumVARIANT* pEnumVariant = nullptr;
    INetFwAuthorizedApplications* pApplications = nullptr;
    AutoIUnknownRelease ReleaseApplications, ReleaseNewEnum;
    AutoIUnknownRelease ReleaseEnumVariant;

    if ( pINetFwProfile == nullptr )
    {
        throw ParameterException( L"pINetFwProfile", __FUNCTION__ );
    }

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Retrieve the authorized application collection.
    hResult = pINetFwProfile->get_AuthorizedApplications( &pApplications );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwProfile::get_AuthorizedApplications", __FUNCTION__ );
    }

    if ( pApplications == nullptr )
    {
        throw NullException( L"pApplications", __FUNCTION__ );
    }
    ReleaseApplications.Set( pApplications );

    hResult = pApplications->get__NewEnum( &pNewEnum );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwAuthorizedApplications::get__NewEnum", __FUNCTION__ );
    }

    if ( pNewEnum == nullptr )
    {
        throw NullException( L"pNewEnum", __FUNCTION__ );
    }
    ReleaseNewEnum.Set( pNewEnum );

    hResult = pNewEnum->QueryInterface( IID_IEnumVARIANT,
                                        reinterpret_cast<void**>( &pEnumVariant ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IID_IEnumVARIANT", "IUnknown::QueryInterface" );
    }

    if ( pEnumVariant == nullptr )
    {
        throw NullException( L"pEnumVariant", __FUNCTION__ );
    }
    ReleaseEnumVariant.Set( pEnumVariant );

    while ( GetNextFwAuthorisedApp( pEnumVariant, &AuthorisedApp ) )
    {
        // Make the audit record
        Record.Reset( PXS_CATEGORY_WINDOWS_FIREWALL );
        Record.Add( PXS_WINDOWS_FIREWALL_NAME, L"Authorised Application" );
        Record.Add( PXS_WINDOWS_FIREWALL_SETTING, AuthorisedApp );
        pRecords->Add( Record );

        AuthorisedApp = PXS_STRING_EMPTY;   // Next
    }
}

//===============================================================================================//
//  Description:
//      Make the data records of the network ports authorised by the
//      Windows Firewall
//
//  Parameters:
//      pINetFwProfile - pointer to the firewall policy interface
//      pRecords       - receives the records
//
//  Returns:
//     void
//===============================================================================================//
void WindowsFirewallInformation::MakeFwAuthorisedPortsRecords( INetFwProfile* pINetFwProfile,
                                                               TArray< AuditRecord >* pRecords )
{
    String      AuthorisedPort;
    HRESULT     hResult = 0;
    AuditRecord Record;
    IUnknown*         pINewEnum        = nullptr;
    IEnumVARIANT*     pIEnumVariant    = nullptr;
    INetFwOpenPorts*  pINetFwOpenPorts = nullptr;
    AutoIUnknownRelease ReleaseINetFwOpenPorts, ReleaseINewEnum;
    AutoIUnknownRelease ReleaseIEnumVariant;

    if ( pINetFwProfile == nullptr )
    {
        throw ParameterException( L"pINetFwProfile", __FUNCTION__ );
    }

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Retrieve the authorized application collection.
    hResult = pINetFwProfile->get_GloballyOpenPorts( &pINetFwOpenPorts );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwProfile::get_GloballyOpenPorts", __FUNCTION__ );
    }

    if ( pINetFwOpenPorts == nullptr )
    {
        throw NullException( L"pINetFwOpenPorts", __FUNCTION__ );
    }
    ReleaseINetFwOpenPorts.Set( pINetFwOpenPorts );

    hResult = pINetFwOpenPorts->get__NewEnum( &pINewEnum );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwOpenPorts::get__NewEnum", __FUNCTION__ );
    }

    if ( pINewEnum == nullptr )
    {
        throw NullException( L"pINewEnum", __FUNCTION__ );
    }
    ReleaseINewEnum.Set( pINewEnum );

    hResult = pINewEnum->QueryInterface( IID_IEnumVARIANT,
                                         reinterpret_cast<void**>( &pIEnumVariant ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IID_IEnumVARIANT", "IUnknown::QueryInterface" );
    }

    if ( pIEnumVariant == nullptr )
    {
        throw NullException( L"pNewEnum", __FUNCTION__ );
    }
    ReleaseIEnumVariant.Set( pIEnumVariant );

    while ( GetNextFwAuthorisedPort( pIEnumVariant, &AuthorisedPort ) )
    {
        // Make the record
        Record.Reset( PXS_CATEGORY_WINDOWS_FIREWALL );
        Record.Add( PXS_WINDOWS_FIREWALL_NAME, L"Authorised Port" );
        Record.Add( PXS_WINDOWS_FIREWALL_SETTING, AuthorisedPort );
        pRecords->Add( Record );

        AuthorisedPort = PXS_STRING_EMPTY;  // Next
    }
}

//===============================================================================================//
//  Description:
//      Make the data records the services authorised by the
//      Windows Firewall
//
//  Parameters:
//      pINetFwProfile  - pointer to the firewall policy interface
//      pRecords        - receives the records
//
//  Returns:
//     void
//===============================================================================================//
void WindowsFirewallInformation::MakeFwAuthorisedSvcsRecords( INetFwProfile* pINetFwProfile,
                                                              TArray< AuditRecord >* pRecords )
{
    String      AuthorisedSvc;
    HRESULT     hResult   = 0;
    AuditRecord Record;
    IUnknown*       pINewEnum       = nullptr;
    IEnumVARIANT*   pIEnumVariant   = nullptr;
    INetFwServices* pINetFwServices = nullptr;
    AutoIUnknownRelease ReleaseINetFwServices, ReleaseINewEnum;
    AutoIUnknownRelease ReleaseIEnumVariant;

    if ( pINetFwProfile == nullptr )
    {
        throw ParameterException( L"pINetFwProfile", __FUNCTION__ );
    }

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Retrieve the authorized services collection.
    hResult = pINetFwProfile->get_Services( &pINetFwServices );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwProfile::get_Services", __FUNCTION__ );
    }

    if ( pINetFwServices == nullptr )
    {
        throw NullException( L"pINetFwServices", __FUNCTION__ );
    }
    ReleaseINetFwServices.Set( pINetFwServices );

    hResult = pINetFwServices->get__NewEnum( &pINewEnum );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwServices::get__NewEnum", __FUNCTION__ );
    }

    if ( pINewEnum == nullptr )
    {
        throw NullException( L"pINewEnum", __FUNCTION__ );
    }
    ReleaseINewEnum.Set( pINewEnum );

    hResult = pINewEnum->QueryInterface( IID_IEnumVARIANT,
                                         reinterpret_cast<void**>( &pIEnumVariant ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IID_IEnumVARIANT", "IUnknown::QueryInterface" );
    }

    if ( pIEnumVariant == nullptr )
    {
        throw NullException( L"pIEnumVariant", __FUNCTION__ );
    }
    ReleaseIEnumVariant.Set( pIEnumVariant );

    while ( GetNextFwAuthorisedSvc( pIEnumVariant, &AuthorisedSvc ) )
    {
        // Make the record
        Record.Reset( PXS_CATEGORY_WINDOWS_FIREWALL );
        Record.Add( PXS_WINDOWS_FIREWALL_NAME, L"Authorised Service" );
        Record.Add( PXS_WINDOWS_FIREWALL_SETTING, AuthorisedSvc );
        pRecords->Add( Record );

        AuthorisedSvc = PXS_STRING_EMPTY;   // Next
    }
}

//===============================================================================================//
//  Description:
//      Make the data records the Windows Firewall status
//
//  Parameters:
//      pINetFwProfile - pointer to the firewall policy interface
//      pRecord        - receives the formatted record
//
//  Remarks:
//
//  Returns:
//     void
//===============================================================================================//
void WindowsFirewallInformation::MakeFwStatusRecord( INetFwProfile* pINetFwProfile,
                                                     AuditRecord* pRecord )
{
    HRESULT   hResult = 0;
    String    Value;
    Formatter Format;
    VARIANT_BOOL enabled = FALSE;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_WINDOWS_FIREWALL );

    if ( pINetFwProfile == nullptr )
    {
        throw ParameterException( L"pINetFwProfile", __FUNCTION__ );
    }

    hResult = pINetFwProfile->get_FirewallEnabled( &enabled );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwProfile::get_FirewallEnabled", __FUNCTION__ );
    }
    Value = Format.Int32YesNo( enabled );      // VARIANT_BOOL = short
    pRecord->Add( PXS_WINDOWS_FIREWALL_NAME, L"Firewall Enabled" );
    pRecord->Add( PXS_WINDOWS_FIREWALL_SETTING, Value );
}

//===============================================================================================//
//  Description:
//      Make a name for the specified open port
//
//  Parameters:
//      pINetFwOpenPort - pointer to the open port's interface
//      pAuthorisedPort - receives the authorised port's name
//
//  Returns:
//      void
//===============================================================================================//
void WindowsFirewallInformation::MakeOpenPortName( INetFwOpenPort* pINetFwOpenPort,
                                                   String* pAuthorisedPort )
{
    BSTR      remoteAddrs = nullptr;
    LONG      portNumber  = 0;
    String    RemoteAddress;
    HRESULT   hResult     = 0;
    Formatter Format;
    NET_FW_IP_PROTOCOL ipProtocol;

    if ( pINetFwOpenPort == nullptr )
    {
        throw ParameterException( L"pINetFwOpenPort", __FUNCTION__ );
    }

    if ( pAuthorisedPort == nullptr )
    {
        throw ParameterException( L"pAuthorisedPort", __FUNCTION__ );
    }
    *pAuthorisedPort = PXS_STRING_EMPTY;

    // Get the port's number
    hResult = pINetFwOpenPort->get_Port( &portNumber );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwOpenPort::get_Port", __FUNCTION__ );
    }

    // Get the port's protocol
    hResult = pINetFwOpenPort->get_Protocol( &ipProtocol );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"INetFwOpenPort::get_Protocol", __FUNCTION__ );
    }

    // Get the port's remote address, UDP does not have one
    if ( ipProtocol != NET_FW_IP_PROTOCOL_UDP )
    {
        // Get the port's remote address
        hResult = pINetFwOpenPort->get_RemoteAddresses( &remoteAddrs );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"INetFwOpenPort::get_RemoteAddresses", __FUNCTION__ );
        }
        AutoSysFreeString AutoSysFreeRemoteAddrs( remoteAddrs );
        RemoteAddress = remoteAddrs;
    }
    pAuthorisedPort->Allocate( 32 );
    TranslateIpProtocol( ipProtocol, pAuthorisedPort );

    // Port Number[:Remote address]
    *pAuthorisedPort += PXS_CHAR_COLON;
    *pAuthorisedPort += Format.Int32( portNumber );
    if ( RemoteAddress.GetLength() )
    {
        *pAuthorisedPort += PXS_CHAR_COLON;
        *pAuthorisedPort += RemoteAddress;
    }
}

//===============================================================================================//
//  Description:
//      Translate the specified enumeration constant of an IP protocol to
//      a name
//
//  Parameters:
//      protocol    - the protocol, NET_FW_IP_PROTOCOL see
//      pIpProtocol - receives the protocol name
//
//  Returns:
//      void
//===============================================================================================//
void WindowsFirewallInformation::TranslateIpProtocol( NET_FW_IP_PROTOCOL protocol,
                                                      String* pIpProtocol )
{
    Formatter Format;

    if ( pIpProtocol == nullptr )
    {
        throw ParameterException( L"pIpProtocol", __FUNCTION__ );
    }

    switch ( protocol )
    {
        default:
            *pIpProtocol = Format.UInt32( protocol );
            break;

        case NET_FW_IP_PROTOCOL_TCP:
            *pIpProtocol = L"TCP";
            break;

        case NET_FW_IP_PROTOCOL_UDP:
            *pIpProtocol = L"UDP";
            break;

        case NET_FW_IP_PROTOCOL_ANY:
            *pIpProtocol = L"ANY";
            break;
    }
}
