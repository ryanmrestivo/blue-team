///////////////////////////////////////////////////////////////////////////////////////////////////
//
// NT Service Information Class Implementation
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
#include "WinAudit/Header Files/NtServiceInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/SystemInformation.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
NtServiceInformation::NtServiceInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
NtServiceInformation::~NtServiceInformation()
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
//      Get services data as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void NtServiceInformation::GetServiceRecords( TArray< AuditRecord >* pRecords )
{
    BOOL    success = FALSE;
    DWORD   i = 0, bytesNeeded = 0, servicesReturned = 0;
    DWORD   resumeHandle = 0, lastError = 0, bytesESS = 0;
    String  Name, ServiceType, CurrentState, StartType;
    String  PathName, StartName, Insert1;
    SC_HANDLE     hSCManager = nullptr;
    AuditRecord   Record;
    AllocateBytes AllocBytes;
    ENUM_SERVICE_STATUS* lpESS = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Allocate a buffer to hold an array of ENUM_SERVICE_STATUS. Each
    // structure has two strings of 256 bytes. Will make it big enough for
    // 100 services. The maximum allowed buffer size is 256K bytes
    bytesESS = 100 * ( sizeof ( ENUM_SERVICE_STATUS ) + 512 );
    lpESS = reinterpret_cast<ENUM_SERVICE_STATUS*>( AllocBytes.New(bytesESS ));

    hSCManager = OpenSCManager( nullptr,     // Local machine
                                nullptr,     // Default database
                                SC_MANAGER_ENUMERATE_SERVICE );
    if ( hSCManager == nullptr )
    {
        throw SystemException( GetLastError(), L"OpenSCManager", __FUNCTION__ );
    }

    // Must catch all errors to free handle
    try
    {
        resumeHandle = 0;     // Must zero on the first call
        do
        {
            servicesReturned = 0;
            memset( lpESS, 0, bytesESS );
            success = EnumServicesStatus( hSCManager,
                                          SERVICE_WIN32 | SERVICE_DRIVER,
                                          SERVICE_STATE_ALL,     // All services
                                          lpESS,
                                          bytesESS,
                                          &bytesNeeded, &servicesReturned, &resumeHandle );
            // Only continue on failure if the error is buffer too small
            lastError = GetLastError();
            if ( ( success == 0 ) && ( lastError != ERROR_MORE_DATA ) )
            {
                // Ensure servicesReturned is zero
                servicesReturned = 0;
                Insert1.SetAnsi( __FUNCTION__ );
                PXSLogSysError1( lastError, L"EnumServicesStatus", Insert1 );
            }

            for ( i = 0; i < servicesReturned; i++ )
            {
                // If no display name use the service name
                Name = lpESS[ i ].lpDisplayName;
                Name.Trim();
                if ( Name.IsEmpty() )
                {
                    Name = lpESS[ i ].lpServiceName;
                    Name.Trim();
                }

                Record.Reset( PXS_CATEGORY_NTSERVICES );
                Record.Add( PXS_NTSERVICES_SERVICE_NAME, Name );

                ServiceType = PXS_STRING_EMPTY;
                TranslateServiceType( lpESS[ i ].ServiceStatus.dwServiceType,
                                      &ServiceType );
                Record.Add( PXS_NTSERVICES_TYPE, ServiceType );

                CurrentState = PXS_STRING_EMPTY;
                TranslateCurrentState( lpESS[ i ].ServiceStatus.dwCurrentState, &CurrentState );
                Record.Add( PXS_NTSERVICES_STATE, CurrentState );

                StartType = PXS_STRING_EMPTY;
                PathName  = PXS_STRING_EMPTY;
                StartName = PXS_STRING_EMPTY;
                GetServiceConfig( hSCManager,
                                  lpESS[ i ].lpServiceName, &StartType, &PathName, &StartName );

                Record.Add( PXS_NTSERVICES_START_MODE, StartType );
                Record.Add( PXS_NTSERVICES_PATH_NAME , PathName  );
                Record.Add( PXS_NTSERVICES_START_NAME, StartName );

                pRecords->Add( Record );
            }
        } while ( ( resumeHandle     != 0 ) &&    // Should not be zero
                  ( servicesReturned != 0 ) &&    // Ensure got at least 1
                  ( lastError == ERROR_MORE_DATA ) );
    }
    catch ( const Exception& )
    {
        CloseServiceHandle( hSCManager );
        throw;
    }
    CloseServiceHandle( hSCManager );
    PXSSortAuditRecords( pRecords, PXS_NTSERVICES_SERVICE_NAME );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get configuration settings for the specified service
//
//  Parameters:
//      pszServiceName - pointer to the service name
//      pStartType     - receives the service type
//      pPathName      - receives the path to the service exe
//      pStartName     - receives the start name
//
//  Remarks:
//      See QUERY_SERVICE_CONFIG
//
//  Returns:
//      void
//===============================================================================================//
void NtServiceInformation::GetServiceConfig( SC_HANDLE hSCManager,
                                             LPCWSTR pszServiceName,
                                             String* pStartType,
                                             String* pPathName, String* pStartName )
{
    DWORD         lastError = 0, bufSize = 0, bytesNeeded = 0;
    Formatter     Format;
    String        Insert2, ServiceName;
    SC_HANDLE     hService = nullptr;
    AllocateBytes AllocBytes;
    QUERY_SERVICE_CONFIG* pQSC = nullptr;

    SERVICE_DELAYED_AUTO_START_INFO sdasi;
    SystemInformation SystemInfo;

    if ( ( pStartType == nullptr ) ||
         ( pPathName  == nullptr ) ||
         ( pStartName == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pStartType = PXS_STRING_EMPTY;
    *pPathName  = PXS_STRING_EMPTY;
    *pStartName = PXS_STRING_EMPTY;

    if ( hSCManager == nullptr )
    {
        throw ParameterException( L"hSCManager", __FUNCTION__ );
    }

    ServiceName = pszServiceName;
    ServiceName.Trim();
    if ( ServiceName.IsEmpty() )
    {
        throw ParameterException( L"pszServiceName", __FUNCTION__ );
    }

    hService = OpenService( hSCManager, ServiceName.c_str(), SERVICE_QUERY_CONFIG );
    if ( hService == nullptr )
    {
        throw SystemException( GetLastError(), ServiceName.c_str(), "OpenService" );
    }

    // Get the required configuration buffer size
    if ( QueryServiceConfig( hService, nullptr, 0, &bytesNeeded ) )
    {
       // Unexpected success with null buffer
       CloseServiceHandle( hService );
       Insert2.SetAnsi( __FUNCTION__ );
       PXSLogAppInfo2( L"QueryServiceConfig said no data for '%%1' in '%%2'.",
                       ServiceName, Insert2 );
       return;
    }

    // Test for expected error
    lastError = GetLastError();
    if ( lastError != ERROR_INSUFFICIENT_BUFFER )
    {
        CloseServiceHandle( hService );
        throw SystemException( lastError,
                               ServiceName.c_str(), "QueryServiceConfig" );
    }

    if ( bytesNeeded == 0 )
    {
       CloseServiceHandle( hService );
       Insert2.SetAnsi( __FUNCTION__ );
       PXSLogAppInfo2( L"QueryServiceConfig said no data for '%%1' in '%%2'.",
                       ServiceName, Insert2 );
       return;
    }

    // Second call, double the amount
    bufSize = PXSMultiplyUInt32( bytesNeeded, 2 );
    pQSC = reinterpret_cast<QUERY_SERVICE_CONFIG*>( AllocBytes.New(bufSize) );
    if ( QueryServiceConfig( hService, pQSC, bufSize, &bytesNeeded ) == 0 )
    {
        CloseServiceHandle( hService );
        throw SystemException( lastError, ServiceName.c_str(), "QueryServiceConfig" );
    }

    // Catch exceptions to free the service handle
    try
    {
        TranslateStartType( pQSC->dwStartType, pStartType );
        *pPathName  = pQSC->lpBinaryPathName;
        *pStartName = pQSC->lpServiceStartName;
    }
    catch ( const Exception& )
    {
        CloseServiceHandle( hService );
        throw;
    }

    // For Vista and newer, determine if the auto start services are delayed.
    // Note, according to MSDN any service can marked as delayed start however,
    // this setting has no effect unless the service is an auto-start service.
    // Since just reporting whatever the configuration is, will test all
    // services rather than those that that are SERVICE_AUTO_START.
    try
    {
        if ( SystemInfo.GetMajorVersion() >= 6 )
        {
            memset( &sdasi, 0, sizeof ( sdasi ) );
            bytesNeeded = 0;
            if ( QueryServiceConfig2( hService,
                                      SERVICE_CONFIG_DELAYED_AUTO_START_INFO,
                                      reinterpret_cast<BYTE*>( &sdasi ),
                                      sizeof ( sdasi ), &bytesNeeded ) )
            {
                if ( sdasi.fDelayedAutostart )
                {
                    *pStartType += L" (Delayed Start)";
                }
            }
            else
            {
                // Continue on error
                PXSLogSysError( GetLastError(), L"QueryServiceConfig2 failed" );
            }
        }
    }
    catch ( const Exception& )
    {
        CloseServiceHandle( hService );
        throw;
    }
    CloseServiceHandle( hService );
}

//===============================================================================================//
//  Description:
//      Translate service's current state into a string
//
//  Parameters:
//      currentState  - current state of service
//      pCurrentState - string object to receive the state
//
//  Returns:
//      void
//===============================================================================================//
void NtServiceInformation::TranslateCurrentState( DWORD currentState, String* pCurrentState )
{
    Formatter Format;

    if ( pCurrentState == nullptr )
    {
        throw ParameterException( L"pCurrentState", __FUNCTION__ );
    }
    *pCurrentState = PXS_STRING_EMPTY;

    if ( currentState == SERVICE_STOPPED )
    {
        *pCurrentState = L"Stopped";
    }
    else if ( currentState == SERVICE_START_PENDING )
    {
        *pCurrentState = L"Start Pending";
    }
    else if ( currentState == SERVICE_STOP_PENDING )
    {
        *pCurrentState = L"Stop Pending";
    }
    else if ( currentState == SERVICE_RUNNING )
    {
        *pCurrentState = L"Running";
    }
    else if ( currentState == SERVICE_CONTINUE_PENDING )
    {
        *pCurrentState = L"Continue Pending";
    }
    else if ( currentState == SERVICE_PAUSE_PENDING )
    {
        *pCurrentState = L"Pause Pending";
    }
    else if ( currentState == SERVICE_PAUSED )
    {
        *pCurrentState = L"Paused";
    }
    else
    {
        *pCurrentState = Format.UInt32( currentState );
    }
}

//===============================================================================================//
//  Description:
//      Translate a service type into a string
//
//  Parameters:
//      serviceType  - type of service
//      pServiceType - string object to receive type
//
//  Returns:
//      void
//===============================================================================================//
void NtServiceInformation::TranslateServiceType( DWORD serviceType, String* pServiceType )
{
    Formatter Format;

    if ( pServiceType == nullptr )
    {
        throw ParameterException( L"pServiceType", __FUNCTION__ );
    }

    if ( serviceType & SERVICE_KERNEL_DRIVER )
    {
        *pServiceType = L"Kernel Driver";
    }
    else if ( serviceType & SERVICE_FILE_SYSTEM_DRIVER )
    {
        *pServiceType = L"System Driver";
    }
    else if ( serviceType & SERVICE_ADAPTER )
    {
        *pServiceType = L"Adapter";
    }
    else if ( serviceType & SERVICE_RECOGNIZER_DRIVER )
    {
        *pServiceType = L"Recognizer Driver";
    }
    else if ( serviceType & SERVICE_WIN32_OWN_PROCESS )
    {
        *pServiceType = L"Own Process";
    }
    else if ( serviceType & SERVICE_WIN32_SHARE_PROCESS )
    {
        *pServiceType = L"Shared Process";
    }
    else if ( serviceType & SERVICE_INTERACTIVE_PROCESS )
    {
        *pServiceType = L"Interactive Process";
    }
    else
    {
        *pServiceType = Format.UInt32( serviceType );
    }
}

//===============================================================================================//
//  Description:
//      Translate a service's start type into a string
//
//  Parameters:
//      startType  - configured start type
//      pStartType - string object to receive start type
//
//  Returns:
//      void
//===============================================================================================//
void NtServiceInformation::TranslateStartType( DWORD startType, String* pStartType )
{
    Formatter Format;

    if ( pStartType == nullptr )
    {
        throw ParameterException( L"pStartType", __FUNCTION__ );
    }
    *pStartType = PXS_STRING_EMPTY;

    if ( startType == SERVICE_BOOT_START )
    {
        *pStartType = L"Boot";
    }
    else if ( startType == SERVICE_SYSTEM_START )
    {
        *pStartType = L"System";
    }
    else if ( startType == SERVICE_AUTO_START )
    {
        *pStartType = L"Automatic";
    }
    else if ( startType == SERVICE_DEMAND_START )
    {
        *pStartType = L"Manual";
    }
    else if ( startType == SERVICE_DISABLED )
    {
        *pStartType = L"Disabled";
    }
    else
    {
        *pStartType = Format.UInt32( startType );
    }
}
