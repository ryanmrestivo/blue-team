///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Kerberos Ticket Information Class Implementation
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
#include "WinAudit/Header Files/KerberosTicketInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/Ddk.h"
#include "WinAudit/Header Files/SecurityInformation.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
KerberosTicketInformation::KerberosTicketInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
KerberosTicketInformation::~KerberosTicketInformation()
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
//      Get the Kerberos Policy as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void KerberosTicketInformation::GetKerberosPolicyRecords(
                                              TArray< AuditRecord >* pRecords )
{
    size_t      i = 0;
    String      DataString;
    AuditRecord Record;
    SystemInformation   SystemInfo;
    WindowsInformation  WindowsInfo;
    SecurityInformation SecurityInfo;

    // Structure to hold desired data points
    struct _STRUCT_KERB_DATA
    {
        LPCWSTR pszPolicy;      // Used for group policy
        LPCWSTR pszKeyName;     // Used for RSOP
        LPCWSTR pszUnit1;       // Used for RSOP
        LPCWSTR pszUnit2;       // Used for RSOP
        VARTYPE varType;        // RSOP data type (VT_BOOL, VT_I4 or VT_BSTR)
    }  Data[] = { { L"Enforce user logon restrictions",
                    L"TicketValidateClient",
                    PXS_STRING_EMPTY,
                    PXS_STRING_EMPTY,
                    VT_BOOL },
                  { L"Maximum lifetime for service ticket",
                    L"MaxServiceAge",
                    L" Minute",
                    L" Minutes",
                    VT_I4 },
                  { L"Maximum lifetime for user ticket",
                    L"MaxTicketAge",
                    L" Hour",
                    L" Hours",
                    VT_I4},
                  { L"Maximum lifetime for user ticket renewal",
                    L"MaxRenewAge",
                    L" Day",
                    L" Days",
                    VT_I4},
                  { L"Maximum tolerance for computer clock synchronization",
                    L"MaxClockSkew",
                    L" Minute",
                    L" Minutes",
                    VT_I4} };

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Cannot proceed with out administrator privileges
    if ( SystemInfo.IsAdministrator() == false )
    {
        throw Exception( PXS_ERROR_TYPE_SYSTEM,
                         ERROR_ACCESS_DENIED, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    for ( i = 0; i < ARRAYSIZE( Data ); i++ )
    {
        DataString = PXS_STRING_EMPTY;
        SecurityInfo.GetRSOPSecuritySetting( Data[ i ].pszKeyName,
                                             Data[ i ].varType, &DataString );

        // For visual display append the unit if its a numerical value
        if ( ( DataString.GetLength() ) && ( Data[ i ].varType == VT_I4 ) )
        {
            if ( DataString.CompareI( PXS_STRING_ONE ) == 0 )
            {
                DataString += Data[ i ].pszUnit1;
            }
            else
            {
                DataString += Data[ i ].pszUnit2;
            }
        }

        Record.Reset( PXS_CATEGORY_KERBEROS_POLICY );
        Record.Add( PXS_KERBEROS_POLICY_POLICY,  Data[ i ].pszPolicy );
        Record.Add( PXS_KERBEROS_POLICY_SETTING, DataString );
        pRecords->Add( Record );
    }
}

//===============================================================================================//
//  Description:
//      Get the Kerberos Tickets for the logged on user as an array of
//      audit records
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void KerberosTicketInformation::GetKerberosTicketRecords(
                                              TArray< AuditRecord >* pRecords )
{
    char        szKerberosNameA[ 32 ] = { 0 };  // "MICROSOFT_KERBEROS_NAME_A"
    size_t      lenBytes = 0;
    ULONG       authenticationPackage = 0, returnBufferLength = 0;
    HANDLE      hLsaHandle = nullptr;
    String      Insert1;
    Formatter   Format;
    NTSTATUS    status, protocolStatus;
    LSA_STRING  lsaPackageName;
    KERB_QUERY_TKT_CACHE_REQUEST   kqtcr;
    KERB_QUERY_TKT_CACHE_RESPONSE* pCacheResponse = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // As only want to query for information will use an untrusted connection
    status = LsaConnectUntrusted( &hLsaHandle );
    if ( STATUS_SUCCESS != status )
    {
        throw SystemException( LsaNtStatusToWinError( status ),
                               L"LsaConnectUntrusted", __FUNCTION__ );
    }

    // Using ANSI as LsaLookupAuthenticationPackage takes LSA_STRING which
    // is single byte
    StringCchCopyA( szKerberosNameA,
                    ARRAYSIZE( szKerberosNameA ),
                    MICROSOFT_KERBEROS_NAME_A  );  // = "Kerberos" in NTSecAPI.h

    lenBytes = strlen( szKerberosNameA );
    lsaPackageName.Length        = PXSCastSizeTToUInt16( lenBytes );
    lsaPackageName.MaximumLength = sizeof ( szKerberosNameA );
    lsaPackageName.Buffer        = szKerberosNameA;
    authenticationPackage        = 0;
    status = LsaLookupAuthenticationPackage( hLsaHandle,
                                             &lsaPackageName, &authenticationPackage );
    if ( STATUS_SUCCESS != status )
    {
        LsaDeregisterLogonProcess( hLsaHandle );
        throw SystemException( LsaNtStatusToWinError( status ),
                                 L"LsaLookupAuthenticationPackage", __FUNCTION__ );
    }

    memset( &kqtcr, 0, sizeof ( kqtcr ) );
    kqtcr.MessageType      = KerbQueryTicketCacheMessage;
    kqtcr.LogonId.LowPart  = 0;             // zero implies current user
    kqtcr.LogonId.HighPart = 0;             // zero implies current user
    protocolStatus         = STATUS_SUCCESS;
    status = LsaCallAuthenticationPackage(
                                   hLsaHandle,
                                   authenticationPackage,
                                   &kqtcr,
                                   sizeof ( kqtcr ),
                                   reinterpret_cast<void**>( &pCacheResponse ),
                                   &returnBufferLength, &protocolStatus );
    if ( STATUS_SUCCESS != status )
    {
        LsaDeregisterLogonProcess( hLsaHandle );
        throw SystemException( LsaNtStatusToWinError( status ),
                               L"LsaCallAuthenticationPackage", __FUNCTION__ );
    }

    if ( pCacheResponse == nullptr )
    {
        LsaDeregisterLogonProcess( hLsaHandle );
        PXSLogAppWarn( L"LsaCallAuthenticationPackage returned no data." );
        return;
    }

    if ( STATUS_SUCCESS != protocolStatus )
    {
        LsaFreeReturnBuffer( pCacheResponse );
        LsaDeregisterLogonProcess( hLsaHandle );
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogNtStatusError1( protocolStatus,
                              L"LsaCallAuthenticationPackage failed in '%%1'.", Insert1 );
        return;
    }

    // Catch as need to clean up
    try
    {
        MakeAuditRecords( pCacheResponse, pRecords );
    }
    catch ( const Exception& )
    {
        LsaFreeReturnBuffer( pCacheResponse );
        LsaDeregisterLogonProcess( hLsaHandle );
        throw;
    }
    LsaFreeReturnBuffer( pCacheResponse );
    LsaDeregisterLogonProcess( hLsaHandle );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Make formatted data string records from the specified Kerberos
//      tickets
//
//  Parameters:
//      pCR      - KERB_QUERY_TKT_CACHE_RESPONSE structure
//      pRecords - receives the formatted data strings
//
//  Returns:
//      void
//===============================================================================================//
void KerberosTicketInformation::MakeAuditRecords( KERB_QUERY_TKT_CACHE_RESPONSE* pCR,
                                                  TArray< AuditRecord >* pRecords )
{
    ULONG       i = 0;
    String      Value;
    FILETIME    ft;
    Formatter   Format;
    AuditRecord Record;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    if ( pCR == nullptr )
    {
        return;     // No tickets so nothing to do
    }

    for ( i = 0; i < pCR->CountOfTickets; i++ )
    {
        Record.Reset( PXS_CATEGORY_KERBEROS_TICKETS );

        Value = Format.UnicodeString( &pCR->Tickets[ i ].ServerName );
        Record.Add( PXS_KERBEROS_TICKS_SERVER_NAME, Value );

        Value = Format.UnicodeString( &pCR->Tickets[ i ].RealmName );
        Record.Add( PXS_KERBEROS_TICKS_REALM_NAME, Value );

        // PXS_KERBEROS_TICKS_START_TIME is optional data
        Value = PXS_STRING_EMPTY;
        if ( ( pCR->Tickets[ i ].StartTime.LowPart  > 0 ) ||
             ( pCR->Tickets[ i ].StartTime.HighPart > 0 )  )
        {
            memset( &ft, 0, sizeof ( ft ) );
            ft.dwLowDateTime  = pCR->Tickets[ i ].StartTime.LowPart;
            ft.dwHighDateTime = PXSCastInt32ToUInt32(
                                        pCR->Tickets[ i ].StartTime.HighPart );
            Value = Format.FileTimeToLocalTimeUser( ft, true );
        }
        Record.Add( PXS_KERBEROS_TICKS_START_TIME, Value );

        Value = PXS_STRING_EMPTY;
        if ( ( pCR->Tickets[ i ].EndTime.LowPart  > 0 ) ||
             ( pCR->Tickets[ i ].EndTime.HighPart > 0 )  )
        {
            memset( &ft, 0, sizeof ( ft ) );
            ft.dwLowDateTime  = pCR->Tickets[ i ].EndTime.LowPart;
            ft.dwHighDateTime = PXSCastInt32ToUInt32(
                                          pCR->Tickets[ i ].EndTime.HighPart );
            Value = Format.FileTimeToLocalTimeUser( ft, true );
        }
        Record.Add( PXS_KERBEROS_TICKS_END_TIME, Value );

        // PXS_KERBEROS_TICKS_RENEW_TIME is only valid if
        // KERB_TICKET_FLAGS_renewable is set
        Value = PXS_STRING_EMPTY;
        if ( KERB_TICKET_FLAGS_renewable & pCR->Tickets[ i ].TicketFlags )
        {
            if ( ( pCR->Tickets[ i ].RenewTime.LowPart  > 0 ) ||
                 ( pCR->Tickets[ i ].RenewTime.HighPart > 0 )  )
            {
                memset( &ft, 0, sizeof ( ft ) );
                ft.dwLowDateTime  = pCR->Tickets[ i ].RenewTime.LowPart;
                ft.dwHighDateTime = PXSCastInt32ToUInt32(
                                        pCR->Tickets[ i ].RenewTime.HighPart );
                Value = Format.FileTimeToLocalTimeUser( ft, true );
            }
        }
        Record.Add( PXS_KERBEROS_TICKS_RENEW_TIME, Value );

        Value = PXS_STRING_EMPTY;
        TranslateEncryptionType( pCR->Tickets[ i ].EncryptionType, &Value );

        Record.Add( PXS_KERBEROS_TICKS_ENCRYPT_TYPE, Value );

        Value = PXS_STRING_EMPTY;
        TranslateFlags( pCR->Tickets[ i ].TicketFlags, &Value );
        Record.Add( PXS_KERBEROS_TICKS_TICKET_FLAGS, Value );

        pRecords->Add( Record );
    }
}

//===============================================================================================//
//  Description:
//      Translate defined Kerberos encryption type
//
//  Parameters:
//      type            - the encryption type
//      pEncryptionType - string object to receive the translation
//
//  Remarks:
//        See KERB_RETRIEVE_TKT_REQUEST Structure
//
//  Returns:
//      void
//===============================================================================================//
void KerberosTicketInformation::TranslateEncryptionType( LONG type,
                                                         String* pEncryptionType )
{
    size_t    i = 0;
    String    Insert2;
    Formatter Format;

    // Structure to hold encryption types
    struct _STRUCT_TYPES
    {
        LONG    type;
        LPCWSTR pszType;
    } Types[] =
        { { KERB_ETYPE_NULL       , L"None"                               },
          { KERB_ETYPE_DES_CBC_CRC, L"DES in cipher-block-chaining mode "
                                    L"with a CRC-32 checksum"              },
          { KERB_ETYPE_DES_CBC_MD4, L"DES in cipher-block-chaining mode "
                                    L"with a MD4 checksum"                 },
          { KERB_ETYPE_DES_CBC_MD5, L"DES in cipher-block-chaining mode "
                                    L"with a MD5 checksum"                 },
          { KERB_ETYPE_RC4_HMAC_NT, L"RC4 stream cipher with a hash-based "
                                    L"Message Authentication Code"         },
          { KERB_ETYPE_RC4_MD4   ,  L"RC4 stream cipher with the MD4 hash "
                                    L"function"                            },
          { KERB_ETYPE_AES128_CTS_HMAC_SHA1_96,
                                    L"KERB_ETYPE_AES128_CTS_HMAC_SHA1_96"  },
          { KERB_ETYPE_AES256_CTS_HMAC_SHA1_96,
                                    L"KERB_ETYPE_AES256_CTS_HMAC_SHA1_96"  } };

    if ( pEncryptionType == nullptr )
    {
        throw ParameterException( L"pEncryptionType", __FUNCTION__ );
    }
    *pEncryptionType = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Types ); i++ )
    {
        if ( type == Types[ i ].type )
        {
            *pEncryptionType = Types[ i ].pszType;
            break;
        }
    }

    // If got no encryption type will use its numerical value
    if ( pEncryptionType->IsEmpty() )
    {
        *pEncryptionType = Format.Int32( type );
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppWarn2( L"Unrecognised encryption type = %%1 in '%%2'.",
                        *pEncryptionType, Insert2 );
    }
}

//===============================================================================================//
//  Description:
//      Translate a defined Kerberos ticket type
//
//  Parameters:
//      flags        - the encryption type
//      pTicketFlags - string object to receive the translation
//
//  Remarks:
//      See KERB_RETRIEVE_TKT_REQUEST Structure
//
//  Returns:
//      void
//===============================================================================================//
void KerberosTicketInformation::TranslateFlags( ULONG flags, String* pTicketFlags )
{
    size_t    i = 0;
    Formatter Format;

    // Structure to Kerberos ticket flags
    struct _STRUCT_FLAGS
    {
        ULONG   flag;
        LPCWSTR pszDescription;
    } Flags[] =
        { { KERB_TICKET_FLAGS_forwardable  ,  L"Forwardable"    },
          { KERB_TICKET_FLAGS_forwarded    ,  L"Forwarded"      },
          { KERB_TICKET_FLAGS_hw_authent   ,  L"HW authent"     },
          { KERB_TICKET_FLAGS_initial      ,  L"Initial"        },
          { KERB_TICKET_FLAGS_invalid      ,  L"Invalid"        },
          { KERB_TICKET_FLAGS_may_postdate ,  L"May postdate"   },
          { KERB_TICKET_FLAGS_ok_as_delegate, L"OK as delegate" },
          { KERB_TICKET_FLAGS_postdated    ,  L"Postdated"      },
          { KERB_TICKET_FLAGS_pre_authent  ,  L"Pre authent"    },
          { KERB_TICKET_FLAGS_proxiable    ,  L"Proxiable"      },
          { KERB_TICKET_FLAGS_proxy        ,  L"Proxy"          },
          { KERB_TICKET_FLAGS_renewable    ,  L"Renewable"      } };

    if ( pTicketFlags == nullptr )
    {
        throw ParameterException( L"pTicketFlags", __FUNCTION__ );
    }
    *pTicketFlags = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Flags ); i++ )
    {
        if ( flags & Flags[ i ].flag )
        {
            if ( pTicketFlags->GetLength() )
            {
                *pTicketFlags += L", ";
            }
            *pTicketFlags += Flags[ i ].pszDescription;
        }
    }
}
