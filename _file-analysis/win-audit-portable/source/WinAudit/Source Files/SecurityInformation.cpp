///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Security Information Class Implementation
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
#include "WinAudit/Header Files/SecurityInformation.h"

// 2. C System Files
#include <gpmgmt.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoCloseHandle.h"
#include "PxsBase/Header Files/AutoIUnknownRelease.h"
#include "PxsBase/Header Files/AutoNetApiBufferFree.h"
#include "PxsBase/Header Files/AutoVariantClear.h"
#include "PxsBase/Header Files/BStr.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/NullException.h"
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
SecurityInformation::SecurityInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
SecurityInformation::~SecurityInformation()
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
//      Get the privileges of the logged-on user as an array of audit records
//
//  Parameters:
//      pRecords - array to hold the privileges
//
//  Remarks:
//      Returns strings such as SeAuditPrivilege. The list of privileges is
//      dependent on the OS version
//
//  Returns:
//      void
//===============================================================================================//
void SecurityInformation::GetLoggedOnUserPrivilegePrecords( TArray< AuditRecord >* pRecords )
{
    wchar_t szName[ 128 ];         // Enough for any privilege name
    wchar_t szDisplayName[ 512 ];  // Enough for any display name
    BYTE*   pTokenInformation = nullptr;
    DWORD   i = 0, lastError  = 0, TokenInformationLength = 0, languageID = 0;
    DWORD   cchName  = 0, ReturnLength = 0, cchDisplayName = 0;
    HANDLE  ProcessHandle = nullptr, TokenHandle = nullptr;
    String  Insert1;
    AuditRecord       Record;
    AllocateBytes     AllocBytes;
    TOKEN_PRIVILEGES* pTokenPrivileges = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();


    ProcessHandle = GetCurrentProcess();
    if ( OpenProcessToken( ProcessHandle, TOKEN_READ, &TokenHandle ) == 0 )
    {
        throw SystemException( GetLastError(), L"OpenProcessToken", __FUNCTION__ );
    }
    AutoCloseHandle CloseTokenHandle( TokenHandle );

    // First call to GetTokenInformation using using 'TokenPrivileges' to get
    // the required buffer size
    ReturnLength = 0;
    if ( GetTokenInformation( TokenHandle,
                              TokenPrivileges,
                              nullptr,          // i.e no buffer
                              0,                // zero size buffer
                              &ReturnLength ) )
    {
        // Success without a buffer
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo1( L"GetTokenInformation said no data in '%%1'", Insert1 );
        return;
    }

    // ERROR_INSUFFICIENT_BUFFER is the expected error result
    lastError = GetLastError();
    if ( lastError != ERROR_INSUFFICIENT_BUFFER )
    {
        throw SystemException( lastError, L"GetTokenInformation", __FUNCTION__ );
    }

     // Call for the second time with a buffer
    ReturnLength      = PXSMultiplyUInt32( ReturnLength, 2 );
    pTokenInformation = AllocBytes.New( ReturnLength );
    TokenInformationLength = ReturnLength;
    ReturnLength = 0;
    if ( GetTokenInformation( TokenHandle,
                              TokenPrivileges,
                              pTokenInformation, TokenInformationLength, &ReturnLength ) == 0 )
    {
        throw SystemException( GetLastError(),
                               L"GetTokenInformation, second call", __FUNCTION__ );
    }

    // Loop through the locally unique identifiers and get the privilege names
    pTokenPrivileges = reinterpret_cast<TOKEN_PRIVILEGES*>( pTokenInformation );
    for ( i = 0; i < pTokenPrivileges->PrivilegeCount; i++ )
    {
        memset( szName, 0, sizeof ( szName ) );
        cchName = ARRAYSIZE( szName );
        if ( LookupPrivilegeName( nullptr,             // Local system
                                  &pTokenPrivileges->Privileges[ i ].Luid,
                                  szName,
                                  &cchName  ) )  // Characters
        {
            szName[ ARRAYSIZE( szName ) - 1 ] = PXS_CHAR_NULL;
            if ( szName[ 0 ] != PXS_CHAR_NULL )
            {
                memset( szDisplayName, 0, sizeof ( szDisplayName ) );
                cchDisplayName = ARRAYSIZE( szDisplayName );    // In characters
                languageID     = GetSystemDefaultLangID();
                if ( LookupPrivilegeDisplayName( nullptr,
                                                 szName,
                                                 szDisplayName, &cchDisplayName, &languageID ) )
                {
                    Record.Reset( PXS_CATEGORY_USERPRIVS);
                    Record.Add( PXS_USERPRIVS_PRIVILEGE_NAME, szDisplayName );
                    pRecords->Add( Record );
                }
                else
                {
                    // No translation, fall back to symbolic name
                    Insert1 = szName;
                    PXSLogSysInfo1( GetLastError(),
                                    L"LookupPrivilegeDisplayName failed for '%%1'.", Insert1 );
                    Record.Reset( PXS_CATEGORY_USERPRIVS);
                    Record.Add( PXS_USERPRIVS_PRIVILEGE_NAME, szName );
                    pRecords->Add( Record );
                }
            }
        }
        else
        {
            // Did not get a privilege name, these are identified by LUID's
            // which are only valid until next system re-boot
            Insert1.SetAnsi( __FUNCTION__ );
            PXSLogSysError1( GetLastError(), L"LookupPrivilegeName in '%%1'", Insert1 );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_USERPRIVS_PRIVILEGE_NAME );
}

//===============================================================================================//
//  Description:
//      Get the Network Time Protocol on the machine as an array audit records
//
//  Parameters:
//      pRecords - array to receive the data strings
//
//  Remarks:
//      There does not appear to be an interface for this. Although
//      if the system has a registered time provider then could use
//      TimeProvOpen to get its callback functions.
//
//  Returns:
//      void
//===============================================================================================//
void SecurityInformation::GetNetworkTimeProtocolRecords( TArray< AuditRecord >* pRecords )
{
    size_t      i = 0, numElements = 0;
    String      Name, Value;
    Registry    RegObject;
    AuditRecord Record;
    TArray< NameValue > NameValues;
    LPCWSTR NTP_PARAMETERS = L"SYSTEM\\CurrentControlSet\\Services\\W32Time\\Parameters";

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    RegObject.Connect( HKEY_LOCAL_MACHINE );
    RegObject.GetNameValueList( NTP_PARAMETERS, &NameValues );
    numElements = NameValues.GetSize();
    for ( i = 0; i < numElements; i++ )
    {
        Record.Reset( PXS_CATEGORY_NET_TIME_PROTOCOL );
        Name = NameValues.Get( i ).GetName();
        if ( Name.GetLength() )
        {
            Value = NameValues.Get( i ).GetValue();
            Record.Add( PXS_NET_TIME_PROTOCOL_NAME   , Name  );
            Record.Add( PXS_NET_TIME_PROTOCOL_SETTING, Value );
            pRecords->Add( Record );
        }
    }
}

//===============================================================================================//
//  Description:
//      Return miscellaneous security settings in the registry
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Remarks:
//      Comparable to [Registry Values] section of output from secedit
//
//  Returns:
//      void
//===============================================================================================//
void SecurityInformation::GetRegistrySecurityValueRecods( TArray< AuditRecord >* pRecords )
{
    BYTE        binaryData[ 4 ];   // Four bytes for "Driver Signing"
    DWORD       i = 0, j = 0;
    String      Value;
    Registry    RegObject;
    Formatter   Format;
    AuditRecord Record;

    // Structure to hold paths to registry paths and value.
    struct _REG_PATHS
    {
        LPCWSTR pszSubkey;
        LPCWSTR pszValueName;
        LPCWSTR pszDisplay;
    } RegPaths[] = {
    { L"Software\\Policies\\Microsoft\\Windows\\Safer\\CodeIdentifiers\\",
      L"AuthenticodeEnabled"        ,
      L"CodeIdentifiers\\AuthenticodeEnabled"                             },
    { L"Software\\Microsoft\\Driver Signing\\",
      L"Policy"                     ,
      L"Driver Signing\\Policy"                                           },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\Application\\",
      L"RestrictGuestAccess"        ,
      L"Eventlog\\Application\\RestrictGuestAccess"                       },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\Security\\",
      L"RestrictGuestAccess"        ,
      L"Eventlog\\Security\\RestrictGuestAccess"                          },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\System\\",
      L"RestrictGuestAccess"        ,
      L"Eventlog\\System\\RestrictGuestAccess"                            },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\Application\\",
      L"Retention"                  ,
      L"Eventlog\\Application\\Retention"                                 },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\Security\\",
      L"Retention"                  ,
      L"Eventlog\\Security\\Retention"                                    },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\System\\",
      L"Retention"                  ,
      L"Eventlog\\System\\Retention"                                      },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\Application\\",
      L"MaxSize"                    ,
      L"Eventlog\\Application\\MaxSize"                                   },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\Security\\",
      L"MaxSize"                    ,
      L"Eventlog\\Security\\MaxSize"                                      },
    { L"System\\CurrentControlSet\\Services\\Eventlog\\System\\",
      L"MaxSize"                    ,
      L"Eventlog\\System\\MaxSize"                                        },
    { L"System\\CurrentControlSet\\Services\\LanManServer\\Parameters\\",
      L"AutoDisconnect"             ,
      L"LanManServer\\Parameters\\AutoDisconnect"                         },
    { L"System\\CurrentControlSet\\Services\\LanManServer\\Parameters\\",
      L"EnableForcedLogOff"         ,
      L"LanManServer\\Parameters\\EnableForcedLogOff"                     },
    { L"System\\CurrentControlSet\\Services\\LanManServer\\Parameters\\",
      L"EnableSecuritySignature"    ,
      L"LanManServer\\Parameters\\EnableSecuritySignature"                },
    { L"System\\CurrentControlSet\\Services\\LanManServer\\Parameters\\",
      L"NullSessionPipes"           ,
      L"LanManServer\\Parameters\\NullSessionPipes"                       },
    { L"System\\CurrentControlSet\\Services\\LanManServer\\Parameters\\",
      L"RequireSecuritySignature"   ,
      L"LanManServer\\Parameters\\RequireSecuritySignature"               },
    { L"System\\CurrentControlSet\\Services\\LanManServer\\Parameters\\",
      L"RestrictNullSessAccess"     ,
      L"LanManServer\\Parameters\\RestrictNullSessAccess"                 },
    { L"System\\CurrentControlSet\\Services\\LanmanWorkstation\\Parameters\\",
      L"EnablePlainTextPassword"    ,
      L"LanmanWorkstation\\Parameters\\EnablePlainTextPassword"           },
    { L"System\\CurrentControlSet\\Services\\LanmanWorkstation\\Parameters\\",
      L"EnableSecuritySignature"    ,
      L"LanmanWorkstation\\Parameters\\EnableSecuritySignature"           },
    { L"System\\CurrentControlSet\\Services\\LanmanWorkstation\\Parameters\\",
      L"RequireSecuritySignature"   ,
      L"LanmanWorkstation\\Parameters\\RequireSecuritySignature"           },
    { L"System\\CurrentControlSet\\Services\\LDAP\\",
      L"LDAPClientIntegrity"        ,
      L"LDAP\\LDAPClientIntegrity"                                        },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"AuditBaseObjects"           ,
      L"Lsa\\AuditBaseObjects"                                            },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"CrashOnAuditFail"           ,
      L"Lsa\\CrashOnAuditFail"                                            },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"DisableDomainCreds"         ,
      L"Lsa\\DisableDomainCreds"                                          },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"EveryoneIncludesAnonymous"  ,
      L"Lsa\\EveryoneIncludesAnonymous"                                   },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"FIPSAlgorithmPolicy\\Enabled",
      L"Lsa\\FIPSAlgorithmPolicy\\Enabled"                                },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"ForceGuest"                 ,
      L"Lsa\\ForceGuest"                                                  },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"FullPrivilegeAuditing"      ,
      L"Lsa\\FullPrivilegeAuditing"                                       },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"LimitBlankPasswordUse"      ,
      L"Lsa\\LimitBlankPasswordUse"                                       },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"LmCompatibilityLevel"       ,
      L"Lsa\\LmCompatibilityLevel"                                        },
    { L"System\\CurrentControlSet\\Control\\Lsa\\MSV1_0\\",
      L"NTLMMinClientSec"           ,
      L"Lsa\\MSV1_0\\NTLMMinClientSec"                                    },
    { L"System\\CurrentControlSet\\Control\\Lsa\\MSV1_0\\",
      L"NTLMMinServerSec"           ,
      L"Lsa\\MSV1_0\\NTLMMinServerSec"                                    },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"NoLMHash"                   ,
      L"Lsa\\NoLMHash"                                                    },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"RestrictAnonymous"          ,
      L"Lsa\\RestrictAnonymous"                                           },
    { L"System\\CurrentControlSet\\Control\\Lsa\\",
      L"RestrictAnonymousSAM"       ,
      L"Lsa\\RestrictAnonymousSAM"                                        },
    { L"System\\CurrentControlSet\\Services\\Netlogon\\Parameters\\",
      L"DisablePasswordChange"      ,
      L"Netlogon\\Parameters\\DisablePasswordChange"                       },
    { L"System\\CurrentControlSet\\Services\\Netlogon\\Parameters\\",
      L"MaximumPasswordAge"         ,
      L"Netlogon\\Parameters\\MaximumPasswordAge"                         },
    { L"System\\CurrentControlSet\\Services\\Netlogon\\Parameters\\",
      L"RequireSignOrSeal"          ,
      L"Netlogon\\Parameters\\RequireSignOrSeal"                          },
    { L"System\\CurrentControlSet\\Services\\Netlogon\\Parameters\\",
      L"RequireStrongKey"           ,
      L"Netlogon\\Parameters\\RequireStrongKey"                           },
    { L"System\\CurrentControlSet\\Services\\Netlogon\\Parameters\\",
      L"SealSecureChannel"          ,
      L"Netlogon\\Parameters\\SealSecureChannel"                          },
    { L"System\\CurrentControlSet\\Services\\Netlogon\\Parameters\\",
      L"SignSecureChannel"          ,
      L"Netlogon\\Parameters\\SignSecureChannel"                          },
    { L"System\\CurrentControlSet\\Services\\NTDS\\Parameters\\",
      L"LDAPServerIntegrity"        ,
      L"NTDS\\Parameters\\LDAPServerIntegrity"                            },
    { L"System\\CurrentControlSet\\Control\\Session Manager\\Kernel\\",
      L"ObCaseInsensitive"          ,
      L"Session Manager\\Kernel\\ObCaseInsensitive"                       },
    { L"System\\CurrentControlSet\\Control\\Session Manager\\",
      L"ProtectionMode"             ,
      L"Session Manager\\ProtectionMode"                                  },
    { L"System\\CurrentControlSet\\Control\\Session Manager\\",
      L"SubSystems\\optional"       ,
      L"Session Manager\\SubSystems\\optional"                            },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"ConsentPromptBehaviorAdmin",
      L"System\\ConsentPromptBehaviorAdmin"                               },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"ConsentPromptBehaviorUser"  ,
      L"System\\ConsentPromptBehaviorUser"                                },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"DisableCAD"                 ,
      L"System\\DisableCAD"                                               },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"DontDisplayLastUserName"    ,
      L"System\\DontDisplayLastUserName"                                  },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"EnableInstallerDetection"   ,
      L"System\\EnableInstallerDetection"                                 },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"EnableLUA"                  ,
      L"System\\EnableLUA"                                                },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"EnableSecureUIAPaths"       ,
      L"System\\EnableSecureUIAPaths"                                     },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"EnableUIADesktopToggle"     ,
      L"System\\EnableUIADesktopToggle"                                   },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"EnableVirtualization"       ,
      L"System\\EnableVirtualization"                                     },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"FilterAdministratorToken"   ,
      L"System\\FilterAdministratorToken"                                 },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"LegalNoticeCaption"         ,
      L"System\\LegalNoticeCaption"                                       },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"LegalNoticeText"            ,
      L"System\\LegalNoticeText"                                          },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"PromptOnSecureDesktop"      ,
      L"System\\PromptOnSecureDesktop"                                    },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"ScForceOption"             ,
      L"System\\ScForceOption"                                            },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"ShutdownWithoutLogon"      ,
      L"System\\ShutdownWithoutLogon"                                     },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"UndockWithoutLogon"        ,
      L"System\\UndockWithoutLogon"                                       },
    { L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\",
      L"ValidateAdminCodeSignatures",
      L"System\\ValidateAdminCodeSignatures"                              },
    { L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon\\",
      L"CachedLogonsCount"         ,
      L"Winlogon\\CachedLogonsCount"                                      },
    { L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon\\",
      L"ForceUnlockLogon"          ,
      L"Winlogon\\ForceUnlockLogon"                                       },
    { L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon\\",
      L"PasswordExpiryWarning"     ,
      L"Winlogon\\PasswordExpiryWarning"                                  },
    { L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon\\",
      L"ScRemoveOption"            ,
      L"Winlogon\\ScRemoveOption"                                        },
    };

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    RegObject.Connect( HKEY_LOCAL_MACHINE );
    for ( i = 0;  i < ARRAYSIZE( RegPaths ); i++ )
    {
        // In the event of an error getting a value, continue to the next one
        try
        {
            // Test for "Driver Signing" as this is binary data
            Value = PXS_STRING_EMPTY;
            if ( CSTR_EQUAL == CompareString(
                                    LOCALE_INVARIANT,
                                    NORM_IGNORECASE,
                                    L"Software\\Microsoft\\Driver Signing\\",
                                    -1, RegPaths[ i ].pszSubkey, - 1 ) )
            {
                // This data seems to be 4 bytes long
                memset( binaryData, 0, sizeof ( binaryData ) );
                RegObject.GetBinaryData( RegPaths[ i ].pszSubkey,
                                         RegPaths[ i ].pszValueName,
                                         binaryData, sizeof ( binaryData ) );
                for ( j = 0; j < sizeof ( binaryData ); j++ )
                {
                    Value += Format.UInt8Hex( binaryData[ j ], false );
                }
            }
            else
            {
                RegObject.GetValueAsString( RegPaths[ i ].pszSubkey,
                                            RegPaths[ i ].pszValueName, &Value );
            }

            // Make the audit record
            Record.Reset( PXS_CATEGORY_REG_SEC_VALUES );
            Record.Add( PXS_REG_SEC_VALS_SUBKEY_SUBKEY, RegPaths[i].pszDisplay);
            Record.Add( PXS_REG_SEC_VALS_SUBKEY_SETTING, Value );
            pRecords->Add( Record );
        }
        catch ( const Exception& eSetting )
        {
            // Log and continue
            PXSLogException( eSetting, __FUNCTION__ );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the resultant set of policy result in HTML format
//
//  Parameters:
//      pRsopHtmlResult - receives the HTML data
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::GetRsopHtmlResult( String* pRsopHtmlResult )
{
    BStr    BStrFilePath, BStrNamespace, BStrLoggingComputer;
    BStr    BStrLoggingUser;
    IGPM*   pIGPM = nullptr;
    wchar_t wzGuid[ 48 ] = { 0 };  // Enough for a GUID
    CLSID   iidIGPM, clsidGPM;
    String  UserName, ComputerName, VarType, EmptyString;
    VARIANT variant;
    HRESULT hResult;
    Formatter   Format;
    IGPMRSOP*   pIGPMRSOP   = nullptr;
    IGPMResult* pIGPMResult = nullptr;
    SystemInformation   SystemInfo;
    AutoIUnknownRelease AutoReleaseIGPM;
    AutoIUnknownRelease AutoReleaseIGPMRSOP;
    AutoIUnknownRelease AutoReleaseIGPMResult;

    if ( pRsopHtmlResult == nullptr )
    {
        throw ParameterException( L"pRsopHtmlResult", __FUNCTION__ );
    }
    *pRsopHtmlResult = PXS_STRING_EMPTY;

    // CLSID_GPM
    StringCchCopy( wzGuid, ARRAYSIZE( wzGuid ), L"{F5694708-88FE-4B35-BABF-E56162D5FBC8}" );
    memset( &clsidGPM , 0 , sizeof ( clsidGPM ) );
    hResult = CLSIDFromString( wzGuid , &clsidGPM );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CLSIDFromString, CLSID_GPM", __FUNCTION__ );
    }

    // IID_IGPM
    StringCchCopy( wzGuid, ARRAYSIZE( wzGuid ), L"{F5FAE809-3BD6-4DA9-A65E-17665B41D763}" );
    memset( &iidIGPM  , 0 , sizeof ( iidIGPM ) );
    hResult = CLSIDFromString( wzGuid , &iidIGPM );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CLSIDFromString, IID_IGPM", __FUNCTION__ );
    }

    hResult = CoCreateInstance( clsidGPM ,
                                nullptr,
                                CLSCTX_INPROC_SERVER, iidIGPM, reinterpret_cast<void**>(&pIGPM) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoCreateInstance", __FUNCTION__ );
    }

    if ( pIGPM == nullptr )
    {
        throw NullException( L"pIGPM", __FUNCTION__ );
    }
    AutoReleaseIGPM.Set( pIGPM );

    // Set to logging mode. The namespace should be "". Avoid using
    // "root\rsop\computer" as IGPMRSOP::GenerateReport may fail
    EmptyString = PXS_STRING_EMPTY;
    BStrNamespace.Allocate( EmptyString );
    hResult = pIGPM->GetRSOP( rsopLogging, BStrNamespace.b_str(), 0, &pIGPMRSOP );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IGPM::GetRSOP", __FUNCTION__ );
    }

    if ( pIGPMRSOP == nullptr )
    {
        throw NullException( L"pIGPMRSOP", __FUNCTION__ );
    }
    AutoReleaseIGPMRSOP.Set( pIGPMRSOP );

    // Set the logging computer
    SystemInfo.GetComputerNetBiosName( &ComputerName );
    BStrLoggingComputer.Allocate( ComputerName );
    hResult = pIGPMRSOP->put_LoggingComputer( BStrLoggingComputer.b_str() );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IGPMRSOP::put_LoggingComputer", __FUNCTION__ );
    }

    // Set the logging username
    SystemInfo.GetCurrentUserName( &UserName );
    BStrLoggingUser.Allocate( UserName );
    hResult = pIGPMRSOP->put_LoggingUser( BStrLoggingUser.b_str() );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IGPMRSOP::put_LoggingUser", __FUNCTION__ );
    }

    // Get the result in XML format
    hResult = pIGPMRSOP->CreateQueryResults();
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IGPMRSOP::CreateQueryResults", __FUNCTION__ );
    }

    // C6309 and C6387 because pvarGPMProgress and pvarGPMCancel are NULL
    // but these are optional input arguments. See gpmgmt.h
    hResult = pIGPMRSOP->GenerateReport( repHTML, nullptr, nullptr, &pIGPMResult );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IGPMRSOP::GenerateReportToFile", __FUNCTION__ );
    }

    if ( pIGPMResult == nullptr )
    {
        throw NullException( L"pIGPMResult", __FUNCTION__ );
    }
    AutoReleaseIGPMResult.Set( pIGPMResult);

    VariantInit( &variant );
    hResult = pIGPMResult->get_Result( &variant );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IGPMResult::get_Result", __FUNCTION__ );
    }
    AutoVariantClear VariantClear( &variant );

    // Test for NULL/EMPTY
    if ( ( variant.vt == VT_NULL ) || ( variant.vt == VT_EMPTY )  )
    {
        *pRsopHtmlResult = L"<html><body>No Information</body></html>";
        return;
    }

    // Expected data type is a binary string
    if ( variant.vt != VT_BSTR )
    {
        VarType = Format.StringInt32( L"VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, VarType.c_str(), __FUNCTION__ );
    }
    *pRsopHtmlResult = variant.bstrVal;
}

//===============================================================================================//
//  Description:
//      Get the specified RSOP security setting
//
//  Parameters:
//      pszKeyName - the policy's key name
//      varType    - the variant data type returned by the WQL query
//      pSetting   - receives setting
//
//  Remarks:
//      The "Setting" property of RSOP_SecuritySettingNumeric is a uint32
//      (CIM_UIN32) but the variant data type returned by a WQL query is VT_I4
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::GetRSOPSecuritySetting( LPCWSTR pszKeyName,
                                                  VARTYPE varType, String* pSetting )
{
    Wmi    WMI;
    String Query, KeyName;
    WindowsInformation WindowsInfo;
    const wchar_t WMI_RSOP_NAMESPACE[] = L"root\\rsop\\computer";

    if ( pSetting == nullptr )
    {
        throw ParameterException( L"pSetting", __FUNCTION__ );
    }
    *pSetting = PXS_STRING_EMPTY;

    // Need a key name to build the WQL query
    KeyName = pszKeyName;
    KeyName.Trim();
    if ( KeyName.IsEmpty() )
    {
       throw ParameterException( L"pszKeyName", __FUNCTION__ );
    }

    // Build the query string, it is in the form: SELECT Setting FROM
    // RSOP_SecuritySettingBoolean WHERE KeyName="PasswordComplexity"
    // RSOP only uses VT_BOOL, VT_I4 and VT_BSTR data types.
    Query = L"SELECT Setting FROM ";
    if ( varType == VT_BOOL )
    {
        Query += L" RSOP_SecuritySettingBoolean ";
    }
    else if ( varType == VT_I4 )
    {
        Query += L" RSOP_SecuritySettingNumeric ";
    }
    else if ( varType == VT_BSTR )
    {
        Query += L" RSOP_SecuritySettingString ";
    }
    else
    {
        throw ParameterException( L"varType", __FUNCTION__ );
    }
    Query += L" WHERE KeyName=\"";
    Query += KeyName;
    Query += L"\"";

    // Query the local machine only
    WMI.Connect( WMI_RSOP_NAMESPACE );
    WMI.ExecQuery( Query.c_str() );
    if ( WMI.Next() )       // Only expecting 1 result
    {
        WMI.Get( L"Setting", pSetting );
    }
}

//===============================================================================================//
//  Description:
//      Get miscellaneous security settings as as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the settings
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::GetSecuritySettingsRecords(
                                              TArray< AuditRecord >* pRecords )
{
    AuditRecord Record;
    TArray< AuditRecord > AuditRecords;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Accounts Settings
    try
    {
        AuditRecords.RemoveAll();
        MakeAccountSettingsRecords( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting accounts information", e, __FUNCTION__ );
    }

    // Account Lockout Policy
    try
    {
        AuditRecords.RemoveAll();
        MakeAccountLockingRecords( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting lockout policy information.", e, __FUNCTION__ );
    }

    // Autologon
    try
    {
        Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
        MakeAutoLogonRecord( &Record );
        pRecords->Add( Record );
     }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting AutoLogon information.", e, __FUNCTION__ );
    }

    // Automatic Updates
    try
    {
        AuditRecords.RemoveAll();
        MakeAutomaticUpdateRecords( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting automatic update data.", e, __FUNCTION__ );
    }

    // Internet Explorer
    try
    {
        AuditRecords.RemoveAll();
        MakeInternetExplorerRecords( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting Internet Explorer data.", e, __FUNCTION__ );
    }

    // Audit Policy
    try
    {
        AuditRecords.RemoveAll();
        MakeAuditPolicyRecords( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting audit policy data.", e, __FUNCTION__ );
    }

    // Get programmes allowed to execute data
    try
    {
        AuditRecords.RemoveAll();
        MakeExecuteDataRecord( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting execute data information.", e, __FUNCTION__ );
    }

    // Network access: Allow anonymous SID/Name translation
    try
    {
        Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
        MakeNetworkAccessRecord( &Record );
        pRecords->Add( Record );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting network access information.", e, __FUNCTION__ );
    }

    // Network Security
    try
    {
        AuditRecords.RemoveAll();
        MakeNetworkSecurityRecords( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting network security information", e, __FUNCTION__ );
    }

    // Password Policy
    try
    {
        AuditRecords.RemoveAll();
        MakePasswordPolicyRecords( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting network security information", e, __FUNCTION__ );
    }

    // Screen Saver Settings for logged on user
    try
    {
        AuditRecords.RemoveAll();
        MakeScreenSaverRecords( &AuditRecords );
        pRecords->Append( AuditRecords );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting network security information", e, __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get system restore points as an array of audit records
//
//  Parameters:
//      pRecords - array to receive formatted data strings
//
//  Remarks:
//      Available on XP with WMI
//      Data seems to be returned in reverse SequenceNumber so do two
//      passes, first to get the sequence numbers than to get the data.
//
//  Returns:
//      void
//===============================================================================================//
void SecurityInformation::GetSystemRestorePointRecords( TArray< AuditRecord >* pRecords )
{
    Wmi         WMI;
    size_t      i = 0, numRestorePoints = 0;
    String      CimDate, CreateTime, Description, SequenceNumber;
    String      Query;
    Formatter   Format;
    StringArray Sequences;
    AuditRecord Record;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Pass 1 - Restore points are in the default name space. WQL does not
    // support the ORDER BY clause
    WMI.Connect( L"root\\default" );
    WMI.ExecQuery( L"Select * from SystemRestore" );
    while ( WMI.Next() )
    {
        SequenceNumber = PXS_STRING_EMPTY;
        WMI.Get( L"SequenceNumber", &SequenceNumber );
        if ( SequenceNumber.GetLength() )
        {
            Sequences.AddUniqueI( SequenceNumber.c_str() );
        }
    }
    WMI.Disconnect();       // Must close because not sure got to the end of the
                            // enumeration, so the record set may still be open

    // Pass 2 - get the restore points by sequence number
    numRestorePoints = Sequences.GetSize();
    WMI.Connect( L"root\\default" );
    for ( i = 0; i < numRestorePoints; i++ )
    {
        WMI.CloseQuery();

        // New query
        Query  = L"Select * from SystemRestore Where SequenceNumber=";
        Query += Sequences.Get( numRestorePoints - i - 1 );
        WMI.ExecQuery( Query.c_str() );
        if ( WMI.Next() )         // There should only be one record
        {
            CimDate     = PXS_STRING_EMPTY;
            Description = PXS_STRING_EMPTY;
            WMI.Get( L"CreationTime", &CreateTime  );
            WMI.Get( L"Description",  &Description );

            // Will truncate the for yyyymmddhhmmss.fraction at the dot for
            // readability
            CreateTime.ReplaceChar( '.', PXS_CHAR_NULL );

            // Make the audit record
            Record.Reset( PXS_CATEGORY_SYSRESTORE );
            Record.Add( PXS_SYSRESTORE_SEQUENCE, Sequences.Get( i ) );
            Record.Add( PXS_SYSRESTORE_CREATE_TIME, CreateTime );
            Record.Add( PXS_SYSRESTORE_DESCRIPTION, Description );
            pRecords->Add( Record );
        }
    }
}

//===============================================================================================//
//  Description:
//        Get the User Rights Assignment policy on the machine
//
//  Parameters:
//        pRecords - array to receive the audit records
//
//  Remarks:
//      The privileges available depend on the OS so LookupPrivilegeDisplayName
//      may not provide a translation. For display purposes want to avoid a mix
//      of  "Se.." constants and descriptions so will use the English strings
//      as documented by MSDN.
//
//  Returns:
//      void
//===============================================================================================//
void SecurityInformation::GetUserRightsAssignmentRecords( TArray< AuditRecord >* pRecords )
{
    ULONG    CountReturned = 0;
    wchar_t  wzPrivilege[ 64 ]  = { 0 };           // Enough for any "Se..."
    size_t   i = 0, numBytes = 0;
    String   Setting, Insert1;
    NTSTATUS status = 0;
    AuditRecord Record;
    LSA_HANDLE  PolicyHandle = nullptr;
    LSA_UNICODE_STRING    UserRights;
    LSA_OBJECT_ATTRIBUTES ObjectAttributes;
    LSA_ENUMERATION_INFORMATION* pEnumerationBuffer = nullptr;

    // Constants are take from MSDN "Account Rights Constants" and "Privilege
    // Constants" ordered to match the output of MMC's User Rights Assignment
    struct _PIVILEGES
    {
        LPCWSTR pwzPrivilege;
        LPCWSTR pszPolicy;
    } Privileges[] = {
        { L"SeTrustedCredManAccessPrivilege",
          L"Access Credential Manager as a trusted caller" },
        { L"SeNetworkLogonRight",
          L"Access this computer from the network"         },
        { L"SeTcbPrivilege",
          L"Act as part of the operating system"           },
        { L"SeMachineAccountPrivilege",
          L"Add workstations to domain"                    },
        { L"SeIncreaseQuotaPrivilege",
          L"Adjust memory quotas for a process"            },
        { L"SeRemoteInteractiveLogonRight",
          L"Allow log on through Terminal Services"        },
        { L"SeBackupPrivilege",
          L"Back up files and directories"                 },
        { L"SeChangeNotifyPrivilege",
          L"Bypass traverse checking"                      },
        { L"SeSystemtimePrivilege",
          L"Change the system time"                        },
        { L"SeTimeZonePrivilege",
          L"Change the time zone"                          },
        { L"SeCreatePagefilePrivilege",
          L"Create a pagefile"                             },
        { L"SeCreateTokenPrivilege",
          L"Create a token object"                         },
        { L"SeCreateGlobalPrivilege",
          L"Create global objects"                         },
        { L"SeCreatePermanentPrivilege",
          L"Create permanent shared objects"               },
        { L"SeCreateSymbolicLinkPrivilege",
          L"Create symbolic links"                         },
        { L"SeDebugPrivilege",
          L"Debug programs"                                },
        { L"SeDenyNetworkLogonRight",
          L"Deny access to this computer from the network" },
        { L"SeDenyBatchLogonRight",
          L"Deny log on as a batch job"                    },
        { L"SeDenyServiceLogonRight",
          L"Deny log on as a service"                      },
        { L"SeDenyServiceLogonRight",
          L"Deny log on locally"                           },
        { L"SeDenyRemoteInteractiveLogonRight",
          L"Deny log on through Terminal Services"         },
        { L"SeEnableDelegationPrivilege",
          L"Enable delegation"                             },
        { L"SeRemoteShutdownPrivilege",
          L"Force shutdown from a remote system"           },
        { L"SeAuditPrivilege",
          L"Generate security audits"                      },
        { L"SeImpersonatePrivilege",
          L"Impersonate a client after authentication"     },
        { L"SeIncreaseWorkingSetPrivilege",
          L"Increase a process working set"                },
        { L"SeIncreaseBasePriorityPrivilege",
          L"Increase scheduling priority"                  },
        { L"SeLoadDriverPrivilege"            ,
          L"Load and unload device drivers"                },
        { L"SeLockMemoryPrivilege",
          L"Lock pages in memory"                          },
        { L"SeBatchLogonRight",
          L"Log on as a batch job"                         },
        { L"SeDenyServiceLogonRight",
          L"Log on as a service"                           },
        { L"SeInteractiveLogonRight",
          L"Log on locally"                                },
        { L"SeSecurityPrivilege",
          L"Manage auditing and security log"              },
        { L"SeManageVolumePrivilege"          ,
          L"Manage the files on a volume"                  },
        { L"SeRelabelPrivilege",
          L"Modify an object label"                        },
        { L"SeSystemEnvironmentPrivilege",
          L"Modify firmware environment values"            },
        { L"SeProfileSingleProcessPrivilege",
          L"Profile single process"                        },
        { L"SeSystemProfilePrivilege",
          L"Profile system performance"                    },
        { L"SeUndockPrivilege",
          L"Remove computer from docking station"          },
        { L"SeAssignPrimaryTokenPrivilege",
          L"Replace a process-level token"                 },
        { L"SeRestorePrivilege",
          L"Restore files and directories"                 },
        { L"SeShutdownPrivilege",
          L"Shut down the system"                          },
        { L"SeSyncAgentPrivilege",
          L"Synchronize directory service data"            },
        { L"SeTakeOwnershipPrivilege",
          L"Take ownership of files or other objects"      } };

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Open the policy, only need view access
    memset( &ObjectAttributes, 0, sizeof ( ObjectAttributes ) );
    status = LsaOpenPolicy( nullptr,       // Local machine
                            &ObjectAttributes,
                            POLICY_LOOKUP_NAMES | POLICY_VIEW_LOCAL_INFORMATION, &PolicyHandle );
    if ( status != ERROR_SUCCESS )
    {
        throw SystemException( LsaNtStatusToWinError( status ), L"LsaOpenPolicy", __FUNCTION__ );
    }

    if ( PolicyHandle == nullptr )
    {
        throw NullException( L"PolicyHandle", __FUNCTION__ );
    }

    // Catch all exceptions to close the policy handle
    try
    {
        Setting.Allocate( 4096 );
        for ( i = 0; i < ARRAYSIZE( Privileges ); i++ )
        {
            Setting = PXS_STRING_EMPTY;
            memset( wzPrivilege, 0, sizeof ( wzPrivilege ) );
            StringCchCopy( wzPrivilege, ARRAYSIZE( wzPrivilege ), Privileges[ i ].pwzPrivilege );
            numBytes = PXSMultiplySizeT( wcslen( wzPrivilege ), sizeof ( wchar_t ) );
            UserRights.Length        = PXSCastSizeTToUInt16( numBytes );
            UserRights.MaximumLength = sizeof ( wzPrivilege );
            UserRights.Buffer        = wzPrivilege;

            // Get the SIDs with this privilege/right. Will ignore the expected
            // errors of STATUS_NO_SUCH_PRIVILEGE which means the privilege does
            // not exist for the operating system and STATUS_NO_MORE_ENTRIES
            // means at the end of the enumeration
            pEnumerationBuffer = nullptr;
            CountReturned = 0;
            status = LsaEnumerateAccountsWithUserRight(
                                                   PolicyHandle,
                                                   &UserRights,
                                                   reinterpret_cast<void**>( &pEnumerationBuffer ),
                                                   &CountReturned );
            if ( status == ERROR_SUCCESS )
            {
                EnumerationInformationToAccounts( pEnumerationBuffer, CountReturned, &Setting );
                LsaFreeMemory( pEnumerationBuffer );
                pEnumerationBuffer = nullptr;            // Reset for next pass
            }
            else
            {
                // Log as a warning as not all privileges exist on all systems
                Insert1 = Privileges[ i ].pwzPrivilege;
                PXSLogNtStatusWarn1( status,
                                     L"LsaEnumerateAccountsWithUserRight failed for '%%1'",
                                     Insert1 );
            }

            // Will add a record even if no one has the privilege/right to
            // mimic the result as show in User Rights Assignment of the
            // MMC's Group Policy
            Record.Reset( PXS_CATEGORY_USER_RIGHTS );
            Record.Add( PXS_USER_RIGHTS_POLICY, Privileges[ i ].pszPolicy );
            Record.Add( PXS_USER_RIGHTS_SETTING, Setting );
            pRecords->Add( Record );
        }
    }
    catch ( const Exception& )
    {
        if ( pEnumerationBuffer )
        {
            LsaFreeMemory( pEnumerationBuffer );
        }
        LsaClose( PolicyHandle );
        throw;
    }
    LsaClose( PolicyHandle );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Lookup the SIDs in the LSA_ENUMERATION_INFORMATION to make a comma
//      separated list of account names
//
//  Parameters:
//      pEnumeration  - pointer to the buffer
//      CountReturned - number of SIDs in the buffer
//      pAccounts     - receives the account names
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::EnumerationInformationToAccounts(
                                                   const LSA_ENUMERATION_INFORMATION* pEnumeration,
                                                   ULONG count, String* pAccounts )
{
    ULONG   i = 0;
    DWORD   cchDomainName = 0, cchName = 0;
    wchar_t szName[ UNLEN + 1 ]   = { 0 };         // Enough for a account name
    wchar_t szDomainName[ MAX_PATH + 1 ] = { 0 };  // Enough for a domain name
    Formatter    Format;
    SID_NAME_USE eUse;

    if ( pAccounts == nullptr )
    {
        throw ParameterException( L"pAccounts", __FUNCTION__ );
    }
    *pAccounts = PXS_STRING_EMPTY;

    if ( pEnumeration == nullptr )
    {
        return;
    }

    for ( i = 0; i < count; i++ )
    {
        if ( pEnumeration[ i ].Sid )
        {
            if ( pAccounts->GetLength() )
            {
                *pAccounts += L", ";
            }
            memset( szName, 0, sizeof ( szName ) );
            memset( szDomainName,  0, sizeof ( szDomainName  ) );
            cchName = ARRAYSIZE( szName );
            cchDomainName  = ARRAYSIZE( szDomainName  );
            if ( LookupAccountSid( nullptr,  // Local computer only
                                   pEnumeration[ i ].Sid,
                                   szName,
                                   &cchName,
                                   szDomainName, &cchDomainName, &eUse ) )
            {
                szName[ ARRAYSIZE( szName ) - 1 ] = PXS_CHAR_NULL;
                *pAccounts += szName;
            }
            else
            {
                // Failed, format as a string
                *pAccounts += Format.SidToString( pEnumeration[ i ].Sid );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Make the audit data records for account lockout settings
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeAccountLockingRecords( TArray< AuditRecord >* pRecords )
{
    LPCWSTR     STR_ACCOUNT_LOCKOUT_POLICY = L"Account Lockout Policy";
    String      StringData, Insert1;
    Formatter   Format;
    AuditRecord Record;
    NET_API_STATUS      status = 0;
    USER_MODALS_INFO_3* pUMI3  = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get information at level 3
    status = NetUserModalsGet( nullptr, 3, reinterpret_cast<BYTE**>( &pUMI3 ) );
    if ( status != NERR_Success )
    {
        throw Exception( PXS_ERROR_TYPE_NETWORK, status, L"NetUserModalsGet", __FUNCTION__ );
    }

    // Always check for NULL buffer
    if ( pUMI3 == nullptr )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo1( L"NetUserModalsGet returned no data in '%%1'.", Insert1 );
        return;
    }
    AutoNetApiBufferFree AutoFreeUMI3( pUMI3 );

    // Account lockout duration
    StringData = L"Not Applicable";
    if ( pUMI3->usrmod3_lockout_threshold > 0 )
    {
        StringData = Format.UInt32( ( pUMI3->usrmod3_lockout_duration / 60 ) );
        if ( ( pUMI3->usrmod3_lockout_duration / 60 ) == 1 )
        {
            StringData += L" Minute";
        }
        else
        {
            StringData += L" Minutes";
        }
    }
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_ACCOUNT_LOCKOUT_POLICY );
    Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Account lockout duration" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
    pRecords->Add( Record );

    // Account lockout threshold
    StringData  = Format.UInt32( ( pUMI3->usrmod3_lockout_threshold ) );
    if ( pUMI3->usrmod3_lockout_threshold == 1 )
    {
        StringData += L" Attempt";
    }
    else
    {
        StringData += L" Attempts";
    }
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_ACCOUNT_LOCKOUT_POLICY );
    Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Account lockout threshold");
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
    pRecords->Add( Record );

    // Reset account lockout counter after
    StringData = L"Not Applicable";
    if ( pUMI3->usrmod3_lockout_threshold > 0 )
    {
        StringData  = Format.UInt32( pUMI3->usrmod3_lockout_observation_window / 60 );
        if ( ( pUMI3->usrmod3_lockout_observation_window / 60 ) == 1 )
        {
            StringData += L" Minute";
        }
        else
        {
            StringData += L" Minutes";
        }
    }
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_ACCOUNT_LOCKOUT_POLICY );
    Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Reset account lockout counter after" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
    pRecords->Add( Record );
}

//===============================================================================================//
//  Description:
//      Make the audit data records for account settings
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeAccountSettingsRecords( TArray< AuditRecord >* pRecords)
{
    String      StringData, AdminName, GuestName;
    AuditRecord Record;
    SystemInformation SystemInfo;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get account names
    SystemInfo.GetBuiltInAccountNameFromRID(DOMAIN_USER_RID_ADMIN, &AdminName);
    SystemInfo.GetBuiltInAccountNameFromRID(DOMAIN_USER_RID_GUEST, &GuestName);

    // Accounts: Administrator account status (secedit's EnableAdminAccount)
    StringData = L"Enabled";
    if ( SystemInfo.IsAccountEnabled( AdminName ) == false )
    {
        StringData = L"Disabled";
    }
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, L"Accounts" );
    Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Administrator account status");
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
    pRecords->Add( Record );

    // Accounts: Guest account status (secedit's EnableGuestAccount)
    StringData = L"Enabled";
    if ( SystemInfo.IsAccountEnabled( GuestName ) == false )
    {
        StringData = L"Disabled";
    }
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, L"Accounts" );
    Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Guest account status" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
    pRecords->Add( Record );

    // Accounts: Rename administrator account (secedit's NewAdministratorName)
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, L"Accounts" );
    Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Rename administrator account");
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, AdminName );
    pRecords->Add( Record );

    // Accounts: Rename guest account (secedit's NewGuestName)
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, L"Accounts" );
    Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Rename guest account" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, GuestName );
    pRecords->Add( Record );
}

//===============================================================================================//
//  Description:
//      Make the formatted audit records for the audit policy settings
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeAuditPolicyRecords( TArray<AuditRecord>* pRecords)
{
    ULONG       i = 0;
    String      Type, Setting;
    NTSTATUS    status = 0;
    LSA_HANDLE  PolicyHandle = nullptr;
    AuditRecord Record;
    LSA_OBJECT_ATTRIBUTES     ObjectAttributes;
    POLICY_AUDIT_EVENT_TYPE   type;
    POLICY_AUDIT_EVENTS_INFO* pAuditEventsInfo = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Open the policy with view access
    memset( &ObjectAttributes, 0, sizeof ( ObjectAttributes ) );
    status = LsaOpenPolicy( nullptr,                       // Local computer
                            &ObjectAttributes, POLICY_VIEW_AUDIT_INFORMATION, &PolicyHandle );
     if ( status != ERROR_SUCCESS )
     {
         throw SystemException( LsaNtStatusToWinError( status ), L"LsaOpenPolicy", __FUNCTION__ );
     }

    // Get the system's auditing rules
    status = LsaQueryInformationPolicy( PolicyHandle,
                                        PolicyAuditEventsInformation,
                                        reinterpret_cast<void**>( &pAuditEventsInfo ) );
    if ( status != ERROR_SUCCESS )
    {
        LsaClose( PolicyHandle );
        throw SystemException( LsaNtStatusToWinError( status ),
                               L"LsaQueryInformationPolicy", __FUNCTION__ );
    }

    // Always check for NULL pointer
    if ( pAuditEventsInfo == nullptr )
    {
        LsaClose( PolicyHandle );
        throw NullException( L"pAuditEventsInfo", __FUNCTION__ );
    }

    // Catch exceptions so can clean up
    try
    {
        for ( i = 0; i < pAuditEventsInfo->MaximumAuditEventCount; i++ )
        {
            // The index of each array element corresponds to an audit event
            // type value in the POLICY_AUDIT_EVENT_TYPE enumeration type.
            type = static_cast<POLICY_AUDIT_EVENT_TYPE>( i );
            Type = PXS_STRING_EMPTY;
            TranslatePolicyAuditEventType( type, &Type );

            // POLICY_AUDIT_EVENT_NONE seems not to be used so if neither
            // "Success" nor "Failure" are set will use "No auditing"
            Setting = PXS_STRING_EMPTY;
            if ( POLICY_AUDIT_EVENT_SUCCESS & pAuditEventsInfo->EventAuditingOptions[i] )
            {
                Setting = L"Success";
            }

            // Test for failure
            if ( POLICY_AUDIT_EVENT_FAILURE & pAuditEventsInfo->EventAuditingOptions[i] )
            {
                if ( Setting.GetLength() )
                {
                    Setting += L", ";
                }
                Setting += L"Failure";
            }

            // Test for No auditing
            if ( Setting.IsEmpty() )
            {
                Setting = L"No auditing";
            }

            // Make the record starting with the table name
            Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
            Record.Add( PXS_SECURITY_SETTINGS_ITEM, L"Audit Policy" );
            Record.Add( PXS_SECURITY_SETTINGS_NAME, Type );
            Record.Add( PXS_SECURITY_SETTINGS_SETTING, Setting );
            pRecords->Add( Record );
        }
    }
    catch ( const Exception& )
    {
        LsaFreeMemory( pAuditEventsInfo );
        LsaClose( PolicyHandle );
        throw;
    }
    LsaFreeMemory( pAuditEventsInfo );
    LsaClose( PolicyHandle );
}

//===============================================================================================//
//  Description:
//      Make an audit data record for the auto logon setting
//
//  Parameters:
//      pRecord - receives the audit record
//
//  Remarks:
//      See MSDN: "How to turn on automatic logon in Windows XP"
//      Article ID 315231
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeAutoLogonRecord( AuditRecord* pRecord )
{
    String   AutoAdminLogon, DefaultUserName;
    String   Enabled, DefaultPassword, SubKey;
    Registry RegObject;

     if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SECUR_SETTINGS );

    // Get WinLogon data from the registry, these determine if
    // auto-logo is enabled
    RegObject.Connect( HKEY_LOCAL_MACHINE );
    SubKey = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon";
    RegObject.GetStringValue( SubKey.c_str(), L"AutoAdminLogon" , &AutoAdminLogon );
    RegObject.GetStringValue( SubKey.c_str(), L"DefaultUserName", &DefaultUserName );
    RegObject.GetStringValue( SubKey.c_str(), L"DefaultPassword", &DefaultPassword );

    // See if autologon enabled
    Enabled = PXS_STRING_EMPTY;
    if ( DefaultUserName.GetLength() && DefaultPassword.GetLength() )
    {
        Enabled = AutoAdminLogon;
    }
    pRecord->Add( PXS_SECURITY_SETTINGS_ITEM   , L"AutoLogon" );
    pRecord->Add( PXS_SECURITY_SETTINGS_NAME   , L"Enabled" );
    pRecord->Add( PXS_SECURITY_SETTINGS_SETTING, Enabled );
}

//===============================================================================================//
//  Description:
//      Make the audit data records for the automatic update settings
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Remarks:
//      Requires Windows XP or Windows 2000 Professional SP3.
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeAutomaticUpdateRecords(
                                                TArray< AuditRecord >* pRecords)
{
    String      UpdateStatus, UpdateDay;
    HRESULT     hResult = 0;
    AuditRecord Record;
    AutoIUnknownRelease ReleaseAutoUpdate;
    AutoIUnknownRelease ReleasepAutoUpdateSettings;
    IAutomaticUpdates*  pAutoUpdate = nullptr;
    IAutomaticUpdatesSettings* pAutoUpdateSettings = nullptr;
    AutomaticUpdatesNotificationLevel level;
    AutomaticUpdatesScheduledInstallationDay day;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // IAutomaticUpdates
    hResult = CoCreateInstance( CLSID_AutomaticUpdates,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_IAutomaticUpdates,
                                reinterpret_cast<void**>( &pAutoUpdate ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CLSID_AutomaticUpdates", "CoCreateInstance");
    }

    // Always test for NULL pointer
    if ( pAutoUpdate == nullptr )
    {
        throw NullException( L"pAutoUpdate", __FUNCTION__ );
    }
    ReleaseAutoUpdate.Set( pAutoUpdate );

    // Retrieve the update settings
    hResult = pAutoUpdate->get_Settings( &pAutoUpdateSettings );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IAutomaticUpdates::get_Settings", __FUNCTION__ );
    }

    if ( pAutoUpdateSettings == nullptr )
    {
        throw NullException( L"pAutoUpdateSettings", __FUNCTION__ );
    }
    ReleasepAutoUpdateSettings.Set( pAutoUpdateSettings );

    // Retrieve the update status.
    hResult = pAutoUpdateSettings->get_NotificationLevel( &level );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult,
                            L"IAutomaticUpdatesSettings::get_NotificationLevel", __FUNCTION__ );
    }

    // Retrieve the update schedule day.
    hResult = pAutoUpdateSettings->get_ScheduledInstallationDay( &day );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult,
                            L"IAutomaticUpdatesSettings::get_ScheduledInstallationDay",
                            __FUNCTION__ );
    }

    // Add records, starting with the table name
    TranslateAutomaticUpdateStatus( level, &UpdateStatus );
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM   , L"Automatic Updates" );
    Record.Add( PXS_SECURITY_SETTINGS_NAME   , L"Update status" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, UpdateStatus );
    pRecords->Add( Record );

    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM   , L"Automatic Updates" );
    Record.Add( PXS_SECURITY_SETTINGS_NAME   , L"Update schedule" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, UpdateDay );
    pRecords->Add( Record );
}

//===============================================================================================//
//  Description:
//      Make the audit data records for those executables that are allowed
//      to execute data
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Remarks:
//      Requires Windows XP +SP2 or 2003 + SP1
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeExecuteDataRecord(TArray< AuditRecord >* pRecords)
{
    HKEY    hKey    = nullptr;
    LONG    result  = 0;
    DWORD   cValues = 0, cchValueName = 0, cbData = 0, i = 0;
    wchar_t szData[ MAX_PATH + 1 ];
    wchar_t szValueName[ MAX_PATH + 1 ];
    String  Insert1, Insert2;
    LPCWSTR EXE_DATA_KEY = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"
                           L"\\AppCompatFlags\\Layers";
    AuditRecord Record;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Open registry, try the machine first, then the user. If key not found
    // there are no programmes allowed to execute data
    result = RegOpenKeyEx( HKEY_LOCAL_MACHINE,
                           EXE_DATA_KEY, 0, KEY_READ, &hKey );
    if ( result != ERROR_SUCCESS )
    {
        Insert1 = EXE_DATA_KEY;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogSysInfo2( static_cast<DWORD>( result ),
                        L"RegOpenKeyEx for path '%%1' in '%%2'.", Insert1, Insert2 );

        hKey   = nullptr;
        result = RegOpenKeyEx( HKEY_CURRENT_USER, EXE_DATA_KEY, 0, KEY_READ, &hKey );
        if ( result != ERROR_SUCCESS )
        {
            PXSLogSysInfo2( static_cast<DWORD>( result ),
                            L"RegOpenKeyEx for path '%%1' in '%%2'.", Insert1, Insert2 );
            return;
        }
    }

    result = RegQueryInfoKey( hKey,
                              nullptr,
                              nullptr,
                              nullptr,
                              nullptr,
                              nullptr, nullptr, &cValues, nullptr, nullptr, nullptr, nullptr );
    if ( result != ERROR_SUCCESS )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysInfo1( (DWORD)result, L"RegQueryInfoKey failed in '%%1'.", Insert1 );
        return;
    }


    // Catch exception to close the registry key
    try
    {
        // Add the entries under each key to the array
        for ( i = 0; i < cValues; i++ )
        {
            // The subkey are listed in the data execution tab of control
            // panel->system->advanced->performance settings. The boxes that
            // are checked are allowed to execute data and appear in the
            // registry as DisableNXShowUI
            cchValueName = ARRAYSIZE( szValueName );   // In characters
            cbData       = sizeof ( szData );                 // In bytes
            memset( szValueName,  0, sizeof ( szValueName ) );
            memset( szData     , 0 , sizeof ( szData  ) );
            if ( ERROR_SUCCESS == RegEnumValue(
                                         hKey,
                                         i,
                                         szValueName,
                                         &cchValueName,
                                         nullptr,
                                         nullptr,
                                         reinterpret_cast<BYTE*>( &szData ), &cbData ) )
            {
                szValueName[ ARRAYSIZE( szValueName ) - 1 ]   = PXS_CHAR_NULL;
                szData[ ARRAYSIZE( szData ) - 1 ] = PXS_CHAR_NULL;
                if ( CSTR_EQUAL == CompareString( LOCALE_INVARIANT,
                                                  NORM_IGNORECASE,
                                                  szData, -1, L"DisableNXShowUI", -1) )
                {
                   Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
                   Record.Add( PXS_SECURITY_SETTINGS_ITEM, L"Execute Data");
                   Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Program" );
                   Record.Add( PXS_SECURITY_SETTINGS_SETTING, szValueName );
                   pRecords->Add( Record );
                }
            }
        }
    }
    catch ( const Exception& )
    {
        RegCloseKey( hKey );
        throw;
    }
    RegCloseKey( hKey );
}

//===============================================================================================//
//  Description:
//      Make the audit data records for Internet Explorer settings
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeInternetExplorerRecords( TArray< AuditRecord >* pRecords )
{
    DWORD   policy = 0;
    size_t  i = 0;
    String  UrlPolicy, Insert1;
    HRESULT hResult  = 0;
    AuditRecord Record;
    AutoIUnknownRelease   ReleaseInternetMgr;
    IInternetZoneManager* pInternetMgr = nullptr;

    // Structure to hold action codes and names
    struct _POLICY
    {
        DWORD   urlAction;
        LPCWSTR pszName;
    } Policies[] = {
        { URLACTION_SCRIPT_RUN          ,  L"Run script"            },
        { URLACTION_ACTIVEX_RUN         ,  L"Run ActiveX"           },
        { URLACTION_HTML_JAVA_RUN       ,  L"Run Java"              },
        { URLACTION_SHELL_FILE_DOWNLOAD ,  L"Download files"        },
        { URLACTION_SHELL_INSTALL_DTITEMS, L"Install desktop items" },
        { URLACTION_SHELL_VERB          ,  L"Launch applications"   } };

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // IInternetZoneManager
    hResult = CoCreateInstance( CLSID_InternetZoneManager,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_IInternetZoneManager,
                                reinterpret_cast<void**>( &pInternetMgr ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CLSID_InternetZoneManager", "CoCreateInstance" );
    }

    // Always test for NULL pointer
    if ( pInternetMgr == nullptr )
    {
        throw NullException( L"pInternetMgr", __FUNCTION__ );
    }
    ReleaseInternetMgr.Set( pInternetMgr );

    // Cycle through structure for the Internet Zone settings
    for ( i = 0; i < ARRAYSIZE( Policies ); i++ )
    {
        policy    = 0;
        UrlPolicy = PXS_STRING_EMPTY;
        hResult   = pInternetMgr->GetZoneActionPolicy( URLZONE_INTERNET,
                                                       Policies[ i ].urlAction,
                                                       reinterpret_cast<BYTE*>( &policy ),
                                                       sizeof ( policy ), URLZONEREG_HKCU );
        if ( SUCCEEDED( hResult ) )
        {
            TranslateInternetActionPolicy( policy, &UrlPolicy );
        }
        else
        {
            // Log and continue to next setting
            Insert1.SetAnsi( __FUNCTION__ );
            PXSLogComWarn1( hResult, L"GetZoneActionPolicy failed in '%%1'.", Insert1 );
        }

        // Add record, starting with the table name
        Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
        Record.Add( PXS_SECURITY_SETTINGS_ITEM   , L"Internet Explorer" );
        Record.Add( PXS_SECURITY_SETTINGS_NAME   ,  Policies[ i ].pszName );
        Record.Add( PXS_SECURITY_SETTINGS_SETTING, UrlPolicy );
        pRecords->Add( Record );
    }
}

//===============================================================================================//
//  Description:
//      Make the audit data record for network access settings
//
//  Parameters:
//      pRecord - receives the audit record
//
//  Remarks:
//      Requires Windows XP and newer
//
//      Network access: Allow anonymous SID/Name translation (secedit's
//      LSAAnonymousNameLookup) despite this value being set in the Locale
//      Security Policy, WMI/RSOP seems to returns nothing. The same behaviour
//      is observed with wbemtest. So will attempt a SID to name translation
//      and test for access denied
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeNetworkAccessRecord( AuditRecord* pRecord )
{
    PSID    pSid = nullptr;
    DWORD   cbSid = 0, cchDomainName = 0, cchName = 0;
    BYTE    bSID[ SECURITY_MAX_SID_SIZE ] = { 0 };
    wchar_t szSetting[ 32 ] = { 0 };          // "Enabled" or "Disabled"
    wchar_t szAccountName[ UNLEN + 1 ] = { 0 };
    wchar_t szDomainName[ MAX_PATH + 1 ] = { 0 };
    SID_NAME_USE  eUse;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SECUR_SETTINGS );

    cchName = ARRAYSIZE( szAccountName );
    if ( GetUserName( szAccountName, &cchName ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetUserName", __FUNCTION__ );
    }
    szAccountName[ ARRAYSIZE( szAccountName ) - 1 ] = PXS_CHAR_NULL;

    // Get the user's SID
    pSid  = (PSID)bSID;
    cbSid = sizeof ( bSID );
    cchDomainName = ARRAYSIZE( szDomainName );
    if ( LookupAccountName( nullptr,        // Local system
                            szAccountName,
                            pSid, &cbSid, szDomainName, &cchDomainName, &eUse ) == 0 )
    {
        throw SystemException( GetLastError(), L"LookupAccountName", __FUNCTION__ );
    }

    // Impersonate the anonymous account, must match with call to RevertToSelf.
    if ( ImpersonateAnonymousToken( GetCurrentThread() ) == 0 )
    {
        throw SystemException( GetLastError(), L"ImpersonateAnonymousToken", __FUNCTION__ );
    }

    // Now get the SID while anonymous impersonation is in effect
    memset( szAccountName, 0, sizeof ( szAccountName ) );
    cchName = ARRAYSIZE( szAccountName );
    memset( szDomainName, 0, sizeof ( szDomainName ) );
    cchDomainName = ARRAYSIZE( szDomainName );
    if ( LookupAccountSid( nullptr,     // Local computer
                           pSid, szAccountName, &cchName, szDomainName, &cchDomainName, &eUse ) )
    {
        StringCchCopy( szSetting, ARRAYSIZE( szSetting ), L"Enabled" );
    }
    else
    {
        // Access denied means "Disabled", anything else is an error
        if ( ERROR_ACCESS_DENIED == GetLastError() )
        {
            StringCchCopy( szSetting, ARRAYSIZE( szSetting ), L"Disabled" );
        }
        else
        {
            RevertToSelf();
            throw SystemException( GetLastError(), L"LookupAccountSid (anonymous)", __FUNCTION__ );
        }
    }
    RevertToSelf();

    // Make audit record
    pRecord->Add( PXS_SECURITY_SETTINGS_ITEM   , L"Network Access" );
    pRecord->Add( PXS_SECURITY_SETTINGS_NAME   , L"Allow anonymous SID/name translation" );
    pRecord->Add( PXS_SECURITY_SETTINGS_SETTING, szSetting );
}

//===============================================================================================//
//  Description:
//      Get the network security settings as an array of audit records
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeNetworkSecurityRecords( TArray< AuditRecord >* pRecords )
{
    DWORD    lmHash = 0;
    String   StringData;
    Registry RegObject;
    AuditRecord Record;
    SystemInformation  SystemInfo;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Network Security: Do not store LAN Manager hash value on next password
    // change.
    StringData = L"Disabled";
    RegObject.Connect( HKEY_LOCAL_MACHINE );
    RegObject.GetDoubleWordValue( L"SYSTEM\\CurrentControlSet\\Control\\Lsa",
                                  L"nolmhash", 0, &lmHash );
    if ( lmHash )
    {
        StringData = L"Enabled";
    }
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, L"Network Security" );
    Record.Add( PXS_SECURITY_SETTINGS_NAME,
                L"Do not store LAN Manager hash value on next password change" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
    pRecords->Add( Record );

    // Network security: Force logoff when logon hours expire
    // (secedit's ForceLogoffWhenHourExpire). Only applies to administrators
    if ( SystemInfo.IsAdministrator() )
    {
        StringData = PXS_STRING_EMPTY;
        GetRSOPSecuritySetting( L"ForceLogoffWhenHourExpire", VT_BOOL, &StringData );
        Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
        Record.Add( PXS_SECURITY_SETTINGS_ITEM, L"Network Security" );
        Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Force logoff when logon hours expire" );
        Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
        pRecords->Add( Record );
    }
}

//===============================================================================================//
//  Description:
//      Make the audit data records for the password policy
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Remarks:
//      The setting RequireLogonToChangePassword no longer exits, see
//      support.microsoft.com/kb/255776
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakePasswordPolicyRecords( TArray< AuditRecord >* pRecords )
{
    LPCWSTR   STR_PASSWORD_POLICY = L"Password Policy";
    String    StringData, Insert1;
    Formatter Format;
    AuditRecord         Record;
    NET_API_STATUS      status = 0;
    SystemInformation   SystemInfo;
    WindowsInformation  WindowsInfo;
    USER_MODALS_INFO_0* pUMI = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Level 0 on local machine
    status = NetUserModalsGet( nullptr, 0, reinterpret_cast<BYTE**>( &pUMI ) );
    if ( ( NERR_Success == status ) && pUMI )
    {
        AutoNetApiBufferFree AutoFreeUMI( pUMI );

        // Catch and log exceptions as want to proceed to using RSOP
        try
        {
            // Enforce password history
            StringData  = Format.UInt32( pUMI->usrmod0_password_hist_len );
            StringData += L" remembered";
            Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
            Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_PASSWORD_POLICY );
            Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Enforce password history" );
            Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
            pRecords->Add( Record );

            // Maximum password age
            if ( pUMI->usrmod0_max_passwd_age == TIMEQ_FOREVER )
            {
                StringData = L"Forever";
            }
            else
            {
                StringData = Format.UInt32(
                                pUMI->usrmod0_max_passwd_age / (24 * 3600) );
            }
            Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
            Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_PASSWORD_POLICY );
            Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Maximum password age");
            Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
            pRecords->Add( Record );

            // Minimum password age
            StringData  = Format.UInt32(
                                   pUMI->usrmod0_min_passwd_age / (24 * 3600) );
            Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
            Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_PASSWORD_POLICY );
            Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Minimum password age");
            Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
            pRecords->Add( Record );

            // Minimum password length
            StringData  = Format.UInt32( pUMI->usrmod0_min_passwd_len );
            StringData += L" Characters";
            Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
            Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_PASSWORD_POLICY );
            Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Minimum password length" );
            Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
            pRecords->Add( Record );
        }
        catch ( const Exception& e )
        {
            // Log and continue
            PXSLogException( L"Error getting password policy.", e, __FUNCTION__ );
        }
    }
    else
    {
       Insert1.SetAnsi( __FUNCTION__ );
       PXSLogNetError1( status, L"NetUserModalsGet failed or returned NULL in '%%1'.", Insert1 );
    }


    // Use policy or RSOP, must be XP or newer with admin privileges
    if ( WindowsInfo.IsWinXPorNewer() && SystemInfo.IsAdministrator() )
    {
        // Password complexity
        StringData = PXS_STRING_EMPTY;
        GetRSOPSecuritySetting( L"PasswordComplexity", VT_BOOL, &StringData );
        Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
        Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_PASSWORD_POLICY );
        Record.Add( PXS_SECURITY_SETTINGS_NAME, L"Password must meet complexity requirements" );
        Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
        pRecords->Add( Record );

        // Store password using reversible encryption for all users in the
        // domain (secedit's ClearTextPassword)
        StringData = PXS_STRING_EMPTY;
        GetRSOPSecuritySetting( L"ClearTextPassword", VT_BOOL, &StringData );
        Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
        Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_PASSWORD_POLICY );
        Record.Add( PXS_SECURITY_SETTINGS_NAME,
                    L"Store password using reversible encryption for all users in the domain" );
        Record.Add( PXS_SECURITY_SETTINGS_SETTING, StringData );
        pRecords->Add( Record );
    }
    else
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogAppWarn1( L"RSOP requires XP or newer with administrator privileges in '%%1'.",
                        Insert1 );
    }
}

//===============================================================================================//
//  Description:
//      Make the audit data records for the screen saver setting of the
//      current user
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Remarks:
//      MSDN: Keys are
//      - ScreenSaveActive REG_SZ Boolean
//          Specifies whether a screen saver should be displayed if the
//          system is not actively being used. Set this value to 1 to use
//          a screen saver; 0 turns off the screen saver.
//
//      - ScreenSaverIsSecure REG_SZ Boolean
//          Specifies whether a password is assigned to the screen saver.
//
//      - ScreenSaveTimeOut REG_SZ seconds
//          Specifies the amount of time that the system must be idle before
//          the screen saver appears.
//
//      - ScreenSaveActive and ScreenSaverIsSecure can be overridden by
//        Win2000 group policy
//
//      On a workstation the settings are at "HKCU\Control Panel\Desktop"
//      for both Win9x/NT. On a domain that has a group policy its at:
//      "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group
//       Policy Objects\{GUID}User\Software\Policies\Microsoft\Windows
//       \Control Panel\Desktop" where the GUID identifies the group policy.
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::MakeScreenSaverRecords( TArray< AuditRecord >* pRecords )
{
    int      seconds = 0;
    size_t   i = 0, numKeys = 0;
    String   SubKey, Value, GPCPKey, GPOKey;
    String   PolicyObjectKey, ControlPanelKey, Temp, Setting;
    Registry RegObject;
    Formatter   Format;
    AuditRecord Record;
    StringArray GPOKeys, GPCPKeys;
    LPCWSTR STR_TYPE = L"Screen Saver";
    LPCWSTR GPO_KEY  = L"Software\\Microsoft\\Windows\\CurrentVersion\\Group Policy Objects\\";
    LPCWSTR GPCP_KEY = L"Software\\Policies\\Microsoft\\Windows\\Control Panel\\";

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // First look for a local policy setting then override with any domain
    // policy. Will take the first user configuration found.
    SubKey = L"Control Panel\\Desktop";
    RegObject.Connect( HKEY_CURRENT_USER );
    RegObject.GetSubKeyList( GPO_KEY, &GPOKeys );
    numKeys = GPOKeys.GetSize();
    for ( i = 0; i < numKeys; i++ )
    {
        GPOKey = GPOKeys.Get( i );
        if ( GPOKey.EndsWithStringI( L"}User" ) )
        {
            // Make the full path
            Temp  = GPO_KEY;
            Temp += GPOKey;
            Temp += L"\\";
            Temp += GPCP_KEY;
            RegObject.GetSubKeyList( Temp.c_str(), &GPCPKeys );
            if ( GPCPKeys.IndexOf( L"Desktop", false ) != PXS_MINUS_ONE )
            {
                SubKey  = Temp;
                SubKey += L"Desktop";
            }
        }
    }
    PXSLogAppInfo1( L"Using registry key '%%1' for screen saver settings.",
                    SubKey );

    // Determine of the screen saver is enabled
    Setting = PXS_STRING_EMPTY;
    RegObject.GetStringValue( SubKey.c_str(), L"ScreenSaveActive", &Setting );

    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM, STR_TYPE );
    Record.Add( PXS_SECURITY_SETTINGS_NAME,  L"Enabled" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, Setting );
    pRecords->Add( Record );

    // Get the timeout
    Setting = PXS_STRING_EMPTY;
    RegObject.GetStringValue( SubKey.c_str(), L"ScreenSaveTimeOut", &Value );
    if ( Value.c_str() )
    {
        seconds = _wtoi( Value.c_str() );
        if ( seconds >= 0 )
        {
            Setting = Format.Int32( seconds );
        }
    }

    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM   , STR_TYPE );
    Record.Add( PXS_SECURITY_SETTINGS_NAME   , L"Timeout" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, Setting );
    pRecords->Add( Record );

    // See if password protected
    Setting = PXS_STRING_EMPTY;
    RegObject.GetStringValue( SubKey.c_str(), L"ScreenSaverIsSecure", &Setting );
    Record.Reset( PXS_CATEGORY_SECUR_SETTINGS );
    Record.Add( PXS_SECURITY_SETTINGS_ITEM   , STR_TYPE );
    Record.Add( PXS_SECURITY_SETTINGS_NAME   ,  L"Password protected" );
    Record.Add( PXS_SECURITY_SETTINGS_SETTING, Setting );
    pRecords->Add( Record );
}

//===============================================================================================//
//  Description:
//      Translate the automatic updates enumeration constant to a day
//
//  Parameters:
//      day        - the enumeration constant
//      pUpdateDay - receives the day
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::TranslateAutomaticUpdateDay(AutomaticUpdatesScheduledInstallationDay day,
                                                      String* pUpdateDay )
{
    Formatter Format;

    if ( pUpdateDay == nullptr )
    {
        throw ParameterException( L"pUpdateDay", __FUNCTION__ );
    }

    switch ( day )
    {
        default:
            *pUpdateDay = Format.UInt32( day );
            break;

        case ausidEveryDay:
            *pUpdateDay = L"Every day";
            break;

        case ausidEverySunday:
            *pUpdateDay = L"Every Sunday";
            break;

        case ausidEveryMonday:
            *pUpdateDay = L"Every Monday";
            break;

        case ausidEveryTuesday:
            *pUpdateDay = L"Every Tuesday";
            break;

        case ausidEveryWednesday:
            *pUpdateDay = L"Every Wednesday";
            break;

        case ausidEveryThursday:
            *pUpdateDay = L"Every Thursday";
            break;

        case ausidEveryFriday:
            *pUpdateDay = L"Every Friday";
            break;

        case ausidEverySaturday:
            *pUpdateDay = L"Every Saturday";
            break;
    }
}

//===============================================================================================//
//  Description:
//      Translate the automatic updates enumeration constant to a notification
//      level aka status
//
//  Parameters:
//      level        - the enumeration constant
//      UpdateStatus - receives the level/status translation
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::TranslateAutomaticUpdateStatus( AutomaticUpdatesNotificationLevel level,
                                                          String* pUpdateStatus )
{
    Formatter Format;

    if ( pUpdateStatus == nullptr )
    {
        throw ParameterException( L"pUpdateStatus", __FUNCTION__ );
    }

    switch ( level )
    {
        default:
            *pUpdateStatus = Format.UInt32( level );
            break;

        case aunlNotConfigured:
            *pUpdateStatus = L"Not configured";
            break;

        case aunlDisabled:
            *pUpdateStatus = L"Disabled";
            break;

        case aunlNotifyBeforeDownload:
            *pUpdateStatus = L"Notify before download";
            break;

        case aunlNotifyBeforeInstallation:
            *pUpdateStatus = L"Notify before installation";
            break;

        case aunlScheduledInstallation:
            *pUpdateStatus = L"Scheduled installation";
            break;
    }
}

//===============================================================================================//
//  Description:
//      Translate the specified URL action policy to a string
//
//  Parameters:
//      policy       - the action policy
//      pActionPolicy- receives the policy translation
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::TranslateInternetActionPolicy( DWORD policy, String* pActionPolicy )
{
    Formatter Format;

    if ( pActionPolicy == nullptr )
    {
        throw ParameterException( L"pActionPolicy", __FUNCTION__ );
    }

    switch ( policy )
    {
        default:
            *pActionPolicy = Format.UInt32( policy );
            break;

        case URLPOLICY_ALLOW:               // (0x00)
            *pActionPolicy = L"Allow";
            break;

        case URLPOLICY_DISALLOW:            // (0x03)
            *pActionPolicy = L"Not allowed";
            break;

        case URLPOLICY_QUERY:               // (0x01)
            *pActionPolicy = L"Prompt user";
            break;

        case URLPOLICY_MASK_PERMISSIONS:    // (0x0f)
            *pActionPolicy = L"Depends on action.";
            break;
    }
}

//===============================================================================================//
//  Description:
//      Translate the policy audit event type enumeration constant to a
//      string
//
//  Parameters:
//      type            - the enumeration constant
//      pAuditEventType - receives the type
//
//  Returns:
//     void
//===============================================================================================//
void SecurityInformation::TranslatePolicyAuditEventType( POLICY_AUDIT_EVENT_TYPE type,
                                                         String* pAuditEventType )
{
    Formatter Format;

    if ( pAuditEventType == nullptr )
    {
        throw ParameterException( L"pAuditEventType", __FUNCTION__ );
    }

    switch ( type )
    {
        default:
            *pAuditEventType = Format.Int32( type );
            break;

        case AuditCategorySystem:                   // = 0
            *pAuditEventType = L"Audit system events";
            break;

        case AuditCategoryLogon:                    // = 1
            *pAuditEventType = L"Audit logon events";
            break;

        case AuditCategoryObjectAccess:             // = 2
            *pAuditEventType = L"Audit object access";
            break;

        case AuditCategoryPrivilegeUse:             // = 3
            *pAuditEventType = L"Audit privilege use";
            break;

        case AuditCategoryDetailedTracking:         // = 4
            *pAuditEventType = L"Audit process tracking";
            break;

        case AuditCategoryPolicyChange:             // = 5
            *pAuditEventType = L"Audit policy change";
            break;

        case AuditCategoryAccountManagement:        // = 6
            *pAuditEventType = L"Audit account management";
            break;

        case AuditCategoryDirectoryServiceAccess:   // = 7
            *pAuditEventType = L"Audit directory service access";
            break;

        case AuditCategoryAccountLogon:             // = 8
            *pAuditEventType = L"Audit account logon events";
            break;
    }
}
