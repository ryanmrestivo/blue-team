///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Object Permission Information Implementation
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
#include "WinAudit/Header Files/ObjectPermissionInformation.h"

// 2. C System Files
#include <AclAPI.h>
#include <Iads.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/Ddk.h"
#include "WinAudit/Header Files/PrinterInformation.h"
#include "WinAudit/Header Files/WindowsNetworkInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ObjectPermissionInformation::ObjectPermissionInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
ObjectPermissionInformation::~ObjectPermissionInformation()
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
//      Get printer and share permissions
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::GetAuditRecords( TArray< AuditRecord >* pRecords )
{
    size_t      i = 0, numElements = 0;
    String      Name;
    StringArray PrinterNames, ShareNames;
    PrinterInformation        PrinterInfo;
    TArray< AuditRecord >     PermissionRecords;
    WindowsNetworkInformation WindowsNetworkInfo;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Printers
    PrinterInfo.GetPrinterNames( &PrinterNames );
    numElements = PrinterNames.GetSize();
    for ( i = 0; i < numElements; i++ )
    {
        Name = PrinterNames.Get( i );
        PermissionRecords.RemoveAll();
        GetObjectPermissionRecords( Name, SE_PRINTER, &PermissionRecords );
        pRecords->Append( PermissionRecords );
    }

    // Shares
    WindowsNetworkInfo.GetShareNames( &ShareNames );
    numElements = ShareNames.GetSize();
    for ( i = 0; i < numElements; i++ )
    {
        Name = ShareNames.Get( i );
        PermissionRecords.RemoveAll();
        GetObjectPermissionRecords( Name, SE_LMSHARE, &PermissionRecords );
        pRecords->Append( PermissionRecords );
    }
    PXSSortAuditRecords( pRecords, PXS_PERMISSIONS_OBJECT_NAME );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the account name from the specified security identifier
//
//  Parameters:
//      pSid         - pointer to the SID
//      pAccountName - receives the account name
//
//  Returns:
//      void, empty string if SID not recognised
//===============================================================================================//
void ObjectPermissionInformation::GetAccountNameFromSid( const PSID pSid, String* pAccountName )
{
    DWORD   lastError = 0, cchName = 0, cchDomainName = 0;
    wchar_t szName[ MAX_PATH + 1 ] = { 0 };
    wchar_t szDomainName[ MAX_PATH + 1 ] = { 0 };
    Formatter    Format;
    SID_NAME_USE eUse = SidTypeUnknown;

    if ( pAccountName == nullptr )
    {
        throw ParameterException( L"pAccountName", __FUNCTION__ );
    }
    *pAccountName = PXS_STRING_EMPTY;

    if ( pSid == nullptr )
    {
        return;     // No account name
    }

    cchName = ARRAYSIZE( szName );
    cchDomainName = ARRAYSIZE( szDomainName );
    if ( LookupAccountSid( nullptr,     // = local machine
                           pSid, szName, &cchName, szDomainName, &cchDomainName, &eUse ) )
    {
        szName[ ARRAYSIZE( szName ) - 1 ] = PXS_CHAR_NULL;
        szDomainName[ ARRAYSIZE( szDomainName ) - 1 ] = PXS_CHAR_NULL;

        // Pre-pend the domain name if have it
        if ( szDomainName[ 0 ] != PXS_CHAR_NULL )
        {
            *pAccountName  = szDomainName;
            *pAccountName += PXS_PATH_SEPARATOR;
        }
        *pAccountName += szName;
    }
    else
    {
        // May not get an account name. e.g. ERROR_NONE_MAPPED. This happens
        // for some ACEs of SE_WINDOWS_OBJECT. In which case will convert
        // the SID to a string
        lastError   = GetLastError();
        *pAccountName = Format.SidToString( pSid );
        PXSLogSysError1( lastError,
                        L"LookupAccountSid failed for SID = '%%1'.", *pAccountName );
    }
}

//===============================================================================================//
//  Description:
//      Get the desktop's security information
//
//  Parameters:
//      ppSecurityDescriptor - receives the security descriptor
//
//  Remarks:
//      The input parameter is overwritten, it is the caller's responsibility
//      to ensure that ppSecurityDescriptor is freed with LocalFree.
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::GetDesktopSecurity( PSECURITY_DESCRIPTOR* ppSecurityDescriptor )
{
    HDESK hDesk     = nullptr;
    DWORD errorCode = 0;

    // Inputs must be nullptr to avoid memory leaks
    if ( ppSecurityDescriptor )
    {
        throw ParameterException( L"ppSecurityDescriptor", __FUNCTION__ );
    }

    hDesk = OpenDesktop( L"Default", 0, FALSE, GENERIC_READ );  // = Desktop
    if ( hDesk == nullptr )
    {
        throw SystemException( GetLastError(), L"OpenDesktop", __FUNCTION__ );
    }

    errorCode = GetSecurityInfo( hDesk,
                                 SE_WINDOW_OBJECT,
                                 DACL_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION,
                                 nullptr,
                                 nullptr, nullptr, nullptr, ppSecurityDescriptor );

    // Finished with the handle, so clean up before testing for error
    CloseDesktop( hDesk );
    if ( errorCode != ERROR_SUCCESS )
    {
        throw SystemException( errorCode, L"GetSecurityInfo", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the security information for a non-window type object
//
//  Parameters:
//      ObjectName  - the object's name
//      objectType - defined enumeration constant, see SE_OBJECT_TYPE
//      ppSecurityDescriptor - receives the security descriptor
//
//  Remarks:
//      Non-window means desktop or workstation, these are handled
//      differently
//
//      ppSecurityDescriptor is overwritten, it is the caller's responsibility
//      to ensure that ppSecurityDescriptor is freed with LocalFree.
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::GetNonWindowSecurity(const String& ObjectName,
                                                       SE_OBJECT_TYPE objectType,
                                                       PSECURITY_DESCRIPTOR* ppSecurityDescriptor )
{
    DWORD    errorCode;
    size_t   numChars;
    wchar_t* pszName   = nullptr;
    AllocateWChars Alloc;

    if ( ObjectName.IsEmpty() )
    {
        throw ParameterException( L"ObjectName", __FUNCTION__ );
    }

    // GetNamedSecurityInfo wants a non-constant string pointer
    numChars = ObjectName.GetLength();
    numChars = PXSAddSizeT( numChars, 1 );
    pszName  = Alloc.New( numChars );
    StringCchCopy( pszName, numChars, ObjectName.c_str() );
    errorCode = GetNamedSecurityInfo(
                         pszName,
                         objectType,
                         DACL_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION,
                         nullptr, nullptr, nullptr, nullptr, ppSecurityDescriptor );
    if ( errorCode != ERROR_SUCCESS )
    {
        throw SystemException( errorCode, L"GetNamedSecurityInfo", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the permissions of the specified object
//
//  Parameters:
//      ObjectName - the object's name
//      objectType - defined enumeration constant of the object type
//      pRecords   - array to receive the data
//
//  Remarks:
//      Rights are tabulated in "Provider-Independent and Windows NT Access
//      Rights" and found in the SDK header files
//
//      Do not sort the permissions because the ordering is important, e.g. if
//      an Access Denied entry is found, then is will override any subsequent
//      entry further down the list.
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::GetObjectPermissionRecords( const String& ObjectName,
                                                              SE_OBJECT_TYPE objectType,
                                                              TArray< AuditRecord >* pRecords )
{
    bool workStation = false;
    PSECURITY_DESCRIPTOR pSecurityDescriptor = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    if ( ObjectName.IsEmpty() )
    {
        throw ParameterException( L"ObjectName", __FUNCTION__ );
    }

    // SE_WINDOW_OBJECT uses GetSecurityInfo, non-window objects use
    // GetNamedSecurityInfo
    if ( objectType == SE_WINDOW_OBJECT )
    {
        // SE_WINDOW_OBJECT is either a Window Station or Desktop
        if ( ObjectName.CompareI( L"Winsta0" ) == 0 )
        {
            workStation = true;
            GetWorkStationSecurity( &pSecurityDescriptor );
        }
        else if ( ObjectName.CompareI( L"Default" ) == 0 )
        {
            workStation = false;
            GetDesktopSecurity( &pSecurityDescriptor );
        }
        else
        {
            throw ParameterException( L"ObjectName", __FUNCTION__ );
        }
    }
    else
    {
        GetNonWindowSecurity( ObjectName, objectType, &pSecurityDescriptor );
    }

    // Catch as must free the security descriptor
    try
    {
        if ( pSecurityDescriptor )
        {
            MakeObjectPermissionRecords( ObjectName,
                                         objectType,
                                         workStation, pSecurityDescriptor, pRecords );
        }
        else
        {
            PXSLogAppInfo1( L"Object '%%1' has no security set.", ObjectName );
        }
    }
    catch ( const Exception& )
    {
        if ( pSecurityDescriptor )
        {
            LocalFree( pSecurityDescriptor );
        }
        throw;
    }

    if ( pSecurityDescriptor )
    {
        LocalFree( pSecurityDescriptor );
    }
}

//===============================================================================================//
//  Description:
//      Get the security information for the workstation
//
//  Parameters:
//      ppSecurityDescriptor - receives the security descriptor
//
//  Remarks:
//      The input parameters is overwritten, it is the caller's responsibility
//      to ensure that ppSecurityDescriptor is freed with LocalFree.
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::GetWorkStationSecurity(
                                                       PSECURITY_DESCRIPTOR* ppSecurityDescriptor )
{
    DWORD   errorCode = 0;
    HWINSTA hWinsta = nullptr;

    // Inputs must be nullptr to avoid memory leaks
    if ( ppSecurityDescriptor )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }

    // Winsta0 = Workstation
    hWinsta = OpenWindowStation( L"Winsta0", FALSE, GENERIC_READ );
    if ( hWinsta == nullptr )
    {
        throw SystemException( GetLastError(), L"OpenWindowStation", __FUNCTION__ );
    }

    errorCode = GetSecurityInfo(
                        hWinsta,
                        SE_WINDOW_OBJECT,
                        DACL_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION,
                        nullptr, nullptr, nullptr, nullptr, ppSecurityDescriptor );

    // Finished with the handle, so clean up before testing for error
    CloseWindowStation( hWinsta );
    if ( errorCode != ERROR_SUCCESS )
    {
        throw SystemException( errorCode, L"GetSecurityInfo", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Determine if the access mask of the specified object give
//      "All Access"
//
//  Parameters:
//      accessMask  - the ACE access mask
//      type        - the type of object
//      workStation - flag to indicate its a windows station object
//
//  Remarks:
//      PRINTER_ALL_ACCESS does not permit Manage Documents so the trustee
//      does not have Full Control so do will not check for it here.
//
//  Returns:
//      true if all access is granted, otherwise false
//===============================================================================================//
bool ObjectPermissionInformation::HasAllAccess( ACCESS_MASK accessMask,
                                                SE_OBJECT_TYPE type, bool workStation )
{
    bool  allAccess  = false;
    DWORD access = 0;

    if ( type == SE_FILE_OBJECT )
    {
        access = FILE_ALL_ACCESS;
    }
    else if ( type == SE_SERVICE )
    {
        access = SERVICE_ALL_ACCESS;
    }
    else if ( type == SE_LMSHARE )
    {
        access = ACCESS_ALL;       // See LMaccess.h
    }
    else if ( type == SE_WINDOW_OBJECT )
    {
        // This can refer to either Workstation or Desktop.
        if ( workStation )
        {
            access = WINSTA_ALL_ACCESS;
        }
        else
        {
            // The Desktop does not have an "ALL_ACCESS" defined in WinUser.h
            // so make one similar to WINSTA_ALL_ACCESS
            access = DESKTOP_READOBJECTS     |
                     DESKTOP_CREATEWINDOW    |
                     DESKTOP_CREATEMENU      |
                     DESKTOP_HOOKCONTROL     |
                     DESKTOP_JOURNALRECORD   |
                     DESKTOP_JOURNALPLAYBACK |
                     DESKTOP_ENUMERATE       |
                     DESKTOP_WRITEOBJECTS    |
                     DESKTOP_SWITCHDESKTOP;
        }
    }

    if ( access & accessMask )
    {
        allAccess = true;
    }

    return allAccess;
}

//===============================================================================================//
//  Description:
//      Get the permissions of the specified object as an array of audit
//      records
//
//  Parameters:
//      ObjectName  - the name of the object
//      type        - the type of object
//      workStation - flag to indicate its a windows station object
//      ppSecurityDescriptor - receives the security descriptor
//      pRecords    - string array to receive the formatted data
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::MakeObjectPermissionRecords(
                                                         const String& ObjectName,
                                                         SE_OBJECT_TYPE type,
                                                         bool workStation,
                                                         PSECURITY_DESCRIPTOR pSecurityDescriptor,
                                                         TArray< AuditRecord >* pRecords )
{
    WORD   i = 0;
    BOOL   fOwnerDefaulted = FALSE, daclPresent = FALSE, daclDefaulted = FALSE;
    PACL   pDacl = nullptr;
    PSID   pSidOwner = nullptr;
    String Owner, ObjectType, AceType, AceFlags, AceMask;
    String Permissions, Trustee;
    Formatter   Format;
    AuditRecord Record;
    ACCESS_ALLOWED_ACE* pACE = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Must have a name
    if ( ObjectName.IsEmpty() )
    {
        throw ParameterException( L"ObjectName", __FUNCTION__ );
    }
    TranslateObjectType( type, &ObjectType );

    // Must have a security descriptor otherwise no security
    if ( pSecurityDescriptor == nullptr )
    {
        Record.Reset( PXS_CATEGORY_PERMISSIONS );
        Record.Add( PXS_PERMISSIONS_OBJECT_NAME,  ObjectName );
        Record.Add( PXS_PERMISSIONS_OBJECT_TYPE,  ObjectType );
        Record.Add( PXS_PERMISSIONS_TRUSTEE    ,  L"No Trustee" );
        Record.Add( PXS_PERMISSIONS_ACE_TYPE   ,  PXS_STRING_EMPTY );
        Record.Add( PXS_PERMISSIONS_PERMISSIONS,  L"No security" );
        Record.Add( PXS_PERMISSIONS_ACE_FLAGS  ,  PXS_STRING_EMPTY );
        Record.Add( PXS_PERMISSIONS_ACCESS_MASK,  PXS_STRING_EMPTY );
        Record.Add( PXS_PERMISSIONS_OBJECT_OWNER, PXS_STRING_EMPTY );
        pRecords->Add( Record );
        return;
    }


    // Get the owning SID so can get the owner's name
    if ( GetSecurityDescriptorOwner( pSecurityDescriptor, &pSidOwner, &fOwnerDefaulted ) )
    {
        GetAccountNameFromSid( pSidOwner, &Owner );
    }
    else
    {
        // Log and continue as this information is optional
        PXSLogSysError( GetLastError(), L"GetSecurityDescriptorOwner failed." );
    }

    // A nullptr DACL is not an error as not all objects are have security. In
    // fact, MSDN: "A value of TRUE for lpbDaclPresent does not mean that
    // pDacl is not nullptr."
    GetSecurityDescriptorDacl( pSecurityDescriptor, &daclPresent, &pDacl, &daclDefaulted );
    if ( pDacl )
    {
        // Get each ACE in the DACL
        for ( i = 0; i < pDacl->AceCount; i++ )
        {
            pACE = nullptr;     // pACE does not need to be freed
            if ( GetAce( pDacl, i, reinterpret_cast<void**>( &pACE ) ) == 0 )
            {
                throw SystemException( GetLastError(), L"GetAce", __FUNCTION__ );
            }

            if ( pACE == nullptr )
            {
                throw NullException( L"pACE", __FUNCTION__ );
            }

            Trustee = PXS_STRING_EMPTY;
            GetAccountNameFromSid( (PSID)&pACE->SidStart, &Trustee );

            AceType = PXS_STRING_EMPTY;
            TranslateAceType( pACE->Header.AceType, &AceType );

            Permissions = PXS_STRING_EMPTY;
            TranslateAccessMask( pACE->Mask, type, workStation, &Permissions );

            AceFlags = PXS_STRING_EMPTY;
            TranslateAceFlags( pACE->Header.AceFlags, &AceFlags );

            AceMask = Format.BinaryUInt32( pACE->Mask );

            // Make the record
            Record.Reset( PXS_CATEGORY_PERMISSIONS );
            Record.Add( PXS_PERMISSIONS_OBJECT_NAME,  ObjectName );
            Record.Add( PXS_PERMISSIONS_OBJECT_TYPE,  ObjectType );
            Record.Add( PXS_PERMISSIONS_TRUSTEE    ,  Trustee );
            Record.Add( PXS_PERMISSIONS_ACE_TYPE   ,  AceType );
            Record.Add( PXS_PERMISSIONS_PERMISSIONS,  Permissions );
            Record.Add( PXS_PERMISSIONS_ACE_FLAGS  ,  AceFlags );
            Record.Add( PXS_PERMISSIONS_ACCESS_MASK,  AceMask );
            Record.Add( PXS_PERMISSIONS_OBJECT_OWNER, Owner );
            pRecords->Add( Record );
        }
    }
    else
    {
        // Object has no permissions so add a single entry to the output array
        // If the object does not have a DACL, the system grants full access
        // to everyone.
        Record.Reset( PXS_CATEGORY_PERMISSIONS );
        Record.Add( PXS_PERMISSIONS_OBJECT_NAME,  ObjectName );
        Record.Add( PXS_PERMISSIONS_OBJECT_TYPE,  ObjectType );
        Record.Add( PXS_PERMISSIONS_TRUSTEE    ,  L"No Trustee" );
        Record.Add( PXS_PERMISSIONS_ACE_TYPE   ,  PXS_STRING_EMPTY );
        Record.Add( PXS_PERMISSIONS_PERMISSIONS,  L"No security" );
        Record.Add( PXS_PERMISSIONS_ACE_FLAGS  ,  PXS_STRING_EMPTY );
        Record.Add( PXS_PERMISSIONS_ACCESS_MASK,  PXS_STRING_EMPTY );
        Record.Add( PXS_PERMISSIONS_OBJECT_OWNER, Owner );
        pRecords->Add( Record );
    }
}

//===============================================================================================//
//  Description:
//      Translate the an access mask to a string
//
//  Parameters:
//      accessMask  - the ACE access mask
//      type        - the type of object
//      workStation - flag to indicate its a windows station object
//      pTranslation- string to receive the translation
//
//  Remarks:
//      For display purposes use the same descriptions as Explorer when
//      possible, display full control first, then standard rights then
//      specific rights. The GENERIC_XXX flags are shown in Explorer as
//      Special Access, so show them last.
//
//      Window station access rights are shown in MSDN: Window Station
//      Security and Access Rights
//
//      Will ignore ACCESS_READ, ACCESS_WRITE, ACCESS_EXEC, ACCESS_DELETE,
//      ACCESS_ATRIB and ACCESS_PERM present in LMaccess.h
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateAccessMask( ACCESS_MASK accessMask,
                                                       SE_OBJECT_TYPE type,
                                                       bool workStation,
                                                       String* pTranslation )
{
    String SpecificRights, GenericRights;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    // Check for all access. However, PRINTER_ALL_ACCESS does not give
    // permit Manage Documents so the trustee does not have Full Control
    // so do will not check for it here
    if ( HasAllAccess( accessMask, type, workStation ) )
    {
        *pTranslation = L"Full Control";
    }
    else if ( type == SE_LMSHARE )
    {
        // Shares are displayed without specific or advance rights in Explorer
        // so handle these separately. Explorer shows these for Read:
        // Read = 00000000000100100000000010101001 = 0x001200A9
        // = ACCESS_READ | ACCESS_EXEC | ACCESS_ATRIB | 0x80 | READ_CONTROL
        //   | SYNCHRONIZE
        // The bit set at position 8 appears to be undocumented or is perhaps
        // FILE_READ_ATTRIBUTES
        if ( accessMask & 0x001200A9 )
        {
            *pTranslation = L"Read";
        }

        // Change = 00000000000100110000000110111111 = 0x001301BF
        // = ACCESS_READ | ACCESS_WRITE | ACCESS_CREATE | ACCESS_EXEC
        //  | ACCESS_DELETE | ACCESS_ATRIB | DELETE | READ_CONTROL | SYNCHRONIZE
        if ( accessMask & 0x001301BF )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += L"Change";
        }
    }
    else
    {
        // Standard rights, bits 16 - 23
        TranslateStandardRights( accessMask & STANDARD_RIGHTS_ALL, pTranslation );

        // Specific rights, bits 0 - 15
        TranslateSpecificRights( accessMask & SPECIFIC_RIGHTS_ALL,
                                 type, workStation, &SpecificRights );
        if ( pTranslation->GetLength() && SpecificRights.GetLength() )
        {
            *pTranslation += L", ";
        }
        *pTranslation += SpecificRights;

        // Generic rights, bits 28 - 31
        TranslateGenericRights( accessMask & 0xF0000000, &GenericRights );
        if ( pTranslation->GetLength() && GenericRights.GetLength() )
        {
            *pTranslation += L", ";
        }
        *pTranslation += GenericRights;
    }
}

//===============================================================================================//
//  Description:
//      Translate the specified ACE flags to a string
//
//  Parameters:
//        aceFlags     - Defined bitmap constant of ACE flags
//        pTranslation - String to receive the ACE flags
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateAceFlags( LONG aceFlags, String* pTranslation )
{
    size_t i = 0;

    // Structure to hold flags
    struct _STRUCT_FLAGS
    {
        LONG    flag;
        LPCWSTR pszFlag;
    } flags[] = { { CONTAINER_INHERIT_ACE    ,  L"Container Inherit" },
                  { FAILED_ACCESS_ACE_FLAG   ,  L"Failed"            },
                  { INHERIT_ONLY_ACE         ,  L"Inherit Only"      },
                  { INHERITED_ACE            ,  L"Inherited"         },
                  { NO_PROPAGATE_INHERIT_ACE ,  L"No Inherit"        },
                  { OBJECT_INHERIT_ACE       ,  L"Object Inherit"    },
                  { SUCCESSFUL_ACCESS_ACE_FLAG, L"Success"           } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( flags ); i++ )
    {
        if ( aceFlags & flags[ i ].flag )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += flags[ i ].pszFlag;
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate the specified ACE type to a string
//
//  Parameters:
//      aceType      - Defined constant of ACE type
//      pTranslation - String to receive the ACE type
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateAceType( LONG type, String* pTranslation )
{
    size_t i = 0;
    Formatter Format;

    // Structure to hold ACE types
    struct _STRUCT_TYPES
    {
        LONG    type;
        LPCWSTR pszType;
    } types[] =
      { { ACCESS_ALLOWED_ACE_TYPE               ,  L"Allow"                },
        { ACCESS_ALLOWED_CALLBACK_ACE_TYPE      ,  L"Allow Callback"       },
        { ACCESS_ALLOWED_CALLBACK_OBJECT_ACE_TYPE, L"Allow Callback Object"},
        { ACCESS_ALLOWED_COMPOUND_ACE_TYPE      ,  L"Allow Compound"       },
        { ACCESS_ALLOWED_OBJECT_ACE_TYPE        ,  L"Allow Object"         },
        { ACCESS_DENIED_ACE_TYPE                ,  L"Deny"                 },
        { ACCESS_DENIED_CALLBACK_ACE_TYPE       ,  L"Deny Callback"        },
        { ACCESS_DENIED_CALLBACK_OBJECT_ACE_TYPE,  L"Deny Callback Object" },
        { ACCESS_DENIED_OBJECT_ACE_TYPE         ,  L"Deny Object"          },
        { SYSTEM_AUDIT_ACE_TYPE                 ,  L"Audit"                },
        { SYSTEM_ALARM_ACE_TYPE                 ,  L"Alarm"                },
        { SYSTEM_ALARM_CALLBACK_ACE_TYPE        ,  L"Alarm Callback"       },
        { SYSTEM_ALARM_CALLBACK_OBJECT_ACE_TYPE ,  L"Alarm Callback Object"},
        { SYSTEM_ALARM_OBJECT_ACE_TYPE          ,  L"Alarm Object"         },
        { SYSTEM_AUDIT_CALLBACK_ACE_TYPE        ,  L"Audit Callback"       },
        { SYSTEM_AUDIT_CALLBACK_OBJECT_ACE_TYPE ,  L"Audit Callback Object"},
        { SYSTEM_AUDIT_OBJECT_ACE_TYPE          ,  L"Audit Object"         }
      };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( types ); i++ )
    {
        if ( type == types[ i ].type )
        {
            *pTranslation = types[ i ].pszType;
            break;
        }
    }

    // If got no ACE type will use its numerical value
    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.Int32( type );
        PXSLogAppWarn1( L"Unrecognised ACE type = %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate the Active Directory Service specific rights bit mask to
//      a string
//
//  Parameters:
//      specificRights - the bit mask of specific rights
//      pTranslation   - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateAdsSpecificRights( DWORD specificRights,
                                                              String* pTranslation )
{
    size_t i = 0;

    // Structure to hold rights
    struct _STRUCT_RIGHTS
    {
        DWORD   right;
        LPCWSTR pszRight;
    } rights[] =
        { { ADS_RIGHT_DS_CREATE_CHILD ,  L"ADS_RIGHT_DS_CREATE_CHILD"   },
          { ADS_RIGHT_DS_DELETE_CHILD ,  L"ADS_RIGHT_DS_DELETE_CHILD"   },
          { ADS_RIGHT_ACTRL_DS_LIST   ,  L"ADS_RIGHT_ACTRL_DS_LIST"     },
          { ADS_RIGHT_DS_SELF         ,  L"ADS_RIGHT_DS_SELF"           },
          { ADS_RIGHT_DS_READ_PROP    ,  L"ADS_RIGHT_DS_READ_PROP"      },
          { ADS_RIGHT_DS_WRITE_PROP   ,  L"ADS_RIGHT_DS_WRITE_PROP"     },
          { ADS_RIGHT_DS_DELETE_TREE  ,  L"ADS_RIGHT_DS_DELETE_TREE"    },
          { ADS_RIGHT_DS_LIST_OBJECT  ,  L"ADS_RIGHT_DS_LIST_OBJECT"    },
          { ADS_RIGHT_DS_CONTROL_ACCESS, L"ADS_RIGHT_DS_CONTROL_ACCESS" }
        };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( rights ); i++ )
    {
        if ( specificRights & rights[ i ].right )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += rights[ i ].pszRight;
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate the desktop specific rights bit mask to a string
//
//  Parameters:
//      specificRights - the bit mask of specific rights
//      pTranslation   - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateDesktopSpecificRights( DWORD specificRights,
                                                                  String* pTranslation )
{
    size_t i = 0;

    // Structure to hold rights
    struct _STRUCT_RIGHTS
    {
        DWORD   right;
        LPCWSTR pszRight;
    } Rights[] = { { DESKTOP_READOBJECTS    ,  L"Read Objects"     },
                    { DESKTOP_CREATEWINDOW  ,  L"Create Window"    },
                    { DESKTOP_CREATEMENU    ,  L"Create Menu"      },
                    { DESKTOP_HOOKCONTROL   ,  L"Hook Control"     },
                    { DESKTOP_JOURNALRECORD ,  L"Journal Record"   },
                    { DESKTOP_JOURNALPLAYBACK, L"Journal Playback" },
                    { DESKTOP_ENUMERATE     ,  L"Enumerate"        },
                    { DESKTOP_WRITEOBJECTS  ,  L"Write Objects"    },
                    { DESKTOP_SWITCHDESKTOP ,  L"Switch Desktop"   } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Rights ); i++ )
    {
        if ( specificRights & Rights[ i ].right )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += Rights[ i ].pszRight;
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate the file specific rights bit mask to a string
//
//  Parameters:
//      specificRights - the bit mask of specific rights
//      pTranslation   - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateFileSpecificRights( DWORD specificRights,
                                                               String* pTranslation )
{
    size_t i = 0;

    // Structure to hold rights
    struct _STRUCT_RIGHTS
    {
        DWORD   right;
        LPCWSTR pszRight;
    } rights[] =
        { { FILE_READ_DATA      ,  L"Read Data / List Folder"        },
          { FILE_WRITE_DATA     ,  L"Write Data / Create Files"      },
          { FILE_APPEND_DATA    ,  L"Append Data / Create Folders"   },
          { FILE_READ_EA        ,  L"Read Extended Attributes"       },
          { FILE_WRITE_EA       ,  L"Write Extended Attributes"      },
          { FILE_EXECUTE        ,  L"Execute File / Traverse Folder" },
          { FILE_DELETE_CHILD   ,  L"Delete Subfolders and Files"    },
          { FILE_READ_ATTRIBUTES,  L"Read Attributes"                },
          { FILE_WRITE_ATTRIBUTES, L"Write Attributes"               } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( rights ); i++ )
    {
        if ( specificRights & rights[ i ].right )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += rights[ i ].pszRight;
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate the specified generic rights bit mask to a string
//
//  Parameters:
//      genericRights - the bit mask of generic rights
//      pTranslation  - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateGenericRights( DWORD genericRights,
                                                          String* pTranslation )
{
    size_t i = 0;

    // Structure to hold rights
    struct _STRUCT_RIGHTS
    {
        DWORD   right;
        LPCWSTR pszRight;
    } rights[] = { { GENERIC_ALL   ,  L"GENERIC_ALL"     },
                   { GENERIC_EXECUTE, L"GENERIC_EXECUTE" },
                   { GENERIC_WRITE ,  L"GENERIC_WRITE"   },
                   { GENERIC_READ  ,  L"GENERIC_READ"    } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( rights ); i++ )
    {
        if ( genericRights & rights[ i ].right )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += rights[ i ].pszRight;
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate a defined security object type constant
//
//  Parameters:
//      type         - the security object type
//      pTranslation - receives the object type
//
//  Remarks:
//      See SE_OBJECT_TYPE
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateObjectType( SE_OBJECT_TYPE type,
                                                       String* pTranslation )
{
    size_t    i = 0;
    Formatter Format;

    // Structure to hold object types
    struct _STRUCT_TYPES
    {
        SE_OBJECT_TYPE type;
        LPCWSTR        pszType;
    } types[] =
        { { SE_UNKNOWN_OBJECT_TYPE   ,  L"Unknown"               },
          { SE_FILE_OBJECT           ,  L"File"                  },
          { SE_SERVICE               ,  L"Service"               },
          { SE_PRINTER               ,  L"Printer"               },
          { SE_REGISTRY_KEY          ,  L"Registry Key"          },
          { SE_LMSHARE               ,  L"Share"                 },
          { SE_KERNEL_OBJECT         ,  L"Kernel"                },
          { SE_DS_OBJECT             ,  L"Directory Service"     },
          { SE_DS_OBJECT_ALL         ,  L"Directory Service All" },
          { SE_PROVIDER_DEFINED_OBJECT, L"Provider Defined"      },
          { SE_WMIGUID_OBJECT        ,  L"WMI"                   },
          { SE_REGISTRY_WOW64_32KEY  ,  L"Registry Key WOW64"    } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( types ); i++ )
    {
        if ( type == types[ i ].type )
        {
            *pTranslation = types[ i ].pszType;
            break;
        }
    }

    // If got no object type will use its numerical value
    if ( pTranslation->IsEmpty() )
    {
        *pTranslation = Format.Int32( type );
        PXSLogAppWarn1( L"Unrecognised object type = %%1.", *pTranslation );
    }
}

//===============================================================================================//
//  Description:
//      Translate the printer specific rights bit mask to a string
//
//  Parameters:
//      specificRights - the bit mask of specific rights
//      pTranslation   - receives the translation
//
//
//  Remarks:
//      Will ignore the individual rights present in Winspool.h e.g.
//      SERVER_ACCESS_ADMINISTER, SERVER_ACCESS_ENUMERATE etc. Instead
//      will use known flag combinations that permit printer functionality
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslatePrinterSpecificRights(
                                                          DWORD specificRights,
                                                          String* pTranslation )
{
    DWORD required = 0;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    if ( specificRights == 0 )
    {
        return;     // No rights
    }

    // Manage is equivalent to PRINTER_ALL_ACCESS
    if ( PRINTER_ALL_ACCESS & specificRights )
    {
        *pTranslation += L"Manage Printers, ";
    }

    // Print requires PRINTER_ACCESS_USE | READ_CONTROL
    required = PRINTER_ACCESS_USE | READ_CONTROL;
    if ( required & specificRights )
    {
        *pTranslation += L"Print, ";
    }

    // Manage Documents requires JOB_ACCESS_ADMINISTER
    //                           | JOB_ACCESS_READ | STANDARD_RIGHTS_REQUIRED
    // JOB_ACCESS_READ = 0x00000020
    required = JOB_ACCESS_ADMINISTER | 0x00000020 | STANDARD_RIGHTS_REQUIRED;
    if ( required & specificRights )
    {
        *pTranslation += L"Manage Documents";
    }

    pTranslation->Trim();
    if ( pTranslation->EndsWithCharacter( PXS_CHAR_COMMA ) )
    {
        pTranslation->Truncate( pTranslation->GetLength() - 1 );
    }
}

//===============================================================================================//
//  Description:
//      Translate the service specific rights bit mask to a string
//
//  Parameters:
//      specificRights - the bit mask of specific rights
//      pTranslation   - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateServiceSpecificRights( DWORD specificRights,
                                                                  String* pTranslation )
{
    size_t i = 0;

    // Structure to hold rights
    struct _STRUCT_RIGHTS
    {
        DWORD   right;
        LPCWSTR pszRight;
    } rights[] =
        { { SERVICE_QUERY_CONFIG       ,  L"Query Config"         },
          { SERVICE_CHANGE_CONFIG      ,  L"Change Config"        },
          { SERVICE_QUERY_STATUS       ,  L"Query Status"         },
          { SERVICE_ENUMERATE_DEPENDENTS, L"Enumerate Dependants" },
          { SERVICE_START              ,  L"Start"                },
          { SERVICE_STOP               ,  L"Stop"                 },
          { SERVICE_INTERROGATE        ,  L"Interrogate"          },
          { SERVICE_USER_DEFINED_CONTROL, L"User Defined Control" } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( rights ); i++ )
    {
        if ( specificRights & rights[ i ].right )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += rights[ i ].pszRight;
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate the specific rights bit mask to a string
//
//  Parameters:
//      specificRights - the bit mask of Specific rights
//      type           - an enumeration constant representing the object type
//      workStation    - flag to indicate this is a workstation type object
//      pTranslation   - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateSpecificRights( DWORD specificRights,
                                                           SE_OBJECT_TYPE type,
                                                           bool workStation,
                                                           String* pTranslation)
{
    String    Insert2;
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    if ( specificRights == 0 )
    {
        return;     // No rights
    }

    if ( type == SE_FILE_OBJECT )
    {
        TranslateFileSpecificRights( specificRights, pTranslation );
    }
    else if ( type == SE_SERVICE )
    {
        TranslateServiceSpecificRights( specificRights, pTranslation );
    }
    else if ( type == SE_PRINTER )
    {
        TranslatePrinterSpecificRights( specificRights, pTranslation );
    }
    else if ( type == SE_WINDOW_OBJECT )
    {
        // This can refer to either Windows station or Desktop
        if ( workStation )
        {
            TranslateWorkstationSpecificRights( specificRights, pTranslation );
        }
        else
        {
            TranslateDesktopSpecificRights( specificRights, pTranslation );
        }
    }
    else if ( type == SE_DS_OBJECT )
    {
        TranslateAdsSpecificRights( specificRights, pTranslation );
    }
    else
    {
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppWarn2( L"Object type %%1 not recognised in '%%2'.",
                        Format.Int32( type ), Insert2 );
    }
}

//===============================================================================================//
//  Description:
//      Translate the specified standard rights bit mask to a string
//
//  Parameters:
//      standardRights - the bit mask of standard rights
//      pTranslation   - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateStandardRights( DWORD standardRights,
                                                           String* pTranslation)
{
    size_t i = 0;

    // Structure to hold rights
    struct _STRUCT_RIGHTS
    {
        DWORD   right;
        LPCWSTR pszRight;
    } rights[] = { { DELETE      , L"Delete"             },
                   { READ_CONTROL, L"Read Permissions"   },
                   { WRITE_DAC   , L"Change Permissions" },
                   { WRITE_OWNER , L"Take Ownership"     },
                   { SYNCHRONIZE , L"Synchronize"        } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    if ( standardRights == 0 )
    {
        return;     // No rights
    }

    for ( i = 0; i < ARRAYSIZE( rights ); i++ )
    {
        if ( standardRights & rights[ i ].right )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += rights[ i ].pszRight;
        }
    }
}

//===============================================================================================//
//  Description:
//      Translate the workstation specific rights bit mask to a string
//
//  Parameters:
//      specificRights - the bit mask of specific rights
//      pTranslation   - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void ObjectPermissionInformation::TranslateWorkstationSpecificRights( DWORD specificRights,
                                                                      String* pTranslation )
{
    size_t i = 0;

    // Structure to hold rights
    struct _STRUCT_RIGHTS
    {
        DWORD   right;
        LPCWSTR pszRight;
    } rights[] = { { WINSTA_ENUMDESKTOPS    ,  L"Enumerate Desktops"  },
                   { WINSTA_READATTRIBUTES  ,  L"Read Attributes"     },
                   { WINSTA_ACCESSCLIPBOARD ,  L"Access Clipboard"    },
                   { WINSTA_CREATEDESKTOP   ,  L"Create Desktop"      },
                   { WINSTA_WRITEATTRIBUTES ,  L"Write Attributes"    },
                   { WINSTA_ACCESSGLOBALATOMS, L"Access Global Atoms" },
                   { WINSTA_EXITWINDOWS     ,  L"Exit Windows"        },
                   { WINSTA_ENUMERATE       ,  L"Enumerate"           },
                   { WINSTA_READSCREEN      ,  L"Read Screen"         } };

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    if ( specificRights == 0 )
    {
        return;     // No rights
    }

    for ( i = 0; i < ARRAYSIZE( rights ); i++ )
    {
        if ( specificRights & rights[ i ].right )
        {
            if ( pTranslation->GetLength() )
            {
                *pTranslation += L", ";
            }
            *pTranslation += rights[ i ].pszRight;
        }
    }
}
