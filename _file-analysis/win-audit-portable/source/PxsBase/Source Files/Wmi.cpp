///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Windows Management Instrumentation Class Implementation
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
#include "PxsBase/Header Files/Wmi.h"

// 2. C System Files
#include <wchar.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/BStr.h"
#include "PxsBase/Header Files/AutoSysFreeString.h"
#include "PxsBase/Header Files/AutoVariantClear.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

//
// Default constructor
//
Wmi::Wmi()
    :m_pIEnumWbemClassObject( nullptr ),
     m_pIWbemClassObject( nullptr ),
     m_pIWbemLocator( nullptr ),
     m_pIWbemServices( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
Wmi::~Wmi()
{
    try
    {
        Disconnect();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
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
//      Close a query
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Wmi::CloseQuery()
{
    // Result set class object
    if ( m_pIWbemClassObject )
    {
        m_pIWbemClassObject->Release();
        m_pIWbemClassObject = nullptr;
    }

    // Query Enumerator
    if ( m_pIEnumWbemClassObject )
    {
        m_pIEnumWbemClassObject->Release();
        m_pIEnumWbemClassObject = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Connect to a WMI name space.
//
//  Parameters:
//      pszNameSpace - pointer to the name space on the computer
//
//  Remarks:
//      ConnectServer() is a blocking call, if no server found may never
//      return. On XP and above can use WBEM_FLAG_CONNECT_USE_MAX_WAIT to
//      specify a maximum of 2 minutes. Think this only applies to remote
//      servers. Documentation says, call will block if server 'broken'.
//
//  Returns:
//      void
//===============================================================================================//
void Wmi::Connect( LPCWSTR pszNameSpace )
{
    BStr    BinaryString;
    String  NameSpace;
    HRESULT hResult = 0;

    // Avoid WBEM_E_INVALID_NAMESPACE
    NameSpace = pszNameSpace;
    NameSpace.Trim();
    if ( NameSpace.IsEmpty() )
    {
        throw ParameterException( L"pszNameSpace", __FUNCTION__ );
    }
    Disconnect();   // Ensure disconnected before proceeding

    // WbemLocator
    hResult = CoCreateInstance( CLSID_WbemLocator,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_IWbemLocator,
                                reinterpret_cast<void**>( &m_pIWbemLocator ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult,
                            L"CoCreateInstance, IID_IWbemLocator",
                            __FUNCTION__ );
    }

    if ( m_pIWbemLocator == nullptr )
    {
        throw NullException( L"m_pIWbemLocator", __FUNCTION__ );
    }


    // For a local connect, credentials not required
    BinaryString.Allocate( NameSpace );
    hResult = m_pIWbemLocator->ConnectServer( BinaryString.b_str(),
                                              nullptr,  // = current user
                                              nullptr,  // = password
                                              nullptr,
                                              WBEM_FLAG_CONNECT_USE_MAX_WAIT,
                                              nullptr,
                                              nullptr,
                                              &m_pIWbemServices );
    if ( FAILED( hResult ) )
    {
        PXSLogAppInfo1( L"Failed to connect to WMI namespace '%%1'.",
                        NameSpace );
        throw ComException( hResult, L"IWbemLocator::ConnectServer", __FUNCTION__ );
    }

    if ( m_pIWbemServices == nullptr )
    {
        throw NullException( L"m_pIWbemServices", __FUNCTION__ );
    }
    ApplyProxySecurity( m_pIWbemServices );
}

//===============================================================================================//
//  Description:
//      Disconnect from the name space and release any used resources
//
//  Parameters:
//      None
//
//  Remarks:
//      Do not throw any exceptions as called by destructor
//
//      Careful! Observed that if CoUninitialzed has been called on this
//      thread then when call Release an an access violation may occur.
//
//  Returns:
//      void
//===============================================================================================//
void Wmi::Disconnect()
{
    // Doing in reverse order of creation
    if ( m_pIWbemClassObject )
    {
        m_pIWbemClassObject->Release();
        m_pIWbemClassObject = nullptr;
    }

    if ( m_pIEnumWbemClassObject )
    {
        m_pIEnumWbemClassObject->Release();
        m_pIEnumWbemClassObject = nullptr;
    }

    if ( m_pIWbemServices )
    {
        m_pIWbemServices->Release();
        m_pIWbemServices = nullptr;
    }

    if ( m_pIWbemLocator )
    {
        m_pIWbemLocator->Release();
        m_pIWbemLocator = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Fill an array with the name/property values of a WMI class object
//      that was created by a query enumerator
//
//  Parameters:
//      pNameValues - an array to receive the name-value pairs
//
//  Returns:
//      void
//===============================================================================================//
void Wmi::GetPropertyValues( TArray< NameValue >* pNameValues )
{
    long       i = 0, lLbound = 0, lUbound = 0;
    BSTR       pElement = nullptr;
    String     Name, Value;
    HRESULT    hResult = 0;
    VARIANT    variant;
    Formatter  Format;
    NameValue  NameValue;
    SAFEARRAY* pstrNames = nullptr;

    if ( pNameValues == nullptr )
    {
        throw ParameterException( L"pNameValues", __FUNCTION__ );
    }
    pNameValues->RemoveAll();

    if ( m_pIWbemClassObject == nullptr )
    {
        throw ParameterException( L"m_pIWbemClassObject", __FUNCTION__ );
    }
    hResult = m_pIWbemClassObject->GetNames( nullptr, WBEM_FLAG_ALWAYS, nullptr, &pstrNames );
    if ( hResult != WBEM_S_NO_ERROR )
    {
        throw ComException( hResult, L"IWbemClassObject::GetNames", __FUNCTION__ );
    }
    if ( pstrNames == nullptr )
    {
        return;     // No names
    }

    hResult = SafeArrayGetLBound( pstrNames, 1, &lLbound );
    if ( FAILED( hResult ) )
    {
        SafeArrayDestroy( pstrNames );
        throw ComException( hResult, L"SafeArrayGetLBound", __FUNCTION__ );
    }

    hResult = SafeArrayGetUBound( pstrNames, 1, &lUbound );
    if ( FAILED( hResult ) )
    {
        SafeArrayDestroy( pstrNames );
        throw ComException( hResult, L"SafeArrayGetUBound", __FUNCTION__ );
    }

    // Catch exceptions to clean up the array
    try
    {
        for ( i = lLbound; i <= lUbound; i++ )
        {
            // SAFEARRAY is an array of type BSTR
            Name     = PXS_STRING_EMPTY;
            Value    = PXS_STRING_EMPTY;
            pElement = nullptr;
            hResult  = SafeArrayGetElement( pstrNames, &i, &pElement );
            if ( SUCCEEDED( hResult ) && pElement )
            {
                AutoSysFreeString AutoSysFreeElement( pElement );
                Name = pElement;

                // Get the value
                VariantInit( &variant );
                hResult = m_pIWbemClassObject->Get( pElement, 0, &variant, nullptr, nullptr );
                if ( SUCCEEDED( hResult ) )
                {
                    AutoVariantClear VariantClear( &variant );
                    PXSVariantValueToString( &variant, &Value );
                }
                else
                {
                    PXSLogComWarn1( hResult, L"IWbemClassObject::Get failed for '%%1'.", Name );
                }
                NameValue.SetNameValue( Name, Value );
                pNameValues->Add( NameValue );
            }
            else
            {
                PXSLogComWarn( hResult, L"SafeArrayGetElement failed." );
            }
        }
    }
    catch ( const Exception& )
    {
        SafeArrayDestroy( pstrNames );
        throw;
    }
    SafeArrayDestroy( pstrNames );
}

//===============================================================================================//
//  Description:
//      Execute the specified query
//
//  Parameters:
//      pszQuery - the query
//
//  Remarks:
//      Forward-only cursor. Cannot execute a query while one is open.
//
//  Returns:
//      void
//===============================================================================================//
void Wmi::ExecQuery( LPCWSTR pszQuery )
{
    BStr    BStrWQL, BStrQuery;
    String  WQL, Query;
    HRESULT hResult = 0;

    // Ensure have a query
    if ( ( pszQuery == nullptr ) || ( *pszQuery == PXS_CHAR_NULL ) )
    {
        throw ParameterException( L"pszQuery", __FUNCTION__ );
    }

    // Must be connected
    if ( m_pIWbemServices == nullptr )
    {
       throw FunctionException( L"m_pIWbemServices", __FUNCTION__ );
    }

    // Do not allow a query if an enumerator is active
    if ( m_pIEnumWbemClassObject )
    {
        throw FunctionException( L"m_pIEnumWbemClassObject", __FUNCTION__ );
    }

    WQL = L"WQL";
    BStrWQL.Allocate( WQL );
    Query = pszQuery;
    BStrQuery.Allocate( Query );
    hResult = m_pIWbemServices->ExecQuery( BStrWQL.b_str(),
                                           BStrQuery.b_str(),
                                           WBEM_FLAG_FORWARD_ONLY
                                           | WBEM_FLAG_RETURN_IMMEDIATELY,
                                           nullptr,
                                           &m_pIEnumWbemClassObject );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pszQuery, "IWbemServices::ExecQuery" );
    }

    // Always check for a NULL enumerator even on success
    if ( m_pIEnumWbemClassObject == nullptr )
    {
        throw NullException( L"m_pIEnumWbemClassObject", __FUNCTION__ );
    }

    // On an enumerator need to set the proxy's security otherwise can get
    // DCOM access denied = 0x8007005
    ApplyProxySecurity( m_pIEnumWbemClassObject );
}

//===============================================================================================//
//  Description:
//      Get the specified property value as a string
//
//  Parameters:
//      pwszName - the name of the property
//      pValue   - receives the value
//
//  Remarks:
//      Must have executed a query and moved on to a row
//
//  Returns:
//      void
//===============================================================================================//
void Wmi::Get( LPCWSTR pwszName, String* pValue )
{
    VARIANT  variant;
    HRESULT  hResult = 0;

    if ( ( pwszName == nullptr ) || ( *pwszName == PXS_CHAR_NULL ) )
    {
        throw ParameterException( L"pwszName", __FUNCTION__ );
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = PXS_STRING_EMPTY;

    // Cannot get a property unless have a class object
    if ( m_pIWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIWbemClassObject", __FUNCTION__ );
    }

    VariantInit( &variant );
    hResult = m_pIWbemClassObject->Get(pwszName, 0, &variant, nullptr, nullptr);
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pwszName, "IWbemClassObject::Get" );
    }
    AutoVariantClear VariantClearVal( &variant );
    PXSVariantToString( &variant, pValue );
}

//===============================================================================================//
//  Description:
//      Get the specified property as a boolean
//
//  Parameters:
//      pwszName - the name of the property
//      pValue   - receives the value
//
//  Remarks:
//      Must have executed a query and moved on to a row
//
//  Returns:
//      true if non-null otherwise false for null or empty
//===============================================================================================//
bool Wmi::GetBool( LPCWSTR pwszName, bool* pValue )
{
    String    ErrorMessage;
    VARIANT   variant;
    HRESULT   hResult = 0;
    Formatter Format;

    if ( ( pwszName == nullptr ) )
    {
        throw ParameterException( L"pwszName", __FUNCTION__ );
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = false;

    // Cannot get a property unless have a class object
    if ( m_pIWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIWbemClassObject", __FUNCTION__ );
    }

    VariantInit( &variant );
    hResult = m_pIWbemClassObject->Get(pwszName, 0, &variant, nullptr, nullptr);
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pwszName, "IWbemClassObject::Get" );
    }
    AutoVariantClear VariantClear( &variant );

    // Test for NULL/EMPTY
    if ( ( variant.vt == VT_NULL ) || ( variant.vt == VT_EMPTY ) )
    {
        return false;
    }

    // Verify its a 32-bit int
    if ( variant.vt != VT_BOOL )
    {
        ErrorMessage  = pwszName;
        ErrorMessage += Format.StringInt32( L", VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );
    }

    if ( variant.boolVal )
    {
        *pValue = true;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the specified property as signed 32-bit integer
//
//  Parameters:
//      pwszName - the name of the property
//      pValue   - receives the value
//
//  Remarks:
//      Must have executed a query and moved on to a row
//
//  Returns:
//      true if non-null otherwise false for null or empty
//===============================================================================================//
bool Wmi::GetInt32( LPCWSTR pwszName, int* pValue )
{
    String    ErrorMessage;
    VARIANT   variant;
    HRESULT   hResult = 0;
    Formatter Format;

    if ( ( pwszName == nullptr ) )
    {
        throw ParameterException( L"pwszName", __FUNCTION__ );
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    // Cannot get a property unless have a class object
    if ( m_pIWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIWbemClassObject", __FUNCTION__ );
    }

    VariantInit( &variant );
    hResult = m_pIWbemClassObject->Get(pwszName, 0, &variant, nullptr, nullptr);
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pwszName, "IWbemClassObject::Get" );
    }
    AutoVariantClear VariantClear( &variant );

    // Test for NULL/EMPTY
    if ( ( variant.vt == VT_NULL ) || ( variant.vt == VT_EMPTY ) )
    {
        return false;
    }

    // Verify its a 32-bit int
    if ( variant.vt != VT_I4 )
    {
        ErrorMessage  = pwszName;
        ErrorMessage += Format.StringInt32( L", VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );
    }
    *pValue = variant.lVal;

    return true;
}

//===============================================================================================//
//  Description:
//      Get the specified property as unsigned 8-bit integer
//
//  Parameters:
//      pwszName - the name of the property
//      pValue   - receives the byte value
//
//  Remarks:
//      Must have executed a query and moved on to a row
//
//  Returns:
//      true if non-null otherwise false for null or empty
//===============================================================================================//
bool Wmi::GetUInt8( LPCWSTR pwszName, BYTE* pValue )
{
    String    ErrorMessage;
    VARIANT   variant;
    HRESULT   hResult = 0;
    Formatter Format;

    if ( pwszName == nullptr )
    {
        throw ParameterException( L"pwszName", __FUNCTION__ );
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    // Cannot get a property unless have a class object
    if ( m_pIWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIWbemClassObject", __FUNCTION__ );
    }

    VariantInit( &variant );
    hResult = m_pIWbemClassObject->Get( pwszName, 0, &variant, nullptr, nullptr );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pwszName, "IWbemClassObject::Get" );
    }
    AutoVariantClear VariantClear( &variant );

    // Test for NULL/EMPTY
    if ( ( variant.vt == VT_NULL ) || ( variant.vt == VT_EMPTY ) )
    {
        return false;
    }

    // Verify its a 32-bit int
    if ( variant.vt != VT_UI1 )
    {
        ErrorMessage  = pwszName;
        ErrorMessage += Format.StringInt32( L", VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );
    }
    *pValue = variant.bVal;

    return true;
}

//===============================================================================================//
//  Description:
//      Get the specified property as byte array
//
//  Parameters:
//      pwszName    - the name of the property
//      pBuffer     - buffer to receive the data
//      bufferBytes - sizeof the buffer in bytes
//
//  Remarks:
//      Must have executed a query and moved on to a row.
//
//  Returns:
//      Number of bytes copied to the buffer. Returns zero if the array
//      is NULL or EMPTY
//===============================================================================================//
DWORD Wmi::GetUInt8Array( LPCWSTR pwszName, BYTE* pBuffer, DWORD bufferBytes )
{
    long      i = 0, lUpper = 0, lLower = 0;
    BYTE      element = 0;
    DWORD     bytesCopied = 0;
    String    ErrorMessage;
    HRESULT   hResult = 0;
    VARIANT   variant;
    Formatter Format;

    if ( pwszName == nullptr )
    {
        throw ParameterException( L"pwszName", __FUNCTION__ );
    }

    if ( pBuffer == nullptr )
    {
        throw ParameterException( L"pBuffer", __FUNCTION__ );
    }
    memset( pBuffer, 0, bufferBytes );

    // Cannot get a property unless have a class object
    if ( m_pIWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIWbemClassObject", __FUNCTION__ );
    }

    VariantInit( &variant );
    hResult = m_pIWbemClassObject->Get( pwszName, 0, &variant, nullptr, nullptr );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pwszName, "IWbemClassObject::Get" );
    }
    AutoVariantClear VariantClear( &variant );

    // Test for NULL/EMPTY
    if ( variant.parray == nullptr )
    {
        return 0;
    }

    // Make sure result is a byte array
    if ( variant.vt != ( VT_ARRAY | VT_UI1 ) )
    {
        ErrorMessage  = pwszName;
        ErrorMessage += Format.StringInt32( L", VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );
    }

    // Get the array's bounds
    hResult = SafeArrayGetLBound( variant.parray, 1, &lLower );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"SafeArrayGetLBound", __FUNCTION__ );
    }

    hResult = SafeArrayGetUBound( variant.parray, 1, &lUpper );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"SafeArrayGetUBound", __FUNCTION__ );
    }

    // Buffer check
    if ( PXSCastLongToUInt32( lUpper - lLower ) >= bufferBytes )
    {
        throw SystemException( ERROR_INSUFFICIENT_BUFFER, L"bufferBytes", __FUNCTION__ );
    }

    // Get a byte, get a byte, get a byte byte, byte
    for ( i = lLower; i <= lUpper; i++ )
    {
        element = 0;
        hResult = SafeArrayGetElement( variant.parray, &i, &element );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"SafeArrayGetElement", __FUNCTION__ );
        }
        pBuffer[ i ] = element;
        bytesCopied = PXSAddUInt32( bytesCopied, 1 );
    }

    return bytesCopied;
}

//===============================================================================================//
//  Description:
//      Get the specified property as unsigned 16-bit integer
//
//  Parameters:
//      pwszName - the name of the property
//      pValue   - receives the value
//
//  Remarks:
//      Must have executed a query and moved on to a row
//
//  Returns:
//      true if non-null otherwise false for null or empty
//===============================================================================================//
bool Wmi::GetUInt16( LPCWSTR pwszName, WORD* pValue )
{
    String    ErrorMessage, Insert2;
    VARIANT   variant;
    HRESULT   hResult = 0;
    Formatter Format;

    if ( pwszName == nullptr )
    {
        throw ParameterException( L"pwszName", __FUNCTION__ );
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    // Cannot get a property unless have a class object
    if ( m_pIWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIWbemClassObject", __FUNCTION__ );
    }

    VariantInit( &variant );
    hResult = m_pIWbemClassObject->Get(pwszName, 0, &variant, nullptr, nullptr);
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pwszName, "IWbemClassObject::Get" );
    }
    AutoVariantClear VariantClear( &variant );

    // Test for NULL/EMPTY
    if ( ( variant.vt == VT_NULL ) || ( variant.vt == VT_EMPTY ) )
    {
        return false;
    }

    // Verify its a 16-bit int, although will accept VT_I4
    if ( variant.vt == VT_UI2 )
    {
        *pValue = variant.uiVal;    // = USHORT
    }
    else if ( variant.vt == VT_I4 )
    {
        ErrorMessage = pwszName;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo2( L"Processing '%%1' as VT_I4 data type in '%%2'.", ErrorMessage, Insert2 );
        *pValue = PXSCastInt32ToUInt16( variant.iVal );
    }
    else
    {
        ErrorMessage  = pwszName;
        ErrorMessage += Format.StringInt32( L", VARTYPE=%%1.", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the specified property as unsigned 32-bit integer
//
//  Parameters:
//      pwszName - the name of the property
//      pValue   - receives the value
//
//  Remarks:
//      Must have executed a query and moved on to a row
//
//  Returns:
//      true if non-null otherwise false for null or empty
//===============================================================================================//
bool Wmi::GetUInt32( LPCWSTR pwszName, DWORD* pValue )
{
    String    ErrorMessage;
    VARIANT   variant;
    HRESULT   hResult = 0;
    Formatter Format;

    if ( pwszName == nullptr )
    {
        throw ParameterException( L"pwszName", __FUNCTION__ );
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    // Cannot get a property unless have a class object
    if ( m_pIWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIWbemClassObject", __FUNCTION__ );
    }

    VariantInit( &variant );
    hResult = m_pIWbemClassObject->Get(pwszName, 0, &variant, nullptr, nullptr);
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pwszName, "IWbemClassObject::Get" );
    }
    AutoVariantClear VariantClear( &variant );

    // Test for NULL/EMPTY
    if ( ( variant.vt == VT_NULL ) || ( variant.vt == VT_EMPTY ) )
    {
        return false;
    }

    // Expecting an unsigned 32-bit int. However, sometimes a query returns a
    // different data type than specified in a MOF. Handle simple cases.
    if ( ( variant.vt != VT_I4 ) && ( variant.vt != VT_UI4 ) )
    {
        ErrorMessage  = pwszName;
        ErrorMessage += Format.StringInt32( L", VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );
    }

    if ( variant.vt == VT_I4 )
    {
        if ( variant.lVal >= 0 )
        {
            *pValue = PXSCastInt32ToUInt32( variant.lVal );
        }
        else
        {
            ErrorMessage  = pwszName;
            ErrorMessage += Format.StringInt32( L", VARTYPE=%%1", variant.vt );
            throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );
        }
    }
    else
    {
        *pValue = variant.ulVal;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the specified property an unsigned 64-bit integer
//
//  Parameters:
//      pwszName - the name of the property
//      pValue   - receives the value
//
//  Remarks:
//      Must have executed a query and moved on to a row
//
//  Returns:
//     true if non-null otherwise false for null or empty
//===============================================================================================//
bool Wmi::GetUInt64( LPCWSTR pwszName, UINT64* pValue )
{
    String    ErrorMessage, Insert2;
    VARIANT   variant;
    HRESULT   hResult  = 0;
    __int64   n64Value = 0;
    Formatter Format;

    if ( pwszName == nullptr )
    {
        throw ParameterException( L"pwszName", __FUNCTION__ );
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = 0;

    // Cannot get a propety unless have a class object
    if ( m_pIWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIWbemClassObject", __FUNCTION__ );
    }

    VariantInit( &variant );
    hResult = m_pIWbemClassObject->Get(pwszName, 0, &variant, nullptr, nullptr);
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, pwszName, "IWbemClassObject::Get" );
    }
    AutoVariantClear VariantClear( &variant );

    // Test for NULL/EMPTY
    if ( ( variant.vt == VT_NULL ) || ( variant.vt == VT_EMPTY ) )
    {
        return false;
    }

    // Verify its a 32-bit int, sometimes 64-bit fields are strings
    if ( variant.vt == VT_UI8 )
    {
        *pValue = variant.ullVal;
    }
    else if ( variant.vt == VT_BSTR )
    {
        ErrorMessage = pwszName;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo2( L"Processing '%%1' as VT_BSTR data type in '%%2'.",
                        ErrorMessage, Insert2 );
        if ( variant.bstrVal )
        {
            n64Value = _wtoi64( variant.bstrVal );
            *pValue  = PXSCastInt64ToUInt64( n64Value );
        }
    }
    else
    {
        ErrorMessage  = pwszName;
        ErrorMessage += Format.StringInt32( L", VARTYPE=%%1", variant.vt );
        throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the next row of a result set
//
//  Parameters:
//      None
//
//  Remarks:
//      The caller must have executed a query
//
//  Returns:
//     true on success, false because no more rows
//===============================================================================================//
bool Wmi::Next()
{
    const   LONG ENUMERATOR_TIMEOUT = 10000;      // milliseconds
    bool    moved     = false;
    ULONG   uReturned = 0;
    HRESULT hResult   = 0;

    // Cannot proceed without an open query
    if ( m_pIEnumWbemClassObject == nullptr )
    {
        throw FunctionException( L"m_pIEnumWbemClassObject", __FUNCTION__ );
    }

    // Release the existing WBEM class object before getting a new row
    if ( m_pIWbemClassObject )
    {
        m_pIWbemClassObject->Release();
        m_pIWbemClassObject = nullptr;
    }
    hResult = m_pIEnumWbemClassObject->Next( ENUMERATOR_TIMEOUT,
                                             1,                 // Get one row
                                             &m_pIWbemClassObject,
                                             &uReturned );
    if ( SUCCEEDED( hResult ) && m_pIWbemClassObject )
    {
        moved = true;
    }

    return moved;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Apply the current WMI security in effect to a proxy's COM interface
//
//  Parameters:
//      pProxy - pointer to the proxy object's interface
//
//  Remarks:
//      Must have already called the ConnectServer method. Need to set the
//      security on newly created objects else often get DCOM access denied
//      (0x80070005) when getting data from remote computers of Win2000 and
//      newer.
//
//  Returns:
//      void
//===============================================================================================//
void Wmi::ApplyProxySecurity( IUnknown* pProxy )
{
    HRESULT hResult = S_OK;

    if ( pProxy == nullptr )
    {
        throw ParameterException( L"pProxy", __FUNCTION__ );
    }

    hResult = CoSetProxyBlanket( pProxy,
                                 RPC_C_AUTHN_WINNT,
                                 RPC_C_AUTHZ_NONE,
                                 nullptr,
                                 RPC_C_AUTHN_LEVEL_CALL,
                                 RPC_C_IMP_LEVEL_IMPERSONATE,
                                 nullptr,       // Must be NULL on Vista+
                                 EOAC_NONE );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult,
                            L"CoSetProxyBlanket", __FUNCTION__ );
    }
}
