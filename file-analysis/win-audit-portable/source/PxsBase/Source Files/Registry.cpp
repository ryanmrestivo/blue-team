///////////////////////////////////////////////////////////////////////////////////////////////////
//
// System Registry Class Implementation
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
// Template Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/Registry.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/ParameterException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Registry::Registry()
         :m_hKey( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
Registry::~Registry()
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
//      Connect to the local computer's registry for read access
//
//  Parameters:
//      hKey - handle to registry key, typically HKEY_CURRENT_USER or
//             HKEY_LOCAL_MACHINE
//
//  Returns:
//      void
//===============================================================================================//
void Registry::Connect( HKEY hKey )
{
    LONG result = 0;

    // Reset class scope variables, disconnect if already connected, this sets m_hKey to NULL
    Disconnect();

    result = RegOpenKeyEx( hKey, nullptr, 0, KEY_READ, &m_hKey );
    if ( result != ERROR_SUCCESS )
    {
        throw SystemException( static_cast<DWORD>( result ), L"RegOpenKeyEx", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Connect to the local computer's registry with the specified access
//
//  Parameters:
//      hKey       - handle to registry key, typically HKEY_CURRENT_USER or HKEY_LOCAL_MACHINE
//      samDesired - desired access mask
//
//  Returns:
//      void
//===============================================================================================//
void Registry::Connect2( HKEY hKey, DWORD samDesired )
{
    LONG result = 0;

    // Reset class scope variables, disconnect if already connected, this sets m_hKey to NULL
    Disconnect();

    result = RegOpenKeyEx( hKey, nullptr, 0, samDesired, &m_hKey );
    if ( result != ERROR_SUCCESS )
    {
        throw SystemException( static_cast<DWORD>( result ), L"RegOpenKeyEx", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Disconnect from the registry
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Registry::Disconnect()
{
    // Close handle if connected
    if ( m_hKey )
    {
        RegCloseKey( m_hKey );
        m_hKey = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Get binary data from the registry
//
//  Parameters:
//      pszSubKey    - pointer to registry subkey name
//      pszValueName - pointer to name of value to get
//      pBinaryData  - buffer to receive the registry data.
//      bufferBytes  - On input the length of the buffer in bytes, on output
//                     the number of bytes written
//  Remarks:
//      Will avoid throwing on registry operation error

//  Returns:
//      Registry error code
//===============================================================================================//
DWORD Registry::GetBinaryData( LPCWSTR pszSubKey,
                               LPCWSTR pszValueName, BYTE* pBinaryData, size_t bufferBytes )
{
    HKEY   hKey   = nullptr;
    LONG   result = ERROR_SUCCESS;
    DWORD  type   = 0, cbData = 0;
    String Insert1, Insert2;

    if ( m_hKey == nullptr )
    {
        throw FunctionException( L"m_hKey", __FUNCTION__ );
    }

    // NB null or empty string for pszValueName is OK
    if ( ( pszSubKey   == nullptr ) ||
         ( pBinaryData == nullptr ) ||
         ( bufferBytes == 0       )  )
    {
        throw ParameterException( L"nullptr/zero", __FUNCTION__ );
    }
    memset( pBinaryData, 0, bufferBytes );
    cbData  = PXSCastSizeTToUInt32( bufferBytes );

    result = RegOpenKeyEx( m_hKey, pszSubKey, 0, KEY_READ, &hKey );
    if ( result != ERROR_SUCCESS )
    {
        /* This can create a lot of messages, will ignore
        Insert1 = pszSubKey;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogSysWarn2( static_cast<DWORD>( result ),
                        L"RegOpenKeyEx for '%%1' in '%%2'.", Insert1, Insert2 );
        */
        return PXSCastLongToUInt32( result );
    }

    result = RegQueryValueEx( hKey,
                              pszValueName, nullptr, &type, pBinaryData, &cbData );
    if ( result == ERROR_SUCCESS )
    {
        // Verify the type
        if ( type != REG_BINARY )
        {
            result = ERROR_INVALID_DATATYPE;
        }
    }
    RegCloseKey( hKey );

    /* This can create a lot of messages, will ignore
    // On error, log and continue, ignore key not found as often caller
    // does not know if it exists.
    if ( ( result != ERROR_SUCCESS ) &&
         ( result != ERROR_FILE_NOT_FOUND ) )
    {
        Insert1 = pszSubKey;
        Insert2 = pszValueName;
        PXSLogSysWarn2( static_cast<DWORD>( result ),
                        L"RegQueryValueEx for '%%1' and '%%2'.", Insert1, Insert2 );
    }
    */
    return PXSCastLongToUInt32( result );
}

//===============================================================================================//
//  Description:
//      Get the value of a DWORD in the registry
//
//  Parameters:
//      pszSubKey        - pointer to registry subkey name
//      pszValueName     - pointer to name of value to get
//      defaultValue     - default value to return in case of error
//      pDoubleWordValue - variable to receive result
//
//  Remarks:
//      Will avoid throwing on registry operation error
//
//  Returns:
//      void
//===============================================================================================//
void Registry::GetDoubleWordValue( LPCWSTR pszSubKey,
                                   LPCWSTR pszValueName,
                                   DWORD defaultValue, DWORD* pDoubleWordValue )
{
    LONG   result = 0;
    HKEY   hKey   = nullptr;
    DWORD  type   = 0, cbData = 0, data = 0;
    String Insert1, Insert2;

    if ( m_hKey == nullptr )
    {
        throw FunctionException( L"m_hKey", __FUNCTION__ );
    }

    if ( pszSubKey == nullptr )
    {
        throw ParameterException( L"pszSubKey", __FUNCTION__ );
    }

    if ( pDoubleWordValue == nullptr )
    {
        throw ParameterException( L"pDoubleWordValue", __FUNCTION__ );
    }
    *pDoubleWordValue = defaultValue;

    result = RegOpenKeyEx( m_hKey, pszSubKey, 0, KEY_READ, &hKey );
    if ( result != ERROR_SUCCESS )
    {
        Insert1 = pszSubKey;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogSysWarn2( static_cast<DWORD>( result ),
                        L"RegOpenKeyEx for '%%1' in '%%2'.", Insert1, Insert2 );
        return;
    }

    cbData = sizeof ( DWORD );    // Its a DWORD, so 4 bytes
    result = RegQueryValueEx( hKey,
                              pszValueName, nullptr, &type, (LPBYTE)&data, &cbData );
    if ( result == ERROR_SUCCESS )
    {
        // Verify the type
        if ( type == REG_DWORD )
        {
            *pDoubleWordValue = data;
        }
        else
        {
            result = ERROR_INVALID_DATATYPE;
        }
    }
    RegCloseKey( hKey );

    // On error, log and continue, ignore key not found as often caller
    // does not know if the key exists.
    if ( ( result != ERROR_SUCCESS ) &&
         ( result != ERROR_FILE_NOT_FOUND ) )
    {
        Insert1 = pszSubKey;
        Insert2 = pszValueName;
        PXSLogSysWarn2( static_cast<DWORD>( result ),
                        L"RegQueryValueEx for '%%1' and '%%2'.", Insert1, Insert2 );
    }
}

//===============================================================================================//
//  Description:
//      Get a list of name value pairs in the registry key
//
//  Parameters:
//      pszSubKey      - registry subkey name
//      pNameValueList - array to receive the name/value pairs
//
//  Remarks:
//      Will avoid throwing on registry operation error
//
//  Returns:
//      void
//===============================================================================================//
void Registry::GetNameValueList( LPCWSTR pszSubKey, TArray< NameValue >* pNameValueList )
{
    HKEY      hKey   = nullptr;
    LONG      result = 0;
    DWORD     cchValueName = 0, cbData = 0, cValues  = 0, i = 0, type = 0;
    BYTE      data[ 2048 ] = { 0 };            // Limit to this, max is 32K
    String    Name, Value, Insert1, Insert2;
    wchar_t   szValueName [ MAX_PATH + 1 ] = { 0 };
    NameValue NameValue;

    if ( m_hKey == nullptr )
    {
        throw FunctionException( L"m_hKey", __FUNCTION__ );
    }

    // Null or empty string for pszValueName is OK
    if ( pszSubKey == nullptr )
    {
        throw ParameterException( L"pszSubKey", __FUNCTION__ );
    }

    if ( pNameValueList == nullptr )
    {
        throw ParameterException( L"pNameValueList", __FUNCTION__ );
    }
    pNameValueList->RemoveAll();

    result = RegOpenKeyEx( m_hKey, pszSubKey, 0, KEY_READ, &hKey );
    if ( result != ERROR_SUCCESS )
    {
        return;    // Almost always because the key does not exist
    }

    result = RegQueryInfoKey( hKey,
                              nullptr,
                              nullptr,
                              nullptr,
                              nullptr,
                              nullptr,
                              nullptr,
                              &cValues,
                              nullptr, nullptr, nullptr, nullptr );
    if ( result != ERROR_SUCCESS )
    {
        RegCloseKey( hKey );
        Insert2 = pszSubKey;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogSysError2( static_cast<DWORD>( result ),
                        L"RegQueryInfoKey error '%%1' in '%%2'.", Insert1, Insert2 );
        return;
    }

    try  // Catch exceptions because need to close registry key
    {
        // Add the entries under each key to the array
        for ( i = 0; i < cValues ; i++ )
        {
            memset( szValueName, 0, sizeof ( szValueName ) );
            memset( data, 0, sizeof ( data ) );
            cchValueName = ARRAYSIZE( szValueName );    // In characters
            cbData = sizeof ( data );                   // In bytes
            type   = REG_NONE;
            result = RegEnumValue( hKey,
                                   i, szValueName, &cchValueName, nullptr, &type, data, &cbData );
            if ( result == ERROR_SUCCESS )
            {
                szValueName[ ARRAYSIZE( szValueName ) - 1 ] = PXS_CHAR_NULL;
                Name = szValueName;
                Name.Trim();
                if ( Name.GetLength() && ( cbData < sizeof ( data ) ) )
                {
                    Value = PXS_STRING_EMPTY;
                    FormatRegistryDataAsString( data, cbData, type, &Value );
                    NameValue.SetNameValue( Name, Value );
                    pNameValueList->Add( NameValue );
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
//        Get the value of a string entry in the registry
//
//  Parameters:
//      pszSubKey       - pointer to registry subkey name
//      pszValueName    - pointer to name of value to get
//      pStringValue    - String object to receive the registry entry
//
//  Remarks:
//      Will avoid throwing on registry operation error
//
//  Returns:
//      void.
//===============================================================================================//
void Registry::GetStringValue( LPCWSTR pszSubKey, LPCWSTR pszValueName, String* pStringValue )
{
    HKEY   hKey   = nullptr;
    LONG   result = 0;
    BYTE*  pData  = nullptr;
    DWORD  cbData = 0, cbRequired = 0, type = REG_NONE;
    String Insert1, Insert2;
    AllocateBytes AllocBytes;

    if ( m_hKey == nullptr )
    {
        throw FunctionException( L"m_hKey", __FUNCTION__ );
    }

    // Null or empty string for pszValueName is OK
    if ( pszSubKey == nullptr )
    {
        throw ParameterException( L"pszSubKey", __FUNCTION__ );
    }

    if ( pStringValue == nullptr )
    {
        throw ParameterException( L"pStringValue", __FUNCTION__ );
    }
    *pStringValue = PXS_STRING_EMPTY;

    hKey   = nullptr;
    result = RegOpenKeyEx( m_hKey, pszSubKey, 0, KEY_READ, &hKey );
    if ( result != ERROR_SUCCESS )
    {
        Insert1 = pszSubKey;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogSysWarn2( static_cast<DWORD>( result ),
                        L"RegOpenKeyEx for '%%1' in '%%2'.", Insert1, Insert2 );
        return;
    }

    // Get the required buffer size
    result = RegQueryValueEx( hKey, pszValueName, nullptr, nullptr, nullptr, &cbRequired );
    if ( result != ERROR_SUCCESS )
    {
        // Almost always because the key does not exist
        return;
    }

    // Catch exceptions because need to release registry key
    try
    {
        cbRequired = PXSAddUInt32( cbRequired, sizeof (wchar_t) );  // +NULL
        pData      = AllocBytes.New( cbRequired );
        cbData     = cbRequired;
        result     = RegQueryValueEx( hKey, pszValueName, nullptr, &type, pData, &cbData );
        if ( result == ERROR_SUCCESS )
        {
            // Validate the data type
            if ( ( type == REG_SZ        ) ||
                 ( type == REG_MULTI_SZ  ) ||
                 ( type == REG_EXPAND_SZ )  )
            {
                FormatRegistryDataAsString( pData, cbRequired, type, pStringValue );
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
//      Get a list of subkeys in the registry
//
//  Parameters:
//      pszSubKey  - pointer to registry subkey name, can be NULL
//      SubKeyList - array to receive the subkeys
//
//  Remarks:
//      Will avoid throwing on registry operation error
//
//  Returns:
//      void
//===============================================================================================//
void Registry::GetSubKeyList( LPCWSTR pszSubKey, StringArray* pSubKeyList ) const
{
    LONG    result = ERROR_SUCCESS;
    HKEY    hKey   = nullptr;
    DWORD   i = 0, cSubKeys = 0, cName = 0;
    String  Key, Insert1, Insert2;
    wchar_t szName[ MAX_PATH + 1 ] = { 0 };

    if ( m_hKey == nullptr )
    {
        throw FunctionException( L"m_hKey", __FUNCTION__ );
    }

    if ( pSubKeyList == nullptr )
    {
        throw ParameterException( L"pSubKeyList", __FUNCTION__ );
    }
    pSubKeyList->RemoveAll();

    // Open the key for read access
    result = RegOpenKeyEx( m_hKey, pszSubKey, 0, KEY_READ, &hKey );
    if ( result != ERROR_SUCCESS )
    {
        Insert1 = pszSubKey;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogSysWarn2( static_cast<DWORD>( result ),
                        L"RegOpenKeyEx for '%%1' in '%%2'.", Insert1, Insert2 );
        return;
    }

    result = RegQueryInfoKey( hKey,
                              nullptr,
                              nullptr,
                              nullptr,
                              &cSubKeys,
                              nullptr,
                              nullptr,
                              nullptr, nullptr, nullptr, nullptr, nullptr );
    if ( result != ERROR_SUCCESS )
    {
        RegCloseKey( hKey );
        Insert1 = pszSubKey;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogSysError2( static_cast<DWORD>( result ),
                         L"RegQueryInfoKey error '%%1' in '%%2'.", Insert1, Insert2 );
        return;
    }

    //  Catch exceptions to clean up
    try
    {
        for ( i = 0; i < cSubKeys ; i++ )
        {
            memset( szName, 0, sizeof ( szName ) );
            cName  = ARRAYSIZE( szName );
            result = RegEnumKeyEx( hKey,
                                   i,
                                   szName,
                                   &cName,  // In chars
                                   nullptr, nullptr, nullptr, nullptr );
            if ( result == ERROR_SUCCESS )
            {
                szName[ ARRAYSIZE( szName ) - 1 ] = PXS_CHAR_NULL;
                Key = szName;
                pSubKeyList->Add( Key );
            }
            else
            {
                // Log it and continue
                Insert1 = pszSubKey;
                Insert2.SetAnsi( __FUNCTION__ );
                PXSLogSysWarn2( static_cast<DWORD>( result ),
                                L"RegEnumKeyEx at '%%1' in '%%2'.", Insert1, Insert2 );
            }
        }
    }
    catch ( const Exception& )
    {
        RegCloseKey(hKey);
        throw;
    }
    RegCloseKey(hKey);
}

//===============================================================================================//
//  Description:
//      Get a value in the registry as a string
//
//  Parameters:
//      pszSubKey    - pointer to registry subkey name
//      pszValueName - pointer to name of value to get, can be NULL
//      pStringValue - String object to receive the registry entry.
//
//  Returns:
//      void
//===============================================================================================//
void Registry::GetValueAsString( LPCWSTR pszSubKey, LPCWSTR pszValueName, String* pStringValue )
{
    HKEY      hKey   = nullptr;
    LONG      result = ERROR_SUCCESS;
    BYTE      data[ 2048 ] = { 0 };     // Enough for practical data, max is 32K
    DWORD     type = REG_NONE, cbData = 0;
    String    Insert1, Insert2;
    Formatter Format;

    if ( m_hKey == nullptr )
    {
        throw FunctionException( L"m_hKey", __FUNCTION__ );
    }

    if ( pStringValue == nullptr )
    {
        throw ParameterException( L"pStringValue", __FUNCTION__ );
    }
    *pStringValue = PXS_STRING_EMPTY;

    result = RegOpenKeyEx( m_hKey, pszSubKey, 0, KEY_READ, &hKey );
    if ( result != ERROR_SUCCESS )
    {
        Insert1 = pszSubKey;
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogSysWarn2( static_cast<DWORD>( result ),
                        L"RegOpenKeyEx for '%%1' in '%%2'.", Insert1, Insert2 );
        return;
    }

    cbData = sizeof ( data );          // In bytes
    result = RegQueryValueEx( hKey, pszValueName, nullptr, &type, data, &cbData );
    if ( result != ERROR_SUCCESS )
    {
        RegCloseKey( hKey );
        if ( result != ERROR_FILE_NOT_FOUND )
        {
            Insert1 = pszSubKey;
            Insert2.SetAnsi( __FUNCTION__ );
            PXSLogSysWarn2( static_cast<DWORD>( result ),
                            L"RegQueryValueEx for '%%1' and '%%2'.", Insert1, Insert2 );
        }
        return;
    }

    try
    {
        if ( cbData < sizeof ( data ) )
        {
            FormatRegistryDataAsString( data, cbData, type, pStringValue );
        }
        else
        {
            PXSLogSysError1( ERROR_INSUFFICIENT_BUFFER,
                             L"RegQueryValueEx said valueLen=%%1 bytes", Format.UInt32( cbData ) );
        }
    }
    catch ( const Exception& )
    {
        RegCloseKey( hKey );
        throw;
    }
    RegCloseKey( hKey );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Format registry binary data as a string
//
//  Parameters:
//      pData       - a byte buffer holding the value
//      dataBytes    - size of the buffer in bytes
//      pStringValue - receives the value as a string
//
//  Returns:
//      void
//===============================================================================================//
void Registry::FormatRegistryBinaryAsString( const BYTE* pData,
                                             DWORD dataBytes, String* pStringValue )
{
    int    cb = PXSCastUInt32ToInt32( dataBytes );
    size_t i  = 0, numChars = 0;
    Formatter Format;

    if ( pStringValue == nullptr )
    {
        throw ParameterException( L"pStringValue", __FUNCTION__ );
    }

    if ( pData == nullptr )
    {
        *pStringValue = L"<null>";
        return;
    }
    *pStringValue = PXS_STRING_EMPTY;

    if ( dataBytes == 0 )
    {
        return;     // No data
    }

    // Sometimes Unicode strings are stored in the registry in binary.
    // These string always seem to NULL terminated so only accept if the
    // binary data ends in 0x0000
    if ( ( dataBytes > 2 ) &&
         ( pData[ dataBytes - 1 ] == 0x00 ) &&
         ( pData[ dataBytes - 2 ] == 0x00 ) &&
         ( IsTextUnicode( pData, cb, nullptr ) ) )
    {
        *pStringValue = reinterpret_cast<LPCWSTR>( pData );
    }
    else
    {
        // Format as 'xx xx xx xx...', i.e. 3 bytes per character.
        numChars = PXSMultiplySizeT( 3, dataBytes );
        pStringValue->Allocate( numChars );
        for ( i = 0; i < dataBytes; i++ )
        {
            *pStringValue += Format.UInt8Hex( pData[ i ], false );
        }
    }
}

//===============================================================================================//
//  Description:
//      Format registry value as a string
//
//  Parameters:
//      pData        - a byte buffer holding the value
//      dataBytes    - size of the buffer in bytes
//      type         - named constant of the type of registry data
//      pStringValue - receives the value as a string
//
//  Returns:
//      void
//===============================================================================================//
void Registry::FormatRegistryDataAsString( const BYTE* pData,
                                           DWORD dataBytes, DWORD type, String* pStringValue )
{
    DWORD     regDWord = 0;
    size_t    numChars = 0;
    wchar_t*  pszValue = nullptr;
    Formatter Format;
    AllocateWChars AllocWChars;

    if ( pStringValue == nullptr )
    {
        throw ParameterException( L"pStringValue", __FUNCTION__ );
    }

    if ( pData == nullptr )
    {
        *pStringValue = L"<null>";
        return;
    }

    switch ( type )
    {
        // Format as binary, including REG_NONE, REG_QWORD and REG_BINARY
        default:
            FormatRegistryBinaryAsString( pData, dataBytes, pStringValue );
            break;

        case REG_DWORD:

            // Make sure input buffer is in range
            if ( dataBytes >= sizeof ( DWORD ) )   // = 4
            {
                regDWord = (DWORD)( ( pData[ 0 ] * 0x00000001 ) +
                                    ( pData[ 1 ] * 0x00000100 ) +
                                    ( pData[ 2 ] * 0x00010000 ) +
                                    ( pData[ 3 ] * 0x01000000 ) );
                *pStringValue = Format.UInt32( regDWord );
            }
            break;

        case REG_EXPAND_SZ:
        case REG_MULTI_SZ:
        case REG_SZ:

            // Make a copy of the data because want to ensure its double
            // NULL terminated
            numChars = dataBytes / sizeof ( wchar_t );
            numChars = PXSAddSizeT( numChars, 2 );  // Two NULLs
            pszValue = AllocWChars.New( numChars );
            memcpy( pszValue, pData, dataBytes );
            pszValue[ numChars - 2 ] = PXS_CHAR_NULL;
            pszValue[ numChars - 1 ] = PXS_CHAR_NULL;
            if ( type == REG_EXPAND_SZ )
            {
                PXSExpandEnvironmentStrings( pszValue, pStringValue );
            }
            else if ( type == REG_MULTI_SZ )
            {
                FormatRegistryMultiSzAsString( pszValue, pStringValue );
            }
            else
            {
                *pStringValue = pszValue;   // REG_SZ
            }
            break;
    }
}

//===============================================================================================//
//  Description:
//      Format a registry REG_MULTI_SZ string as a comma separated string
//
//  Parameters:
//      pszMultiSz   - the REG_MULTI_SZ string
//      pStringValue - receives the value as a string
//
//  Returns:
//      void
//===============================================================================================//
void Registry::FormatRegistryMultiSzAsString( LPCWSTR pszMultiSz, String* pStringValue )
{
    size_t  i = 0, idxStart = 0, numChars;
    Formatter Format;

    if ( pStringValue == nullptr )
    {
        throw ParameterException( L"pStringValue", __FUNCTION__ );
    }

    if ( pszMultiSz == nullptr )
    {
        *pStringValue = L"<null>";
        return;
    }
    *pStringValue = PXS_STRING_EMPTY;

    if ( *pszMultiSz == PXS_CHAR_NULL )
    {
        return;
    }

    // Want to scan up to the NULL to get the last token, have
    // filtered case of string = "" above
    numChars = wcslen( pszMultiSz );
    numChars = PXSAddSizeT( numChars, 1 );
    for ( i = 0; i < numChars; i++ )
    {
        if ( pszMultiSz[ i ] == PXS_CHAR_NULL )
        {
            // Sometimes the string is padded with NULLS
            if ( i > idxStart )
            {
                if ( pStringValue->GetLength() )
                {
                    *pStringValue += L", ";
                }
                *pStringValue += ( pszMultiSz + idxStart );
            }

            // Position the next start after the NULL
            idxStart = ( i + 1 );
        }
    }
}
