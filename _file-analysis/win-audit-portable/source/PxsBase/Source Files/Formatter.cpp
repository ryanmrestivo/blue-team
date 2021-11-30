///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Formatter Class Implementation
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
#include "PxsBase/Header Files/Formatter.h"

// 2. C System Files
#include <float.h>
#include <stdint.h>
#include <math.h>
#include <sddl.h>
#include <time.h>
#include <wchar.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Formatter::Formatter()
          :m_String()
{
}

// Copy constructor
Formatter::Formatter( const Formatter& oFormat )
          :m_String()
{
    *this = oFormat;
}

// Destructor - do not throw any exceptions
Formatter::~Formatter()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
Formatter& Formatter::operator= ( const Formatter& oFormat)
{
    if ( this == &oFormat ) return *this;

    m_String = oFormat.m_String;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Translate and network byte order address to an IPv4 string
//
//  Parameters:
//      address - the address
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::AddressToIPv4( DWORD address )
{
    wchar_t szBuffer[ 32 ] = { 0 };   // Enough for 255.255.255.255

    StringCchPrintf( szBuffer,
                     ARRAYSIZE( szBuffer ),
                     L"%d.%d.%d.%d",
                     LOBYTE( LOWORD( address ) ),
                     HIBYTE( LOWORD( address ) ),
                     LOBYTE( HIWORD( address ) ),
                     HIBYTE( HIWORD( address ) ) );
    m_String = szBuffer;
    return m_String;
}

//===============================================================================================//
//  Description:
//      Format an IPv6 address to a colon-hexadecimal form string.
//
//  Parameters:
//      pAddress - the 16 byte buffer holding the address
//      numBytes - the number of bytes of data, ust be 16
//
//  Remarks:
//      Simple formatting only, no :: compression unless it's unspecified or loopback. No
//      mixed form IPv6:Ipv4 either.
//
//      RtlIpv6AddressToString requires Windows Vista/20008
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::AddressToIPv6( BYTE* pAddress, size_t numBytes )
{
    size_t  i = 0;
    wchar_t szBuffer[ 64 ] = { 0 };   // Enough for hhhh:hhhh:hhhh:hhhh:hhhh:hhhh:hhhh:hhhh

    m_String.Zero();
    if ( pAddress == nullptr )
    {
        return m_String;
    }

    if ( numBytes != 16 )
    {
        throw ParameterException( L"numBytes", __FUNCTION__ );
    }
    m_String.Allocate( 64 );

    // Unspecified and Loopback
    if ( ( pAddress[  0 ] == 0 ) &&
         ( pAddress[  1 ] == 0 ) &&
         ( pAddress[  2 ] == 0 ) &&
         ( pAddress[  3 ] == 0 ) &&
         ( pAddress[  4 ] == 0 ) &&
         ( pAddress[  5 ] == 0 ) &&
         ( pAddress[  6 ] == 0 ) &&
         ( pAddress[  7 ] == 0 ) &&
         ( pAddress[  8 ] == 0 ) &&
         ( pAddress[  9 ] == 0 ) &&
         ( pAddress[ 10 ] == 0 ) &&
         ( pAddress[ 11 ] == 0 ) &&
         ( pAddress[ 12 ] == 0 ) &&
         ( pAddress[ 13 ] == 0 ) &&
         ( pAddress[ 14 ] == 0 ) &&
         ( pAddress[ 15 ] <= 2 )  )
    {
        StringCchPrintf( szBuffer,
                         ARRAYSIZE( szBuffer ), L"%d", pAddress[ 15 ] );
        m_String  = L"::";
        m_String += szBuffer;
        return m_String;
    }

    for ( i = 0; i < 16; i++ )
    {
        szBuffer[ 0 ] = PXS_CHAR_NULL;
        StringCchPrintf( szBuffer,
                         ARRAYSIZE( szBuffer ), L"%02X", pAddress[ i ] );
        m_String += szBuffer;

        if ( ( i % 2 ) && ( i != 15 ) )
        {
            m_String += L":";
        }
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert an ANSI string to a wide string
//
//  Parameters:
//      pszAnsi      - pointer to the ANSI string
//      pwzWide      - pointer to the wide buffer
//      numWideChars - size of wide buffer in characters
//
//  Returns:
//      Number of characters copied into wide buffer
//===============================================================================================//
size_t Formatter::AnsiToWide( const char* pszAnsi, wchar_t* pwzWide, size_t numWideChars )
{
    int charsCopied = 0, cchWideChar = 0;

    m_String = PXS_STRING_EMPTY;

    // Must have at least room for the terminator
    if ( pszAnsi == nullptr )
    {
        return 0;   // Nothing to do
    }

    // Buffer check
    if ( pwzWide == nullptr || numWideChars == 0 )
    {
        throw ParameterException( L"pwzWide/numWideChars", __FUNCTION__ );
    }
    wmemset( pwzWide, 0, numWideChars );

    cchWideChar = PXSCastSizeTToInt32( numWideChars );
    charsCopied = MultiByteToWideChar( CP_ACP,
                                       MB_ERR_INVALID_CHARS, pszAnsi, -1, pwzWide, cchWideChar );
    pwzWide[ numWideChars - 1 ] = PXS_CHAR_NULL;

    // The bytes copied includes the terminator so zero is error
    if ( charsCopied == 0 )
    {
        throw SystemException( GetLastError(), L"MultiByteToWideChar", __FUNCTION__ );
    }

    return PXSCastInt32ToSizeT( charsCopied );
}

//===============================================================================================//
//  Description:
//      Format a bool value to a string
//
//  Parameters:
//      value - the value
//
//  Returns:
//      Reference to formatted string, "1" or "0"
//===============================================================================================//
const String& Formatter::Bool( bool value )
{
    if ( value )
    {
        m_String = PXS_STRING_ONE;
    }
    else
    {
        m_String = PXS_STRING_ZERO;
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a BOOLEAN value to a string, BOOLEAN = BYTE
//
//  Parameters:
//      value - the value
//
//  Returns:
//      Reference to formatted string, "1" or "0"
//===============================================================================================//
const String& Formatter::Boolean( BOOLEAN value )
{
    if ( value )
    {
        m_String = PXS_STRING_ONE;
    }
    else
    {
        m_String = PXS_STRING_ZERO;
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format at size_t value to binary
//
//  Parameters:
//      value - the value to format
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::BinarySizeT( size_t value )
{
    return NumberToBinary( value, 8 * sizeof ( size_t ) );
}

//===============================================================================================//
//  Description:
//      Format an unsigned 8-bit value to binary
//
//  Parameters:
//      value - the value to format
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::BinaryUInt8( BYTE value )
{
    return NumberToBinary( value, 8 );
}

//===============================================================================================//
//  Description:
//      Format an unsigned 32-bit value to binary
//
//  Parameters:
//      value - the value to format
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::BinaryUInt32( DWORD value )
{
    return NumberToBinary( value, 32 );
}

//===============================================================================================//
//  Description:
//      Format at unsigned 64 bit value to binary
//
//  Parameters:
//      value - the value to format
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::BinaryUInt64( UINT64 value )
{
    return NumberToBinary( value, 64 );
}

//===============================================================================================//
//  Description:
//      Create a GUID
//
//  Parameters:
//      none
//
//  Returns:
//      Reference to the GUID string
//===============================================================================================//
const String& Formatter::CreateGuid()
{
    GUID    guid;
    wchar_t wzGuid[ 64 ] = { 0 };       // Enough for a GUID
    HRESULT hResult;

    memset( &guid, 0, sizeof ( guid ) );
    hResult = CoCreateGuid( &guid );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoCreateGuid", __FUNCTION__ );
    }

    if ( StringFromGUID2( guid, wzGuid, ARRAYSIZE( wzGuid ) ) == 0 )
    {
        throw ComException( hResult, L"StringFromGUID2", __FUNCTION__ );
    }
    wzGuid[ ARRAYSIZE( wzGuid ) - 1 ] = PXS_CHAR_NULL;
    m_String = wzGuid;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a date in the user's locale
//
//  Parameters:
//      year  - year
//      month - month
//      day   - the day of the month
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::DateInUserLocale( WORD year, WORD month, WORD day )
{
    wchar_t szDate[ 256 ] = { 0 };       // Big enough for any date style
    SYSTEMTIME  st;

    memset( &st, 0, sizeof ( st ) );
    st.wYear  = year;
    st.wMonth = month;
    st.wDay   = day;

    GetDateFormat( LOCALE_USER_DEFAULT, 0, &st, nullptr, szDate, ARRAYSIZE( szDate ) );
    szDate[ ARRAYSIZE( szDate ) - 1 ] = PXS_CHAR_NULL;  // Terminate
    m_String = szDate;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a double precision (double) value using the %lf specifier
//
//  Parameters:
//      value - the value
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::Double( double value )
{
    wchar_t szBuffer[ 64 ] = { 0 };      // Large enough to hold any double

    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%lf", value);
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Get a double precision value in string format
//
//  Parameters:
//      value       - the value
//      decimalPlaces - the number of decimal places
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::Double( double value, BYTE decimalPlaces )
{
    wchar_t szBuffer[ 64 ] = { 0 };      // Enough to hold any double
    wchar_t szFormat[ 64 ] = { 0 };      // Enough to hold the format specifier

    // Limit the input to typical precision of a double
    if ( decimalPlaces > 15 )
    {
        decimalPlaces = 15;
    }

    // Make the format string in the form %0.nlf where n is the number
    // of decimal places.
    StringCchCopy( szFormat, ARRAYSIZE( szFormat ), L"%0." );
    StringCchPrintf( szFormat + 3, ARRAYSIZE( szFormat ) - 3, L"%u", decimalPlaces );
    StringCchCat( szFormat, ARRAYSIZE( szFormat ), L"lf" );
    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), szFormat, value );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Express a file time as a string in ISO format YYYY-MM-DD hh:mm:ss
//
//  Parameters:
//      fileTime - a FILETIME structure in UTC
//
//  Returns:
//      Reference to formatted string or "" on anything of error
//===============================================================================================//
const String& Formatter::FileTimeToLocalTimeIso( const FILETIME& fileTime )
{
    return FileTimeToString( fileTime, true , true );
}

//===============================================================================================//
//  Description:
//      Format a file time to a string in the user's locale
//
//  Parameters:
//      fileTime - a FILETIME structure in UTC
//      wantTime - if true adds time to the formatted string
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::FileTimeToLocalTimeUser( const FILETIME& fileTime, bool wantTime )
{
    return FileTimeToString( fileTime, wantTime , false );
}

//===============================================================================================//
//  Description:
//      Format a file time to a string
//
//  Parameters:
//      fileTime - a FILETIME structure in UTC
//      wantTime - if true adds time to the formatted string
//      wantISO  - if format is ISO otherwise in user default
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::FileTimeToString( const FILETIME& fileTime, bool wantTime, bool wantISO )
{
    FILETIME   ft;
    SYSTEMTIME st;

    // Filter zero
    m_String = PXS_STRING_EMPTY;
    if ( ( fileTime.dwHighDateTime == 0 ) &&
         ( fileTime.dwLowDateTime  == 0 )  )
    {
        return m_String;
    }

    // File times are in UTC
    memset( &ft, 0, sizeof ( ft ) );
    if ( FileTimeToLocalFileTime( &fileTime, &ft ) == 0 )
    {
        throw SystemException( GetLastError(), L"FileTimeToLocalFileTime", __FUNCTION__ );
    }

    // Convert to system time
    memset( &st, 0, sizeof ( st ) );
    if ( FileTimeToSystemTime( &ft, &st ) == 0 )
    {
        throw SystemException( GetLastError(), L"FileTimeToSystemTime", __FUNCTION__ );
    }

    return SystemTimeToString( st, wantTime , wantISO );
}

//===============================================================================================//
//  Description:
//      Format a single precision (float) value with the %f specifier
//
//  Parameters:
//      value - single precision value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::Float( float value )
{
    wchar_t szBuffer[ 32 ] = { 0 };      // Enough to hold any float

    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%f", value );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a single precision value in string format
//
//  Parameters:
//      value         - the value
//      decimalPlaces - the number of decimal places
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::Float( float value, BYTE decimalPlaces )
{
    wchar_t szFormat[  8 ] = { 0 };      // Enough to hold the format specifier
    wchar_t szBuffer[ 32 ] = { 0 };      // Enough to hold any float

    // Limit the input to typical precision of a float
    if ( decimalPlaces > 7 )
    {
        decimalPlaces = 7;
    }

    // Make the format string, want it in the form %0.nf where n is the
    // number of decimal places.
    StringCchCopy( szFormat, ARRAYSIZE( szFormat ), L"%0." );
    StringCchPrintf( szFormat + 3, ARRAYSIZE( szFormat ) - 3, L"%u", decimalPlaces );
    StringCchCat( szFormat, ARRAYSIZE( szFormat ), L"f" );
    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), szFormat, value );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Get the string from a module corresponding to its message id. The
//      message is stored in the modules Message Table resource.
//
//  Parameters:
//      ModulePath - full path to the module
//      messageID  - the message id
//      Inserts    - array of insertion strings, can be empty
//
//  Remarks:
//      Using FormatMessage with external strings that have incorrect inserts
//      (e.g. eventlog) can result in an access violation so will use
//      FORMAT_MESSAGE_IGNORE_INSERTS then effect any inserts.
//
//  Returns:
//      Reference to formatted string or "" on any king of error
//===============================================================================================//
const String& Formatter::GetModuleString( const String& ModulePath,
                                          DWORD messageID, const StringArray& Inserts )
{
    size_t MAX_INSERTS = 99;
    size_t i = 0, maxInserts = 0;
    wchar_t  szParameter[ 8 ];        // e.g. %1, %2 ...
    wchar_t  szMessage[ 512 ];        // Resource message
    String Message, Insert;
    HINSTANCE hLibrary = nullptr;

    m_String = PXS_STRING_EMPTY;
    if ( ModulePath.IsEmpty() ||
        ( ModulePath.CharAt( 0 ) == PXS_PATH_SEPARATOR ) )  // Disallow UNC
    {
        return m_String;
    }

    hLibrary = LoadLibraryEx( ModulePath.c_str(), nullptr, LOAD_LIBRARY_AS_DATAFILE );
    if ( hLibrary == nullptr )
    {
        return m_String;
    }

    memset( szMessage, 0, sizeof ( szMessage ) );
    FormatMessage( FORMAT_MESSAGE_FROM_HMODULE | FORMAT_MESSAGE_IGNORE_INSERTS,
                   hLibrary,
                   messageID,
                   MAKELANGID( LANG_NEUTRAL, SUBLANG_DEFAULT ),
                   szMessage,
                   ARRAYSIZE( szMessage ),
                   nullptr );

    // Unload library first
    FreeLibrary( hLibrary );
    szMessage[ ARRAYSIZE( szMessage ) - 1 ] = PXS_CHAR_NULL;
    m_String = szMessage;     // Can throw

    // Replace parameters(%1..%99) with insertion strings
    // Limit the number of inserts
    maxInserts = PXSMinSizeT( Inserts.GetSize(), MAX_INSERTS );
    for ( i = 0; i < maxInserts; i++ )
    {
        // Make the substitution parameter, e.g. %2
        memset( szParameter, 0,  sizeof ( szParameter ) );
        szParameter[ 0 ] = '%';
        StringCchPrintf( szParameter + 1, ARRAYSIZE( szParameter ), L"%lu", i + 1 );
        Insert = Inserts.Get( i );
        Message.ReplaceI( szParameter, Insert.c_str() );
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert a GUID to a string, includes the opening and closing braces
//
//  Parameters:
//      guid - the GUID to convert
//
//  Returns:
//      Constant reference to formatted string
//===============================================================================================//
const String& Formatter::GuidToString( const GUID& guid )
{
    wchar_t szGuid[ 64 ] = { 0 };     // Enough to hold 38 char GUID

    // Fixed format, no braces
    StringCchPrintf( szGuid,
                     ARRAYSIZE( szGuid ),
                     L"{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
                     guid.Data1,
                     guid.Data2,
                     guid.Data3,
                     guid.Data4[ 0 ],
                     guid.Data4[ 1 ],
                     guid.Data4[ 2 ],
                     guid.Data4[ 3 ],
                     guid.Data4[ 4 ],
                     guid.Data4[ 5 ],
                     guid.Data4[ 6 ],
                     guid.Data4[ 7 ] );
    m_String  = szGuid;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a hex string as a number
//
//  Parameters:
//      HexString - string object of hex number
//
//  Remarks:
//      Maximum is 8 hex chars i.e. 0xFFFFFFFF, the leading 0x is optional e.g.
//      A    = 10;
//      0xA  = 10;
//
//  Returns:
//      DWORD of string converted to a number, zero on error
//===============================================================================================//
DWORD Formatter::HexStringToNumber( const String& HexString )
{
    bool   success = true;  // Assume success
    size_t i  = 0, maxChars = 0;
    wchar_t  ch = 0;
    DWORD  value = 0, digit = 0, multiplier = 1;
    String Data;

    m_String = PXS_STRING_EMPTY;

    // Make a local copy and clean up
    Data = HexString;
    Data.Trim();
    Data.ToUpperCase();

    // Reverse the string
    Data.Reverse();
    maxChars = Data.GetLength();
    if ( maxChars > 8 )
    {
        maxChars = 8;
    }

    while ( ( i < maxChars ) && ( success == true )  )
    {
        digit = 0;
        ch    = Data.CharAt( i );
        if ( ( ch >= '0' ) && ( ch <= '9' ) )
        {
            digit = PXSCastInt32ToUInt32( ch - '0' );
        }
        else if ( ( ch >= 'A' ) && ( ch <= 'F' ) )
        {
            digit = PXSCastInt32ToUInt32( ch - 55 );
        }
        else
        {
            success = false;    // Non-hexit
            value   = 0;
        }
        value += ( digit * multiplier );

        // Next pass
        if ( i < 8 )
        {
            multiplier *= 16;
        }
        i++;
    }

    return value;
}

//===============================================================================================//
//  Description:
//      Format a signed integer using the %ld specifier
//
//  Parameters:
//      value - the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::Int32( int value )
{
    wchar_t szBuffer[ 16 ] = { 0 };      // Enough to hold any signed integer

    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%ld", value );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a BOOL value to a string, BOOL = int
//
//  Parameters:
//      value - the value
//
//  Returns:
//      Reference to formatted string, "1" or "0"
//===============================================================================================//
const String& Formatter::Int32Bool( BOOL value )
{
    if ( value )
    {
        m_String = PXS_STRING_ONE;
    }
    else
    {
        m_String = PXS_STRING_ZERO;
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a signed 32-bit integer using the 0x%08X specifier
//
//  Parameters:
//      value - integer to format
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::Int32Hex( int value )
{
    wchar_t szBuffer[ 16 ] = { 0 };   // Large enough to hold any hex string

    StringCchPrintf( szBuffer, ARRAYSIZE(szBuffer), L"0x%08X", value );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a value to "Yes" or "No" in English. Used for non-translatable
//      strings such as in the logger so want the result in English.
//
//  Parameters:
//      value - the value, zero = No otherwise Yes
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::Int32YesNo( int value )
{
    if ( value )
    {
        m_String = PXS_STRING_YES;
    }
    else
    {
        m_String = PXS_STRING_NO;
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a 64-bit signed integer with the %ll specifier
//
//  Parameters:
//      value - the 64-bit value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::Int64( __int64 value )
{
    wchar_t szBuffer[ 32 ] = { 0 };  // Enough to hold any 64-bit signed integer

    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%lld", value );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Test if a string forms a valid ISO time
//
//  Parameters:
//      Timestamp   - pointer to the time-stamp string
//      pSystemTime - optional, receives the SYSTEMTIME if the string is valid
//
//  Remarks:
//      YYYY-MM-DD HH:MM:SS, must not have trailing or leading spaces
//
//  Returns:
//      true if valid, else false
//===============================================================================================//
bool Formatter::IsValidIsoTimestamp( const String& Timestamp, SYSTEMTIME* pSystemTime )
{
    int     year = 0, month = 0, day = 0, hour = 0, minute = 0, second = 0;
    LPCWSTR pszTimestamp;

    m_String = PXS_STRING_EMPTY;     // Not used in this method

    // YYYY-MM-DD hh:mm:ss
    // 0123456789012345678 = 19 characters
    if ( Timestamp.GetLength() != 19 )
    {
        return false;
    }
    pszTimestamp = Timestamp.c_str();

    year  = ( 1000 * ( pszTimestamp[ 0 ] - '0' ) ) +
            (  100 * ( pszTimestamp[ 1 ] - '0' ) ) +
            (   10 * ( pszTimestamp[ 2 ] - '0' ) ) +
            (    1 * ( pszTimestamp[ 3 ] - '0' ) );

    month = (   10 * ( pszTimestamp[ 5 ] - '0' ) ) +
            (    1 * ( pszTimestamp[ 6 ] - '0' ) );

    day   = (   10 * ( pszTimestamp[ 8 ] - '0' ) ) +
            (    1 * ( pszTimestamp[ 9 ] - '0' ) );

    hour   = (   10 * ( pszTimestamp[ 11 ] - '0' ) ) +
             (    1 * ( pszTimestamp[ 12 ] - '0' ) );

    minute = (   10 * ( pszTimestamp[ 14 ] - '0' ) ) +
             (    1 * ( pszTimestamp[ 15 ] - '0' ) );

    second = (   10 * ( pszTimestamp[ 17 ] - '0' ) ) +
             (    1 * ( pszTimestamp[ 18 ] - '0' ) );

    if ( IsValidYearMonthDay( year, month, day ) == false )
    {
        return false;
    }

    if ( ( hour < 0   ) || ( hour > 23   ) ||
         ( minute < 0 ) || ( minute > 59 ) ||
         ( second < 0 ) || ( second > 59 )  )
    {
        return false;
    }

    // If the string is valid, fill the input SYSTEMTIME
    if ( pSystemTime )
    {
        pSystemTime->wYear         = PXSCastInt32ToUInt16( year );
        pSystemTime->wMonth        = PXSCastInt32ToUInt16( month );
        pSystemTime->wDayOfWeek    = 0;                     // Not required
        pSystemTime->wDay          = PXSCastInt32ToUInt16( day );
        pSystemTime->wHour         = PXSCastInt32ToUInt16( hour );
        pSystemTime->wMinute       = PXSCastInt32ToUInt16( minute );
        pSystemTime->wSecond       = PXSCastInt32ToUInt16( second );
        pSystemTime->wMilliseconds = 0;                     // Not used
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Determine if a string is a MAC Address
//
//  Parameters:
//      MacAddress  - the string to test, neither leading nor trailing spaces
//                    are not allowed
//
//  Remarks:
//      MAC Address: 00:00:00:00:00:00
//                   01234567891111111
//                             0123456
//  Returns:
//      true if the string is in MAC Address format, else false
//===============================================================================================//
bool Formatter::IsValidMacAddress( const String& MacAddress )
{
    m_String = PXS_STRING_EMPTY;
    if ( MacAddress.GetLength() != 17 )
    {
        return false;
    }

    if ( MacAddress.CompareI( L"00:00:00:00:00:00" ) == 0 )
    {
        return false;
    }

    if ( MacAddress.CompareI( L"FF:FF:FF:FF:FF:FF" ) == 0 )
    {
        return false;
    }

    if ( ( MacAddress.CharAt(  2 ) != ':' ) ||
         ( MacAddress.CharAt(  5 ) != ':' ) ||
         ( MacAddress.CharAt(  8 ) != ':' ) ||
         ( MacAddress.CharAt( 11 ) != ':' ) ||
         ( MacAddress.CharAt( 14 ) != ':' )  )
    {
        return false;
    }

    if ( PXSIsHexitW( MacAddress.CharAt(  0 ) ) &&
         PXSIsHexitW( MacAddress.CharAt(  1 ) ) &&
         PXSIsHexitW( MacAddress.CharAt(  3 ) ) &&
         PXSIsHexitW( MacAddress.CharAt(  4 ) ) &&
         PXSIsHexitW( MacAddress.CharAt(  6 ) ) &&
         PXSIsHexitW( MacAddress.CharAt(  7 ) ) &&
         PXSIsHexitW( MacAddress.CharAt(  9 ) ) &&
         PXSIsHexitW( MacAddress.CharAt( 10 ) ) &&
         PXSIsHexitW( MacAddress.CharAt( 12 ) ) &&
         PXSIsHexitW( MacAddress.CharAt( 13 ) ) &&
         PXSIsHexitW( MacAddress.CharAt( 15 ) ) &&
         PXSIsHexitW( MacAddress.CharAt( 16 ) )  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if a string is a GUID
//
//  Parameters:
//      StringGuid  - the string to test, must not have leading/trailing
//                    white spaces
//  Remarks:
//      Format: {00000000-0000-0000-0000-000000000000}
//              01234567891111111111222222222233333333
//                        0123456789012345678901234567
//
//  Returns:
//      true if valid otherwise else false
//===============================================================================================//
bool Formatter::IsValidStringGuid( const String& StringGuid )
{
    bool    success = true;
    DWORD   i  = 0;
    wchar_t ch = 0;
    LPCWSTR STR_NULL_GUID = L"{00000000-0000-0000-0000-000000000000}";

    m_String = PXS_STRING_EMPTY;
    if ( StringGuid.GetLength() != 38 )
    {
        return false;
    }

    if ( StringGuid.CompareI( STR_NULL_GUID ) == 0 )
    {
        success = false;
    }

    if ( StringGuid.CompareI( L"{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}" ) == 0 )
    {
        success = false;
    }

    while ( ( i < 38 ) && success )
    {
        // Where there is a zero in the test string, the GUID must have a
        // hexit otherwise the formatting characters '{' '-' '}'
        ch = StringGuid.CharAt( i );
        if ( STR_NULL_GUID[ i ] == '0' )
        {
            success = PXSIsHexitW( ch );
        }
        else if ( STR_NULL_GUID[ i ] != ch )
        {
            success = false;
        }
        i++;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Test the data in a SYSTEMTIME structure
//
//  Parameters:
//      systemTime - SYSYEMTIME structure
//
//  Remarks:
//      Does not check the consistency of the Day of Week member
//
//  Returns:
//      true if if valid, else false
//===============================================================================================//
bool Formatter::IsValidSystemTime( const SYSTEMTIME& systemTime )
{
    // Date members
    if ( false == IsValidYearMonthDay( systemTime.wYear, systemTime.wMonth, systemTime.wDay  ) )
    {
        return false;
    }

    if ( systemTime.wDayOfWeek > 6 )
    {
        return false;
    }

    // Time members
    if ( ( systemTime.wHour   > 23 ) ||
         ( systemTime.wMinute > 59 ) ||
         ( systemTime.wSecond > 59 ) ||
         ( systemTime.wMilliseconds > 999 ) )
    {
        return false;
    }

    return true;
}


//===============================================================================================//
//  Description:
//      Test if the specified year, month and day combination are valid
//
//  Parameters:
//      year  - the year, any year is valid but want to check for leap years
//      month - the month
//      day   - the day
//
//
//  Returns:
//      true if the date/time is valid, else false
//===============================================================================================//
bool Formatter::IsValidYearMonthDay( int year, int month, int day )
{
    //  One-based:     -  J   F   M   A   M   J   J   A   S   O   N   D
    int Days[ 13 ] = { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };

    if ( ( month < 1 ) || ( month > 12 ) )
    {
       return false;
    }

    // Non-Leap February
    if ( ( month == 2 ) && ( year % 4 ) && ( day > 28 ) )
    {
        return false;
    }

    // Day in month
    if ( ( day < 1 ) || ( day > Days[ month ] ) )  // Bounds checked above
    {
       return false;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Convert an ISO date to a string formatted in the user's locale
//
//  Parameters:
//      Timestamp - string in ISO format YYYY-MM-DD hh:mm:ss
//      wantTime  - flag to indicate to include the time part
//
//  Remarks:
//      If Null or "" input will no raise an error
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::IsoTimestampToLocale( const String& Timestamp, bool wantTime )
{
    SYSTEMTIME st;

    m_String = PXS_STRING_EMPTY;
    if ( Timestamp.IsEmpty() )
    {
        return m_String;
    }

    // Convert to a system time
    memset( &st, 0, sizeof ( &st ) );
    if ( IsValidIsoTimestamp( Timestamp, &st ) == false )
    {
        throw ParameterException( L"pszTimestamp", __FUNCTION__ );
    }

    return SystemTimeToString( st, wantTime, true );   // true = User locale
}

//===============================================================================================//
//  Description:
//      Convert a language id to a name
//
//  Parameters:
//      languageID - the identifier of the language, e.g. 1033 = English
//
// Remarks
//      There does not seem to a system API for this so will us a table,
//      see MSDN"Locale Identifier Constants and Strings".
//
//      Upper and lower Sorbin have the same primary language ID, everybody
//      knows that!
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::LanguageIdToName( WORD languageID )
{
    BYTE   primaryID = 0;
    size_t i = 0;

    // Hold the ID to name mapping
    struct _ID_NAME
    {
        BYTE    primaryID;
        LPCWSTR pszName;
    } IDs [] = { { LANG_AFRIKAANS           , L"Afrikaans"        } ,
                 { LANG_ALBANIAN            , L"Albanian"         } ,
                 { LANG_ALSATIAN            , L"Alsatian"         } ,
                 { LANG_AMHARIC             , L"Amharic"          } ,
                 { LANG_ARABIC              , L"Arabic"           } ,
                 { LANG_ARMENIAN            , L"Armenian"         } ,
                 { LANG_ASSAMESE            , L"Assamese"         } ,
                 { LANG_AZERI               , L"Azeri"            } ,
                 { LANG_BASQUE              , L"Basque"           } ,
                 { LANG_BELARUSIAN          , L"Belarusian"       } ,
                 { LANG_BENGALI             , L"Bengali"          } ,
                 { LANG_BRETON              , L"Breton"           } ,
                 { LANG_BULGARIAN           , L"Bulgarian"        } ,
                 { LANG_CATALAN             , L"Catalan"          } ,
                 { LANG_CHINESE             , L"Chinese"          } ,
              // { LANG_CHINESE_SIMPLIFIED  , L"Chinese"          } ,
              // { LANG_CHINESE_TRADITIONAL , L"Chinese"          } ,
                 { LANG_CORSICAN            , L"Corsican"         } ,
                 { LANG_CZECH               , L"Czech"            } ,
                 { LANG_DANISH              , L"Danish"           } ,
                 { LANG_DARI                , L"Dari"             } ,
                 { LANG_DIVEHI              , L"Divehi"           } ,
                 { LANG_DUTCH               , L"Dutch"            } ,
                 { LANG_ENGLISH             , L"English"          } ,
                 { LANG_ESTONIAN            , L"Estonian"         } ,
                 { LANG_FAEROESE            , L"Faroese"          } ,
                 { LANG_FARSI               , L"Persian"          } ,
                 { LANG_FILIPINO            , L"Filipino"         } ,
                 { LANG_FINNISH             , L"Finnish"          } ,
                 { LANG_FRENCH              , L"French"           } ,
                 { LANG_FRISIAN             , L"Frisian"          } ,
                 { LANG_GALICIAN            , L"Galician"         } ,
                 { LANG_GEORGIAN            , L"Georgian"         } ,
                 { LANG_GERMAN              , L"German"           } ,
                 { LANG_GREEK               , L"Greek"            } ,
                 { LANG_GREENLANDIC         , L"Greenlandic"      } ,
                 { LANG_GUJARATI            , L"Gujarati"         } ,
                 { LANG_HAUSA               , L"Hausa"            } ,
                 { LANG_HEBREW              , L"Hebrew"           } ,
                 { LANG_HINDI               , L"Hindi"            } ,
                 { LANG_HUNGARIAN           , L"Hungarian"        } ,
                 { LANG_ICELANDIC           , L"Icelandic"        } ,
                 { LANG_IGBO                , L"Igbo"             } ,
                 { LANG_INDONESIAN          , L"Indonesian"       } ,
                 { LANG_INUKTITUT           , L"Inuktitut"        } ,
                 { LANG_INVARIANT           , L"Invariant"        } ,
                 { LANG_IRISH               , L"Irish"            } ,
                 { LANG_ITALIAN             , L"Italian"          } ,
                 { LANG_JAPANESE            , L"Japanese"         } ,
                 { LANG_KANNADA             , L"Kannada"          } ,
                 { LANG_KASHMIRI            , L"Kashmiri"         } ,
                 { LANG_KAZAK               , L"Kazakh"           } ,
                 { LANG_KHMER               , L"Khmer"            } ,
                 { LANG_KICHE               , L"K'iche"           } ,
                 { LANG_KINYARWANDA         , L"Kinyarwanda"      } ,
                 { LANG_KONKANI             , L"Konkani"          } ,
                 { LANG_KOREAN              , L"Korean"           } ,
                 { LANG_KYRGYZ              , L"Kyrgyz"           } ,
                 { LANG_LAO                 , L"Lao"              } ,
                 { LANG_LATVIAN             , L"Latvian"          } ,
                 { LANG_LITHUANIAN          , L"Lithuanian"       } ,
                 { LANG_LOWER_SORBIAN       , L"Sorbian"          } ,
                 { LANG_UPPER_SORBIAN       , L"Sorbian"          } ,
                 { LANG_LUXEMBOURGISH       , L"Luxembourgish"    } ,
                 { LANG_MACEDONIAN          , L"Macedonian"       } ,
                 { LANG_MALAY               , L"Malay"            } ,
                 { LANG_MALAYALAM           , L"Malayalam"        } ,
                 { LANG_MALTESE             , L"Maltese"          } ,
                 { LANG_MANIPURI            , L"Manipuri"         } ,
                 { LANG_MAORI               , L"Maori"            } ,
                 { LANG_MAPUDUNGUN          , L"Mapudungun"       } ,
                 { LANG_MARATHI             , L"Marathi"          } ,
                 { LANG_MOHAWK              , L"Mohawk"           } ,
                 { LANG_MONGOLIAN           , L"Mongolian"        } ,
                 { LANG_NEPALI              , L"Nepali"           } ,
                 { LANG_NEUTRAL             , L"Neutral"          } ,
                 { LANG_NORWEGIAN           , L"Norwegian"        } ,
                 { LANG_OCCITAN             , L"Occitan"          } ,
                 { LANG_ORIYA               , L"Oriya"            } ,
                 { LANG_PASHTO              , L"Pashto"           } ,
                 { LANG_POLISH              , L"Polish"           } ,
                 { LANG_PORTUGUESE          , L"Portuguese"       } ,
                 { LANG_PUNJABI             , L"Punjabi"          } ,
                 { LANG_QUECHUA             , L"Quechua"          } ,
                 { LANG_ROMANIAN            , L"Romanian"         } ,
                 { LANG_ROMANSH             , L"Romansh"          } ,
                 { LANG_RUSSIAN             , L"Russian"          } ,
                 { LANG_SAMI                , L"Sami"             } ,
                 { LANG_SANSKRIT            , L"Sanskrit"         } ,
                 { LANG_SINDHI              , L"Sindhi"           } ,
                 { LANG_SINHALESE           , L"Sinhala"          } ,
                 { LANG_SLOVAK              , L"Slovak"           } ,
                 { LANG_SLOVENIAN           , L"Slovenian"        } ,
                 { LANG_SOTHO               , L"Sesotho"          } ,
                 { LANG_SPANISH             , L"Spanish"          } ,
                 { LANG_SWAHILI             , L"Swahili"          } ,
                 { LANG_SWEDISH             , L"Swedish"          } ,
                 { LANG_SYRIAC              , L"Syriac"           } ,
                 { LANG_TAJIK               , L"Tajik"            } ,
                 { LANG_TAMAZIGHT           , L"Tamazight"        } ,
                 { LANG_TAMIL               , L"Tamil"            } ,
                 { LANG_TATAR               , L"Tatar"            } ,
                 { LANG_TELUGU              , L"Telugu"           } ,
                 { LANG_THAI                , L"Thai"             } ,
                 { LANG_TIBETAN             , L"Tibetan"          } ,
                 { LANG_TIGRIGNA            , L"Tigrigna"         } ,
                 { LANG_TSWANA              , L"Setswana/Tswana"  } ,
                 { LANG_TURKISH             , L"Turkish"          } ,
                 { LANG_TURKMEN             , L"Turkmen"          } ,
                 { LANG_UIGHUR              , L"Uighur"           } ,
                 { LANG_UKRAINIAN           , L"Ukrainian"        } ,
                 { LANG_URDU                , L"Urdu"             } ,
                 { LANG_UZBEK               , L"Uzbek"            } ,
                 { LANG_VIETNAMESE          , L"Vietnamese"       } ,
                 { LANG_WELSH               , L"Welsh"            } ,
                 { LANG_WOLOF               , L"Wolof"            } ,
                 { LANG_XHOSA               , L"Xhosa/isiXhosa"   } ,
                 { LANG_YAKUT               , L"Yakut"            } ,
                 { LANG_YI                  , L"Yi"               } ,
                 { LANG_YORUBA              , L"Yoruba"           } ,
                 { LANG_ZULU                , L"Zulu/isiZulu"     } };

    // Find it
    primaryID = LOBYTE( languageID );
    for ( i = 0; i < ARRAYSIZE( IDs ); i++ )
    {
        if ( primaryID == IDs[ i ].primaryID )
        {
            m_String = IDs[ i ].pszName;
            break;
        }
    }

    // Special cases
    if ( primaryID == LANG_BOSNIAN )  // = LANG_SERBIAN = LANG_CROATIAN = 0x1a
    {
        if ( IsBosnian( languageID ) )
        {
            m_String = L"Bosnian";
        }
        else if ( IsCroatian( languageID ) )
        {
            m_String = L"Croatian";
        }
        else if ( IsSerbian( languageID ) )
        {
            m_String = L"Serbian";
        }
    }

    // If nothing, set the class scope string as Hex
    if ( m_String.IsEmpty() )
    {
        UInt16Hex( languageID );
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Express the local time in ISO time stamp format
//      i.e. YYYY-MM-DD HH:MM:SS format
//
//  Parameters:
//      None
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::LocalTimeInIsoFormat()
{
    SYSTEMTIME st;

    m_String = PXS_STRING_EMPTY;
    memset( &st, 0, sizeof ( st ) );
    GetLocalTime( &st );

    return SystemTimeToString( st, true , true );
}

//===============================================================================================//
//  Description:
//      Express the local time in the format of the user's locale
//
//  Parameters:
//      None
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::LocalTimeInUserLocale()
{
    SYSTEMTIME st;

    m_String = PXS_STRING_EMPTY;
    memset( &st, 0, sizeof ( st ) );
    GetLocalTime( &st );

    return SystemTimeToString( st, true, false );  // false = user locale
}

//===============================================================================================//
//  Description:
//      Format the current system time to hh:mm:ss.ttt
//
//  Parameters:
//      none
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::NowTohhmmssttt()
{
    wchar_t  szBuffer[ 32 ];   // hh:mm:ss
    SYSTEMTIME st;

    memset( &st, 0, sizeof( st ) );
    GetSystemTime( &st );
    StringCchPrintf( szBuffer,
                     ARRAYSIZE( szBuffer ),
                     L"%02u:%02u:%02u.%03u",
                     st.wHour,
                     st.wMinute,
                     st.wSecond,
                     st.wMilliseconds );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format the current system time to YYYYMMDDhhmmsstttt
//
//  Parameters:
//      none
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::NowToYYYYMMDDhhmmsstttt()
{
    wchar_t  szBuffer[ 64 ];   // YYYYMMDDhhmmss
    SYSTEMTIME st;

    memset( &st, 0, sizeof( st ) );
    GetSystemTime( &st );
    StringCchPrintf( szBuffer,
                     ARRAYSIZE( szBuffer ),
                     L"%04u%02u%02u%02u%02u%02u%04u",
                     st.wYear,
                     st.wMonth,
                     st.wDay,
                     st.wHour,
                     st.wMinute,
                     st.wSecond,
                     st.wMilliseconds);
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Express a 64-bit date in ISO format;
//
//  Parameters:
//      dblDate - the OLE date
//
//  Remarks:
//      The DATE type is an 8-byte floating-point number. Days are
//      represented by whole number increments starting with
//      30 December 1899, midnight as time zero. Hour values are
//      expressed as the absolute value of the fractional part of
//      the number.
//      e.g. 38897.557546  =  2006-06-29 14:22:52
//      i.e. 38897 days after 1899-12-30
//           0.557546 * 24 * 3600 seconds after midnight
//
//      VariantTimeToSystemTime is available from 95 OSR2
//      NT3.51 and Win95 OSR1 only have VariantTimeToDosTime
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::OleTimeInIsoFormat( const DATE& oleDate )
{
    String     ErrorMessage;
    Formatter  Format;
    SYSTEMTIME st;

    m_String = PXS_STRING_EMPTY;
    memset( &st, 0, sizeof ( st ) );
    if ( VariantTimeToSystemTime( oleDate, &st ) == 0 )
    {
        ErrorMessage  = L"oleDate = ";
        ErrorMessage += Format.Double( oleDate );
        throw SystemException( ERROR_INVALID_TIME,
                               ErrorMessage.c_str(), "VariantTimeToSystemTime" );
    }

    return SystemTimeToIso( st );
}

//===============================================================================================//
//  Description:
//      Format a pointer address using the 0x%p specifier
//
//  Parameters:
//      pValue - pointer
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::Pointer( const void* pValue )
{
    wchar_t szBuffer[ 64 ] = { 0 };   // Enough to hold any pointer address

    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"0x%p", pValue );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert seconds to a string of DD Days HH Hours MM minutes
//
//  Parameters:
//      seconds - number of seconds to convert
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::SecondsToDDHHMM( time_t seconds )
{
    time_t days = 0, hours = 0, minutes = 0;
    wchar_t  szBuffer[ 64 ] = { 0 };   // Enough for dd hh ss
    String ResourceString;

    m_String.Allocate( 32 );
    m_String = PXS_STRING_EMPTY;

    if ( seconds <= 0 )
    {
        return m_String;
    }

    // Calculate days, hours and seconds
    days    = seconds / ( 24 * 3600);
    hours   = ( seconds % ( 24 * 3600) ) / 3600;
    minutes = ( seconds % 3600 ) / 60;

    // Days
    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%lld", days );
    m_String += szBuffer;
    m_String += PXS_CHAR_SPACE;
    if ( days == 1 )
    {
        PXSGetResourceString( PXS_IDS_143_DAY, &ResourceString );
    }
    else
    {
        PXSGetResourceString( PXS_IDS_144_DAYS, &ResourceString );
    }
    m_String += ResourceString;
    m_String += PXS_CHAR_SPACE;

    // Hours
    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%lld", hours );
    m_String += szBuffer;
    m_String += PXS_CHAR_SPACE;
    if ( hours == 1 )
    {
        PXSGetResourceString( PXS_IDS_141_HOUR, &ResourceString );
    }
    else
    {
        PXSGetResourceString( PXS_IDS_142_HOURS, &ResourceString );
    }
    m_String += ResourceString;
    m_String += PXS_CHAR_SPACE;

    // Minutes
    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%lld", minutes );
    m_String += szBuffer;
    m_String += PXS_CHAR_SPACE;
    if ( hours == 1 )
    {
        PXSGetResourceString( PXS_IDS_139_MINUTE, &ResourceString );
    }
    else
    {
        PXSGetResourceString( PXS_IDS_140_MINUTES, &ResourceString );
    }
    m_String += ResourceString;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Translate a SID to a string
//
//  Parameters:
//      pSid - pointer to the SID
//
//  Remarks:
//      On NT5+ can use ConvertSidToStringSid
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::SidToString( PSID pSid )
{
    wchar_t* pszStringSid = nullptr;

    m_String = PXS_STRING_EMPTY;
    if ( pSid == nullptr )
    {
        return m_String;
    }

    // Must free the string
    try
    {
        if ( ConvertSidToStringSid( pSid, &pszStringSid ) )
        {
            m_String = pszStringSid;
            LocalFree( pszStringSid );
            pszStringSid = nullptr;
        }
    }
    catch ( const Exception& )
    {
        if ( pszStringSid )
        {
            LocalFree( pszStringSid );
        }
        throw;
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Get an size_t integer in string format
//
//  Parameters:
//      value - the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::SizeT( size_t value )
{
    return UInt64( value );
}

//===============================================================================================//
//  Description:
//      Get an size_t integer in string format
//
//  Parameters:
//      value  - the value
//      pArray - receives the formatted number
//
//  Returns:
//      void
//===============================================================================================//
void Formatter::SizeT( size_t value, CharArray* pArray )
{
    char    szValue[ 32 ] = { 0 };  // Enough for "18446744073709551615"
    size_t  length = 0;
    UINT64  un64   = static_cast< UINT64 >( value );
    HRESULT hResult;

    if ( pArray == nullptr )
    {
        throw ParameterException( L"pArray", __FUNCTION__ );
    }
    pArray->Zero();

    hResult = StringCchPrintfA( szValue, sizeof ( szValue ), "%llu", un64 );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"StringCchPrintfA", __FUNCTION__ );
    }
    szValue[ sizeof ( szValue ) - 1 ] = 0x00;
    StringCchLengthA( szValue, STRSAFE_MAX_CCH, &length );
    pArray->Append( szValue, length );
}

//===============================================================================================//
//  Description:
//      Format a byte storage value, e.g. GB or MB in mebibyte style
//
//  Parameters:
//      storageBytes - the value in bytes
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StorageBytes( UINT64 storageBytes )
{
    double  value = 0.0, exponent = 0.0;
    String  LocaleUnitName;
    wchar_t szBuffer[ 64 ] = { 0 };    // Enough for any 64-bit number

    // Handle small values
    if ( storageBytes < 0x00000400 )   // 1KB
    {
        PXSGetResourceString( PXS_IDS_134_BYTES, &LocaleUnitName );
        StringCchPrintf( szBuffer,
			             ARRAYSIZE( szBuffer ),
			             L"%lu", static_cast<WORD>( 0xFFFF & storageBytes) );
        m_String  = szBuffer;
        m_String += PXS_CHAR_SPACE;
        m_String += LocaleUnitName;
        return m_String;
    }

    m_String = PXS_STRING_EMPTY;
    if ( storageBytes > 0x10000000000 )     // 1TB
    {
        exponent = 0x10000000000;
        PXSGetResourceString( PXS_IDS_138_TB, &LocaleUnitName );
    }
    else if ( storageBytes > 0x40000000 )   // 1GB
    {
        exponent = 0x40000000;
        PXSGetResourceString( PXS_IDS_137_GB, &LocaleUnitName );
    }
    else if ( storageBytes > 0x00100000 )   // 1MB
    {
        // Traditionally floppies 1MB = 1024*1000, but will use
        // 1024*1024 is consistent with Windows Explorer
        exponent = 0x00100000;
        PXSGetResourceString( PXS_IDS_136_MB, &LocaleUnitName );
    }
    else
    {
        exponent = 0x00000400;
        PXSGetResourceString( PXS_IDS_135_KB, &LocaleUnitName );
    }

    // Format to 3 significant figures
    value = static_cast<double>( storageBytes ) / exponent;
    if ( value >= 100.0 )
    {
        // No decimal places
        StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%0.01f", value );
    }
    else if ( value >= 10.0 )
    {
        // 1 decimal place
        StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%0.1lf", value );
    }
    else
    {
        // 2 decimal places
        StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%0.2lf", value );
    }
    m_String = szBuffer;
    m_String += LocaleUnitName;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a string with the specified insert
//
//  Parameters:
//      pszString - the string to format
//      Insert1   - insertion string for placeholder %%1
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::String1( LPCWSTR pszString, const String& Insert1 )
{
    String Insert2, Insert3;
    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified inserts
//
//  Parameters:
//      pszString - the string to format
//      Insert1   - insertion string for placeholder %%1
//      Insert2   - insertion string for placeholder %%2
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::String2( LPCWSTR pszString, const String& Insert1, const String& Insert2 )
{
    String Insert3;
    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified inserts
//
//  Parameters:
//      pszString - the string to format
//      Insert1   - insertion string for placeholder %%1
//      Insert2   - insertion string for placeholder %%2
//      Insert3   - insertion string for placeholder %%3
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::String3( LPCWSTR pszString,
                                  const String& Insert1,
                                  const String& Insert2, const String& Insert3)
{
    m_String = pszString;

    // Replace first parameter
    if ( Insert1.c_str() )
    {
        m_String.ReplaceI( L"%%1", Insert1.c_str() );
    }
    else
    {
        // No parameter specified, in case use <null> like printf
        m_String.ReplaceI( L"%%1", L"<null>" );
    }

    // Replace the second parameter
    if ( Insert2.c_str() )
    {
        m_String.ReplaceI( L"%%2", Insert2.c_str() );
    }
    else
    {
        m_String.ReplaceI( L"%%2", L"<null>" );
    }

    // Replace the third parameter
    if ( Insert3.c_str() )
    {
        m_String.ReplaceI( L"%%3", Insert3.c_str() );
    }
    else
    {
        m_String.ReplaceI( L"%%3", L"<null>" );
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a string with the specified signed 16-bit value
//
//  Parameters:
//      pszString - pointer to string to format
//      value   - the value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringInt16( LPCWSTR pszString, short value )
{
    wchar_t szValue[ 16 ] = { 0 };    // Big enough for a short as a string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%d", value );
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified signed 32-bit value
//
//  Parameters:
//      pszString - pointer to string to format
//      value   - the value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringInt32( LPCWSTR pszString, int value )
{
    wchar_t szValue[ 16 ] = { 0 };    // Big enough for an int as a string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%ld", value );
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified pointer
//
//  Parameters:
//      pszString - pointer to string to format
//      pValue     - the value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringPointer( LPCWSTR pszString, const void* pValue )
{
    wchar_t szValue[ 32 ] = { 0 };    // Big enough for a pointer as a string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"0x%p", pValue );
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified size_t value
//
//  Parameters:
//      pszString - pointer to string to format
//      value     - the value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringSizeT( LPCWSTR pszString, size_t value )
{
    // Call the appropriate function depending on the size of size_t
    #if defined ( _WIN64 )

        return StringUInt64( pszString, value );

    #else

        DWORD uValue = PXSCastSizeTToUInt32( value );
        return StringUInt32( pszString, uValue );

    #endif  // _WIN64
}

//===============================================================================================//
//  Description:
//      Convert the specified string to an ANSI string
//
//  Parameters:
//      Text      - the string to convert
//      pszAnsi   - buffer to receive the ANSI string
//      numBytes  - size of the buffer in bytes
//
//  Returns:
//      Number or bytes written to the buffer
//===============================================================================================//
size_t Formatter::StringToAnsi( const String& Text, char* pszAnsi, size_t numBytes )
{
    if ( ( pszAnsi == nullptr ) || ( numBytes == 0 ) )
    {
        throw ParameterException( L"pszAnsi/ansiBytes", __FUNCTION__);
    }
    memset( pszAnsi, 0, numBytes );

    if ( Text.IsEmpty() )
    {
        return 0;     // Nothing to do
    }

    // Use the system's current code page
    return WideToMultiByte( Text.c_str(), pszAnsi, numBytes, CP_ACP );
}

//===============================================================================================//
//  Description:
//      Convert teh input string value to a double
//
//  Parameters:
//      pszString - pointer to string to convert
//
//  Returns:
//      double value, 0.0 on error just like atof
//===============================================================================================//
double Formatter::StringToDouble( const String& Input )
{
    double value;

    if ( Input.GetLength() == 0 )
    {
        throw ParameterException( L"Input", __FUNCTION__ );
    }

    errno = 0;
    value = _wtof( Input.c_str() );
    if ( errno == ERANGE )
    {
        throw BoundsException( Input.c_str(), __FUNCTION__ );
    }

    return value;
}

//===============================================================================================//
//  Description:
//      Convert the input string to a signed 32-bit number.
//
//  Parameters:
//      Input - string to convert to 32-bit value
//
//  Returns:
//      32-bit signed integer
//===============================================================================================//
int Formatter::StringToInt32( const String& Input )
{
    int value;

    if ( Input.GetLength() == 0 )
    {
        throw ParameterException( L"Input", __FUNCTION__ );
    }

    // Note, not handling other forms of zero e.g. +0
    if ( Input.IsOnlyChar( '0' ) )
    {
        return 0;
    }

    errno = 0;
    value = _wtoi( Input.c_str() );
    if ( value == 0 )
    {
        throw ParameterException( Input.c_str(), __FUNCTION__ );
    }

    if ( errno == ERANGE )
    {
        throw BoundsException( Input.c_str(), __FUNCTION__ );
    }

    return value;
}

//===============================================================================================//
//  Description:
//      Convert the input string to a signed 64-bit number.
//
//  Parameters:
//      Input - string to convert to 64-bit value
//
//  Returns:
//      64-bit signed integer
//===============================================================================================//
__int64 Formatter::StringToInt64( const String& Input )
{
    __int64 value;

    if ( Input.GetLength() == 0 )
    {
        throw ParameterException( L"Input", __FUNCTION__ );
    }

    // Note, not handling other forms of zero e.g. +0
    if ( Input.IsOnlyChar( '0' ) )
    {
        return 0;
    }

    errno = 0;
    value = _wtoi64( Input.c_str() );
    if ( value == 0 )
    {
        throw ParameterException( Input.c_str(), __FUNCTION__ );
    }

    if ( errno == ERANGE )
    {
        throw BoundsException( Input.c_str(), __FUNCTION__ );
    }

    return value;
}

//===============================================================================================//
//  Description:
//      Convert the specified string to low ascii. Out of range characters are replaces with
//      defaultChar
//
//  Parameters:
//      Input       - the string to convert
//      defaultChar - used for characters >=0x80
//      pLowAscii   - receive's the low ascii characters
//
//  Returns:
//      number of characters >=0x80, zero if none found.
//===============================================================================================//
size_t Formatter::StringToLowAscii( const String& Input, char defaultChar, CharArray* pLowAscii )
{
    size_t  numHigh = 0, i = 0;
    size_t  length  = Input.GetLength();
    wchar_t wch     = 0;

    if ( pLowAscii == nullptr )
    {
        throw ParameterException( L"pLowAscii", __FUNCTION__ );
    }

    // Preserve NULL
    if ( Input.c_str() == nullptr )
    {
        pLowAscii->Free();
        return 0;
    }
    pLowAscii->Allocate( length );

    for ( i = 0; i < length; i++ )
    {
        wch = Input.CharAt( i );
        if ( wch >= 0x80 )
        {
            pLowAscii->Append( defaultChar );
            numHigh++;
        }
        else
        {
            pLowAscii->Append( static_cast< char > ( 0xFF & wch ) );
        }
    }

    return numHigh;
}

//===============================================================================================//
//  Description:
//      Convert a string to an unsigned 32-bit number.
//
//  Parameters:
//      Input - string to convert to 32-bit value
//
//  Returns:
//      32-bit unsigned integer
//===============================================================================================//
DWORD Formatter::StringToUInt32( const String& Input )
{
    DWORD value;

    if ( Input.GetLength() == 0 )
    {
        throw ParameterException( L"Input", __FUNCTION__ );
    }

    // Note, not handling other forms of zero e.g. +0
    if ( Input.IsOnlyChar( '0' ) )
    {
        return 0;
    }

    errno = 0;
    value = wcstoul( Input.c_str(), nullptr, 10 );
    if ( value == 0 )
    {
        throw ParameterException( Input.c_str(), __FUNCTION__ );
    }

    if ( errno == ERANGE )
    {
        throw BoundsException( Input.c_str(), __FUNCTION__ );
    }

    return value;
}

//===============================================================================================//
//  Description:
//      Convert the input string to an unsigned 64-bit number.
//
//  Parameters:
//      Input - string to convert to unsigned 64-bit value
//
//  Returns:
//      UINT64
//===============================================================================================//
UINT64 Formatter::StringToUInt64( const String& Input )
{
    UINT64 value;

    if ( Input.GetLength() == 0 )
    {
        throw ParameterException( L"Input", __FUNCTION__ );
    }

    // Note, not handling other forms of zero e.g. +0
    if ( Input.IsOnlyChar( '0' ) )
    {
        return 0;
    }

    errno = 0;
    value = _wcstoui64( Input.c_str(), nullptr, 10 );
    if ( value == 0 )
    {
        throw ParameterException( L"Input", __FUNCTION__ );
    }

    if ( errno == ERANGE )
    {
        throw BoundsException( L"Input", __FUNCTION__ );
    }

    return value;
}

//===============================================================================================//
//  Description:
//      Format a string with the specified unsigned 16-bit value
//
//  Parameters:
//      pszString - pointer to string to format
//      value     - the value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringUInt16( LPCWSTR pszString, WORD value )
{
    wchar_t szValue[ 32 ] = { 0 };    // Big enough for a WORD as a string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%u", value );
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified unsigned 32-bit value
//
//  Parameters:
//      pszString - pointer to string to format
//      value    - the value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringUInt32( LPCWSTR pszString, DWORD value )
{
    wchar_t szValue[ 32 ] = { 0 };    // Big enough for a DWORD as a string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%lu", value );
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified unsigned 32-bit value as hex
//
//  Parameters:
//      pszString    - pointer to string to format
//      value        - the value to be displayed in hex format
//      leadingZeroX - true if want the leading 0x

//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the hex value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringUInt32Hex( LPCWSTR pszString, DWORD value, bool leadingZeroX )
{
    wchar_t szValue[ 32 ] = { 0 };    // Big enough for a DWORD as a hex string
    String Insert1, Insert2, Insert3;

    if ( leadingZeroX )
    {
       StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"0x%08X", value );
    }
    else
    {
       StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%08X", value );
    }
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the two specified unsigned 32-bit values
//
//  Parameters:
//      pszString - pointer to string to format
//      value1    - the first value
//      value2    - the second value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value1 and
//      %%2 is replaced with value2
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringUInt32_2( LPCWSTR pszString, DWORD value1, DWORD value2 )
{
    wchar_t  szValue1[ 16 ] = { 0 };    // Big enough for a DWORD as a string
    wchar_t  szValue2[ 16 ] = { 0 };    // Big enough for a DWORD as a string
    String Insert1, Insert2;

    StringCchPrintf( szValue1, ARRAYSIZE( szValue1 ), L"%lu", value1 );
    StringCchPrintf( szValue2, ARRAYSIZE( szValue2 ), L"%lu", value2 );
    Insert1 = szValue1;
    Insert2 = szValue2;

    return String2( pszString, Insert1, Insert2 );
}

//===============================================================================================//
//  Description:
//      Format a string with the three specified unsigned 32-bit values
//
//  Parameters:
//      pszString - pointer to string to format
//      value1    - the first value
//      value2    - the second value
//      value3    - the third value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value1, %%2 with value2 and %%3
//      with value3
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringUInt32_3( LPCWSTR pszString,
                                         DWORD value1, DWORD value2, DWORD value3 )
{
    wchar_t  szValue1[ 16 ] = { 0 };    // Big enough for a DWORD as a string
    wchar_t  szValue2[ 16 ] = { 0 };    // Big enough for a DWORD as a string
    wchar_t  szValue3[ 16 ] = { 0 };    // Big enough for a DWORD as a string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue1, ARRAYSIZE( szValue1 ), L"%lu", value1 );
    StringCchPrintf( szValue2, ARRAYSIZE( szValue2 ), L"%lu", value2 );
    StringCchPrintf( szValue3, ARRAYSIZE( szValue3 ), L"%lu", value3 );
    Insert1 = szValue1;
    Insert2 = szValue2;
    Insert3 = szValue3;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified unsigned 64-bit value
//
//  Parameters:
//      pszString - pointer to string to format
//      value  - the value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringUInt64( LPCWSTR pszString, UINT64 value )
{
    wchar_t szValue[ 32 ] = { 0 };    // Big enough for a uint64 as a string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%llu", value );
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified unsigned 64-bit value as hex
//
//  Parameters:
//      pszString - pointer to string to format
//      value  - the value to be displayed in hex format
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the hex value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringUInt64Hex( LPCWSTR pszString, UINT64 value )
{
    wchar_t szValue[ 32 ] = { 0 };    // Big enough for a uint64 as a hex string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"0x%016llX", value );
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Format a string with the specified unsigned 8-bit value
//
//  Parameters:
//      pszString - pointer to string to format
//      value     - the value
//
//  Remarks:
//      Occurrences of %%1 in pszString are replaced with the value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::StringUInt8( LPCWSTR pszString, BYTE value )
{
    wchar_t szValue[ 8 ] = { 0 };    // Big enough for a BYTE as a string
    String Insert1, Insert2, Insert3;

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%u", value );
    Insert1 = szValue;

    return String3( pszString, Insert1, Insert2, Insert3 );
}

//===============================================================================================//
//  Description:
//      Convert the specified string to wide string
//
//  Parameters:
//      Text     - the string
//      pwzWide  - buffer to receive the Unicode string
//      numChars - size of the buffer in characters
//
//  Remarks:
//      The input buffer pwzWide is terminated on return.
//
//  Returns:
//      Number or characters written to the buffer
//===============================================================================================//
size_t Formatter::StringToWide( const String& Text, wchar_t* pwzWide, size_t numChars )
{
    size_t numCopied = 0;

    if ( ( pwzWide == nullptr ) || ( numChars == 0 ) )
    {
        throw ParameterException( L"pwzWide/numChars", __FUNCTION__ );
    }
    wmemset( pwzWide, 0, numChars );

    if ( Text.IsEmpty() )
    {
        return 0;     // Nothing to do
    }
    StringCchCopy( pwzWide, numChars, Text.c_str() );
    numCopied = wcslen( pwzWide );

    return numCopied;
}

//===============================================================================================//
//  Description:
//      Format a system error number to a string
//
//  Parameters:
//      errorCode - the system error code, ie. returned form GetLastError
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::SystemError( DWORD errorCode )
{
    wchar_t szBuffer[ 1024 ] = { 0 };      // Use a large buffer

    m_String = PXS_STRING_EMPTY;
    FormatMessage( FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_FROM_SYSTEM,
                   nullptr,
                   errorCode,
                   MAKELANGID( LANG_NEUTRAL, SUBLANG_DEFAULT ),
                   szBuffer,
                   ARRAYSIZE( szBuffer ),
                   nullptr );
    szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = PXS_CHAR_NULL;  // Teminate
    m_String = szBuffer;
    m_String.Trim();        // Sometimes there is a CRLF
    if ( ( m_String.GetLength() ) &&
         ( m_String.EndsWithCharacter( PXS_CHAR_DOT ) == false ) )
    {
        m_String += PXS_CHAR_DOT;
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert a SYSTEMTIME structure to ISO format YYYY-MM-DD hh:mm:ss
//
//  Parameters:
//      systemTime - a SYSTEMTIME structure
//
//  Remarks:
//      System Time in this context is the contents of the structure
//      not the computer's system time (UTC)
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::SystemTimeToIso( const SYSTEMTIME& systemTime )
{
    return SystemTimeToString( systemTime, true, true );
}

//===============================================================================================//
//  Description:
//      Format a system time to a string
//
//  Parameters:
//      systemTime  - a SYSTEMTIME structure in UTC
//      wantTime    - if true adds time to the formatted string
//      wantISO     - if format is ISO otherwise in user default
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::SystemTimeToString( const SYSTEMTIME& systemTime,
                                             bool wantTime, bool wantISO )
{
    DWORD      timeFlags = 0;
    wchar_t      szBuffer[ 64 ];   // Enough to hold any date or time value
    LPCWSTR    pszFormatDate = nullptr;
    LPCWSTR    pszFormatTime = nullptr;
    SYSTEMTIME stZero;

    m_String = PXS_STRING_EMPTY;

    // Filter zero
    memset( &stZero, 0, sizeof ( stZero ) );
    if ( memcmp( &systemTime, &stZero , sizeof ( systemTime ) ) == 0 )
    {
        return m_String;
    }

    if ( IsValidSystemTime( systemTime ) == false )
    {
        throw ParameterException( L"pSystemTime", __FUNCTION__ );
    }

    // Format for ISO
    if ( wantISO )
    {
        pszFormatDate = L"yyyy'-'MM'-'dd";
        pszFormatTime = L"hh':'mm':'ss";
        timeFlags = TIME_FORCE24HOURFORMAT;
    }

    // Date
    memset( szBuffer, 0, sizeof ( szBuffer ) );
    if ( GetDateFormat( LOCALE_USER_DEFAULT,
                        0, &systemTime, pszFormatDate, szBuffer, ARRAYSIZE( szBuffer ) ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetDateFormat", __FUNCTION__ );
    }
    szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = PXS_CHAR_NULL;  // Terminate
    m_String = szBuffer;

    // Time
    if ( wantTime )
    {
        memset( szBuffer, 0, sizeof ( szBuffer ) );
        if ( GetTimeFormat( LOCALE_USER_DEFAULT,
                            timeFlags,
                            &systemTime, pszFormatTime, szBuffer, ARRAYSIZE( szBuffer ) ) == 0 )
        {
            throw SystemException( GetLastError(), L"GetTimeFormat", __FUNCTION__ );
        }
        szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = PXS_CHAR_NULL;  // Terminate
        m_String += PXS_CHAR_SPACE;
        m_String += szBuffer;
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a time in the user's locale
//
//  Parameters:
//      hour   - hour
//      minute - minute
//      second - second
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::TimeInUserLocale( DWORD hour, DWORD minute, DWORD second )
{
    wchar_t     szTime[ 32 ] = { 0 };  // Long enough for any time format
    SYSTEMTIME  st;

    if ( ( hour > 23 ) || ( minute > 59 ) || ( second > 59 ) )
    {
        throw ParameterException( L"Hour/Minute/Second", __FUNCTION__);
    }

    // Use GetTimeFormat, it does not require DayOfWeek
    memset( &st, 0, sizeof ( st ) );
    st.wYear    = 2000;
    st.wMonth   = 1;
    st.wDay     = 1;
    st.wHour    = PXSCastUInt32ToUInt16( hour   );
    st.wMinute  = PXSCastUInt32ToUInt16( minute );
    st.wSecond  = PXSCastUInt32ToUInt16( second );

    GetTimeFormat( LOCALE_USER_DEFAULT, 0, &st, nullptr, szTime, ARRAYSIZE( szTime ) );
    szTime[ ARRAYSIZE( szTime ) - 1 ] = PXS_CHAR_NULL;
    m_String = szTime;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert a time_t value to the local time and express it in ISO
//      format YYYY-MM-DD hh:mm:ss
//
//  Parameters:
//      timeT - the time in seconds from 1/1/1970 00:00:00 (UTC)
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::TimeTToLocalTimeInIso( time_t timeT )
{
    FILETIME   FileTime;
    FILETIME   LocalFileTime   = { 0 };
    SYSTEMTIME LocalSystemTime = { 0 };

    m_String = PXS_STRING_EMPTY;

    // Filter for zero, return ""
    if ( timeT == 0 )
    {
        return m_String;
    }

    if ( timeT < 0 )
    {
        throw ParameterException( L"timeT", __FUNCTION__ );
    }
    memset( &FileTime       , 0, sizeof ( FileTime        ) );
    memset( &LocalFileTime  , 0, sizeof ( LocalFileTime   ) );
    memset( &LocalSystemTime, 0, sizeof ( LocalSystemTime ) );

    PXSCastTimeTToFileTime( timeT, &FileTime );
    if ( FileTimeToLocalFileTime( &FileTime, &LocalFileTime ) == 0 )
    {
       throw SystemException( GetLastError(), L"FileTimeToLocalFileTime", __FUNCTION__);
    }

    if ( FileTimeToSystemTime( &LocalFileTime, &LocalSystemTime ) == 0 )
    {
        throw SystemException( GetLastError(), L"FileTimeToSystemTime", __FUNCTION__ );
    }

    return SystemTimeToString( LocalSystemTime, true, true );
}

//===============================================================================================//
//  Description:
//      Format a WORD in using the %d specifier
//
//  Parameters:
//      value - the WORD value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UInt16( WORD value )
{
    wchar_t szBuffer[ 8 ] = { 0 };      // Large enough to hold any WORD

    StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%d", value );
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format an unsigned 16-bit integer using the 0x%04X specifier
//
//  Parameters:
//      value - integer to format
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UInt16Hex( WORD value )
{
    wchar_t szValue[ 16 ] = { 0 };      // Large enough to hold any hex string

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"0x%04X", value );
    m_String = szValue;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a DWORD using the %lu specifier
//
//  Parameters:
//      value - the DWORD value
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::UInt32( DWORD value )
{
    wchar_t szValue[ 16 ] = { 0 };      // Large enough to hold any DWORD

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%lu", value );
    m_String = szValue;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format an unsigned 32-bit integer using the 0x%08X specifier
//
//  Parameters:
//      value     - integer to format
//      leadingZeroX - indicates if want the leading 0x
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UInt32Hex( DWORD value, bool leadingZeroX )
{
    wchar_t szValue[ 16 ] = { 0 };      // Large enough to hold any hex string

    if ( leadingZeroX )
    {
        StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"0x%08X", value);
    }
    else
    {
        StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%08X", value );
    }
    m_String = szValue;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a value to "Yes" or "No" or a translation taken from the
//      application's currently loaded string table.
//
//  Parameters:
//      value - the value, zero = No otherwise Yes
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::UInt32YesNo( DWORD value )
{
    if ( value )
    {
        return Int32YesNo( 1 );
    }

    return Int32YesNo( 0 );
}

//===============================================================================================//
//  Description:
//      Get an unsigned 64-bit signed integer in string format
//
//  Parameters:
//      value - the 64-bit value
//
//  Returns:
//      Reference to formatted string
//===============================================================================================//
const String& Formatter::UInt64( UINT64 value )
{
    wchar_t szValue[ 32 ] = { 0 };  // Enough for a 64-bit unsigned integer

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%llu", value );
    m_String = szValue;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format an unsigned 64-bit integer using the 0x%08X specifier
//
//  Parameters:
//      value        - integer to format
//      leadingZeroX - indicates if want the leading 0x
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UInt64Hex( UINT64 value, bool leadingZeroX )
{
    wchar_t szBuffer[ 32 ] = { 0 };      // Large enough to hold any hex string

    if ( leadingZeroX )
    {
        // Note, ll must be lower case even if want upper case hex
        StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"0x%016llX", value);
    }
    else
    {
        StringCchPrintf( szBuffer, ARRAYSIZE( szBuffer ), L"%016llX", value );
    }
    m_String = szBuffer;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format a byte value using as an integer
//
//  Parameters:
//      value - the byte value
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UInt8( BYTE value )
{
    wchar_t szValue[ 8 ] = { 0 };      // Large enough to hold any byte

    StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%u", value );
    m_String = szValue;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Format an unsigned 8-bit integer using the %02X specifier
//
//  Parameters:
//      value - the byte to format
//      leadingZeroX - flag to indicate want the leading 0x
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UInt8Hex( BYTE value, bool leadingZeroX )
{
    wchar_t szValue[ 8 ] = { 0 };      // Large enough to hold any hex string

    if ( leadingZeroX )
    {
        StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"0x%02X", value );
    }
    else
    {
        StringCchPrintf( szValue, ARRAYSIZE( szValue ), L"%02X", value );
    }
    m_String = szValue;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert a UNICODE_STRING to the encoding in use
//
//  Parameters:
//      pUnicodeString - pointer to the Unicode string
//
//  Remarks:
//      The Buffer member of UNICODE_STRING is not necessarily terminated
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UnicodeString( const UNICODE_STRING* pUnicodeString )
{
    size_t   numChars = 0;
    wchar_t* pwzWide  = nullptr;
    AllocateWChars AllocWChars;

    m_String = PXS_STRING_EMPTY;
    if ( ( pUnicodeString == nullptr ) || ( pUnicodeString->Buffer == nullptr) )
    {
        return m_String;     // Nothing to do
    }
    numChars = pUnicodeString->Length / sizeof ( wchar_t );
    numChars = PXSAddSizeT( numChars, 1 );       // Null terminator
    pwzWide  = AllocWChars.New( numChars );

    memcpy( pwzWide, pUnicodeString->Buffer, pUnicodeString->Length );
    pwzWide[ numChars - 1 ] = PXS_CHAR_NULL;
    m_String = pwzWide;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert the input Unix time to IMF-fixdate format
//
//  Parameters:
//      Utf8Chars - the UTF-8 characters
//
//  Remarks:
//      Format is 29 characters long: Sun, 06 Nov 1994 08:49:37 GMT
//      This is the preferred date format used in HTTP. See RFC 7321, section 7.1.1.1.
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UnixTimeToIMFfixdate( time_t unixTime )
{
    struct  tm newtime;
    wchar_t wzBuffer[ 16 ] = { 0 };     // Enough to hold any signed integer

    m_String.Allocate( 32 );            // Reduce memory allocs on concat
    m_String = PXS_STRING_EMPTY;
    if ( unixTime < 0 )
    {
        throw ParameterException( L"unixTime", __FUNCTION__ );
    }

    memset( &newtime, 0, sizeof ( newtime ) );
    if ( gmtime_s( &newtime, &unixTime ) )
    {
        throw SystemException( ERROR_INVALID_DATA, L"gmtime_s", __FUNCTION__ );
    }

    // day-name is case sensitive
    switch ( newtime.tm_wday )
    {
        default:
            throw SystemException( ERROR_INVALID_DATA, L"tm_wday", __FUNCTION__ );
        case 0: m_String += L"Sun, "; break;
        case 1: m_String += L"Mon, "; break;
        case 2: m_String += L"Tue, "; break;
        case 3: m_String += L"Wed, "; break;
        case 4: m_String += L"Thu, "; break;
        case 5: m_String += L"Fri, "; break;
        case 6: m_String += L"Sat, "; break;
    }

    // day
    if ( ( newtime.tm_mday < 1 ) || ( newtime.tm_mday > 31 ) )
    {
        throw SystemException( ERROR_INVALID_DATA, L"tm_mday", __FUNCTION__ );
    }
    memset( wzBuffer, 0, sizeof ( wzBuffer ) );
    StringCchPrintf( wzBuffer, ARRAYSIZE( wzBuffer ), L"%02ld", newtime.tm_mday );
    wzBuffer[ ARRAYSIZE( wzBuffer ) - 1 ] = PXS_CHAR_NULL;
    m_String += wzBuffer;
    m_String += L" ";

    // month
    switch ( newtime.tm_mon )
    {
        default:
            throw SystemException( ERROR_INVALID_DATA, L"tm_mon", __FUNCTION__ );
        case  0: m_String += L"Jan "; break;
        case  1: m_String += L"Feb "; break;
        case  2: m_String += L"Mar "; break;
        case  3: m_String += L"Apr "; break;
        case  4: m_String += L"May "; break;
        case  5: m_String += L"Jun "; break;
        case  6: m_String += L"Jul "; break;
        case  7: m_String += L"Aug "; break;
        case  8: m_String += L"Sep "; break;
        case  9: m_String += L"Oct "; break;
        case 10: m_String += L"Nov "; break;
        case 11: m_String += L"Dec "; break;
    }

    // year
    newtime.tm_year = PXSAddInt32( newtime.tm_year, 1900 );
    if ( ( newtime.tm_year < 1970 ) || ( newtime.tm_mday > 9999 ) )
    {
        throw SystemException( ERROR_INVALID_DATA, L"tm_year", __FUNCTION__ );
    }
    memset( wzBuffer, 0, sizeof ( wzBuffer ) );
    StringCchPrintf( wzBuffer, ARRAYSIZE( wzBuffer ), L"%04ld", newtime.tm_year );
    wzBuffer[ ARRAYSIZE( wzBuffer ) - 1 ] = PXS_CHAR_NULL;
    m_String += wzBuffer;
    m_String += L" ";

    // hour
    if ( ( newtime.tm_hour < 0 ) || ( newtime.tm_hour > 23 ) )
    {
        throw SystemException( ERROR_INVALID_DATA, L"tm_hour", __FUNCTION__ );
    }
    memset( wzBuffer, 0, sizeof ( wzBuffer ) );
    StringCchPrintf( wzBuffer, ARRAYSIZE( wzBuffer ), L"%02ld", newtime.tm_hour );
    wzBuffer[ ARRAYSIZE( wzBuffer ) - 1 ] = PXS_CHAR_NULL;
    m_String += wzBuffer;
    m_String += L":";

    // minute
    if ( ( newtime.tm_min < 0 ) || ( newtime.tm_min > 59 ) )
    {
        throw SystemException( ERROR_INVALID_DATA, L"tm_min", __FUNCTION__ );
    }
    memset( wzBuffer, 0, sizeof ( wzBuffer ) );
    StringCchPrintf( wzBuffer, ARRAYSIZE( wzBuffer ), L"%02ld", newtime.tm_min );
    wzBuffer[ ARRAYSIZE( wzBuffer ) - 1 ] = PXS_CHAR_NULL;
    m_String += wzBuffer;
    m_String += L":";

    // second
    if ( ( newtime.tm_sec < 0 ) || ( newtime.tm_sec > 59 ) )
    {
        throw SystemException( ERROR_INVALID_DATA, L"tm_sec", __FUNCTION__ );
    }
    memset( wzBuffer, 0, sizeof ( wzBuffer ) );
    StringCchPrintf( wzBuffer, ARRAYSIZE( wzBuffer ), L"%02ld", newtime.tm_sec );
    wzBuffer[ ARRAYSIZE( wzBuffer ) - 1 ] = PXS_CHAR_NULL;
    m_String += wzBuffer;

    // GMT
    m_String += L" GMT";

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert the specified UTF-8 characters to a string
//
//  Parameters:
//      Utf8Chars - the UTF-8 characters
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UTF8ToWide( const CharArray& Utf8Chars )
{
    return UTF8ToWide( Utf8Chars.GetPtr(), Utf8Chars.GetSize() );
}

//===============================================================================================//
//  Description:
//      Convert specified UTF-8 string to a string
//
//  Parameters:
//      pszUtf8  - pointer to the UTF-8
//      numChars - number of chars
//
//  Returns:
//      Reference to the formatted string
//===============================================================================================//
const String& Formatter::UTF8ToWide( const char* pszUtf8, size_t numChars )
{
    int      cchMultiByte, cchWideChar, charsCopied;
    wchar_t* pWideCharStr;
    AllocateWChars Alloc;

    m_String = PXS_STRING_EMPTY;

    if ( ( pszUtf8 == nullptr ) || ( *pszUtf8 == '\0' ) || ( numChars == 0 ) )
    {
        return m_String;   // Nothing to do
    }

    cchMultiByte = PXSCastSizeTToInt32( numChars );
    numChars     = PXSAddSizeT( numChars, 1 );     // Null terminator
    pWideCharStr = Alloc.New( numChars );
    cchWideChar  = PXSCastSizeTToInt32( numChars );
    charsCopied  = MultiByteToWideChar( CP_UTF8,
                                        MB_ERR_INVALID_CHARS,
                                        pszUtf8, cchMultiByte, pWideCharStr, cchWideChar );
    pWideCharStr[ cchWideChar - 1 ] = PXS_CHAR_NULL;

    // The bytes copied includes the terminator so zero is error
    if ( charsCopied == 0 )
    {
        throw SystemException( GetLastError(), L"MultiByteToWideChar", __FUNCTION__ );
    }
    m_String = pWideCharStr;

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert a wide string to a UTF-8 string
//
//  Parameters:
//      lpwzString - pointer to wide string
//      pszUTF8   - pointer to buffer to receive the UTF-8 string
//      utf8Bytes - the size of buffer in bytes
//
//  Returns:
//      The number of bytes copied into UTF-8 buffer
//===============================================================================================//
size_t Formatter::WideToUTF8( LPCWSTR pwzWide, char* pszUTF8, size_t utf8Bytes )
{
    return WideToMultiByte( pwzWide, pszUTF8, utf8Bytes, CP_UTF8 );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Test if the specified language id is Bosnian
//
//  Parameters:
//      languageID - the language ID
//
//  Returns:
//      true if the language id is Bosnian otherwise false
//===============================================================================================//
bool Formatter::IsBosnian( WORD languageID )
{
    BYTE subID = HIBYTE( languageID );

    if ( LOBYTE( languageID ) != LANG_BOSNIAN )
    {
        return false;
    }

    if ( languageID == LANG_BOSNIAN_NEUTRAL )  // 0x781a
    {
        return true;
    }

    if ( ( subID == SUBLANG_BOSNIAN_BOSNIA_HERZEGOVINA_LATIN    ) ||
         ( subID == SUBLANG_BOSNIAN_BOSNIA_HERZEGOVINA_CYRILLIC )  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Test if the specified language id is Croatian
//
//  Parameters:
//      languageID - the language ID
//
//  Returns:
//      true if the language id is Croatian otherwise false
//===============================================================================================//
bool Formatter::IsCroatian( WORD languageID )
{
    BYTE subID = HIBYTE( languageID );

    if ( LOBYTE( languageID ) != LANG_CROATIAN )
    {
        return false;
    }

    if ( ( subID == SUBLANG_CROATIAN_CROATIA ) ||
         ( subID == SUBLANG_CROATIAN_BOSNIA_HERZEGOVINA_LATIN ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Test if the specified language id is Serbian
//
//  Parameters:
//      languageID - the language ID
//
//  Returns:
//      true if the language id is Serbian otherwise false
//===============================================================================================//
bool Formatter::IsSerbian( WORD languageID )
{
    bool   result = false;
    BYTE   subID  = HIBYTE( languageID );
    size_t i = 0;

    if ( LOBYTE( languageID ) != LANG_SERBIAN )
    {
        return false;
    }

    if ( LANG_SERBIAN_NEUTRAL == languageID )     // 0x7c1a
    {
        return true;
    }

    BYTE SubIDS[] = { SUBLANG_SERBIAN_BOSNIA_HERZEGOVINA_LATIN,
                      SUBLANG_SERBIAN_BOSNIA_HERZEGOVINA_CYRILLIC,
                      SUBLANG_SERBIAN_MONTENEGRO_LATIN,
                      SUBLANG_SERBIAN_MONTENEGRO_CYRILLIC,
                      SUBLANG_SERBIAN_SERBIA_LATIN,
                      SUBLANG_SERBIAN_SERBIA_CYRILLIC,
                      SUBLANG_SERBIAN_CROATIA,
                      SUBLANG_SERBIAN_LATIN,
                      SUBLANG_SERBIAN_CYRILLIC };

    for ( i = 0; i < ARRAYSIZE( SubIDS ); i++ )
    {
        if ( subID == SubIDS[ i ] )
        {
            result = true;
            break;
        }
    }

    return result;
}

//===============================================================================================//
//  Description:
//      Format into binary up to a a unsigned 64 bit number
//
//  Parameters:
//      value   - the value to format
//      numBits - the number of bits, 8 for a byte etc.
//
//  Returns:
//      constant string reference
//===============================================================================================//
const String& Formatter::NumberToBinary( UINT64 value, size_t numBits )
{
    size_t i = 0;

    // Allocate some space and reset the string
    m_String.Allocate( 128 );
    m_String = PXS_STRING_EMPTY;

    // Number of bits cannot be more that 64
    if ( numBits > 64 )
    {
        throw ParameterException( L"numBits", __FUNCTION__ );
    }

    for ( i = 0; i < numBits; i++ )
    {
        if ( ( value & ( 1ULL << ( numBits - 1 - i ) ) ) == 0 )
        {
            m_String += PXS_STRING_ZERO;
        }
        else
        {
            m_String += PXS_STRING_ONE;
        }
    }

    return m_String;
}

//===============================================================================================//
//  Description:
//      Convert a wide character string (aka 2-byte, aka Unicode) to a
//      multi-byte string in the specified code page
//
//  Parameters:
//      pwzWide        - pointer to wide character string
//      pszMultiByte   - pointer to a buffer to receive string
//      multiByteChars - size of multi-byte buffer in bytes/chars
//      codePage       - the code page to use for the conversion
//
//  Returns:
//      The number of bytes/chars copied into string
//===============================================================================================//
size_t Formatter::WideToMultiByte( LPCWSTR pwzWide,
                                   char* pszMultiByte, size_t multiByteChars, UINT codePage )
{
    int bytesCopied = 0, nMultiByteChars = 0;

    m_String = PXS_STRING_EMPTY;
    if ( pwzWide == nullptr )
    {
        return 0;   // Nothing to do
    }

    if ( ( pszMultiByte == nullptr ) || ( multiByteChars == 0 ) )
    {
        throw ParameterException( L"MultiByte", __FUNCTION__ );
    }
    memset( pszMultiByte, 0, multiByteChars );

    nMultiByteChars = PXSCastSizeTToInt32( multiByteChars );
    bytesCopied     = WideCharToMultiByte( codePage,
                                           WC_NO_BEST_FIT_CHARS,  // Security
                                           pwzWide,
                                           -1, pszMultiByte, nMultiByteChars, nullptr, nullptr );
    pszMultiByte[ multiByteChars - 1 ] = 0;

    // bytesCopied includes the terminator so zero is error
    if ( bytesCopied == 0 )
    {
        // Try again if WC_NO_BEST_FIT_CHARS caused a problem
        if ( GetLastError() == ERROR_INVALID_FLAGS )
        {
            bytesCopied = WideCharToMultiByte( codePage,
                                               0,
                                               pwzWide,
                                               -1,
                                               pszMultiByte, nMultiByteChars, nullptr, nullptr );
        }
    }

    if ( bytesCopied == 0 )
    {
        throw SystemException( GetLastError(), L"WideCharToMultiByte", __FUNCTION__ );
    }

    return PXSCastInt32ToSizeT( bytesCopied );
}
