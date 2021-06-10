///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Formatter Class Header
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

#ifndef PXSBASE_FORMAT_H_
#define PXSBASE_FORMAT_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files
#include <Ntsecapi.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/CharArray.h"
#include "PxsBase/Header Files/StringT.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Formatter
{
    public:
        // Default constructor
        Formatter();

        // Copy constructor
        Formatter( const Formatter& oFormatter );

        // Destructor
        ~Formatter();

        // Assignment operator
        Formatter& operator= ( const Formatter& oFormatter );

        // Methods

        // Number Formatting
        const String& BinarySizeT( size_t value );
        const String& BinaryUInt8( BYTE value );
        const String& BinaryUInt32( DWORD value );
        const String& BinaryUInt64( UINT64 value );
        const String& Bool( bool value );
        const String& Boolean( BOOLEAN value );
        const String& Double( double value );
        const String& Double( double value, BYTE decimalPlaces );
        const String& Float( float value );
        const String& Float( float value, BYTE decimalPlaces );
        const String& Int32( int value );
        const String& Int32Bool( BOOL value );
        const String& Int32Hex( int value );
        const String& Int32YesNo( int value );
        const String& Int64( __int64 value );
        const String& Pointer( const void* pValue );
        const String& SizeT( size_t value );
                void  SizeT( size_t value, CharArray* pArray );
        const String& StorageBytes( UINT64 value );
        const String& UInt8( BYTE value );
        const String& UInt8Hex( BYTE value, bool leadingZeroX );
        const String& UInt16( WORD value );
        const String& UInt16Hex( WORD value );
        const String& UInt32( DWORD value );
        const String& UInt32Hex( DWORD value, bool leadingZeroX );
        const String& UInt32YesNo( DWORD value );
        const String& UInt64( UINT64 value );
        const String& UInt64Hex( UINT64 value, bool leadingZeroX );

        // String to number conversion
        DWORD   HexStringToNumber( const String& HexString );
        double  StringToDouble( const String& Input );
        int     StringToInt32( const String& Input );
        __int64 StringToInt64( const String& Input );
        DWORD   StringToUInt32( const String& Input );
        UINT64  StringToUInt64( const String& Input );

        // ANSI/Unicode/Utf8 conversion
        size_t AnsiToWide( const char* pszAnsi,
                           wchar_t* pwzWide, size_t numWideChars );
        size_t StringToAnsi( const String& Text,
                             char* pszAnsi, size_t ansiBytes );
        size_t StringToLowAscii( const String& Input, char defaultChar, CharArray* pLowAscii );
        size_t StringToWide( const String& Text,
                             wchar_t* pwzWide, size_t numChars );
const String&  UTF8ToWide( const CharArray& Utf8Chars );
const String&  UTF8ToWide( const char* pszUtf8, size_t numChars );
        size_t WideToUTF8( LPCWSTR pwzWide, char* pszUTF8, size_t utf8Bytes );

        // String Formatting
       const String& GetModuleString( const String& ModulePath,
                                       DWORD messageID, const StringArray& Inserts );
        const String& GuidToString( const GUID& guid );
        const String& SidToString( PSID pSid );
        const String& String1( LPCWSTR pszString, const String& Insert1 );
        const String& String2( LPCWSTR pszString, const String& Insert1, const String& Insert2 );
        const String& String3( LPCWSTR pszString,
                               const String& Insert1,
                               const String& Insert2, const String& Insert3 );
        const String& StringInt16( LPCWSTR pszString, short value );
        const String& StringInt32( LPCWSTR pszString, int value );
        const String& StringPointer( LPCWSTR pszString, const void* pValue );
        const String& StringSizeT( LPCWSTR pszString, size_t value );
        const String& StringUInt8( LPCWSTR pszString, BYTE value );
        const String& StringUInt16( LPCWSTR pszString, WORD value );
        const String& StringUInt32( LPCWSTR pszString, DWORD value );
        const String& StringUInt32_2( LPCWSTR pszString, DWORD dwValue1, DWORD dwValue2 );
        const String& StringUInt32_3( LPCWSTR pszString,
                                      DWORD value1, DWORD value2, DWORD value3 );
        const String& StringUInt32Hex( LPCWSTR pszString, DWORD value, bool leadingZeroX );
        const String& StringUInt64( LPCWSTR pszString, UINT64 value );
        const String& StringUInt64Hex( LPCWSTR pszString, UINT64 value );
        const String& SystemError( DWORD errorCode );
        const String& UnicodeString( const UNICODE_STRING* pUnicodeString );

        // Date-Time functions
        const String& DateInUserLocale( WORD year, WORD month, WORD day );
        const String& FileTimeToLocalTimeIso( const FILETIME& fileTime );
        const String& FileTimeToLocalTimeUser( const FILETIME& fileTime, bool wantTime  );
        const String& FileTimeToString( const FILETIME& fileTime, bool wantTime , bool wantISO );
        const String& IsoTimestampToLocale( const String& Timestamp, bool wantTime );
        const String& LocalTimeInIsoFormat();
        const String& LocalTimeInUserLocale();
        const String& NowTohhmmssttt();
        const String& NowToYYYYMMDDhhmmsstttt();
        const String& OleTimeInIsoFormat( const DATE& oleDate );
        const String& SecondsToDDHHMM( time_t seconds );
        const String& SystemTimeToIso( const SYSTEMTIME& systemTime );
        const String& SystemTimeToString( const SYSTEMTIME& systemTime,
                                          bool wantTime, bool wantISO );
        const String& TimeInUserLocale( DWORD hour, DWORD minute, DWORD second );
        const String& TimeTToLocalTimeInIso( time_t timeT );
        const String& UnixTimeToIMFfixdate( time_t unixTime );

        // Validation Functions
        bool  IsValidIsoTimestamp( const String& Timestamp, SYSTEMTIME* pSystemTime );
        bool  IsValidMacAddress( const String& MacAddress );
        bool  IsValidStringGuid( const String& StringGuid );
        bool  IsValidSystemTime( const SYSTEMTIME& systemTime );
        bool  IsValidYearMonthDay( int year, int month, int day );

        // Others
        const String& AddressToIPv4( DWORD address );
        const String& Formatter::AddressToIPv6( BYTE* pAddress, size_t numBytes );
        const String& CreateGuid();
        const String& LanguageIdToName( WORD languageId );

    protected:
        // Methods

        // Data members

    private:
        // Methods
        bool   IsBosnian( WORD languageID );
        bool   IsCroatian( WORD languageID );
        bool   IsSerbian( WORD languageID );
const String&  NumberToBinary( UINT64 value, size_t numBits );
        size_t WideToMultiByte( LPCWSTR pwzWide,
                                char* pszMultiByte, size_t multiByteChars, UINT codePage );


        // Data members
        String  m_String;
};

#endif  // PXSBASE_FORMAT_H_
