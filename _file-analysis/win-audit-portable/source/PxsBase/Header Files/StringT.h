///////////////////////////////////////////////////////////////////////////////////////////////////
//
// String Class Header
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

#ifndef PXSBASE_STRING_H_
#define PXSBASE_STRING_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards
class CharArray;
class NameValue;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class String
{
    public:
        // Default constructor
        String();

        // Copy constructor
        String( const String& oString );

        // Destructor
        ~String();

        // Operators
        String& operator  = ( const String& oString );
        String& operator += ( const String& oString );

        String& operator  = ( LPCWSTR pszString );
        String& operator += ( LPCWSTR pszString );

        String& operator  = ( wchar_t ch );
        String& operator += ( wchar_t ch );

        // Methods
        void    Allocate( size_t numChars );
        void    Append( LPCWSTR pszString );
        void    AppendChar( wchar_t ch );
        void    AppendChar( wchar_t ch, size_t repeat );
        void    AppendChars( LPCWSTR pszString, size_t numChars );
        void    AppendFrom( size_t idxFrom, String* pInput ) const;
        void    AppendString( const String& Input );
        LPCWSTR c_str() const;
        wchar_t CharAt( size_t index ) const;
        int     Compare( const String& Input, bool caseSensitive ) const;
        int     Compare( LPCWSTR pszString, bool caseSensitive ) const;
        int     CompareI( const String& Input ) const;
        int     CompareI( LPCWSTR pszString ) const;
        size_t  CountOfChar( wchar_t wch );
        void    Empty();
        bool    EndsWithCharacter( wchar_t ch ) const;
        bool    EndsWithStringI( LPCWSTR pszString ) const;
        void    EscapeForHtml();
        void    EscapeForRichText();
        void    FixedWidth( size_t maxWidth, wchar_t ch );
        size_t  GetAnsiMultiByteLength() const;
        size_t  GetLength() const;
        size_t  IndexesOf( LPCWSTR pszString,
                           bool caseSensitive, TArray< size_t >* pIndexes ) const;
        size_t  IndexOf( wchar_t ch, size_t fromIndex ) const;
        size_t  IndexOf( LPCWSTR pszString, bool caseSensitive, size_t fromIndex ) const;
        size_t  IndexOfI( LPCWSTR pszString ) const;
        bool    IsEmpty() const;
        bool    IsNull() const;
        bool    IsOnly8Bit() const;
        bool    IsOnlyChar( wchar_t ch ) const;
        bool    IsOnlyDigits() const;
        bool    IsOnlyUSAscii() const;
        bool    IsOnlyUSAsciiVisibleChar() const;
        size_t  LastIndexOf ( wchar_t ch ) const;
        void    LeftTrim();
        size_t  ReplaceI( LPCWSTR pszOld, LPCWSTR pszNew );
        void    ReplaceChar( wchar_t oldChar, wchar_t newChar );
        size_t  ReplaceChar( wchar_t oldChar, LPCWSTR pszNew );
        void    ReplaceCharAt( size_t index, wchar_t newChar );
        size_t  ReplaceInvalidUTF16( wchar_t newChar ) const;
        size_t  ReplaceWhiteSpaces( wchar_t ch );
        void    Reverse();
        void    RightTrim();
        void    RightTrim( wchar_t wch );
        void    SetAnsi( const char* pszAnsi );
        void    SetAnsiChars( const char* pszAnsi, size_t numChars );
        void    SetCharArray( const CharArray& Chars );
        void    SetCharAt( size_t index, wchar_t ch );
        bool    StartsWith( wchar_t wch ) const;
        bool    StartsWith( const String& Prefix, bool caseSensitive ) const;
        bool    StartsWith( LPCWSTR pszPrefix, bool caseSensitive ) const;
        void    SubString( size_t start, size_t length, String* pSub ) const;
        bool    StartsWithI( LPCWSTR pszPrefix ) const;
        size_t  ToArray( wchar_t ch, StringArray* pTokens ) const;
        size_t  ToArray( const wchar_t* pSeparator, StringArray* pTokens ) const;
        void    ToLowerCase();
        void    ToUpperCase();
        void    ToUSAscii( CharArray* pUSAscii ) const;
        void    Trim();
        void    Truncate( size_t numChars );
        void    Zero();

    protected:
        // Methods

        // Data members

    private:
        // Methods
        void SetMultibyte( const char* pszMB, size_t numChars, UINT codePage );

        // Data members
        size_t   MINIMUM_ALLOCATION;    // Always allocate at least this amount
        size_t   m_uLengthChars;
        size_t   m_uCharsAllocated;
        wchar_t* m_pszString;
};

#endif  // PXSBASE_STRING_H_
