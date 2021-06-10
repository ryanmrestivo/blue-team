///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Char Array Class Header
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

#ifndef PXSBASE_CHAR_ARRAY_H_
#define PXSBASE_CHAR_ARRAY_H_

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

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class CharArray
{
    public:
        // Default constructor
        CharArray();

        // Copy constructor
        CharArray( const CharArray& oCharArray );

        // Assignment operator
        CharArray& operator= ( const CharArray& oCharArray );

        // Destructor
        ~CharArray();

        // Methods
        void    Allocate( size_t numChars );
        void    Append( char ch );
        void    Append( const char* pszString );
        void    Append( const CharArray& Chars );
        void    Append( const char* pChars, size_t numChars );
        char    CharAt( size_t index ) const;
        int     Compare( const CharArray& Input ) const;
        bool    EndsWith( const char* pszSuffix ) const;
        void    Free();
        size_t  Get( size_t idxOffset, char* pBuffer, size_t bufChars ) const;
        size_t  Get( size_t idxOffset, size_t numChars, CharArray* pBuffer ) const;
    const char* GetPtr() const;
        size_t  GetNumAllocated() const;
        size_t  GetSize() const;
        size_t  IndexOf( char ch ) const;
        size_t  IndexOf( char ch, size_t from ) const;
        size_t  IndexOf( const char* pszSearch, bool caseSensitive, size_t from ) const;
        bool    IsNull() const;
        void    LeftShift( size_t count );
        void    LeftTrim();
        void    Set( const char* pszString );
        void    Set( size_t index, char ch );
        void    Truncate( size_t size );
        void    ToString( String* pString ) const;
        void    ToString( size_t idxOffset, size_t numChars, String* pString ) const;
        void    Zero();

    protected:
        // Methods

        // Data members

    private:
        // Methods

        // Data members
        size_t   GROW_BY;           // psuedo-constant
        size_t   m_uSize;
        size_t   m_uAllocated;
        char*    m_pChars;
};

#endif  // PXSBASE_CHAR_ARRAY_H_
