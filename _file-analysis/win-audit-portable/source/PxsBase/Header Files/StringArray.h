///////////////////////////////////////////////////////////////////////////////////////////////////
//
// String Array Class Header
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

#ifndef PXSBASE_STRING_ARRAY_H_
#define PXSBASE_STRING_ARRAY_H_

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
// Defines
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class StringArray
{
    public:
        // Default constructor
        StringArray();

        // Copy constructor
        StringArray( const StringArray& oStringArray );

        // Destructor
        ~StringArray();

       // Assignment operator
        StringArray& operator= ( const StringArray& oStringArray );

        // Methods
        void    Add( LPCWSTR pszString );
        void    Add( const String& Element );
        void    AddArray( const StringArray& Strings );
        bool    AddUniqueI( LPCWSTR pszString );
        int     CompareI( size_t index, LPCWSTR pszString ) const;
        size_t  CountOf( LPCWSTR pszFind, bool caseSensitive ) const;
        LPCWSTR Get( size_t index )const;
        size_t  GetSize() const;
        size_t  IndexOf( LPCWSTR pszFind, bool caseSensitive ) const;
        size_t  IndexOfStartsWithI( LPCWSTR pszPrefix ) const;
        size_t  IndexOfSubStringI( size_t index, LPCWSTR pszFind ) const;
        void    RemoveAll();
        void    Set( size_t index, LPCWSTR pszString );
        void    SetSize( size_t newSize );
        void    Sort( bool ascending );
        bool    StartsWithI( size_t index, wchar_t wch ) const;
        bool    StartsWithI( size_t index, LPCWSTR pszPrefix ) const;
        void    SubString( size_t index, size_t start, size_t length, String* pSub ) const;
        void    ToString( wchar_t wch, String* pString ) const;

    protected:
        // Methods

        // Data members

    private:
        // Methods

        // Data members
        size_t    REALLOC_SIZE;     // Pseudo constant
        size_t    m_uSize;
        size_t    m_uAllocated;
        wchar_t** m_ppStringArray;
};

#endif  // PXSBASE_STRING_ARRAY_H_
