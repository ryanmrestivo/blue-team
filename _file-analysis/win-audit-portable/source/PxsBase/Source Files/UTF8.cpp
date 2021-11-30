///////////////////////////////////////////////////////////////////////////////////////////////////
//
// UTF-8 String Class Implementation
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
#include "PxsBase/Header Files/utf8.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/MemoryException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
UTF8::UTF8()
     :m_uByteLength( 0 ),
      m_pszUtf8( nullptr )
{
}

// Copy constructor
UTF8::UTF8( const UTF8& oUtf8 )
     :m_uByteLength( 0 ),
      m_pszUtf8( nullptr )
{
    *this = oUtf8;
}

// Destructor
UTF8::~UTF8()
{
    if ( m_pszUtf8 )
    {
        delete[] m_pszUtf8;
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation
UTF8& UTF8::operator= ( const UTF8& oUtf8 )
{
    char*   pszNew;
    size_t  numBytes = 0;
    HRESULT hResult;

    if ( this == &oUtf8 ) return *this;

    if ( oUtf8.m_pszUtf8 == nullptr )
    {
        if ( m_pszUtf8 )
        {
            delete [] m_pszUtf8;
        }
        m_uByteLength = 0;

        return *this;
    }

    // Copy
    hResult = StringCchLengthA( oUtf8.m_pszUtf8, STRSAFE_MAX_CCH, &numBytes );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"StringCchLengthA", __FUNCTION__ );
    }
    numBytes = PXSAddSizeT( numBytes, 1 );      // NULL
    pszNew   = new char[ numBytes ];
    if ( pszNew == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    StringCchCopyA( pszNew, numBytes, oUtf8.m_pszUtf8 );

    // Replace
    if ( m_pszUtf8 )
    {
        delete [] m_pszUtf8;
    }
    m_pszUtf8 = pszNew;
    hResult = StringCchLengthA( m_pszUtf8, STRSAFE_MAX_CCH, &m_uByteLength );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"StringCchLengthA", __FUNCTION__ );
    }
    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the char at the specified index
//
//  Parameters:
//      none
//
//  Returns:
//      the single byte char
//===============================================================================================//
char UTF8::GetAt( size_t index ) const
{
    if ( index >= m_uByteLength )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    if ( m_pszUtf8 == nullptr )
    {
        return 0;   // Should not get here
    }

    return m_pszUtf8[ index ];
}

//===============================================================================================//
//  Description:
//      Get the number of bytes in the UTF-8 string not including the NULL
//      terminator
//
//  Parameters:
//      none
//
//  Returns:
//      number of bytes
//===============================================================================================//
size_t UTF8::GetByteLength() const
{
    return m_uByteLength;
}

//===============================================================================================//
//  Description:
//      Set the UTF-8 string using the specified input string not including
//      the NULL terminator
//
//  Parameters:
//      pszString - the input string
//
//  Returns:
//      void
//===============================================================================================//
void UTF8::Set( LPCWSTR pszString )
{
    char*     pszNew   = nullptr;
    size_t    numBytes = 0;
    HRESULT   hResult;
    Formatter Format;

    if ( pszString == nullptr )
    {
        // Reset
        delete[] m_pszUtf8;
        m_uByteLength = 0;
        m_pszUtf8     = nullptr;
        return;
    }

    // Allocate bytes for the UTF8 string, 3 bytes per
    // character + a null terminator
    numBytes = wcslen( pszString );
    numBytes = PXSMultiplySizeT( numBytes, 3 );
    numBytes = PXSAddSizeT( numBytes, 1 );       // Null terminator
    pszNew  = new char[ numBytes ];
    if ( pszNew == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( pszNew, 0, numBytes );

    try
    {
        Format.WideToUTF8( pszString, pszNew, numBytes );
        pszNew[ numBytes - 1 ] = PXS_CHAR_NULL;
    }
    catch ( const Exception& )
    {
        delete[] pszNew;
        throw;
    }

    // Replace
    if ( m_pszUtf8 )
    {
        delete[] m_pszUtf8;
        m_pszUtf8     = nullptr;
        m_uByteLength = 0;
    }
    m_pszUtf8 = pszNew;
    hResult = StringCchLengthA( m_pszUtf8, STRSAFE_MAX_CCH, &m_uByteLength );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"StringCchLengthA", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Return a pointer to the UTF8 string
//
//  Parameters:
//      None
//
//  Returns:
//      pointer to the UTF8 string
//===============================================================================================//
const char* UTF8::u_str() const
{
    return m_pszUtf8;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
