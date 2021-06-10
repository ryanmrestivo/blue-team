///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Char Array Class Implementation
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
#include "PxsBase/Header Files/CharArray.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
CharArray::CharArray()
          :GROW_BY( 512 ),
           m_uSize( 0 ),
           m_uAllocated( 0 ),
           m_pChars( nullptr )
{
}

// Copy constructor
CharArray::CharArray( const CharArray& oCharArray )
          :GROW_BY( 512 ),
           m_uSize( 0 ),
           m_uAllocated( 0 ),
           m_pChars( nullptr )
{
    *this = oCharArray;
}

// Destructor
CharArray::~CharArray()
{
    try
    {
        Free();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
CharArray& CharArray::operator=( const CharArray& oCharArray )
{
    // Disallow self-assignment
    if ( this == &oCharArray ) return *this;

    Zero();
    Append( oCharArray );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Allocate chars for the array
//
//  Parameters:
//      numChars - the number of chars to allocate
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Allocate( size_t numChars )
{
    char*  pChars;
    size_t newSize;

    if ( m_uAllocated == numChars )
    {
        return;     // Nothing to do
    }

    // Special case of freeing all the chars
    if ( numChars == 0 )
    {
        Free();
        return;
    }

    pChars = new char[ numChars ];
    if ( pChars == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( pChars, 0, numChars );

    // Copy in as much as can fit
    newSize = PXSMinSizeT( numChars, m_uSize );
    if ( m_pChars )
    {
        memcpy( pChars, m_pChars, newSize );
    }

    // Replace
    if ( m_pChars )
    {
        delete [] m_pChars;
    }
    m_pChars     = pChars;
    m_uSize      = newSize;
    m_uAllocated = numChars;
}

//===============================================================================================//
//  Description:
//      Append the specified character
//
//  Parameters:
//      ch - the char to append
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Append( char ch )
{
    Append( &ch, 1 );
}

//===============================================================================================//
//  Description:
//      Append the specified NULL terminated string
//
//  Parameters:
//      pszString - the string to append
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Append( const char* pszString )
{
    size_t cch = 0;

    if ( pszString == nullptr )
    {
        return;     // Nothing to do
    }
    StringCchLengthA( pszString, STRSAFE_MAX_CCH, &cch );
    Append( pszString, cch );
}

//===============================================================================================//
//  Description:
//      Append an array to this one
//
//  Parameters:
//      Chars - the chars to append
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Append( const CharArray& Chars )
{
    Append( Chars.m_pChars, Chars.m_uSize );
}

//===============================================================================================//
//  Description:
//      Append chars to this array
//
//  Parameters:
//      pBuffer  - pointer to the the chars
//      numChars - the number of chars
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Append( const char* pBuffer, size_t numChars )
{
    char*  pChars;
    size_t newSize, allocate;

    if ( ( pBuffer == nullptr ) || ( numChars == 0 ) )
    {
        return;     // Nothing to do
    }
    newSize = PXSAddSizeT( m_uSize, numChars );

    // Test if fits into already allocated chars
    if ( m_pChars && ( newSize < m_uAllocated ) )
    {
        memcpy( m_pChars + m_uSize, pBuffer, numChars );
        m_uSize = newSize;
        return;
    }

    // Round up the allocation
    allocate = newSize;
    if ( allocate % GROW_BY )
    {
        allocate = PXSMultiplySizeT( allocate / GROW_BY, GROW_BY );
        allocate = PXSAddSizeT( allocate, GROW_BY );
    }

    // Allocate a bigger array
    pChars = new char[ allocate ];
    if ( pChars == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    if ( m_pChars )
    {
        memcpy( pChars, m_pChars, m_uSize );
    }
    memcpy( pChars + m_uSize, pBuffer, numChars );
    memset( pChars + newSize, 0, allocate - newSize );

    // Replace
    delete [] m_pChars;
    m_pChars     = pChars;
    m_uSize      = newSize;
    m_uAllocated = allocate;
}

//===============================================================================================//
//  Description:
//      Get the character at the specified index
//
//  Parameters:
//      index - the zero-based position of the character
//
//  Returns:
//      void
//===============================================================================================//
char CharArray::CharAt( size_t index ) const
{
    if ( ( m_pChars == nullptr ) || ( index >= m_uSize ) )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    return m_pChars[ index ];
}

//===============================================================================================//
//  Description:
//      Compare this character array with the input array. Test is case sensitive.
//
//  Parameters:
//      Input - the input
//
//  Returns:
//        0 if the arrays are identical
//      < 0 if this array is less than Input
//      > 0 if this array is greater than Input
//===============================================================================================//
int CharArray::Compare( const CharArray& Input ) const
{
    int    result = 0;
    size_t count, inputSize = Input.GetSize();
    const char*   pszInput  = Input.GetPtr();

    if ( ( m_pChars == nullptr ) && ( pszInput == nullptr ) )
    {
        return -1;  // NULL != NULL
    }
    else if ( ( m_pChars == nullptr ) && pszInput )
    {
        return 1;
    }
    else if ( m_pChars && ( pszInput == nullptr ) )
    {
        return -1;
    }

    // Both non-null
    count  = PXSMinSizeT( inputSize, m_uSize );
    result = memcmp( m_pChars, pszInput, count );
    if ( result )
    {
        return result;
    }

    if ( inputSize == m_uSize )
    {
        return 0;   // Both strings of same length
    }

    if ( inputSize > m_uSize )
    {
        return 1;
    }

    return -1;
}

//===============================================================================================//
//  Description:
//      Determine if this character array ends with the specified string. Test is case sensitive.
//
//  Parameters:
//      pszSuffix - the suffix to test for string
//
//  Returns:
//      true if this string ends with the suffix otherwise false
//===============================================================================================//
bool CharArray::EndsWith( const char* pszSuffix ) const
{
    size_t lenSuffix = 0;

    if ( ( m_pChars == nullptr ) || ( pszSuffix == nullptr ) )
    {
        return false;
    }

    StringCchLengthA( pszSuffix , STRSAFE_MAX_CCH, &lenSuffix );
    if ( lenSuffix > m_uSize )
    {
        return false;
    }

    if ( memcmp( pszSuffix, m_pChars + m_uSize - lenSuffix, lenSuffix ) )
    {
        return false;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Free this array's chars
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Free()
{
    if ( m_pChars )
    {
        delete [] m_pChars;
    }
    m_pChars     = nullptr;
    m_uSize      = 0;
    m_uAllocated = 0;
}

//===============================================================================================//
//  Description:
//      Get as many chars from the array as can fit into the input buffer. If
//      the starting index is beyond the end of the buffer, no error is raised
//      with return value set to zero.
//
//  Parameters:
//      idxOffset - zero-based index from which to begin
//      pBuffer   - buffer to receive the chars
//      bufChars  - size of the buffer
//
//  Returns:
//      number of chars copied into the buffer, can be zero.
//===============================================================================================//
size_t CharArray::Get( size_t idxOffset, char* pBuffer, size_t bufChars ) const
{
    size_t numCopied;

    if ( pBuffer == nullptr )
    {
        throw ParameterException( L"pBuffer", __FUNCTION__ );
    }

    if ( ( bufChars  == 0 )       ||    // Nothing to do
         ( idxOffset >= m_uSize ) ||
         ( m_pChars  == nullptr )  )    // No source buffer
    {
        return 0;
    }
    numCopied = m_uSize - idxOffset;
    numCopied = PXSMinSizeT( numCopied, bufChars );
    memcpy( pBuffer, m_pChars + idxOffset, numCopied );

    return numCopied;
}

//===============================================================================================//
//  Description:
//      Get chars from this array. Copies up to the end of the buffer if requested more than
//      avialble but no error is raised.
//
//  Parameters:
//      idxOffset - zero-based index from which to begin
//      numChars  - the number of chars to get
//      pBuffer   - buffer to receive the chars
//
//  Returns:
//      number of chars copied into the buffer, can be zero.
//===============================================================================================//
size_t CharArray::Get( size_t idxOffset, size_t numChars, CharArray* pBuffer ) const
{
    size_t numCopied;

    if ( m_pChars == nullptr )
    {
        throw NullException( L"m_pChars", __FUNCTION__ );
    }

    if ( idxOffset >= m_uSize )
    {
        throw BoundsException( L"idxOffset >= m_uSize", __FUNCTION__ );
    }

    if ( pBuffer == nullptr )
    {
        throw ParameterException( L"pBuffer", __FUNCTION__ );
    }
    pBuffer->Zero();

    numCopied = m_uSize - idxOffset;
    numCopied = PXSMinSizeT( numCopied, numChars );
    pBuffer->Append( m_pChars + idxOffset, numCopied );

    return numCopied;
}

//===============================================================================================//
//  Description:
//      Get the pointer to char buffer
//
//  Parameters:
//      None
//
//  Returns:
//      const char*
//===============================================================================================//
const char* CharArray::GetPtr() const
{
    return m_pChars;
}

//===============================================================================================//
//  Description:
//      Get the size of the array
//
//  Parameters:
//      None
//
//  Returns:
//      number of chars in the array
//===============================================================================================//
size_t CharArray::GetSize() const
{
    return m_uSize;
}

//===============================================================================================//
//  Description:
//      Get the number of chars allocated
//
//  Parameters:
//      None
//
//  Returns:
//      size_t
//===============================================================================================//
size_t CharArray::GetNumAllocated() const
{
    return m_uAllocated;
}

//===============================================================================================//
//  Description:
//      Get the zero-based index of the first occurrence of the specified character. The test is
//      case sensitive
//
//  Parameters:
//      ch - the character to find
//
//  Returns:
//      zero-based index, -1 if not found
//===============================================================================================//
size_t CharArray::IndexOf( char ch ) const
{
    return IndexOf( ch, 0 );
}

//===============================================================================================//
//  Description:
//      Get the zero-based index of the first occurrence of the specified character starting
//      from the specified offset
//      case sensitive
//
//  Parameters:
//      ch - the character to find
//      from - zero-based index to start the search from
//
//  Returns:
//      zero-based index, -1 if not found
//===============================================================================================//
size_t CharArray::IndexOf( char ch, size_t from ) const
{
    size_t i = from, idx = PXS_MINUS_ONE;

    if ( ( m_pChars == nullptr ) || ( i >= m_uSize ) )
    {
        return PXS_MINUS_ONE;
    }

    do
    {
        if ( m_pChars[ i ] == ch )
        {
            idx = i;
        }
        i++;
    } while ( ( idx == PXS_MINUS_ONE ) && ( i < m_uSize ) );

    return idx;
}

//===============================================================================================//
//  Description:
//      Get the zero-based index of the first occurrence of the specified search string.
//      The test is case sensitive
//
//  Parameters:
//      pszSearch     - the string to find
//      caseSensitive - want a case sensitive search
//      from          - the index to start searching from
//
//  Returns:
//      zero-based index, -1 if not found
//===============================================================================================//
size_t CharArray::IndexOf( const char* pszSearch, bool caseSensitive, size_t from ) const
{
    int nCmp = 0;
    size_t idxFound = PXS_MINUS_ONE, lenSearch;

    if ( ( m_pChars == nullptr ) || ( pszSearch == nullptr ) )
    {
        return PXS_MINUS_ONE;
    }

    lenSearch = strlen( pszSearch );
    while ( ( PXSAddSizeT( from, lenSearch ) <= m_uSize ) && ( idxFound == PXS_MINUS_ONE ) )
    {
        nCmp = 0;
        if ( caseSensitive )
        {
            nCmp = memcmp( m_pChars + from, pszSearch, lenSearch );
        }
        else
        {
            nCmp = _strnicmp( m_pChars + from, pszSearch, lenSearch );
        }

        if ( nCmp == 0 )
        {
            idxFound = from;
        }
        from++;
    }

    return idxFound;
}

//===============================================================================================//
//  Description:
//      Determine if the array is NULL
//
//  Parameters:
//      None
//
//  Returns:
//      true if null otherwise false
//===============================================================================================//
bool CharArray::IsNull() const
{
    if ( m_pChars )
    {
        return false;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Shift the chars to the left
//
//  Parameters:
//      count - the number to shift to the left. If exceeds the array size will limit.
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::LeftShift( size_t count )
{
    if ( ( m_pChars == nullptr ) || ( m_uSize == 0 ) || ( count == 0 ) )
    {
        return;     // Nothing to do
    }

    // Limit
    if ( count > m_uSize )
    {
        count = m_uSize;
    }
    memmove( m_pChars, m_pChars + count, m_uSize - count );
    memset( m_pChars + ( m_uSize - count ), 0, count );
    m_uSize -= count;
}

//===============================================================================================//
//  Description:
//      Trim leading white spaces.
//
//  Parameters:
//      nome
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::LeftTrim()
{
    char   ch;
    size_t shift = 0;

    if ( ( m_pChars == nullptr ) || ( m_uSize == 0 ) )
    {
        return;     // Nothing to do
    }

    // Determine how far to shift
    ch = *m_pChars;
    while ( ch && PXSIsWhiteSpaceA( ch ) )
    {
        shift++;
        ch = m_pChars[ shift ];
    }
    LeftShift( shift );
}

//===============================================================================================//
//  Description:
//      Set the characters of this array
//
//  Parameters:
//      pszString - the input string
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Set( const char* pszString )
{
    Zero();
    Append( pszString );
}

//===============================================================================================//
//  Description:
//      Set the character at the specified index
//
//  Parameters:
//      index - zero based index
//      ch    - the character
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Set( size_t index, char ch )
{
    if ( m_pChars == nullptr )
    {
        throw FunctionException( L"m_pChars", __FUNCTION__ );
    }

    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }
    m_pChars[ index ] = ch;
}

//===============================================================================================//
//  Description:
//      Truncate the size of this array to the specified value
//
//  Parameters:
//      size - the new size, no effect if size is larger than the current size.
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Truncate( size_t size )
{
    if ( ( m_pChars == nullptr ) || ( size >= m_uSize ) )
    {
        return;     // Nothing to do
    }
    memset( m_pChars + size, 0, m_uSize - size );
    m_uSize = size;
}

//===============================================================================================//
//  Description:
//      Convert the array to a string, assumes there are only ANSI characcters
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::ToString( String* pString ) const
{
    ToString( 0, m_uSize, pString );
}

//===============================================================================================//
//  Description:
//      Convert part of the array to a string, assumes there are only ANSI characcters
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::ToString( size_t idxOffset, size_t numChars, String* pString ) const
{
    if ( pString == nullptr )
    {
        throw NullException( L"pString", __FUNCTION__ );
    }

    if ( PXSAddSizeT( idxOffset, numChars ) > m_uSize )
    {
        throw BoundsException( L"idxOffset + numChars", __FUNCTION__ );
    }
    pString->SetAnsiChars( m_pChars + idxOffset, numChars );
}

//===============================================================================================//
//  Description:
//      Reset this class, zero any allocated memory and set the size to zero
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void CharArray::Zero()
{
    if ( m_pChars )
    {
        memset( m_pChars, 0, m_uSize );
    }
    m_uSize = 0;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
