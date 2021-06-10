///////////////////////////////////////////////////////////////////////////////////////////////////
//
// String Array Class Implementation
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
#include "PxsBase/Header Files/StringArray.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
StringArray::StringArray()
            :REALLOC_SIZE( 64 ),  // pseudo-constant
             m_uSize( 0 ),
             m_uAllocated( 0 ),
             m_ppStringArray( nullptr )
{
}

// Copy constructor
StringArray::StringArray( const StringArray& oStringArray )
            :REALLOC_SIZE( 64 ),  // pseudo-constant
             m_uSize( 0 ),
             m_uAllocated( 0 ),
             m_ppStringArray( nullptr )
{
    *this = oStringArray;
}

// Destructor
StringArray::~StringArray()
{
    try
    {
        RemoveAll();
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
StringArray& StringArray::operator=( const StringArray& oStringArray )
{
    // Disallow self-assignment
    if ( this == &oStringArray ) return *this;

    RemoveAll();
    AddArray( oStringArray );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add a string to the end of this array
//
//  Parameters:
//      pszString - string pointer
//
//  Returns:
//      void
//===============================================================================================//
void StringArray::Add( LPCWSTR pszString )
{
    size_t newSize = PXSAddSizeT( m_uSize, 1 );

    SetSize( newSize );
    Set( newSize - 1, pszString );  // zero based
}

//===============================================================================================//
//  Description:
//      Add a string object to this array
//
//  Parameters:
//      Element - string object
//
//  Returns:
//      void
//===============================================================================================//
void StringArray::Add( const String& Element )
{
    Add( Element.c_str() );
}

//===============================================================================================//
//  Description:
//      Add a string array to this array
//
//  Parameters:
//      Strings - string array object
//
//  Remarks:
//      Adding one at a time, if this becomes a performance issue can re-work.
//
//  Returns:
//      void
//===============================================================================================//
void StringArray::AddArray( const StringArray& Strings )
{
    size_t  i    = 0;
    size_t  size = Strings.GetSize();
    LPCWSTR pszString = nullptr;

    // Disallow adding an array to itself
    if ( this == &Strings )
    {
        throw FunctionException( L"this = Strings", __FUNCTION__ );
    }

    for ( i = 0; i < size; i++ )
    {
        pszString = Strings.Get( i );
        Add( pszString );
    }
}

//===============================================================================================//
//  Description:
//      Add a string to this array if it is not already present. Comparison
//      is case insensitive
//
//  Parameters:
//      pszString - pointer to string to add
//
//  Returns:
//      true if added the string, otherwise false
//===============================================================================================//
bool StringArray::AddUniqueI( LPCWSTR pszString )
{
    bool success = false;

    // NULL != NULL
    if ( pszString == nullptr )
    {
        return false;
    }

    if ( IndexOf( pszString, false ) == PXS_MINUS_ONE )
    {
         Add( pszString );    // Add to the end
         success = true;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Compare the string at the specified index with the input string. Test is case insensitive.
//
//  Parameters:
//      index     - zero-based array index of the the reference string
//      pszString - string to compare
//
//  Returns:
//      -ve if array string < pszString
//        0 if array string = pszString ( or NULL == NULL )
//      +ve if array string > pszString
//===============================================================================================//
int StringArray::CompareI( size_t index, LPCWSTR pszString ) const
{
    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    return PXSCompareString( m_ppStringArray[ index ], pszString, false );
}

//===============================================================================================//
//  Description:
//      Count the number of occurrences of the specified string in the array
//
//  Parameters:
//      pszFind - the string to find
//      caseSensitive - true if want a case sensitive search otherwise false
//
//  Returns:
//      count of occurrences
//===============================================================================================//
size_t StringArray::CountOf( LPCWSTR pszFind, bool caseSensitive ) const
{
    size_t i = 0, count = 0;

    if ( ( m_ppStringArray == nullptr ) || ( pszFind == nullptr ) )
    {
        return 0;
    }

    for ( i= 0; i < m_uSize; i++ )
    {
        if ( PXSCompareString( pszFind, m_ppStringArray[ i ], caseSensitive ) == 0 )
        {
            count++;
        }
    }

    return count;
}

//===============================================================================================//
//  Description:
//      Get the string at the specified position in the array
//
//  Parameters:
//      index - the zero-based index
//
//  Returns:
//      Constant string pointer
//===============================================================================================//
LPCWSTR StringArray::Get( size_t index ) const
{
    if ( m_ppStringArray == nullptr )
    {
        throw NullException(  L"m_ppStringArray", __FUNCTION__ );
    }

    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    return m_ppStringArray[ index ];
}

//===============================================================================================//
//  Description:
//      Get the size of the array
//
//  Parameters:
//      None
//
//  Returns:
//      The size of the array
//===============================================================================================//
size_t StringArray::GetSize() const
{
    return m_uSize;
}

//===============================================================================================//
//  Description:
//      Get the index of the specified string
//
//  Parameters:
//      pszFind       - pointer to the string to find
//      caseSensitive - flag to indicate if want a case sensitive search
//
//  Returns:
//      Zero based array index, PXS_MINUS_ONE if not found
//===============================================================================================//
size_t StringArray::IndexOf( LPCWSTR pszFind, bool caseSensitive ) const
{
    size_t idxFound = PXS_MINUS_ONE, i = 0;

    if ( m_ppStringArray == nullptr )
    {
        return PXS_MINUS_ONE;
    }

    while ( ( i < m_uSize ) && ( idxFound == PXS_MINUS_ONE ) )
    {
        if ( PXSCompareString( pszFind, m_ppStringArray[ i ], caseSensitive ) == 0 )
        {
            idxFound = i;
        }
        i++;
    }

    return idxFound;
}


//===============================================================================================//
//  Description:
//      Get the index of the first string that starts with the input string. Comparison is
//      case insensitive.
//
//  Parameters:
//      pszFind - pointer to the string to find
//
//  Returns:
//      Zero based array index, PXS_MINUS_ONE if not found
//===============================================================================================//
size_t StringArray::IndexOfStartsWithI( LPCWSTR pszFind ) const
{
    size_t numChars = 0;
    size_t idxFound = PXS_MINUS_ONE, i = 0;

    if ( ( m_ppStringArray == nullptr ) || ( pszFind == nullptr ) )
    {
        return PXS_MINUS_ONE;
    }
    StringCchLength( pszFind, STRSAFE_MAX_CCH, &numChars );

    while ( ( i < m_uSize ) && ( idxFound == PXS_MINUS_ONE ) )
    {
        if ( PXSCompareStringN( pszFind, m_ppStringArray[ i ], numChars, false ) == 0 )
        {
            idxFound = i;
        }
        i++;
    }

    return idxFound;
}

//===============================================================================================//
//  Description:
//      Search the string at idx for the input string. Comparison is case insensitive.
//
//  Parameters:
//      idx     - zero based index of string to search
//      pszFind - pointer to the string to find
//
//  Returns:
//      Zero based array index, PXS_MINUS_ONE if not found
//===============================================================================================//
size_t StringArray::IndexOfSubStringI( size_t index, LPCWSTR pszFind ) const
{
    size_t   length, numChars, i = 0, found = PXS_MINUS_ONE;
    wchar_t* pszString;

    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    if ( ( m_ppStringArray == nullptr ) || ( pszFind == nullptr ) )
    {
        return PXS_MINUS_ONE;
    }
    pszString = m_ppStringArray[ index ];
    length    = wcslen( pszString );
    numChars  = wcslen( pszFind );

    while ( ( found == PXS_MINUS_ONE ) && ( PXSAddSizeT( i, numChars ) <= length ) )
    {
        if ( PXSCompareStringN( pszFind, pszString + i, numChars, false ) == 0 )
        {
            found = i;
        }
        i++;
    };

    return found;
}

//===============================================================================================//
//  Description:
//      Delete all the array's elements
//
//  Parameters:
//      None
//
//   Remarks:
//      Called by destructor so should not throw
//
//  Returns:
//      void
//===============================================================================================//
void StringArray::RemoveAll()
{
    size_t i = 0;

    if ( m_ppStringArray )
    {
        // Delete the individual strings
        for ( i = 0; i < m_uSize; i++ )
        {
            delete[] m_ppStringArray[ i ];
            m_ppStringArray[ i ] = nullptr;
        }

        // Now the array
        delete m_ppStringArray;
    }
    m_ppStringArray = nullptr;
    m_uSize         = 0;
    m_uAllocated    = 0;
}

//===============================================================================================//
//  Description:
//      Set the string at the specified array index
//
//  Parameters:
//      index     - the array index of the string
//      pszString - string pointer
//
//  Returns:
//      void
//===============================================================================================//
void StringArray::Set( size_t index, LPCWSTR pszString )
{
    size_t   numChars = 0;
    wchar_t* pszNew   = nullptr;

    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    if ( m_ppStringArray == nullptr )
    {
        throw NullException( L"m_ppStringArray", __FUNCTION__ );
    }

    // Replace existing string
    if ( pszString )
    {
        numChars = wcslen( pszString );
        numChars = PXSAddSizeT( numChars, 1 );       // Null terminator
        pszNew  = new wchar_t[ numChars ];
        if ( pszNew == nullptr )
        {
            throw MemoryException( __FUNCTION__ );
        }
        wmemset( pszNew, 0, numChars );
        StringCchCopy( pszNew, numChars, pszString );
    }

    if ( m_ppStringArray[ index ] )
    {
        delete[] m_ppStringArray[ index ];
    }
    m_ppStringArray[ index ] = pszNew;
}

//===============================================================================================//
//  Description:
//      Set the the size of the array
//
//  Parameters:
//      newSize - the new size
//
//  Returns:
//      void
//===============================================================================================//
void StringArray::SetSize( size_t newSize )
{
    size_t    i = 0, allocate;
    wchar_t** ppNew = nullptr;

    // Special case
    if ( newSize == 0 )
    {
        RemoveAll();
        return;
    }

    // Handle case of changing the size when have already allocated a
    // sufficient number of elements
    if ( newSize <= m_uAllocated )
    {
        for ( i = newSize; i < m_uAllocated; i++ )
        {
            // Recover memory of any allocated strings
            if ( m_ppStringArray[ i ] )
            {
                delete[] m_ppStringArray[ i ];
                m_ppStringArray[ i ] = nullptr;   // Reset
            }
        }
        m_uSize = newSize;
        return;
    }

    // Grow by REALLOC_SIZE
    allocate = newSize;
    if ( ( allocate % REALLOC_SIZE ) )
    {
        allocate = PXSAddSizeT( allocate, REALLOC_SIZE - ( allocate % REALLOC_SIZE ) );
    }

    // Need to grow the array
    ppNew = new wchar_t*[ allocate ];
    if ( ppNew == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( ppNew, 0, sizeof ( wchar_t* ) * allocate );

    // Replace, reserve any existing elements
    if ( m_ppStringArray )
    {
        memcpy( ppNew, m_ppStringArray, sizeof ( wchar_t* ) * m_uAllocated );
        delete[] m_ppStringArray;
    }
    m_ppStringArray = ppNew;
    m_uAllocated    = allocate;
    m_uSize         = newSize;
}

//===============================================================================================//
//  Description:
//      Sort this array of strings
//
//  Parameters:
//      ascending - true for an ascending sort otherwise false for descending
//
//  Returns:
//      void
//===============================================================================================//
void StringArray::Sort( bool ascending )
{
    size_t   i = 0, sizeT = 0, maxChars = 0, idxOffset = 0, newChars = 0;
    wchar_t* pszStrings = nullptr;
    LPCWSTR  psz = nullptr;
    AllocateWChars  AllocWChars;

    sizeT = GetSize();
    if ( sizeT <= 1 )
    {
        return;       // Nothing to do
    }

    // Get maximum characters in any string
    for ( i = 0; i < sizeT; i++ )
    {
        psz = Get( i );
        if ( psz )
        {
            maxChars = PXSMaxSizeT( maxChars, wcslen( psz ) );
        }
    }

    if ( maxChars == 0 )
    {
        return;     // Nothing to do
    }
    maxChars = PXSAddSizeT( maxChars, 1 );    // Terminator

    // Allocate a fixed-width array of strings and fill it
    newChars   = PXSMultiplySizeT( sizeT, maxChars );
    pszStrings = AllocWChars.New( newChars );
    for ( i = 0; i < sizeT; i++ )
    {
        psz = Get( i );
        if ( psz )
        {
            idxOffset = PXSMultiplySizeT( i, maxChars );
            StringCchCopy( pszStrings + idxOffset, maxChars, psz );
        }
    }

    // Use qsort
    if ( ascending )
    {
        qsort( pszStrings,
               sizeT,
               PXSMultiplySizeT( maxChars, sizeof ( wchar_t ) ),
               PXSQSortStringAscending );
    }
    else
    {
        qsort( pszStrings,
               sizeT,
               PXSMultiplySizeT( maxChars, sizeof ( wchar_t ) ),
               PXSQSortStringDescending );
    }

    // Re-populate the input/output array, this is not exception safe
    RemoveAll();
    for ( i = 0; i < sizeT; i++ )
    {
        idxOffset = PXSMultiplySizeT( i, maxChars );
        Add( pszStrings + idxOffset );
    }
}

//===============================================================================================//
//  Description:
//      Determine if the string at the specified index starts with the input character.
//      The test is case insensitive.
//
//  Parameters:
//      index     - the zero-based array index
//      pszString - the prefix to test for
//
//  Returns:
//      true if the array string starts with the character otherwise false
//===============================================================================================//
bool StringArray::StartsWithI( size_t index, wchar_t wch ) const
{
    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    if ( ( m_ppStringArray == nullptr ) || ( wch == PXS_CHAR_NULL ) )
    {
        return false;
    }

    if ( m_ppStringArray[ index ] && ( *m_ppStringArray[ index ] == wch ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the string at the specified index starts with the input string.
//      The test is case insensitive.
//
//  Parameters:
//      index     - the zero-based array index
//      pszString - the prefix to test for
//
//  Returns:
//      true if the array string starts with the prefix otherwise false
//===============================================================================================//
bool StringArray::StartsWithI( size_t index, LPCWSTR pszPrefix ) const
{
    size_t numChars;

    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    if ( ( m_ppStringArray == nullptr ) || ( pszPrefix == nullptr ) )
    {
        return false;
    }

    numChars = wcslen( pszPrefix );
    if ( ( numChars == 0 ) || ( numChars > wcslen( m_ppStringArray[ index ] ) ) )
    {
        return false;
    }

    if ( PXSCompareStringN( m_ppStringArray[ index ], pszPrefix, numChars, false ) == 0 )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Get a substring from the string at the specified index
//
//  Parameters:
//      index  - the zero based index of the string from which to extract the substring
//      start  - zero-based start position
//      length - number of characters
//      pSub   - string object to receive new string
//
//  Remarks:
//      If length runs beyond this string, no error is raised. The entire
//      amount beyond idxStart is put into the output. This avoids the
//      caller having to determine this string's length.
//
//  Returns:
//      void, zero length string "" if arguments out of range
//===============================================================================================//
void StringArray::SubString( size_t index, size_t start, size_t length, String* pSub ) const
{
    size_t lengthChars;

    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    if ( pSub == nullptr )
    {
        throw ParameterException( L"pSub", __FUNCTION__ );
    }
    pSub->Zero();

    if ( ( m_ppStringArray == nullptr ) || ( m_ppStringArray[ index ] == nullptr ) )
    {
        return;   // Nothing to do
    }

    // Bounds check
    lengthChars = wcslen( m_ppStringArray[ index ] );
    if ( start < lengthChars )
    {
        length = PXSMinSizeT( lengthChars - start, length );
        pSub->AppendChars( m_ppStringArray[ index ] + start, length );
    }
}

//===============================================================================================//
//  Description:
//      Get this string array as a token separated string
//
//  Parameters:
//      wch     - the characer separator
//      pString - receives the token separated string
//
//  Returns:
//      void
//===============================================================================================//
void StringArray::ToString( wchar_t wch, String* pString ) const
{
    size_t i = 0, allocate = 0;

    if ( pString == nullptr )
    {
        throw ParameterException( L"pString", __FUNCTION__ );
    }
    pString->Zero();

    if ( ( m_ppStringArray == nullptr ) || ( m_uSize == 0 ) )
    {
        return;     // Nothing to do
    }

    // Pass 1 to allocate
    for ( i = 0; i < m_uSize; i++ )
    {
        allocate = PXSAddSizeT( allocate, wcslen( m_ppStringArray[ i ] ) );
        allocate = PXSAddSizeT( allocate, 1 );      // the char separator
    }
    allocate = PXSAddSizeT( allocate, 1 );          // the NULL terminator

    // Pass 2 to concat
    for ( i = 0; i < m_uSize; i++ )
    {
        pString->Append( m_ppStringArray[ i ] );
        if ( i < ( m_uSize - 1 ) )
        {
            pString->AppendChar( wch );
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
