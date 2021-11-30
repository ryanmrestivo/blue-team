///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Byte Array Class Implementation
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
#include "PxsBase/Header Files/ByteArray.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/ParameterException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ByteArray::ByteArray()
          :GROW_BY( 512 ),
           m_uSize( 0 ),
           m_uAllocated( 0 ),
           m_pBytes( nullptr )
{
}

// Copy constructor
ByteArray::ByteArray( const ByteArray& oByteArray )
          :GROW_BY( 512 ),
           m_uSize( 0 ),
           m_uAllocated( 0 ),
           m_pBytes( nullptr )
{
    *this = oByteArray;
}

// Destructor
ByteArray::~ByteArray()
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
ByteArray& ByteArray::operator=( const ByteArray& oByteArray )
{
    // Disallow self-assignment
    if ( this == &oByteArray ) return *this;

    Zero();
    Append( oByteArray );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Allocate bytes for the array
//
//  Parameters:
//      numBytes - the number of bytes to allocate
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::Allocate( size_t numBytes )
{
    BYTE*  pBytes;
    size_t newSize;

    if ( m_uAllocated == numBytes )
    {
        return;     // Nothing to do
    }

    // Special case of freeing all the bytes
    if ( numBytes == 0 )
    {
        Free();
        return;
    }

    pBytes = new BYTE[ numBytes ];
    if ( pBytes == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( pBytes, 0, numBytes );

    // Copy in as much as can fit
    newSize = PXSMinSizeT( numBytes, m_uSize );
    if ( m_pBytes )
    {
        memcpy( pBytes, m_pBytes, newSize );
    }

    // Replace
    if ( m_pBytes )
    {
        delete [] m_pBytes;
    }
    m_pBytes     = pBytes;
    m_uSize      = newSize;
    m_uAllocated = numBytes;
}

//===============================================================================================//
//  Description:
//      Append an array to this one
//
//  Parameters:
//      Bytes - the bytes to append
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::Append( const ByteArray& Bytes )
{
    Append( Bytes.m_pBytes, Bytes.m_uSize );
}

//===============================================================================================//
//  Description:
//      Append bytes to this array
//
//  Parameters:
//      pBuffer  - pointer to the the bytes
//      numBytes - the number of bytes
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::Append( const BYTE* pBuffer, size_t numBytes )
{
    BYTE*  pBytes;
    size_t newSize, allocate;

    if ( ( pBuffer == nullptr ) || ( numBytes == 0 ) )
    {
        return;     // Nothing to do
    }
    newSize = PXSAddSizeT( m_uSize, numBytes );

    // Test if fits into already allocated bytes
    if ( m_pBytes && ( newSize < m_uAllocated ) )
    {
        memcpy( m_pBytes + m_uSize, pBuffer, numBytes );
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
    pBytes = new BYTE[ allocate ];
    if ( pBytes == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    if ( m_pBytes )
    {
        memcpy( pBytes, m_pBytes, m_uSize );
    }
    memcpy( pBytes + m_uSize, pBuffer , numBytes );
    memset( pBytes + newSize, 0, allocate - newSize );

    // Replace
    delete [] m_pBytes;
    m_pBytes     = pBytes;
    m_uSize      = newSize;
    m_uAllocated = allocate;
}

//===============================================================================================//
//  Description:
//      Append the specified byte to this array
//
//  Parameters:
//      byte - the byte
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::AppendByte( BYTE byte )
{
    Append( &byte, 1 );
}

//===============================================================================================//
//  Description:
//      Determine if this array begins with the input bytes
//
//  Parameters:
//      pBuffer  - the input bytes
//      numBytes - the number of input bytes
//
//  Returns:
//      true if this array begins with the input bytes
//===============================================================================================//
bool ByteArray::BeginsWith( const BYTE* pBuffer, size_t numBytes ) const
{
    if ( ( m_pBytes == nullptr ) || ( pBuffer == nullptr ) )
    {
        return false;      // NULL != NULL
    }

    if ( numBytes > m_uSize )
    {
        return false;
    }

    if ( memcmp( m_pBytes, pBuffer, numBytes ) )
    {
        return false;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Compare bytes in this array to the input bytes
//
//  Parameters:
//      offset   - zero-based offset in this byte array to start the comparison
//      pBuffer  - the input bytes
//      numBytes - the number of input bytes
//
//  Returns:
//        0 if the arrays are identical
//      < 0 if this array is less than Input
//      > 0 if this array is greater than Input
//===============================================================================================//
int ByteArray::Compare( size_t offset, const BYTE* pBuffer, size_t numBytes ) const
{
    if ( ( m_pBytes == nullptr ) || ( pBuffer == nullptr ) )
    {
        return -1;      // NULL != NULL
    }

    if ( PXSAddSizeT( offset, numBytes ) > m_uSize )
    {
        throw BoundsException( L"offset", __FUNCTION__ );
    }

    return memcmp( m_pBytes + offset, pBuffer, numBytes );
}

//===============================================================================================//
//  Description:
//      Free this array's bytes
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::Free()
{
    if ( m_pBytes )
    {
        delete [] m_pBytes;
    }
    m_pBytes     = nullptr;
    m_uSize      = 0;
    m_uAllocated = 0;
}

//===============================================================================================//
//  Description:
//      Get the byte at the specified zero-based index
//
//  Parameters:
//      idx - zero-based index
//
//  Returns:
//      BTE
//===============================================================================================//
BYTE ByteArray::Get( size_t idx ) const
{

    if ( m_pBytes == nullptr )
    {
        throw FunctionException( L"m_pBytes", __FUNCTION__ );
    }

    if ( idx >= m_uSize )
    {
        throw BoundsException( L"idx >= m_uSize", __FUNCTION__ );
    }

    return m_pBytes[ idx ];
}

//===============================================================================================//
//  Description:
//      Get as many bytes from the array as can fit into the input buffer. If
//      the starting index is beyond the end of the buffer, no error is raised
//      with return value set to zero.
//
//  Parameters:
//      idxOffset - zero-based index from which to begin
//      pBuffer   - buffer to receive the bytes
//      bufBytes  - size of the buffer
//
//  Returns:
//      number of bytes copied into the buffer, can be zero.
//===============================================================================================//
size_t ByteArray::Get( size_t idxOffset, BYTE* pBuffer, size_t bufBytes ) const
{
    size_t numCopied;

    if ( pBuffer == nullptr )
    {
        throw ParameterException( L"pBuffer", __FUNCTION__ );
    }

    if ( ( bufBytes  == 0 )       ||    // Nothing to do
         ( idxOffset >= m_uSize ) ||
         ( m_pBytes  == nullptr )  )    // No source buffer
    {
        return 0;
    }
    numCopied = m_uSize - idxOffset;
    numCopied = PXSMinSizeT( numCopied, bufBytes );
    memcpy( pBuffer, m_pBytes + idxOffset, numCopied );

    return numCopied;
}

//===============================================================================================//
//  Description:
//      Get the bytes starting at the offset.
//
//  Parameters:
//      idxOffset - zero-based index from which to begin
//      Bytes     - receives the bytes
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::Get( size_t idxOffset, ByteArray* pBytes ) const
{
    if ( pBytes == nullptr )
    {
        throw ParameterException( L"pBytes", __FUNCTION__ );
    }
    pBytes->Free();

    //             >= as cannot get zero bytes using an index offset
    if ( idxOffset >= m_uSize )
    {
        throw BoundsException( L"idxOffset", __FUNCTION__ );
    }
    pBytes->Append( m_pBytes + idxOffset, m_uSize - idxOffset );
}

//===============================================================================================//
//  Description:
//      Get as many bytes from the array as can fit into the input buffer. If
//      the starting index is beyond the end of the buffer, no error is raised
//      with return value set to zero.
//
//  Parameters:
//      idxOffset - zero-based index from which to begin
//      numBytes  - number of bytes to get
//      pBytes    - buffer to receive the bytes
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::Get( size_t idxOffset, size_t numBytes, ByteArray* pBytes ) const
{
    size_t end = PXSAddSizeT( idxOffset, numBytes );

    if ( pBytes == nullptr )
    {
        throw ParameterException( L"pBytes", __FUNCTION__ );
    }
    pBytes->Zero();

    if ( end > m_uSize )
    {
        throw BoundsException( L"end", __FUNCTION__ );
    }
    pBytes->Append( m_pBytes + idxOffset, numBytes );
}

//===============================================================================================//
//  Description:
//      Get the number of bytes allocated
//
//  Parameters:
//      None
//
//  Returns:
//      size_t
//===============================================================================================//
size_t ByteArray::GetNumAllocated() const
{
    return m_uAllocated;
}

//===============================================================================================//
//  Description:
//      Get the pointer to the bytes
//
//  Parameters:
//      None
//
//  Returns:
//      BYTE*
//===============================================================================================//
const BYTE* ByteArray::GetPtr() const
{
    return m_pBytes;
}

//===============================================================================================//
//  Description:
//      Get the size of the array
//
//  Parameters:
//      None
//
//  Returns:
//      number of bytes in the array
//===============================================================================================//
size_t ByteArray::GetSize() const
{
    return m_uSize;
}

//===============================================================================================//
//  Description:
//      Get the zero-based index of the first occurrence of the specified bytes.
//      The test is case sensitive
//
//  Parameters:
//      idxOffset - zero based index from which to begin the search
//      pszSearch - the string to find
//      numBytes  - bytes length of pszSearch
//
//  Returns:
//      zero-based index, -1 if not found
//===============================================================================================//
size_t ByteArray::IndexOf( size_t idxOffset, const BYTE* pSearch, size_t numBytes ) const
{
    size_t idxFound = PXS_MINUS_ONE;

    if ( ( m_pBytes == nullptr ) ||
         ( pSearch  == nullptr ) ||
         ( numBytes == 0       ) ||
         ( PXSAddSizeT( idxOffset, numBytes ) > m_uSize ) )
    {
        return PXS_MINUS_ONE;
    }

    while ( ( PXSAddSizeT( idxOffset, numBytes ) <= m_uSize ) && ( idxFound == PXS_MINUS_ONE ) )
    {
        if ( memcmp( m_pBytes + idxOffset, pSearch, numBytes ) == 0 )
        {
            idxFound = idxOffset;
        }
        idxOffset++;
    }

    return idxFound;
}

//===============================================================================================//
//  Description:
//      Shift the bytes to the left
//
//  Parameters:
//      count - the number of bytes to shift. If exceeds the array size will limit.
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::LeftShift( size_t count )
{
    if ( ( m_pBytes == nullptr ) || ( m_uSize == 0 ) || ( count == 0 ) )
    {
        return;     // Nothing to do
    }

    // Limit
    if ( count > m_uSize )
    {
        count = m_uSize;
    }
    memmove( m_pBytes, m_pBytes + count, m_uSize - count );
    memset( m_pBytes + ( m_uSize - count ), 0, count );
    m_uSize -= count;
}

//===============================================================================================//
//  Description:
//      Truncate the size of this array to the specified value
//
//  Parameters:
//      size  - the new size, no effect if size is larger than the current size
//              Any allocated bytes beyond size are zeroed
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::Truncate( size_t size )
{
    if ( ( m_pBytes == nullptr ) || ( size >= m_uSize ) )
    {
        return;     // Nothing to do
    }
    memset( m_pBytes + size, 0, m_uSize - size );
    m_uSize = size;
}

//===============================================================================================//
//  Description:
//      Set the size to zero
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void ByteArray::Zero()
{
    Truncate( 0 );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
