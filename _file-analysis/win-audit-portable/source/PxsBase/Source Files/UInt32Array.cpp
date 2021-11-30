///////////////////////////////////////////////////////////////////////////////////////////////////
//
// UInt32/DWORD Class Implementation
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
#include "PxsBase/Header Files/UInt32Array.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/NullException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
UInt32Array::UInt32Array()
            :REALLOC_SIZE( 64 ),  // pseudo-constant
             m_uSize( 0 ),
             m_uAllocated( 0 ),
             m_pArray( nullptr )
{
}

// Copy constructor
UInt32Array::UInt32Array( const UInt32Array& oArray )
            :REALLOC_SIZE( 64 ),  // pseudo-constant
             m_uSize( 0 ),
             m_uAllocated( 0 ),
             m_pArray( nullptr )
{
    *this = oArray;
}

// Destructor
UInt32Array::~UInt32Array()
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
UInt32Array& UInt32Array::operator=( const UInt32Array& oArray )
{
    // Disallow self-assignment
    if ( this == &oArray ) return *this;

    RemoveAll();
    AddArray( oArray );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add a 32-bit unsigned integer to the end of this array
//
//  Parameters:
//      un32 - the value
//
//  Returns:
//      void
//===============================================================================================//
void UInt32Array::Add( DWORD un32 )
{
    size_t newSize = PXSAddSizeT( m_uSize, 1 );

    SetSize( newSize );
    Set( newSize - 1, un32 );  // zero based
}

//===============================================================================================//
//  Description:
//      Add a 32-bit unsigned integer array to this array
//
//  Parameters:
//      oArray - 32-bit unsigned integer array object
//
//  Remarks:
//      Adding one at a time, if this becomes a performance issue can re-work.
//
//  Returns:
//      void
//===============================================================================================//
void UInt32Array::AddArray( const UInt32Array& oArray )
{
    DWORD  un32 = 0;
    size_t i    = 0;
    size_t size = oArray.GetSize();

    // Disallow adding an array to itself
    if ( this == &oArray )
    {
        throw FunctionException( L"this = oArray", __FUNCTION__ );
    }

    for ( i = 0; i < size; i++ )
    {
        un32 = oArray.Get( i );
        Add( un32 );
    }
}

//===============================================================================================//
//  Description:
//      Add n 32-bit unsigned integer to this array if not already present.
//
//  Parameters:
//      un32 - the value
//
//  Returns:
//      true if added the integer, otherwise false
//===============================================================================================//
bool UInt32Array::AddUnique( DWORD un32 )
{
    bool   success = false, found = false;
    size_t i = 0;

    if ( m_pArray == nullptr )
    {
        Add( un32 );
        return true;
    }

    while ( ( found == false ) && ( i < m_uSize ) )
    {
        if ( un32 == m_pArray[ i ] )
        {
            found = true;
        }
        i++;
    }

    if ( found == false )
    {
        Add( un32 );
        success = true;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Get the 32-bit unsigned integer at the specified index in the array
//
//  Parameters:
//      index - the zero-based index
//
//  Returns:
//      32-bit unsigned integer
//===============================================================================================//
DWORD UInt32Array::Get( size_t index ) const
{
    if ( m_pArray == nullptr )
    {
        throw NullException(  L"m_pArray", __FUNCTION__ );
    }

    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    return m_pArray[ index ];
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
size_t UInt32Array::GetSize() const
{
    return m_uSize;
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
void UInt32Array::RemoveAll()
{
    if ( m_pArray )
    {
        delete [] m_pArray;
    }
    m_pArray     = nullptr;
    m_uSize      = 0;
    m_uAllocated = 0;
}

//===============================================================================================//
//  Description:
//      Set the 32-bit unsigned integer at the specified array index
//
//  Parameters:
//      index - the array index
//      un32  - the 32-bit unsigned integer
//
//  Returns:
//      void
//===============================================================================================//
void UInt32Array::Set( size_t index, DWORD un32 )
{
    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }

    if ( m_pArray == nullptr )
    {
        throw NullException( L"m_pArray", __FUNCTION__ );
    }
    m_pArray[ index ] = un32;
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
void UInt32Array::SetSize( size_t newSize )
{
    DWORD* pNew;
    size_t i = 0, allocate;

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
        // Init the elements to zero
        for ( i = newSize; i < m_uAllocated; i++ )
        {
            m_pArray[ i ] = 0;
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

    pNew = new DWORD[ allocate ];
    if ( pNew == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( pNew, 0, sizeof ( DWORD ) * allocate );

    // Replace, reserve any existing elements
    if ( m_pArray )
    {
        memcpy( pNew, m_pArray, sizeof ( DWORD ) * m_uAllocated );
        delete[] m_pArray;
    }
    m_pArray     = pNew;
    m_uAllocated = allocate;
    m_uSize      = newSize;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
