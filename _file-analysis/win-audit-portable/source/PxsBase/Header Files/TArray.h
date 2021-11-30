///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Template Array Class Header - Inclusion Model
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

#ifndef PXSBASE_TARRAY_H_
#define PXSBASE_TARRAY_H_

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
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

template< class T >
class TArray
{
    public:
        // Default constructor
        TArray();

        // Copy constructor
        TArray( const TArray& oTArray );

        // Destructor
        ~TArray();

        // Assignment operator
        TArray& operator= ( const TArray& oTArray );

        // Methods
        void        Add( const T& Element );
        void        Append( const TArray& NewArray );
        void        Copy( const TArray& NewArray );
        const T&    Get( size_t index ) const;
        size_t      GetSize() const;
        void        Insert( size_t index, const T& Element );
        void        Remove( size_t index );
        void        RemoveAll();
        void        Set( size_t index, const T& Element );
        void        SetSize( size_t newSize );

    protected:
        // Methods

        // Data members
        T*      m_pArray;       // Pointer to array of elements
        size_t  m_uSize;        // Number of elements in array

    private:
        // Methods

        // Data members
};

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Template Array Class Implementation
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
template< class T >
TArray< T >::TArray()
           :m_pArray( nullptr ),
            m_uSize( 0 )
{
}

// Copy constructor
template< class T >
TArray< T >::TArray( const TArray& oTArray )
            :TArray()

{
    *this = oTArray;
}

// Destructor
template< class T >
TArray< T >::~TArray()
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
template< class T >
TArray< T >& TArray< T >::operator= ( const TArray& oTArray )
{
    // Disallow self-assignment
    if ( this == &oTArray ) return *this;

    Copy( oTArray );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add an element to the end of the array
//
//  Parameters:
//      Element - the element to add
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TArray< T >::Add( const T& Element )
{
    // "Insert" at the end
    return Insert( m_uSize, Element );
}

//===============================================================================================//
//  Description:
//      Append an array to the end of this array
//
//  Parameters:
//      NewArray - the array to append
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TArray< T >::Append( const TArray& NewArray )
{
    size_t i = 0, newSize;
    T* pNewArray = nullptr;

    // Cannot append an array to itself
    if ( this == &NewArray )
    {
        throw FunctionException( L"this = NewArray", __FUNCTION__ );
    }
    newSize   = PXSAddSizeT ( m_uSize, NewArray.m_uSize );
    pNewArray = new T[ newSize ];
    if ( pNewArray == nullptr )
    {
         throw MemoryException( __FUNCTION__ );
    }

    try
    {
        for ( i = 0; i < m_uSize; i++ )
        {
            pNewArray[ i ] = m_pArray[ i ];
        }

        // Copy in elements to be appended
        for ( i = 0; i < NewArray.m_uSize; i++ )
        {
            pNewArray[ m_uSize + i ] = NewArray.m_pArray[ i ];
        }
    }
    catch ( const Exception& )
    {
        delete[] pNewArray;
        throw;
    }
    delete[] m_pArray;
    m_pArray = pNewArray;
    m_uSize  = newSize;
}

//===============================================================================================//
//  Description:
//      Copy the contents of the specified array into this array
//
//  Parameters:
//      NewArray - the array to copy
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TArray< T >::Copy ( const TArray& NewArray )
{
    RemoveAll();
    Append( NewArray );
}

//===============================================================================================//
//  Description:
//      Get the element at the specified index
//
//  Parameters:
//      index - the zero-based index to get
//
//  Returns:
//      Reference to the element
//===============================================================================================//
template< class T >
const T& TArray< T >::Get( size_t index ) const
{
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
template< class T >
size_t TArray< T >::GetSize() const
{
    return m_uSize;
}

//===============================================================================================//
//  Description:
//      Insert an element into the array
//
//  Parameters:
//      index   - the zero-based index of where to insert the element
//      element - the element to insert
//
//  Remarks:
//      Can insert at the end.
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TArray< T >::Insert( size_t index, const T& Element )
{
    T* pNewArray = nullptr;
    size_t i = 0, newSize;

    // Bounds check, not using >= in order to "insert" at end of array which is equivalent to append
    if ( index > m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }
    newSize   = PXSAddSizeT( m_uSize, 1 );
    pNewArray = new T[ newSize ];
    if ( pNewArray == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    try
    {
        for ( i = 0; i < index; i++ )
        {
            pNewArray[ i ] = m_pArray[ i ];
        }
        pNewArray[ index ] = Element;

        for ( i = index; i < m_uSize; i++ )
        {
            pNewArray[ i + 1 ] = m_pArray[ i ];
        }
    }
    catch ( const Exception& )
    {
        delete[] pNewArray;
        throw;
    }
    delete[] m_pArray;
    m_pArray = pNewArray;
    m_uSize  = newSize;
}

//===============================================================================================//
//  Description:
//      Remove the element at the specified index from the array
//
//  Parameters:
//      index - the zero-based index of the element to remove
//
//  Remarks:
//      Ignores out-of-bounds errors as the element does not exist.
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TArray< T >::Remove( size_t index )
{
    T* pNewArray = nullptr;
    size_t i = 0, newSize;

    // If out-of-bound will return without error
    if ( m_uSize == 0 ) return;
    if ( index >= m_uSize ) return;

    // Set new size
    newSize = m_uSize - 1;     // From above m_uSize > 0
    if ( newSize == 0 )
    {
        RemoveAll();
        return;
    }

    pNewArray = new T[ newSize ];
    if ( pNewArray == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    try
    {
        for ( i = 0; i < index; i++ )
        {
            pNewArray[ i ] = m_pArray[ i ];
        }

        // Skip the element to be removed and copy in rest
        for ( i = (index + 1); i < m_uSize; i++ )
        {
            pNewArray[ i - 1 ] = m_pArray[ i ];
        }
    }
    catch ( const Exception& )
    {
        delete[] pNewArray;
        throw;
    }
    delete[] m_pArray;
    m_pArray = pNewArray;
    m_uSize  = newSize;
}

//===============================================================================================//
//  Description:
//      Remove all the elements from the array
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TArray< T >::RemoveAll()
{
    delete[] m_pArray;
    m_pArray = nullptr;
    m_uSize  = 0;
}

//===============================================================================================//
//  Description:
//      Set the element at the specified index in the array
//
//  Parameters:
//      index - the zero-based index of the element to set
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TArray< T >::Set( size_t index, const T& Element )
{
    if ( index >= m_uSize )
    {
        throw BoundsException( L"index", __FUNCTION__ );
    }
    m_pArray[ index ] = Element;
}

//===============================================================================================//
//  Description:
//      Set the size of the array
//
//  Parameters:
//      newSize - the new size of the array
//
//  Remarks:
//      Preserves any existing array elements up to newSize
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TArray< T >::SetSize( size_t newSize )
{
    T* pNewArray = nullptr;
    size_t i = 0, numElements = 0;

    if ( newSize == 0 )
    {
        RemoveAll();
        return;
    }
    pNewArray = new T[ newSize ];
    if ( pNewArray == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    try
    {
        // Test if need to create a new array
        if ( m_pArray == nullptr )
        {
            m_pArray = pNewArray;
            m_uSize  = newSize;
        }
        else
        {
            // Array already exists so fill the new array with old data, do
            // not run past end of either the new or old arrays.
            numElements = PXSMinSizeT( newSize, m_uSize );
            for ( i = 0; i < numElements; i++ )
            {
                pNewArray[ i ] = m_pArray[ i ];   // Assignment
            }
            delete[] m_pArray;
            m_pArray = pNewArray;
            m_uSize  = newSize;
        }
    }
    catch ( const Exception& )
    {
        delete[] pNewArray;
        throw;
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////


#endif  // PXSBASE_TARRAY_H_
