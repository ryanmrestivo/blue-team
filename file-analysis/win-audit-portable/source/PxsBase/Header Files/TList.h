///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Template List Class Header - Inclusion Model
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

#ifndef PXSBASE_TLIST_H_
#define PXSBASE_TLIST_H_

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
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Formatter.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

template< class T >
class TList
{
    public:
        // Default constructor
        TList();

        // Copy constructor
        TList( const TList& oTList );

        // Assignment operator
        TList& operator= ( const TList& oTList );

        // Destructor
        ~TList();

        // Methods
        void     AddAtHead( const T& Element );
        bool     Advance();
        void     Append( const T& Element );
        void     AppendList( const TList< T >& List );
        void     End();
        const T& Get() const;
        size_t   GetLength() const;
        T*       GetPointer() const;
        bool     IsEmpty() const;
        void     Insert( const T& Element );
        void     PopHead( T* pElement );
        bool     Previous();
        void     Rewind();
        void     Remove();
        void     RemoveAll();
        void     Set( const T& Element );

        // Data members

    protected:
        // Methods

        // Data members

    private:
        typedef struct _TYPE_NODE
        {
            T* pElement;
            _TYPE_NODE* pNext;
            _TYPE_NODE* pPrev;
        } TYPE_NODE;

        // Methods
        TYPE_NODE* AllocateNode( const T& Element );

        // Data members
        size_t     m_uLength;
        TYPE_NODE* m_pHead;
        TYPE_NODE* m_pTail;
        TYPE_NODE* m_pCurrent;
};

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Template List Class Implementation
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
template< class T >
TList< T >::TList()
           :m_uLength( 0 ),
            m_pHead( nullptr ),
            m_pTail( nullptr ),
            m_pCurrent( nullptr )
{
}

// Copy constructor
template< class T >
TList< T >::TList( const TList& oTList )
           :TList()
{
    *this = oTList;
}

// Destructor
template< class T >
TList< T >::~TList()
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
TList< T >& TList< T >::operator= ( const TList< T >& oTList )
{
    if ( this == &oTList ) return *this;

    RemoveAll();
    AppendList( oTList );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add the specified element at the head of the list
//
//  Parameters:
//      Element - the element to add
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::AddAtHead( const T& Element )
{
    size_t length   = 0;
    TYPE_NODE* pNew = nullptr;

    if ( m_pHead == nullptr )
    {
        Append( Element );
    }
    else
    {
        length         = PXSAddSizeT( m_uLength, 1 );
        pNew           = AllocateNode( Element );
        pNew->pNext    = m_pHead;
        m_pHead->pPrev = pNew;
        m_pHead        = pNew;
        m_pCurrent     = pNew;      // = m_pHead, i.e. position on the new head
        m_uLength      = length;
    }
}

//===============================================================================================//
//  Description:
//      Advance the current position of the list to the next node
//
//  Parameters:
//      None
//
//  Returns:
//      true if advanced, else false (i.e. end of list).
//===============================================================================================//
template< class T >
bool TList< T >::Advance()
{
    if ( m_pCurrent && m_pCurrent->pNext )
    {
        m_pCurrent = m_pCurrent->pNext;
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Add a node to the end of the list. The position is set on the
//      new node
//
//  Parameters:
//      Element - the element to add
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::Append( const T& Element )
{
    size_t length   = PXSAddSizeT( m_uLength, 1 );
    TYPE_NODE* pNew = AllocateNode( Element );

    // Special case of empty list
    if ( m_pHead == nullptr )
    {
        m_pHead        = pNew;
        m_pHead->pPrev = nullptr;   // No nodes before the head
        m_pHead->pNext = nullptr;   // No nodes after the head
    }
    else
    {
        // The tail should be non-null if the list has nodes
        if ( m_pTail == nullptr )
        {
            throw NullException( L"m_pTail", __FUNCTION__ );
        }
        pNew->pPrev    = m_pTail;
        pNew->pNext    = nullptr;   // No nodes after the new one
        m_pTail->pNext = pNew;
    }
    m_pCurrent = pNew;
    m_pTail    = pNew;
    m_uLength  = length;
}

//===============================================================================================//
//  Description:
//      Append the specified list to this list
//
//  Parameters:
//      List - the list to appen
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::AppendList( const TList< T >& List )
{
    TYPE_NODE* pNode = List.m_pHead;

    while ( pNode )
    {
        if ( pNode->pElement )
        {
            Append( *pNode->pElement );
        }
        pNode = pNode->pNext;
    }
}

//===============================================================================================//
//  Description:
//      Set the position to the end of the list
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::End()
{
    m_pCurrent = m_pTail;
}


//===============================================================================================//
//  Description:
//      Get the element at the current position
//
//  Parameters:
//      None
//
//  Returns:
//      Reference to the element
//===============================================================================================//
template< class T >
const T& TList< T >::Get() const
{
    if ( m_pCurrent == nullptr )
    {
        throw FunctionException( L"m_pCurrent", __FUNCTION__ );
    }

    if ( m_pCurrent->pElement == nullptr )
    {
        throw NullException( L"pElement", __FUNCTION__ );
    }

    return *m_pCurrent->pElement;
}

//===============================================================================================//
//  Description:
//      Get the length of this list
//
//  Parameters:
//      None
//
//  Returns:
//     The length
//===============================================================================================//
template< class T >
size_t TList< T >::GetLength() const
{
    return m_uLength;
}

//===============================================================================================//
//  Description:
//      Get a pointer element at the current position
//
//  Parameters:
//      None
//
//  Returns:
//      Pointer to the element. Does not return NULL.
//===============================================================================================//
template< class T >
T* TList< T >::GetPointer() const
{
    if ( m_pCurrent == nullptr )
    {
        throw FunctionException( L"m_pCurrent", __FUNCTION__ );
    }

    if ( m_pCurrent->pElement == nullptr )
    {
        throw NullException( L"pElement", __FUNCTION__ );
    }

    return m_pCurrent->pElement;
}

//===============================================================================================//
//  Description:
//      Inset a node at the current position. The position is set on the new node
//
//  Parameters:
//      Element - the element to add
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::Insert( const T& Element )
{
    size_t  length;
    TYPE_NODE* pNew;
    TYPE_NODE* pPrev;

    // Special case of empty list
    if ( IsEmpty() )
    {
        Append( Element );
        return;
    }

    // Handle inserting at the head, i.e. no previous node
    if ( m_pCurrent == m_pHead )
    {
        AddAtHead( Element );
        return;
    }

    // Insert some where after the head, i.e. m_pCurrent->pPrev is non-NULL
    // Prev            Prev
    // Current   ==>   New
    // Next            Current
    //                 Next
    length = PXSAddSizeT( m_uLength, 1 );
    pNew   = AllocateNode( Element );
    pPrev             = m_pCurrent->pPrev;
    pPrev->pNext      = pNew;
    pNew->pPrev       = pPrev;
    pNew->pNext       = m_pCurrent;
    m_pCurrent->pPrev = pNew;
    m_pCurrent        = pNew;       // Postion on new element
    m_uLength         = length;
}

//===============================================================================================//
//  Description:
//      Determine if the list is empty
//
//  Parameters:
//      None
//
//  Returns:
//      true if empty, otherwise false
//===============================================================================================//
template< class T >
bool TList< T >::IsEmpty() const
{
    if ( m_pHead == nullptr )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Remove the head node from the list.
//
//  Parameters:
//      pElement - receives the element at the head
//
//  Remarks
//      The list cannot be empty. On return the list is positioned on the new head, if any
//
//  Returns:
//      the element that was on head node
//===============================================================================================//
template< class T >
void TList< T >::PopHead( T* pElement )
{
    if ( pElement == nullptr )
    {
        throw ParameterException( L"pElement", __FUNCTION__ );
    }

    Rewind();
    *pElement = Get();
    Remove();
}

//===============================================================================================//
//  Description:
//      Move to the node before the current position
//
//  Parameters:
//      None
//
//  Returns:
//      true if moved, else false (i.e. start of list)
//===============================================================================================//
template< class T >
bool TList< T >::Previous()
{
    if ( m_pCurrent && m_pCurrent->pPrev )
    {
        m_pCurrent = m_pCurrent->pPrev;
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Set the position to the start of the list
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::Rewind()
{
    m_pCurrent = m_pHead;
}

//===============================================================================================//
//  Description:
//      Remove the current node from the list
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::Remove()
{
    TYPE_NODE* pPrev;
    TYPE_NODE* pNext;
    TYPE_NODE* pRemove;

    if ( m_pCurrent == nullptr )
    {
        return;     // Nothing to remove or invalid function
    }
    pRemove = m_pCurrent;

    // Three possibilities
    if ( m_pCurrent == m_pHead )
    {
        m_pHead = m_pHead->pNext;
        if ( m_pHead )
        {
            m_pHead->pPrev = nullptr;
        }
        m_pCurrent = m_pHead;
        if ( ( m_pCurrent == nullptr ) || ( m_pCurrent->pNext == nullptr ) )
        {
            m_pTail = m_pCurrent;
        }
    }
    else if ( m_pCurrent == m_pTail )
    {
        m_pTail        = m_pTail->pPrev;
        m_pTail->pNext = nullptr;
        m_pCurrent     = m_pTail;
    }
    else
    {
        // Removing from in the middle, the head and tail are unchanged
        pPrev = m_pCurrent->pPrev;
        pNext = m_pCurrent->pNext;
        pPrev->pNext = pNext;
        pNext->pPrev = pPrev;
        if ( pPrev )
        {
            m_pCurrent = pPrev;
        }
        else
        {
            m_pCurrent= pNext;
        }
    }

    // Delete
    delete pRemove->pElement;
    delete pRemove;
    if ( m_uLength )
    {
        m_uLength--;
    }
}

//===============================================================================================//
//  Description:
//      Empty the list
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::RemoveAll()
{
    TYPE_NODE* pNext;
    TYPE_NODE* pNode = m_pHead;

    while ( pNode )
    {
        delete pNode->pElement;     // First the element
        pNext = pNode->pNext;
        delete pNode;               // Then the node itself
        pNode = pNext;
    }
    m_pHead    = nullptr;
    m_pTail    = nullptr;
    m_pCurrent = nullptr;
    m_uLength  = 0;
}

//===============================================================================================//
//  Description:
//      Set the element at the current position
//
//  Parameters:
//      Element - the element, this replaces the existing one
//
//  Returns:
//      void
//===============================================================================================//
template< class T >
void TList< T >::Set( const T& Element )
{
    if ( m_pCurrent == nullptr )
    {
        throw FunctionException( L"m_pCurrent", __FUNCTION__ );
    }

    T* pNew = new T( Element );
    if ( pNew == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    delete m_pCurrent->pElement;
    m_pCurrent->pElement = pNew;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Allocate a new node and put a copy of the specified element in it.
//
//  Parameters:
//      Element - the element to copy into the node
//
//  Returns:
//      pointer to the new node, does not return NULL
//===============================================================================================//
template< class T >
typename TList< T >::TYPE_NODE* TList< T >::AllocateNode( const T& Element )
{
    TYPE_NODE* pNew = new TYPE_NODE;
    if ( pNew == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( pNew, 0, sizeof ( TYPE_NODE ) );

    try
    {
        pNew->pElement = new T( Element );
    }
    catch ( const Exception& )
    {
        delete pNew;
        throw;
    }

    if ( pNew->pElement == nullptr )
    {
        delete pNew;
        throw MemoryException( __FUNCTION__ );
    }

    return pNew;
}

#endif  // PXSBASE_TLIST_H_
