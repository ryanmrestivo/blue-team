///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Allocate ANSI Characters Class Implementation
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

// This class is for the allocation of memory. The most convenient usage is to
// declared an instance on the stack. When it does out of scope the memory
// is automatically deleted.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/AllocateChars.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
AllocateChars::AllocateChars()
              :m_pChars( nullptr ),
               m_tNumChars( 0 )
{
}

// Copy constructor - not allowed so no implementation

// Destructor - do not throw any exceptions
AllocateChars::~AllocateChars()
{
    Delete();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Delete the allocate characters
//
//  Parameters:
//      None
//
//  Remarks:
//      Called by destructor, must not throw.
//
//      This method call be called at any time to free the allocated memory.
//
//  Returns:
//      void
//===============================================================================================//
void AllocateChars::Delete()
{
    if ( m_pChars )
    {
        delete[] m_pChars;
        m_pChars = nullptr;
    }
    m_tNumChars = 0;
}

//===============================================================================================//
//  Description:
//      Allocate chars using the new[] operator
//
//  Parameters:
//      numChars - number of chars to allocate
//
//  Remarks:
//      It is an error to allocate if have done already so. The caller
//      must call Delete() to free any allocated the memory before calling
//      this method again otherwise an exception is thrown.
//
//      Memory is initialized to zero
//
//  Returns:
//      void
//===============================================================================================//
char* AllocateChars::New( size_t numChars )
{
    // Test if have already allocated Chars
    if ( m_pChars )
    {
        throw FunctionException( L"m_pChars", __FUNCTION__ );
    }

    // Always create at least 1 byte, to mimic the new operator in the sense
    // that NULL means failure.
    if ( numChars == 0 )
    {
        numChars = 1;
    }

    m_pChars = new char[ numChars ];
    if ( m_pChars == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    m_tNumChars = numChars;
    memset( m_pChars, 0, m_tNumChars );

    return m_pChars;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
