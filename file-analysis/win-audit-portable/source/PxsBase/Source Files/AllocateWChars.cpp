///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Allocate Wide/wchar_t Characters Class Implementation
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
#include "PxsBase/Header Files/AllocateWChars.h"

// 2. C System Files
#include <wchar.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
AllocateWChars::AllocateWChars()
               :m_pWChars( nullptr ),
                m_tNumWChars( 0 )
{
}

// Copy constructor - not allowed so no implementation

// Destructor - do not throw any exceptions
AllocateWChars::~AllocateWChars()
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
//      Delete the allocate wchar_t's
//
//  Parameters:
//      None
//
//  Remarks:
//      Called by destructor, must not throw.
//
//  Returns:
//      void
//===============================================================================================//
void AllocateWChars::Delete()
{
    if ( m_pWChars )
    {
        delete[] m_pWChars;
        m_pWChars = nullptr;
    }
    m_tNumWChars = 0;
}

//===============================================================================================//
//  Description:
//      Allocate wchar_t's using the new[] operator
//
//  Parameters:
//      numWChars - number of wchar_t's to allocate
//
//  Remarks:
//      It is an error to allocate wchar_t's if have done already so. The caller
//      must call Delete() to free any allocated the memory before calling
//      this method again otherwise an exception is thrown.
//
//      Memory is initialized to zero
//
//  Returns:
//      void
//===============================================================================================//
wchar_t* AllocateWChars::New( size_t numWChars )
{
    // Test if have already allocated WChars
    if ( m_pWChars )
    {
        throw FunctionException( L"m_pWChars", __FUNCTION__ );
    }

    // Always create at least 1 byte, to mimic the new operator in the sense
    // that NULL means failure.
    if ( numWChars == 0 )
    {
        numWChars = 1;
    }

    m_pWChars = new wchar_t[ numWChars ];
    if ( m_pWChars == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    m_tNumWChars = numWChars;
    wmemset( m_pWChars, 0, m_tNumWChars );

    return m_pWChars;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
