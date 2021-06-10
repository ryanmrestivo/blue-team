///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Out-Of-Memory Exception Class Implementation
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
#include "PxsBase/Header Files/MemoryException.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Global Defines
///////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////
// new handler for use with _set_new_handler

#ifdef _MSC_VER

    // Suppress compiler warning about unreachable code. Note, do not
    // want to use __declspec(noreturn) as this function returns an int.
    #pragma warning( push )
    #pragma warning ( disable : 4702 )

    int PXSNewHandler( size_t /* size */ )
    {
        throw MemoryException();

        return 0;   // Unreachable, 0 tells new to stop allocation attempts.
    }
    #pragma warning( pop )

#elif ( __GNUC__ )

    void PXSNewHandler()
    {
        // The GCC compiler can only do 1 of 3 things when a new operator
        //  fails to allocate the requested memory:
        //  1. Try again, there is no way to tell GCC to stop trying.
        //  2. Abort or exit the process.
        //  3. Throw an exception which is descended from std::bad_alloc

        throw MemoryException();
    }

#else

    #error Unsupported compiler

#endif

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
MemoryException::MemoryException()
                :Exception( PXS_ERROR_TYPE_SYSTEM,
                            ERROR_OUTOFMEMORY,
                            nullptr,
                            nullptr )
{
}

// Constructor with the throwing method/function
MemoryException::MemoryException( const char* pszFunction )
                :Exception( PXS_ERROR_TYPE_SYSTEM,
                            ERROR_OUTOFMEMORY,
                            nullptr,
                            pszFunction )
{
}

// Copy constructor
MemoryException::MemoryException( const MemoryException& oMemory )
                :Exception( PXS_ERROR_TYPE_SYSTEM,
                            ERROR_OUTOFMEMORY,
                            nullptr,
                            nullptr )
{
    *this = oMemory;
}


//
// Destructor. The exception specifier throw () is required because this
// class is derived from std::bad_alloc whose destructor specifies
// std::~bad_alloc() throw ()
MemoryException::~MemoryException() throw ()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
MemoryException& MemoryException::operator= ( const MemoryException& oMemory )
{
    if ( this == &oMemory ) return *this;

    // Base class
    Exception::operator= ( oMemory );

    m_uErrorCode  = oMemory.m_uErrorCode;
    m_uErrorType  = oMemory.m_uErrorType;
    m_CrashString = oMemory.m_CrashString;
    m_Description = oMemory.m_Description;
    m_Details     = oMemory.m_Details;
    m_Message     = oMemory.m_Message;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
