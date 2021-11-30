///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Mutex/Lock Class Implementation
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
// Include files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/Mutex.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Mutex::Mutex()
      :m_CS()     // non-op
{
    InitializeCriticalSection( &m_CS );
}

// Copy constructor - not allowed so no implementation

// Destructor
Mutex::~Mutex()
{
    DeleteCriticalSection( &m_CS );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

//
// Assignment operator - no implementation as not allowed
//

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Acquire a lock by waiting indefinitely
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Mutex::Lock()
{
    EnterCriticalSection( &m_CS );
}

//===============================================================================================//
//  Description:
//      Acquire a lock
//
//  Parameters:
//      seconds - the time in seconds to wait for the lock, use DWORD_MAX
//                for an infinite wait.
//
//  Returns:
//      void
//===============================================================================================//
void Mutex::TryLock( DWORD seconds )
{
    BOOL  locked;
    DWORD elasped = 0, random = 0;
    SYSTEMTIME st;

    // Try in randomised 1 second intervals
    locked = TryEnterCriticalSection( &m_CS );
    while ( ( locked == 0 ) && ( elasped <= seconds ) )
    {
        // Wait, seed if on first pass
        if ( elasped == 0 )
        {
            memset( &st, 0, sizeof( st ) );
            GetSystemTime(&st);
            srand( st.wMilliseconds );
        }
        random = PXSCastInt32ToUInt32( ( rand() % 200 ) );
        Sleep( 900 + random );
        elasped++;
        locked = TryEnterCriticalSection( &m_CS );
    }

    if ( locked == 0 )
    {
       throw SystemException( ERROR_TIMEOUT, PXS_STRING_EMPTY, __FUNCTION__);
    }
}

//===============================================================================================//
//  Description:
//      Unlock this mutes
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void Mutex::Unlock()
{
    LeaveCriticalSection( &m_CS );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
