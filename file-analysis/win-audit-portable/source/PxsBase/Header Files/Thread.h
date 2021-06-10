///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Thread Class Header
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

#ifndef PXSBASE_THREAD_H_
#define PXSBASE_THREAD_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files
#include <signal.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Thread
{
    public:
        // Default constructor
        Thread();

        // Destructor
        virtual ~Thread();

        // Methods
        DWORD   GetThreadId() const;
        bool    IsCreated() const;
        DWORD   Join();
        void    Run( HWND hWndListener );
        void    Stop();

    protected:
        // Methods

        //                                 //      ACCESS
        // Members                         // Creator      Worker
        sig_atomic_t    m_bRunMT;          // Read-write   Read
        DWORD           m_uThreadIdMT;     // Write-once   Read
        HWND            m_hWndListenerMT;  // Write-once   Read
        HANDLE          m_hDoTaskEventMT;  // Write-once   Read
        HANDLE          m_hStopEventMT;    // Write-once   Read

    private:
        // Copy constructor - not allowed
        Thread( const Thread& oThread );

        // Assignment operator - not allowed
        Thread& operator= ( const Thread& oThread );

        // Methods
        virtual DWORD        RunWorkerThread();
        static  DWORD WINAPI StartRoutine( void* pParameter );

        // Data members
        HANDLE  m_hThread;
};

#endif  // PXSBASE_THREAD_H_
