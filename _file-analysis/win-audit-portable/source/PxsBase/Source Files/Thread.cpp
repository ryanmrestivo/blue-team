///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Thread Class Implementation
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

// As well as the stop flag m_bStopMT, an event m_hStopEventMT can used to
// signal the worker to stop.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/Thread.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/UnhandledException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Thread::Thread()
       :m_bRunMT( FALSE ),
        m_uThreadIdMT( DWORD_MAX ),       // i.e. invalid
        m_hWndListenerMT( nullptr ),
        m_hDoTaskEventMT( nullptr ),
        m_hStopEventMT( nullptr ),
        m_hThread( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
Thread::~Thread()
{
    try
    {
        Join();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }

    // Close this one first in case something went wrong in Join, don't want to
    // close an event being waited on
    if ( m_hThread )
    {
        CloseHandle( m_hThread );
    }

    if ( m_hDoTaskEventMT )
    {
        CloseHandle( m_hDoTaskEventMT );
    }

    if ( m_hStopEventMT )
    {
        CloseHandle( m_hStopEventMT );
    }
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
//      Get this thread's identifier
//
//  Parameters:
//      None
//
//  Returns:
//      the id, -1 if not yest created
//===============================================================================================//
DWORD Thread::GetThreadId() const
{
    return m_uThreadIdMT;
}

//===============================================================================================//
//  Description:
//      Determine if the worker thread has been created
//
//  Parameters:
//      None
//
//  Returns:
//      True if the worker thread is running
//===============================================================================================//
bool Thread::IsCreated() const
{
    if ( m_hThread )
    {
        return true;
    }
    return false;
}

//===============================================================================================//
//  Description:
//      Wait for the thread to exit
//
//  Parameters:
//      None
//
//  Returns:
//      The thread's exit code
//===============================================================================================//
DWORD Thread::Join()
{
    DWORD exitCode = 0;

    if ( m_hThread == nullptr )
    {
        return ERROR_SUCCESS;     // Nothing to do
    }

    // Test if running
    if ( GetExitCodeThread( m_hThread, &exitCode ) )
    {
        if ( exitCode != STILL_ACTIVE )
        {
            return exitCode;
        }
    }

    if ( WaitForSingleObject( m_hThread, INFINITE ) == WAIT_FAILED )
    {
        PXSLogSysError( GetLastError(), L"WaitForSingleObject failed in Thread::Join" );
    }
    GetExitCodeThread( m_hThread, &exitCode );

    return exitCode;
}

//===============================================================================================//
//  Description:
//      Run the thread
//
//  Parameters:
//      hWndListener - optional, handle of window that listens for messages
//                     posted by the worker thread. The caller should ensure
//                     hWndListener remains valid for the lifetime of
//                     the worker thread.
//
//  Returns:
//      void
//===============================================================================================//
void Thread::Run( HWND hWndListener )
{
    // One shot
    if ( m_hThread )
    {
        throw FunctionException( L"m_hThread", __FUNCTION__ );
    }

    // Create the stop event. Note, this is not an event for OVERLAPPED so
    // will use an ordinary auto-reset
    if ( m_hStopEventMT == nullptr )
    {
        m_hStopEventMT = CreateEvent( nullptr,
                                      FALSE,      // Auto-reset
                                      FALSE,      // Not signaled
                                      nullptr );  // No name
        if ( m_hStopEventMT == nullptr )
        {
            throw SystemException( GetLastError(),
                                   L"CreateEvent", __FUNCTION__ );
        }
    }

    // Create the stop event. Note, this is not an event for OVERLAPPED so
    // will use an ordinary auto-reset
    if ( m_hDoTaskEventMT == nullptr )
    {
        m_hDoTaskEventMT = CreateEvent( nullptr,
                                        FALSE,      // Auto-reset
                                        FALSE,      // Not signaled
                                        nullptr );  // No name
        if ( m_hDoTaskEventMT == nullptr )
        {
            throw SystemException( GetLastError(),
                                   L"CreateEvent", __FUNCTION__ );
        }
    }
    m_bRunMT = TRUE;
    m_hWndListenerMT = hWndListener;

    // Create the after setting m_bRunMT so that it continues to run
    m_hThread = CreateThread( nullptr, 0, Thread::StartRoutine, this, 0, &m_uThreadIdMT );
    if ( m_hThread == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateThread", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Tell the thread to stop
//
//  Parameters:
//      None
//
//  Remarks:
//       Worker should poll the flag at reqular intervals
//
//  Returns:
//      void
//===============================================================================================//
void Thread::Stop()
{
    m_bRunMT = FALSE;

    // Signal the stop event in case the worker is waiting
    if ( m_hStopEventMT )
    {
        if ( SetEvent( m_hStopEventMT ) == 0 )
        {
            PXSLogSysError( GetLastError(), L"SetEvent failed in Thread::Stop" );
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Run the worker thread. This method must only be called by the worker.
//
//  Parameters:
//      None
//
//  Remarks
//      Override for useful functionality
//
//  Returns:
//      DWORD system error code
//===============================================================================================//
DWORD Thread::RunWorkerThread()
{
    return ERROR_SUCCESS;
}

//===============================================================================================//
//  Description:
//      Entry point for CreateThread
//
//  Parameters:
//      pParameter - the CreateThread parameter
//
//  Returns:
//      DWORD system error code
//===============================================================================================//
DWORD WINAPI Thread::StartRoutine( void* pParameter )
{
    DWORD   result = ERROR_SUCCESS;
    Thread* pThread;

    if ( pParameter == nullptr )
    {
        return ERROR_INVALID_PARAMETER;
    }
    pThread = reinterpret_cast< Thread* >( pParameter );

    // Catch unexpected exceptions here, if have a window, show a message
    // otherwise log it then abort
    if ( g_pApplication && ( g_pApplication->GetHwndMainFrame() ) )
    {
        __try
        {
            result = pThread->RunWorkerThread();
        }
        __except( PXSShowUnhandledExceptionDialog( GetExceptionInformation() ) )
        {
            ExitProcess( 3 );       // same as abort
        }
    }
    else
    {
        __try
        {
            result = pThread->RunWorkerThread();
        }
        __except( PXSWriteUnhandledExceptionToLog( GetExceptionInformation() ) )
        {
            ExitProcess( 3 );       // same as abort
        }
    }

    return result;
}

