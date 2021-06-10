///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Audit Thread Class Implementation
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
#include "WinAudit/Header Files/AuditThread.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AutoUnlockMutex.h"

// 5. This Project
#include "WinAudit/Header Files/AuditData.h"
#include "WinAudit/Header Files/WinauditFrame.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
AuditThread::AuditThread()
            :m_Mutex(),
             m_pAuditThreadParameterMT( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
AuditThread::~AuditThread()
{
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
//      Set the data for the audit. The caller must ensure that pParameter
//      is valid for the life time of the thread
//
//  Parameters:
//      pParameter - holds data about the audit

//  Remarks:
//      Called by worker: No
//
//  Returns:
//      void
//===============================================================================================//
void AuditThread::SetAuditThreadParameter( AuditThreadParameter* pParameter )
{
    if ( m_hDoTaskEventMT == nullptr )
    {
        throw FunctionException( L"m_hDoTaskEventMT", __FUNCTION__ );
    }

    if ( pParameter == nullptr )
    {
        throw ParameterException( L"pParameter", __FUNCTION__ );
    }

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_pAuditThreadParameterMT = pParameter;

    // Tell the worker there is something do do
    if ( SetEvent( m_hDoTaskEventMT ) == 0 )
    {
        PXSLogSysError( GetLastError(), L"SetEvent failed." );
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
//      Do the audit job. This method must only be called by the worker. A lock
//      must have been aquired before calling it.
//
//  Parameters:
//      pParameter - parameter containg data about the audit job
//
//  Remarks:
//      Called by worker: Yes
//
//  Returns:
//      DWORD system error code
//===============================================================================================//
DWORD AuditThread::DoAuditLocked( AuditThreadParameter* pParameter )
{
    DWORD     result = ERROR_SUCCESS, percentDone = 0, categoryID = 0;
    DWORD     timeRemaining = 0;
    size_t    i = 0, numCategories = 0;
    time_t    now = 0;
    AuditData Auditor;
    TArray< AuditRecord > AuditRecords;

    if ( pParameter == nullptr )
    {
        return ERROR_INVALID_PARAMETER;
    }

    // Ensure any exceptions does not leave this entry procedure
    try
    {
        // COM
        PXSInitializeComOnThread();

        // Do the audit
        numCategories = pParameter->Categories.GetSize();
        while ( m_bRunMT && ( i < numCategories ) && ( result == ERROR_SUCCESS ) )
        {
            // Catch errors in individual categories so can continue the job
            percentDone = PXSCastSizeTToUInt32((100 * (i + 1)) / numCategories);
            try
            {
                // Get the data
                AuditRecords.RemoveAll();
                categoryID = pParameter->Categories.Get( i );
                Auditor.GetCategoryRecords( categoryID, pParameter->LocalTime, &AuditRecords );

                // Test for timeout
                time( &now );
                if ( now < pParameter->timeoutAt )
                {
                    timeRemaining = DWORD_MAX;
                    if ( ( pParameter->timeoutAt - now ) < DWORD_MAX )
                    {
                        timeRemaining = PXSCastTimeTToUInt32(
                                                pParameter->timeoutAt - now );
                    }
                   // Send new data to the application
                   result = pParameter->pWinAuditFrame->AuditThreadCallback(
                                                      timeRemaining,
                                                      percentDone, &AuditRecords, nullptr );
                }
                else
                {
                    result = ERROR_TIMEOUT;
                }
            }
            catch ( const Exception& e )
            {
                // Note, not setting result so can continue to next category
                // Will wait but not for too long
               pParameter->pWinAuditFrame->AuditThreadCallback( 10, percentDone, nullptr, &e );
            }
            i++;
        }
    }
    catch ( const Exception& e )
    {
       result = e.GetErrorCode();
       pParameter->pWinAuditFrame->AuditThreadCallback( 10,
                                                        100,  // Done
                                                        nullptr, &e );
    }
    PostMessage( pParameter->pWinAuditFrame->GetHwnd(), PXS_APP_MSG_AUDIT_THREAD_DONE, 0, 0 );

    if ( m_bRunMT )
    {
        result = ERROR_CANCELLED;
    }
    return result;
}

//===============================================================================================//
//  Description:
//      Run the worker thread. This method must only be called by the worker.
//
//  Parameters:
//      None
//
//  Returns:
//      DWORD system error code
//===============================================================================================//
DWORD AuditThread::RunWorkerThread()
{
    DWORD  WAIT_TIMOUT_MS = 250;
    DWORD  result = 0;
    HANDLE handles[] = { m_hDoTaskEventMT, m_hStopEventMT };
    AuditThreadParameter* pAuditThreadParameter = nullptr;

    if ( ( handles[ 0 ] == nullptr ) || ( handles[ 1 ] == nullptr ) )
    {
        return ERROR_INVALID_FUNCTION;
    }

    // Loop until told to stop
    do
    {
        try
        {
            // Because in a loop will reset the audit thread parameter to
            // NULL so as not to keep cycling as want to use a wait function
            // with a timeout
            m_Mutex.Lock();
            AutoUnlockMutex AutoUnlock( &m_Mutex );
            if ( m_pAuditThreadParameterMT )
            {
                pAuditThreadParameter = m_pAuditThreadParameterMT;
                m_pAuditThreadParameterMT = nullptr;
                DoAuditLocked( pAuditThreadParameter );
            }
        }
        catch( const Exception& e )
        {
            PXSLogException( e, __FUNCTION__ );
        }
        result = WaitForMultipleObjects( ARRAYSIZE( handles ), handles, FALSE, WAIT_TIMOUT_MS );
    }
    while ( m_bRunMT && ( result != WAIT_FAILED ) );

    return result;
}
