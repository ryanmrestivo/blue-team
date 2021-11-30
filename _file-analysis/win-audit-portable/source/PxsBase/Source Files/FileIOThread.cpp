////////////////////////////////////////////////////////////////////////////////
//
// File Input/Output Thread Class Implementation
//
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
//
// Copyright 1987-2014 PARMAVEX SERVICES
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
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
// Remarks
////////////////////////////////////////////////////////////////////////////////

// Performs I/O an entire file, i.e. read/write an entire file
//
// TRANSFER_BUFFER_SIZE = 0x10000 as most disks can transfer at least 64K
// in a single step

////////////////////////////////////////////////////////////////////////////////
// Include Files
////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/FileIOThread.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AutoUnlockMutex.h"
#include "PxsBase/Header Files/ByteArray.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/SystemException.h"

////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
////////////////////////////////////////////////////////////////////////////////

// Default constructor
FileIOThread::FileIOThread()
             :IO_TYPE_NONE( 0 ),
              IO_TYPE_READ( 1 ),
              IO_TYPE_WRITE( 2 ),
              STATE_UNKOWN( 0 ),
              STATE_ERROR( 1 ),
              STATE_WAITING( 2 ),
              STATE_EXECUTING( 3 ),
              STATE_CANCELLED( 4 ),
              STATE_COMPLETED( 5 ),
              m_uTaskCounter( 0 ),
              TRANSFER_BUFFER_SIZE( 0x10000 ),
              m_Mutex(),
              m_pTransferBuffer( nullptr ),
              m_IOTaskListMT()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
FileIOThread::~FileIOThread()
{
    if ( m_pTransferBuffer )
    {
        delete [] m_pTransferBuffer;
    }
}

////////////////////////////////////////////////////////////////////////////////
// Operators
////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation

////////////////////////////////////////////////////////////////////////////////
// Public Methods
////////////////////////////////////////////////////////////////////////////////

//============================================================================//
//  Description:
//      Cancel the specified I/O task
//
//  Parameters:
//      taskID - the task identifier
//
//  Returns:
//      void
//============================================================================//
void FileIOThread::Cancel( size_t taskID )
{
    TYPE_IO_TASK* pTask;

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    pTask = GetTaskFromIDLocked( taskID );
    if ( pTask == nullptr )
    {
        // Unexpected
        PXSLogAppInfo( L"Task was not found for cancel operation." );
        return;
    }
    pTask->state = STATE_CANCELLED;
}

//============================================================================//
//  Description:
//      Get the result of the specified I/O task
//
//  Parameters:
//      taskID     - the task identifier
//      pBuffer    - receives the data, for write operations can be NULL
//      pException - receives any exception raised by the task
//
//  Returns:
//      System error code
//============================================================================//
DWORD FileIOThread::GetResult( size_t taskID,
                               ByteArray* pBuffer, Exception* pException )
{
    DWORD state;
    TYPE_IO_TASK* pTask;

    if ( pException == nullptr )
    {
        throw ParameterException( L"pException", __FUNCTION__ );
    }

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    pTask = GetTaskFromIDLocked( taskID );
    if ( pTask == nullptr )
    {
        // Unexpected
        PXSLogAppInfo( L"No results because the task was not found." );
        return ERROR_INVALID_FUNCTION;
    }

    // Copy data then remove the task
    state       = pTask->state;
    *pException = pTask->e;
    if ( pTask->ioType == IO_TYPE_READ )
    {
        if ( pBuffer == nullptr )
        {
            throw ParameterException( L"pBuffer", __FUNCTION__ );
        }
        *pBuffer = pTask->Buffer;
    }
    m_IOTaskListMT.Remove();        // after his pTask no longer valid

    if ( ( state == STATE_WAITING ) || ( state == STATE_EXECUTING ) )
    {
        return ERROR_INVALID_FUNCTION;      // No results as task not finished
    }
    else if ( state == STATE_CANCELLED )
    {
        return ERROR_CANCELLED;
    }
    else if ( state == STATE_ERROR )
    {
        return pException->GetErrorCode();
    }

    return ERROR_SUCCESS;
}

//============================================================================//
//  Description:
//      Read the specified file
//
//  Parameters:
//      FilePath     - the path to the file
//      hWndListener - the window to inform when the operation completes
//
//  Returns:
//      the identifier of the task
//============================================================================//
size_t FileIOThread::Read( const String& FilePath, HWND hWndListener )
{
    return AppendTask( IO_TYPE_READ, FilePath, nullptr, hWndListener );
}

//============================================================================//
//  Description:
//      Write to the specified file
//
//  Parameters:
//      FilePath     - the path to the file
//      WriteBuffer  - the bytes to write to the file
//      hWndListener - the window to inform when the operation completes
//
//  Returns:
//      the identifier of the task
//============================================================================//
size_t FileIOThread::Write( const String& FilePath,
                            const ByteArray& WriteBuffer, HWND hWndListener )
{
    return AppendTask( IO_TYPE_WRITE, FilePath, &WriteBuffer, hWndListener );
}

////////////////////////////////////////////////////////////////////////////////
// Protected Methods
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
// Private Methods
////////////////////////////////////////////////////////////////////////////////

//============================================================================//
//  Description:
//      Append the specified bytes to the specified task's buffer
//
//  Parameters:
//      taskID   - the task identifier
//      pBuffer  - the bytes to append
//      numBytes - number of bytes to append
//      pState   - receives the state if the task;
//
//  Returns:
//      void
//============================================================================//
void FileIOThread::AppendToTaskBuffer( DWORD taskID,
                                       const BYTE* pBuffer,
                                       size_t numBytes, DWORD* pState )
{
    TYPE_IO_TASK* pTask;

    if ( pState == nullptr )
    {
        throw ParameterException( L"pState", __FUNCTION__ );
    }

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    pTask = GetTaskFromIDLocked( taskID );
    if ( pTask == nullptr )
    {
        *pState = STATE_UNKOWN;
        return;
    }

    if ( pBuffer &&  numBytes )
    {
        pTask->Buffer.Append( pBuffer, numBytes );
    }
    *pState = pTask->state;
}

//============================================================================//
//  Description:
//      Append a taskm to the queue
//
//  Parameters:
//      ioType       - I/O operation type, e.g. read or write
//      FilePath     - the path to the file
//      WriteBuffer  - the bytes to write to the file, not used for read 
//                     operations
//      hWndListener - the window to inform when the operation completes
//
//  Returns:
//      the unique task identifier
//============================================================================//
size_t FileIOThread::AppendTask( DWORD ioType,
                                 const String& FilePath,
                                 const ByteArray* pWriteBuffer,
                                 HWND hWndListener )
{
    File FileObject;
    TYPE_IO_TASK Task;

    if ( m_hDoTaskEvent == nullptr )
    {
        throw FunctionException( L"m_hDoTaskEvent" , __FUNCTION__ );
    }

    if ( ( ioType != IO_TYPE_READ ) && ( ioType != IO_TYPE_WRITE ) )
    {
        throw ParameterException( L"ioType", __FUNCTION__ );
    }

    if ( FilePath.GetLength() == 0 )
    {
        throw SystemException( ERROR_INVALID_NAME,
                               L"FilePath.c_str()", __FUNCTION__ );
    }

    // For read operations need an existing file
    if ( ( ioType == IO_TYPE_READ ) &&
         ( FileObject.Exists( FilePath ) == false ) )
    {
        throw SystemException( ERROR_FILE_NOT_FOUND,
                               FilePath.c_str(), __FUNCTION__ );
    }

    // For write operations need a buffer
    if ( ( ioType == IO_TYPE_WRITE ) && ( pWriteBuffer == nullptr ) )
    {
        throw ParameterException( L"pWriteBuffer", __FUNCTION__ );
    }

    if ( hWndListener == nullptr )
    {
        throw ParameterException( L"hWndListener", __FUNCTION__ );
    }

    // Add the task to the thread's list
    m_uTaskCounter      = PXSAddUInt32( m_uTaskCounter, 1 );
    Task.state        = STATE_WAITING;
    Task.ioType       = ioType;
    Task.taskID       = m_uTaskCounter;
    Task.FilePath     = FilePath;
    if ( pWriteBuffer )
    {
        Task.Buffer   = *pWriteBuffer;
    }
    Task.e.Reset();
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_IOTaskListMT.Append( Task );

    // Tell the worker there is something do do
    if ( SetEvent( m_hDoTaskEvent ) == 0 )
    {
        PXSLogSysError( GetLastError(), L"SetEvent failed." );
    }

    return m_uTaskCounter;
}

//============================================================================//
//  Description:
//      Do the specified task
//
//  Parameters:
//      taskID - the task idntifier
//
//  Remarks:
//      The caller must pass in a copy of the task as this this method has no
//      locking. Doing so would mean the queue would be waiting for this method
//      to complete before being able to add/remove a task.
//
//  Returns:
//      void
//============================================================================//
DWORD FileIOThread::DoTask( size_t taskID )
{
    DWORD  state    = STATE_EXECUTING, ioType = IO_TYPE_NONE;
    size_t numBytes = 0, idxOffset = 0;
    File   FileObject;
    String FilePath;

    // Create on demand a transfer buffer
    if ( m_pTransferBuffer == nullptr )
    {
        m_pTransferBuffer = new BYTE[ TRANSFER_BUFFER_SIZE ];
        if ( m_pTransferBuffer == nullptr )
        {
            throw MemoryException( __FUNCTION__ );
        }
        memset( m_pTransferBuffer, 0, TRANSFER_BUFFER_SIZE );
    }

    // Do the I/O in chunks while polling for a stop or STATE_CANCELLED
    GetTaskProperties( taskID, &ioType, &FilePath );
    if ( ioType == IO_TYPE_READ )
    {
        PXSLogAppInfo1( L"Reading file '%%1'", FilePath );
        FileObject.OpenBinary( FilePath );
        do
        {
            numBytes = FileObject.Read( m_pTransferBuffer,
                                        TRANSFER_BUFFER_SIZE );
            AppendToTaskBuffer( taskID, m_pTransferBuffer, numBytes, &state );
        }
        while ( m_bRunMT && numBytes && ( state == STATE_EXECUTING ) );
    }
    else if ( ioType == IO_TYPE_WRITE )
    {
        // Will permit create on empty buffer to allow for a null write.
        PXSLogAppInfo1( L"Writing file '%%1'", FilePath );
        FileObject.CreateNew( FilePath,
                              0,            // Exclusive access
                              false );      // Assume writing binary data
        do
        {
            numBytes = GetFromTaskBuffer( taskID,
                                          idxOffset,
                                          m_pTransferBuffer,
                                          TRANSFER_BUFFER_SIZE, &state );
            FileObject.Write( m_pTransferBuffer, numBytes );
            idxOffset = PXSAddSizeT( idxOffset, numBytes );
        }
        while ( m_bRunMT && numBytes && ( state == STATE_EXECUTING ) );
        FileObject.Flush();
    }

    // Update the state but preserve STATE_CANCELLED if was set
    if ( m_bRunMT == FALSE )
    {
        state = STATE_CANCELLED;
    }
    else if ( state == STATE_EXECUTING )
    {
        state = STATE_COMPLETED;
    }

    return state;
}

//============================================================================//
//  Description:
//      Copy bytes from the task's buffer into the specified buffer
//
//  Parameters:
//      taskID    - the task identifier
//      idxOffset - zero based index from where to start
//      pBuffer   - receives the bytes
//      bufSize   - the size of the buffer
//      pState    - receives the state if the task;
//
//  Returns:
//      Number of bytes copied from the task's buffer, zero if no more
//      bytes are available
//============================================================================//
size_t FileIOThread::GetFromTaskBuffer( DWORD  taskID,
                                        size_t idxOffset,
                                        BYTE*  pBuffer,
                                        size_t bufSize, DWORD* pState )
{
    size_t bytesRead = 0;
    TYPE_IO_TASK* pTask;

    if ( pState == nullptr )
    {
        throw ParameterException( L"pState", __FUNCTION__ );
    }

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    pTask = GetTaskFromIDLocked( taskID );
    if ( pTask == nullptr )
    {
        *pState = STATE_UNKOWN;
        return 0;
    }

    if ( pBuffer && bufSize )
    {
        bytesRead = pTask->Buffer.Get( idxOffset, pBuffer, bufSize );
    }
    *pState = pTask->state;

    return bytesRead;
}

//============================================================================//
//  Description:
//      Get the task corresponding to the task identifier
//
//  Parameters:
//      taskID - the identifier
//
//  Remarks:
//      The caller must have locked the queue.
//      Side effect: The list is postioned on the item if found otherwise its
//      on the tail.
//
//  Returns:
//      pointer to the task, NULL if not found
//============================================================================//
FileIOThread::TYPE_IO_TASK* FileIOThread::GetTaskFromIDLocked( DWORD taskID )
{
    bool found = false;
    TYPE_IO_TASK* pTask = nullptr;

    if ( m_IOTaskListMT.IsEmpty() )
    {
        return nullptr;
    }

    // If the task has not yet started it will in the queue, so remove it.
    m_IOTaskListMT.Rewind();
    do
    {
        pTask = m_IOTaskListMT.GetPointer();
        if ( pTask && ( pTask->taskID == taskID ) )
        {
            found = true;
        }
    } while ( ( found == false ) && m_IOTaskListMT.Advance() );

    if ( found )
    {
        return pTask;
    }
    return nullptr;
}

//============================================================================//
//  Description:
//      Get the specified tasks' properties
//
//  Parameters:
//      taskID    - the task's identifier
//      pIoType   - optional, receives the I/O type
//      pFilePath - optional, receives the file path
//
//  Returns:
//      named constant representing the state
//============================================================================//
void FileIOThread::GetTaskProperties( DWORD taskID,
                                      DWORD* pIoType, String* pFilePath )
{
    TYPE_IO_TASK* pTask;

    // Reset
    if ( pIoType   ) *pIoType   = IO_TYPE_NONE;
    if ( pFilePath ) *pFilePath = PXS_STRING_EMPTY;

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    pTask = GetTaskFromIDLocked( taskID );
    if ( pTask == nullptr )
    {
        return;
    }
    *pIoType   = pTask->ioType;
    *pFilePath = pTask->FilePath;
}

//============================================================================//
//  Description:
//      Get the state of the specified task
//
//  Parameters:
//      taskID - the task's identifier
//
//  Returns:
//      named constant representing the state
//============================================================================//
DWORD FileIOThread::GetTaskState( DWORD taskID )
{
    TYPE_IO_TASK* pTask;

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    pTask = GetTaskFromIDLocked( taskID );
    if ( pTask == nullptr )
    {
        return STATE_UNKOWN;
    }

    return pTask->taskID;
}

//============================================================================//
//  Description:
//      Get the identifier of the next waiting task in the list/queue. Its
//      state is changed to STATE_EXECUTING.
//
//  Parameters:
//      pTaskID - receives the task
//
//  Remarks:
//      Side effect: The list is postioned on the item if found otherwise its
//      on the tail.
//
//  Returns:
//      true if found a waiting task otherwise false
//============================================================================//
bool FileIOThread::GetNextWaitingTask( DWORD* pTaskID )
{
    bool found = false;
    TYPE_IO_TASK* pTask = nullptr;

    if ( pTaskID == nullptr )
    {
        throw ParameterException( L"pTaskID", __FUNCTION__ );
    }

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    if ( m_IOTaskListMT.IsEmpty() )
    {
        return false;
    }

    m_IOTaskListMT.Rewind();
    do
    {
        pTask = m_IOTaskListMT.GetPointer();
        if ( pTask && ( pTask->state == STATE_WAITING ) )
        {
            found        = true;
            *pTaskID     = pTask->taskID;
            pTask->state = STATE_EXECUTING;
        }
    } while ( ( found == false ) && m_IOTaskListMT.Advance() );

    return found;
}

//============================================================================//
//  Description:
//      Run the worker thread. This method must only be called by the worker.
//
//  Parameters:
//      None
//
//  Returns:
//      DWORD system error code
//============================================================================//
DWORD FileIOThread::RunThread()
{
    DWORD  WAIT_TIMOUT_MS = 1000;
    DWORD  result    = WAIT_TIMEOUT, taskID = 0, state = 0;
    HANDLE handles[] = { m_hDoTaskEvent, m_hStopEvent };
    File   FileObject;

    if ( ( handles[ 0 ] == nullptr ) || ( handles[ 1 ] == nullptr ) )
    {
        return ERROR_INVALID_FUNCTION;
    }

    // Loop until told to stop
    do
    {
        try     // Catch all as want to continue to next task
        {
            while ( GetNextWaitingTask( &taskID ) )
            {
                try
                {
                    state = DoTask( taskID );
                    UpdateTaskState( taskID, state, nullptr );
                }
                catch ( Exception& eTask )
                {
                    UpdateTaskState( taskID, STATE_ERROR, &eTask );
                }

                if ( m_hWndListenerMT )
                {
                    PostMessage( m_hWndListenerMT,
                                 PXS_APP_MSG_FILE_IO_RESULT, taskID, 0 );
                }
            };
        }
        catch( Exception& e )
        {
            PXSLogException( e, __FUNCTION__ );
        }
        result = WaitForMultipleObjects( ARRAYSIZE( handles ),
                                         handles, FALSE, WAIT_TIMOUT_MS );
    }
    while ( m_bRunMT && ( result != WAIT_FAILED ) );

    return result;
}

//============================================================================//
//  Description:
//      Update the specified task's state in the queue
//
//  Parameters:
//      taskID     - the task's identifier
//      state      - the new state
//      pException - optional, an exception assocated with the task
//
//  Returns:
//      void
//============================================================================//
void FileIOThread::UpdateTaskState( DWORD taskID,
                                    DWORD state, const Exception* pException )
{
    TYPE_IO_TASK* pTask;

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    pTask = GetTaskFromIDLocked( taskID );
    if ( pTask == nullptr )
    {
        PXSLogAppInfo( L"Did not update a task because it was not found." );
        return;
    }
    pTask->state = state;
    pTask->e.Reset();
    if ( pException )
    {
        pTask->e = *pException;
    }
}
