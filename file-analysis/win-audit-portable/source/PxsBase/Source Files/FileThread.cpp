///////////////////////////////////////////////////////////////////////////////////////////////////
//
// File Input/Output Thread Class Implementation
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Copyright 1987-2016 PARMAVEX SERVICES
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

// A FIFO queue for disk file I/O operations. A class scope buffer is used to
// transfer data to/from the disk file. The size is comparable to typical
// values of MaximumTransferlength of the host buffer adapter.
//
// Limitations
//  1. I/O is queued at the file level, i.e. one file must be completed before
//     the next begins.
//  2. The entire file needs to fit into memory
//  3. Results are only available on completion, not during progress.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/FileThread.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AutoUnlockMutex.h"
#include "PxsBase/Header Files/ByteArray.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
FileThread::FileThread()
           :IO_TYPE_NONE( 0 ),
            IO_TYPE_READ( 1 ),
            IO_TYPE_WRITE( 2 ),
            m_uTaskCounter( 0 ),
            TRANSFER_BUFFER_SIZE( 0x10000 ),    // 64K
            m_pTransferBuffer( nullptr ),
            m_Mutex(),
            m_TaskListMT()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
FileThread::~FileThread()
{
    if ( m_pTransferBuffer )
    {
        delete [] m_pTransferBuffer;
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
//      Cancel the specified I/O task
//
//  Parameters:
//      taskID - the task identifier
//
//  Remarks:
//      Called by worker: NO
//
//  Returns:
//      void
//===============================================================================================//
void FileThread::Cancel( DWORD taskID )
{
    TYPE_IO_TASK* pTask;

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    pTask = GetTaskFromIDLocked( taskID );
    if ( pTask == nullptr )
    {
        PXSLogAppInfo( L"Task was not cancelled because it was not found." );
        return;
    }

    // If the task is waiting can remove it otherwise set its state
    if ( pTask->state == PXS_STATE_WAITING )
    {
        m_TaskListMT.Remove();
    }
    else
    {
        pTask->state = PXS_STATE_CANCELLED;
    }
}

//===============================================================================================//
//  Description:
//      Get the result of the specified I/O task
//
//  Parameters:
//      taskID     - the task identifier
//      pBuffer    - receives the data, for write operations can be NULL
//      pException - receives any exception raised by the task
//
//  Remarks:
//      Called by worker: NO
//
//  Returns:
//      System error code
//===============================================================================================//
DWORD FileThread::GetResult( DWORD taskID,
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
        PXSLogAppInfo( L"No results because the task was not found." );
        return ERROR_INVALID_FUNCTION;
    }

    if ( ( pTask->state == PXS_STATE_WAITING   ) ||
         ( pTask->state == PXS_STATE_EXECUTING )  )
    {
        // No results yet as the task has not finished
        throw FunctionException( L"pTask->state", __FUNCTION__ );
    }

    // Get the data for read operations
    if ( pTask->state == PXS_STATE_COMPLETED )
    {
        if ( pTask->ioType == IO_TYPE_READ )
        {
            if ( pBuffer == nullptr )
            {
                throw ParameterException( L"pBuffer", __FUNCTION__ );
            }
            *pBuffer = pTask->Buffer;
        }
    }
    state       = pTask->state;
    *pException = pTask->e;
    m_TaskListMT.Remove();      // After removal, pTask is no longer valid

    if ( state == PXS_STATE_CANCELLED )
    {
        return ERROR_CANCELLED;
    }
    else if ( state == PXS_STATE_ERROR )
    {
        return pException->GetErrorCode();
    }

    return ERROR_SUCCESS;
}

//===============================================================================================//
//  Description:
//      Read the specified file
//
//  Parameters:
//      FilePath - the path to the file
//
//  Remarks:
//      Called by worker: NO
//
//  Returns:
//      the identifier of the task
//===============================================================================================//
size_t FileThread::Read( const String& FilePath )
{
    ByteArray Buffer;       // Empty buffer

    return AppendTask( IO_TYPE_READ, FilePath, Buffer );
}

//===============================================================================================//
//  Description:
//      Write to the specified file
//
//  Parameters:
//      FilePath     - the path to the file
//      WriteBuffer  - the bytes to write to the file
//
//  Remarks:
//      Called by worker: NO
//
//  Returns:
//      the identifier of the task
//===============================================================================================//
size_t FileThread::Write( const String& FilePath, const ByteArray& WriteBuffer )
{
    return AppendTask( IO_TYPE_WRITE, FilePath, WriteBuffer );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Append the specified bytes to the specified task's buffer
//
//  Parameters:
//      taskID   - the task identifier
//      pBuffer  - the bytes to append
//      numBytes - number of bytes to append
//      pState   - receives the state of the task;
//
//  Remarks:
//      Called by worker: YES
//
//  Returns:
//      void
//===============================================================================================//
void FileThread::AppendToTaskBuffer( DWORD taskID,
                                     const BYTE* pBuffer, size_t numBytes, DWORD* pState )
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
        *pState = PXS_STATE_UNKNOWN;
        return;
    }

    if ( pBuffer &&  numBytes )
    {
        pTask->Buffer.Append( pBuffer, numBytes );
    }
    *pState = pTask->state;
}

//===============================================================================================//
//  Description:
//      Append a task to the queue
//
//  Parameters:
//      ioType       - I/O operation type, e.g. read or write
//      FilePath     - the path to the file
//      WriteBuffer  - the bytes to write to the file, not used for read 
//                     operations
//
//  Remarks:
//      Called by worker: NO
//
//  Returns:
//      the unique task identifier
//===============================================================================================//
size_t FileThread::AppendTask( DWORD ioType, const String& FilePath, const ByteArray WriteBuffer )
{
    File FileObject;
    TYPE_IO_TASK Task;

    if ( m_hDoTaskEventMT == nullptr )
    {
        throw FunctionException( L"m_hDoTaskEventMT", __FUNCTION__ );
    }

    if ( ( ioType != IO_TYPE_READ ) && ( ioType != IO_TYPE_WRITE ) )
    {
        throw ParameterException( L"ioType", __FUNCTION__ );
    }

    if ( FilePath.GetLength() == 0 )
    {
        throw SystemException( ERROR_INVALID_NAME, L"FilePath.c_str()", __FUNCTION__ );
    }

    // For read operations need an existing file
    if ( ( ioType == IO_TYPE_READ ) &&
         ( FileObject.Exists( FilePath ) == false ) )
    {
        throw SystemException( ERROR_FILE_NOT_FOUND, FilePath.c_str(), __FUNCTION__ );
    }

    // Add the task to the thread's list
    m_uTaskCounter    = PXSAddUInt32( m_uTaskCounter, 1 );
    Task.state        = PXS_STATE_WAITING;
    Task.ioType       = ioType;
    Task.taskID       = m_uTaskCounter;
    Task.FilePath     = FilePath;
    if ( ioType == IO_TYPE_WRITE )
    {
        Task.Buffer = WriteBuffer;
    }
    Task.e.Reset();
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    m_TaskListMT.Append( Task );

    // Tell the worker there is something do do
    if ( SetEvent( m_hDoTaskEventMT ) == 0 )
    {
        PXSLogSysError( GetLastError(), L"SetEvent failed." );
    }

    return m_uTaskCounter;
}

//===============================================================================================//
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
//  Remarks:
//      Called by worker: YES
//
//  Returns:
//      void
//===============================================================================================//
DWORD FileThread::DoTask( DWORD taskID )
{
    DWORD  state    = PXS_STATE_EXECUTING, ioType = IO_TYPE_NONE;
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
        FileObject.OpenBinaryExclusive( FilePath );
        do
        {
            numBytes = FileObject.Read( m_pTransferBuffer, TRANSFER_BUFFER_SIZE );
            AppendToTaskBuffer( taskID, m_pTransferBuffer, numBytes, &state );
        }
        while ( m_bRunMT && numBytes && ( state == PXS_STATE_EXECUTING ) );
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
                                          m_pTransferBuffer, TRANSFER_BUFFER_SIZE, &state );
            FileObject.Write( m_pTransferBuffer, numBytes );
            idxOffset = PXSAddSizeT( idxOffset, numBytes );
        }
        while ( m_bRunMT && numBytes && ( state == PXS_STATE_EXECUTING ) );
        FileObject.Flush();
    }

    if ( m_bRunMT == FALSE )
    {
        state = PXS_STATE_CANCELLED;
    }
    else if ( state == PXS_STATE_EXECUTING )
    {
        state = PXS_STATE_COMPLETED;
    }

    return state;
}

//===============================================================================================//
//  Description:
//      Copy bytes from the task's buffer into the specified buffer
//
//  Parameters:
//      taskID    - the task identifier
//      idxOffset - zero based index from where to start
//      pBuffer   - receives the bytes
//      bufSize   - the size of the buffer
//      pState    - receives the state of the task;
//
//  Remarks:
//      Called by worker: YES
//
//  Returns:
//      Number of bytes copied from the task's buffer, zero if no more
//      bytes are available
//===============================================================================================//
size_t FileThread::GetFromTaskBuffer( DWORD  taskID,
                                      size_t idxOffset,
                                      BYTE*  pBuffer, size_t bufSize, DWORD* pState )
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
        *pState = PXS_STATE_UNKNOWN;
        return 0;
    }

    if ( pBuffer && bufSize )
    {
        bytesRead = pTask->Buffer.Get( idxOffset, pBuffer, bufSize );
    }
    *pState = pTask->state;

    return bytesRead;
}

//===============================================================================================//
//  Description:
//      Get the task corresponding to the task identifier. If the task is found
//      the list is positioned on it otherwise the list is on the tail.
//
//  Parameters:
//      taskID - the identifier
//
//  Remarks:
//      The caller must have locked the queue.
//
//  Remarks:
//      Called by worker: YES
//
//  Returns:
//      pointer to the task, NULL if not found
//===============================================================================================//
FileThread::TYPE_IO_TASK* FileThread::GetTaskFromIDLocked( DWORD taskID )
{
    bool found = false;
    TYPE_IO_TASK* pTask = nullptr;

    if ( m_TaskListMT.IsEmpty() )
    {
        return nullptr;
    }

    m_TaskListMT.Rewind();
    do
    {
        pTask = m_TaskListMT.GetPointer();
        if ( pTask && ( pTask->taskID == taskID ) )
        {
            found = true;
        }
    } while ( ( found == false ) && m_TaskListMT.Advance() );

    if ( found )
    {
        return pTask;
    }
    return nullptr;
}

//===============================================================================================//
//  Description:
//      Get the specified tasks' properties
//
//  Parameters:
//      taskID    - the task's identifier
//      pIoType   - optional, receives the I/O type
//      pFilePath - optional, receives the file path
//
//  Remarks:
//      Called by worker: YES
//
//  Returns:
//      named constant representing the state
//===============================================================================================//
void FileThread::GetTaskProperties( DWORD taskID, DWORD* pIoType, String* pFilePath )
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

//===============================================================================================//
//  Description:
//      Get the identifier of the next waiting task in the list/queue. Its
//      state is changed to STATE_EXECUTING.The list is postioned on the
//      item if found otherwise its on the tail.
//
//  Parameters:
//      pTaskID - receives the task
//
//  Remarks:
//      Called by worker: YES
//
//  Returns:
//      true if found a waiting task otherwise false
//===============================================================================================//
bool FileThread::GetNextWaitingTask( DWORD* pTaskID )
{
    bool found = false;
    TYPE_IO_TASK* pTask = nullptr;

    if ( pTaskID == nullptr )
    {
        throw ParameterException( L"pTaskID", __FUNCTION__ );
    }

    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );
    if ( m_TaskListMT.IsEmpty() )
    {
        return false;
    }

    m_TaskListMT.Rewind();
    do
    {
        pTask = m_TaskListMT.GetPointer();
        if ( pTask && ( pTask->state == PXS_STATE_WAITING ) )
        {
            found        = true;
            *pTaskID     = pTask->taskID;
            pTask->state = PXS_STATE_EXECUTING;
        }
    } while ( ( found == false ) && m_TaskListMT.Advance() );

    return found;
}

//===============================================================================================//
//  Description:
//      Run the worker thread. This method must only be called by the worker.
//
//  Parameters:
//      None
//
//  Remarks:
//      Called by worker: YES
//
//  Returns:
//      DWORD system error code
//===============================================================================================//
DWORD FileThread::RunWorkerThread()
{
    DWORD  WAIT_TIMOUT_MS = 250;
    DWORD  result    = WAIT_TIMEOUT, taskID = 0, state = 0;
    HANDLE handles[] = { m_hDoTaskEventMT, m_hStopEventMT };
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
                    UpdateTaskState( taskID, PXS_STATE_ERROR, &eTask );
                }

                if ( m_hWndListenerMT )
                {
                    PostMessage( m_hWndListenerMT,
                                 PXS_APP_MSG_FILE_IO_RESULT, taskID, 0 );
                }
            }
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

//===============================================================================================//
//  Description:
//      Update the specified task's state in the queue
//
//  Parameters:
//      taskID     - the task's identifier
//      state      - the new state
//      pException - optional, an exception assocated with the task
//
//  Remarks:
//      Called by worker: YES
//
//  Returns:
//      void
//===============================================================================================//
void FileThread::UpdateTaskState( DWORD taskID, DWORD state, const Exception* pException )
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
