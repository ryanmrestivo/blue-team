///////////////////////////////////////////////////////////////////////////////////////////////////
//
// File Input/Output Thread Class Header
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

#ifndef PXSBASE_FILE_IO_THREAD_H_
#define PXSBASE_FILE_IO_THREAD_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/Thread.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/ByteArray.h"
#include "PxsBase/Header Files/Mutex.h"
#include "PxsBase/Header Files/Object.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/TList.h"

// 6. Forwards
class ByteBuffer;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class FileThread : public Thread
{
    public:
        // Default constructor
        FileThread();

        // Destructor
        ~FileThread();

        // Methods
        void    Cancel( DWORD taskID );
        DWORD   GetResult( DWORD taskID, ByteArray* pBuffer, Exception* pException );
        size_t  Read( const String& FilePath );
        size_t  Write( const String& FilePath, const ByteArray& WriteBuffer );

    protected:
        // Methods

        // Data members

    private:

        typedef struct _TYPE_IO_TASK
        {
            DWORD       state;
            DWORD       ioType;
            DWORD       taskID;
            String      FilePath;
            ByteArray   Buffer;
            Exception   e;

            _TYPE_IO_TASK::_TYPE_IO_TASK()
                          :state( 0 ),
                           ioType( 0 ),
                           taskID( 0 ),
                           FilePath(),
                           Buffer(),
                           e()
            {
            }
        }
        TYPE_IO_TASK;

        // Copy constructor - not allowed
        FileThread( const FileThread& oFileThread );

        // Assignment operator - not allowed
        FileThread& operator= ( const FileThread& oFileThread );

        // Methods
        void    AppendToTaskBuffer( DWORD  taskID,
                                    const  BYTE* pBuffer, size_t numBytes, DWORD* pState );
        size_t  AppendTask( DWORD ioType,
                            const String& FilePath, const ByteArray WriteBuffer );
        DWORD   DoTask( DWORD taskID );
        size_t  GetFromTaskBuffer( DWORD  taskID,
                                   size_t idxOffset,
                                   BYTE*  pBuffer, size_t bufSize, DWORD* pState );
        bool    GetNextWaitingTask( DWORD* pTaskID );
 TYPE_IO_TASK*  GetTaskFromIDLocked( DWORD taskID );
        void    GetTaskProperties( DWORD taskID, DWORD* pIoType, String* pFilePath );
        DWORD   RunWorkerThread();
        void    UpdateTaskState( DWORD taskID, DWORD state, const Exception* pException );

        // Data members
        DWORD   IO_TYPE_NONE;
        DWORD   IO_TYPE_READ;
        DWORD   IO_TYPE_WRITE;
        DWORD   m_uTaskCounter;                 // Accessed by creator only
        size_t  TRANSFER_BUFFER_SIZE;
        BYTE*   m_pTransferBuffer;              // Accessed by worker only
        Mutex   m_Mutex;
        TList< TYPE_IO_TASK > m_TaskListMT;     // Multi-threaded member
};

#endif  // PXSBASE_FILE_IO_THREAD_H_
