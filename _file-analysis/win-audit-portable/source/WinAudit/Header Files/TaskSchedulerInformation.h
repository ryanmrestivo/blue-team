///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Task Scheduler Information Class Header
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

#ifndef WINAUDIT_TASK_SCHEDULER_INFORMATION_H_
#define WINAUDIT_TASK_SCHEDULER_INFORMATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files
#include <MSTask.h>
#include <taskschd.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards
class AuditRecord;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class TaskSchedulerInformation
{
    public:
        // Default constructor
        TaskSchedulerInformation();

        // Destructor
        ~TaskSchedulerInformation();

        // Methods
        void  GetScheduledTaskRecords( TArray< AuditRecord >* pRecords );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        TaskSchedulerInformation( const TaskSchedulerInformation& oTaskInfo );

        // Operators - not allowed
        TaskSchedulerInformation& operator= ( const TaskSchedulerInformation& oTaskInfo );

        // Methods
        void GetExecActionProperties( IAction* pAction, String* pExePath, String* pExeArgs );
        void GetRegisteredTaskMaxRunTime( IRegisteredTask* pRegisteredTask, String* pMaxRunTime );
        void GetRegisteredTaskRunCmd( IRegisteredTask* pRegisteredTask, String* pRunCmd );
        void GetRegisteredTaskSchedule( IRegisteredTask* pRegisteredTask, String* pSchedule );
        void GetRegisteredTaskProperties( IRegisteredTask* pRegisteredTask,
                                          String* pName,
                                          String* pState,
                                          String* pSchedule,
                                          String* pIsoNextRun,
                                          String* pRunCmd,
                                          String* pMaxRunTime,
                                          String* pIsoLastRun, String* pExitCode );
        void GetTaskProperties( ITask* pITask ,
                                String* pState,
                                String* pSchedule,
                                String* pIsoNextRun,
                                String* pRunCmd,
                                String* pMaxRunTime, String* pIsoLastRun, String* pxitCode );
        void GetTasksInFolderRecords( ITaskService* pTaskService,
                                      BSTR path, TArray< AuditRecord >* pRecords );
        void GetType1TaskRecords( TArray< AuditRecord >* pRecords );
        void GetType2TaskRecords( TArray< AuditRecord >* pRecords );
        void TranslateScheduledWorkItemStatus( HRESULT hStatus, String* pStatus );
        void TranslateTaskActionType( TASK_ACTION_TYPE type, String* pActionType );
        void TranslateTaskState( TASK_STATE state, String* pTaskState );
        void TranslateTaskTriggerType2( TASK_TRIGGER_TYPE2 type2, String* pTriggerType2 );

        // Data members
};

#endif  // WINAUDIT_TASK_SCHEDULER_INFORMATION_H_
