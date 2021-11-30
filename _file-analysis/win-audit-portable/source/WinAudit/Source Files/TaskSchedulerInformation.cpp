///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Task Scheduler Information Implementation
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

// Vista introduced a new interface for the Task Scheduler so have to use
// interfaces for v1.0 and v2.0 as appropriate.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/TaskSchedulerInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AutoIUnknownRelease.h"
#include "PxsBase/Header Files/AutoSysFreeString.h"
#include "PxsBase/Header Files/BStr.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
TaskSchedulerInformation::TaskSchedulerInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
TaskSchedulerInformation::~TaskSchedulerInformation()
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
//      Get the scheduled tasks as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetScheduledTaskRecords( TArray< AuditRecord >* pRecords )
{
    WindowsInformation WindowsInfo;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // ITask2 or ITask
    if ( WindowsInfo.GetMajorVersion() >= 6 )
    {
        GetType2TaskRecords( pRecords );
    }
    else
    {
        GetType1TaskRecords( pRecords );
    }
    PXSSortAuditRecords( pRecords, PXS_SCHED_TASKS_TASK_NAME );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the executable properties of specified action
//
//  Parameters:
//      pAction  - the action, must be of type TASK_ACTION_EXEC
//      pExePath - receives the executable path
//      pExeArgs - receives the executable arguments
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetExecActionProperties( IAction* pAction,
                                                        String* pExePath, String* pExeArgs )
{
    BSTR     path     = nullptr;
    BSTR     argument = nullptr;
    HRESULT  hResult  = S_OK;
    IExecAction*        pExecAction = nullptr;
    TASK_ACTION_TYPE    actionType;
    AutoIUnknownRelease ReleaseExecAction;

    if ( pAction == nullptr )
    {
        throw ParameterException( L"pAction", __FUNCTION__ );
    }

    if ( ( pExePath == nullptr ) || ( pExeArgs == nullptr ) )
    {
        throw ParameterException( L"pExePath/pExeArgs", __FUNCTION__ );
    }
    *pExePath = PXS_STRING_EMPTY;
    *pExeArgs = PXS_STRING_EMPTY;

    // Verify the action type
    hResult = pAction->get_Type( &actionType );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IAction::get_Type", __FUNCTION__ );
    }

    if ( actionType != TASK_ACTION_EXEC )
    {
        throw ParameterException( L"pAction", __FUNCTION__ );
    }

    hResult = pAction->QueryInterface(IID_IExecAction, reinterpret_cast<void**>(&pExecAction) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IAction::QueryInterface", __FUNCTION__ );
    }

    if ( pExecAction == nullptr )
    {
        throw NullException( L"pExecAction", __FUNCTION__ );
    }
    ReleaseExecAction.Set( pExecAction );

    // Get the exe path
    hResult = pExecAction->get_Path( &path );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IExecAction::get_Path", __FUNCTION__ );
    }
    AutoSysFreeString AutoSysFreePath( path );
    *pExePath = path;

    // Get the arguments
    hResult = pExecAction->get_Arguments( &argument );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IExecAction::get_Arguments", __FUNCTION__ );
    }
    AutoSysFreeString AutoSysFreeArgument( argument );
    *pExeArgs = argument;
}

//===============================================================================================//
//  Description:
//      Get the properties of the specified registered task
//
//  Parameters:
//      pRegisteredTask - the registered task
//      pName           - receives the task's name
//      pStatus         - receives the task's status
//      pSchedule       - receives the task's run schedule
//      pIsoNextRun     - receives the task's time of next run
//      pRunCmd         - receives the task's executable command
//      pMaxRunTime     - receives the task's maximum allowed run time
//      pIsoLastRun     - receives the task's time of the last run
//      pExitCode       - receives the task's the last run's exit code
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetRegisteredTaskProperties( IRegisteredTask* pRegisteredTask,
                                                            String* pName,
                                                            String* pState,
                                                            String* pSchedule,
                                                            String* pIsoNextRun,
                                                            String* pRunCmd,
                                                            String* pMaxRunTime,
                                                            String* pIsoLastRun,
                                                            String* pExitCode )
{
    BSTR       bstrName    = nullptr;
    DATE       nextRunTime = 0.0, lastRunTime = 0.0;
    HRESULT    hResult     = 0, hLastTaskResult = 0;
    Formatter  Format;
    TASK_STATE taskState;

    *pName       = PXS_STRING_EMPTY;
    *pState      = PXS_STRING_EMPTY;
    *pSchedule   = PXS_STRING_EMPTY;
    *pIsoNextRun = PXS_STRING_EMPTY;
    *pRunCmd     = PXS_STRING_EMPTY;
    *pMaxRunTime = PXS_STRING_EMPTY;
    *pIsoLastRun = PXS_STRING_EMPTY;
    *pExitCode   = PXS_STRING_EMPTY;
    if ( pRegisteredTask == nullptr )
    {
        return;
    }

    // Name
    hResult = pRegisteredTask->get_Name( &bstrName );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_Name", __FUNCTION__ );
    }
    AutoSysFreeString AutoSysFreeName( bstrName );
    *pName = bstrName;

    // Status
    hResult = pRegisteredTask->get_State( &taskState );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_State", __FUNCTION__ );
    }
    TranslateTaskState( taskState, pState );

    // Schedule
    GetRegisteredTaskSchedule( pRegisteredTask, pSchedule );

    // Next run time
    hResult = pRegisteredTask->get_NextRunTime( &nextRunTime );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_NextRunTime", __FUNCTION__ );
    }
    *pIsoNextRun = Format.OleTimeInIsoFormat( nextRunTime );

    // Run command
    GetRegisteredTaskRunCmd( pRegisteredTask, pRunCmd );

    // Maximum run time
    GetRegisteredTaskMaxRunTime( pRegisteredTask, pMaxRunTime );

    // Last run
    hResult = pRegisteredTask->get_LastRunTime( &lastRunTime );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_LastRunTime", __FUNCTION__ );
    }
    *pIsoLastRun = Format.OleTimeInIsoFormat( lastRunTime );

    // Exit code
    hResult  = pRegisteredTask->get_LastTaskResult( &hLastTaskResult );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_LastTaskResult", __FUNCTION__ );
    }
    *pExitCode = Format.Int32( hLastTaskResult );
}

//===============================================================================================//
//  Description:
//      Get the maximum run time of the specified task
//
//  Parameters:
//      pRegisteredTask - the task
//      pMaxRunTime     - receives the maximum run time
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetRegisteredTaskMaxRunTime( IRegisteredTask* pRegisteredTask,
                                                            String* pMaxRunTime )
{
    BSTR     executionTimeLimit = nullptr;
    HRESULT  hResult = 0;
    ITaskSettings*      pTaskSettings   = nullptr;
    ITaskDefinition*    pTaskDefinition = nullptr;
    AutoIUnknownRelease ReleaseTaskDefinition, ReleaseTaskSettings;

    if ( pRegisteredTask == nullptr )
    {
        throw ParameterException( L"pRegisteredTask", __FUNCTION__ );
    }

    if ( pMaxRunTime == nullptr )
    {
        throw ParameterException( L"pMaxRunTime", __FUNCTION__ );
    }
    *pMaxRunTime = PXS_STRING_EMPTY;

    // Get the task's definition
    hResult = pRegisteredTask->get_Definition( &pTaskDefinition );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_Definition", __FUNCTION__ );
    }

    if ( pTaskDefinition == nullptr )
    {
        throw NullException( L"pTaskDefinition",  __FUNCTION__ );
    }
    ReleaseTaskDefinition.Set( pTaskDefinition );

    // Get the settings
    hResult = pTaskDefinition->get_Settings( &pTaskSettings );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITaskDefinition::get_Settings", __FUNCTION__ );
    }

    if ( pTaskSettings == nullptr )
    {
        throw NullException( L"pTaskSettings",  __FUNCTION__ );
    }
    ReleaseTaskSettings.Set( pTaskSettings );

    hResult = pTaskSettings->get_ExecutionTimeLimit( &executionTimeLimit );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITaskSettings::get_ExecutionTimeLimit", __FUNCTION__ );
    }
    AutoSysFreeString AutoSysFreeExecutionTimeLimit( executionTimeLimit );
    *pMaxRunTime = executionTimeLimit;
}

//===============================================================================================//
//  Description:
//      Get the run command  for the specified task
//
//  Parameters:
//      pRegisteredTask - the task
//      pRunCmd         - receives the run command
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetRegisteredTaskRunCmd( IRegisteredTask* pRegisteredTask,
                                                        String* pRunCmd )
{
    long    i = 0, count = 0;
    String  ExePath, ExeArgs, ActionType;
    HRESULT hResult = 0;
    TASK_ACTION_TYPE    type;
    IAction*            pAction           = nullptr;
    ITaskDefinition*    pTaskDefinition   = nullptr;
    IActionCollection*  pActionCollection = nullptr;
    AutoIUnknownRelease ReleaseActionCollection, ReleaseTaskDefinition;

    if ( pRunCmd == nullptr )
    {
        throw ParameterException( L"pRunCmd", __FUNCTION__ );
    }
    *pRunCmd = PXS_STRING_EMPTY;

    if ( pRegisteredTask == nullptr )
    {
        return;
    }

    hResult = pRegisteredTask->get_Definition( &pTaskDefinition );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_Definition", __FUNCTION__ );
    }

    if ( pTaskDefinition == nullptr )
    {
        throw NullException( L"pTaskDefinition",  __FUNCTION__ );
    }
    ReleaseTaskDefinition.Set( pTaskDefinition );

    hResult = pTaskDefinition->get_Actions( &pActionCollection );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_Actions", __FUNCTION__ );
    }

    if ( pActionCollection == nullptr )
    {
        throw NullException( L"pActionCollection",  __FUNCTION__ );
    }
    ReleaseActionCollection.Set( pActionCollection );

    hResult = pActionCollection->get_Count( &count );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IActionCollection::get_Count", __FUNCTION__ );
    }

    // This is a cumulative string
    for ( i = 0; i < count; i++ )
    {
        AutoIUnknownRelease ReleaseAction;

        // Get the Action
        pAction = nullptr;
        hResult = pActionCollection->get_Item( i + 1,   // Counting is 1-based
                                               &pAction );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"IActionCollection::get_Item", __FUNCTION__ );
        }

        if ( pAction == nullptr )
        {
            throw NullException( L"pAction",  __FUNCTION__ );
        }
        ReleaseAction.Set( pAction );

        hResult = pAction->get_Type( &type );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"IAction::get_Type", __FUNCTION__ );
        }

        // If its to launch a programme get the exe and its arguments
        if ( type == TASK_ACTION_EXEC )
        {
            GetExecActionProperties( pAction, &ExePath, &ExeArgs );
            if ( pRunCmd->GetLength() )
            {
                *pRunCmd += L", ";
            }
            *pRunCmd += ExePath;
            *pRunCmd += PXS_CHAR_SPACE;
            *pRunCmd += ExeArgs;
        }
        else
        {
            // ... otherwise translate
            TranslateTaskActionType( type, &ActionType );
            if ( pRunCmd->GetLength() )
            {
                *pRunCmd += L", ";
            }
            *pRunCmd += ActionType;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the scheduled for the specified task
//
//  Parameters:
//      pRegisteredTask - the task
//      pSchedule       - receives the schedule
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetRegisteredTaskSchedule( IRegisteredTask* pRegisteredTask,
                                                          String* pSchedule )
{
    long    i = 0, count = 0;
    String  TriggerType2;
    HRESULT hResult = 0;
    TASK_TRIGGER_TYPE2  type;
    ITrigger*           pTrigger           = nullptr;
    ITaskDefinition*    pTaskDefinition    = nullptr;
    ITriggerCollection* pTriggerCollection = nullptr;
    AutoIUnknownRelease ReleaseTaskDefinition;
    AutoIUnknownRelease ReleaseTriggerCollection;

    if ( pSchedule == nullptr )
    {
        throw ParameterException( L"pSchedule", __FUNCTION__ );
    }
    *pSchedule = PXS_STRING_EMPTY;

    if ( pRegisteredTask == nullptr )
    {
        return;
    }

    hResult = pRegisteredTask->get_Definition( &pTaskDefinition );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRegisteredTask::get_Definition", __FUNCTION__ );
    }

    if ( pTaskDefinition == nullptr )
    {
        throw NullException( L"pTaskDefinition",  __FUNCTION__ );
    }
    ReleaseTaskDefinition.Set( pTaskDefinition );

    // Get this task's triggers. Task scheduler 2.0, does not have a method
    // to get a description since there could be multiple triggers doing
    // different actions that may depend on the system state. Will just make
    // a simple description
    hResult = pTaskDefinition->get_Triggers( &pTriggerCollection );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITaskDefinition::get_Triggers", __FUNCTION__ );
    }

    if ( pTriggerCollection == nullptr )
    {
        throw NullException( L"pTriggerCollection",  __FUNCTION__ );
    }
    ReleaseTriggerCollection.Set( pTriggerCollection );

    hResult = pTriggerCollection->get_Count( &count );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITriggerCollection::get_Count", __FUNCTION__ );
    }

    for ( i = 0; i < count; i++ )
    {
        AutoIUnknownRelease ReleaseTrigger;

        pTrigger = nullptr;
        hResult  = pTriggerCollection->get_Item( i + 1,     // 1-based
                                                 &pTrigger );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"ITriggerCollection::get_Item", __FUNCTION__ );
        }

        if ( pTrigger == nullptr )
        {
            throw NullException( L"pTrigger",  __FUNCTION__ );
        }
        ReleaseTrigger.Set( pTrigger );

        hResult = pTrigger->get_Type( &type );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"ITrigger::get_Type", __FUNCTION__ );
        }
        TriggerType2 = PXS_STRING_EMPTY;
        TranslateTaskTriggerType2( type, &TriggerType2 );

        // Make a comma separated string
        if ( pSchedule->GetLength() )
        {
            *pSchedule += L", ";
        }
        *pSchedule += TriggerType2;
    }
}

//===============================================================================================//
//  Description:
//      Get the properties of the specified task
//
//  Parameters:
//      pITask      - the task
//      pState      - receives the task's status
//      pSchedule   - receives the task's run schedule
//      pIsoNextRun - receives the task's time of next run
//      pRunCmd     - receives the task's executable command
//      pMaxRunTime - receives the task's maximum allowed run time
//      pIsoLastRun - receives the task's time of the last run
//      pExitCode   - receives the task's the last run's exit code
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetTaskProperties( ITask* pITask ,
                                                  String* pState,
                                                  String* pSchedule,
                                                  String* pIsoNextRun,
                                                  String* pRunCmd,
                                                  String* pMaxRunTime,
                                                  String* pIsoLastRun,
                                                  String* pExitCode )
{
    WORD     i = 0, lCount = 0;
    DWORD    dwMaxRunTime  = 0, dwExitCode = 0;
    String   Trigger;
    wchar_t* pwszTrigger         = nullptr;
    wchar_t* pwszParameters      = nullptr;
    wchar_t* pwszApplicationName = nullptr;
    HRESULT  hResult = 0, hrStatus = 0;
    Formatter  Format;
    SYSTEMTIME stNextRun, stLastRun;
    IScheduledWorkItem* pIScheduledWorkItem = nullptr;
    AutoIUnknownRelease ReleaseIScheduledWorkItem;

    *pState      = PXS_STRING_EMPTY;
    *pSchedule   = PXS_STRING_EMPTY;
    *pIsoNextRun = PXS_STRING_EMPTY;
    *pRunCmd     = PXS_STRING_EMPTY;
    *pMaxRunTime = PXS_STRING_EMPTY;
    *pIsoLastRun = PXS_STRING_EMPTY;
    *pExitCode   = PXS_STRING_EMPTY;
    if ( pITask == nullptr )
    {
        return;
    }

    hResult = pITask->QueryInterface( IID_IScheduledWorkItem,
                                      reinterpret_cast<void**>( &pIScheduledWorkItem ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITask::QueryInterface", __FUNCTION__ );
    }

    if ( pIScheduledWorkItem == nullptr )
    {
        throw NullException( L"pIScheduledWorkItem",  __FUNCTION__ );
    }
    ReleaseIScheduledWorkItem.Set( pIScheduledWorkItem );

    // Status
    hResult = pIScheduledWorkItem->GetStatus( &hrStatus );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IScheduledWorkItem::GetStatus", __FUNCTION__ );
    }
    TranslateScheduledWorkItemStatus( hrStatus, pState );

    // Schedule
    hResult = pIScheduledWorkItem->GetTriggerCount( &lCount );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IScheduledWorkItem::GetTriggerCount", __FUNCTION__ );
    }

    for ( i = 0; i < lCount; i++ )
    {
        pwszTrigger = nullptr;
        hResult = pIScheduledWorkItem->GetTriggerString( 0, &pwszTrigger );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"IScheduledWorkItem::GetTriggerString", __FUNCTION__ );
        }

        if ( pwszTrigger )
        {
            // Need to call CoTaskMemFree
            try
            {
                Trigger = pwszTrigger;
            }
            catch ( const Exception& )
            {
                CoTaskMemFree( pwszTrigger );
                throw;
            }
            CoTaskMemFree( pwszTrigger );
            pwszTrigger = nullptr;  // Reset

            if ( pSchedule->GetLength() )
            {
                *pSchedule += L", ";
            }
            *pSchedule += Trigger;
        }
    }

    // Next Run
    memset( &stNextRun, 0, sizeof ( stNextRun ) );
    hResult = pITask->GetNextRunTime( &stNextRun );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITask::GetNextRunTime", __FUNCTION__ );
    }
    *pIsoNextRun = Format.SystemTimeToIso( stNextRun );

    // Run command
    hResult = pITask->GetApplicationName( &pwszApplicationName );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITask::GetApplicationName", __FUNCTION__ );
    }

    // Need to call CoTaskMemFree
    try
    {
        *pRunCmd = pwszApplicationName;
    }
    catch ( const Exception& )
    {
        CoTaskMemFree( pwszApplicationName );
        throw;
    }
    CoTaskMemFree( pwszApplicationName );
    pwszApplicationName = nullptr;  // Reset

    // Get the task's parameters
    hResult = pITask->GetParameters( &pwszParameters );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITask::GetParameters", __FUNCTION__ );
    }

    // Need to call CoTaskMemFree
    try
    {
        *pRunCmd += PXS_CHAR_SPACE;
        *pRunCmd += pwszParameters;
    }
    catch ( const Exception& )
    {
        CoTaskMemFree( pwszParameters );
        throw;
    }
    CoTaskMemFree( pwszParameters );
    pwszParameters = nullptr;   // Reset

    // Max run time, result is in milli-seconds so convert to an interval
    hResult = pITask->GetMaxRunTime( &dwMaxRunTime );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITask::GetMaxRunTime", __FUNCTION__ );
    }
    *pMaxRunTime = Format.SecondsToDDHHMM( dwMaxRunTime / 1000 );

    // Last Run
    memset( &stLastRun, 0, sizeof ( stLastRun ) );
    hResult = pITask->GetMostRecentRunTime( &stLastRun );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITask::GetMostRecentRunTime", __FUNCTION__ );
    }

    // Reject zeros
    if ( stLastRun.wYear > 0 )
    {
        *pIsoLastRun = Format.SystemTimeToIso( stLastRun );
    }

    // Exit code, if the task has run otherwise result is failure
    if ( hResult != SCHED_S_TASK_HAS_NOT_RUN )  // The task has not yet run.
    {
        hResult = pITask->GetExitCode( &dwExitCode );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"ITask::GetExitCode", __FUNCTION__ );
        }
        *pExitCode = Format.UInt32( dwExitCode );
    }
}

//===============================================================================================//
//  Description:
//      Get the task in the specified Task Scheduler 2.0 folder
//
//  Parameters:
//      pTaskService - pointer to the task interface
//      path         - the folder path
//      pRecords     - receives the audit records
//
//  Remarks:
//      For Task Scheduler 2.0 only, i.e. Vista
//      Must already have initialized COM and security and connected
//      to a task service
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetTasksInFolderRecords( ITaskService* pTaskService,
                                                        BSTR path,
                                                        TArray< AuditRecord >* pRecords )
{
    long      i = 0, count = 0;
    String    Name, State, Schedule, IsoNextRun, Insert2;
    String    RunCmd, MaxRunTime, IsoLastRun, ExitCode;
    VARIANT   index;
    HRESULT   hResult = S_OK;
    Formatter Format;
    AuditRecord  Record;
    ITaskFolder*     pTaskFolder     = nullptr;
    IRegisteredTask* pRegisteredTask = nullptr;
    IRegisteredTaskCollection* pTaskCollection = nullptr;
    AutoIUnknownRelease ReleaseTaskFolder, ReleaseTaskCollection;

    if ( ( pTaskService == nullptr ) || ( path == nullptr ) )
    {
        throw ParameterException( L"pTaskService/path", __FUNCTION__ );
    }

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get the folder
    hResult = pTaskService->GetFolder( path, &pTaskFolder );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITaskService::GetFolder", __FUNCTION__ );
    }

    if ( pTaskFolder == nullptr )
    {
        throw NullException( L"pTaskFolder", __FUNCTION__ );
    }
    ReleaseTaskFolder.Set( pTaskFolder );

    // Get the tasks in the folder
    hResult = pTaskFolder->GetTasks( 0, &pTaskCollection );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITaskFolder::GetTasks", __FUNCTION__ );
    }

    if ( pTaskCollection == nullptr )
    {
        throw NullException( L"pTaskCollection", __FUNCTION__ );
    }
    ReleaseTaskCollection.Set( pTaskCollection );

    // Get the number of tasks
    count = 0;
    hResult = pTaskCollection->get_Count( &count );
    Insert2 = path;
    PXSLogAppInfo2( L"Found %%1 task(s) in folder '%%2'.", Format.Int32( count ), Insert2 );

    for ( i = 0; i < count; i++)
    {
        AutoIUnknownRelease ReleaseRegisteredTask;

        // Set the variants properties, VariantClear is not required as the
        // VARIANT does not reference allocate resources.
        VariantInit( &index );
        index.vt        = VT_I4;
        index.lVal      = i + 1;    // Index is 1-based;
        pRegisteredTask = nullptr;
        hResult = pTaskCollection->get_Item( index, &pRegisteredTask );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"IRegisteredTaskCollection::get_Item", __FUNCTION__ );
        }

        if ( pRegisteredTask == nullptr )
        {
            throw NullException( L"pRegisteredTask", __FUNCTION__ );
        }
        ReleaseRegisteredTask.Set( pRegisteredTask );

        // Catch so can go to the next task
        try
        {
            Name       = PXS_STRING_EMPTY;
            State      = PXS_STRING_EMPTY;
            Schedule   = PXS_STRING_EMPTY;
            IsoNextRun = PXS_STRING_EMPTY;
            RunCmd     = PXS_STRING_EMPTY;
            MaxRunTime = PXS_STRING_EMPTY;
            IsoLastRun = PXS_STRING_EMPTY;
            ExitCode   = PXS_STRING_EMPTY;
            GetRegisteredTaskProperties( pRegisteredTask,
                                         &Name,
                                         &State,
                                         &Schedule,
                                         &IsoNextRun,
                                         &RunCmd, &MaxRunTime, &IsoLastRun, &ExitCode );
            // Make the audit record
            Record.Reset( PXS_CATEGORY_SCHED_TASKS );
            Record.Add( PXS_SCHED_TASKS_TASK_NAME    , Name );
            Record.Add( PXS_SCHED_TASKS_STATUS       , State );
            Record.Add( PXS_SCHED_TASKS_SCHEDULE     , Schedule );
            Record.Add( PXS_SCHED_TASKS_NEXT_RUN_TIME, IsoNextRun );
            Record.Add( PXS_SCHED_TASKS_RUN_COMMAND  , RunCmd );
            Record.Add( PXS_SCHED_TASKS_MAX_RUN_TIME , MaxRunTime );
            Record.Add( PXS_SCHED_TASKS_LAST_RUN_TIME, IsoLastRun );
            Record.Add( PXS_SCHED_TASKS_LAST_RESULT  , ExitCode );
            pRecords->Add( Record );
        }
        catch ( const Exception& e )
        {
            // Log and continue
            PXSLogException( L"Error getting the data for a task.", e, __FUNCTION__ );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get scheduled tasks using the Task Scheduler 1.0
//
//  Parameters:
//      pRecords - array to receive the audit records
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetType1TaskRecords( TArray< AuditRecord >* pRecords )
{
    CLSID       clsid;
    wchar_t     szName[ MAX_PATH + 1 ] = { 0 };     // Enough for a job name
    ULONG       celtFetched = 0;
    ITask*      pITask      = nullptr;
    wchar_t**   rgpwszNames = nullptr;
    HRESULT     hResult;
    String      State, Schedule, IsoNextRun, RunCmd, MaxRunTime;
    String      IsoLastRun, ExitCode, Name;
    Formatter   Format;
    AuditRecord Record;
    ITaskScheduler* pITaskScheduler = nullptr;
    IEnumWorkItems* pIEnumWorkItems = nullptr;
    AutoIUnknownRelease ReleaseITaskScheduler, ReleaseIEnumWorkItems;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Create an instance of the task manager. For some reason when use
    // CLSID_TaskScheduler the result is "Class No Registered"
    memset( &clsid, 0, sizeof ( clsid ) );
    CLSIDFromString( L"{148BD52A-A2AB-11CE-B11F-00AA00530503}", &clsid );
    hResult = CoCreateInstance( clsid,  // CLSID_TaskScheduler,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_ITaskScheduler,
                                reinterpret_cast<void**>( &pITaskScheduler ) );

    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoCreateInstance", __FUNCTION__ );
    }

    if ( pITaskScheduler == nullptr )
    {
        throw NullException( L"pITaskScheduler",  __FUNCTION__ );
    }
    ReleaseITaskScheduler.Set( pITaskScheduler );

    // Get the work items enumerator
    hResult = pITaskScheduler->Enum( &pIEnumWorkItems );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITaskScheduler::Enum", __FUNCTION__ );
    }

    if ( pIEnumWorkItems == nullptr )
    {
        throw NullException( L"pIEnumWorkItems",  __FUNCTION__ );
    }
    ReleaseIEnumWorkItems.Set( pIEnumWorkItems );

    // Fetch each work item, one at a time.
    hResult = pIEnumWorkItems->Next( 1, &rgpwszNames, &celtFetched  );
    while ( ( hResult == S_OK ) && rgpwszNames )
    {
        AutoIUnknownRelease ReleaseITask;

        // Copy the name and release the system allocated memory
        memset( szName, 0, sizeof ( szName ) );
        StringCchCopy( szName, ARRAYSIZE( szName ), rgpwszNames[ 0 ] );
        CoTaskMemFree( rgpwszNames[ 0 ] );
        CoTaskMemFree( rgpwszNames );
        rgpwszNames = nullptr;   // Reset

        // Get ITask
        pITask  = nullptr;
        hResult = pITaskScheduler->Activate( szName,
                                             IID_ITask, reinterpret_cast<IUnknown**>( &pITask ) );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"ITaskScheduler::Activate", __FUNCTION__ );
        }

        if ( pITask == nullptr )
        {
            throw NullException( L"pITask",  __FUNCTION__ );
        }
        ReleaseITask.Set( pITask );

        State      = PXS_STRING_EMPTY;
        Schedule   = PXS_STRING_EMPTY;
        IsoNextRun = PXS_STRING_EMPTY;
        RunCmd     = PXS_STRING_EMPTY;
        MaxRunTime = PXS_STRING_EMPTY;
        IsoLastRun = PXS_STRING_EMPTY;
        ExitCode   = PXS_STRING_EMPTY;
        GetTaskProperties( pITask ,
                           &State,
                           &Schedule, &IsoNextRun, &RunCmd, &MaxRunTime, &IsoLastRun, &ExitCode );

        // Strip ".job" from the name
        Name = szName;
        if ( Name.EndsWithStringI( L".job" ) )
        {
            Name.Truncate( Name.GetLength() - 4 );
        }

        // Make the audit record
        Record.Reset( PXS_CATEGORY_SCHED_TASKS );
        Record.Add( PXS_SCHED_TASKS_TASK_NAME    , Name );
        Record.Add( PXS_SCHED_TASKS_STATUS       , State );
        Record.Add( PXS_SCHED_TASKS_SCHEDULE     , Schedule );
        Record.Add( PXS_SCHED_TASKS_NEXT_RUN_TIME, IsoNextRun );
        Record.Add( PXS_SCHED_TASKS_RUN_COMMAND  , RunCmd );
        Record.Add( PXS_SCHED_TASKS_MAX_RUN_TIME , MaxRunTime );
        Record.Add( PXS_SCHED_TASKS_LAST_RUN_TIME, IsoLastRun );
        Record.Add( PXS_SCHED_TASKS_LAST_RESULT  , ExitCode );
        pRecords->Add( Record );

        // Next
        celtFetched = 0;
        hResult     = pIEnumWorkItems->Next( 1, &rgpwszNames, &celtFetched );
    }
}

//===============================================================================================//
//  Description:
//      Get scheduled tasks using the Task Scheduler 2.0
//
//  Parameters:
//      pRecords - array to receive audit records
//
//  Remarks:
//      Requires Vista or newer
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::GetType2TaskRecords( TArray< AuditRecord >* pRecords )
{
    long    i = 0, count = 0;
    BSTR    name = nullptr;
    BStr    RootPath;
    CLSID   clsid;
    String  WmiPathSeparator;
    HRESULT hResult;
    VARIANT serverName, user, domain, password, index;
    ITaskFolder*  pRootFolder  = nullptr;
    ITaskFolder*  pTaskFolder  = nullptr;
    ITaskService* pTaskService = nullptr;
    ITaskFolderCollection* pFolderCollection = nullptr;
    AutoIUnknownRelease ReleaseTaskService, ReleaseRootFolder;
    AutoIUnknownRelease ReleaseFolderCollection;
    TArray< AuditRecord > Tasks;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();


    // ITaskService
    memset( &clsid, 0, sizeof ( clsid ) );
    CLSIDFromString( L"{0f87369f-a4e5-4cfc-bd3e-73e6154572dd}", &clsid );
    hResult = CoCreateInstance( clsid,  // CLSID_TaskScheduler,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_ITaskService,
                                reinterpret_cast<void**>( &pTaskService ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoCreateInstance", __FUNCTION__  );
    }

    if ( pTaskService == nullptr )
    {
        throw NullException( L"pTaskService", __FUNCTION__  );
    }
    ReleaseTaskService.Set( pTaskService );

    // Connect, set all parameters to VT_EMPTY
    VariantInit( &serverName );
    VariantInit( &user       );
    VariantInit( &domain     );
    VariantInit( &password   );
    hResult = pTaskService->Connect( serverName, user, domain, password );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITaskService::Connect", __FUNCTION__  );
    }

    // Get the tasks in the root folder
    WmiPathSeparator = L"\\";
    RootPath.Allocate( WmiPathSeparator );
    Tasks.RemoveAll();
    GetTasksInFolderRecords( pTaskService, RootPath.b_str(), &Tasks );
    pRecords->Append( Tasks );

    // Get the root's folder collection
    hResult = pTaskService->GetFolder( RootPath.b_str(), &pRootFolder );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ITaskService::GetFolder, pRootFolder", __FUNCTION__  );
    }

    if ( pRootFolder == nullptr )
    {
        throw NullException( L"pRootFolder", __FUNCTION__  );
    }
    ReleaseRootFolder.Set( pRootFolder );

    hResult = pRootFolder->GetFolders( 0, &pFolderCollection );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult,
                            L"ITaskService::GetFolder, pFolderCollection", __FUNCTION__  );
    }

    if ( pFolderCollection == nullptr )
    {
        throw NullException( L"pFolderCollection", __FUNCTION__  );
    }
    ReleaseFolderCollection.Set( pFolderCollection );

    // Look in each folder
    hResult = pFolderCollection->get_Count( &count );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"TaskFolderCollection::get_Count", __FUNCTION__  );
    }

    for ( i = 1; i <= count; i++ )    // Collections are 1-based
    {
        AutoIUnknownRelease ReleaseTaskFolder;

        // Get the folder
        VariantInit( &index );
        index.vt    = VT_I4;
        index.lVal  = i;
        pTaskFolder = nullptr;
        hResult     = pFolderCollection->get_Item( index, &pTaskFolder );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"ITaskFolderCollection::get_Item", __FUNCTION__  );
        }

        if ( pTaskFolder == nullptr )
        {
            throw NullException( L"pTaskFolder", __FUNCTION__  );
        }
        ReleaseTaskFolder.Set( pTaskFolder );

        // Get the name
        name    = nullptr;
        hResult = pTaskFolder->get_Name( &name );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"ITaskFolder::get_Name", __FUNCTION__  );
        }
        AutoSysFreeString AutoSysFreeName( name );

        // Get the tasks and append to the output
        Tasks.RemoveAll();
        GetTasksInFolderRecords( pTaskService, name, &Tasks );
        pRecords->Append( Tasks );
    }
}

//===============================================================================================//
//  Description:
//    Translate a scheduled work item's status
//
//  Parameters:
//      hStatus  - the status
//      pStatus  - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::TranslateScheduledWorkItemStatus( HRESULT hStatus, String* pStatus )
{
    Formatter Format;

    if ( pStatus == nullptr )
    {
        throw ParameterException( L"pStatus", __FUNCTION__ );
    }

    switch ( hStatus )
    {
        default:
            *pStatus = Format.Int32Hex( hStatus );
            break;

       case SCHED_S_TASK_READY:
            *pStatus = L"Ready";
            break;

       case SCHED_S_TASK_RUNNING:
            *pStatus = L"Running";
            break;

       case SCHED_S_TASK_NOT_SCHEDULED:
            *pStatus = L"Not scheduled";
            break;

       case SCHED_S_TASK_HAS_NOT_RUN:
            *pStatus = L"Never run";
            break;

       case SCHED_S_TASK_DISABLED:
            *pStatus = L"Disabled";
            break;

       case SCHED_S_TASK_NO_MORE_RUNS:
            *pStatus = L"No more runs";
            break;

       case SCHED_S_TASK_NO_VALID_TRIGGERS:
            *pStatus = L"No valid triggers";
            break;

       case SCHED_S_TASK_TERMINATED:
            *pStatus = L"User terminated";
            break;

       case SCHED_E_TASK_NOT_READY:
            *pStatus = L"Not ready";
            break;
    }
}

//===============================================================================================//
//  Description:
//    Translate a task action type enumeration constant
//
//  Parameters:
//      actionType  - enumeration constant
//      pActionType - receives the action type
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::TranslateTaskActionType( TASK_ACTION_TYPE type,
                                                        String* pActionType )
{
    Formatter Format;

    if ( pActionType == nullptr )
    {
        throw ParameterException( L"pActionType", __FUNCTION__ );
    }

    switch ( type )
    {
        default:
            *pActionType = Format.UInt32( type );
            break;

        case TASK_ACTION_EXEC:
            *pActionType = L"Command-line operation.";
            break;

        case TASK_ACTION_COM_HANDLER:
            *pActionType = L"Fires a handler.";
            break;

        case TASK_ACTION_SEND_EMAIL:
            *pActionType = L"Send an e-mail.";
            break;

        case TASK_ACTION_SHOW_MESSAGE:
            *pActionType = L"Show a message box.";
            break;
    }
}

//===============================================================================================//
//  Description:
//    Translate a task state constant
//
//  Parameters:
//      state      - enumeration constant
//      pTaskState - receives the status
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::TranslateTaskState( TASK_STATE state, String* pTaskState )
{
    Formatter Format;

    if ( pTaskState == nullptr )
    {
        throw ParameterException( L"pTaskState", __FUNCTION__ );
    }

    switch( state )
    {
        default:
            *pTaskState = Format.Int32( state );
            break;

        case TASK_STATE_UNKNOWN:
            *pTaskState = L"Unknown";
            break;

        case TASK_STATE_DISABLED:
            *pTaskState = L"Disabled";
            break;

        case TASK_STATE_QUEUED:
            *pTaskState = L"Queued";
            break;

        case TASK_STATE_READY:
            *pTaskState = L"Ready";
            break;

        case TASK_STATE_RUNNING:
            *pTaskState = L"Running";
            break;
    }
}

//===============================================================================================//
//  Description:
//    Translate a task trigger type 2 constant
//
//  Parameters:
//      type          - enumeration constant
//      pTriggerType2 - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void TaskSchedulerInformation::TranslateTaskTriggerType2( TASK_TRIGGER_TYPE2 type2,
                                                          String* pTriggerType2 )
{
    Formatter Format;

    if ( pTriggerType2 == nullptr )
    {
        throw ParameterException( L"pTriggerType2", __FUNCTION__ );
    }

    switch( type2 )
    {
        default:

            if ( type2 == 12 )    //  = TASK_TRIGGER_CUSTOM_TRIGGER_01, present in SDK 8.1
            {
                *pTriggerType2 = L"TASK_TRIGGER_CUSTOM_TRIGGER_01";  // No description in ITrigger docs
            }
            else
            {
                *pTriggerType2 = Format.UInt32( type2 );
            }
            break;

        case TASK_TRIGGER_EVENT:
            *pTriggerType2 = L"When a specific event occurs.";
            break;

        case TASK_TRIGGER_TIME:
            *pTriggerType2 = L"At a specific time of day.";
            break;

        case TASK_TRIGGER_DAILY:
            *pTriggerType2 = L"On a daily schedule.";
            break;

        case TASK_TRIGGER_WEEKLY:
            *pTriggerType2 = L"On a weekly schedule.";
            break;

        case TASK_TRIGGER_MONTHLY:
            *pTriggerType2 = L"On a monthly schedule.";
            break;

        case TASK_TRIGGER_MONTHLYDOW:
            *pTriggerType2 = L"On a monthly day-of-week schedule.";
            break;

        case TASK_TRIGGER_IDLE:
            *pTriggerType2 = L"When the computer goes into an idle state.";
            break;

        case TASK_TRIGGER_REGISTRATION:
            *pTriggerType2 = L"When the task is registered.";
            break;

        case TASK_TRIGGER_BOOT:
            *pTriggerType2 = L"When the computer boots.";
            break;

        case TASK_TRIGGER_LOGON:
            *pTriggerType2 = L"When a specific user logs on.";
            break;

        case TASK_TRIGGER_SESSION_STATE_CHANGE:
            *pTriggerType2 = L"When a specific session state changes.";
            break;
    }
}
